"""MATLABの静的な依存関係を取得し、対応するファイルを比較する。"""

from __future__ import annotations

from collections import defaultdict, deque
from dataclasses import dataclass
import hashlib
import json
import os
from pathlib import Path
import shutil
import subprocess
from tempfile import TemporaryDirectory


@dataclass(frozen=True)
class AnalysisRequest:
    source: Path
    workdir: Path
    search_dirs: tuple[Path, ...] = ()


@dataclass(frozen=True)
class Analysis:
    request: AnalysisRequest
    nodes: dict[Path, tuple[Path, ...]]
    notices: tuple[str, ...]
    error: str = ""


def find_matlab() -> Path | None:
    found = shutil.which("matlab")
    if found:
        return Path(found)
    if os.name == "nt":
        location = Path(os.environ.get("ProgramFiles", "C:/Program Files")) / "MATLAB"
        return max(location.glob("R*/bin/matlab.exe"), default=None)
    return None


def analyze_matlab(executable: Path, request: AnalysisRequest) -> Analysis:
    with TemporaryDirectory(prefix="folder-organizer-matlab-") as work:
        folder = Path(work)
        shutil.copyfile(Path(__file__).with_name("folder_organizer_matlab_scan.m"),
                        folder / "folder_organizer_matlab_scan.m")
        config = {
            "source": str(request.source),
            "workdir": str(request.workdir),
            "searchDirs": [str(path) for path in request.search_dirs],
        }
        (folder / "request.json").write_text(json.dumps(config), encoding="ascii")
        arguments = [str(executable)]
        if os.name == "nt":
            arguments.append("-wait")
        command = "addpath(getenv('FOLDER_ORGANIZER_RUN')); folder_organizer_matlab_scan"
        arguments.extend(["-sd", work, "-batch", command])
        environment = dict(os.environ, FOLDER_ORGANIZER_RUN=work)
        log = folder / "matlab.log"
        try:
            with log.open("wb") as stream:
                process = subprocess.run(
                    arguments, cwd=folder, env=environment, stdin=subprocess.DEVNULL,
                    stdout=stream, stderr=subprocess.STDOUT, timeout=180,
                    creationflags=getattr(subprocess, "CREATE_NO_WINDOW", 0),
                )
        except subprocess.TimeoutExpired:
            raise RuntimeError("MATLABの起動または解析が３分以内に終わりませんでした。") from None

        output = folder / "result.json"
        if process.returncode or not output.is_file():
            with log.open("rb") as stream:
                stream.seek(max(log.stat().st_size - 4000, 0))
                detail = stream.read().decode("utf-8", errors="replace").strip()
            raise RuntimeError(detail or f"MATLABから解析結果を取得できませんでした（終了コード {process.returncode}）。")

        report = json.loads(output.read_text(encoding="utf-8-sig"))
        nodes = {Path(node["file"]): tuple(Path(value) for value in node["dependencies"])
                 for node in report["nodes"]}
        if request.source not in nodes or any(child not in nodes for children in nodes.values() for child in children):
            raise RuntimeError("MATLABから必要なファイル情報がそろった解析結果を取得できませんでした。")
        return Analysis(request, nodes, tuple(report["notices"]))


def dependency_label(path: Path) -> str:
    parts = [path.name]
    for parent in path.parents:
        if parent.name.startswith(("+", "@")) or parent.name.casefold() == "private":
            parts.insert(0, parent.name)
        else:
            break
    return "/".join(parts)


def dependency_report(analyses: list[Analysis]) -> str:
    report = ["MATLABが静的に検出した参照先を、ファイルのバイト単位で比較しました。",
              "実行時に決まる参照先など、静的解析で検出されない依存関係があります。"]
    hashes: dict[Path, bytes | None] = {}
    try:
        common_parent = Path(os.path.commonpath([entry.request.source.parent for entry in analyses]))
    except ValueError:
        common_parent = None

    def source_name(path: Path) -> str:
        return str(path.relative_to(common_parent)) if common_parent is not None else str(path)

    def digest(path: Path) -> bytes | None:
        if path not in hashes:
            try:
                value = hashlib.sha256()
                with path.open("rb") as stream:
                    for block in iter(lambda: stream.read(1024 * 1024), b""):
                        value.update(block)
                hashes[path] = value.digest()
            except OSError:
                hashes[path] = None
        return hashes[path]

    baseline = analyses[0]
    for entry in analyses[1:]:
        left_source, right_source = baseline.request.source, entry.request.source
        report.extend(["", f"=== {source_name(left_source)} ↔ {source_name(right_source)} ==="])
        if baseline.error or entry.error:
            report.append("依存関係を判定できず")
            continue

        pending = deque([(left_source, right_source, "")])
        seen = set()
        details = []
        changed = False
        uncertain = bool(baseline.notices or entry.notices)
        while pending:
            left, right, parent = pending.popleft()
            if (left, right) in seen:
                continue
            seen.add((left, right))
            left_hash, right_hash = digest(left), digest(right)
            if left_hash is None or right_hash is None:
                status = "読み取り不可"
                uncertain = True
            elif left_hash != right_hash:
                status = "内容に差分"
                if parent:
                    changed = True
            else:
                status = "完全一致"
            heading = f"{parent} → {dependency_label(left)}" if parent else "本体"
            details.extend([f"  {heading}  [{status}]", f"    A  {left}", f"    B  {right}"])

            children = []
            for graph, path in ((baseline.nodes, left), (entry.nodes, right)):
                by_name = defaultdict(list)
                for dependency in graph.get(path, ()):
                    by_name[dependency_label(dependency)].append(dependency)
                children.append(by_name)
            first, second = children
            chain = heading if parent else dependency_label(left)
            for label in sorted(first.keys() | second.keys()):
                lpaths, rpaths = first.get(label, []), second.get(label, [])
                if not lpaths or not rpaths:
                    changed = True
                    details.append(f"  {chain} → {label}  [片方でのみ検出]")
                elif len(lpaths) != 1 or len(rpaths) != 1:
                    uncertain = True
                    details.append(f"  {chain} → {label}  [複数候補の対応を判定できず]")
                else:
                    pending.append((lpaths[0], rpaths[0], chain))
                    continue
                details.extend(f"    {side}  {path}" for side, paths in (("A", lpaths), ("B", rpaths))
                               for path in paths)

        if changed:
            report.append("検出された参照先に差分あり")
        if uncertain:
            report.append("参照先の一部を判定できず")
        if not changed and not uncertain:
            report.append("検出された参照先の内容が一致" if len(seen) > 1 else "比較対象となる参照先ファイルは検出されませんでした。")
        report.extend(details)

    report.extend(["", "=== 解析に使った設定と注意 ==="])
    for entry in analyses:
        config = entry.request
        report.extend(["", f"対象  {config.source}", f"作業フォルダ  {config.workdir}"])
        report.extend(f"追加の参照フォルダ  {folder}" for folder in config.search_dirs)
        if entry.error:
            report.append(f"解析できず  {entry.error}")
        report.extend(f"解析の注意  {notice}" for notice in entry.notices)
    return "\n".join(report)
