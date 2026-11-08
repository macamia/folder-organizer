"""選んだフォルダ内のファイルを比較し、グループ単位で整理する。"""

from __future__ import annotations

import difflib
import hashlib
import os
import posixpath
import re
import stat
import threading
import unicodedata
from dataclasses import dataclass
from datetime import datetime
from pathlib import Path
from queue import Empty, Queue
import shutil
from tempfile import mkstemp
from xml.etree import ElementTree
from zipfile import BadZipFile, ZipFile
import tkinter as tk
from tkinter import filedialog, font, messagebox, simpledialog, ttk

from matlab_dependencies import Analysis, AnalysisRequest, analyze_matlab, dependency_report, find_matlab
from matlab_version_diff import change_blocks, code_line_match_rate


TITLE = "フォルダ仕分け"
BACKUP = "backup"
LEGACY_BACKUP = "バックアップ"
TEXT_TYPES = {".m", ".py", ".txt", ".md", ".csv"}
CODE_TYPES = {".m", ".mlx", ".py"}
DOCUMENT_TYPES = {".pdf", ".pptx", ".ppt", ".docx", ".doc", ".xlsx", ".xls"}
SUPPORTED = TEXT_TYPES | CODE_TYPES | DOCUMENT_TYPES
TEXT_LIMIT = 256 * 1024
OFFICE_TYPES = {".docx", ".pptx", ".xlsx"}
XML_LIMIT = 16 * 1024 * 1024
REL_ID = "{http://schemas.openxmlformats.org/officeDocument/2006/relationships}id"
VERSION = re.compile(
    r"(?:[ _.-]+(?:v(?:er)?[ ._-]?\d+(?:\.\d+)*|rev[ ._-]?\d+|\d{1,3}|"
    r"最終版?|確定版?|最新版?|修正版?|旧版|コピー|final|draft|copy)|"
    r"\s*\(\s*(?:\d+|最終版?|修正版?)\s*\))$", re.IGNORECASE
)
INVALID_NAME = re.compile(r'[<>:"/\\|?*\x00-\x1f]|[. ]$')
RESERVED_NAME = re.compile(r"(?:CON|PRN|AUX|NUL|COM[1-9¹²³]|LPT[1-9¹²³])(?:\.|$)", re.IGNORECASE)


def stem_without_version(stem: str) -> str:
    while match := VERSION.search(stem):
        shortened = stem[:match.start()].strip()
        if not shortened:
            break
        stem = shortened
    return stem


def normalized_stem(stem: str) -> str:
    return stem_without_version(unicodedata.normalize("NFKC", stem)).casefold()


def file_stamp(info: os.stat_result) -> tuple[int, ...]:
    return info.st_dev, info.st_ino, info.st_size, info.st_mtime_ns, info.st_ctime_ns


@dataclass(frozen=True)
class File:
    path: Path
    stamp: tuple[int, ...]
    text: str | None
    digest: bytes | None
    shingles: frozenset[tuple[str, str, str]]

    @property
    def ext(self) -> str:
        return self.path.suffix.casefold()

    @property
    def stem(self) -> str:
        return normalized_stem(self.path.stem)


def inspect_file(path: Path) -> File:
    info = path.stat(follow_symlinks=False)
    if not stat.S_ISREG(info.st_mode):
        raise OSError("通常のファイルではありません。")
    text, digest = None, None
    shingles = frozenset()
    if path.suffix.casefold() in TEXT_TYPES and info.st_size <= TEXT_LIMIT:
        with path.open("rb") as stream:
            data = stream.read(TEXT_LIMIT + 1)
        if len(data) <= TEXT_LIMIT:
            digest = hashlib.sha256(data).digest() if data else None
            encodings = ("utf-16",) if data.startswith((b"\xff\xfe", b"\xfe\xff")) else ("utf-8-sig", "cp932")
            for encoding in encodings:
                try:
                    decoded = data.decode(encoding)
                except UnicodeError:
                    continue
                if "\0" not in decoded:
                    text = decoded
                break
            if text is not None:
                lines = [line.strip() for line in text.splitlines() if line.strip()]
                if len(lines) >= 8:
                    shingles = frozenset(zip(lines, lines[1:], lines[2:]))
    return File(path, file_stamp(info), text, digest, shingles)


def similarity(left: File, right: File) -> float | None:
    if left.digest is not None and left.digest == right.digest:
        return 1.0
    if left.shingles and right.shingles:
        shared = len(left.shingles & right.shingles)
        return shared / (len(left.shingles) + len(right.shingles) - shared)
    return None


def office_lines(path: Path) -> list[str]:
    with ZipFile(path) as archive:
        budget = XML_LIMIT

        def xml(name):
            nonlocal budget
            budget -= archive.getinfo(name).file_size
            if budget < 0:
                raise ValueError("Office XML is too large")
            return ElementTree.fromstring(archive.read(name))

        def target_paths(document):
            parent, filename = posixpath.split(document)
            relations = xml(f"{parent}/_rels/{filename}.rels")
            return {
                relation.attrib["Id"]: posixpath.normpath(
                    posixpath.join(parent, target) if not target.startswith("/") else target.lstrip("/")
                )
                for relation in relations
                if relation.get("TargetMode") != "External" and (target := relation.get("Target"))
            }

        if path.suffix.casefold() == ".docx":
            document = xml("word/document.xml")
            return ["".join(node.text or "" for node in paragraph.iterfind(".//{*}t"))
                    for paragraph in document.iterfind(".//{*}p")]

        if path.suffix.casefold() == ".pptx":
            index = "ppt/presentation.xml"
            relations = target_paths(index)
            lines = []
            for number, entry in enumerate(xml(index).iterfind("./{*}sldIdLst/{*}sldId"), start=1):
                slide = xml(relations[entry.attrib[REL_ID]])
                lines.extend(f"[スライド {number}] {text}"
                             for paragraph in slide.iterfind(".//{*}p")
                             if (text := "".join(node.text or "" for node in paragraph.iterfind(".//{*}t"))))
            return lines

        index = "xl/workbook.xml"
        relations = target_paths(index)
        shared = []
        if "xl/sharedStrings.xml" in archive.namelist():
            shared = ["".join(node.text or "" for node in value.iterfind(".//{*}t"))
                      for value in xml("xl/sharedStrings.xml").iterfind("./{*}si")]
        lines = []
        for entry in xml(index).iterfind("./{*}sheets/{*}sheet"):
            sheet = xml(relations[entry.attrib[REL_ID]])
            for cell in sheet.iterfind("./{*}sheetData/{*}row/{*}c"):
                formula = cell.findtext("{*}f")
                value = cell.findtext("{*}v")
                kind = cell.get("t")
                if formula:
                    value = "=" + formula
                elif kind == "s" and value is not None:
                    value = shared[int(value)]
                elif kind == "inlineStr":
                    value = "".join(node.text or "" for node in cell.iterfind("./{*}is//{*}t"))
                if value is not None:
                    safe_value = value.replace("\r\n", "\n").replace("\n", "↵")
                    lines.append(f"[{entry.get('name')}] {cell.get('r', '')}  {safe_value}")
        return lines


def comparison_lines(item: File) -> list[str] | None:
    if item.text is not None:
        return item.text.splitlines()
    if item.ext not in OFFICE_TYPES:
        return None
    try:
        lines = office_lines(item.path)
        if sum(len(line.encode("utf-8")) + 1 for line in lines) <= TEXT_LIMIT:
            return lines
    except (OSError, BadZipFile, RuntimeError, ElementTree.ParseError, KeyError, IndexError, ValueError):
        pass
    return None


def common_positions(versions: list[list[str]]) -> list[set[int]]:
    reference = versions[0]
    shared = list(range(len(reference)))
    positions = []
    for lines in versions[1:]:
        # Match each new version against the subsequence already shared by the earlier versions.
        candidate = [reference[index] for index in shared]
        matcher = difflib.SequenceMatcher(None, candidate, lines,
                                          autojunk=max(len(candidate), len(lines)) > 2000)
        offsets = {shared[left + n]: right + n for left, right, length in matcher.get_matching_blocks()
                   for n in range(length)}
        shared = list(offsets)
        positions.append(offsets)
    return [set(shared), *({offsets[index] for index in shared} for offsets in positions)]


def grouped_with(left: File, right: File) -> bool:
    if left.ext != right.ext:
        return False
    if left.stem == right.stem:
        return True
    score = similarity(left, right)
    return score is not None and score >= 0.75


def scan_paths(folder: Path, include_subfolders: bool, on_error):
    if not include_subfolders:
        yield from sorted(folder.iterdir(), key=lambda item: item.name.casefold())
        return
    for directory, children, files in os.walk(folder, onerror=on_error):
        parent = Path(directory)
        searchable = []
        for name in sorted(children, key=str.casefold):
            if name.startswith('.') or name.casefold() in {BACKUP, LEGACY_BACKUP}:
                continue
            try:
                check_directory(parent / name)
            except OSError as error:
                on_error(error)
                continue
            searchable.append(name)
        children[:] = searchable
        yield from (parent / name for name in sorted(files, key=str.casefold))


def scan(folder: Path, include_subfolders: bool = False) -> tuple[list[list[File]], int]:
    groups: list[list[File]] = []
    skipped = 0

    def skipped_directory(_error):
        nonlocal skipped
        skipped += 1

    for path in scan_paths(folder, include_subfolders, skipped_directory):
        if path.suffix.casefold() not in SUPPORTED or path.name.startswith(("~$", ".")):
            continue
        try:
            item = inspect_file(path)
        except OSError:
            skipped += 1
            continue
        for group in groups:
            if all(grouped_with(item, other) for other in group):
                group.append(item)
                break
        else:
            groups.append([item])
    for group in groups:
        group.sort(key=lambda item: item.stamp[3], reverse=True)
    groups.sort(key=lambda group: (len(group) == 1, group[0].path.name.casefold()))
    return groups, skipped


def content_digest(item: File, cache: dict[Path, bytes]) -> bytes:
    if item.digest is not None:
        return item.digest
    if item.path not in cache:
        digest = hashlib.sha256()
        with item.path.open("rb") as stream:
            for block in iter(lambda: stream.read(1024 * 1024), b""):
                digest.update(block)
        cache[item.path] = digest.digest()
    return cache[item.path]


def find_representatives(folder: Path, group: list[File]) -> tuple[list[tuple[File, str]], int]:
    skipped = 0
    matches = []
    cache: dict[Path, bytes] = {}
    paths = {item.path for item in group}
    by_size: dict[int, list[File]] = {}
    for item in group:
        by_size.setdefault(item.stamp[2], []).append(item)

    def code_lines(item):
        return [line.strip() for line in (item.text or "").splitlines() if line.strip()]

    baselines = [(item, code_lines(item)) for item in group]

    def skip(_error):
        nonlocal skipped
        skipped += 1

    for path in scan_paths(folder, True, skip):
        if path in paths or path.suffix.casefold() != group[0].ext or path.name.startswith(('~$', '.')):
            continue
        try:
            candidate = inspect_file(path)
            identical = any(content_digest(item, cache) == content_digest(candidate, cache)
                            for item in by_size.get(candidate.stamp[2], ()))
        except OSError as error:
            skip(error)
            continue
        lines = code_lines(candidate)
        score = 0.0
        for item, base in baselines:
            score = max(score, similarity(candidate, item) or 0)
            if min(len(lines), len(base)) >= 4:
                matcher = difflib.SequenceMatcher(None, base, lines, autojunk=max(len(lines), len(base)) > 2000)
                if matcher.quick_ratio() >= 0.75:
                    score = max(score, matcher.ratio())
        same_name = any(candidate.stem == item.stem for item in group)
        if identical or score >= 0.75 or same_name:
            reason = "完全一致" if identical else f"内容の類似度 {score:.0%}" if score >= 0.75 else "同系統の名前"
            matches.append((not identical, -score, not same_name, str(path).casefold(), candidate, reason))
    matches.sort(key=lambda row: row[:4])
    return [(row[4], row[5]) for row in matches], skipped


@dataclass(frozen=True)
class Move:
    file: File
    target: Path


def backup_location(folder: Path, label: str, existing: dict[str, Path]) -> Path:
    readable = re.sub(r'[<>:"/\\|?*\x00-\x1f]', "_", label).strip(" .")[:45].rstrip(" .")
    readable = readable or "グループ"
    if RESERVED_NAME.match(readable):
        readable = "_" + readable

    root = existing.get(BACKUP, folder / BACKUP)
    legacy = existing.get(LEGACY_BACKUP)
    if not root.is_dir() and legacy is not None and legacy.is_dir():
        root = legacy
    day = root / datetime.now().date().isoformat()
    group_folder = day / readable
    for directory in (root, day, group_folder):
        if exists(directory):
            check_directory(directory)
    return group_folder


def archive_moves(group: list[File], archived: set[Path], existing: dict[str, Path]) -> list[Move]:
    moves = []
    if archived:
        folder = group[0].path.parent
        backup = backup_location(folder, stem_without_version(group[0].path.stem), existing)
        taken = {path.name.casefold() for path in backup.iterdir()} if backup.is_dir() else set()
        original_names = {path.name.casefold() for path in archived}
        for item in group:
            if item.path not in archived:
                continue
            name = item.path.name
            if name.casefold() in taken:
                number = 2
                while True:
                    name = f"{item.path.stem} ({number}){item.path.suffix}"
                    if name.casefold() not in taken and name.casefold() not in original_names:
                        break
                    number += 1
            taken.add(name.casefold())
            moves.append(Move(item, backup / name))
    return moves


def validate_name(name: str, item: File) -> None:
    if not name or INVALID_NAME.search(name) or RESERVED_NAME.match(name):
        raise ValueError(f"使えないファイル名です。\n{name}")
    if Path(name).suffix.casefold() != item.ext:
        raise ValueError(f"拡張子は元と同じ {item.ext} にしてください。")


def plan_moves(group: list[File], names: dict[Path, str]) -> list[Move]:
    if not names:
        raise ValueError("代表として残すファイルを１件以上選んでください。")
    folder = group[0].path.parent
    if any(item.path.parent != folder for item in group):
        raise ValueError("複数フォルダのグループでは比較ができます。整理するには対象のフォルダを選んでください。")
    existing = {path.name.casefold(): path for path in folder.iterdir()}
    archived = {item.path for item in group if item.path not in names}
    moves = archive_moves(group, archived, existing)
    used = set()
    for item in group:
        if item.path not in names:
            continue
        name = names[item.path].strip()
        validate_name(name, item)
        key = name.casefold()
        if key in used:
            raise ValueError(f"代表の出力名が重複しています。\n{name}")
        used.add(key)
        occupied = existing.get(key)
        if occupied is not None and occupied != item.path and occupied not in archived:
            raise ValueError(f"同じ名前がすでにあります。\n{name}")
        if item.path.name != name:
            moves.append(Move(item, folder / name))
    return moves


def plan_import(group: list[File], representative: File, archived: set[Path], name: str):
    folder = group[0].path.parent
    if any(item.path.parent != folder for item in group):
        raise ValueError("整理先は１つのフォルダを選んでください。")
    if representative.ext != group[0].ext:
        raise ValueError(f"代表版は {group[0].ext} のファイルを選んでください。")
    if representative.path in archived or not archived <= {item.path for item in group}:
        raise ValueError("バックアップに移すファイルの選択を確認してください。")
    validate_name(name, representative)
    existing = {path.name.casefold(): path for path in folder.iterdir()}
    occupied = existing.get(name.casefold())
    if occupied is not None and occupied != representative.path and occupied not in archived:
        raise ValueError(f"同じ名前のファイルを残す設定になっています。\n{name}\n残す名前か、移動対象を変更してください。")
    moves = archive_moves(group, archived, existing)
    if representative.path.parent != folder:
        return moves, folder / name
    if representative.path.name != name:
        moves.append(Move(representative, folder / name))
    return moves, None


def check_directory(path: Path) -> None:
    info = path.lstat()
    reparse = getattr(info, "st_file_attributes", 0) & getattr(stat, "FILE_ATTRIBUTE_REPARSE_POINT", 0)
    if not stat.S_ISDIR(info.st_mode) or reparse:
        raise OSError(f"バックアップ先に通常のフォルダを選べません。\n{path}")


def exists(path: Path) -> bool:
    return path.exists() or path.is_symlink()


def move_without_overwrite(source: Path, target: Path) -> None:
    # Windows の rename は既存の宛先を置き換えない。
    if os.name == "nt":
        source.rename(target)
        return
    # 開発・検証環境でも既存ファイルを置き換えない。
    os.link(source, target)
    try:
        source.unlink()
    except OSError:
        target.unlink()
        raise


def move_file(source: Path, target: Path) -> None:
    if source.parent == target.parent and source.name.casefold() == target.name.casefold():
        # 同じディレクトリ内の大文字小文字だけの変更を扱う。
        temporary = source.with_name(f".folder-organizer-{os.urandom(12).hex()}.tmp")
        move_without_overwrite(source, temporary)
        try:
            move_without_overwrite(temporary, target)
        except OSError as error:
            try:
                move_without_overwrite(temporary, source)
            except OSError as restore_error:
                raise OSError(f"{error}\n一時保存先 {temporary}\n{restore_error}") from error
            raise
    else:
        move_without_overwrite(source, target)


def execute(group: list[File], moves: list[Move], copy: tuple[File, Path] | None = None) -> None:
    for item in group:
        if file_stamp(item.path.stat(follow_symlinks=False)) != item.stamp:
            raise OSError(f"スキャン後にファイルが更新されました。再スキャンしてください。\n{item.path.name}")

    folder = group[0].path.parent
    created = []
    completed = []
    staging = None
    try:
        if copy:
            original, destination = copy
            if file_stamp(original.path.stat(follow_symlinks=False)) != original.stamp:
                raise OSError(f"代表版が更新されました。再選択してください。\n{original.path}")
            descriptor, temporary = mkstemp(prefix=".folder-organizer-", suffix=".tmp", dir=folder)
            staging = Path(temporary)
            os.close(descriptor)
            shutil.copyfile(original.path, staging)
            if file_stamp(original.path.stat(follow_symlinks=False)) != original.stamp:
                raise OSError(f"コピー中に代表版が更新されました。\n{original.path}")
        backup = next((move.target.parent for move in moves if move.target.parent != folder), None)
        if backup is not None:
            directory = folder
            for part in backup.relative_to(folder).parts:
                directory = directory / part
                try:
                    directory.mkdir()
                except FileExistsError:
                    check_directory(directory)
                else:
                    created.append(directory)
        for move in moves:
            source = move.file.path
            if file_stamp(source.stat(follow_symlinks=False)) != move.file.stamp:
                raise OSError(f"ファイルが更新されました。\n{source.name}")
            move_file(source, move.target)
            completed.append(move)
        if copy:
            move_without_overwrite(staging, destination)
    except OSError as error:
        recovery = []
        for move in reversed(completed):
            try:
                move_file(move.target, move.file.path)
            except OSError as restore_error:
                recovery.append(f"{move.target}\n{restore_error}")
        for directory in reversed(created):
            try:
                directory.rmdir()
            except OSError:
                pass
        if recovery:
            raise OSError(f"{error}\n\n元の場所に戻せなかったファイル\n" + "\n".join(recovery)) from error
        raise
    finally:
        if staging is not None:
            try:
                staging.unlink(missing_ok=True)
            except OSError:
                pass


class Organizer(tk.Tk):
    def __init__(self):
        super().__init__()
        families = {family.casefold(): family for family in font.families(self)}
        for face in ("Yu Gothic UI", "Meiryo", "Noto Sans CJK JP"):
            if face.casefold() in families:
                for name in ("TkDefaultFont", "TkTextFont", "TkHeadingFont", "TkMenuFont"):
                    font.nametofont(name, self).configure(family=families[face.casefold()], size=10)
                break
        for face in ("MS Gothic", "Noto Sans Mono CJK JP"):
            if face.casefold() in families:
                font.nametofont("TkFixedFont", self).configure(family=families[face.casefold()], size=10)
                break
        self.title(f"{TITLE} 0.12")
        self.geometry("1050x690")
        self.minsize(800, 540)
        self.folder: Path | None = None
        self.groups: list[list[File]] = []
        self.group: list[File] = []
        self.checks: list[tuple[File, tk.BooleanVar]] = []
        self.result = Queue()
        self.matlab_path: Path | None = None
        self.detail_pairs: list[tuple[File, File]] = []

        outer = ttk.Frame(self, padding=14)
        outer.pack(fill="both", expand=True)
        toolbar = ttk.Frame(outer)
        toolbar.pack(fill="x", pady=(0, 10))
        self.choose_button = ttk.Button(toolbar, text="フォルダを選ぶ…", command=self.choose)
        self.choose_button.pack(side="left")
        self.refresh_button = ttk.Button(toolbar, text="再スキャン", command=self.refresh, state="disabled")
        self.refresh_button.pack(side="left", padx=6)
        self.include_subfolders = tk.BooleanVar(value=False)
        self.recursive_button = ttk.Checkbutton(toolbar, text="サブフォルダも含める",
                                                variable=self.include_subfolders, command=self.refresh)
        self.recursive_button.pack(side="left", padx=6)
        self.location = ttk.Label(toolbar, text="整理するフォルダを選んでください。")
        self.location.pack(side="left", fill="x", expand=True, padx=8)

        panes = ttk.Panedwindow(outer, orient="horizontal")
        panes.pack(fill="both", expand=True)
        left = ttk.Frame(panes, padding=(0, 0, 8, 0))
        right = ttk.Frame(panes, padding=(8, 0, 0, 0))
        panes.add(left, weight=1)
        panes.add(right, weight=3)

        ttk.Label(left, text="グループ候補").pack(anchor="w", pady=(0, 5))
        self.show_single = tk.BooleanVar(value=False)
        ttk.Checkbutton(left, text="単独ファイルも表示", variable=self.show_single,
                        command=self.show_groups).pack(anchor="w", pady=(0, 5))
        self.group_tree = self.tree(left, (("label", "グループ", 190), ("count", "件数", 52)), 14)
        self.group_tree.bind("<<TreeviewSelect>>", self.select_group)
        self.import_button = ttk.Button(left, text="代表バージョンを持ってきて整理",
                                        command=self.choose_import_representative, state="disabled")
        self.import_button.pack(anchor="w", pady=(6, 0))

        ttk.Label(right, text="グループ内のファイル").pack(anchor="w", pady=(0, 5))
        self.group_note = ttk.Label(right, text="左の一覧からグループを選んでください。")
        self.group_note.pack(anchor="w", pady=(0, 5))
        listing = ttk.Frame(right)
        listing.pack(fill="x")
        self.file_canvas = tk.Canvas(listing, height=165, highlightthickness=1,
                                     highlightbackground="#b8b8b8", background="white")
        vertical = ttk.Scrollbar(listing, orient="vertical", command=self.file_canvas.yview)
        horizontal = ttk.Scrollbar(listing, orient="horizontal", command=self.file_canvas.xview)
        self.file_canvas.configure(yscrollcommand=vertical.set, xscrollcommand=horizontal.set)
        self.file_canvas.grid(row=0, column=0, sticky="ew")
        vertical.grid(row=0, column=1, sticky="ns")
        horizontal.grid(row=1, column=0, sticky="ew")
        listing.columnconfigure(0, weight=1)
        self.file_rows = ttk.Frame(self.file_canvas, padding=(5, 3))
        self.rows_window = self.file_canvas.create_window(0, 0, window=self.file_rows, anchor="nw")
        self.file_rows.columnconfigure(0, weight=1)
        self.file_rows.bind("<Configure>", self.resize_file_rows)
        self.file_canvas.bind("<Configure>", self.resize_file_rows)
        self.bind_file_scroll(self.file_canvas)
        self.bind_file_scroll(self.file_rows)

        controls = ttk.Frame(right)
        controls.pack(fill="x", pady=7)
        self.keep_button = ttk.Button(controls, text="代表を選んで整理", command=self.choose_representative, state="disabled")
        self.compare_button = ttk.Button(controls, text="コードを比較する", command=self.compare, state="disabled")
        self.dependency_button = ttk.Button(controls, text="参照先を比較する",
                                             command=self.compare_dependencies, state="disabled")
        self.selection_note = ttk.Label(controls, text="")

        def lay_out_controls(event):
            row, column, remaining = 0, 0, event.width
            for control in (self.keep_button, self.compare_button, self.dependency_button, self.selection_note):
                width = control.winfo_reqwidth() + 6
                if column and width > remaining:
                    row, column, remaining = row + 1, 0, event.width
                control.grid(row=row, column=column, sticky="w", padx=(0, 6), pady=(0, 3))
                column += 1
                remaining -= width

        controls.bind("<Configure>", lay_out_controls)
        for column, control in enumerate((self.keep_button, self.compare_button, self.dependency_button, self.selection_note)):
            control.grid(row=0, column=column, sticky="w", padx=(0, 6), pady=(0, 3))

        self.diff_label = ttk.Label(right, text="差分")
        self.diff_label.pack(anchor="w")
        diff_frame = ttk.Frame(right)
        diff_frame.pack(fill="both", expand=True, pady=(5, 0))
        self.diff = tk.Text(diff_frame, wrap="none", height=10, state="disabled", font="TkFixedFont")
        ybar = ttk.Scrollbar(diff_frame, orient="vertical", command=self.diff.yview)
        xbar = ttk.Scrollbar(diff_frame, orient="horizontal", command=self.diff.xview)
        self.diff.configure(yscrollcommand=ybar.set, xscrollcommand=xbar.set)
        self.diff.grid(row=0, column=0, sticky="nsew")
        ybar.grid(row=0, column=1, sticky="ns")
        xbar.grid(row=1, column=0, sticky="ew")
        diff_frame.rowconfigure(0, weight=1)
        diff_frame.columnconfigure(0, weight=1)
        self.diff.tag_configure("added", background="#e4f5e8")
        self.diff.tag_configure("removed", background="#fbe8e6")
        self.diff.tag_configure("hunk", foreground="#285f9e")
        self.detail_button = ttk.Button(right, text="差分の詳細", command=self.show_diff_details)

        footer = ttk.Frame(outer)
        footer.pack(fill="x", pady=(9, 0))
        self.status = ttk.Label(footer, text="ファイルの分類と比較はこのPC内で行います。", wraplength=600)
        self.status.pack(side="left", fill="x", expand=True)
        self.update_buttons()

    @staticmethod
    def tree(parent, columns, height):
        frame = ttk.Frame(parent)
        frame.pack(fill="both", expand=(height == 14))
        table = ttk.Treeview(frame, columns=[column[0] for column in columns],
                             show="headings", height=height, selectmode="browse")
        for key, label, width in columns:
            table.heading(key, text=label)
            table.column(key, width=width, minwidth=min(width, 45), stretch=key in {"label", "file", "name"})
        bar = ttk.Scrollbar(frame, orient="vertical", command=table.yview)
        horizontal = ttk.Scrollbar(frame, orient="horizontal", command=table.xview)
        table.configure(yscrollcommand=bar.set, xscrollcommand=horizontal.set)
        table.grid(row=0, column=0, sticky="nsew")
        bar.grid(row=0, column=1, sticky="ns")
        horizontal.grid(row=1, column=0, sticky="ew")
        frame.rowconfigure(0, weight=1)
        frame.columnconfigure(0, weight=1)
        return table

    def choose(self):
        selected = filedialog.askdirectory(parent=self, mustexist=True)
        if selected:
            self.folder = Path(selected).resolve()
            self.refresh()

    def refresh(self):
        if self.folder is None:
            return
        self.choose_button.state(["disabled"])
        self.refresh_button.state(["disabled"])
        self.recursive_button.state(["disabled"])
        self.groups = []
        self.show_groups()
        self.location.configure(text=str(self.folder))
        self.status.configure(text="フォルダをスキャンしています…")
        threading.Thread(target=self.scan_worker, args=(self.folder, self.include_subfolders.get()), daemon=True).start()
        self.after(70, self.poll)

    def scan_worker(self, folder, recursive):
        try:
            self.result.put(scan(folder, recursive))
        except Exception as error:
            self.result.put(error)

    def poll(self):
        try:
            result = self.result.get_nowait()
        except Empty:
            self.after(70, self.poll)
            return
        self.choose_button.state(["!disabled"])
        self.refresh_button.state(["!disabled"])
        self.recursive_button.state(["!disabled"])
        if isinstance(result, Exception):
            self.status.configure(text="スキャンできませんでした。")
            messagebox.showerror(TITLE, str(result), parent=self)
            return
        self.groups, skipped = result
        self.show_groups()
        count = sum(len(group) for group in self.groups)
        paired = sum(len(group) > 1 for group in self.groups)
        message = f"対象 {count} 件 ／ 複数ファイルのグループ {paired} 件"
        if skipped:
            message += f" ／ 読み取りをスキップ {skipped} 件"
        self.status.configure(text=message)

    def show_groups(self):
        self.group_tree.delete(*self.group_tree.get_children())
        for index, group in enumerate(self.groups):
            if len(group) > 1 or self.show_single.get():
                label = stem_without_version(group[0].path.stem) + group[0].path.suffix
                self.group_tree.insert("", "end", iid=str(index), values=(label, len(group)))
        available = self.group_tree.get_children()
        if available:
            self.group_tree.selection_set(available[0])
        self.select_group()

    def select_group(self, _event=None):
        selected = self.group_tree.selection()
        group = self.groups[int(selected[0])] if selected else []
        if group is self.group:
            return
        self.group = group
        for widget in self.file_rows.winfo_children():
            widget.destroy()
        self.checks = []
        if group and self.spans_folders():
            self.group_note.configure(text="別フォルダのファイルをまとめて比較できます。整理はフォルダ単位で行います。")
        elif group:
            self.group_note.configure(text="比較するファイル、または代表にするファイルにチェックを入れてください。")
        elif self.groups:
            self.group_note.configure(text="「単独ファイルも表示」を選ぶと各ファイルを確認できます。")
        else:
            self.group_note.configure(text="左の一覧からグループを選んでください。")
        for index, item in enumerate(group):
            checked = tk.BooleanVar(master=self, value=False)
            self.checks.append((item, checked))
            checkbox = ttk.Checkbutton(self.file_rows, text=self.file_name(item),
                                       variable=checked, command=self.update_buttons)
            checkbox.grid(row=index, column=0, sticky="ew", pady=2)
            date = datetime.fromtimestamp(item.stamp[3] / 1_000_000_000).strftime("%Y-%m-%d %H:%M")
            label = ttk.Label(self.file_rows, text=date)
            label.grid(row=index, column=1, sticky="e", padx=(12, 6))
            self.bind_file_scroll(checkbox)
            self.bind_file_scroll(label)
        self.file_canvas.xview_moveto(0)
        self.file_canvas.yview_moveto(0)
        self.update_buttons()

    def resize_file_rows(self, _event):
        width = max(self.file_canvas.winfo_width() - 2, self.file_rows.winfo_reqwidth())
        if int(float(self.file_canvas.itemcget(self.rows_window, "width"))) != width:
            self.file_canvas.itemconfigure(self.rows_window, width=width)
        self.file_canvas.configure(scrollregion=self.file_canvas.bbox("all"))

    def bind_file_scroll(self, widget):
        widget.bind("<MouseWheel>", self.scroll_file_rows)
        widget.bind("<Button-4>", self.scroll_file_rows)
        widget.bind("<Button-5>", self.scroll_file_rows)

    def scroll_file_rows(self, event):
        direction = -1 if event.num == 4 or event.delta > 0 else 1
        self.file_canvas.yview_scroll(direction, "units")
        return "break"

    def checked_files(self) -> list[File]:
        return [item for item, checked in self.checks if checked.get()]

    def file_name(self, item: File) -> str:
        return self.display_path(item.path)

    def display_path(self, path: Path) -> str:
        return str(path.relative_to(self.folder)) if self.folder and path.is_relative_to(self.folder) else str(path)

    def spans_folders(self) -> bool:
        return len({item.path.parent for item in self.group}) > 1

    def update_buttons(self):
        files = self.checked_files()
        count = len(files)
        self.keep_button.configure(state="normal" if count == 1 and not self.spans_folders() else "disabled")
        self.compare_button.configure(state="normal" if count >= 2 else "disabled")
        matlab = count >= 2 and all(item.ext == ".m" for item in files)
        self.dependency_button.configure(state="normal" if matlab else "disabled")
        self.import_button.configure(state="normal" if any(group[0].ext in CODE_TYPES for group in self.groups) else "disabled")
        message = f"{count}件にチェック" if self.group else ""
        self.selection_note.configure(text=message)
        if self.folder is None:
            instruction = "左上の「フォルダを選ぶ…」から、整理したいフォルダを選択してください。"
        else:
            instruction = (
                "「代表を選んで整理」で、残すファイル名とバックアップ先を確認します。\n"
                "「コードを比較する」で、チェックしたファイルの内容や差分を表示します。\n"
                "「参照先を比較する」で、MATLABが検出した補助関数まで比べます。\n\n"
                "１件にチェックすると、同じフォルダ内のグループで代表を指定できます。\n"
                "２件以上にチェックすると比較でき、すべて .m なら参照先も比較できます。\n"
                "左下から、別のフォルダの代表版をコピーして整理できます。"
            )
        self.set_diff(instruction, wrap="word")

    def choose_representative(self):
        item, = self.checked_files()
        prompt = "代表として残すファイル名を入力してください。"
        suggested = stem_without_version(item.path.stem) + item.path.suffix
        if item.ext in CODE_TYPES:
            suggested = item.path.name
            prompt += "\nコードの名前や配置を変えると、呼び出しに影響する場合があります。"
        name = simpledialog.askstring(TITLE, prompt, initialvalue=suggested, parent=self)
        if name is not None:
            self.preview(item, name.strip())

    def compare(self):
        files = sorted(self.checked_files(), key=lambda item: item.stamp[3])
        heading = "  ↔  ".join(self.file_name(item) for item in files) if len(files) == 2 else f"{len(files)}件の比較"
        results = [comparison_lines(item) for item in files]
        versions = [result for result in results if result is not None]
        if len(versions) != len(files):
            details = []
            for item, lines in zip(files, results):
                updated = datetime.fromtimestamp(item.stamp[3] / 1_000_000_000).strftime("%Y-%m-%d %H:%M:%S")
                body = "本文を取得済み" if lines is not None else "本文を取得できない形式、サイズ、または読み取り状態"
                details.append(f"{self.file_name(item)}\n更新日時  {updated}\nサイズ    {item.stamp[2]:,} バイト\n{body}")
            self.set_diff("全件の本文を取得できる場合に本文を比較します。\n\n" + "\n\n".join(details), heading)
        elif len(files) == 2:
            self.show_pair_diff(files, versions, heading)
        else:
            self.show_common_diff(files, versions, heading)
        if len(versions) == len(files) and all(item.ext in CODE_TYPES for item in files):
            self.set_detail_pairs([(files[0], item) for item in files[1:]])

    def choose_import_representative(self):
        destinations = []
        for group in self.groups:
            if group[0].ext in CODE_TYPES:
                by_folder: dict[Path, list[File]] = {}
                for item in group:
                    by_folder.setdefault(item.path.parent, []).append(item)
                destinations.extend(by_folder.values())
        if not destinations:
            return

        window = tk.Toplevel(self)
        window.title("代表バージョンを持ってきて整理")
        window.transient(self)
        window.grab_set()
        frame = ttk.Frame(window, padding=14)
        frame.pack(fill="both", expand=True)
        ttk.Label(frame, text="整理するグループ").pack(anchor="w")
        labels = []
        for group in destinations:
            item = group[0]
            name = stem_without_version(item.path.stem) + item.path.suffix
            location = "現在のフォルダ" if item.path.parent == self.folder else self.display_path(item.path.parent)
            labels.append(f"{name}  ［{location}］  {len(group)}件")
        destination = ttk.Combobox(frame, values=labels, state="readonly", width=65)
        destination.pack(fill="x", pady=(3, 10))
        current = next((index for index, group in enumerate(destinations) if group == self.group), None)
        if current is not None:
            destination.current(current)
        elif not self.spans_folders():
            destination.current(0)

        sources = ttk.Frame(frame)
        sources.pack(fill="x")
        ttk.Label(sources, text="代表バージョン").pack(side="left", padx=(0, 8))
        note = ttk.Label(frame, text="ファイルを指定するか、別のフォルダから候補を探します。")
        note.pack(anchor="w", pady=(8, 3))
        listing = ttk.Frame(frame)
        listing.pack(fill="both", expand=True)
        listing.columnconfigure(0, weight=1)
        listing.rowconfigure(0, weight=1)
        choices = tk.Listbox(listing, height=5, exportselection=False)
        choices.grid(row=0, column=0, sticky="nsew")
        horizontal = ttk.Scrollbar(listing, orient="horizontal", command=choices.xview)
        vertical = ttk.Scrollbar(listing, orient="vertical", command=choices.yview)
        horizontal.grid(row=1, column=0, sticky="ew")
        vertical.grid(row=0, column=1, sticky="ns")
        choices.configure(xscrollcommand=horizontal.set, yscrollcommand=vertical.set)
        candidates: list[tuple[File, str]] = []

        def show_candidates(matches, search_root=None):
            candidates[:] = matches
            choices.delete(0, "end")
            for item, reason in candidates:
                parent = item.path.parent
                location = str(parent.relative_to(search_root)) if search_root and parent.is_relative_to(search_root) else self.display_path(parent)
                choices.insert("end", f"{reason}  |  {item.path.name}  ［{location}］")
            if candidates:
                choices.selection_set(0)
            update_choices()

        def update_choices(_event=None):
            ready = destination.current() >= 0
            choose_button.configure(state="normal" if ready else "disabled")
            search_button.configure(state="normal" if ready else "disabled")
            compare_button.configure(state="normal" if ready and choices.curselection() else "disabled")

        def change_destination(_event=None):
            show_candidates([])
            note.configure(text="ファイルを指定するか、別のフォルダから候補を探します。")

        def choose_file():
            group = destinations[destination.current()]
            ext = group[0].ext
            value = filedialog.askopenfilename(parent=window, title="残す代表版を選ぶ", initialdir=self.folder,
                                               filetypes=[("コード", f"*{ext}")])
            if not value:
                return
            try:
                item = inspect_file(Path(value).resolve())
                if item.ext != ext:
                    raise ValueError(f"{ext} のファイルを選んでください。")
            except (OSError, ValueError) as error:
                messagebox.showerror(TITLE, str(error), parent=window)
                return
            show_candidates([(item, "手動で指定")])
            note.configure(text="選んだ代表版を確認して、比較に進みます。")

        def search_folder():
            value = filedialog.askdirectory(parent=window, title="代表版を探すフォルダ", initialdir=self.folder, mustexist=True)
            if not value:
                return
            search_root = Path(value).resolve()
            group = destinations[destination.current()]
            show_candidates([])
            destination.configure(state="disabled")
            choose_button.configure(state="disabled")
            search_button.configure(state="disabled")
            note.configure(text="選んだフォルダとサブフォルダを探索しています…")
            results = Queue()

            def search():
                try:
                    results.put(find_representatives(search_root, group))
                except Exception as error:
                    results.put(error)

            def poll():
                try:
                    result = results.get_nowait()
                except Empty:
                    window.after(75, poll)
                    return
                destination.configure(state="readonly")
                if isinstance(result, Exception):
                    note.configure(text="フォルダを探索できませんでした。")
                    messagebox.showerror(TITLE, str(result), parent=window)
                    update_choices()
                    return
                matches, skipped = result
                show_candidates(matches, search_root)
                caption = f"候補 {len(matches)}件。完全一致するファイルを先に表示します。" if matches else "似た候補が見つかりませんでした。"
                if skipped:
                    caption += f" 読み取りをスキップ {skipped}件。"
                note.configure(text=caption)

            threading.Thread(target=search, daemon=True).start()
            window.after(75, poll)

        def review():
            selected = choices.curselection()
            if not selected:
                return
            group = destinations[destination.current()]
            representative = candidates[selected[0]][0]
            window.destroy()
            self.review_import(group, representative)

        choose_button = ttk.Button(sources, text="ファイルを選ぶ", command=choose_file)
        choose_button.pack(side="left")
        search_button = ttk.Button(sources, text="フォルダから探す", command=search_folder)
        search_button.pack(side="left", padx=6)
        actions = ttk.Frame(frame)
        actions.pack(fill="x", pady=(8, 0))
        ttk.Button(actions, text="キャンセル", command=window.destroy).pack(side="right")
        compare_button = ttk.Button(actions, text="この代表版で比較", command=review)
        compare_button.pack(side="right", padx=6)
        destination.bind("<<ComboboxSelected>>", change_destination)
        choices.bind("<<ListboxSelect>>", update_choices)
        update_choices()

    def review_import(self, group, representative):
        window = tk.Toplevel(self)
        window.title("代表版と既存の版を比較")
        window.transient(self)
        window.grab_set()
        frame = ttk.Frame(window, padding=14)
        frame.pack(fill="both", expand=True)
        ttk.Label(frame, text=f"代表版  {self.display_path(representative.path)}", wraplength=680).pack(anchor="w")
        ttk.Label(frame, text=f"整理先  {group[0].path.parent}", wraplength=680).pack(anchor="w", pady=(3, 8))
        naming = ttk.Frame(frame)
        naming.pack(fill="x")
        ttk.Label(naming, text="残すファイル名").pack(side="left", padx=(0, 8))
        filename = tk.StringVar(window, value=representative.path.name)
        ttk.Entry(naming, textvariable=filename, width=55).pack(side="left", fill="x", expand=True)
        ttk.Label(frame, text="バックアップに移す版にチェックを付けます。").pack(anchor="w", pady=(10, 3))
        listing = ttk.Frame(frame)
        listing.pack(fill="both", expand=True)
        listing.columnconfigure(0, weight=1)
        listing.rowconfigure(0, weight=1)
        canvas = tk.Canvas(listing, width=650, height=205, highlightthickness=1, highlightbackground="#b8b8b8")
        vertical = ttk.Scrollbar(listing, orient="vertical", command=canvas.yview)
        horizontal = ttk.Scrollbar(listing, orient="horizontal", command=canvas.xview)
        canvas.configure(yscrollcommand=vertical.set, xscrollcommand=horizontal.set)
        canvas.grid(row=0, column=0, sticky="nsew")
        vertical.grid(row=0, column=1, sticky="ns")
        horizontal.grid(row=1, column=0, sticky="ew")
        rows = ttk.Frame(canvas, padding=5)
        canvas.create_window(0, 0, window=rows, anchor="nw")
        rows.bind("<Configure>", lambda _: canvas.configure(scrollregion=canvas.bbox("all")))

        checks = {}
        pairs = []
        cache = {}
        for item in group:
            if item.path == representative.path:
                continue
            variable = tk.BooleanVar(window, value=True)
            checks[item.path] = variable
            try:
                identical = (item.stamp[2] == representative.stamp[2]
                             and content_digest(item, cache) == content_digest(representative, cache))
            except OSError:
                identical = False
            rate = None
            if representative.text is not None and item.text is not None:
                pairs.append((representative, item))
                if identical:
                    summary = "完全一致"
                else:
                    rate = code_line_match_rate(representative.text, item.text, item.ext)
                    summary = f"行一致率 {rate:.1%}" if rate is not None else "比較できるコード行がありません"
            else:
                summary = "完全一致" if identical else "本文の差分を表示できない形式またはサイズ"
            ttk.Checkbutton(rows, text=f"{item.path.name}   {summary}", variable=variable).pack(anchor="w", pady=2)
            if rate is not None and rate < 0.5:
                ttk.Label(rows, text="内容の差が大きく、役割や処理が異なる可能性があります。",
                          foreground="#8c4b13").pack(anchor="w", padx=(22, 0), pady=(0, 4))

        def proceed():
            name = filename.get().strip()
            archived = {path for path, checked in checks.items() if checked.get()}
            try:
                moves, copy_to = plan_import(group, representative, archived, name)
            except (ValueError, OSError) as error:
                messagebox.showerror(TITLE, str(error), parent=window)
                return
            window.destroy()
            self.show_arrangement(group, representative, name, moves, copy_to)

        actions = ttk.Frame(frame)
        actions.pack(fill="x", pady=(8, 0))
        if pairs:
            ttk.Button(actions, text="差分の詳細", command=lambda: self.show_diff_details(pairs, window)).pack(side="left")
        ttk.Button(actions, text="キャンセル", command=window.destroy).pack(side="right")
        ttk.Button(actions, text="バックアップに移す", command=proceed).pack(side="right", padx=6)

    def set_detail_pairs(self, pairs):
        self.detail_pairs = pairs
        if pairs:
            self.detail_button.pack(anchor="e", pady=(5, 0))
        else:
            self.detail_button.pack_forget()

    def show_diff_details(self, pairs=None, parent=None):
        pairs = list(self.detail_pairs if pairs is None else pairs)
        if not pairs:
            return
        window = tk.Toplevel(self)
        window.title("差分の詳細")
        window.geometry("1000x610")
        window.minsize(700, 400)
        window.transient(parent or self)
        if parent is not None:
            window.grab_set()

            def close():
                window.destroy()
                parent.grab_set()

            window.protocol("WM_DELETE_WINDOW", close)
        frame = ttk.Frame(window, padding=12)
        frame.pack(fill="both", expand=True)
        ttk.Label(frame, text="比較対象").pack(anchor="w")
        selector = ttk.Combobox(frame, values=[self.file_name(right) for _, right in pairs], state="readonly")
        selector.pack(fill="x", pady=(2, 6))
        selector.current(0)
        ttk.Label(frame, text="変更箇所を選ぶと、その前後を左右に表示します。").pack(anchor="w")
        change_list = self.tree(frame, (("position", "変更位置", 155), ("label", "変更の目安", 520)), 5)
        panes = ttk.Panedwindow(frame, orient="horizontal")
        panes.pack(fill="both", expand=True, pady=(8, 0))
        views = []
        for color in ("#ffe6e6", "#e4f5e8"):
            panel = ttk.Frame(panes, padding=(3, 0))
            panes.add(panel, weight=1)
            label = ttk.Label(panel, wraplength=460)
            label.pack(anchor="w", pady=(0, 4))
            box = ttk.Frame(panel)
            box.pack(fill="both", expand=True)
            box.rowconfigure(0, weight=1)
            box.columnconfigure(0, weight=1)
            code = tk.Text(box, wrap="none", width=40, height=12, state="disabled", font="TkFixedFont")
            vertical = ttk.Scrollbar(box, orient="vertical", command=code.yview)
            horizontal = ttk.Scrollbar(box, orient="horizontal", command=code.xview)
            code.configure(yscrollcommand=vertical.set, xscrollcommand=horizontal.set)
            code.grid(row=0, column=0, sticky="nsew")
            vertical.grid(row=0, column=1, sticky="ns")
            horizontal.grid(row=1, column=0, sticky="ew")
            code.tag_configure("changed", background=color)
            views.append((label, code))

        changes = []
        source = ([], [])

        def render(_event=None):
            selected = change_list.selection()
            if changes and not selected:
                return
            change = changes[int(selected[0])] if changes else None
            spans = ((change.left_start, change.left_end, change.left_changed),
                     (change.right_start, change.right_end, change.right_changed)) if change else ()
            for index, (_, code) in enumerate(views):
                code.configure(state="normal")
                code.delete("1.0", "end")
                if change:
                    start, end, changed = spans[index]
                    for row in range(max(start - 3, 0), min(end + 3, len(source[index]))):
                        code.insert("end", f"{row + 1:>5} | {source[index][row]}\n", "changed" if row in changed else "")
                    if not source[index]:
                        code.insert("end", "本文は空です。")
                else:
                    code.insert("end", "表示上の行に差分はありません。")
                code.xview_moveto(0)
                code.yview_moveto(0)
                code.configure(state="disabled")

        def load(_event=None):
            nonlocal changes, source
            left, right = pairs[selector.current()]
            for (label, _), item, role in zip(views, (left, right), ("基準", "比較対象")):
                label.configure(text=f"{role}  {self.file_name(item)}")
            source = (left.text.splitlines(), right.text.splitlines())
            changes = change_blocks(*source)
            change_list.delete(*change_list.get_children())
            for index, change in enumerate(changes):
                position = f"左 {change.left_start + 1}行 / 右 {change.right_start + 1}行"
                change_list.insert("", "end", iid=str(index), values=(position, change.summary))
            if changes:
                change_list.selection_set("0")
            render()

        change_list.bind("<<TreeviewSelect>>", render)
        selector.bind("<<ComboboxSelected>>", load)
        load()

    def compare_dependencies(self):
        files = self.checked_files()
        if len(files) < 2 or any(item.ext != ".m" for item in files):
            return
        self.dependency_settings([item.path for item in files])

    def dependency_settings(self, files):
        requests = [AnalysisRequest(path, path.parent) for path in files]
        window = tk.Toplevel(self)
        window.title("参照先を比較する")
        window.transient(self)
        window.grab_set()
        frame = ttk.Frame(window, padding=14)
        frame.pack(fill="both", expand=True)
        names = [self.display_path(path) for path in files]
        ttk.Label(frame, text=f"比較するファイル（{len(names)}件）").pack(anchor="w")
        listing = ttk.Frame(frame)
        listing.pack(fill="x", pady=(4, 8))
        listing.columnconfigure(0, weight=1)
        file_list = tk.Text(listing, width=56, height=min(len(names), 6), wrap="none",
                            font="TkTextFont", takefocus=False)
        file_list.insert("1.0", "\n".join(names))
        file_list.configure(state="disabled")
        file_list.grid(row=0, column=0, sticky="ew")
        horizontal = ttk.Scrollbar(listing, orient="horizontal", command=file_list.xview)
        vertical = ttk.Scrollbar(listing, orient="vertical", command=file_list.yview)
        horizontal.grid(row=1, column=0, sticky="ew")
        if len(names) > 6:
            vertical.grid(row=0, column=1, sticky="ns")
        file_list.configure(xscrollcommand=horizontal.set, yscrollcommand=vertical.set)
        ttk.Label(frame, text="現在のフォルダ内だけなら、そのまま開始。").pack(anchor="w")

        program = self.matlab_path
        if program is None or not program.is_file():
            program = find_matlab()
        missing = ttk.Frame(frame)
        ttk.Label(missing, text="MATLABが見つかりませんでした。").pack(side="left")

        def browse_matlab():
            nonlocal program
            value = filedialog.askopenfilename(parent=window, title="MATLABの実行ファイルを選ぶ",
                                               filetypes=[("MATLAB", "matlab.exe"), ("すべてのファイル", "*")])
            if value:
                program = Path(value)
                update_matlab_choice()

        ttk.Button(missing, text="MATLABを選ぶ…", command=browse_matlab).pack(side="left", padx=(6, 0))

        details = ttk.Frame(frame)
        details.columnconfigure(1, weight=1)
        ttk.Label(details, text="設定するファイル").grid(row=0, column=0, sticky="w", padx=(0, 8))
        selector = ttk.Combobox(details, values=names, state="readonly")
        selector.grid(row=0, column=1, columnspan=2, sticky="ew")
        selector.current(0)
        workdir = tk.StringVar(window)
        ttk.Label(details, text="作業フォルダ").grid(row=1, column=0, sticky="w", pady=(5, 0))
        ttk.Entry(details, textvariable=workdir, state="readonly").grid(row=1, column=1, sticky="ew", pady=(5, 0))
        paths = tk.Listbox(details, height=4, exportselection=False)
        paths.grid(row=2, column=1, sticky="ew", pady=(5, 0))
        ttk.Label(details, text="追加の参照フォルダ").grid(row=2, column=0, sticky="nw", padx=(0, 8), pady=(5, 0))

        def load_settings(_event=None):
            item = requests[selector.current()]
            workdir.set(str(item.workdir))
            paths.delete(0, "end")
            for path in item.search_dirs:
                paths.insert("end", str(path))

        def change_folder():
            current = selector.current()
            item = requests[current]
            value = filedialog.askdirectory(parent=window, initialdir=item.workdir, mustexist=True)
            if value:
                requests[current] = AnalysisRequest(item.source, Path(value), item.search_dirs)
                load_settings()

        def edit_paths(remove=False):
            current = selector.current()
            item = requests[current]
            directories = list(item.search_dirs)
            if remove:
                selected = paths.curselection()
                if not selected:
                    return
                del directories[selected[0]]
            else:
                value = filedialog.askdirectory(parent=window, initialdir=item.workdir, mustexist=True)
                if not value or Path(value) in directories:
                    return
                directories.append(Path(value))
            requests[current] = AnalysisRequest(item.source, item.workdir, tuple(directories))
            load_settings()

        ttk.Button(details, text="変更…", command=change_folder).grid(row=1, column=2, padx=(6, 0), pady=(5, 0))
        path_actions = ttk.Frame(details)
        path_actions.grid(row=2, column=2, sticky="n", padx=(6, 0), pady=(5, 0))
        ttk.Button(path_actions, text="追加…", command=edit_paths).pack(fill="x")
        ttk.Button(path_actions, text="外す", command=lambda: edit_paths(True)).pack(fill="x", pady=(4, 0))
        selector.bind("<<ComboboxSelected>>", load_settings)
        load_settings()

        actions = ttk.Frame(frame)
        actions.pack(fill="x", pady=(10, 0))
        ttk.Button(actions, text="キャンセル", command=window.destroy).pack(side="right")

        def toggle_details():
            if details.winfo_manager():
                details.pack_forget()
                detail_button.configure(text="詳細設定")
            else:
                details.pack(fill="x", before=actions, pady=(10, 0))
                detail_button.configure(text="詳細設定を閉じる")

        detail_button = ttk.Button(actions, text="詳細設定", command=toggle_details)
        detail_button.pack(side="left")

        def update_matlab_choice():
            available = program is not None and program.is_file()
            start_button.configure(state="normal" if available else "disabled")
            if available:
                missing.pack_forget()
            else:
                missing.pack(fill="x", before=details if details.winfo_manager() else actions, pady=(8, 0))

        def start():
            if program is None or not program.is_file():
                update_matlab_choice()
                return
            self.matlab_path = program.resolve()
            window.destroy()
            self.run_dependency_comparison(self.matlab_path, requests)

        start_button = ttk.Button(actions, text="比較を開始", command=start)
        start_button.pack(side="right", padx=6)
        update_matlab_choice()

    def run_dependency_comparison(self, executable, requests):
        window = tk.Toplevel(self)
        window.title("MATLABで参照先を解析中")
        window.transient(self)
        window.grab_set()
        window.protocol("WM_DELETE_WINDOW", lambda: None)
        frame = ttk.Frame(window, padding=14)
        frame.pack(fill="both", expand=True)
        note = ttk.Label(frame, text="MATLABを起動しています…", wraplength=500)
        note.pack(anchor="w")
        progress = ttk.Progressbar(frame, maximum=len(requests), length=420)
        progress.pack(fill="x", pady=(8, 0))
        ttk.Label(frame, text="MATLABで解析中。完了まで数分かかることがあります。").pack(anchor="w", pady=(6, 0))

        events = Queue()

        def run():
            entries = []
            try:
                for index, request in enumerate(requests):
                    events.put((index, f"{index + 1}/{len(requests)}  {request.source}"))
                    try:
                        entry = analyze_matlab(executable, request)
                    except (OSError, RuntimeError, ValueError, KeyError, TypeError) as error:
                        entry = Analysis(request, {}, (), str(error))
                    entries.append(entry)
                report = dependency_report(entries)
            except Exception as error:
                report = f"比較結果を作成できませんでした。\n{error}"
            events.put((None, report))

        def poll_dependencies():
            while True:
                try:
                    index, value = events.get_nowait()
                except Empty:
                    window.after(100, poll_dependencies)
                    return
                if index is None:
                    window.destroy()
                    self.set_diff(value, "MATLABの参照先の比較")
                    return
                progress.configure(value=index)
                note.configure(text=value)

        threading.Thread(target=run, daemon=True).start()
        window.after(100, poll_dependencies)

    def set_diff(self, message, heading="比較結果", color=False, *, wrap="none"):
        self.set_detail_pairs([])
        self.diff_label.configure(text=heading)
        self.diff.configure(state="normal", wrap=wrap)
        self.diff.delete("1.0", "end")
        if color:
            for line in message.splitlines(keepends=True):
                tag = "added" if line.startswith("+") else "removed" if line.startswith("-") else "hunk" if line.startswith("@@") else ""
                self.diff.insert("end", line, tag)
        else:
            self.diff.insert("end", message)
        self.diff.xview_moveto(0)
        self.diff.yview_moveto(0)
        self.diff.configure(state="disabled")

    def show_pair_diff(self, files, versions, heading):
        left, right = files
        if left.digest is not None and left.digest == right.digest:
            self.set_diff("ファイルの中身は完全に一致しています。", heading)
            return
        score = similarity(left, right)
        if score is not None:
            heading += f"   行パターンの類似度 {score:.0%}"
        diff = list(difflib.unified_diff(*versions, self.file_name(left), self.file_name(right), lineterm=""))
        if len(diff) > 1500:
            diff = diff[:1500] + ["", "表示は先頭1500行までです。"]
        self.set_diff("\n".join(diff) or "比較した本文に差分はありません。", heading, color=True)

    def show_common_diff(self, files, versions, heading):
        matches = common_positions(versions)
        common = [f"{index + 1:>5} | {versions[0][index]}" for index in sorted(matches[0])]
        report = [f"比較対象 {len(files)}件", *(f"  {self.file_name(item)}" for item in files),
                  "", f"=== 全ファイルに共通する部分（{len(common)}行） ===",
                  f"行番号は {self.file_name(files[0])} を基準に表示します。", ""]
        report.extend(common or ["共通として対応付けられた行はありません。"])
        report.extend(["", "=== 全ファイルには共通しない部分 ===",
                       "一部のファイルだけに共通する行も、ここに表示します。"])
        for item, lines, shared in zip(files, versions, matches):
            different = [f"{index + 1:>5} | {line}" for index, line in enumerate(lines) if index not in shared]
            report.extend(["", f"--- {self.file_name(item)}（{len(different)}行） ---"])
            report.extend(different or ["該当する行はありません。"])
        self.set_diff("\n".join(report), heading)

    def preview(self, representative, name):
        try:
            moves = plan_moves(self.group, {representative.path: name})
        except (ValueError, OSError) as error:
            messagebox.showerror(TITLE, str(error), parent=self)
            return
        self.show_arrangement(self.group, representative, name, moves)

    def show_arrangement(self, group, representative, name, moves, copy_to=None):
        if not moves and copy_to is None:
            messagebox.showinfo(TITLE, "このグループに反映する変更はありません。", parent=self)
            return
        folder = group[0].path.parent
        window = tk.Toplevel(self)
        window.title("整理内容を確認")
        window.geometry("720x420")
        window.transient(self)
        window.grab_set()
        frame = ttk.Frame(window, padding=14)
        frame.pack(fill="both", expand=True)
        ttk.Label(frame, text="このグループの変更内容").pack(anchor="w", pady=(0, 8))
        if representative.ext in CODE_TYPES:
            ttk.Label(frame, text="コードを呼び出す他のファイルがある場合は、名前と移動先を確認してください。").pack(anchor="w", pady=(0, 5))
        lines = tk.Text(frame, wrap="char", height=14, font="TkTextFont")
        bar = ttk.Scrollbar(frame, orient="vertical", command=lines.yview)
        lines.configure(yscrollcommand=bar.set)
        actions = ttk.Frame(frame)
        actions.pack(side="bottom", fill="x", pady=(8, 0))
        bar.pack(side="right", fill="y")
        lines.pack(side="left", fill="both", expand=True)
        if copy_to is not None:
            lines.insert("end", f"代表版をコピーして残す\n  コピー元 {representative.path}\n"
                                f"  ↳ コピー先フォルダ {copy_to.parent}\n    保存名 {copy_to.name}\n")
        else:
            lines.insert("end", f"代表として残す\n  {representative.path.name}  →  {name}\n")
        lines.insert("end", "\nバックアップへ移動\n")
        archived = [move for move in moves if move.target.parent != folder]
        if archived:
            for move in archived:
                lines.insert("end", f"  {move.file.path.name}\n  ↳ {move.target.relative_to(folder)}\n")
        else:
            lines.insert("end", "  0件\n")
        lines.configure(state="disabled")
        ttk.Button(actions, text="キャンセル", command=window.destroy).pack(side="right")
        copy = (representative, copy_to) if copy_to is not None else None
        ttk.Button(actions, text="この内容で整理する",
                   command=lambda: self.apply(window, group, moves, copy)).pack(side="right", padx=7)

    def apply(self, window, group, moves, copy=None):
        try:
            execute(group, moves, copy)
        except OSError as error:
            messagebox.showerror(TITLE, f"整理を完了できませんでした。\n{error}", parent=window)
        else:
            messagebox.showinfo(TITLE, "選んだグループを整理しました。", parent=window)
        window.destroy()
        self.refresh()


if __name__ == "__main__":
    Organizer().mainloop()
