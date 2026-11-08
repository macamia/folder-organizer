"""MATLABファイルの行差分を、共通する処理を目印にまとめる。"""

from collections import Counter
from dataclasses import dataclass
from difflib import SequenceMatcher
import re

ASSIGNMENT = re.compile(r"^\s*([A-Za-z]\w*)\s*=(?!=)\s*(.*)$")
CONTROL = re.compile(r"^\s*(for|parfor|if|elseif|while|switch|case|break|return)\b")


def code_line_match_rate(left: str, right: str, extension: str) -> float | None:
    def code_lines(text):
        lines = []
        in_comment = False
        prefix = "#" if extension == ".py" else "%"
        for raw in text.splitlines():
            line = raw.strip()
            if extension == ".m" and line == "%{":
                in_comment = True
            elif extension == ".m" and line == "%}":
                in_comment = False
            elif not in_comment and line and not line.startswith(prefix):
                lines.append(line)
        return lines

    first, second = code_lines(left), code_lines(right)
    if not first and not second:
        return None
    return SequenceMatcher(None, first, second, autojunk=max(len(first), len(second)) > 2000).ratio()


@dataclass(frozen=True)
class Change:
    left_start: int
    left_end: int
    right_start: int
    right_end: int
    left_changed: frozenset[int]
    right_changed: frozenset[int]
    summary: str


def change_summary(old: list[str], new: list[str]) -> str:
    def assignments(lines):
        found = {}
        for line in lines:
            if match := ASSIGNMENT.match(line):
                found.setdefault(match[1], []).append(match[2].strip())
        return found

    if all(not line.strip() or line.lstrip().startswith("%") for line in old + new):
        return "コメント行・空行の差分"

    details = []
    before, after = assignments(old), assignments(new)
    for names, label in ((after.keys() - before.keys(), "追加"),
                         (before.keys() - after.keys(), "削除"),
                         ({name for name in before.keys() & after.keys() if before[name] != after[name]}, "変更")):
        if names:
            details.append(f"代入{label} {', '.join(sorted(names))}")

    def controls(lines):
        return Counter(match[1] for line in lines if (match := CONTROL.match(line)))

    before_controls, after_controls = controls(old), controls(new)
    for keyword in sorted(before_controls.keys() | after_controls.keys()):
        difference = after_controls[keyword] - before_controls[keyword]
        if difference:
            details.append(f"{keyword} {'追加' if difference > 0 else '削除'} {abs(difference)}件")
        elif [line.strip() for line in old if (m := CONTROL.match(line)) and m[1] == keyword] != [
                line.strip() for line in new if (m := CONTROL.match(line)) and m[1] == keyword]:
            details.append(f"{keyword} の記述に変更")
    return " / ".join(details) or "コード行に差分"


def change_blocks(left: list[str], right: list[str]) -> list[Change]:
    matcher = SequenceMatcher(None, left, right, autojunk=max(len(left), len(right)) > 2000)
    sections = []
    pending = []
    for opcode in matcher.get_opcodes():
        tag, i, j, _, _ = opcode
        if tag != "equal":
            pending.append(opcode)
        elif pending and sum(bool(line.strip() and not line.lstrip().startswith("%")
                                  and line.strip() != "end") for line in left[i:j]) >= 2:
            sections.append(pending)
            pending = []
    if pending:
        sections.append(pending)

    changes = []
    for section in sections:
        left_changed = frozenset(n for _, start, end, _, _ in section for n in range(start, end))
        right_changed = frozenset(n for _, _, _, start, end in section for n in range(start, end))
        old = [left[n] for n in sorted(left_changed)]
        new = [right[n] for n in sorted(right_changed)]
        changes.append(Change(
            section[0][1], section[-1][2], section[0][3], section[-1][4],
            left_changed, right_changed, change_summary(old, new),
        ))
    return changes


