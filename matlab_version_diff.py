"""コードの比較用に行を整え、差分のまとまりを求める。"""

from collections import Counter
from dataclasses import dataclass
from difflib import SequenceMatcher
from io import StringIO
import re
import tokenize

ASSIGNMENT = re.compile(r"^\s*([A-Za-z]\w*)\s*=(?!=)\s*(.*)$")
CONTROL = re.compile(r"^\s*(for|parfor|if|elseif|while|switch|case|break|return)\b")
MATLAB_MARKERS = re.compile(r"""['"%]|\.\.\.""")
MATLAB_STRINGS = {
    "'": re.compile(r"'(?:''|[^'])*'"),
    '"': re.compile(r'"(?:""|[^"])*"'),
}


def matlab_line(line: str) -> tuple[str, bool]:
    position = 0
    while marker := MATLAB_MARKERS.search(line, position):
        symbol = marker[0]
        start = marker.start()
        if symbol == "%":
            return line[:start].rstrip(), False
        if symbol == "...":
            return line[:start].rstrip(), True
        position = marker.end()
        previous = line[start - 1] if start else ""
        # An apostrophe attached to an operand can be a transpose operator.
        if symbol == "'" and previous and (previous.isalnum() or previous in "_)]}.'"):
            continue
        literal = MATLAB_STRINGS[symbol].match(line, start)
        if literal is None:
            break
        position = literal.end()
    return line, False


def matlab_code_lines(text: str) -> tuple[str, ...]:
    lines = []
    continued_parts = []
    block_depth = 0
    for raw in text.splitlines():
        line = raw.strip()
        if line == "%{":
            block_depth += 1
        elif line == "%}" and block_depth:
            block_depth -= 1
        elif line and not block_depth:
            code, continues = matlab_line(line)
            if code:
                continued_parts.append(code)
                if not continues:
                    lines.append(" ".join(continued_parts))
                    continued_parts.clear()
    if continued_parts:
        lines.append(" ".join(continued_parts))
    return tuple(lines)


def python_code_lines(text: str) -> tuple[str, ...]:
    lines = []
    tokens = []
    depth = 0
    try:
        for token in tokenize.generate_tokens(StringIO(text, newline=None).readline):
            if token.type == tokenize.INDENT:
                depth += 1
            elif token.type == tokenize.DEDENT:
                depth -= 1
            elif token.type == tokenize.NEWLINE:
                if tokens:
                    lines.append(repr((depth, tokens)))
                    tokens.clear()
            elif token.type not in {tokenize.COMMENT, tokenize.NL, tokenize.ENDMARKER}:
                tokens.append(token.string)
    except (tokenize.TokenError, SyntaxError):
        return ()
    return tuple(lines)


def code_line_match_rate(first: tuple[str, ...], second: tuple[str, ...],
                         minimum: float = 0.0) -> float | None:
    if not first and not second:
        return None
    first, second = sorted((first, second))
    matcher = SequenceMatcher(None, first, second, autojunk=max(len(first), len(second)) > 2000)
    if minimum and matcher.quick_ratio() < minimum:
        return 0.0
    return matcher.ratio()


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


