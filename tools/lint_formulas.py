#!/usr/bin/env python3
"""Simple formula text linter for Excel formula modules.

Checks balanced:
- parentheses: ()
- square brackets: []
- curly braces: {}
- double quotes, including Excel doubled-quote escapes

Excel-style block comments `/* ... */` are ignored.
"""

from __future__ import annotations

import argparse
import glob
import re
import sys
from pathlib import Path


def strip_block_comments(text: str) -> str:
    return re.sub(r"/\*.*?\*/", "", text, flags=re.S)


def lint_file(path: Path) -> tuple[bool, str]:
    try:
        text = path.read_text(encoding="utf-8")
    except FileNotFoundError:
        return False, f"{path}: missing file"

    clean = strip_block_comments(text)
    pairs = {"(": ")", "[": "]", "{": "}"}
    opens = set(pairs)
    closes = {v: k for k, v in pairs.items()}
    stack: list[tuple[str, int, int]] = []
    in_string = False
    line = 1
    col = 0
    i = 0

    while i < len(clean):
        ch = clean[i]
        col += 1

        if ch == "\n":
            line += 1
            col = 0
            i += 1
            continue

        if ch == '"':
            if in_string and i + 1 < len(clean) and clean[i + 1] == '"':
                i += 2
                col += 1
                continue
            in_string = not in_string
            i += 1
            continue

        if not in_string:
            if ch in opens:
                stack.append((ch, line, col))
            elif ch in closes:
                if not stack or stack[-1][0] != closes[ch]:
                    return False, f"{path}: unexpected {ch!r} at line {line}, col {col}"
                stack.pop()

        i += 1

    if in_string:
        return False, f"{path}: unterminated double-quoted string"

    if stack:
        ch, line, col = stack[-1]
        return False, f"{path}: unclosed {ch!r} from line {line}, col {col}"

    return True, f"{path}: PASS"


def expand_inputs(inputs: list[str]) -> list[Path]:
    files: list[Path] = []
    for item in inputs:
        if any(ch in item for ch in "*?["):
            matches = sorted(glob.glob(item))
            files.extend(Path(match) for match in matches)
            if not matches:
                files.append(Path(item))
        else:
            files.append(Path(item))
    return files


def main() -> int:
    parser = argparse.ArgumentParser(description="Lint Excel formula-module text files.")
    parser.add_argument("files", nargs="+", help="Formula text files or glob patterns to lint.")
    args = parser.parse_args()

    ok_all = True
    for path in expand_inputs(args.files):
        ok, message = lint_file(path)
        ok_all = ok_all and ok
        print(message)

    return 0 if ok_all else 1


if __name__ == "__main__":
    raise SystemExit(main())
