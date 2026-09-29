"""Check the staged public snapshot before committing or pushing it."""

from __future__ import annotations

import subprocess
import sys
from io import BytesIO
from pathlib import Path

from openpyxl import load_workbook


ROOT = Path(__file__).resolve().parents[1]
FORBIDDEN_PARTS = {
    "config-real", "2026-07-31", "test-output", "reference", "dist",
    "runtime", "wheels", "build", "data", "__pycache__",
}


def main() -> int:
    raw = subprocess.check_output(["git", "diff", "--cached", "--name-only", "-z"], cwd=ROOT)
    paths = [Path(p.decode("utf-8")) for p in raw.split(b"\0") if p]
    if not paths:
        print("没有暂存文件。")
        return 1
    errors: list[str] = []
    books = 0
    for path in paths:
        if any(part.lower() in FORBIDDEN_PARTS for part in path.parts):
            errors.append(f"禁止公开的目录：{path}")
        if path.suffix.lower() in {".csv", ".xls", ".xlsm", ".zip", ".deb", ".exe", ".dll"}:
            errors.append(f"禁止公开的文件类型：{path}")
        if path.suffix.lower() != ".xlsx":
            continue
        if not path.parts or path.parts[0] != "config":
            errors.append(f"工作簿只能位于 config/：{path}")
            continue
        books += 1
        staged = subprocess.check_output(["git", "show", f":{path.as_posix()}"], cwd=ROOT)
        book = load_workbook(BytesIO(staged), read_only=True, data_only=False)
        try:
            for sheet in book:
                for row in sheet.iter_rows(min_row=2):
                    if any(cell.value not in (None, "") for cell in row):
                        errors.append(f"发现表头以下内容：{path} / {sheet.title}")
                        break
        finally:
            book.close()
    if errors:
        print("\n".join(errors))
        return 1
    print(f"公开文件检查通过：{len(paths)} 个暂存文件，{books} 个空表头工作簿。")
    return 0


if __name__ == "__main__":
    sys.exit(main())
