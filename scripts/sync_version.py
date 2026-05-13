#!/usr/bin/env python
"""pre-commit hook: 将 pyproject.toml 中的版本号同步到 src/wordformat/_version.py。

版本号的唯一来源是 pyproject.toml。每次提交时，此 hook 会确保
_version.py 与 pyproject.toml 保持一致。
"""

import re
import sys
from pathlib import Path

VERSION_FILE = Path("src/wordformat/_version.py")
PYPROJECT_FILE = Path("pyproject.toml")
TEMPLATE = '# 此文件由 scripts/sync_version.py 自动生成，请勿手动编辑。\n# 版本号唯一来源为 pyproject.toml。\n__version__ = "{version}"\n'


def read_pyproject_version() -> str | None:
    """从 pyproject.toml 读取版本号。"""
    if not PYPROJECT_FILE.exists():
        print(f"sync_version: {PYPROJECT_FILE} 不存在，跳过", file=sys.stderr)
        return None
    content = PYPROJECT_FILE.read_text(encoding="utf-8")
    match = re.search(
        r'^version\s*=\s*["\']([^"\']+)["\']', content, re.MULTILINE
    )
    if match:
        return match.group(1)
    print("sync_version: 无法从 pyproject.toml 解析版本号", file=sys.stderr)
    return None


def read_version_file() -> str | None:
    """从 _version.py 读取当前版本号。"""
    if not VERSION_FILE.exists():
        return None
    content = VERSION_FILE.read_text(encoding="utf-8")
    match = re.search(r'__version__\s*=\s*["\']([^"\']+)["\']', content)
    if match:
        return match.group(1)
    return None


def main() -> int:
    target = read_pyproject_version()
    if target is None:
        return 0

    current = read_version_file()

    if current == target:
        print(f"sync_version: 版本号一致 ({target})，无需同步")
        return 0

    VERSION_FILE.parent.mkdir(parents=True, exist_ok=True)
    VERSION_FILE.write_text(TEMPLATE.format(version=target), encoding="utf-8")
    print(f"sync_version: 版本号已同步: {current or '(新文件)'} -> {target}")
    print(f"sync_version: 请将 {VERSION_FILE} 重新暂存后再次提交", file=sys.stderr)
    return 1


if __name__ == "__main__":
    sys.exit(main())
