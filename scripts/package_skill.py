#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""package_skill.py — 把本仓库打包成可安装的 `.skill` 文件（纯标准库）。

生成 `dist/dut-thesis-formatter.skill`。GitHub Actions 发布流程也调用它：

    python scripts/package_skill.py
"""
from __future__ import annotations

import zipfile
from pathlib import Path

# 显式白名单——把 CI 配置、缓存、构建产物挡在包外。
# 注意只收 docs/demo.png：showcase.docx / showcase.pdf 是构建产物（几百 KB），
# 会白白撑大技能包，用户不需要。
INCLUDE = [
    'SKILL.md',
    'install-prompt.txt',
    'README.md',
    'LICENSE',
    'references/dut-format-rules.md',
    'scripts/dut_thesis.py',
    'scripts/cli.py',
    'scripts/make_demo.py',
    'scripts/make_showcase.py',
    'scripts/make_demo_real.py',
    'scripts/convert_with_word.sh',
    'scripts/tests/test_smoke.py',
    'docs/demo.png',
    'docs/INSTALL.md',
    'docs/RENDERING.md',
]


def collect(repo: Path) -> list[Path]:
    files: list[Path] = []
    for rel in INCLUDE:
        p = repo / rel
        if p.is_dir():
            for f in sorted(p.rglob('*')):
                if f.is_file() and '__pycache__' not in f.parts and f.suffix != '.pyc':
                    files.append(f)
        elif p.is_file():
            files.append(p)
        else:
            raise SystemExit(f'缺少预期路径: {rel}')
    return files


def main() -> int:
    repo = Path(__file__).resolve().parent.parent
    # 名字必须与 SKILL.md frontmatter 的 `name` 一致，而不是随目录名走——
    # 否则用户把仓库 clone 成别的目录名时，安装后的技能名会对不上。
    name = 'dut-thesis-formatter'
    out_dir = repo / 'dist'
    out_dir.mkdir(exist_ok=True)
    out = out_dir / f'{name}.skill'

    files = collect(repo)
    with zipfile.ZipFile(out, 'w', zipfile.ZIP_DEFLATED) as z:
        for f in files:
            arc = f'{name}/{f.relative_to(repo).as_posix()}'
            z.write(f, arc)
            print(f'  Added: {arc}')

    print(f'Packaged: {out}')
    return 0


if __name__ == '__main__':
    raise SystemExit(main())
