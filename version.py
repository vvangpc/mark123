# -*- coding: utf-8 -*-
"""
version.py — 唯一的版本号来源（Single Source of Truth）

所有地方读版本号都走这里：
    main.py        →  设置 QApplication 版本 + 更新检查时和远端 latest.json 比较
    installer.iss  →  通过 `iscc /DMyAppVersion=x.y` 命令行覆盖；本地直跑 iscc 用本文件默认值
    GitHub Actions →  从 git tag (v3.6) 提取版本号后改写本文件的 __version__，再触发构建

发版流程：
    git tag v3.6
    git push --tags
    → CI 把这里改成 "3.6"，跑 build.py + iscc，产物上传 GitHub Releases + VPS
"""

__version__ = "4.1.1"
