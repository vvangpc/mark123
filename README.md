# 专利附图标记助手 (Patent Marker)

专利附图标记助手是一款使用 PyQt6 和 python-docx 构建的桌面应用程序。主要用于辅助专利撰写和审查人员处理和分析专利申请文件（特别是 `.docx` 格式）中的附图标记和说明书内容。

## 主要特性 ✨

- **专利文档解析**：智能解析 `.docx` 格式的权利要求书和说明书文档。
- **附图标记高亮**：自动识别并高亮显示文档中的附图标记，提升阅读和核对效率。
- **附图标记抽取**：一键提取说明书和权利要求中的所有附图标记，方便生成附图标记列表。
- **深色主题界面**：现代化的深色模式设计，减少长时间阅读的视觉疲劳。
- **高DPI支持**：自动适配高分辨率屏幕，显示清晰亮丽。

## 技术栈 🛠️

- **Python**: >= 3.14
- **GUI 框架**: [PyQt6](https://pypi.org/project/PyQt6/)
- **文档处理**: [python-docx](https://pypi.org/project/python-docx/)
- **依赖管理**: [uv](https://github.com/astral-sh/uv)
- **程序打包**: [PyInstaller](https://pyinstaller.org/)

## 项目结构 📁

```text
d:/AI/mark123/
├── README.md                 # 项目说明文档
├── pyproject.toml            # uv 项目配置及依赖清单
├── uv.lock                   # uv 依赖锁文件
├── main.py                   # 应用程序入口
├── main_window.py            # 主窗口 UI 及交互逻辑
├── doc_parser.py             # Word 文档解析模块
├── annotator.py              # 内容高亮与标注模块
├── mark_extractor.py         # 附图标记提取模块
├── styles.py                 # 深色主题 QSS 样式
└── 专利附图标记助手.spec     # PyInstaller 打包配置
```

## 安装与运行 🚀

本项目使用 `uv` 进行现代化的 Python 环境和依赖管理。

1. **安装包管理器 uv (如果还未安装)**
   可以参考 [uv 官方文档](https://github.com/astral-sh/uv)。
   
2. **克隆/进入项目目录**
   ```bash
   cd d:/AI/mark123
   ```

3. **同步并安装依赖**
   使用 uv 安装 `pyproject.toml` 中的所有依赖项：
   ```bash
   uv sync
   ```
   *如果部分依赖未能成功安装，可以使用 `uv pip install <package_name>` 补充安装，例如 `uv pip install pyqt6 python-docx pyinstaller`。*

4. **运行应用程序**
   ```bash
   uv run python main.py
   ```
   *(或者拖拽一个标准的 `.docx` 专利文档到可执行文件上打开)*

## 打包为可执行文件 📦

如果您想要将项目打包为独立的 Windows `.exe` 程序，可以使用 PyInstaller：

```bash
# 使用提供的 spec 文件进行打包
uv run pyinstaller "专利附图标记助手.spec"
```
打包成功后，可执行文件会被生成在 `dist` 目录下。

## 许可证 📄

本项目仅供学习和内部参考使用。
