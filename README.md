# Word 转 PDF 批量转换工具

基于 Python + Tkinter 的 Windows 桌面工具，支持拖拽批量将 Word 文档（`.doc` / `.docx`）转换为 PDF。

## 功能

- 拖拽或选择文件/文件夹批量转换
- 多线程并发处理，支持大文件
- 转换进度与日志显示
- 打包为独立 exe，无需安装 Python

## 快速使用（exe）

进入 `w_to_PDF` 目录，双击运行 `w_to_p.exe` 即可。

> 需要本机已安装 Microsoft Word。

## 从源码运行

```bash
pip install -r requirements.txt
python w_to_pdf.py
```

## 打包 exe

```bash
python setup.py build
```

打包输出目录为 `build/W_TO_PDF/`，可将整个文件夹复制分发。

## 项目结构

| 文件/目录 | 说明 |
|-----------|------|
| `w_to_pdf.py` | 主程序源码 |
| `setup.py` | cx_Freeze 打包配置 |
| `set.py` | 备用打包配置 |
| `app.ico` | 应用图标 |
| `tkdnd2.9/` | 拖拽功能依赖库 |
| `w_to_PDF/` | 已打包的可执行程序 |

## 环境要求

- Windows 10/11
- Microsoft Word
- Python 3.10+（仅源码运行/打包时需要）
