from cx_Freeze import setup, Executable
import os
import sys

# 设置默认编码
if sys.version_info[0] >= 3:
    import codecs
    sys.stdout = codecs.getwriter('utf-8')(sys.stdout.buffer)

base = None
if os.name == 'nt':
    base = "Win32GUI"

# 获取当前脚本目录
SCRIPT_DIR = os.path.dirname(os.path.abspath(__file__))
PYTHON_INSTALL_DIR = os.path.dirname(sys.executable)

# 初始化包含文件列表
include_files = []

# 1. 包含图标文件
if os.path.exists(os.path.join(SCRIPT_DIR, 'app.ico')):
    include_files.append(('app.ico', 'app.ico'))

# 2. 包含 tkdnd2.9 文件夹（从当前目录）
# 需要递归包含所有文件，确保 tkdnd 库能被正确加载
# 注意：cx_Freeze 会将文件复制到 build_exe 目录的根目录
tkdnd_path = os.path.join(SCRIPT_DIR, 'tkdnd2.9')
if os.path.exists(tkdnd_path):
    # 递归添加 tkdnd2.9 文件夹中的所有文件
    # 目标路径保持为 tkdnd2.9/文件名，确保在打包后与 exe 同级
    for root, dirs, files in os.walk(tkdnd_path):
        for file in files:
            src_path = os.path.join(root, file)
            # 计算相对于 tkdnd_path 的路径
            rel_dir = os.path.relpath(root, tkdnd_path)
            if rel_dir == '.':
                dst_path = os.path.join('tkdnd2.9', file)
            else:
                dst_path = os.path.join('tkdnd2.9', rel_dir, file)
            include_files.append((src_path, dst_path))
            print(f"包含文件: {src_path} -> {dst_path}")

# 3. 包含 tkinterdnd2（从当前目录或 site-packages）
tkinterdnd2_local = os.path.join(SCRIPT_DIR, 'tkinterdnd2')
tkinterdnd2_site = os.path.join(PYTHON_INSTALL_DIR, 'Lib', 'site-packages', 'tkinterdnd2')
if os.path.exists(tkinterdnd2_local):
    include_files.append((tkinterdnd2_local, 'tkinterdnd2'))
elif os.path.exists(tkinterdnd2_site):
    include_files.append((tkinterdnd2_site, 'tkinterdnd2'))

# 4. 包含 tcl/tk DLL（如果存在）
tcl_dll = os.path.join(PYTHON_INSTALL_DIR, 'DLLs', 'tcl86t.dll')
tk_dll = os.path.join(PYTHON_INSTALL_DIR, 'DLLs', 'tk86t.dll')
if os.path.exists(tcl_dll):
    include_files.append(tcl_dll)
if os.path.exists(tk_dll):
    include_files.append(tk_dll)

# 需要包含的完整包（包含子模块）
include_packages = [
    "tkinter",
    "tkinter.ttk",
    "tkinter.filedialog",
    "tkinter.messagebox",
    "tkinterdnd2",
    "comtypes",
    "comtypes.client",
    "comtypes.gen",
    "psutil",
    "pythoncom",
    "pywintypes",
    "concurrent.futures"
]

executables = [
    Executable(
        "w_to_pdf.py",
        base=base,
        target_name="Word转PDF转换工具.exe",
        icon="app.ico" if os.path.exists("app.ico") else None
    )
]

setup(
    name="Word转PDF转换工具",
    version="3.0",
    description="Word转PDF批量转换工具",
    options={
        "build_exe": {
            "packages": include_packages,
            "include_files": include_files,
            "excludes": [
                "tkinter.test",
                "unittest",
                "pydoc",
                "doctest",
                "test",
                "tests"
            ],
            "include_msvcr": True,
            "zip_include_packages": "*",
            "zip_exclude_packages": ["tkdnd2.9"],  # tkdnd2.9 不压缩，确保 Tcl 能找到
            "build_exe": "build/Word转PDF转换工具",
            "optimize": 0,
            "replace_paths": [("*", "")],
            "silent": True,
            # 确保包含所有必要的 DLL
            "bin_includes": [
                "pythoncom310.dll",
                "pywintypes310.dll"
            ],
            # 包含必要的模块
            "includes": [
                "comtypes.client",
                "comtypes.gen",
                "psutil._psutil_windows",
                "pythoncom",
                "pywintypes"
            ]
        }
    },
    executables=executables
)

