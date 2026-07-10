from cx_Freeze import setup, Executable
import sys
import os

# 获取当前脚本目录
SCRIPT_DIR = os.path.dirname(os.path.abspath(__file__))
PYTHON_INSTALL_DIR = os.path.dirname(sys.executable)

# 初始化包含文件列表
include_files = []

# 1. 包含图标文件
if os.path.exists(os.path.join(SCRIPT_DIR, 'app.ico')):
    include_files.append('app.ico')

# 2. 包含 tkdnd2.9 文件夹（递归包含所有文件）
# 需要递归包含所有文件，确保 tkdnd 库能被正确加载
# 注意：cx_Freeze 会将文件复制到 build_exe 目录的根目录
tkdnd_path = os.path.join(SCRIPT_DIR, 'tkdnd2.9')
if os.path.exists(tkdnd_path):
    # 递归添加 tkdnd2.9 文件夹中的所有文件
    # 目标路径保持为 tkdnd2.9/文件名，确保在打包后与 exe 同级
    tkdnd_files_count = 0
    tkdnd_dll_path = None
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
            tkdnd_files_count += 1
            
            # 特别处理 DLL 文件：也复制到 exe 同级目录，确保能被找到
            if file.endswith('.dll') and rel_dir == '.':
                tkdnd_dll_path = src_path
                # 将 DLL 也复制到根目录（与 exe 同级）
                include_files.append((src_path, file))
                print(f"✓ 已包含 tkdnd DLL 到根目录: {file}")
    
    print(f"✓ 已包含 tkdnd2.9 文件夹，共 {tkdnd_files_count} 个文件")
    
    # 验证关键文件是否存在
    pkg_index = os.path.join(tkdnd_path, 'pkgIndex.tcl')
    dll_file = os.path.join(tkdnd_path, 'libtkdnd2.9.5.dll')
    if os.path.exists(pkg_index):
        print(f"✓ 确认 pkgIndex.tcl 存在: {pkg_index}")
    else:
        print(f"✗ 警告: pkgIndex.tcl 不存在: {pkg_index}")
    if os.path.exists(dll_file):
        print(f"✓ 确认 DLL 文件存在: {dll_file}")
    else:
        print(f"✗ 警告: DLL 文件不存在: {dll_file}")
else:
    print(f"✗ 错误: tkdnd2.9 文件夹不存在: {tkdnd_path}")

# 3. 包含整个 tkinterdnd2 包（从 site-packages）
tkinterdnd2_site = os.path.join(PYTHON_INSTALL_DIR, 'lib', 'site-packages', 'tkinterdnd2')
if os.path.exists(tkinterdnd2_site):
    # 递归包含整个 tkinterdnd2 目录
    for root, dirs, files in os.walk(tkinterdnd2_site):
        for file in files:
            src_path = os.path.join(root, file)
            # 计算相对于 tkinterdnd2_site 的路径
            rel_path = os.path.relpath(src_path, tkinterdnd2_site)
            dst_path = os.path.join('lib', 'tkinterdnd2', rel_path)
            include_files.append((src_path, dst_path))
    print(f"✓ 已包含 tkinterdnd2 包: {tkinterdnd2_site}")
else:
    print(f"✗ 警告: tkinterdnd2 包不存在: {tkinterdnd2_site}")

# 4. 包含 tcl/tk DLL（如果存在，确保 Tcl 能正常工作）
tcl_dll = os.path.join(PYTHON_INSTALL_DIR, 'DLLs', 'tcl86t.dll')
tk_dll = os.path.join(PYTHON_INSTALL_DIR, 'DLLs', 'tk86t.dll')
if os.path.exists(tcl_dll):
    include_files.append(tcl_dll)
    print(f"包含 Tcl DLL: {tcl_dll}")
if os.path.exists(tk_dll):
    include_files.append(tk_dll)
    print(f"包含 Tk DLL: {tk_dll}")

# 需要包含的完整包（包含子模块）
include_packages = [
    "tkinter",
    "tkinter.ttk",
    "tkinter.filedialog",
    "tkinter.messagebox",
    "tkinterdnd2",
    "comtypes",
    "comtypes.client",
    "comtypes.gen",  # comtypes 生成的类型定义
    "psutil",
    "pythoncom",     # pywin32 的一部分，用于 COM 操作
    "pywintypes",    # pywin32 的一部分
    "win32api",      # pywin32 的一部分，可能被 comtypes 使用
    "win32com",      # pywin32 的一部分，可能被 comtypes 使用
    "concurrent.futures",
    "threading",     # 显式包含 threading 模块
    "datetime"       # 显式包含 datetime 模块
]

build_options = {
    'packages': include_packages,
    'include_files': include_files,
    'excludes': [
        "tkinter.test",
        "unittest",
        "pydoc",
        "doctest",
        "test",
        "tests"
    ],
    'include_msvcr': True,
    'zip_include_packages': "*",
    'zip_exclude_packages': ["tkdnd2.9", "tkinterdnd2"],  # tkdnd2.9 和 tkinterdnd2 不压缩，确保 Tcl 能找到
    # 确保包含所有必要的 DLL
    'bin_includes': [
        "pythoncom310.dll",
        "pywintypes310.dll",
        "libtkdnd2.9.5.dll",  # tkdnd DLL，确保能被找到
        "win32api.pyd",      # pywin32 模块
        "win32event.pyd",    # pywin32 模块
        "win32process.pyd", # pywin32 模块
        "win32evtlog.pyd"    # pywin32 模块
    ],
    # 包含必要的模块
    'includes': [
        "comtypes.client",
        "comtypes.gen",
        "psutil._psutil_windows",
        "pythoncom",
        "pywintypes",
        "threading",
        "datetime",
        "gc",
        "re",
        "time"
    ]
}

base = 'Win32GUI' if sys.platform == 'win32' else None

setup(
    name='WordToPDF',
    version='2.0',
    description='Word转PDF批量转换工具',
    options={'build_exe': build_options},
    executables=[
        Executable(
            'w_to_pdf.py',
            base=base,
            icon='app.ico',
            target_name='w_to_p',
            copyright='Copyright 2024 Mrzeng',
        )
    ]
)