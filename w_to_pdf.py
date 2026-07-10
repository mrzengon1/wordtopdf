import os
import sys
import threading
from tkinter import Tk, filedialog, messagebox, Button, Label, Listbox, END, Frame, StringVar, Menu
from tkinter.ttk import Progressbar, Scrollbar
from tkinterdnd2 import DND_FILES, TkinterDnD
import comtypes.client
import pythoncom
import time
import re
import concurrent.futures
import gc
import psutil
import datetime
import subprocess

# Windows进程优先级相关
try:
    import win32process
    import win32con
    WIN32_AVAILABLE = True
except ImportError:
    WIN32_AVAILABLE = False

# 常量定义
WORD_EXTENSIONS = ('.docx', '.doc')
LARGE_FILE_THRESHOLD_SIZE = 10 * 1024 * 1024  # 大文件阈值（10MB）- 根据文件大小判断
PDF_FILE_FORMAT = 17  # Word PDF格式代码
DOC_FILE_FORMAT = 0  # Word DOC格式代码

# UI颜色常量
COLORS = {
    'bg_main': '#f7f7f7',
    'bg_white': '#ffffff',
    'bg_light': '#f9f9f9',
    'bg_preview': '#f3f3f3',
    'border': '#e0e0e0',
    'border_dark': '#bdbdbd',
    'text_primary': '#333',
    'text_title': '#1976d2',
    'btn_green': '#4CAF50',
    'btn_green_hover': '#43a047',
    'btn_blue': '#2196F3',
    'btn_blue_hover': '#1976d2',
    'btn_red': '#E53935',
    'btn_red_hover': '#b71c1c',
    'btn_orange': '#FF9800',
    'btn_orange_hover': '#f57c00',
    'btn_primary': '#FF5722',
    'btn_primary_hover': '#d84315',
    'btn_gray': '#607D8B',
    'btn_gray_hover': '#455A64',
    'checkbox_on': '#a5d6a7',
    'checkbox_off': '#e0e0e0',
}

def format_filename(filename):
    """
    格式化文件名，去除非法字符
    """
    filename = filename.replace(" ", "_")
    filename = re.sub(r'[<>:"/\\|?*]', '_', filename)
    return filename

def set_high_priority():
    """设置当前进程为高优先级，提高转换速度"""
    if not WIN32_AVAILABLE:
        return False
    try:
        current_process = win32process.GetCurrentProcess()
        win32process.SetPriorityClass(current_process, win32process.HIGH_PRIORITY_CLASS)
        return True
    except Exception:
        return False

def set_normal_priority():
    """恢复进程优先级为正常"""
    if not WIN32_AVAILABLE:
        return False
    try:
        current_process = win32process.GetCurrentProcess()
        win32process.SetPriorityClass(current_process, win32process.NORMAL_PRIORITY_CLASS)
        return True
    except Exception:
        return False

class PowerModeManager:
    """
    电源模式管理器：控制Windows电源计划
    """
    _original_power_mode = None
    _is_enabled = False
    
    @classmethod
    def get_current_power_mode(cls):
        """获取当前电源模式GUID"""
        try:
            result = subprocess.run(
                ['powercfg', '/getactivescheme'],
                capture_output=True,
                text=True,
                timeout=2
            )
            if result.returncode == 0:
                # 提取GUID，例如："电源方案 GUID: 8c5e7fda-e8bf-4a96-9a85-a6e23a8c635c  (高性能)"
                match = re.search(r'([a-f0-9-]{36})', result.stdout, re.IGNORECASE)
                if match:
                    return match.group(1)
        except Exception:
            pass
        return None
    
    @classmethod
    def initialize(cls):
        """初始化：记录原始电源模式（不切换）"""
        if cls._original_power_mode is None:
            cls._original_power_mode = cls.get_current_power_mode()
    
    @classmethod
    def set_high_performance(cls):
        """切换到高性能模式"""
        if cls._is_enabled:
            return  # 已经设置过了
        
        try:
            # 先保存当前模式
            cls._original_power_mode = cls.get_current_power_mode()
            
            # 切换到高性能模式 (GUID: 8c5e7fda-e8bf-4a96-9a85-a6e23a8c635c)
            result = subprocess.run(
                ['powercfg', '/setactive', '8c5e7fda-e8bf-4a96-9a85-a6e23a8c635c'],
                capture_output=True,
                text=True,
                timeout=2
            )
            if result.returncode == 0:
                cls._is_enabled = True
                print("✓ 已切换到高性能电源模式")
                return True
        except Exception:
            pass
        return False
    
    @classmethod
    def restore_original_mode(cls):
        """恢复原始电源模式"""
        if not cls._is_enabled:
            return
        
        try:
            if cls._original_power_mode:
                result = subprocess.run(
                    ['powercfg', '/setactive', cls._original_power_mode],
                    capture_output=True,
                    text=True,
                    timeout=2
                )
                if result.returncode == 0:
                    print(f"✓ 已恢复原始电源模式")
            cls._is_enabled = False
            cls._original_power_mode = None
        except Exception:
            pass

class WordApp:
    """
    Word应用实例管理，支持多线程安全获取和清理
    优化：预分配内存缓冲区以提高性能
    重要：每个线程必须独立初始化COM，不能跨线程共享COM对象
    """
    _instances = {}
    _lock = threading.Lock()
    _memory_buffers = {}  # 为每个线程预分配内存缓冲区
    _com_initialized = threading.local()  # 线程本地存储，标记每个线程的COM初始化状态
    
    @classmethod
    def _allocate_memory_buffer(cls, thread_id, size_mb=200):
        """
        为每个线程预分配内存缓冲区，增加内存使用以提高性能
        size_mb: 每个线程预分配的内存大小（MB），默认最小200MB
        """
        if thread_id not in cls._memory_buffers:
            try:
                # 预分配内存缓冲区（使用bytearray来真正占用内存）
                buffer_size = size_mb * 1024 * 1024  # 转换为字节
                # 创建一个大的bytearray来真正占用内存
                buffer = bytearray(buffer_size)
                # 填充一些数据确保内存被真正分配
                buffer[0] = 1
                buffer[-1] = 1
                cls._memory_buffers[thread_id] = buffer
                print(f"[内存优化] 线程 {thread_id} 预分配 {size_mb}MB 内存缓冲区")
            except Exception as e:
                print(f"[内存优化] 线程 {thread_id} 内存预分配失败: {e}")
                cls._memory_buffers[thread_id] = None
    
    @classmethod
    def _ensure_com_initialized(cls):
        """
        确保当前线程已初始化COM（线程安全）
        每个线程必须独立初始化COM，不能跨线程共享
        """
        if not hasattr(cls._com_initialized, 'initialized'):
            try:
                # 在当前线程中初始化COM
                pythoncom.CoInitialize()
                cls._com_initialized.initialized = True
            except pythoncom.com_error:
                # 如果已经初始化，会抛出异常，可以忽略
                cls._com_initialized.initialized = True
            except Exception as e:
                print(f"pythoncom.CoInitialize()异常: {e}")
                cls._com_initialized.initialized = False
    
    @classmethod
    def get_instance(cls, thread_id):
        # 确保当前线程已初始化COM（必须在锁外调用，避免死锁）
        cls._ensure_com_initialized()
        
        with cls._lock:
            if thread_id not in cls._instances:
                # 预分配内存缓冲区（增加内存使用以提高性能）
                # 根据系统内存动态调整缓冲区大小，默认最小200MB，往上叠加
                try:
                    memory = psutil.virtual_memory()
                    total_memory_gb = memory.total / (1024 * 1024 * 1024)
                    base_buffer_size_mb = 200  # 默认最小200MB
                    if total_memory_gb >= 16:
                        buffer_size_mb = base_buffer_size_mb + 100  # 高内存系统：200+100=300MB
                    elif total_memory_gb >= 8:
                        buffer_size_mb = base_buffer_size_mb + 50   # 中等内存系统：200+50=250MB
                    else:
                        buffer_size_mb = base_buffer_size_mb        # 小内存系统：200MB
                except Exception:
                    buffer_size_mb = 200  # 默认200MB
                cls._allocate_memory_buffer(thread_id, size_mb=buffer_size_mb)
                
                # 在当前线程中创建Word实例（COM已初始化）
                word = comtypes.client.CreateObject('Word.Application')
                # 性能优化：一次性设置所有选项，减少后续操作开销
                word.Visible = False
                word.DisplayAlerts = False  # 禁用所有提示，加快操作
                word.EnableEvents = False  # 禁用事件，提高性能
                word.ScreenUpdating = False  # 禁用屏幕更新，加快速度
                # 所有Word实例都禁用自动保存和相关优化设置
                try:
                    word.Options.AutoSaveOn = False  # 禁用自动保存
                    word.Options.BackgroundSave = False  # 禁用后台保存
                    word.Options.UpdateLinksAtOpen = False  # 打开时不更新链接，加快打开速度
                    word.Options.CheckGrammarAsYouType = False  # 禁用实时语法检查，提高性能
                    word.Options.CheckSpellingAsYouType = False  # 禁用实时拼写检查，提高性能
                    word.Options.DoNotPromptForConvert = True  # 不提示转换，加快打开
                    # 增加Word内存相关设置以提高性能（不影响PDF输出质量）
                    word.Options.CacheSize = 100  # 增加缓存大小（默认50）
                    word.Options.PictureCacheSize = 100  # 增加图片缓存大小
                    # 注意：不修改Print相关设置，确保PDF输出与Word文档完全一致
                except Exception:
                    pass  # 某些设置可能不支持，忽略错误
                cls._instances[thread_id] = word
            return cls._instances[thread_id]
    
    @classmethod
    def cleanup(cls):
        with cls._lock:
            for word in list(cls._instances.values()):
                try:
                    word.Quit()
                except Exception as e:
                    # 只在调试时输出
                    # print(f"Word实例关闭异常: {e}")
                    pass
                try:
                    del word
                except Exception as e:
                    print(f"Word实例删除异常: {e}")
            cls._instances.clear()
            # 清理内存缓冲区
            cls._memory_buffers.clear()
            try:
                pythoncom.CoUninitialize()
            except Exception as e:
                print(f"pythoncom.CoUninitialize()异常: {e}")
            # 减少垃圾回收频率，让内存保持更长时间以提高性能
            # gc.collect()  # 暂时不强制垃圾回收，保持内存使用

def get_optimal_workers():
    """
    根据CPU和内存情况自动确定最佳线程数
    
    优化策略（更积极的内存使用）：
    1. 考虑CPU核心数（物理核心）
    2. 考虑总内存和可用内存
    3. 考虑系统当前负载
    4. 更积极地使用内存以提高转换速度
    """
    cpu_count = os.cpu_count() or 4  # 如果无法获取，默认4
    memory = psutil.virtual_memory()
    
    # 获取可用内存（更准确）
    available_memory_gb = memory.available / (1024 * 1024 * 1024)
    total_memory_gb = memory.total / (1024 * 1024 * 1024)
    memory_percent = memory.percent
    
    # 获取CPU使用率（考虑当前负载）
    try:
        cpu_percent = psutil.cpu_percent(interval=0.1)
    except Exception:
        cpu_percent = 0  # 如果获取失败，假设CPU空闲
    
    # 更积极的策略：根据内存和CPU情况确定基础线程数
    # 允许使用更多内存以提高转换速度
    if total_memory_gb >= 16:
        # 大内存系统（16GB+）：更积极地使用线程
        if available_memory_gb >= 8 and cpu_percent < 75:
            # 内存充足且CPU不忙：使用更多线程
            base_workers = min(int(cpu_count * 3), 48, cpu_count + 16)
        elif available_memory_gb >= 4 and cpu_percent < 85:
            # 内存中等且CPU较忙：适度增加
            base_workers = min(int(cpu_count * 2.5), 40, cpu_count + 12)
        else:
            # 内存紧张或CPU很忙：仍然使用较多线程
            base_workers = min(int(cpu_count * 2), 32, cpu_count + 8)
    elif total_memory_gb >= 8:
        # 中等内存系统（8-16GB）：适度增加线程
        if available_memory_gb >= 4 and cpu_percent < 75:
            base_workers = min(int(cpu_count * 2.5), 40, cpu_count + 12)
        elif available_memory_gb >= 2 and cpu_percent < 85:
            base_workers = min(int(cpu_count * 2), 32, cpu_count + 8)
        else:
            base_workers = min(int(cpu_count * 1.5), 24, cpu_count + 4)
    else:
        # 小内存系统（<8GB）：适度增加
        if available_memory_gb >= 2 and cpu_percent < 75:
            base_workers = min(int(cpu_count * 2), 24, cpu_count + 6)
        else:
            # 内存紧张或CPU很忙：保守但比之前更积极
            base_workers = min(int(cpu_count * 1.5), 16, cpu_count + 4)
    
    # 更积极的策略：只在内存极度紧张时才减少线程数
    # 允许程序使用更多内存以提高转换速度
    if memory_percent > 95:
        base_workers = max(4, int(base_workers * 0.7))  # 只在内存>95%时减少30%
    elif memory_percent > 90:
        base_workers = max(4, int(base_workers * 0.85))  # 内存>90%时减少15%
    elif memory_percent > 85:
        base_workers = max(4, int(base_workers * 0.95))  # 内存>85%时减少5%
    # 内存使用率<85%时，保持原有线程数，充分利用内存
    
    # 确保至少4个线程，最多不超过64个（提高上限以充分利用内存）
    final_workers = max(4, min(base_workers, 64))
    
    # 静默返回，不输出详细信息
    return final_workers

def get_document_page_count(doc):
    """
    获取Word文档的页数
    
    Args:
        doc: Word文档对象
    
    Returns:
        int: 文档页数，如果获取失败返回0
    """
    try:
        # 使用ComputeStatistics方法获取页数
        # wdStatisticPages = 2
        page_count = doc.ComputeStatistics(2)  # 2表示统计页数
        return int(page_count)
    except Exception as e:
        # 如果获取失败，尝试使用BuiltInDocumentProperties
        try:
            page_count = doc.BuiltInDocumentProperties("Pages").Value
            return int(page_count) if page_count else 0
        except Exception:
            print(f"无法获取文档页数: {e}")
            return 0

def word_to_pdf(input_path, output_path, thread_id, max_retries=2, check_stop=None):
    """
    将Word文档转换为PDF
    
    Args:
        input_path: 输入文件路径
        output_path: 输出文件路径
        thread_id: 线程ID
        max_retries: 最大重试次数（大文件可能需要重试）
        check_stop: 停止检查函数，返回True表示应停止
    
    Returns:
        bool: 转换是否成功
    """
    # 确保当前线程已初始化COM（重要：必须在每个线程中独立初始化）
    WordApp._ensure_com_initialized()
    
    # 获取Word实例（每个线程使用自己的实例）
    word = WordApp.get_instance(thread_id)
    doc = None
    is_large_file = False  # 初始化为False，根据文件大小判断
    
    # 预先处理路径和文件检查，减少重复操作
    input_path = os.path.abspath(input_path)
    if not os.path.exists(input_path):
        return False

    output_path = os.path.abspath(output_path)
    output_dir = os.path.dirname(output_path)
    if not os.path.exists(output_dir):
        os.makedirs(output_dir, exist_ok=True)
    
    # 预先计算文件大小，避免重复调用
    try:
        file_size = os.path.getsize(input_path)
        is_large_file = file_size > LARGE_FILE_THRESHOLD_SIZE
        if is_large_file:
            max_retries = 3  # 大文件使用更多重试次数
        else:
            max_retries = 1  # 普通文件只重试1次，加快速度
    except Exception:
        file_size = 0
        is_large_file = False
        max_retries = 1
    
    for attempt in range(max_retries + 1):
        # 减少停止检查频率（只在关键点检查）
        if check_stop and check_stop():
            return False
        
        try:
            
            # Word COM对象需要按位置传递参数，不能使用关键字参数
            # Documents.Open(FileName, ConfirmConversions, ReadOnly, AddToRecentFiles, 
            #                PasswordDocument, PasswordTemplate, Revert, WritePasswordDocument, 
            #                WritePasswordTemplate, Format, Encoding, Visible, OpenConflictDocument, 
            #                OpenAndRepair, DocumentDirection, NoEncodingDialog, XMLTransform)
            # 优化：使用最简参数，加快打开速度
            doc = word.Documents.Open(
                input_path,        # FileName
                False,             # ConfirmConversions
                True,              # ReadOnly - 只读模式，提高性能
                False,             # AddToRecentFiles - 不添加到最近文件列表
                "",                # PasswordDocument
                "",                # PasswordTemplate
                False,             # Revert
                "",                # WritePasswordDocument
                "",                # WritePasswordTemplate
                0                  # Format - 0表示自动检测格式
            )
            
            # 优化文档处理 - 只设置不影响PDF输出的显示选项
            # 注意：这些设置只影响Word界面显示，不会修改文档内容，也不会影响PDF输出
            try:
                doc.ShowRevisions = False  # 隐藏修订标记（仅显示，不影响文档和PDF）
                doc.TrackRevisions = False  # 禁用跟踪修订（不影响已存在的修订内容）
                doc.AutoSaveOn = False  # 禁用自动保存（不影响文档内容）
                # 隐藏错误标记（仅UI显示，不影响文档内容和PDF输出）
                doc.ShowGrammaticalErrors = False
                doc.ShowSpellingErrors = False
                doc.ShowFormattingErrors = False
            except Exception:
                pass  # 某些文档可能不支持这些属性
            
            # 检查是否已停止（在保存前）
            if check_stop and check_stop():
                doc.Close(SaveChanges=False)
                doc = None
                return False
            
            # Word COM对象需要按位置传递参数
            # Document.SaveAs(FileName, FileFormat, LockComments, Password, AddToRecentFiles, 
            #                WritePassword, ReadOnlyRecommended, EmbedTrueTypeFonts, 
            #                SaveNativePictureFormat, SaveFormsData, CompressPictures, 
            #                EmbedFonts, EmbedLinguisticData, SaveSubsetFonts, 
            #                CompatibilityMode, OptimizeForPrint, ...)
            # 保存为PDF：使用默认设置确保PDF与Word文档完全一致
            # 注意：不修改可能影响PDF质量的参数，确保输出与Word文档一致
            doc.SaveAs(
                output_path,       # FileName
                PDF_FILE_FORMAT,   # FileFormat - 17表示PDF格式
                False,             # LockComments
                "",                # Password
                False,             # AddToRecentFiles - 不添加到最近文件列表
                "",                # WritePassword
                False,             # ReadOnlyRecommended
                True,              # EmbedTrueTypeFonts - 嵌入字体，确保PDF与Word一致
                False,             # SaveNativePictureFormat
                False,             # SaveFormsData
                False,             # CompressPictures - 使用默认压缩
                True,              # EmbedFonts - 嵌入字体，确保PDF与Word一致
                False,             # EmbedLinguisticData
                False,             # SaveSubsetFonts
                False,             # CompatibilityMode
                True              # OptimizeForPrint - 优化打印，确保PDF质量与Word一致
            )
            
            doc.Close(SaveChanges=False)
            doc = None  # 标记已关闭

            # 快速验证输出文件（只检查存在性，不检查大小以加快速度）
            if os.path.exists(output_path):
                # 快速检查：如果文件大小大于0，认为成功（避免详细检查拖慢速度）
                try:
                    if os.path.getsize(output_path) > 0:
                        return True
                    else:
                        # 文件为空，删除并重试
                        try:
                            os.remove(output_path)
                        except Exception:
                            pass
                except Exception:
                    # 如果无法获取文件大小，但文件存在，也认为成功
                    return True
            # 失败时不打印，减少输出开销
            
            # 如果失败且还有重试机会，继续重试
            if attempt < max_retries:
                time.sleep(0.05)  # 最小等待时间，加快重试速度
                continue
            else:
                return False
                
        except Exception as e:
            error_msg = str(e)
            
            # 确保文档被关闭
            if doc is not None:
                try:
                    doc.Close(SaveChanges=False)
                except Exception:
                    pass
                doc = None
            
            # 如果是大文件且是特定错误，可以重试
            if is_large_file and attempt < max_retries:
                # 常见的大文件错误：内存不足、超时等
                retryable_errors = ['内存', 'memory', 'timeout', '超时', 'busy', '忙', 'locked', '锁定']
                if any(err in error_msg.lower() for err in retryable_errors):
                    time.sleep(0.1)  # 最小等待时间，加快重试速度
                    continue
            
            # 最后一次尝试失败，返回False
            if attempt >= max_retries:
                return False
    
    return False

class WordToPDFConverter:
    """
    Word转PDF批量转换器，支持拖拽、批量、进度显示等功能
    """
    
    @staticmethod
    def is_word_file(file_path):
        """检查文件是否为Word文件"""
        return file_path.lower().endswith(WORD_EXTENSIONS)
    
    def add_file_to_list(self, file_path):
        """将文件添加到列表（避免重复）"""
        if file_path not in self.file_paths:
            self.file_paths.append(file_path)
            if not self.destroyed:
                try:
                    self.root.after(0, self.listbox.insert, END, file_path)
                except Exception:
                    pass
            return True
        return False
    
    def find_word_files_in_dir(self, directory):
        """在目录中查找所有Word文件"""
        word_files = []
        for root_dir, _, files in os.walk(directory):
            for file in files:
                if self.is_word_file(file):
                    full_path = os.path.join(root_dir, file)
                    word_files.append(full_path)
        return word_files
    
    def __init__(self):
        self.file_paths = []
        # 创建 TkinterDnD 根窗口
        self.root = TkinterDnD.Tk()
        
        # 记录原始电源模式（用于程序结束时恢复，如果程序改变了的话）
        # 注意：不自动切换高性能模式，让系统自然管理
        PowerModeManager.initialize()
        
        # 获取最优线程数（函数内部会输出详细信息）
        self.max_workers = get_optimal_workers()
        self.conversion_active = False
        self.destroyed = False  # 标记窗口是否已被销毁
        self.start_time = None
        self.log_records = []
        self.folder_preview_files = []
        self.output_dir = StringVar()
        self.output_dir.set("")
        self.auto_open_dir = StringVar()
        self.auto_open_dir.set("1")
        self.setup_ui()
        self.root.protocol("WM_DELETE_WINDOW", self.on_closing)
        # 主窗口和listbox都注册拖拽
        self.root.drop_target_register(DND_FILES)
        self.root.dnd_bind('<<Drop>>', self.on_drop_files)
        self.listbox.drop_target_register(DND_FILES)
        self.listbox.dnd_bind('<<Drop>>', self.on_drop_files)

    def setup_ui(self):
        """
        初始化和布局所有UI控件（美化版）
        """
        self.root.title("Word转PDF批量转换工具 V2.0")
        self.root.geometry("650x700")
        self.root.resizable(True, True)
        self.root.config(bg=COLORS['bg_main'])

        self.main_frame = Frame(self.root, bg=COLORS['bg_white'], highlightbackground=COLORS['border'], highlightthickness=2)
        self.main_frame.grid(row=0, column=0, sticky='nsew', padx=20, pady=20)
        self.root.grid_rowconfigure(0, weight=1)
        self.root.grid_columnconfigure(0, weight=1)

        # 标题
        self.label = Label(self.main_frame, text="Word转PDF批量转换", font=("微软雅黑", 18, "bold"), bg=COLORS['bg_white'], fg=COLORS['text_title'])
        self.label.grid(row=0, column=0, pady=(30, 10), sticky='n')

        # 输出目录选择区
        outdir_frame = Frame(self.main_frame, bg=COLORS['bg_white'])
        outdir_frame.grid(row=1, column=0, pady=(0, 5), sticky='ew')
        outdir_frame.grid_columnconfigure(1, weight=1)
        Label(outdir_frame, text="输出目录:", font=("微软雅黑", 10), bg=COLORS['bg_white']).grid(row=0, column=0, sticky='w')
        self.outdir_entry = Label(outdir_frame, textvariable=self.output_dir, font=("微软雅黑", 10), bg=COLORS['bg_light'], anchor='w', relief='sunken', width=40)
        self.outdir_entry.grid(row=0, column=1, sticky='ew', padx=5)
        Button(outdir_frame, text="选择", command=self.choose_output_dir, font=("微软雅黑", 10), bg=COLORS['btn_blue'], fg="white", relief="flat").grid(row=0, column=2, padx=5)
        self.auto_open_check = Button(outdir_frame, text="✔ 转换后自动打开输出目录", font=("微软雅黑", 9), bg=COLORS['checkbox_off'], relief="flat", command=self.toggle_auto_open)
        self.auto_open_check.grid(row=1, column=0, columnspan=3, sticky='w', pady=2)
        self.update_auto_open_btn()

        # 文件夹内容预览区（默认隐藏）
        self.folder_preview_frame = Frame(self.main_frame, bg=COLORS['bg_preview'], highlightbackground=COLORS['border_dark'], highlightthickness=1)
        self.folder_preview_frame.grid(row=2, column=0, sticky='ew', pady=(0, 8))
        self.folder_preview_frame.grid_remove()
        self.folder_preview_label = Label(self.folder_preview_frame, text="", font=("微软雅黑", 10), bg=COLORS['bg_preview'])
        self.folder_preview_label.pack(anchor='w', pady=(5, 2), padx=5)
        preview_list_frame = Frame(self.folder_preview_frame, bg=COLORS['bg_preview'])
        preview_list_frame.pack(fill='both', expand=True, padx=5)
        self.folder_preview_listbox = Listbox(preview_list_frame, selectmode='multiple', font=("微软雅黑", 10), width=80, height=8)
        self.folder_preview_listbox.pack(side='left', fill='both', expand=True)
        preview_scrollbar = Scrollbar(preview_list_frame, orient="vertical", command=self.folder_preview_listbox.yview)
        preview_scrollbar.pack(side='right', fill='y')
        self.folder_preview_listbox.config(yscrollcommand=preview_scrollbar.set)
        preview_btn_frame = Frame(self.folder_preview_frame, bg=COLORS['bg_preview'])
        preview_btn_frame.pack(fill='x', pady=5)
        Button(preview_btn_frame, text="全选", command=lambda: self.folder_preview_listbox.select_set(0, END), width=8).pack(side='left', padx=5)
        Button(preview_btn_frame, text="全不选", command=lambda: self.folder_preview_listbox.select_clear(0, END), width=8).pack(side='left', padx=5)
        Button(preview_btn_frame, text="添加到列表", command=self.add_selected_folder_files, width=12, bg=COLORS['btn_green'], fg="white").pack(side='right', padx=5)
        Button(preview_btn_frame, text="取消", command=self.hide_folder_preview, width=10).pack(side='right', padx=5)

        # 按钮区
        btn_frame = Frame(self.main_frame, bg=COLORS['bg_white'])
        btn_frame.grid(row=3, column=0, pady=10, sticky='ew')
        for i in range(4):
            btn_frame.grid_columnconfigure(i, weight=1)
        self.add_button = Button(btn_frame, text="➕ 添加文件", command=self.add_files, bg=COLORS['btn_green'], fg="white", font=("微软雅黑", 11, "bold"), relief="flat", activebackground=COLORS['btn_green_hover'], cursor="hand2")
        self.add_button.grid(row=0, column=0, padx=8, sticky='ew')
        self.add_folder_button = Button(btn_frame, text="📂 选择文件夹", command=self.add_folder, bg=COLORS['btn_blue'], fg="white", font=("微软雅黑", 11, "bold"), relief="flat", activebackground=COLORS['btn_blue_hover'], cursor="hand2")
        self.add_folder_button.grid(row=0, column=1, padx=8, sticky='ew')
        self.delete_button = Button(btn_frame, text="🗑️ 删除选中", command=self.delete_selected_files, bg=COLORS['btn_red'], fg="white", font=("微软雅黑", 11, "bold"), relief="flat", activebackground=COLORS['btn_red_hover'], cursor="hand2")
        self.delete_button.grid(row=0, column=2, padx=8, sticky='ew')
        self.clear_button = Button(btn_frame, text="🧹 清空列表", command=self.clear_files, bg=COLORS['btn_orange'], fg="white", font=("微软雅黑", 11, "bold"), relief="flat", activebackground=COLORS['btn_orange_hover'], cursor="hand2")
        self.clear_button.grid(row=0, column=3, padx=8, sticky='ew')

        # 文件列表
        list_frame = Frame(self.main_frame, bg=COLORS['bg_white'])
        list_frame.grid(row=4, column=0, pady=(10, 5), sticky='nsew')
        self.main_frame.grid_rowconfigure(4, weight=1)
        self.main_frame.grid_columnconfigure(0, weight=1)
        list_frame.grid_rowconfigure(0, weight=1)
        list_frame.grid_columnconfigure(0, weight=1)
        self.listbox = Listbox(list_frame, font=("微软雅黑", 10), selectmode='multiple', bd=2, relief="solid", highlightthickness=0, bg=COLORS['bg_light'], fg=COLORS['text_primary'])
        self.listbox.grid(row=0, column=0, sticky='nsew')
        scrollbar = Scrollbar(list_frame, orient="vertical", command=self.listbox.yview)
        scrollbar.grid(row=0, column=1, sticky='ns')
        self.listbox.config(yscrollcommand=scrollbar.set)

        # 分隔线
        sep = Frame(self.main_frame, height=2, bg=COLORS['border'])
        sep.grid(row=5, column=0, sticky='ew', pady=10)

        # 进度区
        progress_frame = Frame(self.main_frame, bg=COLORS['bg_white'])
        progress_frame.grid(row=6, column=0, pady=(10, 0), sticky='ew')
        progress_frame.grid_columnconfigure(1, weight=1)
        self.progress_label = Label(progress_frame, text="转换进度：", font=("微软雅黑", 11), bg=COLORS['bg_white'], fg=COLORS['text_primary'])
        self.progress_label.grid(row=0, column=0, padx=(0, 10), sticky='w')
        self.progress = Progressbar(progress_frame, mode='determinate', style="TProgressbar")
        self.progress.grid(row=0, column=1, sticky='ew')
        self.percent_label = Label(progress_frame, text="0%", font=("微软雅黑", 10), bg=COLORS['bg_white'], fg=COLORS['text_primary'])
        self.percent_label.grid(row=0, column=2, padx=(10, 0), sticky='e')

        # 开始转换/停止转换按钮（动态切换）
        self.convert_button = Button(self.main_frame, text="🚀 开始转换", command=self.start_conversion, bg=COLORS['btn_primary'], fg="white", font=("微软雅黑", 13, "bold"), relief="flat", activebackground=COLORS['btn_primary_hover'], cursor="hand2")
        self.convert_button.grid(row=7, column=0, pady=25, sticky='ew')

        # 导出日志按钮
        self.export_log_button = Button(self.main_frame, text="📄 导出日志", command=self.export_log, bg=COLORS['btn_gray'], fg="white", font=("微软雅黑", 10), relief="flat", activebackground=COLORS['btn_gray_hover'], cursor="hand2")
        self.export_log_button.grid(row=8, column=0, pady=(0, 10), sticky='ew')

        # 进度条样式
        try:
            from tkinter import ttk
            style = ttk.Style()
            style.theme_use('clam')
            style.configure('red.Horizontal.TProgressbar', background=COLORS['btn_primary'], thickness=18, troughcolor=COLORS['bg_main'], bordercolor=COLORS['bg_main'])
            self.progress.config(style='red.Horizontal.TProgressbar')
        except Exception:
            pass

        # 文件列表右键菜单
        self.listbox.bind("<Button-3>", self.show_listbox_menu)
        self.listbox_menu = Menu(self.listbox, tearoff=0)
        self.listbox_menu.add_command(label="移除选中", command=self.delete_selected_files)
        self.listbox_menu.add_command(label="打开所在目录", command=self.open_selected_file_dir)

    def on_closing(self):
        self.conversion_active = False
        self.destroyed = True  # 标记窗口正在被销毁
        WordApp.cleanup()
        # 恢复原始电源模式（如果程序改变过的话）
        PowerModeManager.restore_original_mode()
        try:
            self.root.destroy()
        except Exception:
            pass  # 忽略销毁时的异常

    def add_files(self):
        if self.conversion_active:
            messagebox.showwarning("警告", "转换进行中，请等待完成后再添加文件！")
            return
        try:
            files = filedialog.askopenfilenames(
                title="选择 Word 文件",
                filetypes=[("Word 文件", "*.docx;*.doc")],
            )
            added_count = 0
            for file in files:
                if self.add_file_to_list(file):
                    added_count += 1
            if added_count > 0:
                messagebox.showinfo("添加完成", f"成功添加 {added_count} 个文件")
        except Exception as e:
            print(f"添加文件时出错: {e}")
            messagebox.showerror("错误", f"添加文件时出错: {e}")

    def filter_conflict_files(self, word_files, folder_path):
        # 检查主目录下的Word文件
        main_dir_files = {os.path.basename(f): f for f in word_files if os.path.dirname(f) == folder_path}
        sub_dir_conflicts = []
        for f in word_files:
            if os.path.dirname(f) != folder_path and os.path.basename(f) in main_dir_files:
                sub_dir_conflicts.append(f)
        if sub_dir_conflicts:
            msg = "检测到以下子目录存在与主目录同名的文件：\n"
            for f in sub_dir_conflicts:
                msg += f"目录: {os.path.dirname(f)}\n文件: {os.path.basename(f)}\n\n"
            msg += "是否删除这些子目录下的同名文件？\n选择'是'则不添加这些文件，选择'否'则全部添加。"
            if messagebox.askyesno("同名文件处理", msg):
                word_files = [f for f in word_files if f not in sub_dir_conflicts]
        return word_files

    def add_folder(self):
        if self.conversion_active:
            messagebox.showwarning("警告", "转换进行中，请等待完成后再添加文件！")
            return
        try:
            folder_path = filedialog.askdirectory(title="选择文件夹")
            if not folder_path:
                return
            
            # 自动全选模式新逻辑
            all_word_files = self.find_word_files_in_dir(folder_path)

            if not all_word_files:
                messagebox.showinfo("提示", "该文件夹下没有可用的Word文件！")
                return

            # 直接处理冲突文件并添加
            filtered_files = self.filter_conflict_files(all_word_files, folder_path)
            added_count = 0
            for file in filtered_files:
                if self.add_file_to_list(file):
                    added_count += 1

            if added_count > 0:
                messagebox.showinfo("添加完成", f"成功添加 {added_count} 个文件")
            else:
                messagebox.showinfo("提示", "没有新文件被添加（可能已存在或冲突）")

        except Exception as e:
            messagebox.showerror("错误", f"添加文件夹时出错: {e}")

    def show_folder_preview(self, file_list, folder_path):
        self.folder_preview_files = file_list
        self.folder_preview_label.config(text=f"文件夹: {folder_path}\n共找到 {len(file_list)} 个Word文件，请选择要添加的文件：")
        self.folder_preview_listbox.delete(0, END)
        for f in file_list:
            self.folder_preview_listbox.insert(END, f)
        self.folder_preview_frame.grid()

    def hide_folder_preview(self):
        self.folder_preview_frame.grid_remove()
        self.folder_preview_files = []
        self.folder_preview_listbox.delete(0, END)
        self.folder_preview_label.config(text="")

    def add_selected_folder_files(self):
        selected = self.folder_preview_listbox.curselection()
        count = 0
        selected_files = [self.folder_preview_files[idx] for idx in selected]
        # 新增：同名文件处理
        if selected_files:
            folder_path = os.path.dirname(os.path.commonprefix(selected_files))
            selected_files = self.filter_conflict_files(selected_files, folder_path)
        for file in selected_files:
            if self.add_file_to_list(file):
                count += 1
        self.hide_folder_preview()
        if count == 0:
            messagebox.showinfo("提示", "没有新文件被添加（可能已在列表中）")

    def clear_files(self):
        if self.conversion_active:
            messagebox.showwarning("警告", "转换进行中，请等待完成后再清空列表！")
            return
        self.file_paths.clear()
        if not self.destroyed:
            try:
                self.root.after(0, self.listbox.delete, 0, END)
            except Exception:
                pass

    def delete_selected_files(self):
        selected_indices = list(self.listbox.curselection())
        if not selected_indices:
            messagebox.showinfo("提示", "请先选中要删除的文件！")
            return
        for idx in reversed(selected_indices):
            if not self.destroyed:
                try:
                    self.root.after(0, self.listbox.delete, idx)
                except Exception:
                    pass
            del self.file_paths[idx]

    def preprocess_files(self):
        """预处理文件：排序、验证（优化：先处理小文件，提高整体速度）"""
        # 按文件大小排序，先处理小文件（小文件转换快，可以快速看到进度）
        # 这样可以提高用户体验，同时小文件不会占用太多资源
        valid_files = []
        file_sizes = []
        for file in self.file_paths:
            if os.path.exists(file) and os.access(file, os.R_OK):
                try:
                    size = os.path.getsize(file)
                    valid_files.append(file)
                    file_sizes.append(size)
                except Exception:
                    # 文件无法访问，跳过（不打印，减少输出开销）
                    pass
            # 文件不存在或无法访问，跳过（不打印，减少输出开销）
        
        # 按文件大小排序：小文件在前，大文件在后（小文件转换快，先完成可以提高整体速度）
        if valid_files:
            sorted_files = sorted(zip(valid_files, file_sizes), key=lambda x: x[1])
            self.file_paths = [f[0] for f in sorted_files]
        else:
            self.file_paths = valid_files
        
        # 注意：大文件判断会在转换时根据文件大小（>10MB）进行
        return bool(valid_files)

    def start_conversion(self):
        if self.conversion_active:
            messagebox.showwarning("警告", "转换已在进行中！")
            return
            
        if not self.file_paths:
            messagebox.showwarning("警告", "请先选择文件！")
            return

        if not self.preprocess_files():
            messagebox.showerror("错误", "没有有效的文件可以转换！")
            return

        self.conversion_active = True
        # 将按钮改为停止按钮
        self.convert_button.config(text="⏹️ 停止转换", command=self.stop_conversion, state="normal", bg=COLORS['btn_red'], activebackground=COLORS['btn_red_hover'])
        self.label.config(text="正在转换，已用时: 0分0秒", fg=COLORS['btn_primary'])

        self.start_time = time.time()
        self.progress['value'] = 0
        self.progress['maximum'] = len(self.file_paths)
        self.percent_label.config(text="0%")

        # 启动定时器实时刷新已用时间
        self.update_elapsed_time()

        threading.Thread(target=self.convert_files, daemon=True).start()
    
    def stop_conversion(self):
        """停止转换"""
        if not self.conversion_active:
            return
        if messagebox.askyesno("确认停止", "确定要停止当前转换吗？\n\n已完成的文件会保留，未完成的文件将不会转换。"):
            self.conversion_active = False
            self.convert_button.config(text="正在停止...", state="disabled")
            self.label.config(text="正在停止转换...", fg=COLORS['btn_red'])
            print("用户请求停止转换")

    def update_elapsed_time(self):
        # 检查窗口是否已被销毁
        if self.destroyed or not self.conversion_active or not self.start_time:
            return
        try:
            elapsed_time = int(time.time() - self.start_time)
            minutes = elapsed_time // 60
            seconds = elapsed_time % 60
            self.label.config(
                text=f"正在转换，已用时: {minutes}分{seconds}秒",
                fg=COLORS['btn_primary']
            )
            # 只有在窗口未被销毁时才调度下一次更新
            if not self.destroyed:
                self.root.after(1000, self.update_elapsed_time)
        except Exception:
            # 如果窗口已被销毁，忽略异常
            pass

    def process_single_file(self, file_path, thread_id):
        """处理单个文件的转换"""
        # 检查是否已停止（快速检查）
        if not self.conversion_active:
            return file_path, None, False, "转换已停止"
        
        try:
            # 预先处理路径，减少重复操作
            file_path = os.path.normpath(file_path)
            
            # 输出目录逻辑（预先计算）
            if self.output_dir.get():
                folder_path = self.output_dir.get()
            else:
                folder_path = os.path.dirname(file_path)
            
            # 生成输出文件名（预先计算）
            base_name = os.path.splitext(os.path.basename(file_path))[0]
            base_pdf_name = format_filename(base_name + ".pdf")
            output_path = os.path.normpath(os.path.join(folder_path, base_pdf_name))
            
            # 如果文件已存在，添加时间戳（预先检查）
            if os.path.exists(output_path):
                now_str = datetime.datetime.now().strftime("_%Y%m%d_%H%M%S")
                name, ext = os.path.splitext(base_pdf_name)
                output_path = os.path.normpath(os.path.join(folder_path, f"{name}{now_str}{ext}"))
            
            # 注意：大文件判断在word_to_pdf函数内部通过文件大小进行
            # 这里使用默认重试次数，word_to_pdf函数会根据实际文件大小调整
            max_retries = 1  # 默认重试次数，word_to_pdf会根据文件大小自动调整
            
            # 传递停止检查函数（使用lambda减少函数调用开销）
            check_stop = lambda: not self.conversion_active
            
            success = word_to_pdf(file_path, output_path, thread_id, max_retries=max_retries, check_stop=check_stop)
            return file_path, output_path, success
        except Exception as e:
            return file_path, None, False, str(e)

    def convert_files(self):
        failed_files = []
        failed_reasons = {}
        results = []
        completed_count = 0
        output_dirs = set()
        self.log_records = []
        total_size = sum(os.path.getsize(f) for f in self.file_paths if os.path.exists(f))
        
        # 预分配全局内存缓冲区以提高性能（增加内存使用）
        global_memory_buffer = None
        try:
            # 根据系统内存动态调整全局缓冲区大小
            # 默认最小值为200MB，根据系统内存往上叠加
            memory = psutil.virtual_memory()
            total_memory_gb = memory.total / (1024 * 1024 * 1024)
            base_buffer_size_mb = 200  # 默认最小200MB
            if total_memory_gb >= 16:
                global_buffer_size_mb = base_buffer_size_mb + 400  # 高内存系统：200+400=600MB
            elif total_memory_gb >= 8:
                global_buffer_size_mb = base_buffer_size_mb + 200  # 中等内存系统：200+200=400MB
            else:
                global_buffer_size_mb = base_buffer_size_mb  # 小内存系统：200MB
            
            buffer_size = global_buffer_size_mb * 1024 * 1024
            global_memory_buffer = bytearray(buffer_size)
            # 填充一些数据确保内存被真正分配
            global_memory_buffer[0] = 1
            global_memory_buffer[-1] = 1
            print(f"[内存优化] 预分配全局内存缓冲区 {global_buffer_size_mb}MB")
        except Exception as e:
            print(f"[内存优化] 全局内存缓冲区预分配失败: {e}")
        
        # 减少垃圾回收频率，保持内存使用以提高性能
        gc.set_threshold(1000, 15, 15)  # 进一步提高GC阈值，减少GC频率，加快转换速度
        
        # 注意：大文件判断在转换时根据文件大小进行（>10MB）
        # 根据文件数量动态调整并发数（限制最大线程数，避免资源耗尽导致闪退）
        file_count = len(self.file_paths)
        
        # 限制最大线程数，避免文件很多时资源耗尽导致闪退
        # 根据系统内存动态调整最大线程数上限
        try:
            memory = psutil.virtual_memory()
            total_memory_gb = memory.total / (1024 * 1024 * 1024)
            if total_memory_gb >= 16:
                max_workers_limit = 40  # 大内存系统：最多40个线程
            elif total_memory_gb >= 8:
                max_workers_limit = 30  # 中等内存系统：最多30个线程
            else:
                max_workers_limit = 20  # 小内存系统：最多20个线程
        except Exception:
            max_workers_limit = 25  # 默认最多25个线程
        
        if file_count > 100:
            # 文件非常多时，限制线程数，避免资源耗尽
            adjusted_workers = min(self.max_workers + 8, int(self.max_workers * 1.5), max_workers_limit)
        elif file_count > 50:
            # 文件很多时，适度增加线程数
            adjusted_workers = min(self.max_workers + 6, int(self.max_workers * 1.4), max_workers_limit)
        elif file_count > 20:
            # 文件较多时，适度增加
            adjusted_workers = min(self.max_workers + 4, int(self.max_workers * 1.3), max_workers_limit)
        elif file_count > 10:
            # 文件中等时，稍微增加
            adjusted_workers = min(self.max_workers + 2, int(self.max_workers * 1.2), max_workers_limit)
        else:
            # 文件较少时，使用基础线程数
            adjusted_workers = min(self.max_workers, max_workers_limit)
        
        # 减少输出，提高性能（只在调试时输出）
        # print(f"[转换配置] 文件总数: {file_count}, 使用并发线程数: {adjusted_workers}")
        
        # 注意：不再预加载Word实例，因为COM对象必须在每个线程中独立初始化
        # 每个线程会在第一次调用word_to_pdf时自动创建自己的Word实例
        
        try:
            # 使用优化的线程池，提高CPU利用率
            # 注意：Word COM对象是线程安全的，可以安全地并发使用
            with concurrent.futures.ThreadPoolExecutor(
                max_workers=adjusted_workers,
                thread_name_prefix="WordConverter"
            ) as executor:
                future_to_file = {}
                # 批量提交任务，添加异常处理防止闪退
                for thread_id, file_path in enumerate(self.file_paths):
                    try:
                        future = executor.submit(self.process_single_file, file_path, thread_id)
                        future_to_file[future] = file_path
                    except Exception as e:
                        # 如果提交任务失败，记录错误但继续处理其他文件
                        failed_files.append(file_path)
                        failed_reasons[file_path] = f"任务提交失败: {str(e)}"
                        self.log_records.append(f"[{self._now()}][失败] {file_path} - 任务提交失败")
                
                # 处理完成的任务
                for future in concurrent.futures.as_completed(future_to_file):
                    if not self.conversion_active:
                        # 取消所有未完成的任务
                        try:
                            for f in future_to_file:
                                if not f.done():
                                    f.cancel()
                        except Exception:
                            pass
                        break
                    
                    file_path = future_to_file.get(future)
                    if file_path is None:
                        continue
                    
                    try:
                        # 注意：不强制设置超时，防止大文件（>5分钟）被误判为卡死
                        # 如果你希望限制单个文件的最大转换时间，可以在这里传入 timeout 参数
                        result = future.result()
                        results.append(result)
                        if not result[2]:
                            failed_files.append(file_path)
                            failed_reasons[file_path] = result[3] if len(result) > 3 else "转换失败"
                            self.log_records.append(f"[{self._now()}][失败] {file_path} - {failed_reasons[file_path]}")
                        else:
                            self.log_records.append(f"[{self._now()}][成功] {file_path}")
                    except concurrent.futures.TimeoutError:
                        failed_files.append(file_path)
                        failed_reasons[file_path] = "转换超时（超过5分钟）"
                        self.log_records.append(f"[{self._now()}][失败] {file_path} - 转换超时")
                    except Exception as e:
                        failed_files.append(file_path)
                        failed_reasons[file_path] = str(e)
                        self.log_records.append(f"[{self._now()}][失败] {file_path} - {e}")
                    
                    completed_count += 1
                    
                    # 每处理50个文件，清理一次Word实例，防止内存累积
                    if completed_count % 50 == 0:
                        try:
                            WordApp.cleanup()
                            gc.collect()
                        except Exception:
                            pass
                    
                    # 只有在窗口未被销毁时才更新UI（进一步减少更新频率以提高性能）
                    if not self.destroyed:
                        try:
                            total_files = len(self.file_paths)
                            # 每完成10个文件或达到整数百分比时才更新UI，大幅减少UI更新开销
                            if completed_count % 10 == 0 or completed_count == total_files:
                                self.root.after(0, self.progress.config, {'value': completed_count})
                                percent = int(completed_count * 100 / total_files)
                                self.root.after(0, self.percent_label.config, {'text': f"{percent}%"})
                                # 只在完成时或每20个文件更新一次UI，进一步减少CPU占用
                                if completed_count % 20 == 0 or completed_count == total_files:
                                    self.root.after(0, self.root.update_idletasks)
                        except Exception:
                            # 如果窗口已被销毁，忽略异常
                            pass
            output_files = []
            for _, output_path, success, *_ in results:
                if success and output_path and os.path.exists(output_path):
                    output_files.append(output_path)
                    output_dirs.add(os.path.dirname(output_path))
            # 结果弹窗（只有在窗口未被销毁时才显示）
            if not self.destroyed:
                try:
                    total_time = time.time() - self.start_time
                    was_stopped = completed_count < len(self.file_paths)
                    remaining_count = len(self.file_paths) - completed_count
                    
                    if was_stopped:
                        stat = f"转换已停止\n\n总文件数: {len(self.file_paths)}\n已完成: {completed_count}\n未完成: {remaining_count}\n成功: {len(output_files)}\n失败: {len(failed_files)}\n已用时间: {int(total_time)//60}分{int(total_time)%60}秒"
                        if failed_files:
                            error_message = "以下文件转换失败:\n\n"
                            for file in failed_files:
                                error_message += f"文件: {os.path.basename(file)}\n原因: {failed_reasons.get(file, '未知错误')}\n\n"
                            messagebox.showwarning("转换已停止", stat + "\n\n" + error_message)
                        else:
                            messagebox.showinfo("转换已停止", stat)
                    else:
                        stat = f"总文件数: {len(self.file_paths)}\n成功: {len(output_files)}\n失败: {len(failed_files)}\n总耗时: {int(total_time)//60}分{int(total_time)%60}秒\n总大小: {total_size/1024/1024:.2f}MB"
                        if failed_files:
                            error_message = "以下文件转换失败:\n\n"
                            for file in failed_files:
                                error_message += f"文件: {os.path.basename(file)}\n原因: {failed_reasons.get(file, '未知错误')}\n\n"
                            messagebox.showerror("转换失败", stat + "\n\n" + error_message)
                        else:
                            messagebox.showinfo("完成", stat + "\n\n输出目录:\n" + "\n".join(output_dirs))
                    
                    # 自动打开输出目录（仅在完成时，停止时不自动打开）
                    if not was_stopped and self.auto_open_dir.get() == "1" and output_dirs:
                        for output_dir in output_dirs:
                            try:
                                os.startfile(output_dir)
                            except Exception as e:
                                print(f"无法打开输出目录 {output_dir}: {e}")
                except Exception:
                    # 如果窗口已被销毁，忽略异常
                    pass
        except Exception as e:
            # 捕获所有未处理的异常，防止程序闪退
            error_msg = f"转换过程中发生未处理的错误: {str(e)}"
            print(error_msg)
            if not self.destroyed:
                try:
                    messagebox.showerror("错误", f"转换过程中发生错误:\n{error_msg}\n\n已完成的文件会保留。")
                except Exception:
                    pass
        finally:
            was_stopped = not self.conversion_active
            self.conversion_active = False
            # 清理Word实例，释放资源
            try:
                WordApp.cleanup()
                gc.collect()
            except Exception:
                pass
            # 只有在窗口未被销毁时才更新UI
            if not self.destroyed:
                try:
                    # 恢复按钮状态和文本
                    self.root.after(0, self.convert_button.config, {
                        'text': "🚀 开始转换",
                        'command': self.start_conversion,
                        'state': "normal",
                        'bg': COLORS['btn_primary'],
                        'activebackground': COLORS['btn_primary_hover']
                    })
                    if self.start_time:
                        elapsed_time = int(time.time() - self.start_time)
                        minutes = elapsed_time // 60
                        seconds = elapsed_time % 60
                        if was_stopped:
                            # 如果被停止，显示停止信息
                            self.root.after(0, self.label.config, {'text': f"转换已停止，已用时: {minutes}分{seconds}秒", 'fg': COLORS['btn_red']})
                        else:
                            self.root.after(0, self.label.config, {'text': f"总用时: {minutes}分{seconds}秒", 'fg': COLORS['text_primary']})
                    else:
                        self.root.after(0, self.label.config, {'text': "选择 Word 文件进行批量转换", 'fg': COLORS['text_primary']})
                except Exception:
                    # 如果窗口已被销毁，忽略异常
                    pass
            WordApp.cleanup()
            # 清理全局内存缓冲区
            global_memory_buffer = None
            # 恢复默认GC阈值
            gc.set_threshold(700, 10, 10)
            # 最后进行一次垃圾回收
            gc.collect()

    def on_drop_files(self, event):
        print("拖拽事件触发", event.data)
        # event.data 可能是一个或多个文件/文件夹路径
        paths = self.root.tk.splitlist(event.data)
        for path in paths:
            # 兼容大括号和空格
            clean_path = path.strip().strip('{}')
            if os.path.isdir(clean_path):
                print("拖入文件夹:", clean_path)
                word_files = self.find_word_files_in_dir(clean_path)
                if word_files:
                    word_files = self.filter_conflict_files(word_files, clean_path)
                    for full_path in word_files:
                        self.add_file_to_list(full_path)
            elif os.path.isfile(clean_path):
                print("拖入文件:", clean_path)
                if self.is_word_file(clean_path):
                    self.add_file_to_list(clean_path)
            else:
                print("未识别的拖拽路径:", clean_path)

    def choose_output_dir(self):
        dir_path = filedialog.askdirectory(title="选择输出目录")
        if dir_path:
            self.output_dir.set(dir_path)

    def toggle_auto_open(self):
        if self.auto_open_dir.get() == "1":
            self.auto_open_dir.set("0")
        else:
            self.auto_open_dir.set("1")
        self.update_auto_open_btn()

    def update_auto_open_btn(self):
        if self.auto_open_dir.get() == "1":
            self.auto_open_check.config(bg=COLORS['checkbox_on'], text="✔ 转换后自动打开输出目录")
        else:
            self.auto_open_check.config(bg=COLORS['checkbox_off'], text="✖ 转换后自动打开输出目录")

    def export_log(self):
        if not self.log_records:
            messagebox.showinfo("提示", "暂无可导出的日志！")
            return
        file_path = filedialog.asksaveasfilename(title="保存日志", defaultextension=".txt", filetypes=[("文本文件", "*.txt")])
        if file_path:
            try:
                with open(file_path, 'w', encoding='utf-8') as f:
                    f.write("\n".join(self.log_records))
                messagebox.showinfo("导出成功", f"日志已保存到: {file_path}")
            except Exception as e:
                messagebox.showerror("导出失败", f"日志导出失败: {e}")

    def show_listbox_menu(self, event):
        selected_indices = self.listbox.curselection()
        if not selected_indices:
            return
        menu = self.listbox_menu
        menu.post(event.x_root, event.y_root)

    def open_selected_file_dir(self):
        selected_indices = self.listbox.curselection()
        if not selected_indices:
            return
        selected_file = self.listbox.get(selected_indices[0])
        if os.path.isfile(selected_file):
            os.startfile(os.path.dirname(selected_file))

    def _now(self):
        return datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S")

    def run(self):
        self.root.mainloop()

if __name__ == "__main__":
    converter = WordToPDFConverter()
    converter.run()