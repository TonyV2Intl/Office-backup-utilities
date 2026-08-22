import win32com.client as win32   #导入win32com.client库，用于通过COM接口与Microsoft Office应用程序交互

import os   #导入os模块，用于处理文件和目录路径操作
import shutil   #导入shutil模块，用于复制文件并保留元数据
import time   #导入time模块，用于时间相关操作
import datetime   #导入datetime库，用于计算备份所用时间
from collections import defaultdict  #导入collections库的defaultdict方法，用于跟踪单个文件的跳过次数
import hashlib   #导入hashlib库，用于计算文件MD5值
import json  #导入json库，用于处理配置文件的读写
import ctypes   #导入ctypes库，用于调用Windows API函数
import subprocess  #导入subprocess模块，用于启动新进程
import sys  #导入sys模块，用于处理系统相关操作
import threading  #导入threading模块，用于多线程操作
import traceback  #导入traceback模块，用于获取详细的异常信息



#设定默认配置文件
default_config = {
    #指定备份路径，注意路径中的反斜杠需要转义
    "ppt_backup_path": "C:\\Officebackup\\pptbackup",   #PPT、WPS备份路径
    "word_backup_path": "C:\\Officebackup\\wordbackup",   #Word备份路径
    #指定轮询间隔
    "interval": 60,   #指定执行完一轮操作后等待的时间间隔，单位为秒（默认60秒）
    #功能开启或禁用
    "ppt_backup_enable": True,   #PPT备份功能
    "word_backup_enable": True,   #Word备份功能
    "wps_backup_enable": True,   #WPS备份功能
    #文件夹精确备份功能
    "accurate_backup_enable": False,
    "accurate_backup_source_path": "",
    "accurate_backup_target_path": "",
    #控制台行为与日志保存设置
    "show_console_window_at_startup": True,   #程序启动时显示控制台窗口，True为显示（默认），False为隐藏
    "save_log": True,   #是否保存日志到OBUlatest.log文件，True为保存（默认），False为不保存
    "archive_previous_log": True,   #是否在程序启动时归档之前的日志（重命名为OBUprevious.log），True为归档（默认），False为直接覆盖
    #超时设置
    "backup_timeout": 600,   #备份操作超时时间，单位为秒（默认10分钟）
}
startup_warnings = []   #初始化启动警告列表，日志系统准备好后再输出
config_file_path = 'OBU6.3Core.json'   #配置文件路径
try:   #读取配置文件
    with open('OBU6.3Core.json', 'r', encoding='utf-8') as f:   #尝试读取配置文件（只读）
        config = json.load(f)
    if not isinstance(config, dict):   #检查配置内容是否为JSON对象
        raise TypeError('configuration root must be a JSON object')
    config_changed = False
    for key, value in default_config.items():   #如果现有配置文件有缺漏，根据默认配置项自动补全
        if key not in config:
            config[key] = value
            config_changed = True
    if config_changed:   #如果配置文件有新增项，写回配置文件
        try:
            with open(config_file_path, 'w', encoding='utf-8') as f:
                json.dump(config, f, indent=4, ensure_ascii=False)
        except OSError as e:
            startup_warnings.append('Failed to update configuration file ' + config_file_path + ': ' + str(e))
except FileNotFoundError:   #若配置文件不存在，使用默认配置并尝试创建
    config = default_config.copy()   #使用默认配置
    try:
        with open(config_file_path, 'w', encoding='utf-8') as f:   #在当前目录下根据默认配置文件创建
            json.dump(config, f, indent=4, ensure_ascii=False)   #写入默认配置文件
    except OSError as e:
        startup_warnings.append('Failed to create default configuration file ' + config_file_path + ': ' + str(e))
except (json.JSONDecodeError, UnicodeDecodeError, OSError, TypeError) as e:   #配置损坏、不可读取或格式不正确
    config = default_config.copy()   #使用默认配置，但保留损坏文件
    backup_path = config_file_path + '.' + datetime.datetime.now().strftime('%Y%m%d%H%M%S') + '.bak'
    try:
        os.rename(config_file_path, backup_path)   #将损坏配置文件重命名为备份
        startup_warnings.append('Configuration file ' + config_file_path + ' is invalid (' + type(e).__name__ + ': ' + str(e) + '); preserved as ' + backup_path)
    except OSError as backup_error:
        startup_warnings.append('Configuration file ' + config_file_path + ' is invalid (' + type(e).__name__ + ': ' + str(e) + '); failed to preserve backup (' + str(backup_error) + ')')
    try:
        with open(config_file_path, 'w', encoding='utf-8') as f:   #尝试写入默认配置文件
            json.dump(config, f, indent=4, ensure_ascii=False)   #写入默认配置
    except OSError as write_error:
        startup_warnings.append('Failed to create replacement configuration file ' + config_file_path + ': ' + str(write_error))



log_file = None   #初始化日志文件句柄，日志打开失败时保持为空
file_logging_enabled = False   #标记本次会话是否启用文件日志
log_write_error_reported = False   #标记是否已经报告过日志写入失败
if config.get('save_log'):   #检查是否启用日志保存功能
    try:
        if os.path.exists('OBUlatest.log'):   #如果日志文件存在
            if config.get('archive_previous_log'):   #如果启用归档功能
                # 将旧日志重命名为OBUprevious.log
                if os.path.exists('OBUprevious.log'):
                    os.remove('OBUprevious.log')
                os.rename('OBUlatest.log', 'OBUprevious.log')
            else:   #如果禁用归档功能，直接删除旧日志
                os.remove('OBUlatest.log')
        log_file = open('OBUlatest.log', 'a', encoding='utf-8')   #以追加模式打开日志文件
        file_logging_enabled = True   #标记文件日志已启用
        # 写入版权信息和开始运行时间戳到控制台和日志文件
        header = 'Office Backup Utilities 6.3 Core\nCopyright (C) 2024-2026 TonyV2Intl\nSession starts at: ' + time.strftime('%Y-%m-%d %H:%M:%S')
        print(header + '\n')
        log_file.write(header + '\n\n')
        log_file.flush()   #刷新文件缓冲区，确保日志消息立即写入文件
    except (OSError, IOError) as e:
        startup_warnings.append('Failed to initialize file logging: ' + type(e).__name__ + ': ' + str(e) + '; continuing with console-only logging')
        config['save_log'] = False   #仅在内存中禁用文件日志
        log_file = None

def log_print(msg):   #定义日志打印函数
    global log_file, file_logging_enabled, log_write_error_reported   #声明日志状态变量
    global runid    #声明全局变量runid，以便在函数内修改其值
    runid+=1   #运行计数器累加
    log_msg= time.strftime('[%H:%M:%S-#') + str(runid) + '] ' + msg   # 打印带时间戳的日志消息到控制台
    print(log_msg)   # 打印日志消息到控制台
    if config.get('save_log') and file_logging_enabled and log_file is not None and not log_file.closed:   #如果启用日志保存功能，则将日志消息写入日志文件
        try:
            log_file.write(log_msg + '\n')   # 将日志消息写入日志文件
            log_file.flush()   #刷新文件缓冲区，确保日志消息立即写入文件
        except (OSError, IOError, ValueError) as e:
            file_logging_enabled = False   #写入失败后禁用本次会话的文件日志
            config['save_log'] = False   #仅在内存中禁用文件日志
            if not log_write_error_reported:
                log_write_error_reported = True
                print('[ERROR] File logging disabled after write failure: ' + type(e).__name__ + ': ' + str(e), file=sys.stderr)



# 根据配置显示或隐藏控制台窗口
console_visible = config.get('show_console_window_at_startup')   #获取控制台窗口初始状态参数
console_window = ctypes.windll.kernel32.GetConsoleWindow()   #获取控制台窗口句柄
if not console_visible:
    ctypes.windll.user32.ShowWindow(console_window, 0)   #隐藏控制台窗口



#初始化变量
runid=0   #初始化运行计数器
file_skip_count = defaultdict(int)   #使用字典记录每个文件的跳过次数（替代原全局skippedtime）
SaveAs_method_activated = defaultdict(bool)  # 使用字典记录每个文件是否已激活SaveAs方法
Existed_in_this_session = defaultdict(bool)  # 使用字典记录每个文件是否在本次运行中出现过，让之前会话中已经备份过的文件在程序重启后正常进行第一次备份
accurate_backup_running = False  # 精确备份线程运行标志

def log_exception(context, error):   #记录异常类型、消息和完整堆栈
    log_print(context + ': ' + type(error).__name__ + ': ' + str(error) + '\n' + traceback.format_exc())

def validate_positive_number(value, default_value, setting_name):   #校验必须为正数的配置项
    try:
        checked_value = float(value)
        if checked_value > 0:
            return checked_value
    except (TypeError, ValueError):
        pass
    log_print('Invalid ' + setting_name + ' value ' + repr(value) + ', using default ' + str(default_value))
    return default_value

for startup_warning in startup_warnings:   #回放配置和日志初始化阶段产生的启动警告
    log_print('STARTUP WARNING: ' + startup_warning)

#从配置文件读取变量
sleeptime=validate_positive_number(config.get('interval'), 60, 'interval')   #轮询间隔（默认为60秒）
config['interval'] = sleeptime   #仅在内存中保存校验后的轮询间隔
config['backup_timeout'] = validate_positive_number(config.get('backup_timeout'), 600, 'backup_timeout')   #备份超时时间（默认为600秒）
ppt_save_folder=config.get('ppt_backup_path')   #ppt备份路径
word_save_folder=config.get('word_backup_path')   #word备份路径



if config.get('accurate_backup_enable'):  # 检查精确备份功能是否启用
    source_path = config.get('accurate_backup_source_path')   #获取源文件夹路径
    target_path = config.get('accurate_backup_target_path')   #获取目标文件夹路径
    if not source_path or not target_path:   #如果精确备份功能开启但源路径为空或目标路径为空，则当前会话禁用精确备份功能
        log_print("Accurate backup disabled for this session, source path or target path for accurate backup is empty, please provide valid paths in the configuration file")
        config['accurate_backup_enable'] = False   #当前会话禁用（不修改配置文件）



# 超时装饰器函数 - 主线程执行函数，子线程计时
def timeout(seconds, config_key=None):
    def decorator(func):
        def wrapper(*args, **kwargs):
            timeout_value = seconds
            if config_key:
                timeout_value = config.get(config_key, seconds)
                timeout_value = validate_positive_number(timeout_value, seconds, config_key)
            
            # 超时标志
            timeout_occurred = [False]
            
            # 计时线程函数
            def timer_thread():
                time.sleep(timeout_value)
                if not timeout_occurred[0]:
                    timeout_occurred[0] = True
                    # 先打印超时信息（此时日志文件还未关闭）
                    log_print(f"Function {func.__name__} exceeded timeout of {timeout_value} seconds, restarting program")
                    # 关闭日志文件，解除占用
                    if 'log_file' in globals() and log_file is not None and not log_file.closed:
                        try:
                            log_file.close()
                            # 注意：关闭后不要再调用 log_print
                        except Exception as close_error:
                            print('[ERROR] Failed to close log before restart: ' + type(close_error).__name__ + ': ' + str(close_error) + '\n' + traceback.format_exc(), file=sys.stderr)
                    # 启动新实例
                    try:
                        # 获取当前运行的可执行文件路径
                        # sys.argv[0] 在打包为exe后指向实际的exe文件
                        current_exe = sys.argv[0]
                        if not os.path.isabs(current_exe):
                            current_exe = os.path.abspath(current_exe)
                        
                        # 启动新实例
                        subprocess.Popen([current_exe])
                        time.sleep(1)  # 给新进程启动时间
                    except Exception as e:
                        # 如果sys.argv[0]失败，尝试其他方法
                        primary_error_traceback = traceback.format_exc()
                        try:
                            # 尝试使用__file__（适用于未打包的情况）
                            script_path = os.path.abspath(__file__)
                            subprocess.Popen([sys.executable, script_path])
                            time.sleep(1)
                        except Exception as fallback_error:
                            print('[ERROR] Failed to restart program using primary and fallback paths.\nPrimary error: ' + type(e).__name__ + ': ' + str(e) + '\n' + primary_error_traceback + '\nFallback error: ' + type(fallback_error).__name__ + ': ' + str(fallback_error) + '\n' + traceback.format_exc(), file=sys.stderr)
                    # 强制退出
                    os._exit(1)
            
            # 启动计时线程
            timer = threading.Thread(target=timer_thread)
            timer.daemon = True  # 守护线程，主线程退出时自动退出
            timer.start()
            
            try:
                # 执行实际函数
                result = func(*args, **kwargs)
                # 标记执行完成
                timeout_occurred[0] = True
                return result
            except Exception as e:
                log_exception('Error in ' + func.__name__, e)
                timeout_occurred[0] = True
                return None
        return wrapper
    return decorator

# 计算文件MD5值的函数
def calculate_md5(file_path):  # 计算文件的MD5值
    hash_md5 = hashlib.md5()
    try:
        # 使用更大的块大小以提高性能
        with open(file_path, "rb") as f:
            for chunk in iter(lambda: f.read(8192), b""):  # 使用8192字节的块大小
                hash_md5.update(chunk)
        return hash_md5.hexdigest()
    except Exception as e:
        log_exception('Error calculating MD5 for ' + file_path, e)
        return None

# 辅助函数：检查并移除文件只读属性
def remove_readonly(file_path):
    try:
        if os.path.exists(file_path):
            # 检查文件是否只读
            attrs = os.stat(file_path).st_mode
            # 移除只读属性
            if not (attrs & 0o200):  # 如果是只读的
                os.chmod(file_path, attrs | 0o200)  # 添加写入权限
                log_print('Removed readonly attribute from ' + file_path)
    except Exception as e:
        log_exception('Error removing readonly from ' + file_path, e)

def _backup_open_files(save_folder, app_progid, use_get_object, collection_attr, file_type_label):   #通用备份函数，参数化不同Office应用的差异
    try:   #开始异常捕获
        if not os.path.exists(save_folder):   #检查备份目录是否存在
            os.makedirs(save_folder)   #若不存在则创建备份目录（包括所有必要的父目录）
            log_print('Target backup folder not found, created: ' + save_folder + ' successfully')   #打印成功创建备份目录的信息
        
        if use_get_object:   #判断是否使用GetObject方式获取COM对象（WPS用）
            app = win32.GetObject(Class=app_progid)   #使用GetObject获取已运行的Office应用实例
        else:   #使用Dispatch方式（PPT/Word用）
            app = win32.Dispatch(app_progid)   #启动或连接到Office应用实例
        file_collection = getattr(app, collection_attr)   #获取应用中所有打开的文档集合

        any_backup_performed = False   #标记本轮是否有任何备份操作
        
        for target_file in file_collection:   #遍历所有打开的文档
            try:
                target_file_path = target_file.FullName   #获取文件的完整路径
                target_file_name = os.path.basename(target_file_path)   #提取文件名
                backup_file_path = os.path.join(save_folder, target_file_name)   #生成备份文件路径

                if os.path.exists(backup_file_path):   #检查备份文件是否已存在
                    if SaveAs_method_activated[target_file_name] == True:   #如果SaveAs方法已被激活
                        log_print(target_file_name + ' has already existed in ' + save_folder + ', skipped backup (SaveAs method activated)')   #打印跳过信息
                        continue   #跳过此次备份

                    original_md5 = calculate_md5(target_file_path)   #计算源文件的MD5值
                    backup_md5 = calculate_md5(backup_file_path)   #计算备份文件的MD5值

                    if original_md5 and backup_md5 and original_md5 == backup_md5:   #两个MD5都成功计算且相同
                        log_print(target_file_name + ' has already existed in ' + save_folder + ', skipped backup (MD5 match)')   #打印跳过信息
                        continue   #跳过此次备份
                    elif original_md5 is None:   #源文件MD5计算失败（可能文件找不到）
                        log_print(target_file_name + ' source file not found, skipping this backup')   #打印跳过信息
                        continue   #跳过此次备份
                    else:   #MD5值不同，文件已修改
                        log_print(target_file_name + ' has changed, backup will begin soon (MD5 mismatch)')   #打印文件变更信息
                
                Existed_in_this_session[target_file_name] = True   #标记该文件在本次会话中出现过
                log_print('Start to backup ' + target_file_name + ' to ' + save_folder)   #打印备份开始信息
                remove_readonly(backup_file_path)   #如果目标文件存在，先移除只读属性
                copy_start_time = datetime.datetime.now()   #记录复制操作开始时间
                shutil.copy2(target_file_path, backup_file_path)   #复制文件到备份文件夹，保留元数据
                copy_end_time = datetime.datetime.now()   #记录复制操作结束时间
                copy_used_time = copy_end_time - copy_start_time   #计算复制所用时间

                current_time = time.time()   #获取当前时间
                os.utime(backup_file_path, (os.path.getatime(backup_file_path), current_time))   #设置修改时间为备份发生的时间

                file_skip_count[target_file_name] = 0   #重置该文件的跳过计数器
                any_backup_performed = True   #标记本轮有备份操作
                log_print('Successfully backuped ' + target_file_name + ' to ' + save_folder + ' in ' + str(copy_used_time) + ' s')   #打印备份成功信息
            except Exception as e:
                try:
                    failed_file_name = target_file.FullName
                except Exception:
                    failed_file_name = '<unknown document>'
                log_exception('Failed to backup open ' + file_type_label + ' ' + str(failed_file_name), e)

        if not any_backup_performed and len(file_collection) == 0:   #没有可备份文件
            log_print('No ' + file_type_label + ' available now (Normal request)')   #打印无文件信息

    except FileNotFoundError as e:   #捕获移动存储介质移除导致的文件未找到异常，使用SaveAs方法备份
        if 'file_collection' not in locals():   #如果获取文档集合前就失败
            log_exception('File not found before opening ' + file_type_label + ' document collection', e)
            return   #无法使用SaveAs方式继续
        if not os.path.exists(save_folder):   #检查备份目录是否存在
            try:
                os.makedirs(save_folder)   #若不存在则创建备份目录
                log_print('Target backup folder not found, created: ' + save_folder + ' successfully')   #打印创建成功信息
            except Exception as folder_error:
                log_exception('Failed to create target backup folder ' + save_folder, folder_error)
                return

        for idx in range(1, file_collection.Count + 1):   #遍历文档集合（使用索引方式）
            try:
                target_file = file_collection.Item(idx)   #获取当前文档对象
                target_file_path = target_file.FullName   #获取文件完整路径
                target_file_name = os.path.basename(target_file_path)   #提取文件名
                backup_file_path = os.path.join(save_folder, target_file_name)   #生成备份路径
                log_print('Start to backup ' + target_file_name + ' to ' + save_folder + ' using SaveAs method')   #打印SaveAs备份开始信息
                remove_readonly(backup_file_path)   #移除目标文件只读属性
                save_start_time = datetime.datetime.now()   #记录保存操作开始时间
                target_file.SaveAs(backup_file_path)   #使用SaveAs方法保存文档到指定路径
                save_end_time = datetime.datetime.now()   #记录保存操作结束时间
                save_used_time = save_end_time - save_start_time   #计算保存所用时间
                SaveAs_method_activated[target_file_name] = True   #标记该文件已激活SaveAs方法
                log_print('Detected access control, activated SaveAs method, successfully backuped ' + target_file_name + ' to ' + save_folder + ' in ' + str(save_used_time) + ' s')   #打印SaveAs备份成功信息
            except Exception as save_error:
                try:
                    failed_file_name = target_file.FullName
                except Exception:
                    failed_file_name = '<unknown document>'
                log_exception('Failed to SaveAs open ' + file_type_label + ' ' + str(failed_file_name), save_error)
                
    except Exception as e:   #捕获其他所有异常
        if type(e).__name__ == 'com_error':   #如果是COM错误（无打开的应用实例）
            log_print('No ' + file_type_label + ' available now (' + app_progid.split('.')[0] + ' application not detected)')   #打印应用未检测到信息
        else:   #其他类型的错误
            log_exception('Exception while backing up open ' + file_type_label, e)


@timeout(seconds=600, config_key='backup_timeout')  # 添加10分钟超时机制
def save_open_ppt_files(ppt_save_folder):   #定义ppt保存函数，参数ppt_save_folder是备份文件的存储路径
    _backup_open_files(
        save_folder=ppt_save_folder,
        app_progid='PowerPoint.Application',
        use_get_object=False,
        collection_attr='Presentations',
        file_type_label='ppt'
    )




@timeout(seconds=600, config_key='backup_timeout')  # 添加10分钟超时机制
def save_open_word_files(word_save_folder):   #定义word保存函数，参数word_save_folder是备份文件的存储路径
    _backup_open_files(
        save_folder=word_save_folder,
        app_progid='Word.Application',
        use_get_object=False,
        collection_attr='Documents',
        file_type_label='doc'
    )




@timeout(seconds=600, config_key='backup_timeout')  # 添加10分钟超时机制
def save_open_WPS_files(ppt_save_folder):   #定义WPS保存函数，参数ppt_save_folder是备份文件的存储路径
    _backup_open_files(
        save_folder=ppt_save_folder,
        app_progid='KWPP.Application',
        use_get_object=True,
        collection_attr='Presentations',
        file_type_label='WPS ppt'
    )




def accurate_backup():   #定义精确备份函数
    global accurate_backup_running
    accurate_backup_running = True
    try:
        source_path = config.get('accurate_backup_source_path')
        target_path = config.get('accurate_backup_target_path')
        if not source_path or not target_path:
            log_print('Accurate backup source or target path is empty, skipped')
            return
        if os.path.exists(source_path):   #检查源文件夹是否存在
            log_print('Start accurate backup from ' + source_path + ' to ' + target_path)   #打印精确备份开始信息
            copy_start_time=datetime.datetime.now()   #记录复制操作开始时间
            shutil.copytree(source_path, target_path, dirs_exist_ok=True)  # 复制源文件夹及其内容到目标文件夹
            copy_end_time=datetime.datetime.now()   #记录复制操作结束时间
            copy_used_time=copy_end_time-copy_start_time  #计算复制所用时间
            log_print(f'Accurate backup completed successfully from {source_path} to {target_path} in {copy_used_time} s')  # 打印精确备份完成信息
            
            config['accurate_backup_enable'] = False   #当前会话禁用精确备份功能
            try:
                with open('OBU6.3Core.json', 'w', encoding='utf-8') as f:
                    json.dump(config, f, indent=4, ensure_ascii=False)
                log_print('Accurate backup disabled after successful backup')
            except Exception as e:
                log_exception('Failed to update config file', e)
        else:
            log_print('Source path for accurate backup does not exist: ' + source_path + ', wait for the next request')  # 打印源文件夹不存在信息，等待下次请求
    except Exception as e:
        log_exception('Accurate backup failed', e)  # 打印精确备份失败信息
    finally:
        accurate_backup_running = False





def global_exception_handler(exctype, value, tb):   #处理全局未捕获异常
    if issubclass(exctype, KeyboardInterrupt):   #正常响应Ctrl+C
        sys.__excepthook__(exctype, value, tb)
        return
    error_msg = ''.join(traceback.format_exception(exctype, value, tb))
    try:
        log_print('[ERROR] ' + error_msg)
    except Exception as log_error:
        print('[ERROR] Failed to record uncaught exception: ' + type(log_error).__name__ + ': ' + str(log_error) + '\n' + traceback.format_exc(), file=sys.stderr)

def threading_exception_handler(args):   #处理守护线程未捕获异常
    if issubclass(args.exc_type, KeyboardInterrupt):   #正常响应Ctrl+C
        threading.__excepthook__(args)
        return
    try:
        global_exception_handler(args.exc_type, args.exc_value, args.exc_traceback)
    except Exception as e:
        print('[ERROR] Thread exception handler failed: ' + type(e).__name__ + ': ' + str(e) + '\n' + traceback.format_exc(), file=sys.stderr)

sys.excepthook = global_exception_handler   #注册主线程异常处理器
threading.excepthook = threading_exception_handler   #注册守护线程异常处理器

print('Program initialization completed, entering main loop\n')
log_print('Program initialization completed, entering main loop')

while True:   #主线程无限循环，防止程序退出
    if config.get('ppt_backup_enable'):   #检查PPT备份功能是否启用
        save_open_ppt_files(ppt_save_folder)   #启动线程
    if config.get('word_backup_enable'):   #检查Word备份功能是否启用
        save_open_word_files(word_save_folder)   #启动线程
    if config.get('wps_backup_enable'):   #检查WPS备份功能是否启用
        save_open_WPS_files(ppt_save_folder)   #启动线程
    if config.get('accurate_backup_enable') and not accurate_backup_running:  # 检查精确备份功能是否启用且未在运行
        backup_thread = threading.Thread(target=accurate_backup)
        backup_thread.daemon = True
        backup_thread.start()
    time.sleep(sleeptime)   #等待指定时间后继续轮询
