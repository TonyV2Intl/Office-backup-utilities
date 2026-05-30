import win32com.client as win32   #导入win32com.client库，用于通过COM接口与Microsoft Office应用程序交互
import pystray   #导入pystray库，用于创建系统托盘图标
from pystray import MenuItem as item   #从pystray库中导入MenuItem类，用于创建托盘菜单项
from PIL import Image   #导入PIL库的Image模块，用于处理图标图像（pystray需要）

import os   #导入os模块，用于处理文件和目录路径操作
import shutil   #导入shutil模块，用于复制文件并保留元数据
import time   #导入time模块，用于时间相关操作
import datetime   #导入datetime库，用于计算备份所用时间
from collections import defaultdict  #导入collections库的defaultdict方法，用于跟踪单个文件的跳过次数
import hashlib   #导入hashlib库，用于计算文件MD5值
import sys   #导入sys模块，用于处理系统相关操作
import traceback   #导入traceback模块，用于获取详细的异常信息
import threading  #导入threading库，用于多线程操作
import json  #导入json库，用于处理配置文件的读写
import ctypes   #导入ctypes库，用于调用Windows API函数
import subprocess  #导入subprocess模块，用于启动新进程
import asyncio  #导入asyncio模块，用于异步操作





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
    "upload_to_openlist_enable": True,   #上传到OpenList功能
    #OpenList参数
    "openlist_url": "",   #OpenList服务器URL
    "openlist_username": "",   #OpenList用户名
    "openlist_password": "",   #OpenList密码
    "openlist_target_folder": "",   #目标文件夹路径，根目录用"/"表示
    #文件夹精确备份功能
    "accurate_backup_enable": False,
    "accurate_backup_source_path": "",
    "accurate_backup_target_path": "",
    #托盘图标、控制台行为与日志保存设置
    #"tray_left_click_behavior": "open_console",   #托盘图标左键点击行为，选项有"open_console"（打开控制台）和"exit_program"（退出程序）（无法生效）
    "hide_tray_icon": False,   #是否隐藏托盘图标，True为隐藏，False为显示（默认）
    "show_console_window_at_startup": False,   #程序启动时显示控制台窗口，True为显示，False为隐藏（默认）
    "save_log": True,   #是否保存日志到OBUlatest.log文件，True为保存（默认），False为不保存
    "archive_previous_log": True,   #是否在程序启动时归档之前的日志（重命名为OBUprevious.log），True为归档（默认），False为直接覆盖
    "log_abnormal_upload": True,   #是否记录上传异常的文件到OBUabnormal.txt，True为记录（默认），False为不记录
    #超时和重试设置
    "backup_timeout": 600,   #备份操作超时时间，单位为秒（默认10分钟）
    "upload_retry_wait": 30,   #上传重试等待时间，单位为秒（默认30秒）
    "upload_max_retries": ""   #上传最大重试次数，默认为空，表示无限次重试
}
try:   #读取配置文件
    with open('OBU6.1.json', 'r', encoding='utf-8') as f:   #尝试读取配置文件（只读）
        config = json.load(f)
    config_changed = False
    for key, value in default_config.items():   #如果现有配置文件有缺漏，根据默认配置项自动补全
        if key not in config:
            config[key] = value
            config_changed = True
    if config_changed:   #如果配置文件有新增项，写回配置文件
        with open('OBU6.1.json', 'w', encoding='utf-8') as f:
            json.dump(config, f, indent=4, ensure_ascii=False)
except (FileNotFoundError, json.JSONDecodeError):   #若配置文件不存在或无法解析
    config = default_config   #使用默认配置
    with open('OBU6.1.json', 'w', encoding='utf-8') as f:   #在当前目录下根据默认配置文件创建（写入）
        json.dump(config, f, indent=4, ensure_ascii=False)   #写入默认配置文件



if config.get('save_log'):   #检查是否启用日志保存功能
    if os.path.exists('OBUlatest.log'):   #如果日志文件存在
        if config.get('archive_previous_log'):   #如果启用归档功能
            # 将旧日志重命名为OBUprevious.log
            if os.path.exists('OBUprevious.log'):
                os.remove('OBUprevious.log')
            os.rename('OBUlatest.log', 'OBUprevious.log')
        else:   #如果禁用归档功能，直接删除旧日志
            os.remove('OBUlatest.log')
    log_file = open('OBUlatest.log', 'a', encoding='utf-8')   #以追加模式打开日志文件
    # 写入版权信息和开始运行时间戳到控制台和日志文件
    header = 'Office Backup Utilities 6.1\nCopyright (C) 2024-2026 TonyV2Intl\nSession starts at: ' + time.strftime('%Y-%m-%d %H:%M:%S')
    print(header + '\n')
    log_file.write(header + '\n\n')
    log_file.flush()   #刷新文件缓冲区，确保日志消息立即写入文件

def log_print(msg, source='main'):   #定义日志打印函数
    global runid    #声明全局变量runid，以便在函数内修改其值
    runid+=1   #运行计数器累加
    log_msg= time.strftime('[%H:%M:%S-#') + str(runid) + '-' + source + '] ' + msg   # 打印带时间戳和来源的日志消息到控制台
    print(log_msg)   # 打印日志消息到控制台
    if config.get('save_log'):   #如果启用日志保存功能，则将日志消息写入日志文件
        log_file.write(log_msg + '\n')   # 将日志消息写入日志文件
        log_file.flush()   #刷新文件缓冲区，确保日志消息立即写入文件



console_visible = config.get('show_console_window_at_startup')   #获取控制台窗口初始状态参数（默认为隐藏）
console_window = ctypes.windll.kernel32.GetConsoleWindow()   #获取控制台窗口句柄
if not console_visible:
    ctypes.windll.user32.ShowWindow(console_window, 0)   #隐藏控制台窗口



#初始化变量
runid=0   #初始化运行计数器
file_skip_count = defaultdict(int)   #使用字典记录每个文件的跳过次数（替代原全局skippedtime）
SaveAs_method_activated = defaultdict(bool)  # 使用字典记录每个文件是否已激活SaveAs方法
Existed_in_this_session = defaultdict(bool)  # 使用字典记录每个文件是否在本次运行中出现过，让之前会话中已经备份过的文件在程序重启后正常进行第一次备份
upload_queue = []  # 初始化上传队列
#从配置文件读取变量
sleeptime=config.get('interval')   #轮询间隔（默认为60秒）
ppt_save_folder=config.get('ppt_backup_path')   #ppt备份路径
word_save_folder=config.get('word_backup_path')   #word备份路径
'''behavior = config.get('tray_left_click_behavior')  # 托盘图标左键点击行为（默认为打开控制台）（无法生效）'''



try:   #尝试导入AList3SDK
    from alist import AListUser, AListAsync   #导入AList3SDK，用于与OpenList/AList服务交互（1.3.1版本默认是同步API，需要指定使用AListAsync类已进行异步操作）
except ImportError:
    log_print("alist3 not found, upload function disabled for this session")
    config['upload_to_openlist_enable'] = False   #当前会话禁用上传功能（不修改配置文件）

# 从配置文件读取OpenList变量
openlist_url = config.get('openlist_url')
openlist_username = config.get('openlist_username')
openlist_password = config.get('openlist_password')
openlist_target_folder = config.get('openlist_target_folder')



def upload_to_openlist_thread():   #在单独线程中执行上传操作
    retry_count = 0
    max_retries = config.get('upload_max_retries')
    if max_retries is None or max_retries == "":
        max_retries = float('inf')  # 无限次重试
    else:
        max_retries = int(max_retries)

    upload_retry_wait = config.get('upload_retry_wait', 30)
    if upload_retry_wait is None or upload_retry_wait == "":
        upload_retry_wait = 30

    while True:
        # 当队列中有文件且上传功能启用时执行上传
        if not upload_queue or not config.get('upload_to_openlist_enable'):
            log_print('Upload queue is empty or upload disabled, upload thread exit', source='openlist')
            break

        log_print('Upload thread started, processing ' + str(len(upload_queue)) + ' file(s)', source='openlist')

        # 复制队列以避免在迭代时修改
        current_queue = list(upload_queue)
        failed_in_this_round = False

        for (upload_file, upload_source_path) in current_queue:
            log_print('Start to upload ' + upload_file + ' to OpenList', source='openlist')   #打印上传开始信息
            upload_start_time=datetime.datetime.now()   #记录上传操作开始时间
            
            try:
                # 定义异步上传函数
                async def async_upload():
                    # 初始化AList客户端和用户（使用异步API）
                    user = AListUser(openlist_username, openlist_password)
                    client = AListAsync(openlist_url)
                    upload_result = False
                    
                    # 构造目标文件路径
                    target_file_path = os.path.join(openlist_target_folder, upload_file).replace(os.sep, '/')
                    
                    # 登录
                    login_result = await client.login(user)
                    if login_result:
                        log_print('Login to OpenList successfully', source='openlist')
                    else:
                        log_print('Login to OpenList failed', source='openlist')
                        return False
                    
                    # 检查目标文件夹是否有效（使用mkdir测试）
                    try:
                        await client.mkdir(openlist_target_folder)
                        log_print('Target folder validated: ' + openlist_target_folder, source='openlist')
                    except Exception as e:
                        # 说明路径无效，当前会话禁用上传功能
                        log_print('Target folder invalid: ' + openlist_target_folder + ', error: ' + str(e), source='openlist')
                        log_print('Upload function disabled for this session, please check target folder path  is valid in the configuration file', source='openlist')
                        config['upload_to_openlist_enable'] = False   #当前会话禁用（不修改配置文件）
                        return False
                    
                    # 检查文件是否已存在，存在则删除
                    try:
                        await client.remove(target_file_path)
                        log_print('Existing file in OpenList deleted successfully: ' + upload_file, source='openlist')
                    except Exception as e:
                        log_print('No matching file found in OpenList or delete failed: ' + str(e) + ', upload will continue', source='openlist')
                    
                    # 上传文件
                    log_print('Uploading: ' + upload_file + ', file size: ' + str(round(os.path.getsize(upload_source_path) / 1024, 2)) + ' KB', source='openlist')
                    
                    # 使用 client.upload 方法，添加错误处理
                    try:
                        upload_result = await client.upload(target_file_path, upload_source_path)
                        if upload_result:
                            log_print('Upload to OpenList successfully: ' + upload_file, source='openlist')
                        else:
                            log_print('Upload to OpenList failed', source='openlist')
                    except Exception as e:
                        error_str = str(e)
                        log_print('Upload failed: ' + error_str, source='openlist')
                        log_print('Upload failed, will retry in next upload', source='openlist')
                    
                    return upload_result
                
                # 运行异步上传函数
                upload_result = asyncio.run(async_upload())
                
                # 检查上传是否成功
                if upload_result:
                    # 上传成功，从队列中移除文件
                    upload_end_time=datetime.datetime.now()   #记录上传操作结束时间
                    upload_used_time=upload_end_time-upload_start_time   #计算上传所用时间
                    log_print('Upload to OpenList finished: ' + upload_file + ' in ' + str(upload_used_time) + ' s', source='openlist')
                    if (upload_file, upload_source_path) in upload_queue:
                        upload_queue.remove((upload_file, upload_source_path))
                else:
                    # 上传失败，保留文件在队列中，等待下次上传
                    upload_end_time=datetime.datetime.now()   #记录上传操作结束时间
                    upload_used_time=upload_end_time-upload_start_time   #计算上传所用时间
                    log_print('Upload to OpenList failed: ' + upload_file + ' in ' + str(upload_used_time) + ' s, will retry in next upload', source='openlist')
                    failed_in_this_round = True
                
            except Exception as e:
                error_str = str(e)
                log_print('Upload to OpenList failed: ' + error_str, source='openlist')
                log_print('Traceback: ' + traceback.format_exc(), source='openlist')
                # 记录异常文件到OBUabnormal.txt
                if config.get('log_abnormal_upload'):
                    with open('OBUabnormal.txt', 'a', encoding='utf-8') as f:
                        f.write(f"{time.strftime('%Y-%m-%d %H:%M:%S')} - {upload_file} - {error_str}\n")
                # 发生错误时，保留文件在队列中，等待下次上传
                upload_end_time=datetime.datetime.now()   #记录上传操作结束时间
                upload_used_time=upload_end_time-upload_start_time   #计算上传所用时间
                log_print('Upload to OpenList failed: ' + upload_file + ' in ' + str(upload_used_time) + ' s, will retry in next upload', source='openlist')
                failed_in_this_round = True
        
        # 处理完当前队列后，检查队列是否为空（所有文件都处理完了）
        if not upload_queue:
            log_print('All files uploaded successfully, upload thread exit', source='openlist')
            break

        # 检查队列是否还有文件（上传失败的）
        if failed_in_this_round:
            retry_count += 1
            if max_retries != float('inf') and retry_count >= max_retries:
                # 达到最大重试次数，放弃剩余文件
                log_print('Max retries reached, ' + str(len(upload_queue)) + ' file(s) failed to upload', source='openlist')
                break
            if max_retries == float('inf'):
                log_print('Some files failed to upload, waiting ' + str(upload_retry_wait) + ' seconds before retry (retry ' + str(retry_count) + ')', source='openlist')
            else:
                log_print('Some files failed to upload, waiting ' + str(upload_retry_wait) + ' seconds before retry (' + str(retry_count) + '/' + str(int(max_retries)) + ')', source='openlist')
            time.sleep(upload_retry_wait)  # 等待后重试
        else:
            # 队列非空但没有失败的文件（新添加的文件），继续处理
            log_print('New files detected in queue, continuing upload', source='openlist')
    
    log_print('Upload thread finished', source='openlist')

def upload_to_openlist():   #启动上传线程
    if not config.get('upload_to_openlist_enable'):   #检查上传功能是否启用
        return   #如果未启用，直接返回
    
    # 只有当队列中有文件且没有上传线程在运行时才启动
    global upload_thread
    if upload_queue and ('upload_thread' not in globals() or not upload_thread.is_alive()):
        # 创建并启动上传线程
        upload_thread = threading.Thread(target=upload_to_openlist_thread)
        log_print('Upload thread starting', source='openlist')
        upload_thread.daemon = True  # 设置为守护线程，随主程序终止而结束
        upload_thread.start()


if config.get('upload_to_openlist_enable'):   #检查上传功能是否启用
    if not openlist_url or not openlist_username or not openlist_target_folder:   #检查OpenList配置是否完整（密码可为空）
        log_print('OpenList URL, username or target folder is empty, force disabled upload function, please provide valid credentials in the configuration file')
        config['upload_to_openlist_enable'] = False   #强制禁用上传功能
    else:
        # 启动上传线程
        upload_to_openlist()

if config.get('accurate_backup_enable'):  # 检查精确备份功能是否启用
    source_path = config.get('accurate_backup_source_path')   #获取源文件夹路径
    target_path = config.get('accurate_backup_target_path')   #获取目标文件夹路径
    if not source_path and target_path:   #如果精确备份功能开启但源路径为空或目标路径为空，则当前会话禁用精确备份功能
        log_print("Accurate backup disabled for this session, source path or target path for accurate backup is empty, please provide valid paths in the configuration file")
        config['accurate_backup_enable'] = False   #当前会话禁用（不修改配置文件）



# 超时装饰器函数 - 主线程执行函数，子线程计时
def timeout(seconds, config_key=None):
    def decorator(func):
        def wrapper(*args, **kwargs):
            timeout_value = seconds
            if config_key:
                timeout_value = config.get(config_key, seconds)
                if timeout_value is None or timeout_value == "":
                    timeout_value = seconds
            
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
                    if 'log_file' in globals():
                        try:
                            log_file.close()
                            # 注意：关闭后不要再调用 log_print
                        except:
                            pass
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
                        try:
                            # 尝试使用__file__（适用于未打包的情况）
                            script_path = os.path.abspath(__file__)
                            subprocess.Popen([sys.executable, script_path])
                            time.sleep(1)
                        except:
                            pass
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
                log_print(f"Error in {func.__name__}: {str(e)}")
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
        log_print('Error calculating MD5 for ' + file_path + ': ' + str(e))
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
                log_print(f"Removed readonly attribute from {file_path}")
    except Exception as e:
        log_print(f"Error removing readonly from {file_path}: {e}")





@timeout(seconds=600, config_key='backup_timeout')  #添加10分钟超时机制
def save_open_ppt_files(ppt_save_folder):   #定义ppt保存函数，参数ppt_save_folder是备份文件的存储路径
    global upload_queue  # 声明全局上传队列变量
    try:
        if not os.path.exists(ppt_save_folder):   #检查ppt备份目录是否存在
            os.makedirs(ppt_save_folder)   #若不存在则创建备份目录（包括所有必要的父目录）
            log_print('Target backup folder not found, created: ' + ppt_save_folder + ' successfully')   #打印成功创建ppt备份目录的信息
        
        ppt_app=win32.Dispatch('PowerPoint.Application')   #启动一个PowerPoint实例
        presentations = ppt_app.Presentations   #获取当前PowerPoint实例中所有打开的演示文稿集合

        any_backup_performed = False   #标记本轮是否有任何备份操作（替代原haveppt）
        
        for ppt in presentations:   #遍历集合
            ppt_path = ppt.FullName   #获取PPT文件的完整路径
            ppt_name = os.path.basename(ppt_path)   #提取文件名
            new_ppt_path = os.path.join(ppt_save_folder, ppt_name)   #生成备份路径

            if os.path.exists(new_ppt_path):   #检查备份文件是否已存在
                if SaveAs_method_activated[ppt_name] == True:   #如果SaveAs方法已被激活，则不再使用复制方法
                    log_print(ppt_name + ' has already existed in ' + ppt_save_folder + ', skipped backup (SaveAs method activated)')   #打印跳过信息
                    continue   #跳过此次备份
                
                # 计算原始文件和备份文件的MD5值
                original_md5 = calculate_md5(ppt_path)
                backup_md5 = calculate_md5(new_ppt_path)
                
                # 只有当两个MD5都成功计算且相同时才跳过
                if original_md5 and backup_md5 and original_md5 == backup_md5:
                    # MD5值相同，跳过备份
                    log_print(ppt_name + ' has already existed in ' + ppt_save_folder + ', skipped backup (MD5 match)')   #打印跳过信息
                    continue   #跳过此次备份
                elif original_md5 is None:
                    # 源文件找不到，跳过这次备份
                    log_print(ppt_name + ' source file not found, skipping this backup')
                    continue
                else:
                    # MD5值不同，需要备份
                    log_print(ppt_name + ' has changed, backup will begin soon (MD5 mismatch)')
            
            Existed_in_this_session[ppt_name] = True   #标记该文件在本次会话中出现过
            log_print('Start to backup ' + ppt_name + ' to ' + ppt_save_folder)   #打印备份开始信息
            # 如果目标文件存在，先移除只读属性
            remove_readonly(new_ppt_path)
            copy_start_time=datetime.datetime.now()   #记录复制操作开始时间
            shutil.copy2(ppt_path, new_ppt_path)   #复制PPT到备份文件夹，并尝试保留元数据（如修改时间等）
            copy_end_time=datetime.datetime.now()   #记录复制操作结束时间
            copy_used_time=copy_end_time-copy_start_time  #计算复制所用时间

            modified_time=os.path.getmtime(new_ppt_path)   #获取备份文件的修改时间
            current_time=time.time()   #获取当前时间
            os.utime(new_ppt_path, (modified_time, current_time))   #将 修改时间 存储到 访问时间（参数1），将 当前系统时间 设为 修改时间（参数2），方便文件系统根据修改时间排序

            file_skip_count[ppt_name] = 0   #重置该文件的跳过计数器
            any_backup_performed = True   #标记本轮有备份操作
            log_print(f'Successfully backuped {ppt_name} to {ppt_save_folder} in {copy_used_time} s')   #打印备份成功信息

            upload_queue.append((ppt_name,new_ppt_path))  # 将文件名和备份路径添加到上传队列
        upload_to_openlist()  # 启动上传线程

        if not any_backup_performed and len(presentations) == 0:   #检查变量值，如果没有可备份PPT，打印此条信息
            log_print('No ppt available now (Normal request)')   #打印运行信息

    except FileNotFoundError:   #捕获由于U盘等移动存储介质被移除而导致的“文件未找到”异常，使用2.0版本中的SaveAs方法进行备份
        if not os.path.exists(ppt_save_folder):   #检查ppt备份目录是否存在
            os.makedirs(ppt_save_folder)   #若不存在则创建备份目录（包括所有必要的父目录）
            log_print('Target backup folder not found, created: ' + ppt_save_folder + ' successfully')   #打印成功创建ppt备份目录的信息

        for idx in range(1, presentations.Count + 1):   #遍历PPT实例集合
            ppt = presentations.Item(idx)   #获取当前PPT实例
            ppt_path = ppt.FullName   #获取PPT文件的完整路径
            ppt_name = os.path.basename(ppt_path)   #提取文件名
            new_ppt_path = os.path.join(ppt_save_folder, ppt_name)   #生成备份路径
            log_print('Start to backup ' + ppt_name + ' to ' + ppt_save_folder + ' using SaveAs method')   #打印备份开始信息
            # 如果目标文件存在，先移除只读属性
            remove_readonly(new_ppt_path)
            savestarttime=datetime.datetime.now()   #记录保存操作开始时间
            ppt.SaveAs(new_ppt_path)   #使用SaveAs方法保存当前PPT实例到指定路径
            saveendtime=datetime.datetime.now()   #记录保存操作结束时间
            saveusedtime=saveendtime-savestarttime  #计算保存所用时间
            SaveAs_method_activated[ppt_name] = True   #标记该文件已激活SaveAs方法，后续不再备份
            log_print('Detected access control, activated SaveAs method, successfully backuped ' + ppt_name + ' to ' + ppt_save_folder + ' in ' + str(saveusedtime) + ' s')   #打印备份成功信息

            upload_queue.append((ppt_name,new_ppt_path))  # 将文件名和备份路径添加到上传队列
            upload_to_openlist()  # 启动上传线程       
    except Exception as e:   #获取其他错误类型
            if type(e).__name__=='com_error':   #捕获无打开的PowerPoint实例而产生的的异常
                log_print('No ppt available now (PowerPoint application not detected)')   #打印异常信息
            else:   #打印出其他错误并继续轮询
                log_print('Exception: ' + type(e).__name__ + ', request continue')   #打印异常信息



@timeout(seconds=600, config_key='backup_timeout')   #添加10分钟超时机制
def save_open_word_files(word_save_folder):   #定义word保存函数，参数word_save_folder是备份文件的存储路径
    global upload_queue  # 声明全局上传队列变量
    try:
        if not os.path.exists(word_save_folder):   #检查word备份目录是否存在
            os.makedirs(word_save_folder)   #若不存在则创建备份目录（包括所有必要的父目录）
            log_print('Target backup folder not found, created: ' + word_save_folder + ' successfully')   #打印成功创建word备份目录的信息
        
        word_app = win32.Dispatch('Word.Application')   #启动一个Word实例，若启用独立实例则无法获取当前已经打开的Word实例信息
        documents = word_app.Documents   #获取当前Word实例中所有打开的文档集合

        any_backup_performed = False   #标记本轮是否有任何备份操作（替代原havedoc）
            
        for doc in documents:   #遍历集合
            doc_path = doc.FullName   #获取Word文件的完整路径
            doc_name = os.path.basename(doc_path)   #提取文件名
            new_doc_path = os.path.join(word_save_folder, doc_name)   #生成备份路径

            if os.path.exists(new_doc_path):   #检查备份文件是否已存在
                if SaveAs_method_activated[doc_name] == True:   #如果SaveAs方法已被激活，则不再使用复制方法
                    log_print(doc_name + ' has already existed in ' + word_save_folder + ', skipped backup (SaveAs method activated)')   #打印跳过信息
                    continue   #跳过此次备份
                
                # 计算原始文件和备份文件的MD5值
                original_md5 = calculate_md5(doc_path)
                backup_md5 = calculate_md5(new_doc_path)
                
                # 只有当两个MD5都成功计算且相同时才跳过
                if original_md5 and backup_md5 and original_md5 == backup_md5:
                    # MD5值相同，跳过备份
                    log_print(doc_name + ' has already existed in ' + word_save_folder + ', skipped backup (MD5 match)')   #打印跳过信息
                    continue   #跳过此次备份
                elif original_md5 is None:
                    # 源文件找不到，跳过这次备份
                    log_print(doc_name + ' source file not found, skipping this backup')
                    continue
                else:
                    # MD5值不同，需要备份
                    log_print(doc_name + ' has changed, backup will begin soon (MD5 mismatch)')
            
            Existed_in_this_session[doc_name] = True   #标记该文件在本次会话中出现过
            log_print('Start to backup ' + doc_name + ' to ' + word_save_folder)   #打印备份开始信息
            # 如果目标文件存在，先移除只读属性
            remove_readonly(new_doc_path)
            copy_start_time=datetime.datetime.now()   #记录复制操作开始时间
            shutil.copy2(doc_path, new_doc_path)   #复制文档到备份文件夹，并尝试保留元数据（如修改时间等）
            copy_end_time=datetime.datetime.now()   #记录复制操作结束时间
            copy_used_time=copy_end_time-copy_start_time  #计算复制所用时间

            modified_time=os.path.getmtime(new_doc_path)   #获取备份文件的修改时间
            current_time=time.time()   #获取当前时间
            os.utime(new_doc_path, (modified_time, current_time))   #将修改时间存储到访问时间（参数1），创建时间存储到修改时间（参数2），方便文件系统根据修改时间排序

            file_skip_count[doc_name] = 0   #重置该文件的跳过计数器
            any_backup_performed = True   #标记本轮有备份操作
            log_print('Successfully backuped ' + doc_name + ' to ' + word_save_folder + ' in ' + str(copy_used_time) +' s')   #打印备份成功信息

            upload_queue.append((doc_name,new_doc_path))  # 将文件名和备份路径添加到上传队列
        upload_to_openlist()  # 启动上传线程

        if not any_backup_performed and len(documents) == 0:   #检查变量值，如果没有可备份PPT，打印此条信息
                log_print('No doc available now')

    except FileNotFoundError:   #捕获由于U盘等移动存储介质被移除而导致的“文件未找到”异常，使用2.0版本中的SaveAs方法进行备份
        if not os.path.exists(word_save_folder):   #检查word备份目录是否存在
            os.makedirs(word_save_folder)   #若不存在则创建备份目录（包括所有必要的父目录）
            log_print('Target backup folder not found, created: ' + word_save_folder + ' successfully')   #打印成功创建word备份目录的信息
    
        for idx in range(1, documents.Count + 1):   #遍历文档实例集合
            doc = documents.Item(idx)   #获取当前文档实例
            doc_path = doc.FullName   #获取Word文件的完整路径
            doc_name = os.path.basename(doc_path)   #提取文件名
            new_doc_path = os.path.join(word_save_folder, doc_name)   #生成备份路径
            log_print('Start to backup ' + doc_name + ' to ' + word_save_folder + ' using SaveAs method')   #打印备份开始信息
            # 如果目标文件存在，先移除只读属性
            remove_readonly(new_doc_path)
            save_start_time=datetime.datetime.now()   #记录保存操作开始时间
            doc.SaveAs(new_doc_path)   #使用SaveAs方法保存当前文档实例到指定路径
            save_end_time=datetime.datetime.now()   #记录保存操作结束时间
            save_used_time=save_end_time-save_start_time  #计算保存所用时间
            SaveAs_method_activated[doc_name] = True   #标记该文件已激活SaveAs方法，后续不再备份
            log_print('Detected access control, activated SaveAs method, successfully backuped ' + doc_name + ' to ' + word_save_folder + ' in ' + str(save_used_time) + ' s')   #打印备份成功信息

            upload_queue.append((doc_name,new_doc_path))  # 将文件名和备份路径添加到上传队列
            upload_to_openlist()  # 启动上传线程
    except Exception as e:   #获取其他错误类型
            if type(e).__name__=='com_error':   #捕获无打开的PowerPoint实例而产生的的异常
                log_print('No doc available now (Word application not detected)')   #打印带时间戳和运行次数的异常信息
            else:   #打印出其他错误并继续轮询
                log_print('Exception: ' + type(e).__name__ + ', request continue')   #打印带时间戳和运行次数的异常信息



@timeout(seconds=600, config_key='backup_timeout')  #添加10分钟超时机制
def save_open_WPS_files(ppt_save_folder):   #定义WPS保存函数，参数ppt_save_folder是备份文件的存储路径
    global upload_queue  # 声明全局上传队列变量
    try:
        if not os.path.exists(ppt_save_folder):   #检查ppt备份目录是否存在
            os.makedirs(ppt_save_folder)   #若不存在则创建备份目录（包括所有必要的父目录）
            log_print('Target backup folder not found, created: ' + ppt_save_folder + ' successfully')   #打印成功创建ppt备份目录的信息
        
        WPS_app=win32.GetObject(Class='KWPP.Application')   #捕获当前打开的WPS演示实例
        WPSpresentations = WPS_app.Presentations   #获取当前WPS实例中所有打开的演示文稿集合

        any_backup_performed = False   #标记本轮是否有任何备份操作（替代原haveppt）
        
        for ppt in WPSpresentations:   #遍历集合
            WPS_ppt_path = ppt.FullName   #获取PPT文件的完整路径
            WPS_ppt_name = os.path.basename(WPS_ppt_path)   #提取文件名
            WPS_new_ppt_path = os.path.join(ppt_save_folder, WPS_ppt_name)   #生成备份路径

            if os.path.exists(WPS_new_ppt_path):   #检查备份文件是否已存在
                if SaveAs_method_activated[WPS_ppt_name] == True:   #如果SaveAs方法已被激活，则不再使用复制方法
                    log_print(WPS_ppt_name + ' has already existed in ' + ppt_save_folder + ', skipped backup (SaveAs method activated)')   #打印带时间戳和运行次数的跳过信息
                    continue   #跳过此次备份
                
                # 计算原始文件和备份文件的MD5值
                original_md5 = calculate_md5(WPS_ppt_path)
                backup_md5 = calculate_md5(WPS_new_ppt_path)
                
                # 只有当两个MD5都成功计算且相同时才跳过
                if original_md5 and backup_md5 and original_md5 == backup_md5:
                    # MD5值相同，跳过备份
                    log_print(WPS_ppt_name + ' has already existed in ' + ppt_save_folder + ', skipped backup (MD5 match)')   #打印带时间戳和运行次数的跳过信息
                    continue   #跳过此次备份
                elif original_md5 is None:
                    # 源文件找不到，跳过这次备份
                    log_print(WPS_ppt_name + ' source file not found, skipping this backup')
                    continue
                else:
                    # MD5值不同，需要备份
                    log_print(WPS_ppt_name + ' has changed, backup will begin soon (MD5 mismatch)')
            
            Existed_in_this_session[WPS_ppt_name] = True   #标记该文件在本次会话中出现过
            log_print('Start to backup ' + WPS_ppt_name + ' to ' + ppt_save_folder)   #打印备份开始信息
            # 如果目标文件存在，先移除只读属性
            remove_readonly(WPS_new_ppt_path)
            copystarttime=datetime.datetime.now()   #记录复制操作开始时间
            shutil.copy2(WPS_ppt_path, WPS_new_ppt_path)   #复制PPT到备份文件夹，并尝试保留元数据（如修改时间等）
            copyendtime=datetime.datetime.now()   #记录复制操作结束时间
            copyusedtime=copyendtime-copystarttime  #计算复制所用时间

            modified_time=os.path.getmtime(WPS_new_ppt_path)   #获取备份文件的修改时间
            create_time=os.path.getctime(WPS_new_ppt_path)   #获取备份文件的创建时间
            os.utime(WPS_new_ppt_path, (modified_time, create_time))   #将修改时间存储到访问时间（参数1），创建时间存储到修改时间（参数2），方便文件系统根据修改时间排序

            file_skip_count[WPS_ppt_name] = 0   #重置该文件的跳过计数器
            any_backup_performed = True   #标记本轮有备份操作
            log_print('Successfully backuped ' + WPS_ppt_name + ' to ' + ppt_save_folder + ' in ' + str(copyusedtime) +' s')   #打印带时间戳和运行次数的备份成功信息

            upload_queue.append((WPS_ppt_name,WPS_new_ppt_path))  # 将文件名和备份路径添加到上传队列
            upload_to_openlist()  # 启动上传线程

        if not any_backup_performed and len(WPSpresentations) == 0:   #检查变量值，如果没有可备份PPT，打印此条信息
            log_print('No WPS ppt available now (Normal request)')   #打印带时间戳和运行次数的运行信息

    except FileNotFoundError:   #捕获由于U盘等移动存储介质被移除而导致的“文件未找到”异常，使用2.0版本中的SaveAs方法进行备份
        if not os.path.exists(ppt_save_folder):   #检查ppt备份目录是否存在
            os.makedirs(ppt_save_folder)   #若不存在则创建备份目录（包括所有必要的父目录）
            log_print('Target backup folder not found, created: ' + ppt_save_folder + ' successfully')   #打印成功创建ppt备份目录的信息
        
        for idx in range(1, WPSpresentations.Count + 1):   #遍历PPT实例集合
            ppt = WPSpresentations.Item(idx)   #获取当前PPT实例
            WPS_ppt_path = ppt.FullName   #获取PPT文件的完整路径
            WPS_ppt_name = os.path.basename(WPS_ppt_path)   #提取文件名
            WPS_new_ppt_path = os.path.join(ppt_save_folder, WPS_ppt_name)   #生成备份路径
            log_print('Start to backup ' + WPS_ppt_name + ' to ' + ppt_save_folder + ' using SaveAs method')   #打印备份开始信息
            # 如果目标文件存在，先移除只读属性
            remove_readonly(WPS_new_ppt_path)
            savestarttime=datetime.datetime.now()   #记录保存操作开始时间
            ppt.SaveAs(WPS_new_ppt_path)   #使用SaveAs方法保存当前PPT实例到指定路径
            saveendtime=datetime.datetime.now()   #记录保存操作结束时间
            saveusedtime=saveendtime-savestarttime  #计算保存所用时间
            SaveAs_method_activated[WPS_ppt_name] = True   #标记该文件已激活SaveAs方法，后续不再备份
            log_print('Detected access control, activated SaveAs method, successfully backuped ' + WPS_ppt_name + ' to ' + ppt_save_folder + ' in ' + str(saveusedtime) + ' s')   #打印备份成功信息

            upload_queue.append((WPS_ppt_name,WPS_new_ppt_path))  # 将文件名和备份路径添加到上传队列
        upload_to_openlist()  # 启动上传线程  
    except Exception as e:   #获取其他错误类型
            if type(e).__name__=='com_error':   #捕获无打开的WPS实例而产生的的异常
                log_print('No ppt available now (WPS application not detected)')   #打印异常信息
            else:   #打印出其他错误并继续轮询
                log_print('Exception: ' + type(e).__name__ + ', request continue')   #打印异常信息



@timeout(seconds=600, config_key='backup_timeout')   #添加10分钟超时机制
def accurate_backup():   #定义精确备份函数
    try:
        if os.path.exists(source_path):   #检查源文件夹是否存在
            log_print('Start accurate backup from ' + source_path + ' to ' + target_path)   #打印精确备份开始信息
            copy_start_time=datetime.datetime.now()   #记录复制操作开始时间
            shutil.copytree(source_path, target_path)  # 复制源文件夹及其内容到目标文件夹
            copy_end_time=datetime.datetime.now()   #记录复制操作结束时间
            copy_used_time=copy_end_time-copy_start_time  #计算复制所用时间
            log_print(f'Accurate backup completed successfully from {source_path} to {target_path} in {copy_used_time} s')  # 打印精确备份完成信息
        else:
            log_print('Source path for accurate backup does not exist: ' + source_path + ', wait for the next request')  # 打印源文件夹不存在信息，等待下次请求
    except Exception as e:
        log_print('Accurate backup failed: ' + str(e))  # 打印精确备份失败信息
    


def toggle_console():   #切换控制台窗口的显示/隐藏状态
    global console_visible   #声明全局变量console_visible，以便在函数内修改其值
    console_window = ctypes.windll.kernel32.GetConsoleWindow()   #获取控制台窗口句柄
    if console_visible:   #隐藏控制台窗口
        ctypes.windll.user32.ShowWindow(console_window, 0)   #隐藏控制台窗口
        console_visible = False
    else:   #显示控制台窗口
        ctypes.windll.user32.ShowWindow(console_window, 1)  # SW_SHOWNORMAL
        console_visible = True

def exit_program(icon):   #退出程序
    icon.stop()
    os._exit(0)
'''
def on_clicked(icon):   #左键单击事件处理（无法生效）
    global behavior   #声明全局变量behavior，以便在函数内修改其值
    if behavior == 'open_console':   #切换控制台窗口显示/隐藏状态
        toggle_console()
    elif behavior == 'exit_program':   #退出程序
        exit_program(icon)
'''

# Base64编码的图标数据
ICON_BASE64 = '''
AAABAAEAgIAAAAEAIAAoCAEAFgAAACgAAACAAAAAAAEAAAEAIAAAAAAAAAABAMMOAADDDgAAAAAAAAAAAAD///////////////////////////////////////////////////////////////////////////////////////////////////////////7+/v/9/f7/+/z9//j6+//1+Pn/8/b3//D19v/v9PX/7vP1/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/u8/X/7/T1//D19v/z9vf/9fj5//j6+//7/P3//f3+//7+/v///////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////v////3+/v/7/Pz/9/n6//P2+P/w9Pb/7/T1/+/z9f/u8/X/7vP0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7vP0/+7z9f/v8/X/7/T1//D09v/z9vj/9/n6//v8/P/9/v7//v///////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////v////39/v/5+/v/9Pf4//D09v/u8/X/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/u8/X/8PT2//T3+P/5+/v//f3+//7///////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////3+/v/5+/v/8/b4/+/09f/u8/T/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7vP0/+/09f/z9vj/+fv7//39/v////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////7+/v/7/Pz/9fj5//D09v/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+7y9P/v9PX/9Pf4//r8/P/+/v7////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////9/v7/+Pr6//H19//u8/X/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/u8/X/8fX3//j6+v/9/v7//////////////////////////////////////////////////////////////////////////////////////////////////////////////////P39//b4+f/v9PX/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7/T1//b4+f/8/f3///////////////////////////////////////////////////////////////////////////////////////////////////////v8/P/09/j/7/P1/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7/P1//T3+P/7/Pz////////////////////////////////////////////////////////////////////////////////////////////7/Pz/8/b3/+7z9f/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7vP1//P29//7/Pz/////////////////////////////////////////////////////////////////////////////////+/z8//P29//u8/T/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7vP0//P29//7/Pz///////////////////////////////////////////////////////////////////////z9/f/09/j/7vP1/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7vP1//T3+P/8/f3////////////////////////////////////////////////////////////9/v7/9fj5/+7z9f/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7/P1//b4+f/9/v7//////////////////////////////////////////////////v7+//j6+v/v9PX/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7vL0/+7y9P/v8vT/7/L0/+/y9P/w8/T/8PP0//Dz9P/w8/T/8PP0//Dz9P/w8/T/8PP0//Dz9P/v8/T/7/L0/+/y9P/v8vT/7vL0/+7y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7/T1//j6+v/+/v7////////////////////////////////////////////6/Pz/8fX3/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7vL0/+/y9P/v8vT/7/L0/+ry9P/l8PT/4e/1/97v9f/a7vX/1+31/9Xt9v/U7fb/0+32/9Ls9v/R7Pb/0uz2/9Ps9v/U7fb/1e32/9ft9f/a7vX/3e71/+Dv9f/l8PX/6vH0/+7y9P/v8/T/7/L0/+7y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/8fX3//v8/P///////////////////////////////////////f3+//T3+f/u8/X/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7/L0/+7y9P/o8fT/4PD1/9ju9f/O7Pb/vOj3/6fk+P+X4fn/i976/37b+/9y2fz/a9f8/2jW/f9j1f3/YNT9/1zU/f9f1P3/YtT9/2bV/f9p1vz/cdj8/3za+/+G3Pv/k976/6Ti+P+45/f/yuv2/9bt9f/e7/X/5vD0/+3y9P/v8vT/7vL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/u8/X/9fj5//3+/v////////////////////////////7////4+/v/7/T2/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/u8vT/7vL0/+rx9P/f7/X/0Oz2/7Xn+P+U4fr/dtv8/2HX/f9W1P7/UdP+/0zS/v9J0f7/RdD//0LP//8/z///Ps7//z3O//88zf//O83//zvN//87zf//PMz//z3M//8/zf//Qc3//0TN/v9Hzv7/S8/+/1DP/v9Y0f3/bdX8/4rc+v+q4/j/yer2/9zu9f/o8fT/7vL0/+7y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/w9Pb/+fv7//7//////////////////////////P3+//P2+P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7vL0/+ry9P/b7/X/v+n3/5Lh+v9q2fz/V9b+/03U/v9F0v//P9D//z7Q//8+0P//PtD//z7Q//8+z///Ps///z7P//8+zv//Ps7//z3O//89zv//Pc3//zzN//88zf//O8z//zvM//86zP//Osz//znL//84y///N8r//zfK//88y///Q8z+/0zO/v9d0vz/g9r6/7Hk+P/V7fX/6PH0/+7y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/z9vj//f3+//////////////////7////5+/v/7/T1/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7vL0/+3y9P/h8PX/wOr3/4ng+v9f2P3/TdT+/0PS//9A0v//QNH//0HR//9B0f//QdH//0DQ//9A0P//QND//z/Q//8/z///P8///z7P//8+zv//Ps7//z3O//89zv//Pc3//zzN//88zf//PMz//zvM//87zP//O8z//zrL//86y///Osv//znK//84yv//N8r//zjK//9Ay/7/UM7+/3fX+/+y5ff/3O71/+3y9P/u8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+/09f/5+/v//v///////////////f7+//T3+P/u8/T/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/s8vT/1+71/6Dk+f9m2v3/TNX+/0LT//9C0v//QtL//0LS//9C0v//QtL//0HR//9B0f//QdH//0DQ//9A0P//QND//z/Q//8/z///P8///z7P//8+zv//Ps7//z3O//89zv//Pc3//zzN//88zf//PMz//zvM//87zP//O8z//zrL//86y///Osv//znK//85yv//Ocr//zfJ//83yf//Pcr//1XQ/f+Q3fr/0ev2/+zy9P/u8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7vP0//T3+P/9/v7////////////7/Pz/8PT2/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7PL0/9Ht9v+N4fr/WNf+/0TU//9D1P//RNP//0PT//9D0///Q9L//0LS//9C0v//QtL//0HR//9B0f//QdH//0DQ//9A0P//QND//z/Q//8/z///P8///z7P//8+zv//Ps7//z3O//89zv//Pc3//zzN//88zf//PMz//zvM//87zP//O8z//zrL//86y///Ocv//znK//85yv//Ocr//zjK//83yf//N8n//0jN/v+A2fv/z+v2/+zy9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/8PT2//v8/P///////v7+//f5+v/u8/X/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/W7vX/jeH6/1PX/v9E1P//RNT//0TU//9E1P//RNP//0PT//9D0///Q9L//0LS//9C0v//QtL//0HR//9B0f//QdH//0DQ//9A0P//QND//z/Q//8/z///P8///z7P//8+zv//Ps7//z3O//89zv//Pc3//zzN//88zf//PMz//zvM//87zP//Ocv//zvM//8/zP//Pcv//zjK//84yv//Ocr//zjK//84yf//N8n//0fM/v+J2/r/2e31/+7y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/u8/X/9/n6//7+/v/9/f7/8/b4/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/u8vT/4fD1/6Dl+f9Y2P7/RNX//0XV//9F1P//RdT//0TU//9E1P//RNP//0PT//9D0///Q9L//0LS//9C0v//QtL//0HR//9B0f//QdH//0DQ//9A0P//QND//z/Q//8/z///P8///z7P//8+zv//Ps7//z3O//89zv//Pc3//zzN//88zf//O8z//z7M//9O0P7/ZdX9/3XY+/9u1vz/VdH9/0HM//85yv//Ocr//zjJ//84yf//N8n//1LP/f+v5Pj/5/H0/+7y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/z9vj//f3+//v8/f/w9Pb/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+vy9P/F6/f/adz9/0bW//9F1f//RdX//0XV//9F1P//RNT//0TU//9E1P//RNP//0PT//9D0///Q9L//0LS//9C0v//QtL//0HR//9B0f//QdH//0DQ//9A0P//QND//z/Q//8/z///P8///z7P//8+zv//Ps7//z3O//89zv//Pc3//zzN//8/zf//YdT9/6Ti+f/P7Pb/2e71/9Xt9f+75/f/e9n7/0bN/v85yv//Ocr//zjJ//83yf//Psv//33Z+//a7vX/7/L0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0//D09v/7/P3/+Pr7/+/09f/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/u8vT/4fD1/5jk+v9S2P7/Rdb//0bW//9G1f//RdX//0XV//9F1P//RNT//0TU//9E1P//Q9P//0PT//9D0///Q9L//0LS//9C0v//QtL//0HR//9B0f//QdH//0DQ//9A0P//QND//z/Q//8/z///P8///z7P//8+zv//Ps7//z3O//89zf//PM3//1bT/f+t5fj/4vD1/+/y9P/w8/T/7/P0/+vy9P/K6vb/dNf7/z/M//84yv//Ocr//zjJ//83yf//WtH9/8zq9v/u8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7/T1//j6+//1+Pn/7/P1/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+/y9P/W7vb/dN78/0nX//9G1v//Rtb//0bW//9G1f//RdX//0XV//9F1P//RNT//0TU//9E1P//Q9P//0PT//9D0///QtL//0LS//9C0v//QtL//0HR//9B0f//QdH//0DQ//9A0P//QND//z/P//8/z///P8///z7P//8+zv//Ps7//zzO//9Ez///f9v7/9ju9f/v8vT/7fL0/+3y9P/t8vT/7vL0/+jx9P+u5Pj/Ts/+/zfK//85yv//Ocr//zXJ//9Rz/7/v+j3/+zy9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/v8/X/9fj5//P29//u8/X/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7vL0/83t9v9j2/3/Rdb//0fW//9H1v//Rtb//0bW//9G1f//RdX//0XV//9F1P//RNT//0TU//9E1P//Q9P//0PT//9D0///QtL//0LS//9C0v//QdH//0HR//9B0f//QdH//0DQ//9A0P//QND//z/P//8/z///P8///z7P//8+zv//PM7//0zR/v+b4fn/4vD1/+7y9P/t8vT/7fL0/+3y9P/t8vT/7/L0/83r9v9Z0v3/OMr//znK//85yv//Nsn//1DP/v+85/f/6/L0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+7z9f/z9vf/8PX2/+7z9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/s8vT/xez2/1/b/f9F1///R9f//0fW//9H1v//Rtb//0bW//9G1f//RdX//0XV//9F1P//RNT//0TU//9E0///Q9P//0PT//9D0///QtL//0LS//9C0v//QdH//0HR//9B0f//QNH//0DQ//9A0P//QND//z/P//8/z///P8///z7P//88zv//TNH+/53h+f/j8PX/7vL0/+3y9P/t8vT/7fL0/+3y9P/v8/T/zuz2/1nS/f84y///Osv//znK//82yf//Uc/+/73n9//r8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7vP0//D19v/v9PX/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+zy9P/E7Pb/X9v9/0bX//9I1///R9f//0fW//9H1v//Rtb//0bV//9G1f//RdX//0XV//9F1P//RNT//0TU//9E0///Q9P//0PT//9D0///QtL//0LS//9C0v//QdH//0HR//9B0f//QNH//0DQ//9A0P//P9D//z/P//8/z///P8///z3P//9G0P7/g9z7/9ru9f/v8vT/7fL0/+3y9P/t8vT/7vL0/+rx9P+z5vj/UND+/zjL//86y///Osv//zfK//9Rz/7/vuj3/+vy9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7/T1/+7z9f/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7PL0/8Xs9v9g3P3/Rtf//0jX//9I1///R9f//0fW//9H1v//Rtb//0bV//9G1f//RdX//0XV//9F1P//RNT//0TU//9E0///Q9P//0PT//9D0///QtL//0LS//9C0v//QdH//0HR//9B0f//QNH//0DQ//9A0P//P9D//z/P//8/z///Ps///z7P//9b1P3/tOb4/+Xw9P/v8vT/7/L0/+/y9P/t8vT/z+z2/3za+/9Czf//Osv//zrL//86y///N8r//1HQ/v++6Pf/6/L0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/u8/X/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/s8vT/xez2/2Dc/f9G1///SNf//0jX//9I1///R9f//0fW//9H1v//Rtb//0bV//9G1f//RdX//0XV//9F1P//RNT//0TU//9E0///Q9P//0PT//9D0///QtL//0LS//9C0v//QdH//0HR//9B0f//QNH//0DQ//9A0P//P9D//z/P//8/z///Ps///0LQ//9p1/z/r+X4/9Tt9v/c7/X/2u71/8Tp9/+G3Pv/S8/+/zvM//87zP//O8v//zrL//83yv//UtD+/77o9//r8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+zy9P/F7Pb/YNz9/0fY//9J2P//SNf//0jX//9I1///R9f//0fW//9H1v//Rtb//0bV//9G1f//RdX//0XV//9F1P//RNT//0TU//9E0///Q9P//0PT//9D0///QtL//0LS//9C0v//QdH//0HR//9B0f//QNH//0DQ//9A0P//P9D//z/P//8/z///Ps///0LP//9V0/7/cdj8/4Pc+/972vv/X9T9/0fP/v88zf//O8z//zvM//87zP//O8v//zjK//9S0P7/v+j3/+vy9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7PL0/8Xs9v9g3P3/R9j//0nY//9J2P//SNf//0jX//9I1///R9f//0fW//9H1v//Rtb//0bV//9G1f//RdX//0XV//9F1P//RNT//0TU//9E0///Q9P//0PT//9D0///QtL//0LS//9C0v//QdH//0HR//9B0f//QNH//0DQ//9A0P//P9D//z/P//8/z///Ps///z3P//9Bz///RdD//0PP//89zf//O83//zzN//88zf//O8z//zvM//87zP//OMv//1LQ/v+/6Pf/7PL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/s8vT/xez2/2Hc/f9H2P//Sdj//0nY//9J2P//SNf//0jX//9I1///R9f//0fW//9H1v//Rtb//0bV//9G1f//RdX//0XV//9F1P//RNT//0TU//9E0///Q9P//0PT//9D0///QtL//0LS//9C0v//QdH//0HR//9B0f//QNH//0DQ//9A0P//P9D//z/P//8/z///Ps///z3P//89zv//PM7//z3O//89zf//PM3//zzN//88zP//O8z//zvM//84y///U9H+/7/o9//s8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+zy9P/F7Pb/Yd39/0jZ//9K2f//Sdj//0nY//9J2P//SNf//0jX//9I1///R9f//0fW//9H1v//Rtb//0bV//9G1f//RdX//0XV//9F1P//RNT//0TU//9E0///Q9P//0PT//9D0///QtL//0LS//9C0v//QdH//0HR//9B0f//QNH//0DQ//9A0P//P9D//z/P//8/z///Ps///z7O//8+zv//Pc7//z3O//89zf//PM3//zzN//88zP//O8z//zjL//9T0f7/v+j3/+zy9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7PL0/8Xt9v9h3f3/SNn//0rZ//9K2f//Sdj//0nY//9J2P//SNf//0jX//9I1///R9f//0fW//9H1v//Rtb//0bV//9G1f//RdX//0XV//9F1P//RNT//0TU//9E0///Q9P//0PT//9B0///QdL//0DS//9A0f//QNH//z/R//8/0f//P9D//z7Q//8+0P//Ps///z3P//89z///Pc7//zzO//88zv//PM7//zvN//87zf//O83//zrM//86zP//N8v//1LR/v/A6Pf/7PL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/s8vT/xe32/2Hd/f9I2f//Stn//0rZ//9J2f//Sdj//0nY//9J2P//SNf//0jX//9I1///R9f//0fW//9H1v//Rtb//0bV//9G1f//RdX//0XU//9F1P//RNT//0TU//9E0///RtP//07V/v9Q1f7/T9T+/0/U/v9P1P7/T9T+/07U/v9O0/7/TtP+/03T/v9N0/7/TdL+/0zS/v9M0v7/TNL+/0zR/v9L0f7/S9H+/0vR/v9K0P7/StD+/0rQ/v9Hz/7/X9T9/8Tp9v/s8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+zy9P/F7fb/Yt79/0jZ//9K2f//Stn//0rZ//9J2f//Sdj//0nY//9I2P//SNf//0jX//9I1///R9b//0fW//9H1v//Rtb//0bV//9G1f//RdX//0XU//9F1P//RNT//0TU//9Y1/3/j+D6/53j+f+b4/n/m+P5/5vj+f+b4/n/muP5/5rj+f+a4vn/muL5/5ri+f+a4vn/meL5/5ni+f+Z4vn/meH5/5nh+f+Z4fn/mOH5/5jh+f+Y4fn/mOH5/5fg+f+k4vj/2O71/+zy9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8/X/7vP2/+709v/u9Pf/7vT3/+709//u9Pf/7vT3/+709//u9Pf/7vT2/+7z9v/t8vT/7PL0/8Xt9v9i3v3/Sdr//0va//9K2f//Stn//0rZ//9J2P//Sdj//0nY//9I2P//SNf//0jX//9H1///R9b//0fW//9H1v//Rtb//0bV//9G1f//RdX//0XU//9F1P//RNT//2ja/f/L7Pb/5PD1/+Lw9f/i8PX/4vD1/+Lw9f/i8PX/4vD1/+Lw9f/i8PX/4vD1/+Lw9f/i8PX/4vD1/+Lw9f/i8PX/4vD1/+Lw9f/i8PX/4vD1/+Lw9f/i8PX/4fD1/+Pw9f/q8fT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9f/u8/b/7fL0/+nr6//m5uL/5OLe/+Pi3P/j4tz/4+Lc/+Pi3P/j4dz/4+Hc/+Ph3P/j4t3/5+jm/+zw8v/s8vT/xu32/2Le/f9J2v//S9r//0va//9K2f//Stn//0rZ//9J2P//Sdj//0nY//9I2P//SNf//0jX//9H1///R9b//0fW//9G1v//Rtb//0bV//9G1f//RdX//0XU//9F1P//atv9/9Xu9v/w8/T/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+zy9P/s8vT/7PL0/+zy9P/s8vT/7PL0/+zy9P/s8vT/7PL0/+zy9P/s8vT/7fL0/+/y9P/v8/T/7vL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8/X/7PDx/+bm4v/c1cv/zbyl/8Cmhf+4mXT/tpZv/7WWb/+1lm//tZZv/7WVb/+1lW//tJRu/7eZdf/LuqX/6Onn/+3z9f/G7fb/Yt79/0na//9L2v//S9r//0va//9K2f//Stn//0rZ//9J2P//Sdj//0nY//9I2P//SNf//0jX//9H1///R9b//0fW//9G1v//Rtb//0bV//9F1f//RdX//0XU//9h2f3/sOf4/8Tr9v/C6vb/wur2/8Lq9v/C6vb/wur2/8Lq9v/C6vb/wur2/8Lq9v/C6vb/wur2/8Lq9v/B6vb/wer2/8Hq9v/B6vb/wer2/8Hp9v/B6fb/wOn2/8Dp9v/A6fb/wOn2/8Dp9v/A6fb/wOn2/8Do9v/A6Pb/wOj2/8Do9v/A6Pb/wOj3/8Do9//G6fb/z+v2/9ft9f/h7/X/6/L0/+7y9P/u8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+jq6P/Z0cP/waeH/6+JW/+nfEn/pHZA/6FzPP+gcjv/oHE7/6BxO/+fcTv/n3E7/55xOv+ecDn/onZC/76lhf/m5eL/7fP2/8bt9v9i3v3/Str//0za//9L2v//S9r//0va//9K2f//Stn//0rZ//9J2P//Sdj//0nY//9I2P//SNf//0jX//9H1///R9b//0fW//9G1v//Rtb//0bV//9F1f//RdX//0rV//9Y2P7/W9j9/1vY/f9b2P3/Wtf9/1rX/f9a1/3/Wtb9/1nW/f9Z1v3/Wdb9/1nW/f9Y1f3/WNX9/1jV/f9X1f3/V9X9/1fU/f9X1P3/VtT9/1bT/f9W0/3/VdP9/1XT/f9V0v3/VNL9/1TS/f9U0f3/VNH9/1PR/f9T0f3/U9D9/1PQ/f9S0P3/UtD+/1PQ/f9b0v3/ctb7/5be+f++6Pf/2u71/+rx9P/u8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/l5eH/zryk/7KNX/+ld0H/oXA2/59vNf+gbzb/n3A2/59vNv+fbzb/n282/55vNv+ebzb/nW42/5xtNf+gdD//vaOD/+bl4v/t8/b/xu32/2Pe/f9K2///TNv//0za//9L2v//S9r//0va//9K2f//Stn//0rZ//9J2P//Sdj//0nY//9I2P//SNf//0jX//9H1///R9b//0fW//9G1v//Rtb//0bV//9F1f//RNX//0PU//9C1P//QtT//0HT//9B0///QNP//0DS//9A0v//QNL//z/R//8/0f//P9H//z7R//8+0P//PtD//z3P//89z///Pc///zzP//88zv//PM7//zvO//87zf//O83//zrN//86zP//Osz//znM//85zP//Ocv//zjL//84y///OMr//zfK//83yv//Nsr//zjK//89y///Rsz+/1fQ/f+D2vr/vuj3/+Lw9f/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/5OPf/8m0mf+tg1L/onE3/6FwNv+hcTf/oXA3/6BwN/+gcDf/oHA3/59wN/+fbzf/n283/55vN/+ebzf/nW42/6F0QP++o4T/5uXi/+3z9v/G7fb/Y9/9/0rb//9M2///TNv//0za//9L2v//S9r//0va//9K2f//Stn//0rZ//9J2P//Sdj//0nY//9I2P//SNf//0jX//9H1///R9b//0fW//9G1v//Rtb//0bV//9F1f//RdX//0XU//9E1P//RNT//0TU//9D0///Q9P//0PT//9C0v//QtL//0LS//9C0v//QdH//0HR//9B0f//QND//0DQ//9A0P//P8///z/P//8/z///Ps///z7O//8+zv//Pc7//z3N//89zf//PM3//zzN//88zP//O8z//zvM//87y///Osv//zrL//86y///Ocr//zjK//83yf//N8n//0HL//9e0f3/oOH5/9nt9f/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+bm4//KtZr/rIJP/6NxN/+icTb/onE3/6FxN/+hcTf/oXA3/6BwN/+gcDf/oHA3/59wN/+fbzf/n283/55vN/+dbjb/oXVA/76jhP/m5eL/7fP2/8bt9v9j3/3/S9v//03b//9M2///TNv//0za//9L2v//S9r//0va//9K2f//Stn//0rZ//9J2P//Sdj//0nY//9I2P//SNf//0jX//9H1///R9b//0fW//9G1v//Rtb//0bV//9F1f//RdX//0XU//9E1P//RNT//0TU//9D0///Q9P//0PT//9C0v//QtL//0LS//9B0f//QdH//0HR//9B0f//QND//0DQ//9A0P//P8///z/P//8/z///Ps///z7O//8+zv//Pc7//z3N//89zf//PM3//zzN//88zP//O8z//zvM//87y///Osv//zrL//86y///Ocr//znK//85yv//N8n//zjJ//9Qzv7/kN36/9bt9f/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3z9f/q7ez/0sSv/6+HVf+jcjf/o3I2/6NyN/+icTf/onE3/6FxN/+hcTf/oXA3/6BwN/+gcDf/oHA3/59wN/+fbzf/n283/55uNv+idUD/vqSE/+bl4v/t8/b/xu32/2Tf/f9L3P//Tdz//03b//9M2///TNv//0za//9L2v//S9r//0va//9K2f//Stn//0rZ//9J2P//Sdj//0nY//9I2P//SNf//0jX//9H1///R9b//0fW//9G1v//Rtb//0bV//9F1f//RdX//0XU//9E1P//RNT//0TT//9D0///Q9P//0PT//9C0v//QtL//0LS//9B0f//QdH//0HR//9A0f//QND//0DQ//9A0P//P8///z/P//8/z///Ps///z7O//8+zv//Pc7//z3N//89zf//PM3//zzN//88zP//O8z//zvM//87y///Osv//zrL//86y///Ocr//znK//85yv//OMn//zfJ//9Nzv7/k976/9ru9f/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fLz/9/az/+4lmv/pXQ6/6NxNv+jcjf/o3I3/6NyN/+icTf/onE3/6FxN/+hcTf/oXA3/6BwN/+gcDf/oHA3/59wN/+fbzf/nm42/6J1QP++pIT/5uXi/+3z9v/G7fb/ZN/9/0vc//9N3P//Tdz//03b//9M2///TNv//0za//9L2v//S9r//0va//9K2f//Stn//0rZ//9J2P//Sdj//0nY//9I2P//SNf//0jX//9H1///R9b//0fW//9G1v//RtX//0bV//9F1f//RdX//0XU//9E1P//RNT//0TT//9D0///Q9P//0PT//9C0v//QtL//0LS//9B0f//QdH//0HR//9A0f//QND//0DQ//8/0P//P8///z/P//8/z///Ps///z7O//8+zv//Pc7//z3N//89zf//PM3//zzN//88zP//O8z//zvM//87y///Osv//zrL//86y///Ocr//znK//85yv//OMn//zfJ//9Rz/7/pOH5/+Pw9f/u8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+7z9v/n6OX/yLGU/6p9Rf+kcjb/pHI3/6RyN/+jcjf/o3I3/6NyN/+icTf/onE3/6FxN/+hcTf/oXA3/6BwN/+gcDf/oHA3/59wN/+fbjb/o3VA/7+khP/m5eL/7fP2/8bt9v9k4P3/TNz//07c//9N3P//Tdz//03b//9M2///TNv//0za//9L2v//S9r//0va//9K2f//Stn//0rZ//9J2P//Sdj//0nY//9I1///SNf//0jX//9H1///R9b//0fW//9G1v//RtX//0bV//9F1f//RdX//0XU//9E1P//RNT//0TT//9D0///Q9P//0PT//9C0v//QtL//0LS//9B0f//QdH//0HR//9A0f//QND//0DQ//8/0P//P8///z/P//8+z///Ps///z7O//8+zv//Pc7//z3N//89zf//PM3//zzN//88zP//O8z//zvM//87y///Osv//zrL//86y///Ocr//znK//85yv//OMn//znJ//9j0vz/wej3/+vx9P/u8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fHz/93Wy/+0jV7/pXM3/6VzN/+kczf/pHI3/6RyN/+jcjf/o3I3/6NyN/+icTf/onE3/6FxN/+hcTf/oXA3/6BwN/+gcDf/oHA3/59vNv+jdUD/v6SE/+bl4v/t8/b/x+32/2Xg/f9M3P//Ttz//07c//9N3P//Tdz//03b//9M2///TNv//0za//9L2v//S9r//0vZ//9K2f//Stn//0rZ//9J2P//Sdj//0nY//9I1///SNf//0jX//9H1///R9b//0fW//9G1v//RtX//0bV//9F1f//RdX//0XU//9E1P//RNT//0TT//9D0///Q9P//0PT//9C0v//QtL//0LS//9B0f//QdH//0HR//9A0f//QND//0DQ//8/0P//P8///z/P//8+z///Ps///z7O//89zv//Pc7//z3N//89zf//PM3//zzN//88zP//O8z//zvM//87y///Osv//zrL//86y///Ocr//znK//85yv//N8n//0PL/v+H2/r/2+71/+7y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+7z9v/o6uj/y7aa/6t9Rv+lcjX/pXM3/6VzN/+kczf/pHI3/6RyN/+jcjf/o3I3/6NyN/+icTf/onE3/6FxN/+hcTf/oXA3/6BwN/+gcDf/n282/6N2QP+/pIT/5uXi/+3z9v/H7fb/ZeD9/0zd//9O3f//Ttz//07c//9N3P//Tdv//03b//9M2///TNv//0za//9L2v//S9r//0vZ//9K2f//Stn//0rZ//9J2P//Sdj//0nY//9I1///SNf//0jX//9H1///R9b//0fW//9G1v//RtX//0bV//9F1f//RdX//0XU//9E1P//RNT//0TT//9D0///Q9P//0PT//9C0v//QtL//0LS//9B0f//QdH//0HR//9A0f//QND//0DQ//8/0P//P8///z/P//8+z///Ps///z7O//89zv//Pc7//z3N//88zf//PM3//zzN//88zP//O8z//zvM//87y///Osv//zrL//86y///Ocr//znK//85yv//OMn//1bQ/f+55vf/6vH0/+7y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7vP2/+Lf2P+6lmr/p3U5/6ZzNv+mczf/pXM3/6VzN/+kczf/pHI3/6RyN/+jcjf/o3I3/6NyN/+icTf/onE3/6FxN/+hcTf/oXA3/6BwN/+gbzb/pHZA/7+lhP/m5eL/7vP2/8nt9v9m4P3/TN3//07d//9O3f//Ttz//07c//9N3P//Tdv//03b//9M2///TNv//0za//9L2v//S9r//0vZ//9K2f//Stn//0rZ//9J2P//Sdj//0nY//9I1///SNf//0jX//9H1///R9b//0fW//9G1v//RtX//0bV//9F1f//RdX//0XU//9E1P//RNT//0TT//9D0///Q9P//0PT//9C0v//QtL//0LS//9B0f//QdH//0HR//9A0f//QND//0DQ//8/0P//P8///z/P//8+z///Ps///z7O//89zv//Pc7//z3N//88zf//PM3//zzN//87zP//O8z//zvM//87y///Osv//zrL//86yv//Ocr//znK//83yf//Qsz+/4jb+v/d7vX/7/L0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9f/s7/D/2My8/7CFUP+mczX/pnQ3/6ZzN/+mczf/pXM3/6VzN/+kczf/pHI3/6RyN/+jcjf/o3I3/6NyN/+icTf/onE3/6FxN/+hcTf/oXA3/6BvNv+kdkD/v6SD/+bl4v/v9Pb/z+/2/2jh/f9N3f//T93//07d//9O3f//Ttz//03c//9N3P//Tdv//03b//9M2///TNv//0za//9L2v//S9r//0vZ//9K2f//Stn//0rZ//9J2P//Sdj//0nY//9I1///SNf//0jX//9H1///R9b//0fW//9G1v//RtX//0bV//9F1f//RdX//0XU//9E1P//RNT//0TT//9D0///Q9P//0PT//9C0v//QtL//0LS//9B0f//QdH//0HR//9A0f//QND//0DQ//8/0P//P8///z/P//8+z///Ps///z7O//89zv//Pc7//z3N//88zf//PM3//zzM//87zP//O8z//zvM//86y///Osv//zrL//86yv//Ocr//znK//85yv//X9L8/8rq9v/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7vP2/+jp5//Ks5b/rH1E/6ZzNf+ndDf/pnQ3/6ZzN/+mczf/pXM3/6VzN/+kczf/pHI3/6RyN/+jcjf/o3I3/6NxN/+icTf/onE3/6FxN/+hcTf/oHA2/6N1P/+9on//5eXh//D09v/V7/b/b+L9/07e//9P3f//T93//07d//9O3f//Ttz//03c//9N3P//Tdv//0zb//9M2///TNv//0za//9L2v//S9r//0vZ//9K2f//Stn//0rZ//9J2P//Sdj//0nY//9I1///SNf//0jX//9H1///R9b//0fW//9G1v//RtX//0bV//9F1f//RdX//0XU//9E1P//RNT//0TT//9D0///Q9P//0PT//9C0v//QtL//0LS//9B0f//QdH//0HR//9A0f//QND//0DQ//8/0P//P8///z/P//8+z///Ps7//z7O//89zv//Pc7//z3N//88zf//PM3//zzM//87zP//O8z//zvM//86y///Osv//zrL//85yv//Ocr//zfK//9Mzv7/p+L4/+Xw9P/u8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/u9Pb/5ePd/76cc/+pdzz/p3Q2/6d0N/+ndDf/pnQ3/6ZzN/+mczf/pXM3/6VzN/+kczf/pHI3/6RyN/+jcjf/o3I3/6NxN/+icTf/onE3/6FxN/+hcDb/o3Q8/7qcd//k497/8PT2/9rv9f+B5Pz/U97//0/e//9P3f//T93//07d//9O3f//Ttz//03c//9N3P//Tdv//0zb//9M2///TNv//0va//9L2v//S9r//0vZ//9K2f//Stn//0rZ//9J2P//Sdj//0nY//9I1///SNf//0jX//9H1///R9b//0fW//9G1v//RtX//0bV//9F1f//RdX//0XU//9E1P//RNT//0TT//9D0///Q9P//0PT//9C0v//QtL//0LS//9B0f//QdH//0HR//9A0P//QND//0DQ//8/0P//P8///z/P//8+z///Ps7//z7O//89zv//Pc7//z3N//88zf//PM3//zzM//87zP//O8z//zvM//86y///Osv//zrL//85yv//OMr//0LM/v+D2/v/3O71/+/y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+7z9v/g29L/tYxb/6h0Nv+odDf/p3Q3/6d0N/+ndDf/pnQ3/6ZzN/+lczf/pXM3/6VzN/+kczf/pHI3/6RyN/+jcjf/o3I3/6NxN/+icTf/onE3/6FxN/+icjn/tZNp/+Pg2v/v9Pf/4fD1/5vo+v9a3/7/Tt7//0/e//9P3f//T93//07d//9O3f//Ttz//03c//9N3P//Tdv//0zb//9M2///TNv//0va//9L2v//S9r//0rZ//9K2f//Stn//0rZ//9J2P//Sdj//0nY//9I1///SNf//0jX//9H1///R9b//0fW//9G1v//RtX//0bV//9F1f//RdX//0XU//9E1P//RNT//0TT//9D0///Q9P//0PS//9C0v//QtL//0LS//9B0f//QdH//0HR//9A0P//QND//0DQ//8/0P//P8///z/P//8+z///Ps7//z7O//89zv//Pc7//z3N//88zf//PM3//zzM//87zP//O8z//zvM//86y///Osv//zrL//85yv//Osr//2fU/P/T7Pb/7/P0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vX/7PDx/9nNvf+xhU//p3M0/6h1N/+odDf/p3Q3/6d0N/+ndDf/pnQ3/6ZzN/+lczf/pXM3/6VzN/+kczf/pHI3/6RyN/+jcjf/o3I3/6NxN/+icTf/onE3/6BwNf+uh1f/3tnO/+7z9f/p8vT/vez3/2bh/f9P3v//UN7//0/e//9P3f//T93//07d//9O3f//Ttz//03c//9N3P//Tdv//0zb//9M2///TNv//0va//9L2v//S9r//0rZ//9K2f//Stn//0nZ//9J2P//Sdj//0nY//9I1///SNf//0jX//9H1///R9b//0fW//9G1v//RtX//0bV//9F1f//RdT//0XU//9E1P//RNT//0TT//9D0///Q9P//0PS//9C0v//QtL//0LS//9B0f//QdH//0HR//9A0P//QND//0DQ//8/0P//P8///z/P//8+z///Ps7//z7O//89zv//Pc7//z3N//88zf//PM3//zzM//87zP//O8z//zvM//86y///Osv//zrL//83yv//VND9/8Ho9//s8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3z9f/q7Ov/0b6l/6+BSf+odDX/qHU3/6h1N/+odDf/p3Q3/6d0N/+ndDf/pnQ3/6ZzN/+lczf/pXM3/6VzN/+kczf/pHI3/6RyN/+jcjf/o3I3/6NxN/+icTf/oW81/6qATP/TxrP/6+7u/+/z9f/Y7/X/heX8/1Tf//9P3v//UN7//0/e//9P3f//T93//07d//9O3f//Ttz//03c//9N3P//Tdv//0zb//9M2///TNv//0va//9L2v//S9r//0rZ//9K2f//Stn//0nZ//9J2P//Sdj//0jY//9I1///SNf//0jX//9H1v//R9b//0fW//9G1v//RtX//0bV//9F1f//RdT//0XU//9E1P//RNT//0TT//9D0///Q9P//0PS//9C0v//QtL//0LS//9B0f//QdH//0HR//9A0P//QND//0DQ//8/0P//P8///z/P//8+z///Ps7//z7O//89zv//Pc7//z3N//88zf//PM3//zzM//87zP//O8z//zvM//86y///Osv//zjK//9Mzv7/qeP4/+bw9P/u8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7vP2/+jp5v/Ls5T/rn5E/6l0Nf+pdTf/qHU3/6h1N/+odDf/p3Q3/6d0N/+ndDf/pnQ3/6ZzN/+lczf/pXM3/6VzN/+kczf/pHI3/6RyN/+jcjf/o3I3/6NxN/+icDb/pnhB/8Kpif/m5uP/7/P2/+jx9P+47Pj/auL9/1Df//9Q3v//UN7//0/e//9P3f//T93//07d//9O3f//Ttz//03c//9N3P//Tdv//0zb//9M2///TNr//0va//9L2v//S9r//0rZ//9K2f//Stn//0nY//9J2P//Sdj//0jY//9I1///SNf//0fX//9H1v//R9b//0fW//9G1v//RtX//0bV//9F1f//RdT//0XU//9E1P//RNT//0TT//9D0///Q9P//0PS//9C0v//QtL//0LS//9B0f//QdH//0HR//9A0P//QND//0DQ//8/0P//P8///z/P//8+z///Ps7//z7O//89zv//Pc7//z3N//88zf//PM3//zzM//87zP//O8z//zvM//86y///OMv//0jO/v+X3/n/4e/1/+7y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/u8/b/5+bi/8aphv+tfED/qXU2/6l1N/+pdTf/qHU3/6h1N/+odDf/p3Q3/6d0N/+ndDf/pnQ3/6ZzN/+lczf/pXM3/6VzN/+kczf/pHI3/6RyN/+jcjf/o3I3/6NxN/+icTf/soxf/97Yz//t8vT/7vL0/9/w9f+f6fr/YuH+/0/f//9Q3v//UN7//0/e//9P3f//T93//07d//9O3P//Ttz//03c//9N3P//Tdv//0zb//9M2///TNr//0va//9L2v//S9r//0rZ//9K2f//Stn//0nY//9J2P//Sdj//0jY//9I1///SNf//0fX//9H1v//R9b//0bW//9G1v//RtX//0bV//9F1f//RdT//0XU//9E1P//RNT//0TT//9D0///Q9P//0PS//9C0v//QtL//0LS//9B0f//QdH//0HR//9A0P//QND//0DQ//8/0P//P8///z/P//8+z///Ps7//z7O//89zv//Pc7//z3N//88zf//PM3//zzM//87zP//O8z//zvM//85y///RM3+/4jc+v/d7/X/7/L0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+709v/l5N7/waB3/6x6Pf+qdTb/qnY3/6l1N/+pdTf/qHU3/6h1N/+odDf/p3Q3/6d0N/+ndDf/pnQ3/6ZzN/+lczf/pXM3/6VzN/+kczf/pHI3/6RyN/+jcjf/o3I3/6JwNf+pfEj/zLmg/+nr6f/u8/X/7fL0/9jw9v+Z6Pr/Y+H+/1De//9P3v//UN7//0/e//9P3f//T93//07d//9O3P//Ttz//03c//9N3P//Tdv//0zb//9M2///TNr//0va//9L2v//S9r//0rZ//9K2f//Stn//0nY//9J2P//Sdj//0jY//9I1///SNf//0fX//9H1v//R9b//0bW//9G1v//RtX//0XV//9F1f//RdT//0XU//9E1P//RNT//0TT//9D0///Q9P//0PS//9C0v//QtL//0LS//9B0f//QdH//0HR//9A0P//QND//0DQ//8/0P//P8///z/P//8+z///Ps7//z7O//89zv//Pc7//z3N//88zf//PM3//zzM//87zP//O8z//zrM//9AzP//edj7/9nu9f/v8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7vT3/+Ti2/+9mGv/rHg5/6p2N/+qdjf/qnY3/6l1N/+pdTf/qHU3/6h1N/+odDf/p3Q3/6d0N/+ndDf/pnQ3/6ZzN/+lczf/pXM3/6VzN/+kczf/pHI3/6RyN/+jcjf/o3E2/6RzOf+0kWX/3NXJ/+3y8//u8/X/7fL0/9nw9f+l6fn/b+P9/1jf//9P3v//Tt7//0/e//9P3f//T93//07d//9O3P//Ttz//03c//9N3P//Tdv//0zb//9M2///TNr//0va//9L2v//S9r//0rZ//9K2f//Stn//0nY//9J2P//Sdj//0jY//9I1///SNf//0fX//9H1v//R9b//0bW//9G1v//RtX//0XV//9F1f//RdT//0TU//9E1P//RNT//0TT//9D0///Q9P//0PS//9C0v//QtL//0LS//9B0f//QdH//0HR//9A0P//QND//0DQ//8/0P//P8///z/P//8+z///Ps7//z7O//89zv//Pc7//z3N//88zf//PM3//zzM//87zP//Osz//z3M//9t1vz/1u31//Dz9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/u9Pf/5OHa/7uUZf+rdzj/q3Y3/6p2N/+qdjf/qnY3/6l1N/+pdTf/qHU3/6h1N/+odDf/p3Q3/6d0N/+ndDf/pnQ3/6ZzN/+lczf/pXM3/6VzN/+kczf/pHI3/6RyN/+jcjf/o3E2/6d5Qv/BpoT/4uDZ/+3z9f/u8vX/7fL0/+Hx9f/C7ff/kub6/23i/f9e4P7/Vd///0/e//9N3f//Td3//0zc//9M3P//TNz//0vc//9L2///S9v//0rb//9K2v//Str//0na//9J2v//Sdn//0jZ//9I2f//SNj//0fY//9H2P//R9j//0bX//9G1///Rtf//0XW//9F1v//Rdb//0TW//9E1f//RNX//0TV//9E1f//RdT//0TU//9E1P//RNT//0PT//9D0///Q9P//0PS//9C0v//QtL//0LS//9B0f//QdH//0HR//9A0P//QND//0DQ//8/0P//P8///z/P//8+z///Ps7//z7O//89zv//Pc3//z3N//88zf//PM3//zzM//87zP//PMz//2fV/P/U7fb/8PP0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+709//k4Nj/upFg/6t2Nv+rdjf/q3Y3/6p2N/+qdjf/qnY3/6l1N/+pdTf/qHU3/6h1N/+odDf/p3Q3/6d0N/+ndDf/pnQ3/6ZzN/+lczf/pXM3/6VzN/+kczf/pHI3/6RyN/+jcjf/o3I3/6p+Sv/Gr5H/4+Da/+3y8//u8/X/7vL0/+ry9P/e8PX/ye73/6jp+f+H5fv/ceL9/2bh/f9j4P7/Yd/+/2Df/v9g3/7/YN7+/1/e/v9f3v7/X97+/1/d/v9e3f7/Xt3+/17d/v9d3f7/Xdz+/13c/v9d3P7/XNz+/1zb/v9c2/7/W9v+/1vb/v9b2v7/Wtr+/1ra/v9a2f7/Wdn+/1fZ/v9U2P7/UNf+/0rW//9E1f//Q9T//0TU//9E1P//RNT//0PT//9D0///Q9P//0LS//9C0v//QtL//0LS//9B0f//QdH//0HR//9A0P//QND//0DQ//8/z///P8///z/P//8+z///Ps7//z7O//89zv//Pc3//z3N//88zf//PM3//zzM//87zP//YtT9/9Pt9v/w8/T/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7vT3/+Pf1/+5jlv/q3Y1/6t3N/+rdjf/q3Y3/6p2N/+qdjf/qnY3/6l1N/+pdTf/qHU3/6h1N/+odDf/p3Q3/6d0N/+ndDf/pnQ3/6ZzN/+lczf/pXM3/6VzN/+kczf/pHI3/6RyN/+jcjf/o3I3/6p+Sv/Cp4f/3djN/+rt7P/u8/X/7vP2/+/z9f/s8vT/5PH1/9zw9f/U7/b/yO32/77s9/+46/j/t+r4/7fq+P+36vj/t+r4/7fq+P+36vj/t+r4/7bp+P+26fj/tun4/7bp+P+26fj/tun4/7bp+P+26fj/ten4/7Xp+P+16fj/ten4/7Xp+P+16Pj/tOj4/7To+P+y6Pj/rOf4/6Pl+f+T4vr/fN77/2Ta/f9U1/7/SNX//0PU//9E1P//RNP//0PT//9D0///Q9P//0LS//9C0v//QtL//0HR//9B0f//QdH//0HR//9A0P//QND//0DQ//8/z///P8///z/P//8+z///Ps7//z7O//89zv//Pc3//z3N//88zf//PM3//zrM//9d0/3/0uz2//Dz9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/u9Pf/49/W/7iMWP+rdTT/rHc3/6t3N/+rdjf/q3Y3/6p2N/+qdjf/qXY3/6l1N/+pdTf/qHU3/6h1N/+odDf/p3Q3/6d0N/+ndDf/pnQ3/6ZzN/+lczf/pXM3/6VzN/+kczf/pHI3/6RyN/+jcjf/o3I3/6h6RP+3lWz/0MCq/+He1v/o6ef/6+/v/+7z9f/v9Pf/8PT3//D09v/t9Pb/6/P3/+nz9//p8/f/6fP3/+nz9//p8/f/6fP3/+nz9//p8/f/6fP3/+nz9//p8/f/6fP3/+nz9//p8/f/6fP3/+nz9//p8/f/6fP3/+nz9//p8/f/6fP3/+nz9//p8/f/6fL2/+jy9v/m8fX/4/D1/9/w9f/Z7/X/yez3/6Tm+f923fz/Vdf+/0XU//9D1P//RNP//0PT//9D0///Q9P//0LS//9C0v//QtL//0HR//9B0f//QdH//0DR//9A0P//QND//z/Q//8/z///P8///z/P//8+z///Ps7//z7O//89zv//Pc3//z3N//88zf//Osz//1vT/f/R7Pb/8PP0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+709//j39f/uI1Z/6x1Nf+sdzf/rHc3/6t3N/+rdjf/q3Y3/6p2N/+qdjf/qXY3/6l1N/+pdTf/qHU3/6h1N/+odDf/p3Q3/6d0N/+ndDf/pnQ3/6ZzN/+lczf/pXM3/6VzN/+kczf/pHI3/6RyN/+jcjf/o3E2/6R0O/+rf0z/t5Rr/8exlf/Vybj/3tnP/+Lf2f/j4dz/5OPd/+Tj3v/l497/5ePe/+Xj3v/l497/5ePe/+Xj3v/l497/5ePe/+Tj3v/k497/5OPe/+Tj3v/k497/5OPe/+Tj3v/k497/5OPe/+Tj3v/k497/5OPe/+Tj3v/k497/5OPf/+Xk4P/m5uT/6evp/+zw8P/v8/b/7/T2/+/z9P/t8vT/5PH1/9Ht9v+g5fn/ZNr9/0fV//9E1P//RNP//0PT//9D0///Q9P//0LS//9C0v//QtL//0HR//9B0f//QdH//0DR//9A0P//QND//z/Q//8/z///P8///z7P//8+z///Ps7//z7O//89zv//Pc3//z3N//86zf//XNT9/9Hs9v/w8/T/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7vT3/+Pg2P+6kF7/rHY2/6x3N/+sdzf/rHc3/6t3N/+rdjf/q3Y3/6p2N/+qdjf/qXY3/6l1N/+pdTf/qHU3/6h1N/+odDf/p3Q3/6d0N/+ndDf/pnQ3/6ZzN/+lczf/pXM3/6VzN/+kczf/pHI3/6RyN/+jcjf/o3E2/6JwNf+jczr/p3pE/6qATv+uh1f/s49k/7aVbf+5mnP/upx3/7qdd/+6nHf/uZx3/7mcd/+5nHf/uJx3/7icd/+4nHf/uJx3/7ecd/+3m3f/t5t3/7ebd/+2m3f/tpt3/7aad/+2mnf/tZp3/7Wad/+1mXf/tZl3/7WZd/+1mnj/t55+/72ni//HtqH/08q8/9/d1v/m6Ob/7PDx/+7z9f/u8/T/7vL0/+Hw9f+16Pj/bdv8/0jV//9D1P//RNP//0PT//9D0///Q9P//0LS//9C0v//QtL//0HR//9B0f//QdH//0DR//9A0P//QND//z/Q//8/z///P8///z7P//8+z///Ps7//z3O//89zv//Pc3//zzN//9h1f3/0+z2//Dz9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/u9Pf/5OHZ/72UY/+teDf/rXc3/6x3N/+sdzf/rHc3/6t3N/+rdjf/q3Y3/6p2N/+qdjf/qXY3/6l1N/+pdTf/qHU3/6h1N/+odDf/p3Q3/6d0N/+ndDf/pnQ3/6ZzN/+lczf/pXM3/6VzN/+kczf/pHI3/6RyN/+jcjf/o3I3/6NxNv+icDX/oW81/6BwNf+hcTj/onI6/6JzPP+idD3/onQ9/6J0Pf+hcz3/oXM9/6BzPf+gcz3/oHI9/59yPf+fcjz/n3I8/55yPP+ecTz/nnE8/51xPP+dcTz/nHE8/5xwPP+ccDz/m3A8/5twPP+bbzz/mm88/5pvPP+bcD7/nHNB/552R/+ifFD/rI1n/8Ctk//Vz8L/5ebi/+zx8//u8/X/7vL0/+bx9P+46ff/adr9/0bU//9D1P//RNP//0PT//9D0///Q9P//0LS//9C0v//QtL//0HR//9B0f//QdH//0DR//9A0P//QND//z/Q//8/z///P8///z7P//8+z///Ps7//z3O//89zv//Pc3//2XW/P/U7fb/8PP0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+709//k4tr/vpho/655Of+tdzf/rXc3/6x3N/+sdzf/rHc3/6t3N/+rdjf/q3Y3/6p2N/+qdjf/qXY3/6l1N/+pdTf/qHU3/6h1N/+odDf/p3Q3/6d0N/+ndDf/pnQ3/6ZzN/+lczf/pXM3/6VzN/+kczf/pHI3/6RyN/+jcjf/o3I3/6NxN/+icTf/onE3/6FxN/+hcDb/oHA2/6BvNv+gbzb/n282/59vNv+ebzb/nm42/55uNv+dbjb/nW42/51tNv+cbTX/nG01/5xtNf+bbTX/m2w1/5psNf+abDX/mmw1/5lsNf+ZazX/mWs1/5hrNf+YazX/mGo1/5dqNf+XajX/lmk0/5VpNP+Xazj/nHND/6WDWP+9qY7/3NjO/+vu7//t8vT/7vL0/+Pw9f+p5vn/Wdj+/0TU//9E1P//RNP//0PT//9D0///Q9P//0LS//9C0v//QtL//0HR//9B0f//QdH//0DR//9A0P//QND//z/Q//8/z///P8///z7P//8+z///Ps7//z3O//8+zv//atf8/9Xt9f/w8/T/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7vT3/+Xj3f/BnXD/r3o7/654Nv+teDf/rXc3/6x3N/+sdzf/rHc3/6t3N/+rdjf/q3Y3/6p2N/+qdjf/qXY3/6l1N/+pdTf/qHU3/6h1N/+odDf/p3Q3/6d0N/+ndDf/pnQ3/6ZzN/+lczf/pXM3/6VzN/+kczf/pHI3/6RyN/+jcjf/o3I3/6NxN/+icTf/onE3/6FxN/+hcTf/oXA3/6BwN/+gcDf/oHA3/59wN/+fbzf/nm83/55vN/+ebzf/nW43/51uN/+dbjb/nG42/5xuNv+cbTb/m202/5ttNv+abTb/mm02/5psNv+ZbDb/mWw2/5lsNv+Yazb/mGs2/5hrNv+Xazb/l2s2/5ZqNv+VaTT/lmo2/5t0Q/+vk2//1Mu+/+rt7f/t8vT/7vL0/9fu9v+F4Pv/TNX//0PU//9E1P//RNP//0PT//9D0///Q9P//0LS//9C0v//QtL//0HR//9B0f//QdH//0DR//9A0P//QND//z/Q//8/z///P8///z7P//8+z///Pc7//0HP//9y2fz/1+71/+/y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/u8/b/5+bg/8anf/+xfT//rnc2/654N/+teDf/rXc3/6x3N/+sdzf/rHc3/6t3N/+rdjf/q3Y3/6p2N/+qdjf/qXY3/6l1N/+pdTf/qHU3/6h1N/+odDf/p3Q3/6d0N/+ndDf/pnQ3/6ZzN/+lczf/pXM3/6VzN/+kczf/pHI3/6RyN/+jcjf/o3I3/6NxN/+icTf/onE3/6FxN/+hcTf/oXA3/6BwN/+gcDf/oHA3/59wN/+fbzf/nm83/55vN/+ebzf/nW43/51uN/+dbjb/nG42/5xuNv+cbTb/m202/5ttNv+abTb/mm02/5psNv+ZbDb/mWw2/5lsNv+Yazb/mGs2/5drNv+Xazb/l2s2/5ZqNv+Wajb/lWk1/5huO/+sjWf/1cy//+vv7//u8/X/6fH0/7jp9/9c2f7/Q9T//0TU//9E1P//RNP//0PT//9D0///Q9P//0LS//9C0v//QtL//0HR//9B0f//QdH//0DR//9A0P//QND//z/Q//8/z///P8///z7P//89zv//Rc///3/c+//b7vX/7/L0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+7z9v/o6OX/zLKQ/7OBQ/+udzb/rng3/654N/+teDf/rXc3/6x3N/+sdzf/rHc3/6t3N/+rdjf/q3Y3/6p2N/+qdjf/qXY3/6l1N/+pdTf/qHU3/6h1N/+odDf/p3Q3/6d0N/+ndDf/pnQ3/6ZzN/+lczf/pXM3/6VzN/+kczf/pHI3/6RyN/+jcjf/o3I3/6NxN/+icTf/onE3/6FxN/+hcTf/oXA3/6BwN/+gcDf/oHA3/59wN/+fbzf/nm83/55vN/+ebzf/nW43/51uNv+dbjb/nG42/5xuNv+cbTb/m202/5ttNv+abTb/mm02/5psNv+ZbDb/mWw2/5lsNv+Yazb/mGs2/5drNv+Xazb/l2s2/5ZqNv+Wajb/lWk1/5huPP+xlXL/3dnS/+3x8//v8vT/1u72/3je/P9J1f//RNT//0TU//9E1P//RNP//0PT//9D0///Q9P//0LS//9C0v//QtL//0HR//9B0f//QdH//0DQ//9A0P//QND//z/Q//8/z///P8///z3O//9I0P7/jt76/97v9f/v8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fP1/+rr6v/SvaL/tIRI/654Nf+ueDf/rng3/654N/+teDf/rXc3/6x3N/+sdzf/rHc3/6t3N/+rdjf/q3Y3/6p2N/+qdjf/qXY3/6l1N/+pdTf/qHU3/6h1N/+odDf/p3Q3/6d0N/+ndDf/pnQ3/6ZzN/+lczf/pXM3/6VzN/+kczf/pHI3/6RyN/+jcjf/o3I3/6JxN/+icTf/onE3/6FxN/+hcTf/oXA3/6BwN/+gcDf/oHA3/59wN/+fbzf/nm83/55vN/+ebzf/nW43/51uNv+dbjb/nG42/5xuNv+cbTb/m202/5ttNv+abTb/mm02/5psNv+ZbDb/mWw2/5lsNv+Yazb/mGs2/5drNv+Xazb/l2s2/5ZqNv+Wajb/lWk1/511Rv/CsJf/5+jm/+/z9f/j8PX/n+X5/1PX/v9D1P//RdT//0TU//9E1P//RNP//0PT//9D0///Q9L//0LS//9C0v//QtL//0HR//9B0f//QdH//0DQ//9A0P//QND//z/Q//8/z///Pc///0zS/v+c4vn/4vD1/+7y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vX/6+/v/9nKt/+3h07/rnc1/695N/+ueDf/rng3/614N/+teDf/rXc3/6x3N/+sdzf/rHc3/6t3N/+rdjf/q3Y3/6p2N/+qdjf/qXY3/6l1N/+pdTf/qHU3/6h1N/+odDf/p3Q3/6d0N/+ndDf/pnQ3/6ZzN/+lczf/pXM3/6VzN/+kczf/pHI3/6RyN/+jcjf/o3I3/6JxN/+icTf/onE3/6FxN/+hcTf/oXA3/6BwN/+gcDf/oHA3/59wN/+fbzf/nm83/55vN/+ebzf/nW43/51uNv+dbjb/nG42/5xuNv+cbTb/m202/5ttNv+abTb/mm02/5psNv+ZbDb/mWw2/5lsNv+Yazb/mGs2/5drNv+Xazb/l2s2/5ZqNv+Wajb/lmo2/6iIX//b187/7fL0/+vy9P+/6vf/XNn+/0PU//9F1P//RdT//0TU//9E1P//RNP//0PT//9D0///Q9L//0LS//9C0v//QtL//0HR//9B0f//QdH//0DQ//9A0P//QND//z/Q//89z///UtP+/6/m+P/n8fT/7vL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8/X/4drP/7uPWv+veDb/r3k3/695N/+ueDf/rng3/614N/+teDf/rXc3/6x3N/+sdzf/rHc3/6t3N/+rdjf/q3Y3/6p2N/+qdjf/qXY3/6l1N/+pdTf/qHU3/6h1N/+odDf/p3Q3/6d0N/+ndDf/pnQ3/6ZzN/+lczf/pXM3/6VzN/+kczf/pHI3/6RyN/+jcjf/o3I3/6JxN/+icTf/onE3/6FxN/+hcTf/oXA3/6BwN/+gcDf/oHA3/59wN/+fbzf/nm83/55vN/+ebzf/nW43/51uNv+dbjb/nG42/5xuNv+cbTb/m202/5ttNv+abTb/mm02/5psNv+ZbDb/mWw2/5lsNv+Yazb/mGs2/5drNv+Xazb/l2s2/5ZqNv+VaTT/nndI/8m7pv/p7Ov/7/P1/9Lt9v9p2/3/RdX//0XV//9F1P//RdT//0TU//9E1P//RNP//0PT//9D0///Q9L//0LS//9C0v//QtL//0HR//9B0f//QdH//0DQ//9A0P//QND//z3P//9a1f3/xer3/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+709v/m493/w59z/7F8PP+veTf/r3k3/695N/+ueDf/rng3/614N/+teDf/rXc3/6x3N/+sdzf/rHc3/6t3N/+rdjf/q3Y3/6p2N/+qdjf/qXY3/6l1N/+pdTf/qHU3/6h1N/+odDf/p3Q3/6d0N/+ndDf/pnQ3/6ZzN/+lczf/pXM3/6VzN/+kczf/pHI3/6RyN/+jcjf/o3I3/6JxN/+icTf/onE3/6FxN/+hcTf/oXA3/6BwN/+gcDf/oHA3/59wN/+fbzf/nm83/55vN/+ebzf/nW43/51uNv+dbjb/nG42/5xuNv+cbTb/m202/5ttNv+abTb/mm02/5psNv+ZbDb/mWw2/5lsNv+Yazb/mGs2/5drNv+Xazb/l2s2/5ZpNf+acD//uaCC/+Tl4f/w9Pb/2O71/3je/P9J1v//RdX//0XV//9F1P//RdT//0TU//9E1P//RNP//0PT//9D0///Q9L//0LS//9C0v//QtL//0HR//9B0f//QdH//0DQ//9A0P//QdD//2vY/P/U7fb/7/L0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7vP2/+np5v/OtpX/tIJF/694Nv+veTj/r3k3/695N/+ueDf/rng3/614N/+teDf/rXc3/6x3N/+sdzf/rHc3/6t3N/+rdjf/q3Y3/6p2N/+qdjf/qXY3/6l1N/+pdTf/qHU3/6h1N/+odDf/p3Q3/6d0N/+ndDf/pnQ3/6ZzN/+lczf/pXM3/6VzN/+kczf/pHI3/6RyN/+jcjf/o3I3/6JxN/+icTf/onE3/6FxN/+hcTf/oXA3/6BwN/+gcDf/oHA3/59wN/+fbzf/nm83/55vN/+ebzf/nW43/51uNv+dbjb/nG42/5xuNv+cbTb/m202/5ttNv+abTb/mm02/5psNv+ZbDb/mWw2/5lsNv+Yazb/mGs2/5drNv+Xazb/l2o2/5drOP+tjWf/4eDa//D09//b7/X/guD7/0zX//9F1f//RtX//0XV//9F1P//RdT//0TU//9E1P//RNP//0PT//9D0///Q9L//0LS//9C0v//QtL//0HR//9B0f//QdH//z/Q//9I0v7/ht77/9zv9f/v8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vX/7PDx/9zQvv+5i1P/r3g2/7B5OP+veTj/r3k3/695N/+ueDf/rng3/614N/+teDf/rXc3/6x3N/+sdzf/rHc3/6t3N/+rdjf/q3Y3/6p2N/+qdjf/qXY3/6l1N/+pdTf/qHU3/6h1N/+odDf/p3Q3/6d0N/+ndDf/pnQ3/6ZzN/+lczf/pXM3/6VzN/+kczf/pHI3/6RyN/+jcjf/o3I3/6JxN/+icTf/onE3/6FxN/+hcTf/oXA3/6BwN/+gcDf/oHA3/59wN/+fbzf/nm83/55vN/+ebzf/nW43/51uNv+dbjb/nG42/5xuNv+cbTb/m202/5ttNv+abTb/mm02/5psNv+ZbDb/mWw2/5lsNv+Yazb/mGs2/5drNv+Xazb/lmk0/6aCV//e29T/8PT3/93v9f+J4fv/Ttf//0XW//9G1f//RdX//0XV//9F1P//RdT//0TU//9E1P//RNP//0PT//9D0///Q9L//0LS//9C0v//QtL//0HR//9B0f//P9D//1HU/v+o5Pj/5fD1/+7y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/u8/b/5eLa/8Ofcf+yfDz/sHk3/7B5OP+veTj/r3k3/695N/+ueDf/rng3/614N/+teDf/rXc3/6x3N/+sdzf/rHc3/6t3N/+rdjf/q3Y3/6p2N/+qdjf/qXY3/6l1N/+pdTf/qHU3/6h1N/+odDf/p3Q3/6d0N/+ndDf/pnQ3/6ZzN/+lczf/pXM3/6VzN/+kczf/pHI3/6RyN/+jcjf/o3I3/6JxN/+icTf/onE3/6FxN/+hcTf/oXA3/6BwN/+gcDf/oHA3/59wN/+fbzf/nm83/55vN/+ebzf/nW43/51uNv+dbjb/nG42/5xuNv+cbTb/m202/5ttNv+abTb/mm02/5psNv+ZbDb/mWw2/5lsNv+Yazb/mGs2/5drNv+WaTP/pH9S/9vXzf/v8/X/3e/1/4rh+/9P1/7/Rdb//0bW//9G1f//RdX//0XV//9F1P//RNT//0TU//9E1P//RNP//0PT//9D0///Q9L//0LS//9C0v//QtL//0HR//9A0f//Ydf9/8jr9//t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3z9f/q7Or/076i/7eGSv+weTf/sHk4/7B5OP+veTj/r3k3/695N/+ueDf/rng3/614N/+teDf/rXc3/6x3N/+sdzf/rHc3/6t3N/+rdjf/q3Y3/6p2N/+qdjf/qXY3/6l1N/+pdTf/qHU3/6h1N/+odDf/p3Q3/6d0N/+ndDf/pnQ3/6ZzN/+lczf/pXM3/6VzN/+kczf/pHI3/6RyN/+jcjf/o3I3/6JxN/+icTf/onE3/6FxN/+hcTf/oXA3/6BwN/+gcDf/oHA3/59wN/+fbzf/nm83/55vN/+ebzf/nW43/51uNv+dbjb/nG42/5xuNv+cbTb/m202/5ttNv+abTb/mm02/5psNv+ZbDb/mWw2/5lsNv+Yazb/mGs2/5ZpM/+jflH/2dPH/+/y8//d7/X/i+H7/0/X/v9G1v//Rtb//0bW//9G1f//RdX//0XV//9F1P//RNT//0TU//9E1P//Q9P//0PT//9D0///Q9L//0LS//9C0v//QdL//0jT//+D3vv/2+/1/+/y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/i3NP/wZts/7J8O/+weTf/sHk4/7B5OP+veTj/r3k3/695N/+ueDf/rng3/614N/+teDf/rXc3/6x3N/+sdzf/rHc3/6t3N/+rdjf/q3Y3/6p2N/+qdjf/qXY3/6l1N/+pdTf/qHU3/6h1N/+odDf/p3Q3/6d0N/+ndDf/pnQ3/6ZzN/+lczf/pXM3/6VzN/+kczf/pHI3/6RyN/+jcjf/o3I3/6JxN/+icTf/onE3/6FxN/+hcTf/oXA3/6BwN/+gcDf/oHA3/59wN/+fbzf/nm83/55vN/+ebzf/nW43/51uNv+dbjb/nG42/5xuNv+bbTb/m202/5ttNv+abTb/mm02/5psNv+ZbDb/mWw2/5lsNv+Yazb/lmkz/6N+UP/Y0cb/7vLz/93v9f+L4fv/T9j+/0bW//9H1v//Rtb//0bW//9G1f//RdX//0XV//9F1P//RNT//0TU//9E1P//Q9P//0PT//9D0///QtL//0LS//9B0v//WNb+/7Ln+P/o8fT/7vL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fP1/+rs6//VwKb/uIhO/7B5N/+wejj/sHk4/7B5OP+veTj/r3k3/695N/+ueDf/rng3/614N/+teDf/rXc3/6x3N/+sdzf/rHc3/6t3N/+rdjf/q3Y3/6p2N/+qdjf/qXY3/6l1N/+pdTf/qHU3/6h1N/+odDf/p3Q3/6d0N/+ndDf/pnQ3/6ZzN/+lczf/pXM3/6VzN/+kczf/pHI3/6RyN/+jcjf/o3I3/6JxN/+icTf/onE3/6FxN/+hcTf/oXA3/6BwN/+gcDf/oHA3/59wN/+fbzf/nm83/55vN/+ebzf/nW43/51uNv+dbjb/nG42/5xuNv+bbTb/m202/5ttNv+abTb/mm02/5psNv+ZbDb/mWw2/5lsNv+XaTP/pH5Q/9jRxv/u8vP/3fD1/4vi+/9Q2P7/Rtf//0fW//9H1v//Rtb//0bW//9G1f//RdX//0XV//9F1P//RNT//0TU//9E1P//Q9P//0PT//9D0///QdL//0bT//963Pz/1e72/+7y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+Xi2//JqYH/tIBB/7B5N/+wejj/sHk4/7B5OP+veTj/r3k3/694N/+ueDf/rng3/614N/+teDf/rXc3/6x3N/+sdzf/rHc3/6t3N/+rdjf/q3Y3/6p2N/+qdjf/qXY3/6l1N/+pdTf/qHU3/6h1N/+odDf/p3Q3/6d0N/+mdDf/pnQ3/6ZzN/+lczf/pXM3/6VzN/+kczf/pHI3/6RyN/+jcjf/o3I3/6JxN/+icTf/onE3/6FxN/+hcTf/oXA3/6BwN/+gcDf/oHA3/59wN/+fbzf/nm83/55vN/+ebzf/nW43/51uNv+dbjb/nG42/5xuNv+bbTb/m202/5ttNv+abTb/mm02/5psNv+ZbDb/mWw2/5dqM/+kflD/2NHG/+7y8//d8PX/i+L7/1DY/v9H1///R9f//0fW//9H1v//Rtb//0bW//9G1f//RdX//0XV//9F1P//RNT//0TU//9E0///Q9P//0PT//9C0///W9f9/7Ln+P/n8fT/7vL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vX/7PDy/9/WyP/Bm2v/sn08/7B5N/+wejj/sHk4/7B5OP+veTj/r3k3/694N/+ueDf/rng3/614N/+teDf/rXc3/6x3N/+sdzf/rHc3/6t3N/+rdjf/q3Y3/6p2N/+qdjf/qXU3/6l1N/+pdTf/qHU3/6h1N/+odDf/p3Q3/6d0N/+mdDf/pnQ3/6ZzN/+lczf/pXM3/6VzN/+kczf/pHI3/6RyN/+jcjf/o3I3/6JxN/+icTf/onE3/6FxN/+hcTf/oXA3/6BwN/+gcDf/oHA3/59wN/+fbzf/nm83/55vN/+ebzf/nW43/51uNv+dbjb/nG42/5xuNv+bbTb/m202/5ttNv+abTb/mm02/5psNv+ZbDb/mGoz/6R/UP/Y0sb/7vLz/93w9f+M4vv/UNn+/0fX//9I1///R9f//0fW//9H1v//Rtb//0bV//9G1f//RdX//0XV//9F1P//RNT//0TU//9E0///QtP//07V/v+M4Pr/2u71/+7y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vX/6+7u/9zPvf/AmWj/s309/7B5N/+wejj/sHk4/7B5OP+veTj/r3k3/694N/+ueDf/rng3/614N/+teDf/rXc3/6x3N/+sdzf/rHc3/6t3N/+rdjf/q3Y3/6p2N/+qdjf/qXU3/6l1N/+pdTf/qHU3/6h1N/+odDf/p3Q3/6d0N/+mdDf/pnQ3/6ZzN/+lczf/pXM3/6VzN/+kczf/pHI3/6RyN/+jcjf/o3I3/6JxN/+icTf/onE3/6FxN/+hcTf/oXA3/6BwN/+gcDf/oHA3/59wN/+fbzf/nm83/55vN/+ebzf/nW43/51uNv+dbjb/nG42/5xuNv+bbTb/m202/5ttNv+abTb/mm02/5psNv+YajP/pX9Q/9nSxv/u8vP/3fD1/4zi+/9R2f7/R9f//0jX//9I1///R9f//0fW//9H1v//Rtb//0bV//9G1f//RdX//0XV//9F1P//RNT//0PU//9J1P//dt38/8rs9v/s8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vX/6+3t/9zPvf/DnnD/tYJF/7F6N/+weTf/sHk4/7B5OP+veTj/r3k3/694N/+ueDf/rng3/614N/+teDf/rXc3/6x3N/+sdzf/rHc3/6t3N/+rdjf/q3Y3/6p2N/+qdjf/qXU3/6l1N/+pdTf/qHU3/6h1N/+odDf/p3Q3/6d0N/+mdDf/pnQ3/6ZzN/+lczf/pXM3/6VzN/+kcjf/pHI3/6RyN/+jcjf/o3I3/6JxN/+icTf/onE3/6FxN/+hcTf/oXA3/6BwN/+gcDf/oHA3/59wN/+fbzf/nm83/55vN/+ebzf/nW43/51uNv+dbjb/nG42/5xuNv+bbTb/m202/5ttNv+abTb/mm02/5hqM/+lf1D/2dLG/+7y8//d8PX/jOL7/1HZ/v9I2P//SNf//0jX//9I1///R9f//0fW//9H1v//Rtb//0bV//9G1f//RdX//0XV//9E1P//SdX//3Hc/P++6vf/6fH0/+7y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/6+/v/+Hazv/MsIz/u41X/7R/Qf+weTf/r3g2/694Nv+veDf/r3g2/654Nv+udzb/rXc2/613Nv+tdzb/rHc2/6x2Nv+sdjb/q3Y2/6t2Nv+rdjb/qnU2/6p1Nv+pdTb/qXU2/6l0Nv+odDb/qHQ2/6h0Nv+ndDb/p3M2/6ZzNv+mczb/pnM2/6VzNv+lcjb/pXI2/6RyNv+kcjb/pHE2/6NxNv+jcTb/onE2/6JxNv+icTf/onE3/6FxN/+hcTf/oXA3/6BwN/+gcDf/oHA3/59wN/+fbzf/nm83/55vN/+ebzf/nW43/51uNv+dbjb/nG42/5xuNv+bbTb/m202/5ttNv+abTb/mWsz/6V/UP/Z0sb/7vLz/93w9f+M4/v/Udn+/0jY//9J2P//SNf//0jX//9I1///R9f//0fW//9H1v//Rtb//0bV//9E1f//RdX//1DX/v933vz/vur3/+fx9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vX/7fHz/+fm4f/bzbv/yquG/7ySXf+3iE7/tYJG/7N/Qv+yfkD/sn4//7F9P/+xfT//sH0//7B9P/+wfT//r3w//698P/+vfD//r3w//658P/+uez//rns//617P/+tez//rHs//6x6P/+sej//q3o//6t6P/+rej//qnk//6p5P/+peT//qXk//6l5P/+oeD//qHg//6h4P/+neD//p3c//6d3P/+meD//pXY+/6NyOf+icTf/onE3/6FxN/+hcTf/oXA3/6BwN/+gcDf/oHA3/59vN/+fbzf/nm83/55vN/+ebzf/nW43/51uNv+dbjb/nG42/5xuNv+bbTb/m202/5ttNv+ZazP/pn9Q/9nSxv/u8vP/3fD1/4vj+/9Q2v//Rtj//0fY//9H1///Rtf//0bX//9G1///Rdb//0XW//9F1v//R9b//1DX/v9i2v3/keL6/8vs9v/p8fT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fP1/+vv8P/n5uL/4tvR/9nKtv/PuJn/yq2I/8epgf/HqID/x6iA/8eogP/Gp4D/xqeA/8angP/Gp4D/xaeA/8WngP/Fp4D/xaeA/8WmgP/FpoD/xKaA/8SmgP/EpoD/xKaA/8SmgP/DpoD/w6aA/8OlgP/CpYD/wqWA/8KlgP/CpYD/wqWA/8GlgP/BpYD/waWA/8GkgP/ApID/wKSA/8Clgf+8nXb/qX1H/6JxN/+icTf/onE3/6FxN/+hcTf/oXA3/6BwN/+gcDf/oHA3/59vN/+fbzf/nm83/55vN/+ebzf/nW43/51uNv+dbjb/nG42/5xuNv+bbTb/m202/5prM/+mgFD/2dLG/+7y8//f8PX/luX6/2Ld/f9Z2/7/Wtv+/1rb/v9Z2v7/Wdr+/1na/v9Z2v7/Wdr+/17a/f9v3fz/k+P6/73q9//c7/X/7PL0/+7y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL1/+7z9v/t8/X/6+/v/+nq5//n5+P/5+bh/+fm4f/n5uH/5+bh/+fm4f/n5uH/5ubh/+bm4f/m5uH/5ubh/+bl4f/m5eH/5uXh/+bl4f/m5eH/5uXh/+bl4f/m5eH/5uXh/+bl4f/m5eH/5uXh/+bl4f/m5eH/5uXh/+bl4f/m5eH/5uXh/+bl4f/m5eH/5uXh/+bl4f/l5eH/5+fj/9zUyP+yjF3/o3I3/6JxN/+icTf/onE3/6FxN/+hcTf/oXA3/6BwN/+gcDf/oHA3/59vN/+fbzf/nm83/55vN/+ebzf/nW43/51uNv+dbjb/nG42/5xuNv+bbTb/mmsz/6aAUP/Z0sb/7fLz/+fx9f/I7Pb/sej4/67o+f+u6Pj/rej4/63o+P+t6Pj/rej4/63o+P+x6Pj/wOr3/9Lu9v/f8PX/6vH0/+7y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vX/7fP2/+7z9v/u8/b/7vP2/+7z9v/u8/b/7vP2/+7z9v/u8/b/7vP2/+7z9v/u8/b/7vT3/+/1+f/v9fn/7/X5/+/1+f/v9fn/7/X5/+/1+f/v9fn/7/X5/+/1+f/v9fn/7/X5/+/1+f/v9fn/7/X5/+/1+f/v9fn/7/X5/+/1+f/v9fn/7/X5/+/1+f/w+Pz/5OPd/7WQY/+jcjf/o3I3/6JxN/+icTf/onE3/6FxN/+hcTf/oXA3/6BwN/+gcDf/n3A3/59vN/+fbzf/nm83/55vN/+ebzf/nW43/51uNv+dbjb/nG42/5xuNv+aazP/poBQ/9nSxv/t8fP/7fL0/+nx9P/n8fT/5vH1/+bx9f/m8fX/5vH1/+bx9f/m8fX/5vH0/+fx9P/r8vT/7/L0/+/y9P/u8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/r7e3/5eLc/+Pf1//j4Nj/4+DY/+Pg2P/j4Nj/49/Y/+Pf2P/j39j/49/Y/+Pf2P/j39j/49/Y/+Pf2P/j39j/49/Y/+Pf2P/j39j/4t/Y/+Lf2P/i39j/4t/Y/+Ph2v/Zz8H/sopb/6RyN/+jcjf/o3I3/6JxN/+icTf/onE3/6FxN/+hcTf/oXA3/6BwN/+gcDf/n3A3/59vN/+fbzf/nm83/55vN/+ebzf/nW43/51uNv+dbjb/nG42/5tsM/+ngFD/2dLG/+3x8//t8vT/7vL0/+7y9P/u8vT/7vL0/+7y9P/u8vT/7vL0/+7y9P/u8vT/7vL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fHz/+Dazv/BnnL/uY9b/7qQXv+6kF7/upBe/7mQXv+5kF7/uI9e/7iPXv+4j17/uI9e/7ePXv+3j17/t49e/7eOXv+2jl7/to5e/7aOXv+1jl7/tY5e/7WNXv+0jV7/tI1e/7KJWP+oeED/pHI3/6RyN/+jcjf/o3I3/6JxN/+icTf/onE3/6FxN/+hcTf/oXA3/6BwN/+gcDf/n3A3/59vN/+fbzf/nm83/55vN/+ebzf/nW43/51uNv+dbjb/m2wz/6eAUP/Z0sb/7fHz/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8fP/3dPE/7aIUP+sdTP/rXc2/6x2Nv+sdjb/rHY2/6t2Nv+rdjb/qnU2/6p1Nv+pdTb/qXU2/6l0Nv+odDb/qHQ2/6h0Nv+nczb/p3M2/6ZzNv+mczb/pnM2/6VyNv+lcjb/pXI2/6VzN/+kcjf/pHI3/6RyN/+jcjf/o3I3/6JxN/+icTf/onE3/6FxN/+hcTf/oXA3/6BwN/+gcDf/n3A3/59vN/+fbzf/nm83/55vN/+ebzf/nW43/51uNv+cbDP/qIBQ/9nSxv/t8fP/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3x8//d08T/tolR/6x2NP+teDf/rXc3/6x3N/+sdzf/rHc3/6t3N/+rdjf/qnY3/6p2N/+qdjf/qXU3/6l1N/+pdTf/qHU3/6h1N/+odDf/p3Q3/6d0N/+mdDf/pnQ3/6ZzN/+lczf/pXM3/6VzN/+kcjf/pHI3/6RyN/+jcjf/o3I3/6JxN/+icTf/onE3/6FxN/+hcTf/oXA3/6BwN/+gcDf/n3A3/59vN/+fbzf/nm83/55vN/+ebzf/nW43/5xsM/+ogFD/2dLG/+3x8//t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fHz/93TxP+3iVH/rXY0/614N/+teDf/rXc3/6x3N/+sdzf/rHc3/6t3N/+rdjf/qnY3/6p2N/+qdjf/qXU3/6l1N/+pdTf/qHU3/6h1N/+odDf/p3Q3/6d0N/+mdDf/pnQ3/6ZzN/+lczf/pXM3/6VzN/+kcjf/pHI3/6RyN/+jcjf/o3I3/6JxN/+icTf/onE3/6FxN/+hcTf/oXA3/6BwN/+gcDf/n3A3/59vN/+fbzf/nm83/55vN/+ebzf/nGw0/6iAUf/Z0sb/7fHz/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/s8fL/3dPD/7eJUf+tdjT/rng3/614N/+teDf/rXc3/6x3N/+sdzf/q3Y2/6p1Nf+qdDT/qnU1/6p1Nv+qdjf/qXU3/6l1N/+pdTf/qHU3/6h1N/+odDf/p3Q3/6d0N/+mdDf/pnQ3/6ZzN/+lczf/pXM3/6VzN/+kcjf/pHI3/6RyN/+jcjf/o3I3/6JxN/+icTf/onE3/6FxN/+hcTf/oXA3/6BwN/+gcDf/n3A3/59vN/+fbzf/nm83/55vN/+dbTT/qIFR/9nSxv/t8fP/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+zx8v/d0sP/t4lR/652NP+ueDf/rng3/614N/+teDf/rXc2/6x4OP+wf0P/s4ZN/7SHUP+xgkn/rXo9/6p2N/+qdjf/qXU3/6l1N/+pdTf/qHU3/6h1N/+odDf/p3Q3/6d0N/+mdDf/pnQ3/6ZzN/+lczf/pXM3/6VzN/+kcjf/pHI3/6RyN/+jcjf/o3I3/6JxN/+icTf/onE3/6FxN/+hcTf/oXA3/6BwN/+gcDf/n3A3/59vN/+fbzf/nm83/51tNP+pgVH/2dLG/+3x8//t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7PHy/93Sw/+3iVH/rnc0/694N/+ueDf/rng3/614N/+uejv/uI1Z/8qwj//Yyrf/29C//9PBqP/Bn3X/sIFH/6p2OP+qdjf/qXU3/6l1N/+pdTf/qHU3/6h1N/+odDf/p3Q3/6d0N/+mdDf/pnQ3/6ZzN/+lczf/pXM3/6VzN/+kcjf/pHI3/6RyN/+jcjf/o3I3/6JxN/+icTf/onE3/6FxN/+hcTf/oXA3/6BwN/+gcDf/n3A3/59vN/+fbzf/nW00/6mBUf/Z0sb/7fHz/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/s8fL/3dLD/7iJUf+udzX/r3k3/694N/+ueDf/rnk5/7qQXP/XyLP/5+fj/+vv8P/s8PH/6u3s/+Ld1P/IrYv/r39G/6p1Nv+qdjf/qXU3/6l1N/+pdTf/qHU3/6h1N/+odDf/p3Q3/6d0N/+mdDf/pnQ3/6ZzN/+lczf/pXM3/6VzN/+kcjf/pHI3/6NyN/+jcjf/o3I3/6JxN/+icTf/onE3/6FxN/+hcTf/oXA3/6BwN/+gcDf/n3A3/59vN/+ebTT/qYFR/9rSxv/t8fP/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/u8/X/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+zx8v/d0sP/uIlR/693Nf+veTj/r3k3/653Nf+zgkb/0Lmc/+jp5v/u8/b/7fP1/+3y9f/t8/X/7fP1/+Da0P+8l2n/q3g6/6p2Nv+qdjf/qXU3/6l1N/+pdTf/qHU3/6h1N/+odDf/p3Q3/6d0N/+mdDf/pnQ3/6ZzN/+lczf/pXM3/6VzN/+kcjf/pHI3/6NyN/+jcjf/o3I3/6JxN/+icTf/onE3/6FxN/+hcTf/oXA3/6BwN/+gcDf/n3A3/55tNP+qgVH/2dLG/+3x8//t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7vP1/+/09f/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7PHy/93Swv+4ilH/r3g1/7B5OP+veTj/rnc1/7iLVP/f18r/7fL0/+3y9P/t8vT/7fL0/+3y9P/u8/b/6Onm/8uykf+vfkP/qnU2/6p2N/+qdjf/qXU3/6l1N/+pdTf/qHU3/6h1N/+odDf/p3Q3/6d0N/+mdDf/pnQ3/6ZzN/+lczf/pXM3/6VzN/+kcjf/pHI3/6NyN/+jcjf/o3I3/6JxN/+icTf/onE3/6FxN/+hcTf/oXA3/6BwN/+gcDf/nm40/6qBUf/Z0sX/7fHz/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/v9PX/8PX2/+7z9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/s8fL/3dLC/7mKUf+veDX/sHk4/7B5OP+veDb/uo9a/+Pf1v/u9Pf/7fL0/+3y9P/t8vT/7fL0/+3z9f/p6+n/0Lyh/7GCSP+qdTX/qnY3/6p2N/+qdjf/qXU3/6l1N/+pdTf/qHU3/6h1N/+odDf/p3Q3/6d0N/+mdDf/pnQ3/6ZzN/+lczf/pXM3/6VzN/+kcjf/pHI3/6NyN/+jcjf/o3I3/6JxN/+icTf/onE3/6FxN/+hcTf/oXA3/6BwN/+fbjT/q4NT/9zVyv/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7vP0//D19v/z9vf/7vP1/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+zx8v/d0sL/uYpR/7B4Nf+wejj/sHk4/693Nv+4i1T/3tbI/+3y9P/t8vT/7fL0/+3y9P/t8vT/7vP2/+jp5v/LsZD/sH9D/6t2Nv+rdjf/qnY3/6p2N/+qdjf/qXU3/6l1N/+pdTf/qHU3/6h1N/+odDf/p3Q3/6d0N/+mdDf/pnQ3/6ZzN/+lczf/pXM3/6VzN/+kcjf/pHI3/6NyN/+jcjf/o3I3/6JxN/+icTf/onE3/6FxN/+hcTf/oXA3/6BvNv+wil7/4NzV/+7z9v/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/u8/X/8/b3//X4+f/v8/X/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/9/WyP+6jFT/sHg1/7F6OP+wejj/r3g2/7WDR//PuJr/6Ojl/+7z9v/t8/X/7fL1/+7z9v/t8vX/4NnO/72WaP+teTr/rHc2/6t3N/+rdjf/qnY3/6p2N/+qdjf/qXU3/6l1N/+pdTf/qHU3/6h1N/+odDf/p3Q3/6d0N/+mdDf/pnQ3/6ZzN/+lczf/pXM3/6VzN/+kcjf/pHI3/6NyN/+jcjf/o3I3/6JxN/+icTf/onE3/6FxN/+hcDb/pHU+/72hfv/l5OD/7vP2/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+/z9f/1+Pn/+Pr7/+/09f/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/u8/b/5ODX/8CZaP+yezr/sXo4/7F6OP+wejj/sHo6/7uPW//Wxq//5+Xh/+vu7v/s7/D/6uzr/+Hb0f/IrIf/sYBE/6x2Nv+sdzf/rHc3/6t3N/+rdjf/qnY3/6p2N/+qdjf/qXU3/6l1N/+pdTf/qHU3/6h1N/+odDf/p3Q3/6d0N/+mdDf/pnQ3/6ZzN/+lczf/pXM3/6VzN/+kcjf/pHI3/6NyN/+jcjf/o3I3/6JxN/+icTf/onE3/6FwNv+qgE7/0cGt/+rt7P/t8/X/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7/T1//j6+//7/P3/8PT2/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+7z9v/p6eb/z7aW/7aDRv+xeTb/sXo4/7F6OP+weTj/sXs7/7qNV//KrYr/18aw/9nMuP/SvaL/wZ1x/7KBRv+teDj/rXc3/6x3N/+sdzf/rHc3/6t3N/+rdjf/qnY3/6p2N/+qdjf/qXU3/6l1N/+pdTf/qHU3/6h1N/+odDf/p3Q3/6d0N/+mdDf/pnQ3/6ZzN/+lczf/pXM3/6VzN/+kcjf/pHI3/6NyN/+jcjf/o3I3/6JxN/+icDb/pHU9/7ydeP/j39n/7fL1/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/w9Pb/+/z9//39/v/z9vj/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+zx8v/g2Mv/wZlo/7J8Ov+xeTf/sXo4/7F6OP+weTf/sHo4/7OAQ/+2hkz/todO/7SDSP+wezz/rng3/614N/+teDf/rXc3/6x3N/+sdzf/rHc3/6t3N/+rdjf/qnY3/6p2N/+qdjf/qXU3/6l1N/+pdTf/qHU3/6h1N/+odDf/p3Q3/6d0N/+mdDf/pnQ3/6ZzN/+lczf/pXM3/6VzN/+kcjf/pHI3/6NyN/+jcjf/onE2/6NyOf+zjWH/2M2//+zv8P/t8/X/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0//P2+P/9/f7//v7+//f5+v/u8/X/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fP1/+rt7P/YyLL/vpJd/7J7Ov+xeTf/sXo4/7F6OP+wejj/sHg3/694Nv+udzX/rng1/654Nv+ueDf/rng3/614N/+teDf/rXc3/6x3N/+sdzf/rHc3/6t3N/+rdjf/qnY3/6p2N/+qdjf/qXU3/6l1N/+pdTf/qHU3/6h1N/+odDf/p3Q3/6d0N/+mdDf/pnQ3/6ZzN/+lczf/pXM3/6VzN/+kcjf/pHI3/6NxNv+kdDr/soxe/9LEsf/p6+r/7fL1/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/u8/X/9/n6//7+/v//////+/z8//D09v/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+nq6P/Yx7D/wJZj/7R/P/+xeTf/sXk3/7F6OP+wejj/sHk4/7B5OP+veTj/r3k3/654N/+ueDf/rng3/614N/+teDf/rXc3/6x3N/+sdzf/rHc3/6t3N/+rdjf/qnY3/6p2N/+qdjf/qXU3/6l1N/+pdTf/qHU3/6h1N/+ndDf/p3Q3/6d0N/+mdDf/pnQ3/6ZzN/+lczf/pXM3/6RyNv+kcjb/qHpD/7eVa//VyLb/6erp/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0//D09v/7/Pz////////////9/v7/9Pf4/+7z9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+rs6v/d0sH/yKd+/7mJUP+zfT3/sHk2/7B5Nv+weTj/sHk4/7B5OP+veTj/r3k3/654N/+ueDf/rng3/614N/+teDf/rXc3/6x3N/+sdzf/rHc3/6t3N/+rdjf/qnY3/6p2N/+qdjf/qXU3/6l1N/+pdTf/qHU3/6h1N/+ndDf/p3Q3/6d0N/+mdDf/pnM2/6VyNf+lczb/qHk//7CHVf/Eqor/3NbK/+rt7f/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/u8/T/9Pf4//3+/v////////////7////5+/v/7/T1/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL1/+zw8f/l49z/18au/8Wid/+6i1P/tYJE/7F8Ov+weDb/r3g2/694N/+veDf/r3k3/654N/+ueDf/rng3/614N/+teDf/rXc3/6x3N/+sdzf/rHc3/6t3N/+rdjf/qnY3/6p2N/+qdjf/qXU3/6l1N/+pdTf/qHU3/6d0Nv+nczX/pnM1/6ZzNf+odzz/rH5G/7OKWP/CpoP/18u6/+Xl4f/s8fL/7fP1/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+/09f/5+/v//v////////////////////z9/v/z9vj/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL1/+3z9f/r7u3/5ePc/9vPvf/OtJP/wZtr/7qMVP+2hUr/tIBD/7J9Pv+veTj/rnc1/612NP+tdjT/rXc1/613Nf+sdzX/rHY1/6x2Nf+rdjX/q3Y1/6t2Nf+qdTX/qnU1/6l1Nf+pdDT/qHM0/6hzNP+pdjj/q3k+/6x9RP+vgUv/s4pX/72cc//Lt5v/3NPG/+bl4P/r7+//7fP1/+3y9f/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/8/b3//39/f///////////////////////v////j7+//v9Pb/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3z9f/t8/b/7PDw/+np5v/l4dr/3tXG/9XCqf/MsI7/xaN4/76WZP+5jVb/t4lR/7aHTv+1hUv/tINI/7ODSP+ygkb/soFG/7GBRv+xgUb/sYFG/7GCSP+xgkj/sYNL/7KFTv+zh1H/tIpW/7qUZf/Bonv/yrKT/9PErv/e1sv/5OLd/+jq6P/s8fL/7vP2/+3z9f/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+/09v/4+/v//v///////////////////////////////f3+//T3+f/u8/X/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vX/7vP2/+7z9v/t8fP/6u3r/+jo5P/m5N//5OHZ/+Hb0P/d1MT/2cy5/9bFrf/Tv6T/0b2h/8+5nP/OuJr/zria/864mv/PuZ3/0Lyh/9G+pP/Uw63/2Mu5/93Uxv/h29L/4+DZ/+bl4P/o6eb/6u3t/+3y9P/u8/b/7vP1/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/u8/X/9Pj5//3+/v//////////////////////////////////////+vz8//H19//t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8/X/7vP2/+709v/u9Pf/7vP2/+3x8//s7/D/6+3t/+rs6v/q6+n/6ero/+nq6P/p6uj/6ero/+nq6P/p6+n/6uzq/+rt7f/r7/D/7fHz/+7z9v/u9Pf/7vT2/+7z9v/t8/X/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0//H19//6/Pz////////////////////////////////////////////+/v7/+Pr6/+/09f/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9f/t8/X/7fP1/+3z9f/t8/X/7vP1/+7z9f/u8/X/7fP1/+3z9f/t8/X/7fP1/+3y9f/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/v9PX/+Pn6//7+/v/////////////////////////////////////////////////9/v7/9fj5/+7z9f/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7/P1//b4+f/9/f7////////////////////////////////////////////////////////////8/f3/9Pf4/+7z9f/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+7z9f/09/j//P39///////////////////////////////////////////////////////////////////////7/Pz/8/b3/+7z9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/u8vT/8/b3//v8/P/////////////////////////////////////////////////////////////////////////////////7/Pz/8/b3/+7z9f/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7vP1//P29//6/Pz////////////////////////////////////////////////////////////////////////////////////////////7/Pz/9Pf4/+7z9f/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+7z9f/09/j/+/z8///////////////////////////////////////////////////////////////////////////////////////////////////////8/f3/9fj5/+/09f/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/v9PX/9fj5//z9/f/////////////////////////////////////////////////////////////////////////////////////////////////////////////////9/v7/+Pr6//H19//u8/X/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/u8/X/8fX3//j6+v/9/v7////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////+/v7/+/z8//X4+f/w9Pb/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7/T1//T3+P/6/Pz//v7+/////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////f7+//n7+//z9vf/7/T1/+7z9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/u8/T/7/T1//P2+P/4+/v//f3+/////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////v////z9/f/5+/v/9Pf4//D09v/u8/X/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/u8/X/8PT2//T3+P/5+/v//P3+//7///////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////7////9/v7/+/z8//f5+v/z9vj/8PT2/+/09f/v8/X/7vP1/+7z9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+7z9P/u8/X/7/P1/+/09f/w9Pb/8/b4//f5+v/7/Pz//f7+//7//////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////v7+//39/v/7/P3/+Pr7//X4+f/z9vf/8PX2/+/09f/u8/X/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+3y9P/t8vT/7fL0/+7z9f/v9PX/8PX2//P29//1+Pn/+Pr7//v8/f/9/f7//v7+////////////////////////////////////////////////////////////////////////////////////////////////////////////AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA=
'''

try:   ##尝试使用base64编码的图标
        import base64
        import io
        # 解码base64数据
        icon_data = base64.b64decode(ICON_BASE64.strip())
        # 使用io.BytesIO创建一个内存文件对象
        icon_file = io.BytesIO(icon_data)
        # 使用PIL打开图标
        image = Image.open(icon_file)
except Exception as e:
    log_print(f'Error loading icon: {str(e)}, using default icon')
    image = Image.new('RGB', (64, 64), color=(255, 255, 255))   #创建一个简单的白色方块作为默认图标

menu = (item('Show/Hide console window', toggle_console), item('Exit program', exit_program))   #创建右键菜单

icon = pystray.Icon("office_backup_utilities", image, "Office Backup Utilities", menu)   #创建托盘图标对象
'''icon.on_left_click = on_clicked   #绑定左键单击事件处理函数（无法生效）'''

# 根据配置决定是否启动托盘图标
if not config.get('hide_tray_icon'):
    icon_task = threading.Thread(target=icon.run)   #创建托盘图标线程
    icon_task.daemon = True   #设置为守护线程（随主程序终止而自动结束）
    icon_task.start()   #启动托盘图标线程

# 全局异常处理函数
def global_exception_handler(exctype, value, tb):   #处理全局未捕获异常
    # 构建完整的错误信息，包括 Traceback (most recent call last):
    error_msg = "".join(traceback.format_exception(exctype, value, tb))
    
    # 打印到控制台
    print(f"[ERROR] {error_msg}")
    
    # 写入日志文件
    if config.get('save_log'):
        log_msg = time.strftime('[%H:%M:%S]') + ' [ERROR] ' + error_msg
        log_file.write(log_msg + '\n')
        log_file.flush()

# 设置全局异常处理器
sys.excepthook = global_exception_handler





print('Program initialization completed, entering main loop\n')
if config.get('save_log'):
    log_file.write('Program initialization completed, entering main loop\n\n')
    log_file.flush()   #刷新文件缓冲区，确保日志消息立即写入文件

while True:   #主线程无限循环，防止程序退出
    if config.get('ppt_backup_enable'):   #检查PPT备份功能是否启用
        save_open_ppt_files(ppt_save_folder)   #启动线程
    if config.get('word_backup_enable'):   #检查Word备份功能是否启用
        save_open_word_files(word_save_folder)   #启动线程
    if config.get('wps_backup_enable'):   #检查WPS备份功能是否启用
        save_open_WPS_files(ppt_save_folder)   #启动线程
    if config.get('accurate_backup_enable'):  # 检查精确备份功能是否启用
        accurate_backup()  # 启动线程
    time.sleep(sleeptime)   #等待指定时间后继续轮询
