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
    #指定备份路径，r表示取原始字符串，需要更改请更改引号内部分
    "ppt_backup_path": "C:\\Officebackup\\pptbackup",   #PPT、WPS备份路径
    "word_backup_path": "C:\\Officebackup\\wordbackup",   #Word备份路径
    #指定间隔时间，单位为秒
    "interval": 60,   #指定所有操作的轮询时间间隔，单位为秒（默认60秒）
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
    "show_console_window_at_startup": False,   #程序启动时显示控制台窗口，True为显示，False为隐藏（默认）
    "save_log": True,   #是否保存日志到OBUlatest.log文件，True为保存（默认），False为不保存
    "archive_previous_log": True,   #是否在程序启动时归档之前的日志，True为归档（默认），False为直接覆盖
    #超时和重试设置
    "backup_timeout": 600,   #备份操作超时时间，单位为秒（默认10分钟）
    "upload_retry_wait": 30,   #上传重试等待时间，单位为秒（默认30秒）
    "upload_max_retries": ""   #上传最大重试次数，默认空表示无限次重试
}
try:   #读取配置文件
    with open('OBU6.0.json', 'r', encoding='utf-8') as f:   #尝试读取配置文件（只读）
        config = json.load(f)
    config_changed = False
    for key, value in default_config.items():   #如果现有配置文件有缺漏，根据默认配置项自动补全
        if key not in config:
            config[key] = value
            config_changed = True
    if config_changed:   #如果配置文件有新增项，写回配置文件
        with open('OBU6.0.json', 'w', encoding='utf-8') as f:
            json.dump(config, f, indent=4, ensure_ascii=False)
except (FileNotFoundError, json.JSONDecodeError):   #若配置文件不存在或无法解析
    config = default_config   #使用默认配置
    with open('OBU6.0.json', 'w', encoding='utf-8') as f:   #在当前目录下根据默认配置文件创建（写入）
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
    log_print("alist3 not found, force disabled upload function")
    config['upload_to_openlist_enable'] = False   #强制禁用上传功能

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
                        # 说明路径无效，禁用上传功能
                        log_print('Target folder invalid: ' + openlist_target_folder + ', error: ' + str(e), source='openlist')
                        log_print('Disabling upload function, please check target folder path in config file', source='openlist')
                        config['upload_to_openlist_enable'] = False
                        with open('OBU6.0.json', 'w', encoding='utf-8') as f:
                            json.dump(config, f, indent=4, ensure_ascii=False)
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
                        log_print('Upload failed with error: ' + str(e), source='openlist')
                        # 检查是否是 504 超时错误
                        if '504' in str(e) or 'timeout' in str(e).lower():
                            log_print('Server timeout error, will retry in next upload', source='openlist')
                        else:
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
                log_print('Upload to OpenList failed: ' + str(e), source='openlist')
                log_print('Traceback: ' + traceback.format_exc(), source='openlist')
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


if not openlist_url or not openlist_username or not openlist_password:   #检查OpenList配置是否完整
    log_print('OpenList URL, username or password is empty, force disabled upload function, please provide valid credentials in the configuration file')
    config['upload_to_openlist_enable'] = False   #强制禁用上传功能
else:
    # 启动上传线程
    upload_to_openlist()

if config.get('accurate_backup_enable'):  # 检查精确备份功能是否启用
    source_path = config.get('accurate_backup_source_path')   #获取源文件夹路径
    target_path = config.get('accurate_backup_target_path')   #获取目标文件夹路径
    if not source_path and target_path:   #如果精确备份功能开启但源路径为空或目标路径为空，则强制禁用精确备份功能
        log_print("Source path or target path for accurate backup is empty, force disabled accurate backup function, please provide valid paths in the configuration file")
        with open('OBU6.0.json', 'w', encoding='utf-8') as f:   #将禁用精确备份功能写入配置文件
                config['accurate_backup_enable'] = False   #强制禁用精确备份功能
                json.dump(config, f, indent=4, ensure_ascii=False)   #写入更新后的配置文件



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
                        # 构建完整的命令
                        script_path = os.path.abspath(__file__)
                        command = [sys.executable, script_path]
                        # 直接启动，不打印日志
                        subprocess.Popen(command)
                        time.sleep(1)  # 给新进程启动时间
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
                
                if original_md5 and backup_md5 and original_md5 == backup_md5:
                    # MD5值相同，跳过备份
                    log_print(ppt_name + ' has already existed in ' + ppt_save_folder + ', skipped backup (MD5 match)')   #打印跳过信息
                    continue   #跳过此次备份
                else:
                    # MD5值不同，需要备份
                    log_print(ppt_name + ' has changed, backup will begin soon (MD5 mismatch)')
            
            Existed_in_this_session[ppt_name] = True   #标记该文件在本次会话中出现过
            log_print('Start to backup ' + ppt_name + ' to ' + ppt_save_folder)   #打印备份开始信息
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
            log_print('Start to backup ' + ppt_name + ' to ' + ppt_save_folder)   #打印备份开始信息
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
                
                if original_md5 and backup_md5 and original_md5 == backup_md5:
                    # MD5值相同，跳过备份
                    log_print(doc_name + ' has already existed in ' + word_save_folder + ', skipped backup (MD5 match)')   #打印跳过信息
                    continue   #跳过此次备份
                else:
                    # MD5值不同，需要备份
                    log_print(doc_name + ' has changed, backup will begin soon (MD5 mismatch)')

            Existed_in_this_session[doc_name] = True   #标记该文件在本次会话中出现过
            log_print('Start to backup ' + doc_name + ' to ' + word_save_folder)   #打印备份开始信息
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
            log_print('Start to backup ' + doc_name + ' to ' + word_save_folder)   #打印备份开始信息
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
                
                if original_md5 and backup_md5 and original_md5 == backup_md5:
                    # MD5值相同，跳过备份
                    log_print(WPS_ppt_name + ' has already existed in ' + ppt_save_folder + ', skipped backup (MD5 match)')   #打印带时间戳和运行次数的跳过信息
                    continue   #跳过此次备份
                else:
                    # MD5值不同，需要备份
                    log_print(WPS_ppt_name + ' has changed, backup will begin soon (MD5 mismatch)')

            Existed_in_this_session[WPS_ppt_name] = True   #标记该文件在本次会话中出现过
            log_print('Start to backup ' + WPS_ppt_name + ' to ' + ppt_save_folder)   #打印备份开始信息
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
            log_print('Start to backup ' + WPS_ppt_name + ' to ' + ppt_save_folder)   #打印备份开始信息
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

# SVG图标代码 - 请将PythonLight.svg的内容粘贴到这里
SVG_ICON_CODE = '''
<?xml version="1.0" encoding="UTF-8"?><svg xmlns="http://www.w3.org/2000/svg" xmlns:xlink="http://www.w3.org/1999/xlink" width="128" height="128" viewBox="0 0 128 128"><title>PythonLight 128x128</title><image width="128" height="128" xlink:href="data:image/png;base64,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"/></svg>
'''

try:   #尝试加载图标文件
    # 尝试使用svglib加载嵌入的SVG图标
    try:
        from svglib.svglib import svg2rlg
        from reportlab.graphics import renderPM
        import io
        # 使用io.StringIO创建一个内存文件对象
        svg_file = io.StringIO(SVG_ICON_CODE)
        drawing = svg2rlg(svg_file)
        # 尝试使用不同的后端渲染
        try:
            image = renderPM.drawToPIL(drawing)
        except Exception as render_error:
            # 如果渲染失败，使用默认图标
            log_print(f'Failed to render SVG: {str(render_error)}, using default icon')
            image = Image.new('RGB', (64, 64), color=(255, 255, 255))
    except ImportError as import_error:
        # 如果svglib或reportlab未安装，使用默认图标
        log_print(f'Import error: {str(import_error)}, using default icon')
        image = Image.new('RGB', (64, 64), color=(255, 255, 255))
    except Exception as e:
        # 如果SVG加载失败，使用默认图标
        log_print(f'Failed to load SVG icon: {str(e)}, using default icon')
        image = Image.new('RGB', (64, 64), color=(255, 255, 255))
except Exception as e:
    log_print(f'Error loading icon: {str(e)}, using default icon')
    image = Image.new('RGB', (64, 64), color=(255, 255, 255))   #创建一个简单的白色方块作为默认图标

menu = (item('Show/Hide console window', toggle_console), item('Exit program', exit_program))   #创建右键菜单

icon = pystray.Icon("office_backup_utilities", image, "Office Backup Utilities", menu)   #创建托盘图标对象
'''icon.on_left_click = on_clicked   #绑定左键单击事件处理函数（无法生效）'''





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
