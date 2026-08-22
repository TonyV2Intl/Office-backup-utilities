import os   #导入os模块，用于处理文件和目录路径操作
import shutil   #导入shutil模块，用于复制文件并保留元数据
import time   #导入time模块，用于时间相关操作
import win32com.client as win32   #导入win32com.client库，用于通过COM接口与Microsoft Office应用程序交互
import datetime   #导入datetime库，用于计算备份所用时间
from collections import defaultdict  #导入collections库的defaultdict方法，用于跟踪单个文件的跳过次数

import threading  #导入threading库，用于多线程操作
import json  #导入json库，用于处理配置文件的读写
import sys  #导入sys模块，用于处理系统异常
import traceback  #导入traceback模块，用于获取详细的异常信息

import pystray   #导入pystray库，用于创建系统托盘图标
from pystray import MenuItem as item   #从pystray库中导入MenuItem类，用于创建托盘菜单项
from PIL import Image   #导入PIL库的Image模块，用于处理图标图像
import ctypes   #导入ctypes库，用于调用Windows API函数





#设定默认配置文件
default_config = {
    #指定备份路径，r表示取原始字符串，需要更改请更改引号内部分
    "ppt_backup_path": "C:\\Officebackup\\pptbackup",   #PPT、WPS备份路径
    "word_backup_path": "C:\\Officebackup\\wordbackup",   #Word备份路径
    #指定间隔时间，单位为秒
    "interval": 60,   #指定所有操作的轮询时间间隔，单位为秒（默认60秒）
    "max_skipping_time": 15,   #指定连续跳过次数（默认15次）
    #功能开启或禁用
    "ppt_backup_enable": True,   #PPT备份功能
    "word_backup_enable": True,   #Word备份功能
    "wps_backup_enable": True,   #WPS备份功能
    "upload_to_123pan_enable": True,   #上传到123云盘功能
    #123云盘参数
    "client_id": "",   #123云盘API用户ID
    "client_secret": "",   #123云盘API用户密钥
    "access_token": "",   #123云盘访问令牌，程序会自动获取并写入
    "folder_id": 0,   #目标文件夹ID，可以从浏览器地址栏获取，根目录用0表示
    #文件夹精确备份功能
    "accurate_backup_enable": False,
    "accurate_backup_source_path": "",
    "accurate_backup_target_path": "",
    #托盘图标、控制台行为与日志保存设置
    #"tray_left_click_behavior": "open_console",   #托盘图标左键点击行为，选项有"open_console"（打开控制台）和"exit_program"（退出程序）（无法生效）
    "show_console_window_at_startup": False,   #程序启动时显示控制台窗口，True为显示，False为隐藏（默认）
    "save_log": True   #是否保存日志到latest.log文件，True为保存（默认），False为不保存
}
startup_warnings = []   #初始化启动警告列表，日志系统准备好后再输出
config_file_path = 'OfficebackupSingleConfig.json'   #配置文件路径
try:   #读取配置文件
    with open('OfficebackupSingleConfig.json', 'r', encoding='utf-8') as f:   #尝试读取配置文件（只读）
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
        with open(config_file_path, 'w', encoding='utf-8') as f:
            json.dump(config, f, indent=4, ensure_ascii=False)
    except OSError as write_error:
        startup_warnings.append('Failed to create replacement configuration file ' + config_file_path + ': ' + str(write_error))



log_file = None   #初始化日志文件句柄，日志打开失败时保持为空
file_logging_enabled = False   #标记本次会话是否启用文件日志
log_write_error_reported = False   #标记是否已经报告过日志写入失败
if config.get('save_log'):   #检查是否启用日志保存功能
    try:
        if os.path.exists('latest.log'):   #如果日志文件存在，则删除旧日志文件
            os.remove('latest.log')
        log_file = open('latest.log', 'a', encoding='utf-8')   #以追加模式打开日志文件
        file_logging_enabled = True   #标记文件日志已启用
    except (OSError, IOError) as e:
        startup_warnings.append('Failed to initialize file logging: ' + type(e).__name__ + ': ' + str(e) + '; continuing with console-only logging')
        config['save_log'] = False   #仅在内存中禁用文件日志
def log_print(msg):   #定义日志打印函数
    global runid, file_logging_enabled, log_write_error_reported    #声明全局变量
    runid+=1   #运行计数器累加
    log_msg= time.strftime('[%H:%M:%S-#') + str(runid) + '] ' + msg   # 打印带时间戳的日志消息到控制台
    print(log_msg)   # 打印日志消息到控制台
    if config.get('save_log') and file_logging_enabled and log_file is not None and not log_file.closed:   #如果启用日志保存功能，则将日志消息写入日志文件
        try:
            log_file.write(log_msg + '\n')   # 将日志消息写入日志文件
            log_file.flush()   #刷新文件缓冲区，确保日志消息立即写入文件
        except (OSError, IOError, ValueError) as e:
            file_logging_enabled = False
            config['save_log'] = False
            if not log_write_error_reported:
                log_write_error_reported = True
                print('[ERROR] File logging disabled after write failure: ' + type(e).__name__ + ': ' + str(e), file=sys.stderr)



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
max_skipping_time=config.get('max_skipping_time')   #连续跳过次数（默认为15次）
ppt_save_folder=config.get('ppt_backup_path')   #ppt备份路径
word_save_folder=config.get('word_backup_path')   #word备份路径
client_id = config.get('client_id', '')  # 123云盘用户ID
client_secret = config.get('client_secret', '')  # 123云盘用户密钥
folder_id = config.get('folder_id', 0)  # 123云盘目标文件夹ID
'''behavior = config.get('tray_left_click_behavior')  # 托盘图标左键点击行为（默认为打开控制台）（无法生效）'''



try:   #尝试导入pan123库
    from pan123.auth import get_access_token   #导入pan123.auth库，用于获取123云盘的access_token
    from pan123 import Pan123   #导入pan123库，用于与123云盘交互
except ImportError:
    log_print("pan123 not found, force disabled upload function")
    config['upload_to_123pan_enable'] = False   #强制禁用上传功能

if not client_id or not client_secret:   #检查client_id和client_secret是否为空，若为空则强制禁用上传功能
    log_print("Client ID or Client Secret is empty, force disabled upload function, please provide valid credentials in the configuration file")
    config['upload_to_123pan_enable'] = False   #强制禁用上传功能

def request_access_token():   #定义获取access_token函数
    global access_token, pan, token_aquired   #声明全局变量access_token和pan，以便在函数内修改其值
    try:   #尝试获取123云盘的access_token
        access_token = get_access_token(client_id, client_secret)
        pan = Pan123(access_token)
        log_print('Access_token of 123Pan acquired successfully')   #打印获取123云盘的access_token成功的信息

        token_aquired=True   #标记token获取成功

        config['access_token'] = access_token   #更新配置字典中的access_token
        try:
            with open(config_file_path, 'w', encoding='utf-8') as f:   #将access_token写入配置文件
                json.dump(config, f, indent=4, ensure_ascii=False)   #写入更新后的配置文件
            log_print('Access_token of 123Pan saved to json file successfully')   #打印保存access_token到配置文件成功的信息
        except OSError as save_error:
            log_exception('Failed to save newly acquired 123Pan access_token; continuing with token in memory', save_error)
    except Exception as e:
        token_aquired=False   #标记token获取失败
        if config.get('access_token'):   #如果配置文件中已有access_token，则尝试使用该token
            access_token = config.get('access_token')  #尝试从配置文件中读取之前保存的access_token
            pan = Pan123(access_token)
            log_exception('Failed to acquire access_token of 123Pan: temporarily use the token in json file, token request will continue after a while', e)
        else:
            access_token = ""  #将access_token定义为空字符串
            log_exception('Failed to acquire access_token of 123Pan: do not find token in json file either, token request will continue after a while', e)

if config['upload_to_123pan_enable'] == True:   #如果上传功能开启，则获取access_token
    request_access_token()

if config.get('accurate_backup_enable'):  # 检查精确备份功能是否启用
    source_path = config.get('accurate_backup_source_path')   #获取源文件夹路径
    target_path = config.get('accurate_backup_target_path')   #获取目标文件夹路径
    if not source_path and target_path:   #如果精确备份功能开启但源路径为空或目标路径为空，则强制禁用精确备份功能
        log_print("Source path or target path for accurate backup is empty, force disabled accurate backup function, please provide valid paths in the configuration file")
        config['accurate_backup_enable'] = False   #强制禁用精确备份功能
        try:
            with open(config_file_path, 'w', encoding='utf-8') as f:   #将禁用精确备份功能写入配置文件
                config['accurate_backup_enable'] = False   #强制禁用精确备份功能
                json.dump(config, f, indent=4, ensure_ascii=False)   #写入更新后的配置文件
        except OSError as e:
            log_exception('Failed to save accurate backup disable setting', e)





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
                if file_skip_count[ppt_name] < max_skipping_time and Existed_in_this_session[ppt_name] == True:  # 仅当同一文件连续跳过规定次数，且在本次会话中出现过时才允许重新备份
                    file_skip_count[ppt_name] += 1   #该文件的跳过计数器累加
                    if file_skip_count[ppt_name] == max_skipping_time:   # 如果跳过次数达到规定次数，打印提示信息
                        log_print(ppt_name + ' has already existed in ' + ppt_save_folder + ', skipped backup (skipped times: ' + str(file_skip_count[ppt_name]) + ', this file will be backed up again during the next request)')   #打印跳过信息
                    else:
                        log_print(ppt_name + ' has already existed in ' + ppt_save_folder + ', skipped backup (skipped times: ' + str(file_skip_count[ppt_name]) + ')')   #打印跳过信息
                    continue   #跳过此次备份
            
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
    except Exception as e:   #获取其他错误类型
            if type(e).__name__=='com_error':   #捕获无打开的PowerPoint实例而产生的的异常
                log_print('No ppt available now (PowerPoint application not detected)')   #打印异常信息
            else:   #打印出其他错误并继续轮询
                log_exception('Exception while backing up ppt, request continue', e)



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
                if file_skip_count[doc_name] < max_skipping_time and Existed_in_this_session[doc_name] == True:  # 仅当同一文件连续跳过规定次数，且在本次会话中出现过时才允许重新备份
                    file_skip_count[doc_name] += 1   #该文件的跳过计数器累加
                    if file_skip_count[doc_name] == max_skipping_time:   # 如果跳过次数达到规定次数，打印提示信息
                        log_print(doc_name + ' has already existed in ' + word_save_folder + ', skipped backup (skipped times: ' + str(file_skip_count[doc_name]) + ', this file will be backed up again during the next request)')   #打印跳过信息
                    else:
                        log_print(doc_name + ' has already existed in ' + word_save_folder + ', skipped backup (skipped times: ' + str(file_skip_count[doc_name]) + ')')   #打印跳过信息
                    continue   #跳过此次备份

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
    except Exception as e:   #获取其他错误类型
            if type(e).__name__=='com_error':   #捕获无打开的PowerPoint实例而产生的的异常
                log_print('No doc available now (Word application not detected)')   #打印带时间戳和运行次数的异常信息
            else:   #打印出其他错误并继续轮询
                log_exception('Exception while backing up doc, request continue', e)



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
                if file_skip_count[WPS_ppt_name] < max_skipping_time and Existed_in_this_session[WPS_ppt_name] == True:  # 仅当同一文件连续跳过规定次数，且在本次会话中出现过时才允许重新备份
                    file_skip_count[WPS_ppt_name] += 1   #该文件的跳过计数器累加
                    if file_skip_count[WPS_ppt_name] == max_skipping_time:   # 如果跳过次数达到规定次数，打印提示信息
                        log_print(WPS_ppt_name + ' has already existed in ' + ppt_save_folder + ', skipped backup (skipped times: ' + str(file_skip_count[WPS_ppt_name]) + ', this file will be backed up again during the next request)')   #打印带时间戳和运行次数的跳过信息
                    else:
                        log_print(WPS_ppt_name + ' has already existed in ' + ppt_save_folder + ', skipped backup (skipped times: ' + str(file_skip_count[WPS_ppt_name]) + ')')   #打印带时间戳和运行次数的跳过信息
                    continue   #跳过此次备份

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
    except Exception as e:   #获取其他错误类型
            if type(e).__name__=='com_error':   #捕获无打开的WPS实例而产生的的异常
                log_print('No ppt available now (WPS application not detected)')   #打印异常信息
            else:   #打印出其他错误并继续轮询
                log_exception('Exception while backing up WPS ppt, request continue', e)



def upload_to_123pan():   #定义上传函数
    global upload_queue, token_aquired, access_token  #声明全局上传队列变量
    if not token_aquired:   #如果token获取失败，则尝试重新获取
        request_access_token()
        if not token_aquired and not access_token:   #如果token获取仍然失败且配置文件中没有access_token，则等待一段时间后重试
            return   #跳过本次上传操作，继续下一轮循环
    for (upload_file, upload_source_path) in list(upload_queue):  #遍历队列快照，避免修改队列时跳过项目
        log_print('Start to upload ' + upload_file + ' to 123Pan')   #打印上传开始信息
        upload_start_time=datetime.datetime.now()   #记录上传操作开始时间
        try:
            try:
                response = pan.file.list(parent_file_id=folder_id, search_data=upload_file, search_mode=1, limit=1)   #尝试在云盘内精确搜索文件，检查文件是否已经上传，限制返回1个结果
                file_id = response.get("lastFileId", [])   #获取文件列表
                if file_id != -1:   #如果找到了匹配的文件（不是返回-1），则进行删除操作
                    pan.file.trash([file_id])  #移到回收站
                    log_print('Existing file in 123Pan deleted successfully: ' + upload_file + ' (File_ID: ' + str(file_id) + ')')
                else:
                    log_print('No matching file found in 123Pan: ' + upload_file + ', skip delete operation')
            except Exception as e:
                log_exception('Delete operation failed for ' + upload_file + ', upload will continue', e)

            pan.file.upload(parent_file_id=folder_id, file_path=upload_source_path)   #上传当前文件到123云盘
            log_print('Upload to 123Pan successfully: ' + upload_file)
            if (upload_file, upload_source_path) in upload_queue:
                upload_queue.remove((upload_file, upload_source_path))   #仅从上传成功的队列中移除文件
        except Exception as e:
            log_exception('Upload to 123Pan failed for ' + upload_file + '; file remains queued for retry', e)
        upload_end_time=datetime.datetime.now()   #记录上传操作结束时间
        upload_used_time=upload_end_time-upload_start_time   #计算上传所用时间
        log_print('Upload to 123Pan finished: ' + upload_file + ' in ' + str(upload_used_time) + ' s')
    if not upload_queue:
        log_print('Upload queue has been cleared')



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
        log_exception('Accurate backup failed', e)  # 打印精确备份失败信息
    


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
try:   #尝试加载图标文件
    image = Image.open('PythonLight.ico')   #图标文件路径
except FileNotFoundError:
    log_print('Icon file not found, using a white cube as default icon')
    image = Image.new('RGB', (64, 64), color=(255, 255, 255))   #创建一个简单的白色方块作为默认图标

menu = (item('Show/Hide console window', toggle_console), item('Exit program', exit_program))   #创建右键菜单

icon = pystray.Icon("office_backup_utilities", image, "Office Backup Utilities", menu)   #创建托盘图标对象
'''icon.on_left_click = on_clicked   #绑定左键单击事件处理函数（无法生效）'''


def global_exception_handler(exctype, value, tb):   #处理全局未捕获异常
    if issubclass(exctype, KeyboardInterrupt):   #正常响应Ctrl+C
        sys.__excepthook__(exctype, value, tb)
        return
    try:
        log_print('[ERROR] ' + ''.join(traceback.format_exception(exctype, value, tb)))
    except Exception as log_error:
        print('[ERROR] Failed to record uncaught exception: ' + type(log_error).__name__ + ': ' + str(log_error) + '\n' + traceback.format_exc(), file=sys.stderr)

def threading_exception_handler(args):   #处理守护线程未捕获异常
    if issubclass(args.exc_type, KeyboardInterrupt):   #正常响应Ctrl+C
        threading.__excepthook__(args)
        return
    global_exception_handler(args.exc_type, args.exc_value, args.exc_traceback)

sys.excepthook = global_exception_handler   #注册主线程异常处理器
threading.excepthook = threading_exception_handler   #注册守护线程异常处理器




icon_task = threading.Thread(target=icon.run)   #创建托盘图标线程
icon_task.daemon = True   #设置为守护线程（随主程序终止而自动结束）
icon_task.start()   #启动托盘图标线程

while True:   #主线程无限循环，防止程序退出
    try:
        if config.get('ppt_backup_enable'):   #检查PPT备份功能是否启用
            save_open_ppt_files(ppt_save_folder)   #启动线程
        if config.get('word_backup_enable'):   #检查Word备份功能是否启用
            save_open_word_files(word_save_folder)   #启动线程
        if config.get('wps_backup_enable'):   #检查WPS备份功能是否启用
            save_open_WPS_files(ppt_save_folder)   #启动线程
        if config.get('upload_to_123pan_enable'):   #检查上传到123云盘功能是否启用
            upload_to_123pan()   #启动线程
        if config.get('accurate_backup_enable'):  # 检查精确备份功能是否启用
            accurate_backup()  # 启动线程
    except Exception as e:
        log_exception('Main loop iteration failed', e)
    time.sleep(sleeptime)   #等待指定时间后继续轮询
