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
    #文件夹精确备份功能
    "accurate_backup_enable": False,
    "accurate_backup_source_path": "",
    "accurate_backup_target_path": "",
    #日志保存设置
    "save_log": True,   #是否保存日志到OBUlatest.log文件，True为保存（默认），False为不保存
    "archive_previous_log": True,   #是否在程序启动时归档之前的日志，True为归档（默认），False为直接覆盖
    #超时设置
    "backup_timeout": 600,   #备份操作超时时间，单位为秒（默认10分钟）
}
try:   #读取配置文件
    with open('OBU6.0Core.json', 'r', encoding='utf-8') as f:   #尝试读取配置文件（只读）
        config = json.load(f)
    config_changed = False
    for key, value in default_config.items():   #如果现有配置文件有缺漏，根据默认配置项自动补全
        if key not in config:
            config[key] = value
            config_changed = True
    if config_changed:   #如果配置文件有新增项，写回配置文件
        with open('OBU6.0Core.json', 'w', encoding='utf-8') as f:
            json.dump(config, f, indent=4, ensure_ascii=False)
except (FileNotFoundError, json.JSONDecodeError):   #若配置文件不存在或无法解析
    config = default_config   #使用默认配置
    with open('OBU6.0Core.json', 'w', encoding='utf-8') as f:   #在当前目录下根据默认配置文件创建（写入）
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
def log_print(msg):   #定义日志打印函数
    global runid    #声明全局变量runid，以便在函数内修改其值
    runid+=1   #运行计数器累加
    log_msg= time.strftime('[%H:%M:%S-#') + str(runid) + '] ' + msg   # 打印带时间戳的日志消息到控制台
    print(log_msg)   # 打印日志消息到控制台
    if config.get('save_log'):   #如果启用日志保存功能，则将日志消息写入日志文件
        log_file.write(log_msg + '\n')   # 将日志消息写入日志文件
        log_file.flush()   #刷新文件缓冲区，确保日志消息立即写入文件



# 默认隐藏控制台窗口
console_window = ctypes.windll.kernel32.GetConsoleWindow()   #获取控制台窗口句柄
ctypes.windll.user32.ShowWindow(console_window, 0)   #隐藏控制台窗口



#初始化变量
runid=0   #初始化运行计数器
file_skip_count = defaultdict(int)   #使用字典记录每个文件的跳过次数（替代原全局skippedtime）
SaveAs_method_activated = defaultdict(bool)  # 使用字典记录每个文件是否已激活SaveAs方法
Existed_in_this_session = defaultdict(bool)  # 使用字典记录每个文件是否在本次运行中出现过，让之前会话中已经备份过的文件在程序重启后正常进行第一次备份
#从配置文件读取变量
sleeptime=config.get('interval')   #轮询间隔（默认为60秒）
ppt_save_folder=config.get('ppt_backup_path')   #ppt备份路径
word_save_folder=config.get('word_backup_path')   #word备份路径



if config.get('accurate_backup_enable'):  # 检查精确备份功能是否启用
    source_path = config.get('accurate_backup_source_path')   #获取源文件夹路径
    target_path = config.get('accurate_backup_target_path')   #获取目标文件夹路径
    if not source_path and target_path:   #如果精确备份功能开启但源路径为空或目标路径为空，则强制禁用精确备份功能
        log_print("Source path or target path for accurate backup is empty, force disabled accurate backup function, please provide valid paths in the configuration file")
        with open('OBU6.0Core.json', 'w', encoding='utf-8') as f:   #将禁用精确备份功能写入配置文件
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





@timeout(seconds=600, config_key='backup_timeout')  # 添加10分钟超时机制
def save_open_ppt_files(ppt_save_folder):   #定义ppt保存函数，参数ppt_save_folder是备份文件的存储路径
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
                
    except Exception as e:   #获取其他错误类型
            if type(e).__name__=='com_error':   #捕获无打开的PowerPoint实例而产生的的异常
                log_print('No ppt available now (PowerPoint application not detected)')   #打印异常信息
            else:   #打印出其他错误并继续轮询
                log_print('Exception: ' + type(e).__name__ + ', request continue')   #打印异常信息




@timeout(seconds=600, config_key='backup_timeout')  # 添加10分钟超时机制
def save_open_word_files(word_save_folder):   #定义word保存函数，参数word_save_folder是备份文件的存储路径
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
                
    except Exception as e:   #获取其他错误类型
            if type(e).__name__=='com_error':   #捕获无打开的PowerPoint实例而产生的的异常
                log_print('No doc available now (Word application not detected)')   #打印带时间戳和运行次数的异常信息
            else:   #打印出其他错误并继续轮询
                log_print('Exception: ' + type(e).__name__ + ', request continue')   #打印带时间戳和运行次数的异常信息




@timeout(seconds=600, config_key='backup_timeout')  # 添加10分钟超时机制
def save_open_WPS_files(ppt_save_folder):   #定义WPS保存函数，参数ppt_save_folder是备份文件的存储路径
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
                
    except Exception as e:   #获取其他错误类型
            if type(e).__name__=='com_error':   #捕获无打开的WPS实例而产生的的异常
                log_print('No ppt available now (WPS application not detected)')   #打印异常信息
            else:   #打印出其他错误并继续轮询
                log_print('Exception: ' + type(e).__name__ + ', request continue')   #打印异常信息




@timeout(seconds=600, config_key='backup_timeout')  # 添加10分钟超时机制
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
