import os   #导入os模块，用于处理文件和目录路径操作
import shutil   #导入shutil模块，用于复制文件并保留元数据
import time   #导入time模块，用于时间相关操作
import win32com.client as win32   #导入win32com.client库，用于通过COM接口与Microsoft Office应用程序交互
import datetime   #导入datetime库，用于计算备份所用时间
from collections import defaultdict  #导入collections库的defaultdict方法，用于跟踪单个文件的跳过次数

import threading  #导入threading库，用于多线程操作
import json  #导入json库，用于处理配置文件的读写

import pystray   #导入pystray库，用于创建系统托盘图标
from pystray import MenuItem as item   #从pystray库中导入MenuItem类，用于创建托盘菜单项
from PIL import Image   #导入PIL库的Image模块，用于处理图标图像
import ctypes   #导入ctypes库，用于调用Windows API函数





#设定默认配置文件
default_config = {
    #指定备份路径，r表示取原始字符串，需要更改请更改引号内部分
    "ppt_backup_path": "C:\\Officebackup\\pptbckup",   #PPT、WPS备份路径
    "word_backup_path": "C:\\Officebackup\\wordbackup",   #Word备份路径
    #指定间隔时间，单位为秒
    "interval": 60,   #指定所有操作的轮询时间间隔，单位为秒（默认60秒）
    "max_skipping_time": 15,   #指定连续跳过次数（默认15次）
    #功能开启或禁用
    "ppt_backup_enable": True,   #PPT备份功能
    "word_backup_enable": True,   #Word备份功能
    "wps_backup_enable": True,   #WPS备份功能
    "upload_to_ftp_enable": True,   #上传到FTP服务器功能
    #FTP服务器参数
    "ftp_host": "",   #FTP服务器地址（FRP穿透后的公网地址）
    "ftp_port": 21,   #FTP服务器端口（FRP映射的端口）
    "ftp_username": "",   #FTP用户名
    "ftp_password": "",   #FTP密码
    "ftp_target_path": "",   #FTP目标路径（NAS上的路径）
    #文件夹精确备份功能
    "accurate_backup_enable": False,
    "accurate_backup_source_path": "",
    "accurate_backup_target_path": "",
    #托盘图标、控制台行为与日志保存设置
    #"tray_left_click_behavior": "open_console",   #托盘图标左键点击行为，选项有"open_console"（打开控制台）和"exit_program"（退出程序）（无法生效）
    "show_console_window_at startup": False,   #程序启动时显示控制台窗口，True为显示，False为隐藏（默认）
    "save_log": True   #是否保存日志到latest.log文件，True为保存（默认），False为不保存
}
try:   #读取配置文件
    with open('OfficebackupSingleConfig.json', 'r', encoding='utf-8') as f:   #尝试读取配置文件（只读）
        config = json.load(f)
    for key, value in default_config.items():   #如果现有配置文件有缺漏，根据默认配置项自动补全
        if key not in config:
            config[key] = value
except (FileNotFoundError, json.JSONDecodeError):   #若配置文件不存在或无法解析
    config = default_config   #使用默认配置
    with open('OfficebackupSingleConfig.json', 'w', encoding='utf-8') as f:   #在当前目录下根据默认配置文件创建（写入）
        json.dump(config, f, indent=4, ensure_ascii=False)   #写入默认配置文件



if config.get('save_log'):   #检查是否启用日志保存功能
    if os.path.exists('latest.log'):   #如果日志文件存在，则删除旧日志文件
        os.remove('latest.log')
    log_file = open('latest.log', 'a', encoding='utf-8')   #以追加模式打开日志文件
def log_print(msg):   #定义日志打印函数
    global runid    #声明全局变量runid，以便在函数内修改其值
    runid+=1   #运行计数器累加
    log_msg= time.strftime('[%H:%M:%S-#') + str(runid) + '] ' + msg   # 打印带时间戳的日志消息到控制台
    print(log_msg)   # 打印日志消息到控制台
    if config.get('save_log'):   #如果启用日志保存功能，则将日志消息写入日志文件
        log_file.write(log_msg + '\n')   # 将日志消息写入日志文件
        log_file.flush()   #刷新文件缓冲区，确保日志消息立即写入文件



console_visible = config.get('show_console_window_at startup')   #获取控制台窗口初始状态参数（默认为隐藏）
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
max_skipping_time=config.get('max_skipping_time')   #连续跳过次数（默认为15次）
ppt_save_folder=config.get('ppt_backup_path')   #ppt备份路径
word_save_folder=config.get('word_backup_path')   #word备份路径

#FTP相关配置
ftp_enable = config.get('upload_to_ftp_enable')
ftp_host = config.get('ftp_host')
ftp_port = config.get('ftp_port')
ftp_username = config.get('ftp_username')
ftp_password = config.get('ftp_password')
ftp_target_path = config.get('ftp_target_path')

#检查FTP配置是否完整
if ftp_enable:
    if not ftp_host or not ftp_username or not ftp_password:
        log_print("FTP configuration is incomplete, force disabled upload function, please provide valid credentials in the configuration file")
        config['upload_to_ftp_enable'] = False
        ftp_enable = False

if config.get('accurate_backup_enable'):  # 检查精确备份功能是否启用
    source_path = config.get('accurate_backup_source_path')   #获取源文件夹路径
    target_path = config.get('accurate_backup_target_path')   #获取目标文件夹路径
    if not source_path and target_path:   #如果精确备份功能开启但源路径为空或目标路径为空，则强制禁用精确备份功能
        log_print("Source path or target path for accurate backup is empty, force disabled accurate backup function, please provide valid paths in the configuration file")
        with open('OfficebackupSingleConfig.json', 'w', encoding='utf-8') as f:   #将禁用精确备份功能写入配置文件
                config['accurate_backup_enable'] = False   #强制禁用精确备份功能
                json.dump(config, f, indent=4, ensure_ascii=False)   #写入更新后的配置文件





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
                log_print('Exception: ' + type(e).__name__ + ', request continue')   #打印异常信息



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
                log_print('Exception: ' + type(e).__name__ + ', request continue')   #打印带时间戳和运行次数的异常信息



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
                log_print('Exception: ' + type(e).__name__ + ', request continue')   #打印异常信息



def upload_to_ftp():   #定义FTP上传函数
    import ftplib  #导入ftplib库，用于FTP操作
    global upload_queue  #声明全局上传队列变量
    
    if not ftp_enable:  #如果FTP功能未启用，则返回
        return
        
    # 自定义FTP类，修复PASV响应中的IP地址并验证被动端口范围
    class FixedPASVFTP(ftplib.FTP):
        def __init__(self, host, *args, passive_port_range=None, **kwargs):
            super().__init__(*args, **kwargs)
            self.real_host = host  # 保存真实的公网IP
            self.passive_port_range = passive_port_range  # 被动端口范围
        
        def makepasv(self):
            """
            重写makepasv方法，解析PASV响应并替换IP地址，验证被动端口范围
            - 将本地IP地址替换为公网IP
            - 验证被动端口是否在指定范围内
            """
            # 发送PASV命令并获取响应
            response = self.sendcmd('PASV')
            log_print(f"DEBUG: PASV response: {response}")
            
            # 手动解析PASV响应
            import re
            match = re.search(r'227.*?\((\d+),(\d+),(\d+),(\d+),(\d+),(\d+)\)', response)
            if not match:
                raise ftplib.error_reply(response)
            
            # 提取原始IP和端口
            h1, h2, h3, h4, p1, p2 = map(int, match.groups())
            original_host = f"{h1}.{h2}.{h3}.{h4}"
            final_port = p1 * 256 + p2
            
            log_print(f"DEBUG: Original PASV info: host={original_host}, port={final_port}")
            
            # 验证被动端口是否在指定范围内（55752-55753）
            passive_port_range = range(55752, 55753)
            if final_port < passive_port_range.start or final_port > passive_port_range.stop - 1:
                log_print(f"WARNING: PASV port {final_port} is not in expected range {passive_port_range}")
            else:
                log_print(f"DEBUG: PASV port {final_port} is within expected range {passive_port_range}")
            
            # 替换本地IP为真实公网IP
            if original_host in ('127.0.0.1', 'localhost', '0.0.0.0'):
                final_host = self.real_host
                log_print(f"DEBUG: Replaced PASV IP: {original_host} -> {final_host}")
            else:
                final_host = original_host
                log_print(f"DEBUG: Using original PASV IP: {final_host}")
            
            return final_host, final_port
        
    while upload_queue:  #当上传队列不为空时
        # 使用列表副本遍历，避免在遍历过程中修改列表
        for (upload_file, upload_source_path) in list(upload_queue):
            log_print('Start to upload ' + upload_file + ' to FTP server')   #打印上传开始信息
            upload_start_time=datetime.datetime.now()   #记录上传操作开始时间
            
            try:
                # 连接FTP服务器
                ftp = FixedPASVFTP(ftp_host)
                ftp.connect(ftp_host, ftp_port, timeout=60)  # 增加超时时间
                ftp.login(ftp_username, ftp_password)
                log_print('Successfully connected to FTP server')
                
                # 使用被动模式
                ftp.set_pasv(True)
                log_print('FTP passive mode enabled')
                
                # 切换到目标目录
                if ftp_target_path:
                    try:
                        ftp.cwd(ftp_target_path)
                        log_print('Successfully changed to target directory: ' + ftp_target_path)
                    except ftplib.error_perm as e:
                        log_print(f'Failed to change directory to {ftp_target_path}: {e}, creating directory...')
                        # 如果目录不存在则创建
                        dirs = ftp_target_path.split('/')
                        current_dir = ''
                        for dir in dirs:
                            if dir:
                                current_dir += '/' + dir
                                try:
                                    ftp.cwd(current_dir)
                                except ftplib.error_perm:
                                    ftp.mkd(current_dir)
                                    ftp.cwd(current_dir)
                        log_print('Successfully created target directory: ' + ftp_target_path)
                
                # 检查文件是否已存在
                try:
                    ftp.size(upload_file)  #尝试获取文件大小，如果成功则文件存在
                    log_print('File already exists on FTP server: ' + upload_file + ', deleting...')
                    ftp.delete(upload_file)  #删除已存在的文件
                    log_print('Successfully deleted existing file: ' + upload_file)
                except ftplib.error_perm:
                    log_print('No existing file found on FTP server: ' + upload_file + ', proceeding with upload')
                
                # 上传文件
                with open(upload_source_path, 'rb') as f:
                    try:
                        # 设置底层套接字超时时间
                        ftp.sock.settimeout(30)  # 设置30秒超时
                        try:
                            ftp.storbinary('STOR ' + upload_file, f)
                            log_print('Upload to FTP server successfully: ' + upload_file)
                        except (TimeoutError, ConnectionResetError, ftplib.error_temp) as e:
                            # 捕获超时等临时错误，需要重新连接检查
                            log_print('⚠ File upload timed out, attempting to reconnect and check...')
                            log_print(f'Error: {e}')
                            # 关闭当前连接（对于超时情况，直接关闭socket而不是发送QUIT命令）
                            try:
                                ftp.quit()
                            except Exception as quit_error:
                                log_print(f"   WARNING: Failed to quit cleanly, closing socket directly: {quit_error}")
                                # 直接关闭socket
                                try:
                                    ftp.sock.close()
                                except Exception as close_error:
                                    log_print(f"   WARNING: Failed to close socket: {close_error}")
                            
                            # 重新建立FTP连接
                            log_print("   Reconnecting to FTP server...")
                            ftp = FixedPASVFTP(ftp_host)
                            ftp.connect(ftp_host, ftp_port, timeout=60)
                            ftp.login(ftp_username, ftp_password)
                            log_print("   ✓ Reconnected to FTP server")
                            
                            # 重新切换到目标目录
                            if ftp_target_path:
                                ftp.cwd(ftp_target_path)
                                log_print(f"   ✓ Changed to directory: {ftp_target_path}")
                            
                            # 检查文件是否实际已上传
                            try:
                                file_size = ftp.size(upload_file)
                                log_print(f'✓ File actually exists on server, size: {file_size} bytes')
                            except ftplib.error_perm:
                                log_print('✗ File not found on server after timeout')
                    except Exception as e:
                        log_print(f'⚠ Unexpected error during upload: {e}')
                # 关闭FTP连接
                ftp.quit()
                
            except Exception as e:
                log_print('FTP upload failed: ' + str(e))
            
            upload_end_time=datetime.datetime.now()   #记录上传操作结束时间
            upload_used_time=upload_end_time-upload_start_time   #计算上传所用时间
            log_print('FTP upload finished: ' + upload_file + ' in ' + str(upload_used_time) + ' s')
            upload_queue.remove((upload_file, upload_source_path))   #从上传队列中移除已处理的文件
    log_print('FTP upload queue has been cleared')



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
try:   #尝试加载图标文件
    image = Image.open('PythonLight.ico')   #图标文件路径
except FileNotFoundError:
    log_print('Icon file not found, using a white cube as default icon')
    image = Image.new('RGB', (64, 64), color=(255, 255, 255))   #创建一个简单的白色方块作为默认图标

menu = (item('Show/Hide console window', toggle_console), item('Exit program', exit_program))   #创建右键菜单

icon = pystray.Icon("office_backup_utilities", image, "Office Backup Utilities", menu)   #创建托盘图标对象
'''icon.on_left_click = on_clicked   #绑定左键单击事件处理函数（无法生效）'''





icon_task = threading.Thread(target=icon.run)   #创建托盘图标线程
icon_task.daemon = True   #设置为守护线程（随主程序终止而自动结束）
icon_task.start()   #启动托盘图标线程

while True:   #主线程无限循环，防止程序退出
    if config.get('ppt_backup_enable'):   #检查PPT备份功能是否启用
        save_open_ppt_files(ppt_save_folder)   #启动线程
    if config.get('word_backup_enable'):   #检查Word备份功能是否启用
        save_open_word_files(word_save_folder)   #启动线程
    if config.get('wps_backup_enable'):   #检查WPS备份功能是否启用
        save_open_WPS_files(ppt_save_folder)   #启动线程
    if config.get('upload_to_ftp_enable'):   #检查上传到FTP服务器功能是否启用
        upload_to_ftp()   #启动线程
    if config.get('accurate_backup_enable'):  # 检查精确备份功能是否启用
        accurate_backup()  # 启动线程
    time.sleep(sleeptime)   #等待指定时间后继续轮询
