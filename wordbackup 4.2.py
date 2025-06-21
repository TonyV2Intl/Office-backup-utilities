import os   #导入os模块，用于处理文件和目录路径操作
import shutil   #导入shutil模块，用于复制文件并保留元数据
import time   #导入time模块，用于添加延时和时间戳
import win32com.client as win32   #导入win32com.client库，用于通过COM接口与Microsoft Office应用程序交互
import datetime   #导入datetime库，用于计算备份所用时间
from collections import defaultdict  #导入collections库的defaultdict方法，用于跟踪单个文件的跳过次数

runid=0   #初始化运行计数器
file_skip_count = defaultdict(int)   #使用字典记录每个文件的跳过次数（替代原全局skippedtime）
SaveAs_method_activated = defaultdict(bool)  # 使用字典记录每个文件是否已激活SaveAs方法

def save_open_word_files(save_folder):   #定义保存函数，参数save_folder是备份文件的存储路径

    global runid   #声明全局变量runid，以便在函数内修改其值
    global documents   #声明全局变量documents，以便在函数内修改其值
    global new_doc_path   #声明全局变量new_doc_path，以便在函数内修改其值
    global doc_name   #声明全局变量doc_name，以便在函数内修改其值

    if not os.path.exists(save_folder):   #检查备份目录是否存在
        os.makedirs(save_folder)   #若不存在则创建备份目录（包括所有必要的父目录）
        runid+=1   #运行计数器累加
        print(time.strftime('[%H:%M:%S-#') + str(runid) + '] Target backup folder not found, created: ' + save_folder + ' successfully')   #打印成功创建备份目录的时间戳和运行次数

    word_app = win32.Dispatch('Word.Application')   #启动一个Word实例，若启用独立实例则无法获取当前已经打开的Word实例信息
    documents = word_app.Documents   #获取当前Word实例中所有打开的文档集合

    any_backup_performed = False   #标记本轮是否有任何备份操作（替代原havedoc）
        
    for doc in documents:   #遍历集合
        doc_path = doc.FullName   #获取Word文件的完整路径
        doc_name = os.path.basename(doc_path)   #提取文件名
        new_doc_path = os.path.join(save_folder, doc_name)   #生成备份路径

        if os.path.exists(new_doc_path):   #检查备份文件是否已存在
            if SaveAs_method_activated[doc_name] == True:   #如果SaveAs方法已被激活，则不再使用复制方法
                runid+=1
                print(time.strftime('[%H:%M:%S-#') + str(runid) + '] ' + doc_name + ' has already existed in ' + save_folder + ', skipped backup (SaveAs method activated)')   #打印带时间戳和运行次数的跳过信息
                continue   #跳过此次备份
            if file_skip_count[doc_name] < 5:  # 仅当同一文件连续跳过5次时才允许重新备份
                file_skip_count[doc_name] += 1   #该文件的跳过计数器累加
                runid+=1   #运行计数器累加
                if file_skip_count[doc_name] == 5:   # 如果跳过次数达到5次，打印提示信息
                    print(time.strftime('[%H:%M:%S-#') + str(runid) + '] ' + doc_name + ' has already existed in ' + save_folder + ', skipped backup (skipped times: ' + str(file_skip_count[doc_name]) + ', this file will be backed up again during the next request)')   #打印带时间戳和运行次数的跳过信息
                else:
                    print(time.strftime('[%H:%M:%S-#') + str(runid) + '] ' + doc_name + ' has already existed in ' + save_folder + ', skipped backup (skipped times: ' + str(file_skip_count[doc_name]) + ')')   #打印带时间戳和运行次数的跳过信息
                continue   #跳过此次备份

        runid+=1   #运行计数器累加
        print(time.strftime('[%H:%M:%S-#') + str(runid) + '] Start to backup ' + doc_name + ' to ' + save_folder)   #打印备份开始信息
        copystarttime=datetime.datetime.now()   #记录复制操作开始时间
        shutil.copy2(doc_path, new_doc_path)   #复制文档到备份文件夹，并尝试保留元数据（如修改时间等）
        copyendtime=datetime.datetime.now()   #记录复制操作结束时间
        copyusedtime=copyendtime-copystarttime  #计算复制所用时间

        modified_time=os.path.getmtime(new_doc_path)   #获取备份文件的修改时间
        create_time=os.path.getctime(new_doc_path)   #获取备份文件的创建时间
        os.utime(new_doc_path, (modified_time, create_time))   #将修改时间存储到访问时间（参数1），创建时间存储到修改时间（参数2），方便文件系统根据修改时间排序

        file_skip_count[doc_name] = 0   #重置该文件的跳过计数器
        any_backup_performed = True   #标记本轮有备份操作
        runid+=1   #运行计数器累加
        print(time.strftime('[%H:%M:%S-#') + str(runid) + '] Successfully backuped ' + doc_name + ' to ' + save_folder + ' in ' + str(copyusedtime) +' s')   #打印带时间戳和运行次数的备份成功信息
            
    if not any_backup_performed and len(documents) == 0:   #检查变量值，如果没有可备份PPT，打印此条信息
            runid+=1   #运行计数器累加
            print(time.strftime('[%H:%M:%S-#') + str(runid) + '] No doc available now')

sleeptime=180   #每3分钟轮询一次        
save_folder=r'C:\wordbackup'   #指定备份路径，r表示取原始字符串，需要更改请更改引号内部分

while True:   #无限循环
    try:
        save_open_word_files(save_folder)   #调用函数
        time.sleep(sleeptime)   #等待下次轮询
    except FileNotFoundError:   #捕获由于U盘等移动存储介质被移除而导致的“文件未找到”异常，使用2.0版本中的SaveAs方法进行备份
        for idx in range(1, documents.Count + 1):   #遍历文档实例集合
            doc = documents.Item(idx)   #获取当前文档实例
            if not os.path.exists(save_folder):   #再次检查备份目录是否存在
                os.makedirs(save_folder)   #若不存在则创建备份目录（包括所有必要的父目录）
                runid+=1   #运行计数器累加
                print(time.strftime('[%H:%M:%S-#') + str(runid) + '] Target backup folder not found, created: ' + save_folder + ' successfully')   #打印成功创建备份目录的时间戳和运行次数
            runid+=1   #运行计数器累加
            print(time.strftime('[%H:%M:%S-#') + str(runid) + '] Start to backup ' + doc_name + ' to ' + save_folder)   #打印备份开始信息
            savestarttime=datetime.datetime.now()   #记录保存操作开始时间
            doc.SaveAs(new_doc_path)   #使用SaveAs方法保存当前文档实例到指定路径
            saveendtime=datetime.datetime.now()   #记录保存操作结束时间
            saveusedtime=saveendtime-savestarttime  #计算保存所用时间
            runid+=1   #运行计数器累加
            SaveAs_method_activated[doc_name] = True   #标记该文件已激活SaveAs方法，后续不再备份
            print(time.strftime('[%H:%M:%S-#') + str(runid) + '] Detected access control, activated SaveAs method, successfully backuped ' + doc_name + ' to ' + save_folder + ' in ' + str(saveusedtime) + ' s')   #打印备份成功信息
            time.sleep(sleeptime)   #等待下次轮询
    except Exception as e:   #获取其他错误类型
        if type(e).__name__=='com_error':   #捕获无打开的PowerPoint实例而产生的的异常
            runid+=1   #运行计数器累加
            print(time.strftime('[%H:%M:%S-#') + str(runid) + '] No doc available now (Word application not detected)')   #打印带时间戳和运行次数的异常信息
            time.sleep(sleeptime)   #等待下次轮询
        else:   #打印出其他错误并继续轮询
            runid+=1   #运行计数器累加
            print(time.strftime('[%H:%M:%S-#') + str(runid) + '] Exception: ' + type(e).__name__ + ', request continue')   #打印带时间戳和运行次数的异常信息
            time.sleep(sleeptime)   #等待下次轮询