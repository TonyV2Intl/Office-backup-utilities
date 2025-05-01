import os   #导入os模块，用于处理文件和目录路径操作
import shutil   #导入shutil模块，用于复制文件并保留元数据
import time   #导入time模块，用于添加延时和时间戳
import win32com.client as win32   #导入win32com.client库，用于通过COM接口与Microsoft Office应用程序交互
import datetime   #导入datetime库，用于计算备份所用时间

runid=0   #初始化运行计数器

def save_open_word_files(save_folder):   #定义保存函数，参数save_folder是备份文件的存储路径

    global runid   #声明全局变量runid，以便在函数内修改其值
    global havedoc   #声明全局变量havedoc，以便在函数内修改其值
    global documents   #声明全局变量documents，以便在函数内修改其值
    global new_doc_path   #声明全局变量new_doc_path，以便在函数内修改其值
    global doc_name   #声明全局变量doc_name，以便在函数内修改其值

    if not os.path.exists(save_folder):   #检查备份目录是否存在
        os.makedirs(save_folder)   #若不存在则创建备份目录（包括所有必要的父目录）

    word_app = win32.Dispatch('Word.Application')   #启动一个Word实例，若启用独立实例则无法获取当前已经打开的Word实例信息
    
    while True:   #第一层（定义层）无限循环
        documents = word_app.Documents   #获取当前Word实例中所有打开的文档集合

        havedoc=0   #创建一个变量，用以标记是否有文档可供备份
        
        for doc in documents:   #遍历集合
            doc_path = doc.FullName   #获取Word文件的完整路径
            doc_name = os.path.basename(doc_path)   #提取文件名
            new_doc_path = os.path.join(save_folder, doc_name)   #生成备份路径

            if os.path.exists(new_doc_path):   #检查备份文件是否已存在
                runid+=1   #运行计数器累加
                runiddisplay=str(runid)   #将运行计数器转换为字符串，用于显示
                print(time.strftime('[%H:%M:%S-#') + runiddisplay + ']' + doc_name + ' has already existed in ' + save_folder + ' , skipped backup')   #打印带时间戳和运行次数的跳过信息
                havedoc=1   #更新变量值，标记备份询问成功，无需打印没有可备份文档的信息
                continue   #因为备份文件已存在，所以跳过备份

            copystarttime=datetime.datetime.now()   #记录复制操作开始时间
            shutil.copy2(doc_path, new_doc_path)   #复制文档到备份文件夹，并尝试保留元数据（如修改时间等）
            copyendtime=datetime.datetime.now()   #记录复制操作结束时间
            copyusedtime=copyendtime-copystarttime  #计算复制所用时间
            copyusedtimedisplay=str(copyusedtime)   #将时间差转换为字符串，用于显示

            modified_time=os.path.getmtime(new_doc_path)   #获取备份文件的修改时间
            access_time=os.path.getatime(new_doc_path)   #获取备份文件的创建时间
            os.utime(new_doc_path, (modified_time, access_time))   #将修改时间存储到访问时间（参数1），访问时间存储到修改时间（参数2），方便文件系统根据修改时间排序

            havedoc=1   #更新变量值，标记备份操作成功，无需打印没有可备份文档的信息
            runid+=1   #运行计数器累加
            runiddisplay=str(runid)   #将运行计数器转换为字符串，用于显示
            print(time.strftime('[%H:%M:%S-#') + runiddisplay + '] Successfully backuped ' + doc_name + ' to ' + save_folder + ' in ' + copyusedtimedisplay +' s')   #打印带时间戳和运行次数的备份成功信息
            
        if havedoc==0:   #检查变量值，如果没有可备份文档，打印此条信息
            runid+=1   #运行计数器累加
            runiddisplay=str(runid)   #将运行计数器转换为字符串，用于显示
            print(time.strftime('[%H:%M:%S-#') + runiddisplay + '] No doc available now')
        
        time.sleep(180)   # 每3分钟轮询一次
        
save_folder=r'C:\wordbackup'   #指定备份路径，r表示取原始字符串，需要更改请更改引号内部分
while True:   #第二层（调用层）无限循环
    try:
        save_open_word_files(save_folder)   #调用函数
    except FileNotFoundError:   #捕获由于U盘等移动存储介质被移除而导致的“文件未找到”异常，使用2.0版本中的SaveAs方法进行备份
        for idx in range(1, documents.Count + 1):   #遍历文档实例集合
            doc = documents.Item(idx)   #获取当前文档实例
            if not os.path.exists(save_folder):   #再次检查备份目录是否存在
                os.makedirs(save_folder)   #若不存在则创建备份目录（包括所有必要的父目录）
            savestarttime=datetime.datetime.now()   #记录保存操作开始时间
            doc.SaveAs(new_doc_path)   #使用SaveAs方法保存当前文档实例到指定路径
            saveendtime=datetime.datetime.now()   #记录保存操作结束时间
            saveusedtime=saveendtime-savestarttime  #计算保存所用时间
            saveusedtimedisplay=str(saveusedtime)   #将时间差转换为字符串，用于显示
            runid+=1   #运行计数器累加
            runiddisplay=str(runid)   #将运行计数器转换为字符串，用于显示
            print(time.strftime('[%H:%M:%S-#') + runiddisplay + '] Detected access control, activated SaveAs method, successfully backuped ' + doc_name + ' to ' + save_folder + ' in ' + saveusedtimedisplay + ' s')   #打印备份成功信息
    except Exception as e:   #如果需备份文档所在的Word实例被关闭，第一层循环内会因找不到实例而报错，通过第二层循环重新调用函数继续轮询
        runid+=1   #运行计数器累加
        runiddisplay=str(runid)   #将运行计数器转换为字符串，用于显示
        print(time.strftime('[%H:%M:%S-#') + runiddisplay + '] Detected doc close, request continue (Exception: ' + type(e).__name__ + ')')   #打印带时间戳和运行次数的关闭文档信息