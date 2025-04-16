import os   #导入os模块，用于处理文件和目录路径操作
import shutil   #导入shutil模块，用于复制文件并保留元数据
import time   #导入time模块，用于添加延时和时间戳
import win32com.client as win32   #导入win32com.client库，用于通过COM接口与Microsoft Office应用程序交互

def save_open_word_files(save_folder):   #定义保存函数，参数save_folder是备份文件的存储路径
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
            shutil.copy2(doc_path, new_doc_path)   #复制Word文件到备份文件夹，并尝试保留元数据（如修改时间等）
            print(time.strftime('[%H:%M:%S]') + 'Successfully backuped ' + doc_name + ' to ' + save_folder)   #打印带时间戳的成功信息
            havedoc=1   #更新变量值，标记备份操作成功，无需打印没有可备份文档的信息
        
        if havedoc==0:   #检查变量值，如果没有可备份文档，打印此条信息
            print(time.strftime('[%H:%M:%S]') +'No doc available now')
        
        time.sleep(180)   # 每3分钟轮询一次
        
save_folder=r'C:wordbackup'   #指定备份路径，r表示取原始字符串，需要更改请更改引号内部分
while True:   #第二层（调用层）无限循环
    try:
        save_open_word_files(save_folder)   #调用函数
    except:
        print(time.strftime('[%H:%M:%S]') + 'Detected doc close, request continue')
        #如果需备份文档所在的Word实例被关闭，第一层循环内会因找不到实例而报错，通过第二层循环重新调用函数继续轮询
