import os   #导入os模块，用于处理文件和目录路径操作
import shutil   #导入shutil模块，用于复制文件并保留元数据
import time   #导入time模块，用于添加延时和时间戳
import win32com.client as win32   #导入win32com.client库，用于通过COM接口与Microsoft Office应用程序交互

def save_open_ppt_files(save_folder):   #定义保存函数，参数save_folder是备份文件的存储路径
    if not os.path.exists(save_folder):   #检查备份目录是否存在
        os.makedirs(save_folder)   #若不存在则创建备份目录（包括所有必要的父目录）

    ppt_app=win32.DispatchEx('PowerPoint.Application')   #启动一个独立的PowerPoint实例

    while True:   #第一层（定义层）无限循环
        presentations = ppt_app.Presentations   #获取当前PowerPoint实例中所有打开的演示文稿集合

        for ppt in presentations:   #遍历集合
            ppt_path = ppt.FullName   #获取PPT文件的完整路径
            ppt_name = os.path.basename(ppt_path)   #提取文件名
            new_ppt_path = os.path.join(save_folder, ppt_name)   #生成备份路径 
            shutil.copy2(ppt_path, new_ppt_path)   #复制PPT文件到备份文件夹，并尝试保留元数据（如修改时间等）
            print(time.strftime('[%H:%M:%S]') + 'Successfully backuped ' + ppt_name + ' to ' + save_folder)   #打印带时间戳的成功信息
        
        time.sleep(180)   # 每3分钟轮询一次
        
save_folder=r'D:\tech\pptbackup'   #指定备份路径，r表示取原始字符串，需要更改请更改引号内部分
while True:   #第二层（调用层）无限循环
    try:
        save_open_ppt_files(save_folder)   #调用函数
    except:
        print(time.strftime('[%H:%M:%S]') + 'No ppt available now')
        #修复2018版希沃（Windows7和Office2016环境下）由于没有活动的ppt窗口而出现pywintypes.com_error报错后直接跳出第一层循环结束程序运行的问题
