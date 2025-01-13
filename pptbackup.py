import os
import time
import win32com.client as win32

def save_open_ppt_files(save_folder):
    if not os.path.exists(save_folder):
        os.makedirs(save_folder)

    ppt_app=win32.Dispatch('PowerPoint.Application')
    
    while True:
        # 获取所有打开的PPT文档
        presentations = ppt_app.Presentations
        for idx in range(1, presentations.Count + 1):
            ppt = presentations.Item(idx)
            ppt_path = ppt.FullName
            ppt_name, _ = os.path.splitext(os.path.basename(ppt_path))
            new_ppt_path = os.path.join(save_folder, ppt_name + '_backup.pptx')
            
            # 检查是否已经备份过（避免重复备份）
            if not os.path.exists(new_ppt_path):
                ppt.SaveAs(new_ppt_path, FileFormat=32)  # 32 表示 ppSaveAsOpenXMLPresentation (.pptx)
                print('Successfully backuped ' + ppt_name + ' to ' + save_folder)
        
        # 等待一段时间再检查（避免频繁检查占用资源）
        time.sleep(180)  # 每300秒检查一次
        print('request complete')
        

# 示例用法
save_folder=r'D:\tech\pptbackup'
for i in range(1,20000):
    try:
        save_open_ppt_files(save_folder)
    except:
        print('no ppt available now')
    

