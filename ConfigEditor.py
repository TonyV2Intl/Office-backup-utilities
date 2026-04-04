import tkinter as tk
from tkinter import ttk, messagebox, filedialog, scrolledtext
import json
import os
import copy
import time

class ConfigEditor:
    def __init__(self, root):
        self.root = root
        self.root.title("配置文件编辑器")
        self.root.geometry("854x550")
        self.root.resizable(True, True)
        
        # 设置中文字体
        self.font = ("SimHei", 10)
        
        # 版本配置信息
        self.version_configs = {
            "5.0": {
                "config_file": "Officebackup.json",
                "cloud_section": "123云盘参数",
                "cloud_params": ["client_id", "client_secret", "access_token", "folder_id"]
            },
            "5.1": {
                "config_file": "OBU5.1.json",
                "cloud_section": "OpenList参数",
                "cloud_params": ["openlist_url", "openlist_username", "openlist_password", "openlist_target_folder"]
            },
            "5.1Core": {
                "config_file": "OBU5.1Core.json",
                "cloud_section": None,
                "cloud_params": []
            }
        }
        
        # 配置数据
        self.current_version = "5.0"
        self.config_data = {}
        self.original_config = {}
        self.history = []  # 用于撤销/恢复操作
        self.history_index = -1
        self.unsaved_changes = False
        self.runid = 0
        
        # 创建界面
        self.create_widgets()
        self.create_log_widget()
        
        # 自定义样式
        self.setup_styles()
        
        # 初始隐藏配置界面
        self.config_frame.pack_forget()
        self.status_var.set("请选择版本")
    
    def create_widgets(self):
        # 主框架
        main_frame = ttk.Frame(self.root)
        main_frame.pack(fill=tk.BOTH, expand=True)
        
        # 顶部框架
        top_frame = ttk.Frame(main_frame, padding="10")
        top_frame.pack(fill=tk.X)
        
        # 版本选择
        ttk.Label(top_frame, text="版本选择:", font=self.font).grid(row=0, column=0, sticky=tk.W, padx=5, pady=5)
        
        self.version_var = tk.StringVar(value="")
        version_combobox = ttk.Combobox(
            top_frame, 
            textvariable=self.version_var, 
            values=["", "5.0", "5.1", "5.1Core"],
            state="readonly",
            width=10
        )
        version_combobox.grid(row=0, column=1, sticky=tk.W, padx=5, pady=5)
        version_combobox.bind("<<ComboboxSelected>>", lambda e: self.on_version_change())
        
        # 选项卡控件
        self.notebook = ttk.Notebook(main_frame)
        self.notebook.pack(fill=tk.BOTH, expand=True)
        
        # 配置编辑区域
        self.config_frame = ttk.Frame(self.notebook)
        self.notebook.add(self.config_frame, text="配置编辑")
        
        # 创建滚动条
        scrollbar = ttk.Scrollbar(self.config_frame)
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        
        # 创建画布用于滚动
        self.canvas = tk.Canvas(self.config_frame, yscrollcommand=scrollbar.set)
        self.canvas.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        scrollbar.config(command=self.canvas.yview)
        
        # 创建配置项容器
        self.config_content = ttk.Frame(self.canvas)
        self.config_content.pack(fill=tk.BOTH, expand=True)
        
        # 将配置内容框架添加到画布
        self.canvas.create_window((0, 0), window=self.config_content, anchor=tk.NW, tags="content")
        
        # 配置滚动区域
        def update_scrollregion(event):
            self.canvas.configure(scrollregion=self.canvas.bbox("all"))
        
        self.config_content.bind("<Configure>", update_scrollregion)
        
        # 控制按钮区域
        button_frame = ttk.Frame(main_frame, padding="10")
        button_frame.pack(fill=tk.X)
        
        ttk.Button(button_frame, text="撤销", command=self.undo, width=10).pack(side=tk.LEFT, padx=5)
        ttk.Button(button_frame, text="恢复", command=self.redo, width=10).pack(side=tk.LEFT, padx=5)
        ttk.Button(button_frame, text="保存", command=self.save_config, width=10, style="Green.TButton").pack(side=tk.LEFT, padx=5)
        ttk.Button(button_frame, text="退出", command=self.on_exit, width=10).pack(side=tk.LEFT, padx=5)
    
    def on_version_change(self):
        new_version = self.version_var.get()
        if new_version:
            if new_version != self.current_version:
                # 检查是否有未保存的更改
                if self.unsaved_changes:
                    if not messagebox.askyesno("提示", "当前有未保存的更改，是否继续切换版本？"):
                        self.version_var.set(self.current_version)
                        return
                
                self.current_version = new_version
                self.load_config()
                self.log_message(f"已加载版本 {new_version} 的配置文件")
        else:
            # 隐藏配置界面
            self.status_var.set("请选择版本")
            self.log_message("请选择要编辑的配置版本")
    
    def create_log_widget(self):
        # 日志显示区域
        log_frame = ttk.LabelFrame(self.root, text="操作日志")
        log_frame.pack(fill=tk.BOTH, expand=True)
        self.log_text = scrolledtext.ScrolledText(log_frame, wrap=tk.WORD, width=100, height=15)
        self.log_text.pack(fill=tk.BOTH, expand=True)
        
        # 状态栏
        self.status_var = tk.StringVar(value="就绪")
        status_bar = ttk.Label(self.root, textvariable=self.status_var, relief=tk.SUNKEN, anchor=tk.W)
        status_bar.pack(fill=tk.X, side=tk.BOTTOM)
    
    def setup_styles(self):
        style = ttk.Style()
        
        # 绿色按钮样式
        style.configure("Green.TButton", foreground="green", font=("SimHei", 10))
        
        # 红色按钮样式
        style.configure("Red.TButton", foreground="red", font=("SimHei", 10))
        
        # 选项卡样式
        style.configure("TNotebook.Tab", padding=(10, 1), font=("SimHei", 10))
    
    def log_message(self, message):
        self.runid += 1
        timestamp = time.strftime('[%H:%M:%S-#') + str(self.runid) + ']'
        self.log_text.insert(tk.END, f"{timestamp} {message}\n")
        self.log_text.see(tk.END)
        self.root.update()
    
    def load_config(self):
        config_file = self.version_configs[self.current_version]["config_file"]
        
        try:
            if os.path.exists(config_file):
                with open(config_file, 'r', encoding='utf-8') as f:
                    self.config_data = json.load(f)
                self.log_message(f"成功加载配置文件: {config_file}")
            else:
                # 如果文件不存在，使用默认配置
                self.config_data = self.get_default_config()
                # 保存默认配置到文件
                self.save_config_to_file(self.config_data, config_file)
                self.log_message(f"未找到配置文件，已创建默认配置: {config_file}")
            
            # 保存原始配置用于比较
            self.original_config = copy.deepcopy(self.config_data)
            # 重置历史记录
            self.history = [copy.deepcopy(self.config_data)]
            self.history_index = 0
            self.unsaved_changes = False
            
            # 更新配置界面
            self.update_config_ui()
            self.status_var.set(f"已加载版本 {self.current_version} 的配置")
            
        except Exception as e:
            messagebox.showerror("错误", f"加载配置文件失败: {str(e)}")
            self.log_message(f"加载配置文件失败: {str(e)}")
            self.status_var.set("加载配置失败")
    
    def get_default_config(self):
        # 根据版本返回默认配置
        if self.current_version == "5.0":
            return {
                "ppt_backup_path": "C:\\Officebackup\\pptbckup",
                "word_backup_path": "C:\\Officebackup\\wordbackup",
                "interval": 60,
                "max_skipping_time": 15,
                "ppt_backup_enable": True,
                "word_backup_enable": True,
                "wps_backup_enable": True,
                "upload_to_123pan_enable": True,
                "client_id": "",
                "client_secret": "",
                "access_token": "",
                "folder_id": 0,
                "accurate_backup_enable": False,
                "accurate_backup_source_path": "",
                "accurate_backup_target_path": "",
                "show_console_window_at_startup": False,
                "save_log": True
            }
        elif self.current_version == "5.1":
            return {
                "ppt_backup_path": "C:\\Officebackup\\pptbckup",
                "word_backup_path": "C:\\Officebackup\\wordbackup",
                "interval": 60,
                "max_skipping_time": 15,
                "ppt_backup_enable": True,
                "word_backup_enable": True,
                "wps_backup_enable": True,
                "upload_to_openlist_enable": True,
                "openlist_url": "",
                "openlist_username": "",
                "openlist_password": "",
                "openlist_target_folder": "",
                "accurate_backup_enable": False,
                "accurate_backup_source_path": "",
                "accurate_backup_target_path": "",
                "show_console_window_at_startup": False,
                "save_log": True
            }
        else:  # 5.1Core
            return {
                "ppt_backup_path": "C:\\Officebackup\\pptbckup",
                "word_backup_path": "C:\\Officebackup\\wordbackup",
                "interval": 60,
                "max_skipping_time": 15,
                "ppt_backup_enable": True,
                "word_backup_enable": True,
                "wps_backup_enable": True,
                "accurate_backup_enable": False,
                "accurate_backup_source_path": "",
                "accurate_backup_target_path": "",
                "save_log": True
            }
    
    def update_config_ui(self):
        # 清空现有配置项
        for widget in self.config_content.winfo_children():
            widget.destroy()
        
        # 创建配置项分组
        sections = {
            "备份路径": ["ppt_backup_path", "word_backup_path"],
            "时间设置": ["interval", "max_skipping_time"],
            "功能开关": ["ppt_backup_enable", "word_backup_enable", "wps_backup_enable"]
        }
        
        # 添加云盘相关配置（如果适用）
        cloud_section = self.version_configs[self.current_version]["cloud_section"]
        cloud_params = self.version_configs[self.current_version]["cloud_params"]
        if cloud_section and cloud_params:
            if self.current_version == "5.0":
                sections["功能开关"].append("upload_to_123pan_enable")
            elif self.current_version == "5.1":
                sections["功能开关"].append("upload_to_openlist_enable")
            sections[cloud_section] = cloud_params
        
        # 添加其他配置
        sections["精确备份"] = ["accurate_backup_enable", "accurate_backup_source_path", "accurate_backup_target_path"]
        
        # 添加控制台和日志设置（如果适用）
        if self.current_version != "5.1Core":
            sections["界面与日志"] = ["show_console_window_at_startup", "save_log"]
        else:
            sections["日志设置"] = ["save_log"]
        
        # 为每个分组创建框架
        for section_name, keys in sections.items():
            section_frame = ttk.LabelFrame(self.config_content, text=section_name, padding="10")
            section_frame.pack(fill=tk.X, padx=5, pady=5, anchor=tk.W)
            
            for key in keys:
                if key in self.config_data:
                    self.create_config_item(section_frame, key)
    
    def create_config_item(self, parent, key):
        value = self.config_data[key]
        frame = ttk.Frame(parent)
        frame.pack(fill=tk.X, pady=1)
        
        # 标签
        label = ttk.Label(frame, text=key, width=30, font=self.font)
        label.pack(side=tk.LEFT, padx=5)
        
        # 根据值类型创建不同的输入控件
        if isinstance(value, bool):
            var = tk.BooleanVar(value=value)
            var.trace_add("write", lambda *args, k=key, v=var: self.on_config_change(k, v.get()))
            ttk.Checkbutton(frame, variable=var).pack(side=tk.LEFT, padx=5)
        elif isinstance(value, int):
            var = tk.StringVar(value=str(value))
            var.trace_add("write", lambda *args, k=key, v=var: self.on_config_change(k, int(v.get()) if v.get().isdigit() else 0))
            ttk.Entry(frame, textvariable=var, width=20).pack(side=tk.LEFT, padx=5)
        else:  # 字符串
            var = tk.StringVar(value=str(value))
            var.trace_add("write", lambda *args, k=key, v=var: self.on_config_change(k, v.get()))
            
            # 为路径类型的配置项添加浏览按钮
            if "_path" in key:
                entry_frame = ttk.Frame(frame)
                entry_frame.pack(side=tk.LEFT, fill=tk.X, expand=True)
                ttk.Entry(entry_frame, textvariable=var, width=40).pack(side=tk.LEFT, fill=tk.X, expand=True, padx=5)
                ttk.Button(entry_frame, text="浏览...", command=lambda v=var: self.browse_path(v)).pack(side=tk.LEFT, padx=5)
            else:
                ttk.Entry(frame, textvariable=var, width=50).pack(side=tk.LEFT, fill=tk.X, expand=True, padx=5)
    
    def browse_path(self, var):
        path = filedialog.askdirectory()
        if path:
            var.set(path)
    
    def on_config_change(self, key, value):
        self.config_data[key] = value
        self.unsaved_changes = True
        
        # 更新历史记录
        self.history = self.history[:self.history_index + 1]
        self.history.append(copy.deepcopy(self.config_data))
        self.history_index += 1
        
        # 限制历史记录长度
        if len(self.history) > 50:
            self.history.pop(0)
            self.history_index -= 1
        
        self.status_var.set("有未保存的更改")
    
    def save_config(self):
        config_file = self.version_configs[self.current_version]["config_file"]
        
        # 备份原始配置
        backup_file = config_file + ".bak"
        if os.path.exists(config_file):
            try:
                import shutil
                shutil.copy2(config_file, backup_file)
                self.log_message(f"已创建配置备份: {backup_file}")
            except Exception as e:
                messagebox.showwarning("警告", f"无法创建备份文件: {str(e)}")
                self.log_message(f"无法创建备份文件: {str(e)}")
        
        # 保存配置
        try:
            self.save_config_to_file(self.config_data, config_file)
            self.original_config = copy.deepcopy(self.config_data)
            self.unsaved_changes = False
            self.status_var.set(f"配置已保存到: {config_file}")
            self.log_message(f"配置保存成功: {config_file}")
            messagebox.showinfo("成功", "配置保存成功！")
        except Exception as e:
            messagebox.showerror("错误", f"保存配置失败: {str(e)}")
            self.log_message(f"保存配置失败: {str(e)}")
            self.status_var.set("保存失败")
    
    def save_config_to_file(self, config, file_path):
        with open(file_path, 'w', encoding='utf-8') as f:
            json.dump(config, f, indent=4, ensure_ascii=False)
    
    def undo(self):
        if self.history_index > 0:
            self.history_index -= 1
            self.config_data = copy.deepcopy(self.history[self.history_index])
            self.unsaved_changes = True
            self.update_config_ui()
            self.status_var.set("已撤销更改")
            self.log_message("已撤销更改")
    
    def redo(self):
        if self.history_index < len(self.history) - 1:
            self.history_index += 1
            self.config_data = copy.deepcopy(self.history[self.history_index])
            self.unsaved_changes = True
            self.update_config_ui()
            self.status_var.set("已恢复更改")
            self.log_message("已恢复更改")
    
    def on_exit(self):
        if self.unsaved_changes:
            if not messagebox.askyesno("提示", "当前有未保存的更改，是否退出？"):
                return
        self.log_message("程序正在退出...")
        self.root.destroy()

if __name__ == "__main__":
    root = tk.Tk()
    
    try:
        root.iconbitmap(default='PythonLight.ico')
    except:
        pass
    
    app = ConfigEditor(root)
    app.log_message("配置文件编辑器已启动")
    app.log_message("请选择要编辑的配置版本")
    root.mainloop()