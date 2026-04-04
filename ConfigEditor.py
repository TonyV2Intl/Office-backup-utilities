import tkinter as tk
from tkinter import ttk, messagebox, filedialog
import json
import os
import copy

class ConfigEditor:
    def __init__(self, root):
        self.root = root
        self.root.title("Office Backup Utility Config Editor")
        self.root.geometry("600x400")
        self.root.resizable(True, True)
        # 设置窗口最小尺寸
        self.root.minsize(600, 400)
        
        # 设置中文字体
        self.font = ("SimHei", 10)
        
        # 版本配置信息
        self.version_configs = {
            "5.0": {
                "config_file": "OfficebackupSingleConfig.json",
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
        
        # 键名显示模式
        self.key_name_mode = "simple"  # 默认显示简明键名
        self.key_name_button = None
        
        # 创建界面
        self.create_widgets()
        self.create_status_bar()
        
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
        
        # 左侧版本选择区域
        version_frame = ttk.Frame(top_frame)
        version_frame.pack(side=tk.LEFT, padx=5, pady=5, fill=tk.Y)
        
        # 版本选择
        ttk.Label(version_frame, text="版本选择:", font=self.font).pack(side=tk.LEFT, padx=5, fill=tk.Y)
        
        self.version_var = tk.StringVar(value="")
        version_combobox = ttk.Combobox(
            version_frame, 
            textvariable=self.version_var, 
            values=["", "5.0", "5.1", "5.1Core"],
            state="readonly",
            width=10,
            takefocus=False
        )
        version_combobox.pack(side=tk.LEFT, padx=5, fill=tk.Y)
        version_combobox.bind("<<ComboboxSelected>>", lambda e: self.on_version_change())
        
        # 右侧按钮区域
        button_frame = ttk.Frame(top_frame)
        button_frame.pack(side=tk.RIGHT, padx=5, pady=5, fill=tk.Y)
        
        # 其他按钮
        ttk.Button(button_frame, text="启动", command=self.start_program, takefocus=False,width=5).pack(side=tk.RIGHT, padx=5, fill=tk.Y)
        # 键名显示模式切换按钮
        initial_text = "切换到原始键名" if self.key_name_mode == "simple" else "切换到简明键名"
        self.key_name_button = ttk.Button(button_frame, text=initial_text, command=self.toggle_key_name_mode, takefocus=False)
        self.key_name_button.pack(side=tk.RIGHT, padx=5, fill=tk.Y)
        ttk.Button(button_frame, text="重做(下一步)", command=self.redo, takefocus=False).pack(side=tk.RIGHT, padx=5, fill=tk.Y)
        ttk.Button(button_frame, text="撤销(上一步)", command=self.undo, takefocus=False).pack(side=tk.RIGHT, padx=5, fill=tk.Y)
        
        # 配置编辑区域
        self.config_frame = ttk.Frame(main_frame, padding="10")
        self.config_frame.pack(fill=tk.BOTH, expand=True)
        
        # 创建滚动条
        scrollbar = ttk.Scrollbar(self.config_frame, takefocus=False)
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        
        # 创建画布用于滚动
        self.canvas = tk.Canvas(self.config_frame, yscrollcommand=scrollbar.set, takefocus=0, highlightthickness=0)
        self.canvas.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        
        # 为滚动条设置命令
        scrollbar.config(command=self.canvas.yview)
        
        # 创建配置项容器
        self.config_content = ttk.Frame(self.canvas)
        # 使用fill=tk.X确保配置内容框架宽度随画布自动调整
        self.config_content.pack(fill=tk.BOTH, expand=True)
        
        # 将配置内容框架添加到画布
        self.canvas.create_window((0, 0), window=self.config_content, anchor=tk.NW, tags="content")
        
        # 配置滚动区域和画布宽度调整
        def update_scrollregion(event):
            # 更新滚动区域
            self.canvas.configure(scrollregion=self.canvas.bbox("all"))
            # 更新画布内框架的宽度，使其与画布宽度一致
            self.canvas.itemconfig("content", width=self.canvas.winfo_width())
        
        self.config_content.bind("<Configure>", update_scrollregion)
        # 监听画布大小变化
        self.canvas.bind("<Configure>", update_scrollregion)
        
        # 为画布添加鼠标滚轮事件，实现滚动功能
        def on_mouse_wheel(event):
            # 根据鼠标滚轮的方向调整滚动位置
            # 使用不同的滚动单位，使滚动更流畅
            scroll_amount = int(-event.delta / 120)
            self.canvas.yview_scroll(scroll_amount, "units")
            # 打印调试信息
            print(f"Mouse wheel delta: {event.delta}, scroll amount: {scroll_amount}")
            # 防止事件冒泡，确保事件只被处理一次
            return "break"
        
        # 绑定鼠标滚轮事件到画布
        self.canvas.bind("<MouseWheel>", on_mouse_wheel)
        
        # 为配置内容框架添加鼠标滚轮事件，确保在框架上滚动也能生效
        self.config_content.bind("<MouseWheel>", on_mouse_wheel)
        
        # 为所有分组框架也添加鼠标滚轮事件
        def bind_wheel_events():
            # 延迟绑定，确保所有控件都已创建
            def bind_recursive(widget):
                for child in widget.winfo_children():
                    if isinstance(child, (ttk.Frame, ttk.LabelFrame)):
                        # 为框架绑定鼠标滚轮事件
                        child.bind("<MouseWheel>", on_mouse_wheel)
                        # 递归处理子控件
                        bind_recursive(child)
            
            self.root.after(100, lambda: bind_recursive(self.config_content))
        
        # 绑定滚轮事件
        self.bind_wheel_events = bind_wheel_events
        
        # 为所有框架添加点击事件，当点击空白处时移除焦点
        def on_frame_click(event):
            # 检查点击的是否是框架本身（即空白处）
            if isinstance(event.widget, (ttk.Frame, ttk.LabelFrame)):
                # 直接使用root窗口获取焦点，这样可以移除所有控件的焦点
                self.root.focus_set()
                # 打印调试信息
                print(f"Clicked on {event.widget}, removing focus")
        
        # 为所有框架绑定点击事件
        def bind_frame_click_events():
            # 延迟绑定，确保所有控件都已创建
            def bind_recursive(widget):
                for child in widget.winfo_children():
                    if isinstance(child, (ttk.Frame, ttk.LabelFrame)):
                        # 为框架绑定点击事件
                        child.bind("<Button-1>", on_frame_click)
                        # 递归处理子控件
                        bind_recursive(child)
            
            self.root.after(100, lambda: bind_recursive(self.config_content))
        
        # 绑定框架点击事件
        self.bind_frame_click_events = bind_frame_click_events
        
        # 为所有输入控件绑定焦点事件，确保焦点停留在输入控件上
        def bind_focus_events():
            # 延迟绑定，确保所有控件都已创建
            self.root.after(100, lambda: self._bind_focus_events_recursive(self.config_content))
        
        self.bind_focus_events = bind_focus_events
    
    def on_version_change(self):
        new_version = self.version_var.get()
        if new_version:
            # 检查是否有未保存的更改
            if self.unsaved_changes and new_version != self.current_version:
                if not messagebox.askyesno("提示", "当前有未保存的更改，是否继续切换版本？"):
                    self.version_var.set(self.current_version)
                    return
            
            self.current_version = new_version
            # 显示配置界面
            self.config_frame.pack(fill=tk.BOTH, expand=True)
            self.load_config()
            self.status_var.set(f"已加载版本 {new_version} 的配置文件")
        else:
            # 隐藏配置界面
            self.config_frame.pack_forget()
            self.status_var.set("请选择版本")
    
    def create_status_bar(self):
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
    

    
    def load_config(self):
        config_file = self.version_configs[self.current_version]["config_file"]
        
        try:
            if os.path.exists(config_file):
                with open(config_file, 'r', encoding='utf-8') as f:
                    self.config_data = json.load(f)
                self.status_var.set(f"成功加载配置文件: {config_file}")
            else:
                # 如果文件不存在，使用默认配置
                self.config_data = self.get_default_config()
                # 保存默认配置到文件
                self.save_config_to_file(self.config_data, config_file)
                self.status_var.set(f"未找到配置文件，已创建默认配置: {config_file}")
            
            # 保存原始配置用于比较
            self.original_config = copy.deepcopy(self.config_data)
            # 重置历史记录
            self.history = [copy.deepcopy(self.config_data)]
            self.history_index = 0
            self.unsaved_changes = False
            
            # 更新配置界面
            self.update_config_ui()
            
        except Exception as e:
            messagebox.showerror("错误", f"加载配置文件失败: {str(e)}")
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
    
    def _bind_focus_events_recursive(self, widget):
        # 递归地为所有输入控件绑定焦点事件
        for child in widget.winfo_children():
            # 检查是否是输入控件
            if isinstance(child, (ttk.Entry, ttk.Checkbutton, ttk.Combobox)):
                # 为输入控件绑定焦点事件，确保焦点停留在控件上
                pass
            elif isinstance(child, ttk.Frame) or isinstance(child, ttk.LabelFrame):
                # 递归处理子框架
                self._bind_focus_events_recursive(child)
    
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
            section_frame.pack(fill=tk.X, padx=5, pady=5)
            
            for key in keys:
                if key in self.config_data:
                    self.create_config_item(section_frame, key)
        
        # 绑定焦点事件，防止画布获得焦点
        self.bind_focus_events()
        # 绑定框架点击事件，实现点击空白处去除焦点的功能
        self.bind_frame_click_events()
        # 绑定滚轮事件，实现滚动功能
        self.bind_wheel_events()
    
    def create_config_item(self, parent, key):
        value = self.config_data[key]
        frame = ttk.Frame(parent)
        frame.pack(fill=tk.X, pady=1)
        
        # 标签
        display_name = self.get_key_display_name(key)
        label = ttk.Label(frame, text=display_name, font=self.font)
        label.pack(side=tk.LEFT, padx=5)
        
        # 根据值类型创建不同的输入控件
        if isinstance(value, bool):
            var = tk.BooleanVar(value=value)
            var.trace_add("write", lambda *args, k=key, v=var: self.on_config_change(k, v.get()))
            ttk.Checkbutton(frame, variable=var, takefocus=False).pack(side=tk.LEFT, padx=5)
        elif isinstance(value, int):
            var = tk.StringVar(value=str(value))
            var.trace_add("write", lambda *args, k=key, v=var: self.on_config_change(k, int(v.get()) if v.get().isdigit() else 0))
            entry = ttk.Entry(frame, textvariable=var, takefocus=False)
            entry.pack(side=tk.LEFT, fill=tk.X, expand=True, padx=5)
        else:  # 字符串
            var = tk.StringVar(value=str(value))
            var.trace_add("write", lambda *args, k=key, v=var: self.on_config_change(k, v.get()))
            
            # 为路径类型的配置项添加浏览按钮
            if "_path" in key:
                entry_frame = ttk.Frame(frame)
                entry_frame.pack(side=tk.LEFT, fill=tk.X, expand=True)
                entry = ttk.Entry(entry_frame, textvariable=var, takefocus=False)
                entry.pack(side=tk.LEFT, fill=tk.X, expand=True, padx=5)
                ttk.Button(entry_frame, text="浏览...", command=lambda v=var: self.browse_path(v), takefocus=False).pack(side=tk.LEFT, padx=5)
            else:
                entry = ttk.Entry(frame, textvariable=var, takefocus=False)
                entry.pack(side=tk.LEFT, fill=tk.X, expand=True, padx=5)
    
    def browse_path(self, var):
        path = filedialog.askdirectory()
        if path:
            # 确保路径使用Windows格式的反斜杠
            path = path.replace('/', '\\')
            var.set(path)
    
    def on_config_change(self, key, value):
        # 处理路径格式，确保Windows路径在JSON中正确保存
        if "_path" in key and value:
            # 确保路径使用正确的Windows格式
            value = os.path.normpath(value)
        
        self.config_data[key] = value
        
        # 更新历史记录
        self.history = self.history[:self.history_index + 1]
        self.history.append(copy.deepcopy(self.config_data))
        self.history_index += 1
        
        # 限制历史记录长度
        if len(self.history) > 50:
            self.history.pop(0)
            self.history_index -= 1
        
        # 自动保存配置
        self.save_config()
    
    def save_config(self):
        config_file = self.version_configs[self.current_version]["config_file"]
        
        # 保存配置
        try:
            self.save_config_to_file(self.config_data, config_file)
            self.original_config = copy.deepcopy(self.config_data)
            self.unsaved_changes = False
            self.status_var.set(f"配置已保存到: {config_file}")
        except Exception as e:
            messagebox.showerror("错误", f"保存配置失败: {str(e)}")
            self.status_var.set("保存失败")
    
    def save_config_to_file(self, config, file_path):
        with open(file_path, 'w', encoding='utf-8') as f:
            json.dump(config, f, indent=4, ensure_ascii=False)
    
    def undo(self):
        if self.history_index > 0:
            self.history_index -= 1
            self.config_data = copy.deepcopy(self.history[self.history_index])
            self.update_config_ui()
            self.status_var.set("已撤销更改")
            # 自动保存配置
            self.save_config()
    
    def redo(self):
        if self.history_index < len(self.history) - 1:
            self.history_index += 1
            self.config_data = copy.deepcopy(self.history[self.history_index])
            self.update_config_ui()
            self.status_var.set("已恢复更改")
            # 自动保存配置
            self.save_config()
    
    def toggle_key_name_mode(self):
        # 切换键名显示模式
        if self.key_name_mode == "original":
            self.key_name_mode = "simple"
            if self.key_name_button:
                self.key_name_button.config(text="切换到原始键名")
        else:
            self.key_name_mode = "original"
            if self.key_name_button:
                self.key_name_button.config(text="切换到简明键名")
        
        # 更新配置界面
        self.update_config_ui()
        # 显示中文状态
        if self.key_name_mode == "simple":
            self.status_var.set("已切换到简明键名模式")
        else:
            self.status_var.set("已切换到原始键名模式")
    
    def get_key_display_name(self, key):
        # 映射原始键名到中文名称
        key_map = {
            "ppt_backup_path": "PPT备份路径",
            "word_backup_path": "Word备份路径",
            "interval": "轮询间隔(秒)",
            "max_skipping_time": "相同文件最大跳过次数",
            "ppt_backup_enable": "启用PPT备份",
            "word_backup_enable": "启用Word备份",
            "wps_backup_enable": "启用WPS备份",
            "upload_to_123pan_enable": "启用123云盘上传",
            "client_id": "123云盘客户端ID",
            "client_secret": "123云盘客户端密钥",
            "access_token": "123云盘访问令牌",
            "folder_id": "123云盘文件夹ID",
            "upload_to_openlist_enable": "启用OpenList上传",
            "openlist_url": "OpenList服务器地址",
            "openlist_username": "OpenList用户名",
            "openlist_password": "OpenList密码",
            "openlist_target_folder": "OpenList目标文件夹",
            "accurate_backup_enable": "启用精确备份",
            "accurate_backup_source_path": "精确备份源路径",
            "accurate_backup_target_path": "精确备份目标路径",
            "show_console_window_at_startup": "启动时显示控制台",
            "save_log": "保存日志"
        }
        
        if self.key_name_mode == "simple" and key in key_map:
            return key_map[key]
        else:
            return key
    
    def start_program(self):
        # 启动对应版本的程序
        version = self.version_var.get()
        if not version:
            messagebox.showerror("错误", "请先选择版本")
            return
        
        # 根据版本确定程序文件
        program_files = {
            "5.0": "OfficebackupSingle5.0.py",
            "5.1": "Officebackup5.1.py",
            "5.1Core": "Officebackup5.1Core.py"
        }
        
        if version in program_files:
            program_file = program_files[version]
            if os.path.exists(program_file):
                # 启动程序
                try:
                    import subprocess
                    subprocess.Popen(["python", program_file])
                    self.status_var.set(f"已启动{version}版本程序")
                    messagebox.showinfo("成功", f"已启动{version}版本程序")
                except Exception as e:
                    messagebox.showerror("错误", f"启动程序失败: {str(e)}")
                    self.status_var.set("启动程序失败")
            else:
                messagebox.showerror("错误", f"程序文件不存在: {program_file}")
                self.status_var.set("程序文件不存在")
        else:
            messagebox.showerror("错误", "无效的版本")
            self.status_var.set("无效的版本")
    
    def on_exit(self):
        self.root.destroy()

if __name__ == "__main__":
    root = tk.Tk()
    
    try:
        root.iconbitmap(default='PythonLight.ico')
    except:
        pass
    
    app = ConfigEditor(root)
    root.mainloop()