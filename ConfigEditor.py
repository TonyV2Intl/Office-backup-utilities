import tkinter as tk
from tkinter import ttk, messagebox, filedialog
import json
import os
import copy

class ConfigEditor:
    def __init__(self, root):
        self.root = root
        self.root.title("Office Backup Utility Config Editor")
        self.root.geometry("800x400")
        self.root.resizable(True, True)
        # 设置窗口最小尺寸
        self.root.minsize(800, 400)
        
        # 使用系统默认字体
        self.font = None
        
        # 版本配置信息
        self.version_configs = {
            "5.0": {
                "config_file": "OfficebackupSingleConfig.json",
                "cloud_section": "123云盘参数",
                "cloud_params": ["client_id", "client_secret", "access_token", "folder_id"]
            },
            "6.2": {
                "config_file": "OBU6.2.json",
                "cloud_section": "OpenList参数",
                "cloud_params": ["openlist_url", "openlist_username", "openlist_password", "openlist_target_folder"]
            },
            "6.2Core": {
                "config_file": "OBU6.2Core.json",
                "cloud_section": None,
                "cloud_params": []
            }
        }
        
        # 配置数据
        self.config_data = {}
        self.original_config = {}
        self.history = []  # 用于撤销/恢复操作
        self.history_index = -1
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
        self.status_var.set("请选择版本")
    
    def create_widgets(self):
        # 主框架
        main_frame = ttk.Frame(self.root)
        main_frame.pack(fill=tk.BOTH, expand=True)
        
        # 选项卡控件（放在最顶部）
        self.notebook = ttk.Notebook(main_frame)
        self.notebook.pack(fill=tk.BOTH, expand=True)
        
        # 配置编辑选项卡
        self.config_frame = ttk.Frame(self.notebook)
        self.notebook.add(self.config_frame, text="配置编辑")
        
        # 连通性测试选项卡
        self.test_frame = ttk.Frame(self.notebook)
        self.notebook.add(self.test_frame, text="连通性测试")
        
        # 顶部框架（版本选择和按钮）
        top_frame = ttk.Frame(self.config_frame, padding="10")
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
            values=["5.0", "6.2", "6.2Core"],
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
        ttk.Button(button_frame, text="一键启动", command=self.start_program, takefocus=False, width=10).pack(side=tk.RIGHT, padx=5, fill=tk.Y)
        # 键名显示模式切换按钮
        initial_text = "切换到原始键名" if self.key_name_mode == "simple" else "切换到简明键名"
        self.key_name_button = ttk.Button(button_frame, text=initial_text, command=self.toggle_key_name_mode, takefocus=False)
        self.key_name_button.pack(side=tk.RIGHT, padx=5, fill=tk.Y)
        ttk.Button(button_frame, text="重做(下一步)", command=self.redo, takefocus=False).pack(side=tk.RIGHT, padx=5, fill=tk.Y)
        ttk.Button(button_frame, text="撤销(上一步)", command=self.undo, takefocus=False).pack(side=tk.RIGHT, padx=5, fill=tk.Y)
        ttk.Button(button_frame, text="恢复默认配置", command=self.reset_to_default_config, takefocus=False).pack(side=tk.RIGHT, padx=5, fill=tk.Y)
        ttk.Button(button_frame, text="刷新", command=self.refresh_config, takefocus=False).pack(side=tk.RIGHT, padx=5, fill=tk.Y)
        
        # 创建连通性测试界面
        self.create_test_widgets()
        
        # 配置编辑区域
        
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
            #print(f"Mouse wheel delta: {event.delta}, scroll amount: {scroll_amount}")
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
                #print(f"Clicked on {event.widget}, removing focus")
        
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
    
    def create_test_widgets(self):
        """创建连通性测试界面"""
        # 主测试框架
        test_main_frame = ttk.Frame(self.test_frame, padding="10")
        test_main_frame.pack(fill=tk.BOTH, expand=True)
        
        # COM 接口测试板块
        com_test_frame = ttk.LabelFrame(test_main_frame, text="COM 接口测试", padding="10")
        com_test_frame.pack(fill=tk.X, padx=5, pady=5)
        
        # COM 测试按钮
        ttk.Button(com_test_frame, text="测试 COM 接口", command=self.test_com_interfaces, takefocus=False).pack(side=tk.LEFT, padx=5, pady=5)
        
        # COM 测试结果显示区域
        self.com_test_results_frame = ttk.Frame(com_test_frame)
        self.com_test_results_frame.pack(fill=tk.BOTH, expand=True, padx=5, pady=5)
        
        # 初始显示等待测试状态
        self.ppt_label = ttk.Label(self.com_test_results_frame, text="PowerPoint: 等待测试")
        self.ppt_label.pack(anchor=tk.W, padx=5, pady=2)
        
        self.word_label = ttk.Label(self.com_test_results_frame, text="Word: 等待测试")
        self.word_label.pack(anchor=tk.W, padx=5, pady=2)
        
        self.wps_label = ttk.Label(self.com_test_results_frame, text="WPS: 等待测试")
        self.wps_label.pack(anchor=tk.W, padx=5, pady=2)
        
        # 连通性测试板块
        conn_test_frame = ttk.LabelFrame(test_main_frame, text="网络存储连通性测试", padding="10")
        conn_test_frame.pack(fill=tk.X, padx=5, pady=5)
        
        # 左侧控件容器
        left_frame = ttk.Frame(conn_test_frame)
        left_frame.pack(side=tk.LEFT, padx=5, pady=5, fill=tk.Y)
        
        # 版本选择
        version_frame = ttk.Frame(left_frame)
        version_frame.pack(pady=2, fill=tk.X)
        ttk.Label(version_frame, text="版本选择:", font=self.font).pack(side=tk.LEFT, padx=5, fill=tk.Y)
        
        self.test_version_var = tk.StringVar(value="")
        test_version_combobox = ttk.Combobox(
            version_frame, 
            textvariable=self.test_version_var, 
            values=["5.0", "6.2"],
            state="readonly",
            width=10,
            takefocus=False
        )
        test_version_combobox.pack(side=tk.LEFT, padx=5, fill=tk.Y)
        
        # 连通性测试按钮
        ttk.Button(left_frame, text="测试云盘连通性", command=self.test_cloud_connection, takefocus=False).pack(pady=5, fill=tk.X)
        
        # 连通性测试结果显示区域
        self.conn_test_results_frame = ttk.Frame(conn_test_frame)
        self.conn_test_results_frame.pack(fill=tk.BOTH, expand=True, padx=5, pady=5)
    
    def test_com_interfaces(self):
        """测试 COM 接口"""
        # 记录开始时间
        import time
        start_time = time.time()
        
        # 清空结果显示
        for widget in self.com_test_results_frame.winfo_children():
            widget.destroy()
        
        # 测试 PowerPoint
        try:
            import win32com.client
            powerpoint = win32com.client.Dispatch("PowerPoint.Application")
            if powerpoint.Visible:
                presentations = powerpoint.Presentations
                if presentations.Count > 0:
                    # 找到文件，显示绿色
                    ppt_frame = ttk.LabelFrame(self.com_test_results_frame, text="PowerPoint: 已打开")
                    ppt_frame.pack(fill=tk.X, padx=5, pady=2)
                    for i in range(1, presentations.Count + 1):
                        presentation = presentations(i)
                        ttk.Label(ppt_frame, text=f"  - {presentation.Name}", foreground="green").pack(anchor=tk.W, padx=10, pady=1)
                else:
                    # 已打开但无文件，显示绿色
                    ttk.Label(self.com_test_results_frame, text="PowerPoint: 已打开，但无文件", foreground="green").pack(anchor=tk.W, padx=5, pady=2)
            else:
                # 未打开，显示红色
                ttk.Label(self.com_test_results_frame, text="PowerPoint: 未打开", foreground="red").pack(anchor=tk.W, padx=5, pady=2)
        except Exception as e:
            # 错误，显示红色
            ttk.Label(self.com_test_results_frame, text=f"PowerPoint: 错误 - {str(e)}", foreground="red").pack(anchor=tk.W, padx=5, pady=2)
        
        # 测试 Word
        try:
            import win32com.client
            word = win32com.client.Dispatch("Word.Application")
            if word.Visible:
                documents = word.Documents
                if documents.Count > 0:
                    # 找到文件，显示绿色
                    word_frame = ttk.LabelFrame(self.com_test_results_frame, text="Word: 已打开")
                    word_frame.pack(fill=tk.X, padx=5, pady=2)
                    for i in range(1, documents.Count + 1):
                        document = documents(i)
                        ttk.Label(word_frame, text=f"  - {document.Name}", foreground="green").pack(anchor=tk.W, padx=10, pady=1)
                else:
                    # 已打开但无文件，显示绿色
                    ttk.Label(self.com_test_results_frame, text="Word: 已打开，但无文件", foreground="green").pack(anchor=tk.W, padx=5, pady=2)
            else:
                # 未打开，显示红色
                ttk.Label(self.com_test_results_frame, text="Word: 未打开", foreground="red").pack(anchor=tk.W, padx=5, pady=2)
        except Exception as e:
            # 错误，显示红色
            ttk.Label(self.com_test_results_frame, text=f"Word: 错误 - {str(e)}", foreground="red").pack(anchor=tk.W, padx=5, pady=2)
        
        # 测试 WPS
        try:
            import win32com.client
            wps = win32com.client.GetObject(Class="KWPP.Application")
            presentations = wps.Presentations
            if presentations.Count > 0:
                # 找到文件，显示绿色
                wps_frame = ttk.LabelFrame(self.com_test_results_frame, text="WPS: 已打开")
                wps_frame.pack(fill=tk.X, padx=5, pady=2)
                for i in range(1, presentations.Count + 1):
                    presentation = presentations(i)
                    ttk.Label(wps_frame, text=f"  - {presentation.Name}", foreground="green").pack(anchor=tk.W, padx=10, pady=1)
            else:
                # 已打开但无文件，显示绿色
                ttk.Label(self.com_test_results_frame, text="WPS: 已打开，但无文件", foreground="green").pack(anchor=tk.W, padx=5, pady=2)
        except Exception as e:
            # 未打开或错误，显示红色
            ttk.Label(self.com_test_results_frame, text=f"WPS: 未打开或错误 - {str(e)}", foreground="red").pack(anchor=tk.W, padx=5, pady=2)
        
        # 计算测试用时
        end_time = time.time()
        elapsed_time = end_time - start_time
        # 显示测试用时
        ttk.Label(self.com_test_results_frame, text=f"测试用时: {elapsed_time:.3f} 秒").pack(anchor=tk.W, padx=5, pady=5)
    
    def test_cloud_connection(self):
        """测试云盘连通性"""
        import threading
        import time
        start_time = time.time()
        
        # 清空结果显示
        for widget in self.conn_test_results_frame.winfo_children():
            widget.destroy()
        
        # 先显示正在测试的提示
        ttk.Label(self.conn_test_results_frame, text="正在测试连通性...", foreground="blue").pack(anchor=tk.W, padx=5, pady=2)
        
        # 获取当前选择的版本
        version = self.test_version_var.get()
        if not version:
            def show_no_version():
                for widget in self.conn_test_results_frame.winfo_children():
                    widget.destroy()
                ttk.Label(self.conn_test_results_frame, text="请先选择版本", foreground="red").pack(anchor=tk.W, padx=5, pady=2)
                elapsed_time = time.time() - start_time
                ttk.Label(self.conn_test_results_frame, text=f"测试用时: {elapsed_time:.3f} 秒").pack(anchor=tk.W, padx=5, pady=5)
            self.root.after(0, show_no_version)
            return
        
        # 加载配置文件
        try:
            config_file = self.version_configs[version]["config_file"]
            if os.path.exists(config_file):
                with open(config_file, 'r', encoding='utf-8') as f:
                    config = json.load(f)
            else:
                def show_no_config():
                    for widget in self.conn_test_results_frame.winfo_children():
                        widget.destroy()
                    ttk.Label(self.conn_test_results_frame, text=f"配置文件 {config_file} 不存在", foreground="red").pack(anchor=tk.W, padx=5, pady=2)
                    elapsed_time = time.time() - start_time
                    ttk.Label(self.conn_test_results_frame, text=f"测试用时: {elapsed_time:.3f} 秒").pack(anchor=tk.W, padx=5, pady=5)
                self.root.after(0, show_no_config)
                return
        except Exception as e:
            def show_load_error():
                for widget in self.conn_test_results_frame.winfo_children():
                    widget.destroy()
                ttk.Label(self.conn_test_results_frame, text=f"加载配置文件失败: {str(e)}", foreground="red").pack(anchor=tk.W, padx=5, pady=2)
                elapsed_time = time.time() - start_time
                ttk.Label(self.conn_test_results_frame, text=f"测试用时: {elapsed_time:.3f} 秒").pack(anchor=tk.W, padx=5, pady=5)
            self.root.after(0, show_load_error)
            return
        
        # 在后台线程中执行测试
        def run_test_in_thread():
            
            # 测试 123 云盘 (版本 5.0)
            if version == "5.0":
                client_id = config.get("client_id")
                client_secret = config.get("client_secret")
                access_token = config.get("access_token")
                
                if not client_id or not client_secret or not access_token:
                    def show_incomplete_params():
                        for widget in self.conn_test_results_frame.winfo_children():
                            widget.destroy()
                        ttk.Label(self.conn_test_results_frame, text="123云盘参数不完整", foreground="red").pack(anchor=tk.W, padx=5, pady=2)
                        elapsed_time = time.time() - start_time
                        ttk.Label(self.conn_test_results_frame, text=f"测试用时: {elapsed_time:.3f} 秒").pack(anchor=tk.W, padx=5, pady=5)
                    self.root.after(0, show_incomplete_params)
                else:
                    # 测试 123 云盘连通性
                    try:
                        import requests
                        url = "https://open.123pan.com/api/v1/file/list"
                        headers = {"Authorization": f"Bearer {access_token}"}
                        response = requests.get(url, headers=headers, timeout=10)
                        
                        def show_success():
                            for widget in self.conn_test_results_frame.winfo_children():
                                widget.destroy()
                            result_frame = ttk.LabelFrame(self.conn_test_results_frame, text="123云盘: 连通性测试成功")
                            result_frame.pack(fill=tk.X, padx=5, pady=2)
                            ttk.Label(result_frame, text=f"  Client ID: {client_id}").pack(anchor=tk.W, padx=10, pady=1)
                            ttk.Label(result_frame, text=f"  测试状态: 成功").pack(anchor=tk.W, padx=10, pady=1)
                            ttk.Label(result_frame, text=f"  响应状态码: {response.status_code}").pack(anchor=tk.W, padx=10, pady=1)
                            elapsed_time = time.time() - start_time
                            ttk.Label(self.conn_test_results_frame, text=f"测试用时: {elapsed_time:.3f} 秒").pack(anchor=tk.W, padx=5, pady=5)
                        
                        def show_failure():
                            for widget in self.conn_test_results_frame.winfo_children():
                                widget.destroy()
                            ttk.Label(self.conn_test_results_frame, text=f"123云盘: 连通性测试失败 - {response.status_code}", foreground="red").pack(anchor=tk.W, padx=5, pady=2)
                            elapsed_time = time.time() - start_time
                            ttk.Label(self.conn_test_results_frame, text=f"测试用时: {elapsed_time:.3f} 秒").pack(anchor=tk.W, padx=5, pady=5)
                        
                        self.root.after(0, show_success if response.status_code == 200 else show_failure)
                    except Exception as e:
                        def show_error():
                            for widget in self.conn_test_results_frame.winfo_children():
                                widget.destroy()
                            ttk.Label(self.conn_test_results_frame, text=f"123云盘: 连通性测试错误 - {str(e)}", foreground="red").pack(anchor=tk.W, padx=5, pady=2)
                            elapsed_time = time.time() - start_time
                            ttk.Label(self.conn_test_results_frame, text=f"测试用时: {elapsed_time:.3f} 秒").pack(anchor=tk.W, padx=5, pady=5)
                        self.root.after(0, show_error)
            
            # 测试 OpenList (版本 6.2)
            elif version == "6.2":
                openlist_url = config.get("openlist_url")
                if openlist_url:
                    openlist_url = openlist_url.rstrip('/')
                openlist_username = config.get("openlist_username")
                openlist_password = config.get("openlist_password")
                openlist_target_folder = config.get("openlist_target_folder", "/")
                if openlist_target_folder:
                    openlist_target_folder = openlist_target_folder.rstrip('/')
                    if not openlist_target_folder.startswith('/'):
                        openlist_target_folder = '/' + openlist_target_folder
                if not openlist_target_folder or openlist_target_folder == '':
                    openlist_target_folder = '/'
                
                if not openlist_url or not openlist_username:
                    def show_incomplete_params():
                        for widget in self.conn_test_results_frame.winfo_children():
                            widget.destroy()
                        ttk.Label(self.conn_test_results_frame, text="OpenList参数不完整", foreground="red").pack(anchor=tk.W, padx=5, pady=2)
                        elapsed_time = time.time() - start_time
                        ttk.Label(self.conn_test_results_frame, text=f"测试用时: {elapsed_time:.3f} 秒").pack(anchor=tk.W, padx=5, pady=5)
                    self.root.after(0, show_incomplete_params)
                else:
                    # 测试 OpenList 连通性
                    try:
                        from alist import AListAsync, AListUser
                        import asyncio
                        import tempfile
                        import os as os_module
                        
                        # 先显示登录成功信息
                        result_frame = None
                        def show_login_success():
                            nonlocal result_frame
                            for widget in self.conn_test_results_frame.winfo_children():
                                widget.destroy()
                            result_frame = ttk.LabelFrame(self.conn_test_results_frame, text="OpenList: 登录测试成功")
                            result_frame.pack(fill=tk.X, padx=5, pady=2)
                            ttk.Label(result_frame, text=f"  服务器 URL: {openlist_url}").pack(anchor=tk.W, padx=10, pady=1)
                            ttk.Label(result_frame, text=f"  用户名: {openlist_username}").pack(anchor=tk.W, padx=10, pady=1)
                            ttk.Label(result_frame, text=f"  上传测试进行中...", foreground="blue").pack(anchor=tk.W, padx=10, pady=1)
                        
                        # 更新上传测试结果
                        def update_upload_result(result_data):
                            nonlocal result_frame
                            if not result_frame:
                                return
                            
                            # 移除之前的"上传测试进行中..."标签
                            for widget in result_frame.winfo_children():
                                if widget.cget("text") == "  上传测试进行中...":
                                    widget.destroy()
                                    break
                            
                            # 添加新的上传测试结果
                            if result_data['upload_success']:
                                if result_data['delete_success']:
                                    ttk.Label(result_frame, text=f"  上传/删除测试: 成功", foreground="green").pack(anchor=tk.W, padx=10, pady=1)
                                else:
                                    ttk.Label(result_frame, text=f"  上传测试: 成功 (删除失败: {result_data['delete_error']})", foreground="orange").pack(anchor=tk.W, padx=10, pady=1)
                            else:
                                if result_data['upload_error']:
                                    ttk.Label(result_frame, text=f"  上传测试: 错误 - {result_data['upload_error']}", foreground="red").pack(anchor=tk.W, padx=10, pady=1)
                                else:
                                    ttk.Label(result_frame, text=f"  上传测试: 失败", foreground="red").pack(anchor=tk.W, padx=10, pady=1)
                            
                            # 添加测试用时
                            elapsed_time = time.time() - start_time
                            ttk.Label(self.conn_test_results_frame, text=f"测试用时: {elapsed_time:.3f} 秒").pack(anchor=tk.W, padx=5, pady=5)
                        
                        # 异步测试函数
                        async def test_upload_and_delete(client):
                            result_data = {
                                'openlist_url': openlist_url,
                                'username': openlist_username,
                                'upload_success': False,
                                'delete_success': False,
                                'upload_error': None,
                                'delete_error': None
                            }
                            
                            # 创建测试文件
                            test_file_name = ".obu_connectivity_test.tmp"
                            test_file_path = os_module.path.join(tempfile.gettempdir(), test_file_name)
                            with open(test_file_path, 'w') as f:
                                f.write("Office Backup Utility Connectivity Test File\n")
                                f.write(f"Test Time: {time.strftime('%Y-%m-%d %H:%M:%S')}\n")
                            
                            if openlist_target_folder == '/':
                                target_file_path = '/' + test_file_name
                            else:
                                target_file_path = openlist_target_folder + '/' + test_file_name
                            
                            try:
                                result_data['upload_success'] = await client.upload(target_file_path, test_file_path)
                                if result_data['upload_success']:
                                    try:
                                        await client.remove(target_file_path)
                                        result_data['delete_success'] = True
                                    except Exception as delete_error:
                                        result_data['delete_error'] = str(delete_error)
                            except Exception as upload_error:
                                result_data['upload_error'] = str(upload_error)
                            
                            # 删除本地临时文件
                            try:
                                os_module.remove(test_file_path)
                            except Exception:
                                pass
                            
                            return result_data
                        
                        # 运行异步测试并更新结果
                        def run_async_test():
                            # 先登录
                            loop = asyncio.new_event_loop()
                            asyncio.set_event_loop(loop)
                            try:
                                user = AListUser(openlist_username, openlist_password)
                                client = AListAsync(openlist_url)
                                login_result = loop.run_until_complete(client.login(user))
                                
                                # 显示登录成功信息
                                self.root.after(0, show_login_success)
                                
                                # 运行上传测试
                                result_data = loop.run_until_complete(test_upload_and_delete(client))
                                self.root.after(0, lambda: update_upload_result(result_data))
                            finally:
                                loop.close()
                        
                        threading.Thread(target=run_async_test, daemon=True).start()
                        
                    except Exception as e:
                        def show_openlist_error():
                            import traceback
                            traceback.print_exc()
                            for widget in self.conn_test_results_frame.winfo_children():
                                widget.destroy()
                            ttk.Label(self.conn_test_results_frame, text=f"OpenList: 连通性测试错误 - {str(e)}", foreground="red").pack(anchor=tk.W, padx=5, pady=2)
                            elapsed_time = time.time() - start_time
                            ttk.Label(self.conn_test_results_frame, text=f"测试用时: {elapsed_time:.3f} 秒").pack(anchor=tk.W, padx=5, pady=5)
                        self.root.after(0, show_openlist_error)
        
        # 启动后台线程
        thread = threading.Thread(target=run_test_in_thread, daemon=True)
        thread.start()
    
    def on_version_change(self):
        new_version = self.version_var.get()
        if new_version:
            self.current_version = new_version
            # 显示配置界面
            self.load_config()
            self.status_var.set(f"已加载版本 {new_version} 的配置文件")
        else:
            # 隐藏配置界面
            self.status_var.set("请选择版本")
    
    def refresh_config(self):
        if not hasattr(self, 'current_version') or not self.current_version:
            messagebox.showinfo("提示", "请先选择版本")
            return
        
        self.load_config()
        self.status_var.set(f"已刷新版本 {self.current_version} 的配置文件")
    
    def create_status_bar(self):
        # 状态栏
        self.status_var = tk.StringVar(value="就绪")
        status_bar = ttk.Label(self.root, textvariable=self.status_var, relief=tk.SUNKEN, anchor=tk.W)
        status_bar.pack(fill=tk.X, side=tk.BOTTOM)
    
    def setup_styles(self):
        style = ttk.Style()
        
        # 绿色按钮样式
        style.configure("Green.TButton", foreground="green")
        
        # 红色按钮样式
        style.configure("Red.TButton", foreground="red")
        
        # 选项卡样式
        style.configure("TNotebook.Tab", padding=(10, 1))
    

    
    def load_config(self):
        config_file = self.version_configs[self.current_version]["config_file"]
        
        try:
            config_changed = False
            if os.path.exists(config_file):
                with open(config_file, 'r', encoding='utf-8') as f:
                    self.config_data = json.load(f)
                # 检查是否有缺失的配置项
                default_config = self.get_default_config()
                for key, value in default_config.items():
                    if key not in self.config_data:
                        self.config_data[key] = value
                        config_changed = True
                if config_changed:
                    # 保存补全后的配置
                    self.save_config_to_file(self.config_data, config_file)
                    self.status_var.set(f"配置文件已更新，新增了缺失的配置项: {config_file}")
                else:
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
            
            # 更新配置界面
            self.update_config_ui()
            
        except Exception as e:
            messagebox.showerror("错误", f"加载配置文件失败: {str(e)}")
            self.status_var.set("加载配置失败")
    
    def get_default_config(self):
        # 根据版本返回默认配置
        if self.current_version == "5.0":
            return {
                "ppt_backup_path": "C:\\Officebackup\\pptbackup",
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
        elif self.current_version == "6.2":
            return {
                "ppt_backup_path": "C:\\Officebackup\\pptbackup",
                "word_backup_path": "C:\\Officebackup\\wordbackup",
                "interval": 60,
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
                "hide_tray_icon": False,
                "show_console_window_at_startup": False,
                "save_log": True,
                "archive_previous_log": True,
                "log_abnormal_upload": False,
                "backup_timeout": 600,
                "upload_retry_wait": 30,
                "upload_max_retries": ""
            }
        elif self.current_version == "6.2Core":
            return {
                "ppt_backup_path": "C:\\Officebackup\\pptbackup",
                "word_backup_path": "C:\\Officebackup\\wordbackup",
                "interval": 60,
                "ppt_backup_enable": True,
                "word_backup_enable": True,
                "wps_backup_enable": True,
                "accurate_backup_enable": False,
                "accurate_backup_source_path": "",
                "accurate_backup_target_path": "",
                "show_console_window_at_startup": True,
                "save_log": True,
                "archive_previous_log": True,
                "backup_timeout": 600
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
            elif self.current_version == "6.2":
                sections["功能开关"].append("upload_to_openlist_enable")
            sections[cloud_section] = cloud_params
        
        # 添加其他配置
        sections["精确备份"] = ["accurate_backup_enable", "accurate_backup_source_path", "accurate_backup_target_path"]
        
        # 添加控制台和日志设置
        if self.current_version == "6.2":
            sections["界面与日志"] = ["hide_tray_icon", "show_console_window_at_startup", "save_log", "archive_previous_log", "log_abnormal_upload"]
        else:
            sections["控制台与日志"] = ["show_console_window_at_startup", "save_log", "archive_previous_log"]
        
        # 添加超时和重试设置
        if self.current_version == "6.2":
            sections["超时与重试"] = ["backup_timeout", "upload_retry_wait", "upload_max_retries"]
        else:
            sections["超时设置"] = ["backup_timeout"]
        
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
                if key == "accurate_backup_target_path":
                    ttk.Button(entry_frame, text="自动补全", command=lambda v=var: self.auto_complete_target_path(v), takefocus=False).pack(side=tk.LEFT, padx=2)
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
    
    def auto_complete_target_path(self, var):
        source_path = self.config_data.get("accurate_backup_source_path", "")
        if not source_path:
            messagebox.showinfo("提示", "请先设置精确备份源路径")
            return
        
        source_path = os.path.normpath(source_path)
        folder_name = os.path.basename(source_path)
        if not folder_name:
            messagebox.showinfo("提示", "无法从源路径提取文件夹名称")
            return
        
        current_target = var.get()
        if not current_target:
            messagebox.showinfo("提示", "请先设置目标路径")
            return
        
        current_target = os.path.normpath(current_target)
        new_target = os.path.join(current_target, folder_name + "-backup")
        new_target = new_target.replace('/', '\\')
        var.set(new_target)
    
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

    def reset_to_default_config(self):
        # 确认操作
        if messagebox.askyesno("确认", "确定要恢复默认配置吗？当前配置将被覆盖。"):
            # 先把当前配置添加到历史记录，以便撤销
            self.history = self.history[:self.history_index + 1]
            self.history.append(copy.deepcopy(self.config_data))
            self.history_index = len(self.history) - 1
            if len(self.history) > 50:
                self.history.pop(0)
                self.history_index -= 1

            default_config = self.get_default_config()
            self.config_data = copy.deepcopy(default_config)
            self.save_config()
            self.original_config = copy.deepcopy(self.config_data)
            self.history.append(copy.deepcopy(self.config_data))
            self.history_index = len(self.history) - 1
            if len(self.history) > 50:
                self.history.pop(0)
                self.history_index -= 1
            self.update_config_ui()
            self.status_var.set("已恢复默认配置（可撤销）")
    
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
            "hide_tray_icon": "隐藏托盘图标",
            "show_console_window_at_startup": "启动时显示控制台",
            "save_log": "保存日志",
            "archive_previous_log": "归档之前的日志",
            "log_abnormal_upload": "记录上传异常文件",
            "backup_timeout": "备份超时时间(秒)",
            "upload_retry_wait": "上传重试等待(秒)",
            "upload_max_retries": "上传最大重试次数"
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
            "5.0": "OfficebackupSingle5.0",
            "6.2": "Officebackup6.2",
            "6.2Core": "Officebackup6.2Core"
        }
        
        if version in program_files:
            base_name = program_files[version]
            # 优先尝试启动 py 文件
            py_file = f"{base_name}.py"
            if os.path.exists(py_file):
                # 启动 py 文件
                try:
                    import subprocess
                    subprocess.Popen(["python", py_file])
                    self.status_var.set(f"已启动{version}版本程序")
                    messagebox.showinfo("成功", f"已启动{version}版本程序")
                    return
                except Exception as e:
                    messagebox.showerror("错误", f"启动程序失败: {str(e)}")
                    self.status_var.set("启动程序失败")
                    return
            
            # 如果 py 文件不存在，尝试启动 exe 文件
            exe_file = f"{base_name}.exe"
            if os.path.exists(exe_file):
                # 启动 exe 文件
                try:
                    import subprocess
                    subprocess.Popen([exe_file])
                    self.status_var.set(f"已启动{version}版本程序")
                    messagebox.showinfo("成功", f"已启动{version}版本程序")
                except Exception as e:
                    messagebox.showerror("错误", f"启动程序失败: {str(e)}")
                    self.status_var.set("启动程序失败")
            else:
                messagebox.showerror("错误", f"程序文件不存在: {py_file} 或 {exe_file}")
                self.status_var.set("程序文件不存在")
        else:
            messagebox.showerror("错误", "无效的版本")
            self.status_var.set("无效的版本")
    
    def on_exit(self):
        self.root.destroy()

if __name__ == "__main__":
    root = tk.Tk()
    app = ConfigEditor(root)
    root.mainloop()