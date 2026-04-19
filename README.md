**一个用来备份当前打开的PPT和Word文档并上传网盘的Python小程序**<br>
**A Python-based mini program for backing up currently opened PPT and Word documents and uploading them to cloud drives**<br>
<br>
源代码以MIT协议公开，全部有注释，欢迎研究与改进<br>
The source code is open to the public under the MIT license and is fully annotated. Study and modifications are welcome.<br>
<br>
<br>
<br>
在命令行中执行以下命令安装依赖：
```
pip install -U Pillow pystray pywin32 alist3==1.3.1 svglib rlpycairo
```

或使用requirements.txt安装所有依赖：
```
pip install -r requirements.txt
```
**稳定版本：5.0（推荐，最新，3软件备份功能合一，可对接网盘）、4.2（推荐，经过长期测试 运行稳定，shutil.copy2为主，SaveAs备用）或2.0（SaveAs，有实时备份修改内容的需求可尝试使用，但未经过测试）**<br>
pptbackup的WPS版本**只支持WPS专业版**，**官网最新个人版不可用**，WPS2019教育考试专用版经实测可用，可至 https://hellowindows.cn/ 中的Office/WPS分区下载，来自此网站的WPS2019教育考试专用版123云盘链接：https://www.123pan.com/s/ZrzA-2UZgh<br>
wordbackup**不支持WPS**，后续也没有支持计划（会抛出attribute error异常）<br>
<br>
**5.0更新内容：** <br>
**·新增自动上传至网盘功能（使用第三方pan123库），目前支持123云盘**，可以自行修改相关源码以适配其他网盘<br>
·将PPT、Word、WPS三个软件的备份功能整合到单个程序内，~~引入内建Threading库多线程功能，三个软件的备份、网盘上传和托盘图标分布在5个独立子线程中；将死循环和异常处理都移动到了各线程内部，保证持续运行；三个备份功能与上传功能之间均有1秒钟间隔（注：不能通过在控制台Ctrl+C结束程序，因为KeyboardInterrupt只能结束主线程，而所有功能都分布在不同的子线程中）~~（目前多线程版本有Bug，在主线程结束后子线程内无法捕获到PPT与Word实例，故没有添加至Release，可至仓库内OfficebackupMulti5.0 (Not in use).py查看；Release中的OfficebackupSingle5.0则是将所有函数放到死循环中顺序执行）<br>
**·新增通过json文件修改配置功能（使用内建json库）**，支持自动创建与补全，替代了原先的硬编码方式；配置文件包括：备份路径、轮询间隔、功能开关、网盘API对接、托盘图标左键行为设置）<br>
**·默认隐藏控制台窗口（使用内建ctypes库）**，不再需要vbs脚本隐藏，可以在配置文件内修改默认行为（仅Win7、Win10可隐藏命令行窗口，Win11只能最小化到任务栏）<br>
**·新增托盘图标（使用内建pystray库）**，右键图标能弹出菜单，可以隐藏/显示控制台窗口或退出程序<br>
**·新增日志存储功能**，会在程序运行目录下创建latest.log（是纯文本文件，可用任意文本编辑器打开，如记事本），所有在命令行中print的内容将会通过log_file.write()方法自动写入文件保存，**每次启动程序都会自动删除同目录下已有的latest.log文件，如需保存请在程序重新运行前将旧的日志文件重命名或移动到其他目录**；同时**改进了日志输出逻辑**，将所有日志信息传入到新定义的log_print()函数中，**统一进行时间戳追加、写入.log文件等操作**<br>
<br>
<br>
<br>
<br>
<br>
**5.0版本使用方法：**<br>
1.下载Release中的OfficebackupSingle5.0.exe与PythonLight.ico，放在同一目录下（若不下载图标，托盘图标将设置为一个白色方块；此exe程序使用nuitka打包，支持Windows7及以上系统）<br>
2.下载OfficebackupSingleConfig.json（或运行一次程序，此配置文件将自动生成），按需修改其中的值：<br>
```
{
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
    "upload_to_123pan_enable": True,   #上传到123云盘功能

    #123云盘参数
    "client_id": "",   #123云盘API用户ID
    "client_secret": "",   #123云盘API用户密钥
    "access_token": "",   #123云盘访问令牌，程序会自动获取并写入
    "folder_id": 0,   #目标文件夹ID，可以从浏览器地址栏获取，根目录用0表示

    #文件夹精确备份功能
    "accurate_backup_enable": False,
    "accurate_backup_source_path": "",
    "accurate_backup_target_path": "",

    #托盘图标、控制台行为与日志保存设置
    #"tray_left_click_behavior": "open_console",   #托盘图标左键点击行为，选项有"open_console"（打开控制台）和"exit_program"（退出程序）（无法生效）
    "show_console_window_at startup": False,   #程序启动时显示控制台窗口，True为显示，False为隐藏（默认）
    "save_log": True   #是否保存日志到latest.log文件，True为保存（默认），False为不保存
}
```
3.（可选）把源码放在某个隐蔽的角落，右键创建一个快捷方式，按win+r打开运行框，输入shell:startup，这个文件夹是Windows启动项的文件夹，把快捷方式丢进去就可以实现开机自动以最小化窗口运行**（不能直接将程序放在此文件夹中，否则不生效）**<br>
4.（可选，不建议）如果想要完全静默运行，不弹出托盘图标，新建一个文本文档，输入以下代码，其中引号部分要换成Python文件的路径（Windows10/11可以直接右键，点复制文件路径）：<br>
```
Set ws = CreateObject("Wscript.Shell")
ws.run "OfficebackupSingle5.0.exe",vbhide
```
保存退出，将文件后缀名由.txt改为.vbs，弹出提示框确认更改；最后与第3步一样，为VBS文件创建快捷方式，放到shell:startup文件夹内<br>
此种方式若需要停止程序运行，只能到任务管理器中结束所有Python相关进程，不方便监控程序运行情况，请酌情使用<br>
<br>
**123云盘API对接方式：**<br>
1.如果没有123云盘账号，访问https://www.123pan.cn/ 或 https://www.123pan.com/ ，使用手机号注册123云盘账号<br>
2.登录云盘，点击页面右上角头像，进入个人中心<br>
3.复制个人账号ID，备用<br>
4.访问https://www.123pan.cn/developer （或通过主界面右下角 工具中心-开发者工具-开放平台 进入），勾选并点击“我已阅读并同意，下一步”<br>
5.填写相关信息：<br>
123云盘账号（手机号）：刚刚注册/登录用的手机号<br>
**云盘UID：刚刚复制的个人账号ID**<br>
邮箱：稍后API密钥将发送到这个邮箱<br>
**对接123云盘OpenAPI的应用场景简述：我的产品是一个自动备份文档的小程序，帮助用户定期备份文档，防止数据丢失。希望对接123云盘实现文档自动上传服务<br>
目标用户简述：有文档备份需求的用户**<br>
对接的产品或服务是否已上线：是/否均可<br>
6.审核通过后（基本上是1分钟内秒过，不用等待很久），检查填写的邮箱，API密钥应该会发送到邮箱中<br>
<br>
<br>
<br>
**操作过程图片：**<br>
<img width="563" height="379" alt="image" src="https://github.com/user-attachments/assets/15d63508-af87-4bd6-bf82-49d9074fd55e" />
<img width="589" height="791" alt="image" src="https://github.com/user-attachments/assets/a2a80e9f-c111-45e1-84ca-196197462875" />
<img width="1239" height="963" alt="image" src="https://github.com/user-attachments/assets/38a9449c-55ab-4aea-9fe9-9c372912c5ac" />
<img width="730" height="1295" alt="image" src="https://github.com/user-attachments/assets/26fd34ac-dacf-4fff-890e-ed5d00a07ea8" />
<img width="2125" height="601" alt="image" src="https://github.com/user-attachments/assets/b02263c2-5f5a-4ae4-b9bc-9da9e9cdf6c9" />
<br>
<br>
<br>
<br>
<br>
**​4.2版本使用方法：**<br>
A：直接下载exe文件，无需安装Python环境或运行库即可运行，但只能使用默认的备份路径（"C:\pptbackup"或"C:\wordbackup"）和轮询周期（180秒），无法修改，需要修改路径请使用B方法<br>
B：1.安装好Python环境以及pywin32库（在cmd里运行pip install pywin32）<br>
2.自行修改倒数第7行save_folder变量指定的保存路径（前面的r表示绝对路径），默认在C盘根目录下创建pptbackup文件夹；倒数第9行的time.sleep()里可以自行调节轮询的时间周期，单位是秒，默认为3分钟一次（180秒）<br>
3.（可选）把源码放在某个隐蔽的角落，右键创建一个快捷方式（可以右键打开属性，设置为打开时最小化窗口），按win+r打开运行框，输入shell:startup，这个文件夹是Windows启动项的文件夹，把快捷方式丢进去就可以实现开机自动以最小化窗口运行**（不能直接将程序放在此文件夹中，否则不生效）**
4.（可选，不建议）如果想要完全静默运行，不弹出运行框，先将Python文件名中的空格删除（例如将pptbackup 4.0.py改为pptbackup4.0.py，否则无法索引到），新建一个文本文档，输入以下代码，其中引号部分要换成Python文件的路径（Windows10/11可以直接右键，点复制文件路径）：
```
Set ws = CreateObject("Wscript.Shell")
ws.run "C:\pptbackup4.0.py",vbhide
```
保存退出，将文件后缀名由.txt改为.vbs，弹出提示框确认更改；最后与第3步一样，为VBS文件创建快捷方式，放到shell:startup文件夹内<br>
此种方式若需要停止程序运行，只能到任务管理器中结束所有Python相关进程，不方便监控程序运行情况，请酌情使用<br>
<br>
<br>
<br>
<br>
<br>
做这个程序是因为之前有这样的烦恼：老师一下课就把PPT关掉，U盘拔走，上课没记完的笔记就没法再补了，以后老师也不一定会把PPT公开。所以就想写一个程序，能自动将打开的PPT存到本地，最好是后台静默运行而不是通过模拟点击来实现，不然会~~被老师发现~~打扰老师上课，于是这个程序就应运而生了。它能自动识别打开的PPT或Word文档，并将其保存至指定路径<br>
<br>
**使用此程序前请注意，老师的课件都有知识产权，倾注了每个老师的心血，在自己的班级内与同学分享是可以的，但是请不要外传，因使用不当而导致的后果请自行承担**<br>
<br>
3.0版及以上程序的基本原理：使用win32com.client库，通过COM接口与Office应用程序交互，获取文件路径，再使用shutil库进行文件的复制操作（所以3.0版及以上程序只能备份已保存的文件，正在编辑中的部分无法保存，2.0版程序使用的是SaveAs方法，也许可以进行实时备份，请自行测试）<br>
<br>
<br>
<br>
<br>
<br>
**完整更新日志（与Release描述相同）** <br>
<br>
**2024-12-28（1.0发布）** <br>
<br>
**2025-01-29（建立仓库，2.0发布）** <br>
·修复了Windows7系统上由于没有活动的ppt窗口而出现pywintypes.com_error报错后直接结束程序运行的问题；通过try…except语句和while死循环确保正常运行<br>
·修复了在Office 2016上封装PPT为PDF文件时单独弹出进度条窗口打断全屏放映的问题；现在取消了封装步骤，会直接将PPT以原格式（即.ppt或.pptx）保存，不会打断全屏放映<br>
<br>
**2025-04-16（3.0发布）** <br>
·应同学要求增加了Word文档备份程序，仓库相应进行更名<br>
·代码重构：由于PPT和Word的COM实例行为有差异（新增PowerPoint独立实例默认附加到现有实例，而 Word 的独立实例则无法获取实际打开的文档），而且执行SaveAs之后文档打开路径会自动变更为备份路径，导致文档内超链接失效等问题。为了双版本代码逻辑统一，现在**使用shutil库直接把需备份复制到指定目录**，代替了之前版本的SaveAs方法（所以3.0版程序只能备份已保存的文件，正在编辑中的部分无法保存，如需实时备份可尝试2.0版本）<br>
·所有代码后都添加了注释，方便理解与修改<br>
<br>
**2025-05-01（4.0发布）** <br>
**应该是最后一次大更新，重要修改点已经加粗**<br>
·加入了运行次数计数器，在时间戳后显示<br>
**·加入了对同名文件的检查，现在只会复制第一次，后续检查到同名文件存在会直接跳过复制操作**<br>
·加入了计时器，能显示复制操作的用时<br>
**·加入了对文件修改时间、访问时间的修改，便于排序和查看**（大部分文件系统都使用修改时间来排序，而不是创建时间）：没有此操作时，文件的创建时间和访问时间是备份操作执行的时间，而修改时间是最后一次文档被保存的时间；此操作将修改时间赋给访问时间，~~用于看老师的课件是什么时候做的~~，再将访问时间赋给修改时间，用于排序<br>
·增加了Windows10及以上系统在没有可供备份的PPT时的提示，并标注为正常轮询（Normal Request），之前的版本是静默运行；相应的，由于pywintypes.com_error报错而输出的提示则标注为异常（Exception），并打印出具体异常类型<br>
**·修复了在PPT/Word实例仍处于打开状态时，存储在U盘等移动存储介质上的文件 在U盘被直接物理移除后 进行复制操作 而导致FileNotFoundError异常 不断报错的问题，在第二层（调用层）循环中用except语句捕获此异常，再使用2.0版本中的SaveAs方法进行备份，同时还适用于处理 在程序运行过程中 目标文件夹被删除 而导致的报错**（使用此种方法备份后的输出为"Detected access control, activated SaveAs method, successfully backuped……"，即”检测到访问控制，激活SaveAs方法，成功备份了……“）<br>
<br>
**2025-05-06（4.0修改）** <br>
·pptbackup增加了对WPS专业版的支持<br>
·将三个版本使用Pyinstaller封装成了.exe可执行文件<br>
关于这两条更新的详情见文档开头<br>
<br>
**2025-05-11（4.1发布）** <br>
**第三条更改将会使程序在处理FileNotFound异常时报错，请勿使用此版本！**<br>
·取消了定义层的死循环，将等待时间放到了调用层，逻辑更合理<br>
·删除了多余变量：runiddisplay和copyusedtimedisplay，使用str(runid)和str(copyusedtime)代替<br>
·捕获进程部分将Dispatch/DispatchEx('')更改为GetObject(Class='')<br>
·添加了skippedtime变量，如果备份文件已经存在且跳过次数少于5次（<=5），跳过此次备份操作，否则继续备份<br>
<br>
**2025-06-01（4.2预发布）** <br>
**2025-06-21（4.2发布）** <br>
**·导入collections库的defaultdict方法，用于跟踪单个文件的跳过次数，防止变量污染；为每个需要备份文件都生成file_skip_count和SaveAs_method_activated两个变量，替代了原来的全局skippedtime变量** <br>
·捕获进程部分的更改还原，PPT/Word仍然将使用Dispatch/DispatchEx('')，WPS继续使用GetObject(Class='')<br>
·为创建备份文件夹和开始备份操作增加了输出<br>
<br>
**2025-09（5.0开始开发）** <br>
**2025-10-25（5.0虚拟机测试通过）** <br>
**·新增自动上传至网盘功能（使用第三方pan123库），目前支持123云盘**，可以自行修改相关源码以适配其他网盘<br>
·将PPT、Word、WPS三个软件的备份功能整合到单个程序内，~~引入内建Threading库多线程功能，三个软件的备份、网盘上传和托盘图标分布在5个独立子线程中；将死循环和异常处理都移动到了各线程内部，保证持续运行；三个备份功能与上传功能之间均有1秒钟间隔（注：不能通过在控制台Ctrl+C结束程序，因为KeyboardInterrupt只能结束主线程，而所有功能都分布在不同的子线程中）~~（目前多线程版本有Bug，在主线程结束后子线程内无法捕获到PPT与Word实例，故没有添加至Release，可至仓库内OfficebackupMulti5.0 (Not in use).py查看；Release中的OfficebackupSingle5.0则是将所有函数放到死循环中顺序执行）<br>
**·新增通过json文件修改配置功能（使用内建json库）**，支持自动创建与补全，替代了原先的硬编码方式；配置文件包括：备份路径、轮询间隔、功能开关、网盘API对接、托盘图标左键行为设置）<br>
**·默认隐藏控制台窗口（使用内建ctypes库）**，不再需要vbs脚本隐藏，可以在配置文件内修改默认行为（仅Win7、Win10可隐藏命令行窗口，Win11只能最小化到任务栏）<br>
**·新增托盘图标（使用内建pystray库）**，右键图标能弹出菜单，可以隐藏/显示控制台窗口或退出程序<br>
**·新增日志存储功能**，会在程序运行目录下创建latest.log（是纯文本文件，可用任意文本编辑器打开，如记事本），所有在命令行中print的内容将会通过log_file.write()方法自动写入文件保存，**每次启动程序都会自动删除同目录下已有的latest.log文件，如需保存请在程序重新运行前将旧的日志文件重命名或移动到其他目录**；同时**改进了日志输出逻辑**，将所有日志信息传入到新定义的log_print()函数中，**统一进行时间戳追加、写入.log文件等操作**<br>
**2025-10-31晚（5.0实装测试后改进）** <br>
**2025-11-01凌晨（5.0发布）** <br>
·修复了开机自启后程序无法第一时间联网获取云盘token、标记token_aquired=False时出现的逻辑问题，**确保在任何情况下token_aquired和acccess_token变量都有定义**，避免在上传函数内第二次获取token时出现变量未定义错误导致程序直接终止
