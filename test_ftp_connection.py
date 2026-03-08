#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
FTP Connection Test Script
用于测试FTP服务器的连通性和基本操作
"""

import ftplib
import os
import time
import tempfile

class FixedPASVFTP(ftplib.FTP):
    """
    自定义FTP类，修复PASV响应中的IP地址并验证被动端口范围
    - 替换本地IP为真实公网IP
    - 验证被动端口是否在指定范围内
    """
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
        print(f"DEBUG: PASV response: {response}")
        
        # 手动解析PASV响应
        import re
        match = re.search(r'227.*?\((\d+),(\d+),(\d+),(\d+),(\d+),(\d+)\)', response)
        if not match:
            raise ftplib.error_reply(response)
        
        # 提取原始IP和端口
        h1, h2, h3, h4, p1, p2 = map(int, match.groups())
        original_host = f"{h1}.{h2}.{h3}.{h4}"
        final_port = p1 * 256 + p2
        
        print(f"DEBUG: Original PASV info: host={original_host}, port={final_port}")
        
        # 验证被动端口是否在指定范围内
        if self.passive_port_range:
            if final_port < self.passive_port_range.start or final_port > self.passive_port_range.stop - 1:
                print(f"WARNING: PASV port {final_port} is not in expected range {self.passive_port_range}")
            else:
                print(f"DEBUG: PASV port {final_port} is within expected range {self.passive_port_range}")
        
        # 替换本地IP为真实公网IP
        if original_host in ('127.0.0.1', 'localhost', '0.0.0.0'):
            final_host = self.real_host
            print(f"DEBUG: Replaced PASV IP: {original_host} -> {final_host}")
        else:
            final_host = original_host
            print(f"DEBUG: Using original PASV IP: {final_host}")
        
        return final_host, final_port


def test_ftp_connection(host, port=21, username='', password='', target_path='', timeout=60, active_mode=False, passive_port_range=None):
    """
    测试FTP服务器连接和基本操作
    
    参数:
        host: FTP服务器地址
        port: FTP服务器端口，默认21（同时作为被动端口）
        username: FTP用户名
        password: FTP密码
        target_path: 测试操作的目标路径
        timeout: 连接超时时间，默认30秒
        active_mode: 是否使用主动模式，默认False（被动模式）
        passive_port_range: 被动端口范围，默认None
    
    返回:
        bool: 测试是否成功
    """
    
    print("=" * 60)
    print(f"FTP Connection Test")
    print(f"Server: {host}:{port}")
    print(f"Username: {username}")
    print(f"Target Path: {target_path or 'Root Directory'}")
    print("=" * 60)
    
    ftp = None
    test_file_name = f"ftp_test_{int(time.time())}.txt"
    test_file_content = f"FTP Connection Test File\nCreated: {time.strftime('%Y-%m-%d %H:%M:%S')}\nHost: {host}:{port}\nUsername: {username}"
    
    try:
        # 1. 测试连接
        print("1. Testing connection...")
        # 创建FTP对象，验证被动端口范围
        ftp = FixedPASVFTP(host, passive_port_range=passive_port_range)
        ftp.set_debuglevel(2)  # 设置调试级别，显示详细的FTP命令和响应
        ftp.connect(host, port, timeout=timeout)
        print(f"   ✓ Connected to {host}:{port}")
        
        # 2. 测试登录
        print("2. Testing login...")
        ftp.login(username, password)
        print(f"   ✓ Logged in as {username}")
        
        # 3. 设置传输模式
        print(f"3. Setting transfer mode...")
        if active_mode:
            ftp.set_pasv(False)
            print(f"   ✓ Active mode enabled")
        else:
            ftp.set_pasv(True)
            print(f"   ✓ Passive mode enabled")
        
        # 3. 测试获取欢迎信息
        welcome_msg = ftp.getwelcome()
        print(f"   Server welcome: {welcome_msg.strip()}")
        
        # 4. 测试获取当前目录
        current_dir = ftp.pwd()
        print(f"   Current directory: {current_dir}")
        
        # 5. 测试目录切换
        if target_path:
            print(f"3. Testing directory change to {target_path}...")
            try:
                ftp.cwd(target_path)
                print(f"   ✓ Changed to directory: {target_path}")
            except ftplib.error_perm as e:
                print(f"   ✗ Failed to change directory: {e}")
                print(f"   Attempting to create directory: {target_path}")
                # 创建多级目录
                dirs = target_path.split('/')
                current_dir = ''
                for dir in dirs:
                    if dir:
                        current_dir += '/' + dir
                        try:
                            ftp.cwd(current_dir)
                        except ftplib.error_perm:
                            ftp.mkd(current_dir)
                            ftp.cwd(current_dir)
                print(f"   ✓ Created and changed to directory: {target_path}")
        
        # 6. 测试列出目录内容
        print("4. Testing directory listing...")
        files = []
        ftp.retrlines('NLST', files.append)
        print(f"   ✓ Directory contains {len(files)} items")
        if files:
            print(f"   First 5 items: {files[:5]}")
        
        # 7. 测试上传文件
        print(f"5. Testing file upload: {test_file_name}...")
        with tempfile.NamedTemporaryFile(mode='w', delete=False, suffix='.txt') as f:
            f.write(test_file_content)
            temp_file_path = f.name
        
        try:
            with open(temp_file_path, 'rb') as f:
                # 设置底层套接字超时时间
                ftp.sock.settimeout(30)  # 设置30秒超时
                try:
                    ftp.storbinary(f'STOR {test_file_name}', f)
                    print(f"   ✓ Uploaded file: {test_file_name}")
                except (TimeoutError, ConnectionResetError, ftplib.error_temp) as e:
                    # 捕获超时等临时错误，需要重新连接
                    print(f"   ⚠ File upload timed out, attempting to reconnect...")
                    print(f"   Error: {e}")
                    # 关闭当前连接（对于超时情况，直接关闭socket而不是发送QUIT命令）
                    try:
                        ftp.quit()
                    except Exception as quit_error:
                        print(f"   WARNING: Failed to quit cleanly, closing socket directly: {quit_error}")
                        # 直接关闭socket
                        try:
                            ftp.sock.close()
                        except Exception as close_error:
                            print(f"   WARNING: Failed to close socket: {close_error}")
                    
                    # 重新建立FTP连接
                    print(f"   Reconnecting to FTP server...")
                    ftp = FixedPASVFTP(host, passive_port_range=passive_port_range)
                    ftp.connect(host, port, timeout=timeout)
                    ftp.login(username, password)
                    print(f"   ✓ Reconnected to {host}:{port}")
                    
                    # 重新切换到目标目录
                    if target_path:
                        ftp.cwd(target_path)
                        print(f"   ✓ Changed to directory: {target_path}")
        finally:
            os.unlink(temp_file_path)
        
        # 8. 测试文件大小（添加异常处理）
        print(f"6. Testing file size...")
        file_size = None
        try:
            file_size = ftp.size(test_file_name)
            print(f"   ✓ File size: {file_size} bytes")
        except (ftplib.error_perm, BrokenPipeError, ConnectionResetError, TimeoutError) as e:
            print(f"   ✗ Failed to get file size: {e}")
        
        # 9. 测试下载文件（添加异常处理）
        print(f"7. Testing file download: {test_file_name}...")
        download_path = f"download_{test_file_name}"
        download_success = False
        try:
            with open(download_path, 'wb') as f:
                ftp.retrbinary(f'RETR {test_file_name}', f.write)
            print(f"   ✓ Downloaded file: {download_path}")
            download_success = True
            
            # 验证下载内容
            with open(download_path, 'r') as f:
                downloaded_content = f.read()
            if downloaded_content == test_file_content:
                print(f"   ✓ Downloaded content matches original")
            else:
                print(f"   ✗ Downloaded content mismatch")
        except (ftplib.error_perm, BrokenPipeError, ConnectionResetError, TimeoutError) as e:
            print(f"   ✗ Failed to download file: {e}")
        finally:
            if os.path.exists(download_path):
                os.unlink(download_path)
        
        # 10. 测试删除文件（添加异常处理）
        print(f"8. Testing file deletion: {test_file_name}...")
        try:
            ftp.delete(test_file_name)
            print(f"   ✓ Deleted file: {test_file_name}")
        except (ftplib.error_perm, BrokenPipeError, ConnectionResetError, TimeoutError) as e:
            print(f"   ✗ Failed to delete file: {e}")
        
        # 11. 测试断开连接（添加异常处理）
        print("9. Testing disconnection...")
        try:
            ftp.quit()
            print(f"   ✓ Disconnected successfully")
        except Exception as e:
            print(f"   ✗ Failed to disconnect: {e}")
        
        print("=" * 60)
        print("🎉 All tests passed! FTP connection is working correctly.")
        print("=" * 60)
        return True
        
    except ftplib.all_errors as e:
        print(f"   ✗ FTP Error: {e}")
        print("=" * 60)
        print(f"❌ Test failed! FTP connection issue detected.")
        print("=" * 60)
        return False
    except Exception as e:
        print(f"   ✗ Unexpected Error: {e}")
        print("=" * 60)
        print(f"❌ Test failed! Unexpected error occurred.")
        print("=" * 60)
        return False
    finally:
        if ftp and hasattr(ftp, 'sock') and ftp.sock:
            try:
                ftp.quit()
            except:
                pass


if __name__ == "__main__":
    # 默认测试配置，可以根据需要修改
    # 设置被动端口范围
    PASSIVE_PORT_RANGE = range(55752, 55753)
    
    TEST_CONFIG = {
        'host': 'cn-zj-nb-1.lcf.im',  # FRP穿透后的公网IP
        'port': 10814,  # FRP映射的FTP端口
        'username': 'Tony',
        'password': '090721Version2',
        'target_path': '/Tony/课件备份 高二',  # NAS上的目标路径
        'active_mode': False,  # 使用被动模式，服务器已配置masquerade_address
        'passive_port_range': PASSIVE_PORT_RANGE  # 传递被动端口范围
    }
    
    # 运行测试
    test_ftp_connection(**TEST_CONFIG)
