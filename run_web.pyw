#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
电池健康检测程序 - 静默启动版本（双击运行）
"""

import os
import sys
import subprocess
import webbrowser
import logging
import platform  # 添加 platform 模块导入
from pathlib import Path

# 配置日志
log_file = os.path.join(os.getcwd(), "battery_health.log")
logging.basicConfig(
    level=logging.DEBUG,
    format='%(asctime)s - %(levelname)s - %(message)s',
    handlers=[
        logging.FileHandler(log_file),
        logging.StreamHandler()
    ]
)
logger = logging.getLogger(__name__)

# 检测是否在可执行文件中运行
is_frozen = getattr(sys, 'frozen', False)
logger.info(f"运行环境: {'可执行文件' if is_frozen else 'Python脚本'}")

# 获取基础路径
try:
    if is_frozen:
        # PyInstaller 打包后，文件会被解压到临时目录
        base_path = sys._MEIPASS
        logger.info(f"临时目录: {base_path}")
    else:
        # 正常运行时，使用脚本所在目录
        base_path = str(Path(__file__).parent.absolute())
        logger.info(f"脚本目录: {base_path}")
    
    # 切换到基础目录
    os.chdir(base_path)
    logger.info(f"当前工作目录: {os.getcwd()}")
except Exception as e:
    logger.error(f"获取路径时出错: {e}")
    input("按回车键退出...")
    sys.exit(1)

def check_port(port):
    """检查端口是否被占用"""
    try:
        import socket
        with socket.socket(socket.AF_INET, socket.SOCK_STREAM) as s:
            s.settimeout(1)
            result = s.connect_ex(("localhost", port))
            return result == 0  # 0 表示端口被占用
    except Exception as e:
        logger.error(f"检查端口时出错: {e}")
        return False

def release_port(port):
    """释放被占用的端口"""
    logger.info(f"检测到端口 {port} 被占用，正在尝试释放...")
    
    try:
        # 在Windows上使用netstat查找占用端口的进程
        if os.name == 'nt':
            result = subprocess.run(
                ["netstat", "-ano"],
                capture_output=True,
                text=True
            )
            
            # 解析输出查找占用端口的进程
            for line in result.stdout.splitlines():
                if f":{port}" in line and "LISTENING" in line:
                    # 提取PID
                    parts = line.split()
                    if len(parts) >= 5:
                        pid = parts[-1]
                        logger.info(f"找到占用端口的进程: PID {pid}")
                        
                        # 尝试终止进程
                        try:
                            subprocess.run(
                                ["taskkill", "/PID", pid, "/F"],
                                capture_output=True,
                                text=True
                            )
                            logger.info(f"✓ 成功终止进程 PID {pid}")
                            return True
                        except Exception as e:
                            logger.error(f"⚠ 无法终止进程: {e}")
                            return False
        else:
            logger.warning("⚠ 自动释放端口功能仅在Windows上支持")
            return False
            
    except Exception as e:
        logger.error(f"⚠ 释放端口时出错: {e}")
        return False
    
    logger.warning("⚠ 未找到占用端口的进程")
    return False

def run_battery_checker():
    """运行电池检测程序"""
    if is_frozen:
        # 在可执行文件中运行，直接导入模块
        logger.info("在可执行文件中运行，直接导入电池检测模块...")
        try:
            # 在可执行文件中，battery_health_checker.py 会被解压到临时目录
            # 我们需要动态导入
            import importlib.util
            module_path = os.path.join(base_path, "battery_health_checker.py")
            logger.info(f"电池检测模块路径: {module_path}")
            logger.info(f"模块文件存在: {os.path.exists(module_path)}")
            
            spec = importlib.util.spec_from_file_location("battery_health_checker", module_path)
            battery_health_checker = importlib.util.module_from_spec(spec)
            spec.loader.exec_module(battery_health_checker)
            BatteryHealthChecker = battery_health_checker.BatteryHealthChecker
            
            checker = BatteryHealthChecker()
            checker.run(save_report=True)
            return True
        except Exception as e:
            logger.error(f"错误：{e}")
            import traceback
            traceback.print_exc()
            return False
    else:
        # 在Python脚本中运行，使用subprocess
        logger.info("在Python脚本中运行，使用subprocess...")
        try:
            result = subprocess.run(
                [sys.executable, "battery_health_checker.py"],
                capture_output=False,
                text=True
            )
            return result.returncode == 0
        except Exception as e:
            logger.error(f"错误：{e}")
            import traceback
            traceback.print_exc()
            return False

def main():
    """主函数"""
    logger.info("="*60)
    logger.info("        电池健康检测报告 - Web查看器")
    logger.info("="*60)
    
    # 步骤1：生成报告
    logger.info("[1/2] 正在检测电池信息...")
    
    if not run_battery_checker():
        logger.error("错误：电池检测失败")
        input("按回车键退出...")
        return

    logger.info("✓ 报告生成完成！")

    # 步骤2：启动服务器
    logger.info("[2/2] 正在启动Web服务器...")
    
    try:
        import http.server
        import socketserver
        
        PORT = 8000
        
        # 检查端口是否被占用
        logger.info(f"检查端口 {PORT} 是否被占用...")
        if check_port(PORT):
            # 尝试释放端口
            logger.info(f"端口 {PORT} 被占用，尝试释放...")
            if not release_port(PORT):
                logger.error(f"错误: 端口 {PORT} 已被占用且无法自动释放")
                logger.error("请手动关闭占用该端口的程序后重试")
                input("按回车键退出...")
                return
            
            # 等待片刻让端口完全释放
            import time
            logger.info("等待端口释放...")
            time.sleep(1)
        
        logger.info(f"准备启动服务器在端口 {PORT}...")
        
        # 检查文件是否存在
        index_file = os.path.join(base_path, "index.html")
        logger.info(f"检查 index.html 文件: {os.path.exists(index_file)}")
        if os.path.exists(index_file):
            logger.info(f"index.html 文件大小: {os.path.getsize(index_file)} 字节")
        
        # 检查报告文件
        report_file = os.path.join(base_path, "battery_health_report.json")
        logger.info(f"检查电池报告文件: {os.path.exists(report_file)}")
        if os.path.exists(report_file):
            logger.info(f"报告文件大小: {os.path.getsize(report_file)} 字节")
        
        # 自定义HTTP处理器
        class CustomHandler(http.server.SimpleHTTPRequestHandler):
            def log_message(self, format, *args):
                # 记录HTTP请求
                logger.info(f"HTTP请求: {args[0]} {args[1]}")
        
        logger.info(f"创建TCP服务器...")
        with socketserver.TCPServer(("", PORT), CustomHandler) as httpd:
            logger.info(f"服务器已启动在端口 {PORT}")
            logger.info(f"服务器地址: http://localhost:{PORT}")
            logger.info(f"正在打开浏览器...")
            logger.info(f"请访问: http://localhost:{PORT}/index.html")

            # 打开浏览器
            try:
                webbrowser.open(f'http://localhost:{PORT}/index.html')
                logger.info("✓ 浏览器已打开")
            except Exception as e:
                logger.error(f"⚠ 无法自动打开浏览器: {e}")
                logger.error(f"请手动访问: http://localhost:{PORT}/index.html")

            # 运行服务器
            logger.info("服务器开始运行...")
            httpd.serve_forever()
            
    except OSError as e:
        if e.errno == 10048:
            logger.error(f"错误: 端口 {PORT} 已被占用")
            logger.error("请关闭占用该端口的程序后重试")
        else:
            logger.error(f"错误: {e}")
            import traceback
            traceback.print_exc()
        input("按回车键退出...")
    except KeyboardInterrupt:
        logger.info("服务器已停止")
    except Exception as e:
        logger.error(f"错误: {e}")
        import traceback
        traceback.print_exc()
        input("按回车键退出...")

if __name__ == "__main__":
    main()
