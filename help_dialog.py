# -*- coding: utf-8 -*-
"""
帮助对话框模块
包含GitHub连接检测功能
"""

import sys
import requests
import socket
import urllib.parse
from datetime import datetime

from PyQt5.QtWidgets import (
    QDialog, QVBoxLayout, QHBoxLayout, QLabel, QPushButton, QTextEdit,
    QGroupBox, QProgressBar, QFrame, QMessageBox
)
from PyQt5.QtCore import Qt, QTimer, pyqtSignal, QThread
from PyQt5.QtGui import QFont, QColor, QPalette


class GitHubConnectionChecker(QThread):
    """GitHub连接检测线程"""
    connection_result = pyqtSignal(dict)  # 连接结果信号
    
    def __init__(self, github_api_url):
        super().__init__()
        self.github_api_url = github_api_url
    
    def _safe_encode_text(self, text):
        """安全编码文本，处理所有编码问题"""
        if text is None:
            return ""
        
        try:
            # 如果是字节类型，尝试解码
            if isinstance(text, bytes):
                # 尝试UTF-8解码
                try:
                    return text.decode('utf-8', errors='ignore')
                except:
                    # 如果UTF-8失败，尝试其他常见编码
                    try:
                        return text.decode('gbk', errors='ignore')
                    except:
                        return text.decode('latin-1', errors='ignore')
            
            # 如果是字符串，确保是UTF-8
            if isinstance(text, str):
                # 尝试重新编码为UTF-8来确保编码正确
                return text.encode('utf-8', errors='ignore').decode('utf-8', errors='ignore')
            
            # 其他类型转换为字符串
            return str(text)
            
        except Exception as e:
            # 如果所有编码都失败，返回安全文本
            return f"编码错误: {type(e).__name__}"
    
    def _network_diagnosis(self):
        """网络诊断，检查基础网络连接"""
        diagnosis = {}
        
        try:
            # 1. 检查DNS解析
            parsed_url = urllib.parse.urlparse(self.github_api_url)
            hostname = parsed_url.hostname
            
            try:
                ip_address = socket.gethostbyname(hostname)
                diagnosis['dns_resolve'] = f"成功 - {hostname} -> {ip_address}"
            except socket.gaierror as e:
                diagnosis['dns_resolve'] = f"失败 - {str(e)}"
            
            # 2. 检查基本网络连接
            try:
                # 尝试连接到GitHub的常用端口
                sock = socket.socket(socket.AF_INET, socket.SOCK_STREAM)
                sock.settimeout(5)
                result = sock.connect_ex(('github.com', 443))
                sock.close()
                
                if result == 0:
                    diagnosis['tcp_connect'] = "成功 - 可以连接到github.com:443"
                else:
                    diagnosis['tcp_connect'] = f"失败 - 错误代码: {result}"
            except Exception as e:
                diagnosis['tcp_connect'] = f"异常 - {str(e)}"
            
        except Exception as e:
            diagnosis['diagnosis_error'] = f"诊断过程出错: {str(e)}"
        
        return diagnosis
    
    def run(self):
        """执行连接检测"""
        result = {
            'timestamp': datetime.now().strftime('%Y-%m-%d %H:%M:%S'),
            'status': 'unknown',
            'response_time': 0,
            'error_message': '',
            'details': {},
            'network_diagnosis': {}
        }
        
        # 先进行网络诊断
        result['network_diagnosis'] = self._network_diagnosis()
        
        # 在最外层捕获所有可能的编码错误
        try:
            # 开始计时
            start_time = datetime.now()
            
            # 在VPN环境下使用更安全的请求设置
            session = requests.Session()
            
            # 添加更安全的请求头（使用英文User-Agent避免编码问题）
            headers = {
                'User-Agent': 'RefundManager/1.0',
                'Accept': 'application/vnd.github.v3+json',
                'Accept-Charset': 'utf-8',
                'Accept-Encoding': 'identity'  # 禁用压缩，避免编码问题
            }
            
            # 发送请求到GitHub API
            response = session.get(
                self.github_api_url,
                timeout=10,  # 10秒超时
                headers=headers,
                verify=True,  # 启用SSL验证
                allow_redirects=True
            )
            
            # 计算响应时间
            end_time = datetime.now()
            response_time = (end_time - start_time).total_seconds() * 1000  # 毫秒
            
            result['response_time'] = response_time
            
            if response.status_code == 200:
                result['status'] = 'connected'
                result['details'] = {
                    'status_code': response.status_code,
                    'content_type': response.headers.get('Content-Type', ''),
                    'rate_limit': response.headers.get('X-RateLimit-Limit', '未知'),
                    'rate_remaining': response.headers.get('X-RateLimit-Remaining', '未知'),
                    'server': response.headers.get('Server', '未知'),
                    'content_length': len(response.content)
                }
            else:
                result['status'] = 'error'
                result['error_message'] = f"HTTP错误: {response.status_code}"
                
                # 详细检查响应内容
                content_details = {
                    'status_code': response.status_code,
                    'content_length': len(response.content),
                    'content_type': response.headers.get('Content-Type', ''),
                    'encoding': response.encoding
                }
                
                # 尝试多种方式获取响应内容
                try:
                    # 方法1: 直接获取文本
                    response_text = response.text[:500]
                    content_details['response_text'] = self._safe_encode_text(response_text)
                except Exception as text_error:
                    content_details['text_error'] = f"文本获取失败: {text_error}"
                
                try:
                    # 方法2: 获取原始字节
                    raw_content = response.content[:200]
                    content_details['raw_bytes'] = str(raw_content)
                except Exception as bytes_error:
                    content_details['bytes_error'] = f"字节获取失败: {bytes_error}"
                
                result['details'] = content_details
                
        except Exception as outer_exception:
            # 在最外层捕获所有异常，包括编码错误
            try:
                # 尝试获取异常类型和基本信息
                exception_type = type(outer_exception).__name__
                
                # 安全地获取异常信息
                try:
                    # 先尝试直接字符串化
                    error_str = str(outer_exception)
                    # 如果包含编码错误，说明是异常对象本身的问题
                    if "latin-1" in error_str or "encode" in error_str:
                        error_msg = "异常对象包含无法编码的字符"
                        result['details'] = {
                            'exception_type': exception_type,
                            'encoding_issue': '异常对象字符串化失败',
                            'raw_exception_repr': repr(outer_exception)[:200]
                        }
                    else:
                        error_msg = self._safe_encode_text(error_str)
                except Exception as encode_error:
                    # 如果字符串化也失败，说明是严重的编码问题
                    error_msg = f"严重编码错误: {type(encode_error).__name__}"
                    result['details'] = {
                        'exception_type': exception_type,
                        'encode_error_type': type(encode_error).__name__,
                        'raw_exception_repr': repr(outer_exception)[:100]
                    }
                
                # 根据异常类型设置状态
                if isinstance(outer_exception, requests.exceptions.Timeout):
                    result['status'] = 'timeout'
                    result['error_message'] = f"连接超时: {error_msg}"
                elif isinstance(outer_exception, requests.exceptions.ConnectionError):
                    result['status'] = 'connection_error'
                    result['error_message'] = f"连接错误: {error_msg}"
                elif isinstance(outer_exception, requests.exceptions.RequestException):
                    result['status'] = 'request_error'
                    result['error_message'] = f"请求错误: {error_msg}"
                else:
                    result['status'] = 'unknown_error'
                    result['error_message'] = f"未知错误: {error_msg}"
                    
            except Exception as final_error:
                # 如果连异常处理都失败，使用最安全的错误信息
                result['status'] = 'critical_error'
                result['error_message'] = "严重错误: 无法处理异常"
                result['details'] = {
                    'final_error': type(final_error).__name__
                }
        
        # 发送结果
        self.connection_result.emit(result)


class HelpDialog(QDialog):
    """帮助对话框"""
    
    def __init__(self, parent=None, github_api_url=""):
        super().__init__(parent)
        self.parent = parent
        self.github_api_url = github_api_url
        self.connection_checker = None
        self.connection_logs = []
        
        self.setup_ui()
        self.setup_connections()
        
        # 设置窗口属性
        self.setWindowTitle("帮助与连接检测")
        self.setMinimumSize(600, 500)
        self.setMaximumSize(800, 600)
        
    def setup_ui(self):
        """设置界面"""
        layout = QVBoxLayout(self)
        
        # 标题
        title_label = QLabel("帮助与连接检测")
        title_font = QFont()
        title_font.setPointSize(14)
        title_font.setBold(True)
        title_label.setFont(title_font)
        title_label.setAlignment(Qt.AlignCenter)
        layout.addWidget(title_label)
        
        # 分隔线
        line = QFrame()
        line.setFrameShape(QFrame.HLine)
        line.setFrameShadow(QFrame.Sunken)
        layout.addWidget(line)
        
        # GitHub连接检测区域
        self.setup_github_section(layout)
        
        # 分隔线
        line2 = QFrame()
        line2.setFrameShape(QFrame.HLine)
        line2.setFrameShadow(QFrame.Sunken)
        layout.addWidget(line2)
        
        # 帮助信息区域
        self.setup_help_section(layout)
        
        # 按钮区域
        self.setup_buttons(layout)
        
    def setup_github_section(self, layout):
        """设置GitHub连接检测区域"""
        github_group = QGroupBox("GitHub连接检测")
        github_layout = QVBoxLayout(github_group)
        
        # 状态显示区域
        status_layout = QHBoxLayout()
        
        # 状态标签
        self.status_label = QLabel("状态: 未检测")
        self.status_label.setFont(QFont("Microsoft YaHei", 10))
        status_layout.addWidget(self.status_label)
        
        # 状态指示器
        self.status_indicator = QLabel("●")
        self.status_indicator.setFont(QFont("Microsoft YaHei", 16))
        self.status_indicator.setStyleSheet("color: gray;")
        status_layout.addWidget(self.status_indicator)
        
        status_layout.addStretch()
        
        # 响应时间
        self.response_time_label = QLabel("响应时间: --")
        self.response_time_label.setFont(QFont("Microsoft YaHei", 9))
        status_layout.addWidget(self.response_time_label)
        
        github_layout.addLayout(status_layout)
        
        # 进度条
        self.progress_bar = QProgressBar()
        self.progress_bar.setVisible(False)
        self.progress_bar.setRange(0, 0)  # 无限进度条
        github_layout.addWidget(self.progress_bar)
        
        # 检测按钮
        button_layout = QHBoxLayout()
        
        self.check_button = QPushButton("检测连接")
        self.check_button.setStyleSheet("""
            QPushButton {
                background-color: #4CAF50;
                color: white;
                border: none;
                padding: 8px 16px;
                border-radius: 4px;
                font-weight: bold;
            }
            QPushButton:hover {
                background-color: #45a049;
            }
            QPushButton:disabled {
                background-color: #cccccc;
                color: #666666;
            }
        """)
        button_layout.addWidget(self.check_button)
        
        self.clear_logs_button = QPushButton("清空日志")
        self.clear_logs_button.setStyleSheet("""
            QPushButton {
                background-color: #f44336;
                color: white;
                border: none;
                padding: 8px 16px;
                border-radius: 4px;
            }
            QPushButton:hover {
                background-color: #da190b;
            }
        """)
        button_layout.addWidget(self.clear_logs_button)
        
        button_layout.addStretch()
        github_layout.addLayout(button_layout)
        
        # 日志显示区域
        log_label = QLabel("检测日志:")
        log_label.setFont(QFont("Microsoft YaHei", 10, QFont.Bold))
        github_layout.addWidget(log_label)
        
        self.log_text = QTextEdit()
        self.log_text.setReadOnly(True)
        self.log_text.setMaximumHeight(150)
        self.log_text.setStyleSheet("""
            QTextEdit {
                background-color: #f5f5f5;
                border: 1px solid #ddd;
                border-radius: 4px;
                padding: 8px;
                font-family: Consolas, 'Courier New', monospace;
                font-size: 9px;
            }
        """)
        github_layout.addWidget(self.log_text)
        
        layout.addWidget(github_group)
        
    def setup_help_section(self, layout):
        """设置帮助信息区域"""
        help_group = QGroupBox("帮助信息")
        help_layout = QVBoxLayout(help_group)
        
        help_text = QTextEdit()
        help_text.setReadOnly(True)
        help_text.setHtml("""
        <h3>GitHub连接检测说明</h3>
        <p><b>功能用途：</b>检测当前系统与GitHub服务器的网络连接状态，用于调试自动更新功能。</p>
        
        <h4>连接状态说明：</h4>
        <ul>
        <li><span style='color: green;'>● 已连接</span> - 系统可以正常访问GitHub</li>
        <li><span style='color: red;'>● 连接错误</span> - 网络连接存在问题</li>
        <li><span style='color: orange;'>● 超时</span> - 连接响应时间过长</li>
        <li><span style='color: gray;'>● 未检测</span> - 尚未进行连接检测</li>
        </ul>
        
        <h4>常见问题排查：</h4>
        <ol>
        <li><b>连接超时：</b>检查网络连接，尝试重启路由器</li>
        <li><b>HTTP错误：</b>可能是GitHub服务暂时不可用</li>
        <li><b>连接错误：</b>检查防火墙设置，确保允许程序访问网络</li>
        <li><b>响应时间过长：</b>网络状况不佳，建议稍后重试</li>
        </ol>
        
        <p><i>提示：如果连接检测持续失败，请检查网络设置或联系网络管理员。</i></p>
        """)
        
        help_layout.addWidget(help_text)
        layout.addWidget(help_group)
        
    def setup_buttons(self, layout):
        """设置按钮区域"""
        button_layout = QHBoxLayout()
        
        button_layout.addStretch()
        
        self.close_button = QPushButton("关闭")
        self.close_button.setStyleSheet("""
            QPushButton {
                background-color: #2196F3;
                color: white;
                border: none;
                padding: 8px 16px;
                border-radius: 4px;
                min-width: 80px;
            }
            QPushButton:hover {
                background-color: #0b7dda;
            }
        """)
        button_layout.addWidget(self.close_button)
        
        layout.addLayout(button_layout)
        
    def setup_connections(self):
        """设置信号连接"""
        self.check_button.clicked.connect(self.check_connection)
        self.clear_logs_button.clicked.connect(self.clear_logs)
        self.close_button.clicked.connect(self.accept)
        
    def check_connection(self):
        """检测连接"""
        if not self.github_api_url:
            QMessageBox.warning(self, "配置错误", "GitHub API地址未配置")
            return
            
        # 禁用按钮，显示进度条
        self.check_button.setEnabled(False)
        self.progress_bar.setVisible(True)
        
        # 更新状态
        self.update_status("检测中...", "orange")
        
        # 创建并启动检测线程
        self.connection_checker = GitHubConnectionChecker(self.github_api_url)
        self.connection_checker.connection_result.connect(self.on_connection_result)
        self.connection_checker.finished.connect(self.on_check_finished)
        self.connection_checker.start()
        
        # 添加日志
        self.add_log("开始检测GitHub连接...")
        
    def on_connection_result(self, result):
        """连接检测结果处理"""
        # 解析结果
        status = result['status']
        response_time = result['response_time']
        error_message = result['error_message']
        timestamp = result['timestamp']
        
        # 更新状态显示
        if status == 'connected':
            self.update_status("已连接", "green")
            status_text = f"✅ 连接成功 - 响应时间: {response_time:.0f}ms"
        elif status in ['timeout', 'connection_error', 'request_error']:
            self.update_status("连接错误", "red")
            status_text = f"❌ {error_message}"
        else:
            self.update_status("未知错误", "red")
            status_text = f"⚠️ {error_message}"
        
        # 更新响应时间
        if response_time > 0:
            self.response_time_label.setText(f"响应时间: {response_time:.0f}ms")
        else:
            self.response_time_label.setText("响应时间: --")
        
        # 添加详细日志
        log_entry = f"[{timestamp}] {status_text}\n"
        
        # 添加网络诊断信息
        if 'network_diagnosis' in result and result['network_diagnosis']:
            log_entry += "📡 网络诊断信息:\n"
            for key, value in result['network_diagnosis'].items():
                log_entry += f"    {key}: {value}\n"
        
        if result['details']:
            log_entry += "📋 响应详情:\n"
            for key, value in result['details'].items():
                log_entry += f"    {key}: {value}\n"
        
        self.add_log(log_entry)
        
        # 保存日志记录
        self.connection_logs.append({
            'timestamp': timestamp,
            'status': status,
            'response_time': response_time,
            'error_message': error_message,
            'details': result['details'],
            'network_diagnosis': result.get('network_diagnosis', {})
        })
        
    def on_check_finished(self):
        """检测完成"""
        # 启用按钮，隐藏进度条
        self.check_button.setEnabled(True)
        self.progress_bar.setVisible(False)
        
        # 清理线程
        if self.connection_checker:
            self.connection_checker.quit()
            self.connection_checker.wait()
            self.connection_checker = None
            
    def update_status(self, text, color):
        """更新状态显示"""
        self.status_label.setText(f"状态: {text}")
        self.status_indicator.setStyleSheet(f"color: {color};")
        
    def add_log(self, message):
        """添加日志"""
        current_text = self.log_text.toPlainText()
        new_text = f"{current_text}\n{message}" if current_text else message
        self.log_text.setPlainText(new_text)
        
        # 自动滚动到底部
        cursor = self.log_text.textCursor()
        cursor.movePosition(cursor.End)
        self.log_text.setTextCursor(cursor)
        
    def clear_logs(self):
        """清空日志"""
        self.log_text.clear()
        self.connection_logs.clear()
        self.add_log("日志已清空")
        
    def get_connection_summary(self):
        """获取连接摘要信息"""
        if not self.connection_logs:
            return "尚未进行连接检测"
            
        latest_log = self.connection_logs[-1]
        status_map = {
            'connected': '✅ 连接正常',
            'timeout': '⏰ 连接超时',
            'connection_error': '❌ 连接错误',
            'request_error': '⚠️ 请求错误',
            'unknown_error': '❓ 未知错误'
        }
        
        status_text = status_map.get(latest_log['status'], '❓ 未知状态')
        
        summary = f"最新检测: {latest_log['timestamp']}\n"
        summary += f"状态: {status_text}\n"
        
        if latest_log['response_time'] > 0:
            summary += f"响应时间: {latest_log['response_time']:.0f}ms\n"
            
        if latest_log['error_message']:
            summary += f"错误信息: {latest_log['error_message']}"
            
        return summary


if __name__ == "__main__":
    # 测试代码
    from PyQt5.QtWidgets import QApplication
    
    app = QApplication(sys.argv)
    
    # 测试GitHub API地址
    test_api_url = "https://api.github.com/repos/244727418/shouhou/releases/latest"
    
    dialog = HelpDialog(github_api_url=test_api_url)
    dialog.show()
    
    sys.exit(app.exec_())