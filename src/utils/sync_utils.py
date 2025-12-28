"""
Sync Utilities - مزامنة البيانات بين جهازين على نفس الشبكة المحلية
"""

import socket
import threading
import json
import os
import zipfile
import tempfile
import shutil
from datetime import datetime


# Default sync port
SYNC_PORT = 5555
BUFFER_SIZE = 8192
HEADER_SIZE = 10


def get_local_ip():
    """الحصول على عنوان IP المحلي للجهاز"""
    try:
        s = socket.socket(socket.AF_INET, socket.SOCK_DGRAM)
        s.connect(("8.8.8.8", 80))
        ip = s.getsockname()[0]
        s.close()
        return ip
    except Exception:
        return "127.0.0.1"


def get_data_folder():
    """الحصول على مجلد البيانات"""
    return os.path.join(os.path.expanduser("~"), "Documents", "alswaife")


def create_backup_zip(progress_callback=None):
    """إنشاء ملف مضغوط من البيانات"""
    data_folder = get_data_folder()
    if not os.path.exists(data_folder):
        return None
    
    temp_dir = tempfile.gettempdir()
    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    zip_path = os.path.join(temp_dir, f"alswaife_sync_{timestamp}.zip")
    
    try:
        with zipfile.ZipFile(zip_path, 'w', zipfile.ZIP_DEFLATED) as zipf:
            total_files = sum([len(files) for _, _, files in os.walk(data_folder)])
            processed = 0
            
            for root, dirs, files in os.walk(data_folder):
                for file in files:
                    file_path = os.path.join(root, file)
                    arcname = os.path.relpath(file_path, data_folder)
                    zipf.write(file_path, arcname)
                    processed += 1
                    if progress_callback:
                        progress_callback(processed / total_files * 100)
        
        return zip_path
    except Exception as e:
        print(f"[ERROR] Failed to create backup: {e}")
        return None


def extract_backup_zip(zip_path, progress_callback=None):
    """استخراج ملف مضغوط إلى مجلد البيانات"""
    data_folder = get_data_folder()
    
    try:
        # إنشاء نسخة احتياطية قبل الاستبدال
        backup_folder = data_folder + "_backup_" + datetime.now().strftime("%Y%m%d_%H%M%S")
        if os.path.exists(data_folder):
            shutil.copytree(data_folder, backup_folder)
        
        with zipfile.ZipFile(zip_path, 'r') as zipf:
            total_files = len(zipf.namelist())
            for i, file in enumerate(zipf.namelist()):
                zipf.extract(file, data_folder)
                if progress_callback:
                    progress_callback((i + 1) / total_files * 100)
        
        return True, backup_folder
    except Exception as e:
        print(f"[ERROR] Failed to extract backup: {e}")
        return False, None


class SyncServer:
    """خادم المزامنة - يستقبل البيانات من جهاز آخر"""
    
    def __init__(self, port=SYNC_PORT):
        self.port = port
        self.server_socket = None
        self.running = False
        self.on_progress = None
        self.on_complete = None
        self.on_error = None
    
    def start(self):
        """بدء الخادم"""
        try:
            self.server_socket = socket.socket(socket.AF_INET, socket.SOCK_STREAM)
            self.server_socket.setsockopt(socket.SOL_SOCKET, socket.SO_REUSEADDR, 1)
            self.server_socket.bind(('0.0.0.0', self.port))
            self.server_socket.listen(1)
            self.running = True
            
            thread = threading.Thread(target=self._accept_connections)
            thread.daemon = True
            thread.start()
            
            return True, get_local_ip()
        except Exception as e:
            return False, str(e)
    
    def stop(self):
        """إيقاف الخادم"""
        self.running = False
        if self.server_socket:
            try:
                self.server_socket.close()
            except:
                pass
    
    def _accept_connections(self):
        """قبول الاتصالات الواردة"""
        while self.running:
            try:
                self.server_socket.settimeout(1.0)
                try:
                    client_socket, address = self.server_socket.accept()
                    self._handle_client(client_socket, address)
                except socket.timeout:
                    continue
            except Exception as e:
                if self.running and self.on_error:
                    self.on_error(str(e))
                break
    
    def _handle_client(self, client_socket, address):
        """معالجة اتصال العميل"""
        try:
            # استقبال حجم الملف
            header = client_socket.recv(HEADER_SIZE).decode('utf-8')
            file_size = int(header.strip())
            
            # استقبال الملف
            temp_dir = tempfile.gettempdir()
            zip_path = os.path.join(temp_dir, "received_sync.zip")
            
            received = 0
            with open(zip_path, 'wb') as f:
                while received < file_size:
                    chunk = client_socket.recv(min(BUFFER_SIZE, file_size - received))
                    if not chunk:
                        break
                    f.write(chunk)
                    received += len(chunk)
                    if self.on_progress:
                        self.on_progress(received / file_size * 50)  # 50% للاستقبال
            
            # استخراج الملف
            def extract_progress(p):
                if self.on_progress:
                    self.on_progress(50 + p * 0.5)  # 50% للاستخراج
            
            success, backup_path = extract_backup_zip(zip_path, extract_progress)
            
            # إرسال تأكيد
            if success:
                client_socket.send(b"OK")
                if self.on_complete:
                    self.on_complete(True, "تم استقبال البيانات بنجاح")
            else:
                client_socket.send(b"FAIL")
                if self.on_error:
                    self.on_error("فشل في استخراج البيانات")
            
            # حذف الملف المؤقت
            try:
                os.remove(zip_path)
            except:
                pass
                
        except Exception as e:
            if self.on_error:
                self.on_error(str(e))
        finally:
            client_socket.close()


class SyncClient:
    """عميل المزامنة - يرسل البيانات إلى جهاز آخر"""
    
    def __init__(self):
        self.on_progress = None
        self.on_complete = None
        self.on_error = None
    
    def send_data(self, target_ip, port=SYNC_PORT):
        """إرسال البيانات إلى الجهاز الهدف"""
        thread = threading.Thread(target=self._send_data_thread, args=(target_ip, port))
        thread.daemon = True
        thread.start()
    
    def _send_data_thread(self, target_ip, port):
        """خيط إرسال البيانات"""
        try:
            # إنشاء ملف مضغوط
            def zip_progress(p):
                if self.on_progress:
                    self.on_progress(p * 0.3)  # 30% للضغط
            
            zip_path = create_backup_zip(zip_progress)
            if not zip_path:
                if self.on_error:
                    self.on_error("فشل في إنشاء ملف النسخ الاحتياطي")
                return
            
            # الاتصال بالخادم
            client_socket = socket.socket(socket.AF_INET, socket.SOCK_STREAM)
            client_socket.settimeout(30)
            client_socket.connect((target_ip, port))
            
            # إرسال حجم الملف
            file_size = os.path.getsize(zip_path)
            header = f"{file_size:<{HEADER_SIZE}}"
            client_socket.send(header.encode('utf-8'))
            
            # إرسال الملف
            sent = 0
            with open(zip_path, 'rb') as f:
                while sent < file_size:
                    chunk = f.read(BUFFER_SIZE)
                    if not chunk:
                        break
                    client_socket.send(chunk)
                    sent += len(chunk)
                    if self.on_progress:
                        self.on_progress(30 + (sent / file_size * 70))  # 70% للإرسال
            
            # انتظار التأكيد
            response = client_socket.recv(10).decode('utf-8')
            
            if response.startswith("OK"):
                if self.on_complete:
                    self.on_complete(True, "تم إرسال البيانات بنجاح")
            else:
                if self.on_error:
                    self.on_error("فشل في استقبال البيانات على الجهاز الآخر")
            
            client_socket.close()
            
            # حذف الملف المؤقت
            try:
                os.remove(zip_path)
            except:
                pass
                
        except socket.timeout:
            if self.on_error:
                self.on_error("انتهت مهلة الاتصال - تأكد من أن الجهاز الآخر في وضع الاستقبال")
        except ConnectionRefusedError:
            if self.on_error:
                self.on_error("تم رفض الاتصال - تأكد من أن الجهاز الآخر في وضع الاستقبال")
        except Exception as e:
            if self.on_error:
                self.on_error(f"خطأ في الاتصال: {str(e)}")


def discover_devices(port=SYNC_PORT, timeout=3):
    """اكتشاف الأجهزة المتاحة على الشبكة"""
    devices = []
    local_ip = get_local_ip()
    
    # الحصول على نطاق الشبكة
    ip_parts = local_ip.split('.')
    if len(ip_parts) != 4:
        return devices
    
    network_prefix = '.'.join(ip_parts[:3])
    
    def check_device(ip):
        try:
            sock = socket.socket(socket.AF_INET, socket.SOCK_STREAM)
            sock.settimeout(0.5)
            result = sock.connect_ex((ip, port))
            sock.close()
            if result == 0 and ip != local_ip:
                devices.append(ip)
        except:
            pass
    
    threads = []
    for i in range(1, 255):
        ip = f"{network_prefix}.{i}"
        t = threading.Thread(target=check_device, args=(ip,))
        t.start()
        threads.append(t)
    
    for t in threads:
        t.join(timeout=timeout)
    
    return devices
