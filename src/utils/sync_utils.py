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
import hashlib
from datetime import datetime


# Default sync port
SYNC_PORT = 5555
COMPARE_PORT = 5556
UDP_PORT = 5557
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


def get_file_hash(file_path):
    """حساب hash للملف"""
    try:
        hasher = hashlib.md5()
        with open(file_path, 'rb') as f:
            for chunk in iter(lambda: f.read(8192), b''):
                hasher.update(chunk)
        return hasher.hexdigest()
    except:
        return None


def get_file_info(file_path, base_folder):
    """الحصول على معلومات الملف"""
    try:
        stat = os.stat(file_path)
        rel_path = os.path.relpath(file_path, base_folder)
        return {
            'path': rel_path,
            'size': stat.st_size,
            'modified': stat.st_mtime,
            'hash': get_file_hash(file_path)
        }
    except:
        return None


def scan_local_files():
    """فحص الملفات المحلية وإرجاع قائمة بمعلوماتها"""
    data_folder = get_data_folder()
    files_info = {}
    
    if not os.path.exists(data_folder):
        return files_info
    
    for root, dirs, files in os.walk(data_folder):
        for file in files:
            file_path = os.path.join(root, file)
            info = get_file_info(file_path, data_folder)
            if info:
                files_info[info['path']] = info
    
    return files_info


def compare_files(local_files, remote_files):
    """مقارنة الملفات المحلية والبعيدة وإرجاع الفروقات"""
    differences = []
    
    all_paths = set(local_files.keys()) | set(remote_files.keys())
    
    for path in all_paths:
        local = local_files.get(path)
        remote = remote_files.get(path)
        
        if local and not remote:
            # ملف موجود محلياً فقط
            differences.append({
                'path': path,
                'status': 'local_only',
                'status_text': 'موجود محلياً فقط',
                'local_size': local['size'],
                'remote_size': 0,
                'local_modified': local['modified'],
                'remote_modified': 0
            })
        elif remote and not local:
            # ملف موجود على الجهاز البعيد فقط
            differences.append({
                'path': path,
                'status': 'remote_only',
                'status_text': 'موجود على الجهاز الآخر فقط',
                'local_size': 0,
                'remote_size': remote['size'],
                'local_modified': 0,
                'remote_modified': remote['modified']
            })
        elif local['hash'] != remote['hash']:
            # الملف مختلف
            if local['modified'] > remote['modified']:
                status = 'local_newer'
                status_text = 'النسخة المحلية أحدث'
            else:
                status = 'remote_newer'
                status_text = 'النسخة البعيدة أحدث'
            
            differences.append({
                'path': path,
                'status': status,
                'status_text': status_text,
                'local_size': local['size'],
                'remote_size': remote['size'],
                'local_modified': local['modified'],
                'remote_modified': remote['modified']
            })
    
    return differences


def create_selective_zip(file_paths, progress_callback=None):
    """إنشاء ملف مضغوط من ملفات محددة"""
    data_folder = get_data_folder()
    if not os.path.exists(data_folder):
        return None
    
    temp_dir = tempfile.gettempdir()
    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    zip_path = os.path.join(temp_dir, f"alswaife_selective_{timestamp}.zip")
    
    try:
        with zipfile.ZipFile(zip_path, 'w', zipfile.ZIP_DEFLATED) as zipf:
            total_files = len(file_paths)
            for i, rel_path in enumerate(file_paths):
                file_path = os.path.join(data_folder, rel_path)
                if os.path.exists(file_path):
                    zipf.write(file_path, rel_path)
                if progress_callback:
                    progress_callback((i + 1) / total_files * 100)
        
        return zip_path
    except Exception as e:
        print(f"[ERROR] Failed to create selective backup: {e}")
        return None


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


class BroadcastServer:
    """خادم البث - يعلن عن وجود الجهاز على الشبكة"""
    
    def __init__(self, port=UDP_PORT):
        self.port = port
        self.sock = None
        self.running = False
        
    def start(self):
        """بدء خادم البث"""
        try:
            self.sock = socket.socket(socket.AF_INET, socket.SOCK_DGRAM)
            self.sock.setsockopt(socket.SOL_SOCKET, socket.SO_REUSEADDR, 1)
            self.sock.bind(('0.0.0.0', self.port))
            self.running = True
            
            thread = threading.Thread(target=self._listen)
            thread.daemon = True
            thread.start()
        except Exception as e:
            print(f"Failed to start broadcast server: {e}")
            
    def stop(self):
        """إيقاف خادم البث"""
        self.running = False
        if self.sock:
            try:
                self.sock.close()
            except:
                pass

    def _listen(self):
        """الاستماع لرسائل الاستكشاف"""
        while self.running:
            try:
                data, addr = self.sock.recvfrom(1024)
                if data.decode('utf-8').strip() == "DISCOVER_AL_SWAIFE":
                    # الرد برسالة تأكيد
                    response = f"AL_SWAIFE_DEVICE:{socket.gethostname()}"
                    self.sock.sendto(response.encode('utf-8'), addr)
            except:
                pass


class SyncServer:
    """خادم المزامنة - يستقبل البيانات من جهاز آخر"""
    
    def __init__(self, port=SYNC_PORT):
        self.port = port
        self.server_socket = None
        self.running = False
        self.on_progress = None
        self.on_complete = None
        self.on_error = None
        self.broadcast_server = BroadcastServer()
    
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
            
            self.broadcast_server.start()
            
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
        self.broadcast_server.stop()
    
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


def discover_devices(timeout=2.0):
    """اكتشاف الأجهزة باستخدام UDP Broadcast"""
    devices = []
    local_ip = get_local_ip()
    
    try:
        sock = socket.socket(socket.AF_INET, socket.SOCK_DGRAM)
        sock.setsockopt(socket.SOL_SOCKET, socket.SO_BROADCAST, 1)
        sock.settimeout(timeout)
        
        message = "DISCOVER_AL_SWAIFE"
        sock.sendto(message.encode('utf-8'), ('255.255.255.255', UDP_PORT))
        
        start_time = datetime.now()
        while (datetime.now() - start_time).total_seconds() < timeout:
            try:
                data, addr = sock.recvfrom(1024)
                ip = addr[0]
                if ip != local_ip:
                    devices.append(ip)
            except socket.timeout:
                break
            except:
                continue
                
        sock.close()
    except Exception as e:
        print(f"Discovery error: {e}")
        
    return list(set(devices))


class CompareServer:
    """خادم المقارنة - يستقبل طلبات المقارنة ويرسل معلومات الملفات"""
    
    def __init__(self, port=COMPARE_PORT):
        self.port = port
        self.server_socket = None
        self.running = False
        self.on_compare_request = None
        self.on_error = None
    
    def start(self):
        """بدء خادم المقارنة"""
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
        """معالجة طلب المقارنة"""
        try:
            # استقبال نوع الطلب
            request_type = client_socket.recv(20).decode('utf-8').strip()
            
            if request_type == "GET_FILES_INFO":
                # إرسال معلومات الملفات المحلية
                local_files = scan_local_files()
                data = json.dumps(local_files).encode('utf-8')
                
                # إرسال حجم البيانات ثم البيانات
                header = f"{len(data):<{HEADER_SIZE}}"
                client_socket.send(header.encode('utf-8'))
                client_socket.sendall(data)
                
            elif request_type == "RECEIVE_FILES":
                # استقبال ملفات محددة
                # استقبال حجم الملف
                header = client_socket.recv(HEADER_SIZE).decode('utf-8')
                file_size = int(header.strip())
                
                # استقبال الملف
                temp_dir = tempfile.gettempdir()
                zip_path = os.path.join(temp_dir, "received_selective.zip")
                
                received = 0
                with open(zip_path, 'wb') as f:
                    while received < file_size:
                        chunk = client_socket.recv(min(BUFFER_SIZE, file_size - received))
                        if not chunk:
                            break
                        f.write(chunk)
                        received += len(chunk)
                
                # استخراج الملفات
                success, _ = extract_backup_zip(zip_path)
                
                # إرسال تأكيد
                if success:
                    client_socket.send(b"OK")
                else:
                    client_socket.send(b"FAIL")
                
                try:
                    os.remove(zip_path)
                except:
                    pass
                    
        except Exception as e:
            if self.on_error:
                self.on_error(str(e))
        finally:
            client_socket.close()


class CompareClient:
    """عميل المقارنة - يطلب معلومات الملفات ويرسل الملفات المحددة"""
    
    def __init__(self):
        self.on_compare_complete = None
        self.on_send_progress = None
        self.on_send_complete = None
        self.on_error = None
    
    def get_remote_files_info(self, target_ip, port=COMPARE_PORT):
        """الحصول على معلومات الملفات من الجهاز البعيد"""
        thread = threading.Thread(target=self._get_remote_files_thread, args=(target_ip, port))
        thread.daemon = True
        thread.start()
    
    def _get_remote_files_thread(self, target_ip, port):
        """خيط الحصول على معلومات الملفات"""
        try:
            client_socket = socket.socket(socket.AF_INET, socket.SOCK_STREAM)
            client_socket.settimeout(30)
            client_socket.connect((target_ip, port))
            
            # إرسال طلب معلومات الملفات
            request = f"{'GET_FILES_INFO':<20}"
            client_socket.send(request.encode('utf-8'))
            
            # استقبال حجم البيانات
            header = client_socket.recv(HEADER_SIZE).decode('utf-8')
            data_size = int(header.strip())
            
            # استقبال البيانات
            data = b''
            while len(data) < data_size:
                chunk = client_socket.recv(min(BUFFER_SIZE, data_size - len(data)))
                if not chunk:
                    break
                data += chunk
            
            client_socket.close()
            
            remote_files = json.loads(data.decode('utf-8'))
            local_files = scan_local_files()
            
            # مقارنة الملفات
            differences = compare_files(local_files, remote_files)
            
            if self.on_compare_complete:
                self.on_compare_complete(differences, target_ip)
                
        except socket.timeout:
            if self.on_error:
                self.on_error("انتهت مهلة الاتصال - تأكد من أن الجهاز الآخر في وضع المقارنة")
        except ConnectionRefusedError:
            if self.on_error:
                self.on_error("تم رفض الاتصال - تأكد من أن الجهاز الآخر في وضع المقارنة")
        except Exception as e:
            if self.on_error:
                self.on_error(f"خطأ في الاتصال: {str(e)}")
    
    def send_selected_files(self, target_ip, file_paths, port=COMPARE_PORT):
        """إرسال ملفات محددة إلى الجهاز البعيد"""
        thread = threading.Thread(target=self._send_files_thread, args=(target_ip, file_paths, port))
        thread.daemon = True
        thread.start()
    
    def _send_files_thread(self, target_ip, file_paths, port):
        """خيط إرسال الملفات"""
        try:
            # إنشاء ملف مضغوط من الملفات المحددة
            def zip_progress(p):
                if self.on_send_progress:
                    self.on_send_progress(p * 0.3)
            
            zip_path = create_selective_zip(file_paths, zip_progress)
            if not zip_path:
                if self.on_error:
                    self.on_error("فشل في إنشاء ملف النسخ الاحتياطي")
                return
            
            # الاتصال بالخادم
            client_socket = socket.socket(socket.AF_INET, socket.SOCK_STREAM)
            client_socket.settimeout(60)
            client_socket.connect((target_ip, port))
            
            # إرسال نوع الطلب
            request = f"{'RECEIVE_FILES':<20}"
            client_socket.send(request.encode('utf-8'))
            
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
                    if self.on_send_progress:
                        self.on_send_progress(30 + (sent / file_size * 70))
            
            # انتظار التأكيد
            response = client_socket.recv(10).decode('utf-8')
            
            if response.startswith("OK"):
                if self.on_send_complete:
                    self.on_send_complete(True, f"تم إرسال {len(file_paths)} ملف بنجاح")
            else:
                if self.on_error:
                    self.on_error("فشل في استقبال البيانات على الجهاز الآخر")
            
            client_socket.close()
            
            try:
                os.remove(zip_path)
            except:
                pass
                
        except socket.timeout:
            if self.on_error:
                self.on_error("انتهت مهلة الاتصال")
        except Exception as e:
            if self.on_error:
                self.on_error(f"خطأ في الإرسال: {str(e)}")
