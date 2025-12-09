"""
Kompyuter ekranini yozib olish va ishchi faolligini monitoring qilish tizimi (Optimallashtirilgan)
- CRM kirishlarini aniqlash va sanash
- Telefon foydalanishini aniqlash va sanash
- Vaqt belgilari bilan barcha faollikni yozish
- Mijozlar bilan ishlashni kuzatish
- Kompyuter foydalanishini monitoring qilish
- Faqat muhim voqealarda video yozib olish
"""

import cv2
import mss
import numpy as np
from ultralytics import YOLO
import pygetwindow as gw
import pytesseract
import psutil
from datetime import datetime
import json
import os
import threading
import time
import pandas as pd
from flask import Flask, render_template, jsonify, send_from_directory, Response, request
from flask_cors import CORS
import re
import subprocess
import platform


class ActivityMonitor:
    def __init__(self, camera_url=None, crm_keywords=None, output_dir="activity_logs", web_port=5000):
        """
        ActivityMonitor - ishchi faolligini monitoring qilish tizimi
        
        Args:
            camera_url: RTSP kamera URL
            crm_keywords: CRM tizimini aniqlash uchun kalit so'zlar ro'yxati
            output_dir: Log fayllarini saqlash papkasi
            web_port: Web dashboard porti
        """
        self.camera_url = camera_url
        self.crm_keywords = crm_keywords or ["crm", "client", "mijoz", "customer", "salesforce", "hubspot", "bitrix"]
        self.output_dir = output_dir
        self.web_port = web_port
        self.create_output_dir()
        
        # YOLO modelini yuklash (telefon aniqlash uchun)
        self.model = YOLO("yolov8n.pt")
        
        # Faollik ma'lumotlarini saqlash
        self.activities = []
        self.crm_access_count = 0
        self.phone_usage_count = 0
        self.client_interactions = []
        self.computer_usage_sessions = []
        self.website_visits = []  # Sayt/sahifa tashriflari
        self.website_count = {}  # Har bir sayt uchun hisoblagich
        self.process_activities = []  # Jarayon faolliklari
        self.process_count = {}  # Har bir jarayon uchun hisoblagich
        self.last_active_process = None  # So'nggi faol jarayon
        
        # Vaqt belgilari
        self.current_session_start = None
        self.last_crm_access_time = None
        self.last_phone_detection_time = None
        self.last_client_interaction_time = None
        self.last_website_title = None  # So'nggi sayt/sahifa
        self.last_website_time = None
        
        # Threading
        self.is_running = False
        self.camera_monitoring_thread = None
        self.activity_tracking_thread = None
        
        # Video yozib olish (faqat muhim voqealarda)
        self.is_recording = False
        self.video_writer = None
        self.recording_start_time = None
        self.recording_event = None
        self.recording_duration = 30  # Sekundlarda (muhim voqea uchun)
        self.video_recording_thread = None
        
        # Web server
        self.app = Flask(__name__, template_folder='templates', static_folder='static')
        CORS(self.app)
        self.setup_routes()
        
    def create_output_dir(self):
        """Chiqish papkasini yaratish"""
        dirs = [self.output_dir, 
                os.path.join(self.output_dir, "videos"),
                os.path.join(self.output_dir, "screenshots"),
                "templates", "static"]
        for d in dirs:
            if not os.path.exists(d):
                os.makedirs(d)
    
    def start_video_recording(self, event_type, event_info=""):
        """Muhim voqea uchun video yozib olishni boshlash"""
        if self.is_recording:
            return  # Allaqachon yozib olinmoqda
        
        try:
            self.is_recording = True
            self.recording_event = event_type
            self.recording_start_time = datetime.now()
            
            # Monitor ma'lumotlarini olish (asosiy threadda)
            with mss.mss() as temp_sct:
                monitor = temp_sct.monitors[1]
                monitor_info = {
                    "width": monitor["width"],
                    "height": monitor["height"],
                    "top": monitor["top"],
                    "left": monitor["left"]
                }
            
            # Video fayl nomi (voqea turi va vaqt bilan)
            # Format keyinroq codec ga qarab o'zgartiriladi
            video_filename = os.path.join(
                self.output_dir, 
                "videos", 
                f"{event_type}_{datetime.now().strftime('%Y-%m-%d_%H-%M-%S')}.mp4"
            )
            
            # Video filename ni saqlash (keyinroq o'zgartirish mumkin)
            self.current_video_filename = video_filename
            
            # Video codec tanlash - Windows uchun ishlaydigan codec'lar
            fps = 10
            fourcc = None
            
            # Windows uchun ishlaydigan codec'larni sinab ko'rish
            # H.264 va avc1 ba'zi sistemalarda ishlamaydi (libopenh264 muammosi)
            # mp4v va XVID ko'proq ishlaydi
            codecs_to_try = ['mp4v', 'XVID', 'MJPG']  # H.264 va avc1 ni olib tashladik
            
            for codec in codecs_to_try:
                try:
                    fourcc = cv2.VideoWriter_fourcc(*codec)
                    # Test yozuvchi yaratish
                    test_file = video_filename.replace('.mp4', f'_test_{codec}.mp4')
                    test_writer = cv2.VideoWriter(test_file, fourcc, fps, (monitor_info["width"], monitor_info["height"]))
                    if test_writer.isOpened():
                        test_writer.release()
                        # Test faylni o'chirish
                        try:
                            if os.path.exists(test_file):
                                os.remove(test_file)
                        except:
                            pass
                        print(f"[VIDEO] Codec tanlandi: {codec}")
                        break
                    else:
                        fourcc = None
                except Exception as e:
                    continue
            
            # Agar hech biri ishlamasa, mp4v default (eng keng tarqalgan)
            if fourcc is None:
                fourcc = cv2.VideoWriter_fourcc(*'mp4v')
                print(f"[VIDEO] Default codec ishlatilmoqda: mp4v")
            
            self.video_writer = cv2.VideoWriter(video_filename, fourcc, fps, (monitor_info["width"], monitor_info["height"]))
            
            if not self.video_writer.isOpened():
                # Agar hali ham ochilmasa, XVID bilan urinib ko'rish
                print(f"[VIDEO] mp4v ishlamadi, XVID sinab ko'rilmoqda...")
                fourcc = cv2.VideoWriter_fourcc(*'XVID')
                # XVID uchun .avi format ishlatish
                video_filename = video_filename.replace('.mp4', '.avi')
                self.video_writer = cv2.VideoWriter(video_filename, fourcc, fps, (monitor_info["width"], monitor_info["height"]))
                
                if self.video_writer.isOpened():
                    print(f"[VIDEO] XVID codec bilan .avi formatida yozilmoqda")
                else:
                    # Oxirgi variant - MJPG
                    print(f"[VIDEO] XVID ishlamadi, MJPG sinab ko'rilmoqda...")
                    fourcc = cv2.VideoWriter_fourcc(*'MJPG')
                    video_filename = video_filename.replace('.avi', '.avi')
                    self.video_writer = cv2.VideoWriter(video_filename, fourcc, fps, (monitor_info["width"], monitor_info["height"]))
            
            print(f"[VIDEO] {event_type} uchun yozib olish boshlandi: {video_filename} (Codec: {fourcc})")
            
            # Video filename ni saqlash
            self.current_video_filename = video_filename
            
            # Video yozib olish thread (monitor_info ni uzatish, MSS emas)
            self.video_recording_thread = threading.Thread(
                target=self._record_video_worker,
                args=(monitor_info, fps, video_filename),
                daemon=True
            )
            self.video_recording_thread.start()
            
        except Exception as e:
            print(f"Video yozib olishni boshlashda xatolik: {e}")
            import traceback
            traceback.print_exc()
            self.is_recording = False
    
    def _record_video_worker(self, monitor_info, fps, video_filename):
        """Video yozib olish worker thread (thread-safe)"""
        start_time = time.time()
        frame_count = 0
        frames = []  # Frame'larni saqlash
        
        # Har bir thread uchun alohida MSS obyekti yaratish (thread-safe)
        sct = None
        try:
            sct = mss.mss()
            monitor = {
                "top": monitor_info["top"],
                "left": monitor_info["left"],
                "width": monitor_info["width"],
                "height": monitor_info["height"]
            }
            
            while self.is_recording and (time.time() - start_time) < self.recording_duration:
                try:
                    screenshot = sct.grab(monitor)
                    img = np.array(screenshot)
                    img = cv2.cvtColor(img, cv2.COLOR_BGRA2BGR)
                    
                    # Frame'larni saqlash (backup uchun)
                    if len(frames) < 300:  # Faqat birinchi 300 frame (xotira tejash)
                        frames.append(img.copy())
                    
                    # Videoga yozish
                    if self.video_writer and self.video_writer.isOpened():
                        self.video_writer.write(img)
                    
                    frame_count += 1
                    time.sleep(1.0 / fps)
                except Exception as e:
                    print(f"Frame yozishda xatolik: {e}")
                    time.sleep(0.1)
                    continue
            
            # Video yozib olishni to'xtatish
            if self.video_writer:
                self.video_writer.release()
            
            # Agar video yozilmagan bo'lsa, qayta yozish
            if not os.path.exists(video_filename) or os.path.getsize(video_filename) == 0:
                print(f"[VIDEO] Video yozilmadi, qayta urinib ko'rilmoqda...")
                
                # mp4v codec bilan qayta urinib ko'rish
                fourcc = cv2.VideoWriter_fourcc(*'mp4v')
                temp_writer = cv2.VideoWriter(video_filename, fourcc, fps, (monitor_info["width"], monitor_info["height"]))
                
                if temp_writer.isOpened() and len(frames) > 0:
                    for frame in frames:
                        temp_writer.write(frame)
                    temp_writer.release()
                    
                    # Video faylni tekshirish
                    if os.path.exists(video_filename) and os.path.getsize(video_filename) > 0:
                        print(f"[VIDEO] Video mp4v codec bilan qayta yozildi ({len(frames)} frame)")
                    else:
                        # XVID bilan urinib ko'rish (.avi format)
                        print(f"[VIDEO] mp4v ishlamadi, XVID sinab ko'rilmoqda...")
                        video_filename_avi = video_filename.replace('.mp4', '.avi')
                        fourcc = cv2.VideoWriter_fourcc(*'XVID')
                        temp_writer = cv2.VideoWriter(video_filename_avi, fourcc, fps, (monitor_info["width"], monitor_info["height"]))
                        
                        if temp_writer.isOpened() and len(frames) > 0:
                            for frame in frames:
                                temp_writer.write(frame)
                            temp_writer.release()
                            
                            if os.path.exists(video_filename_avi) and os.path.getsize(video_filename_avi) > 0:
                                print(f"[VIDEO] Video XVID codec bilan .avi formatida yozildi ({len(frames)} frame)")
                                video_filename = video_filename_avi
                        else:
                            print(f"[VIDEO] Xatolik: Video writer ochilmadi")
                else:
                    print(f"[VIDEO] Xatolik: Video writer ochilmadi")
            
            duration = time.time() - start_time
            if os.path.exists(video_filename):
                file_size = os.path.getsize(video_filename) / (1024 * 1024)  # MB
                print(f"[VIDEO] Yozib olish to'xtatildi: {video_filename} ({duration:.1f}s, {file_size:.2f}MB, {frame_count} frame)")
            else:
                print(f"[VIDEO] Xatolik: Video fayl yaratilmadi")
            
        except Exception as e:
            print(f"Video yozib olishda xatolik: {e}")
            import traceback
            traceback.print_exc()
        finally:
            # MSS obyektini yopish
            if sct:
                try:
                    sct.close()
                except:
                    pass
            self.is_recording = False
            self.video_writer = None
    
    def stop_video_recording(self):
        """Video yozib olishni to'xtatish"""
        self.is_recording = False
        if self.video_recording_thread:
            self.video_recording_thread.join(timeout=2)
    
    def detect_crm_access(self):
        """CRM tizimiga kirishni aniqlash"""
        try:
            active_windows = gw.getWindowsWithTitle("")
            
            for window in active_windows:
                if window.visible:
                    window_title = window.title.lower()
                    
                    for keyword in self.crm_keywords:
                        if keyword in window_title:
                            current_time = datetime.now()
                            
                            if (self.last_crm_access_time is None or 
                                (current_time - self.last_crm_access_time).total_seconds() > 5):
                                
                                self.crm_access_count += 1
                                self.last_crm_access_time = current_time
                                
                                activity = {
                                    "type": "CRM_ACCESS",
                                    "timestamp": current_time.strftime("%Y-%m-%d %H:%M:%S"),
                                    "window_title": window.title,
                                    "count": self.crm_access_count
                                }
                                self.activities.append(activity)
                                self.save_activity(activity)
                                
                                # Video yozib olishni boshlash
                                self.start_video_recording("CRM", window.title)
                                
                                print(f"[CRM] {current_time.strftime('%H:%M:%S')} - CRM ga kirildi (Jami: {self.crm_access_count})")
                                return True
        except Exception as e:
            print(f"CRM aniqlashda xatolik: {e}")
        return False
    
    def detect_phone_usage(self, frame):
        """Kamera orqali telefon foydalanishini aniqlash"""
        try:
            results = self.model(frame)[0]
            
            for box in results.boxes:
                cls = int(box.cls[0])
                conf = float(box.conf[0])
                class_name = self.model.names[cls].lower()
                
                if ("phone" in class_name or "cell" in class_name) and conf > 0.5:
                    current_time = datetime.now()
                    
                    if (self.last_phone_detection_time is None or 
                        (current_time - self.last_phone_detection_time).total_seconds() > 3):
                        
                        self.phone_usage_count += 1
                        self.last_phone_detection_time = current_time
                        
                        activity = {
                            "type": "PHONE_USAGE",
                            "timestamp": current_time.strftime("%Y-%m-%d %H:%M:%S"),
                            "confidence": round(conf, 2),
                            "count": self.phone_usage_count
                        }
                        self.activities.append(activity)
                        self.save_activity(activity)
                        
                        # Video yozib olishni boshlash
                        self.start_video_recording("PHONE", f"Confidence: {conf:.2f}")
                        
                        print(f"[TELEFON] {current_time.strftime('%H:%M:%S')} - Telefon ishlatildi (Jami: {self.phone_usage_count})")
                        return True
        except Exception as e:
            print(f"Telefon aniqlashda xatolik: {e}")
        
        return False
    
    def detect_client_interactions(self):
        """Mijozlar bilan ishlashni aniqlash (yaxshilangan)"""
        try:
            active_window = gw.getActiveWindow()
            if not active_window or not active_window.visible:
                return
            
            window_title = active_window.title
            window_title_lower = window_title.lower()
            
            # Kengroq kalit so'zlar ro'yxati
            client_keywords = [
                "client", "mijoz", "customer", "contact", "lead", "deal", "order", "buyurtma",
                "klient", "клиент", "müşteri", "pelanggan", "cliente",
                "prospect", "potential", "opportunity", "sales", "sotuv",
                "contract", "shartnoma", "agreement", "kelishuv",
                "invoice", "hisob", "payment", "to'lov", "tolov",
                "account", "hisob", "profile", "profil",
                "name", "ism", "phone", "telefon", "email", "pochta",
                "address", "manzil", "company", "kompaniya", "tashkilot"
            ]
            
            # Window title da kalit so'zlarni qidirish
            found_keyword = None
            for keyword in client_keywords:
                if keyword in window_title_lower:
                    found_keyword = keyword
                    break
            
            # Agar window title da topilmasa, OCR yordamida ekran matnini o'qish
            if not found_keyword:
                try:
                    screenshot = self.capture_window_screenshot(active_window)
                    if screenshot is not None:
                        # OCR yordamida matnni o'qish
                        text = pytesseract.image_to_string(screenshot, lang='eng+uzb+rus')
                        text_lower = text.lower()
                        
                        # Matnda kalit so'zlarni qidirish
                        for keyword in client_keywords:
                            if keyword in text_lower:
                                found_keyword = keyword
                                break
                except Exception as e:
                    pass  # OCR xatolik bo'lsa, davom etish
            
            # Agar kalit so'z topilsa
            if found_keyword:
                current_time = datetime.now()
                
                # Takrorlanmasligi uchun tekshirish (window title o'zgarmagan bo'lsa)
                if (self.last_client_interaction_time is None or 
                    (current_time - self.last_client_interaction_time).total_seconds() > 10 or
                    window_title != getattr(self, '_last_client_window_title', '')):
                    
                    interaction = {
                        "type": "CLIENT_INTERACTION",
                        "timestamp": current_time.strftime("%Y-%m-%d %H:%M:%S"),
                        "window_title": window_title,
                        "keyword": found_keyword,
                        "detection_method": "title" if found_keyword in window_title_lower else "ocr"
                    }
                    
                    self.client_interactions.append(interaction)
                    self.activities.append(interaction)
                    self.save_activity(interaction)
                    
                    self.last_client_interaction_time = current_time
                    self._last_client_window_title = window_title
                    
                    # Video yozib olishni boshlash
                    safe_name = re.sub(r'[<>:"/\\|?*]', '_', window_title)[:50]
                    self.start_video_recording("CLIENT", safe_name)
                    
                    print(f"[MIJOZ] {current_time.strftime('%H:%M:%S')} - Mijoz bilan ishlash: {window_title} (kalit: {found_keyword})")
        
        except Exception as e:
            print(f"Mijoz aniqlashda xatolik: {e}")
    
    def capture_window_screenshot(self, window):
        """Oyna skrinshotini olish"""
        try:
            if window and window.visible:
                with mss.mss() as sct:
                    monitor = {
                        "top": window.top,
                        "left": window.left,
                        "width": window.width,
                        "height": window.height
                    }
                    screenshot = sct.grab(monitor)
                    img = np.array(screenshot)
                    img = cv2.cvtColor(img, cv2.COLOR_BGRA2BGR)
                    return img
        except:
            pass
        return None
    
    def monitor_computer_usage(self):
        """Kompyuter foydalanishini monitoring qilish"""
        try:
            active_windows = gw.getWindowsWithTitle("")
            has_active_window = any(w.visible for w in active_windows)
            current_time = datetime.now()
            
            if has_active_window:
                if self.current_session_start is None:
                    self.current_session_start = current_time
                    session = {
                        "start_time": current_time.strftime("%Y-%m-%d %H:%M:%S"),
                        "end_time": None,
                        "duration_seconds": 0
                    }
                    self.computer_usage_sessions.append(session)
            else:
                if self.current_session_start is not None:
                    duration = (current_time - self.current_session_start).total_seconds()
                    if self.computer_usage_sessions:
                        self.computer_usage_sessions[-1]["end_time"] = current_time.strftime("%Y-%m-%d %H:%M:%S")
                        self.computer_usage_sessions[-1]["duration_seconds"] = duration
                    self.current_session_start = None
        except Exception as e:
            print(f"Kompyuter monitoring xatolik: {e}")
    
    def save_activity(self, activity):
        """Faollikni JSON faylga saqlash"""
        try:
            log_file = os.path.join(self.output_dir, f"activities_{datetime.now().strftime('%Y-%m-%d')}.json")
            
            if os.path.exists(log_file):
                with open(log_file, 'r', encoding='utf-8') as f:
                    data = json.load(f)
            else:
                data = []
            
            data.append(activity)
            
            with open(log_file, 'w', encoding='utf-8') as f:
                json.dump(data, f, ensure_ascii=False, indent=2)
        except Exception as e:
            print(f"Faollikni saqlashda xatolik: {e}")
    
    def camera_monitoring_worker(self):
        """Kamera monitoring thread"""
        if not self.camera_url:
            return
        
        try:
            cap = cv2.VideoCapture(self.camera_url)
            if not cap.isOpened():
                print(f"[KAMERA] Kameraga ulanib bo'lmadi: {self.camera_url}")
                return
            
            print(f"[KAMERA] Kamera monitoring boshlandi")
            
            while self.is_running:
                ret, frame = cap.read()
                if not ret:
                    time.sleep(1)
                    continue
                
                self.detect_phone_usage(frame)
                time.sleep(0.5)
            
            cap.release()
        except Exception as e:
            print(f"Kamera monitoring xatolik: {e}")
    
    def get_active_process_info(self):
        """Faol jarayon ma'lumotlarini olish"""
        try:
            active_window = gw.getActiveWindow()
            if not active_window or not active_window.visible:
                return None
            
            window_title = active_window.title
            process_name = None
            process_path = None
            
            try:
                # Window title dan jarayon nomini olish
                if " - " in window_title:
                    parts = window_title.split(" - ")
                    if len(parts) > 1:
                        process_name = parts[-1].lower()
                else:
                    # Window title dan jarayon nomini aniqlash
                    process_name = window_title.lower()
                
                # psutil yordamida jarayon ma'lumotlarini olish
                for proc in psutil.process_iter(['pid', 'name', 'exe', 'cmdline']):
                    try:
                        if proc.info['name'] and window_title:
                            proc_name_lower = proc.info['name'].lower()
                            # Browser jarayonlarini aniqlash
                            browsers = ["chrome.exe", "firefox.exe", "msedge.exe", "opera.exe", 
                                       "safari.exe", "brave.exe", "yandex.exe"]
                            if any(browser in proc_name_lower for browser in browsers):
                                process_name = proc.info['name']
                                process_path = proc.info.get('exe', '')
                                break
                    except (psutil.NoSuchProcess, psutil.AccessDenied):
                        continue
                
                return {
                    "window_title": window_title,
                    "process_name": process_name or "Noma'lum",
                    "process_path": process_path or ""
                }
            except Exception as e:
                return {
                    "window_title": window_title,
                    "process_name": "Noma'lum",
                    "process_path": ""
                }
        except Exception as e:
            return None
    
    def detect_process_activity(self):
        """Jarayon faolligini aniqlash"""
        try:
            process_info = self.get_active_process_info()
            if not process_info:
                return
            
            current_time = datetime.now()
            process_name = process_info.get("process_name", "Noma'lum")
            window_title = process_info.get("window_title", "")
            
            # Agar jarayon o'zgarmagan bo'lsa, yangi faollik emas
            if (self.last_active_process != process_name or 
                self.last_active_process is None or
                (current_time - datetime.now()).total_seconds() > 3):
                
                # Yangi jarayon faolligi
                if process_name not in self.process_count:
                    self.process_count[process_name] = 0
                
                self.process_count[process_name] += 1
                
                activity = {
                    "type": "PROCESS_ACTIVITY",
                    "timestamp": current_time.strftime("%Y-%m-%d %H:%M:%S"),
                    "process_name": process_name,
                    "window_title": window_title,
                    "process_path": process_info.get("process_path", ""),
                    "count": self.process_count[process_name]
                }
                
                self.process_activities.append(activity)
                self.activities.append(activity)
                self.save_activity(activity)
                
                self.last_active_process = process_name
                
                # Muhim jarayonlar uchun video yozib olish
                important_processes = ["chrome", "firefox", "edge", "opera", "excel", "word", 
                                     "notepad", "code", "pycharm", "idea"]
                if any(imp in process_name.lower() for imp in important_processes):
                    safe_name = re.sub(r'[<>:"/\\|?*]', '_', process_name)[:50]
                    self.start_video_recording("PROCESS", safe_name)
                
                print(f"[JARAYON] {current_time.strftime('%H:%M:%S')} - {process_name} (Jami: {self.process_count[process_name]} marta)")
        
        except Exception as e:
            print(f"Jarayon aniqlashda xatolik: {e}")
    
    def detect_website_visits(self):
        """Sayt/sahifa tashriflarini aniqlash"""
        try:
            active_window = gw.getActiveWindow()
            if not active_window or not active_window.visible:
                return
            
            window_title = active_window.title
            
            # Browser oynalarini aniqlash
            browsers = ["chrome", "firefox", "edge", "opera", "safari", "brave", "yandex"]
            is_browser = any(browser in window_title.lower() for browser in browsers)
            
            if is_browser:
                # Sayt nomini olish (title dan)
                # Format: "Sayt nomi - Browser" yoki "Sayt nomi | Browser"
                site_name = window_title
                for separator in [" - ", " | ", " — "]:
                    if separator in site_name:
                        site_name = site_name.split(separator)[0].strip()
                        break
                
                # Agar sayt nomi o'zgarmagan bo'lsa, yangi tashrif emas
                current_time = datetime.now()
                
                if (self.last_website_title != site_name or 
                    self.last_website_title is None or
                    (self.last_website_time and (current_time - self.last_website_time).total_seconds() > 5)):
                    
                    # Yangi sayt/sahifa tashrifi
                    if site_name not in self.website_count:
                        self.website_count[site_name] = 0
                    
                    self.website_count[site_name] += 1
                    
                    visit = {
                        "type": "WEBSITE_VISIT",
                        "timestamp": current_time.strftime("%Y-%m-%d %H:%M:%S"),
                        "site_name": site_name,
                        "window_title": window_title,
                        "visit_count": self.website_count[site_name]
                    }
                    
                    self.website_visits.append(visit)
                    self.activities.append(visit)
                    self.save_activity(visit)
                    
                    self.last_website_title = site_name
                    self.last_website_time = current_time
                    
                    # Video yozib olishni boshlash
                    safe_filename = re.sub(r'[<>:"/\\|?*]', '_', site_name)[:50]  # Xavfsiz fayl nomi
                    self.start_video_recording("WEBSITE", safe_filename)
                    
                    print(f"[SAYT] {current_time.strftime('%H:%M:%S')} - {site_name} (Jami: {self.website_count[site_name]} marta)")
        
        except Exception as e:
            print(f"Sayt aniqlashda xatolik: {e}")
    
    def activity_tracking_worker(self):
        """Faollikni kuzatish thread"""
        while self.is_running:
            try:
                self.detect_crm_access()
                self.detect_client_interactions()
                self.detect_website_visits()  # Sayt monitoring
                self.detect_process_activity()  # Jarayon monitoring qo'shildi
                self.monitor_computer_usage()
                time.sleep(2)
            except Exception as e:
                print(f"Faollik kuzatishda xatolik: {e}")
                time.sleep(2)
    
    def setup_routes(self):
        """Web server route'larini sozlash"""
        
        @self.app.route('/')
        def index():
            return render_template('dashboard.html')
        
        @self.app.route('/api/stats')
        def get_stats():
            """Real-time statistika"""
            total_computer_time = sum(s.get("duration_seconds", 0) for s in self.computer_usage_sessions)
            
            return jsonify({
                "crm_access_count": self.crm_access_count,
                "phone_usage_count": self.phone_usage_count,
                "client_interactions_count": len(self.client_interactions),
                "website_visits_count": len(self.website_visits),
                "unique_websites_count": len(self.website_count),
                "computer_sessions_count": len(self.computer_usage_sessions),
                "total_computer_time_hours": round(total_computer_time / 3600, 2),
                "is_recording": self.is_recording,
                "recording_event": self.recording_event if self.is_recording else None
            })
        
        @self.app.route('/api/websites')
        def get_websites():
            """Sayt tashriflari ro'yxati"""
            return jsonify({
                "visits": self.website_visits[-50:] if len(self.website_visits) > 50 else self.website_visits,
                "counts": self.website_count
            })
        
        @self.app.route('/api/activities')
        def get_activities():
            """So'nggi faolliklar"""
            recent_activities = self.activities[-50:] if len(self.activities) > 50 else self.activities
            return jsonify(recent_activities)
        
        @self.app.route('/api/videos')
        def get_videos():
            """Video fayllar ro'yxati"""
            video_dir = os.path.join(self.output_dir, "videos")
            videos = []
            if os.path.exists(video_dir):
                for f in os.listdir(video_dir):
                    # Barcha video formatlarni qo'llab-quvvatlash
                    if f.endswith(('.mp4', '.avi', '.mov', '.mkv')):
                        filepath = os.path.join(video_dir, f)
                        file_size = os.path.getsize(filepath) / (1024 * 1024)  # MB
                        mod_time = datetime.fromtimestamp(os.path.getmtime(filepath))
                        videos.append({
                            "filename": f,
                            "size_mb": round(file_size, 2),
                            "created": mod_time.strftime("%Y-%m-%d %H:%M:%S"),
                            "format": f.split('.')[-1]
                        })
            return jsonify(sorted(videos, key=lambda x: x["created"], reverse=True))
        
        @self.app.route('/videos/<filename>')
        def serve_video(filename):
            """Video fayllarni xizmat qilish (to'g'ri headers bilan)"""
            video_path = os.path.join(self.output_dir, "videos", filename)
            
            if not os.path.exists(video_path):
                return "Video topilmadi", 404
            
            # Video fayl hajmini tekshirish
            file_size = os.path.getsize(video_path)
            if file_size == 0:
                return "Video fayl bo'sh", 404
            
            # Video formatini aniqlash
            file_ext = filename.split('.')[-1].lower()
            mime_types = {
                'mp4': 'video/mp4',
                'avi': 'video/x-msvideo',
                'mov': 'video/quicktime',
                'mkv': 'video/x-matroska'
            }
            mime_type = mime_types.get(file_ext, 'video/mp4')
            
            # Range request qo'llab-quvvatlash (video streaming uchun)
            range_header = request.headers.get('Range', None)
            
            if not range_header:
                # Oddiy yuklash
                response = send_from_directory(
                    os.path.join(self.output_dir, "videos"), 
                    filename,
                    mimetype=mime_type
                )
                response.headers['Accept-Ranges'] = 'bytes'
                response.headers['Content-Length'] = str(file_size)
                return response
            
            # Range request uchun streaming
            byte_start = 0
            byte_end = file_size - 1
            
            range_match = re.search(r'bytes=(\d+)-(\d*)', range_header)
            if range_match:
                byte_start = int(range_match.group(1))
                if range_match.group(2):
                    byte_end = int(range_match.group(2))
            
            content_length = byte_end - byte_start + 1
            
            def generate():
                try:
                    with open(video_path, 'rb') as f:
                        f.seek(byte_start)
                        remaining = content_length
                        while remaining:
                            chunk_size = min(1024 * 1024, remaining)  # 1MB chunks
                            chunk = f.read(chunk_size)
                            if not chunk:
                                break
                            remaining -= len(chunk)
                            yield chunk
                except Exception as e:
                    print(f"Video streaming xatolik: {e}")
            
            response = Response(
                generate(),
                206,  # Partial Content
                {
                    'Content-Type': mime_type,
                    'Accept-Ranges': 'bytes',
                    'Content-Length': str(content_length),
                    'Content-Range': f'bytes {byte_start}-{byte_end}/{file_size}',
                    'Cache-Control': 'no-cache'
                }
            )
            return response
    
    def start_monitoring(self):
        """Monitoringni boshlash"""
        if self.is_running:
            return
        
        self.is_running = True
        print("=" * 60)
        print("MONITORING TIZIMI ISHGA TUSHDI")
        print("=" * 60)
        
        # Threadlarni boshlash
        if self.camera_url:
            self.camera_monitoring_thread = threading.Thread(target=self.camera_monitoring_worker, daemon=True)
            self.camera_monitoring_thread.start()
        
        self.activity_tracking_thread = threading.Thread(target=self.activity_tracking_worker, daemon=True)
        self.activity_tracking_thread.start()
        
        print(f"\n✅ Monitoring ishlamoqda...")
        print(f"🌐 Web Dashboard: http://localhost:{self.web_port}")
        print(f"   Yoki tarmoqda: http://0.0.0.0:{self.web_port}")
        print("\nTo'xtatish uchun Ctrl+C bosing\n")
    
    def stop_monitoring(self):
        """Monitoringni to'xtatish"""
        if not self.is_running:
            return
        
        self.is_running = False
        self.stop_video_recording()
        
        if self.current_session_start:
            current_time = datetime.now()
            duration = (current_time - self.current_session_start).total_seconds()
            if self.computer_usage_sessions:
                self.computer_usage_sessions[-1]["end_time"] = current_time.strftime("%Y-%m-%d %H:%M:%S")
                self.computer_usage_sessions[-1]["duration_seconds"] = duration
        
        self.generate_report()
    
    def generate_report(self):
        """Hisobot yaratish"""
        print("\n" + "=" * 60)
        print("HISOBOT")
        print("=" * 60)
        
        total_crm = self.crm_access_count
        total_phone = self.phone_usage_count
        total_client = len(self.client_interactions)
        total_sessions = len(self.computer_usage_sessions)
        total_time = sum(s.get("duration_seconds", 0) for s in self.computer_usage_sessions) / 3600
        
        print(f"\n📊 STATISTIKA:")
        print(f"  • CRM ga kirishlar: {total_crm} marta")
        print(f"  • Telefon foydalanish: {total_phone} marta")
        print(f"  • Mijozlar bilan ishlash: {total_client} marta")
        print(f"  • Kompyuter sessiyalari: {total_sessions} marta")
        print(f"  • Jami kompyuter vaqti: {total_time:.2f} soat")
        
        self.save_to_excel()
    
    def save_to_excel(self):
        """Excel faylga saqlash"""
        try:
            excel_file = os.path.join(self.output_dir, f"activity_report_{datetime.now().strftime('%Y-%m-%d_%H-%M-%S')}.xlsx")
            
            with pd.ExcelWriter(excel_file, engine='openpyxl') as writer:
                if any(a["type"] == "CRM_ACCESS" for a in self.activities):
                    crm_df = pd.DataFrame([a for a in self.activities if a["type"] == "CRM_ACCESS"])
                    crm_df.to_excel(writer, sheet_name="CRM Kirishlar", index=False)
                
                if any(a["type"] == "PHONE_USAGE" for a in self.activities):
                    phone_df = pd.DataFrame([a for a in self.activities if a["type"] == "PHONE_USAGE"])
                    phone_df.to_excel(writer, sheet_name="Telefon Foydalanish", index=False)
                
                if self.client_interactions:
                    client_df = pd.DataFrame(self.client_interactions)
                    client_df.to_excel(writer, sheet_name="Mijozlar bilan ishlash", index=False)
                
                if self.computer_usage_sessions:
                    session_df = pd.DataFrame(self.computer_usage_sessions)
                    session_df.to_excel(writer, sheet_name="Kompyuter Sessiyalari", index=False)
                
                if self.website_visits:
                    website_df = pd.DataFrame(self.website_visits)
                    website_df.to_excel(writer, sheet_name="Sayt Tashriflari", index=False)
                    
                    # Saytlar bo'yicha statistika
                    website_stats = pd.DataFrame([
                        {"Sayt nomi": site, "Tashriflar soni": count}
                        for site, count in self.website_count.items()
                    ])
                    website_stats = website_stats.sort_values("Tashriflar soni", ascending=False)
                    website_stats.to_excel(writer, sheet_name="Saytlar Statistika", index=False)
                
                if self.process_activities:
                    process_df = pd.DataFrame(self.process_activities)
                    process_df.to_excel(writer, sheet_name="Jarayon Faolliklari", index=False)
                    
                    # Jarayonlar bo'yicha statistika
                    process_stats = pd.DataFrame([
                        {"Jarayon nomi": proc, "Ishlatish soni": count}
                        for proc, count in self.process_count.items()
                    ])
                    process_stats = process_stats.sort_values("Ishlatish soni", ascending=False)
                    process_stats.to_excel(writer, sheet_name="Jarayonlar Statistika", index=False)
                
                stats_data = {
                    "Ko'rsatkich": [
                        "CRM ga kirishlar (marta)",
                        "Telefon foydalanish (marta)",
                        "Mijozlar bilan ishlash (marta)",
                        "Sayt tashriflari (marta)",
                        "Unikal saytlar (ta)",
                        "Kompyuter sessiyalari (marta)",
                        "Jami kompyuter vaqti (soat)"
                    ],
                    "Qiymat": [
                        self.crm_access_count,
                        self.phone_usage_count,
                        len(self.client_interactions),
                        len(self.website_visits),
                        len(self.website_count),
                        len(self.computer_usage_sessions),
                        round(sum(s.get("duration_seconds", 0) for s in self.computer_usage_sessions) / 3600, 2)
                    ]
                }
                stats_df = pd.DataFrame(stats_data)
                stats_df.to_excel(writer, sheet_name="Umumiy Statistika", index=False)
            
            print(f"\n📄 Excel hisobot: {excel_file}")
        except Exception as e:
            print(f"Excel saqlashda xatolik: {e}")
    
    def run_web_server(self):
        """Web serverni ishga tushirish"""
        self.app.run(host='0.0.0.0', port=self.web_port, debug=False, threaded=True)


def main():
    """Asosiy funksiya"""
    print("=" * 60)
    print("KOMPYUTER EKRANINI YOZIB OLISH VA FAOLLIK MONITORING TIZIMI")
    print("=" * 60)
    
    camera_url = "rtsp://admin:LKUHDN@192.168.10.63:554/h264"  # None qilib qo'yish mumkin
    crm_keywords = ["crm", "client", "mijoz", "customer", "salesforce", "hubspot", "bitrix"]
    
    monitor = ActivityMonitor(
        camera_url=camera_url,
        crm_keywords=crm_keywords,
        output_dir="activity_logs",
        web_port=5000
    )
    
    try:
        # Web serverni alohida threadda ishga tushirish
        web_thread = threading.Thread(target=monitor.run_web_server, daemon=True)
        web_thread.start()
        time.sleep(1)  # Server ishga tushishini kutish
        
        # Monitoringni boshlash
        monitor.start_monitoring()
        
        while True:
            time.sleep(1)
    
    except KeyboardInterrupt:
        print("\n\nDastur to'xtatilmoqda...")
        monitor.stop_monitoring()
        print("\nDastur to'xtatildi. Xayr!")


if __name__ == "__main__":
    main()
