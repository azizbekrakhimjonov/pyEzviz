"""
Ma'lumotlarni ko'rish va tahlil qilish dasturi
"""

import json
import os
import pandas as pd
from datetime import datetime
import glob


class DataViewer:
    def __init__(self, data_dir="activity_logs"):
        """Ma'lumotlarni ko'rish klassi"""
        self.data_dir = data_dir
    
    def list_available_files(self):
        """Mavjud fayllarni ko'rsatish"""
        print("\n" + "=" * 60)
        print("MAVJUD MA'LUMOTLAR")
        print("=" * 60)
        
        # JSON fayllar
        json_files = glob.glob(os.path.join(self.data_dir, "activities_*.json"))
        if json_files:
            print(f"\n📄 JSON Fayllar ({len(json_files)} ta):")
            for f in sorted(json_files):
                file_size = os.path.getsize(f) / 1024  # KB
                print(f"  • {os.path.basename(f)} ({file_size:.1f} KB)")
        else:
            print("\n📄 JSON Fayllar: Topilmadi")
        
        # Excel fayllar
        excel_files = glob.glob(os.path.join(self.data_dir, "activity_report_*.xlsx"))
        if excel_files:
            print(f"\n📊 Excel Hisobotlar ({len(excel_files)} ta):")
            for f in sorted(excel_files, reverse=True):  # Eng yangisidan
                file_size = os.path.getsize(f) / 1024  # KB
                mod_time = datetime.fromtimestamp(os.path.getmtime(f))
                print(f"  • {os.path.basename(f)}")
                print(f"    └─ {mod_time.strftime('%Y-%m-%d %H:%M:%S')} ({file_size:.1f} KB)")
        else:
            print("\n📊 Excel Hisobotlar: Topilmadi")
        
        # Video fayllar
        video_dir = os.path.join(self.data_dir, "videos")
        if os.path.exists(video_dir):
            video_files = glob.glob(os.path.join(video_dir, "*.mp4"))
            if video_files:
                print(f"\n🎥 Video Fayllar ({len(video_files)} ta):")
                for f in sorted(video_files, reverse=True):  # Eng yangisidan
                    file_size = os.path.getsize(f) / (1024 * 1024)  # MB
                    mod_time = datetime.fromtimestamp(os.path.getmtime(f))
                    print(f"  • {os.path.basename(f)}")
                    print(f"    └─ {mod_time.strftime('%Y-%m-%d %H:%M:%S')} ({file_size:.1f} MB)")
            else:
                print("\n🎥 Video Fayllar: Topilmadi")
        else:
            print("\n🎥 Video Fayllar: Topilmadi")
        
        print("\n" + "=" * 60)
    
    def view_json_data(self, date=None):
        """JSON ma'lumotlarini ko'rsatish"""
        if date:
            json_file = os.path.join(self.data_dir, f"activities_{date}.json")
        else:
            # Eng yangi faylni topish
            json_files = glob.glob(os.path.join(self.data_dir, "activities_*.json"))
            if not json_files:
                print("❌ JSON fayllar topilmadi!")
                return
            json_file = max(json_files, key=os.path.getmtime)
        
        if not os.path.exists(json_file):
            print(f"❌ Fayl topilmadi: {json_file}")
            return
        
        print(f"\n📄 Ma'lumotlar: {os.path.basename(json_file)}")
        print("=" * 60)
        
        try:
            with open(json_file, 'r', encoding='utf-8') as f:
                data = json.load(f)
            
            if not data:
                print("Fayl bo'sh")
                return
            
            # Faollik turlarini guruhlash
            activities_by_type = {}
            for activity in data:
                activity_type = activity.get("type", "UNKNOWN")
                if activity_type not in activities_by_type:
                    activities_by_type[activity_type] = []
                activities_by_type[activity_type].append(activity)
            
            # Har bir tur bo'yicha ko'rsatish
            for activity_type, activities in activities_by_type.items():
                print(f"\n📋 {activity_type} ({len(activities)} marta):")
                print("-" * 60)
                
                for i, activity in enumerate(activities[:20], 1):  # Birinchi 20 tasi
                    timestamp = activity.get("timestamp", "Noma'lum")
                    if activity_type == "CRM_ACCESS":
                        window_title = activity.get("window_title", "Noma'lum")
                        count = activity.get("count", 0)
                        print(f"  {i}. [{timestamp}] {window_title} (#{count})")
                    elif activity_type == "PHONE_USAGE":
                        confidence = activity.get("confidence", 0)
                        count = activity.get("count", 0)
                        print(f"  {i}. [{timestamp}] Aniqlik: {confidence:.2f} (#{count})")
                    elif activity_type == "CLIENT_INTERACTION":
                        window_title = activity.get("window_title", "Noma'lum")
                        keyword = activity.get("keyword", "")
                        print(f"  {i}. [{timestamp}] {window_title} (kalit: {keyword})")
                    else:
                        print(f"  {i}. [{timestamp}] {activity}")
                
                if len(activities) > 20:
                    print(f"  ... va yana {len(activities) - 20} ta")
            
            print("\n" + "=" * 60)
            
        except Exception as e:
            print(f"❌ Xatolik: {e}")
    
    def view_excel_report(self, filename=None):
        """Excel hisobotni ko'rsatish"""
        if filename:
            excel_file = os.path.join(self.data_dir, filename)
        else:
            # Eng yangi faylni topish
            excel_files = glob.glob(os.path.join(self.data_dir, "activity_report_*.xlsx"))
            if not excel_files:
                print("❌ Excel fayllar topilmadi!")
                return
            excel_file = max(excel_files, key=os.path.getmtime)
        
        if not os.path.exists(excel_file):
            print(f"❌ Fayl topilmadi: {excel_file}")
            return
        
        print(f"\n📊 Excel Hisobot: {os.path.basename(excel_file)}")
        print("=" * 60)
        
        try:
            # Excel faylni o'qish
            excel_data = pd.ExcelFile(excel_file)
            
            print(f"\n📑 Jadvallar ({len(excel_data.sheet_names)} ta):")
            for sheet_name in excel_data.sheet_names:
                df = pd.read_excel(excel_file, sheet_name=sheet_name)
                print(f"\n  📋 {sheet_name} ({len(df)} qator):")
                print("-" * 60)
                
                # Birinchi 10 qatorni ko'rsatish
                if len(df) > 0:
                    print(df.head(10).to_string(index=False))
                    if len(df) > 10:
                        print(f"\n  ... va yana {len(df) - 10} qator")
                else:
                    print("  Bo'sh")
            
            print("\n" + "=" * 60)
            print(f"\n💡 To'liq ma'lumotlarni ko'rish uchun Excel faylni oching:")
            print(f"   {excel_file}")
            
        except Exception as e:
            print(f"❌ Xatolik: {e}")
    
    def show_statistics(self, date=None):
        """Umumiy statistika ko'rsatish"""
        print("\n" + "=" * 60)
        print("UMUMIY STATISTIKA")
        print("=" * 60)
        
        # Barcha JSON fayllarni o'qish
        json_files = glob.glob(os.path.join(self.data_dir, "activities_*.json"))
        if not json_files:
            print("❌ Ma'lumotlar topilmadi!")
            return
        
        all_activities = []
        for json_file in json_files:
            try:
                with open(json_file, 'r', encoding='utf-8') as f:
                    data = json.load(f)
                    all_activities.extend(data)
            except:
                continue
        
        if not all_activities:
            print("Ma'lumotlar topilmadi")
            return
        
        # Statistika hisoblash
        crm_count = sum(1 for a in all_activities if a.get("type") == "CRM_ACCESS")
        phone_count = sum(1 for a in all_activities if a.get("type") == "PHONE_USAGE")
        client_count = sum(1 for a in all_activities if a.get("type") == "CLIENT_INTERACTION")
        
        # Vaqt bo'yicha guruhlash
        dates = {}
        for activity in all_activities:
            timestamp = activity.get("timestamp", "")
            if timestamp:
                date_str = timestamp.split()[0]  # Faqat sana
                if date_str not in dates:
                    dates[date_str] = {"crm": 0, "phone": 0, "client": 0}
                activity_type = activity.get("type")
                if activity_type == "CRM_ACCESS":
                    dates[date_str]["crm"] += 1
                elif activity_type == "PHONE_USAGE":
                    dates[date_str]["phone"] += 1
                elif activity_type == "CLIENT_INTERACTION":
                    dates[date_str]["client"] += 1
        
        print(f"\n📊 JAMI STATISTIKA:")
        print(f"  • CRM ga kirishlar: {crm_count} marta")
        print(f"  • Telefon foydalanish: {phone_count} marta")
        print(f"  • Mijozlar bilan ishlash: {client_count} marta")
        print(f"  • Jami faolliklar: {len(all_activities)} marta")
        
        if dates:
            print(f"\n📅 KUNLIK STATISTIKA:")
            for date_str in sorted(dates.keys(), reverse=True)[:10]:  # Eng yangi 10 kuni
                stats = dates[date_str]
                print(f"  • {date_str}:")
                print(f"    - CRM: {stats['crm']} marta")
                print(f"    - Telefon: {stats['phone']} marta")
                print(f"    - Mijozlar: {stats['client']} marta")
        
        print("\n" + "=" * 60)
    
    def open_excel_file(self, filename=None):
        """Excel faylni ochish"""
        import subprocess
        import platform
        
        if filename:
            excel_file = os.path.join(self.data_dir, filename)
        else:
            excel_files = glob.glob(os.path.join(self.data_dir, "activity_report_*.xlsx"))
            if not excel_files:
                print("❌ Excel fayllar topilmadi!")
                return
            excel_file = max(excel_files, key=os.path.getmtime)
        
        if not os.path.exists(excel_file):
            print(f"❌ Fayl topilmadi: {excel_file}")
            return
        
        try:
            if platform.system() == 'Windows':
                os.startfile(excel_file)
            elif platform.system() == 'Darwin':  # macOS
                subprocess.call(['open', excel_file])
            else:  # Linux
                subprocess.call(['xdg-open', excel_file])
            print(f"✅ Excel fayl ochildi: {excel_file}")
        except Exception as e:
            print(f"❌ Xatolik: {e}")
            print(f"   Qo'lda oching: {excel_file}")


def main():
    """Asosiy funksiya"""
    print("=" * 60)
    print("MA'LUMOTLARNI KO'RISH VA TAHLIL QILISH")
    print("=" * 60)
    
    viewer = DataViewer()
    
    while True:
        print("\n📋 MENYU:")
        print("  1. Mavjud fayllarni ko'rsatish")
        print("  2. JSON ma'lumotlarini ko'rsatish (eng yangi)")
        print("  3. Excel hisobotni ko'rsatish (eng yangi)")
        print("  4. Umumiy statistika")
        print("  5. Excel faylni ochish (eng yangi)")
        print("  0. Chiqish")
        
        choice = input("\nTanlov kiriting (0-5): ").strip()
        
        if choice == "1":
            viewer.list_available_files()
        elif choice == "2":
            viewer.view_json_data()
        elif choice == "3":
            viewer.view_excel_report()
        elif choice == "4":
            viewer.show_statistics()
        elif choice == "5":
            viewer.open_excel_file()
        elif choice == "0":
            print("\nXayr!")
            break
        else:
            print("❌ Noto'g'ri tanlov!")


if __name__ == "__main__":
    main()

