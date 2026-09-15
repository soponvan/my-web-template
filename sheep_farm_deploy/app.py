#!/usr/bin/env python3
# -*- coding: utf-8 -*-
# โปรแกรมบริหารจัดการฟาร์มแกะ - Web App
# Sheep Farm Management System - Web Application

from flask import Flask, render_template, request, jsonify, send_file, session, redirect, url_for
from functools import wraps
from werkzeug.security import generate_password_hash, check_password_hash
import json
import os
import secrets
import tempfile
import threading
from datetime import datetime
from pathlib import Path
from io import BytesIO
try:
    import psycopg2
    from psycopg2.extras import Json
except ImportError:
    psycopg2 = None
    Json = None
from reportlab.lib.pagesizes import letter, A4
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.lib.units import inch
from reportlab.platypus import SimpleDocTemplate, Table, TableStyle, Paragraph, Spacer, PageBreak
from reportlab.lib import colors
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from docx import Document
from docx.shared import Pt, RGBColor, Inches
from docx.enum.text import WD_ALIGN_PARAGRAPH

app = Flask(__name__)
app.config['JSON_AS_ASCII'] = False
app.secret_key = os.environ.get('SECRET_KEY', secrets.token_hex(32))

# ไฟล์เก็บข้อมูล
BASE_DIR = Path(__file__).resolve().parent
DATA_DIR = BASE_DIR / "sheep_farm_data"
SHEEP_FILE = DATA_DIR / "sheep_records.json"
FEEDING_FILE = DATA_DIR / "feeding_records.json"
HEALTH_FILE = DATA_DIR / "health_records.json"
BREEDING_FILE = DATA_DIR / "breeding_records.json"
FEED_STOCK_FILE = DATA_DIR / "feed_stock.json"
USERS_FILE = DATA_DIR / "users.json"
DATABASE_URL = os.environ.get('DATABASE_URL')

# ระบบสิทธิ์: admin และ user ทั่วไป
ROLE_PERMISSIONS = {
    'admin': ['read', 'write', 'delete', 'manage_users'],
    'user': ['read', 'write']
}

def login_required(f):
    @wraps(f)
    def decorated_function(*args, **kwargs):
        if 'user_id' not in session:
            return jsonify({'status': 'error', 'message': 'กรุณาเข้าสู่ระบบก่อน', 'redirect': '/login'}), 401
        return f(*args, **kwargs)
    return decorated_function

def admin_required(f):
    @wraps(f)
    def decorated_function(*args, **kwargs):
        if 'user_id' not in session:
            return jsonify({'status': 'error', 'message': 'กรุณาเข้าสู่ระบบก่อน', 'redirect': '/login'}), 401
        if session.get('role') != 'admin':
            return jsonify({'status': 'error', 'message': 'คุณไม่มีสิทธิ์ในการดำเนินการนี้'}), 403
        return f(*args, **kwargs)
    return decorated_function

def permission_required(permission):
    def decorator(f):
        @wraps(f)
        def decorated_function(*args, **kwargs):
            if 'user_id' not in session:
                return jsonify({'status': 'error', 'message': 'กรุณาเข้าสู่ระบบก่อน', 'redirect': '/login'}), 401
            role = session.get('role', 'user')
            if permission not in ROLE_PERMISSIONS.get(role, []):
                return jsonify({'status': 'error', 'message': f'คุณไม่มีสิทธิ์: {permission}'}), 403
            return f(*args, **kwargs)
        return decorated_function
    return decorator

class SheepFarmManager:
    def __init__(self):
        self._save_lock = threading.RLock()
        self.create_data_directory()
        if DATABASE_URL:
            self.initialize_database()
        self.load_all_data()
    
    def create_data_directory(self):
        DATA_DIR.mkdir(exist_ok=True)

    def get_database_connection(self):
        if psycopg2 is None:
            raise RuntimeError('DATABASE_URL requires psycopg2-binary to be installed')
        return psycopg2.connect(DATABASE_URL)

    def initialize_database(self):
        with self.get_database_connection() as connection:
            with connection.cursor() as cursor:
                cursor.execute('''
                    CREATE TABLE IF NOT EXISTS farm_data (
                        data_key TEXT PRIMARY KEY,
                        data JSONB NOT NULL,
                        updated_at TIMESTAMPTZ NOT NULL DEFAULT NOW()
                    )
                ''')
    
    def load_all_data(self):
        self.sheep_records = self.load_json(SHEEP_FILE)
        self.feeding_records = self.load_json(FEEDING_FILE)
        self.health_records = self.load_json(HEALTH_FILE)
        self.breeding_records = self.load_json(BREEDING_FILE)
        self.feed_stock = self.load_json(FEED_STOCK_FILE)
    
    def load_json(self, filepath):
        if DATABASE_URL:
            with self.get_database_connection() as connection:
                with connection.cursor() as cursor:
                    cursor.execute('SELECT data FROM farm_data WHERE data_key = %s', (filepath.name,))
                    row = cursor.fetchone()
                    if row:
                        return row[0]

            if filepath.exists():
                data = self.load_local_json(filepath)
                if data:
                    self.save_json(filepath, data)
                return data
            return {}

        return self.load_local_json(filepath)

    def load_local_json(self, filepath):
        if filepath.exists():
            try:
                with open(filepath, 'r', encoding='utf-8') as f:
                    return json.load(f)
            except:
                return {}
        return {}
    
    def save_json(self, filepath, data):
        if DATABASE_URL:
            with self._save_lock:
                with self.get_database_connection() as connection:
                    with connection.cursor() as cursor:
                        cursor.execute('''
                            INSERT INTO farm_data (data_key, data, updated_at)
                            VALUES (%s, %s, NOW())
                            ON CONFLICT (data_key) DO UPDATE SET
                                data = EXCLUDED.data,
                                updated_at = NOW()
                        ''', (filepath.name, Json(data)))
            return

        filepath.parent.mkdir(parents=True, exist_ok=True)
        temp_path = None
        with self._save_lock:
            try:
                with tempfile.NamedTemporaryFile(
                    mode='w', encoding='utf-8', dir=filepath.parent,
                    prefix=f'.{filepath.name}.', suffix='.tmp', delete=False
                ) as temp_file:
                    temp_path = Path(temp_file.name)
                    json.dump(data, temp_file, ensure_ascii=False, indent=2)
                    temp_file.flush()
                    os.fsync(temp_file.fileno())
                os.replace(temp_path, filepath)
            finally:
                if temp_path and temp_path.exists():
                    temp_path.unlink()
    
    def register_sheep(self, sheep_id, gender, dob, breed, weight, animal_type='แกะ', tag_color='-'):
        if sheep_id in self.sheep_records:
            return {"status": "error", "message": f"{animal_type or 'แกะ'} ID นี้มีอยู่แล้ว!"}
        
        self.sheep_records[sheep_id] = {
            "id": sheep_id,
            "animal_type": animal_type or "แกะ",
            "tag_color": tag_color or "-",
            "gender": gender,
            "dob": dob,
            "breed": breed,
            "weight": weight,
            "status": "ปกติ",
            "registered_date": datetime.now().strftime("%Y-%m-%d")
        }
        
        self.save_json(SHEEP_FILE, self.sheep_records)
        return {"status": "success", "message": f"บันทึก{animal_type or 'แกะ'} {sheep_id} สำเร็จ!"}
    
    def get_all_sheep(self):
        records = []
        for s in self.sheep_records.values():
            item = dict(s)
            if 'animal_type' not in item or not item['animal_type']:
                item['animal_type'] = 'แกะ'
            if 'tag_color' not in item or not item['tag_color']:
                item['tag_color'] = '-'
            records.append(item)
        return records
    
    def edit_sheep(self, sheep_id, gender=None, dob=None, breed=None, weight=None, status=None, death_date=None, animal_type=None, tag_color=None):
        if sheep_id not in self.sheep_records:
            return {"status": "error", "message": "ไม่พบข้อมูลสัตว์ ID นี้!"}
        
        if animal_type:
            self.sheep_records[sheep_id]["animal_type"] = animal_type
        elif "animal_type" not in self.sheep_records[sheep_id]:
            self.sheep_records[sheep_id]["animal_type"] = "แกะ"
            
        if tag_color is not None:
            self.sheep_records[sheep_id]["tag_color"] = tag_color or "-"
            
        if gender:
            self.sheep_records[sheep_id]["gender"] = gender
        if dob:
            self.sheep_records[sheep_id]["dob"] = dob
        if breed:
            self.sheep_records[sheep_id]["breed"] = breed
        if weight:
            self.sheep_records[sheep_id]["weight"] = weight
        if status:
            self.sheep_records[sheep_id]["status"] = status
        
        if status in ["ตาย", "หาย"] and death_date:
            self.sheep_records[sheep_id]["death_date"] = death_date
        elif status not in ["ตาย", "หาย"] and "death_date" in self.sheep_records[sheep_id]:
            # Clear death date if status changes back
            del self.sheep_records[sheep_id]["death_date"]
        
        self.save_json(SHEEP_FILE, self.sheep_records)
        curr_type = self.sheep_records[sheep_id].get("animal_type", "แกะ")
        return {"status": "success", "message": f"แก้ไขข้อมูล{curr_type} {sheep_id} สำเร็จ!"}
    
    def delete_sheep(self, sheep_id):
        if sheep_id not in self.sheep_records:
            return {"status": "error", "message": "ไม่พบข้อมูลสัตว์ ID นี้!"}
        
        curr_type = self.sheep_records[sheep_id].get("animal_type", "แกะ")
        # Delete sheep record
        del self.sheep_records[sheep_id]
        
        # Delete health records for this sheep
        if sheep_id in self.health_records:
            del self.health_records[sheep_id]
        
        # Save changes
        self.save_json(SHEEP_FILE, self.sheep_records)
        self.save_json(HEALTH_FILE, self.health_records)
        
        return {"status": "success", "message": f"ลบ{curr_type} {sheep_id} สำเร็จ!"}
    
    def record_feeding(self, sheep_ids, feed_type, quantity, notes):
        feeding_record = {
            "date": datetime.now().strftime("%Y-%m-%d %H:%M"),
            "sheep_ids": sheep_ids,
            "feed_type": feed_type,
            "quantity": quantity,
            "notes": notes
        }
        
        record_id = f"feeding_{datetime.now().strftime('%Y%m%d_%H%M%S_%f')}_{secrets.token_hex(4)}"
        self.feeding_records[record_id] = feeding_record
        
        self.save_json(FEEDING_FILE, self.feeding_records)
        return {"status": "success", "message": "บันทึกการให้อาหารสำเร็จ!"}
    
    def get_feeding_history(self):
        records = list(self.feeding_records.items())[-10:]
        return [{"id": k, "data": v} for k, v in records]
    
    def edit_feeding(self, record_id, feed_type=None, quantity=None, notes=None):
        if record_id not in self.feeding_records:
            return {"status": "error", "message": "ไม่พบประวัติการให้อาหารนี้!"}
        
        if feed_type:
            self.feeding_records[record_id]["feed_type"] = feed_type
        if quantity:
            self.feeding_records[record_id]["quantity"] = quantity
        if notes:
            self.feeding_records[record_id]["notes"] = notes
        
        self.save_json(FEEDING_FILE, self.feeding_records)
        return {"status": "success", "message": "แก้ไขประวัติการให้อาหารสำเร็จ!"}
    
    def delete_feeding(self, record_id):
        if record_id not in self.feeding_records:
            return {"status": "error", "message": "ไม่พบประวัติการให้อาหารนี้!"}
        
        del self.feeding_records[record_id]
        self.save_json(FEEDING_FILE, self.feeding_records)
        return {"status": "success", "message": "ลบประวัติการให้อาหารสำเร็จ!"}
    
    def record_health(self, sheep_id, health_type, status=None, symptoms=None, temperature=None, vaccine_name=None, injection_number=None, vaccine_date=None, next_dose=None, death_date=None):
        if sheep_id not in self.sheep_records:
            return {"status": "error", "message": "ไม่พบสัตว์ ID นี้!"}
        
        if health_type == "examination":
            health_record = {
                "date": datetime.now().strftime("%Y-%m-%d"),
                "type": "ตรวจสุขภาพ",
                "status": status,
                "symptoms": symptoms,
                "temperature": temperature,
                "death_date": death_date if status in ["ตาย", "หาย"] else None
            }
            self.sheep_records[sheep_id]["status"] = status
            if status in ["ตาย", "หาย"] and death_date:
                self.sheep_records[sheep_id]["death_date"] = death_date
            elif status not in ["ตาย", "หาย"] and "death_date" in self.sheep_records[sheep_id]:
                del self.sheep_records[sheep_id]["death_date"]
        else:
            health_record = {
                "date": vaccine_date,
                "type": "วัคซีน",
                "vaccine_name": vaccine_name,
                "injection_number": injection_number,
                "next_dose_date": next_dose
            }
        
        if sheep_id not in self.health_records:
            self.health_records[sheep_id] = []
        
        self.health_records[sheep_id].append(health_record)
        
        self.save_json(HEALTH_FILE, self.health_records)
        self.save_json(SHEEP_FILE, self.sheep_records)
        return {"status": "success", "message": "บันทึกสุขภาพสำเร็จ!"}
    
    def get_health_history(self, sheep_id):
        if sheep_id not in self.health_records:
            return []
        return self.health_records[sheep_id]
    
    def edit_health(self, sheep_id, record_index, health_type, status=None, symptoms=None, temperature=None, vaccine_name=None, injection_number=None, vaccine_date=None, next_dose=None, death_date=None):
        if sheep_id not in self.health_records or record_index < 0 or record_index >= len(self.health_records[sheep_id]):
            return {"status": "error", "message": "ไม่พบประวัติสุขภาพนี้!"}
        
        record = self.health_records[sheep_id][record_index]
        
        if health_type == "examination":
            if status:
                record["status"] = status
                if status in ["ตาย", "หาย"] and death_date:
                    record["death_date"] = death_date
                    self.sheep_records[sheep_id]["death_date"] = death_date
                elif status not in ["ตาย", "หาย"] and "death_date" in record:
                    del record["death_date"]
                    if "death_date" in self.sheep_records[sheep_id]:
                        del self.sheep_records[sheep_id]["death_date"]
            if symptoms is not None:
                record["symptoms"] = symptoms
            if temperature:
                record["temperature"] = temperature
            if status:
                self.sheep_records[sheep_id]["status"] = status
        else:
            if vaccine_name:
                record["vaccine_name"] = vaccine_name
            if injection_number:
                record["injection_number"] = injection_number
            if vaccine_date:
                record["date"] = vaccine_date
            if next_dose:
                record["next_dose_date"] = next_dose
        
        self.save_json(HEALTH_FILE, self.health_records)
        self.save_json(SHEEP_FILE, self.sheep_records)
        return {"status": "success", "message": "แก้ไขประวัติสุขภาพสำเร็จ!"}
    
    def delete_health(self, sheep_id, record_index):
        if sheep_id not in self.health_records or record_index < 0 or record_index >= len(self.health_records[sheep_id]):
            return {"status": "error", "message": "ไม่พบประวัติสุขภาพนี้!"}
        
        del self.health_records[sheep_id][record_index]
        
        # If no more records for this sheep, remove the sheep key
        if len(self.health_records[sheep_id]) == 0:
            del self.health_records[sheep_id]
        
        self.save_json(HEALTH_FILE, self.health_records)
        return {"status": "success", "message": "ลบประวัติสุขภาพสำเร็จ!"}
    
    def plan_breeding(self, female_id, male_id, breeding_date, expected_date):
        if female_id not in self.sheep_records or male_id not in self.sheep_records:
            return {"status": "error", "message": "ไม่พบสัตว์ ID นี้!"}
        
        if self.sheep_records[female_id]["gender"] != "ตัวเมีย":
            return {"status": "error", "message": "ต้องเป็นสัตว์ตัวเมีย!"}
        
        breeding_record = {
            "female_id": female_id,
            "male_id": male_id,
            "breeding_date": breeding_date,
            "expected_date": expected_date,
            "status": "รอติดตาม",
            "lambs": []
        }
        
        record_id = f"breeding_{female_id}_{datetime.now().strftime('%Y%m%d_%H%M%S_%f')}_{secrets.token_hex(4)}"
        self.breeding_records[record_id] = breeding_record
        
        self.save_json(BREEDING_FILE, self.breeding_records)
        return {"status": "success", "message": f"บันทึกการผสมพันธุ์สำเร็จ!"}
    
    def edit_breeding_plan(self, record_id, female_id=None, male_id=None, breeding_date=None, expected_date=None):
        if record_id not in self.breeding_records:
            return {"status": "error", "message": "ไม่พบแผนผสมพันธุ์นี้!"}
        
        if female_id:
            if female_id not in self.sheep_records:
                return {"status": "error", "message": "ไม่พบสัตว์ตัวเมีย ID นี้!"}
            if self.sheep_records[female_id]["gender"] != "ตัวเมีย":
                return {"status": "error", "message": "ต้องเป็นสัตว์ตัวเมีย!"}
            self.breeding_records[record_id]["female_id"] = female_id
            
        if male_id:
            if male_id not in self.sheep_records:
                return {"status": "error", "message": "ไม่พบสัตว์ตัวผู้ ID นี้!"}
            self.breeding_records[record_id]["male_id"] = male_id
            
        if breeding_date:
            self.breeding_records[record_id]["breeding_date"] = breeding_date
        if expected_date:
            self.breeding_records[record_id]["expected_date"] = expected_date
            
        self.save_json(BREEDING_FILE, self.breeding_records)
        return {"status": "success", "message": "แก้ไขแผนผสมพันธุ์สำเร็จ!"}
    
    def get_pending_breeding(self):
        return [{"id": k, "data": v} for k, v in self.breeding_records.items() if v['status'] == "รอติดตาม"]
    
    def record_lambing(self, record_id, lambing_date, lambs_data):
        if record_id not in self.breeding_records:
            return {"status": "error", "message": "ไม่พบ Record ID นี้!"}
        
        mother = self.sheep_records.get(self.breeding_records[record_id]['female_id'], {})
        mother_animal_type = mother.get('animal_type', 'แกะ')
        
        for lamb in lambs_data:
            self.sheep_records[lamb["id"]] = {
                "id": lamb["id"],
                "animal_type": mother_animal_type,
                "gender": lamb["gender"],
                "dob": lambing_date,
                "breed": mother.get('breed', 'ไม่ระบุ'),
                "weight": lamb["weight"],
                "status": "ปกติ",
                "registered_date": datetime.now().strftime("%Y-%m-%d"),
                "mother": self.breeding_records[record_id]['female_id'],
                "father": self.breeding_records[record_id]['male_id']
            }
        
        self.breeding_records[record_id]["status"] = "คลอดแล้ว"
        self.breeding_records[record_id]["lambing_date"] = lambing_date
        self.breeding_records[record_id]["lambs"] = lambs_data
        
        self.save_json(SHEEP_FILE, self.sheep_records)
        self.save_json(BREEDING_FILE, self.breeding_records)
        return {"status": "success", "message": f"บันทึกการคลอดสำเร็จ!"}
    
    def edit_lambing(self, record_id, lambing_date=None, lambs_data=None):
        if record_id not in self.breeding_records:
            return {"status": "error", "message": "ไม่พบ Record ID นี้!"}
        
        mother = self.sheep_records.get(self.breeding_records[record_id]['female_id'], {})
        mother_animal_type = mother.get('animal_type', 'แกะ')
        
        if lambing_date:
            self.breeding_records[record_id]["lambing_date"] = lambing_date
        
        if lambs_data:
            # Delete old lamb records first
            old_lambs = self.breeding_records[record_id].get('lambs', [])
            for old_lamb in old_lambs:
                if old_lamb["id"] in self.sheep_records:
                    del self.sheep_records[old_lamb["id"]]
            
            # Add new lamb records
            for lamb in lambs_data:
                self.sheep_records[lamb["id"]] = {
                    "id": lamb["id"],
                    "animal_type": mother_animal_type,
                    "gender": lamb["gender"],
                    "dob": lambing_date if lambing_date else self.breeding_records[record_id].get("lambing_date"),
                    "breed": mother.get('breed', 'ไม่ระบุ'),
                    "weight": lamb["weight"],
                    "status": "ปกติ",
                    "registered_date": datetime.now().strftime("%Y-%m-%d"),
                    "mother": self.breeding_records[record_id]['female_id'],
                    "father": self.breeding_records[record_id]['male_id']
                }
            
            self.breeding_records[record_id]["lambs"] = lambs_data
        
        self.save_json(SHEEP_FILE, self.sheep_records)
        self.save_json(BREEDING_FILE, self.breeding_records)
        return {"status": "success", "message": "แก้ไขบันทึกการคลอดสำเร็จ!"}
    
    def delete_lambing(self, record_id):
        if record_id not in self.breeding_records:
            return {"status": "error", "message": "ไม่พบ Record ID นี้!"}
        
        # Delete lamb records from sheep_records
        lambs = self.breeding_records[record_id].get('lambs', [])
        for lamb in lambs:
            if lamb["id"] in self.sheep_records:
                del self.sheep_records[lamb["id"]]
        
        # Delete breeding record
        del self.breeding_records[record_id]
        
        self.save_json(SHEEP_FILE, self.sheep_records)
        self.save_json(BREEDING_FILE, self.breeding_records)
        return {"status": "success", "message": "ลบบันทึกการคลอดสำเร็จ!"}
    
    def get_feed_stock(self):
        return self.feed_stock
    
    def add_feed_stock(self, feed_type, quantity, min_level):
        self.feed_stock[feed_type] = {
            "quantity": quantity,
            "min_level": min_level,
            "unit": "กท.",
            "last_updated": datetime.now().strftime("%Y-%m-%d")
        }
        self.save_json(FEED_STOCK_FILE, self.feed_stock)
        return {"status": "success", "message": "บันทึกข้อมูลอาหารสำเร็จ!"}
    
    def update_feed_stock(self, feed_type, operation, quantity):
        if feed_type not in self.feed_stock:
            return {"status": "error", "message": "ไม่พบประเภทอาหารนี้!"}
        
        current = float(self.feed_stock[feed_type]['quantity'])
        
        if operation == "+":
            new_quantity = current + float(quantity)
        elif operation == "-":
            new_quantity = current - float(quantity)
        else:
            return {"status": "error", "message": "ตัวการอัปเดตไม่ถูกต้อง!"}
        
        self.feed_stock[feed_type]['quantity'] = str(new_quantity)
        self.feed_stock[feed_type]['last_updated'] = datetime.now().strftime("%Y-%m-%d")
        
        self.save_json(FEED_STOCK_FILE, self.feed_stock)
        return {"status": "success", "message": f"อัปเดตปริมาณสำเร็จ! ปริมาณใหม่: {new_quantity} กท."}
    
    def generate_report(self):
        # สัตว์ที่ยังมีชีวิต (ไม่ตาย/หาย)
        active_animals = [s for s in self.sheep_records.values() if s.get('status') not in ['ตาย', 'หาย']]
        total_animals = len(active_animals)
        
        # แกะ
        sheep_active = [s for s in active_animals if s.get('animal_type', 'แกะ') == 'แกะ']
        total_sheep = len(sheep_active)
        male_sheep = sum(1 for s in sheep_active if s.get('gender') == 'ตัวผู้')
        female_sheep = sum(1 for s in sheep_active if s.get('gender') == 'ตัวเมีย')
        
        # แพะ
        goat_active = [s for s in active_animals if s.get('animal_type') == 'แพะ']
        total_goats = len(goat_active)
        male_goats = sum(1 for s in goat_active if s.get('gender') == 'ตัวผู้')
        female_goats = sum(1 for s in goat_active if s.get('gender') == 'ตัวเมีย')
        
        # สถานะภาพรวม
        healthy = sum(1 for s in self.sheep_records.values() if s.get('status') == 'ปกติ')
        sick = sum(1 for s in self.sheep_records.values() if s.get('status') == 'ป่วย')
        dead = sum(1 for s in self.sheep_records.values() if s.get('status') == 'ตาย')
        missing = sum(1 for s in self.sheep_records.values() if s.get('status') == 'หาย')
        
        # สถิติการตาย/หาย แยกตามชนิด
        dead_sheep = sum(1 for s in self.sheep_records.values() if s.get('animal_type', 'แกะ') == 'แกะ' and s.get('status') == 'ตาย')
        missing_sheep = sum(1 for s in self.sheep_records.values() if s.get('animal_type', 'แกะ') == 'แกะ' and s.get('status') == 'หาย')
        dead_goats = sum(1 for s in self.sheep_records.values() if s.get('animal_type') == 'แพะ' and s.get('status') == 'ตาย')
        missing_goats = sum(1 for s in self.sheep_records.values() if s.get('animal_type') == 'แพะ' and s.get('status') == 'หาย')
        
        total_breeding = len(self.breeding_records)
        successful_breeding = sum(1 for b in self.breeding_records.values() if b.get('status') == 'คลอดแล้ว')
        
        # ลูกสัตว์เกิดใหม่
        total_lambs = 0
        total_kids = 0
        for b in self.breeding_records.values():
            mother = self.sheep_records.get(b.get('female_id', ''), {})
            m_type = mother.get('animal_type', 'แกะ')
            cnt = len(b.get('lambs', []))
            if m_type == 'แพะ':
                total_kids += cnt
            else:
                total_lambs += cnt
        total_offspring = total_lambs + total_kids
        
        active_health_records = {k: v for k, v in self.health_records.items() if self.sheep_records.get(k, {}).get("status") not in ["ตาย", "หาย"]}
        vaccine_records = [
            {
                "animal_id": animal_id,
                "animal_type": self.sheep_records.get(animal_id, {}).get('animal_type', 'แกะ'),
                "vaccine_name": record.get('vaccine_name', '-'),
                "injection_number": record.get('injection_number', '-'),
                "date": record.get('date', '-'),
                "next_dose_date": record.get('next_dose_date', '-')
            }
            for animal_id, records in active_health_records.items()
            for record in records
            if record.get('type') == 'วัคซีน'
        ]

        return {
            "total_animals": total_animals,
            "total_sheep": total_sheep,
            "male_sheep": male_sheep,
            "female_sheep": female_sheep,
            "total_goats": total_goats,
            "male_goats": male_goats,
            "female_goats": female_goats,
            "healthy": healthy,
            "sick": sick,
            "dead": dead,
            "missing": missing,
            "dead_sheep": dead_sheep,
            "missing_sheep": missing_sheep,
            "dead_goats": dead_goats,
            "missing_goats": missing_goats,
            "total_breeding": total_breeding,
            "successful_breeding": successful_breeding,
            "total_lambs": total_lambs,
            "total_kids": total_kids,
            "total_offspring": total_offspring,
            "feeding_records": len(self.feeding_records),
            "health_exams": sum(1 for records in active_health_records.values() for r in records if r.get('type') == 'ตรวจสุขภาพ'),
            "vaccines": sum(1 for records in active_health_records.values() for r in records if r.get('type') == 'วัคซีน'),
            "vaccine_records": vaccine_records,
            "feed_stock": self.feed_stock,
            "date": datetime.now().strftime("%Y-%m-%d")
        }
    
    def export_to_pdf(self):
        """Export full report to PDF with comprehensive details"""
        try:
            buffer = BytesIO()
            doc = SimpleDocTemplate(buffer, pagesize=A4, encoding='utf-8', topMargin=0.5*inch, bottomMargin=0.5*inch)
            story = []
            
            styles = getSampleStyleSheet()
            title_style = ParagraphStyle(
                'CustomTitle',
                parent=styles['Heading1'],
                fontSize=20,
                textColor=colors.HexColor('#06b6d4'),
                spaceAfter=10,
                alignment=1,
                fontName='Helvetica-Bold'
            )
            
            heading_style = ParagraphStyle(
                'CustomHeading',
                parent=styles['Heading2'],
                fontSize=12,
                textColor=colors.HexColor('#06b6d4'),
                spaceAfter=8,
                spaceBefore=8,
                fontName='Helvetica-Bold'
            )
            
            subheading_style = ParagraphStyle(
                'SubHeading',
                parent=styles['Heading3'],
                fontSize=10,
                textColor=colors.HexColor('#0d9488'),
                spaceAfter=6,
                spaceBefore=6
            )
            
            normal_style = ParagraphStyle(
                'CustomNormal',
                parent=styles['Normal'],
                fontSize=9,
                spaceAfter=4
            )
            
            # Title
            story.append(Paragraph("รายงานบริหารจัดการฟาร์มแกะและแพะ กองบิน 23", title_style))
            story.append(Paragraph(f"สร้างเมื่อ: {datetime.now().strftime('%d/%m/%Y เวลา %H:%M:%S น.')}", normal_style))
            story.append(Spacer(1, 0.2*inch))
            
            report_data = self.generate_report()
            
            # Section 1: สถิติแกะและแพะ
            story.append(Paragraph("📊 1. สถิติแกะและแพะ", heading_style))
            sheep_stats = [
                ["หมวดหมู่", "จำนวน"],
                ["สัตว์ทั้งหมด", str(report_data['total_animals'])],
                ["แกะทั้งหมด", f"{report_data['total_sheep']} ตัว (ผู้ {report_data['male_sheep']} / เมีย {report_data['female_sheep']})"],
                ["แพะทั้งหมด", f"{report_data['total_goats']} ตัว (ผู้ {report_data['male_goats']} / เมีย {report_data['female_goats']})"],
                ["ปกติ", str(report_data['healthy'])],
                ["ป่วย", str(report_data['sick'])],
                ["ตาย", str(report_data['dead'])],
                ["หาย", str(report_data['missing'])]
            ]
            
            sheep_table = Table(sheep_stats, colWidths=[3*inch, 2.5*inch])
            sheep_table.setStyle(TableStyle([
                ('BACKGROUND', (0, 0), (-1, 0), colors.HexColor('#06b6d4')),
                ('TEXTCOLOR', (0, 0), (-1, 0), colors.whitesmoke),
                ('ALIGN', (0, 0), (-1, -1), 'CENTER'),
                ('FONTNAME', (0, 0), (-1, 0), 'Helvetica-Bold'),
                ('FONTSIZE', (0, 0), (-1, 0), 11),
                ('BOTTOMPADDING', (0, 0), (-1, 0), 12),
                ('BACKGROUND', (0, 1), (-1, -1), colors.lightcyan),
                ('GRID', (0, 0), (-1, -1), 1, colors.black),
                ('FONTSIZE', (0, 1), (-1, -1), 10)
            ]))
            story.append(sheep_table)
            story.append(Spacer(1, 0.15*inch))
            
            # Detailed Sheep List
            if self.sheep_records:
                story.append(Paragraph("รายชื่อสัตว์แต่ละตัว (แกะและแพะ)", subheading_style))
                sheep_list_data = [["รหัส ID", "ชนิด", "สีแท็กติดหูแกะ", "เพศ", "สายพันธุ์", "น้ำหนัก", "สถานะ"]]
                for sheep in list(self.sheep_records.values())[:15]:  # Show first 15
                    sheep_list_data.append([
                        sheep.get('id', '-'),
                        sheep.get('animal_type', 'แกะ'),
                        sheep.get('tag_color', '-'),
                        sheep.get('gender', '-'),
                        sheep.get('breed', '-'),
                        f"{sheep.get('weight', '-')} กก.",
                        sheep.get('status', '-')
                    ])
                
                sheep_list_table = Table(sheep_list_data, colWidths=[1.0*inch, 0.8*inch, 0.9*inch, 0.8*inch, 1.2*inch, 0.9*inch, 0.8*inch])
                sheep_list_table.setStyle(TableStyle([
                    ('BACKGROUND', (0, 0), (-1, 0), colors.HexColor('#0d9488')),
                    ('TEXTCOLOR', (0, 0), (-1, 0), colors.whitesmoke),
                    ('ALIGN', (0, 0), (-1, -1), 'CENTER'),
                    ('FONTSIZE', (0, 0), (-1, -1), 8),
                    ('BOTTOMPADDING', (0, 0), (-1, 0), 8),
                    ('BACKGROUND', (0, 1), (-1, -1), colors.lightgrey),
                    ('GRID', (0, 0), (-1, -1), 1, colors.black)
                ]))
                story.append(sheep_list_table)
                story.append(Spacer(1, 0.15*inch))
            
            # Section 2: การผสมพันธุ์
            story.append(Paragraph("👶 2. สถิติการผสมพันธุ์", heading_style))
            breeding_stats = [
                ["หมวดหมู่", "จำนวน"],
                ["แผนผสมพันธุ์ทั้งหมด", str(report_data['total_breeding'])],
                ["ผสมพันธุ์สำเร็จ", str(report_data['successful_breeding'])],
                ["ลูกแกะทั้งหมด", str(report_data['total_lambs'])],
                ["ลูกแพะทั้งหมด", str(report_data.get('total_kids', 0))]
            ]
            
            breeding_table = Table(breeding_stats, colWidths=[3*inch, 2*inch])
            breeding_table.setStyle(TableStyle([
                ('BACKGROUND', (0, 0), (-1, 0), colors.HexColor('#0d9488')),
                ('TEXTCOLOR', (0, 0), (-1, 0), colors.whitesmoke),
                ('ALIGN', (0, 0), (-1, -1), 'CENTER'),
                ('FONTNAME', (0, 0), (-1, 0), 'Helvetica-Bold'),
                ('FONTSIZE', (0, 0), (-1, 0), 11),
                ('BOTTOMPADDING', (0, 0), (-1, 0), 12),
                ('BACKGROUND', (0, 1), (-1, -1), colors.lightgreen),
                ('GRID', (0, 0), (-1, -1), 1, colors.black),
                ('FONTSIZE', (0, 1), (-1, -1), 10)
            ]))
            story.append(breeding_table)
            story.append(Spacer(1, 0.15*inch))
            
            # Section 3: สุขภาพและอาหาร
            story.append(Paragraph("💉 3. สถิติสุขภาพและอาหาร", heading_style))
            health_stats = [
                ["หมวดหมู่", "จำนวน"],
                ["ตรวจสุขภาพ", str(report_data['health_exams'])],
                ["ฉีดวัคซีน", str(report_data['vaccines'])],
                ["บันทึกการให้อาหาร", str(report_data['feeding_records'])]
            ]
            
            health_table = Table(health_stats, colWidths=[3*inch, 2*inch])
            health_table.setStyle(TableStyle([
                ('BACKGROUND', (0, 0), (-1, 0), colors.HexColor('#1399bd')),
                ('TEXTCOLOR', (0, 0), (-1, 0), colors.whitesmoke),
                ('ALIGN', (0, 0), (-1, -1), 'CENTER'),
                ('FONTNAME', (0, 0), (-1, 0), 'Helvetica-Bold'),
                ('FONTSIZE', (0, 0), (-1, 0), 11),
                ('BOTTOMPADDING', (0, 0), (-1, 0), 12),
                ('BACKGROUND', (0, 1), (-1, -1), colors.lightyellow),
                ('GRID', (0, 0), (-1, -1), 1, colors.black),
                ('FONTSIZE', (0, 1), (-1, -1), 10)
            ]))
            story.append(health_table)
            story.append(Spacer(1, 0.15*inch))

            vaccine_list = [["รหัสสัตว์", "ชนิด", "วัคซีน", "เข็มที่", "วันที่ฉีด", "ครั้งถัดไป"]]
            for vaccine in report_data.get('vaccine_records', []):
                vaccine_list.append([
                    vaccine.get('animal_id', '-'),
                    vaccine.get('animal_type', '-'),
                    vaccine.get('vaccine_name', '-'),
                    vaccine.get('injection_number', '-'),
                    vaccine.get('date', '-'),
                    vaccine.get('next_dose_date', '-')
                ])
            if len(vaccine_list) > 1:
                story.append(Paragraph("รายการวัคซีน", subheading_style))
                vaccine_table = Table(vaccine_list, colWidths=[0.85*inch, 0.7*inch, 1.25*inch, 0.65*inch, 1.0*inch, 1.0*inch])
                vaccine_table.setStyle(TableStyle([
                    ('BACKGROUND', (0, 0), (-1, 0), colors.HexColor('#1399bd')),
                    ('TEXTCOLOR', (0, 0), (-1, 0), colors.whitesmoke),
                    ('ALIGN', (0, 0), (-1, -1), 'CENTER'),
                    ('FONTSIZE', (0, 0), (-1, -1), 8),
                    ('GRID', (0, 0), (-1, -1), 1, colors.black)
                ]))
                story.append(vaccine_table)
                story.append(Spacer(1, 0.15*inch))
            
            # Section 4: รายการแกะและแพะจำหน่าย (ตาย/หาย)
            story.append(Paragraph("⚰️ 4. รายการสัตว์จำหน่าย (ตาย/หาย)", heading_style))
            dead_missing_data = [["รหัส ID", "ชนิด", "เพศ", "สถานะ", "วันที่ตาย/หาย"]]
            for sheep in self.sheep_records.values():
                if sheep.get('status') in ['ตาย', 'หาย']:
                    dead_missing_data.append([
                        sheep.get('id', '-'),
                        sheep.get('animal_type', 'แกะ'),
                        sheep.get('gender', '-'),
                        sheep.get('status', '-'),
                        sheep.get('death_date', '-')
                    ])
            
            if len(dead_missing_data) > 1:
                dm_table = Table(dead_missing_data, colWidths=[1.1*inch, 0.9*inch, 0.9*inch, 1.1*inch, 1.6*inch])
                dm_table.setStyle(TableStyle([
                    ('BACKGROUND', (0, 0), (-1, 0), colors.HexColor('#6366f1')),
                    ('TEXTCOLOR', (0, 0), (-1, 0), colors.whitesmoke),
                    ('ALIGN', (0, 0), (-1, -1), 'CENTER'),
                    ('FONTSIZE', (0, 0), (-1, -1), 9),
                    ('GRID', (0, 0), (-1, -1), 1, colors.black)
                ]))
                story.append(dm_table)
            else:
                story.append(Paragraph("ไม่มีข้อมูล", normal_style))
            story.append(Spacer(1, 0.15*inch))
            
            # Section 4: สถานะคลังอาหาร
            if self.feed_stock:
                story.append(Paragraph("📦 4. สถานะคลังอาหาร", heading_style))
                feed_stock_data = [["ประเภทอาหาร", "ปริมาณ (กท.)", "ระดับขั้นต่ำ"]]
                for feed_type, data in self.feed_stock.items():
                    feed_stock_data.append([
                        feed_type,
                        str(data.get('quantity', '0')),
                        str(data.get('min_level', '0'))
                    ])
                
                feed_table = Table(feed_stock_data, colWidths=[2*inch, 1.5*inch, 1.5*inch])
                feed_table.setStyle(TableStyle([
                    ('BACKGROUND', (0, 0), (-1, 0), colors.HexColor('#059669')),
                    ('TEXTCOLOR', (0, 0), (-1, 0), colors.whitesmoke),
                    ('ALIGN', (0, 0), (-1, -1), 'CENTER'),
                    ('FONTSIZE', (0, 0), (-1, -1), 9),
                    ('BOTTOMPADDING', (0, 0), (-1, 0), 8),
                    ('BACKGROUND', (0, 1), (-1, -1), colors.lightgreen),
                    ('GRID', (0, 0), (-1, -1), 1, colors.black)
                ]))
                story.append(feed_table)
            
            doc.build(story)
            buffer.seek(0)
            return buffer
            
        except Exception as e:
            print(f"Error exporting PDF: {e}")
            return None
    
    def export_to_excel(self):
        """Export full report to Excel with multiple sheets"""
        try:
            wb = Workbook()
            
            # Define styles
            header_fill = PatternFill(start_color="06b6d4", end_color="06b6d4", fill_type="solid")
            header_font = Font(bold=True, color="FFFFFF", size=11)
            title_font = Font(bold=True, size=14, color="06b6d4")
            border = Border(
                left=Side(style='thin'),
                right=Side(style='thin'),
                top=Side(style='thin'),
                bottom=Side(style='thin')
            )
            
            # Remove default sheet and create summarysheet
            ws_summary = wb.active
            ws_summary.title = "สรุปรายงาน"
            
            # Title
            ws_summary['A1'] = "รายงานบริหารจัดการฟาร์มแกะ กองบิน 23"
            ws_summary['A1'].font = title_font
            ws_summary.merge_cells('A1:D1')
            
            ws_summary['A2'] = f"สร้างเมื่อ: {datetime.now().strftime('%d/%m/%Y เวลา %H:%M:%S')} น."
            ws_summary.merge_cells('A2:D2')
            
            report_data = self.generate_report()
            
            # Summary Statistics
            ws_summary['A4'] = "สถิติแกะและแพะ"
            ws_summary['A4'].font = header_font
            ws_summary['A4'].fill = header_fill
            
            summary_data = [
                ["หมวดหมู่", "จำนวน"],
                ["แกะทั้งหมด", report_data['total_sheep']],
                ["ตัวผู้", report_data['male_sheep']],
                ["ตัวเมีย", report_data['female_sheep']],
                ["แพะทั้งหมด", report_data['total_goats']],
                ["แพะตัวผู้", report_data['male_goats']],
                ["แพะตัวเมีย", report_data['female_goats']],
                ["ปกติ", report_data['healthy']],
                ["ป่วย", report_data['sick']],
                ["ตาย", report_data['dead']],
                ["หาย", report_data['missing']],
                ["", ""],
                ["สถิติการผสมพันธุ์", "จำนวน"],
                ["แผนผสมพันธุ์ทั้งหมด", report_data['total_breeding']],
                ["ผสมพันธุ์สำเร็จ", report_data['successful_breeding']],
                ["ลูกแกะทั้งหมด", report_data['total_lambs']],
                ["", ""],
                ["สถิติสุขภาพและอาหาร", "จำนวน"],
                ["ตรวจสุขภาพ", report_data['health_exams']],
                ["ฉีดวัคซีน", report_data['vaccines']],
                ["บันทึกการให้อาหาร", report_data['feeding_records']]
            ]
            
            for row_idx, row_data in enumerate(summary_data, start=5):
                for col_idx, value in enumerate(row_data, start=1):
                    cell = ws_summary.cell(row=row_idx, column=col_idx, value=value)
                    if row_data[0] in ["หมวดหมู่", "สถิติแกะและแพะ", "สถิติการผสมพันธุ์", "สถิติสุขภาพและอาหาร"]:
                        cell.font = Font(bold=True, color="FFFFFF")
                        cell.fill = header_fill
                    cell.border = border
            
            ws_summary.column_dimensions['A'].width = 30
            ws_summary.column_dimensions['B'].width = 15
            
            # Sheet 2: Detailed Sheep and Goat List
            ws_sheep = wb.create_sheet("แกะและแพะ")
            ws_sheep['A1'] = "รายชื่อและข้อมูลแกะและแพะ"
            ws_sheep['A1'].font = title_font
            ws_sheep.merge_cells('A1:I1')
            
            sheep_headers = ["รหัส ID", "ชนิดสัตว์", "สีแท็กติดหูแกะ", "เพศ", "วันเกิด", "สายพันธุ์", "น้ำหนัก (กก.)", "สถานะ", "วันลงทะเบียน"]
            for col_idx, header in enumerate(sheep_headers, start=1):
                cell = ws_sheep.cell(row=3, column=col_idx, value=header)
                cell.font = header_font
                cell.fill = header_fill
                cell.border = border
            
            for row_idx, sheep in enumerate(self.sheep_records.values(), start=4):
                ws_sheep.cell(row=row_idx, column=1, value=sheep.get('id', '-')).border = border
                ws_sheep.cell(row=row_idx, column=2, value=sheep.get('animal_type', 'แกะ')).border = border
                ws_sheep.cell(row=row_idx, column=3, value=sheep.get('tag_color', '-')).border = border
                ws_sheep.cell(row=row_idx, column=4, value=sheep.get('gender', '-')).border = border
                ws_sheep.cell(row=row_idx, column=5, value=sheep.get('dob', '-')).border = border
                ws_sheep.cell(row=row_idx, column=6, value=sheep.get('breed', '-')).border = border
                ws_sheep.cell(row=row_idx, column=7, value=sheep.get('weight', '-')).border = border
                ws_sheep.cell(row=row_idx, column=8, value=sheep.get('status', '-')).border = border
                ws_sheep.cell(row=row_idx, column=9, value=sheep.get('registered_date', '-')).border = border
            
            for col in ['A', 'B', 'C', 'D', 'E', 'F', 'G', 'H', 'I']:
                ws_sheep.column_dimensions[col].width = 15
            
            # Sheet 3: Feeding History
            ws_feeding = wb.create_sheet("ประวัติการให้อาหาร")
            ws_feeding['A1'] = "ประวัติการให้อาหาร"
            ws_feeding['A1'].font = title_font
            ws_feeding.merge_cells('A1:E1')
            
            feeding_headers = ["วันที่", "ประเภทอาหาร", "ปริมาณ (กท.)", "รหัสแกะ", "หมายเหตุ"]
            for col_idx, header in enumerate(feeding_headers, start=1):
                cell = ws_feeding.cell(row=3, column=col_idx, value=header)
                cell.font = header_font
                cell.fill = header_fill
                cell.border = border
            
            for row_idx, (record_id, record) in enumerate(self.feeding_records.items(), start=4):
                ws_feeding.cell(row=row_idx, column=1, value=record.get('date', '-')).border = border
                ws_feeding.cell(row=row_idx, column=2, value=record.get('feed_type', '-')).border = border
                ws_feeding.cell(row=row_idx, column=3, value=record.get('quantity', '-')).border = border
                ws_feeding.cell(row=row_idx, column=4, value=', '.join(record.get('sheep_ids', []))).border = border
                ws_feeding.cell(row=row_idx, column=5, value=record.get('notes', '-')).border = border
            
            for col in ['A', 'B', 'C', 'D', 'E']:
                ws_feeding.column_dimensions[col].width = 18
            
            # Sheet 4: Health History
            ws_health = wb.create_sheet("สุขภาพ")
            ws_health['A1'] = "ประวัติสุขภาพและวัคซีน"
            ws_health['A1'].font = title_font
            ws_health.merge_cells('A1:E1')
            
            health_headers = ["รหัสแกะ", "ประเภท", "วันที่", "รายละเอียด", "โน้ต"]
            for col_idx, header in enumerate(health_headers, start=1):
                cell = ws_health.cell(row=3, column=col_idx, value=header)
                cell.font = header_font
                cell.fill = header_fill
                cell.border = border
            
            row_idx = 4
            for sheep_id, records in self.health_records.items():
                sheep_info = self.sheep_records.get(sheep_id, {})
                if sheep_info.get('status') in ['ตาย', 'หาย']:
                    continue
                    
                for record in records:
                    ws_health.cell(row=row_idx, column=1, value=sheep_id).border = border
                    ws_health.cell(row=row_idx, column=2, value=record.get('type', '-')).border = border
                    ws_health.cell(row=row_idx, column=3, value=record.get('date', '-')).border = border
                    
                    if record.get('type') == 'ตรวจสุขภาพ':
                        detail = f"สถานะ: {record.get('status', '-')}, อุณหภูมิ: {record.get('temperature', '-')}°C"
                    else:
                        detail = f"วัคซีน: {record.get('vaccine_name', '-')}, เข็มที่: {record.get('injection_number', '-')}"
                    
                    ws_health.cell(row=row_idx, column=4, value=detail).border = border
                    ws_health.cell(row=row_idx, column=5, value=record.get('symptoms', '-')).border = border
                    row_idx += 1
            
            for col in ['A', 'B', 'C', 'D', 'E']:
                ws_health.column_dimensions[col].width = 20
            
            # Sheet 5: Breeding History
            ws_breeding = wb.create_sheet("ผสมพันธุ์")
            ws_breeding['A1'] = "ประวัติการผสมพันธุ์"
            ws_breeding['A1'].font = title_font
            ws_breeding.merge_cells('A1:F1')
            
            breeding_headers = ["แม่ (ID)", "พ่อ (ID)", "วันผสม", "วันคลอดคาดหวัง", "สถานะ", "ลูกแกะ"]
            for col_idx, header in enumerate(breeding_headers, start=1):
                cell = ws_breeding.cell(row=3, column=col_idx, value=header)
                cell.font = header_font
                cell.fill = header_fill
                cell.border = border
            
            for row_idx, record in enumerate(self.breeding_records.values(), start=4):
                ws_breeding.cell(row=row_idx, column=1, value=record.get('female_id', '-')).border = border
                ws_breeding.cell(row=row_idx, column=2, value=record.get('male_id', '-')).border = border
                ws_breeding.cell(row=row_idx, column=3, value=record.get('breeding_date', '-')).border = border
                ws_breeding.cell(row=row_idx, column=4, value=record.get('expected_date', '-')).border = border
                ws_breeding.cell(row=row_idx, column=5, value=record.get('status', '-')).border = border
                
                lambs_count = len(record.get('lambs', []))
                ws_breeding.cell(row=row_idx, column=6, value=f"{lambs_count} ตัว").border = border
            
            for col in ['A', 'B', 'C', 'D', 'E', 'F']:
                ws_breeding.column_dimensions[col].width = 16
            
            # Sheet 6: Feed Stock
            ws_feed = wb.create_sheet("คลังอาหาร")
            ws_feed['A1'] = "สถานะคลังอาหาร"
            ws_feed['A1'].font = title_font
            ws_feed.merge_cells('A1:D1')
            
            feed_headers = ["ประเภทอาหาร", "ปริมาณปัจจุบัน (กก.)", "ระดับขั้นต่ำ (กก.)", "วันอัปเดตล่าสุด"]
            for col_idx, header in enumerate(feed_headers, start=1):
                cell = ws_feed.cell(row=3, column=col_idx, value=header)
                cell.font = header_font
                cell.fill = header_fill
                cell.border = border
            
            for row_idx, (feed_type, data) in enumerate(self.feed_stock.items(), start=4):
                ws_feed.cell(row=row_idx, column=1, value=feed_type).border = border
                ws_feed.cell(row=row_idx, column=2, value=data.get('quantity', '-')).border = border
                ws_feed.cell(row=row_idx, column=3, value=data.get('min_level', '-')).border = border
                ws_feed.cell(row=row_idx, column=4, value=data.get('last_updated', '-')).border = border
            
            for col in ['A', 'B', 'C', 'D']:
                ws_feed.column_dimensions[col].width = 20
                
            # Sheet 7: Dead/Missing Sheep
            ws_dead = wb.create_sheet("แกะจำหน่าย")
            ws_dead['A1'] = "รายการแกะจำหน่าย (ตาย/หาย)"
            ws_dead['A1'].font = title_font
            ws_dead.merge_cells('A1:D1')
            
            dead_headers = ["รหัส ID", "เพศ", "สถานะ", "วันที่ตาย/หาย"]
            for col_idx, header in enumerate(dead_headers, start=1):
                cell = ws_dead.cell(row=3, column=col_idx, value=header)
                cell.font = header_font
                cell.fill = header_fill
                cell.border = border
                
            row_idx = 4
            for sheep in self.sheep_records.values():
                if sheep.get('status') in ['ตาย', 'หาย']:
                    ws_dead.cell(row=row_idx, column=1, value=sheep.get('id', '-')).border = border
                    ws_dead.cell(row=row_idx, column=2, value=sheep.get('gender', '-')).border = border
                    ws_dead.cell(row=row_idx, column=3, value=sheep.get('status', '-')).border = border
                    ws_dead.cell(row=row_idx, column=4, value=sheep.get('death_date', '-')).border = border
                    row_idx += 1
            
            for col in ['A', 'B', 'C', 'D']:
                ws_dead.column_dimensions[col].width = 15
            
            buffer = BytesIO()
            wb.save(buffer)
            buffer.seek(0)
            return buffer
            
        except Exception as e:
            print(f"Error exporting Excel: {e}")
            return None

class UserManager:
    def __init__(self):
        DATA_DIR.mkdir(exist_ok=True)
        self.load_users()
    
    def load_users(self):
        if USERS_FILE.exists():
            with open(USERS_FILE, 'r', encoding='utf-8') as f:
                self.users = json.load(f)
            changed = False
            for user in self.users.values():
                if user.get('role') in ('staff', 'viewer'):
                    user['role'] = 'user'
                    changed = True
            if changed:
                self.save_users()
        else:
            # สร้าง admin เริ่มต้น
            self.users = {
                'admin': {
                    'username': 'admin',
                    'password': generate_password_hash('admin123'),
                    'role': 'admin',
                    'display_name': 'ผู้ดูแลระบบ',
                    'created_at': datetime.now().strftime('%Y-%m-%d %H:%M:%S'),
                    'active': True
                }
            }
            self.save_users()
    
    def save_users(self):
        with open(USERS_FILE, 'w', encoding='utf-8') as f:
            json.dump(self.users, f, ensure_ascii=False, indent=2)
    
    def authenticate(self, username, password):
        user = self.users.get(username)
        if not user:
            return None
        if not user.get('active', True):
            return None
        if check_password_hash(user['password'], password):
            return user
        return None
    
    def create_user(self, username, password, role, display_name):
        if username in self.users:
            return {'status': 'error', 'message': 'ชื่อผู้ใช้นี้มีอยู่แล้ว!'}
        if role not in ROLE_PERMISSIONS:
            return {'status': 'error', 'message': 'ระดับสิทธิ์ไม่ถูกต้อง'}
        self.users[username] = {
            'username': username,
            'password': generate_password_hash(password),
            'role': role,
            'display_name': display_name,
            'created_at': datetime.now().strftime('%Y-%m-%d %H:%M:%S'),
            'active': True
        }
        self.save_users()
        return {'status': 'success', 'message': f'สร้างผู้ใช้ {display_name} สำเร็จ!'}
    
    def update_user(self, username, data, current_user):
        if username not in self.users:
            return {'status': 'error', 'message': 'ไม่พบผู้ใช้'}
        if username == 'admin' and current_user != 'admin':
            return {'status': 'error', 'message': 'ไม่สามารถแก้ไข admin ได้'}
        user = self.users[username]
        if 'display_name' in data:
            user['display_name'] = data['display_name']
        if 'role' in data and data['role'] in ROLE_PERMISSIONS:
            user['role'] = data['role']
        if 'password' in data and data['password']:
            user['password'] = generate_password_hash(data['password'])
        if 'active' in data:
            user['active'] = data['active']
        self.save_users()
        return {'status': 'success', 'message': 'อัปเดตข้อมูลสำเร็จ!'}
    
    def delete_user(self, username, current_user):
        if username == 'admin':
            return {'status': 'error', 'message': 'ไม่สามารถลบบัญชี admin ได้'}
        if username == current_user:
            return {'status': 'error', 'message': 'ไม่สามารถลบบัญชีตัวเองได้'}
        if username not in self.users:
            return {'status': 'error', 'message': 'ไม่พบผู้ใช้'}
        del self.users[username]
        self.save_users()
        return {'status': 'success', 'message': 'ลบผู้ใช้สำเร็จ!'}
    
    def get_all_users(self):
        result = []
        for u in self.users.values():
            result.append({
                'username': u['username'],
                'display_name': u['display_name'],
                'role': u['role'],
                'created_at': u['created_at'],
                'active': u.get('active', True)
            })
        return result

manager = SheepFarmManager()
user_manager = UserManager()

@app.route('/login')
def login_page():
    if 'user_id' in session:
        return redirect('/')
    return render_template('login.html')

@app.route('/')
def index():
    if 'user_id' not in session:
        return redirect('/login')
    return render_template('index.html')

# ===== AUTH API =====
@app.route('/api/login', methods=['POST'])
def api_login():
    data = request.json
    username = data.get('username', '').strip()
    password = data.get('password', '')
    user = user_manager.authenticate(username, password)
    if user:
        session['user_id'] = username
        session['role'] = user['role']
        session['display_name'] = user['display_name']
        return jsonify({
            'status': 'success',
            'message': f'ยินดีต้อนรับ, {user["display_name"]}!',
            'user': {'username': username, 'role': user['role'], 'display_name': user['display_name']}
        })
    return jsonify({'status': 'error', 'message': 'ชื่อผู้ใช้หรือรหัสผ่านไม่ถูกต้อง'}), 401

@app.route('/api/auth/change-password', methods=['POST'])
@login_required
def change_password():
    data = request.json
    old_password = data.get('old_password')
    new_password = data.get('new_password')
    username = session['user_id']
    
    user = user_manager.authenticate(username, old_password)
    if not user:
        return jsonify({'status': 'error', 'message': 'รหัสผ่านเดิมไม่ถูกต้อง'}), 401
    
    if len(new_password) < 6:
        return jsonify({'status': 'error', 'message': 'รหัสผ่านใหม่ต้องมีความยาวอย่างน้อย 6 ตัวอักษร'}), 400
        
    user_manager.users[username]['password'] = generate_password_hash(new_password)
    user_manager.save_users()
    return jsonify({'status': 'success', 'message': 'เปลี่ยนรหัสผ่านสำเร็จ!'})

@app.route('/api/logout', methods=['POST'])
def api_logout():
    session.clear()
    return jsonify({'status': 'success', 'message': 'ออกจากระบบเรียบร้อยแล้ว'})

@app.route('/api/me', methods=['GET'])
def api_me():
    if 'user_id' not in session:
        return jsonify({'status': 'error', 'message': 'Not logged in'}), 401
    return jsonify({
        'status': 'success',
        'user': {
            'username': session['user_id'],
            'role': session['role'],
            'display_name': session['display_name']
        }
    })

# ===== USER MANAGEMENT API (Admin only) =====
@app.route('/api/users', methods=['GET'])
@admin_required
def get_users():
    return jsonify(user_manager.get_all_users())

@app.route('/api/users/create', methods=['POST'])
@admin_required
def create_user():
    data = request.json
    result = user_manager.create_user(
        data['username'], data['password'], data['role'], data['display_name']
    )
    return jsonify(result)

@app.route('/api/users/update', methods=['POST'])
@admin_required
def update_user():
    data = request.json
    result = user_manager.update_user(data['username'], data, session['user_id'])
    return jsonify(result)

@app.route('/api/users/delete/<username>', methods=['DELETE'])
@admin_required
def delete_user(username):
    result = user_manager.delete_user(username, session['user_id'])
    return jsonify(result)

@app.route('/api/sheep', methods=['GET'])
@login_required
def get_sheep():
    return jsonify(manager.get_all_sheep())

@app.route('/api/sheep/register', methods=['POST'])
@permission_required('write')
def register_sheep():
    data = request.json
    result = manager.register_sheep(
        data['sheep_id'],
        data['gender'],
        data['dob'],
        data['breed'],
        data['weight'],
        data.get('animal_type', 'แกะ'),
        data.get('tag_color', '-')
    )
    return jsonify(result)

@app.route('/api/sheep/edit', methods=['POST'])
@permission_required('write')
def edit_sheep():
    data = request.json
    result = manager.edit_sheep(
        data['sheep_id'],
        data.get('gender'),
        data.get('dob'),
        data.get('breed'),
        data.get('weight'),
        data.get('status'),
        data.get('death_date'),
        data.get('animal_type'),
        data.get('tag_color')
    )
    return jsonify(result)

@app.route('/api/sheep/delete/<sheep_id>', methods=['DELETE'])
@permission_required('delete')
def delete_sheep(sheep_id):
    result = manager.delete_sheep(sheep_id)
    return jsonify(result)

@app.route('/api/feeding', methods=['GET'])
@login_required
def get_feeding():
    return jsonify(manager.get_feeding_history())

@app.route('/api/feeding/record', methods=['POST'])
@permission_required('write')
def record_feeding():
    data = request.json
    result = manager.record_feeding(
        data['sheep_ids'],
        data['feed_type'],
        data['quantity'],
        data['notes']
    )
    return jsonify(result)

@app.route('/api/feeding/edit', methods=['POST'])
@permission_required('write')
def edit_feeding():
    data = request.json
    result = manager.edit_feeding(
        data['record_id'],
        data.get('feed_type'),
        data.get('quantity'),
        data.get('notes')
    )
    return jsonify(result)

@app.route('/api/feeding/delete/<record_id>', methods=['DELETE'])
@permission_required('delete')
def delete_feeding(record_id):
    result = manager.delete_feeding(record_id)
    return jsonify(result)

@app.route('/api/health/<sheep_id>', methods=['GET'])
@login_required
def get_health(sheep_id):
    return jsonify(manager.get_health_history(sheep_id))

@app.route('/api/health/record', methods=['POST'])
@permission_required('write')
def record_health():
    data = request.json
    result = manager.record_health(
        data['sheep_id'],
        data['health_type'],
        data.get('status'),
        data.get('symptoms'),
        data.get('temperature'),
        data.get('vaccine_name'),
        data.get('injection_number'),
        data.get('vaccine_date'),
        data.get('next_dose'),
        data.get('death_date')
    )
    return jsonify(result)

@app.route('/api/health/edit', methods=['POST'])
@permission_required('write')
def edit_health():
    data = request.json
    result = manager.edit_health(
        data['sheep_id'],
        int(data['record_index']),
        data['health_type'],
        data.get('status'),
        data.get('symptoms'),
        data.get('temperature'),
        data.get('vaccine_name'),
        data.get('injection_number'),
        data.get('vaccine_date'),
        data.get('next_dose'),
        data.get('death_date')
    )
    return jsonify(result)

@app.route('/api/health/delete/<sheep_id>/<int:record_index>', methods=['DELETE'])
@permission_required('delete')
def delete_health(sheep_id, record_index):
    result = manager.delete_health(sheep_id, record_index)
    return jsonify(result)

@app.route('/api/breeding', methods=['GET'])
@login_required
def get_breeding():
    return jsonify(manager.get_pending_breeding())

@app.route('/api/breeding/plan', methods=['POST'])
@permission_required('write')
def plan_breeding():
    data = request.json
    result = manager.plan_breeding(
        data['female_id'],
        data['male_id'],
        data['breeding_date'],
        data['expected_date']
    )
    return jsonify(result)
 
@app.route('/api/breeding/edit', methods=['POST'])
@permission_required('write')
def edit_breeding_plan():
    data = request.json
    result = manager.edit_breeding_plan(
        data['record_id'],
        data.get('female_id'),
        data.get('male_id'),
        data.get('breeding_date'),
        data.get('expected_date')
    )
    return jsonify(result)

@app.route('/api/lambing', methods=['POST'])
@permission_required('write')
def record_lambing():
    data = request.json
    result = manager.record_lambing(
        data['record_id'],
        data['lambing_date'],
        data['lambs']
    )
    return jsonify(result)

@app.route('/api/lambing/edit', methods=['POST'])
@permission_required('write')
def edit_lambing():
    data = request.json
    result = manager.edit_lambing(
        data['record_id'],
        data.get('lambing_date'),
        data.get('lambs')
    )
    return jsonify(result)

@app.route('/api/lambing/delete/<record_id>', methods=['DELETE'])
@permission_required('delete')
def delete_lambing(record_id):
    result = manager.delete_lambing(record_id)
    return jsonify(result)

@app.route('/api/feed-stock', methods=['GET'])
@login_required
def get_feed_stock():
    return jsonify(manager.get_feed_stock())

@app.route('/api/feed-stock/add', methods=['POST'])
@permission_required('write')
def add_feed_stock():
    data = request.json
    result = manager.add_feed_stock(
        data['feed_type'],
        data['quantity'],
        data['min_level']
    )
    return jsonify(result)

@app.route('/api/feed-stock/update', methods=['POST'])
@permission_required('write')
def update_feed_stock():
    data = request.json
    result = manager.update_feed_stock(
        data['feed_type'],
        data['operation'],
        data['quantity']
    )
    return jsonify(result)

@app.route('/api/report', methods=['GET'])
@login_required
def generate_report():
    return jsonify(manager.generate_report())

@app.route('/api/report/export-pdf', methods=['GET'])
@login_required
def export_report_pdf():
    pdf_buffer = manager.export_to_pdf()
    if pdf_buffer:
        return send_file(
            pdf_buffer,
            mimetype='application/pdf',
            as_attachment=True,
            download_name=f'sheep_farm_report_{datetime.now().strftime("%Y%m%d_%H%M%S")}.pdf'
        )
    return jsonify({"error": "ไม่สามารถสร้างไฟล์ PDF ได้"}), 500

@app.route('/api/report/export-excel', methods=['GET'])
@login_required
def export_report_excel():
    excel_buffer = manager.export_to_excel()
    if excel_buffer:
        return send_file(
            excel_buffer,
            mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
            as_attachment=True,
            download_name=f'sheep_farm_report_{datetime.now().strftime("%Y%m%d_%H%M%S")}.xlsx'
        )
    return jsonify({"error": "ไม่สามารถสร้างไฟล์ Excel ได้"}), 500

if __name__ == '__main__':
    app.run(debug=True, host='0.0.0.0', port=5000)
