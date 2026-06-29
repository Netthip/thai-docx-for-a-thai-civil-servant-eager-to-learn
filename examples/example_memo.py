#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""ตัวอย่าง: สร้าง 'บันทึกข้อความ' ด้วย thai_official.create_memo()

รัน:  python examples/example_memo.py
ได้ไฟล์ examples/example_memo.docx
"""
import os
import sys
import datetime

sys.path.insert(0, os.path.join(os.path.dirname(__file__), "..", "scripts"))
from thai_official import create_memo, thai_date  # noqa: E402

OUT = os.path.join(os.path.dirname(__file__), "example_memo.docx")

create_memo(
    OUT,
    department="กองคลัง สำนักงานเลขานุการกรม",
    doc_number="ศธ ๐๒๐๑/๔๕๖",
    date=thai_date(datetime.date(2026, 6, 29)),   # ระบุวันที่ตายตัวให้ผลลัพธ์คงที่
    subject="ขออนุมัติเบิกค่าใช้จ่ายในการเดินทางไปราชการ",
    to="ผู้อำนวยการกองคลัง",
    body=[
        "ด้วยข้าพเจ้ามีความจำเป็นต้องเดินทางไปราชการเพื่อเข้าร่วมประชุมเชิงปฏิบัติการ "
        "เรื่องการพัฒนาระบบงานสารบรรณอิเล็กทรอนิกส์ ในระหว่างวันที่ ๑ ถึงวันที่ ๓ "
        "กรกฎาคม ๒๕๖๙ ณ จังหวัดเชียงใหม่ โดยมีค่าใช้จ่ายเป็นค่าเบี้ยเลี้ยง ค่าที่พัก "
        "และค่าพาหนะ รวมเป็นเงินทั้งสิ้น ๘,๕๐๐ บาท (แปดพันห้าร้อยบาทถ้วน)",
        "จึงเรียนมาเพื่อโปรดพิจารณาอนุมัติ จะเป็นพระคุณยิ่ง",
    ],
    signer_name="นางสาวเนตรทิพย์ ใจดี",
    signer_position="นักวิชาการเงินและบัญชีชำนาญการ",
)
print("created:", OUT)
