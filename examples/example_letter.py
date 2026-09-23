#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""ตัวอย่าง: สร้าง 'หนังสือภายนอก' ด้วย thai_official.create_letter()

รัน:  python examples/example_letter.py
ได้ไฟล์ examples/example_letter.docx
"""
import os
import sys
import datetime

sys.path.insert(0, os.path.join(os.path.dirname(__file__), "..", "scripts"))
from thai_official import create_letter, thai_date  # noqa: E402

OUT = os.path.join(os.path.dirname(__file__), "example_letter.docx")

create_letter(
    OUT,
    doc_number="ศธ ๐๒๐๑/๑๒๓",
    sender_lines=[
        "สำนักงานเลขานุการกรม",
        "ถนนราชดำเนินนอก เขตดุสิต",
        "กรุงเทพมหานคร ๑๐๓๐๐",
    ],
    date=thai_date(datetime.date(2026, 6, 29)),
    subject="ขอเชิญเข้าร่วมประชุมเชิงปฏิบัติการ",
    to="อธิการบดีมหาวิทยาลัยเชียงใหม่",
    reference="หนังสือกรม ที่ ศธ ๐๒๐๑/ว ๘๙ ลงวันที่ ๑ มิถุนายน ๒๕๖๙",
    enclosure="กำหนดการประชุม จำนวน ๑ ชุด",
    body=[
        "ด้วยกรมมีกำหนดจัดประชุมเชิงปฏิบัติการเรื่องการพัฒนาระบบงานสารบรรณ"
        "อิเล็กทรอนิกส์ ระหว่างวันที่ ๑ ถึงวันที่ ๓ กรกฎาคม ๒๕๖๙ "
        "ณ โรงแรมในจังหวัดเชียงใหม่ เพื่อยกระดับการจัดการเอกสารของหน่วยงานในภูมิภาค",
        "ในการนี้ กรมพิจารณาเห็นว่าบุคลากรในสังกัดของท่านมีความเหมาะสม "
        "จึงขอเรียนเชิญส่งผู้แทนเข้าร่วมการประชุมตามวัน เวลา และสถานที่ดังกล่าว "
        "โดยสามารถแจ้งรายชื่อผู้เข้าร่วมได้ที่กองคลัง ภายในวันที่ ๒๕ มิถุนายน ๒๕๖๙",
        "จึงเรียนมาเพื่อโปรดพิจารณา และขอขอบคุณมา ณ โอกาสนี้",
    ],
    signer_name="นายสมชาย รักงาน",
    signer_position="เลขานุการกรม",
    contact_lines=[
        "กองคลัง",
        "โทร. ๐ ๒๒๐๒ ๑๒๓๔",
        "โทรสาร ๐ ๒๒๐๒ ๑๒๓๕",
    ],
)
print("created:", OUT)
