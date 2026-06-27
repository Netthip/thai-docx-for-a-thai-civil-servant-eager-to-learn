---
name: thai-docx
description: |
  สร้างไฟล์ Word (.docx) ที่ภาษาไทยเขียนเต็มบรรทัดไม่ตัดบรรทัดก่อนเวลา ใช้ skill นี้ทุกครั้งที่ต้องสร้างเอกสาร Word ที่มีเนื้อหาภาษาไทย ไม่ว่าจะเป็นรายงาน บทความ เอกสารวิชาการ บันทึกข้อความ หรือเอกสารทั่วไป ใช้ได้กับข้อความภาษาไทยล้วนและไทย-อังกฤษผสม
  Trigger เมื่อ: ผู้ใช้ต้องการสร้าง Word ที่มีภาษาไทย, "ทำ docx ภาษาไทย", "สร้างเอกสาร Word", "ข้อความไทยไม่เต็มบรรทัด", "ข้อความตัดบรรทัดผิด", "Thai text line break", "สร้างรายงาน", "เขียนเอกสาร"
  ใช้ skill นี้แทน docx skill ปกติเมื่อเอกสารมีภาษาไทย เพราะ docx-js (JavaScript) ไม่จัดการ Thai line breaking ได้ดี และ python-docx ปกติก็ตั้งฟอนต์ฝั่ง Complex Script ของไทยไม่ครบ — สคริปต์ใน skill นี้แก้ทั้งสองเรื่องให้แล้ว
---

# Thai DOCX — เอกสาร Word ภาษาไทยที่เต็มบรรทัด ฟอนต์ถูก (ฉบับเบสิก)

> ฉบับเต็มสำหรับข้าราชการ (ฟอนต์แห่งชาติ + เอกสารราชการ/สารบรรณ + สารบัญ/ตาราง/เลขหน้า) อยู่ที่ branch `main`

## ปัญหา 2 ข้อที่ skill นี้แก้

1. **ไทยตัดบรรทัดก่อนเวลา** — ภาษาไทยไม่มีช่องว่างระหว่างคำ Word จึงไม่รู้จะตัดบรรทัดตรงไหน เลยตัดเร็วเกินไป เกิดบรรทัดสั้น ๆ เสียพื้นที่ขอบขวา
2. **ฟอนต์ไทยไม่เปลี่ยนตามที่สั่ง** — python-docx ตั้งฟอนต์ผ่าน `run.font.name` ให้แค่ฝั่ง `w:ascii`/`w:hAnsi` **ไม่ตั้งฝั่ง `w:cs` (Complex Script) ที่อักษรไทยใช้จริง**

## วิธีแก้

1. แทรก **Zero-Width Space (U+200B)** ระหว่างคำไทย (ตัดคำด้วย pythainlp) — มองไม่เห็น แต่บอก Word ว่าตัดบรรทัดตรงไหนได้
2. ตั้งฟอนต์ผ่าน XML ให้ครบทั้ง `w:ascii`/`w:hAnsi`/**`w:cs`**/`w:eastAsia` รวมขนาด (`w:szCs`) และตัวหนา/เอียง (`w:bCs`/`w:iCs`)

## ติดตั้ง

```bash
pip install pythainlp python-docx
```

## ใช้งานเร็ว (CLI)

```bash
python scripts/thai_docx.py input.md -o output.docx --preset default --title "รายงาน"
```

- input รองรับ markdown ง่าย ๆ: `#`/`##`/`###` = หัวข้อ; `-`/`*` = bullet; `1.` = เลขข้อ; `**ตัวหนา**`
- บรรทัดว่างคั่นย่อหน้า

## ใช้งานผ่านโค้ด

```python
from thai_docx import create_docx

paragraphs = [
    {"text": "รายงานประจำปี", "type": "title"},
    {"text": "บทที่ ๑ บทนำ", "type": "heading1"},
    {"text": "เนื้อหาย่อหน้าภาษาไทย รองรับ **ตัวหนา** และ *ตัวเอียง*", "type": "body"},
    {"text": "ข้อแรก", "type": "bullet"},
    {"type": "table", "header": True, "rows": [["รายการ", "จำนวน"], ["ก", "100"]]},
]

create_docx(
    paragraphs, "output.docx",
    preset="default",       # "default" | "book"
    font_size=16,
    line_spacing=1.5,
)
```

ชนิด paragraph: `title`, `subtitle`, `heading1`–`heading3`, `body`, `bullet`, `number`, `quote`, `caption`, `table`, `image`, `pagebreak`

## กฎสำคัญ

- **ใช้ python-docx (Python) เท่านั้น ห้ามใช้ docx-js (JavaScript)** สำหรับเอกสารไทย
- **ใช้การจัดชิดซ้าย (LEFT)** — JUSTIFIED จะยืดช่องว่างน่าเกลียด
- `insert_zwsp()` ถูกเรียกอัตโนมัติใน `create_docx()` แล้ว ไม่ต้องเรียกเอง

## ฟังก์ชันที่ใช้ได้ (scripts/thai_docx.py)

- `insert_zwsp(text, engine="newmm")` — แทรก ZWS ระหว่างคำไทย
- `set_run_font(run, name, size, bold, italic, color)` — ตั้งฟอนต์ครบฝั่ง cs
- `create_docx(paragraphs, output_path, ...)` — สร้างไฟล์
- `parse_markdown(raw)` — แปลง markdown/ข้อความ → รายการ paragraph
