import streamlit as st
import os
import requests
import json
from docx import Document
from PIL import Image
import pytesseract
from pdf2image import convert_from_path

# ==========================================
# CONFIG & INITIALIZATION
# ==========================================
TEMPLATE_DIR = "hr_templates"
OUTPUT_DIR = "generated_docs"

for folder in [TEMPLATE_DIR, OUTPUT_DIR]:
    if not os.path.exists(folder):
        os.makedirs(folder)

st.set_page_config(page_title="HR AI Document Automation", layout="wide")
st.title("📄 ระบบจัดการและเขียนเอกสาร HR อัตโนมัติด้วย AI")
st.subheader("ลดงานคีย์ข้อมูล 100% ด้วย Open-Source LLM (ใช้งานฟรี ไม่มีลิมิต)")

# ==========================================
# FUNCTIONS
# ==========================================

# 1. ฟังก์ชันดึงข้อความจากไฟล์ (OCR)
def extract_text_from_file(file_path):
    ext = os.path.splitext(file_path)[1].lower()
    text = ""
    if ext in ['.jpg', '.jpeg', '.png']:
        text = pytesseract.image_to_string(Image.open(file_path), lang='tha+eng')
    elif ext == '.pdf':
        pages = convert_from_path(file_path)
        for page in pages:
            text += pytesseract.image_to_string(page, lang='tha+eng') + "\n"
    return text

# 2. ฟังก์ชันเรียก Open-Source AI (Ollama) เพื่อดึงข้อมูลเป็น JSON
def extract_data_with_ai(raw_text, keys_needed):
    url = "http://localhost:11434/api/generate"
    
    # สร้าง Prompt บังคับให้ AI ตอบกลับมาเป็น JSON ตาม Key ที่ระบุเท่านั้น
    prompt = f"""
    คุณคือผู้ช่วย HR อัจฉริยะ หน้าที่ของคุณคืออ่านข้อความที่ได้จากการสแกนเอกสารดังต่อไปนี้ 
    แล้วดึงข้อมูลสำคัญออกมาตามหัวข้อที่กำหนดให้ในรูปแบบ JSON object เท่านั้น ห้ามเขียนคำอธิบายอื่นใดเพิ่มเติม

    หัวข้อที่ต้องดึง (Keys): {keys_needed}

    ข้อความจากเอกสาร:
    \"\"\"
    {raw_text}
    \"\"\"

    ตอบกลับเป็นโครงสร้าง JSON เท่านั้น:
    """
    
    data = {
        "model": "typhoon-m1", # หรือ llama3.1
        "prompt": prompt,
        "stream": False,
        "format": "json"
    }
    
    try:
        response = requests.post(url, json=data)
        result = response.json()
        return json.loads(result['response'])
    except Exception as e:
        st.error(f"เกิดข้อผิดพลาดในการเชื่อมต่อ AI: {e}")
        return None

# 3. ฟังก์ชันแทนที่คำใน Word Document (.docx)
def fill_template(template_path, data, output_path):
    doc = Document(template_path)
    
    # วนลูปค้นหาและแทนที่ข้อความในย่อหน้า (Paragraphs)
    for p in doc.paragraphs:
        for key, value in data.items():
            placeholder = f"{{{{{key}}}}}" # หาตัวแปรในรูปแบบ {{key}}
            if placeholder in p.text:
                p.text = p.text.replace(placeholder, str(value))
                
    # วนลูปค้นหาและแทนที่ข้อความในตาราง (Tables)
    for table in doc.tables:
        for row in table.rows:
            for cell in row.cells:
                for p in cell.paragraphs:
                    for key, value in data.items():
                        placeholder = f"{{{{{key}}}}}"
                        if placeholder in p.text:
                            p.text = p.text.replace(placeholder, str(value))
                            
    doc.save(output_path)

# ==========================================
# UI SIDEBAR: TEMPLATE MANAGEMENT (ข้อ 2 & ข้อ 5)
# ==========================================
with st.sidebar:
    st.header("⚙️ จัดการเทมเพลตบริษัท")
    
    # อัปโหลดเทมเพลตใหม่
    uploaded_tpl = st.file_uploader("อัปโหลดเทมเพลตใหม่ (.docx)", type=["docx"])
    if uploaded_tpl is not None:
        tpl_path = os.path.join(TEMPLATE_DIR, uploaded_tpl.name)
        with open(tpl_path, "wb") as f:
            f.write(uploaded_tpl.getbuffer())
        st.success(f"บันทึกเทมเพลต {uploaded_tpl.name} แล้ว!")
        st.rerun()

    st.divider()
    
    # ลบเทมเพลต
    st.subheader("🗑️ ลบเทมเพลตที่ไม่ใช้")
    all_tpls = os.listdir(TEMPLATE_DIR)
    if all_tpls:
        tpl_to_delete = st.selectbox("เลือกไฟล์ที่จะลบ:", all_tpls, key="delete_box")
        if st.button("🔴 ยืนยันการลบไฟล์"):
            os.remove(os.path.join(TEMPLATE_DIR, tpl_to_delete))
            st.success(f"ลบ {tpl_to_delete} สำเร็จ")
            st.rerun()
    else:
        st.caption("ยังไม่มีเทมเพลตในระบบ")

# ==========================================
# MAIN UI: RUNNING AUTOMATION (ข้อ 1, 3, 4)
# ==========================================
available_templates = os.listdir(TEMPLATE_DIR)

if not available_templates:
    st.info("💡 เริ่มต้นใช้งานโดยการอัปโหลดไฟล์เทมเพลตเอกสารที่เมนูด้านซ้ายก่อนครับ")
else:
    # 1. ขั้นตอนเลือกเทมเพลตที่จะใช้งาน (ข้อ 3)
    st.header("Step 1: เลือกเทมเพลตและระบุฟิลด์ข้อมูล")
    selected_template = st.selectbox("เลือกเทมเพลตเอกสารที่จะใช้:", available_templates, key="main_select")
    
    # ให้ HR กำหนดหัวข้อที่อยู่ในเทมเพลตนั้นๆ (เพื่อบอก AI ให้ดึงได้ตรงจุด)
    # เช่น ในไฟล์ Word มีพิมพ์ว่า {{Name}}, {{ID}}, {{Salary}}
    st.markdown("*ระบุชื่อตัวแปรที่อยู่ในไฟล์เทมเพลต (คั่นด้วยเครื่องหมายจุลภาค `,`)*")
    fields_input = st.text_input("ตัวแปรที่ต้องการให้ AI เติมลงในเอกสาร:", "Name, ID_Number, Address, Position, Salary")
    fields_list = [f.strip() for f in fields_input.split(",")]

    st.divider()

    # 2. ขั้นตอนอัปโหลดเอกสารดิบนำเข้า (ข้อ 4)
    st.header("Step 2: อัปโหลดเอกสาร Input ต้นทาง")
    input_file = st.file_uploader("อัปโหลดเอกสารต้นทาง (เช่น รูปถ่ายใบสมัครงาน, PDF เรซูเม่, บัตรประชาชน)", type=["png", "jpg", "jpeg", "pdf"])

    if input_file is not None:
        # เซฟไฟล์ชั่วคราวลงเครื่อง
        temp_input_path = os.path.join(OUTPUT_DIR, input_file.name)
        with open(temp_input_path, "wb") as f:
            f.write(input_file.getbuffer())
            
        if st.button("🚀 เริ่มต้นดึงข้อมูลและสร้างเอกสารอัตโนมัติ"):
            with st.status("🤖 กำลังประมวลผลระบบอัตโนมัติ...", expanded=True) as status:
                
                # ข้อ 1 & 4: ระบบรันอัตโนมัติแทนคน
                status.write("⏳ 1. กำลังอ่านข้อความจากไฟล์เอกสาร (OCR)...")
                raw_text = extract_text_from_file(temp_input_path)
                
                status.write("🧠 2. กำลังส่งให้ Open-Source AI วิเคราะห์ดึงข้อมูลสำคัญ...")
                extracted_json = extract_data_with_ai(raw_text, fields_list)
                
                if extracted_json:
                    status.write("📝 3. ข้อมูลที่ AI ดึงมาได้ (กำลังกรอกลงเทมเพลต):")
                    status.json(extracted_json)
                    
                    # กำหนดชื่อไฟล์ผลลัพธ์
                    output_filename = f"Filled_{selected_template}"
                    output_path = os.path.join(OUTPUT_DIR, output_filename)
                    template_path = os.path.join(TEMPLATE_DIR, selected_template)
                    
                    status.write("💾 4. กำลังสร้างไฟล์เอกสารฉบับใหม่...")
                    fill_template(template_path, extracted_json, output_path)
                    
                    status.update(label="🎉 ดำเนินการสำเร็จเรียบร้อย!", state="complete", expanded=True)
                    
                    # แสดงปุ่มให้ดาวน์โหลดผลลัพธ์
                    st.success(f"สร้างเอกสารสำเร็จตามเทมเพลตที่ตั้งไว้!")
                    with open(output_path, "rb") as file:
                        st.download_button(
                            label="📥 ดาวน์โหลดเอกสารสำเร็จรูป (.docx)",
                            data=file,
                            file_name=output_filename,
                            mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
                        )
                else:
                    status.update(label="❌ การดึงข้อมูลผิดพลาด", state="error")
                    st.error("AI ไม่สามารถอ่านข้อมูลหรือแปลงเป็น JSON ได้ กรุณาลองใหม่อีกครั้ง")
