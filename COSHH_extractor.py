import fitz  # PyMuPDF
import pytesseract
from pytesseract import Output
from PIL import Image
import re
import os
import pandas as pd
import io

def extract_text_with_ocr(pdf_path):
    doc = fitz.open(pdf_path)
    full_text = ""
    for page in doc:
        text = page.get_text()
        if text.strip():
            full_text += text + "\n"
        else:
            pix = page.get_pixmap(dpi=300)
            img = Image.open(io.BytesIO(pix.tobytes("png")))
            ocr_text = pytesseract.image_to_string(img, config='--psm 6', output_type=Output.STRING)
            full_text += ocr_text + "\n"
    return full_text

def extract_coshh_data(data):
    entry = {}
    entry['Chemical Name'] = re.search(r'Product name\s+(.*?)\n', data, re.IGNORECASE)
    entry['Chemical Name'] = entry['Chemical Name'].group(1).strip() if entry['Chemical Name'] else "N/A"
    entry['Hazard Classification'] = "; ".join(set(re.findall(r'(H\d{3}.*?)\.', data)))
    
    exposure_terms = []
    for term in ["inhalation", "skin", "eye"]:
        if term in data.lower():
            exposure_terms.append(term.capitalize() + " contact" if term != "inhalation" else "Inhalation")
    entry['Exposure Risks'] = ", ".join(exposure_terms)

    controls = []
    if "ventilation" in data.lower():
        controls.append("Local exhaust ventilation")
    if "goggles" in data.lower() or "face shield" in data.lower():
        controls.append("Eye/face protection")
    if "gloves" in data.lower():
        controls.append("Protective gloves (EN 374)")
    if "respiratory" in data.lower():
        controls.append("Respiratory protection")
    entry['Control Measures'] = ", ".join(controls)

    emergency = []
    if "fresh air" in data.lower():
        emergency.append("Move to fresh air")
    if "rinse" in data.lower():
        emergency.append("Rinse eyes/skin")
    if "vomiting" in data.lower():
        emergency.append("Do not induce vomiting")
    if "fire" in data.lower():
        emergency.append("Use CO₂, foam, dry powder")
    entry['Emergency Procedures'] = ", ".join(emergency)

    entry['Additional Recommendations'] = "Store tightly sealed; Avoid environmental release; Train staff"
    return entry

def batch_process():
    folder_path = os.path.dirname(os.path.abspath(__file__))
    records = []
    for filename in os.listdir(folder_path):
        if filename.lower().endswith(".pdf"):
            path = os.path.join(folder_path, filename)
            try:
                raw_text = extract_text_with_ocr(path)
                record = extract_coshh_data(raw_text)
                record['Source File'] = filename
                records.append(record)
            except Exception as e:
                print(f"❌ Failed on {filename}: {e}")

    df = pd.DataFrame(records)
    output_path = os.path.join(folder_path, "COSHH_output.xlsx")
    df.to_excel(output_path, index=False)
    print(f"✅ Extraction complete. Output saved to {output_path}")

# 🔧 Run the processor
if __name__ == "__main__":
    batch_process()
