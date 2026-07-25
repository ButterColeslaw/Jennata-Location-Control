from fastapi import FastAPI, UploadFile, File, Form
from fastapi.responses import HTMLResponse, FileResponse
from datetime import datetime
import pandas as pd
import fitz  # PyMuPDF
import tempfile
import io
import re
import os

app = FastAPI()

TODAY_STR = datetime.now().strftime("%Y%m%d")
NEW_FILENAME = f"LOC-{TODAY_STR}.pdf"

HTML_FORM = f"""
<!DOCTYPE html>
<html lang="th">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>Auto Loc. Control</title>
    <style>
        body {{ font-family: 'Segoe UI', Tahoma, Geneva, Verdana, sans-serif; background-color: #f4f6f9; margin: 0; padding: 20px; }}
        .container {{ max-width: 520px; margin: 40px auto; background: #fff; padding: 30px; border-radius: 8px; box-shadow: 0 4px 6px rgba(0,0,0,0.1); }}
        h2 {{ text-align: center; color: #333; margin-bottom: 20px; }}
        .form-group {{ margin-bottom: 15px; }}
        label {{ display: block; margin-bottom: 5px; font-weight: bold; color: #555; }}
        input[type="text"], input[type="number"], input[type="file"] {{ width: 100%; padding: 10px; border: 1px solid #ccc; border-radius: 4px; box-sizing: border-box; }}
        .help-text {{ font-size: 12px; color: #777; margin-top: 3px; }}
        
        .btn-group {{ display: flex; gap: 10px; margin-top: 15px; }}
        button {{ flex: 1; padding: 12px; border: none; border-radius: 4px; font-size: 15px; font-weight: bold; cursor: pointer; color: white; }}
        
        .btn-preview {{ background-color: #6c757d; }}
        .btn-preview:hover {{ background-color: #5a6268; }}
        
        .btn-download {{ background-color: #0d6efd; }}
        .btn-download:hover {{ background-color: #0b5ed7; }}
        
        .modal {{ display: none; position: fixed; z-index: 1000; left: 0; top: 0; width: 100%; height: 100%; background-color: rgba(0,0,0,0.6); }}
        .modal-content {{ background-color: #fff; margin: 2% auto; padding: 15px; width: 85%; height: 90%; border-radius: 8px; position: relative; display: flex; flex-direction: column; }}
        .close-btn {{ position: absolute; top: 10px; right: 20px; font-size: 28px; font-weight: bold; color: #aaa; cursor: pointer; }}
        .close-btn:hover {{ color: #000; }}
        iframe {{ width: 100%; height: 100%; border: none; margin-top: 10px; }}
    </style>
</head>
<body>
    <div class="container">
        <h2>ระบบประมวลผลใบคุมโลเคชัน</h2>
        <form id="processForm" enctype="multipart/form-data">
            <div class="form-group">
                <label>ไฟล์ Excel นำออก by LOC. (.xls / .xlsx)</label>
                <input type="file" name="excel_file" accept=".xls,.xlsx" required>
            </div>
            <div class="form-group">
                <label>ไฟล์ PDF ใบคุมโลเคชันต้นฉบับ</label>
                <input type="file" name="pdf_file" accept=".pdf" required>
            </div>
            <div class="form-group">
                <label>โลเคชันเริ่มต้นเต็มรูปแบบ:</label>
                <input type="text" name="example_loc" placeholder="26-0xx-xxxxx" required>
                <div class="help-text">ตัวอย่าง 26-099-00099</div>
            </div>
            <div class="form-group">
                <label>ยอด Adjust (ถ้าไม่มีให้ใส่ 0)</label>
                <input type="number" name="adjust_qty" value="0" required>
            </div>
            
            <div class="btn-group">
                <button type="button" class="btn-preview" onclick="submitForm('preview')">Preview เอกสาร</button>
                <button type="button" class="btn-download" onclick="submitForm('download')">ประมวลผลและดาวน์โหลด</button>
            </div>
        </form>
    </div>

    <div id="previewModal" class="modal">
        <div class="modal-content">
            <span class="close-btn" onclick="closeModal()">&times;</span>
            <h3 style="margin: 0 0 10px 0;">ตัวอย่างเอกสาร ({NEW_FILENAME})</h3>
            <iframe id="pdfViewer"></iframe>
        </div>
    </div>

    <script>
        async function submitForm(mode) {{
            const form = document.getElementById('processForm');
            if (!form.checkValidity()) {{
                form.reportValidity();
                return;
            }}

            const formData = new FormData(form);
            
            try {{
                const response = await fetch('/process', {{
                    method: 'POST',
                    body: formData
                }});

                if (!response.ok) {{
                    const errText = await response.text();
                    alert("เกิดข้อผิดพลาด: " + errText);
                    return;
                }}

                const blob = await response.blob();
                const blobUrl = URL.createObjectURL(blob);

                if (mode === 'preview') {{
                    document.getElementById('pdfViewer').src = blobUrl;
                    document.getElementById('previewModal').style.display = 'block';
                }} else if (mode === 'download') {{
                    const a = document.createElement('a');
                    a.href = blobUrl;
                    a.download = "{NEW_FILENAME}";
                    document.body.appendChild(a);
                    a.click();
                    a.remove();
                }}
            }} catch (err) {{
                alert("เกิดข้อผิดพลาดในการเชื่อมต่อเซิร์ฟเวอร์");
            }}
        }}

        function closeModal() {{
            document.getElementById('previewModal').style.display = 'none';
            document.getElementById('pdfViewer').src = '';
        }}
    </script>
</body>
</html>
"""

def parse_prefix_and_padding(example_loc_str):
    cleaned = example_loc_str.strip()
    match = re.match(r'^(.*?)(\d+)$', cleaned)
    if match:
        prefix_base = match.group(1)
        digits_sample = match.group(2)
        return prefix_base, len(digits_sample)
    return cleaned, 0

def read_excel_flexibly(file_bytes, filename):
    try:
        tables = pd.read_html(io.BytesIO(file_bytes))
        if tables:
            return tables[0]
    except Exception:
        pass

    try:
        return pd.read_excel(io.BytesIO(file_bytes), engine='openpyxl')
    except Exception:
        pass

    try:
        return pd.read_excel(io.BytesIO(file_bytes), engine='xlrd')
    except Exception:
        pass

    return pd.read_excel(io.BytesIO(file_bytes))

def format_qty_str(val: int) -> str:
    return "-" if val == 0 else str(val)

@app.get("/", response_class=HTMLResponse)
async def read_root():
    return HTMLResponse(content=HTML_FORM)

@app.post("/process")
async def process_files(
    excel_file: UploadFile = File(...),
    pdf_file: UploadFile = File(...),
    example_loc: str = Form(...),
    adjust_qty: int = Form(0)
):
    temp_dir = tempfile.mkdtemp()
    pdf_path = os.path.join(temp_dir, pdf_file.filename)
    output_pdf_path = os.path.join(temp_dir, NEW_FILENAME)

    try:
        prefix_base, digit_len = parse_prefix_and_padding(example_loc)

        # 1. อ่านไฟล์ Excel
        excel_bytes = await excel_file.read()
        df_raw = read_excel_flexibly(excel_bytes, excel_file.filename)

        with open(pdf_path, "wb") as f:
            f.write(await pdf_file.read())

        header_idx = None
        for idx, row in df_raw.iterrows():
            row_str = [str(s) for s in row.values]
            if any("Location" in s for s in row_str):
                header_idx = idx
                break

        if header_idx is not None:
            df_data = df_raw.iloc[header_idx + 1:].copy()
            df_data.columns = [str(c) for c in df_raw.iloc[header_idx].values]
        else:
            df_data = df_raw.copy()

        loc_col = [c for c in df_data.columns if "Location" in str(c)]
        count_col = [c for c in df_data.columns if "Count" in str(c)]

        if not loc_col or not count_col:
            raise ValueError("ไม่พบคอลัมน์ 'Location' หรือ 'Count' ในไฟล์ Excel")

        df_data = df_data.dropna(subset=[loc_col[0], count_col[0]])
        df_data['Count_Clean'] = pd.to_numeric(df_data[count_col[0]], errors='coerce').fillna(0)
        df_data['Location_Clean'] = df_data[loc_col[0]].astype(str).str.strip()

        location_summary = df_data.groupby('Location_Clean')['Count_Clean'].sum().to_dict()

        qty_1401 = int(location_summary.get('1401', 0))
        qty_1402 = int(location_summary.get('1402', 0))
        total_sale_counting = qty_1401 + qty_1402

        # 2. จัดการไฟล์ PDF
        doc = fitz.open(pdf_path)
        front_locations_total = 0
        back_locations_total = 0
        onhand_amount = 0

        FONT_NAME = "hebo"
        FONT_SIZE = 11.0
        NAVY_BLUE = (0.0, 0.15, 0.55)
        OBLIQUE_MATRIX = fitz.Matrix(1, 0, 0.2, 1, 0, 0)

        # ดึงยอด Onhand
        for page in doc:
            text = page.get_text()
            match = re.search(r'Onhand\s*:\s*(\d+)', text)
            if match:
                onhand_amount = int(match.group(1))
                break

        found_front_page = False

        # 3. วนลูปประมวลผล PDF ทีละหน้า
        for page_idx in range(len(doc) - 1):
            page = doc[page_idx]
            words = page.get_text("words")
            page_qty_sum = 0
            
            qty_words = [w for w in words if "Qty" in w[4]]

            for q_w in qty_words:
                q_x0, q_y0, q_x1, q_y1 = q_w[0], q_w[1], q_w[2], q_w[3]
                
                short_loc = None
                candidates = []
                for w in words:
                    w_x0, w_y0, w_x1, w_text = w[0], w[1], w[2], w[4].strip()
                    if abs(w_y0 - q_y0) < 5 and w_x1 < q_x0:
                        if w_text.isdigit() and (q_x0 - w_x1) < 120:
                            if w_text not in ['1401', '1402']:
                                candidates.append((w_x1, w_text))
                
                if candidates:
                    candidates.sort(key=lambda x: x[0], reverse=True)
                    short_loc = candidates[0][1]

                if short_loc:
                    if digit_len > 0 and len(short_loc) < digit_len:
                        formatted_num = short_loc.zfill(digit_len)
                    else:
                        formatted_num = short_loc

                    full_loc = f"{prefix_base}{formatted_num}"
                    count_qty = int(location_summary.get(full_loc, 0))
                    page_qty_sum += count_qty

                    display_text = format_qty_str(count_qty)

                    page.insert_text(
                        (q_x1 + 12, q_y1 - 3.5), 
                        display_text, 
                        fontname=FONT_NAME, 
                        fontsize=FONT_SIZE, 
                        color=NAVY_BLUE,
                        morph=(fitz.Point(q_x1 + 12, q_y1 - 3.5), OBLIQUE_MATRIX)
                    )

            if not found_front_page:
                front_locations_total += page_qty_sum
            else:
                back_locations_total += page_qty_sum

            # ค้นหาข้อความคำว่า Total (Front) โดยไม่สนการเว้นวรรค หรือเครื่องหมาย Colon
            matches_front = page.search_for("Total (Front)")
            if matches_front:
                m_rect = matches_front[0]
                # วางตัวเลขเยื้องจากคำว่า Total (Front) ไปทางขวา 110pt
                page.insert_text(
                    (m_rect.x0 + 110, m_rect.y1 - 3.5), 
                    format_qty_str(front_locations_total), 
                    fontname=FONT_NAME, 
                    fontsize=FONT_SIZE, 
                    color=NAVY_BLUE,
                    morph=(fitz.Point(m_rect.x0 + 110, m_rect.y1 - 3.5), OBLIQUE_MATRIX)
                )
                found_front_page = True

            # ค้นหาข้อความคำว่า Total (Back) โดยไม่สนการเว้นวรรค หรือเครื่องหมาย Colon
            matches_back = page.search_for("Total (Back)")
            if matches_back:
                m_rect = matches_back[0]
                # วางตัวเลขเยื้องจากคำว่า Total (Back) ไปทางขวา 110pt
                page.insert_text(
                    (m_rect.x0 + 110, m_rect.y1 - 3.5), 
                    format_qty_str(back_locations_total), 
                    fontname=FONT_NAME, 
                    fontsize=FONT_SIZE, 
                    color=NAVY_BLUE,
                    morph=(fitz.Point(m_rect.x0 + 110, m_rect.y1 - 3.5), OBLIQUE_MATRIX)
                )

        # คำนวณยอดสรุปรวม
        normal_locations_total = front_locations_total + back_locations_total
        grand_total = normal_locations_total + total_sale_counting + adjust_qty
        count_val = grand_total
        diff_val = count_val - onhand_amount

        # 4. เติมข้อมูลตารางสรุปในหน้าสุดท้าย
        last_page = doc[-1]

        def write_header_item(keyword, value_str):
            matches = last_page.search_for(keyword)
            if matches:
                rect = matches[0]
                val_len = len(str(value_str))
                if val_len <= 2:
                    base_offset = 8
                elif val_len == 3:
                    base_offset = 5
                else:
                    base_offset = 3

                last_page.insert_text(
                    (rect.x1 + base_offset, rect.y1 - 3.5), 
                    str(value_str), 
                    fontname=FONT_NAME, 
                    fontsize=FONT_SIZE, 
                    color=NAVY_BLUE,
                    morph=(fitz.Point(rect.x1 + base_offset, rect.y1 - 3.5), OBLIQUE_MATRIX)
                )

        def write_table_cell(keyword, value_str, align_x=238, y_drop=1.0):
            matches = last_page.search_for(keyword)
            if matches:
                rect = matches[0]
                last_page.insert_text(
                    (align_x, rect.y1 - y_drop), 
                    str(value_str), 
                    fontname=FONT_NAME, 
                    fontsize=FONT_SIZE, 
                    color=NAVY_BLUE,
                    morph=(fitz.Point(align_x, rect.y1 - y_drop), OBLIQUE_MATRIX)
                )

        sale_pending_str = format_qty_str(qty_1401)
        diff_str = f"-{abs(diff_val)}" if diff_val < 0 else (format_qty_str(diff_val))

        write_header_item("Sale pending :", sale_pending_str)
        write_header_item("Amount :", format_qty_str(onhand_amount))
        write_header_item("Count :", format_qty_str(count_val))
        write_header_item("Diff :", diff_str)

        write_table_cell("1401 # Before Counting", format_qty_str(qty_1401), align_x=238, y_drop=1.0)
        write_table_cell("1402 # Counting", format_qty_str(qty_1402), align_x=238, y_drop=1.0)
        write_table_cell("Total # Sale counting", format_qty_str(total_sale_counting), align_x=238, y_drop=1.0)

        write_table_cell("1701 # Diff 1", format_qty_str(adjust_qty), align_x=238, y_drop=1.0)
        write_table_cell("Total # Adjust", format_qty_str(adjust_qty), align_x=238, y_drop=1.0)

        write_table_cell("Grand Total", format_qty_str(grand_total), align_x=232, y_drop=1.0)

        doc.save(output_pdf_path)
        doc.close()

        return FileResponse(
            path=output_pdf_path,
            filename=NEW_FILENAME,
            media_type="application/pdf"
        )

    except Exception as e:
        return HTMLResponse(content=f"<h3>เกิดข้อผิดพลาดในการประมวลผล: {str(e)}</h3>", status_code=500)

if __name__ == "__main__":
    import uvicorn
    uvicorn.run("main:app", host="127.0.0.1", port=8000, reload=True)