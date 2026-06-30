import streamlit as st
import google.generativeai as genai
import fitz  # PyMuPDF
import pandas as pd
import openpyxl  # Xử lý định dạng Excel
from openpyxl.styles import Font
import io
import time
import tempfile
import os
import math
import re
import csv
import json
from docx import Document
from docx.shared import Pt
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml.ns import qn

# =====================================================================
# 1. KHỞI TẠO BỘ NHỚ ĐỆM (SESSION STATE) CHO ỨNG DỤNG
# =====================================================================
if 'processed_data' not in st.session_state: st.session_state.processed_data = None
if 'export_mode_used' not in st.session_state: st.session_state.export_mode_used = None
if 'preview_data' not in st.session_state: st.session_state.preview_data = None
if 'trans_processed_data' not in st.session_state: st.session_state.trans_processed_data = None
if 'trans_file_name' not in st.session_state: st.session_state.trans_file_name = None

st.set_page_config(page_title="AI Document Tool", layout="wide")

# =====================================================================
# 2. HÀM TIỆN ÍCH DỊCH THUẬT NÂNG CẤP KÈM TIẾN ĐỘ % THỰC TẾ
# =====================================================================
def translate_texts_batch(texts_list, target_lang, model, skip_english, skip_abbreviations):
    """Gom nhóm văn bản và gửi qua Gemini để dịch kèm các bộ lọc thông minh"""
    if not texts_list:
        return {}
    
    mapping = {}
    chunk_size = 40  # Mẻ dịch tối ưu để bảo vệ giới hạn Quota API
    
    # Thiết lập bộ đếm hiển thị tiến độ trên giao diện
    percent_text = st.empty()
    progress_bar = st.progress(0)
    status = st.empty()
    
    for i in range(0, len(texts_list), chunk_size):
        chunk = texts_list[i:i+chunk_size]
        
        # Cập nhật bộ đếm % hoàn thành thực tế theo thời gian thực
        current_processed = min(i + chunk_size, len(texts_list))
        percent_completed = int((current_processed / len(texts_list)) * 100)
        
        percent_text.markdown(f"### 📊 Tiến độ hoàn thành dịch: `{percent_completed}%`")
        status.text(f"⏳ Đang biên dịch câu số {current_processed}/{len(texts_list)}...")
        progress_bar.progress(current_processed / len(texts_list))
        
     # Xây dựng Prompt dịch thuật tích hợp bộ lọc điều kiện và chống dịch ngược
        prompt = f"""
        Bạn là dịch giả kỹ thuật chuyên nghiệp. Nhiệm vụ của bạn là đồng nhất mảng JSON văn bản sau hoàn toàn sang {target_lang}.
        
        QUY TẮC QUAN TRỌNG NHẤT:
        1. XỬ LÝ ĐA NGÔN NGỮ (CHỐNG DỊCH NGƯỢC): Văn bản đầu vào đang bị trộn lẫn nhiều ngôn ngữ.
           -> BẮT BUỘC: Nếu một chuỗi (hoặc câu) ĐÃ LÀ {target_lang} rồi, bạn PHẢI GIỮ NGUYÊN toàn bộ chuỗi đó. Tuyệt đối không được dịch ngược nó sang Tiếng Việt hay bất kỳ ngôn ngữ nào khác.
           -> Chỉ tiến hành dịch những chuỗi CHƯA PHẢI là {target_lang} sang {target_lang}.
        2. BẢO TOÀN DỮ LIỆU: Giữ nguyên các thuật ngữ kỹ thuật, tên riêng (như địa danh Xaysomboun, Vientiane, tên công ty...), thông số đo lường và con số.
        3. ĐỊNH DẠNG ĐẦU RA: BẮT BUỘC trả về DUY NHẤT một mảng JSON chứa các chuỗi kết quả. Mảng này PHẢI CÓ ĐÚNG {len(chunk)} phần tử và giữ TUYỆT ĐỐI NGUYÊN THỨ TỰ so với đầu vào.
        4. LÀM SẠCH: KHÔNG kèm theo các ký tự định dạng khối mã Markdown (như ```json) hay bất kỳ câu chữ giải thích nào khác ngoài mảng JSON.
        """
        
        # Áp dụng bộ lọc tùy chọn từ phía người dùng
        if skip_english:
            prompt += "\n- CẤM DỊCH TIẾNG ANH: Nếu phát hiện chuỗi văn bản hoàn toàn bằng Tiếng Anh (hoặc là các mã hiệu kỹ thuật dạng ký tự Latinh), bạn phải GIỮ NGUYÊN GỐC, tuyệt đối không dịch."
        if skip_abbreviations:
            prompt += "\n- CẤM DỊCH TỪ VIẾT TẮT: Nếu gặp các từ viết tắt viết hoa chuyên ngành (Ví dụ: BOQ, OPBRC, PRTC, BTCT, VAT, XD, USD, ...), bạn phải GIỮ NGUYÊN KÝ TỰ, tuyệt đối không dịch nghĩa chúng."
        
        prompt += f"\n\nINPUT JSON:\n{json.dumps(chunk, ensure_ascii=False)}"
        
        max_retries = 3
        for attempt in range(max_retries):
            try:
                response = model.generate_content(prompt)
                raw_text = response.text.strip().replace("```json", "").replace("```", "").strip()
                translated_chunk = json.loads(raw_text)
                
                if len(translated_chunk) == len(chunk):
                    for orig, trans in zip(chunk, translated_chunk):
                        mapping[orig] = trans
                    break
            except Exception as e:
                if attempt == max_retries - 1:
                    for orig in chunk: mapping[orig] = orig
                time.sleep(3)
                
        time.sleep(2.5)  # Giãn cách an toàn chủ động chống lỗi 429
        
    status.empty()
    progress_bar.empty()
    percent_text.empty()
    return mapping

def set_docx_font(run, font_name):
    """Hỗ trợ cấu hình font chữ chuẩn xác trong Word ngăn lỗi ô vuông"""
    run.font.name = font_name
    run._element.rPr.rFonts.set(qn('w:eastAsia'), font_name)

# =====================================================================
# 3. GIAO DIỆN CHÍNH APPLICATION
# =====================================================================
st.sidebar.title("🛠️ CÔNG CỤ XỬ LÝ AI")
app_mode = st.sidebar.radio("Chọn chức năng hoạt động:", ["📄 1. Chức năng OCR (Đọc PDF)", "🌐 2. Chức năng Dịch (Excel/Word)"])

st.sidebar.markdown("---")
api_key = st.sidebar.text_input("🔑 Nhập Gemini API Key:", type="password")
model_choice = st.sidebar.selectbox("🤖 Chọn AI Model:", ["gemini-3.5-flash", "gemini-2.5-pro"])

if not api_key:
    st.warning("Vui lòng nhập API Key ở thanh sidebar bên trái để sử dụng phần mềm.")
    st.stop()

genai.configure(api_key=api_key)
model = genai.GenerativeModel(model_choice)

# =====================================================================
# MODULE 1: OCR PDF (XUẤT SẠCH, BỎ THỂ RÁC HTML)
# =====================================================================
if "OCR" in app_mode:
    st.title("📄 Ứng dụng OCR PDF sang Excel/Word")
    
    batch_size = st.sidebar.number_input("Số trang mỗi mẻ (Batch Size):", min_value=1, max_value=20, value=3)
    export_mode = st.sidebar.radio("Lựa chọn định dạng xuất file:", ["📊 Xuất EXCEL (Chuyên trị Bảng biểu)", "📝 Xuất WORD (Giữ nguyên Định dạng)"])
    remove_graphics = st.sidebar.checkbox("🚫 Tự động bỏ qua Logo, Con dấu đỏ, Chữ ký", value=True)
    
    uploaded_file = st.file_uploader("Tải lên file PDF cần OCR", type=["pdf"])

    if uploaded_file and st.button("Bắt đầu xử lý OCR"):
        st.session_state.processed_data = None
        try:
            doc = fitz.open(stream=uploaded_file.read(), filetype="pdf")
            total_pages = len(doc)
            total_batches = math.ceil(total_pages / batch_size)
            st.info(f"📁 Hệ thống nhận diện PDF gồm: {total_pages} trang.")
            
            all_dataframes, all_text_chunks = [], []
            progress_bar = st.progress(0)
            status_text = st.empty()

            for batch_idx in range(total_batches):
                start_page = batch_idx * batch_size
                end_page = min(start_page + batch_size, total_pages) - 1
                status_text.info(f"⏳ Đang OCR cấu trúc Phần {batch_idx + 1}/{total_batches}...")
                
                chunk_pdf = fitz.open()
                chunk_pdf.insert_pdf(doc, from_page=start_page, to_page=end_page)
                fd, tmp_file_path = tempfile.mkstemp(suffix=".pdf")
                os.close(fd) 
                chunk_pdf.save(tmp_file_path)
                chunk_pdf.close()
                
                ai_response_text = "" 
                for attempt in range(3):
                    try:
                        gemini_file = genai.upload_file(path=tmp_file_path, display_name=f"OCR_{batch_idx}")
                        while gemini_file.state.name == "PROCESSING":
                            time.sleep(2)
                            gemini_file = genai.get_file(gemini_file.name)
                            
                        if "EXCEL" in export_mode:
                            prompt = "Bạn là chuyên gia trích xuất dữ liệu. Trích xuất BẢNG DỮ LIỆU. BẮT BUỘC dùng định dạng TSV (phân cách bằng đúng 1 dấu TAB \\t). Trả về dữ liệu thô."
                        else:
                            prompt = "Chép lại nội dung văn bản. Giữ nguyên định dạng lề. In đậm bằng **văn bản**. Tiêu đề dùng dấu #. Bảng biểu kẻ bằng định dạng hàng dọc Markdown (|). Tuyệt đối không sinh thẻ HTML <br>."
                        
                        if remove_graphics: prompt += "\n- QUAN TRỌNG: TUYỆT ĐỐI BỎ QUA không đọc logo, hình ảnh con dấu, hoặc chữ ký."

                        response = model.generate_content([prompt, gemini_file])
                        ai_response_text = response.text.strip()
                        genai.delete_file(gemini_file.name)
                        if os.path.exists(tmp_file_path): os.remove(tmp_file_path)
                        break 
                    except Exception as e:
                        if attempt == 2: st.error(f"Lỗi kết nối API: {e}")
                        time.sleep(5)

                if ai_response_text:
                    if "EXCEL" in export_mode:
                        csv_text = ai_response_text.replace("```tsv", "").replace("```csv", "").replace("```", "").strip()
                        try:
                            df = pd.read_csv(io.StringIO(csv_text), sep="\t", quoting=csv.QUOTE_NONE, on_bad_lines='skip', engine='python')
                            all_dataframes.append(df)
                        except: pass
                    else:
                        # Làm sạch hoàn toàn thẻ rác HTML để tránh lỗi chữ <br> trong Word
                        cl_text = ai_response_text.replace("```markdown", "").replace("```", "")
                        cl_text = re.sub(r'<br\s*/?>', '\n', cl_text, flags=re.IGNORECASE).strip()
                        all_text_chunks.append(cl_text)
                
                progress_bar.progress((batch_idx + 1) / total_batches)
                time.sleep(4)

            doc.close()
            status_text.success("🎉 Quy trình xử lý OCR đã hoàn thành thành công!")

            # ĐÓNG GÓI CHUẨN ĐỊNH DẠNG ĐẦU RA
            if "EXCEL" in export_mode and all_dataframes:
                final_df = pd.concat(all_dataframes, ignore_index=True)
                excel_output = io.BytesIO()
                with pd.ExcelWriter(excel_output, engine='xlsxwriter') as writer:
                    final_df.to_excel(writer, index=False, sheet_name='DuLieu_OCR')
                st.session_state.processed_data = excel_output.getvalue()
                st.session_state.export_mode_used = "EXCEL"
                st.session_state.preview_data = final_df.head(20)
            elif "WORD" in export_mode and all_text_chunks:
                word_doc = Document()
                full_text = "\n\n".join(all_text_chunks)
                
                in_table = False
                current_table = None

                for line in full_text.split('\n'):
                    line = line.strip()
                    if not line:
                        in_table = False
                        continue
                    
                    # Thuật toán tự dựng lưới bảng Grid trong Microsoft Word
                    if line.startswith('|') and line.endswith('|'):
                        cells = [c.strip() for c in line.split('|')[1:-1]]
                        if all(c.replace('-', '').replace(':', '').strip() == '' for c in cells):
                            continue
                            
                        if not in_table:
                            in_table = True
                            current_table = word_doc.add_table(rows=1, cols=len(cells))
                            current_table.style = 'Table Grid'
                            row_cells = current_table.rows[0].cells
                            for i, cell in enumerate(cells):
                                if i < len(row_cells):
                                    row_cells[i].text = cell.replace('**', '')
                        else:
                            row_cells = current_table.add_row().cells
                            for i, cell in enumerate(cells):
                                if i < len(row_cells):
                                    row_cells[i].text = cell.replace('**', '')
                        continue
                    else:
                        in_table = False
                    
                    align = WD_ALIGN_PARAGRAPH.LEFT
                    if "<center>" in line and "</center>" in line:
                        align = WD_ALIGN_PARAGRAPH.CENTER
                        line = line.replace("<center>", "").replace("</center>", "")
                    
                    if line.startswith('#'):
                        level = len(line) - len(line.lstrip('#'))
                        text_content = line.lstrip('#').strip()
                        p = word_doc.add_heading(text_content, level=min(level, 9))
                        p.alignment = align
                        continue
                    
                    p = word_doc.add_paragraph()
                    p.alignment = align
                    parts = re.split(r'(\*\*.*?\*\*)', line)
                    for part in parts:
                        if part.startswith('**') and part.endswith('**'):
                            run = p.add_run(part[2:-2])
                            run.bold = True
                        else:
                            p.add_run(part)
                            
                word_output = io.BytesIO()
                word_doc.save(word_output)
                st.session_state.processed_data = word_output.getvalue()
                st.session_state.export_mode_used = "WORD"
                st.session_state.preview_data = full_text
        except Exception as e:
            st.error(f"❌ Lỗi hệ thống: {e}")

    # ===== HIỂN THỊ NÚT TẢI TIÊU CHUẨN PHÙ HỢP HOÀN HẢO VỚI LAUNCHER LAUNCHER =====
    if st.session_state.processed_data:
        st.markdown("---")
        st.success("✅ **KẾT QUẢ ĐÃ SẴN SÀNG!**")
        
        if st.session_state.export_mode_used == "EXCEL":
            st.dataframe(st.session_state.preview_data)
            st.download_button(
                label="📥 Tải về file EXCEL (.xlsx)",
                data=st.session_state.processed_data,
                file_name="Ket_Qua_OCR.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                use_container_width=True
            )
        else:
            with st.expander("Hiển thị dữ liệu văn bản thô (Xem trước)"):
                st.text(st.session_state.preview_data)
            st.download_button(
                label="📥 Tải về file WORD (.docx)",
                data=st.session_state.processed_data,
                file_name="Ket_Qua_OCR.docx",
                mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                use_container_width=True
            )

# =====================================================================
# MODULE 2: DỊCH TÀI LIỆU (BẢO TOÀN TRẠNG THÁI MERGE & BOLD)
# =====================================================================
elif "Dịch" in app_mode:
    st.title("🌐 Ứng dụng Dịch Tài Liệu (Bảo toàn cấu trúc ô gốc)")
    
    col1, col2 = st.columns(2)
    with col1:
        target_lang = st.selectbox("Lựa chọn ngôn ngữ đích cần dịch sang:", ["Tiếng Lào", "Tiếng Việt", "Tiếng Anh"])
    with col2:
        font_map = {"Tiếng Lào": "Saysettha Laos OT", "Tiếng Việt": "Times New Roman", "Tiếng Anh": "Times New Roman"}
        target_font = font_map[target_lang]
        st.info(f"🔤 Font hệ thống áp dụng cho bản dịch: **{target_font}**")

    st.sidebar.subheader("⚙️ Bộ lọc dịch thuật nâng cao")
    skip_english = st.sidebar.checkbox("🔤 Giữ nguyên gốc Tiếng Anh (Không dịch)", value=True)
    skip_abbreviations = st.sidebar.checkbox("🔽 Giữ nguyên từ viết tắt (BOQ, BTCT, ...)", value=True)

    trans_file = st.file_uploader("Tải lên tệp tin Excel (.xlsx) hoặc Word (.docx) cần biên dịch", type=["xlsx", "docx"])

    if trans_file:
        file_ext = trans_file.name.split('.')[-1].lower()
        
        # --- LUỒNG XỬ LÝ FILE EXCEL ---
        if file_ext == "xlsx":
            wb = openpyxl.load_workbook(trans_file)
            sheet_names = wb.sheetnames
            selected_sheet = st.selectbox("Chọn tên Sheet cần tiến hành dịch:", sheet_names)
            
            if st.button(f"🚀 Tiến hành dịch Sheet '{selected_sheet}'"):
                st.session_state.trans_processed_data = None
                st.session_state.trans_file_name = None
                ws = wb[selected_sheet]
                
                # Quét trích xuất chữ không làm vỡ các ô đã Merge
                texts_to_trans = set()
                for row in ws.iter_rows():
                    for cell in row:
                        if cell.value and isinstance(cell.value, str):
                            val = str(cell.value).strip()
                            if val and not val.isnumeric() and len(val) > 1:
                                texts_to_trans.add(val)
                
                texts_list = list(texts_to_trans)
                st.info(f"🔎 Đã phát hiện {len(texts_list)} cụm chuỗi văn bản độc lập. Đang chuẩn bị dịch thuật...")
                
                translation_map = translate_texts_batch(texts_list, target_lang, model, skip_english, skip_abbreviations)
                
                # Đổ ngược dữ liệu dịch vào bảng tính (Bảo toàn định dạng cũ)
                for row in ws.iter_rows():
                    for cell in row:
                        if cell.value and isinstance(cell.value, str):
                            val = str(cell.value).strip()
                            if val in translation_map:
                                cell.value = translation_map[val]
                                old_font = cell.font
                                is_bold = old_font.bold if old_font else False
                                is_italic = old_font.italic if old_font else False
                                cell.font = Font(name=target_font, bold=is_bold, italic=is_italic)
                
                excel_io = io.BytesIO()
                wb.save(excel_io)
                st.session_state.trans_processed_data = excel_io.getvalue()
                st.session_state.trans_file_name = f"Dich_{target_lang}_{trans_file.name}"
                st.success("🎉 Dịch cấu trúc Sheet hoàn thành! File đã sẵn sàng để tải về.")

        # --- LUỒNG XỬ LÝ FILE WORD ---
        elif file_ext == "docx":
            trans_mode = st.radio("Cấu hình chế độ hiển thị bản dịch Word:", ["Dịch thay thế (Ghi đè nội dung gốc)", "Dịch song song (Chèn dòng dịch in đậm dưới gốc)"])
            
            if st.button("🚀 Tiến hành dịch file Word"):
                st.session_state.trans_processed_data = None
                st.session_state.trans_file_name = None
                doc = Document(trans_file)
                
                all_paragraphs = list(doc.paragraphs)
                for table in doc.tables:
                    for row in table.rows:
                        for cell in row.cells:
                            all_paragraphs.extend(cell.paragraphs)
                
                texts_to_trans = []
                for p in all_paragraphs:
                    val = p.text.strip()
                    if val and not val.isnumeric() and len(val) > 1 and val not in texts_to_trans:
                        texts_to_trans.append(val)
                        
                st.info(f"🔎 Tổng số đoạn văn bản phát hiện: {len(texts_to_trans)} đoạn. Đang kết nối dịch thuật...")
                translation_map = translate_texts_batch(texts_to_trans, target_lang, model, skip_english, skip_abbreviations)
                
                for p in all_paragraphs:
                    orig_val = p.text.strip()
                    if orig_val in translation_map:
                        trans_val = translation_map[orig_val]
                        if "song song" in trans_mode:
                            run = p.add_run("\n" + trans_val)
                            run.bold = True
                            set_docx_font(run, target_font)
                        else:
                            p.clear()
                            run = p.add_run(trans_val)
                            set_docx_font(run, target_font)

                word_io = io.BytesIO()
                doc.save(word_io)
                st.session_state.trans_processed_data = word_io.getvalue()
                st.session_state.trans_file_name = f"Dich_{target_lang}_{trans_file.name}"
                st.success("🎉 Bản dịch tài liệu Word đã sẵn sàng!")

    # ===== HIỂN THỊ NÚT TẢI CHO DỊCH THUẬT =====
    if st.session_state.trans_processed_data:
        st.markdown("---")
        mime_type = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet" if st.session_state.trans_file_name.endswith("xlsx") else "application/vnd.openxmlformats-officedocument.wordprocessingml.document"
        
        st.download_button(
            label="📥 BẤM VÀO ĐÂY ĐỂ TẢI BẢN DỊCH VỀ MÁY",
            data=st.session_state.trans_processed_data,
            file_name=st.session_state.trans_file_name,
            mime=mime_type,
            use_container_width=True
        )
