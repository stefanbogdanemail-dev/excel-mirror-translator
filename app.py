import io
import re
import copy
import unicodedata
import pandas as pd
import streamlit as st
from openpyxl import load_workbook, Workbook
from openpyxl.styles import PatternFill, Font, Border, Side
from deep_translator import GoogleTranslator

st.set_page_config(page_title="Excel Mirror Translator", page_icon="📊", layout="wide")

BORDER_MISSING = Border(
    left=Side(style='medium', color='FFC000'),
    right=Side(style='medium', color='FFC000'),
    top=Side(style='medium', color='FFC000'),
    bottom=Side(style='medium', color='FFC000')
)

BORDER_NOT_FOUND = Border(
    left=Side(style='medium', color='FF0000'),
    right=Side(style='medium', color='FF0000'),
    top=Side(style='medium', color='FF0000'),
    bottom=Side(style='medium', color='FF0000')
)

def normalize_text(text: str) -> str:
    if not isinstance(text, str):
        text = str(text)
    text = text.lower().strip()
    replacements = {'ă': 'a', 'â': 'a', 'î': 'i', 'ș': 's', 'ț': 't', 'ş': 's', 'ţ': 't'}
    for k, v in replacements.items():
        text = text.replace(k, v)
    nfd = unicodedata.normalize('NFD', text)
    text = ''.join(c for c in nfd if unicodedata.category(c) != 'Mn')
    text = re.sub(r'\s+', ' ', text)
    return text.strip()

def is_purely_numeric(val) -> bool:
    if val is None:
        return False
    if isinstance(val, (int, float)):
        return True
    s = str(val).strip()
    has_digit = any(c.isdigit() for c in s)
    only_chars = bool(re.match(r'^[\s\d.,\-–%/:]+$', s))
    return (has_digit and only_chars) or (s == "-")

def clone_cell_style(src_cell, dst_cell):
    if src_cell.has_style:
        dst_cell.font = copy.copy(src_cell.font)
        dst_cell.fill = copy.copy(src_cell.fill)
        dst_cell.border = copy.copy(src_cell.border)
        dst_cell.alignment = copy.copy(src_cell.alignment)
        dst_cell.number_format = copy.copy(src_cell.number_format)

def translate_with_google(text: str) -> tuple[str, str]:
    try:
        translated = GoogleTranslator(source='ro', target='en').translate(text)
        if not translated:
            return text, "NOT FOUND"
        if translated.lower() == text.lower() and len(text) <= 4:
            return text, "NOT FOUND"
        return translated, "MISSING"
    except Exception:
        return text, "NOT FOUND"

def run_translation_pipeline(input_bytes, dict_bytes):
    wb = load_workbook(io.BytesIO(input_bytes), data_only=False)
    wb_vals = load_workbook(io.BytesIO(input_bytes), data_only=True)
    ext_wb = load_workbook(io.BytesIO(dict_bytes)) if dict_bytes else None

    # 1. Încărcare dicționar
    dictionary = {}
    sheets_to_read = []
    if ext_wb:
        for s in ext_wb.sheetnames:
            sheets_to_read.append(ext_wb[s])
    for s in wb.sheetnames:
        if normalize_text(s) in ["dictionar", "dictionary", "glosar"]:
            sheets_to_read.append(wb[s])

    for sheet in sheets_to_read:
        for r in range(1, sheet.max_row + 1):
            src = sheet.cell(r, 1).value
            trg = sheet.cell(r, 2).value
            if src and trg:
                dictionary[normalize_text(str(src))] = str(trg).strip()

    out_wb = Workbook()
    out_wb.remove(out_wb.active)

    review_rows = []
    stats = {"MATCH": 0, "MISSING": 0, "NOT FOUND": 0, "SKIP": 0}
    unique_new_terms = {}

    source_sheets = wb.sheetnames

    for sheet_name in source_sheets:
        src_sheet = wb[sheet_name]
        val_sheet = wb_vals[sheet_name]
        s_norm = normalize_text(sheet_name)

        if s_norm in ["dictionar", "dictionary", "glosar"] or sheet_name.startswith('_'):
            new_s = out_wb.create_sheet(title=sheet_name[:31])
            for r in range(1, src_sheet.max_row + 1):
                for c in range(1, src_sheet.max_column + 1):
                    cell = src_sheet.cell(r, c)
                    dst = new_s.cell(r, c, cell.value)
                    clone_cell_style(cell, dst)
            continue

        orig_twin = out_wb.create_sheet(title=sheet_name[:31])
        for r in range(1, src_sheet.max_row + 1):
            for c in range(1, src_sheet.max_column + 1):
                cell = src_sheet.cell(r, c)
                dst = orig_twin.cell(r, c, cell.value)
                clone_cell_style(cell, dst)

        tr_title = f"{sheet_name[:25]} (EN)"[:31]
        tr_twin = out_wb.create_sheet(title=tr_title)

        for col_letter, col_dim in src_sheet.column_dimensions.items():
            tr_twin.column_dimensions[col_letter].width = col_dim.width
        for row_idx, row_dim in src_sheet.row_dimensions.items():
            tr_twin.row_dimensions[row_idx].height = row_dim.height
        for m_range in src_sheet.merged_cells.ranges:
            tr_twin.merge_cells(str(m_range))

        merged_non_anchors = set()
        for rng in src_sheet.merged_cells.ranges:
            for r in range(rng.min_row, rng.max_row + 1):
                for c in range(rng.min_col, rng.max_col + 1):
                    if not (r == rng.min_row and c == rng.min_col):
                        merged_non_anchors.add((r, c))

        for r in range(1, src_sheet.max_row + 1):
            for c in range(1, src_sheet.max_column + 1):
                src_c = src_sheet.cell(r, c)
                val_c = val_sheet.cell(r, c)
                dst_c = tr_twin.cell(r, c)
                clone_cell_style(src_c, dst_c)

                if (r, c) in merged_non_anchors:
                    stats["SKIP"] += 1
                    continue

                val = val_c.value
                if val is None or str(val).strip() == "" or is_purely_numeric(val):
                    dst_c.value = src_c.value
                    stats["SKIP"] += 1
                    continue

                val_str = str(val).strip()
                val_norm = normalize_text(val_str)

                # Prioritate Dicționar
                if val_norm in dictionary:
                    dst_c.value = dictionary[val_norm]
                    stats["MATCH"] += 1
                else:
                    tr_text, status = translate_with_google(val_str)
                    dst_c.value = tr_text
                    stats[status] += 1
                    if status == "MISSING":
                        dst_c.border = BORDER_MISSING
                        review_rows.append((tr_title, r, c, val_str, tr_text, "MISSING"))
                        if val_str not in unique_new_terms:
                            unique_new_terms[val_str] = tr_text
                    else:
                        dst_c.border = BORDER_NOT_FOUND
                        review_rows.append((tr_title, r, c, val_str, tr_text, "NOT FOUND"))

    # Tab _Review Summary
    summary = out_wb.create_sheet(title="_Review Summary", index=0)
    summary.append(["Sheet", "Row", "Column", "Original", "Translation", "Status"])
    for col_idx in range(1, 7):
        c = summary.cell(1, col_idx)
        c.font = Font(bold=True)
        c.fill = PatternFill(start_color="D9D9D9", fill_type="solid")

    for row_data in review_rows:
        summary.append(list(row_data))
        status_cell = summary.cell(summary.max_row, 6)
        color = "FF0000" if row_data[5] == "NOT FOUND" else "FFC000"
        status_cell.fill = PatternFill(start_color=color, fill_type="solid")

    return out_wb, stats, unique_new_terms

# --- UI Streamlit ---
st.title("📊 Excel Mirror Translator cu Validare")
st.write("Traducere automată prioritară pe dicționar, cu fallback Google Translate și revizuire manuală.")

col1, col2 = st.columns(2)
with col1:
    uploaded_file = st.file_uploader("1. Fișier Excel de Tradus (.xlsx)", type=["xlsx"])
with col2:
    uploaded_dict = st.file_uploader("2. Dicționar Extern (.xlsx)", type=["xlsx"])

if "translated_wb" not in st.session_state:
    st.session_state.translated_wb = None
    st.session_state.new_terms_df = None
    st.session_state.stats = None

if st.button("🚀 Pasul 1: Generează Traducerea", type="primary"):
    if not uploaded_file:
        st.error("Încarcă un fișier Excel!")
    else:
        with st.spinner("Se procesează fișierul..."):
            wb, stats, new_terms = run_translation_pipeline(
                uploaded_file.getvalue(),
                uploaded_dict.getvalue() if uploaded_dict else None
            )
            st.session_state.translated_wb = wb
            st.session_state.stats = stats
            
            # Construim DataFrame pentru revizuire
            df_data = []
            for src, trg in new_terms.items():
                df_data.append({"Validează": True, "Termen Română": src, "Traducere Engleză (Editabil)": trg})
            
            st.session_state.new_terms_df = pd.DataFrame(df_data)

if st.session_state.translated_wb is not None:
    stats = st.session_state.stats
    st.success(f"Statistici: Dicționar (MATCH): {stats['MATCH']} | Noi Google Translate (MISSING): {stats['MISSING']} | Netraduse (NOT FOUND): {stats['NOT FOUND']}")

    st.markdown("### 📝 Pasul 2: Revizuiește termenii noi pentru Dicționar")
    st.info("Termenii de mai jos au fost traduși automat. Poți modifica traducerea direct în tabel sau debifa termenii pe care NU dorești să îi salvezi în dicționar.")

    if not st.session_state.new_terms_df.empty:
        edited_df = st.data_editor(
            st.session_state.new_terms_df,
            column_config={
                "Validează": st.column_config.CheckboxColumn("Salvează în Dicționar?", default=True),
                "Termen Română": st.column_config.TextColumn("Termen Română", disabled=True),
                "Traducere Engleză (Editabil)": st.column_config.TextColumn("Traducere Engleză (Editează dacă e greșit)")
            },
            hide_index=True,
            use_container_width=True
        )
    else:
        st.write("Toți termenii au existat deja în dicționar. Nu este nevoie de revizuire!")
        edited_df = pd.DataFrame()

    # Buton Final de Descărcare
    st.markdown("### 💾 Pasul 3: Descarcă fișierul cu dicționarul actualizat")
    
    if st.button("Finalizează și Pregătește Fișierul"):
        out_wb = copy.copy(st.session_state.translated_wb)

        # Adăugăm doar termenii bifați de utilizator
        if not edited_df.empty:
            dict_sheet = None
            for s in out_wb.sheetnames:
                if normalize_text(s) in ["dictionar", "dictionary", "glosar"]:
                    dict_sheet = out_wb[s]
                    break
            if not dict_sheet:
                dict_sheet = out_wb.create_sheet(title="Dictionar")
                dict_sheet.append(["TERMEN SURSA", "TRADUCERE", "STATUS"])

            dict_sheet.append([])
            h = dict_sheet.cell(dict_sheet.max_row + 1, 1, "TERMENI NOI - CONFIRMATI DE UTILIZATOR")
            h.font = Font(bold=True, italic=True)

            for _, row in edited_df.iterrows():
                if row["Validează"]:
                    dict_sheet.append([row["Termen Română"], row["Traducere Engleză (Editabil)"], "Confirmat"])

        output_buffer = io.BytesIO()
        out_wb.save(output_buffer)
        output_buffer.seek(0)

        st.download_button(
            label="📥 Descarcă Fișierul Excel Final (.xlsx)",
            data=output_buffer,
            file_name=f"translated_{uploaded_file.name}",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
