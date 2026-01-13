import streamlit as st
import pandas as pd
from docxtpl import DocxTemplate, RichText
import io
import zipfile
import re

# ===============================
# 0️⃣ 基礎設定
# ===============================
try:
    st.set_page_config(page_title="節能績效計劃書生成器", page_icon="🏭", layout="wide")
except Exception:
    pass

st.title("📊 節能績效計劃書生成器：2.3")
st.markdown("""

### ⚠️ 重要使用說明

1.  **Word 模板變數寫法：** `{{變數名稱}}` (程式會自動變紅字，不要加 r)
2.  **Excel 設定：**
    * **Sheet 1**: 變數設定 (A欄名稱, B欄數值)。
    * **Sheet 2+**: 表格資料 (Sheet 名稱需對應 Word 標籤)。

""")

# ===============================
# 1️⃣ 格式規則
# ===============================
FORMAT_RULES = {
    "me_prefix": {"description": "ME 類：千分位 + 保留原始小數"},
    "decimal_2": {"keywords": ["_rate", "elec_price", "new_cop_std", "new_eff_std"], "description": "2 位小數"},
    "decimal_1": {"keywords": ["_year"], "description": "1 位小數"},
    "integer": {"description": "整數（預設）"},
}

def clean_text(val):
    if pd.isna(val): return ""
    s = str(val).strip()
    if s.lower() in ["nan", "none", "nat", ""]: return ""
    return s

def process_value_to_richtext(val, key_name=""):
    val_str = clean_text(val)
    if val_str == "": return ""
    if any(x in val_str for x in ["~", "CH", "CWP", "HP", "/", "New", "new"]): return val_str
    
    try:
        float_val = float(val_str)
        key_lower = str(key_name).lower()
        formatted = val_str

        if key_lower.startswith("me_"):
            if "." in val_str:
                parts = val_str.split(".")
                formatted = f"{int(parts[0]):,}.{parts[1]}"
            else:
                formatted = f"{int(float_val):,}"
        elif any(k in key_lower for k in FORMAT_RULES["decimal_2"]["keywords"]):
            formatted = f"{float_val:,.2f}"
        elif any(k in key_lower for k in FORMAT_RULES["decimal_1"]["keywords"]):
            formatted = f"{float_val:,.1f}"
        else:
            formatted = f"{float_val:,.0f}"

        rt = RichText()
        rt.add(formatted, color="FF0000", bold=False)
        return rt
    except ValueError:
        return val_str

# ==========================================
# 2️⃣ 核心邏輯：通用表格讀取
# ==========================================
def get_clean_table_data(excel_file, sheet_name):
    try:
        # 1. 預讀找標題
        df_preview = pd.read_excel(excel_file, sheet_name=sheet_name, header=None, nrows=20, dtype=str)
        header_row = -1
        
        target_name = ["名稱", "name", "設備名稱"]
        target_no = ["no", "編號", "設備編號", "那台冰水主機代號"]
        
        for i, row in df_preview.iterrows():
            row_clean = [str(x).strip().lower() for x in row.values]
            row_str = " ".join(row_clean)
            if any(k in row_str for k in target_name) and any(k in row_str for k in target_no):
                header_row = i
                break
        
        if header_row == -1: return []
            
        # 2. 正式讀取
        df = pd.read_excel(excel_file, sheet_name=sheet_name, header=header_row, dtype=str)
        df = df.loc[:, ~df.columns.str.contains('^Unnamed', na=False)]
        df.columns = [str(c).strip() for c in df.columns]
        
        col_map = {}
        for c in df.columns:
            c_low = c.lower()
            if c in target_name: col_map['name'] = c
            if c in target_no: col_map['no'] = c
            if 'name' not in col_map and ('名稱' in c or 'name' in c_low): col_map['name'] = c
            if 'no' not in col_map and ('代號' in c or '編號' in c or 'no' in c_low): col_map['no'] = c

        if 'name' not in col_map or 'no' not in col_map: return []

        df['standard_name'] = df[col_map['name']]
        df['standard_no'] = df[col_map['no']]

        df = df.dropna(subset=['standard_name', 'standard_no'])
        df = df[~df['standard_name'].str.contains('名稱|Equipment|name', case=False, na=False)]
        df = df[~df['standard_name'].str.lower().isin(['nan', 'none', ''])]

        results = []
        for _, row in df.iterrows():
            row_dict = {}
            for col in df.columns:
                if col in ['standard_name', 'standard_no']: continue
                val = clean_text(row[col])
                if val.endswith(".0"):
                    try: val = str(int(float(val)))
                    except: pass
                row_dict[col] = val
            
            row_dict['name'] = clean_text(row[col_map['name']])
            row_dict['no'] = clean_text(row[col_map['no']])
            results.append(row_dict)
            
        return results
    except:
        return []

# ==========================================
# 3️⃣ 核心邏輯：彈性處理 (分系統編號)
# ==========================================
def process_dynamic_context(context, excel_file):
    xls = pd.ExcelFile(excel_file)
    all_sheets = xls.sheet_names
    
    groups = { "before": [], "after": [] }
    
    def get_sort_weight(name):
        if "主機" in name or "chiller" in name.lower(): return 1
        if "泵" in name or "pump" in name.lower(): return 2
        if "水塔" in name or "tower" in name.lower(): return 3
        return 4

    # 讀取並分類
    for sheet in all_sheets:
        if sheet == "變數": continue
        
        data = get_clean_table_data(excel_file, sheet)
        if not data: continue
        
        if "改善前" in sheet:
            groups["before"].append((sheet, data))
        elif "改善後" in sheet:
            groups["after"].append((sheet, data))
        else:
            context[sheet] = data

    groups["before"].sort(key=lambda x: get_sort_weight(x[0]))
    groups["after"].sort(key=lambda x: get_sort_weight(x[0]))

    # 編號邏輯
    def apply_numbering(sheet_list):
        pm_counter = 1
        fm_counter = 1
        t_counter = 1
        
        # 1. PM 編號
        for sheet_name, items in sheet_list:
            for item in items:
                item['pm'] = f"PM{pm_counter}"; pm_counter += 1
        
        # 2. FM/T 編號 (分系統)
        chiller_lists = []
        for sheet_name, items in sheet_list:
            if "主機" in sheet_name or "chiller" in sheet_name.lower():
                chiller_lists.append(items)
        
        # 冰水側
        for items in chiller_lists:
            for item in items:
                item['evap_fm'] = f"FM{fm_counter}"; fm_counter += 1
                item['evap_t_out'] = f"T{t_counter}"; 
                item['evap_t_in'] = f"T{t_counter+1}"; t_counter += 2
        # 冷卻水側
        for items in chiller_lists:
            for item in items:
                item['cond_fm'] = f"FM{fm_counter}"; fm_counter += 1
                item['cond_t_out'] = f"T{t_counter}"; 
                item['cond_t_in'] = f"T{t_counter+1}"; t_counter += 2

        # 3. 寫回 Context + 水泵分流
        for sheet_name, items in sheet_list:
            context[sheet_name] = items
            
            if "泵" in sheet_name or "pump" in sheet_name.lower():
                ice_pumps, cool_pumps, zone_pumps, other_pumps = [], [], [], []
                
                for item in items:
                    name_str = str(item.get('name', ''))
                    no_str = str(item.get('no', '')).upper() # 轉大寫比對
                    
                    # === 精確拆分邏輯 ===
                    # 1. 區域泵 (ZP or 區域)
                    if 'ZP' in no_str or '區域' in name_str:
                        zone_pumps.append(item)
                    # 2. 冷卻水泵 (CWP or 冷卻)
                    elif 'CWP' in no_str or '冷卻' in name_str:
                        cool_pumps.append(item)
                    # 3. 冰水泵 (CHP or 冰水)
                    elif 'CHP' in no_str or '冰水' in name_str:
                        ice_pumps.append(item)
                    # 4. 其他
                    else:
                        other_pumps.append(item)
                
                context[f"{sheet_name}_冰水"] = ice_pumps
                context[f"{sheet_name}_冷卻"] = cool_pumps
                context[f"{sheet_name}_區域"] = zone_pumps
                context[f"{sheet_name}_其他"] = other_pumps

    apply_numbering(groups["before"])
    apply_numbering(groups["after"])

    return context

# ===============================
# 4️⃣ 主 UI
# ===============================
col1, col2 = st.columns(2)
with col1:
    uploaded_excel = st.file_uploader("1️⃣ 上傳 Excel", type="xlsx")
with col2:
    uploaded_templates = st.file_uploader("2️⃣ 上傳 Word 模板", type="docx", accept_multiple_files=True)

if uploaded_excel and uploaded_templates:
    if st.button("🚀 生成報告", type="primary"):
        try:
            context = {}
            st.toast("處理資料中...")
            
            # A. 變數
            try:
                xl = pd.ExcelFile(uploaded_excel)
                s_name = "變數" if "變數" in xl.sheet_names else xl.sheet_names[0]
                df_var = xl.parse(s_name, header=None)
                for i, row in df_var.iterrows():
                    if pd.isna(row[0]): continue
                    key = str(row[0]).strip()
                    val = row[1] if len(row) > 1 else ""
                    context[key] = process_value_to_richtext(val, key)
            except: pass

            # B. 設備處理
            context = process_dynamic_context(context, uploaded_excel)

            # C. 渲染 Word
            zip_buffer = io.BytesIO()
            with zipfile.ZipFile(zip_buffer, "w") as zf:
                for tpl in uploaded_templates:
                    tpl.seek(0)
                    doc = DocxTemplate(tpl)
                    doc.render(context)
                    out = io.BytesIO()
                    doc.save(out)
                    zf.writestr(f"Result_{tpl.name}", out.getvalue())
            
            st.success("✅ 報告生成成功！")
            

            st.download_button("📦 下載結果 (ZIP)", zip_buffer.getvalue(), "Reports.zip", "application/zip")
            
        except Exception as e:
            st.error(f"發生錯誤: {e}")