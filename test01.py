import streamlit as st
import pandas as pd
from docxtpl import DocxTemplate, RichText
import io

# ---------------- 處理數字邏輯（只用在「單一變數」Sheet） ----------------
def process_value_to_richtext(val, key_name=""):
    if pd.isna(val):
        return ""

    val_str = str(val).strip()
    if val_str == "":
        return ""

    if "~" in val_str or "～" in val_str:
        rt = RichText()
        rt.add(val_str, color="000000", bold=False) # 強制黑色、不加粗
        return rt

    is_number = False
    float_val = 0.0

    try:
        # 排除日期格式邏輯
        if "/" not in val_str:
            # 處理負號邏輯 (避免將 2023-01-01 誤判為負數)
            if "-" in val_str:
                if val_str.count("-") == 1 and val_str.startswith("-"):
                    float_val = float(val_str)
                    is_number = True
                else:
                    is_number = False
            else:
                float_val = float(val_str)
                is_number = True
    except ValueError:
        is_number = False

    if is_number:
        key_lower = str(key_name).strip().lower()

        if key_lower.startswith("me_"):
            if "." in val_str:
                parts = val_str.split(".")
                integer_part = parts[0]
                decimal_part = parts[1]
                formatted_int = "{:,}".format(int(integer_part))
                formatted_str = f"{formatted_int}.{decimal_part}"
            else:
                formatted_str = "{:,}".format(int(float_val))

        elif (
            key_lower.endswith("_rate")
            or "elec_price" in key_lower
            or "new_cop_std" in key_lower
            or "new_eff_std" in key_lower
        ):
            formatted_str = "{:,.2f}".format(float_val)

        elif key_lower.endswith("_year"):
            formatted_str = "{:,.1f}".format(float_val)

        else:
            formatted_str = "{:,.0f}".format(float_val)

        rt = RichText()
        rt.add(formatted_str, color="FF0000", bold=False)
        return rt

    return val_str


# ---------------- 主程式 ----------------
st.set_page_config(page_title="節能績效計劃書生成器", page_icon="📊")

st.title("📊 HWsmart節能績效計劃書生成器")
st.markdown("""
此工具支援 **Excel 表格同步** 功能：

1. **單一變數**（例如：COP、效率、kWh 等）標示為 **紅字**。
   - 請放在 Excel Sheet 的 第一個分頁中。  
   - 第 1 欄為「變數名稱」，第 2 欄為「數值」，其餘欄位會被忽略。  
   - 在 Word 中使用：`{{r 變數名稱}}`。

2. **表格資料（例如：改善前冰水機、改善前水泵…）** - 每個表格放在獨立的 Sheet，**Sheet 名稱 = Word 中的變數名稱**。
   - Word 表格內使用（搭配 docxtpl 的 row 擴充）：  
     開頭列：`{%tr for row in 改善前冰水機 %}`  
     中間：`{{ row.欄位名 }}`  
     結尾列：`{%tr endfor %}`

3. **RichText（紅字）** - 單一變數只要 Python 端處理成 RichText，Word 模板要寫成 `{{r 變數}}`。
""")

col1, col2 = st.columns(2)
with col1:
    uploaded_word = st.file_uploader("1️⃣ 上傳 Word 模板 (.docx)", type="docx")
with col2:
    uploaded_excel = st.file_uploader("2️⃣ 上傳 Excel 數據 (.xlsx)", type="xlsx")

if uploaded_word and uploaded_excel:
    st.divider()

    if st.button("🚀 開始生成報告", type="primary"):
        try:
            uploaded_word.seek(0)
            uploaded_excel.seek(0)

            word_bytes = uploaded_word.read()
            excel_bytes = uploaded_excel.read()

            excel_io = io.BytesIO(excel_bytes)
            excel_file = pd.ExcelFile(excel_io)

            context = {}
            st.toast("🔍 正在解析 Excel 資料...")

            for i, sheet_name in enumerate(excel_file.sheet_names):

                # 1) 變數 Sheet（第 1 張）：套用紅字格式化
                if i == 0:
                    df_var = excel_file.parse(sheet_name=sheet_name, header=None)
                    for _, row in df_var.iterrows():
                        if pd.isna(row[0]):
                            continue
                        key = str(row[0]).strip()
                        val = row[1]
                        context[key] = process_value_to_richtext(val, key_name=key)

                # 2) 表格 Sheet（其餘）：完全不更動值（只把 NaN 變成 ""）
                else:
                    df = excel_file.parse(sheet_name=sheet_name)

                    # ✅ 只刪除整列全空（不改任何 cell 值）
                    df = df.dropna(how="all")

                    # ✅ 欄位名 strip（不影響值）
                    df.columns = [str(c).strip() for c in df.columns]

                    table_list = []
                    for _, row in df.iterrows():
                        row_dict = {}
                        for col_name in df.columns:
                            v = row[col_name]
                            # ✅ 唯一處理：NaN → ""（避免 Word 顯示 nan）
                            row_dict[col_name] = "" if pd.isna(v) else v
                        table_list.append(row_dict)

                    context[sheet_name] = table_list

            # 渲染 Word
            doc = DocxTemplate(io.BytesIO(word_bytes))
            doc.render(context)

            output_buffer = io.BytesIO()
            doc.save(output_buffer)
            doc_bytes = output_buffer.getvalue()

            # 檔名邏輯
            download_name = "報告測試.docx"
            file_name_var = context.get("檔名", None)

            if isinstance(file_name_var, RichText):
                download_name = "Generated_Report.docx"
            elif isinstance(file_name_var, str) and file_name_var.strip():
                download_name = f"{file_name_var.strip()}.docx"

            st.session_state["generated_doc"] = doc_bytes
            st.session_state["download_name"] = download_name
            st.success("✅ 報告生成成功！請點擊下方按鈕下載。")

        except Exception as e:
            st.error(f"❌ 發生錯誤：{e}")

    if "generated_doc" in st.session_state:
        st.download_button(
            label="📥 下載生成的報告",
            data=st.session_state["generated_doc"],
            file_name=st.session_state["download_name"],
            mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
        )
