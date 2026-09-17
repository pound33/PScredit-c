import streamlit as st
import pdfplumber
import pandas as pd
import re

def process_pdf(uploaded_file):
    all_data = []
    headers = None

    with pdfplumber.open(uploaded_file) as pdf:
        for page in pdf.pages:
            table = page.extract_table()
            if table:
                # 移除全為空的列
                table = [row for row in table if not all(cell is None for cell in row)]
                if not table:
                    continue

                if headers is None:
                    # 第一頁，設定表頭並強制清除所有空白與換行，確保欄位名稱正確
                    headers = [re.sub(r'\s+', '', str(cell)) if cell else '' for cell in table[0]]
                    all_data.extend(table[1:])
                else:
                    first_row = [re.sub(r'\s+', '', str(cell)) if cell else '' for cell in table[0]]
                    if first_row == headers:
                        all_data.extend(table[1:])
                    else:
                        all_data.extend(table)

    if not headers or not all_data:
        return None, "無法從 PDF 中解析出有效的表格格式。"

    # 統一每一列的長度，避免因為 PDF 結尾總計表欄位太少導致後續處理錯位
    expected_cols = len(headers)
    normalized_data = []
    for row in all_data:
        # 將 None 轉為空字串
        row = [str(cell) if cell is not None else '' for cell in row]
        if len(row) < expected_cols:
            row = row + [''] * (expected_cols - len(row))
        elif len(row) > expected_cols:
            row = row[:expected_cols]
        normalized_data.append(row)

    df = pd.DataFrame(normalized_data, columns=headers)
    df = df.fillna('')
    
    # 【關鍵修正】：剃除 PDF 結尾的「依課程類別彙整」總計表
    # 真實的課程一定會有「課程名稱」，而總計表跑到這個欄位時會是空白的
    if '課程名稱' in df.columns:
        df = df[df['課程名稱'].astype(str).str.strip() != '']
    
    return df, None

def categorize_and_calculate(row):
    org_val = str(row.get('主辦單位', ''))
    course_val = str(row.get('課程名稱', ''))
    reviewer_val = str(row.get('審查單位', ''))
    
    # 使用正規表達式去除所有(全形/半形)空白與換行符號
    org_string = re.sub(r'\s+', '', org_val)
    course_string = re.sub(r'\s+', '', course_val)
    reviewer_string = re.sub(r'\s+', '', reviewer_val)
    
    # 組合字串，徹底解決 PDF 欄位內容左右溢出的錯位問題
    combined_string = org_string + course_string + reviewer_string
    
    try:
        score = float(row.get('有效積分', 0))
    except:
        score = 0.0

    # ==== A類判斷邏輯 ====
    a_keywords = ["贋", "贗", "復牙科"]
    if any(kw in combined_string for kw in a_keywords):
        return pd.Series(['A類', score])

    # ==== B類判斷邏輯 ====
    b_keywords = [
        "醫學院", "醫學大學",
        "校友會", "校友總會", "牙友學會",
        "長庚醫院", "台大醫院", "總醫院", "奇美醫院", "成大醫院",
        "童綜合醫院", "中國附醫", "北醫附醫", "馬偕醫院", "高雄長庚",
        "高醫附醫", "花蓮慈濟",
        "中華民國口腔顎面外科學會", "中華民國齒顎矯正學會", "中華民國家庭牙醫學會",
        "中華民國兒童牙醫學會", "台灣牙周病醫學會", "中華民國牙髓病學會",
        "台灣特殊需求者口腔醫學會", "牙體復形科", "中華民國牙體復形學會"
    ]
    
    is_b_class = any(kw in combined_string for kw in b_keywords)
    score_multiplier = 1.0

    # 規則 4：中華牙醫學會年會
    if "中華牙醫學會年會" in combined_string:
        is_b_class = True
        score_multiplier = 1.0 / 3.0

    if is_b_class:
        return pd.Series(['B類', score * score_multiplier])

    # ==== 待判定邏輯 ====
    return pd.Series(['待判定', score])

def main():
    st.set_page_config(page_title="學分分析工具", layout="wide")
    st.title("鳥專科醫師換証A/B類學分自動分類計算")
    st.write("請上傳轉檔後的「學分整理結果.pdf」，系統將自動為您區分 A類(贋復)、B類(符合目前正面表述主辦單位) 及待判定學分，並精算總和。")

    uploaded_file = st.file_uploader("選擇 PDF 檔案", type="pdf")

    if uploaded_file is not None:
        with st.spinner("正在解析 PDF 並套用分類規則..."):
            df, error = process_pdf(uploaded_file)
            
            if error:
                st.error(error)
            else:
                if '主辦單位' not in df.columns or '有效積分' not in df.columns:
                    st.error("解析失敗：找不到「主辦單位」或「有效積分」欄位。")
                    st.write("目前抓取到的欄位為：", df.columns.tolist())
                    return

                # 執行分類運算
                df[['分類', '核算後積分']] = df.apply(categorize_and_calculate, axis=1)

                df_a = df[df['分類'] == 'A類']
                df_b = df[df['分類'] == 'B類']
                df_pending = df[df['分類'] == '待判定']

                total_a = df_a['核算後積分'].sum()
                total_b = df_b['核算後積分'].sum()
                total_pending = df_pending['核算後積分'].sum()

                st.success("檔案解析與分類成功！")
                
                col1, col2, col3 = st.columns(3)
                with col1:
                    st.metric(label="✅ A類學分 (贗復) 總計", value=f"{total_a:.2f} 分")
                with col2:
                    st.metric(label="✅ B類學分 (其他認可) 總計", value=f"{total_b:.2f} 分")
                with col3:
                    st.metric(label="⚠️ 待判定學分 總計", value=f"{total_pending:.2f} 分")
                
                st.divider()

                tab1, tab2, tab3, tab4 = st.tabs(["A類清單", "B類清單", "待判定清單", "查看原始所有資料"])
                
                with tab1:
                    st.subheader(f"A類 課程清單 (共 {len(df_a)} 筆)")
                    st.dataframe(df_a, use_container_width=True, hide_index=True)
                
                with tab2:
                    st.subheader(f"B類 課程清單 (共 {len(df_b)} 筆)")
                    st.info("💡 提示：若課程包含「中華牙醫學會年會」，核算後積分已自動乘以 1/3。")
                    st.dataframe(df_b, use_container_width=True, hide_index=True)
                
                with tab3:
                    st.subheader(f"待判定 課程清單 (共 {len(df_pending)} 筆)")
                    st.warning("這些項目不符合 A 類與 B 類的正面表述字詞，請人工檢查是否為有效學分。")
                    st.dataframe(df_pending, use_container_width=True, hide_index=True)
                    
                with tab4:
                    st.subheader(f"原始所有資料 (共 {len(df)} 筆)")
                    st.dataframe(df, use_container_width=True, hide_index=True)

if __name__ == "__main__":
    main()