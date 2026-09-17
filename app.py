import streamlit as st
import pdfplumber
import pandas as pd

def process_pdf(uploaded_file):
    all_data = []
    headers = None

    with pdfplumber.open(uploaded_file) as pdf:
        for page in pdf.pages:
            table = page.extract_table()
            if table:
                table = [row for row in table if not all(cell is None for cell in row)]
                if not table:
                    continue

                if headers is None:
                    headers = [str(cell).replace('\n', '') if cell else '' for cell in table[0]]
                    all_data.extend(table[1:])
                else:
                    first_row = [str(cell).replace('\n', '') if cell else '' for cell in table[0]]
                    if first_row == headers:
                        all_data.extend(table[1:])
                    else:
                        all_data.extend(table)

    if not headers or not all_data:
        return None, "無法從 PDF 中解析出有效的表格格式。"

    df = pd.DataFrame(all_data, columns=headers)
    df = df.replace(r'\n', '', regex=True)
    df = df.fillna('')
    
    return df, None

def categorize_and_calculate(row):
    org_string = str(row.get('主辦單位', '')).replace(' ', '')
    course_string = str(row.get('課程名稱', '')).replace(' ', '')
    
    try:
        score = float(row.get('有效積分', 0))
    except:
        score = 0.0

    # ==== A類判斷邏輯 ====
    a_keywords = ["中華民國贗復牙科學會", "中華民國復牙科學會", "中華民國贗復牙會", "復牙科"]
    if any(kw in org_string for kw in a_keywords):
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
    
    is_b_class = any(kw in org_string for kw in b_keywords)
    score_multiplier = 1.0

    # 規則 4：中華牙醫學會年會
    if "中華牙醫學會年會" in org_string or "中華牙醫學會年會" in course_string:
        is_b_class = True
        score_multiplier = 1.0 / 3.0

    if is_b_class:
        return pd.Series(['B類', score * score_multiplier])

    # ==== 待判定邏輯 ====
    return pd.Series(['待判定', score])

def main():
    st.set_page_config(page_title="學分分析工具", layout="wide")
    st.title("🦷 學分自動分析計算工具")
    st.write("上傳「學分整理結果.pdf」，系統將自動為您區分 A類(贗復)、B類(指定單位/規則) 及待判定學分，並精算總和。")

    uploaded_file = st.file_uploader("選擇 PDF 檔案", type="pdf")

    if uploaded_file is not None:
        with st.spinner("正在解析 PDF 並套用分類規則..."):
            df, error = process_pdf(uploaded_file)
            
            if error:
                st.error(error)
            else:
                if '主辦單位' not in df.columns or '有效積分' not in df.columns:
                    st.error("解析失敗：找不到「主辦單位」或「有效積分」欄位。")
                    return

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