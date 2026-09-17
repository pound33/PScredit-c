import io
import re
import pandas as pd
import streamlit as st

def evaluate_credit(org_name: str, points: float):
    """判定主辦單位類別、折算率與採認積分"""
    org = str(org_name).strip()
    pts = float(points) if pd.notnull(points) else 0.0

    # 1. 判定 A 類（相容「贋」與「贗」）
    if re.search(r"中華民國[贋贗]復牙科學會", org):
        return "A類", 1.0, pts, "中華民國贋復牙科學會"

    # 2. 判定 B 類條件 4：中華牙醫學會年會（折算 1/3）
    if "中華牙醫學會年會" in org:
        b_pts = round(pts * (1 / 3), 2)
        return "B類", 1 / 3, b_pts, "中華牙醫學會年會 (折算 1/3)"

    # 3. 判定 B 類條件 1：醫學院
    if "醫學院" in org:
        return "B類", 1.0, pts, "醫學院體系"

    # 4. 判定 B 類條件 2：校友會 或 校友總會
    if "校友會" in org or "校友總會" in org:
        return "B類", 1.0, pts, "校友會/校友總會"

    # 5. 判定 B 類條件 3：指定醫學中心與醫院名單
    target_hospitals = [
        "長庚醫院", "台大醫院", "總醫院", "奇美醫院", "成大醫院",
        "童綜合醫院", "中國附醫", "北醫附醫", "馬偕醫院", "高雄長庚",
        "高醫附醫", "花蓮慈濟"
    ]
    for hosp in target_hospitals:
        if hosp in org:
            return "B類", 1.0, pts, f"採認醫院 ({hosp})"

    # 6. 判定 B 類條件 5：八大專科/學會
    target_societies = [
        "中華民國口腔顎面外科學會",
        "中華民國齒顎矯正學會",
        "中華民國家庭牙醫學會",
        "中華民國兒童牙醫學會",
        "台灣牙周病醫學會",
        "中華民國牙髓病學會",
        "台灣特殊需求者口腔醫學會牙體復形科",
        "中華民國牙體復形學會"
    ]
    for soc in target_societies:
        if soc in org:
            return "B類", 1.0, pts, f"採認學會 ({soc})"

    # 非 A、B 類
    return "其他", 0.0, 0.0, "非採認項目"


def process_credit_summary(df_org_summary: pd.DataFrame):
    """傳入主辦單位統計總表，進行 A/B 類標籤分類、積分統計與 UI/Excel 輸出"""
    # 確保欄位名稱標準化 (假設欄位為 '主辦單位' 與 '專業課程' 或 '積分')
    org_col = next((c for c in df_org_summary.columns if "主辦單位" in c), "主辦單位")
    pts_col = next((c for c in df_org_summary.columns if any(k in c for k in ["專業課程", "專業", "積分"])), "積分")

    # 進行每一筆資料的判定
    results = df_org_summary.apply(
        lambda row: evaluate_credit(row[org_col], row[pts_col]),
        axis=1,
        result_type="expand"
    )
    
    df_calc = df_org_summary.copy()
    df_calc["採認類別"] = results[0]
    df_calc["折算率"] = results[1]
    df_calc["採認積分"] = results[2]
    df_calc["採認依據"] = results[3]

    # 指標數值計算
    score_a = df_calc.loc[df_calc["採認類別"] == "A類", "採認積分"].sum()
    score_b = df_calc.loc[df_calc["採認類別"] == "B類", "採認積分"].sum()
    total_prof_score = df_calc[pts_col].sum()
    total_ab_score = score_a + score_b

    # 呈現 4 格自適應指標卡片 (Metric Cards)
    col1, col2, col3, col4 = st.columns(4)
    with col1:
        st.metric(label="A 類有效積分", value=f"{score_a:.2f} 分", help="中華民國贋復牙科學會有效積分")
    with col2:
        st.metric(label="B 類採認積分", value=f"{score_b:.2f} 分", help="符合五大採認條件且經折算後的有效積分")
    with col3:
        st.metric(label="專業課程總積分", value=f"{total_prof_score:.2f} 分", help="所有參加之專業課程原始累計總積分")
    with col4:
        st.metric(label="A + B 類合計", value=f"{total_ab_score:.2f} 分", help="供專科換照或評鑑對照之重點積分")

    # 產生 Excel 檔案供下載
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        # Sheet 1: 統計指標摘要
        summary_data = {
            "指標項目": ["A 類有效積分", "B 類採認積分", "專業課程總積分", "A + B 類合計"],
            "積分數值": [round(score_a, 2), round(score_b, 2), round(total_prof_score, 2), round(total_ab_score, 2)],
            "說明": [
                "中華民國贋復牙科學會有效積分",
                "符合五大條款（含年會 1/3 折算）",
                "原始所有專業學分總合",
                "專科換照審查重點積分"
            ]
        }
        pd.DataFrame(summary_data).to_excel(writer, sheet_name="積分彙整指標", index=False)
        
        # Sheet 2: 完整明細表
        df_calc.to_excel(writer, sheet_name="主辦單位採認明細", index=False)

    excel_bytes = output.getvalue()

    # 下載按鈕
    st.download_button(
        label="📥 下載 A/B 類積分統計 Excel 報表",
        data=excel_bytes,
        file_name="學分統計_AB類採認彙整表.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        use_container_width=True
    )

    return df_calc