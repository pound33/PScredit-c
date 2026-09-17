import io
import re
import openpyxl
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
import pandas as pd
import pdfplumber
import streamlit as st

# 頁面配置
st.set_page_config(
    page_title="鳥專科換證積分分析工具",
    page_icon="🦷",
    layout="wide",
)

st.title("鳥專科醫師繼續教育積分自動整理器 (學分整理專用版)")
st.caption("專門支援「學分整理結果」格式：過濾專業課程、辨識 A/B 類學分，並匯出雙頁籤 Excel 報表。")

# 1. 畫面首要條件：要求使用者先上傳檔案
uploaded_file = st.file_uploader(
    "請先上傳您的「學分整理結果.pdf」",
    type=["pdf"],
    help="請上傳完整積分明細 PDF 檔",
)

if uploaded_file is None:
    st.info("👈 請於上方上傳您的 PDF 紀錄檔案以啟動統計分析。")
    st.stop()


def extract_records_from_pdf(file_bytes) -> list[dict]:
    records = []
    full_text = ""
    
    with pdfplumber.open(file_bytes) as pdf:
        for page in pdf.pages:
            t = page.extract_text()
            if t: 
                full_text += t + "\n"
                
            # 引擎 1：結構化表格擷取 (精準應對「學分整理結果」格式)
            tables = page.extract_tables()
            for table in tables:
                if not table: continue
                header_idx = -1
                for i, row in enumerate(table):
                    row_str = "".join([str(c) for c in row if c])
                    if "課程類別" in row_str and "有效積分" in row_str:
                        header_idx = i
                        break
                
                if header_idx != -1:
                    for row in table[header_idx+1:]:
                        if len(row) < 5: continue
                        cat = str(row[0]).strip() if row[0] else ""
                        
                        # 嚴格過濾：只抓取專業課程
                        if "專業課程" not in cat: continue
                        
                        try: pts = float(str(row[1]).strip())
                        except: pts = 0.0
                        
                        review = str(row[2]).strip() if len(row) > 2 and row[2] else ""
                        host = str(row[3]).strip() if len(row) > 3 and row[3] else ""
                        course_name = str(row[4]).strip() if len(row) > 4 and row[4] else ""
                        date_str = str(row[5]).strip() if len(row) > 5 and row[5] else ""
                        
                        records.append({
                            "有效積分": pts,
                            "無效積分": 0.0,
                            "審查單位": review,
                            "主辦單位": host,
                            "課程名稱": course_name,
                            "raw_lines": [date_str],
                            "is_new_format": True
                        })
    
    # 若表格引擎成功擷取資料，直接進入後處理
    if records:
        return [finalize_record(r) for r in records]

    # 引擎 2：舊版純文字擷取 (若不是新版表格，退回備用機制)
    start_match = re.search(r"◎\s*參加課程積分", full_text)
    if start_match:
        sub_text = full_text[start_match.start() :]
        end_matches = list(re.finditer(r"\n\s*◎", sub_text))
        if len(end_matches) > 1:
            sub_text = sub_text[: end_matches[1].start()]

        lines = [line.strip() for line in sub_text.split("\n") if line.strip()]
        current_record = None

        for line in lines:
            if any(k in line for k in ["有效總積分", "課程類別", "審查單位", "衛生福利部", "人員類別"]):
                continue

            m1 = re.match(r"^專業課程\s+([0-9.]+)\s+([0-9.]+)\s+(\S+)\s+(\S+)\s*(.*)$", line)
            m2 = re.match(r"^([0-9.]+)\s+([0-9.]+)\s+專業課程\s+(\S+)\s+(\S+)\s*(.*)$", line)

            if m1 or m2:
                if current_record: records.append(current_record)
                m = m1 if m1 else m2
                v_pts, inv_pts, review, host, rest = m.groups()
                current_record = {
                    "有效積分": float(v_pts),
                    "無效積分": float(inv_pts),
                    "審查單位": review,
                    "主辦單位": host,
                    "課程名稱": rest,
                    "raw_lines": [],
                    "is_new_format": False
                }
            else:
                if any(line.startswith(c) for c in ["專業品質", "專業相關法規", "專業倫理"]):
                    if current_record:
                        records.append(current_record)
                        current_record = None
                    continue
                if current_record:
                    current_record["raw_lines"].append(line)

        if current_record:
            records.append(current_record)

    return [finalize_record(r) for r in records]


def finalize_record(record: dict) -> dict:
    is_new = record.pop("is_new_format", False)
    raw_tail = " ".join(record.pop("raw_lines"))

    # 提取課程日期與年份
    dates = re.findall(r"(\d{4}/\d{1,2}/\d{1,2})", raw_tail)
    course_date = dates[0] if dates else "未知"
    year = course_date.split("/")[0] if dates else "未知"

    host = record.get("主辦單位", "")
    review = record.get("審查單位", "")
    course_name = record.get("課程名稱", "")
    original_pts = record.get("有效積分", 0.0)

    # 智慧校正：如果表格擷取時將「主辦單位」與「課程名稱」顛倒，自動還原
    if is_new:
        if any(course_name.endswith(s) for s in ["學會", "公會", "醫院", "大學", "中心", "聯盟"]) and \
           not any(host.endswith(s) for s in ["學會", "公會", "醫院", "大學", "中心", "聯盟"]):
            host, course_name = course_name, host

    # 徹底清除所有空白與不可見換行，避免 PDF 排版干擾比對
    host_clean = re.sub(r'\s+', '', host)
    review_clean = re.sub(r'\s+', '', review)
    course_clean = re.sub(r'\s+', '', course_name)
    
    course_category = "待判定學分" 
    final_pts = original_pts

    # 【A 類無敵字根】：加入容錯特徵，專門捕捉 贗/贋/膺 異體字與 PDF 漏字狀況
    host_review = f"{host_clean}_{review_clean}"
    a_keywords = ["贗復", "贋復", "膺復", "復牙科"]
    is_a_class = any(kw in host_review for kw in a_keywords)

    if is_a_class:
        course_category = "A"
        # 【強制歸戶】：不管原文是贗復、贋復、還是漏字，一律強制正名為統一格式
        host = "中華民國贋復牙科學會"
    else:
        # 【B 類規則】：正面表述驗證
        is_b_class = False
        
        # 規則 4（中華牙醫學會年會）- 需查驗課程名稱與主辦單位
        if "中華牙醫學會年會" in course_clean or "中華牙醫學會年會" in host_review:
            is_b_class = True
            final_pts = original_pts / 3.0
            
        # 若未命中規則 4，則檢驗其他主辦單位關鍵字
        if not is_b_class:
            b_keywords = [
                "醫學院", "醫學大學", "校友會", "校友總會", "牙友學會",
                "長庚", "台大", "總醫院", "奇美", "成大", 
                "童綜合", "中國附醫", "北醫", "馬偕", "高醫", "慈濟",
                "口腔顎面外科", "齒顎矯正", "家庭牙醫", "兒童牙醫",
                "牙周病", "牙髓病", "特殊需求", "牙體復形"
            ]
            
            # B 類比對：只查主辦單位與審查單位，避免將課程標題中的醫院誤判
            if any(kw in host_review for kw in b_keywords):
                is_b_class = True
                
        if is_b_class:
            course_category = "B"

    # 清洗課程名稱，保留單一空白以利閱讀
    if not is_new:
        cleaned_tail = re.sub(r"\d{4}/\d{1,2}/\d{1,2}(?:\s+\d{1,2}:\d{2})?", "", raw_tail)
        cleaned_tail = re.sub(r"\b(D1|D2|AG|A|B|C|F|G|H)\b", "", cleaned_tail)
        cleaned_tail = re.sub(r"[AB]\s*類", "", cleaned_tail)
        full_title = re.sub(r"\s+", " ", (course_name + " " + cleaned_tail).strip())
    else:
        full_title = re.sub(r"\s+", " ", course_name.strip())
        
    if not full_title and host:
        full_title = "專業課程"

    return {
        "主辦單位": host.strip(),
        "課程名稱": full_title,
        "類別": course_category,
        "有效積分": final_pts,
        "課程日期": course_date,
        "年度": year,
    }


def build_excel(records: list[dict], stream: io.BytesIO):
    org_dict = {}
    for r in records:
        org_dict.setdefault(r["主辦單位"], []).append(r)
    sorted_orgs = sorted(
        org_dict.items(),
        key=lambda x: sum(i["有效積分"] for i in x[1]),
        reverse=True,
    )

    wb = openpyxl.Workbook()
    thin_border = Border(
        left=Side(style="thin", color="D3D3D3"),
        right=Side(style="thin", color="D3D3D3"),
        top=Side(style="thin", color="D3D3D3"),
        bottom=Side(style="thin", color="D3D3D3"),
    )
    double_bottom = Border(
        left=Side(style="thin", color="D3D3D3"),
        right=Side(style="thin", color="D3D3D3"),
        top=Side(style="thin", color="D3D3D3"),
        bottom=Side(style="double", color="1B365D"),
    )

    # 頁籤 1: 主辦單位統計總表
    ws1 = wb.active
    ws1.title = "主辦單位統計總表"
    ws1.views.sheetView[0].showGridLines = True
    ws1["A1"] = "牙醫師繼續教育積分 — 專業課程統計總表 (含待判定學分)"
    ws1["A1"].font = Font(name="微軟正黑體", size=14, bold=True, color="1B365D")
    ws1.append([])

    headers_ws1 = [
        "主辦單位名稱",
        "修習堂數",
        "A類積分",
        "B類積分",
        "待判定積分",
        "有效積分加總",
        "積分佔比 (%)",
        "課程年份區間",
    ]
    ws1.append(headers_ws1)

    for col in range(1, 9):
        c = ws1.cell(row=3, column=col)
        c.font = Font(name="微軟正黑體", size=10, bold=True, color="FFFFFF")
        c.fill = PatternFill("solid", fgColor="1B365D")
        c.alignment = Alignment(horizontal="center", vertical="center")

    row = 4
    for org, items in sorted_orgs:
        pts = sum(i["有效積分"] for i in items)
        a_pts = sum(i["有效積分"] for i in items if i["類別"] == "A")
        b_pts = sum(i["有效積分"] for i in items if i["類別"] == "B")
        p_pts = sum(i["有效積分"] for i in items if i["類別"] == "待判定學分")
        years = sorted(list(set(i["年度"] for i in items)))
        yr_str = f"{years[0]} ~ {years[-1]}" if len(years) > 1 else years[0]

        ws1.cell(row=row, column=1, value=org).font = Font(name="微軟正黑體", size=10, bold=True)
        ws1.cell(row=row, column=2, value=len(items)).alignment = Alignment(horizontal="center")
        ws1.cell(row=row, column=3, value=round(a_pts, 2)).alignment = Alignment(horizontal="right")
        ws1.cell(row=row, column=4, value=round(b_pts, 2)).alignment = Alignment(horizontal="right")
        ws1.cell(row=row, column=5, value=round(p_pts, 2)).alignment = Alignment(horizontal="right")
        ws1.cell(row=row, column=6, value=round(pts, 2)).alignment = Alignment(horizontal="right")
        ws1.cell(row=row, column=7, value=f"=F{row}/F{len(sorted_orgs)+4}")
        ws1.cell(row=row, column=8, value=yr_str).alignment = Alignment(horizontal="center")

        for c_idx in [3, 4, 5, 6]:
            ws1.cell(row=row, column=c_idx).number_format = "#,##0.00"
        ws1.cell(row=row, column=7).number_format = "0.0%"
        for c in range(1, 9):
            ws1.cell(row=row, column=c).border = thin_border
        row += 1

    # 總計行
    ws1.cell(row=row, column=1, value="總計").font = Font(name="微軟正黑體", size=10, bold=True, color="1B365D")
    ws1.cell(row=row, column=2, value=f"=SUM(B4:B{row-1})").alignment = Alignment(horizontal="center")
    ws1.cell(row=row, column=3, value=f"=SUM(C4:C{row-1})").alignment = Alignment(horizontal="right")
    ws1.cell(row=row, column=4, value=f"=SUM(D4:D{row-1})").alignment = Alignment(horizontal="right")
    ws1.cell(row=row, column=5, value=f"=SUM(E4:E{row-1})").alignment = Alignment(horizontal="right")
    ws1.cell(row=row, column=6, value=f"=SUM(F4:F{row-1})").alignment = Alignment(horizontal="right")
    ws1.cell(row=row, column=7, value=1).alignment = Alignment(horizontal="right")

    for c_idx in [3, 4, 5, 6]:
        ws1.cell(row=row, column=c_idx).number_format = "#,##0.00"
    ws1.cell(row=row, column=7).number_format = "0.0%"
    for c in range(1, 9):
        ws1.cell(row=row, column=c).border = double_bottom

    ws1.column_dimensions["A"].width = 28
    ws1.column_dimensions["B"].width = 12
    ws1.column_dimensions["C"].width = 14
    ws1.column_dimensions["D"].width = 14
    ws1.column_dimensions["E"].width = 14
    ws1.column_dimensions["F"].width = 16
    ws1.column_dimensions["G"].width = 14
    ws1.column_dimensions["H"].width = 18

    # 頁籤 2: 課程明細
    ws2 = wb.create_sheet(title="專業課程明細(依主辦單位分類)")
    ws2.views.sheetView[0].showGridLines = True
    curr_row = 1

    for idx, (org, items) in enumerate(sorted_orgs, 1):
        org_pts = sum(i["有效積分"] for i in items)
        ws2.merge_cells(start_row=curr_row, start_column=1, end_row=curr_row, end_column=6)
        banner = ws2.cell(row=curr_row, column=1, value=f"【大類 {idx}】 {org} （共 {len(items)} 堂，小計：{org_pts:.2f} 分）")
        banner.font = Font(name="微軟正黑體", size=11, bold=True, color="FFFFFF")
        fill_b = PatternFill("solid", fgColor="1B365D") if idx % 2 == 1 else PatternFill("solid", fgColor="006666")
        for c in range(1, 7):
            ws2.cell(row=curr_row, column=c).fill = fill_b
        curr_row += 1

        detail_headers = ["項次", "課程日期", "課程名稱 / 主題", "類別", "主辦單位", "有效積分"]
        for c_idx, h in enumerate(detail_headers, 1):
            cell = ws2.cell(row=curr_row, column=c_idx, value=h)
            cell.fill = PatternFill("solid", fgColor="2C4D75")
            cell.font = Font(name="微軟正黑體", size=10, bold=True, color="FFFFFF")
            cell.alignment = Alignment(horizontal="center")
        curr_row += 1

        start_r = curr_row
        for item_idx, item in enumerate(items, 1):
            ws2.cell(row=curr_row, column=1, value=item_idx).alignment = Alignment(horizontal="center")
            ws2.cell(row=curr_row, column=2, value=item["課程日期"]).alignment = Alignment(horizontal="center")
            ws2.cell(row=curr_row, column=3, value=item["課程名稱"]).alignment = Alignment(wrap_text=True)
            ws2.cell(row=curr_row, column=4, value=item["類別"]).alignment = Alignment(horizontal="center")
            ws2.cell(row=curr_row, column=5, value=org)
            ws2.cell(row=curr_row, column=6, value=item["有效積分"]).number_format = "#,##0.00"
            for c in range(1, 7):
                ws2.cell(row=curr_row, column=c).border = thin_border
            curr_row += 1

        # 小計行
        ws2.merge_cells(start_row=curr_row, start_column=1, end_row=curr_row, end_column=5)
        ws2.cell(row=curr_row, column=1, value=f"{org} — 小計").alignment = Alignment(horizontal="right")
        sub_sum = ws2.cell(row=curr_row, column=6, value=f"=SUM(F{start_r}:F{curr_row-1})")
        sub_sum.number_format = "#,##0.00"
        sub_sum.font = Font(name="微軟正黑體", bold=True)
        for c in range(1, 7):
            ws2.cell(row=curr_row, column=c).border = thin_border
        curr_row += 2

    ws2.column_dimensions["A"].width = 8
    ws2.column_dimensions["B"].width = 14
    ws2.column_dimensions["C"].width = 56
    ws2.column_dimensions["D"].width = 12
    ws2.column_dimensions["E"].width = 24
    ws2.column_dimensions["F"].width = 14

    wb.save(stream)


# 2. 資料解析與指標呈現
with st.spinner("正在解析 PDF 並計算 A/B 類積分..."):
    records = extract_records_from_pdf(uploaded_file)

if not records:
    st.error("未能從上傳的 PDF 中解析出任何「專業課程」資料，請確認檔案格式是否正確。")
    st.stop()

df = pd.DataFrame(records)

# 聚合統計
total_pts = df["有效積分"].sum()
a_pts = df[df["類別"] == "A"]["有效積分"].sum()
b_pts = df[df["類別"] == "B"]["有效積分"].sum()
p_pts = df[df["類別"] == "待判定學分"]["有效積分"].sum()
total_courses = len(df)
total_hosts = df["主辦單位"].nunique()

# 3. 五格自適應指標卡片 (Metric Cards)
st.subheader("📊 積分採認總覽")
c1, c2, c3, c4, c5 = st.columns(5)
with c1:
    st.metric(
        label="專業課程總積分",
        value=f"{total_pts:.2f} 分",
        help="所有專業課程採認有效積分加總",
    )
with c2:
    a_pct = (a_pts / total_pts * 100) if total_pts > 0 else 0
    st.metric(
        label="A 類積分",
        value=f"{a_pts:.2f} 分",
        delta=f"佔比 {a_pct:.1f}%",
        help="嚴格符合贗復學會",
    )
with c3:
    b_pct = (b_pts / total_pts * 100) if total_pts > 0 else 0
    st.metric(
        label="B 類積分",
        value=f"{b_pts:.2f} 分",
        delta=f"佔比 {b_pct:.1f}%",
        delta_color="off",
        help="正面表列之特定主辦機構",
    )
with c4:
    p_pct = (p_pts / total_pts * 100) if total_pts > 0 else 0
    st.metric(
        label="待判定學分",
        value=f"{p_pts:.2f} 分",
        delta=f"佔比 {p_pct:.1f}%",
        delta_color="inverse",
        help="未符合A或B類條件，需人工檢閱",
    )
with c5:
    st.metric(
        label="修習堂數 / 機構",
        value=f"{total_courses} 堂",
        delta=f"涵蓋 {total_hosts} 個主辦單位",
        delta_color="off",
    )

st.divider()

# 4. 主辦單位分組聚合表與詳細清單
tab1, tab2 = st.tabs(["🏛️ 主辦單位聚合統計", "📄 完整課程明細"])

with tab1:
    summary_df = (
        df.groupby("主辦單位")
        .agg(
            修習堂數=("課程名稱", "count"),
            A類積分=("有效積分", lambda s: s[df.loc[s.index, "類別"] == "A"].sum()),
            B類積分=("有效積分", lambda s: s[df.loc[s.index, "類別"] == "B"].sum()),
            待判定積分=("有效積分", lambda s: s[df.loc[s.index, "類別"] == "待判定學分"].sum()),
            有效積分總計=("有效積分", "sum"),
        )
        .sort_values(by="有效積分總計", ascending=False)
        .reset_index()
    )
    summary_df["積分佔比"] = (
        summary_df["有效積分總計"] / total_pts
    ).map(lambda x: f"{x:.1%}")
    st.dataframe(summary_df, use_container_width=True, hide_index=True)

with tab2:
    st.dataframe(
        df[["課程日期", "主辦單位", "課程名稱", "類別", "有效積分"]],
        use_container_width=True,
        hide_index=True,
    )

# 5. 匯出 Excel 報表下載按鈕
excel_stream = io.BytesIO()
build_excel(records, excel_stream)
excel_stream.seek(0)

st.download_button(
    label="📥 下載完整統計 Excel 報表 (.xlsx)",
    data=excel_stream,
    file_name="牙醫師繼續教育積分_學分整理專用版.xlsx",
    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
)