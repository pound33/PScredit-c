import re

def classify_course(host_name, title_name, points):
    """
    依主辦單位與課程名稱，判斷歸屬分類 (A類 / B類 / 其他) 並計算折抵後積分
    """
    host = str(host_name).strip()
    title = str(title_name).strip()
    pts = float(points)

    # ------------------ 1. A 類判斷 ------------------
    if "贗復" in host or "贋復" in host:
        return {
            "分類": "A類",
            "命中原因": "主辦單位為贗復牙科學會",
            "採認積分": pts,
            "折算備註": "全額採認"
        }

    # ------------------ 2. B 類判斷 ------------------
    b_reasons = []
    factor = 1.0

    # 規則 1: 醫學院/大學
    if any(k in host for k in ["醫學院", "醫學大學", "醫學院牙醫學系"]):
        b_reasons.append("符合[醫學院/大學]")

    # 規則 2: 校友會 / 牙友會
    if any(k in host for k in ["校友會", "校友總會", "牙友會", "牙友學會"]):
        b_reasons.append("符合[校友會/牙友會]")

    # 規則 3: 指定教學醫院
    hospitals = [
        "長庚醫院", "長庚紀念醫院", "台大醫院", "總醫院", "奇美醫院",
        "成大醫院", "童綜合醫院", "中國附醫", "北醫附醫", "馬偕醫院",
        "高雄長庚", "高醫附醫", "花蓮慈濟"
    ]
    if any(k in host for k in hospitals):
        b_reasons.append("符合[指定醫院]")

    # 規則 4: 中華牙醫學會年會 (積分 1/3)
    if "中華牙醫學會年會" in host or "中華牙醫學會年會" in title:
        b_reasons.append("中華牙醫學會年會(1/3折算)")
        factor = 1.0 / 3.0

    # 規則 5: 部定專科學會
    societies = [
        "中華民國口腔顎面外科學會",
        "中華民國齒顎矯正學會",
        "中華民國家庭牙醫學會",
        "中華民國兒童牙醫學會",
        "台灣牙周病醫學會",
        "中華民國牙髓病學會",
        "台灣特殊需求者口腔醫學會牙體復形科",
        "中華民國牙體復形學會"
    ]
    if any(k in host for k in societies):
        b_reasons.append("符合[部定專科學會]")

    if b_reasons:
        return {
            "分類": "B類",
            "命中原因": "、".join(b_reasons),
            "採認積分": round(pts * factor, 2),
            "折算備註": "1/3折算" if factor < 1.0 else "全額採認"
        }

    # ------------------ 3. 其他 ------------------
    return {
        "分類": "其他",
        "命中原因": "未符合A/B類條件",
        "採認積分": 0.0,
        "折算備註": "不計入A/B"
    }