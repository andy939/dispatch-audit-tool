# 快速檢視本資料夾內派工案件 xlsx 匯出檔的統計分布，並(重新)產生 data.js 內的捷運站資料。
# 案件資料本身改為直接在網頁（index.html）用「匯入本月 xlsx 案件檔」上傳，不需要透過本程式匯入。
import json
import glob
import openpyxl


def parse_dt(s):
    if not s:
        return ""
    return str(s).replace("\n", " ").strip()


def clean(v):
    if v is None:
        return ""
    return str(v).strip()


def load_xlsx(fname):
    wb = openpyxl.load_workbook(fname, data_only=True)
    ws = wb.worksheets[0]
    rows = list(ws.iter_rows(min_row=1, values_only=True))
    # 系統原始匯出檔前幾列通常是報表標題/列印資訊，掃描找出含有「案件編號」的那一列當標題列
    header_idx = next((i for i, r in enumerate(rows) if r and "案件編號" in [clean(c) for c in r]), None)
    if header_idx is None:
        print(f"  [略過] {fname}：找不到「案件編號」欄位標題")
        return []
    header = rows[header_idx]
    cases = []
    for r in rows[header_idx + 1:]:
        if not r or not r[0]:
            continue
        d = dict(zip(header, r))
        case_id_raw = clean(d.get("案件編號"))
        if not case_id_raw:
            continue
        address2 = clean(d.get("發生地址2"))
        address1 = clean(d.get("發生地址"))
        cases.append({
            "id": case_id_raw.split("\n")[0],
            "idRaw": case_id_raw,
            "status": clean(d.get("案件狀態")),
            "category": clean(d.get("案件項目")),
            "agency": clean(d.get("承辦機關")),
            "district": clean(d.get("行政區")),
            "address": address2 or address1,
            "createdAt": parse_dt(d.get("成案時間")),
            "dueAt": parse_dt(d.get("限辦時間")),
            "closedAt": parse_dt(d.get("結案時間")),
            "content": clean(d.get("案件內容")),
            "handling": clean(d.get("案件處理")),
            "level": clean(d.get("案件等級")),
            "registrar": clean(d.get("立案人員")),
            "satisfaction": clean(d.get("滿意度選項")),
            "feedback": clean(d.get("民眾意見")),
            "sourceFile": fname,
        })
    return cases


def main():
    all_cases = []
    seen = set()
    for fname in glob.glob("*.xlsx"):
        for c in load_xlsx(fname):
            if c["id"] in seen:
                continue
            seen.add(c["id"])
            all_cases.append(c)

    print(f"總案件數: {len(all_cases)}")
    from collections import Counter
    print("案件項目分布:", Counter(c["category"] for c in all_cases).most_common())
    print("行政區分布:", Counter(c["district"] for c in all_cases).most_common())

    try:
        mrt = json.load(open("mrt_stations.json", encoding="utf-8"))
    except FileNotFoundError:
        mrt = []

    with open("data.js", "w", encoding="utf-8") as f:
        f.write("// 自動產生檔案，請勿手動編輯：地圖視窗 (map.html) 用的捷運站座標。\n")
        f.write("window.MRT_STATIONS = ")
        json.dump(mrt, f, ensure_ascii=False)
        f.write(";\n")
    print("已寫入 data.js")


if __name__ == "__main__":
    main()
