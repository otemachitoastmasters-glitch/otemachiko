import requests
from playwright.sync_api import sync_playwright
from bs4 import BeautifulSoup
from openpyxl import load_workbook
from openpyxl.utils import get_column_letter
from openpyxl.drawing.image import Image
import datetime

def fetch_latest_mtgid(base_url="https://tmcsupport.coresv.com/otemachiko/"):
    headers = {
        "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/115.0 Safari/537.36"
    }
    response = requests.get("https://tmcsupport.coresv.com/otemachiko/", headers=headers)
    
    response = requests.get(base_url)
    soup = BeautifulSoup(response.text, "html.parser")

    # 全てのtrを取得（1行目はthなのでスキップ）
    rows = soup.select("table.tableCommon tr")[1:]

    latest_mtgid = None
    latest_date = None

    for row in rows:
        cols = row.find_all("td")
        if len(cols) < 2:
            continue

        # 日付と onclick 属性取得
        date_text = cols[0].text.strip()
        link_tag = cols[1].find("a")
        if not link_tag or "onclick" not in link_tag.attrs:
            continue

        onclick = link_tag["onclick"]
        if "showDetail" not in onclick:
            continue

        try:
            mtg_id = int(onclick.split("showDetail(")[1].split(")")[0])
            meeting_date = datetime.strptime(date_text, "%Y/%m/%d")
        except Exception as e:
            continue

        # 最も未来の会議日程（今日以降）を探す
        if meeting_date >= datetime.today():
            if latest_date is None or meeting_date < latest_date:
                latest_date = meeting_date
                latest_mtgid = mtg_id

    //if latest_mtgid == None:
    //    latest_mtgid = 77
        
    return latest_mtgid
    
def generate_agenda_excel_from_url(mtgid, template_path="meeting_agenda_template.xlsx", output_path: str = "generated_agenda.xlsx") -> str:

    html_url = f"https://tmcsupport.coresv.com/otemachiko/mtgDetailReadonly.php?mtgid={mtgid}"
    print(f"🔗 Fetching agenda from: {html_url}")
    res = requests.get(html_url)
    soup = BeautifulSoup(res.content, "html.parser")

    # 会議情報取得
    header_table = soup.find("table", class_="tableCommon")
    rows = header_table.find_all("tr")
    mtg_info = rows[1].find_all("td")
    date = mtg_info[0].text.strip()
    title = mtg_info[1].text
    venue = mtg_info[3].text.strip()
    room = mtg_info[4].text.strip()
    
    meeting_title = title
    meeting_datetime = date
    
    # ゲスト取得
    guest = ""
    for table in soup.find_all("table", class_="tableCommon"):
        th = table.find("th")
        if th and "Guests" in th.text:
            guest_td = table.find("td")
            guest = guest_td.get_text(strip=True)
            break
    
    # アジェンダ表取得
    agenda_table = soup.find("table", class_="tableCommon mainTbl")
    agenda = []
    role_name_map = {}
    evaluator_map = {}
    speech_path_map = {}
    speech_title_map = {}
    theme = ""
    for tr in agenda_table.find_all("tr")[1:]:
        tds = tr.find_all("td")
        if len(tds) >= 3:
            role = tds[0].text.strip()
            name = tds[1].text.strip()
            detail = tds[2].text.strip()
            title = tds[3].text.strip() if len(tds) > 3 else ""
            agenda.append([role, name, detail, title])
            role_name_map[role] = name
            if "Theme" in role:
                theme = detail
            if "Speech" in role:
                speech_path_map[role] = detail.split("\n")[-2] if "\n" in detail else ""
                speech_title_map[role] = title
    
            if "Evaluator" in role:
                evaluator_map[role] = detail
            
    # Excelテンプレートを読み込み
    wb = load_workbook("meeting_agenda_template.xlsx")
    ws = wb.active
    ws.title = "Agenda"

    img = Image('toastmasters_logo.jpg')
    # サイズの調整（ピクセル単位）
    img.width = 150  # 幅(px)
    img.height = 105  # 高さ(px)

    # 貼り付け位置（例：A1セル）
    ws.add_image(img, 'B1')  # A1セルに貼り付け（位置は調整してください）
    
    # 💡 すべての結合セルを解除
    if ws.merged_cells.ranges:
        print("⚠️ Unmerging cells...")
    for merged_range in list(ws.merged_cells.ranges):
        ws.unmerge_cells(str(merged_range))
    
    # 🔄 タイトル・日時を書き込む（例：A2セル想定）
    ws["F3"] = theme
    ws["K3"] = f"{meeting_title}　{meeting_datetime}"
    ws["K4"] = f"{venue} {room}"
    
    # TMOE, WOE, Ah-Counter, Grammarian, PC manager
    ws["I9"] = f"{role_name_map['Toastmaster of the Evening']}"
    ws["I10"] = f"{role_name_map['Word of the Evening']}"
    ws["I11"] = f"{role_name_map['Ah-Counter']}"
    ws["I12"] = f"{role_name_map['Grammarian']}"
    ws["I13"] = f"{role_name_map['Timer']}"
    ws["I14"] = f"{role_name_map['PC Manager (Vote Counter)']}"
    
    # Table Topic, Prepared Speech
    ws["I16"] = f"{role_name_map['Table Topics Master']}"
    ws["E25"] = f"「{speech_title_map['Speech1']}"
    ws["I25"] = f"{speech_path_map['Speech1']}"
    ws["I26"] = f"{role_name_map['Speech1']}"
    ws["E27"] = f"「{speech_title_map['Speech2']}"
    ws["I27"] = f"{speech_path_map['Speech2']}"
    ws["I28"] = f"{role_name_map['Speech2']}"
    ws["E29"] = f"「{speech_title_map['Speech3']}"
    ws["I29"] = f"{speech_path_map['Speech3']}"
    ws["I30"] = f"{role_name_map['Speech3']}"
    ws["E31"] = f"「{speech_title_map.get('Speech4', '')}"
    ws["I31"] = f"{speech_path_map.get('Speech4','')}"
    ws["I32"] = f"{role_name_map.get('Speech4','')}"
    
    # GE, Evaluators
    ws["I37"] = f"{role_name_map['General Evaluator']}"
    ws["E38"] = f"{evaluator_map['Evaluator1']}"
    ws["I38"] = f"{role_name_map['Evaluator1']}"
    ws["E39"] = f"{evaluator_map['Evaluator2']}"
    ws["I39"] = f"{role_name_map['Evaluator2']}"
    ws["E40"] = f"{evaluator_map['Evaluator3']}"
    ws["I40"] = f"{role_name_map['Evaluator3']}"
    ws["E41"] = f"{evaluator_map.get('Evaluator4','')}"
    ws["I41"] = f"{role_name_map.get('Evaluator4','')}"
    # WOE, Ah-Counter, Grammarian report
    ws["I43"] = f"{role_name_map['Word of the Evening']}"
    ws["I44"] = f"{role_name_map['Ah-Counter']}"
    ws["I45"] = f"{role_name_map['Grammarian']}"
    
    # 保存
    output_path = meeting_title + "_agenda.xlsx"
    wb.save(output_path)
    print(f"✅ Saved Excel to: {output_path}")
    return output_path
