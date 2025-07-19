import requests
from bs4 import BeautifulSoup
from openpyxl import load_workbook
from openpyxl.utils import get_column_letter
import datetime

def generate_agenda_excel_from_url(mtgid: str, template_path="meeting_agenda_template.xlsx", output_path: str = "generated_agenda.xlsx") -> str:
    
    html_url = "https://tmcsupport.coresv.com/otemachiko/mtgDetailReadonly.php?mtgid=" + mtgid
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
    # 日付・タイトル取得
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
    for tr in agenda_table.find_all("tr")[1:]:
        tds = tr.find_all("td")
        if len(tds) >= 3:
            role = tds[0].text.strip()
            name = tds[1].text.strip()
            detail = tds[2].text.strip()
            title = tds[3].text.strip() if len(tds) > 3 else ""
            agenda.append([role, name, detail, title])
            role_name_map[role] = name

    # Excelテンプレートを読み込み
    wb = load_workbook(template_path)
    ws = wb.active

    # 💡 すべての結合セルを解除
    if ws.merged_cells.ranges:
        print("⚠️ Unmerging cells...")
    for merged_range in list(ws.merged_cells.ranges):
        ws.unmerge_cells(str(merged_range))

    # Excelテンプレートを読み込み
    wb = load_workbook("meeting_agenda_template.xlsx")
    ws = wb.active
    ws.title = "Agenda"
    
    # 💡 すべての結合セルを解除
    if ws.merged_cells.ranges:
        print("⚠️ Unmerging cells...")
    for merged_range in list(ws.merged_cells.ranges):
        ws.unmerge_cells(str(merged_range))
    
    # 🔄 タイトル・日時を書き込む（例：A2セル想定）
    ws["J3"] = f"{meeting_title}　{meeting_datetime}"
    ws["J4"] = f"{venue} {room}"
    
    # TMOE, WOE, Ah-Counter, Grammarian, PC manager
    ws["I9"] = f"{role_name_map['Toastmaster of the Evening']}"
    ws["I10"] = f"{role_name_map['Word of the Evening']}"
    ws["I11"] = f"{role_name_map['Ah-Counter']}"
    ws["I12"] = f"{role_name_map['Grammarian']}"
    ws["I13"] = f"{role_name_map['PC Manager (Vote Counter)']}"
    
    # Table Topic, Prepared Speech
    ws["I16"] = f"{role_name_map['Table Topics Master']}"
    ws["I26"] = f"{role_name_map['Speech1']}"
    ws["I28"] = f"{role_name_map['Speech2']}"
    ws["I30"] = f"{role_name_map['Speech3']}"
    
    # GE, Evaluators
    ws["I37"] = f"{role_name_map['General Evaluator']}"
    ws["I38"] = f"{role_name_map['Evaluator1']}"
    ws["I39"] = f"{role_name_map['Evaluator2']}"
    ws["I40"] = f"{role_name_map['Evaluator3']}"
    
    # WOE, Ah-Counter, Grammarian report
    ws["I42"] = f"{role_name_map['Word of the Evening']}"
    ws["I43"] = f"{role_name_map['Ah-Counter']}"
    ws["I44"] = f"{role_name_map['Grammarian']}"

    # 保存
    wb.save(output_path)
    print(f"✅ Saved Excel to: {output_path}")
    return output_path
