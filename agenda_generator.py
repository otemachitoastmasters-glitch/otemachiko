from bs4 import BeautifulSoup
import requests
import openpyxl
from openpyxl.styles import Alignment
from openpyxl.utils import get_column_letter

def generate_agenda_excel_from_url(mtgid: str, output_path: str = "generated_agenda.xlsx") -> str:
    
    html_url = "https://tmcsupport.coresv.com/otemachiko/mtgDetailReadonly.php?mtgid=" + mtgid
    print(f"🔗 Fetching agenda from: {html_url}")
    res = requests.get(html_url)
    soup = BeautifulSoup(res.content, "html.parser")

    # 会議情報取得
    header_table = soup.find("table", class_="tableCommon")
    rows = header_table.find_all("tr")
    mtg_info = rows[1].find_all("td")
    date = mtg_info[0].text.strip()
    title = mtg_info[1].text.strip()
    venue = mtg_info[3].text.strip()
    room = mtg_info[4].text.strip()

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
    for tr in agenda_table.find_all("tr")[1:]:
        tds = tr.find_all("td")
        if len(tds) >= 3:
            role = tds[0].text.strip()
            name = tds[1].text.strip()
            detail = tds[2].text.strip()
            title = tds[3].text.strip() if len(tds) > 3 else ""
            agenda.append([role, name, detail, title])

    # Excel作成
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Agenda"

    ws.append(["Meeting Date", date])
    ws.append(["Meeting Title", title])
    ws.append(["Venue", venue])
    ws.append(["Room", room])
    ws.append(["Guests", guest])
    ws.append([])
    ws.append(["Role", "Name", "Details", "Speech Title"])

    for row in agenda:
        ws.append(row)

    for col in range(1, 5):
        col_letter = get_column_letter(col)
        ws.column_dimensions[col_letter].width = 30
        for cell in ws[col_letter]:
            cell.alignment = Alignment(wrap_text=True, vertical="top")

    wb.save(output_path)
    print(f"✅ Saved Excel to: {output_path}")
    return output_path
