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

    # 日付・タイトル取得
    meeting_title = soup.find("div", class_="agendaTitle").text.strip()
    meeting_datetime = soup.find("div", class_="agendaDatetime").text.strip()

    # スピーカー情報取得（仮：テーブル例）
    rows = soup.select("table.tableNormal tbody tr")

    # Excelテンプレートを読み込み
    wb = load_workbook(template_path)
    ws = wb.active

    # 💡 すべての結合セルを解除
    if ws.merged_cells.ranges:
        print("⚠️ Unmerging cells...")
    for merged_range in list(ws.merged_cells.ranges):
        ws.unmerge_cells(str(merged_range))

    # 🔄 タイトル・日時を書き込む（例：A2セル想定）
    ws["A2"] = f"{meeting_title}　{meeting_datetime}"

    # 💡 テーブル情報を反映（ここはデモ用：本番は項目に応じて座標調整要）
    start_row = 10  # 実際の開始位置に合わせて調整
    for i, row in enumerate(rows):
        cols = row.find_all("td")
        if len(cols) >= 2:
            time = cols[0].text.strip()
            role = cols[1].text.strip()
            member = cols[2].text.strip() if len(cols) >= 3 else ""
            # 書き込む位置を調整（例：列B、C、D）
            ws.cell(row=start_row + i, column=2).value = time
            ws.cell(row=start_row + i, column=3).value = role
            ws.cell(row=start_row + i, column=4).value = member

    # 保存
    wb.save(output_path)
    print(f"✅ Saved Excel to: {output_path}")
    return output_path
