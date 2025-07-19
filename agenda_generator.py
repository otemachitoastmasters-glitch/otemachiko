from bs4 import BeautifulSoup
import requests
import openpyxl
from openpyxl.styles import Alignment
from openpyxl.utils import get_column_letter

def generate_agenda_excel_from_url(mtgid: str, template_path, output_path: str = "generated_agenda.xlsx") -> str:
    
    html_url = "https://tmcsupport.coresv.com/otemachiko/mtgDetailReadonly.php?mtgid=" + mtgid
    print(f"🔗 Fetching agenda from: {html_url}")
    try:
        # Webページ取得
        res = requests.get(html_url)
        res.raise_for_status()
        soup = BeautifulSoup(res.text, "html.parser")

        # Excelテンプレート読み込み
        wb = openpyxl.load_workbook(template_path)
        ws = wb.active

        # ヘッダー情報取得
        headers = soup.find_all("h2")
        if len(headers) >= 2:
            ws["D1"] = headers[0].text.strip()  # 例: The 185th meeting　2025/07/23（Wed）
            ws["D2"] = headers[1].text.strip()  # 例: hybrid　St. Luke’s Garden Tower 15F

        # テーブルの取得（各セッション情報）
        tables = soup.find_all("table")
        start_row = 7  # 書き込み開始行

        for table in tables:
            rows = table.find_all("tr")
            for r_idx, row in enumerate(rows):
                tds = row.find_all("td")
                for c_idx, cell in enumerate(tds):
                    text = cell.text.strip()
                    ws.cell(row=start_row + r_idx, column=2 + c_idx).value = text
            start_row += len(rows) + 1  # 1行空けて次のセクションへ

        wb.save(output_path)
        print(f"✅ Saved Excel to: {output_path}")
        return output_path
        
    #return output_path
