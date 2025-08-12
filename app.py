from flask import Flask, request, send_file
from agenda_generator import convert_excel_to_pdf, fetch_latest_mtgid, fetch_first_mtgid_by_showdetail, generate_agenda_excel_from_url 
import os

app = Flask(__name__)

@app.route("/generate-pdf/")
def generate_pdf():
    excel_path = generate_agenda_excel_from_url(...)
    pdf_path = excel_path.replace(".xlsx", ".pdf")

    convert_excel_to_pdf(excel_path, pdf_path)

    return send_file(pdf_path, as_attachment=True)

@app.route("/generate/")
def generate_agenda():

    mtgid = fetch_first_mtgid_by_showdetail()
    print(f"🔗 Fetching agenda of {mtgid}")
    output_path = generate_agenda_excel_from_url(mtgid, "meeting_agenda_template.xlsx")
    print(f"✅ Saved Excel to: {output_path}")

    # ✅ ダウンロードさせるレスポンス
    return send_file(output_path,
                     mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                     as_attachment=True,
                     download_name=output_path)
    
if __name__ == "__main__":
    app.run(host='0.0.0.0', port=10000)
