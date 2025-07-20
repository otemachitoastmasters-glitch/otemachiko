from flask import Flask, request, send_file
from agenda_generator import generate_agenda_excel_from_url 
import os

app = Flask(__name__)

@app.route("/generate/")
def generate_agenda():

    print(f"🔗 Fetching agenda ")
    output_path = generate_agenda_excel_from_url("meeting_agenda_template.xlsx")
    print(f"✅ Saved Excel to: {output_path}")

    # ✅ ダウンロードさせるレスポンス
    return send_file(output_path,
                     mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                     as_attachment=True,
                     download_name=output_path)
    
if __name__ == "__main__":
    app.run(host='0.0.0.0', port=10000)
