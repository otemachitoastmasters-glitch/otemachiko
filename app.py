from flask import Flask, request, send_file
from agenda_generator import generate_agenda_excel_from_url 
import os

app = Flask(__name__)

@app.route("/generate/")
def generate_agenda():
    mtgid = request.args.get("mtgid")
    if not mtgid:
        return "Missing mtgid parameter", 400

    print(f"🔗 Fetching agenda from: {mtgid}")
    output_path = generate_agenda_excel_from_url(mtgid)
    print(f"✅ Saved Excel to: {output_path}")

    # ✅ ダウンロードさせるレスポンス
    return send_file(output_path,
                     mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                     as_attachment=True,
                     download_name="Generated_Agenda.xlsx")
    
if __name__ == "__main__":
    app.run(host='0.0.0.0', port=10000)
