from flask import Flask, request, send_file
from agenda_generator import generate_agenda_excel_from_url

app = Flask(__name__)

@app.route("/generate", methods=["POST"])
def generate():
    try:
        print("✅ Request received")
        print(request.json)

        # 本処理
        return "done"

    except Exception as e:
        print("❌ Error:", e)
        return str(e), 500
    
if __name__ == "__main__":
    app.run(host='0.0.0.0', port=10000)
