from flask import Flask, request, jsonify
from agenda_generator import generate_agenda_excel_from_url  # 関数のあるファイル名に置き換えてください

app = Flask(__name__)

@app.route("/generate/", methods=["GET"])

def generate_agenda():
    url = request.args.get("url")
    if not url:
        return "Missing URL", 400
    try:
        output_path = generate_agenda_excel_from_url(url)
        return send_file(output_path, as_attachment=True)
    except Exception as e:
        import traceback
        traceback.print_exc()
        return f"Error: {str(e)}", 500

if __name__ == "__main__":
    app.run(host='0.0.0.0', port=10000)
