from flask import Flask, request, jsonify
from your_module import generate_agenda_excel_from_url  # 関数のあるファイル名に置き換えてください

app = Flask(__name__)

@app.route("/generate/", methods=["GET"])
def generate():
    url = request.args.get("url")
    if not url:
        return jsonify({"error": "Missing URL"}), 400

    try:
        # アジェンダExcelを生成して、ファイルのURLを返す想定
        file_url = generate_agenda_excel_from_url(url)
        return jsonify({"message": "success", "file_url": file_url})
    except Exception as e:
        return jsonify({"error": str(e)}), 500

if __name__ == "__main__":
    app.run(host='0.0.0.0', port=10000)
