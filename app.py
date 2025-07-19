from flask import Flask

app = Flask(__name__)

@app.route("/")
def hello():
    return "Hello from Render!"
    
@app.route("/generate", methods=["POST"])
def generate():
    data = request.get_json()
    html_url = data.get("html_url")
    # アジェンダExcel/PDFを生成（事前に定義）
    filepath = generate_agenda_excel_from_url(html_url)
    public_url = upload_to_gdrive_or_s3(filepath)
    return jsonify({"file": public_url})
    
if __name__ == "__main__":
    app.run(host='0.0.0.0', port=10000)
