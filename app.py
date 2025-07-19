from flask import Flask, request, send_file
from agenda_generator import generate_agenda_excel_from_url

app = Flask(__name__)

@app.route("/generate", methods=["GET"])
def generate():
    url = request.args.get("url")
    if not url:
        return "Missing 'url' parameter", 400

    output_path = generate_agenda_excel_from_url(url)
    return send_file(output_path, as_attachment=True)
    
if __name__ == "__main__":
    app.run(host='0.0.0.0', port=10000)
