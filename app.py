from flask import Flask, render_template, request, send_file
from docxtpl import DocxTemplate
import os

app = Flask(__name__)

@app.route('/')
def index():
    return render_template('form.html')

@app.route('/generate', methods=['POST'])
def generate():
    data = request.get_json()
    doc = DocxTemplate("TEMPLATE_ACT.docx")
    doc.render(data)
    filename = f"Акт_№{data['act_number']}_от_{data['day']}_{data['month']}.docx"
    filepath = os.path.join("/tmp", filename)
    doc.save(filepath)
    return send_file(filepath, as_attachment=True)
