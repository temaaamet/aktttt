from flask import Flask, render_template, request, send_file
from docxtpl import DocxTemplate
import os
from num2words import num2words

app = Flask(__name__)

@app.route('/')
def index():
    return render_template('form.html')

@app.route('/generate', methods=['POST'])
def generate():
    data = request.get_json()

    # Расчёт суммы
    qty = int(data['qty'])
    rub = int(data['rub'])
    kop = int(data['kop'])
    rub_per_unit = float(f"{rub}.{kop:02d}")
    sum_total = round(qty * rub_per_unit, 2)

    # Генерация суммы прописью (рубли и копейки отдельно)
    rub_sum = int(sum_total)
    kop_sum = int(round((sum_total - rub_sum) * 100))

    rub_words = num2words(rub_sum, lang='ru')
    kop_words = num2words(kop_sum, lang='ru')

    sum_total_words = f"{rub_words} рублей {kop_words} копеек"

    # Данные для шаблона
    context = {
        "act_number": data['act_number'],
        "act_day": data['day'],
        "act_month": data['month'],
        "contract_number": data['contract_number'],
        "contractor_name": data['contractor_name'],
        "contractor_inn": data['contractor_inn'],
        "qty": qty,
        "rub_per_unit": rub_per_unit,
        "sum_total": sum_total,
        "sum_total_words": sum_total_words,
    }

    doc = DocxTemplate("TEMPLATE_ACT.docx")
    doc.render(context)

    filename = f"Акт_№{data['act_number']}_от_{data['day']}_{data['month']}.docx"
    filepath = os.path.join("/tmp", filename)
    doc.save(filepath)
    return send_file(filepath, as_attachment=True)
