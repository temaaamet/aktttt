from flask import Flask, render_template, request, send_file
from docxtpl import DocxTemplate
import io

app = Flask(__name__)

@app.route('/')
def index():
    return render_template('form.html')

@app.route('/generate', methods=['POST'])
def generate():
    data = request.get_json()

    # Inputs
    act_number = data['act_number']
    act_day = data.get('day')          # from form: "Число создания акта"
    act_month = data.get('month')      # from form: "Месяц создания акта"
    contract_day = data['contract_day']
    contract_month = data['contract_month']
    contract_number = data['contract_number']

    # FIO -> Title Case (каждое слово с заглавной)
    contractor_name = " ".join(w.capitalize() for w in data['contractor_name'].split())
    contractor_inn = data['contractor_inn']

    # Pricing
    qty = int(data['qty'])
    rub = int(data['rub'])
    kop = int(data['kop'])
    price_float = rub + kop / 100
    total = round(qty * price_float, 2)

    # Форматы для шаблона
    rub_per_unit = f"{rub}.{kop:02d}"            # например '120.50'
    total_str = f"{total:.2f}"                   # например '241.00'
    # По требованию: не склонять, а выводить цифрами с сокращениями
    total_words_caps = f"{int(total):d} РУБ. {int(round((total - int(total)) * 100)):02d} КОП."

    context = {
        "act_number": act_number,
        "act_day": act_day,
        "act_month": act_month,
        "contract_day": contract_day,
        "contract_month": contract_month,
        "contract_number": contract_number,
        "contractor_name": contractor_name,
        "contractor_inn": contractor_inn,
        "qty": qty,
        "rub_per_unit": rub_per_unit,
        "sum_total": total_str,
        "sum_total_words": total_words_caps
    }

    doc = DocxTemplate("TEMPLATE_ACT.docx")
    doc.render(context)

    buffer = io.BytesIO()
    doc.save(buffer)
    buffer.seek(0)

    download_name = f"Акт_№{act_number}_от_{act_day}_{act_month}.docx"
    return send_file(
        buffer,
        as_attachment=True,
        download_name=download_name,
        mimetype="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
    )

if __name__ == '__main__':
    app.run(host='0.0.0.0', port=5000, debug=True)
