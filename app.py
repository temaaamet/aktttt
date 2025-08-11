from flask import Flask, render_template, request, send_file
from docxtpl import DocxTemplate
from num2words import num2words
import io

app = Flask(__name__)

@app.route('/')
def index():
    return render_template('form.html')

@app.route('/generate', methods=['POST'])
def generate():
    data = request.get_json()

    # Преобразуем ФИО исполнителя — каждое слово с заглавной буквы
    contractor_name = " ".join(word.capitalize() for word in data['contractor_name'].split())

    # Сумма
    rub = int(data['rub'])
    kop = int(data['kop'])
    qty = int(data['qty'])
    total = rub * qty + (kop * qty) / 100

    # Прописью (все заглавные)
    total_words = num2words(total, lang='ru', to='currency').upper()

    # Загружаем шаблон
    doc = DocxTemplate("TEMPLATE_ACT.docx")

    context = {
        'act_number': data['act_number'],
        'act_day': data['act_day'],
        'act_month': data['act_month'],
        'contract_day': data['contract_day'],
        'contract_month': data['contract_month'],
        'contract_number': data['contract_number'],
        'contractor_name': contractor_name,
        'contractor_inn': data['contractor_inn'],
        'qty': qty,
        'rub': rub,
        'kop': f"{kop:02}",
        'total': f"{total:.2f}",
        'total_words': total_words
    }

    doc.render(context)

    buffer = io.BytesIO()
    doc.save(buffer)
    buffer.seek(0)

    return send_file(
        buffer,
        as_attachment=True,
        download_name="Акт.docx",
        mimetype="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
    )

if __name__ == '__main__':
    app.run(host='0.0.0.0', port=5000)
