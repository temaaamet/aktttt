from flask import Flask, render_template, request, send_file
from docxtpl import DocxTemplate
from num2words import num2words
import io

app = Flask(__name__)

def money_to_words(rub, kop):
    rub = int(rub)
    kop = int(kop)

    def get_ruble_word(n):
        if 11 <= n % 100 <= 14:
            return "РУБЛЕЙ"
        if n % 10 == 1:
            return "РУБЛЬ"
        if 2 <= n % 10 <= 4:
            return "РУБЛЯ"
        return "РУБЛЕЙ"

    def get_kopeck_word(n):
        if 11 <= n % 100 <= 14:
            return "КОПЕЕК"
        if n % 10 == 1:
            return "КОПЕЙКА"
        if 2 <= n % 10 <= 4:
            return "КОПЕЙКИ"
        return "КОПЕЕК"

    rub_words = num2words(rub, lang='ru').upper()
    kop_words = num2words(kop, lang='ru').upper()

    return f"{rub_words} {get_ruble_word(rub)} {kop_words} {get_kopeck_word(kop)}"

@app.route("/")
def index():
    return render_template("form.html")

@app.route("/generate", methods=["POST"])
def generate():
    data = request.get_json()

    rub = int(data['rub'])
    kop = int(data['kop'])
    qty = int(data['qty'])
    price = round(rub + kop / 100, 2)

    contractor_name = ' '.join(word.capitalize() for word in data['contractor_name'].split())

    context = {
        "act_number": data['act_number'],
        "act_day": data['day'],
        "act_month": data['month'],
        "contract_day": data['contract_day'],
        "contract_month": data['contract_month'],
        "contract_number": data['contract_number'],
        "contractor_name": contractor_name,
        "contractor_inn": data['contractor_inn'],
        "qty": qty,
        "rub": rub,
        "kop": kop,
        "price": price,
        "sum_text": money_to_words(rub, kop)
    }

    doc = DocxTemplate("TEMPLATE_ACT.docx")
    doc.render(context)

    byte_io = io.BytesIO()
    doc.save(byte_io)
    byte_io.seek(0)

    return send_file(
        byte_io,
        as_attachment=True,
        download_name="Акт.docx",
        mimetype="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
    )

if __name__ == "__main__":
    app.run(debug=True)

