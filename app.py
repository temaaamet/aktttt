from flask import Flask, render_template, request, send_file, jsonify
from docxtpl import DocxTemplate, InlineImage
from docx.shared import Mm
from num2words import num2words
import io
import os

app = Flask(__name__)

def money_to_words_caps(rub: int, kop: int) -> str:
    # Слова для рублей
    def r(n):
        if 11 <= n % 100 <= 14: return "РУБЛЕЙ"
        if n % 10 == 1: return "РУБЛЬ"
        if 2 <= n % 10 <= 4: return "РУБЛЯ"
        return "РУБЛЕЙ"
    # Слова для копеек
    def k(n):
        if 11 <= n % 100 <= 14: return "КОПЕЕК"
        if n % 10 == 1: return "КОПЕЙКА"
        if 2 <= n % 10 <= 4: return "КОПЕЙКИ"
        return "КОПЕЕК"

    # Женский род для последних слов числительного копеек (одна/две), кроме 11–14
    def num_words_caps_fem(n: int) -> str:
        s = num2words(int(n), lang='ru')
        parts = s.strip().split()
        if parts and not (11 <= n % 100 <= 14):
            if parts[-1] == "один":
                parts[-1] = "одна"
            elif parts[-1] == "два":
                parts[-1] = "две"
        return " ".join(parts).upper()

    rub = int(rub); kop = int(kop)
    rub_words = num2words(rub, lang='ru').upper()
    kop_words = num_words_caps_fem(kop)
    return f"{rub_words} {r(rub)} {kop_words} {k(kop)}"

@app.route("/")
def index():
    return render_template("form.html")  # form.html должен лежать в templates/

@app.route("/generate", methods=["POST"])
def generate():
    try:
        act_number = request.form["act_number"].strip()
        act_day = str(int(request.form["day"]))
        act_month = request.form["month"].strip()
        contract_day = str(int(request.form["contract_day"]))
        contract_month = request.form["contract_month"].strip()
        contract_number = request.form["contract_number"].strip()
        contractor_name = " ".join(w[:1].upper() + w[1:].lower() for w in request.form["contractor_name"].split())
        contractor_inn = request.form["contractor_inn"].strip()
        qty = int(request.form["qty"])
        total_rub_input = int(request.form["rub"])   # выручка (руб)
        total_kop_input = int(request.form["kop"])   # выручка (коп)
    except Exception as e:
        return jsonify({"error": f"Неверные данные формы: {e}"}), 400

    if not contractor_inn.isdigit() or len(contractor_inn) != 12:
        return jsonify({"error": "ИНН должен содержать ровно 12 цифр"}), 400
    if qty <= 0:
        return jsonify({"error": "Количество машин должно быть больше 0"}), 400
    if total_rub_input < 0 or not (0 <= total_kop_input <= 99):
        return jsonify({"error": "Выручка: руб ≥ 0, коп 0–99"}), 400

    # Общая выручка в копейках
    total_cents = total_rub_input * 100 + total_kop_input

    # Цена за 1 машину (в копейках) = выручка / qty, округляем до копейки
    unit_cents = round(total_cents / qty)
    unit_rub = unit_cents // 100
    unit_kop = unit_cents % 100

    # Итог для документа — введённая выручка
    total_rub = total_cents // 100
    total_kop = total_cents % 100

    # Готовим шаблон
    try:
        doc = DocxTemplate("TEMPLATE_ACT.docx")
    except Exception as e:
        return jsonify({"error": f"Не найден TEMPLATE_ACT.docx рядом с app.py: {e}"}), 500

    # Подпись (необязательно)
    signature_img = ""
    file = request.files.get("signature")
    if file and file.filename:
        tmp_path = os.path.join("/tmp", file.filename)
        file.save(tmp_path)
        signature_img = InlineImage(doc, tmp_path, width=Mm(40))  # подстрой ширину по вкусу

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
        "rub_per_unit": f"{unit_rub}.{unit_kop:02d}",        # цена одной машины
        "sum_total":   f"{total_rub}.{total_kop:02d}",       # выручка за месяц
        "sum_total_words": money_to_words_caps(total_rub, total_kop),
        "signature": signature_img,                          # {{ signature }} в шаблоне где нужно
    }

    doc.render(context)
    buf = io.BytesIO()
    doc.save(buf)
    buf.seek(0)

    filename = f"Акт_№{act_number}_от_{act_day}_{act_month}.docx"
    return send_file(
        buf,
        as_attachment=True,
        download_name=filename,
        mimetype="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
    )

if __name__ == "__main__":
    app.run(debug=True)
