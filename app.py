from flask import Flask, render_template, request, send_file, jsonify
from docxtpl import DocxTemplate
from num2words import num2words
import io

app = Flask(__name__)

def money_to_words(rub: int, kop: int) -> str:
    def r(n):
        if 11 <= n % 100 <= 14: return "РУБЛЕЙ"
        if n % 10 == 1: return "РУБЛЬ"
        if 2 <= n % 10 <= 4: return "РУБЛЯ"
        return "РУБЛЕЙ"
    def k(n):
        if 11 <= n % 100 <= 14: return "КОПЕЕК"
        if n % 10 == 1: return "КОПЕЙКА"
        if 2 <= n % 10 <= 4: return "КОПЕЙКИ"
        return "КОПЕЕК"
    rub = int(rub); kop = int(kop)
    return f"{num2words(rub, lang='ru').upper()} {r(rub)} {num2words(kop, lang='ru').upper()} {k(kop)}"

@app.route("/")
def index():
    return render_template("index.html")  # index.html в templates/

@app.route("/generate", methods=["POST"]
)
def generate():
    try:
        d = request.get_json(force=True)
    except Exception:
        return jsonify({"error": "Некорректный JSON"}), 400

    need = ["act_number","day","month",
            "contract_number","contract_day","contract_month",
            "contractor_name","contractor_inn",
            "qty","rub","kop"]
    miss = [k for k in need if k not in d or str(d[k]).strip() == ""]
    if miss:
        return jsonify({"error": "Отсутствуют поля: " + ", ".join(miss)}), 400

    try:
        act_number = str(d["act_number"]).strip()
        act_day    = str(int(d["day"]))
        act_month  = str(d["month"])
        c_number   = str(d["contract_number"]).strip()
        c_day      = str(int(d["contract_day"]))
        c_month    = str(d["contract_month"])
        name_raw   = str(d["contractor_name"]).strip()
        inn        = str(d["contractor_inn"]).strip()
        qty        = int(d["qty"])                     # количество машин
        total_rub_input = int(d["rub"])               # выручка за месяц (руб)
        total_kop_input = int(d["kop"])               # выручка за месяц (коп)
    except Exception as e:
        return jsonify({"error": f"Неверные типы полей: {e}"}), 400

    if not inn.isdigit() or len(inn) != 12:
        return jsonify({"error": "ИНН: ровно 12 цифр"}), 400
    if qty <= 0:
        return jsonify({"error": "Количество должно быть > 0"}), 400
    if total_rub_input < 0 or not (0 <= total_kop_input <= 99):
        return jsonify({"error": "Выручка: rub >= 0, kop 0–99"}), 400

    contractor_name = " ".join(w[:1].upper() + w[1:].lower() for w in name_raw.split())

    # Общая выручка (в копейках)
    total_cents = total_rub_input * 100 + total_kop_input

    # Цена за 1 машину = выручка / qty (округляем до копеек)
    unit_cents = round(total_cents / qty)
    unit_rub = unit_cents // 100
    unit_kop = unit_cents % 100

    # Итог для документа = та же общая выручка
    total_rub = total_cents // 100
    total_kop = total_cents % 100

    context = {
        "act_number": act_number,
        "act_day": act_day,
        "act_month": act_month,

        "contract_number": c_number,
        "contract_day": c_day,
        "contract_month": c_month,

        "contractor_name": contractor_name,
        "contractor_inn": inn,

        "qty": qty,
        "rub_per_unit": f"{unit_rub}.{unit_kop:02d}",     # цена оклейки одной машины
        "sum_total":   f"{total_rub}.{total_kop:02d}",    # выручка за месяц (итог)
        "sum_total_words": money_to_words(total_rub, total_kop),
    }

    try:
        doc = DocxTemplate("TEMPLATE_ACT.docx")
    except Exception as e:
        return jsonify({"error": f"Нет TEMPLATE_ACT.docx рядом с app.py: {e}"}), 500

    doc.render(context)
    buf = io.BytesIO(); doc.save(buf); buf.seek(0)

    filename = f"Акт_№{act_number}_от_{act_day}_{act_month}.docx"
    return send_file(buf, as_attachment=True, download_name=filename,
                     mimetype="application/vnd.openxmlformats-officedocument.wordprocessingml.document")

if __name__ == "__main__":
    app.run(debug=True)
