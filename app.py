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
    return render_template("index.html")  # index.html должен лежать в templates/

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
            "rub","kop"]
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
        rub        = int(d["rub"])
        kop        = int(d["kop"])
    except Exception as e:
        return jsonify({"error": f"Неверные типы полей: {e}"}), 400

    if not inn.isdigit() or len(inn) != 12:
        return jsonify({"error": "ИНН: ровно 12 цифр"}), 400
    if rub < 0 or not (0 <= kop <= 99):
        return jsonify({"error": "Проверьте rub (>=0) и kop (0–99)"}), 400

    contractor_name = " ".join(w[:1].upper() + w[1:].lower() for w in name_raw.split())

    # Цена = итог (количество фиксировано 1)
    price = rub + kop / 100.0
    total_kop_all = int(round(price * 100))
    total_rub = total_kop_all // 100
    total_kop = total_kop_all % 100

    context = {
        "act_number": act_number,
        "act_day": act_day,
        "act_month": act_month,

        "contract_number": c_number,
        "contract_day": c_day,
        "contract_month": c_month,

        "contractor_name": contractor_name,
        "contractor_inn": inn,

        # Плейсхолдеры из DOCX
        "rub_per_unit": f"{rub}.{kop:02d}",                 # цена за 1 (и она же единственная)
        "sum_total": f"{total_rub}.{total_kop:02d}",        # итог = цена
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
