from flask import Flask, render_template, request, send_file, jsonify
from docxtpl import DocxTemplate
from num2words import num2words
import io

app = Flask(__name__)

# === Полные слова КАПС + корректный женский род для копеек ===
def money_to_words_caps(rub: int, kop: int) -> str:
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
    rub_words = num2words(int(rub), lang='ru').upper()
    # женский род для последнего слова (одна/две), кроме 11–14
    kop_int = int(kop)
    kop_words = num2words(kop_int, lang='ru').split()
    if kop_words and not (11 <= kop_int % 100 <= 14):
        if kop_words[-1] == "один": kop_words[-1] = "одна"
        elif kop_words[-1] == "два": kop_words[-1] = "две"
    kop_words = " ".join(kop_words).upper()
    return f"{rub_words} {r(int(rub))} {kop_words} {k(kop_int)}"

@app.route('/')
def index():
    return render_template('form.html')  # или index.html — по твоему проекту

@app.route('/generate', methods=['POST'])
def generate():
    # Принимаем и JSON, и multipart (на будущее)
    data = None
    if request.is_json:
        data = request.get_json(silent=True) or {}
        getter = data.get
    else:
        getter = request.form.get

    # --- Поля ---
    act_number = (getter('act_number') or '').strip()
    act_day = (getter('day') or '').strip()
    act_month = (getter('month') or '').strip()
    contract_day = (getter('contract_day') or '').strip()
    contract_month = (getter('contract_month') or '').strip()
    contract_number = (getter('contract_number') or '').strip()
    contractor_name = " ".join((getter('contractor_name') or '').strip().split())
    contractor_name = " ".join(w[:1].upper() + w[1:].lower() for w in contractor_name.split())
    contractor_inn = (getter('contractor_inn') or '').strip()
    try:
        qty = int(getter('qty') or 0)
        # ВАЖНО: rub/kop — это ВЫРУЧКА ЗА МЕСЯЦ, не цена за единицу
        total_rub_input = int(getter('rub') or 0)
        total_kop_input = int(getter('kop') or 0)
    except Exception as e:
        return jsonify({"error": f"Неверные типы полей: {e}"}), 400

    # --- Валидация ---
    missing = [k for k in ["act_number","day","month","contract_day","contract_month","contract_number",
                           "contractor_name","contractor_inn"] if not locals()[k if k!='contractor_name' else 'contractor_name']]
    if missing:
        return jsonify({"error": "Отсутствуют поля: " + ", ".join(missing)}), 400
    if not contractor_inn.isdigit() or len(contractor_inn) != 12:
        return jsonify({"error": "ИНН: ровно 12 цифр"}), 400
    if qty <= 0:
        return jsonify({"error": "Количество должно быть > 0"}), 400
    if total_rub_input < 0 or not (0 <= total_kop_input <= 99):
        return jsonify({"error": "Выручка: rub >= 0, kop 0–99"}), 400

    # --- Расчёты без плавающей точки ---
    total_cents = total_rub_input * 100 + total_kop_input               # выручка за месяц
    unit_cents = int(round(total_cents / qty))                           # цена за 1 машину (округление до копейки)
    unit_rub, unit_kop = divmod(unit_cents, 100)
    total_rub, total_kop = divmod(total_cents, 100)

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
        "rub_per_unit": f"{unit_rub}.{unit_kop:02d}",                  # цена одной машины
        "sum_total":   f"{total_rub}.{total_kop:02d}",                 # выручка за месяц
        "sum_total_words": money_to_words_caps(total_rub, total_kop),  # КАПС полными словами
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
