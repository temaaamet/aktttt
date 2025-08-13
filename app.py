from flask import Flask, render_template, request, send_file, jsonify
from docxtpl import DocxTemplate, InlineImage
from docx.shared import Mm
from num2words import num2words
import io, os

app = Flask(__name__)

def money_to_words_caps(rub: int, kop: int) -> str:
    def rub_word(n):
        if 11 <= n % 100 <= 14: return "РУБЛЕЙ"
        if n % 10 == 1: return "РУБЛЬ"
        if 2 <= n % 10 <= 4: return "РУБЛЯ"
        return "РУБЛЕЙ"
    def kop_word(n):
        if 11 <= n % 100 <= 14: return "КОПЕЕК"
        if n % 10 == 1: return "КОПЕЙКА"
        if 2 <= n % 10 <= 4: return "КОПЕЙКИ"
        return "КОПЕЕК"
    def num_ru(n: int, fem_last: bool = False) -> str:
        n = int(n)
        s = num2words(n, lang="ru")
        if fem_last and not (11 <= n % 100 <= 14):
            parts = s.split()
            if parts:
                if parts[-1] == "один": parts[-1] = "одна"
                elif parts[-1] == "два": parts[-1] = "две"
            s = " ".join(parts)
        return s.upper()
    rub = int(rub); kop = int(kop)
    return f"{num_ru(rub)} {rub_word(rub)} {num_ru(kop, fem_last=True)} {kop_word(kop)}"

@app.route('/')
def index():
    return render_template('form.html')

@app.route('/generate', methods=['POST'])
def generate():
    # поддерживаем и JSON, и обычную форму
    if request.is_json:
        d = request.get_json(silent=True) or {}
        get = d.get
        file_get = lambda name: None
    else:
        get = request.form.get
        file_get = request.files.get

    try:
        act_number      = (get('act_number') or '').strip()
        act_day         = (get('day') or '').strip()
        act_month       = (get('month') or '').strip()
        contract_day    = (get('contract_day') or '').strip()
        contract_month  = (get('contract_month') or '').strip()
        contract_number = (get('contract_number') or '').strip()
        contractor_name = ' '.join((get('contractor_name') or '').strip().split())
        contractor_name = ' '.join(w[:1].upper() + w[1:].lower() for w in contractor_name.split())
        contractor_inn  = (get('contractor_inn') or '').strip()

        qty             = int(get('qty') or 0)          # количество машин
        total_rub_input = int(get('rub') or 0)          # выручка руб.
        total_kop_input = int(get('kop') or 0)          # выручка коп.
    except Exception as e:
        return jsonify({'error': f'Неверные данные: {e}'}), 400

    # валидации
    if not all([act_number, act_day, act_month, contract_day, contract_month, contract_number, contractor_name, contractor_inn]):
        return jsonify({'error': 'Заполните все поля'}), 400
    if not (contractor_inn.isdigit() and len(contractor_inn) == 12):
        return jsonify({'error': 'ИНН: ровно 12 цифр'}), 400
    if qty <= 0 or total_rub_input < 0 or not (0 <= total_kop_input <= 99):
        return jsonify({'error': 'qty>0, руб>=0, коп 0–99'}), 400

    # расчёты в копейках (без float)
    total_cents = total_rub_input * 100 + total_kop_input
    unit_cents  = int(round(total_cents / qty))
    unit_rub, unit_kop = divmod(unit_cents, 100)
    total_rub, total_kop = divmod(total_cents, 100)

    # открываем шаблон ДО вставки картинок
    try:
        doc = DocxTemplate('TEMPLATE_ACT.docx')
    except Exception as e:
        return jsonify({'error': f'Не найден TEMPLATE_ACT.docx: {e}'}), 500

    # подпись (необязательно)
    signature = ""
    f = file_get('signature')
    if f and f.filename:
        tmp = os.path.join('/tmp', f.filename)
        f.save(tmp)
        signature = InlineImage(doc, tmp, width=Mm(35))

    context = {
        'act_number': act_number,
        'act_day': act_day,
        'act_month': act_month,
        'contract_day': contract_day,
        'contract_month': contract_month,
        'contract_number': contract_number,
        'contractor_name': contractor_name,
        'contractor_inn': contractor_inn,
        'qty': qty,
        'rub_per_unit': f'{unit_rub}.{unit_kop:02d}',      # цена за 1 машину
        'sum_total':    f'{total_rub}.{total_kop:02d}',     # общая выручка
        'sum_total_words': money_to_words_caps(total_rub, total_kop),
        'signature': signature,
    }

    doc.render(context)
    buf = io.BytesIO(); doc.save(buf); buf.seek(0)
    return send_file(
        buf,
        as_attachment=True,
        download_name=f'Акт_№{act_number}_от_{act_day}_{act_month}.docx',
        mimetype='application/vnd.openxmlformats-officedocument.wordprocessingml.document'
    )

if __name__ == '__main__':
    port = int(os.getenv('PORT', '5050'))
    app.run(host='0.0.0.0', port=port, debug=True)
