from flask import Flask, render_template_string, request
import os

app = Flask(__name__)

# HTMLデザイン（入力画面と結果表示）
HTML_TEMPLATE = """
<!DOCTYPE html>
<html>
<head>
    <title>建設コスト計算</title>
    <style>
        body { font-family: sans-serif; margin: 40px; line-height: 1.6; }
        .container { max-width: 500px; margin: auto; padding: 20px; border: 1px solid #ddd; border-radius: 10px; }
        input { width: 100%; padding: 10px; margin: 10px 0; box-sizing: border-box; }
        button { width: 100%; padding: 10px; background-color: #007bff; color: white; border: none; border-radius: 5px; cursor: pointer; }
        .result { margin-top: 20px; padding: 15px; background-color: #f8f9fa; border-left: 5px solid #007bff; }
    </style>
</head>
<body>
    <div class="container">
        <h2>建設コスト計算ツール</h2>
        <form method="POST">
            <label>延床面積 (㎡) を入力してください：</label>
            <input type="number" name="area" step="0.01" placeholder="例: 100" required>
            <button type="submit">計算する</button>
        </form>

        {% if result %}
        <div class="result">
            <h3>計算結果</h3>
            <p>延床面積: {{ area }} ㎡</p>
            <p><strong>概算建設コスト: ¥{{ "{:,.0f}".format(result) }}</strong></p>
            <p><small>※坪単価 80万円（税込 88万円）で仮計算しています</small></p>
        </div>
        {% endif %}
    </div>
</body>
</html>
"""

@app.route('/', methods=['GET', 'POST'])
def home():
    result = None
    area = None
    if request.method == 'POST':
        # 入力された面積を取得
        area = float(request.form.get('area', 0))
        # 簡易計算ロジック（例：坪単価80万 × 3.3で㎡単価換算など）
        # ここにConstructionCost12.pyのロジックを今後移植します
        unit_price_per_sqm = 800000 / 3.305785
        result = area * unit_price_per_sqm * 1.1  # 1.1は消費税

    return render_template_string(HTML_TEMPLATE, result=result, area=area)

if __name__ == "__main__":
    port = int(os.environ.get("PORT", 5000))
    app.run(host='0.0.0.0', port=port)