from flask import Flask
import os

# Webアプリの本体を作成
app = Flask(__name__)

@app.route('/')
def home():
    # ブラウザに表示されるメッセージ
    return """
    <h1>建設コスト計算ツール (Web版)</h1>
    <p>現在、システムをWeb用に移植中です。</p>
    <p>Renderでのデプロイが成功すれば、この文字が表示されます！</p>
    """

if __name__ == "__main__":
    # Renderが指定するポート番号を取得して起動
    port = int(os.environ.get("PORT", 5000))
    app.run(host='0.0.0.0', port=port)