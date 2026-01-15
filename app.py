from flask import Flask
import os

app = Flask(__name__)

@app.route('/')
def hello():
    return "<h1>建設コスト計算ツール</h1><p>現在、Web版への移植作業中です！</p>"

if __name__ == "__main__":
    port = int(os.environ.get("PORT", 5000))
    app.run(host='0.0.0.0', port=port)