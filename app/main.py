import logging
from flask import Flask, render_template, request, jsonify
import string
import random

# 詳細なログ設定
#フォーマットを指定することで、いつ、どのレベルの、どのメッセージが記録されたかを明確にする
handler = logging.StreamHandler()
handler.setLevel(logging.INFO)
handler.setFormatter(logging.Formatter('%(asctime)s - %(name)s - %(levelname)s - %(message)s'))

logger = logging.getLogger(__name__)
logger.setLevel(logging.INFO)
logger.addHandler(handler)


app = Flask(__name__)

@app.route('/')
def index():
    """メインページを表示します。"""
    logger.info("トップページが表示されました。")
    return render_template('index.html')

def generate_password_logic(base_word, length, char_types):
    """パスワード生成のコアロジック"""
    logger.info(f"パスワード生成開始: 単語='{base_word}', 長さ={length}, 文字種={char_types}")

    substitutions = {'l': '1', 'o': '0', 's': '5', 'q': '9', 'b': '8'}

    # 1. 文字を変換
    substituted_word = ""
    for char in base_word:
        # 大文字・小文字を問わず変換
        if char.lower() in substitutions:
            substituted_word += substitutions[char.lower()]
        else:
            substituted_word += char

    logger.info(f"文字変換後: '{substituted_word}'")

    # 2. 文字プールを作成
    character_pool = ""
    if 'lowercase' in char_types:
        character_pool += string.ascii_lowercase
    if 'uppercase' in char_types:
        character_pool += string.ascii_uppercase
    if 'numbers' in char_types:
        character_pool += string.digits
    if 'symbols' in char_types:
        # 一般的な記号に限定
        character_pool += "!@#$%^&*()_+-=[]{}|;:,.<>?"

    # character_poolから、変換後の単語に含まれる文字を除外（多様性のため）
    for char in substituted_word:
        character_pool = character_pool.replace(char, "")

    # 3. 足りない文字を追加
    password_chars = list(substituted_word)
    num_to_add = length - len(password_chars)

    if num_to_add > 0 and character_pool:
        for _ in range(num_to_add):
            password_chars.append(random.choice(character_pool))

    # 長すぎる場合は切り詰める
    elif num_to_add < 0:
        password_chars = password_chars[:length]

    # 4. シャッフル
    random.shuffle(password_chars)

    final_password = "".join(password_chars)
    logger.info(f"生成されたパスワード: '{final_password}'")

    return final_password


@app.route('/generate', methods=['POST'])
def generate_password_route():
    """パスワード生成APIのエンドポイント"""
    try:
        data = request.get_json()
        logger.info(f"APIリクエスト受信: {data}")

        base_word = data.get('base_word', '')
        length = int(data.get('length', 12))
        char_types = data.get('char_types', [])

        if not base_word:
            logger.warning("APIリクエストでbase_wordが指定されませんでした。")
            return jsonify({"error": "単語を入力してください。"}), 400
        if not char_types:
            logger.warning("APIリクエストでchar_typesが指定されませんでした。")
            return jsonify({"error": "使用する文字の種類を1つ以上選択してください。"}), 400

        password = generate_password_logic(base_word, length, char_types)

        return jsonify({"password": password})
    except Exception as e:
        logger.error(f"パスワード生成中にエラーが発生: {e}", exc_info=True)
        return jsonify({"error": "サーバーエラーが発生しました。"}), 500


if __name__ == '__main__':
    # docker-compose.ymlでflask runコマンドを使っているため、このブロックは直接実行時のみ使用される
    app.run(debug=True)
