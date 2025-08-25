document.addEventListener('DOMContentLoaded', () => {
    // 要素の取得
    const form = document.getElementById('password-form');
    const customLengthInput = document.getElementById('custom-length');
    const lengthOptions = document.querySelectorAll('input[name="length_option"]');
    const resultArea = document.getElementById('result-area');
    const passwordResultEl = document.getElementById('password-result');
    const errorMessageEl = document.getElementById('error-message');
    const copyButton = document.getElementById('copy-button');

    // 文字数「その他」の処理
    lengthOptions.forEach(radio => {
        radio.addEventListener('change', () => {
            customLengthInput.disabled = radio.value !== 'other';
        });
    });

    // フォーム送信時の処理
    form.addEventListener('submit', async (e) => {
        e.preventDefault(); // デフォルトの送信をキャンセル

        // フォームデータの取得
        const baseWord = document.getElementById('base-word').value;
        const selectedCharTypes = [...document.querySelectorAll('input[name="char_type"]:checked')].map(el => el.value);

        let length;
        const selectedLengthOption = document.querySelector('input[name="length_option"]:checked').value;
        if (selectedLengthOption === 'other') {
            length = parseInt(customLengthInput.value, 10);
        } else {
            length = parseInt(selectedLengthOption, 10);
        }

        // バリデーション
        if (!baseWord) {
            showError("ベースとなる単語を入力してください。");
            return;
        }
        if (selectedCharTypes.length === 0) {
            showError("使用する文字の種類を1つ以上選択してください。");
            return;
        }

        // APIリクエストの送信
        try {
            const response = await fetch('/generate', {
                method: 'POST',
                headers: {
                    'Content-Type': 'application/json',
                },
                body: JSON.stringify({
                    base_word: baseWord,
                    length: length,
                    char_types: selectedCharTypes,
                }),
            });

            const data = await response.json();

            if (response.ok) {
                showPassword(data.password);
            } else {
                showError(data.error || '不明なエラーが発生しました。');
            }
        } catch (error) {
            console.error('Fetch Error:', error);
            showError('サーバーとの通信に失敗しました。');
        }
    });

    // パスワード表示処理
    function showPassword(password) {
        resultArea.style.display = 'block';
        passwordResultEl.textContent = password;
        errorMessageEl.textContent = '';
        errorMessageEl.style.display = 'none';
    }

    // エラー表示処理
    function showError(message) {
        resultArea.style.display = 'block';
        passwordResultEl.textContent = '';
        errorMessageEl.textContent = message;
        errorMessageEl.style.display = 'block';
    }

    // コピーボタンの処理
    copyButton.addEventListener('click', () => {
        const password = passwordResultEl.textContent;
        if (password) {
            navigator.clipboard.writeText(password).then(() => {
                // コピー成功時のフィードバック
                copyButton.textContent = '✅';
                setTimeout(() => {
                    copyButton.textContent = '📋';
                }, 1500);
            }).catch(err => {
                console.error('コピーに失敗しました', err);
                alert('コピーに失敗しました。');
            });
        }
    });
});
