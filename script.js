document.addEventListener('DOMContentLoaded', () => {
    // --- DOM要素の取得 ---
    const prefectureSelect = document.getElementById('prefecture');
    const municipalityInput = document.getElementById('municipality');
    const searchBtn = document.getElementById('searchBtn');
    const resultArea = document.getElementById('resultArea');
    const errorArea = document.getElementById('errorArea');
    const resultFare = document.getElementById('resultFare');
    const resultDays = document.getElementById('resultDays');
    const resultRelayFee = document.getElementById('resultRelayFee');
    const resultTotal = document.getElementById('resultTotal');

    // --- 定数 ---
    const RELAY_FEE_AMOUNT = 1000;
    const prefectures = [
        "北海道", "青森県", "岩手県", "宮城県", "秋田県", "山形県", "福島県", "茨城県", "栃木県", "群馬県",
        "埼玉県", "千葉県", "東京都", "神奈川県", "新潟県", "富山県", "石川県", "福井県", "山梨県", "長野県",
        "岐阜県", "静岡県", "愛知県", "三重県", "滋賀県", "京都府", "大阪府", "兵庫県", "奈良県", "和歌山県",
        "鳥取県", "島根県", "岡山県", "広島県", "山口県", "徳島県", "香川県", "愛媛県", "高知県", "福岡県",
        "佐賀県", "長崎県", "熊本県", "大分県", "宮崎県", "鹿児島県", "沖縄県"
    ];

    // --- マスターデータの読み込み ---
    const fetchData = async () => {
        try {
            const [fareRes, daysRes, relayRes] = await Promise.all([
                fetch('./data/fare.json'),
                fetch('./data/days.json'),
                fetch('./data/relay.json')
            ]);
            const fareMaster = await fareRes.json();
            const daysMaster = await daysRes.json();
            const relayFeeMaster = await relayRes.json();
            return { fareMaster, daysMaster, relayFeeMaster };
        } catch (error) {
            console.error('マスターデータの読み込みに失敗しました:', error);
            errorArea.textContent = 'データの読み込みに失敗しました。ページを再読み込みしてください。';
            errorArea.classList.remove('hidden');
            return null;
        }
    };

    // --- 初期化処理 ---
    const initialize = (masters) => {
        if (!masters) return;

        const { fareMaster, daysMaster, relayFeeMaster } = masters;

        // 都道府県プルダウンの生成
        let optionsHtml = '<option value="">-- 選択してください --</option>';
        prefectures.forEach(pref => {
            optionsHtml += `<option value="${pref}">${pref}</option>`;
        });
        prefectureSelect.innerHTML = optionsHtml;

        // 検索ボタンのイベントリスナー
        searchBtn.addEventListener('click', () => {
            const selectedPrefecture = prefectureSelect.value;
            const enteredMunicipality = municipalityInput.value.trim().replace(/\s/g, ''); // 入力から空白を除去

            if (!selectedPrefecture || !enteredMunicipality) {
                errorArea.textContent = '都道府県と市区町村を両方入力・選択してください。';
                errorArea.classList.remove('hidden');
                resultArea.classList.add('hidden');
                return;
            }

            // 1. 運賃の検索
            let fare = 0;
            let fareText = '';
            const fareData = fareMaster[selectedPrefecture];
            if (fareData === -1) {
                fareText = '都度確認';
                fare = 0;
            } else if (fareData !== undefined) {
                fare = fareData;
                fareText = `¥ ${fare.toLocaleString()}`;
            } else {
                fareText = 'データなし';
            }

            // 2. 日数の検索 (改善ロジック)
            let daysText = 'データなし';
            let daysData = daysMaster.find(item => {
                const normalizedPrefecture = item.prefecture.replace(/\s/g, '');
                const normalizedMasterMunicipality = item.municipality.replace(/\s/g, '');
                return normalizedPrefecture === selectedPrefecture && enteredMunicipality.startsWith(normalizedMasterMunicipality);
            });

            if (!daysData && selectedPrefecture === '沖縄県') {
                 daysData = daysMaster.find(item => item.prefecture === '沖縄県' && item.municipality === '');
            }

            if (daysData) {
                daysText = daysData.daysText;
                if (daysData.confirmationNeeded) {
                    daysText += ' <span class="text-red-500 font-semibold">※日数要確認</span>';
                }
            }

            // 3. 中継料の検索
            let relayFee = 0;
            let relayFeeText = '';
            const relayFeeData = relayFeeMaster.find(item => {
                if (item.prefecture !== selectedPrefecture) {
                    return false;
                }
                const normalizedMasterMunicipality = item.municipality.replace(/\s/g, '');
                 if (enteredMunicipality.startsWith(normalizedMasterMunicipality)) {
                    if (!item.partialArea || enteredMunicipality === normalizedMasterMunicipality) {
                        return true;
                    }
                    const partials = item.partialArea.replace(/[()（）\s※▲「」]/g, '').split(/、|・/);
                    return partials.some(p => p && enteredMunicipality.includes(p));
                }
                return false;
            });

            if (relayFeeData) {
                relayFee = RELAY_FEE_AMOUNT;
                const partialText = relayFeeData.partialArea ? relayFeeData.partialArea.trim() : '';
                if (partialText) {
                    relayFeeText = `<span class="text-red-600">${partialText}<br>+ ¥ ${relayFee.toLocaleString()}</span>`;
                } else {
                    relayFeeText = `<span class="text-red-600">あり<br>+ ¥ ${relayFee.toLocaleString()}</span>`;
                }
            } else {
                relayFee = 0;
                relayFeeText = 'なし';
            }

            // 4. 合計金額の計算
            let totalText = '';
            if (fareData === -1) {
                totalText = '都度確認';
            } else {
                const totalAmount = fare + relayFee;
                totalText = `¥ ${totalAmount.toLocaleString()}`;
            }

            // --- 結果表示 ---
            resultFare.innerHTML = fareText;
            resultDays.innerHTML = daysText;
            resultRelayFee.innerHTML = relayFeeText;
            resultTotal.innerHTML = totalText;

            resultArea.classList.remove('hidden');
            errorArea.classList.add('hidden');
        });
    };

    // --- 実行 ---
    fetchData().then(masters => {
        initialize(masters);
    });
});
