// 公約チェッカーの設定。機種入替などで変わる部分はここを編集してください。
window.KOYAKU_CONFIG = {
  // ①の対象5機種。pattern は機種名に含まれる文字（正規表現）。
  // 全角/半角・スペース・記号の違いは自動で吸収してから照合します。
  targets: [
    { label: '東京喰種', pattern: '東京喰種' },
    { label: 'SAO2', pattern: 'ソードアートオンライン(II|2)' },
    { label: 'GOD', pattern: 'ミリオンゴッド' },
    { label: '戦国乙女5', pattern: '戦国乙女5' },
    { label: 'Lカバネリ', pattern: 'カバネリ' },
  ],

  // 紫②「114.9%機種」とみなす設定6機械割のしきい値（%）
  payout6Threshold: 114.9,

  // 補填した台の塗り色（Excel出力時）
  suppFill: 'FF9DC3E6',
};
