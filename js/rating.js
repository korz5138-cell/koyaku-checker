/* 取材の評価（☆1〜3）と、共有用の文。
 * 出率 = セーフの合計 ÷ アウトの合計。対象 = 仕掛け台（⑤⑥。補填した台も含む）。1台は1回だけ数える。
 */
(function () {
  const RT = {};
  const sum = (list, f) => list.reduce((s, m) => s + (f(m) || 0), 0);
  const hasData = m => m.out != null && m.safe != null && m.out > 0;

  RT.rate = function (machines, isTarget, type, equiv) {
    const cfg = (window.KOYAKU_CONFIG.rating || {})[equiv ? 'equiv' : 'normal'] || [110, 107];
    const targets = machines.filter(isTarget);
    const tData = targets.filter(hasData);
    const allData = machines.filter(m => m.out != null && m.safe != null);
    const out = sum(tData, m => m.out), safe = sum(tData, m => m.safe);
    // 表示と同じ小数1桁で判定する（画面に 109.0% と出ているのに ☆3 にならない、を防ぐ）
    const rate = out > 0 ? Math.round((safe / out) * 1000) / 10 : null;
    const totalDiff = sum(allData, m => m.safe - m.out);
    const tDiff = sum(tData, m => m.safe - m.out);
    let stars = null, capped = false;
    if (rate != null) {
      stars = rate >= cfg[0] ? 3 : rate > cfg[1] ? 2 : 1;
      // 極だけは、全体差枚がプラスでなければ ☆3 でも ☆2
      if (type === 'kiwami' && stars === 3 && !(totalDiff > 0)) { stars = 2; capped = true; }
    }
    return {
      stars, capped, rate, thresholds: cfg,
      nTarget: targets.length, nTargetData: tData.length,
      targetAvgDiff: tData.length ? tDiff / tData.length : null,
      totalDiff, avgDiff: allData.length ? totalDiff / allData.length : null, nAll: allData.length,
    };
  };

  const signed = n => (n > 0 ? '+' : n < 0 ? '-' : '') + Math.abs(Math.round(n)).toLocaleString('ja-JP');
  RT.signed = signed;

  // 共有用の文（ユーザーの見本の形）
  RT.shareText = function ({ date, store, lottery, r }) {
    const lines = [];
    if (date) lines.push(`${+date.slice(5, 7)}月${+date.slice(8, 10)}日`, '');
    if (store) lines.push(store, '');
    if (lottery != null && lottery !== '') lines.push(`抽選${lottery}人`);
    lines.push(`評価${r.stars ? '☆'.repeat(r.stars) : '－'}`, '');
    lines.push(`総差枚${signed(r.totalDiff)}枚 平均差枚${r.avgDiff != null ? signed(r.avgDiff) : '－'}枚`, '');
    lines.push(`・1台あたりの平均差枚数${r.targetAvgDiff != null ? signed(r.targetAvgDiff) : '－'}枚(平均出率${r.rate != null ? r.rate.toFixed(1) : '－'}%)`);
    return lines.join('\n');
  };

  window.KoyakuRating = RT;
})();
