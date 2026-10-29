/* 対象の台番号を、メールなどのテキストやファイルから読み取る（pision から始めたとき、その台に色を付ける）。
 *  テキスト例：「・GOD：789-792/810/815-818」「・リコリコ：318/320/-321/323」「382 384 386 387」
 *  ファイル例：色や⭕️の付いた台の一覧（ホールの出力の抜き出し） ／ 「対象機種・台数・台番」の表（台番に番号を並べたもの）
 */
(function () {
  const TG = {};
  const P = window.KoyakuParser;

  // 1行ずつ読む。「：」の後ろだけを見る。ラベルの無い行は、数字と区切りだけの行のときだけ読む（文章や日付の行は読まない）
  TG.parseText = function (text) {
    const nums = [];
    for (let line of String(text || '').normalize('NFKC').split(/\r?\n/)) {
      line = line.trim();
      if (!line) continue;
      let body = line;
      const ci = line.lastIndexOf(':');
      if (ci >= 0) body = line.slice(ci + 1);
      body = body.replace(/番台|番|台/g, ' ');
      if (!/\d/.test(body) || /[^\d\s\/,、・\-‐−~〜.]/.test(body)) continue;
      for (const tok of body.split(/[\s\/,、・.]+/)) {
        // 「-321」のように前に区切りが付いたものは 321、「789-792」は範囲
        const m = tok.match(/^[-‐−~〜]?(\d+)(?:[-‐−~〜](\d+))?$/);
        if (!m) continue;
        const a = +m[1], b = m[2] ? +m[2] : null;
        if (b != null && b >= a && b - a <= 60) for (let n = a; n <= b; n++) nums.push(n);
        else { nums.push(a); if (b != null) nums.push(b); }
      }
    }
    return [...new Set(nums)];
  };

  const txt = cell => String(P.cellValue(cell) ?? '').normalize('NFKC').trim();

  // ファイル（ExcelJS のブック）から台番号
  TG.fromWorkbook = function (wb, fileName) {
    // 1) 結果報告シート・ホールの出力など、⭕️や色で印の付いた台
    try {
      const p = P.parseWorkbook(wb, fileName);
      const marked = p.machines.filter(m => m.orig).map(m => m.no);
      if (marked.length && marked.length < p.machines.length) return { numbers: marked, how: '印（色・⭕️）の付いた台' };
    } catch (e) { /* 表の形が違えば次へ */ }
    // 2) 「台番」の見出しの列。色の付いたセルがあればそれだけ、無ければ全部（「382 384 386」と並べたセルも読む）
    for (const ws of wb.worksheets) {
      if (ws.state && ws.state !== 'visible') continue;
      const lastR = Math.min(ws.rowCount, (ws.actualRowCount || ws.rowCount) + 200, 5000);
      const lastC = Math.min(ws.columnCount, 40);
      for (let r = 1; r <= Math.min(lastR, 40); r++) {
        for (let c = 1; c <= lastC; c++) {
          const h = txt(ws.getCell(r, c)).replace(/[▲▼\s]/g, '');
          if (!/^台番号?$|^台NO$/i.test(h)) continue;
          const cells = [];
          for (let rr = r + 1; rr <= lastR; rr++) {
            const cell = ws.getCell(rr, c);
            const t = txt(cell);
            if (t) cells.push({ t, filled: !!P.fillKey(cell) });
          }
          const filled = cells.filter(x => x.filled);
          const use = filled.length && filled.length < cells.length ? filled : cells;
          const numbers = TG.parseText(use.map(x => x.t).join('\n'));
          if (numbers.length) return { numbers, how: use === filled ? '「台番」の列の色の付いた台' : '「台番」の列' };
        }
      }
    }
    // 3) どこにも見つからなければ、全部の文字をテキストとして読む
    const lines = [];
    for (const ws of wb.worksheets) {
      const lastR = Math.min(ws.rowCount, (ws.actualRowCount || ws.rowCount) + 200, 5000);
      for (let r = 1; r <= lastR; r++) {
        const row = [];
        for (let c = 1; c <= Math.min(ws.columnCount, 40); c++) { const t = txt(ws.getCell(r, c)); if (t) row.push(t); }
        if (row.length) lines.push(row.join(' '));
      }
    }
    return { numbers: TG.parseText(lines.join('\n')), how: 'ファイルの文字' };
  };

  window.KoyakuTargets = TG;
})();
