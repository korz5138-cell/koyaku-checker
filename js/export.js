/* チェック結果・補填をExcelに書き出す。
 * 元のファイルに補填分を書き込み（補填の台は色を変える）、「公約チェック結果」シートを追加する。
 */
(function () {
  const R = window.KoyakuRules, CFG = window.KOYAKU_CONFIG;
  const X = {};
  const solid = argb => ({ type: 'pattern', pattern: 'solid', fgColor: { argb } });
  const STATUS = { ok: '達成', ng: '未達', skip: '対象外' };
  // 読み込んだファイルでは同じ書式のセルがスタイルを共有しているため、複製してから塗る
  const setFill = (cell, fill) => { cell.style = Object.assign({}, cell.style, { fill }); };

  function pickByTag(items, labelOf, tag) {
    const hit = items.find(x => R.TAGS[tag].col.test(labelOf(x)));
    if (hit) return hit;
    return items.find(x => R.TAGS.other.col.test(labelOf(x))) || items[items.length - 1];
  }

  function uniqueName(wb, base) {
    let name = base, i = 2;
    while (wb.getWorksheet(name)) name = `${base}${i++}`;
    return name;
  }

  X.build = async function (S) {
    const p = S.parsed, res = S.result;
    const wb = new ExcelJS.Workbook();
    await wb.xlsx.load(S.buffer.slice(0));
    const ws = wb.getWorksheet(p.dataSheet);
    const fill = solid(CFG.suppFill);
    const tagOf = m => res.assign.get(m) || 'other';
    const supp = p.machines.filter(m => S.supp.has(m.i));
    const blocks = p.blocks ? p.blocks.map(b => ({ ...b })) : null;

    for (const m of supp) {
      const tag = tagOf(m), label = R.TAGS[tag].short;
      if (p.format === 'A' && p.catCols.length) {
        const cc = pickByTag(p.catCols, c => c.label, tag);
        const cell = ws.getCell(m.row, cc.c);
        cell.value = '⭕️';
        setFill(cell, fill);
        setFill(ws.getCell(m.row, p.cols.no), fill);
      } else if (p.format === 'C' && blocks) {
        const mws = wb.getWorksheet(p.marksSheet);
        const b = pickByTag(blocks, x => x.title, tag);
        const r = ++b.lastRow;
        const put = (col, v) => { if (col) { const c = mws.getCell(r, col); c.value = v; setFill(c, fill); } };
        put(b.cols.no, m.no); put(b.cols.name, m.name); put(b.cols.out, m.out); put(b.cols.safe, m.safe);
        put(b.cols.diff, m.rawDiff); put(b.cols.games, m.games);
      } else {
        for (let x = p.cols.no; x <= p.cols.last; x++) setFill(ws.getCell(m.row, x), fill);
        ws.getCell(m.row, p.annC || p.cols.last + 1).value = `補填（${label}）`;
      }
    }

    // 結果シート
    const rs = wb.addWorksheet(uniqueName(wb, '公約チェック結果'));
    rs.columns = [{ width: 10 }, { width: 44 }, { width: 16 }, { width: 8 }, { width: 10 }, { width: 10 }, { width: 10 }, { width: 10 }];
    const bold = row => { row.font = { bold: true }; return row; };
    const t = rs.addRow([`S級ホール調査（${R.TYPES[S.type].label}）公約チェック結果`]);
    t.font = { bold: true, size: 14 };
    const origCnt = p.machines.filter(m => m.orig).length;
    rs.addRow([`${p.fileName}　設置${S.ctx.N}台　⑤⑥ ${origCnt}台${supp.length ? `＋補填${supp.length}台` : ''}`]);
    rs.addRow([]);
    const head = bold(rs.addRow(['項目', '内容', '判定']));
    head.eachCell(c => setFill(c, solid('FFE7EAF0')));
    for (const r of res.reqs) {
      const row = bold(rs.addRow([r.no, r.title, STATUS[r.status]]));
      row.getCell(3).font = { bold: true, color: { argb: r.status === 'ok' ? 'FF1F8A4C' : r.status === 'ng' ? 'FFC8322B' : 'FF6B7280' } };
      for (const l of r.lines) rs.addRow(['', (l.s === 'ng' ? '✕ ' : l.s === 'ok' ? '✓ ' : '') + l.t]);
    }
    rs.addRow([]);
    const th = bold(rs.addRow(['台番号', '機種名', '数えた公約', '補填', 'アウト', 'セーフ', '差枚', 'G数']));
    th.eachCell(c => setFill(c, solid('FFE7EAF0')));
    for (const m of p.machines) {
      if (!m.orig && !S.supp.has(m.i)) continue;
      const isSupp = S.supp.has(m.i);
      const row = rs.addRow([m.no, m.name, R.TAGS[tagOf(m)].short, isSupp ? '○' : '', m.out, m.safe, m.diff, m.games]);
      if (isSupp) row.eachCell({ includeEmpty: true }, c => setFill(c, fill));
    }

    const buf = await wb.xlsx.writeBuffer();
    return new Blob([buf], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' });
  };

  window.KoyakuExport = X;
})();
