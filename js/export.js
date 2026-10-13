/* チェック結果をExcelに書き出す。
 * 読み込んだ書式にかかわらず、公式の結果報告シート（template/report_template.xlsx）の形で出力する。
 * 取材の種類のシートに全台データを入れ、数えた公約の列に⭕️を付ける。補填した台は色を変える。
 */
(function () {
  const R = window.KoyakuRules, CFG = window.KOYAKU_CONFIG;
  const X = {};
  const TEMPLATE_URL = 'template/report_template.xlsx';
  const SHEET = { kiwami: '極', murasaki: '紫', aka: '赤', ao: '青' };
  const HEAD_ROW = 11;
  const TEMPLATE_LAST_ROW = 557; // テンプレートの数式が入っている最終行

  const solid = argb => ({ type: 'pattern', pattern: 'solid', fgColor: { argb } });
  // 読み込んだファイルでは同じ書式のセルがスタイルを共有しているため、複製してから塗る
  const setFill = (cell, fill) => { cell.style = Object.assign({}, cell.style, { fill }); };
  const text = cell => {
    let v = cell.value;
    if (v && typeof v === 'object') v = v.richText ? v.richText.map(t => t.text).join('') : (v.result ?? v.text ?? '');
    return String(v ?? '').normalize('NFKC').trim();
  };

  // テンプレートの集計欄の数式を計算する（保護ビューやプレビューでも値が見えるよう、計算結果も書き込むため）
  function evalFormula(ws, f) {
    const num = addr => {
      const v = ws.getCell(addr).value;
      const x = v && typeof v === 'object' ? v.result : v;
      return typeof x === 'number' ? x : null;
    };
    const colNums = (col, r1, r2) => {
      const xs = [];
      for (let r = r1; r <= r2; r++) { const x = num(col + r); if (x != null) xs.push(x); }
      return xs;
    };
    let m;
    if ((m = f.match(/^COUNT\(([A-Z]+)(\d+):\1(\d+)\)$/))) return colNums(m[1], +m[2], +m[3]).length;
    if ((m = f.match(/^AVERAGE\(([A-Z]+)(\d+):\1(\d+)\)$/))) {
      const xs = colNums(m[1], +m[2], +m[3]);
      return xs.length ? xs.reduce((s, x) => s + x, 0) / xs.length : undefined;
    }
    if ((m = f.match(/^([A-Z]+\d+)-([A-Z]+\d+)$/))) {
      const a = num(m[1]), b = num(m[2]);
      return a != null && b != null ? a - b : undefined;
    }
    if ((m = f.match(/^ROUNDUP\(([A-Z]+\d+),0\)$/))) {
      const a = num(m[1]);
      return a != null ? Math.ceil(Math.round(a * 1e9) / 1e9) : undefined;
    }
    if ((m = f.match(/^\(([A-Z]+\d+)\*([\d.]+)\)(?:\/([\d.]+))?$/))) {
      const a = num(m[1]);
      return a != null ? (a * +m[2]) / (m[3] ? +m[3] : 1) : undefined;
    }
    if ((m = f.match(/^LOOKUP\(([A-Z]+\d+),\{([\d,]+)\},\{(.+)\}\)$/))) {
      const a = num(m[1]);
      const keys = m[2].split(',').map(Number), vals = m[3].split(',').map(s => s.replace(/"/g, ''));
      let i = -1;
      keys.forEach((k, j) => { if (a != null && a >= k) i = j; });
      return i >= 0 ? vals[i] : undefined;
    }
    return undefined;
  }

  function pickByTag(items, tag) {
    const hit = items.find(x => R.TAGS[tag].col.test(x.label));
    if (hit) return hit;
    return items.find(x => R.TAGS.other.col.test(x.label)) || items[items.length - 1];
  }

  function findCell(ws, pred, maxRow) {
    for (let r = 1; r <= maxRow; r++) {
      for (let c = 1; c <= 30; c++) {
        const cell = ws.getCell(r, c);
        if (pred(text(cell))) return { r, c, cell };
      }
    }
    return null;
  }

  X.build = async function (S) {
    const p = S.parsed, res = S.result;
    const resp = await fetch(TEMPLATE_URL, { cache: 'no-cache' });
    if (!resp.ok) throw new Error('テンプレートを読み込めませんでした');
    const wb = new ExcelJS.Workbook();
    await wb.xlsx.load(await resp.arrayBuffer());
    const ws = wb.getWorksheet(SHEET[S.type]);

    // 見出し（11行目）から列の位置を読む
    const col = {};
    const cats = [];
    for (let c = 1; c <= 30; c++) {
      const t = text(ws.getCell(HEAD_ROW, c));
      if (!t) continue;
      if (t === '台番号') col.no = c;
      else if (t === '機種名') col.name = c;
      else if (t === 'アウト') col.out = c;
      else if (t === 'セーフ') col.safe = c;
      else if (t === 'BB') col.bb = c;
      else if (t === 'RB') col.rb = c;
      else if (t.includes('ゲーム')) col.games = c;
      else if (t.startsWith('差枚')) col.diff = c;
      if (!col.no) cats.push({ c, label: t });
    }
    const letter = c => ws.getColumn(c).letter;
    const fill = solid(CFG.suppFill);
    const tagOf = m => res.assign.get(m) || 'other';

    let r = HEAD_ROW + 1;
    for (const m of p.machines) {
      const put = (c, v) => { if (c && v != null) ws.getCell(r, c).value = v; };
      put(col.no, m.no);
      put(col.name, m.name);
      put(col.out, m.out);
      put(col.safe, m.safe);
      put(col.bb, m.bb);
      put(col.rb, m.rb);
      put(col.games, m.games);
      if (col.diff) {
        ws.getCell(r, col.diff).value = { formula: `${letter(col.safe)}${r}-${letter(col.out)}${r}`, result: (m.safe || 0) - (m.out || 0) };
      }
      const supp = S.supp.has(m.i);
      if (m.orig || supp) {
        const cell = ws.getCell(r, pickByTag(cats, tagOf(m)).c);
        cell.value = '⭕️';
        if (supp) {
          setFill(cell, fill);
          setFill(ws.getCell(r, col.no), fill);
        }
      }
      r++;
    }

    // テンプレートの行数を超える店舗は、集計の数式の範囲を広げる
    const lastRow = r - 1;
    if (lastRow > TEMPLATE_LAST_ROW) {
      for (let rr = 4; rr <= 8; rr++) {
        ws.getRow(rr).eachCell(cell => {
          const v = cell.value;
          if (v && typeof v === 'object' && v.formula && v.formula.includes(String(TEMPLATE_LAST_ROW))) {
            cell.value = { formula: v.formula.split(String(TEMPLATE_LAST_ROW)).join(String(lastRow)) };
          }
        });
      }
    }

    // データのない行の差枚の数式は 0（空欄の引き算）
    if (col.diff) {
      for (let rr = lastRow + 1; rr <= TEMPLATE_LAST_ROW; rr++) {
        const v = ws.getCell(rr, col.diff).value;
        if (v && typeof v === 'object' && v.formula) ws.getCell(rr, col.diff).value = { formula: v.formula, result: 0 };
      }
    }
    // 集計欄（設置数・平均・全台系必要台数など）の計算結果を入れる。参照し合う数式があるので数回まわす
    for (let pass = 0; pass < 3; pass++) {
      for (let rr = 4; rr <= 8; rr++) {
        for (let c = 1; c <= 20; c++) {
          const cell = ws.getCell(rr, c);
          const v = cell.value;
          if (!v || typeof v !== 'object' || !v.formula) continue;
          const result = evalFormula(ws, v.formula);
          if (result !== undefined) cell.value = { formula: v.formula, result };
        }
      }
    }
    // Excelで開いたときにも計算し直させる
    wb.calcProperties = Object.assign({}, wb.calcProperties, { fullCalcOnLoad: true });

    // 見出し行にフィルターを付ける（テンプレートの極シートと同じ範囲の形）
    ws.autoFilter = `A${HEAD_ROW}:${ws.getColumn(ws.columnCount).letter}${Math.max(lastRow, TEMPLATE_LAST_ROW)}`;
    // テンプレートに残っている極シート用のフィルター範囲の名前は、書き出すとシートとの対応が外れるため消す（Excelが開くときに作り直す）
    wb.definedNames.model = wb.definedNames.model.filter(x => x.name !== '_xlnm._FilterDatabase');

    // 朝一入場・店舗様コメントを引き継ぐ
    for (const [label, v] of Object.entries(p.meta.entry || {})) {
      const hit = v != null && findCell(ws, t => t === label, 10);
      if (hit) ws.getCell(hit.r, hit.c + 1).value = v;
    }
    if (p.meta.comment) {
      const hit = findCell(ws, t => t.startsWith('※店長カスタム'), 10);
      if (hit) hit.cell.value = p.meta.comment;
    }

    // 開いたときに該当の取材のシートが表示されるようにする（複数シートの同時選択は避ける）
    const idx = wb.worksheets.indexOf(ws);
    wb.views = [{ x: 0, y: 0, width: 20000, height: 12000, firstSheet: 0, activeTab: idx, visibility: 'visible' }];
    for (const sh of wb.worksheets) {
      sh.views = (sh.views && sh.views.length ? sh.views : [{}]).map(v => ({ ...v, tabSelected: sh === ws }));
    }

    const buf = await wb.xlsx.writeBuffer();
    return new Blob([buf], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' });
  };

  window.KoyakuExport = X;
})();
