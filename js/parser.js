/* 結果報告シート(Excel)の読み込み。書式は自動で判別する。
 *  A: 公式テンプレ（台番号の左に公約ごとの⭕️列）
 *  B: ホールの出力データ（仕掛け台を色で塗り、右端の列に「全台/並び/1/2」のメモ）
 *  C: 公約ごとのブロックに台を並べ、全台データは別シート
 */
(function () {
  const P = {};

  // 同じ店舗内で機種をまとめるためのキー（表記はファイル内で揃っている前提）
  P.normKey = s => String(s ?? '').normalize('NFKC').replace(/\s+/g, '');
  // 機種名の照合用（全角半角・スペース・記号の違いを吸収）
  P.normMatch = s => String(s ?? '').normalize('NFKC').toUpperCase()
    .replace(/[\s・･\/\\\-‐－–—!?:;~〜～'"、。,.()（）「」『』\[\]【】＜＞<>]/g, '');

  // 機種名の中心部分（DMMとホールの表記を突き合わせる用）
  // 読み仮名の括弧・記号を外し、Ⅱ→2 などに揃え、L/S/スマスロ/パチスロ等の頭文字を外す
  const ROMAN = { VIII: '8', VII: '7', VI: '6', IV: '4', V: '5', III: '3', II: '2' };
  P.normCore = s => {
    let t = String(s ?? '').normalize('NFKC').toUpperCase()
      .replace(/[（(][^）)]*[）)]/g, '')
      .replace(/[\s・･\/\\\-‐－–—!?:;~〜～'"、。,.「」『』\[\]【】＜＞<>☆★]/g, '')
      .replace(/(?<![A-Z])(VIII|VII|VI|IV|V|III|II)/g, r => ROMAN[r]);
    let prev;
    do {
      prev = t;
      t = t.replace(/^(スマスロ|パチスロ|スロット|LB(?![A-Z])|SB(?![A-Z])|L(?![A-Z])|S(?![A-Z]))/, '');
    } while (t !== prev);
    return t;
  };

  function cellValue(cell) {
    let v = cell.value;
    if (v == null) return null;
    if (typeof v === 'object') {
      if (v instanceof Date) return v;
      if (v.richText) return v.richText.map(t => t.text).join('');
      if ('result' in v) {
        v = v.result;
        if (v == null || (typeof v === 'object' && !(v instanceof Date))) return null;
        return v;
      }
      if ('text' in v) return v.text;
      return null;
    }
    return v;
  }
  const txt = cell => {
    const v = cellValue(cell);
    return v == null ? '' : String(v).normalize('NFKC').trim();
  };

  function toInt(v) {
    if (v == null) return null;
    if (typeof v === 'number') return Math.abs(v - Math.round(v)) < 1e-9 ? Math.round(v) : null;
    const s = String(v).normalize('NFKC').trim();
    return /^\d+$/.test(s) ? parseInt(s, 10) : null;
  }
  function toNum(v) {
    if (v == null || v === '') return null;
    if (typeof v === 'number') return isFinite(v) ? v : null;
    const s = String(v).normalize('NFKC').replace(/[,\s]/g, '');
    return /^-?\d+(\.\d+)?$/.test(s) ? parseFloat(s) : null;
  }

  // ⭕️などの印か
  function isCircle(v) {
    if (v == null) return false;
    const s = String(v).normalize('NFKC').trim();
    return /^[⭕○〇◯●◎✓✔oO⑤⑥56]/.test(s);
  }
  const isSlave = cell => cell.isMerged && cell.master && cell.master !== cell;

  // セルの塗り色を文字列キーにする。白・塗りなしは null
  function fillKey(cell) {
    const f = cell.fill;
    if (!f || f.type !== 'pattern' || !f.pattern || f.pattern === 'none') return null;
    const c = f.fgColor || {};
    if (c.argb) {
      const rgb = String(c.argb).slice(-6).toUpperCase();
      return rgb === 'FFFFFF' ? null : 'rgb:' + rgb;
    }
    if (c.theme != null) {
      if (c.theme === 0 || c.theme === 1) return null; // 白/黒系（縞模様など）
      return 'theme:' + c.theme + ':' + (c.tint || 0);
    }
    if (c.indexed != null) {
      if ([1, 9, 64, 65].includes(c.indexed)) return null;
      return 'idx:' + c.indexed;
    }
    return null;
  }

  const NO_HEADERS = ['台番号', '台番', '台NO', '台No', '台no', '台No.', '台NO.'];

  // 「台番号」と、その右3列以内の「機種名」の組を探す
  function findHeaders(ws) {
    const res = [];
    const maxR = Math.min(ws.rowCount, 60);
    const maxC = Math.min(ws.columnCount, 150);
    for (let r = 1; r <= maxR; r++) {
      for (let c = 1; c <= maxC; c++) {
        // 並べ替えの印（台番▲など）は外して比べる
        if (!NO_HEADERS.includes(txt(ws.getCell(r, c)).replace(/[▲▼△▽↑↓\s]/g, ''))) continue;
        for (let d = 1; d <= 3; d++) {
          if (txt(ws.getCell(r, c + d)).includes('機種')) { res.push({ r, c }); break; }
        }
      }
    }
    return res;
  }

  function mapCols(ws, r, c, endC) {
    const cols = { no: c, last: c };
    for (let x = c + 1; x < endC; x++) {
      const t = txt(ws.getCell(r, x)).toUpperCase();
      if (!t) {
        if (x > cols.last + 1) break; // 2列続けて空なら表の終わり
        continue;
      }
      cols.last = x;
      if (!cols.name && t.includes('機種')) cols.name = x;
      else if (!cols.out && t.includes('アウト')) cols.out = x;
      else if (!cols.safe && (t.includes('セーフ') || t === 'BO' || t.includes('払出'))) cols.safe = x;
      else if (!cols.diff && t.includes('差')) { cols.diff = x; cols.diffLabel = t; }
      // 「総G数」があればそれを優先（「通常G数」はボーナス中を含まないので使わない）
      else if (/総G|総ゲーム/.test(t)) cols.games = x;
      else if (!cols.games && (t.includes('ゲーム') || t === 'G数' || t.includes('回転'))) cols.games = x;
      else if (!cols.bb && (t === 'BB' || t.startsWith('特賞') || t.startsWith('BIG'))) cols.bb = x;
      else if (!cols.rb && t === 'RB') cols.rb = x;
      else if (!cols.setting && t.includes('設定')) cols.setting = x;
    }
    return cols;
  }

  function readRows(ws, hdrRow, cols) {
    const rows = [];
    for (let r = hdrRow + 1; r <= ws.rowCount; r++) {
      const no = toInt(cellValue(ws.getCell(r, cols.no)));
      if (no == null) continue;
      const g = x => (x ? cellValue(ws.getCell(r, x)) : null);
      const name = String(g(cols.name) ?? '').trim();
      if (!name) continue;
      const out = toNum(g(cols.out));
      const rawDiff = toNum(g(cols.diff));
      let safe = toNum(g(cols.safe));
      // セーフの列がなく「差玉」「差数」（ホール目線＝アウト−セーフ）がある場合はセーフを逆算する
      if (safe == null && out != null && rawDiff != null && /差玉|差数|^差$/.test(cols.diffLabel || '')) safe = out - rawDiff;
      rows.push({
        row: r, no, name,
        out, safe, rawDiff,
        // 差枚はお客さん目線（プラス＝出た）に揃える
        diff: out != null && safe != null ? safe - out : rawDiff,
        games: toNum(g(cols.games)),
        bb: toNum(g(cols.bb)), rb: toNum(g(cols.rb)),
        setting: toInt(g(cols.setting)),
      });
    }
    return rows;
  }

  // 取材の種類をテキストから推定
  function guessType(texts) {
    for (const t of texts) {
      if (!t) continue;
      const s = String(t).normalize('NFKC');
      if (/極|虹/.test(s)) return 'kiwami';
      if (/紫/.test(s)) return 'murasaki';
      if (/赤/.test(s)) return 'aka';
      if (/青/.test(s)) return 'ao';
    }
    return null;
  }

  function topTexts(ws, rows) {
    const out = [];
    for (let r = 1; r <= Math.min(rows, ws.rowCount); r++) {
      for (let c = 1; c <= Math.min(ws.columnCount, 20); c++) {
        const t = txt(ws.getCell(r, c));
        if (t && /S級|調査|結果報告/.test(t)) out.push(t);
      }
    }
    return out;
  }

  P.parseWorkbook = function (wb, fileName) {
    const tables = [];
    const blockSheets = [];
    const warnings = [];

    for (const ws of wb.worksheets) {
      if (ws.state && ws.state !== 'visible') continue;
      const hs = findHeaders(ws);
      if (!hs.length) continue;
      const byRow = new Map();
      for (const h of hs) {
        if (!byRow.has(h.r)) byRow.set(h.r, []);
        byRow.get(h.r).push(h);
      }
      for (const [r, list] of byRow) {
        list.sort((a, b) => a.c - b.c);
        if (list.length >= 2) {
          // 書式C：公約ごとのブロック
          const blocks = list.map((h, i) => {
            const endC = i + 1 < list.length ? list[i + 1].c : h.c + 15;
            const cols = mapCols(ws, r, h.c, endC);
            let title = '';
            for (let rr = r - 1; rr >= Math.max(1, r - 3) && !title; rr--) {
              for (let x = h.c; x < endC && !title; x++) title = txt(ws.getCell(rr, x));
            }
            return { c: h.c, endC, cols, title, rows: readRows(ws, r, cols) };
          });
          blockSheets.push({ ws, r, blocks, count: blocks.reduce((s, b) => s + b.rows.length, 0) });
        } else {
          const h = list[0];
          const cols = mapCols(ws, r, h.c, h.c + 20);
          tables.push({ ws, r, cols, rows: readRows(ws, r, cols) });
        }
      }
    }

    tables.sort((a, b) => b.rows.length - a.rows.length);
    const data = tables[0];
    if (!data || !data.rows.length) {
      throw new Error('台番号・機種名の入った全台データが見つかりませんでした。');
    }
    const ws = data.ws;
    const marksSheet = blockSheets.filter(b => b.count > 0).sort((a, b) => b.count - a.count)[0] || null;

    // 台番号の重複チェック
    const seen = new Map();
    const rows = [];
    for (const r of data.rows) {
      if (seen.has(r.no)) { warnings.push(`台番号 ${r.no} が重複しています（${r.row}行目は無視）`); continue; }
      seen.set(r.no, r);
      rows.push(r);
    }

    let format, catCols = [], annC = null;
    const machines = rows.map((r, i) => ({ ...r, i, key: P.normKey(r.name), orig: false, declared: null }));

    if (marksSheet) {
      format = 'C';
      const byNo = new Map(machines.map(m => [m.no, m]));
      for (const b of marksSheet.blocks) {
        for (const e of b.rows) {
          const m = byNo.get(e.no);
          if (!m) { warnings.push(`「${b.title.slice(0, 12)}」の台番号 ${e.no} が全台データにありません`); continue; }
          m.orig = true;
          if (!m.declared) m.declared = b.title;
        }
      }
    } else {
      // 台番号より左の列に見出しがあれば書式A
      for (let x = 1; x < data.cols.no; x++) {
        const label = txt(ws.getCell(data.r, x));
        if (label && !/部店|店舗|ホール|店名/.test(label)) catCols.push({ c: x, label });
      }
      let circles = 0;
      for (const m of machines) {
        for (const cc of catCols) {
          const cell = ws.getCell(m.row, cc.c);
          if (!isSlave(cell) && isCircle(cellValue(cell))) {
            m.orig = true;
            if (!m.declared) m.declared = cc.label;
            circles++;
          }
        }
      }
      format = catCols.length && (circles > 0 || !machines.some(m => fillKey(ws.getCell(m.row, data.cols.no)))) ? 'A' : 'B';

      if (format === 'B') {
        catCols = [];
        // 塗り色の頻度（データ行の6割以上に使われている色は背景色とみなす）
        const freq = new Map();
        const keyOf = m => fillKey(ws.getCell(m.row, data.cols.no)) || fillKey(ws.getCell(m.row, data.cols.name));
        for (const m of machines) {
          const k = keyOf(m);
          if (k) freq.set(k, (freq.get(k) || 0) + 1);
        }
        for (const m of machines) {
          const k = keyOf(m);
          if (k && freq.get(k) <= machines.length * 0.6) m.orig = true;
        }
        // メモ列：表の右側で文字が入っている列
        for (let x = data.cols.last + 1; x <= data.cols.last + 4 && annC == null; x++) {
          if (machines.some(m => txt(ws.getCell(m.row, x)) && toNum(cellValue(ws.getCell(m.row, x))) == null)) annC = x;
        }
        if (annC == null) annC = data.cols.last + 1;
        // メモは同じ機種の続き（並びは続く仕掛け台）にも当てはめる
        let cur = null;
        for (let i = 0; i < machines.length; i++) {
          const m = machines[i];
          const t = txt(ws.getCell(m.row, annC));
          if (t) cur = { text: t.replace(/[・、]$/, ''), key: m.key, run: /並/.test(t) };
          else if (cur) {
            const prev = machines[i - 1];
            const contiguous = prev && Math.abs(prev.no - m.no) === 1;
            const keep = cur.run ? (m.orig && contiguous) : (m.key === cur.key);
            if (!keep) cur = null;
          }
          if (cur && m.orig) m.declared = cur.text;
        }
      }
    }

    // 朝一入場の人数と店舗様コメント（公式テンプレから読み込んだときは出力に引き継ぐ）
    // 出力ファイル名の初期値用に、店舗名（部店名の列）とシート上部の文字（抽出条件の日付など）も拾う
    const meta = { entry: {}, comment: null, store: null, topText: '' };
    for (let x = 1; x <= Math.min(ws.columnCount, 30); x++) {
      if (/部店名|店舗名|ホール名|^店名$/.test(txt(ws.getCell(data.r, x))) && rows.length) {
        const v = String(cellValue(ws.getCell(rows[0].row, x)) ?? '').trim();
        if (v) { meta.store = v; break; }
      }
    }
    const tops = [];
    for (const sh of [ws, marksSheet && marksSheet.ws].filter(Boolean)) {
      for (let r = 1; r <= Math.min(3, sh.rowCount); r++) {
        for (let c = 1; c <= Math.min(sh.columnCount, 20); c++) { const t = txt(sh.getCell(r, c)); if (t) tops.push(t); }
      }
    }
    meta.topText = tops.join(' ');
    for (let r = 1; r <= Math.min(10, ws.rowCount); r++) {
      for (let c = 1; c <= Math.min(30, ws.columnCount); c++) {
        const cell = ws.getCell(r, c);
        if (isSlave(cell)) continue;
        const t = txt(cell);
        if (['抽選(予約)', '再整列', '一般'].includes(t)) meta.entry[t] = toNum(cellValue(ws.getCell(r, c + 1)));
        else if (t.startsWith('※店長カスタム') && !meta.comment) meta.comment = String(cellValue(cell));
      }
    }

    const typeGuess = guessType([
      format === 'C' ? marksSheet.ws.name : ws.name,
      ...topTexts(format === 'C' ? marksSheet.ws : ws, 5),
      fileName,
    ]);

    return {
      fileName,
      format,
      formatLabel: { A: '公式テンプレ（⭕️列）', B: 'ホール出力（色塗り）', C: 'ブロック形式' }[format],
      dataSheet: ws.name,
      marksSheet: marksSheet ? marksSheet.ws.name : null,
      hdrRow: data.r,
      cols: data.cols,
      catCols,
      annC,
      blocks: marksSheet ? marksSheet.blocks.map(b => ({ c: b.c, cols: b.cols, title: b.title, lastRow: Math.max(marksSheet.r, ...b.rows.map(e => e.row)) })) : null,
      typeGuess,
      machines,
      meta,
      warnings,
    };
  };

  window.KoyakuParser = P;
})();
