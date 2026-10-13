/* データファイル（ビューアで書き出した「日付_店名_台データオンライン.csv / .xlsx」など）を読み、
 * pision と同じ形（unitId・model.name・bb・rb・games）にそろえる。照合と反映は KoyakuPision.apply を使う。
 */
(function () {
  const XD = {};
  const norm = s => String(s ?? '').normalize('NFKC').trim();
  const toNum = v => {
    if (v == null || v === '') return null;
    if (typeof v === 'number') return isFinite(v) ? v : null;
    const s = norm(v).replace(/,/g, '');
    return /^-?\d+(\.\d+)?$/.test(s) ? parseFloat(s) : null;
  };

  function parseCSV(text) {
    const rows = [];
    let row = [], cell = '', q = false;
    for (let i = 0; i < text.length; i++) {
      const ch = text[i];
      if (q) {
        if (ch === '"' && text[i + 1] === '"') { cell += '"'; i++; }
        else if (ch === '"') q = false;
        else cell += ch;
      } else if (ch === '"') q = true;
      else if (ch === ',') { row.push(cell); cell = ''; }
      else if (ch === '\n' || ch === '\r') {
        if (ch === '\r' && text[i + 1] === '\n') i++;
        row.push(cell); rows.push(row); row = []; cell = '';
      } else cell += ch;
    }
    if (cell || row.length) { row.push(cell); rows.push(row); }
    return rows;
  }

  async function readRows(file) {
    const buf = await file.arrayBuffer();
    if (/\.csv$/i.test(file.name)) {
      let text = new TextDecoder('utf-8').decode(buf);
      if (text.includes('�')) text = new TextDecoder('shift_jis').decode(buf); // Excelで保存したCSV（Shift_JIS）
      return parseCSV(text.replace(/^﻿/, ''));
    }
    const wb = new ExcelJS.Workbook();
    await wb.xlsx.load(buf);
    const ws = wb.worksheets.find(w => w.state !== 'hidden') || wb.worksheets[0];
    const rows = [];
    ws.eachRow({ includeEmpty: false }, row => {
      const vals = [];
      for (let c = 1; c <= ws.columnCount; c++) {
        let v = row.getCell(c).value;
        if (v && typeof v === 'object') v = v.result ?? v.text ?? (v.richText ? v.richText.map(t => t.text).join('') : null);
        vals.push(v);
      }
      rows.push(vals);
    });
    return rows;
  }

  // ファイル名「20261012_ピーアーク千葉駅前_台データオンライン.csv」から日付・店名・取得元
  XD.metaFromName = function (name) {
    const base = norm(name).replace(/\.(csv|xlsx?)$/i, '');
    const m = base.match(/^(20\d{2})(\d{2})(\d{2})_(.+?)(?:_([^_]+))?$/);
    if (!m) return { date: null, store: null, source: null };
    return { date: `${m[1]}-${m[2]}-${m[3]}`, store: m[4], source: m[5] || null };
  };

  XD.read = async function (file) {
    const rows = await readRows(file);
    const hi = rows.findIndex(r => r.some(v => /^台番号?$|^台NO$/i.test(norm(v).replace(/[▲▼\s]/g, ''))));
    if (hi < 0) throw new Error('「台番号」の列が見つかりませんでした。ビューアで書き出したファイルか確認してください');
    const head = rows[hi].map(v => norm(v).toUpperCase());
    const col = re => head.findIndex(h => re.test(h));
    const c = {
      no: col(/^台番号?$|^台NO$/),
      name: col(/機種/),
      bb: col(/^BB$|^特賞|^BIG/),
      rb: col(/^RB$|^REG/),
      games: (() => { const t = col(/累計G|総G|総ゲーム/); return t >= 0 ? t : col(/ゲーム数|^G数$/); })(),
    };
    const details = [];
    for (const r of rows.slice(hi + 1)) {
      const no = toNum(r[c.no]);
      if (no == null || !Number.isInteger(no)) continue;
      details.push({
        unitId: no,
        displayName: c.name >= 0 ? norm(r[c.name]) : '',
        model: { name: c.name >= 0 ? norm(r[c.name]) : '' },
        bb: c.bb >= 0 ? toNum(r[c.bb]) : null,
        rb: c.rb >= 0 ? toNum(r[c.rb]) : null,
        games: c.games >= 0 ? toNum(r[c.games]) : null,
      });
    }
    if (!details.length) throw new Error('台のデータが見つかりませんでした');
    return { details, meta: XD.metaFromName(file.name) };
  };

  // ---- 公約チェッカーのサイトに載っている台データオンラインのデータ（このPCのビューアが取得のたびに載せる） ----
  XD.base = 'daidata/';
  XD.index = async function () {
    let r;
    try { r = await fetch(XD.base + 'index.json', { cache: 'no-cache' }); } catch (e) { throw new Error('データの一覧を読み込めませんでした（ネット接続を確認してください）'); }
    if (r.status === 404) return { stores: [] };
    if (!r.ok) throw new Error(`データの一覧を読み込めませんでした（${r.status}）`);
    return r.json();
  };
  XD.load = async function (id, date) {
    let r;
    try { r = await fetch(`${XD.base}${encodeURIComponent(id)}/${date}.json`, { cache: 'no-cache' }); } catch (e) { throw new Error('データを読み込めませんでした（ネット接続を確認してください）'); }
    if (!r.ok) throw new Error('この店・日付のデータがありません');
    const j = await r.json();
    return {
      details: (j.units || []).map(u => ({ unitId: u.no, displayName: u.name, model: { name: u.name }, bb: u.bb, rb: u.rb, games: u.games })),
      meta: { date: j.date, store: j.store, source: j.source },
    };
  };
  // 出力ファイル名の店舗名から店を推定
  XD.guessStore = function (stores, name) {
    const n = norm(name).replace(/\s+/g, '');
    if (!n) return null;
    return stores.find(s => norm(s.name).replace(/\s+/g, '') === n)
      || stores.find(s => { const x = norm(s.name).replace(/\s+/g, ''); return x.includes(n) || n.includes(x); }) || null;
  };

  window.KoyakuExtData = XD;
})();
