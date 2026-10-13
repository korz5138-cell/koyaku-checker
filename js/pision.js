/* pision.io との連携。
 * 台番号でpisionのデータと照合し、シートで空欄のBB・RB・G数を埋め、機種名をpisionの正式名に置き換える。
 * APIキーは各自のブラウザ（localStorage）にだけ保存する。公開サイトやファイルには入れない。
 */
(function () {
  const PS = {};
  const BASE = 'https://www.pision.io';
  const KEY_STORE = 'koyaku.pision.key';
  const HALLS_STORE = 'koyaku.pision.halls';
  const HALLS_TTL = 24 * 3600 * 1000;

  const PREFS = ['', '北海道', '青森県', '岩手県', '宮城県', '秋田県', '山形県', '福島県', '茨城県', '栃木県', '群馬県',
    '埼玉県', '千葉県', '東京都', '神奈川県', '新潟県', '富山県', '石川県', '福井県', '山梨県', '長野県', '岐阜県',
    '静岡県', '愛知県', '三重県', '滋賀県', '京都府', '大阪府', '兵庫県', '奈良県', '和歌山県', '鳥取県', '島根県',
    '岡山県', '広島県', '山口県', '徳島県', '香川県', '愛媛県', '高知県', '福岡県', '佐賀県', '長崎県', '熊本県',
    '大分県', '宮崎県', '鹿児島県', '沖縄県'];

  const ls = {
    get(k) { try { return localStorage.getItem(k); } catch (e) { return null; } },
    set(k, v) { try { localStorage.setItem(k, v); } catch (e) { /* 保存できなくても続ける */ } },
    del(k) { try { localStorage.removeItem(k); } catch (e) { /* 同上 */ } },
  };

  PS.getKey = () => (ls.get(KEY_STORE) || '').trim();
  PS.setKey = k => (k ? ls.set(KEY_STORE, k.trim()) : ls.del(KEY_STORE));

  // 通信（テストのときは差し替えられるようにしておく）
  PS.request = async function (path) {
    const key = PS.getKey();
    if (!key) throw new Error('pisionのAPIキーを入力して「保存」を押してください');
    let r;
    try {
      r = await fetch(BASE + path, { headers: { 'X-Api-Key': key, Accept: 'application/json' } });
    } catch (e) {
      throw new Error('pisionに接続できませんでした（ネット接続を確認してください）');
    }
    if (r.status === 401 || r.status === 403) throw new Error('pisionのAPIキーが正しくありません');
    if (r.status === 404) throw new Error('この店・日付のデータがpisionにありません（非公開の店か、まだデータが作られていない日です）');
    if (!r.ok) throw new Error(`pisionからデータを取れませんでした（${r.status}）`);
    return r.json();
  };

  PS.prefName = id => PREFS[id] || '';
  PS.hallLabel = h => `${h.name}（${PS.prefName(h.prefectureId)}）#${h.id}`;

  // 店舗一覧（1日1回だけ取り直す）
  PS.cachedHalls = function () {
    try {
      const c = JSON.parse(ls.get(HALLS_STORE) || 'null');
      return c && Date.now() - c.at < HALLS_TTL ? c.halls : null;
    } catch (e) { return null; }
  };
  PS.halls = async function () {
    const cached = PS.cachedHalls();
    if (cached) return cached;
    const j = await PS.request('/api/v2/halls');
    const halls = (j.halls || []).map(h => ({ id: h.id, name: h.name, prefectureId: h.prefectureId }));
    ls.set(HALLS_STORE, JSON.stringify({ at: Date.now(), halls }));
    return halls;
  };

  PS.results = async function (hallId, date) {
    const j = await PS.request(`/api/v2/halls/${encodeURIComponent(hallId)}/results/${date}`);
    return j.details || [];
  };

  // 店名から店を推定（完全一致 → 含む／含まれる の順。表記ゆれは全角半角・空白を吸収）
  const norm = s => String(s || '').normalize('NFKC').replace(/\s+/g, '').toUpperCase();
  PS.guessHall = function (halls, store) {
    const s = norm(store);
    if (!s || !halls) return null;
    const exact = halls.find(h => norm(h.name) === s);
    if (exact) return exact;
    let best = null, bestScore = 0;
    for (const h of halls) {
      const n = norm(h.name);
      const score = n.includes(s) ? s.length / n.length : s.includes(n) ? n.length / s.length : 0;
      if (score > bestScore) { best = h; bestScore = score; }
    }
    return bestScore >= 0.5 ? best : null;
  };
  PS.findHall = function (halls, text) {
    const m = String(text || '').match(/#(\d+)\s*$/);
    if (m) return halls.find(h => String(h.id) === m[1]) || null;
    const s = norm(text);
    return halls.find(h => norm(h.name) === s) || null;
  };

  // 台番号で照合して反映。元の値は戻せるように残す
  PS.apply = function (machines, models, details) {
    const byNo = new Map(details.map(d => [Number(d.unitId), d]));
    const stat = { total: machines.length, matched: 0, bb: 0, rb: 0, games: 0, renamed: 0, missing: [], extra: 0 };
    const renamedModels = new Set();
    for (const m of machines) {
      const d = byNo.get(m.no);
      if (!d) { stat.missing.push(m.no); continue; }
      stat.matched++;
      if (!m.pisionBackup) m.pisionBackup = { name: m.name, bb: m.bb, rb: m.rb, games: m.games };
      if (m.bb == null && d.bb != null) { m.bb = d.bb; stat.bb++; }
      if (m.rb == null && d.rb != null) { m.rb = d.rb; stat.rb++; }
      if (m.games == null && d.games != null) { m.games = d.games; stat.games++; }
      const name = ((d.model && d.model.name) || d.displayName || '').trim();
      if (name && name !== m.name) { m.name = name; renamedModels.add(m.model); }
    }
    for (const g of renamedModels) g.name = g.units[0].name;
    stat.renamed = renamedModels.size;
    stat.extra = details.length - stat.matched;
    return stat;
  };

  PS.undo = function (machines, models) {
    for (const m of machines) {
      if (!m.pisionBackup) continue;
      Object.assign(m, m.pisionBackup);
      delete m.pisionBackup;
    }
    for (const g of models) g.name = g.units[0].name;
  };

  window.KoyakuPision = PS;
})();
