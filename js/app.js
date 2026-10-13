(function () {
  const P = window.KoyakuParser, R = window.KoyakuRules, X = window.KoyakuExport, CFG = window.KOYAKU_CONFIG;
  const PAY = window.KOYAKU_PAYOUT6 || { machines: [] };
  const $ = s => document.querySelector(s);
  const esc = s => String(s ?? '').replace(/[&<>"']/g, c => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;' }[c]));
  const fmtNum = n => (n == null ? '' : Math.round(n).toLocaleString('ja-JP'));
  const fmtDiff = n => (n == null ? '' : `<span class="${n >= 0 ? 'plus' : 'minus'}">${n > 0 ? '+' : ''}${fmtNum(n)}</span>`);
  const STATUS = { ok: '達成', ng: '未達', skip: '対象外' };

  // G数。シートにない場合（ホール出力の書式など）はアウト÷3で推定し「約」を付ける
  const gamesOf = m => (m.games != null ? { v: m.games, est: false } : m.out != null ? { v: m.out / 3, est: true } : null);
  const fmtGames = g => (g ? `${g.est ? '約' : ''}${fmtNum(g.v)}G` : '');
  const avg = list => (list.length ? list.reduce((s, x) => s + x, 0) / list.length : null);
  // 元のシートの印（赤＝⑥・黄＝⑤で塗り分けたファイルのときは設定を表示）
  const settingBadge = m => (m.setting === 6 ? '<span class="badge badge--s6">⑥</span>'
    : m.setting === 5 ? '<span class="badge badge--s5">⑤</span>'
      : '<span class="badge badge--orig">⑤⑥</span>');
  function avgLine(units, label) {
    const d = avg(units.map(m => m.diff).filter(v => v != null));
    const gs = units.map(gamesOf).filter(Boolean);
    const g = gs.length ? { v: avg(gs.map(x => x.v)), est: gs.some(x => x.est) } : null;
    if (d == null && !g) return '';
    return `<div class="cgroup__avg">${esc(label)}：差枚 ${d != null ? fmtDiff(d) : '－'} ／ ${g ? fmtGames(g) : '－'}</div>`;
  }

  const store = {
    get(k, d) { try { const v = localStorage.getItem(k); return v ? JSON.parse(v) : d; } catch (e) { return d; } },
    set(k, v) { try { localStorage.setItem(k, JSON.stringify(v)); } catch (e) { /* 保存できなくても動作は続ける */ } },
  };

  const S = {
    parsed: null, ctx: null, buffer: null,
    type: 'kiwami', typeSource: 'auto', // auto: ファイルから判定 / unknown: 判定できず / manual: 手動で選択
    targets: [], supp: new Set(), pins: new Map(), lenient: false,
    result: null, filter: 'marked', search: '',
  };

  // ---------- 114.9%機種 ----------
  const payRules = PAY.machines
    .filter(x => x.payout6 != null)
    .map(x => ({ ...x, core: P.normCore(x.name) }))
    .filter(x => x.core.length >= 2);
  let ov1149 = store.get('koyaku.1149', {});
  const isSubseq = (s, t) => { let i = 0; for (const ch of t) if (ch === s[i]) i++; return i === s.length; };
  const shortest = list => list.reduce((a, b) => (!a || b.core.length < a.core.length ? b : a), null);
  const payMatch = name => {
    const h = P.normCore(name);
    // 1) ホールの機種名にDMMの機種名が含まれる（長く一致したものを優先。旧作と新作の取り違え防止）
    let best = null;
    for (const x of payRules) if (h.includes(x.core) && (!best || x.core.length > best.core.length)) best = x;
    if (best) return best;
    // 2) ホール側が途中で切れている（末尾の型式コード「L8」「CT」などは外す）
    const t = h.replace(/[A-Z][A-Z0-9]{0,4}$/, '');
    if (t.length >= 4) {
      best = shortest(payRules.filter(x => x.core.includes(t)));
      if (best) return best;
    }
    // 3) 略称（「バジリスク絆2天膳BLACK」など、文字が順番どおりに含まれる）
    if (t.length >= 6) best = shortest(payRules.filter(x => x.core.startsWith(t.slice(0, 2)) && isSubseq(t, x.core)));
    return best;
  };
  const payCache = new Map();
  const payInfo = g => {
    if (!payCache.has(g.name)) payCache.set(g.name, payMatch(g.name));
    return payCache.get(g.name);
  };
  const payoutOf = g => (payInfo(g) ? payInfo(g).payout6 : null);
  const ovKey = g => P.normMatch(g.name);
  const is1149 = g => {
    const k = ovKey(g);
    if (k in ov1149) return ov1149[k];
    const p = payoutOf(g);
    return p != null && p >= CFG.payout6Threshold;
  };

  // ---------- 読み込み ----------
  // 古い形式（.xls）はSheetJSで読み、ExcelJSのブックに移してから同じ処理にかける
  const SHEETJS_URL = 'https://cdnjs.cloudflare.com/ajax/libs/xlsx/0.18.5/xlsx.full.min.js';
  function loadSheetJS() {
    if (window.XLSX) return Promise.resolve();
    return new Promise((resolve, reject) => {
      const s = document.createElement('script');
      s.src = SHEETJS_URL;
      s.onload = resolve;
      s.onerror = () => reject(new Error('.xls を読むための部品を読み込めませんでした（ネット接続を確認してください）'));
      document.head.appendChild(s);
    });
  }

  async function xlsToWorkbook(buffer) {
    await loadSheetJS();
    const src = XLSX.read(buffer, { type: 'array', cellStyles: true });
    const wb = new ExcelJS.Workbook();
    src.SheetNames.forEach((name, i) => {
      const sws = src.Sheets[name];
      const ws = wb.addWorksheet(name.replace(/[\[\]:*?\/\\]/g, '_').slice(0, 31) || `Sheet${i + 1}`);
      const info = src.Workbook && src.Workbook.Sheets && src.Workbook.Sheets[i];
      if (info && info.Hidden) ws.state = 'hidden';
      if (!sws || !sws['!ref']) return;
      const range = XLSX.utils.decode_range(sws['!ref']);
      for (let r = range.s.r; r <= range.e.r; r++) {
        for (let c = range.s.c; c <= range.e.c; c++) {
          const cell = sws[XLSX.utils.encode_cell({ r, c })];
          if (!cell) continue;
          const x = ws.getCell(r + 1, c + 1);
          if (cell.v != null) x.value = cell.t === 'd' ? new Date(cell.v) : cell.v;
          const s = cell.s;
          if (s && s.patternType && s.patternType !== 'none' && s.fgColor && s.fgColor.rgb) {
            x.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF' + String(s.fgColor.rgb).slice(-6) } };
          }
        }
      }
      for (const m of sws['!merges'] || []) {
        try { ws.mergeCells(m.s.r + 1, m.s.c + 1, m.e.r + 1, m.e.c + 1); } catch (e) { /* 重なる結合は無視 */ }
      }
    });
    return wb;
  }

  async function loadBuffer(buffer, fileName) {
    const head = new Uint8Array(buffer, 0, 2);
    const isZip = head[0] === 0x50 && head[1] === 0x4B; // .xlsx は "PK" で始まる
    let wb;
    if (isZip) {
      wb = new ExcelJS.Workbook();
      await wb.xlsx.load(buffer.slice(0));
    } else {
      wb = await xlsToWorkbook(buffer);
    }
    const parsed = P.parseWorkbook(wb, fileName);
    S.parsed = parsed;
    S.buffer = buffer;
    S.ctx = R.buildCtx(parsed.machines);
    S.supp = new Set();
    S.pins = new Map();
    S.type = parsed.typeGuess || 'kiwami';
    S.typeSource = parsed.typeGuess ? 'auto' : 'unknown';
    S.lenient = false;
    resolveTargets();
    $('#out-date').value = guessDate(parsed);
    $('#out-store').value = guessStore(parsed);
    resetPision();
    $('#drop').hidden = true;
    $('#app').hidden = false;
    $('#btn-reset').hidden = false;
    renderSettings();
    run();
    return summaryText();
  }

  // ---------- 出力ファイル名（1010_店名_極.xlsx） ----------
  const pad2 = n => String(n).padStart(2, '0');
  const validMD = (m, d) => m >= 1 && m <= 12 && d >= 1 && d <= 31;

  // ファイル名やシート上部（抽出条件の年月日など）から日付を推定。見つからなければ今日
  function guessDate(p) {
    const today = new Date();
    const nfkc = t => (t ? t.normalize('NFKC') : '');
    const names = [p.fileName.replace(/\.xls[xm]?$/i, ''), p.dataSheet, p.marksSheet].map(nfkc).filter(Boolean);
    const all = [nfkc(p.meta.topText), ...names];
    let y = null, md = null, m;
    // 1) 「20261005」のような年月日（シート上部の抽出条件など）
    for (const t of all) {
      if ((m = t.match(/(20\d{2})(\d{2})(\d{2})/)) && validMD(+m[2], +m[3])) { y = +m[1]; md = [+m[2], +m[3]]; break; }
    }
    // 2) 「9月30日」
    if (!md) for (const t of all) if ((m = t.match(/(\d{1,2})月(\d{1,2})日/)) && validMD(+m[1], +m[2])) { md = [+m[1], +m[2]]; break; }
    // 3) ファイル名・シート名の「10.5」「1011」（データの数字と取り違えないよう、上部の文字は見ない）
    if (!md) {
      for (const t of names) {
        if ((m = t.match(/(?:^|[^\d])(\d{1,2})[.\/](\d{1,2})(?!\d)/)) && validMD(+m[1], +m[2])) { md = [+m[1], +m[2]]; break; }
        if ((m = t.match(/(?:^|[^\d])(\d{2})(\d{2})(?!\d)/)) && validMD(+m[1], +m[2])) { md = [+m[1], +m[2]]; break; }
      }
    }
    if (!md) return `${today.getFullYear()}-${pad2(today.getMonth() + 1)}-${pad2(today.getDate())}`;
    if (!y) {
      // 年がない場合は今年。2か月以上先の日付になるなら去年とみなす
      y = today.getFullYear();
      if (new Date(y, md[0] - 1, md[1]) - today > 60 * 864e5) y--;
    }
    return `${y}-${pad2(md[0])}-${pad2(md[1])}`;
  }

  // 「部店名」の列か、ファイル名から日付・取材名などを除いた残りを店舗名の初期値にする
  function guessStore(p) {
    if (p.meta.store) return p.meta.store;
    return p.fileName.normalize('NFKC')
      .replace(/\.xls[xm]?$/i, '')
      .replace(/_?(補填済|チェック結果|チェック済)(\s*\(\d+\))?$/, '')
      .replace(/S級ホール調査|調査結果報告シート|結果報告シート|結果報告|ホール調査|\d+月以降/g, '')
      .replace(/[（(][^）)]*[）)]/g, '')
      .replace(/\d{1,2}月\d{1,2}日|20\d{6}|\d{1,2}[.\/]\d{1,2}|\b\d{3,4}\b|^\d{3,4}/g, '')
      .replace(/データ|S級|極|虹|紫|赤|青/g, '')
      .replace(/[\s_・\-]+/g, ' ')
      .trim();
  }

  function outBaseName() {
    const d = $('#out-date').value; // yyyy-mm-dd
    const mmdd = d ? d.slice(5, 7) + d.slice(8, 10) : '';
    const store = $('#out-store').value.trim().replace(/[\\/:*?"<>|]/g, '').replace(/\s+/g, '');
    return [mmdd, store, R.TYPES[S.type].label].filter(Boolean).join('_');
  }

  function renderOutName() {
    $('#out-preview').textContent = outBaseName() + '.xlsx';
    const d = $('#out-date').value;
    $('#pision-date').textContent = d ? `${+d.slice(5, 7)}/${+d.slice(8, 10)}` : '未入力';
  }

  // ---------- pision連携 ----------
  const PS = window.KoyakuPision;
  let psHalls = PS.cachedHalls();

  function renderPisionKey() {
    const has = !!PS.getKey();
    $('#pision-key').value = '';
    $('#pision-key').placeholder = has ? '保存済み（変えるときだけ入力）' : 'pisionのAPIキー';
    $('#pision-key-state').textContent = has ? '保存済み' : '未設定';
    $('#pision-key-state').classList.toggle('is-ok', has);
    $('#pision-key-del').hidden = !has;
  }
  function fillHallList() {
    if (psHalls) $('#pision-halls').innerHTML = psHalls.map(h => `<option value="${esc(PS.hallLabel(h))}"></option>`).join('');
  }
  async function ensureHalls() {
    if (!psHalls) { psHalls = await PS.halls(); fillHallList(); }
    return psHalls;
  }
  // 店舗の初期値は「出力ファイル名」の店舗名から
  function prefillHall() {
    if (!psHalls || $('#pision-hall').value) return;
    const h = PS.guessHall(psHalls, $('#out-store').value);
    if (h) $('#pision-hall').value = PS.hallLabel(h);
  }
  function pisionMessage(text, isError) {
    const el = $('#pision-result');
    el.hidden = !text;
    el.textContent = text || '';
    el.classList.toggle('is-error', !!isError);
  }
  function resetPision() {
    S.pision = null;
    $('#pision-hall').value = '';
    $('#pision-undo').hidden = true;
    $('#pision-badge').hidden = true;
    pisionMessage('');
    prefillHall();
  }

  async function fetchPision() {
    const btn = $('#pision-fetch');
    btn.disabled = true;
    pisionMessage('pisionからデータを取得しています…');
    try {
      const halls = await ensureHalls();
      const text = $('#pision-hall').value;
      const hall = PS.findHall(halls, text) || PS.guessHall(halls, text || $('#out-store').value);
      if (!hall) throw new Error('店舗を一覧から選んでください（店名を入力すると候補が出ます）');
      $('#pision-hall').value = PS.hallLabel(hall);
      const date = $('#out-date').value;
      if (!date) throw new Error('「出力ファイル名」の日付を入れてください');
      const details = await PS.results(hall.id, date);
      if (S.pision) PS.undo(S.parsed.machines, S.ctx.models); // 取り直すときは元の値から
      const st = PS.apply(S.parsed.machines, S.ctx.models, details);
      S.pision = { hall, date, st };
      payCache.clear();
      renderSettings();
      run();
      const parts = [
        `${hall.name}・${+date.slice(5, 7)}/${+date.slice(8, 10)} のpisionデータと照合しました（${st.matched}/${st.total}台が一致）。`,
        `BB ${st.bb}台・RB ${st.rb}台・G数 ${st.games}台を補い、機種名を${st.renamed}機種置き換えました。`,
      ];
      if (st.missing.length) parts.push(`pisionに無かった台番号：${st.missing.slice(0, 12).join('、')}${st.missing.length > 12 ? ` ほか${st.missing.length - 12}台` : ''}`);
      if (st.matched < st.total * 0.5) parts.push('※一致した台が少ないので、店舗と日付が合っているか確認してください。');
      pisionMessage(parts.join(' '), st.matched < st.total * 0.5);
      $('#pision-undo').hidden = false;
      $('#pision-badge').hidden = false;
      $('#pision-badge').textContent = '取得済み';
    } catch (e) {
      pisionMessage(e.message || String(e), true);
    } finally {
      btn.disabled = false;
    }
  }

  function undoPision() {
    PS.undo(S.parsed.machines, S.ctx.models);
    S.pision = null;
    payCache.clear();
    renderSettings();
    run();
    $('#pision-undo').hidden = true;
    $('#pision-badge').hidden = true;
    pisionMessage('pisionで補った値と機種名を元に戻しました。');
  }

  function resolveTargets() {
    S.targets = CFG.targets.map(t => {
      const re = new RegExp(t.pattern);
      const hits = S.ctx.models.filter(g => re.test(P.normMatch(g.name))).sort((a, b) => b.n - a.n);
      return { label: t.label, model: hits[0] || null };
    });
  }

  function run() {
    const isMarked = m => m.orig || S.supp.has(m.i);
    const isSupp = m => S.supp.has(m.i);
    S.result = R.solve(S.ctx, { type: S.type, isMarked, isSupp, targets: S.targets, lenient: S.lenient, is1149 });
    render();
  }

  // ---------- 描画 ----------
  function render() {
    const p = S.parsed, res = S.result;
    const origCnt = p.machines.filter(m => m.orig).length;
    $('#info-name').textContent = p.fileName;
    $('#info-meta').innerHTML = [
      `書式 <b>${esc(p.formatLabel)}</b>`,
      `シート <b>${esc(p.dataSheet)}</b>`,
      `設置 <b>${S.ctx.N}台</b>（${S.ctx.models.length}機種）`,
      p.settingColors
        ? `⑥ <b>${p.machines.filter(m => m.orig && m.setting === 6).length}台</b>・⑤ <b>${p.machines.filter(m => m.orig && m.setting === 5).length}台</b>（赤＝⑥・黄＝⑤）${S.supp.size ? `＋補填 <b>${S.supp.size}台</b>` : ''}`
        : `⑤⑥ <b>${origCnt}台</b>${S.supp.size ? `＋補填 <b>${S.supp.size}台</b>` : ''}`,
    ].map(s => `<span>${s}</span>`).join('');
    document.querySelectorAll('.type').forEach(b => b.setAttribute('aria-checked', String(b.dataset.type === S.type)));
    const hint = $('#type-hint');
    hint.textContent = {
      auto: '（ファイル名から自動判定・ボタンで切り替えできます）',
      unknown: '（自動判定できませんでした。選んでください）',
      manual: '（手動で選択）',
    }[S.typeSource];
    hint.classList.toggle('is-warn', S.typeSource === 'unknown');

    const warns = [];
    if (S.typeSource === 'unknown') warns.push('取材の種類をファイルから判定できませんでした。「取材の種類」で正しい種類を選んでください。');
    if (!origCnt) warns.push('仕掛け台（⭕️や色塗り）が見つかりませんでした。シートの書式を確認してください。');
    for (const w of p.warnings.slice(0, 5)) warns.push(w);
    if (p.warnings.length > 5) warns.push(`ほか ${p.warnings.length - 5} 件の注意があります`);
    $('#warnings').innerHTML = warns.map(w => `<div class="notice">${esc(w)}</div>`).join('');

    const judged = res.reqs.filter(r => r.status !== 'skip');
    const ok = judged.filter(r => r.status === 'ok').length;
    const allOk = ok === judged.length;
    $('#summary-score').innerHTML =
      `<span class="summary__big ${allOk ? 'is-ok' : 'is-ng'}">${allOk ? 'すべて達成' : `未達 ${judged.length - ok}件`}</span>` +
      `<span class="summary__sub">${R.TYPES[S.type].label}取材 ・ 達成 ${ok} / ${judged.length}</span>`;
    $('#btn-download').textContent = S.supp.size ? `補填版Excelをダウンロード（${S.supp.size}台）` : 'チェック結果をExcelでダウンロード';
    renderOutName();

    renderSupp();
    $('#reqs').innerHTML = res.reqs.map(renderReq).join('');
    renderTable();
  }

  function renderSupp() {
    const el = $('#supp');
    const list = S.parsed.machines.filter(m => S.supp.has(m.i));
    el.hidden = !list.length;
    if (!list.length) { el.innerHTML = ''; return; }
    const tagOf = m => R.TAGS[S.result.assign.get(m) || 'other'].short;
    el.innerHTML = `<div class="supp__head">補填した台（${list.length}台）</div><div class="chips">` +
      list.map(m => `<span class="chip">${m.no} ${esc(m.name)} <small>${esc(tagOf(m))}</small><button title="補填を取り消す" data-unsupp="${m.i}">×</button></span>`).join('') +
      '</div>';
  }

  function machineRow(m) {
    const orig = m.orig, supp = S.supp.has(m.i);
    return `<label class="mrow${orig ? ' is-orig' : ''}${supp ? ' is-supp' : ''}">` +
      `<input type="checkbox" data-i="${m.i}"${orig || supp ? ' checked' : ''}${orig ? ' disabled' : ''}>` +
      `<span class="mrow__no">${m.no}</span>` +
      `<span class="mrow__name" title="${esc(m.name)}">${esc(m.name)}${orig ? ' ' + settingBadge(m) : ''}</span>` +
      `<span class="mrow__diff">${fmtDiff(m.diff)}</span>` +
      `<span class="mrow__g">${fmtGames(gamesOf(m))}</span>` +
      `</label>`;
  }

  // チェックを入れた候補グループは、判定し直しても消えないように固定して表示し続ける
  function pinGroup(reqId, key) {
    const id = `${reqId}|${key}`;
    if (S.pins.has(id)) return;
    const r = S.result.reqs.find(x => x.id === reqId);
    const g = r && (r.groups || []).find(x => x.key === key);
    if (g) S.pins.set(id, { reqId, key, group: { ...g, machines: [...g.machines] } });
  }

  function groupsFor(r) {
    const current = r.status === 'ng' ? (r.groups || []) : [];
    const pins = [...S.pins.values()].filter(p => p.reqId === r.id);
    const pinned = pins.map(p => {
      const cur = current.find(g => g.key === p.key);
      // チェックした台は「未選択の台」から外れるので、固定したときの台と合わせて並べる
      const machines = [...new Set([...p.group.machines, ...(cur ? cur.machines : [])])];
      p.group = { ...(cur || p.group), machines };
      return cur ? { ...cur, machines, pinned: true } : { ...p.group, needText: '補填済み', done: true, pinned: true };
    });
    return [...pinned, ...current.filter(g => !pins.some(p => p.key === g.key))];
  }

  function renderReq(r) {
    const lines = r.lines.map(l => `<li class="${l.s}">${esc(l.t)}</li>`).join('');
    let cands = '';
    const groups = groupsFor(r);
    if (r.status === 'ng' || groups.length) {
      const head = r.status === 'ng'
        ? '補填候補<small>チェックを入れた台が補填されます（差枚の良い順）</small>'
        : '補填した台<small>チェックを外すと取り消せます</small>';
      cands = `<div class="cands"><h4 class="cands__title">${head}</h4>` +
        (groups.length
          ? `<div class="cgroups">${groups.map(g => {
            const selectable = g.machines.filter(m => !m.orig && !S.supp.has(m.i));
            return `<div class="cgroup${g.pinned ? ' is-pinned' : ''}" data-req="${esc(r.id)}" data-key="${esc(g.key)}"><div class="cgroup__head"><span class="cgroup__title">${esc(g.title)}</span>` +
              `<span class="need${g.done ? ' need--done' : ''}">${esc(g.needText)}</span>` +
              (g.selectAll && selectable.length ? `<button class="btn btn--sm" data-sel="${selectable.map(m => m.i).join(',')}">まとめて選ぶ</button>` : '') +
              `</div>` +
              (g.model ? avgLine(g.model.units, `機種平均（${g.model.n}台）`) : g.avgUnits ? avgLine(g.avgUnits, `${g.avgUnits.length}台の平均`) : '') +
              (g.note ? `<div class="cgroup__note">${esc(g.note)}</div>` : '') +
              (g.machines.length ? `<div class="mlist">${g.machines.map(machineRow).join('')}</div>` : '') +
              `</div>`;
          }).join('')}</div>`
          : `<div class="cgroup__note">補填できる候補が見つかりませんでした。</div>`) +
        `</div>`;
    }
    return `<article class="req req--${r.status}">` +
      `<div class="req__head"><span class="req__no">${esc(r.no)}</span><h3 class="req__title">${esc(r.title)}</h3>` +
      `<span class="pill pill--${r.status}">${STATUS[r.status]}</span></div>` +
      `<ul class="req__lines">${lines}</ul>` +
      (r.warn ? `<div class="req__warn">${esc(r.warn)}</div>` : '') +
      cands + `</article>`;
  }

  function renderTable() {
    const assign = S.result.assign;
    const q = S.search.trim().normalize('NFKC').toUpperCase();
    const rows = S.parsed.machines.filter(m => {
      if (S.filter === 'marked' && !(m.orig || S.supp.has(m.i))) return false;
      if (q && !String(m.no).includes(q) && !P.normMatch(m.name).includes(P.normMatch(q))) return false;
      return true;
    });
    $('#tbl').innerHTML =
      '<thead><tr><th>⑤⑥</th><th>台番号</th><th>機種名</th><th>台数</th><th>数えた公約</th><th>差枚</th><th>G数</th><th>シート上の区分</th></tr></thead><tbody>' +
      rows.map(m => {
        const supp = S.supp.has(m.i), on = m.orig || supp;
        const tag = on ? R.TAGS[assign.get(m) || 'other'].short : '';
        return `<tr class="${supp ? 'is-supp' : ''}">` +
          `<td><input type="checkbox" data-i="${m.i}"${on ? ' checked' : ''}${m.orig ? ' disabled' : ''}></td>` +
          `<td class="num">${m.no}</td><td class="name">${esc(m.name)}</td><td class="num">${m.model.n}</td>` +
          `<td>${m.orig ? settingBadge(m) + ' ' : ''}${tag ? `<span class="tag">${esc(tag)}</span>` : ''}${supp ? ' <span class="badge badge--supp">補填</span>' : ''}</td>` +
          `<td class="num">${fmtDiff(m.diff)}</td><td class="num">${fmtGames(gamesOf(m))}</td>` +
          `<td>${esc((m.declared || '').replace(/\s+/g, ' ').slice(0, 24))}</td></tr>`;
      }).join('') + '</tbody>';
  }

  function renderSettings() {
    const models = [...S.ctx.models].sort((a, b) => a.name.localeCompare(b.name, 'ja'));
    const opts = sel => '<option value="">（未設置）</option>' +
      models.map(g => `<option value="${esc(g.key)}"${sel === g ? ' selected' : ''}>${esc(g.name)}（${g.n}台）</option>`).join('');
    const big = models.filter(g => g.n >= 3).sort((a, b) => (is1149(b) - is1149(a)) || a.name.localeCompare(b.name, 'ja'));
    $('#settings-body').innerHTML =
      `<div class="set-sec"><h3>①の対象5機種</h3><p>機種名から自動で選んでいます。違っていたら選び直してください。</p><div class="tgt-grid">` +
      S.targets.map((t, i) => `<label class="tgt"><b>${esc(t.label)}</b><select data-target="${i}">${opts(t.model)}</select></label>`).join('') +
      `</div></div>` +
      `<div class="set-sec"><h3>114.9%機種（紫②で使用）</h3><p>3台以上設置の機種です。設定6の機械割が${CFG.payout6Threshold}%以上の機種に自動でチェックしています${PAY.updated ? `（${esc(PAY.source)}・${esc(PAY.updated)}時点）` : ''}。変更はこのブラウザに保存されます。</p><div class="chk-grid">` +
      big.map(g => {
        const info = payInfo(g), on = is1149(g);
        return `<label class="chk${on ? ' is-on' : ''}"><input type="checkbox" data-1149="${esc(g.key)}"${on ? ' checked' : ''}>${esc(g.name)}（${g.n}台）` +
          (info ? ` <small title="DMM：${esc(info.name)}">⑥ ${info.payout6}%</small>` : ' <small>DMMに該当なし</small>') + `</label>`;
      }).join('') +
      `</div></div>` +
      `<div class="set-sec"><h3>仕掛けの被り</h3><p>初期設定では1台を複数の公約に数えません（極②は除く）。青は「①を含めず」のため常に別扱いです。</p>` +
      `<label class="chk"><input type="checkbox" id="lenient"${S.lenient ? ' checked' : ''}>①の「⑥あり」は、全台系など他の公約の台と兼用してよい</label></div>`;
  }

  // ---------- 操作 ----------
  function toggleSupp(i, on) {
    if (on) S.supp.add(i); else S.supp.delete(i);
    // 補填した台が1台も残っていない固定グループは外す
    for (const [id, p] of S.pins) if (!p.group.machines.some(m => S.supp.has(m.i))) S.pins.delete(id);
    run();
  }

  document.addEventListener('change', e => {
    const t = e.target;
    if (t.matches('input[data-i]')) {
      const grp = t.closest('.cgroup');
      if (grp && t.checked) pinGroup(grp.dataset.req, grp.dataset.key);
      return toggleSupp(Number(t.dataset.i), t.checked);
    }
    if (t.matches('select[data-target]')) {
      const key = t.value;
      S.targets[Number(t.dataset.target)].model = S.ctx.models.find(g => g.key === key) || null;
      return run();
    }
    if (t.matches('input[data-1149]')) {
      const g = S.ctx.models.find(x => x.key === t.dataset['1149']);
      ov1149[ovKey(g)] = t.checked;
      store.set('koyaku.1149', ov1149);
      t.closest('.chk').classList.toggle('is-on', t.checked);
      return run();
    }
    if (t.id === 'lenient') { S.lenient = t.checked; return run(); }
  });

  document.addEventListener('click', e => {
    const b = e.target.closest('button');
    if (!b) return;
    if (b.dataset.type) {
      if (b.dataset.type !== S.type || S.typeSource === 'unknown') {
        S.type = b.dataset.type;
        S.typeSource = 'manual';
        S.pins.clear();
        run();
        toast(`${R.TYPES[S.type].label}取材の公約で判定し直しました`);
      }
      return;
    }
    if (b.dataset.sel) {
      const grp = b.closest('.cgroup');
      if (grp) pinGroup(grp.dataset.req, grp.dataset.key);
      b.dataset.sel.split(',').forEach(i => S.supp.add(Number(i)));
      return run();
    }
    if (b.dataset.unsupp) return toggleSupp(Number(b.dataset.unsupp), false);
    if (b.dataset.filter) {
      S.filter = b.dataset.filter;
      document.querySelectorAll('[data-filter]').forEach(x => x.classList.toggle('is-on', x === b));
      return renderTable();
    }
  });
  $('#tbl-search').addEventListener('input', e => { S.search = e.target.value; renderTable(); });
  $('#out-date').addEventListener('input', renderOutName);
  $('#out-store').addEventListener('input', renderOutName);

  renderPisionKey();
  fillHallList();
  $('#pision-key-save').addEventListener('click', async () => {
    const v = $('#pision-key').value.trim();
    if (!v) return pisionMessage('APIキーを入力してから「保存」を押してください', true);
    PS.setKey(v);
    renderPisionKey();
    pisionMessage('');
    try { await ensureHalls(); prefillHall(); } catch (e) { pisionMessage(e.message, true); }
  });
  $('#pision-key-del').addEventListener('click', () => { PS.setKey(''); renderPisionKey(); });
  $('#pision').addEventListener('toggle', async () => {
    if (!$('#pision').open || !PS.getKey()) return;
    try { await ensureHalls(); prefillHall(); } catch (e) { pisionMessage(e.message, true); }
  });
  $('#pision-fetch').addEventListener('click', fetchPision);
  $('#pision-undo').addEventListener('click', undoPision);

  // ---------- ファイル ----------
  const drop = $('#drop'), input = $('#file');
  async function handleFile(file) {
    if (!file) return;
    const err = $('#drop-error');
    err.hidden = true;
    try {
      await loadBuffer(await file.arrayBuffer(), file.name);
    } catch (ex) {
      console.error(ex);
      $('#app').hidden = true;
      drop.hidden = false;
      err.textContent = `読み込めませんでした：${ex.message || ex}`;
      err.hidden = false;
    }
  }
  drop.addEventListener('click', () => input.click());
  drop.addEventListener('keydown', e => { if (e.key === 'Enter' || e.key === ' ') { e.preventDefault(); input.click(); } });
  input.addEventListener('change', () => { handleFile(input.files[0]); input.value = ''; });
  ['dragenter', 'dragover'].forEach(ev => document.addEventListener(ev, e => { e.preventDefault(); drop.classList.add('is-over'); }));
  ['dragleave', 'drop'].forEach(ev => document.addEventListener(ev, e => { e.preventDefault(); if (ev === 'drop' || e.target === drop) drop.classList.remove('is-over'); }));
  document.addEventListener('drop', e => handleFile(e.dataTransfer.files[0]));
  $('#btn-reset').addEventListener('click', () => input.click());

  // ---------- 出力 ----------
  function toast(msg) {
    const t = $('#toast');
    t.textContent = msg;
    t.hidden = false;
    clearTimeout(toast.timer);
    toast.timer = setTimeout(() => { t.hidden = true; }, 2200);
  }

  function summaryText() {
    const res = S.result, p = S.parsed;
    const lines = [`【公約チェック】${outBaseName()} 設置${S.ctx.N}台`];
    for (const r of res.reqs) {
      const mark = { ok: '✅', ng: '❌', skip: '➖' }[r.status];
      lines.push(`${mark} ${r.no} ${r.title}`);
      if (r.status === 'ng') r.lines.filter(l => l.s === 'ng').slice(0, 4).forEach(l => lines.push(`　・${l.t}`));
    }
    const supp = p.machines.filter(m => S.supp.has(m.i));
    if (supp.length) {
      lines.push(`補填 ${supp.length}台：`);
      supp.forEach(m => lines.push(`　${m.no} ${m.name}（${R.TAGS[res.assign.get(m) || 'other'].short}）`));
    }
    return lines.join('\n');
  }

  $('#btn-copy').addEventListener('click', async () => {
    try { await navigator.clipboard.writeText(summaryText()); toast('結果をコピーしました'); }
    catch (e) { toast('コピーできませんでした'); }
  });

  $('#btn-download').addEventListener('click', async () => {
    const btn = $('#btn-download');
    btn.disabled = true;
    try {
      const blob = await X.build(S);
      const a = document.createElement('a');
      a.href = URL.createObjectURL(blob);
      a.download = outBaseName() + '.xlsx';
      document.body.appendChild(a);
      a.click();
      setTimeout(() => { URL.revokeObjectURL(a.href); a.remove(); }, 1000);
    } catch (ex) {
      console.error(ex);
      toast(`Excelを作れませんでした：${ex.message || ex}`);
    } finally {
      btn.disabled = false;
    }
  });

  // 動作確認用
  window.KoyakuApp = { loadBuffer, state: S, summaryText, payMatch };
})();
