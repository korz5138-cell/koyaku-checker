/* 公約の判定ロジック。
 * ⭕️（色塗り）の台は⑤⑥、①の対象機種に数える台は⑥とみなす。
 * 1台を複数の公約に数えない（被りなし）。ただし極②（4台以上機種に1台以上）は他の公約の台でも満たせる。
 * 「どの台をどの公約に数えるか」は組み合わせを試して、達成数が最も多くなる割り振りを採用する。
 */
(function () {
  const R = {};
  const ceilDiv = (a, b) => Math.floor((a + b - 1) / b);
  R.ceilDiv = ceilDiv;

  R.TYPES = {
    kiwami: { label: '極' },
    murasaki: { label: '紫' },
    aka: { label: '赤' },
    ao: { label: '青' },
  };

  // 区分（どの公約に数えたか）と、書式Aで⭕️を付ける列の見出しキーワード
  R.TAGS = {
    t6: { short: '①⑥', col: /喰種|各機種|6あり/ },
    thalf: { short: '①1/2', col: /喰種|各機種|6あり/ },
    k2: { short: '②', col: /4台以上/ },
    k3: { short: '③全台', col: /全台/ },
    k4: { short: '④並び', col: /並び/ },
    a2: { short: '②全台系', col: /全台|4%/ },
    a3: { short: '③1/2', col: /1\/2機種|^1\/2$/ },
    a4: { short: '④2台以下', col: /2台設置以下|2台以下/ },
    m2: { short: '②114.9%全⑥', col: /114/ },
    m4: { short: '④2台以下並び', col: /2台設置以下|2台以下/ },
    o2: { short: '②並び', col: /並び/ },
    other: { short: 'その他', col: /その他|単品|他公約/ },
  };

  R.buildCtx = function (machines) {
    const map = new Map();
    for (const m of machines) {
      let g = map.get(m.key);
      if (!g) { g = { key: m.key, name: m.name, units: [] }; map.set(m.key, g); }
      g.units.push(m);
      m.model = g;
    }
    for (const g of map.values()) g.n = g.units.length;
    // 並びの判定：シート上で隣の行かつ台番号が連番
    for (let i = 0; i < machines.length; i++) {
      const a = machines[i], b = machines[i + 1];
      const adj = !!b && Math.abs(b.no - a.no) === 1;
      a.next = adj ? b : null;
      if (i === 0) a.prev = null;
      if (b) b.prev = adj ? a : null;
    }
    return { machines, models: [...map.values()], N: machines.length };
  };

  const markedCnt = (g, isMarked) => g.units.reduce((s, m) => s + (isMarked(m) ? 1 : 0), 0);
  const isFull = (g, isMarked) => g.n >= 3 && g.units.every(isMarked);
  const byDiffDesc = (a, b) => (b.diff ?? -1e9) - (a.diff ?? -1e9);
  const unmarkedOf = (g, isMarked) => g.units.filter(m => !isMarked(m)).sort(byDiffDesc);
  const shortNames = ms => [...new Set(ms.map(m => m.name))].join('・');

  function runsOf(ctx, pred) {
    const out = [];
    let cur = null;
    for (const m of ctx.machines) {
      if (pred(m)) {
        if (cur && cur[cur.length - 1].next === m) cur.push(m);
        else { cur = [m]; out.push(cur); }
      } else cur = null;
    }
    return out;
  }
  function segsOf(runs, size) {
    const segs = [];
    for (const r of runs) for (let i = 0; i + size <= r.length; i += size) segs.push(r.slice(i, i + size));
    return segs;
  }

  // 機種 g から k 台を公約に割り当てる。4台並びを崩さない台を優先する
  function consume(g, k, tag, st, isMarked) {
    const free = m => isMarked(m) && !st.used.has(m);
    const picked = [];
    for (let t = 0; t < k; t++) {
      let best = null, bestLoss = Infinity, bestLen = Infinity;
      for (const m of g.units) {
        if (!free(m)) continue;
        let a = 0, p = m.prev; while (p && free(p)) { a++; p = p.prev; }
        let b = 0, q = m.next; while (q && free(q)) { b++; q = q.next; }
        const L = a + b + 1;
        const loss = Math.floor(L / 4) - Math.floor(a / 4) - Math.floor(b / 4);
        if (loss < bestLoss || (loss === bestLoss && L < bestLen)) { best = m; bestLoss = loss; bestLen = L; }
      }
      if (!best) break;
      st.used.set(best, tag);
      picked.push(best);
    }
    return picked;
  }

  // ---------- ① 対象5機種 ----------
  function evalTargets(o, st, half, strict) {
    const { isMarked } = o;
    if (half) consume(half, ceilDiv(half.n, 2), 'thalf', st, isMarked);
    return o.targets.map(t => {
      const g = t.model;
      if (!g) return { t, status: 'none' };
      if (g === half) return { t, status: 'ok', via: 'half' };
      if (strict) {
        const p = consume(g, 1, 't6', st, isMarked);
        return { t, status: p.length ? 'ok' : 'ng', machine: p[0] };
      }
      // 兼用OK：⑤⑥が1台でもあれば達成（割り当ては最後に決める）
      return { t, status: g.units.some(isMarked) ? 'ok' : 'ng', lenient: true };
    });
  }

  // 兼用OKの場合の表示用割り当てと、1/2機種の残りの台の区分付け
  function finishTargets(items, half, st, isMarked) {
    for (const x of items) {
      if (!x.lenient || x.status !== 'ok') continue;
      const g = x.t.model;
      const free = g.units.find(m => isMarked(m) && !st.used.has(m));
      if (free) { st.used.set(free, 't6'); x.machine = free; }
      else { x.machine = g.units.find(isMarked); x.shared = true; }
    }
    if (half) for (const m of half.units) if (isMarked(m) && !st.used.has(m)) st.used.set(m, 'thalf');
  }

  function reqTargets(no, items, o) {
    const ng = items.filter(x => x.status === 'ng');
    const none = items.filter(x => x.status === 'none');
    return {
      id: 't1', no: no + '-1', title: '対象5機種すべてに⑥あり',
      status: ng.length ? 'ng' : 'ok',
      warn: none.length ? `未設置の対象機種：${none.map(x => x.t.label).join('、')}（設定で機種を選べます）` : null,
      lines: items.map(x => {
        const nm = x.t.label;
        if (x.status === 'none') return { s: 'warn', t: `${nm}：未設置` };
        const g = x.t.model;
        if (x.status === 'ok') {
          const where = x.via === 'half' ? '1/2機種' : `${x.machine.no}番${x.shared ? '・他の公約と兼用' : ''}`;
          return { s: 'ok', t: `${nm}：⑥あり（${where}）` };
        }
        const c = markedCnt(g, o.isMarked);
        return { s: 'ng', t: `${nm}：⑥なし（${g.n}台中 ⑤⑥${c}台${c ? '・すべて他の公約で使用' : ''}）` };
      }),
      // 空き台がなく補填できない未達は重く数え、補填できる割り振りを優先させる
      need: ng.reduce((s, x) => s + (x.t.model.units.some(m => !o.isMarked(m)) ? 1 : 50), 0),
      cands: () => ng.map(x => {
        const g = x.t.model;
        const um = unmarkedOf(g, o.isMarked);
        return {
          key: 'm:' + g.key, title: `${x.t.label}（${g.name}・${g.n}台）`, needText: 'あと1台', machines: um, model: g,
          note: um.length ? null : '空いている台がありません（全台が他の公約で使われています）',
        };
      }),
    };
  }

  function reqHalf(no, half, o, excluded) {
    const base = { id: 't2', no: no + '-2', title: '対象機種のうち1機種が1/2以上⑤⑥' };
    if (half) {
      const c = markedCnt(half, o.isMarked);
      return { ...base, status: 'ok', need: 0, lines: [{ s: 'ok', t: `${labelOf(o, half)}：${c}/${half.n}台（必要${ceilDiv(half.n, 2)}台）` }] };
    }
    const opts = o.targets.filter(t => t.model && !excluded.has(t.model)).map(t => {
      const g = t.model, c = markedCnt(g, o.isMarked);
      return { t, g, c, need: ceilDiv(g.n, 2) - c };
    }).sort((a, b) => a.need - b.need);
    return {
      ...base, status: 'ng', need: opts.length ? Math.max(1, opts[0].need) : 99,
      lines: opts.length
        ? opts.map(x => ({ s: 'ng', t: `${x.t.label}：${x.c}/${x.g.n}台（${x.need > 0 ? `1/2まであと${x.need}台` : '他の公約で使用中'}）` }))
        : [{ s: 'ng', t: '対象機種が設置されていません' }],
      cands: () => opts.filter(x => x.need > 0).map(x => ({
        key: 'm:' + x.g.key, title: `${x.t.label}（${x.g.name}・${x.g.n}台・⑤⑥${x.c}台）`, needText: `あと${x.need}台で1/2`, model: x.g,
        machines: unmarkedOf(x.g, o.isMarked),
      })),
    };
  }
  const labelOf = (o, g) => (o.targets.find(t => t.model === g) || { label: g.name }).label;

  // ---------- 補填候補 ----------
  // 全台系にできる機種（あと何台で全台か）
  function zenCands(ctx, o, excluded, filter, extra) {
    return ctx.models
      .filter(g => g.n >= 3 && !excluded.has(g) && !isFull(g, o.isMarked) && (!filter || filter(g)))
      .map(g => ({ g, c: markedCnt(g, o.isMarked), need: g.n - markedCnt(g, o.isMarked) }))
      .sort((a, b) => a.need - b.need || b.g.n - a.g.n)
      .slice(0, 10)
      .map(x => ({
        key: 'm:' + x.g.key, title: `${x.g.name}（${x.g.n}台・⑤⑥${x.c}台）`, needText: `あと${x.need}台で全台${extra ? extra(x.g) : ''}`, model: x.g,
        machines: unmarkedOf(x.g, o.isMarked), selectAll: true,
      }));
  }

  // n台並びを作れる場所（追加が少ない順）
  function runCands(ctx, o, st, size, keepTag, filter) {
    const blocked = m => st.used.has(m) && st.used.get(m) !== keepTag;
    const base = m => o.isMarked(m) && !blocked(m) && (!filter || filter(m));
    const baseline = segsOf(runsOf(ctx, base), size).length;
    const wins = [];
    const ms = ctx.machines;
    for (let i = 0; i + size <= ms.length; i++) {
      const w = ms.slice(i, i + size);
      let ok = true;
      for (let j = 0; j < size - 1 && ok; j++) if (w[j].next !== w[j + 1]) ok = false;
      if (!ok || w.some(blocked) || (filter && !w.every(filter))) continue;
      const add = w.filter(m => !o.isMarked(m));
      if (!add.length) continue;
      const addSet = new Set(add);
      const n = segsOf(runsOf(ctx, m => base(m) || addSet.has(m)), size).length;
      if (n <= baseline) continue;
      wins.push({ w, add, diff: add.reduce((s, m) => s + (m.diff || 0), 0) });
    }
    wins.sort((a, b) => a.add.length - b.add.length || b.diff - a.diff);
    const taken = new Set(), out = [];
    for (const x of wins) {
      if (x.w.some(m => taken.has(m))) continue;
      x.w.forEach(m => taken.add(m));
      out.push({
        key: `w:${x.w[0].i}:${size}`, title: `${x.w[0].no}〜${x.w[x.w.length - 1].no}（${shortNames(x.w)}）`,
        needText: `あと${x.add.length}台で${size}台並び`, machines: x.w, selectAll: true, avgUnits: x.w,
      });
      if (out.length >= 10) break;
    }
    return out;
  }

  // どの公約にも数えなかった⑤⑥の台の区分（ホールの付け方に合わせる）
  // 対象5機種の台は①、極では4台以上の機種の台は②、それ以外は「その他」
  function finishOthers(ctx, st, isMarked, o, bigTag) {
    const targets = new Set(o.targets.map(t => t.model).filter(Boolean));
    for (const m of ctx.machines) {
      if (!isMarked(m) || st.used.has(m)) continue;
      st.used.set(m, targets.has(m.model) ? 't6' : bigTag && m.model.n >= 4 ? bigTag : 'other');
    }
    return st.used;
  }

  const score = reqs => reqs.reduce((s, r) => s + (r.status === 'ok' ? 1000 : r.status === 'ng' ? -Math.min(r.need, 999) : 0), 0);

  function halfOptions(o) {
    return [null, ...o.targets.map(t => t.model).filter(g => g && markedCnt(g, o.isMarked) >= ceilDiv(g.n, 2))];
  }

  // 同点なら、補填した台をなるべく使わない割り振り（＝ホールが付けた⭕️を活かす）を選ぶ。
  // 補填の途中で割り振りが入れ替わり、選んでいる候補が消えるのを防ぐため
  const suppCost = (o, models) => models.reduce((s, g) => s + (g ? g.units.filter(o.isSupp).length : 0), 0);

  function pickBest(cands) {
    let best = null;
    for (const c of cands) {
      const s = score(c.reqs);
      if (!best || s > best.score || (s === best.score && c.cost < best.cost)) best = { ...c, score: s };
    }
    for (const r of best.reqs) r.groups = r.status === 'ng' && r.cands ? r.cands() : [];
    return best;
  }

  function pairs(list) {
    const out = [];
    for (let i = 0; i < list.length; i++) for (let j = i + 1; j < list.length; j++) out.push([list[i], list[j]]);
    return out;
  }

  // ---------- 極 ----------
  function solveKiwami(ctx, o) {
    const evals = [];
    for (const half of halfOptions(o)) {
      const eligible = ctx.models.filter(g => g !== half && isFull(g, o.isMarked));
      const combos = eligible.length <= 2 ? [eligible] : pairs(eligible);
      for (const zen of combos) evals.push(() => evalKiwami(ctx, o, half, zen));
    }
    return pickBest(evals.map(f => f()));
  }

  function evalKiwami(ctx, o, half, zen) {
    const { isMarked } = o;
    const strict = !o.lenient;
    const st = { used: new Map() };
    const free = m => isMarked(m) && !st.used.has(m);
    for (const g of zen) for (const m of g.units) st.used.set(m, 'k3');
    const items = evalTargets(o, st, half, strict);
    const segs = segsOf(runsOf(ctx, free), 4);
    for (const s of segs) for (const m of s) st.used.set(m, 'k4');
    finishTargets(items, half, st, isMarked);
    const big = ctx.models.filter(g => g.n >= 4);
    const uncovered = big.filter(g => !g.units.some(isMarked));
    for (const g of big) {
      if (g.units.some(m => st.used.has(m))) continue;
      const m = g.units.find(free);
      if (m) st.used.set(m, 'k2');
    }
    finishOthers(ctx, st, isMarked, o, 'k2');

    const targetSet = new Set(o.targets.map(t => t.model).filter(Boolean));
    const zenExcl = new Set([...zen, half, ...(strict ? targetSet : [])].filter(Boolean));
    const missZen = Math.max(0, 2 - zen.length);
    const zenNeed = missZen ? ctx.models.filter(g => g.n >= 3 && !zenExcl.has(g)).map(g => g.n - markedCnt(g, isMarked)).sort((a, b) => a - b).slice(0, missZen).reduce((s, x) => s + x, 0) || 99 : 0;

    const reqs = [
      reqTargets('①', items, o),
      reqHalf('①', half, o, new Set(zen)),
      {
        id: 'k2', no: '②', title: '4台以上設置の全機種に⑤⑥',
        status: uncovered.length ? 'ng' : 'ok', need: uncovered.length,
        lines: [
          { s: uncovered.length ? 'ng' : 'ok', t: `4台以上設置：${big.length}機種 ／ ⑤⑥なし：${uncovered.length}機種` },
          ...uncovered.map(g => ({ s: 'ng', t: `${g.name}（${g.n}台）` })),
        ],
        cands: () => uncovered.map(g => ({ key: 'm:' + g.key, title: `${g.name}（${g.n}台）`, needText: 'あと1台', machines: unmarkedOf(g, isMarked), model: g })),
      },
      {
        id: 'k3', no: '③', title: '3台以上機種の全台⑤⑥ ×2機種',
        status: missZen ? 'ng' : 'ok', need: zenNeed,
        lines: zen.length ? zen.map(g => ({ s: 'ok', t: `全台系：${g.name}（${g.n}台）` })) : [{ s: 'ng', t: '全台系なし' }],
        cands: () => zenCands(ctx, o, zenExcl),
      },
      {
        id: 'k4', no: '④', title: '4台並び⑤⑥ ×2箇所',
        status: segs.length >= 2 ? 'ok' : 'ng', need: Math.max(0, 2 - segs.length) * 2,
        lines: segs.length ? segs.map(s => ({ s: 'ok', t: `${s[0].no}〜${s[3].no}（${shortNames(s)}）` })) : [{ s: 'ng', t: '4台並びなし' }],
        cands: () => runCands(ctx, o, st, 4, 'k4'),
      },
    ];
    return { reqs, assign: st.used, cost: suppCost(o, [half, ...zen]) };
  }

  // ---------- 紫 ----------
  function solveMurasaki(ctx, o) {
    const evals = [];
    for (const half of halfOptions(o)) {
      const m2opts = [null, ...ctx.models.filter(g => g !== half && o.is1149(g) && isFull(g, o.isMarked))];
      for (const m2 of m2opts) evals.push(evalMurasaki(ctx, o, half, m2));
    }
    return pickBest(evals);
  }

  function evalMurasaki(ctx, o, half, m2) {
    const { isMarked } = o;
    const strict = !o.lenient;
    const st = { used: new Map() };
    if (m2) for (const m of m2.units) st.used.set(m, 'm2');
    const items = evalTargets(o, st, half, strict);
    const small = m => m.model.n <= 2;
    const runs = runsOf(ctx, m => small(m) && isMarked(m) && !st.used.has(m)).filter(r => r.length >= 2);
    const ok4 = runs.some(r => r.length >= 3) || runs.length >= 2;
    for (const r of runs) for (const m of r) st.used.set(m, 'm4');
    finishTargets(items, half, st, isMarked);
    finishOthers(ctx, st, isMarked, o);

    const flagged = ctx.models.filter(g => g.n >= 3 && o.is1149(g));
    const targetSet = new Set(o.targets.map(t => t.model).filter(Boolean));
    const excl = new Set([half, ...(strict ? targetSet : [])].filter(Boolean));
    const smallCount = ctx.machines.filter(small).length;

    const reqs = [
      reqTargets('①', items, o),
      reqHalf('①', half, o, new Set([m2].filter(Boolean))),
      {
        id: 'm2', no: '②', title: '3台以上設置の114.9%機種から全⑥',
        status: m2 ? 'ok' : 'ng',
        need: m2 ? 0 : (flagged.filter(g => !excl.has(g)).map(g => g.n - markedCnt(g, isMarked)).sort((a, b) => a - b)[0] || 99),
        lines: m2
          ? [{ s: 'ok', t: `全⑥：${m2.name}（${m2.n}台）` }]
          : flagged.length
            ? [{ s: 'ng', t: `114.9%機種（3台以上）：${flagged.map(g => `${g.name}(${markedCnt(g, isMarked)}/${g.n})`).join('、')}` }]
            : [{ s: 'ng', t: '114.9%機種（3台以上設置）が見つかりません。設定の「114.9%機種」を確認してください' }],
        cands: () => zenCands(ctx, o, excl, g => o.is1149(g)),
      },
      {
        id: 'm3', no: '③', title: '10台以上の列から1/2⑤⑥の列（列数は設置台数で変動）',
        status: 'skip', need: 0, lines: [{ s: 'skip', t: '列の判定は行いません（対象外）' }],
      },
      {
        id: 'm4', no: '④', title: '2台設置以下に⑤⑥の3台並び or 2台並び×2',
        status: smallCount ? (ok4 ? 'ok' : 'ng') : 'skip', need: ok4 ? 0 : 2,
        lines: smallCount
          ? (runs.length ? runs.map(r => ({ s: ok4 ? 'ok' : 'ng', t: `${r.length}台並び：${r[0].no}〜${r[r.length - 1].no}（${shortNames(r)}）` })) : [{ s: 'ng', t: '2台設置以下の並びなし' }])
          : [{ s: 'skip', t: '2台設置以下の機種がありません' }],
        cands: () => {
          const out = runCands(ctx, o, st, 3, 'm4', small);
          if (runs.length === 1) out.push(...runCands(ctx, o, st, 2, 'm4', small));
          return out.slice(0, 12);
        },
      },
    ];
    return { reqs, assign: st.used, cost: suppCost(o, [half, m2]) };
  }

  // ---------- 赤 ----------
  function solveAka(ctx, o) {
    const strict = !o.lenient;
    const targetSet = new Set(o.targets.map(t => t.model).filter(Boolean));
    const evals = [];
    for (const half of halfOptions(o)) {
      const r3opts = [null, ...ctx.models.filter(g => g !== half && g.n >= 3 && markedCnt(g, o.isMarked) >= ceilDiv(g.n, 2))];
      for (const r3 of r3opts) {
        const full = ctx.models.filter(g => g !== half && g !== r3 && isFull(g, o.isMarked));
        const fullT = strict ? full.filter(g => targetSet.has(g)) : [];
        const fixed = full.filter(g => !fullT.includes(g));
        for (let mask = 0; mask < (1 << fullT.length); mask++) {
          const zen = [...fixed, ...fullT.filter((g, i) => mask & (1 << i))];
          evals.push(evalAka(ctx, o, half, r3, zen));
        }
      }
    }
    return pickBest(evals);
  }

  function evalAka(ctx, o, half, r3, zen) {
    const { isMarked } = o;
    const strict = !o.lenient;
    const st = { used: new Map() };
    for (const g of zen) for (const m of g.units) st.used.set(m, 'a2');
    // ③の1/2機種は半分だけ使う（残りは①の⑥に回せる）
    if (r3) consume(r3, ceilDiv(r3.n, 2), 'a3', st, isMarked);
    const items = evalTargets(o, st, half, strict);
    const small = ctx.machines.filter(m => m.model.n <= 2);
    const smallMarked = small.filter(m => isMarked(m) && !st.used.has(m));
    for (const m of smallMarked) st.used.set(m, 'a4');
    finishTargets(items, half, st, isMarked);
    if (r3) for (const m of r3.units) if (isMarked(m) && !st.used.has(m)) st.used.set(m, 'a3');
    finishOthers(ctx, st, isMarked, o);

    const N = ctx.N, R2 = ceilDiv(N * 4, 100);
    const have2 = zen.reduce((s, g) => s + g.n, 0);
    const R4 = ceilDiv(small.length, 5);
    const targetSet = new Set(o.targets.map(t => t.model).filter(Boolean));
    const used = new Set([...zen, half, r3].filter(Boolean));
    // ③の機種も全台系の候補に残す（全台にすれば割り振りを組み直す）
    const zenExcl = new Set([...zen, half, ...(strict ? targetSet : [])].filter(Boolean));

    const r3opts = ctx.models.filter(g => g.n >= 3 && !used.has(g))
      .map(g => {
        const c = markedCnt(g, isMarked);
        // 被りなしなら、対象5機種は①の⑥の1台を残す必要がある
        return { g, c, need: ceilDiv(g.n, 2) - c + (strict && targetSet.has(g) ? 1 : 0) };
      })
      .filter(x => x.need > 0)
      .sort((a, b) => a.need - b.need);

    const reqs = [
      reqTargets('①', items, o),
      reqHalf('①', half, o, new Set([...zen, r3].filter(Boolean))),
      {
        id: 'a2', no: '②', title: '総設置の4%を全台系で配置（3台以上機種）',
        status: have2 >= R2 ? 'ok' : 'ng', need: Math.max(0, R2 - have2),
        lines: [
          { s: have2 >= R2 ? 'ok' : 'ng', t: `全台系 計${have2}台 ／ 必要${R2}台（設置${N}台×4%）` },
          ...zen.map(g => ({ s: 'ok', t: `${g.name}（${g.n}台）` })),
        ],
        cands: () => zenCands(ctx, o, zenExcl, null, g => `（+${g.n}台）`),
      },
      {
        id: 'a3', no: '③', title: '1/2で⑤⑥が1機種（3台以上機種）',
        status: r3 ? 'ok' : 'ng', need: r3 ? 0 : (r3opts[0] ? r3opts[0].need : 99),
        lines: r3
          ? [{ s: 'ok', t: `${r3.name}：${markedCnt(r3, isMarked)}/${r3.n}台（必要${ceilDiv(r3.n, 2)}台）` }]
          : [{ s: 'ng', t: '1/2機種なし' }],
        cands: () => r3opts.slice(0, 10).map(x => ({
          key: 'm:' + x.g.key, title: `${x.g.name}（${x.g.n}台・⑤⑥${x.c}台）`, needText: `あと${x.need}台で1/2`, machines: unmarkedOf(x.g, isMarked), model: x.g,
        })),
      },
      {
        id: 'a4', no: '④', title: '2台設置以下の1/5が⑤⑥',
        status: small.length ? (smallMarked.length >= R4 ? 'ok' : 'ng') : 'skip',
        need: Math.max(0, R4 - smallMarked.length),
        lines: small.length
          ? [{ s: smallMarked.length >= R4 ? 'ok' : 'ng', t: `2台設置以下 ${small.length}台 → 必要${R4}台 ／ ⑤⑥${smallMarked.length}台` }]
          : [{ s: 'skip', t: '2台設置以下の機種がありません' }],
        cands: () => [{
          key: 'small', title: '2台設置以下の台', needText: `あと${R4 - smallMarked.length}台`,
          machines: small.filter(m => !isMarked(m)).sort(byDiffDesc),
        }],
      },
    ];
    return { reqs, assign: st.used, cost: suppCost(o, [half, r3, ...zen]) };
  }

  // ---------- 青 ----------
  function solveAo(ctx, o) {
    return pickBest(halfOptions(o).map(half => evalAo(ctx, o, half)));
  }

  function evalAo(ctx, o, half) {
    const { isMarked } = o;
    const st = { used: new Map() };
    const free = m => isMarked(m) && !st.used.has(m);
    // 青は「上記含めず」なので①の台は必ず別扱い
    const items = evalTargets(o, st, half, true);
    const segs = segsOf(runsOf(ctx, free), 4);
    for (const s of segs) for (const m of s) st.used.set(m, 'o2');
    finishTargets(items, half, st, isMarked);
    finishOthers(ctx, st, isMarked, o);
    const N = ctx.N, need = ceilDiv(N, 50); // N×8%÷4 の切り上げ
    const reqs = [
      reqTargets('①', items, o),
      reqHalf('①', half, o, new Set()),
      {
        id: 'o2', no: '②', title: '①以外で総設置の8%分を⑤⑥の4台並び',
        status: segs.length >= need ? 'ok' : 'ng', need: Math.max(0, need - segs.length) * 2,
        lines: [
          { s: segs.length >= need ? 'ok' : 'ng', t: `4台並び ${segs.length}箇所 ／ 必要${need}箇所（設置${N}台×8%÷4）` },
          ...segs.map(s => ({ s: 'ok', t: `${s[0].no}〜${s[3].no}（${shortNames(s)}）` })),
        ],
        cands: () => runCands(ctx, o, st, 4, 'o2'),
      },
    ];
    return { reqs, assign: st.used, cost: suppCost(o, [half]) };
  }

  R.solve = function (ctx, o) {
    const fn = { kiwami: solveKiwami, murasaki: solveMurasaki, aka: solveAka, ao: solveAo }[o.type];
    return fn(ctx, { isSupp: () => false, ...o });
  };

  window.KoyakuRules = R;
})();
