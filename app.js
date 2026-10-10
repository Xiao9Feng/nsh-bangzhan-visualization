/* 帮战数据分析 · 成员五维 / 联赛对比
   纯前端：SheetJS 读 Excel + ECharts 出图。双击 index.html 即可运行。 */
(function () {
  'use strict';

  /* ================= 字段定义 =================
     已适配「逆水寒帮会战报」原生导出的列名，别名可自行扩充 */
  const FIELDS = [
    { key: 'name',   label: '角色名', type: 'text', required: true,
      aliases: ['角色名','玩家名字','玩家名','昵称','成员','名字','姓名','玩家','角色','id','uid'] },
    { key: 'school', label: '门派', type: 'text', aliases: ['职业','门派','流派','心法','宗派'] },
    { key: 'guild',  label: '帮会', type: 'text', aliases: ['帮会','帮派','公会','阵营','势力','所属帮会','帮派名'] },
    { key: 'damage', label: '人伤', type: 'num',
      aliases: ['对玩家伤害','人伤','人伤害','对人伤害','伤害','输出','总伤害','造成伤害','伤害量','dmg','damage'] },
    { key: 'armorBreak', label: '人伤卸甲', type: 'num', aliases: ['人伤卸甲','对人卸甲'] },
    { key: 'tower',  label: '塔伤', type: 'num',
      aliases: ['对建筑伤害','塔伤','塔伤害','对塔伤害','建筑伤害','拆塔','推塔'] },
    { key: 'towerBreak', label: '破塔卸甲', type: 'num', aliases: ['破塔卸甲','对建筑卸甲'] },
    { key: 'tank',   label: '承伤', type: 'num', aliases: ['承受伤害','承伤','受到伤害','承伤害','taken'] },
    { key: 'heal',   label: '治疗', type: 'num', aliases: ['治疗值','治疗','治疗量','奶量','治疗输出','heal'] },
    { key: 'death',  label: '重伤', type: 'num',
      aliases: ['重伤','重伤次数','死亡','死亡数','阵亡','走火','death','deaths'] },
    { key: 'kill',   label: '击败', type: 'num',
      aliases: ['击败/清泉','击败/清泉值','击败','击败数','击杀','击杀数','杀敌','kill','kills'] },
    { key: 'assist', label: '助攻', type: 'num', aliases: ['助攻','助攻数','救援','assist','assists'] },
    { key: 'revive', label: '复活', type: 'num', aliases: ['复活/清泉','复活','救援次数'] },
    { key: 'resource', label: '资源', type: 'num', aliases: ['资源','资源值'] },
    { key: 'fireBone', label: '焚骨', type: 'num', aliases: ['焚骨'] },
    { key: 'score',  label: '积分', type: 'num', aliases: ['积分','战功','得分','贡献','贡献值','score','points'] }
  ];

  /* ================= 五维维度 & 口径 ================= */
  const DIM = {
    damage:       { label: '人伤',   fmt: 'num', get: (r) => r.damage },
    tower:        { label: '塔伤',   fmt: 'num', get: (r) => r.tower },
    tank:         { label: '承伤',   fmt: 'num', get: (r) => r.tank },
    heal:         { label: '治疗',   fmt: 'num', get: (r) => r.heal },
    kill:         { label: '击杀',   fmt: 'int', get: (r) => r.kill },
    assist:       { label: '助攻',   fmt: 'int', get: (r) => r.assist },
    death:        { label: '重伤',   fmt: 'int', get: (r) => r.death },
    survive:      { label: '生存',   fmt: 'idx', get: (r) => 1 / (1 + (r.death || 0)) },
    perDeathDmg:  { label: '每死输出', fmt: 'num', get: (r) => r.damage / Math.max(1, r.death) },
    perDeathTank: { label: '每死承伤', fmt: 'num', get: (r) => r.tank / Math.max(1, r.death) },
    kda:          { label: 'KDA',   fmt: 'kda', get: (r) => (r.kill + r.assist) / Math.max(1, r.death) }
  };
  const TABS = [
    { key: 'all', label: '综合', dims: ['damage', 'survive', 'tank', 'tower', 'assist'] },
    { key: 'atk', label: '进攻', dims: ['damage', 'tower', 'kill', 'perDeathDmg', 'assist'] },
    { key: 'def', label: '防守', dims: ['tank', 'perDeathTank', 'survive', 'heal', 'perDeathDmg'] }
  ];

  const SCHOOL_COLOR = {
    '铁衣': '#5b8ff9', '素问': '#5ad8a6', '神相': '#9270ca', '碎梦': '#f2676b',
    '玄机': '#f6bd16', '血河': '#e8684a', '九灵': '#5d7092', '龙吟': '#ff9d4d',
    '沧澜': '#6dc8ec', '潮光': '#269a99', '鸿音': '#c78ae8',
    '鸿音·D': '#f2676b', '鸿音·奶': '#5ad8a6'
  };
  const PALETTE = ['#5b8ff9', '#5ad8a6', '#f6bd16', '#9270ca', '#f2676b',
    '#ff9d4d', '#6dc8ec', '#e8684a', '#269a99', '#c78ae8'];
  const colorOf = (n) => {
    if (!n) return '#9aa5bd';
    if (SCHOOL_COLOR[n]) return SCHOOL_COLOR[n];
    let h = 0; for (let i = 0; i < n.length; i++) h = (h * 31 + n.charCodeAt(i)) % 997;
    return PALETTE[h % PALETTE.length];
  };

  const state = {
    workbook: null, sheetNames: [], sheetIndex: 0,
    headers: [], rawRows: [], rows: [], mapping: {},
    sortKey: 'damage', sortDir: 'desc',
    modalPlayer: null, modalTab: 'all',
    view: 'members', myGuild: null, aggMode: 'total', guildCount: 0
  };
  const charts = {};
  let modalChart = null;
  const $ = (id) => document.getElementById(id);
  /* 默认「我方帮会」：人数最多的帮会（开源版不绑定任何具体帮会） */
  const topGuild = () => { const c = {}; state.rows.forEach((r) => { if (r.guild) c[r.guild] = (c[r.guild] || 0) + 1; }); return Object.keys(c).sort((a, b) => c[b] - c[a])[0] || null; };

  /* ================= 基础工具 ================= */
  const norm = (s) => String(s == null ? '' : s).replace(/[\s\u3000*()（）\[\]【】:：\/\\]/g, '').toLowerCase();
  const esc = (s) => String(s == null ? '' : s)
    .replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;').replace(/"/g, '&quot;');

  function parseNum(v) {
    if (v == null || v === '') return 0;
    if (typeof v === 'number') return isFinite(v) ? v : 0;
    let s = String(v).trim(); if (!s) return 0;
    if (s.indexOf('/') >= 0) s = s.split('/')[0];    // 「击败/清泉」这类斜杠列取左值
    s = s.replace(/,/g, '');
    let mult = 1;
    if (s.includes('亿')) { mult = 1e8; s = s.replace(/亿/g, ''); }
    else if (s.includes('万')) { mult = 1e4; s = s.replace(/万/g, ''); }
    s = s.replace(/[^\d.\-+eE]/g, '');
    const n = parseFloat(s);
    return isFinite(n) ? n * mult : 0;
  }
  function formatNum(n) {
    n = Number(n) || 0; const a = Math.abs(n);
    if (a >= 1e8) return (n / 1e8).toFixed(2) + '亿';
    if (a >= 1e4) return (n / 1e4).toFixed(2) + '万';
    return Math.round(n).toLocaleString('zh-CN');
  }
  const fmtInt = (n) => Math.round(Number(n) || 0).toLocaleString('zh-CN');
  function fmtVal(key, v) {
    const f = (DIM[key] || {}).fmt || 'num';
    if (f === 'int') return fmtInt(v);
    if (f === 'idx' || f === 'kda') return (Number(v) || 0).toFixed(2);
    return formatNum(v);
  }
  let toastTimer = null;
  function toast(m) {
    const t = $('toast'); t.textContent = m; t.hidden = false;
    clearTimeout(toastTimer); toastTimer = setTimeout(() => { t.hidden = true; }, 3200);
  }

  /* ================= 读文件 / 解析 ================= */
  /* 文本解码：优先 UTF-8，出现乱码替换符再回退 GBK */
  function decodeText(buf) {
    let s = '';
    try { s = new TextDecoder('utf-8', { fatal: false }).decode(buf); } catch (e) { s = ''; }
    if (!s || s.indexOf('\uFFFD') >= 0) {
      try { s = new TextDecoder('gbk').decode(buf); } catch (e2) { /* 保持原样 */ }
    }
    return s;
  }

  function handleFile(file) {
    const isText = /\.(csv|txt)$/i.test(file.name);
    const rd = new FileReader();
    rd.onload = (e) => {
      try {
        let wb;
        if (isText) {
          // 先自己按 UTF-8/GBK 解码，再交给 SheetJS；raw:true 防止它把 " 5/1" 当成日期
          const buf = new Uint8Array(e.target.result);
          wb = XLSX.read(decodeText(buf), { type: 'string', raw: true });
        } else {
          wb = XLSX.read(new Uint8Array(e.target.result), { type: 'array' });
        }
        state.workbook = wb; state.sheetNames = wb.SheetNames; state.sheetIndex = 0;
        $('srcInfo').textContent = file.name + ' · ' + new Date().toLocaleString('zh-CN');
        loadSheet(0);
      } catch (err) { toast('解析失败：' + err.message); }
    };
    rd.readAsArrayBuffer(file);   // 文本文件也在回调里自行解码
  }

  function loadSheet(idx) {
    state.sheetIndex = idx;
    const ws = state.workbook.Sheets[state.sheetNames[idx]];
    const aoa = XLSX.utils.sheet_to_json(ws, { header: 1, raw: true, defval: '', blankrows: false });
    const t = unifyAoa(aoa);
    if (!t.rows.length) { toast('这个工作表里没解析出数据行。'); return; }
    state.headers = t.headers;
    state.rawRows = t.rows;
    state.guildCount = t.guildCount;
    state.mapping = autoMap(state.headers);
    renderMapping();
    analyze();
  }

  /* ================= 表格规整 =================
     支持两种输入：
     1) 普通单表：第一行是表头
     2) 逆水寒帮会战报原生导出：多个帮会分段，每段第一行是 ["帮会名", 人数]，
        第二行是列名，之后是数据。帮会名会被补成一列。 */
  function unifyAoa(aoa) {
    const clean = aoa.filter((r) => r.some((c) => c !== '' && c != null));
    if (!clean.length) return { headers: [], rows: [], guildCount: 0 };

    const numLike = (v) => typeof v === 'number' ||
      (typeof v === 'string' && /^\d+(\.\d+)?$/.test(v.trim()));
    const sectionIdx = [];
    clean.forEach((r, i) => {
      const ne = r.filter((c) => c !== '' && c != null);
      if (ne.length === 2 && numLike(ne[1]) &&
          !/^\d+(\.\d+)?$/.test(String(ne[0]).trim())) sectionIdx.push(i);
    });
    const isSectioned = sectionIdx.length > 0 && sectionIdx.every((i) =>
      clean[i + 1] && clean[i + 1].filter((c) => c !== '' && c != null).length >= 3);

    if (isSectioned) {
      const base = clean[sectionIdx[0] + 1].map((h, i) =>
        String(h == null ? '' : h).trim() || ('列' + (i + 1)));
      const headers = ['帮会'].concat(base);
      const rows = [];
      sectionIdx.forEach((start, si) => {
        const guild = String(clean[start][0]).trim();
        const end = si + 1 < sectionIdx.length ? sectionIdx[si + 1] : clean.length;
        for (let i = start + 2; i < end; i++) {
          const r = clean[i];
          if (!r || !r.some((c) => c !== '' && c != null)) continue;
          const out = new Array(headers.length).fill('');
          out[0] = guild;
          for (let j = 0; j < base.length; j++) out[j + 1] = r[j] === undefined ? '' : r[j];
          rows.push(out);
        }
      });
      return { headers, rows, guildCount: sectionIdx.length };
    }

    const headers = clean[0].map((h, i) => String(h == null ? '' : h).trim() || ('列' + (i + 1)));
    const rows = clean.slice(1).map((r) => {
      const out = new Array(headers.length).fill('');
      for (let j = 0; j < headers.length; j++) out[j] = r[j] === undefined ? '' : r[j];
      return out;
    });
    return { headers, rows, guildCount: 0 };
  }

  function renderMapping() {
    const opts = (sel) => '<option value="">— 不使用 —</option>' + state.headers.map((h, i) =>
      `<option value="${i}" ${String(sel) === String(i) ? 'selected' : ''}>${esc(h)}</option>`).join('');
    $('mapGrid').innerHTML = FIELDS.map((f) =>
      `<div class="map-item"><label>${f.label}${f.required ? ' <b>*</b>' : ''}</label>` +
      `<select class="input" data-key="${f.key}">${opts(state.mapping[f.key])}</select></div>`).join('');
    $('mapGrid').querySelectorAll('select').forEach((s) => {
      s.onchange = () => { state.mapping[s.dataset.key] = s.value === '' ? '' : Number(s.value); };
    });
  }

  function autoMap(headers) {
    const hN = headers.map(norm);
    const map = {}; FIELDS.forEach((f) => { map[f.key] = ''; });
    const used = new Set();
    FIELDS.forEach((f) => {                       // 第一轮：精确匹配
      for (const a of f.aliases) {
        const i = hN.indexOf(norm(a));
        if (i >= 0 && !used.has(i)) { map[f.key] = i; used.add(i); return; }
      }
    });
    FIELDS.forEach((f) => {                       // 第二轮：包含匹配
      if (map[f.key] !== '') return;
      for (const a of f.aliases) {
        const an = norm(a); if (an.length < 2) continue;
        const i = hN.findIndex((h, x) => h && h.includes(an) && !used.has(x));
        if (i >= 0) { map[f.key] = i; used.add(i); return; }
      }
    });
    return map;
  }

  /* ================= 计算 ================= */
  function analyze() {
    const rows = state.rawRows.map((r) => {
      const o = {};
      FIELDS.forEach((f) => {
        const i = state.mapping[f.key];
        const v = (i === '' || i == null) ? '' : r[i];
        o[f.key] = f.type === 'num' ? parseNum(v) : String(v == null ? '' : v).trim();
      });
      return o;
    }).filter((o) => o.name);
    /* 鸿音自动细分：治疗值 > 0 且 >（人伤 + 塔伤）判为「奶鸿」，否则「D 鸿」 */
    rows.forEach((o) => { if (o.school === '鸿音') o.school = (o.heal > 0 && o.heal > o.damage + o.tower) ? '鸿音·奶' : '鸿音·D'; });
    if (!rows.length) {
      $('mapPanel').hidden = false;
      renderMapping();
      toast('没认出来数据：请检查表头是不是第一行，或在「字段设置」里手动指定「角色名」。');
      return;
    }
    state.rows = rows;

    const sum = (k) => rows.reduce((a, b) => a + (b[k] || 0), 0);
    state.totals = { count: rows.length, damage: sum('damage'), tower: sum('tower'),
      tank: sum('tank'), heal: sum('heal'), kill: sum('kill'), assist: sum('assist'),
      death: sum('death'), score: sum('score') };
    state.rows.forEach((o) => { o.dmgPct = state.totals.damage ? o.damage / state.totals.damage : 0; });

    if (!state.myGuild) {
      state.myGuild = topGuild();
    }

    $('emptyHint').hidden = true;
    $('memberWrap').hidden = false;
    $('leagueEmpty').hidden = true;
    $('leagueWrap').hidden = false;
    $('btnExport').disabled = false;
    $('btnImport').textContent = '换个文件';

    fillFilters(); fillGuildSelect(); renderMembers();
    if (state.view === 'league') renderLeague();
    toast('已解析 ' + rows.length + ' 名成员' +
      (state.guildCount > 1 ? '（合并 ' + state.guildCount + ' 个帮会）' : ''));
  }

  function availableDims() {
    return Object.keys(DIM).filter((k) => state.rows.some((r) => DIM[k].get(r) > 0));
  }
  function tabDims(tabKey) {
    const tab = TABS.find((t) => t.key === tabKey) || TABS[0];
    const av = availableDims();
    let ds = tab.dims.filter((d) => av.indexOf(d) >= 0);
    if (ds.length < 3) ds = av.filter((d) => d !== 'death' && d !== 'kda').slice(0, 5);
    return ds.slice(0, 5);
  }

  /* ================= 视图一：成员表 ================= */
  function fillFilters() {
    const schools = [...new Set(state.rows.map((r) => r.school).filter(Boolean))];
    const guilds = [...new Set(state.rows.map((r) => r.guild).filter(Boolean))];
    $('filterSchool').innerHTML = '<option value="">全部门派</option>' +
      schools.map((s) => `<option value="${esc(s)}">${esc(s)}</option>`).join('');
    $('filterGuild').innerHTML = '<option value="">全部帮会</option>' +
      guilds.map((s) => `<option value="${esc(s)}">${esc(s)}</option>`).join('');
  }

  function memberCols() {
    const cols = [{ key: 'name', label: '角色名', txt: true },
      { key: 'school', label: '门派', txt: true, pill: 'sch' },
      { key: 'guild', label: '帮会', txt: true, pill: 'gd' }];
    ['damage', 'tower', 'tank', 'death', 'kill', 'assist', 'heal', 'score'].forEach((k) => {
      if (state.mapping[k] !== '' && state.mapping[k] != null) {
        cols.push({ key: k, label: FIELDS.find((f) => f.key === k).label, dim: k });
      }
    });
    if (state.mapping.kill !== '' || state.mapping.death !== '' || state.mapping.assist !== '') {
      cols.push({ key: 'kda', label: 'KDA', dim: 'kda' });
    }
    if (state.mapping.tank !== '' && state.mapping.death !== '') {
      cols.push({ key: 'perDeathTank', label: '每死承伤', dim: 'perDeathTank' });
    }
    return cols;
  }

  function renderMembers() {
    const cols = memberCols();
    const kw = $('search').value.trim().toLowerCase();
    const fs = $('filterSchool').value, fg = $('filterGuild').value;
    let rows = state.rows.filter((r) => {
      if (fs && r.school !== fs) return false;
      if (fg && r.guild !== fg) return false;
      if (kw && !(r.name + ' ' + r.school + ' ' + r.guild).toLowerCase().includes(kw)) return false;
      return true;
    });
    const k = state.sortKey, dir = state.sortDir === 'asc' ? 1 : -1;
    rows.sort((a, b) => {
      const va = DIM[k] ? DIM[k].get(a) : a[k], vb = DIM[k] ? DIM[k].get(b) : b[k];
      if (typeof va === 'string' || typeof vb === 'string') return String(va).localeCompare(String(vb), 'zh-CN') * dir;
      return ((va || 0) - (vb || 0)) * dir;
    });

    $('memCount').textContent = `共 ${rows.length} / ${state.rows.length} 人`;
    $('memTable').innerHTML =
      '<thead><tr><th class="t">#</th>' +
      cols.map((c) => `<th data-k="${c.key}" class="${c.txt ? 't' : ''}">${c.label}` +
        (state.sortKey === c.key ? `<span class="ar">${state.sortDir === 'asc' ? '▲' : '▼'}</span>` : '') +
        '</th>').join('') +
      '<th>分析</th></tr></thead><tbody>' +
      rows.map((r, i) => '<tr data-name="' + esc(r.name) + '">' +
        '<td class="t" style="color:#9aa5bd">' + (i + 1) + '</td>' +
        cols.map((c) => {
          if (c.key === 'name') return `<td class="t"><span class="nm">${esc(r.name)}</span></td>`;
          if (c.pill) return `<td class="t"><span class="pill ${c.pill}">${esc(r[c.key] || '—')}</span></td>`;
          return `<td>${fmtVal(c.dim || c.key, DIM[c.key] ? DIM[c.key].get(r) : r[c.key])}</td>`;
        }).join('') +
        '<td><button class="dim-btn">五维</button></td></tr>').join('') +
      '</tbody>';

    $('memTable').querySelectorAll('thead th[data-k]').forEach((th) => {
      th.onclick = () => {
        const kk = th.dataset.k;
        if (state.sortKey === kk) state.sortDir = state.sortDir === 'asc' ? 'desc' : 'asc';
        else { state.sortKey = kk; state.sortDir = (kk === 'name' || kk === 'school' || kk === 'guild') ? 'asc' : 'desc'; }
        renderMembers();
      };
    });
    $('memTable').querySelectorAll('tbody tr').forEach((tr) => {
      tr.onclick = () => openModal(tr.dataset.name);
    });
  }

  /* ================= 个人雷达弹窗 ================= */
  function openModal(name) {
    const me = state.rows.find((r) => r.name === name);
    if (!me) return;
    state.modalPlayer = name;
    $('mName').textContent = me.name;
    $('mSchool').textContent = [me.school, me.guild].filter(Boolean).join(' · ') || '—';
    $('mAvatar').textContent = (me.school || me.name || '战').slice(0, 1);
    $('mAvatar').style.background = 'linear-gradient(140deg,' + colorOf(me.school) + 'cc,' + colorOf(me.school) + ')';
    $('mTabs').innerHTML = TABS.map((t) =>
      `<button data-t="${t.key}" class="${t.key === state.modalTab ? 'active' : ''}">${t.label}</button>`).join('');
    $('mTabs').querySelectorAll('button').forEach((b) => {
      b.onclick = () => { state.modalTab = b.dataset.t; openModal(name); };
    });
    $('overlay').hidden = false;
    drawModal();
  }

  function dimValues(rows, key) {
    const n = rows.length || 1;
    return rows.reduce((a, r) => a + DIM[key].get(r), 0) / n;
  }

  function drawModal() {
    const me = state.rows.find((r) => r.name === state.modalPlayer);
    if (!me) return;
    const mates = me.school ? state.rows.filter((r) => r.school === me.school) : state.rows;
    const dims = tabDims(state.modalTab);
    if (!dims.length) { $('mTable').innerHTML = '<p style="color:#9aa5bd">没有可用的维度数据。</p>'; return; }

    const base = {};   // 全队人均作为 100 基准
    dims.forEach((d) => { base[d] = dimValues(state.rows, d) || 1; });
    const val = (rows, d) => +(dimValues(rows, d) / base[d] * 100).toFixed(1);

    const meVals = dims.map((d) => val([me], d));
    const mateVals = dims.map((d) => val(mates, d));
    const teamVals = dims.map(() => 100);
    const pool = meVals.concat(mateVals);
    const max = niceMax(Math.max.apply(null, pool) * 1.1);

    if (modalChart) { try { modalChart.dispose(); } catch (e) {} }
    modalChart = echarts.init($('mRadar'));
    modalChart.setOption({
      tooltip: { trigger: 'item', backgroundColor: '#26314b', borderWidth: 0,
        textStyle: { color: '#fff', fontSize: 12 },
        formatter: (p) => {
          let s = '<b>' + esc(p.name) + '</b><br/>';
          dims.forEach((d, i) => { s += DIM[d].label + '：' + p.value[i] + '（均值=100）<br/>'; });
          return s;
        } },
      legend: { bottom: 0, icon: 'roundRect', itemWidth: 10, itemHeight: 10,
        textStyle: { color: '#6b7794', fontSize: 12 } },
      radar: {
        center: ['50%', '47%'], radius: '62%',
        indicator: dims.map((d) => ({ name: DIM[d].label, max })),
        axisName: { color: '#38445f', fontSize: 12, fontWeight: 600 },
        splitLine: { lineStyle: { color: '#e5eaf2' } },
        splitArea: { areaStyle: { color: ['#fbfcfe', '#ffffff'] } },
        axisLine: { lineStyle: { color: '#e5eaf2' } }
      },
      series: [{
        type: 'radar', symbolSize: 4,
        data: [
          { name: me.name, value: meVals,
            lineStyle: { width: 2.2, color: '#2f7cf6' }, itemStyle: { color: '#2f7cf6' },
            areaStyle: { color: 'rgba(47,124,246,.22)' } },
          { name: (me.school || '门派') + '均值', value: mateVals,
            lineStyle: { width: 1.8, color: '#ff9f43' }, itemStyle: { color: '#ff9f43' },
            areaStyle: { color: 'rgba(0,0,0,0)' } },
          { name: '全队均值', value: teamVals,
            lineStyle: { width: 1.4, type: 'dashed', color: '#b9c3d6' },
            itemStyle: { color: '#b9c3d6' }, symbol: 'none' }
        ]
      }]
    });

    const teamTotal = (d) => state.rows.reduce((a, r) => a + DIM[d].get(r), 0);
    $('mTable').innerHTML =
      '<div class="mt-head"><span>数据维度</span><span style="text-align:right">本人</span>' +
      '<span style="text-align:right">占比</span><span style="text-align:right">门派均值</span>' +
      '<span style="text-align:right">全队均值</span><span style="text-align:right">对比门派</span></div>' +
      dims.map((d) => {
        const mine = DIM[d].get(me);
        const mAve = dimValues(mates, d);
        const tAve = dimValues(state.rows, d);
        const tot = teamTotal(d);
        const share = (d === 'survive' || d === 'kda' || d === 'perDeathDmg' || d === 'perDeathTank')
          ? '—' : (tot ? (mine / tot * 100).toFixed(2) + '%' : '—');
        const diff = mAve ? (mine / mAve - 1) : 0;
        const cls = Math.abs(diff) < 0.02 ? 'mid' : (diff > 0 ? 'up' : 'down');
        return '<div class="mt-row">' +
          '<span class="k">' + DIM[d].label + '</span>' +
          '<span class="v">' + fmtVal(d, mine) + '</span>' +
          '<span class="m">' + share + '</span>' +
          '<span class="m">' + fmtVal(d, mAve) + '</span>' +
          '<span class="m">' + fmtVal(d, tAve) + '</span>' +
          '<span class="' + cls + '">' + (diff > 0 ? '+' : '') + (diff * 100).toFixed(0) + '%</span>' +
          '</div>';
      }).join('');
    $('mNote').textContent = '口径：雷达数值以「全队人均」为 100 基准，越高越强；对比门派 = 与同门派人均的差异。';
  }

  function niceMax(v) {
    if (!(v > 0)) return 1;
    const p = Math.pow(10, Math.floor(Math.log10(v)));
    const step = p / 2;
    return Math.ceil(v / step) * step;
  }

  /* ================= 视图二：联赛对比 ================= */
  function fillGuildSelect() {
    const guilds = [...new Set(state.rows.map((r) => r.guild).filter(Boolean))];
    if (!guilds.length) { $('myGuild').innerHTML = '<option>（无帮会列）</option>'; return; }
    if (guilds.indexOf(state.myGuild) < 0) {
      state.myGuild = topGuild() || guilds[0];
    }
    $('myGuild').innerHTML = guilds.map((g) =>
      `<option value="${esc(g)}" ${g === state.myGuild ? 'selected' : ''}>${esc(g)}</option>`).join('');
  }

  function leagueGroups() {
    const guilds = [...new Set(state.rows.map((r) => r.guild).filter(Boolean))];
    if (guilds.length >= 2) {
      const mine = state.myGuild && guilds.indexOf(state.myGuild) >= 0 ? state.myGuild : guilds[0];
      const other = guilds.find((g) => g !== mine);
      return {
        a: state.rows.filter((r) => r.guild === mine), aName: mine,
        b: state.rows.filter((r) => r.guild === other), bName: other
      };
    }
    // 没有帮会列 → 按“门派排序前后一半”分，纯粹给个对比兜底
    const half = Math.ceil(state.rows.length / 2);
    return { a: state.rows.slice(0, half), aName: '我方', b: state.rows.slice(half), bName: '敌方' };
  }

  function sumOf(rows, k) { return rows.reduce((a, r) => a + (r[k] || 0), 0); }
  function avgOf(rows, key) { return rows.length ? rows.reduce((a, r) => a + DIM[key].get(r), 0) / rows.length : 0; }

  function renderLeague() {
    const G = leagueGroups();
    $('pieATitle').textContent = G.aName + ' 阵容人数';
    $('pieBTitle').textContent = G.bName + ' 阵容人数';
    $('leagueCount').textContent = `${G.aName} ${G.a.length} 人 · ${G.bName} ${G.b.length} 人`;

    const cards = [
      { k: '总击杀', key: 'kill', f: fmtInt },
      { k: '总人伤', key: 'damage', f: formatNum },
      { k: '总塔伤', key: 'tower', f: formatNum },
      { k: '总重伤', key: 'death', f: fmtInt },
      { k: '总承伤', key: 'tank', f: formatNum },
      { k: '总治疗', key: 'heal', f: formatNum },
      { k: 'KDA', key: 'kda', f: (v) => v.toFixed(2) },
      { k: '每死承伤', key: 'perDeathTank', f: formatNum }
    ].filter((c) => sumOf(state.rows, c.key === 'kda' || c.key === 'perDeathTank' ? 'kill' : c.key) ||
      (state.totals[c.key] || 0) > 0);

    $('vsKpis').innerHTML = cards.slice(0, 4).map((c) => {
      const a = c.key === 'kda' ? avgOf(G.a, 'kda') : (c.key === 'perDeathTank' ? avgOf(G.a, 'perDeathTank') : sumOf(G.a, c.key));
      const b = c.key === 'kda' ? avgOf(G.b, 'kda') : (c.key === 'perDeathTank' ? avgOf(G.b, 'perDeathTank') : sumOf(G.b, c.key));
      return `<div class="vkpi"><div class="k">${c.k}</div>
        <div class="row"><span class="a">${c.f(a)}</span>
        <span class="vs">VS</span><span class="b">${c.f(b)}</span></div></div>`;
    }).join('');

    // 团队能力雷达
    const dims = tabDims('all');
    const base = {}; dims.forEach((d) => { base[d] = dimValues(state.rows, d) || 1; });
    const aV = dims.map((d) => +(dimValues(G.a, d) / base[d] * 100).toFixed(1));
    const bV = dims.map((d) => +(dimValues(G.b, d) / base[d] * 100).toFixed(1));
    const max = niceMax(Math.max.apply(null, aV.concat(bV)) * 1.1);

    const chR = initChart('vsRadar');
    if (chR) chR.setOption({
      tooltip: { trigger: 'item', backgroundColor: '#26314b', borderWidth: 0, textStyle: { color: '#fff', fontSize: 12 } },
      legend: { bottom: 0, icon: 'roundRect', itemWidth: 10, itemHeight: 10, textStyle: { color: '#6b7794', fontSize: 12 } },
      radar: {
        center: ['50%', '46%'], radius: '64%',
        indicator: dims.map((d) => ({ name: DIM[d].label, max })),
        axisName: { color: '#38445f', fontSize: 12, fontWeight: 600 },
        splitLine: { lineStyle: { color: '#e5eaf2' } },
        splitArea: { areaStyle: { color: ['#fbfcfe', '#ffffff'] } },
        axisLine: { lineStyle: { color: '#e5eaf2' } }
      },
      series: [{ type: 'radar', symbolSize: 4, data: [
        { name: G.aName, value: aV, lineStyle: { width: 2, color: '#2f7cf6' },
          itemStyle: { color: '#2f7cf6' }, areaStyle: { color: 'rgba(47,124,246,.2)' } },
        { name: G.bName, value: bV, lineStyle: { width: 2, color: '#f2676b' },
          itemStyle: { color: '#f2676b' }, areaStyle: { color: 'rgba(242,103,107,.16)' } }
      ] }]
    }, true);

    pieChart('vsPieA', G.a);
    pieChart('vsPieB', G.b);

    const unit = state.aggMode === 'total' ? '总量' : '人均';
    setTitle('chDmg', '各门派 人伤产出（' + unit + '）');
    setTitle('chKda', '各门派 KDA');
    setTitle('chDeath', '各门派 重伤人数（' + unit + '）');
    setTitle('chPerDeath', '各门派 每死输出（总人伤÷总重伤）');
    setTitle('chTank', '各门派 承伤（' + unit + '）');
    setTitle('chTower', '各门派 塔伤（' + unit + '）');

    groupBar('vsDmg', G, 'damage');
    groupBar('vsKda', G, 'kda');
    groupBar('vsDeath', G, 'death');
    groupBar('vsPerDeath', G, 'perDeathDmg');
    groupBar('vsTank', G, 'tank');
    groupBar('vsTower', G, 'tower');
  }

  function setTitle(id, text) {
    const el = document.getElementById(id);
    if (el) el.textContent = text;
  }

  function pieChart(id, rows) {
    const ch = initChart(id); if (!ch) return;
    const agg = {};
    rows.forEach((r) => { const k = r.school || '未填写'; agg[k] = (agg[k] || 0) + 1; });
    const data = Object.keys(agg).map((k) => ({ name: k + ' ' + agg[k], value: agg[k], itemStyle: { color: colorOf(k) } }))
      .sort((a, b) => b.value - a.value);
    ch.setOption({
      tooltip: { trigger: 'item', backgroundColor: '#26314b', borderWidth: 0,
        textStyle: { color: '#fff', fontSize: 12 }, formatter: (p) => p.name + '：' + p.value + ' 人（' + p.percent + '%）' },
      legend: { bottom: 0, left: 'center', type: 'scroll', icon: 'circle',
        itemWidth: 8, itemHeight: 8, itemGap: 6, pageIconSize: 9, pageTextStyle: { fontSize: 10 },
        textStyle: { color: '#6b7794', fontSize: 10 } },
      series: [{ type: 'pie', radius: ['46%', '66%'], center: ['50%', '42%'], data,
        label: { show: false }, labelLine: { show: false },
        emphasis: { scale: true, scaleSize: 6 },
        itemStyle: { borderColor: '#fff', borderWidth: 2 } }]
    }, true);
  }

  const AGG_KEYS = ['damage', 'tower', 'tank', 'heal', 'kill', 'assist', 'death', 'score'];
  /* 按门派汇总：kda/每死类 = 总量分子÷总量分母；其余按口径取总量或人均 */
  function schoolStat(rows, metric, mode) {
    const g = {};
    rows.forEach((r) => {
      const k = r.school || '未填写';
      const a = g[k] || (g[k] = { n: 0, damage: 0, tower: 0, tank: 0, heal: 0, kill: 0, assist: 0, death: 0, score: 0 });
      a.n++;
      AGG_KEYS.forEach((x) => { a[x] += r[x] || 0; });
    });
    const out = {};
    Object.keys(g).forEach((k) => {
      const a = g[k];
      let v;
      if (metric === 'kda') v = (a.kill + a.assist) / Math.max(1, a.death);
      else if (metric === 'perDeathDmg') v = a.damage / Math.max(1, a.death);
      else if (metric === 'perDeathTank') v = a.tank / Math.max(1, a.death);
      else { v = a[metric] || 0; if (mode !== 'total') v = v / Math.max(1, a.n); }
      out[k] = v;
    });
    return out;
  }

  function groupBar(id, G, key) {
    const ch = initChart(id); if (!ch) return;
    if (!state.rows.some((r) => DIM[key].get(r) > 0)) {
      ch.clear();
      ch.setOption({ title: { text: '无数据', left: 'center', top: 'middle',
        textStyle: { color: '#c3cbdb', fontSize: 13, fontWeight: 'normal' } } });
      return;
    }
    const isRatio = (key === 'kda' || key === 'perDeathDmg' || key === 'perDeathTank');
    const mode = isRatio ? 'total' : state.aggMode;
    const A = schoolStat(G.a, key, mode), B = schoolStat(G.b, key, mode);
    const cats = [...new Set(Object.keys(A).concat(Object.keys(B)))];
    const mk = (src) => cats.map((c) => (c in src ? +src[c].toFixed(2) : 0));
    const fmt = (v) => (key === 'kda' ? Number(v).toFixed(2) : formatNum(v));

    const w = (ch.getWidth ? ch.getWidth() : 600) || 600;
    const rotate = w < cats.length * 34 ? 30 : 0;

    ch.setOption({
      grid: { left: 6, right: 16, top: 30, bottom: rotate ? 22 : 4, containLabel: true },
      tooltip: { trigger: 'axis', axisPointer: { type: 'shadow' }, backgroundColor: '#26314b',
        borderWidth: 0, textStyle: { color: '#fff', fontSize: 12 },
        formatter: (p) => p.map((x) => x.marker + x.seriesName + '：' + fmt(x.value)).join('<br/>') },
      legend: { top: 0, left: 'center', icon: 'roundRect', itemWidth: 10, itemHeight: 10,
        textStyle: { color: '#6b7794', fontSize: 11 } },
      xAxis: { type: 'category', data: cats, axisTick: { show: false },
        axisLine: { lineStyle: { color: '#e5eaf2' } },
        axisLabel: { color: '#6b7794', fontSize: 11, interval: 0, rotate, margin: 8 } },
      yAxis: { type: 'value', axisLine: { show: false },
        splitLine: { lineStyle: { color: '#eef2f8' } },
        axisLabel: { color: '#9aa5bd', fontSize: 10, formatter: (v) => (key === 'kda' ? Number(v).toFixed(0) : formatNum(v)) } },
      series: [
        { name: G.aName, type: 'bar', barMaxWidth: rotate ? 9 : 18, data: mk(A),
          itemStyle: { color: '#2f7cf6', borderRadius: [3, 3, 0, 0] } },
        { name: G.bName, type: 'bar', barMaxWidth: rotate ? 9 : 18, data: mk(B),
          itemStyle: { color: '#f2676b', borderRadius: [3, 3, 0, 0] } }
      ]
    }, true);
  }

  function initChart(id) {
    const el = $(id);
    if (!el || !el.clientWidth) return null;
    if (!charts[id]) charts[id] = echarts.init(el);
    return charts[id];
  }

  /* ================= 导出 ================= */
  function exportResult() {
    if (!state.rows.length) return;
    const out = state.rows.map((r) => {
      const o = { '角色名': r.name, '门派': r.school, '帮会': r.guild };
      ['damage', 'tower', 'tank', 'kill', 'assist', 'death', 'heal', 'score'].forEach((k) => {
        if (state.mapping[k] !== '' && state.mapping[k] != null) o[FIELDS.find((f) => f.key === k).label] = Math.round(r[k]);
      });
      o['KDA'] = +((DIM.kda.get(r)).toFixed(2));
      o['每死输出'] = Math.round(DIM.perDeathDmg.get(r));
      o['每死承伤'] = Math.round(DIM.perDeathTank.get(r));
      o['伤害占比%'] = +((r.dmgPct * 100).toFixed(2));
      return o;
    });
    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, XLSX.utils.json_to_sheet(out), '帮战分析');
    const d = new Date();
    const st = `${d.getFullYear()}${String(d.getMonth() + 1).padStart(2, '0')}${String(d.getDate()).padStart(2, '0')}`;
    XLSX.writeFile(wb, `帮战数据分析_${st}.xlsx`);
  }

  /* ================= 示例数据（仿逆水寒帮会战报原生导出格式） ================= */
  function loadDemo() {
    const schools = ['铁衣','素问','神相','碎梦','玄机','血河','九灵','龙吟','沧澜','潮光','鸿音'];
    const sur = '赵钱孙李周吴郑王冯陈褚卫蒋沈韩杨朱秦尤许何吕施张孔曹严华金魏陶姜';
    const giv = ['无双','青崖','山月','知秋','不渡','归尘','听雪','长歌','凤鸣','逐星','惊鸿','半盏','浮生','夜白','南栀','轻语'];
    const HEAD = ['玩家名字','职业','击败/清泉','助攻','资源','对玩家伤害','人伤卸甲',
      '对建筑伤害','破塔卸甲','治疗值','承受伤害','重伤','复活/清泉','焚骨'];
    const aoa = [];
    let k = 0;
    [['示例帮会甲', 55], ['示例帮会乙', 60]].forEach(([guild, n]) => {
      aoa.push([guild, n]);
      aoa.push(HEAD.slice());
      for (let i = 0; i < n; i++, k++) {
        const school = schools[k % schools.length];
        const isHeal = school === '素问', isTank = school === '铁衣' || school === '血河';
        const hongHeal = school === '鸿音' && Math.random() < 0.5;   // 一半鸿音当奶鸿
        const healer = isHeal || hongHeal;
        const kill = Math.round(Math.random() * (isHeal ? 3 : 20));
        const death = Math.round(Math.random() * 12);
        const assist = Math.round(Math.random() * (isHeal ? 200 : 120));
        const damage = healer ? Math.round(Math.random() * 5e5) : Math.round((isTank ? 3e6 : 9e6) + Math.random() * 2.4e7);
        const tower = healer ? Math.round(Math.random() * 8e5) : Math.round(Math.random() * (isTank ? 4e6 : 3.4e7));
        const tank = isTank ? Math.round(8e6 + Math.random() * 2.2e7) : Math.round(Math.random() * 1.2e7);
        const heal = healer ? Math.round(1.2e7 + Math.random() * 1.4e7) : Math.round(Math.random() * 5e5);
        aoa.push([
          sur[k % sur.length] + giv[k % giv.length], school,
          ' ' + kill + '/' + Math.round(Math.random() * 6), assist, 0,
          damage, 0, tower, 0, heal, tank, death,
          Math.round(Math.random() * 18), Math.round(Math.random() * 60 * (Math.random() < .4 ? 1 : 0))
        ]);
      }
    });
    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, XLSX.utils.aoa_to_sheet(aoa), '战报');
    state.workbook = wb; state.sheetNames = wb.SheetNames; state.sheetIndex = 0;
    state.myGuild = null;
    $('srcInfo').textContent = '示例数据（仿帮会战报导出） · ' + new Date().toLocaleString('zh-CN');
    loadSheet(0);
  }

  /* ================= 视图切换 / 事件 ================= */
  function switchView(v) {
    state.view = v;
    $('view-members').hidden = v !== 'members';
    $('view-league').hidden = v !== 'league';
    $('nav').querySelectorAll('.nav-btn').forEach((b) => b.classList.toggle('active', b.dataset.view === v));
    if (v === 'league' && state.rows.length) renderLeague();
  }

  function bind() {
    $('btnImport').onclick = () => $('fileInput').click();
    $('btnPick').onclick = () => $('fileInput').click();
    $('fileInput').onchange = (e) => { if (e.target.files[0]) handleFile(e.target.files[0]); };
    $('btnDemo').onclick = loadDemo;
    $('btnExport').onclick = exportResult;
    $('mClose').onclick = () => { $('overlay').hidden = true; if (modalChart) { modalChart.dispose(); modalChart = null; } };
    $('overlay').onclick = (e) => { if (e.target === $('overlay')) $('mClose').click(); };
    document.addEventListener('keydown', (e) => { if (e.key === 'Escape' && !$('overlay').hidden) $('mClose').click(); });

    $('nav').querySelectorAll('.nav-btn').forEach((b) => { b.onclick = () => switchView(b.dataset.view); });
    $('btnMap').onclick = () => {
      const p = $('mapPanel');
      p.hidden = !p.hidden;
      if (!p.hidden) renderMapping();
    };
    $('btnApplyMap').onclick = () => { analyze(); if (state.rows.length) $('mapPanel').hidden = true; };
    $('search').oninput = renderMembers;
    $('filterSchool').onchange = renderMembers;
    $('filterGuild').onchange = renderMembers;
    $('myGuild').onchange = () => { state.myGuild = $('myGuild').value; renderLeague(); };
    $('aggMode').onchange = () => { state.aggMode = $('aggMode').value; renderLeague(); };

    // 整页拖拽
    ['dragenter', 'dragover'].forEach((ev) => document.addEventListener(ev, (e) => {
      e.preventDefault();
      const box = document.querySelector('.empty-box'); if (box) box.classList.add('drag');
    }));
    document.addEventListener('dragleave', (e) => {
      const box = document.querySelector('.empty-box'); if (box && e.target === document) box.classList.remove('drag');
    });
    document.addEventListener('drop', (e) => {
      e.preventDefault();
      const box = document.querySelector('.empty-box'); if (box) box.classList.remove('drag');
      const f = e.dataTransfer && e.dataTransfer.files[0];
      if (f) handleFile(f);
    });

    let t = null;
    const relayout = () => { clearTimeout(t); t = setTimeout(() => {
      Object.values(charts).forEach((c) => { try { c.resize(); } catch (e) {} });
      if (modalChart) { try { modalChart.resize(); } catch (e) {} }
    }, 140); };
    window.addEventListener('resize', relayout);
  }

  bind();
  if (location.search.indexOf('demo') >= 0) {
    loadDemo();
    if (location.search.indexOf('league') >= 0) switchView('league');
    else if (location.search.indexOf('modal') >= 0) openModal(state.rows[0].name);
  }
})();
