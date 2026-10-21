/* ===================== 学生作业与考试成绩登记系统 ===================== */
(function () {
  "use strict";
  const $ = (id) => document.getElementById(id);
  const STORE_KEY = "scoreSystem.v1";

  /* ---------- 状态与持久化（纯静态版：仅本地存储，无后端服务器） ---------- */
  let state = { students: [], items: [], scores: {} };
  let currentItemId = null;
  let voice = { active: false, currentItem: null };

  // 纯静态部署（如 GitHub Pages）没有后端，数据保存在访问设备的浏览器 localStorage 中。
  // 跨设备/跨网络同步请用「导出 JSON / 导入 JSON」按钮搬运数据。
  async function load() {
    try {
      const raw = localStorage.getItem(STORE_KEY);
      if (raw) state = JSON.parse(raw);
    } catch (e) {}
    state.students = state.students || [];
    state.items = state.items || [];
    state.scores = state.scores || {};
    try { localStorage.setItem(STORE_KEY, JSON.stringify(state)); } catch (e) {}
  }
  function save() {
    try { localStorage.setItem(STORE_KEY, JSON.stringify(state)); flashSaved(); } catch (e) {}
  }
  let savedTimer = null;
  function flashSaved() {
    const s = $("saveState");
    s.textContent = "已保存 " + new Date().toLocaleTimeString("zh-CN");
    clearTimeout(savedTimer);
    savedTimer = setTimeout(() => (s.textContent = "已保存"), 1500);
  }
  function uid() { return Date.now().toString(36) + Math.random().toString(36).slice(2, 6); }

  /* ---------- 工具 ---------- */
  function escapeXml(s) {
    return String(s).replace(/[&<>"']/g, (c) => ({ "&": "&amp;", "<": "&lt;", ">": "&gt;", '"': "&quot;", "'": "&apos;" }[c]));
  }
  function trunc(s, n) { s = String(s); return s.length > n ? s.slice(0, n) + "…" : s; }
  function toast(msg, type) {
    const t = $("toast"); t.textContent = msg; t.className = "toast";
    clearTimeout(t._t); t._t = setTimeout(() => t.classList.add("hidden"), 2600);
  }
  const NAME_FILLERS = ["分", "的", "同学", "让", "给", "帮", "登记", "成绩", "录入", "一下",
    "这个", "那个", "把", "到", "是", "了", "我要", "我想", "请", "帮我", "记", "上", "们"];
  function normalizeName(s) {
    s = String(s).replace(/[\s,，。.!！?？、~～"'’”()（）:：]/g, "");
    for (const f of NAME_FILLERS) s = s.split(f).join("");
    return s.trim();
  }
  function normalizeItem(s) {
    s = String(s).replace(/[\s,，。.!！?？、~～"'’”()（）:：]/g, "");
    for (const f of ["登记", "录入", "成绩", "把", "给", "帮", "我要", "我想", "请", "帮我", "的分数", "分", "的"]) s = s.split(f).join("");
    return s.trim();
  }
  function lcsLen(a, b) {
    const n = a.length, m = b.length;
    const dp = Array.from({ length: n + 1 }, () => new Array(m + 1).fill(0));
    for (let i = 1; i <= n; i++)
      for (let j = 1; j <= m; j++)
        dp[i][j] = a[i - 1] === b[j - 1] ? dp[i - 1][j - 1] + 1 : Math.max(dp[i - 1][j], dp[i][j - 1]);
    return dp[n][m];
  }
  function cnToNumber(str) {
    const d = { "零": 0, "一": 1, "壹": 1, "二": 2, "贰": 2, "两": 2, "三": 3, "叁": 3,
      "四": 4, "肆": 4, "五": 5, "伍": 5, "六": 6, "陆": 6, "七": 7, "柒": 7,
      "八": 8, "捌": 8, "九": 9, "玖": 9 };
    const m = str.match(/[零一二两三四五六七八九十百千壹贰叁肆伍陆柒捌玖拾佰仟]+/);
    if (!m) return null;
    const s = m[0];
    let total = 0, section = 0, num = 0;
    for (const ch of s) {
      if (ch === "百" || ch === "佰") { section = (num || 1) * 100; total += section; num = 0; section = 0; }
      else if (ch === "千" || ch === "仟") { section = (num || 1) * 1000; total += section; num = 0; section = 0; }
      else if (ch === "十" || ch === "拾") { total += (num || 1) * 10; num = 0; }
      else { num = d[ch]; }
    }
    total += num;
    return { value: total, raw: s };
  }

  /* ---------- 名单导入 ---------- */
  function parseRoster(text) {
    const lines = text.split(/\r?\n/).map((s) => s.trim()).filter(Boolean);
    let added = 0;
    for (const line of lines) {
      const parts = line.split(/[\t,，]/).map((s) => s.trim()).filter(Boolean);
      let cls = "未分班", name;
      if (parts.length >= 2) { cls = parts[0]; name = parts.slice(1).join(""); }
      else { name = parts[0]; }
      if (!name) continue;
      if (state.students.some((s) => s.name === name && s.class === cls)) continue;
      state.students.push({ id: uid(), name, class: cls });
      added++;
    }
    return added;
  }

  /* ---------- Excel(.xlsx) 导入：零依赖自研解析 ---------- */
  function u16(dv, o) { return dv.getUint16(o, true); }
  function u32(dv, o) { return dv.getUint32(o, true); }
  async function inflateRaw(bytes) {
    if (typeof DecompressionStream === "undefined") throw new Error("no_deflate");
    const ds = new DecompressionStream("deflate-raw");
    const stream = new Response(bytes).body.pipeThrough(ds);
    const buf = await new Response(stream).arrayBuffer();
    return new Uint8Array(buf);
  }
  function colToIndex(col) {
    let n = 0;
    for (const ch of col) n = n * 26 + (ch.charCodeAt(0) - 64);
    return n - 1;
  }
  function isClassLike(x) { return /班|年级|级|小组|组别/.test(x) || /^\s*[\d一二三四五六七八九十]+\s*$/.test(x); }
  // 从文件名推导班级名：形如 高一三班/高三2班/七年级二班/一班/二班/2026级
  function classNameFromFilename(fn) {
    const base = String(fn || "").replace(/\.xlsx$/i, "").trim();
    if (!base) return "未分班";
    const m = base.match(/(?:高|初|中|职|大)?\s*(?:[0-9]+|[一二三四五六七八九十百千]+)?\s*(?:[0-9]+|[一二三四五六七八九十百千]+)\s*(?:班|年级|级)/);
    if (m) return m[0].replace(/\s+/g, "");
    const m2 = base.match(/[0-9]+\s*(?:年级|级)/);
    if (m2) return m2[0].replace(/\s+/g, "");
    const idx = base.search(/[班年级]/);
    if (idx > 0 && /[0-9高初中小职一二三四五六七八九十百千]/.test(base[idx - 1])) return base.slice(0, idx + 1);
    return base;
  }
  // 判断单元格是否像「人名」（2~4 个汉字，可带一个间隔号；排除班级/性别/备注/电话/表头等）
  function isNameLike(x) {
    const v = String(x || "").trim();
    if (!v) return false;
    const low = v.toLowerCase();
    if (/班|年级|级|小组|性别|男|女|备注|电话|家长|联系|名单|花名册|表|姓名|名字|学生|name|student|学号|id|序号|编号|考号|成绩|分数/.test(low)) return false;
    return /^[一-龥]{2,4}(·[一-龥]{1,2})?$/.test(v);
  }
  // 判断单元格是否为表头词（整行都是表头词时整行跳过）
  function isHeaderCell(c) {
    const v = String(c || "").trim().toLowerCase();
    return /^(姓名|名字|名称|学生姓名|学生|学子|name|student|学号|考号|学籍|序号|编号|id|班级|年级|班|级|小组|组别|性别|男|女|备注|电话|家长|家长电话|联系方式)$/.test(v);
  }
  // 解析 xlsx 行 -> [{cls, name, sid}]；班级用传入的 className（来自文件名），
  // 仅提取「姓名」与「学号」，自动跳过表头行与无关列，避免把“姓名/性别”等无关信息导成学生。
  function xlsxRowsToLines(rows, className) {
    if (!rows.length) return [];
    // 1) 扫描表头行（含 姓名/name 或 学号/id 关键词），确定姓名列与学号列索引，并跳过该表头及之前的标题行
    let nameCol = -1, sidCol = -1;
    for (let r = 0; r < rows.length; r++) {
      const lower = rows[r].map((c) => String(c || "").trim().toLowerCase());
      const hasNameH = lower.some((h) => /^(姓名|名字|名称|学生姓名|name|student)$/.test(h));
      const hasIdH = lower.some((h) => /^(学号|考号|学籍|序号|编号|id)$/.test(h));
      if (hasNameH || hasIdH) {
        nameCol = lower.findIndex((h) => /^(姓名|名字|名称|学生姓名|name|student)$/.test(h));
        sidCol = lower.findIndex((h) => /^(学号|考号|学籍|序号|编号|id)$/.test(h));
        rows = rows.slice(r + 1);
        break;
      }
    }
    const out = [];
    for (const row of rows) {
      const cells = row.map((c) => String(c == null ? "" : c).trim());
      if (!cells.some(Boolean)) continue;
      // 整行都是表头词 -> 跳过（防止多行表头/残留表头混入）
      if (cells.every((c) => isHeaderCell(c))) continue;
      let name = "", sid = "";
      if (nameCol >= 0 && cells[nameCol] != null) name = cells[nameCol];
      if (sidCol >= 0 && cells[sidCol] != null) sid = cells[sidCol];
      if (!name) {
        const ni = cells.findIndex((c) => isNameLike(c));
        if (ni >= 0) name = cells[ni];
      }
      if (!sid) {
        const idc = cells.findIndex((c) => /^\d{4,}$/.test(c));
        if (idc >= 0) sid = cells[idc];
      }
      name = (name || "").trim();
      // 只接受像人名的单元格；挡掉表头词、纯数字、班级/性别等无关内容
      if (!name || !isNameLike(name)) continue;
      sid = (sid || "").trim();
      out.push({ cls: className, name, sid });
    }
    return out;
  }
  // 按 [{cls, name, sid}] 批量加入学生（按 姓名+班级 去重）
  function importStudents(list) {
    let added = 0, withSid = 0;
    for (const it of list) {
      const cls = (it.cls || "未分班").trim() || "未分班";
      const name = (it.name || "").trim();
      const sid = (it.sid || "").trim();
      if (!name) continue;
      if (state.students.some((s) => s.name === name && s.class === cls)) continue;
      state.students.push({ id: uid(), name, class: cls, sid });
      added++;
      if (sid) withSid++;
    }
    return { added, withSid };
  }
  // 解析 .xlsx（仅首个工作表），返回 2D 字符串数组
  async function parseXlsx(arrayBuffer) {
    const dv = new DataView(arrayBuffer);
    const bytes = new Uint8Array(arrayBuffer);
    const len = bytes.length;
    // 1) 定位 EOCD
    let eocd = -1;
    const minPos = Math.max(0, len - 65557);
    for (let i = len - 22; i >= minPos; i--) {
      if (dv.getUint32(i, true) === 0x06054b50) { eocd = i; break; }
    }
    if (eocd < 0) throw new Error("非有效的 xlsx(zip) 文件");
    const cdCount = u16(dv, eocd + 10);
    const cdOffset = u32(dv, eocd + 16);
    // 2) 遍历中央目录
    const entries = {};
    let p = cdOffset;
    for (let i = 0; i < cdCount; i++) {
      if (dv.getUint32(p, true) !== 0x02014b50) break;
      const comp = u16(dv, p + 10);
      const compSize = u32(dv, p + 20);
      const fnLen = u16(dv, p + 28);
      const exLen = u16(dv, p + 30);
      const cmLen = u16(dv, p + 32);
      const name = new TextDecoder().decode(bytes.subarray(p + 46, p + 46 + fnLen));
      const lho = u32(dv, p + 42);
      entries[name] = { offset: lho, comp, compSize };
      p += 46 + fnLen + exLen + cmLen;
    }
    // 3) 抽取 sharedStrings 与首个工作表
    let sheetName = null;
    for (const k of Object.keys(entries)) if (/^xl\/worksheets\/sheet\d+\.xml$/.test(k)) { sheetName = k; break; }
    if (!sheetName) throw new Error("未找到工作表");
    const need = ["xl/sharedStrings.xml", sheetName];
    const files = {};
    for (const k of need) {
      const e = entries[k]; if (!e) continue;
      const lh = e.offset;
      const lfnLen = u16(dv, lh + 26);
      const lexLen = u16(dv, lh + 28);
      const dataStart = lh + 30 + lfnLen + lexLen;
      let data = bytes.subarray(dataStart, dataStart + e.compSize);
      if (e.comp === 8) data = await inflateRaw(data);
      files[k] = new TextDecoder("utf-8").decode(data);
    }
    // 4) 共享字符串表
    const shared = [];
    const sst = files["xl/sharedStrings.xml"];
    if (sst) {
      const siRe = /<si>([\s\S]*?)<\/si>/g; let m;
      while ((m = siRe.exec(sst))) {
        let t = "", tm; const tRe = /<t[^>]*>([\s\S]*?)<\/t>/g;
        while ((tm = tRe.exec(m[1]))) t += tm[1];
        shared.push(t);
      }
    }
    // 5) 工作表 -> 行/列
    const sheetXml = files[sheetName] || "";
    const rows = [];
    const rowRe = /<row[^>]*>([\s\S]*?)<\/row>/g; let rm;
    while ((rm = rowRe.exec(sheetXml))) {
      const rowInner = rm[1];
      const cells = {};
      const cRe = /<c\s+([^>]*?)\/>|<c\s+([\s\S]*?)<\/c>/g; let cm;
      while ((cm = cRe.exec(rowInner))) {
        const attr = cm[1] || cm[2] || "";
        const rM = /r="([A-Z]+\d+)"/.exec(attr);
        if (!rM) continue;
        const col = colToIndex(rM[1].replace(/\d+$/, ""));
        const tM = /t="([^"]*)"/.exec(attr);
        const type = tM ? tM[1] : "";
        let val = "";
        if (type === "s") {
          const vM = /<v>(\d+)<\/v>/.exec(cm[0]);
          if (vM) val = shared[+vM[1]] || "";
        } else if (type === "inlineStr") {
          const tM2 = /<t[^>]*>([\s\S]*?)<\/t>/g; let t2;
          while ((t2 = tM2.exec(cm[0]))) val += t2[1];
        } else {
          const vM = /<v>([\s\S]*?)<\/v>/.exec(cm[0]);
          if (vM) val = vM[1];
        }
        cells[col] = val;
      }
      const maxCol = Object.keys(cells).length ? Math.max(...Object.keys(cells).map(Number)) : -1;
      const arr = [];
      for (let c = 0; c <= maxCol; c++) arr.push(cells[c] || "");
      rows.push(arr);
    }
    return rows;
  }
  function renderRoster() {
    const classes = {};
    state.students.forEach((s) => { (classes[s.class] = classes[s.class] || []).push(s); });
    const box = $("rosterList");
    box.innerHTML = "";
    const keys = Object.keys(classes);
    if (!keys.length) { box.innerHTML = '<div class="empty">暂无学生，请先导入名单</div>'; }
    keys.forEach((c) => {
      const div = document.createElement("div");
      div.className = "roster-class";
      const h = document.createElement("h4"); h.textContent = c + "（" + classes[c].length + "人）";
      div.appendChild(h);
      classes[c].forEach((s) => {
        const chip = document.createElement("span");
        chip.className = "stu-chip";
        const nm = document.createElement("span"); nm.textContent = s.sid ? s.name + " · " + s.sid : s.name;
        const del = document.createElement("button"); del.textContent = "×";
        del.onclick = () => {
          if (!confirm("删除学生「" + s.name + "」？")) return;
          state.students = state.students.filter((x) => x.id !== s.id);
          for (const k in state.scores) if (state.scores[k]) delete state.scores[k][s.id];
          save(); renderRoster(); renderAllSelectors(); renderGrid();
        };
        chip.appendChild(nm); chip.appendChild(del); div.appendChild(chip);
      });
      box.appendChild(div);
    });
    const total = state.students.length;
    const clsCount = keys.length;
    $("rosterStats").textContent = total + " 名学生 · " + clsCount + " 个班级";
  }

  /* ---------- 项目（作业/考试） ---------- */
  function ensureItemSmart(name) {
    const q = normalizeItem(name);
    if (!q) return null;
    let best = null, bestScore = 0;
    for (const it of state.items) {
      const sc = lcsLen(q, normalizeItem(it.name)) / Math.max(q.length, it.name.length);
      if (sc > bestScore) { bestScore = sc; best = it; }
    }
    if (best && bestScore >= 0.6) return best;
    const it = { id: uid(), name: q, type: "作业", unit: "", date: today() };
    state.items.push(it);
    return it;
  }
  function getItem(id) { return state.items.find((i) => i.id === id); }
  function today() { const d = new Date(); return d.getFullYear() + "-" + String(d.getMonth() + 1).padStart(2, "0") + "-" + String(d.getDate()).padStart(2, "0"); }
  function renderItemSelect() {
    const sel = $("itemSelect");
    const cur = sel.value;
    sel.innerHTML = "";
    if (!state.items.length) { sel.innerHTML = '<option value="">（暂无，请新建）</option>'; }
    state.items.slice().sort((a, b) => (a.date || "").localeCompare(b.date || "")).forEach((it) => {
      const o = document.createElement("option"); o.value = it.id;
      o.textContent = it.name + (it.type ? "【" + it.type + "】" : "");
      sel.appendChild(o);
    });
    if (cur && getItem(cur)) sel.value = cur;
    else if (state.items.length) { sel.value = state.items[0].id; }
    currentItemId = sel.value || null;
    renderItemMeta();
  }
  function renderItemMeta() {
    const it = getItem(currentItemId);
    const box = $("itemMeta");
    if (!it) { box.innerHTML = '<span class="tag">未选择项目</span>'; return; }
    box.innerHTML = '当前：<b>' + escapeXml(it.name) + '</b> ｜ 类型：' + (it.type || "—") +
      " ｜ 单元：" + (it.unit || "—") + " ｜ 日期：" + (it.date || "—");
  }

  /* ---------- 成绩矩阵 ---------- */
  function getScore(itemId, stuId) {
    return state.scores[itemId] && state.scores[itemId][stuId] != null ? state.scores[itemId][stuId] : null;
  }
  function setScore(itemId, stuId, val) {
    state.scores[itemId] = state.scores[itemId] || {};
    if (val == null || val === "") delete state.scores[itemId][stuId];
    else state.scores[itemId][stuId] = Number(val);
    save();
  }
  function renderGrid() {
    const wrap = $("grid");
    if (!state.students.length) { wrap.innerHTML = '<div class="empty">请先在「学生名单」导入学生</div>'; return; }
    const items = state.items.slice().sort((a, b) => (a.date || "").localeCompare(b.date || ""));
    let html = "<table><thead><tr><th>学生</th>";
    items.forEach((it) => { html += "<th>" + escapeXml(it.name) + (it.type ? "<br><small>" + it.type + "</small>" : "") + "</th>"; });
    html += "</tr></thead><tbody>";
    state.students.forEach((s) => {
      html += '<tr><td class="name">' + escapeXml(s.name) + (s.sid ? ' <small style="color:var(--muted)">#' + escapeXml(s.sid) + "</small>" : "") + ' <small style="color:var(--muted)">' + escapeXml(s.class) + "</small></td>";
      items.forEach((it) => {
        const v = getScore(it.id, s.id);
        html += '<td class="' + (v != null ? "filled" : "") + '"><input type="number" min="0" max="150" step="0.5" data-i="' + it.id + '" data-s="' + s.id + '" value="' + (v != null ? v : "") + '" /></td>';
      });
      html += "</tr>";
    });
    html += "</tbody></table>";
    wrap.innerHTML = html;
    wrap.querySelectorAll("input[data-i]").forEach((inp) => {
      inp.addEventListener("change", () => {
        setScore(inp.dataset.i, inp.dataset.s, inp.value === "" ? null : inp.value);
        inp.parentElement.className = inp.value === "" ? "" : "filled";
      });
    });
  }

  /* ---------- 语音解析 ---------- */
  // 判断是否为“作业/考试名称”声明（无数字且含关键词）
  function isItemText(text) {
    return !/\d/.test(text) && /作业|考试|测验|月考|期中|期末|练习|小测|测评|统考|单元|测试|课|quiz|test|exam/i.test(text);
  }
  // 核心：先用名单锚定姓名，再解析剩余部分为成绩（阿拉伯数字优先，其次中文数字）
  // 返回 {score, student, rest, conf}；student 为已锚定的学生对象（可能为 null）
  function parseScoreUtterance(text, students) {
    // 1) 阿拉伯数字优先
    const m = text.match(/(\d+(?:\.\d+)?)/);
    if (m) {
      const score = parseFloat(m[1]);
      const rest = normalizeName(text.replace(m[1], " "));
      const ex = students.find((s) => rest.includes(s.name));
      return { score, student: ex || null, rest, conf: ex ? 0.96 : null };
    }
    // 2) 名单精确锚定 + 中文数字成绩
    for (const s of students) {
      if (text.includes(s.name)) {
        const rest = text.replace(s.name, "");
        const cn = cnToNumber(rest);
        return { score: cn ? cn.value : null, student: s, rest: normalizeName(rest), conf: cn ? 0.96 : null };
      }
    }
    // 3) 模糊锚定 + 中文数字成绩
    const q = normalizeName(text);
    const fm = matchStudent(q, students);
    if (fm.match) {
      const rest = text.replace(fm.match.name, "");
      const cn = cnToNumber(rest);
      return { score: cn ? cn.value : null, student: fm.match, rest: normalizeName(rest), conf: cn ? Math.max(0.7, fm.bestScore) : null };
    }
    // 4) 整句为中文数字（无姓名，且整句就是数字，避免“第一单元测试”这类被误判）
    const cleaned = text.replace(/分/g, "").trim();
    const cn = cnToNumber(cleaned);
    if (cn && cn.raw === cleaned) return { score: cn.value, student: null, rest: "", conf: null };
    return { score: null, student: null, rest: q, conf: null };
  }
  function matchStudent(query, students) {
    const q = normalizeName(query);
    const cands = [];
    let best = null, bestScore = 0;
    if (!q) return { match: null, bestScore: 0, candidates: [] };
    for (const s of students) {
      const name = normalizeName(s.name);
      let sc;
      if (q.includes(name) || name.includes(q)) sc = 0.96;
      else {
        const l = lcsLen(q, name);
        sc = l / Math.max(q.length, name.length);
        if (q[0] === name[0]) sc += 0.2;
        if (Math.abs(q.length - name.length) <= 1) sc += 0.1;
      }
      sc = Math.min(sc, 1);
      cands.push({ student: s, score: sc });
      if (sc > bestScore) { bestScore = sc; best = s; }
    }
    cands.sort((a, b) => b.score - a.score);
    return { match: bestScore >= 0.45 ? best : null, bestScore, candidates: cands.slice(0, 3) };
  }
  function logLine(msg, cls) {
    const box = $("voiceLog");
    const d = document.createElement("div");
    if (cls) d.className = cls;
    d.textContent = "· " + msg;
    box.prepend(d);
  }
  function showCandidates(cands) {
    const box = $("candBox");
    box.innerHTML = ""; box.classList.remove("hidden");
    cands.forEach((c) => {
      const b = document.createElement("button");
      b.textContent = c.student.name + " (" + c.score.toFixed(2) + ")";
      b.onclick = () => {
        if (voice.currentItem && pendingScore != null) {
          setScore(voice.currentItem.id, c.student.id, pendingScore);
          logLine("✅ 手动确认 " + c.student.name + " ← " + pendingScore, "ok");
          renderGrid();
        }
        box.classList.add("hidden");
      };
      box.appendChild(b);
    });
  }
  let pendingScore = null;
  let pendingItem = null;
  function handleUtterance(text) {
    const t = String(text || "").trim();
    // 1) 作业/考试名称声明
    if (isItemText(t)) {
      const it = ensureItemSmart(t);
      if (it) {
        voice.currentItem = it; currentItemId = it.id; pendingItem = it;
        renderItemSelect(); renderGrid(); renderAllSelectors();
        logLine("🗂 已设定项目：" + it.name, "ok");
        setVoiceStatus("当前项目：" + it.name + " ｜ 请说「学生姓名 成绩」");
      }
      return;
    }
    if (!voice.currentItem) { logLine("⚠ 请先说出作业/考试名称", "warn"); return; }
    // 2) 成绩录入
    const r = parseScoreUtterance(t, state.students);
    if (r.score == null) { logLine("⚠ 未识别到成绩：" + t, "warn"); return; }
    let student = r.student;
    let conf = r.conf;
    if (!student) {
      // 未锚定到姓名 → 对剩余文本做模糊匹配
      const fm = matchStudent(r.rest, state.students);
      if (!fm.match) {
        pendingScore = r.score;
        logLine("⚠ 未在名单匹配到「" + r.rest + "」，请点选：", "warn");
        showCandidates(fm.candidates);
        return;
      }
      student = fm.match; conf = fm.bestScore;
    }
    pendingScore = r.score;
    const flag = (conf != null && conf < 0.7) ? "（低置信度 " + conf.toFixed(2) + "）" : "";
    if (conf != null && conf < 0.7) showCandidates(matchStudent(r.rest || student.name, state.students).candidates);
    setScore(voice.currentItem.id, student.id, r.score);
    logLine("✅ " + student.name + " ← " + r.score + " ［" + voice.currentItem.name + "］" + flag, "ok");
    renderGrid();
  }

  /* ---------- Web Speech API ---------- */
  let recog = null;
  function initRecog() {
    const SR = window.SpeechRecognition || window.webkitSpeechRecognition;
    if (!SR) return false;
    recog = new SR();
    recog.lang = "zh-CN"; recog.continuous = true; recog.interimResults = true;
    recog.onresult = (e) => {
      let interim = "", finalText = "";
      for (let i = e.resultIndex; i < e.results.length; i++) {
        const r = e.results[i];
        if (r.isFinal) finalText += r[0].transcript; else interim += r[0].transcript;
      }
      $("liveTranscript").textContent = interim || finalText;
      if (finalText) handleUtterance(finalText);
    };
    recog.onerror = (e) => { toast("语音错误：" + e.error, "warn"); };
    recog.onend = () => { if (voice.active) { try { recog.start(); } catch (_) {} } else setVoiceUI(); };
    return true;
  }
  function setVoiceStatus(s) { $("voiceStatus").textContent = s; }
  function setVoiceUI() {
    const b = $("btnVoice");
    b.classList.toggle("listening", voice.active);
    b.textContent = voice.active ? "■ 停止语音登记" : "开始语音登记";
  }
  function startVoice() {
    if (!recog && !initRecog()) {
      toast("当前浏览器不支持语音识别，请用 Chrome/Edge，或使用下方模拟输入", "warn");
      $("simInput").focus();
      return;
    }
    voice.active = true; setVoiceUI();
    if (!voice.currentItem && !currentItemId) {
      setVoiceStatus("请先说出本次要登记的成绩名称，例如「第一单元测试」");
    } else {
      voice.currentItem = getItem(currentItemId) || voice.currentItem;
      setVoiceStatus("当前项目：" + (voice.currentItem ? voice.currentItem.name : "未设定") + " ｜ 请说「学生姓名 成绩」");
    }
    try { recog.start(); } catch (_) {}
  }
  function stopVoice() { voice.active = false; if (recog) try { recog.stop(); } catch (_) {} setVoiceUI(); setVoiceStatus("已停止。"); }

  /* ---------- 横向对比 ---------- */
  function renderAllSelectors() {
    // 学生
    const hs = $("hStudent"); const hv = hs.value;
    hs.innerHTML = state.students.map((s) => '<option value="' + s.id + '">' + escapeXml(s.name) + "（" + escapeXml(s.class) + "）</option>").join("");
    if (hv && getItem) {} if (state.students.find((s) => s.id === hv)) hs.value = hv;
    // 项目
    const vs = $("vItem"); const vv = vs.value;
    vs.innerHTML = state.items.map((i) => '<option value="' + i.id + '">' + escapeXml(i.name) + "【" + (i.type || "") + "】</option>").join("");
    if (state.items.find((i) => i.id === vv)) vs.value = vv;
    // 班级（纵向）
    const classes = [...new Set(state.students.map((s) => s.class))];
    const vcs = $("vClassSel"); const vcv = vcs.value;
    vcs.innerHTML = classes.map((c) => '<option value="' + escapeXml(c) + '">' + escapeXml(c) + "</option>").join("");
    if (classes.includes(vcv)) vcs.value = vcv;
  }
  function selectedItemsForStudent() {
    const scope = $("hScope").value;
    let items = state.items.slice();
    if (scope === "unit") {
      const on = [...document.querySelectorAll("#hUnits .chip.on")].map((c) => c.dataset.u);
      items = items.filter((i) => on.includes(i.unit));
    } else if (scope === "range") {
      const f = $("hFrom").value, t = $("hTo").value;
      items = items.filter((i) => (!f || (i.date || "") >= f) && (!t || (i.date || "") <= t));
    }
    items.sort((a, b) => (a.date || "").localeCompare(b.date || "") || a.name.localeCompare(b.name));
    return items;
  }
  function renderHorizontal() {
    const stu = state.students.find((s) => s.id === $("hStudent").value);
    if (!stu) { $("hLine").innerHTML = '<div class="empty">无学生</div>'; return; }
    const items = selectedItemsForStudent();
    if (!items.length) { $("hLine").innerHTML = '<div class="empty">无匹配项目</div>'; $("hRadar").innerHTML = ""; $("hSummary").innerHTML = ""; return; }
    const stuVals = items.map((i) => getScore(i.id, stu.id));
    // 班级均值
    const clsStus = state.students.filter((s) => s.class === stu.class);
    const clsVals = items.map((i) => {
      const vs = clsStus.map((s) => getScore(i.id, s.id)).filter((v) => v != null);
      return vs.length ? +(vs.reduce((a, b) => a + b, 0) / vs.length).toFixed(1) : null;
    });
    svgLine($("hLine"), items.map((i) => i.name), [
      { label: stu.name + "（" + stu.class + "）", color: "#4f8cff", values: stuVals },
      { label: stu.class + " 班级均值", color: "#f5a623", values: clsVals },
    ]);
    // 雷达：按单元均分
    const unitsMap = {};
    items.forEach((i) => { const u = i.unit || i.name; (unitsMap[u] = unitsMap[u] || []).push(i); });
    const unitLabels = Object.keys(unitsMap);
    if (unitLabels.length >= 3) {
      const stuRadar = unitLabels.map((u) => {
        const vs = unitsMap[u].map((i) => getScore(i.id, stu.id)).filter((v) => v != null);
        return vs.length ? +(vs.reduce((a, b) => a + b, 0) / vs.length).toFixed(1) : 0;
      });
      const clsRadar = unitLabels.map((u) => {
        const vs = [];
        unitsMap[u].forEach((i) => clsStus.forEach((s) => { const v = getScore(i.id, s.id); if (v != null) vs.push(v); }));
        return vs.length ? +(vs.reduce((a, b) => a + b, 0) / vs.length).toFixed(1) : 0;
      });
      svgRadar($("hRadar"), unitLabels, [
        { label: stu.name, color: "#4f8cff", values: stuRadar },
        { label: "班级均值", color: "#f5a623", values: clsRadar },
      ]);
    } else { $("hRadar").innerHTML = '<div class="empty">单元不足 3 个，无法绘制雷达图</div>'; }
    // 汇总
    const all = stuVals.filter((v) => v != null);
    const avg = all.length ? (all.reduce((a, b) => a + b, 0) / all.length).toFixed(1) : "—";
    const max = all.length ? Math.max(...all) : "—";
    const min = all.length ? Math.min(...all) : "—";
    $("hSummary").innerHTML = "<b>" + escapeXml(stu.name) + "</b> 共 " + all.length + " 条成绩 ｜ 均值 <b>" + avg +
      "</b> ｜ 最高 <b>" + max + "</b> ｜ 最低 <b>" + min + "</b> ｜ 项目数 " + items.length;
  }

  /* ---------- 纵向对比 ---------- */
  function renderVertical() {
    const it = getItem($("vItem").value);
    if (!it) { $("vStats").innerHTML = '<div class="empty">无项目</div>'; return; }
    let pop = state.students;
    if ($("vScope").value === "class") pop = state.students.filter((s) => s.class === $("vClassSel").value);
    const recs = pop.map((s) => ({ s, v: getScore(it.id, s.id) })).filter((r) => r.v != null);
    if (!recs.length) { $("vStats").innerHTML = '<div class="empty">该项目暂无成绩</div>'; $("vHist").innerHTML = ""; $("vRank").innerHTML = ""; $("vTable").innerHTML = ""; return; }
    const vals = recs.map((r) => r.v);
    const n = vals.length;
    const avg = (vals.reduce((a, b) => a + b, 0) / n).toFixed(1);
    const max = Math.max(...vals), min = Math.min(...vals);
    const sorted = vals.slice().sort((a, b) => a - b);
    const mid = n % 2 ? sorted[(n - 1) / 2] : ((sorted[n / 2 - 1] + sorted[n / 2]) / 2).toFixed(1);
    const pass = vals.filter((v) => v >= 60).length, excel = vals.filter((v) => v >= 85).length;
    $("vStats").innerHTML =
      stat(n, "人数") + stat(avg, "平均分") + stat(max, "最高分") + stat(min, "最低分") +
      stat(mid, "中位数") + stat((pass / n * 100).toFixed(0) + "%", "及格率(≥60)") + stat((excel / n * 100).toFixed(0) + "%", "优秀率(≥85)");
    // 分布
    const buckets = [[0, 60, "<60"], [60, 70, "60-69"], [70, 80, "70-79"], [80, 90, "80-89"], [90, 101, "90-100"]];
    const colors = ["#ff5c72", "#f5a623", "#4f8cff", "#2dd4bf", "#36c98d"];
    const counts = buckets.map((b) => vals.filter((v) => v >= b[0] && v < b[1]).length);
    svgBarsV($("vHist"), buckets.map((b) => b[2]), counts, colors);
    // 排名
    const ranked = recs.slice().sort((a, b) => b.v - a.v);
    svgBarsH($("vRank"), ranked.map((r) => r.s.name), ranked.map((r) => r.v), "#4f8cff");
    // 表
    let html = "<table><thead><tr><th>排名</th><th>姓名</th><th>班级</th><th>成绩</th></tr></thead><tbody>";
    ranked.forEach((r, i) => {
      const rankCls = i === 0 ? "rank-1" : i === 1 ? "rank-2" : i === 2 ? "rank-3" : "";
      html += '<tr><td class="' + rankCls + '">' + (i + 1) + '</td><td>' + escapeXml(r.s.name) + (r.s.sid ? ' <small style="color:var(--muted)">#' + escapeXml(r.s.sid) + "</small>" : "") + "</td><td>" + escapeXml(r.s.class) + "</td><td>" + r.v + "</td></tr>";
    });
    html += "</tbody></table>";
    $("vTable").innerHTML = html;
  }
  function stat(v, l) { return '<div class="stat"><div class="v">' + v + '</div><div class="l">' + l + "</div></div>"; }

  /* ===================== SVG 图表 ===================== */
  function svgLine(container, labels, series) {
    const W = 660, H = 340, pl = 44, pr = 18, pt = 18, pb = 54;
    const iw = W - pl - pr, ih = H - pt - pb, ymax = 100, n = labels.length;
    const x = (i) => (n <= 1 ? pl + iw / 2 : pl + (iw * i) / (n - 1));
    const y = (v) => pt + ih * (1 - v / ymax);
    let g = "";
    for (let v = 0; v <= 100; v += 20) {
      const yy = y(v);
      g += '<line x1="' + pl + '" y1="' + yy + '" x2="' + (W - pr) + '" y2="' + yy + '" stroke="#2c3344"/>';
      g += '<text x="' + (pl - 8) + '" y="' + (yy + 4) + '" fill="#8b93a7" font-size="11" text-anchor="end">' + v + "</text>";
    }
    labels.forEach((lb, i) => { g += '<text x="' + x(i) + '" y="' + (H - pb + 18) + '" fill="#8b93a7" font-size="11" text-anchor="middle">' + escapeXml(trunc(lb, 6)) + "</text>"; });
    let ly = pt;
    series.forEach((s) => {
      let pts = "", dots = "";
      s.values.forEach((v, i) => { if (v == null) return; const px = x(i), py = y(v); pts += px + "," + py + " "; dots += '<circle cx="' + px + '" cy="' + py + '" r="3.5" fill="' + s.color + '"/><text x="' + px + '" y="' + (py - 8) + '" fill="' + s.color + '" font-size="10" text-anchor="middle">' + v + "</text>"; });
      g += '<polyline points="' + pts.trim() + '" fill="none" stroke="' + s.color + '" stroke-width="2.5" stroke-linejoin="round"/>' + dots;
      g += '<rect x="' + (W - pr - 130) + '" y="' + ly + '" width="10" height="10" fill="' + s.color + '"/><text x="' + (W - pr - 114) + '" y="' + (ly + 9) + '" fill="#e6e9ef" font-size="11">' + escapeXml(s.label) + "</text>";
      ly += 16;
    });
    container.innerHTML = '<svg viewBox="0 0 ' + W + " " + H + '" style="width:100%;height:auto">' + g + "</svg>";
  }
  function svgRadar(container, labels, series) {
    const W = 420, H = 380, cx = W / 2, cy = H / 2 + 6, R = 125, n = labels.length;
    if (n < 3) { container.innerHTML = '<div class="empty">单元不足 3 个，无法绘制雷达图</div>'; return; }
    const ang = (i) => -Math.PI / 2 + (i * 2 * Math.PI) / n;
    const pt = (i, r) => [cx + r * Math.cos(ang(i)), cy + r * Math.sin(ang(i))];
    let g = "";
    [20, 40, 60, 80, 100].forEach((v) => {
      let p = ""; for (let i = 0; i < n; i++) { const [x, y] = pt(i, (R * v) / 100); p += x + "," + y + " "; }
      g += '<polygon points="' + p.trim() + '" fill="none" stroke="#2c3344"/>';
    });
    labels.forEach((lb, i) => {
      const [x, y] = pt(i, R); g += '<line x1="' + cx + '" y1="' + cy + '" x2="' + x + '" y2="' + y + '" stroke="#2c3344"/>';
      const [lx, ly] = pt(i, R + 18); g += '<text x="' + lx + '" y="' + ly + '" fill="#8b93a7" font-size="11" text-anchor="middle">' + escapeXml(trunc(lb, 6)) + "</text>";
    });
    series.forEach((s) => {
      let p = ""; for (let i = 0; i < n; i++) { const v = Math.max(0, Math.min(100, s.values[i] || 0)); const [x, y] = pt(i, (R * v) / 100); p += x + "," + y + " "; }
      g += '<polygon points="' + p.trim() + '" fill="' + s.color + '22" stroke="' + s.color + '" stroke-width="2"/>';
      for (let i = 0; i < n; i++) { const v = Math.max(0, Math.min(100, s.values[i] || 0)); const [x, y] = pt(i, (R * v) / 100); g += '<circle cx="' + x + '" cy="' + y + '" r="2.5" fill="' + s.color + '"/>'; }
    });
    let ly = 14; series.forEach((s) => { g += '<rect x="10" y="' + ly + '" width="10" height="10" fill="' + s.color + '"/><text x="26" y="' + (ly + 9) + '" fill="#e6e9ef" font-size="11">' + escapeXml(s.label) + "</text>"; ly += 16; });
    container.innerHTML = '<svg viewBox="0 0 ' + W + " " + H + '" style="width:100%;height:auto">' + g + "</svg>";
  }
  function svgBarsH(container, labels, values, color) {
    const rowH = 26, top = 10, left = 92, W = 620, H = top + labels.length * rowH + 10;
    const max = Math.max(1, ...values);
    let g = "";
    labels.forEach((lb, i) => {
      const v = values[i], yy = top + i * rowH, bw = (W - left - 60) * (v / max);
      g += '<text x="' + (left - 8) + '" y="' + (yy + rowH / 2 + 4) + '" fill="#e6e9ef" font-size="12" text-anchor="end">' + escapeXml(trunc(lb, 8)) + "</text>";
      g += '<rect x="' + left + '" y="' + (yy + 3) + '" width="' + Math.max(2, bw) + '" height="' + (rowH - 8) + '" rx="4" fill="' + color + '"/>';
      g += '<text x="' + (left + Math.max(2, bw) + 6) + '" y="' + (yy + rowH / 2 + 4) + '" fill="#8b93a7" font-size="11">' + v + "</text>";
    });
    container.innerHTML = '<svg viewBox="0 0 ' + W + " " + H + '" style="width:100%;height:auto">' + g + "</svg>";
  }
  function svgBarsV(container, labels, values, colors) {
    const W = 620, H = 300, pl = 40, pr = 14, pt = 14, pb = 46, iw = W - pl - pr, ih = H - pt - pb;
    const max = Math.max(1, ...values), n = labels.length, gap = iw / n, bw = gap * 0.6;
    let g = "";
    for (let v = 0; v <= max; v += Math.max(1, Math.ceil(max / 5))) {
      const yy = pt + ih * (1 - v / max);
      g += '<line x1="' + pl + '" y1="' + yy + '" x2="' + (W - pr) + '" y2="' + yy + '" stroke="#2c3344"/><text x="' + (pl - 6) + '" y="' + (yy + 4) + '" fill="#8b93a7" font-size="10" text-anchor="end">' + v + "</text>";
    }
    labels.forEach((lb, i) => {
      const v = values[i], x = pl + gap * i + (gap - bw) / 2, bh = (ih * v) / max, y = pt + ih - bh;
      g += '<rect x="' + x + '" y="' + y + '" width="' + bw + '" height="' + bh + '" rx="4" fill="' + (colors[i] || "#4f8cff") + '"/>';
      g += '<text x="' + (x + bw / 2) + '" y="' + (y - 5) + '" fill="#e6e9ef" font-size="11" text-anchor="middle">' + v + "</text>";
      g += '<text x="' + (x + bw / 2) + '" y="' + (H - pb + 18) + '" fill="#8b93a7" font-size="11" text-anchor="middle">' + escapeXml(lb) + "</text>";
    });
    container.innerHTML = '<svg viewBox="0 0 ' + W + " " + H + '" style="width:100%;height:auto">' + g + "</svg>";
  }

  /* ===================== 事件绑定 ===================== */
  function bind() {
    // tabs
    document.querySelectorAll(".tab").forEach((t) => t.onclick = () => {
      document.querySelectorAll(".tab").forEach((x) => x.classList.remove("active"));
      document.querySelectorAll(".panel").forEach((x) => x.classList.remove("active"));
      t.classList.add("active");
      $("tab-" + t.dataset.tab).classList.add("active");
    });
    // roster
    $("btnImport").onclick = () => {
      const n = parseRoster($("rosterImport").value);
      save(); renderRoster(); renderAllSelectors(); renderGrid();
      toast("已导入 " + n + " 名学生");
    };
    $("fileImport").onchange = (e) => {
      const f = e.target.files[0]; if (!f) return;
      const r = new FileReader();
      r.onload = () => { const n = parseRoster(r.result); save(); renderRoster(); renderAllSelectors(); renderGrid(); toast("文件导入 " + n + " 名学生"); };
      r.readAsText(f, "utf-8");
    };
    $("fileImportXlsx").onchange = async (e) => {
      const files = Array.from(e.target.files || []); if (!files.length) return;
      let totalAdded = 0, totalSid = 0, filesOk = 0;
      const problems = [];
      for (const f of files) {
        if (!/\.xlsx$/i.test(f.name)) { problems.push(f.name + "（非 .xlsx）"); continue; }
        try {
          const className = classNameFromFilename(f.name);
          const buf = await f.arrayBuffer();
          const rows = await parseXlsx(buf);
          const list = xlsxRowsToLines(rows, className);
          const r = importStudents(list);
          totalAdded += r.added; totalSid += r.withSid; filesOk++;
        } catch (err) {
          console.error(err);
          problems.push(f.name + "（" + (err && err.message ? err.message : err) + "）");
        }
      }
      save(); renderRoster(); renderAllSelectors(); renderGrid();
      let msg = "Excel 导入 " + totalAdded + " 名学生" + (totalSid ? "（含学号 " + totalSid + " 条）" : "") + "，共 " + filesOk + " 个文件";
      if (problems.length) msg += "；跳过：" + problems.join("；");
      toast(msg, problems.length && !filesOk ? "warn" : undefined);
      e.target.value = "";
    };
    $("btnAdd").onclick = () => {
      const name = $("newName").value.trim(), cls = $("newClass").value.trim() || "未分班";
      if (!name) return toast("请输入姓名");
      if (state.students.some((s) => s.name === name && s.class === cls)) return toast("该生已存在");
      state.students.push({ id: uid(), name, class: cls });
      $("newName").value = ""; save(); renderRoster(); renderAllSelectors(); renderGrid();
    };
    // item
    $("itemSelect").onchange = () => { currentItemId = $("itemSelect").value; renderItemMeta(); };
    $("btnNewItem").onclick = () => {
      $("itemForm").classList.toggle("hidden");
      $("itemName").value = getItem(currentItemId) ? getItem(currentItemId).name : "";
      $("itemType").value = getItem(currentItemId) ? (getItem(currentItemId).type || "作业") : "作业";
      $("itemUnit").value = getItem(currentItemId) ? (getItem(currentItemId).unit || "") : "";
      $("itemDate").value = getItem(currentItemId) ? (getItem(currentItemId).date || today()) : today();
    };
    $("btnDelItem").onclick = () => {
      const it = getItem(currentItemId);
      if (!it) return toast("请先在上方选择一个作业/考试", "warn");
      if (!confirm("删除「" + it.name + "」？该项目的所有成绩也会一并删除。")) return;
      state.items = state.items.filter((x) => x.id !== it.id);
      if (state.scores[it.id]) delete state.scores[it.id];
      currentItemId = null;
      save(); renderItemSelect(); renderGrid(); renderAllSelectors(); toast("已删除：" + it.name);
    };
    $("btnCancelItem").onclick = () => $("itemForm").classList.add("hidden");
    $("itemForm").onsubmit = (e) => {
      e.preventDefault();
      const name = $("itemName").value.trim(); if (!name) return;
      const data = { name, type: $("itemType").value, unit: $("itemUnit").value.trim(), date: $("itemDate").value || today() };
      if (getItem(currentItemId)) Object.assign(getItem(currentItemId), data);
      else { const it = Object.assign({ id: uid() }, data); state.items.push(it); currentItemId = it.id; }
      $("itemForm").classList.add("hidden"); save(); renderItemSelect(); renderGrid(); renderAllSelectors(); toast("已保存项目");
    };
    // voice
    $("btnVoice").onclick = () => { voice.active ? stopVoice() : startVoice(); };
    $("btnSim").onclick = () => { const t = $("simInput").value.trim(); if (t) handleUtterance(t); };
    $("simInput").addEventListener("keydown", (e) => { if (e.key === "Enter") { const t = $("simInput").value.trim(); if (t) handleUtterance(t); } });
    // horizontal
    $("hScope").onchange = () => {
      $("hUnitBox").classList.toggle("hidden", $("hScope").value !== "unit");
      $("hRangeBox").classList.toggle("hidden", $("hScope").value !== "range");
      if ($("hScope").value === "unit") renderUnitChips();
    };
    $("btnHRender").onclick = renderHorizontal;
    // vertical
    $("vScope").onchange = () => { $("vClassBox").classList.toggle("hidden", $("vScope").value !== "class"); };
    $("btnVRender").onclick = renderVertical;
    // data
    $("btnExport").onclick = () => {
      const blob = new Blob([JSON.stringify(state, null, 2)], { type: "application/json" });
      const a = document.createElement("a"); a.href = URL.createObjectURL(blob);
      a.download = "成绩数据_" + today() + ".json"; a.click();
    };
    $("btnImportJson").onclick = () => $("fileImportJson").click();
    $("fileImportJson").onchange = (e) => {
      const f = e.target.files[0]; if (!f) return;
      const r = new FileReader();
      r.onload = () => { try { state = JSON.parse(r.result); save(); renderAll(); toast("数据已导入"); } catch (err) { toast("JSON 解析失败", "warn"); } };
      r.readAsText(f, "utf-8");
    };
  }
  function renderUnitChips() {
    const units = [...new Set(state.items.map((i) => i.unit).filter(Boolean))];
    const box = $("hUnits"); box.innerHTML = "";
    units.forEach((u) => {
      const c = document.createElement("span"); c.className = "chip on"; c.dataset.u = u; c.textContent = u;
      c.onclick = () => c.classList.toggle("on");
      box.appendChild(c);
    });
  }
  function renderAll() {
    renderRoster(); renderItemSelect(); renderGrid(); renderAllSelectors(); renderUnitChips();
  }

  /* ===================== 启动 ===================== */
  load().then(() => {
    bind();
    renderAll();
  });
})();
