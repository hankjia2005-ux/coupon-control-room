"use client";

import { useEffect, useMemo, useRef, useState, type FormEvent } from "react";
import * as XLSX from "xlsx";
import { supabase } from "../lib/supabase";

export type CouponRow = {
  id: string;
  sku: string;
  usable: boolean;
  manager: string;
  storeId: string;
  category1: string;
  category2: string;
  category3: string;
  brand: string;
  statusSource?: "source" | "manual";
  updatedBy?: string;
  updatedAt?: string;
};

type GroupKey = "manager" | "storeId" | "brand";
type FilterState = { search: string; status: "all" | "blocked" | "usable"; manager: string; storeId: string; brand: string };

const aliases: Record<string, string[]> = {
  sku: ["上传内容", "sku", "SKU", "商品编码"],
  usable: ["是否可用东券", "usable", "可用东券"],
  manager: ["采销/类目运营人员", "manager", "负责人"],
  storeId: ["店铺id", "店铺ID", "storeId", "店铺"],
  category1: ["一级类目", "category1"],
  category2: ["二级类目", "category2"],
  category3: ["三级类目", "category3"],
  brand: ["品牌", "brand"],
};

function getField(row: Record<string, unknown>, field: string) {
  for (const alias of aliases[field] ?? [field]) {
    const direct = Object.keys(row).find((key) => key.replace(/^\uFEFF/, "").trim() === alias);
    if (direct) return row[direct];
  }
  return "";
}

function expandScientific(value: unknown) {
  let text = String(value ?? "").trim().replace(/,/g, "");
  if (!text) return "";
  if (/^[-+]?\d+\.0+$/.test(text)) text = text.replace(/\.0+$/, "");
  const match = text.match(/^([-+]?)(\d+(?:\.\d+)?)[eE]([-+]?\d+)$/);
  if (!match) return text;
  const sign = match[1];
  const raw = match[2];
  const digits = raw.replace(".", "");
  const decimalAt = (raw.indexOf(".") === -1 ? raw.length : raw.indexOf(".")) + Number(match[3]);
  if (decimalAt <= 0) return `${sign}0.${"0".repeat(Math.abs(decimalAt))}${digits}`;
  if (decimalAt >= digits.length) return `${sign}${digits}${"0".repeat(decimalAt - digits.length)}`;
  return `${sign}${digits.slice(0, decimalAt)}.${digits.slice(decimalAt)}`;
}

function normalizeRows(input: unknown[]): CouponRow[] {
  return input.map((raw, index) => {
    const row = (raw ?? {}) as Record<string, unknown>;
    const sku = expandScientific(getField(row, "sku"));
    const usableRaw = String(getField(row, "usable") ?? "").trim().toLowerCase();
    return {
      id: String(row.id ?? `${sku}-${index}`),
      sku,
      usable: ["是", "true", "1", "yes", "可用"].includes(usableRaw),
      manager: String(getField(row, "manager") ?? "").trim() || "未标注",
      storeId: String(getField(row, "storeId") ?? "").trim() || "未标注",
      category1: String(getField(row, "category1") ?? "").trim() || "—",
      category2: String(getField(row, "category2") ?? "").trim() || "—",
      category3: String(getField(row, "category3") ?? "").trim() || "—",
      brand: String(getField(row, "brand") ?? "").trim() || "未标注",
      statusSource: row.statusSource === "manual" ? "manual" as const : "source" as const,
      updatedBy: typeof row.updatedBy === "string" ? row.updatedBy : undefined,
      updatedAt: typeof row.updatedAt === "string" ? row.updatedAt : undefined,
    };
  }).filter((row) => row.sku);
}

function parseCsv(text: string) {
  const source = text.replace(/^\uFEFF/, "");
  const rows: string[][] = [];
  let row: string[] = [];
  let cell = "";
  let quoted = false;
  for (let i = 0; i < source.length; i += 1) {
    const char = source[i];
    const next = source[i + 1];
    if (char === '"' && quoted && next === '"') { cell += '"'; i += 1; continue; }
    if (char === '"') { quoted = !quoted; continue; }
    if (char === "," && !quoted) { row.push(cell); cell = ""; continue; }
    if ((char === "\n" || char === "\r") && !quoted) {
      if (char === "\r" && next === "\n") i += 1;
      row.push(cell); cell = "";
      if (row.some((value) => value.trim())) rows.push(row);
      row = [];
      continue;
    }
    cell += char;
  }
  if (cell || row.length) { row.push(cell); rows.push(row); }
  const headers = rows.shift()?.map((header) => header.trim().replace(/^\uFEFF/, "")) ?? [];
  return rows.map((values) => headers.reduce<Record<string, string>>((result, header, index) => {
    result[header] = values[index] ?? "";
    return result;
  }, {}));
}

function formatNumber(value: number) {
  return new Intl.NumberFormat("zh-CN").format(value);
}

function formatPercent(value: number) {
  return `${(value * 100).toFixed(2)}%`;
}

function escapeCsv(value: unknown) {
  const text = String(value ?? "");
  return /[",\n]/.test(text) ? `"${text.replace(/"/g, '""')}"` : text;
}

export default function DashboardClient({
  initialRows,
  initialSource,
  initialUpdatedAt,
  canEdit,
  userName,
}: {
  initialRows: CouponRow[];
  initialSource: string;
  initialUpdatedAt: string;
  canEdit: boolean;
  userName: string;
}) {
  const fileInput = useRef<HTMLInputElement>(null);
  const [rows, setRows] = useState<CouponRow[]>(() => normalizeRows(initialRows));
  const [meta, setMeta] = useState({ sourceName: initialSource, updatedAt: initialUpdatedAt, shared: false });
  const [filters, setFilters] = useState<FilterState>({ search: "", status: "all", manager: "all", storeId: "all", brand: "all" });
  const [groupBy, setGroupBy] = useState<GroupKey>("manager");
  const [page, setPage] = useState(1);
  const [pageSize, setPageSize] = useState(15);
  const [busy, setBusy] = useState(false);
  const [toast, setToast] = useState("");
  const [signedInEmail, setSignedInEmail] = useState("");
  const [authEmail, setAuthEmail] = useState("");
  const [authOpen, setAuthOpen] = useState(false);
  const [authBusy, setAuthBusy] = useState(false);

  const editable = !supabase ? canEdit : Boolean(signedInEmail);
  const displayUserName = signedInEmail || userName;

  useEffect(() => {
    const client = supabase;
    if (!client) return;
    let active = true;

    function applySnapshot(snapshot: { rows?: unknown; source_name?: string; updated_at?: string } | null) {
      if (!active || !snapshot || !Array.isArray(snapshot.rows)) return;
      setRows(normalizeRows(snapshot.rows));
      setMeta({
        sourceName: snapshot.source_name || "云端共享数据",
        updatedAt: snapshot.updated_at ? new Date(snapshot.updated_at).toLocaleString("zh-CN", { hour12: false }) : "刚刚",
        shared: true,
      });
    }

    async function refreshCloud(seedIfEmpty = false, actor = "") {
      const { data, error } = await client!.from("coupon_current").select("id, rows, source_name, updated_at").eq("id", 1).maybeSingle();
      if (error) {
        if (active) showToast("云端数据暂不可用，请检查 Supabase 表结构。");
        return;
      }
      if (data) {
        applySnapshot(data);
        return;
      }
      if (!seedIfEmpty) return;
      const { data: seeded, error: seedError } = await client!.from("coupon_current").upsert({
        id: 1,
        rows: normalizeRows(initialRows),
        source_name: initialSource,
        updated_at: new Date().toISOString(),
        updated_by: actor || "首个登录用户",
      }, { onConflict: "id" }).select("id, rows, source_name, updated_at").single();
      if (seedError) {
        if (active) showToast("云端初始化失败，请确认已执行 Supabase 建表 SQL。");
        return;
      }
      applySnapshot(seeded);
    }

    client.auth.getSession().then(({ data }) => {
      if (!active) return;
      const email = data.session?.user.email ?? "";
      setSignedInEmail(email);
      void refreshCloud(Boolean(email), email);
    });

    const { data: authListener } = client.auth.onAuthStateChange((_event, session) => {
      const email = session?.user.email ?? "";
      setSignedInEmail(email);
      void refreshCloud(Boolean(email), email);
    });

    const channel = client.channel("coupon-current-live")
      .on("postgres_changes", { event: "UPDATE", schema: "public", table: "coupon_current", filter: "id=eq.1" }, (payload) => {
        applySnapshot(payload.new as { rows?: unknown; source_name?: string; updated_at?: string });
      })
      .subscribe();
    const timer = window.setInterval(() => void refreshCloud(false), 20000);

    return () => {
      active = false;
      window.clearInterval(timer);
      void client.removeChannel(channel);
      authListener.subscription.unsubscribe();
    };
  }, []);

  useEffect(() => { setPage(1); }, [filters.search, filters.status, filters.manager, filters.storeId, filters.brand, pageSize]);

  useEffect(() => {
    if (!toast) return;
    const timer = window.setTimeout(() => setToast(""), 3200);
    return () => window.clearTimeout(timer);
  }, [toast]);

  const options = useMemo(() => ({
    managers: [...new Set(rows.map((row) => row.manager))].sort((a, b) => a.localeCompare(b, "zh-CN")),
    stores: [...new Set(rows.map((row) => row.storeId))].sort((a, b) => a.localeCompare(b, "zh-CN")),
    brands: [...new Set(rows.map((row) => row.brand))].sort((a, b) => a.localeCompare(b, "zh-CN")),
  }), [rows]);

  const filtered = useMemo(() => {
    const query = filters.search.trim().toLowerCase();
    return rows.filter((row) => {
      const matchesQuery = !query || [row.sku, row.manager, row.storeId, row.category1, row.category2, row.category3, row.brand].some((value) => value.toLowerCase().includes(query));
      const matchesStatus = filters.status === "all" || (filters.status === "blocked" ? !row.usable : row.usable);
      return matchesQuery && matchesStatus && (filters.manager === "all" || row.manager === filters.manager) && (filters.storeId === "all" || row.storeId === filters.storeId) && (filters.brand === "all" || row.brand === filters.brand);
    });
  }, [rows, filters]);

  const stats = useMemo(() => {
    const blocked = filtered.filter((row) => !row.usable).length;
    return { total: filtered.length, blocked, usable: filtered.length - blocked, rate: filtered.length ? blocked / filtered.length : 0 };
  }, [filtered]);

  const ranking = useMemo(() => {
    const map = new Map<string, { label: string; total: number; blocked: number }>();
    filtered.forEach((row) => {
      const label = row[groupBy] || "未标注";
      const current = map.get(label) ?? { label, total: 0, blocked: 0 };
      current.total += 1;
      if (!row.usable) current.blocked += 1;
      map.set(label, current);
    });
    return [...map.values()].map((item) => ({ ...item, rate: item.blocked / item.total })).sort((a, b) => b.rate - a.rate || b.blocked - a.blocked).slice(0, 8);
  }, [filtered, groupBy]);

  const totalPages = Math.max(1, Math.ceil(filtered.length / pageSize));
  const visibleRows = filtered.slice().sort((a, b) => Number(a.usable) - Number(b.usable) || a.sku.localeCompare(b.sku)).slice((page - 1) * pageSize, page * pageSize);

  function showToast(message: string) { setToast(message); }

  function updateFilter(key: keyof FilterState, value: string) {
    setFilters((current) => ({ ...current, [key]: value }));
  }

  async function sendMagicLink(event: FormEvent<HTMLFormElement>) {
    event.preventDefault();
    if (!supabase || !authEmail.trim()) return;
    setAuthBusy(true);
    const { error } = await supabase.auth.signInWithOtp({
      email: authEmail.trim(),
      options: { emailRedirectTo: window.location.href.split("#")[0] },
    });
    setAuthBusy(false);
    showToast(error ? error.message : "登录链接已发送，请检查邮箱后点击链接回来。");
    if (!error) setAuthOpen(false);
  }

  async function signOut() {
    if (!supabase) return;
    await supabase.auth.signOut();
    setSignedInEmail("");
    showToast("已退出，当前为只读浏览模式。");
  }

  async function handleImport(file: File) {
    if (supabase && !signedInEmail) {
      showToast("导入今日数据前，请先用邮箱 magic link 登录。");
      return;
    }
    setBusy(true);
    try {
      const extension = file.name.toLowerCase().split(".").pop();
      const buffer = await file.arrayBuffer();
      const rawRows = extension === "xlsx" || extension === "xls"
        ? (() => { const workbook = XLSX.read(buffer, { type: "array", cellDates: false }); return XLSX.utils.sheet_to_json(workbook.Sheets[workbook.SheetNames[0]], { defval: "", raw: false }); })()
        : parseCsv(new TextDecoder("gb18030").decode(buffer));
      const normalized = normalizeRows(rawRows as unknown[]).map((row) => ({ ...row, statusSource: "source" as const, updatedBy: undefined, updatedAt: undefined }));
      let syncedAt = new Date().toISOString();
      if (supabase) {
        const { data, error } = await supabase.from("coupon_current").upsert({
          id: 1,
          rows: normalized,
          source_name: file.name,
          updated_at: syncedAt,
          updated_by: signedInEmail || "在线用户",
        }, { onConflict: "id" }).select("source_name, updated_at").single();
        if (error) throw new Error(error.message || "上传失败");
        syncedAt = data.updated_at;
      } else {
        const response = await fetch("/api/upload", { method: "POST", headers: { "Content-Type": "application/json" }, body: JSON.stringify({ rows: normalized, sourceName: file.name }) });
        const result = await response.json();
        if (!response.ok) throw new Error(result.error || "上传失败");
        syncedAt = result.updatedAt;
      }
      setRows(normalized);
      setMeta({ sourceName: file.name, updatedAt: new Date(syncedAt).toLocaleString("zh-CN", { hour12: false }), shared: true });
      showToast(`已同步 ${formatNumber(normalized.length)} 条 SKU，所有同事刷新后可见。`);
    } catch (error) {
      showToast(error instanceof Error ? error.message : "文件读取失败，请检查列名后重试。");
    } finally {
      setBusy(false);
      if (fileInput.current) fileInput.current.value = "";
    }
  }

  async function handleStatusChange(id: string, usable: boolean) {
    if (supabase && !signedInEmail) {
      showToast("请先登录后再修改 SKU 状态。");
      return;
    }
    const previous = rows.find((row) => row.id === id);
    if (!previous || previous.usable === usable) return;
    setRows((current) => current.map((row) => row.id === id ? { ...row, usable, statusSource: "manual" } : row));
    try {
      if (supabase) {
        const { data, error } = await supabase.rpc("update_coupon_status", { p_id: id, p_usable: usable, p_updated_by: signedInEmail || "在线用户" });
        if (error) throw new Error(error.message || "状态更新失败");
        const snapshot = Array.isArray(data) ? data[0] : data;
        if (snapshot?.rows) setRows(normalizeRows(snapshot.rows));
        if (snapshot?.updated_at) setMeta((current) => ({ ...current, updatedAt: new Date(snapshot.updated_at).toLocaleString("zh-CN", { hour12: false }), shared: true }));
      } else {
        const response = await fetch("/api/row", { method: "POST", headers: { "Content-Type": "application/json" }, body: JSON.stringify({ id, usable }) });
        const result = await response.json();
        if (!response.ok) throw new Error(result.error || "状态更新失败");
        setRows((current) => current.map((row) => row.id === id ? result.row : row));
        setMeta((current) => ({ ...current, updatedAt: result.updatedAt, shared: true }));
      }
      showToast("状态已保存，在线同事刷新后即可同步看到。");
    } catch (error) {
      setRows((current) => current.map((row) => row.id === id ? previous : row));
      showToast(error instanceof Error ? error.message : "状态更新失败，请重试。");
    }
  }

  function exportCsv() {
    const header = ["SKU", "是否可用东券", "店铺id", "一级类目", "二级类目", "三级类目", "品牌", "采销/类目运营人员", "状态来源"];
    const body = filtered.map((row) => [row.sku, row.usable ? "是" : "否", row.storeId, row.category1, row.category2, row.category3, row.brand, row.manager, row.statusSource === "manual" ? "人工修订" : "原始检测"]);
    const csv = [header, ...body].map((line) => line.map(escapeCsv).join(",")).join("\r\n");
    const url = URL.createObjectURL(new Blob(["\uFEFF" + csv], { type: "text/csv;charset=utf-8" }));
    const link = document.createElement("a");
    link.href = url; link.download = `东券禁券监查_${new Date().toISOString().slice(0, 10)}.csv`; link.click();
    URL.revokeObjectURL(url);
  }

  const scopeLabel = [filters.manager !== "all" ? filters.manager : "", filters.storeId !== "all" ? `店铺 ${filters.storeId}` : "", filters.brand !== "all" ? `品牌 ${filters.brand}` : "", filters.status === "blocked" ? "仅被禁券" : filters.status === "usable" ? "仅可用" : ""].filter(Boolean).join(" · ") || "全量监查";

  return (
    <div className="app-shell">
      <header className="site-header">
        <div className="brand-lockup"><div className="brand-overline">JD / OPERATIONS INTELLIGENCE</div><div className="brand-title">东券禁券 <span>·</span> 今日监查</div></div>
        <div className="header-actions"><div className="sync-status"><span className="sync-dot" />共享在线数据<span className="sync-time">更新于 {meta.updatedAt}</span></div><button className="button button-secondary" onClick={exportCsv}>导出筛选结果</button><button className="button button-primary" onClick={() => fileInput.current?.click()} disabled={busy || (Boolean(supabase) && !signedInEmail)}>{busy ? "同步中…" : "导入今日数据"}</button>{supabase && (signedInEmail ? <div className="auth-identity"><span>{signedInEmail}</span><button className="auth-link" onClick={signOut}>退出</button></div> : <button className="button button-login" onClick={() => setAuthOpen((value) => !value)}>邮箱登录</button>)}<input ref={fileInput} type="file" accept=".csv,.xlsx,.xls" hidden onChange={(event) => event.target.files?.[0] && handleImport(event.target.files[0])} />{authOpen && !signedInEmail && <form className="auth-popover" onSubmit={sendMagicLink}><strong>无密码登录看板</strong><span>输入工作邮箱，接收一次性登录链接。</span><input type="email" value={authEmail} onChange={(event) => setAuthEmail(event.target.value)} placeholder="name@jd.com" required /><button className="button button-primary" type="submit" disabled={authBusy}>{authBusy ? "发送中…" : "发送登录链接"}</button></form>}</div>
      </header>

      <main className="page-wrap">
        <section className="hero-section">
          <div><div className="section-kicker">DAILY COUPON CONTROL / {meta.shared ? "SHARED LIVE" : "READY TO SYNC"}</div><h1>把今天的禁券风险，<em>一眼看清。</em></h1><p>以 SKU 为最小监查单元，快速定位被禁券商品，按负责人、店铺与品牌拆解风险；状态修改会同步到所有在线用户。</p><div className="hero-tags"><span>禁券口径 <strong>是否可用东券 = 否</strong></span><span>数据源 <strong>{meta.sourceName}</strong></span></div></div>
          <aside className="risk-summary"><div className="risk-summary-top"><span>当前筛选范围</span><span className="risk-state">{stats.blocked ? "需要关注" : "状态稳定"}</span></div><strong>{formatNumber(stats.blocked)}</strong><div className="risk-summary-bottom"><span>被禁券 SKU</span><span>{formatPercent(stats.rate)} · {scopeLabel}</span></div></aside>
        </section>

        <section className="metric-grid" aria-label="今日核心指标"><article><span className="metric-label">监查 SKU</span><strong>{formatNumber(stats.total)}</strong><small>当前筛选范围</small></article><article className="metric-alert"><span className="metric-label">被禁券 SKU</span><strong>{formatNumber(stats.blocked)}</strong><small>是否可用东券 = 否</small></article><article className="metric-safe"><span className="metric-label">可用东券 SKU</span><strong>{formatNumber(stats.usable)}</strong><small>是否可用东券 = 是</small></article><article className="metric-rate"><span className="metric-label">禁券比例</span><strong>{formatPercent(stats.rate)}</strong><small>被禁券 / 总 SKU</small></article></section>

        <section className="filter-bar"><div className="search-box"><span>⌕</span><input value={filters.search} onChange={(event) => updateFilter("search", event.target.value)} placeholder="搜索 SKU / 店铺 ID / 品牌 ID / 负责人 / 类目 ID" aria-label="搜索" /><kbd>/</kbd></div><div className="filters"><select value={filters.status} onChange={(event) => updateFilter("status", event.target.value)} aria-label="券状态"><option value="all">全部券状态</option><option value="blocked">仅被禁券</option><option value="usable">仅可用东券</option></select><select value={filters.manager} onChange={(event) => updateFilter("manager", event.target.value)} aria-label="负责人"><option value="all">全部负责人</option>{options.managers.map((value) => <option key={value} value={value}>{value}</option>)}</select><select value={filters.storeId} onChange={(event) => updateFilter("storeId", event.target.value)} aria-label="店铺 ID"><option value="all">全部店铺 ID</option>{options.stores.map((value) => <option key={value} value={value}>{value}</option>)}</select><select value={filters.brand} onChange={(event) => updateFilter("brand", event.target.value)} aria-label="品牌 ID"><option value="all">全部品牌 ID</option>{options.brands.map((value) => <option key={value} value={value}>{value}</option>)}</select><button className="clear-button" onClick={() => setFilters({ search: "", status: "all", manager: "all", storeId: "all", brand: "all" })}>清除</button></div></section>

        <section className="insight-grid"><article className="card ranking-card"><div className="card-heading"><div><div className="section-kicker">RISK CONCENTRATION</div><h2>异常密度排行</h2></div><select value={groupBy} onChange={(event) => setGroupBy(event.target.value as GroupKey)} aria-label="排行维度"><option value="manager">按负责人</option><option value="storeId">按店铺</option><option value="brand">按品牌</option></select></div><div className="rank-note"><span className="legend-dot" />禁券比例 <span>· 点击排行项快速筛选</span></div><div className="ranking-list">{ranking.map((item, index) => <button className="ranking-item" key={item.label} onClick={() => updateFilter(groupBy, item.label)}><span className="rank-no">{String(index + 1).padStart(2, "0")}</span><span className="rank-name">{item.label}</span><span className="rank-track"><i style={{ width: `${Math.max(item.rate ? 4 : 0, (item.rate / Math.max(...ranking.map((entry) => entry.rate), 0.01)) * 100)}%` }} /></span><span className="rank-meta"><b>{formatPercent(item.rate)}</b><small>{item.blocked}/{item.total}</small></span></button>)}</div></article><article className="card source-card"><div className="source-badge"><span className="sync-dot" /> SHARED DATA SOURCE</div><h2>今日数据工作台</h2><p>每天导入新文件后，当前数据会覆盖为最新快照；同事打开同一个链接，就能看到同一份在线结果。</p><div className="source-detail"><div><span>当前用户</span><strong>{displayUserName}</strong></div><div><span>最近同步</span><strong>{meta.updatedAt}</strong></div><div><span>状态修改</span><strong>{editable ? "已开放在线修订" : "邮箱登录后可修订"}</strong></div></div><div className="source-note">原始检测与人工修订会在明细中分别标识，方便日常核对。</div></article></section>

        <section className="table-card card"><div className="table-heading"><div><div className="section-kicker">SKU REGISTER</div><h2>SKU 明细台账</h2></div><div className="table-actions"><span>命中 {formatNumber(stats.total)} / {formatNumber(rows.length)}</span><select value={pageSize} onChange={(event) => setPageSize(Number(event.target.value))} aria-label="每页显示数量"><option value="15">15 / 页</option><option value="30">30 / 页</option><option value="60">60 / 页</option></select></div></div><div className="table-note">状态可直接在线修改 · 被禁券默认置顶 · 修改会同步至共享数据源</div><div className="table-scroll"><table><thead><tr><th>状态</th><th>SKU</th><th>店铺 ID</th><th>类目 ID（一级 / 二级 / 三级）</th><th>品牌 ID</th><th>负责人</th><th>来源</th></tr></thead><tbody>{visibleRows.length ? visibleRows.map((row) => <tr key={row.id} className={!row.usable ? "blocked-row" : ""}><td><select disabled={!editable || busy} className={`status-control ${row.usable ? "is-usable" : "is-blocked"}`} value={row.usable ? "usable" : "blocked"} onChange={(event) => handleStatusChange(row.id, event.target.value === "usable")} aria-label={`${row.sku} 券状态`}><option value="blocked">被禁券</option><option value="usable">可用东券</option></select></td><td className="sku-text">{row.sku}</td><td className="mono-text">{row.storeId}</td><td className="category-text">{row.category1} <i>/</i> {row.category2} <i>/</i> {row.category3}</td><td className="brand-text">{row.brand}</td><td className="manager-text">{row.manager}</td><td><span className={`source-pill ${row.statusSource === "manual" ? "manual" : "source"}`}>{row.statusSource === "manual" ? "人工修订" : "原始检测"}</span></td></tr>) : <tr><td colSpan={7} className="empty-row">当前筛选范围没有匹配 SKU。</td></tr>}</tbody></table></div><div className="table-footer"><span>显示第 {filtered.length ? (page - 1) * pageSize + 1 : 0}–{Math.min(page * pageSize, filtered.length)} 条</span><div className="pagination"><button disabled={page <= 1} onClick={() => setPage((value) => Math.max(1, value - 1))}>上一页</button><strong>{page} / {totalPages}</strong><button disabled={page >= totalPages} onClick={() => setPage((value) => Math.min(totalPages, value + 1))}>下一页</button></div></div></section>
      </main>
      <footer className="site-footer"><span>COUPON CONTROL ROOM</span><span>共享在线数据 · 每日导入 · 支持状态协同修订</span></footer>
      {toast && <div className="toast" role="status">{toast}</div>}
    </div>
  );
}
