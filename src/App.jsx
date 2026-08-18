import { useMemo, useRef, useState } from "react";
import {
  Activity,
  Award,
  AlertCircle,
  BarChart3,
  CalendarDays,
  Check,
  CheckCircle2,
  ChevronLeft,
  ChevronRight,
  CircleDollarSign,
  ClipboardCheck,
  Download,
  FileCheck2,
  FileSpreadsheet,
  Filter,
  Info,
  Layers3,
  LockKeyhole,
  RefreshCw,
  Search,
  ShieldCheck,
  Store,
  Target,
  TrendingUp,
  UploadCloud,
  UsersRound,
  X,
} from "lucide-react";
import {
  Area,
  AreaChart,
  Bar,
  BarChart,
  CartesianGrid,
  Cell,
  Legend,
  LabelList,
  Pie,
  PieChart,
  ResponsiveContainer,
  Tooltip,
  XAxis,
  YAxis,
} from "recharts";
import {
  applyFilters,
  exportWorkbook,
  groupBy,
  parseFile,
  summarize,
  weekly,
} from "./dataEngine";

const nf = new Intl.NumberFormat("zh-TW");
const pf = new Intl.NumberFormat("zh-TW", {
  style: "percent",
  maximumFractionDigits: 1,
});
const cf = new Intl.NumberFormat("zh-TW", {
  style: "currency",
  currency: "TWD",
  maximumFractionDigits: 0,
});
const n = (value) => nf.format(Number(value || 0));
const p = (value) => pf.format(Number(value || 0));
const c = (value) => cf.format(Number(value || 0));
const COLORS = ["#e21b2d", "#f59e0b", "#111827", "#8b5cf6", "#0f766e", "#94a3b8"];
const ATTRIBUTE_LABELS = {
  A: "軍營",
  B: "金融機構",
  C: "公司",
  H: "醫療院所",
  I: "政府／公營機構",
  P: "個人",
  S: "學校／教育機構",
};
const attributeLabel = (value) => {
  const raw = String(value ?? "").normalize("NFKC").trim();
  if (!raw) return "未分類";
  const code = raw.match(/^([A-Za-z])(?:\s*[｜|／/-].*)?$/)?.[1]?.toUpperCase();
  if (!code) return raw;
  return ATTRIBUTE_LABELS[code] ? `${code}｜${ATTRIBUTE_LABELS[code]}` : `${code}｜其他`;
};

const memberGroup = (value) => {
  const normalized = String(value ?? "").normalize("NFKC").trim().toLowerCase();
  if (["是", "y", "yes", "1", "會員", "true", "v", "✓", "✔"].includes(normalized)) return "會員";
  if (["否", "n", "no", "0", "非會員", "false"].includes(normalized)) return "非會員";
  return "未知";
};

const amountBand = (amount) => {
  if (amount == null) return "金額缺漏／異常";
  if (amount < 5000) return "未滿 5,000";
  if (amount < 10000) return "5,000–9,999";
  if (amount < 20000) return "10,000–19,999";
  if (amount < 50000) return "20,000–49,999";
  if (amount < 100000) return "50,000–99,999";
  return "100,000 以上";
};

function Brand({ compact = false }) {
  return (
    <div className={`brand ${compact ? "brand--compact" : ""}`}>
      <div className="brand-bars" aria-label="三條紅色直槓識別">
        <i />
        <i />
        <i />
      </div>
      <div>
        <b>KFC</b>
        {!compact && <span>LARGE ORDER VISIT ANALYTICS</span>}
      </div>
    </div>
  );
}

function Dropzone({ onFile, loading, compact = false }) {
  const input = useRef(null);
  const [dragging, setDragging] = useState(false);
  const open = () => !loading && input.current?.click();
  return (
    <div
      className={`dropzone ${compact ? "compact" : ""} ${dragging ? "dragging" : ""}`}
      role="button"
      tabIndex={0}
      aria-label="上傳 Excel 活頁簿"
      onClick={open}
      onKeyDown={(event) => {
        if (event.key === "Enter" || event.key === " ") open();
      }}
      onDragOver={(event) => event.preventDefault()}
      onDragEnter={(event) => {
        event.preventDefault();
        setDragging(true);
      }}
      onDragLeave={() => setDragging(false)}
      onDrop={(event) => {
        event.preventDefault();
        setDragging(false);
        if (event.dataTransfer.files?.[0]) onFile(event.dataTransfer.files[0]);
      }}
    >
      <input
        ref={input}
        hidden
        type="file"
        accept=".xlsx,.xls"
        onChange={(event) => event.target.files?.[0] && onFile(event.target.files[0])}
      />
      <span className="drop-icon">
        {loading ? <RefreshCw className="spin" /> : <UploadCloud />}
      </span>
      <div>
        <strong>{loading ? "正在解析活頁簿…" : compact ? "更新 Excel" : "拖曳 Excel 到這裡"}</strong>
        {!compact && (
          <>
            <span>或點擊選擇最新版大單拜訪資料</span>
            <small>支援 .xlsx／.xls，資料只在此瀏覽器處理</small>
          </>
        )}
      </div>
    </div>
  );
}

function Start({ onFile, loading, error }) {
  return (
    <main className="start">
      <section>
        <div className="eyebrow">營運決策分析工具 · BROWSER-ONLY</div>
        <h1>
          KFC 大單拜訪
          <br />
          <em>從 Excel 到營運洞察</em>
        </h1>
        <p>
          上傳最新版活頁簿後，自動辨識明細、重新判斷成功拜訪、更新階層 KPI
          與互動圖表，不需修改路徑或重新部署。
        </p>
        {error && (
          <div className="alert">
            <AlertCircle /> {error}
          </div>
        )}
        <Dropzone onFile={onFile} loading={loading} />
        <div className="privacy">
          <LockKeyhole /> 檔案不會傳到伺服器；重新整理後即清除。
        </div>
      </section>
      <aside className="features">
        {[
          ["01", FileCheck2, "自動辨識資料", "優先辨識 all 工作表與必要欄位，容忍換行與常見別名。"],
          ["02", Target, "重新判斷成功", "固定依「訂餐日期 > 電訪日期」計算，不依賴來源公式。"],
          ["03", Layers3, "四層績效分析", "區域、AC、餐廳與電訪人員逐層找出轉換落差。"],
          ["04", Download, "可稽核匯出", "輸出分析摘要、清理明細與資料診斷供回查。"],
        ].map(([step, Icon, title, copy]) => (
          <article key={step}>
            <div>
              <span>{step}</span>
              <Icon />
            </div>
            <h2>{title}</h2>
            <p>{copy}</p>
          </article>
        ))}
      </aside>
    </main>
  );
}

function Kpi({ Icon, label, value, note, tone = "", group }) {
  return (
    <article className={`kpi ${tone}`}>
      <div>
        <span>{group && <em>{group}</em>}{label}</span>
        <Icon />
      </div>
      <strong>{value}</strong>
      <small>{note}</small>
    </article>
  );
}

function ChartTip({ active, payload, label }) {
  if (!active || !payload?.length) return null;
  return (
    <div className="chart-tip">
      <strong>{label}</strong>
      {payload.map((item) => (
        <div key={item.dataKey}>
          <span style={{ color: item.color }}>{item.name}</span>
          <b>
            {String(item.dataKey).includes("Rate")
              ? p(item.value)
              : String(item.dataKey).includes("Amount")
                ? c(item.value)
                : n(item.value)}
          </b>
        </div>
      ))}
    </div>
  );
}

function Filters({ dataset, filters, setFilters }) {
  const regions = useMemo(
    () => [...new Set(dataset.rows.map((row) => row.region).filter(Boolean))].sort(),
    [dataset],
  );
  const acs = useMemo(
    () =>
      [
        ...new Set(
          dataset.rows
            .filter((row) => !filters.region || row.region === filters.region)
            .map((row) => row.ac)
            .filter(Boolean),
        ),
      ].sort(),
    [dataset, filters.region],
  );
  const stores = useMemo(
    () =>
      [
        ...new Set(
          dataset.rows
            .filter(
              (row) =>
                (!filters.region || row.region === filters.region) &&
                (!filters.ac || row.ac === filters.ac),
            )
            .map((row) => row.store)
            .filter(Boolean),
        ),
      ].sort(),
    [dataset, filters.ac, filters.region],
  );
  return (
    <section className="filters">
      <div className="filter-title">
        <Filter /> 分析條件
      </div>
      <label>
        <span>電訪日期起</span>
        <input
          type="date"
          min={dataset.dateRange.min}
          max={dataset.dateRange.max}
          value={filters.start}
          onChange={(event) => setFilters({ ...filters, start: event.target.value })}
        />
      </label>
      <label>
        <span>電訪日期迄</span>
        <input
          type="date"
          min={dataset.dateRange.min}
          max={dataset.dateRange.max}
          value={filters.end}
          onChange={(event) => setFilters({ ...filters, end: event.target.value })}
        />
      </label>
      <label>
        <span>區域</span>
        <select
          value={filters.region}
          onChange={(event) =>
            setFilters({ ...filters, region: event.target.value, ac: "", store: "" })
          }
        >
          <option value="">全部區域</option>
          {regions.map((value) => (
            <option key={value}>{value}</option>
          ))}
        </select>
      </label>
      <label>
        <span>AC</span>
        <select
          value={filters.ac}
          onChange={(event) =>
            setFilters({ ...filters, ac: event.target.value, store: "" })
          }
        >
          <option value="">全部 AC</option>
          {acs.map((value) => (
            <option key={value}>{value}</option>
          ))}
        </select>
      </label>
      <label>
        <span>餐廳</span>
        <select
          value={filters.store}
          onChange={(event) => setFilters({ ...filters, store: event.target.value })}
        >
          <option value="">全部餐廳</option>
          {stores.map((value) => (
            <option key={value}>{value}</option>
          ))}
        </select>
      </label>
      <label>
        <span>成功狀態</span>
        <select
          value={filters.success}
          onChange={(event) => setFilters({ ...filters, success: event.target.value })}
        >
          <option value="all">全部</option>
          <option value="yes">成功拜訪</option>
          <option value="no">尚未成功</option>
        </select>
      </label>
      <label className="check-filter">
        <input
          type="checkbox"
          checked={filters.excludeDuplicates}
          onChange={(event) =>
            setFilters({ ...filters, excludeDuplicates: event.target.checked })
          }
        />
        排除疑似重複
      </label>
    </section>
  );
}

function RegionBattleChart({ leaders, market }) {
  const scores = [...leaders.map((item) => item.score), market?.score].filter(Number.isFinite);
  const rawMin = Math.min(0, ...scores);
  const rawMax = Math.max(0, ...scores);
  const padding = Math.max((rawMax - rawMin) * 0.08, Math.abs(rawMax) * 0.03, 0.01);
  const min = rawMin < 0 ? rawMin - padding : 0;
  const max = rawMax + padding;
  const range = max - min || 1;
  const position = (value) => ((value - min) / range) * 100;
  const zero = position(0);
  const marketPosition = Number.isFinite(market?.score) ? position(market.score) : null;

  return (
    <div className="region-battle-chart">
      <div className="region-chart-legend"><i />六區表現<span /><b>全市場基準線</b></div>
      {leaders.map((leader) => {
        const valuePosition = Number.isFinite(leader.score) ? position(leader.score) : zero;
        const left = Math.min(zero, valuePosition);
        const width = Math.max(Math.abs(valuePosition - zero), Number.isFinite(leader.score) ? 1.5 : 0);
        return (
          <div className="region-chart-row" key={leader.name}>
            <div className="region-chart-label"><b>{leader.name}</b><small>{leader.detail}</small></div>
            <div className="region-chart-plot">
              <div className="region-chart-track">
                <i className="zero-line" style={{ left: `${zero}%` }} />
                {marketPosition != null && <i className="market-line" style={{ left: `${marketPosition}%` }} />}
                <span className={leader.score < 0 ? "negative" : ""} style={{ left: `${left}%`, width: `${width}%` }} />
              </div>
            </div>
            <strong>{leader.value}</strong>
          </div>
        );
      })}
    </div>
  );
}

function Overview({ rows, filters }) {
  const summary = useMemo(() => summarize(rows), [rows]);
  const periodLabel = `${String(filters.start || "").replaceAll("-", "/")}–${String(filters.end || "").replaceAll("-", "/")}`;
  const weeks = useMemo(() => weekly(rows), [rows]);
  const regions = useMemo(
    () => groupBy(rows, "region").sort((a, b) => b.successRate - a.successRate),
    [rows],
  );
  const customers = useMemo(
    () => groupBy(rows, "customerType").sort((a, b) => b.visits - a.visits).slice(0, 6),
    [rows],
  );
  const acLeaders = useMemo(
    () =>
      groupBy(rows, "ac")
        .filter((item) => item.name !== "未分類" && item.visits >= 20)
        .sort(
          (a, b) =>
            b.successes - a.successes ||
            b.successRate - a.successRate ||
            b.successAmount - a.successAmount,
        )
        .slice(0, 6),
    [rows],
  );
  const storeLeaders = useMemo(
    () =>
      groupBy(rows, "store")
        .filter((item) => item.name !== "未分類" && item.visits >= 5)
        .map((item) => ({
          ...item,
          topOrder: rows
            .filter((row) => row.isSuccess && row.amount != null && (row.store || "未分類") === item.name)
            .sort((a, b) => b.amount - a.amount)[0],
        }))
        .sort(
          (a, b) =>
            b.successes - a.successes ||
            b.successRate - a.successRate ||
            b.successAmount - a.successAmount,
        )
        .slice(0, 6),
    [rows],
  );
  const topSuccessOrders = useMemo(
    () => rows
      .filter((row) => row.isSuccess && row.amount != null)
      .sort((a, b) => b.amount - a.amount)
      .slice(0, 8),
    [rows],
  );
  const regionContributions = useMemo(() => {
    const totalSuccesses = rows.filter((row) => row.isSuccess).length;
    const totalSuccessAmount = rows.reduce(
      (sum, row) => sum + (row.isSuccess && row.amount != null ? row.amount : 0),
      0,
    );

    return groupBy(rows, "region")
      .filter((region) => region.name !== "未分類")
      .map((region) => {
        const regionRows = rows.filter((row) => (row.region || "未分類") === region.name);
        const topAc = groupBy(regionRows, "ac")
          .filter((item) => item.name !== "未分類")
          .sort((a, b) => b.successAmount - a.successAmount || b.successes - a.successes)[0];
        const topStore = groupBy(regionRows, "store")
          .filter((item) => item.name !== "未分類")
          .sort((a, b) => b.successAmount - a.successAmount || b.successes - a.successes)[0];

        return {
          ...region,
          successShare: totalSuccesses ? region.successes / totalSuccesses : 0,
          amountShare: totalSuccessAmount ? region.successAmount / totalSuccessAmount : 0,
          topAc,
          topAcShare: region.successAmount && topAc ? topAc.successAmount / region.successAmount : 0,
          topStore,
          topStoreShare: region.successAmount && topStore ? topStore.successAmount / region.successAmount : 0,
        };
      })
      .sort((a, b) => b.successAmount - a.successAmount);
  }, [rows]);
  const hierarchyLeaders = useMemo(() => {
    const totalAmount = summary.successAmount;
    const make = (key, minimumVisits) => groupBy(rows, key)
      .filter((item) => item.name !== "未分類" && item.visits >= minimumVisits)
      .map((item) => ({ ...item, amountShare: totalAmount ? item.successAmount / totalAmount : 0 }))
      .sort((a, b) => b.successAmount - a.successAmount || b.successes - a.successes)
      .slice(0, 5);
    return { region: make("region", 1), ac: make("ac", 20), store: make("store", 5) };
  }, [rows, summary.successAmount]);
  const levelChampions = useMemo(() => {
    const createLevel = (key, minimumVisits) => {
      const items = groupBy(rows, key)
        .filter((item) => item.name !== "未分類" && item.visits >= minimumVisits)
        .map((item) => ({ ...item, amountShare: summary.successAmount ? item.successAmount / summary.successAmount : 0 }));
      const select = (metric) => [...items].sort((a, b) => b[metric] - a[metric] || b.successAmount - a.successAmount)[0];
      return {
        amount: select("successAmount"),
        count: select("successes"),
        rate: select("successRate"),
      };
    };
    return {
      region: createLevel("region", 1),
      ac: createLevel("ac", 20),
      store: createLevel("store", 5),
    };
  }, [rows, summary.successAmount]);
  const topOrderForStore = (storeName) => rows
    .filter((row) => row.isSuccess && row.amount != null && (row.store || "未分類") === storeName)
    .sort((a, b) => b.amount - a.amount)[0];
  const quickInsights = useMemo(() => {
    const newRows = rows.filter((row) => /新|new/i.test(row.customerType || ""));
    const oldRows = rows.filter((row) => /舊|old/i.test(row.customerType || ""));
    const newSummary = summarize(newRows);
    const oldSummary = summarize(oldRows);
    const topAttribute = groupBy(rows, "attribute")
      .filter((item) => item.name !== "未分類")
      .sort((a, b) => b.successAmount - a.successAmount)[0];

    const memberComparison = (key) => {
      const member = summarize(rows.filter((row) => memberGroup(row[key]) === "會員"));
      const nonMember = summarize(rows.filter((row) => memberGroup(row[key]) === "非會員"));
      return { member, nonMember, gap: member.successRate - nonMember.successRate };
    };
    const pk = memberComparison("pkMember");
    const line = memberComparison("lineMember");
    const topMotive = groupBy(rows, "motive")
      .filter((item) => item.name !== "未分類")
      .sort((a, b) => b.successes - a.successes || b.successAmount - a.successAmount)[0];
    const successfulWithAmount = rows.filter((row) => row.isSuccess && row.amount != null);
    const bands = new Map();
    successfulWithAmount.forEach((row) => {
      const key = amountBand(row.amount);
      bands.set(key, (bands.get(key) || 0) + 1);
    });
    const topBand = [...bands].sort((a, b) => b[1] - a[1])[0];
    const lineAmountShare = summary.successAmount
      ? line.member.successAmount / summary.successAmount
      : 0;

    return [
      {
        Icon: UsersRound,
        title: "新客成功率",
        value: p(newSummary.successRate),
        text: `${n(newSummary.successes)}／${n(newSummary.visits)} 筆成功；舊客 ${p(oldSummary.successRate)}。`,
      },
      {
        Icon: Award,
        title: "成功金額最高客戶屬性",
        value: topAttribute ? attributeLabel(topAttribute.name) : "無有效資料",
        text: topAttribute ? `成功消費金額 ${c(topAttribute.successAmount)}。` : "目前沒有可判定資料。",
      },
      {
        Icon: ShieldCheck,
        title: "PK 會員成功率差",
        value: `${pk.gap >= 0 ? "+" : ""}${p(pk.gap)}`,
        text: `會員 ${p(pk.member.successRate)}｜非會員 ${p(pk.nonMember.successRate)}。`,
      },
      {
        Icon: CheckCircle2,
        title: "LINE 會員表現",
        value: p(line.member.successRate),
        text: `LINE 會員成功率｜成功金額占比 ${p(lineAmountShare)}。`,
      },
      {
        Icon: ClipboardCheck,
        title: "主要成功用餐動機",
        value: topMotive?.name || "無有效資料",
        text: topMotive ? `成功 ${n(topMotive.successes)} 筆，為目前主要動機。` : "目前沒有可判定資料。",
      },
      {
        Icon: BarChart3,
        title: "成功訂單主要金額區間",
        value: topBand?.[0] || "無有效資料",
        text: topBand ? `${n(topBand[1])} 筆，占成功訂單 ${p(topBand[1] / successfulWithAmount.length)}。` : "目前沒有有效成功金額。",
      },
    ];
  }, [rows, summary.successAmount]);
  const battlefieldSections = useMemo(() => {
    const configs = [
      { key: "region", title: "By 區域（RC）", kicker: "RC BATTLEFIELD", minimum: 0, limit: 6, meta: "六區完整比較｜全市場值為比較基準" },
      { key: "ac", title: "By AC", kicker: "AC BATTLEFIELD", minimum: 20, limit: 3, meta: "TOP 3｜最低 20 筆拜訪｜獨立統計" },
      { key: "store", title: "By 餐廳", kicker: "STORE BATTLEFIELD", minimum: 5, limit: 3, meta: "TOP 3｜最低 5 筆拜訪｜獨立統計" },
    ];
    const newCustomer = (row) => /新|new/i.test(row.customerType || "");
    const ranked = (items, limit, keepAll = false) => items
      .filter((item) => keepAll || Number.isFinite(item.score))
      .sort((a, b) => {
        const aScore = Number.isFinite(a.score) ? a.score : -Infinity;
        const bScore = Number.isFinite(b.score) ? b.score : -Infinity;
        return bScore - aScore || b.tie - a.tie;
      })
      .slice(0, limit);

    return configs.map((config) => {
      const eligible = groupBy(rows, config.key)
        .filter((item) => item.name !== "未分類" && item.visits >= config.minimum)
        .map((item) => ({ name: item.name, allRows: rows.filter((row) => (row[config.key] || "未分類") === item.name) }));
      const build = (title, Icon, evaluate) => ({
        title,
        Icon,
        market: evaluate(rows),
        leaders: ranked(eligible.map(({ name, allRows }) => ({ name, tie: summarize(allRows).successAmount, ...evaluate(allRows) })), config.limit, config.key === "region"),
      });

      const metrics = [
        build("新客成功率", UsersRound, (allRows) => {
          const result = summarize(allRows.filter(newCustomer));
          return { score: result.visits ? result.successRate : NaN, value: p(result.successRate), detail: `${n(result.successes)}／${n(result.visits)} 筆成功` };
        }),
        build("客戶屬性成功金額", Award, (allRows) => {
          const top = groupBy(allRows, "attribute").filter((item) => item.name !== "未分類").sort((a, b) => b.successAmount - a.successAmount)[0];
          return { score: top?.successAmount ?? NaN, value: c(top?.successAmount), detail: top ? attributeLabel(top.name) : "無有效資料" };
        }),
        build("PK 會員成功率差", ShieldCheck, (allRows) => {
          const member = summarize(allRows.filter((row) => memberGroup(row.pkMember) === "會員"));
          const nonMember = summarize(allRows.filter((row) => memberGroup(row.pkMember) === "非會員"));
          const gap = member.successRate - nonMember.successRate;
          return { score: member.visits && nonMember.visits ? gap : NaN, value: `${gap >= 0 ? "+" : ""}${p(gap)}`, detail: `會員 ${p(member.successRate)}｜非會員 ${p(nonMember.successRate)}` };
        }),
        build("LINE 會員表現", CheckCircle2, (allRows) => {
          const total = summarize(allRows);
          const member = summarize(allRows.filter((row) => memberGroup(row.lineMember) === "會員"));
          const share = total.successAmount ? member.successAmount / total.successAmount : 0;
          return { score: member.visits ? member.successRate : NaN, value: p(member.successRate), detail: `成功金額占該層級 ${p(share)}` };
        }),
        build("主要成功用餐動機", ClipboardCheck, (allRows) => {
          const top = groupBy(allRows, "motive").filter((item) => item.name !== "未分類").sort((a, b) => b.successes - a.successes || b.successAmount - a.successAmount)[0];
          return { score: top?.successes ?? NaN, value: `${n(top?.successes)} 筆`, detail: top?.name || "無有效資料" };
        }),
        build("成功訂單主要金額區間", BarChart3, (allRows) => {
          const counts = new Map();
          allRows.filter((row) => row.isSuccess && row.amount != null).forEach((row) => counts.set(amountBand(row.amount), (counts.get(amountBand(row.amount)) || 0) + 1));
          const top = [...counts].sort((a, b) => b[1] - a[1])[0];
          return { score: top?.[1] ?? NaN, value: `${n(top?.[1])} 筆`, detail: top?.[0] || "無有效資料" };
        }),
      ];
      return { ...config, metrics };
    });
  }, [rows]);
  const amountCoverage = summary.successes ? summary.amountRows / summary.successes : 0;
  const pending = summary.visits - summary.successes;
  const topRateRegion = regions[0];
  const topAmountRegion = regionContributions[0];
  const secondAmountRegion = regionContributions[1];
  const topSingleOrder = topSuccessOrders[0];
  const topAcByAmount = levelChampions.ac.amount;
  const topStoreByAmount = levelChampions.store.amount;
  const managementActions = [
    topAmountRegion && `區域層級由 ${topAmountRegion.name} 貢獻最高，占全市場成功金額 ${p(topAmountRegion.amountShare)}，優先拆解該區客群與開發方式。`,
    topAcByAmount && `AC 層級由 ${topAcByAmount.name} 成功金額最高，請獨立整理其高價值客群與轉換話術。`,
    topStoreByAmount && `餐廳層級由 ${topStoreByAmount.name} 成功金額最高，請追蹤訂單來源與可複製商圈。`,
    secondAmountRegion && topAmountRegion && `${topAmountRegion.name}領先第二名 ${c(topAmountRegion.successAmount - secondAmountRegion.successAmount)}，請拆解客群、話術與訂單來源。`,
    topSingleOrder && `最高單筆 ${c(topSingleOrder.amount)} 來自 ${topSingleOrder.store || "未分類"}／${attributeLabel(topSingleOrder.attribute)}／${topSingleOrder.target || "未提供對象"}，建議建立相似名單。`,
    amountCoverage < 1 && `仍有 ${n(summary.missingAmount)} 筆成功資料缺金額，先補齊以避免貢獻度低估。`,
  ].filter(Boolean).slice(0, 5);
  const pieLabel = ({ name, percent }) =>
    percent >= 0.06 ? `${name} ${p(percent)}` : "";
  return (
    <div className="stack">
      <section className="kpi-command">
        <div className="kpi-command-head">
          <div><span>MARKET KPI COMMAND</span><h2>全市場核心 KPI 戰情</h2></div>
          <b><CalendarDays />分析期間｜{periodLabel}</b>
        </div>
        <div className="kpi-grid">
          <Kpi Icon={CircleDollarSign} group="成果規模" label="成功消費金額" value={c(summary.successAmount)} note={`${n(summary.amountRows)} 筆有效金額｜平均 ${c(summary.averageAmount)}`} tone="red" />
          <Kpi Icon={Target} group="轉換效率" label="拜訪成功率" value={p(summary.successRate)} note={`成功 ${n(summary.successes)}／有效拜訪 ${n(summary.visits)}`} tone="dark" />
          <Kpi Icon={CheckCircle2} group="成果規模" label="成功拜訪筆數" value={n(summary.successes)} note={`占全部拜訪 ${p(summary.visits ? summary.successes / summary.visits : 0)}`} />
          <Kpi Icon={TrendingUp} group="單筆價值" label="成功平均金額" value={c(summary.averageAmount)} note={`成功總額 ÷ ${n(summary.amountRows)} 筆有效金額`} />
          <Kpi Icon={Activity} group="市場觸及" label="有效拜訪筆數" value={n(summary.visits)} note={`目前篩選期間共 ${n(summary.visits)} 筆`} />
          <Kpi Icon={UsersRound} group="前端互動" label="同意受訪率" value={p(summary.consentRate)} note={`${n(summary.consented)} 筆同意｜${n(summary.visits - summary.consented)} 筆未同意`} />
          <Kpi Icon={AlertCircle} group="轉換缺口" label="未轉換拜訪" value={n(pending)} note={`未轉換率 ${p(summary.visits ? pending / summary.visits : 0)}`} />
          <Kpi Icon={FileCheck2} group="資料品質" label="成功金額完整率" value={p(amountCoverage)} note={`${n(summary.amountRows)}／${n(summary.successes)} 筆具金額`} tone={amountCoverage >= .98 ? "quality" : "warning"} />
        </div>
        <div className="kpi-pulse">
          <div><span>區域金額冠軍</span><b>{topAmountRegion?.name || "—"}</b><strong>{c(topAmountRegion?.successAmount)}</strong></div>
          <div><span>AC 金額冠軍</span><b>{topAcByAmount?.name || "—"}</b><strong>{c(topAcByAmount?.successAmount)}</strong></div>
          <div><span>餐廳金額冠軍</span><b>{topStoreByAmount?.name || "—"}</b><strong>{c(topStoreByAmount?.successAmount)}</strong></div>
          <div><span>市場轉換缺口</span><b>{n(pending)} 筆未成功</b><strong>{p(summary.visits ? pending / summary.visits : 0)}</strong></div>
        </div>
      </section>
      {!rows.length ? (
        <div className="no-data">
          <Search />
          <h2>目前條件沒有資料</h2>
          <p>請調整日期或其他篩選條件。</p>
        </div>
      ) : (
        <>
          <section className="executive-snapshot">
            <article className="panel executive-leaders">
              <PanelTitle kicker="EXECUTIVE SNAPSHOT" title="各層級營運冠軍速覽" period={periodLabel} meta="四項獨立排名，不代表隸屬關係" />
              <div className="executive-leader-grid">
                {[
                  ["區域貢獻冠軍", topAmountRegion?.name, c(topAmountRegion?.successAmount), `全市場 ${p(topAmountRegion?.amountShare)}｜成功率 ${p(topAmountRegion?.successRate)}`],
                  ["AC 貢獻冠軍", topAcByAmount?.name, c(topAcByAmount?.successAmount), `成功 ${n(topAcByAmount?.successes)} 筆｜成功率 ${p(topAcByAmount?.successRate)}`],
                  ["餐廳貢獻冠軍", topStoreByAmount?.name, c(topStoreByAmount?.successAmount), `成功 ${n(topStoreByAmount?.successes)} 筆｜成功率 ${p(topStoreByAmount?.successRate)}`],
                  ["最高成功單筆", topSingleOrder?.store || "未提供餐廳", c(topSingleOrder?.amount), `${attributeLabel(topSingleOrder?.attribute)}｜${topSingleOrder?.target || "未提供對象"}`],
                ].map(([label, name, value, note], index) => (
                  <div className="executive-leader" key={label}>
                    <i>{index + 1}</i><span>{label}</span><strong>{name || "無有效資料"}</strong><b>{value}</b><small>{note}</small>
                  </div>
                ))}
              </div>
            </article>
            <article className="panel executive-actions">
              <PanelTitle kicker="DECISION ACTION" title="管理行動提醒" period={periodLabel} meta="依目前篩選結果自動產生" />
              <ol>
                {managementActions.map((action, index) => <li key={action}><b>{String(index + 1).padStart(2, "0")}</b><span>{action}</span></li>)}
              </ol>
            </article>
          </section>
          <section className="chart-grid">
            <article className="panel wide">
              <PanelTitle kicker="WEEKLY TREND" title="每週拜訪與成功趨勢" period={periodLabel} meta="週一為每週起始日" />
              <div className="chart">
                <ResponsiveContainer>
                  <AreaChart data={weeks} margin={{ top: 30, right: 18, left: -8, bottom: 0 }}>
                    <defs>
                      <linearGradient id="visit" x1="0" y1="0" x2="0" y2="1">
                        <stop offset="5%" stopColor="#111827" stopOpacity=".18" />
                        <stop offset="95%" stopColor="#111827" stopOpacity="0" />
                      </linearGradient>
                      <linearGradient id="success" x1="0" y1="0" x2="0" y2="1">
                        <stop offset="5%" stopColor="#e21b2d" stopOpacity=".3" />
                        <stop offset="95%" stopColor="#e21b2d" stopOpacity="0" />
                      </linearGradient>
                    </defs>
                    <CartesianGrid vertical={false} strokeDasharray="3 3" />
                    <XAxis dataKey="name" axisLine={false} tickLine={false} fontSize={13} />
                    <YAxis axisLine={false} tickLine={false} fontSize={12} />
                    <Tooltip content={<ChartTip />} />
                    <Legend iconType="circle" iconSize={8} />
                    <Area name="拜訪筆數" dataKey="visits" stroke="#111827" fill="url(#visit)" strokeWidth={3}>
                      <LabelList dataKey="visits" position="top" fill="#111827" fontSize={12} fontWeight={800} />
                    </Area>
                    <Area name="成功筆數" dataKey="successes" stroke="#e21b2d" fill="url(#success)" strokeWidth={3}>
                      <LabelList dataKey="successes" position="insideTop" fill="#e21b2d" fontSize={12} fontWeight={800} />
                    </Area>
                  </AreaChart>
                </ResponsiveContainer>
              </div>
            </article>
            <article className="panel">
              <PanelTitle kicker="CONVERSION" title="區域成功率" period={periodLabel} />
              <div className="chart">
                <ResponsiveContainer>
                  <BarChart data={regions} layout="vertical" margin={{ left: 4, right: 58 }}>
                    <CartesianGrid horizontal={false} />
                    <XAxis type="number" tickFormatter={p} axisLine={false} tickLine={false} fontSize={12} />
                    <YAxis type="category" dataKey="name" width={62} axisLine={false} tickLine={false} fontSize={13} fontWeight={700} />
                    <Tooltip content={<ChartTip />} />
                    <Bar name="成功率" dataKey="successRate" fill="#e21b2d" radius={[0, 8, 8, 0]} maxBarSize={26}>
                      <LabelList dataKey="successRate" position="right" formatter={p} fill="#b51224" fontSize={13} fontWeight={800} />
                    </Bar>
                  </BarChart>
                </ResponsiveContainer>
              </div>
            </article>
          </section>
          <section className="lower-grid">
            <article className="panel">
              <PanelTitle kicker="CUSTOMER MIX" title="客戶類型結構" period={periodLabel} />
              <div className="donut">
                <ResponsiveContainer>
                  <PieChart>
                    <Pie data={customers} dataKey="visits" nameKey="name" innerRadius={54} outerRadius={82} paddingAngle={2} label={pieLabel} labelLine={false}>
                      {customers.map((item, index) => (
                        <Cell key={item.name} fill={COLORS[index % COLORS.length]} />
                      ))}
                    </Pie>
                    <Tooltip formatter={(value) => `${n(value)} 筆`} />
                  </PieChart>
                </ResponsiveContainer>
                <div className="donut-total">
                  <b>{n(summary.visits)}</b>
                  <span>總拜訪</span>
                </div>
              </div>
              <div className="legend-list">
                {customers.map((item, index) => (
                  <div key={item.name}>
                    <span>
                      <i style={{ background: COLORS[index % COLORS.length] }} />
                      {item.name}
                    </span>
                    <b>{p(item.visits / summary.visits)}</b>
                  </div>
                ))}
              </div>
            </article>
            <article className="panel">
              <PanelTitle kicker="FUNNEL" title="拜訪轉換漏斗" period={periodLabel} />
              <div className="funnel">
                {[
                  ["有效拜訪", summary.visits, 100],
                  ["同意受訪", summary.consented, Math.max(35, summary.consentRate * 100)],
                  ["成功訂餐", summary.successes, Math.max(22, summary.successRate * 100)],
                ].map(([label, value, width], index) => (
                  <div key={label}>
                    <span>{label}</span>
                    <b className={`funnel-${index}`} style={{ width: `${width}%` }}>
                      {n(value)}
                    </b>
                  </div>
                ))}
              </div>
            </article>
            <article className="panel quick">
              <div className="quick-heading">
                <PanelTitle kicker="MARKET OVERVIEW" title="全市場整體洞察" period={periodLabel} meta="只呈現市場結果，不混合組織層級" />
                <div className="brand-stripes" aria-hidden="true"><i /><i /><i /></div>
              </div>
              <div className="quick-insights">
                {topRateRegion && (
                  <div className="insight insight--hero">
                    <div className="mascot mascot-0" aria-hidden="true" />
                    <div className="insight-copy">
                      <span><TrendingUp />區域成功率領先</span>
                      <strong>{topRateRegion.name} {p(topRateRegion.successRate)}</strong>
                      <p>{n(topRateRegion.successes)} 筆成功，占全市場成功筆數 {p(topRateRegion.successes / summary.successes)}。</p>
                    </div>
                  </div>
                )}
                {topAmountRegion && (
                  <div className="insight insight--hero">
                    <div className="mascot mascot-1" aria-hidden="true" />
                    <div className="insight-copy">
                      <span><CircleDollarSign />區域成功金額最高</span>
                      <strong>{topAmountRegion.name} {p(topAmountRegion.amountShare)}</strong>
                      <p>{c(topAmountRegion.successAmount)}；領先第二名 {c(topAmountRegion.successAmount - (secondAmountRegion?.successAmount || 0))}。</p>
                    </div>
                  </div>
                )}
                {quickInsights.map(({ Icon, title, value, text }, index) => (
                  <div className="insight" key={title}>
                    <div className={`mascot mascot-${index + 2}`} aria-hidden="true" />
                    <div className="insight-copy">
                      <span><Icon />{title}</span>
                      <strong>{value}</strong>
                      <p>{text}</p>
                    </div>
                  </div>
                ))}
              </div>
            </article>
          </section>
          <section className="battlefield-stack">
            {battlefieldSections.map((section) => (
              <article className="panel battlefield-panel" key={section.key}>
                <PanelTitle kicker={section.kicker} title={`${section.title}｜六大面向戰情排行`} period={periodLabel} meta={section.meta} />
                <div className="battlefield-note"><Info />本區只比較同一層級，不表示區域、AC、餐廳之間存在組織隸屬關係。</div>
                <div className="battlefield-grid">
                  {section.metrics.map(({ title, Icon, leaders, market }) => (
                    <section className="battlefield-metric" key={title}>
                      <header><Icon /><strong>{title}</strong><small>{periodLabel}</small></header>
                      {section.key === "region" && (
                        <div className="market-benchmark">
                          <span>全市場值</span>
                          <b>{market.value}</b>
                          <small>{market.detail}</small>
                        </div>
                      )}
                      {section.key === "region" ? (
                        <RegionBattleChart leaders={leaders} market={market} />
                      ) : <div className="battlefield-ranking">
                        {leaders.length ? leaders.map((leader, index) => (
                          <div className="battlefield-rank" key={leader.name}>
                            <i className={index === 0 ? "top" : ""}>{index + 1}</i>
                            <div><b>{leader.name}</b><small>{leader.detail}</small></div>
                            <strong>{leader.value}</strong>
                          </div>
                        )) : <p>目前沒有符合樣本門檻的資料。</p>}
                      </div>}
                    </section>
                  ))}
                </div>
              </article>
            ))}
          </section>
          <section className="panel contribution-panel">
            <PanelTitle kicker="REGIONAL CONTRIBUTION" title="六區成功貢獻拆解" period={periodLabel} meta="區域層級獨立比較｜成功金額由高至低" />
            <div className="contribution-table">
              <table>
                <thead>
                  <tr><th>區域</th><th>拜訪筆數</th><th>成功率</th><th>成功筆數／占比</th><th>成功金額／占比</th><th>成功平均金額</th></tr>
                </thead>
                <tbody>
                  {regionContributions.map((item, index) => (
                    <tr key={item.name}>
                      <td><i className={`performance-rank ${index < 3 ? "top" : ""}`}>{index + 1}</i><strong>{item.name}</strong></td>
                      <td><b>{n(item.visits)} 筆</b></td>
                      <td><b className="metric-emphasis">{p(item.successRate)}</b></td>
                      <td><b>{n(item.successes)} 筆</b><small>全市場 {p(item.successShare)}</small></td>
                      <td><b>{c(item.successAmount)}</b><small>全市場 {p(item.amountShare)}</small></td>
                      <td><b>{c(item.averageAmount)}</b></td>
                    </tr>
                  ))}
                </tbody>
              </table>
            </div>
          </section>
          <section className="panel hierarchy-leader-panel">
            <PanelTitle kicker="MARKET COMMAND VIEW" title="各層級關鍵指標冠軍" period={periodLabel} meta="區域、AC、餐廳分開判讀｜成功率採最低樣本門檻，避免小樣本失真" />
            <div className="hierarchy-leader-grid">
              {[
                ["區域表現優", levelChampions.region, "全市場區域", false],
                ["AC 表現優", levelChampions.ac, "最低 20 筆拜訪", false],
                ["餐廳表現優", levelChampions.store, "最低 5 筆拜訪", true],
              ].map(([title, champions, note, isStore]) => (
                <article className="hierarchy-leader-card" key={title}>
                  <header><strong>{title}</strong><span>{note}</span></header>
                  <div className="champion-list">
                    {[
                      ["成功金額冠軍", champions.amount, "amount"],
                      ["成功筆數冠軍", champions.count, "count"],
                      ["成功率冠軍", champions.rate, "rate"],
                    ].map(([label, item, metric]) => {
                      const order = isStore && item ? topOrderForStore(item.name) : null;
                      return (
                        <div className="champion-row" key={label}>
                          <div className="champion-title"><Award /><span>{label}</span></div>
                          <strong>{item?.name || "無有效資料"}</strong>
                          <div className="champion-kpis">
                            <b>{metric === "amount" ? c(item?.successAmount) : metric === "count" ? `${n(item?.successes)} 筆` : p(item?.successRate)}</b>
                            <span>成功金額 {c(item?.successAmount)}｜成功 {n(item?.successes)} 筆｜成功率 {p(item?.successRate)}</span>
                          </div>
                          {order && (
                            <div className="champion-source">
                              <em>最高單筆 {c(order.amount)}</em>
                              <span>屬性 {attributeLabel(order.attribute)}｜來自 {order.target || "未提供對象"}</span>
                              <span>{order.region || "未分類"}｜{order.ac || "未分類"}｜Excel 第 {n(order.rowNumber)} 列</span>
                            </div>
                          )}
                        </div>
                      );
                    })}
                  </div>
                </article>
              ))}
            </div>
          </section>
          <section className="highlight-grid">
            <PerformanceLeaders
              Icon={Award}
              kicker="AC HIGHLIGHTS"
              title="AC 表現亮眼名單"
              data={acLeaders}
              threshold="最低 20 筆拜訪"
              period={periodLabel}
            />
            <PerformanceLeaders
              Icon={Store}
              kicker="RESTAURANT HIGHLIGHTS"
              title="餐廳表現亮眼名單"
              data={storeLeaders}
              threshold="最低 5 筆拜訪"
              period={periodLabel}
            />
          </section>
          <section className="panel top-orders-panel">
            <PanelTitle kicker="TOP SUCCESS ORDERS" title="最佳成功訂單金額來源" period={periodLabel} meta="依成功訂單金額由高至低｜顯示前 8 筆" />
            <div className="top-orders-table">
              <table>
                <thead><tr><th>排名</th><th>成功金額</th><th>區域</th><th>AC</th><th>餐廳</th><th>屬性</th><th>客戶／電訪對象</th><th>電訪日期</th><th>訂餐日期</th><th>成功天數差</th><th>來源資料列</th></tr></thead>
                <tbody>
                  {topSuccessOrders.map((row, index) => (
                    <tr key={row.id}>
                      <td><i className={`performance-rank ${index < 3 ? "top" : ""}`}>{index + 1}</i></td>
                      <td><strong>{c(row.amount)}</strong></td><td>{row.region || "未分類"}</td><td>{row.ac || "未分類"}</td><td>{row.store || "未分類"}</td><td>{attributeLabel(row.attribute)}</td><td>{row.target || "—"}</td>
                      <td>{row.callDateKey || "—"}</td><td>{row.orderDateKey || "—"}</td>
                      <td>{row.callDate && row.orderDate ? `${Math.round((row.orderDate - row.callDate) / 86400000)} 天` : "—"}</td><td>第 {n(row.rowNumber)} 列</td>
                    </tr>
                  ))}
                </tbody>
              </table>
            </div>
          </section>
        </>
      )}
    </div>
  );
}

function PerformanceLeaders({ Icon, kicker, title, data, threshold, period }) {
  return (
    <article className="panel performance-panel">
      <PanelTitle kicker={kicker} title={title} period={period} meta={`${threshold}｜依成功筆數排序`} />
      {data.length ? (
        <div className="performance-list">
          {data.map((item, index) => {
            const topOrder = title.includes("餐廳") ? item.topOrder : null;
            return (
            <div className="performance-row" key={item.name}>
              <i className={`performance-rank ${index < 3 ? "top" : ""}`}>
                {index + 1}
              </i>
              <div className="performance-name">
                <Icon />
                <div>
                  <b>{item.name}</b>
                  <span>{n(item.visits)} 筆拜訪</span>
                </div>
              </div>
              <div className="performance-metric">
                <small>成功筆數</small>
                <strong>{n(item.successes)}</strong>
              </div>
              <div className="performance-metric rate">
                <small>拜訪成功率</small>
                <strong>{p(item.successRate)}</strong>
              </div>
              <div className="performance-metric amount">
                <small>成功金額</small>
                <strong>{c(item.successAmount)}</strong>
              </div>
              {topOrder && (
                <div className="performance-detail">
                  <small>最高單筆 {c(topOrder.amount)}</small>
                  <b>{attributeLabel(topOrder.attribute)}｜{topOrder.target || "未提供對象"}</b>
                  <span>{topOrder.region || "未分類"}／{topOrder.ac || "未分類"}／第 {n(topOrder.rowNumber)} 列</span>
                </div>
              )}
            </div>
            );
          })}
        </div>
      ) : (
        <div className="performance-empty">目前篩選範圍沒有符合最低樣本門檻的名單。</div>
      )}
    </article>
  );
}

function PanelTitle({ kicker, title, meta, period }) {
  return (
    <div className="panel-title">
      <div>
        <span>{kicker}</span>
        <div className="panel-heading-line"><h2>{title}</h2>{period && <b className="period-badge"><CalendarDays />{period}</b>}</div>
      </div>
      {meta && <small>{meta}</small>}
    </div>
  );
}

function Hierarchy({ rows, onDrill }) {
  const [dimension, setDimension] = useState("region");
  const [sort, setSort] = useState("successAmount");
  const labels = { region: "區域", ac: "AC", store: "餐廳", caller: "電訪人員" };
  const data = useMemo(
    () => groupBy(rows, dimension).sort((a, b) => b[sort] - a[sort]),
    [dimension, rows, sort],
  );
  return (
    <div className="stack">
      <div className="subnav">
        <div>
          {Object.entries(labels).map(([key, label]) => (
            <button className={dimension === key ? "active" : ""} key={key} onClick={() => setDimension(key)}>
              {label}
            </button>
          ))}
        </div>
        <label>
          排序
          <select value={sort} onChange={(event) => setSort(event.target.value)}>
            <option value="successAmount">成功消費金額</option>
            <option value="successRate">成功率</option>
            <option value="successes">成功筆數</option>
            <option value="visits">拜訪筆數</option>
          </select>
        </label>
      </div>
      <section className="chart-grid">
        <article className="panel wide">
          <PanelTitle kicker="TOP 12" title={`${labels[dimension]}績效排名`} meta="點擊表格可帶入明細" />
          <div className="chart tall">
            <ResponsiveContainer>
              <BarChart data={data.slice(0, 12)} margin={{ top: 10, right: 10, left: -8, bottom: 48 }}>
                <CartesianGrid vertical={false} />
                <XAxis dataKey="name" angle={-30} textAnchor="end" interval={0} fontSize={10} />
                <YAxis
                  tickFormatter={sort === "successAmount" ? (value) => `${Math.round(value / 1000)}K` : sort === "successRate" ? p : undefined}
                  fontSize={10}
                />
                <Tooltip content={<ChartTip />} />
                <Bar
                  name={{ successAmount: "成功消費金額", successRate: "成功率", successes: "成功筆數", visits: "拜訪筆數" }[sort]}
                  dataKey={sort}
                  fill="#e21b2d"
                  radius={[8, 8, 0, 0]}
                />
              </BarChart>
            </ResponsiveContainer>
          </div>
        </article>
        <article className="panel leader">
          <span>LEADER</span>
          <strong>01</strong>
          <h2>{data[0]?.name || "—"}</h2>
          <dl>
            <div><dt>成功消費金額</dt><dd>{c(data[0]?.successAmount)}</dd></div>
            <div><dt>成功率</dt><dd>{p(data[0]?.successRate)}</dd></div>
            <div><dt>成功／拜訪</dt><dd>{n(data[0]?.successes)}／{n(data[0]?.visits)}</dd></div>
          </dl>
        </article>
      </section>
      <article className="panel table-panel">
        <PanelTitle kicker="DETAIL RANKING" title="完整階層績效表" meta={`${n(data.length)} 個分類`} />
        <div className="table-scroll">
          <table>
            <thead>
              <tr><th>排名</th><th>{labels[dimension]}</th><th>拜訪</th><th>成功</th><th>成功率</th><th>成功消費金額</th><th>成功平均金額</th><th /></tr>
            </thead>
            <tbody>
              {data.map((item, index) => (
                <tr key={item.name} onClick={() => onDrill(dimension, item.name)}>
                  <td><i className={index < 3 ? "rank top" : "rank"}>{String(index + 1).padStart(2, "0")}</i></td>
                  <td><b>{item.name}</b></td>
                  <td>{n(item.visits)}</td>
                  <td>{n(item.successes)}</td>
                  <td>{p(item.successRate)}</td>
                  <td>{c(item.successAmount)}</td>
                  <td>{c(item.averageAmount)}</td>
                  <td className="arrow">→</td>
                </tr>
              ))}
            </tbody>
          </table>
        </div>
      </article>
    </div>
  );
}

function Details({ rows }) {
  const [search, setSearch] = useState("");
  const [page, setPage] = useState(1);
  const found = useMemo(
    () =>
      applyFilters(rows, { search }).sort(
        (a, b) => b.callDateKey.localeCompare(a.callDateKey) || b.rowNumber - a.rowNumber,
      ),
    [rows, search],
  );
  const pages = Math.max(1, Math.ceil(found.length / 20));
  const current = Math.min(page, pages);
  const visible = found.slice((current - 1) * 20, current * 20);
  return (
    <article className="panel table-panel">
      <div className="detail-head">
        <PanelTitle kicker="AUDITABLE DETAIL" title="拜訪明細" />
        <label className="search">
          <Search />
          <input
            value={search}
            placeholder="搜尋區域、AC、餐廳、對象…"
            onChange={(event) => {
              setSearch(event.target.value);
              setPage(1);
            }}
          />
          {search && <button onClick={() => setSearch("")}><X /></button>}
        </label>
      </div>
      <div className="result-note">顯示 {n(found.length)} 筆；畫面不顯示電話、手機與統編。</div>
      <div className="table-scroll">
        <table className="detail-table">
          <thead>
            <tr><th>來源列</th><th>區域</th><th>AC</th><th>餐廳</th><th>Code</th><th>電訪日期</th><th>客戶類型</th><th>電訪對象</th><th>訂餐日期</th><th>消費金額</th><th>成功</th><th>品質</th></tr>
          </thead>
          <tbody>
            {visible.map((row) => (
              <tr key={row.id}>
                <td>{row.rowNumber}</td><td>{row.region}</td><td>{row.ac}</td><td><b>{row.store}</b></td><td>{row.code || "—"}</td><td>{row.callDateKey || "—"}</td><td>{row.customerType}</td><td className="clip">{row.target || "—"}</td><td>{row.orderDateKey || "—"}</td><td>{row.amount == null ? "—" : c(row.amount)}</td>
                <td><span className={row.isSuccess ? "pill success" : "pill"}>{row.isSuccess && <Check />}{row.isSuccess ? "成功" : "未成功"}</span></td>
                <td><span className={row.issues.length ? "quality warn" : "quality"}>{row.issues.length ? `${row.issues.length} 項` : "正常"}</span></td>
              </tr>
            ))}
          </tbody>
        </table>
      </div>
      <div className="pagination">
        <span>第 {current}／{pages} 頁</span>
        <div>
          <button disabled={current <= 1} onClick={() => setPage(current - 1)}><ChevronLeft /></button>
          <button disabled={current >= pages} onClick={() => setPage(current + 1)}><ChevronRight /></button>
        </div>
      </div>
    </article>
  );
}

function Diagnostics({ dataset }) {
  const d = dataset.diagnostics;
  return (
    <div className="stack">
      <section className="diag-banner">
        <ShieldCheck />
        <div><h2>資料結構已完成檢查</h2><p>分析「{d.selectedSheet}」工作表第 {d.headerRow} 列欄名，共匯入 {n(d.importedRows)} 筆明細。</p></div>
        <b>{d.sourceSuccessAvailable && d.sourceMismatch === 0 ? "成功判斷一致" : "請檢查警示"}</b>
      </section>
      <section className="diag-grid">
        {[
          ["疑似重複資料", d.duplicateCount, "Code＋序號＋日期"],
          ["無效電訪日期", d.invalidCallDate, "影響期間分析"],
          ["缺少訂餐日期", d.missingOrderDate, "視為尚未成功"],
          ["成功但缺金額", d.successMissingAmount, "不計入成功金額"],
          ["來源成功差異", d.sourceMismatch, "系統重新比對"],
        ].map(([label, value, note]) => (
          <article key={label} className={value ? "has-issue" : ""}><span>{label}</span><b>{n(value)}</b><small>{note}</small></article>
        ))}
      </section>
      <section className="diag-layout">
        <article className="panel table-panel">
          <PanelTitle kicker="COLUMN MAPPING" title="欄位辨識結果" />
          <div className="table-scroll">
            <table>
              <thead><tr><th>分析欄位</th><th>必要性</th><th>來源欄名</th><th>狀態</th></tr></thead>
              <tbody>
                {d.headerMapping.map((item) => (
                  <tr key={item.key}><td><b>{item.label}</b></td><td>{item.required ? "必要" : "選用"}</td><td>{item.actual || "—"}</td><td><span className={item.found ? "mapped" : "optional"}>{item.found ? <CheckCircle2 /> : <Info />}{item.found ? "已辨識" : "未提供"}</span></td></tr>
                ))}
              </tbody>
            </table>
          </div>
        </article>
        <aside>
          <article className="panel source">
            <PanelTitle kicker="WORKBOOK" title="來源活頁簿" />
            <dl>
              <div><dt>檔案名稱</dt><dd>{dataset.fileName}</dd></div>
              <div><dt>工作表</dt><dd>{d.sheetNames.join("、")}</dd></div>
              <div><dt>採用工作表</dt><dd>{d.selectedSheet}</dd></div>
              <div><dt>來源欄數</dt><dd>{d.sourceColumns}</dd></div>
            </dl>
          </article>
          <article className="panel privacy-card">
            <LockKeyhole />
            <h2>隱私與資料安全</h2>
            <p>活頁簿由瀏覽器直接解析。完整清理匯出會保留來源欄位，請依公司規範妥善保存。</p>
          </article>
        </aside>
      </section>
    </div>
  );
}

export default function App() {
  const [dataset, setDataset] = useState(null);
  const [loading, setLoading] = useState(false);
  const [error, setError] = useState("");
  const [tab, setTab] = useState("overview");
  const [filters, setFilters] = useState({
    start: "",
    end: "",
    region: "",
    ac: "",
    store: "",
    success: "all",
    excludeDuplicates: true,
  });
  const upload = async (file) => {
    setError("");
    if (!/\.(xlsx|xls)$/i.test(file.name)) return setError("請選擇 .xlsx 或 .xls 活頁簿。");
    if (file.size > 50 * 1024 * 1024) return setError("檔案超過 50 MB，請先精簡後再試。");
    setLoading(true);
    try {
      const parsed = await parseFile(file);
      setDataset(parsed);
      setFilters({
        start: parsed.dateRange.min,
        end: parsed.dateRange.max,
        region: "",
        ac: "",
        store: "",
        success: "all",
        excludeDuplicates: true,
      });
      setTab("overview");
    } catch (reason) {
      setError(reason?.message || "解析失敗，請檢查檔案與欄位。");
    } finally {
      setLoading(false);
    }
  };
  const rows = useMemo(
    () => (dataset ? applyFilters(dataset.rows, filters) : []),
    [dataset, filters],
  );
  const drill = (dimension, value) => {
    if (dimension === "region") setFilters({ ...filters, region: value, ac: "", store: "" });
    if (dimension === "ac") setFilters({ ...filters, ac: value, store: "" });
    if (dimension === "store") setFilters({ ...filters, store: value });
    setTab("details");
  };
  const tabs = [
    ["overview", BarChart3, "分析總覽"],
    ["hierarchy", Layers3, "階層分析"],
    ["details", FileSpreadsheet, "拜訪明細"],
    ["diagnostics", ClipboardCheck, "資料診斷"],
  ];
  return (
    <div className="app">
      <header>
        <div className="header-inner">
          <Brand compact={Boolean(dataset)} />
          {dataset && (
            <>
              <div className="app-title"><b>KFC 大單拜訪分析</b><span>營運資料決策儀表板</span></div>
              <div className="file-meta"><FileSpreadsheet /><div><b>{dataset.fileName}</b><span>{dataset.selectedSheet} · {n(dataset.rows.length)} 筆 · {dataset.dateRange.min}～{dataset.dateRange.max}</span></div></div>
              <Dropzone onFile={upload} loading={loading} compact />
            </>
          )}
        </div>
      </header>
      {!dataset ? (
        <Start onFile={upload} loading={loading} error={error} />
      ) : (
        <>
          <nav>
            <div className="nav-inner">
              <div className="tabs">
                {tabs.map(([id, Icon, label]) => (
                  <button className={tab === id ? "active" : ""} key={id} onClick={() => setTab(id)}>
                    <Icon /> {label}
                    {id === "diagnostics" && dataset.diagnostics.issueRows.length > 0 && <i>{n(dataset.diagnostics.issueRows.length)}</i>}
                  </button>
                ))}
              </div>
              <div className="exports">
                <button onClick={() => exportWorkbook({ dataset, rows, filters })}><Download />匯出篩選結果</button>
                <button className="primary" onClick={() => exportWorkbook({ dataset, rows, filters, complete: true })}><FileSpreadsheet />匯出完整清理資料</button>
              </div>
            </div>
          </nav>
          <main className="dashboard">
            {error && <div className="alert"><AlertCircle />{error}</div>}
            <Filters dataset={dataset} filters={filters} setFilters={setFilters} />
            <div className="page-title">
              <div><span>{tabs.find(([id]) => id === tab)?.[2]}</span><h1>{filters.region || "全市場"}{filters.ac ? `／${filters.ac}` : ""}{filters.store ? `／${filters.store}` : ""}</h1></div>
              <div><Info /> 成功拜訪定義：<b>訂餐日期 &gt; 電訪日期</b></div>
            </div>
            {tab === "overview" && <Overview rows={rows} filters={filters} />}
            {tab === "hierarchy" && <Hierarchy rows={rows} onDrill={drill} />}
            {tab === "details" && <Details rows={rows} />}
            {tab === "diagnostics" && <Diagnostics dataset={dataset} />}
          </main>
        </>
      )}
      <footer><span>Carrie Chen®kfctaiwan.com</span><span>KFC 大單拜訪分析 · 瀏覽器本機運算</span></footer>
    </div>
  );
}
