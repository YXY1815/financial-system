'use client';

import React, { useMemo, useRef, useState } from 'react';
import * as XLSX from 'xlsx';
import {
  Activity,
  AlertTriangle,
  BarChart3,
  Building2,
  Download,
  Landmark,
  LogIn,
  Plus,
  Trash2,
  TrendingUp,
  Upload,
  Wallet
} from 'lucide-react';

type LoanType = 'equal_payment' | 'equal_principal' | 'interest_only' | 'bullet';
type PageKey = 'dashboard' | 'analysis' | 'financing' | 'valuation' | 'budget';

type FinancialInput = {
  revenue: number;
  cogs: number;
  opProfit: number;
  netProfit: number;
  totalAssets: number;
  totalLiabilities: number;
  currentAssets: number;
  currentLiabilities: number;
  inventory: number;
  ar: number;
  ap: number;
  ocf: number;
  capex: number;
  cash: number;
  interest: number;
};

type BudgetProduct = {
  id: number;
  name: string;
  price: number;
  unitCost: number;
  q1: number;
  q2: number;
  q3: number;
  q4: number;
};

type ExistingLoan = {
  id: number;
  bank: string;
  product: string;
  principal: number;
  balance: number;
  annualRate: number;
  maturity: string;
  type: LoanType;
};

const benchmarks: Record<string, { grossMargin: number; currentRatio: number; debtRatio: number; roa: number; ocfRatio: number }> = {
  manufacturing: { grossMargin: 22, currentRatio: 1.6, debtRatio: 58, roa: 6, ocfRatio: 0.9 },
  retail: { grossMargin: 28, currentRatio: 1.3, debtRatio: 62, roa: 7, ocfRatio: 0.95 },
  software: { grossMargin: 68, currentRatio: 2.2, debtRatio: 35, roa: 12, ocfRatio: 1.1 },
  healthcare: { grossMargin: 45, currentRatio: 1.8, debtRatio: 48, roa: 8, ocfRatio: 1.0 },
  logistics: { grossMargin: 18, currentRatio: 1.4, debtRatio: 60, roa: 5, ocfRatio: 0.88 },
  food: { grossMargin: 25, currentRatio: 1.5, debtRatio: 55, roa: 7, ocfRatio: 0.92 }
};

const initialFin: FinancialInput = {
  revenue: 50000000,
  cogs: 36000000,
  opProfit: 6200000,
  netProfit: 4800000,
  totalAssets: 42000000,
  totalLiabilities: 24000000,
  currentAssets: 18000000,
  currentLiabilities: 12000000,
  inventory: 5000000,
  ar: 6000000,
  ap: 3800000,
  ocf: 5600000,
  capex: 1800000,
  cash: 4200000,
  interest: 900000
};

const initialBudgetProducts: BudgetProduct[] = [
  { id: 1, name: 'A产品', price: 1200, unitCost: 720, q1: 1200, q2: 1350, q3: 1500, q4: 1650 },
  { id: 2, name: 'B产品', price: 800, unitCost: 460, q1: 1800, q2: 1900, q3: 2100, q4: 2250 },
  { id: 3, name: 'C产品', price: 1500, unitCost: 880, q1: 600, q2: 700, q3: 850, q4: 900 }
];

const initialExistingLoans: ExistingLoan[] = [
  { id: 1, bank: '中国银行', product: '流动资金贷款', principal: 8000000, balance: 5200000, annualRate: 4.35, maturity: '2027-06-30', type: 'equal_payment' },
  { id: 2, bank: '建设银行', product: '固定资产贷款', principal: 12000000, balance: 9800000, annualRate: 4.9, maturity: '2028-12-31', type: 'equal_principal' },
  { id: 3, bank: '招商银行', product: '短期周转贷款', principal: 5000000, balance: 5000000, annualRate: 5.2, maturity: '2026-09-30', type: 'interest_only' }
];

const currency = (n: number) => new Intl.NumberFormat('zh-CN', { style: 'currency', currency: 'CNY', maximumFractionDigits: 0 }).format(Number.isFinite(n) ? n : 0);
const pct = (n: number, d = 2) => `${(Number.isFinite(n) ? n : 0).toFixed(d)}%`;
const quarters = ['2026Q1', '2026Q2', '2026Q3', '2026Q4'];

function calcFinancialMetrics(fin: FinancialInput) {
  const equity = Math.max(fin.totalAssets - fin.totalLiabilities, 0.0001);
  const grossMargin = fin.revenue ? ((fin.revenue - fin.cogs) / fin.revenue) * 100 : 0;
  const currentRatio = fin.currentLiabilities ? fin.currentAssets / fin.currentLiabilities : 0;
  const debtRatio = fin.totalAssets ? (fin.totalLiabilities / fin.totalAssets) * 100 : 0;
  const roa = fin.totalAssets ? (fin.netProfit / fin.totalAssets) * 100 : 0;
  const roe = fin.netProfit / equity * 100;
  const ocfToNetProfit = fin.netProfit ? fin.ocf / fin.netProfit : 0;
  const fcf = fin.ocf - fin.capex;
  return { grossMargin, currentRatio, debtRatio, roa, roe, ocfToNetProfit, fcf };
}

function calcHealthScore(metrics: ReturnType<typeof calcFinancialMetrics>, industry: string) {
  const b = benchmarks[industry];
  let score = 0;
  score += Math.min((metrics.grossMargin / b.grossMargin) * 20, 20);
  score += Math.min((metrics.currentRatio / b.currentRatio) * 15, 15);
  score += Math.min((b.debtRatio / Math.max(metrics.debtRatio, 1)) * 15, 15);
  score += Math.min((metrics.roa / b.roa) * 20, 20);
  score += Math.min((metrics.ocfToNetProfit / b.ocfRatio) * 15, 15);
  return Math.max(0, Math.min(100, Math.round(score)));
}

function calcLoanSchedule({ principal, annualRate, years, type }: { principal: number; annualRate: number; years: number; type: LoanType }) {
  const P = principal || 0;
  const r = (annualRate || 0) / 100 / 12;
  const n = Math.max(0, Math.round((years || 0) * 12));
  const rows: Array<{ period: number; payment: number; principal: number; interest: number; balance: number }> = [];
  if (!P || !n) return { rows, payment: 0, totalInterest: 0, totalPayment: 0 };
  if (r === 0) {
    const principalPaid = P / n;
    for (let i = 1; i <= n; i += 1) {
      rows.push({ period: i, payment: principalPaid, principal: principalPaid, interest: 0, balance: Math.max(0, P - principalPaid * i) });
    }
    return { rows, payment: principalPaid, totalInterest: 0, totalPayment: P };
  }
  let balance = P;
  let totalInterest = 0;
  if (type === 'equal_payment') {
    const factor = Math.pow(1 + r, n);
    const payment = (P * r * factor) / (factor - 1);
    for (let i = 1; i <= n; i += 1) {
      const interest = balance * r;
      const principalPaid = payment - interest;
      balance = Math.max(0, balance - principalPaid);
      totalInterest += interest;
      rows.push({ period: i, payment, principal: principalPaid, interest, balance });
    }
    return { rows, payment, totalInterest, totalPayment: P + totalInterest };
  }
  if (type === 'equal_principal') {
    const principalPaid = P / n;
    for (let i = 1; i <= n; i += 1) {
      const interest = balance * r;
      const payment = principalPaid + interest;
      balance = Math.max(0, balance - principalPaid);
      totalInterest += interest;
      rows.push({ period: i, payment, principal: principalPaid, interest, balance });
    }
    return { rows, payment: rows[0]?.payment || 0, totalInterest, totalPayment: P + totalInterest };
  }
  if (type === 'interest_only') {
    const monthlyInterest = P * r;
    for (let i = 1; i <= n; i += 1) {
      const principalPaid = i === n ? P : 0;
      totalInterest += monthlyInterest;
      rows.push({ period: i, payment: monthlyInterest + principalPaid, principal: principalPaid, interest: monthlyInterest, balance: i === n ? 0 : P });
    }
    return { rows, payment: monthlyInterest, totalInterest, totalPayment: P + totalInterest };
  }
  const totalInterestBullet = P * r * n;
  for (let i = 1; i <= n; i += 1) {
    rows.push({ period: i, payment: i === n ? P + totalInterestBullet : 0, principal: i === n ? P : 0, interest: i === n ? totalInterestBullet : 0, balance: i === n ? 0 : P });
  }
  return { rows, payment: rows[n - 1]?.payment || 0, totalInterest: totalInterestBullet, totalPayment: P + totalInterestBullet };
}

function calcDCF(v: { fcf: number; growth1: number; growth2: number; years1: number; years2: number; wacc: number; terminalGrowth: number }) {
  let current = v.fcf;
  const cashflows: Array<{ year: number; fcf: number; pv: number }> = [];
  let year = 1;
  const d = v.wacc / 100;
  const tg = v.terminalGrowth / 100;
  for (let i = 0; i < v.years1; i += 1, year += 1) {
    current *= 1 + v.growth1 / 100;
    cashflows.push({ year, fcf: current, pv: current / Math.pow(1 + d, year) });
  }
  for (let i = 0; i < v.years2; i += 1, year += 1) {
    current *= 1 + v.growth2 / 100;
    cashflows.push({ year, fcf: current, pv: current / Math.pow(1 + d, year) });
  }
  const terminalValue = d > tg ? (current * (1 + tg)) / (d - tg) : 0;
  const terminalPV = terminalValue / Math.pow(1 + d, Math.max(1, year - 1));
  return { cashflows, terminalValue, terminalPV, enterpriseValue: cashflows.reduce((s, c) => s + c.pv, 0) + terminalPV };
}

function calcRollingBudget(cfg: { products: BudgetProduct[]; fixedManufacturing: number; opex: number; collectionRate: number; paymentRate: number; startingCash: number }) {
  let rollingCash = cfg.startingCash;
  const qs = quarters.map((label, idx) => {
    const q = `q${idx + 1}` as 'q1' | 'q2' | 'q3' | 'q4';
    const sales = cfg.products.reduce((s, p) => s + p.price * p[q], 0);
    const variableCost = cfg.products.reduce((s, p) => s + p.unitCost * p[q], 0);
    const grossProfit = sales - variableCost;
    const operatingProfit = grossProfit - cfg.fixedManufacturing - cfg.opex;
    const cashIn = sales * (cfg.collectionRate / 100);
    const cashOut = variableCost * (cfg.paymentRate / 100) + cfg.fixedManufacturing + cfg.opex;
    const netCash = cashIn - cashOut;
    const beginningCash = rollingCash;
    const endingCash = beginningCash + netCash;
    rollingCash = endingCash;
    return { quarter: label, sales, variableCost, grossProfit, operatingProfit, cashIn, cashOut, netCash, beginningCash, endingCash, financingNeed: Math.max(0, -endingCash) };
  });
  const totals = qs.reduce((a, q) => ({
    sales: a.sales + q.sales,
    operatingProfit: a.operatingProfit + q.operatingProfit,
    netCash: a.netCash + q.netCash,
    financingNeed: Math.max(a.financingNeed, q.financingNeed),
    endingCash: q.endingCash
  }), { sales: 0, operatingProfit: 0, netCash: 0, financingNeed: 0, endingCash: cfg.startingCash });
  return { quarters: qs, totals };
}

function loanSummary(loans: ExistingLoan[]) {
  const totalPrincipal = loans.reduce((s, i) => s + i.principal, 0);
  const totalBalance = loans.reduce((s, i) => s + i.balance, 0);
  const avgRate = totalBalance ? loans.reduce((s, i) => s + i.balance * i.annualRate, 0) / totalBalance : 0;
  const shortTermBalance = loans.filter((i) => i.maturity <= '2026-12-31').reduce((s, i) => s + i.balance, 0);
  return { totalPrincipal, totalBalance, avgRate, shortTermBalance };
}

function workbookRows(workbook: XLSX.WorkBook) {
  const first = workbook.SheetNames[0];
  if (!first) return [] as Record<string, unknown>[];
  return XLSX.utils.sheet_to_json<Record<string, unknown>>(workbook.Sheets[first], { defval: '' });
}

function selfCheck() {
  const loan = calcLoanSchedule({ principal: 120000, annualRate: 6, years: 1, type: 'equal_payment' });
  console.assert(loan.rows.length === 12, 'loan periods');
  const wb = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(wb, XLSX.utils.json_to_sheet([{ revenue: 1000 }]), 'A');
  console.assert(Number(workbookRows(wb)[0].revenue) === 1000, 'xlsx round trip');
}
selfCheck();

function cardStyle(): React.CSSProperties {
  return { background: '#fff', border: '1px solid #e2e8f0', borderRadius: 16, padding: 16, boxShadow: '0 1px 2px rgba(0,0,0,0.03)' };
}

function Kpi({ title, value, sub, icon: Icon }: { title: string; value: string; sub?: string; icon: React.ComponentType<{ size?: number }> }) {
  return (
    <div style={cardStyle()}>
      <div style={{ display: 'flex', justifyContent: 'space-between', gap: 12 }}>
        <div>
          <div style={{ fontSize: 13, color: '#64748b' }}>{title}</div>
          <div style={{ fontSize: 24, fontWeight: 700, marginTop: 8 }}>{value}</div>
          {sub ? <div style={{ fontSize: 12, color: '#64748b', marginTop: 4 }}>{sub}</div> : null}
        </div>
        <div style={{ background: '#f1f5f9', borderRadius: 12, padding: 10, height: 40 }}><Icon size={20} /></div>
      </div>
    </div>
  );
}

export default function Page() {
  const fileRef = useRef<HTMLInputElement | null>(null);
  const [loggedIn, setLoggedIn] = useState(false);
  const [page, setPage] = useState<PageKey>('dashboard');
  const [industry, setIndustry] = useState('manufacturing');
  const [login, setLogin] = useState({ company: '星云制造集团', user: 'admin', password: '123456' });
  const [message, setMessage] = useState('已支持真实 .xlsx / .xls / .csv 导入导出。');
  const [fin, setFin] = useState<FinancialInput>(initialFin);
  const [loan, setLoan] = useState({ principal: 10000000, annualRate: 5.6, years: 3, type: 'equal_payment' as LoanType });
  const [valuation, setValuation] = useState({ fcf: 3800000, growth1: 12, growth2: 6, years1: 3, years2: 3, wacc: 10, terminalGrowth: 3 });
  const [budgetCfg, setBudgetCfg] = useState({ fixedManufacturing: 900000, opex: 1200000, collectionRate: 85, paymentRate: 92, startingCash: 3000000 });
  const [budgetProducts, setBudgetProducts] = useState<BudgetProduct[]>(initialBudgetProducts);
  const [existingLoans, setExistingLoans] = useState<ExistingLoan[]>(initialExistingLoans);

  const metrics = useMemo(() => calcFinancialMetrics(fin), [fin]);
  const score = useMemo(() => calcHealthScore(metrics, industry), [metrics, industry]);
  const newLoan = useMemo(() => calcLoanSchedule(loan), [loan]);
  const dcf = useMemo(() => calcDCF(valuation), [valuation]);
  const budget = useMemo(() => calcRollingBudget({ ...budgetCfg, products: budgetProducts }), [budgetCfg, budgetProducts]);
  const currentLoans = useMemo(() => loanSummary(existingLoans), [existingLoans]);

  const setNum = <T extends object>(setter: React.Dispatch<React.SetStateAction<T>>, key: keyof T) => (e: React.ChangeEvent<HTMLInputElement>) => {
    const v = e.target.value;
    setter((prev) => ({ ...prev, [key]: v === '' ? 0 : Number(v) }));
  };

  const exportAnalysisXlsx = () => {
    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, XLSX.utils.json_to_sheet([fin]), '财务输入');
    XLSX.utils.book_append_sheet(wb, XLSX.utils.json_to_sheet([
      { 指标: '毛利率', 值: pct(metrics.grossMargin) },
      { 指标: '流动比率', 值: metrics.currentRatio.toFixed(2) },
      { 指标: '资产负债率', 值: pct(metrics.debtRatio) },
      { 指标: 'ROA', 值: pct(metrics.roa) },
      { 指标: 'ROE', 值: pct(metrics.roe) },
      { 指标: '健康评分', 值: score }
    ]), '财务指标');
    XLSX.writeFile(wb, 'financial_analysis_export.xlsx');
    setMessage('已导出真实 .xlsx 财务分析文件。');
  };

  const exportBudgetXlsx = () => {
    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, XLSX.utils.json_to_sheet(budgetProducts.map((p) => ({ 产品: p.name, 单价: p.price, 单位成本: p.unitCost, Q1: p.q1, Q2: p.q2, Q3: p.q3, Q4: p.q4 }))), '预算产品');
    XLSX.utils.book_append_sheet(wb, XLSX.utils.json_to_sheet(budget.quarters.map((q) => ({ 季度: q.quarter, 销售收入: q.sales, 营业利润: q.operatingProfit, 净现金流: q.netCash, 期末现金: q.endingCash, 融资需求: q.financingNeed }))), '预算汇总');
    XLSX.writeFile(wb, 'rolling_budget_export.xlsx');
    setMessage('已导出真实 .xlsx 预算文件。');
  };

  const importFile = async (e: React.ChangeEvent<HTMLInputElement>) => {
    const file = e.target.files?.[0];
    if (!file) return;
    try {
      const name = file.name.toLowerCase();
      if (name.endsWith('.xlsx') || name.endsWith('.xls')) {
        const wb = XLSX.read(await file.arrayBuffer(), { type: 'array' });
        const rows = workbookRows(wb);
        if (!rows.length) throw new Error('empty workbook');
        const first = rows[0] as Record<string, unknown>;
        const keys = Object.keys(first);
        if (keys.includes('revenue')) {
          const next: Partial<FinancialInput> = {};
          keys.forEach((k) => { (next as Record<string, number>)[k] = Number(first[k] || 0); });
          setFin((prev) => ({ ...prev, ...next }));
          setMessage('已从 Excel 导入财务分析数据。');
        } else if (keys.includes('产品') || keys.includes('单价') || keys.includes('Q1')) {
          setBudgetProducts(rows.map((row, idx) => {
            const r = row as Record<string, unknown>;
            return {
              id: Date.now() + idx,
              name: String(r['产品'] || r.name || `导入产品${idx + 1}`),
              price: Number(r['单价'] || r.price || 0),
              unitCost: Number(r['单位成本'] || r.unitCost || 0),
              q1: Number(r.Q1 || r.q1 || 0),
              q2: Number(r.Q2 || r.q2 || 0),
              q3: Number(r.Q3 || r.q3 || 0),
              q4: Number(r.Q4 || r.q4 || 0)
            };
          }));
          setMessage('已从 Excel 导入预算产品数据。');
        } else {
          setMessage('Excel 导入失败：无法识别模板字段。');
        }
      } else {
        const text = await file.text();
        const lines = text.split(/\r?\n/).filter(Boolean);
        const header = lines[0].split(',').map((s) => s.trim());
        if (header.includes('revenue')) {
          const values = lines[1].split(',');
          const next: Partial<FinancialInput> = {};
          header.forEach((k, i) => { (next as Record<string, number>)[k] = Number(values[i] || 0); });
          setFin((prev) => ({ ...prev, ...next }));
          setMessage('已从 CSV 导入财务分析数据。');
        } else if (header.includes('产品') || header.includes('product')) {
          setBudgetProducts(lines.slice(1).map((line, idx) => {
            const [name2, price, unitCost, q1, q2, q3, q4] = line.split(',');
            return { id: Date.now() + idx, name: name2, price: Number(price), unitCost: Number(unitCost), q1: Number(q1), q2: Number(q2), q3: Number(q3), q4: Number(q4) };
          }));
          setMessage('已从 CSV 导入预算产品数据。');
        } else {
          setMessage('CSV 导入失败：无法识别模板字段。');
        }
      }
    } catch (err) {
      console.error(err);
      setMessage('导入失败：请检查文件格式。');
    } finally {
      e.target.value = '';
    }
  };

  const addLoan = () => setExistingLoans((items) => [...items, { id: Date.now(), bank: '新银行', product: '新增贷款', principal: 3000000, balance: 3000000, annualRate: 4.8, maturity: '2027-12-31', type: 'equal_payment' }]);
  const updateLoan = (id: number, field: keyof ExistingLoan, value: string) => setExistingLoans((items) => items.map((it) => it.id === id ? { ...it, [field]: ['bank', 'product', 'maturity', 'type'].includes(field) ? value : Number(value) } as ExistingLoan : it));
  const removeLoan = (id: number) => setExistingLoans((items) => items.filter((it) => it.id !== id));
  const addProduct = () => setBudgetProducts((items) => [...items, { id: Date.now(), name: `新产品${items.length + 1}`, price: 1000, unitCost: 600, q1: 500, q2: 550, q3: 600, q4: 650 }]);
  const updateProduct = (id: number, field: keyof BudgetProduct, value: string) => setBudgetProducts((items) => items.map((it) => it.id === id ? { ...it, [field]: field === 'name' ? value : Number(value) } as BudgetProduct : it));
  const removeProduct = (id: number) => setBudgetProducts((items) => items.filter((it) => it.id !== id));

  const nav = (key: PageKey, label: string) => (
    <button onClick={() => setPage(key)} style={{ width: '100%', textAlign: 'left', padding: '12px 14px', borderRadius: 12, border: 'none', background: page === key ? '#0f172a' : 'transparent', color: page === key ? '#fff' : '#334155', cursor: 'pointer' }}>{label}</button>
  );

  if (!loggedIn) {
    return (
      <div style={{ minHeight: '100vh', background: 'linear-gradient(135deg,#020617,#0f172a,#1e293b)', padding: 24, color: '#fff' }}>
        <div style={{ maxWidth: 1200, margin: '0 auto', display: 'grid', gap: 32, alignItems: 'center', minHeight: '92vh', gridTemplateColumns: '1.2fr 1fr' }}>
          <div>
            <div style={{ display: 'inline-block', background: 'rgba(255,255,255,0.08)', padding: '6px 12px', borderRadius: 999, fontSize: 12 }}>专业版财务分析与经营决策系统</div>
            <h1 style={{ fontSize: 44, lineHeight: 1.15 }}>支持真实 Excel 读写的小型网页系统</h1>
            <p style={{ color: '#cbd5e1', maxWidth: 720 }}>已内置登录页、首页仪表盘、财务分析、企业银行贷款管理、融资测算、DCF 估值、多产品多季度滚动预算，以及真实 .xlsx / .xls / .csv 导入导出。</p>
          </div>
          <div style={{ ...cardStyle(), borderRadius: 24 }}>
            <div style={{ display: 'flex', alignItems: 'center', gap: 8, fontSize: 28, fontWeight: 700 }}><LogIn size={24} />登录系统</div>
            <div style={{ marginTop: 20, display: 'grid', gap: 12 }}>
              <label>企业名称<input value={login.company} onChange={(e) => setLogin((s) => ({ ...s, company: e.target.value }))} style={inputStyle} /></label>
              <label>账号<input value={login.user} onChange={(e) => setLogin((s) => ({ ...s, user: e.target.value }))} style={inputStyle} /></label>
              <label>密码<input type="password" value={login.password} onChange={(e) => setLogin((s) => ({ ...s, password: e.target.value }))} style={inputStyle} /></label>
              <button onClick={() => setLoggedIn(true)} style={primaryBtn}>进入系统</button>
            </div>
          </div>
        </div>
      </div>
    );
  }

  return (
    <div style={{ minHeight: '100vh', display: 'flex' }}>
      <aside style={{ width: 260, background: '#fff', borderRight: '1px solid #e2e8f0', padding: 20 }}>
        <div style={{ fontSize: 12, color: '#64748b' }}>专业版系统</div>
        <h2 style={{ marginTop: 8 }}>{login.company}</h2>
        <div style={{ color: '#64748b', fontSize: 14, marginBottom: 20 }}>经营分析与决策中心</div>
        <div style={{ display: 'grid', gap: 8 }}>
          {nav('dashboard', '首页仪表盘')}
          {nav('analysis', '财务分析')}
          {nav('financing', '融资管理')}
          {nav('valuation', '企业估值')}
          {nav('budget', '全面预算')}
        </div>
        <div style={{ ...cardStyle(), marginTop: 20, background: '#0f172a', color: '#fff' }}>
          <div style={{ fontSize: 13 }}>当前登录</div>
          <div style={{ marginTop: 10, fontWeight: 700 }}>{login.user}</div>
          <button onClick={() => setLoggedIn(false)} style={{ ...secondaryBtn, background: '#e2e8f0', marginTop: 14 }}>退出登录</button>
        </div>
      </aside>
      <main style={{ flex: 1, padding: 24 }}>
        <div style={{ display: 'flex', justifyContent: 'space-between', gap: 16, alignItems: 'center', flexWrap: 'wrap' }}>
          <div>
            <h1 style={{ margin: 0 }}>{pageTitles[page]}</h1>
            <div style={{ marginTop: 8, color: '#64748b', fontSize: 14 }}>当前版本已支持真实 .xlsx / .xls / .csv 读写。</div>
          </div>
          <div style={{ display: 'flex', gap: 8, flexWrap: 'wrap' }}>
            <input ref={fileRef} type="file" accept=".xlsx,.xls,.csv" style={{ display: 'none' }} onChange={importFile} />
            <button onClick={() => fileRef.current?.click()} style={secondaryBtn}><Upload size={16} /> 导入 Excel/CSV</button>
            <button onClick={exportAnalysisXlsx} style={secondaryBtn}><Download size={16} /> 导出分析 Excel</button>
            <button onClick={exportBudgetXlsx} style={primaryBtn}><Download size={16} /> 导出预算 Excel</button>
          </div>
        </div>
        <div style={{ ...cardStyle(), marginTop: 16, fontSize: 14, color: '#475569' }}>{message}</div>

        {page === 'dashboard' && (
          <div style={{ marginTop: 20, display: 'grid', gap: 16 }}>
            <div style={{ display: 'grid', gridTemplateColumns: 'repeat(4,minmax(0,1fr))', gap: 16 }}>
              <Kpi title="财务健康评分" value={`${score} / 100`} sub="综合评分模型" icon={Activity} />
              <Kpi title="企业估值" value={currency(dcf.enterpriseValue)} sub="DCF 模型输出" icon={TrendingUp} />
              <Kpi title="融资总成本" value={currency(newLoan.totalInterest)} sub="当前融资方案" icon={Landmark} />
              <Kpi title="预算期末现金" value={currency(budget.totals.endingCash)} sub="年度滚动预算" icon={Wallet} />
            </div>
            <div style={{ display: 'grid', gridTemplateColumns: '2fr 1fr', gap: 16 }}>
              <div style={cardStyle()}>
                <h3>经营总览</h3>
                <div style={{ display: 'grid', gridTemplateColumns: 'repeat(4,minmax(0,1fr))', gap: 12 }}>
                  <Mini title="营业收入" value={currency(fin.revenue)} />
                  <Mini title="毛利率" value={pct(metrics.grossMargin)} />
                  <Mini title="年度预算收入" value={currency(budget.totals.sales)} />
                  <Mini title="最大融资缺口" value={currency(budget.totals.financingNeed)} />
                </div>
              </div>
              <div style={cardStyle()}>
                <h3>企业档案</h3>
                <InfoRow k="企业名称" v={login.company} />
                <InfoRow k="所属行业" v={industry} />
                <InfoRow k="预算产品数" v={String(budgetProducts.length)} />
                <InfoRow k="预算周期" v="4 个季度" />
              </div>
            </div>
          </div>
        )}

        {page === 'analysis' && (
          <div style={{ marginTop: 20, display: 'grid', gridTemplateColumns: '1.1fr 1.4fr', gap: 16 }}>
            <div style={cardStyle()}>
              <h3>三大报表录入</h3>
              <div style={grid2}>
                {Object.entries(fin).map(([k, v]) => (
                  <label key={k}>{k}<input type="number" value={v} onChange={setNum(setFin, k as keyof FinancialInput)} style={inputStyle} /></label>
                ))}
                <label style={{ gridColumn: '1 / -1' }}>行业对标
                  <select value={industry} onChange={(e) => setIndustry(e.target.value)} style={inputStyle as React.CSSProperties}>
                    {Object.keys(benchmarks).map((k) => <option key={k} value={k}>{k}</option>)}
                  </select>
                </label>
              </div>
            </div>
            <div style={{ display: 'grid', gap: 16 }}>
              <div style={{ display: 'grid', gridTemplateColumns: 'repeat(3,minmax(0,1fr))', gap: 16 }}>
                <Kpi title="毛利率" value={pct(metrics.grossMargin)} sub="盈利能力" icon={BarChart3} />
                <Kpi title="流动比率" value={metrics.currentRatio.toFixed(2)} sub="偿债能力" icon={Landmark} />
                <Kpi title="经营现金流/净利润" value={metrics.ocfToNetProfit.toFixed(2)} sub="现金流质量" icon={Wallet} />
              </div>
              <div style={cardStyle()}>
                <h3>财务健康度评分</h3>
                <div style={{ height: 12, background: '#e2e8f0', borderRadius: 999, overflow: 'hidden', marginTop: 12 }}><div style={{ width: `${score}%`, background: score >= 80 ? '#16a34a' : score >= 60 ? '#eab308' : '#dc2626', height: '100%' }} /></div>
                <div style={{ marginTop: 10 }}>{score < 60 ? <span style={{ color: '#dc2626' }}><AlertTriangle size={16} /> 预警：建议优化负债结构与回款能力</span> : `当前评分：${score}/100`}</div>
              </div>
            </div>
          </div>
        )}

        {page === 'financing' && (
          <div style={{ marginTop: 20, display: 'grid', gap: 16 }}>
            <div style={{ display: 'grid', gridTemplateColumns: 'repeat(4,minmax(0,1fr))', gap: 16 }}>
              <Kpi title="贷款余额总额" value={currency(currentLoans.totalBalance)} sub="企业当前银行贷款" icon={Building2} />
              <Kpi title="贷款本金总额" value={currency(currentLoans.totalPrincipal)} sub="历史授信累计" icon={Landmark} />
              <Kpi title="加权平均利率" value={pct(currentLoans.avgRate)} sub="按余额加权" icon={TrendingUp} />
              <Kpi title="年内到期余额" value={currency(currentLoans.shortTermBalance)} sub="2026 年内需关注" icon={AlertTriangle} />
            </div>
            <div style={cardStyle()}>
              <div style={{ display: 'flex', justifyContent: 'space-between', alignItems: 'center' }}>
                <h3>企业当前银行贷款管理</h3>
                <button onClick={addLoan} style={primaryBtn}><Plus size={16} /> 新增贷款</button>
              </div>
              <div style={{ overflowX: 'auto' }}>
                <table style={tableStyle}><thead><tr><th>合作银行</th><th>融资产品</th><th>贷款本金</th><th>贷款余额</th><th>年利率%</th><th>到期日</th><th>还款方式</th><th>操作</th></tr></thead><tbody>
                  {existingLoans.map((item) => (
                    <tr key={item.id}>
                      <td><input value={item.bank} onChange={(e) => updateLoan(item.id, 'bank', e.target.value)} style={cellInput} /></td>
                      <td><input value={item.product} onChange={(e) => updateLoan(item.id, 'product', e.target.value)} style={cellInput} /></td>
                      <td><input type="number" value={item.principal} onChange={(e) => updateLoan(item.id, 'principal', e.target.value)} style={cellInput} /></td>
                      <td><input type="number" value={item.balance} onChange={(e) => updateLoan(item.id, 'balance', e.target.value)} style={cellInput} /></td>
                      <td><input type="number" value={item.annualRate} onChange={(e) => updateLoan(item.id, 'annualRate', e.target.value)} style={cellInput} /></td>
                      <td><input type="date" value={item.maturity} onChange={(e) => updateLoan(item.id, 'maturity', e.target.value)} style={cellInput} /></td>
                      <td><select value={item.type} onChange={(e) => updateLoan(item.id, 'type', e.target.value)} style={cellInput as React.CSSProperties}><option value="equal_payment">等额本息</option><option value="equal_principal">等额本金</option><option value="interest_only">先息后本</option><option value="bullet">到期一次性还本付息</option></select></td>
                      <td><button onClick={() => removeLoan(item.id)} style={{ ...secondaryBtn, padding: '8px 10px' }}><Trash2 size={14} /></button></td>
                    </tr>
                  ))}
                </tbody></table>
              </div>
            </div>
            <div style={{ display: 'grid', gridTemplateColumns: '1fr 1.5fr', gap: 16 }}>
              <div style={cardStyle()}>
                <h3>新增融资测算</h3>
                <div style={{ display: 'grid', gap: 12 }}>
                  <label>贷款本金<input type="number" value={loan.principal} onChange={setNum(setLoan, 'principal')} style={inputStyle} /></label>
                  <label>年利率（%）<input type="number" value={loan.annualRate} onChange={setNum(setLoan, 'annualRate')} style={inputStyle} /></label>
                  <label>贷款年限<input type="number" value={loan.years} onChange={setNum(setLoan, 'years')} style={inputStyle} /></label>
                  <label>还款方式<select value={loan.type} onChange={(e) => setLoan((s) => ({ ...s, type: e.target.value as LoanType }))} style={inputStyle as React.CSSProperties}><option value="equal_payment">等额本息</option><option value="equal_principal">等额本金</option><option value="interest_only">先息后本</option><option value="bullet">到期一次性还本付息</option></select></label>
                </div>
              </div>
              <div style={cardStyle()}>
                <h3>还款计划表（前24期）</h3>
                <div style={{ display: 'grid', gridTemplateColumns: 'repeat(3,minmax(0,1fr))', gap: 12, marginBottom: 12 }}>
                  <Mini title="首期月供" value={currency(newLoan.payment)} />
                  <Mini title="总利息" value={currency(newLoan.totalInterest)} />
                  <Mini title="总还款额" value={currency(newLoan.totalPayment)} />
                </div>
                <div style={{ maxHeight: 420, overflow: 'auto' }}>
                  <table style={tableStyle}><thead><tr><th>期数</th><th>月供</th><th>本金</th><th>利息</th><th>剩余本金</th></tr></thead><tbody>
                    {newLoan.rows.slice(0, 24).map((r) => <tr key={r.period}><td>{r.period}</td><td>{currency(r.payment)}</td><td>{currency(r.principal)}</td><td>{currency(r.interest)}</td><td>{currency(r.balance)}</td></tr>)}
                  </tbody></table>
                </div>
              </div>
            </div>
          </div>
        )}

        {page === 'valuation' && (
          <div style={{ marginTop: 20, display: 'grid', gridTemplateColumns: '1fr 1.5fr', gap: 16 }}>
            <div style={cardStyle()}>
              <h3>DCF 参数</h3>
              <div style={grid2}>
                {Object.entries(valuation).map(([k, v]) => <label key={k}>{k}<input type="number" value={v} onChange={setNum(setValuation, k as keyof typeof valuation)} style={inputStyle} /></label>)}
              </div>
            </div>
            <div style={cardStyle()}>
              <div style={{ display: 'grid', gridTemplateColumns: 'repeat(3,minmax(0,1fr))', gap: 16, marginBottom: 16 }}>
                <Kpi title="企业价值 EV" value={currency(dcf.enterpriseValue)} icon={TrendingUp} />
                <Kpi title="终值现值" value={currency(dcf.terminalPV)} icon={BarChart3} />
                <Kpi title="终值" value={currency(dcf.terminalValue)} icon={Wallet} />
              </div>
              <table style={tableStyle}><thead><tr><th>年份</th><th>预测 FCF</th><th>现值 PV</th></tr></thead><tbody>
                {dcf.cashflows.map((cf) => <tr key={cf.year}><td>第 {cf.year} 年</td><td>{currency(cf.fcf)}</td><td>{currency(cf.pv)}</td></tr>)}
              </tbody></table>
            </div>
          </div>
        )}

        {page === 'budget' && (
          <div style={{ marginTop: 20, display: 'grid', gap: 16 }}>
            <div style={{ display: 'grid', gridTemplateColumns: '1fr 1.7fr', gap: 16 }}>
              <div style={cardStyle()}>
                <h3>滚动预算配置</h3>
                <div style={{ display: 'grid', gap: 12 }}>
                  <label>固定制造费用/季度<input type="number" value={budgetCfg.fixedManufacturing} onChange={setNum(setBudgetCfg, 'fixedManufacturing')} style={inputStyle} /></label>
                  <label>期间费用/季度<input type="number" value={budgetCfg.opex} onChange={setNum(setBudgetCfg, 'opex')} style={inputStyle} /></label>
                  <label>回款率%<input type="number" value={budgetCfg.collectionRate} onChange={setNum(setBudgetCfg, 'collectionRate')} style={inputStyle} /></label>
                  <label>付款率%<input type="number" value={budgetCfg.paymentRate} onChange={setNum(setBudgetCfg, 'paymentRate')} style={inputStyle} /></label>
                  <label>期初现金<input type="number" value={budgetCfg.startingCash} onChange={setNum(setBudgetCfg, 'startingCash')} style={inputStyle} /></label>
                </div>
              </div>
              <div style={cardStyle()}>
                <div style={{ display: 'flex', justifyContent: 'space-between', alignItems: 'center' }}>
                  <h3>产品预算明细</h3>
                  <button onClick={addProduct} style={primaryBtn}><Plus size={16} /> 新增产品</button>
                </div>
                <div style={{ overflowX: 'auto' }}>
                  <table style={tableStyle}><thead><tr><th>产品</th><th>单价</th><th>单位成本</th><th>Q1销量</th><th>Q2销量</th><th>Q3销量</th><th>Q4销量</th><th>操作</th></tr></thead><tbody>
                    {budgetProducts.map((p) => <tr key={p.id}><td><input value={p.name} onChange={(e) => updateProduct(p.id, 'name', e.target.value)} style={cellInput} /></td><td><input type="number" value={p.price} onChange={(e) => updateProduct(p.id, 'price', e.target.value)} style={cellInput} /></td><td><input type="number" value={p.unitCost} onChange={(e) => updateProduct(p.id, 'unitCost', e.target.value)} style={cellInput} /></td><td><input type="number" value={p.q1} onChange={(e) => updateProduct(p.id, 'q1', e.target.value)} style={cellInput} /></td><td><input type="number" value={p.q2} onChange={(e) => updateProduct(p.id, 'q2', e.target.value)} style={cellInput} /></td><td><input type="number" value={p.q3} onChange={(e) => updateProduct(p.id, 'q3', e.target.value)} style={cellInput} /></td><td><input type="number" value={p.q4} onChange={(e) => updateProduct(p.id, 'q4', e.target.value)} style={cellInput} /></td><td><button onClick={() => removeProduct(p.id)} style={{ ...secondaryBtn, padding: '8px 10px' }}><Trash2 size={14} /></button></td></tr>)}
                  </tbody></table>
                </div>
              </div>
            </div>
            <div style={{ display: 'grid', gridTemplateColumns: 'repeat(4,minmax(0,1fr))', gap: 16 }}>
              <Kpi title="年度销售收入" value={currency(budget.totals.sales)} icon={BarChart3} />
              <Kpi title="年度营业利润" value={currency(budget.totals.operatingProfit)} icon={TrendingUp} />
              <Kpi title="年度净现金流" value={currency(budget.totals.netCash)} icon={Wallet} />
              <Kpi title="最大融资缺口" value={currency(budget.totals.financingNeed)} icon={Landmark} />
            </div>
            <div style={cardStyle()}>
              <h3>季度滚动预算结果</h3>
              <table style={tableStyle}><thead><tr><th>季度</th><th>销售收入</th><th>营业利润</th><th>现金流入</th><th>现金流出</th><th>净现金流</th><th>期初现金</th><th>期末现金</th><th>融资需求</th></tr></thead><tbody>
                {budget.quarters.map((q) => <tr key={q.quarter}><td>{q.quarter}</td><td>{currency(q.sales)}</td><td>{currency(q.operatingProfit)}</td><td>{currency(q.cashIn)}</td><td>{currency(q.cashOut)}</td><td>{currency(q.netCash)}</td><td>{currency(q.beginningCash)}</td><td>{currency(q.endingCash)}</td><td>{currency(q.financingNeed)}</td></tr>)}
              </tbody></table>
            </div>
          </div>
        )}
      </main>
    </div>
  );
}

const pageTitles: Record<PageKey, string> = {
  dashboard: '首页仪表盘',
  analysis: '财务分析模块',
  financing: '融资管理模块',
  valuation: '企业估值模块',
  budget: '全面预算模块'
};

function InfoRow({ k, v }: { k: string; v: string }) {
  return <div style={{ display: 'flex', justifyContent: 'space-between', padding: '6px 0', borderBottom: '1px solid #f1f5f9' }}><span style={{ color: '#64748b' }}>{k}</span><span style={{ fontWeight: 600 }}>{v}</span></div>;
}

function Mini({ title, value }: { title: string; value: string }) {
  return <div style={{ background: '#f8fafc', border: '1px solid #e2e8f0', borderRadius: 14, padding: 14 }}><div style={{ fontSize: 13, color: '#64748b' }}>{title}</div><div style={{ marginTop: 8, fontWeight: 700 }}>{value}</div></div>;
}

const inputStyle: React.CSSProperties = { width: '100%', marginTop: 6, border: '1px solid #cbd5e1', borderRadius: 10, padding: '10px 12px', background: '#fff' };
const primaryBtn: React.CSSProperties = { display: 'inline-flex', alignItems: 'center', gap: 6, border: 'none', borderRadius: 12, padding: '10px 14px', background: '#0f172a', color: '#fff', cursor: 'pointer' };
const secondaryBtn: React.CSSProperties = { display: 'inline-flex', alignItems: 'center', gap: 6, border: '1px solid #cbd5e1', borderRadius: 12, padding: '10px 14px', background: '#fff', color: '#0f172a', cursor: 'pointer' };
const grid2: React.CSSProperties = { display: 'grid', gridTemplateColumns: '1fr 1fr', gap: 12 };
const tableStyle: React.CSSProperties = { width: '100%', borderCollapse: 'collapse', marginTop: 10 };
const cellInput: React.CSSProperties = { width: '100%', border: '1px solid #cbd5e1', borderRadius: 8, padding: '8px 10px' };
