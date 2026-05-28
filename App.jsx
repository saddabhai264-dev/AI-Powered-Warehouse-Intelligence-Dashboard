import { useState, useMemo, useRef, useEffect } from "react";
import * as XLSX from "xlsx";
import { BarChart, Bar, XAxis, YAxis, Tooltip, ResponsiveContainer, PieChart, Pie, Cell } from "recharts";

const DEMO_DATA = [
  {OrderID:3000,OrderDate:"2024-05-05",Branch:"Qasimabad",Company:"Guard Rice",Product:"Basmati Rice 50kg",Category:"Groceries",Quantity:80,UnitPrice_PKR:8500,Cost_PKR:7200,ExpiryDate:"2025-07-15",IsReturn:"No",ReturnQty:0,ReturnReason:""},
  {OrderID:3001,OrderDate:"2024-04-29",Branch:"Mirpur",Company:"Habib Oil",Product:"Desi Ghee 1kg",Category:"Edible Oil",Quantity:88,UnitPrice_PKR:2800,Cost_PKR:2400,ExpiryDate:"2025-07-11",IsReturn:"No",ReturnQty:0,ReturnReason:""},
  {OrderID:3002,OrderDate:"2024-10-28",Branch:"Thand Sarak",Company:"Unilever PK",Product:"Lifebuoy Soap 6pcs",Category:"Personal Care",Quantity:40,UnitPrice_PKR:380,Cost_PKR:300,ExpiryDate:"2026-06-01",IsReturn:"Yes",ReturnQty:4,ReturnReason:"Wrong Item"},
  {OrderID:3003,OrderDate:"2024-07-02",Branch:"Latifabad",Company:"Seasons",Product:"Sunflower Oil 3L",Category:"Edible Oil",Quantity:20,UnitPrice_PKR:1200,Cost_PKR:980,ExpiryDate:"2025-06-23",IsReturn:"No",ReturnQty:0,ReturnReason:""},
  {OrderID:3004,OrderDate:"2024-11-17",Branch:"Qasimabad",Company:"Unilever PK",Product:"Head & Shoulders 400ml",Category:"Personal Care",Quantity:10,UnitPrice_PKR:850,Cost_PKR:680,ExpiryDate:"2026-11-12",IsReturn:"No",ReturnQty:0,ReturnReason:""},
  {OrderID:3005,OrderDate:"2024-03-12",Branch:"Latifabad",Company:"Nestle PK",Product:"Nestle Tea 1kg",Category:"Beverages",Quantity:60,UnitPrice_PKR:1800,Cost_PKR:1500,ExpiryDate:"2025-03-11",IsReturn:"Yes",ReturnQty:8,ReturnReason:"Expired"},
  {OrderID:3006,OrderDate:"2024-06-20",Branch:"Mirpur",Company:"Tapal Ent",Product:"Tapal Tea 500g",Category:"Beverages",Quantity:45,UnitPrice_PKR:950,Cost_PKR:780,ExpiryDate:"2025-06-19",IsReturn:"No",ReturnQty:0,ReturnReason:""},
  {OrderID:3007,OrderDate:"2024-08-15",Branch:"Thand Sarak",Company:"Kolson Ltd",Product:"Kolson Spaghetti 400g",Category:"Snacks",Quantity:120,UnitPrice_PKR:180,Cost_PKR:140,ExpiryDate:"2025-02-14",IsReturn:"Yes",ReturnQty:20,ReturnReason:"Expired"},
  {OrderID:3008,OrderDate:"2024-09-01",Branch:"Mirpur",Company:"Dalda Foods",Product:"Cooking Oil 5L",Category:"Edible Oil",Quantity:35,UnitPrice_PKR:1950,Cost_PKR:1650,ExpiryDate:"2025-08-31",IsReturn:"No",ReturnQty:0,ReturnReason:""},
  {OrderID:3009,OrderDate:"2024-02-10",Branch:"Qasimabad",Company:"National Foods",Product:"Daal Maash 10kg",Category:"Groceries",Quantity:50,UnitPrice_PKR:3200,Cost_PKR:2700,ExpiryDate:"2025-02-09",IsReturn:"Yes",ReturnQty:10,ReturnReason:"Quality Issue"},
  {OrderID:3010,OrderDate:"2024-12-05",Branch:"Latifabad",Company:"P&G Pakistan",Product:"Ariel 500g",Category:"Household",Quantity:30,UnitPrice_PKR:420,Cost_PKR:340,ExpiryDate:"2026-12-04",IsReturn:"No",ReturnQty:0,ReturnReason:""},
  {OrderID:3011,OrderDate:"2024-01-22",Branch:"Mirpur",Company:"Habib Oil",Product:"Desi Ghee 1kg",Category:"Edible Oil",Quantity:55,UnitPrice_PKR:2800,Cost_PKR:2400,ExpiryDate:"2025-01-21",IsReturn:"Yes",ReturnQty:15,ReturnReason:"Expired"},
  {OrderID:3012,OrderDate:"2024-07-30",Branch:"Thand Sarak",Company:"Unilever PK",Product:"Surf Excel 1kg",Category:"Household",Quantity:25,UnitPrice_PKR:650,Cost_PKR:520,ExpiryDate:"2026-07-29",IsReturn:"No",ReturnQty:0,ReturnReason:""},
  {OrderID:3013,OrderDate:"2024-05-18",Branch:"Latifabad",Company:"Guard Rice",Product:"Basmati Rice 50kg",Category:"Groceries",Quantity:40,UnitPrice_PKR:8500,Cost_PKR:7200,ExpiryDate:"2025-05-17",IsReturn:"Yes",ReturnQty:5,ReturnReason:"Damaged Packaging"},
  {OrderID:3014,OrderDate:"2024-10-10",Branch:"Qasimabad",Company:"Unilever PK",Product:"Lifebuoy Soap 6pcs",Category:"Personal Care",Quantity:60,UnitPrice_PKR:380,Cost_PKR:300,ExpiryDate:"2026-10-09",IsReturn:"No",ReturnQty:0,ReturnReason:""},
  {OrderID:3015,OrderDate:"2024-04-03",Branch:"Mirpur",Company:"Sufi Oil",Product:"Cooking Oil 5L",Category:"Edible Oil",Quantity:70,UnitPrice_PKR:1950,Cost_PKR:1650,ExpiryDate:"2025-04-02",IsReturn:"Yes",ReturnQty:12,ReturnReason:"Quality Issue"},
  {OrderID:3016,OrderDate:"2024-11-28",Branch:"Latifabad",Company:"English Biscuit",Product:"Peek Freans Biscuit",Category:"Snacks",Quantity:200,UnitPrice_PKR:120,Cost_PKR:90,ExpiryDate:"2025-05-27",IsReturn:"No",ReturnQty:0,ReturnReason:""},
  {OrderID:3017,OrderDate:"2024-08-08",Branch:"Thand Sarak",Company:"Falak Rice",Product:"Basmati Rice 50kg",Category:"Groceries",Quantity:30,UnitPrice_PKR:8500,Cost_PKR:7200,ExpiryDate:"2025-08-07",IsReturn:"No",ReturnQty:0,ReturnReason:""},
  {OrderID:3018,OrderDate:"2024-03-25",Branch:"Qasimabad",Company:"Lipton",Product:"Lipton Yellow 200g",Category:"Beverages",Quantity:80,UnitPrice_PKR:480,Cost_PKR:380,ExpiryDate:"2025-03-24",IsReturn:"Yes",ReturnQty:6,ReturnReason:"Expired"},
  {OrderID:3019,OrderDate:"2024-09-15",Branch:"Mirpur",Company:"Reckitt PK",Product:"Vim Bar 6pcs",Category:"Household",Quantity:150,UnitPrice_PKR:280,Cost_PKR:220,ExpiryDate:"2026-09-14",IsReturn:"No",ReturnQty:0,ReturnReason:""},
  {OrderID:3020,OrderDate:"2024-06-12",Branch:"Latifabad",Company:"Nestle PK",Product:"Nestle Tea 1kg",Category:"Beverages",Quantity:40,UnitPrice_PKR:1800,Cost_PKR:1500,ExpiryDate:"2025-06-11",IsReturn:"No",ReturnQty:0,ReturnReason:""},
  {OrderID:3021,OrderDate:"2024-02-28",Branch:"Thand Sarak",Company:"GSK PK",Product:"Head & Shoulders 400ml",Category:"Personal Care",Quantity:25,UnitPrice_PKR:850,Cost_PKR:680,ExpiryDate:"2025-02-27",IsReturn:"Yes",ReturnQty:3,ReturnReason:"Wrong Item"},
  {OrderID:3022,OrderDate:"2024-10-01",Branch:"Qasimabad",Company:"Habib Oil",Product:"Desi Ghee 1kg",Category:"Edible Oil",Quantity:100,UnitPrice_PKR:2800,Cost_PKR:2400,ExpiryDate:"2025-09-30",IsReturn:"No",ReturnQty:0,ReturnReason:""},
  {OrderID:3023,OrderDate:"2024-07-14",Branch:"Mirpur",Company:"Shan Foods",Product:"Daal Chana 10kg",Category:"Groceries",Quantity:45,UnitPrice_PKR:2800,Cost_PKR:2300,ExpiryDate:"2025-07-13",IsReturn:"No",ReturnQty:0,ReturnReason:""},
  {OrderID:3024,OrderDate:"2024-01-05",Branch:"Latifabad",Company:"Seasons",Product:"Sunflower Oil 3L",Category:"Edible Oil",Quantity:60,UnitPrice_PKR:1200,Cost_PKR:980,ExpiryDate:"2024-12-04",IsReturn:"Yes",ReturnQty:18,ReturnReason:"Expired"},
  {OrderID:3025,OrderDate:"2024-12-20",Branch:"Thand Sarak",Company:"P&G Pakistan",Product:"Ariel 500g",Category:"Household",Quantity:40,UnitPrice_PKR:420,Cost_PKR:340,ExpiryDate:"2026-12-19",IsReturn:"No",ReturnQty:0,ReturnReason:""},
  {OrderID:3026,OrderDate:"2024-05-28",Branch:"Qasimabad",Company:"Continental Biscuit",Product:"Peek Freans Biscuit",Category:"Snacks",Quantity:300,UnitPrice_PKR:120,Cost_PKR:90,ExpiryDate:"2025-05-22",IsReturn:"Yes",ReturnQty:30,ReturnReason:"Expired"},
  {OrderID:3027,OrderDate:"2024-08-22",Branch:"Mirpur",Company:"Dalda Foods",Product:"Desi Ghee 1kg",Category:"Edible Oil",Quantity:75,UnitPrice_PKR:2800,Cost_PKR:2400,ExpiryDate:"2025-08-21",IsReturn:"No",ReturnQty:0,ReturnReason:""},
  {OrderID:3028,OrderDate:"2024-11-03",Branch:"Latifabad",Company:"Tapal Ent",Product:"Tapal Tea 500g",Category:"Beverages",Quantity:50,UnitPrice_PKR:950,Cost_PKR:780,ExpiryDate:"2026-11-02",IsReturn:"No",ReturnQty:0,ReturnReason:""},
  {OrderID:3029,OrderDate:"2024-04-17",Branch:"Thand Sarak",Company:"Kolson Ltd",Product:"Kolson Spaghetti 400g",Category:"Snacks",Quantity:90,UnitPrice_PKR:180,Cost_PKR:140,ExpiryDate:"2025-04-16",IsReturn:"Yes",ReturnQty:15,ReturnReason:"Damaged Packaging"},
];

const fmt = (n) => {
  if (n === undefined || n === null || isNaN(n)) return "PKR 0";
  return n >= 1000000 ? `PKR ${(n/1000000).toFixed(2)}M` : n >= 1000 ? `PKR ${(n/1000).toFixed(0)}K` : `PKR ${n}`;
};

const TODAY = new Date().toISOString().slice(0, 10);

const getExpiryStatus = (days) => {
  if (days < 0) return { label: "Expired", color: "#E74C3C", bg: "#2E0A0A" };
  if (days <= 7) return { label: "Critical", color: "#E67E22", bg: "#2A180A" };
  if (days <= 30) return { label: "Warning", color: "#F39C12", bg: "#2A210A" };
  if (days <= 60) return { label: "Watch", color: "#5DADE2", bg: "#0A1B2E" };
  return { label: "OK", color: "#2ECC71", bg: "#0B2E1A" };
};

const COLORS = ["#185FA5", "#27AE60", "#E67E22", "#E74C3C", "#9B59B6", "#16A085"];

// Normalize a key: lowercase, remove spaces/underscores/dashes/brackets/parens
const normalizeKey = (k) => String(k).toLowerCase().replace(/[\s_\-().\/\\]+/g, "");

// Fix 3: Flexible column matching — handles spaces, underscores, dashes, case differences
const getVal = (obj, keys, defaultVal = null) => {
  if (!obj || typeof obj !== "object") return defaultVal;
  const objKeys = Object.keys(obj);
  const normObjKeys = objKeys.map(k => normalizeKey(k));
  for (let requestedKey of keys) {
    const normRequested = normalizeKey(requestedKey);
    const idx = normObjKeys.indexOf(normRequested);
    if (idx !== -1 && obj[objKeys[idx]] !== undefined && obj[objKeys[idx]] !== null && obj[objKeys[idx]] !== "") {
      return obj[objKeys[idx]];
    }
  }
  // Fallback: partial/contains match
  for (let requestedKey of keys) {
    const normRequested = normalizeKey(requestedKey);
    const idx = normObjKeys.findIndex(k => k.includes(normRequested) || normRequested.includes(k));
    if (idx !== -1 && obj[objKeys[idx]] !== undefined && obj[objKeys[idx]] !== null && obj[objKeys[idx]] !== "") {
      return obj[objKeys[idx]];
    }
  }
  return defaultVal;
};

// Fix 2: Parse numbers that may have commas, PKR prefix, spaces, currency symbols
const parseNumber = (val, defaultVal = 0) => {
  if (val === null || val === undefined || val === "") return defaultVal;
  if (typeof val === "number") return isNaN(val) ? defaultVal : val;
  const cleaned = String(val).replace(/[PKRpkr₨,\s]+/g, "").trim();
  const n = parseFloat(cleaned);
  return isNaN(n) ? defaultVal : n;
};

// Fix 1: Parse Excel dates — handles serial numbers, DD-MM-YYYY, DD/MM/YYYY, ISO, JS Date objects
const parseExcelDate = (rawExpiry, fallbackDays = 90) => {
  const today = new Date();
  const fallback = new Date(today.getTime() + fallbackDays * 86400000);

  if (!rawExpiry && rawExpiry !== 0) return fallback;

  // Already a JS Date (from cellDates:true)
  if (rawExpiry instanceof Date) {
    return isNaN(rawExpiry.getTime()) ? fallback : rawExpiry;
  }

  // Excel serial number (e.g. 45318)
  if (typeof rawExpiry === "number") {
    // Excel epoch: Dec 30, 1899. Adjust for Excel's 1900 leap year bug.
    const excelEpoch = new Date(Date.UTC(1899, 11, 30));
    const d = new Date(excelEpoch.getTime() + rawExpiry * 86400000);
    return isNaN(d.getTime()) ? fallback : d;
  }

  const str = String(rawExpiry).trim();

  // DD-MM-YYYY or DD/MM/YYYY
  const dmy = str.match(/^(\d{1,2})[-\/](\d{1,2})[-\/](\d{4})$/);
  if (dmy) {
    const d = new Date(Number(dmy[3]), Number(dmy[2]) - 1, Number(dmy[1]));
    return isNaN(d.getTime()) ? fallback : d;
  }

  // MM-DD-YYYY or MM/DD/YYYY
  const mdy = str.match(/^(\d{1,2})[-\/](\d{1,2})[-\/](\d{4})$/);
  if (mdy) {
    const d = new Date(Number(mdy[3]), Number(mdy[1]) - 1, Number(mdy[2]));
    return isNaN(d.getTime()) ? fallback : d;
  }

  // YYYY-MM-DD ISO format
  const iso = new Date(str);
  if (!isNaN(iso.getTime())) return iso;

  return fallback;
};

// Fix 4: Parse IsReturn — handles Yes/No, true/false, 1/0, TRUE/FALSE, "yes", "1"
const parseIsReturn = (val) => {
  if (val === null || val === undefined) return "No";
  if (typeof val === "boolean") return val ? "Yes" : "No";
  if (typeof val === "number") return val === 1 ? "Yes" : "No";
  const s = String(val).trim().toLowerCase();
  if (s === "yes" || s === "true" || s === "1" || s === "y") return "Yes";
  return "No";
};

const parseLocalAIQuery = (query, currentData) => {
  const q = query.toLowerCase().trim();
  const expired = currentData.filter(r => r.DaysToExpiry < 0);
  const critical = currentData.filter(r => r.DaysToExpiry >= 0 && r.DaysToExpiry <= 7);
  const returns = currentData.filter(r => r.IsReturn === "Yes");
  const lossFromExpiry = expired.reduce((s, r) => s + (r.Cost_PKR * r.Quantity), 0);
  const totalRev = currentData.reduce((s, r) => s + r.Revenue, 0);
  const branchExpiry = {};
  currentData.forEach(r => {
    if (!branchExpiry[r.Branch]) branchExpiry[r.Branch] = 0;
    if (r.DaysToExpiry < 30) branchExpiry[r.Branch]++;
  });
  const worstBranch = Object.entries(branchExpiry).sort((a, b) => b[1] - a[1])[0];
  const productReturns = {};
  returns.forEach(r => { productReturns[r.Product] = (productReturns[r.Product] || 0) + r.ReturnQty; });
  const mostReturned = Object.entries(productReturns).sort((a, b) => b[1] - a[1])[0];
  const returnsLoss = returns.reduce((s, r) => s + (r.ReturnQty * r.Cost_PKR), 0);

  if (q.includes("expire") || q.includes("expiry") || q.includes("loss") || q.includes("nuqsan")) {
    return `MaxBachat ke warehouse mein abhi **${expired.length} items** expire ho chuke hain — estimated loss: **PKR ${lossFromExpiry.toLocaleString()}**.\n\n**${critical.length} items** agle 7 din mein expire hone wale hain (risk: **${fmt(critical.reduce((s,r) => s + r.Cost_PKR * r.Quantity, 0))}**).\n\n💡 In items ko discount sale ya supplier return par lagayein — foran action lein.`;
  }
  if (q.includes("branch") || q.includes("latifabad") || q.includes("mirpur") || q.includes("qasimabad") || q.includes("thand")) {
    const bs = {};
    currentData.forEach(r => {
      if (!bs[r.Branch]) bs[r.Branch] = { rev: 0, exp: 0, ret: 0 };
      bs[r.Branch].rev += r.Revenue;
      if (r.DaysToExpiry < 0) bs[r.Branch].exp++;
      if (r.IsReturn === "Yes") bs[r.Branch].ret++;
    });
    let txt = `### Branch Comparison:\n\n`;
    Object.entries(bs).forEach(([name, s]) => {
      txt += `• **${name}**: Revenue: ${fmt(s.rev)} | Expired: ${s.exp} | Returns: ${s.ret}\n`;
    });
    if (worstBranch) txt += `\n⚠️ Highest risk: **${worstBranch[0]}** — ${worstBranch[1]} items 30 din mein expire ho rahe hain.`;
    return txt;
  }
  if (q.includes("return") || q.includes("wapas") || q.includes("quality")) {
    return `Returns se **PKR ${returnsLoss.toLocaleString()}** ka loss hua hai. Total **${returns.length} orders** return hue.\n\nSabse zyada return: **${mostReturned ? mostReturned[0] : "N/A"}** (${mostReturned ? mostReturned[1] : 0} units).\n\n💡 Goods receiving par expiry date check karna laazmi karein.`;
  }
  if (q.includes("action") || q.includes("plan") || q.includes("suggestion") || q.includes("mashwara") || q.includes("kya karein")) {
    return `### Is Hafte Ka Action Plan:\n\n1. **FIFO:** ${critical.length} critical items ko pehle nikaalein\n2. **Flash Sale:** 30-day expiry wale items par 15-30% discount\n3. **Supplier Meeting:** ${mostReturned ? mostReturned[0] : "Returns"} wale supplier se baat karein\n4. **Branch Audit:** ${worstBranch ? worstBranch[0] : "Highest risk branch"} par stock taking karein`;
  }
  if (q.includes("hello") || q.includes("hi") || q.includes("assalam") || q.includes("helo")) {
    return `Assalam-o-Alaikum! Main MaxBachat ka AI Warehouse Assistant hoon.\n\nAap mujhse yeh pooch sakte hain:\n• Kaunse items expire ho rahe hain?\n• Branch comparison kya hai?\n• Returns ka kya haal hai?\n• Is hafte kya karna chahiye?`;
  }
  return `Aapke warehouse mein **${currentData.length} records** hain. Total Revenue: **${fmt(totalRev)}** | Expired: **${expired.length}** items.\n\nKoi specific sawal poochiye — branch, expiry, returns, ya action plan ke baare mein.`;
};

const generateAIInsights = (data) => {
  const expired = data.filter(r => r.DaysToExpiry < 0);
  const critical = data.filter(r => r.DaysToExpiry >= 0 && r.DaysToExpiry <= 7);
  const returns = data.filter(r => r.IsReturn === "Yes");
  const expiredReturns = returns.filter(r => r.ReturnReason === "Expired");
  const branchExpiry = {};
  data.forEach(r => {
    if (!branchExpiry[r.Branch]) branchExpiry[r.Branch] = 0;
    if (r.DaysToExpiry < 30) branchExpiry[r.Branch]++;
  });
  const worstBranch = Object.entries(branchExpiry).sort((a,b) => b[1]-a[1])[0];
  const productReturns = {};
  returns.forEach(r => { productReturns[r.Product] = (productReturns[r.Product] || 0) + 1; });
  const mostReturned = Object.entries(productReturns).sort((a,b) => b[1]-a[1])[0];
  const lossFromExpiry = expired.reduce((s,r) => s + (r.Cost_PKR * r.Quantity), 0);
  return [
    { type:"danger", icon:"⚠️", title:`${expired.length} items already expired`, detail:`Estimated loss: ${fmt(lossFromExpiry)}. FIFO system strictly follow karein — purana stock pehle nikalna chahiye.` },
    { type:"warning", icon:"🔴", title:`${critical.length} items expire in 7 days`, detail:`Fori action lein. Loss risk: ${fmt(critical.reduce((s,r)=>s+r.Cost_PKR*r.Quantity,0))}. Discount sale ya supplier return karein.` },
    { type:"info", icon:"📦", title:`${mostReturned ? mostReturned[0] : 'N/A'} — most returned product`, detail:`${mostReturned ? mostReturned[1] : 0} returns. Supplier quality check karein.` },
    { type:"warning", icon:"🏭", title:`${worstBranch ? worstBranch[0] : 'N/A'} — highest expiry risk branch`, detail:`${worstBranch ? worstBranch[1] : 0} items 30 din mein expire. Branch incharge ko alert karein.` },
    { type:"danger", icon:"🔄", title:`${expiredReturns.length} returns due to expiry`, detail:`FIFO nahi follow karne ka direct nataija. Har return avoidable loss tha.` },
  ];
};

export default function App() {
  const [tab, setTab] = useState("dashboard");
  const [data, setData] = useState(() => {
    const today = new Date();
    return DEMO_DATA.map(item => {
      const expDate = new Date(item.ExpiryDate);
      const days = Math.round((expDate - today) / 86400000);
      return { ...item, DaysToExpiry: days, Revenue: item.Quantity * item.UnitPrice_PKR };
    });
  });
  const [uploading, setUploading] = useState(false);
  const [uploadMsg, setUploadMsg] = useState("");
  const [filterBranch, setFilterBranch] = useState("All");
  const [filterStatus, setFilterStatus] = useState("All");
  const [aiLoading, setAiLoading] = useState(false);
  const [aiChat, setAiChat] = useState([]);
  const [aiInput, setAiInput] = useState("");
  const fileRef = useRef();

  useEffect(() => {
    const link = document.createElement("link");
    link.href = "https://fonts.googleapis.com/css2?family=DM+Sans:wght@400;600;700&display=swap";
    link.rel = "stylesheet";
    document.head.appendChild(link);
    const style = document.createElement("style");
    style.textContent = `
      .mb-tab { transition: all 0.2s ease-in-out; }
      .mb-tab:hover { background: rgba(243,156,18,0.2) !important; color: #F39C12 !important; }
      .mb-tab-active:hover { background: #E58E0D !important; color: #0D1B2A !important; }
      .mb-btn { transition: all 0.2s ease-in-out; }
      .mb-btn:hover { background: #E58E0D !important; transform: translateY(-1px); box-shadow: 0 4px 12px rgba(243,156,18,0.25); }
      .mb-btn:active { transform: translateY(0); }
      .mb-card { transition: all 0.25s cubic-bezier(0.25,0.8,0.25,1); }
      .mb-card:hover { transform: translateY(-2px); box-shadow: 0 8px 24px rgba(0,0,0,0.35); border-color: #F39C1240 !important; }
      .mb-kpi { transition: all 0.2s ease-in-out; }
      .mb-kpi:hover { transform: scale(1.025); box-shadow: 0 4px 15px rgba(0,0,0,0.25); }
      .mb-row-hover { transition: background 0.15s ease; }
      .mb-row-hover:hover { background: rgba(255,255,255,0.04) !important; }
      .mb-chat-msg { animation: mbSlideUp 0.25s ease-out; }
      @keyframes mbSlideUp { from { opacity:0; transform:translateY(8px); } to { opacity:1; transform:translateY(0); } }
      .mb-scrollable::-webkit-scrollbar { width:8px; height:8px; }
      .mb-scrollable::-webkit-scrollbar-track { background:#16293D; border-radius:4px; }
      .mb-scrollable::-webkit-scrollbar-thumb { background:#1B4F72; border-radius:4px; }
      .mb-scrollable::-webkit-scrollbar-thumb:hover { background:#F39C12; }
    `;
    document.head.appendChild(style);
    return () => { document.head.removeChild(link); document.head.removeChild(style); };
  }, []);

  const branches = useMemo(() => ["All", ...new Set(data.map(r => r.Branch))], [data]);
  const filtered = useMemo(() => data.filter(r => {
    const branchOk = filterBranch === "All" || r.Branch === filterBranch;
    const statusOk = filterStatus === "All" || getExpiryStatus(r.DaysToExpiry).label === filterStatus;
    return branchOk && statusOk;
  }), [data, filterBranch, filterStatus]);

  const stats = useMemo(() => ({
    expired: data.filter(r => r.DaysToExpiry < 0).length,
    critical: data.filter(r => r.DaysToExpiry >= 0 && r.DaysToExpiry <= 7).length,
    returns: data.filter(r => r.IsReturn === "Yes").length,
    lossRisk: data.filter(r => r.DaysToExpiry < 30).reduce((s,r) => s + r.Cost_PKR * r.Quantity, 0),
    totalRev: data.reduce((s,r) => s + r.Revenue, 0),
    total: data.length,
  }), [data]);

  const branchChart = useMemo(() => {
    const m = {};
    data.forEach(r => {
      if (!m[r.Branch]) m[r.Branch] = { branch: r.Branch, revenue: 0, expiring: 0, returns: 0 };
      m[r.Branch].revenue += r.Revenue;
      if (r.DaysToExpiry < 30) m[r.Branch].expiring++;
      if (r.IsReturn === "Yes") m[r.Branch].returns++;
    });
    return Object.values(m);
  }, [data]);

  const categoryChart = useMemo(() => {
    const m = {};
    data.forEach(r => { m[r.Category] = (m[r.Category] || 0) + r.Revenue; });
    return Object.entries(m).map(([name, value]) => ({ name, value }));
  }, [data]);

  const insights = useMemo(() => generateAIInsights(data), [data]);

  const handleUpload = (e) => {
    const file = e.target.files[0];
    if (!file) return;
    setUploading(true);
    setUploadMsg("File parh raha hoon...");
    const reader = new FileReader();
    reader.onload = (ev) => {
      try {
        const wb = XLSX.read(new Uint8Array(ev.target.result), { type: "array", cellDates: true });
        const rows = XLSX.utils.sheet_to_json(wb.Sheets[wb.SheetNames[0]]);
        const today = new Date();
        const processed = rows.map((r, i) => {
          // Fix 2: parseNumber handles "8,500", "PKR 1,200", spaces, currency symbols
          const unitPrice = parseNumber(getVal(r, ["UnitPrice_PKR","Unit Price (PKR)","Unit Price","UnitPrice","Price","unit_price","unitprice"], 0));
          const costPrice = parseNumber(getVal(r, ["Cost_PKR","Cost (PKR)","Cost","cost_pkr","costprice","cost"], 0));
          const qty = parseNumber(getVal(r, ["Quantity","Qty","qty","quantity"], 0));

          // Fix 1: parseExcelDate handles serial numbers, DD-MM-YYYY, DD/MM/YYYY
          const rawExpiry = getVal(r, ["ExpiryDate","Expiry Date","Expiry","expiry_date","expirydate","expiry"], null);
          const expDate = parseExcelDate(rawExpiry, 90);
          const days = Math.round((expDate - today) / 86400000);

          // Fix 4: parseIsReturn handles true/false, 1/0, yes/no strings
          const rawReturn = getVal(r, ["IsReturn","Is Return","isreturn","is_return","return"], "No");
          const isReturn = parseIsReturn(rawReturn);

          return {
            OrderID: getVal(r, ["OrderID","Order ID","orderid","order_id"]) || 4000 + i,
            OrderDate: getVal(r, ["OrderDate","Order Date","orderdate","order_date"]) || TODAY,
            Branch: String(getVal(r, ["Branch","Region","branch","region"], "Main")),
            Company: String(getVal(r, ["Company","Supplier","company","supplier"], "Unknown")),
            Product: String(getVal(r, ["Product","Item","product","item"]) || `Item ${i+1}`),
            Category: String(getVal(r, ["Category","category"], "General")),
            Quantity: qty,
            UnitPrice_PKR: unitPrice,
            Cost_PKR: costPrice,
            ExpiryDate: expDate.toISOString().slice(0,10),
            IsReturn: isReturn,
            ReturnQty: parseNumber(getVal(r, ["ReturnQty","Return Qty","returnqty","return_qty"], 0)),
            ReturnReason: String(getVal(r, ["ReturnReason","Return Reason","returnreason","return_reason"], "")),
            DaysToExpiry: days,
            Revenue: qty * unitPrice,
          };
        });
        setData(processed);
        setUploadMsg(`✅ ${processed.length} records load ho gayi!`);
      } catch (err) {
        console.error(err);
        setUploadMsg("❌ File format theek nahi — columns check karo");
      }
      setUploading(false);
    };
    reader.readAsArrayBuffer(file);
  };

  const askAI = async () => {
    if (!aiInput.trim()) return;
    const q = aiInput;
    setAiInput("");
    setAiChat(c => [...c, { role: "user", text: q }]);
    setAiLoading(true);
    const summary = `MaxBachat Warehouse Summary:\n- Total: ${stats.total} items\n- Revenue: PKR ${stats.totalRev.toLocaleString()}\n- Expired: ${stats.expired}\n- Critical (7d): ${stats.critical}\n- Returns: ${stats.returns}\n- At-risk loss: PKR ${stats.lossRisk.toLocaleString()}\n- Branches: ${branches.filter(b=>b!=="All").join(", ")}`;
    try {
      const res = await fetch("https://api.anthropic.com/v1/messages", {
        method: "POST",
        headers: { "Content-Type": "application/json" },
        body: JSON.stringify({
          model: "claude-sonnet-4-20250514",
          max_tokens: 1000,
          system: "Tu MaxBachat Pakistan ka expert warehouse AI assistant hai. Warehouse data diya gaya hai. Seedha practical Urdu/Hinglish mein jawab de. Numbers PKR mein batao. 2-3 paragraphs se zyada mat likho.",
          messages: [{ role: "user", content: `${summary}\n\nSawal: ${q}` }]
        })
      });
      const d = await res.json();
      const text = d.content?.[0]?.text;
      if (text) { setAiChat(c => [...c, { role: "ai", text }]); }
      else { throw new Error("No response"); }
    } catch {
      setAiChat(c => [...c, { role: "ai", text: parseLocalAIQuery(q, data) }]);
    }
    setAiLoading(false);
  };

  const TABS = [
    { id:"dashboard", label:"Dashboard" },
    { id:"inventory", label:"Inventory" },
    { id:"returns", label:"Returns" },
    { id:"ai", label:"AI Assistant" },
  ];

  const s = {
    wrap: { fontFamily:"'DM Sans',sans-serif", background:"#0D1B2A", minHeight:"100vh", color:"#fff" },
    header: { background:"linear-gradient(135deg,#0D1B2A,#1A3A5C)", padding:"20px 24px 0", borderBottom:"2px solid #F39C12" },
    brand: { fontSize:10, color:"#F39C12", letterSpacing:4, fontWeight:700, textTransform:"uppercase" },
    title: { margin:"4px 0 2px", fontSize:22, fontWeight:700, color:"#fff" },
    sub: { color:"#AED6F1", fontSize:11 },
    tabs: { display:"flex", gap:6, marginTop:16 },
    tab: (a) => ({ padding:"8px 20px", borderRadius:"20px 20px 0 0", border:"none", cursor:"pointer", fontSize:12, fontWeight:600, fontFamily:"'DM Sans',sans-serif", background:a?"#F39C12":"rgba(255,255,255,0.08)", color:a?"#0D1B2A":"#AED6F1" }),
    content: { padding:"20px" },
    kpiGrid: { display:"grid", gridTemplateColumns:"repeat(auto-fit,minmax(140px,1fr))", gap:12, marginBottom:20 },
    kpi: (c) => ({ background:"#16293D", borderRadius:12, padding:"14px 16px", border:`1px solid ${c}30` }),
    kpiVal: (c) => ({ fontSize:22, fontWeight:700, color:c }),
    kpiLabel: { color:"#AED6F1", fontSize:11, marginTop:4 },
    card: { background:"#16293D", borderRadius:14, padding:18, border:"1px solid #1B4F72", marginBottom:16 },
    cardTitle: { color:"#F39C12", fontSize:12, fontWeight:700, letterSpacing:1, textTransform:"uppercase", marginBottom:14 },
    grid2: { display:"grid", gridTemplateColumns:"repeat(auto-fit,minmax(280px,1fr))", gap:16 },
    badge: (s) => ({ background:s.bg, color:s.color, padding:"3px 10px", borderRadius:999, fontSize:11, fontWeight:600, display:"inline-block" }),
    table: { width:"100%", borderCollapse:"collapse", fontSize:12 },
    th: { padding:"8px 10px", color:"#F39C12", textAlign:"left", fontSize:11, borderBottom:"1px solid #1B4F72" },
    td: { padding:"10px", borderBottom:"1px solid #1B4F7240", color:"#ddd", verticalAlign:"middle" },
    insight: (t) => ({ borderLeft:`3px solid ${t==="danger"?"#E74C3C":t==="warning"?"#E67E22":"#5DADE2"}`, background:t==="danger"?"#2E0A0A":t==="warning"?"#2A1E0A":"#0A1C2E", borderRadius:"0 10px 10px 0", padding:"12px 14px", marginBottom:10 }),
    insightTitle: (t) => ({ color:t==="danger"?"#E74C3C":t==="warning"?"#E67E22":"#5DADE2", fontWeight:700, fontSize:13, marginBottom:4 }),
    insightDetail: { color:"#AED6F1", fontSize:12, lineHeight:1.6 },
    chatMsg: (r) => ({ background:r==="user"?"#1B4F72":"#142D1E", borderRadius:10, padding:"10px 14px", marginBottom:10, fontSize:13, lineHeight:1.6, color:"#fff", maxWidth:"85%", marginLeft:r==="user"?"auto":0, border:r==="user"?"1px solid #2471A3":"1px solid #1A3E25" }),
    input: { width:"100%", background:"#0D1B2A", border:"1px solid #1B4F72", borderRadius:8, padding:"10px 14px", color:"#fff", fontSize:13, fontFamily:"'DM Sans',sans-serif", outline:"none" },
    btn: { background:"#F39C12", border:"none", borderRadius:8, padding:"10px 20px", color:"#0D1B2A", fontWeight:700, fontSize:13, cursor:"pointer", fontFamily:"'DM Sans',sans-serif" },
    select: { background:"#16293D", border:"1px solid #1B4F72", borderRadius:8, padding:"6px 12px", color:"#AED6F1", fontSize:12, fontFamily:"'DM Sans',sans-serif" },
  };

  const TT = (p) => <Tooltip contentStyle={{ background:"#0D1B2A", border:"1px solid #F39C12", borderRadius:8, fontFamily:"'DM Sans',sans-serif", fontSize:11 }} cursor={{ fill:"rgba(255,255,255,0.03)" }} {...p} />;

  return (
    <div style={s.wrap}>
      <header style={s.header}>
        <div style={{ marginBottom:16 }}>
          <div style={s.brand}>📦 MaxBachat Wholesale — AI System</div>
          <h1 style={s.title}>Warehouse Intelligence Dashboard</h1>
          <div style={s.sub}>Latifabad • Mirpur • Qasimabad • Thand Sarak | Powered by AI | {TODAY}</div>
          <div style={{ marginTop:6, display:"flex", alignItems:"center", gap:8 }}>
            <span style={{ background:"#F39C12", borderRadius:999, padding:"3px 14px", fontSize:11, fontWeight:700, color:"#0D1B2A" }}>
              Designed &amp; Developed by Saddam Hussain
            </span>
            <span style={{ fontSize:10, color:"#7F8C8D" }}>© 2025 — All Rights Reserved</span>
          </div>
        </div>
        <div style={{ display:"flex", alignItems:"center", gap:12, marginBottom:14, flexWrap:"wrap" }}>
          <button onClick={() => fileRef.current?.click()} style={{ ...s.btn, fontSize:11, padding:"7px 16px", background:uploading?"#555":"#F39C12" }} className="mb-btn">
            📂 Excel Upload
          </button>
          <input ref={fileRef} type="file" accept=".xlsx,.xls" onChange={handleUpload} style={{ display:"none" }} />
          {uploadMsg && <span style={{ fontSize:11, color:uploadMsg.startsWith("✅")?"#2ECC71":"#E74C3C" }}>{uploadMsg}</span>}
          <span style={{ fontSize:10, color:"#7F8C8D", marginLeft:"auto" }}>Demo data loaded — apna Excel upload karo</span>
        </div>
        <nav style={s.tabs}>
          {TABS.map(t => (
            <button key={t.id} style={s.tab(tab===t.id)} className={`mb-tab ${tab===t.id?"mb-tab-active":""}`} onClick={() => setTab(t.id)}>
              {t.label}
            </button>
          ))}
        </nav>
      </header>

      <main style={s.content}>
        {tab === "dashboard" && (
          <section>
            <div style={s.kpiGrid}>
              {[
                { val:fmt(stats.totalRev), label:"Total Revenue", color:"#2ECC71" },
                { val:stats.expired, label:"Expired Items 🔴", color:"#E74C3C" },
                { val:stats.critical, label:"Critical (7 days) ⚠️", color:"#E67E22" },
                { val:stats.returns, label:"Returns 🔄", color:"#9B59B6" },
                { val:fmt(stats.lossRisk), label:"At-Risk Loss (30d)", color:"#E74C3C" },
                { val:stats.total, label:"Total Records", color:"#F39C12" },
              ].map((k,i) => (
                <div key={i} style={s.kpi(k.color)} className="mb-kpi">
                  <div style={s.kpiVal(k.color)}>{k.val}</div>
                  <div style={s.kpiLabel}>{k.label}</div>
                </div>
              ))}
            </div>
            <div style={s.grid2}>
              <div style={s.card} className="mb-card">
                <div style={s.cardTitle}>Branch Revenue</div>
                <ResponsiveContainer width="100%" height={180}>
                  <BarChart data={branchChart}>
                    <XAxis dataKey="branch" tick={{ fill:"#AED6F1", fontSize:10 }} axisLine={false} tickLine={false} />
                    <YAxis tickFormatter={v=>`${(v/1000000).toFixed(1)}M`} tick={{ fill:"#7F8C8D", fontSize:9 }} axisLine={false} tickLine={false} width={50} />
                    <TT formatter={v=>fmt(v)} />
                    <Bar dataKey="revenue" name="Revenue" radius={[4,4,0,0]}>
                      {branchChart.map((_,i) => <Cell key={i} fill={COLORS[i%COLORS.length]} />)}
                    </Bar>
                  </BarChart>
                </ResponsiveContainer>
              </div>
              <div style={s.card} className="mb-card">
                <div style={s.cardTitle}>Category Revenue</div>
                <ResponsiveContainer width="100%" height={180}>
                  <PieChart>
                    <Pie data={categoryChart} dataKey="value" nameKey="name" cx="50%" cy="50%" outerRadius={70} innerRadius={30} label={({name,percent})=>`${(percent*100).toFixed(0)}%`} labelLine={false} fontSize={9}>
                      {categoryChart.map((_,i) => <Cell key={i} fill={COLORS[i%COLORS.length]} />)}
                    </Pie>
                    <TT formatter={v=>fmt(v)} />
                  </PieChart>
                </ResponsiveContainer>
              </div>
            </div>
            <div style={s.card} className="mb-card">
              <div style={s.cardTitle}>🤖 AI Insights — Auto Generated</div>
              {insights.map((ins,i) => (
                <div key={i} style={s.insight(ins.type)}>
                  <div style={s.insightTitle(ins.type)}>{ins.icon} {ins.title}</div>
                  <div style={s.insightDetail}>{ins.detail}</div>
                </div>
              ))}
            </div>
          </section>
        )}

        {tab === "inventory" && (
          <section>
            <div style={{ display:"flex", gap:10, marginBottom:16, flexWrap:"wrap" }}>
              <select style={s.select} value={filterBranch} onChange={e=>setFilterBranch(e.target.value)}>
                {branches.map(b => <option key={b}>{b}</option>)}
              </select>
              <select style={s.select} value={filterStatus} onChange={e=>setFilterStatus(e.target.value)}>
                {["All","Expired","Critical","Warning","Watch","OK"].map(x => <option key={x}>{x}</option>)}
              </select>
              <span style={{ fontSize:11, color:"#7F8C8D", alignSelf:"center" }}>{filtered.length} records</span>
            </div>
            <div style={s.card} className="mb-card">
              <div style={{ overflowX:"auto" }} className="mb-scrollable">
                <table style={s.table}>
                  <thead>
                    <tr>{["Product","Branch","Company","Qty","Unit Price","Expiry Date","Days Left","Status"].map(h=><th key={h} style={s.th}>{h}</th>)}</tr>
                  </thead>
                  <tbody>
                    {filtered.slice(0,50).map((r,i) => {
                      const st = getExpiryStatus(r.DaysToExpiry);
                      return (
                        <tr key={i} style={{ background:r.DaysToExpiry<0?"#2E0A0A":r.DaysToExpiry<=7?"#2A180A":"transparent" }} className="mb-row-hover">
                          <td style={s.td}><strong style={{ color:"#fff" }}>{r.Product}</strong></td>
                          <td style={s.td}>{r.Branch}</td>
                          <td style={s.td}>{r.Company}</td>
                          <td style={s.td}>{r.Quantity}</td>
                          <td style={{ ...s.td, color:"#2ECC71", fontWeight:600 }}>PKR {r.UnitPrice_PKR.toLocaleString()}</td>
                          <td style={s.td}>{r.ExpiryDate}</td>
                          <td style={{ ...s.td, color:r.DaysToExpiry<0?"#E74C3C":r.DaysToExpiry<=30?"#E67E22":"#AED6F1", fontWeight:600 }}>
                            {r.DaysToExpiry<0?`${Math.abs(r.DaysToExpiry)}d ago`:`${r.DaysToExpiry}d`}
                          </td>
                          <td style={s.td}><span style={s.badge(st)}>{st.label}</span></td>
                        </tr>
                      );
                    })}
                  </tbody>
                </table>
              </div>
            </div>
          </section>
        )}

        {tab === "returns" && (
          <section>
            <div style={s.kpiGrid}>
              {[
                { val:stats.returns, label:"Total Returns", color:"#E74C3C" },
                { val:data.filter(r=>r.ReturnReason==="Expired").length, label:"Due to Expiry", color:"#E67E22" },
                { val:data.filter(r=>r.ReturnReason==="Quality Issue").length, label:"Quality Issues", color:"#9B59B6" },
                { val:data.filter(r=>r.ReturnReason==="Damaged Packaging").length, label:"Damaged Pack", color:"#F39C12" },
              ].map((k,i) => (
                <div key={i} style={s.kpi(k.color)} className="mb-kpi">
                  <div style={s.kpiVal(k.color)}>{k.val}</div>
                  <div style={s.kpiLabel}>{k.label}</div>
                </div>
              ))}
            </div>
            <div style={s.card} className="mb-card">
              <div style={s.cardTitle}>Return Records</div>
              <div style={{ overflowX:"auto" }} className="mb-scrollable">
                <table style={s.table}>
                  <thead>
                    <tr>{["Product","Branch","Company","Return Qty","Reason","Loss (Cost)"].map(h=><th key={h} style={s.th}>{h}</th>)}</tr>
                  </thead>
                  <tbody>
                    {data.filter(r=>r.IsReturn==="Yes").map((r,i) => (
                      <tr key={i} className="mb-row-hover">
                        <td style={s.td}>{r.Product}</td>
                        <td style={s.td}>{r.Branch}</td>
                        <td style={s.td}>{r.Company}</td>
                        <td style={{ ...s.td, color:"#E74C3C", fontWeight:600 }}>{r.ReturnQty}</td>
                        <td style={s.td}><span style={{ background:"#2E0A0A", color:"#E74C3C", padding:"3px 10px", borderRadius:999, fontSize:11, fontWeight:600 }}>{r.ReturnReason}</span></td>
                        <td style={{ ...s.td, color:"#E74C3C", fontWeight:600 }}>PKR {(r.ReturnQty*r.Cost_PKR).toLocaleString()}</td>
                      </tr>
                    ))}
                  </tbody>
                </table>
              </div>
            </div>
          </section>
        )}

        {tab === "ai" && (
          <div style={s.card} className="mb-card">
            <div style={s.cardTitle}>🤖 AI Warehouse Assistant</div>
            <div style={{ fontSize:12, color:"#AED6F1", marginBottom:16, lineHeight:1.6 }}>
              Apne warehouse ke baare mein kuch bhi poochho — expiry, returns, branch comparison, action plan.
            </div>
            {aiChat.length === 0 && (
              <div style={{ display:"flex", flexWrap:"wrap", gap:8, marginBottom:16 }}>
                {["Kaunse items expire ho rahe hain?","Latifabad branch ka kya haal hai?","Returns kum karne ke liye kya karein?","Is hafte action plan kya hai?"].map((q,i) => (
                  <button key={i} onClick={() => setAiInput(q)} style={{ background:"#0D1B2A", border:"1px solid #1B4F72", borderRadius:8, padding:"8px 12px", color:"#AED6F1", fontSize:11, cursor:"pointer", fontFamily:"'DM Sans',sans-serif" }} className="mb-btn">{q}</button>
                ))}
              </div>
            )}
            <div style={{ maxHeight:350, overflowY:"auto", marginBottom:14 }} className="mb-scrollable">
              {aiChat.map((m,i) => (
                <div key={i} style={s.chatMsg(m.role)} className="mb-chat-msg">
                  <div style={{ fontSize:10, color:m.role==="user"?"#AED6F1":"#2ECC71", fontWeight:700, marginBottom:4 }}>{m.role==="user"?"Aap":"🤖 AI"}</div>
                  <div style={{ whiteSpace:"pre-wrap" }}>{m.text}</div>
                </div>
              ))}
              {aiLoading && <div style={{ ...s.chatMsg("ai"), color:"#F39C12" }} className="mb-chat-msg">🤖 Soch raha hoon...</div>}
            </div>
            <div style={{ display:"flex", gap:8 }}>
              <input style={s.input} value={aiInput} onChange={e=>setAiInput(e.target.value)} onKeyDown={e=>e.key==="Enter"&&!aiLoading&&askAI()} placeholder="Kuch bhi poochho apne warehouse ke baare mein..." />
              <button onClick={askAI} disabled={aiLoading} style={s.btn} className="mb-btn">Send</button>
            </div>
          </div>
        )}
      </main>

      <footer style={{ textAlign:"center", padding:14, borderTop:"1px solid #1B4F72", fontSize:11, color:"#7F8C8D" }}>
        MaxBachat Wholesale • AI-Powered Warehouse System • Hyderabad, Sindh • Designed by Saddam Hussain • © 2026
      </footer>
    </div>
  );
}
