# AI-Powered-Warehouse-Intelligence-Dashboard
AI-Powered Warehouse Intelligence Dashboard for MaxBachat Wholesale, Hyderabad


![React](https://img.shields.io/badge/React-18-61DAFB?style=flat&logo=react)
![License](https://img.shields.io/badge/License-MIT-green)
![Status](https://img.shields.io/badge/Status-Live-brightgreen)

---

## 🧠 What Is This?

A real-time warehouse intelligence dashboard built for **MaxBachat Wholesale**, 
covering 4 branches across Hyderabad — Latifabad, Mirpur, Qasimabad, and Thand Sarak.

It tracks inventory expiry, returns, revenue, and gives AI-powered insights 
in **Urdu/Hinglish** — designed for Pakistan's wholesale market.

---

## ✨ Features

- 📊 **Live KPI Dashboard** — Revenue, Expired Items, Critical Alerts, Returns
- ⏰ **Expiry Tracking** — Color-coded: Expired / Critical / Warning / Watch / OK
- 🔄 **Returns Management** — Reason tracking with loss calculation (PKR)
- 🤖 **AI Warehouse Assistant** — Ask questions in Urdu/Hinglish, get smart answers
- 📂 **Excel Upload** — Upload any `.xlsx` file with flexible column name matching
- 📈 **Branch & Category Charts** — Recharts-powered visual analytics

---

## 🛠️ Tech Stack

| Technology | Purpose |
|---|---|
| React 18 | Frontend UI |
| XLSX (SheetJS) | Excel file parsing |
| Recharts | Data visualization |
| Claude API (Anthropic) | AI assistant |

---

## 🐛 Bugs Fixed (v2.0)

| Bug | Root Cause | Fix Applied |
|---|---|---|
| DaysToExpiry showing 0 | Excel serial numbers & DD-MM-YYYY not parsed | `parseExcelDate()` helper |
| Total Revenue showing PKR 0 | Commas in numbers (`"8,500"`) broke `Number()` | `parseNumber()` strips symbols |
| Column headers not mapping | Spaces/underscores mismatch (`Unit_Price` vs `Unit Price`) | `normalizeKey()` fuzzy match |
| Returns showing 0 | `true`/`1`/`"yes"` not recognized as "Yes" | `parseIsReturn()` multi-format |

---

## 🚀 How to Run

```bash
# 1. Clone the repo
git clone https://github.com/YOUR_USERNAME/maxbachat-warehouse-dashboard.git

# 2. Install dependencies
npm install

# 3. Run locally
npm start
```

---

## 📋 Excel Upload Format

Your `.xlsx` file can use any of these column name styles:

| Field | Accepted Column Names |
|---|---|
| Price | `UnitPrice_PKR`, `Unit Price`, `unit_price`, `Price` |
| Expiry | `ExpiryDate`, `Expiry Date`, `expiry_date`, `Expiry` |
| Returns | `IsReturn`, `Is Return`, `isreturn` — values: `Yes/No`, `true/false`, `1/0` |
| Cost | `Cost_PKR`, `Cost (PKR)`, `cost_pkr`, `Cost` |

---

## 👨‍💻 Developer

**Saddam Hussain**  
Full Stack Developer — Hyderabad, Sindh, Pakistan  
Built with ❤️ for Pakistan's wholesale business community

---

*© 2025 MaxBachat Wholesale — All Rights Reserved*
