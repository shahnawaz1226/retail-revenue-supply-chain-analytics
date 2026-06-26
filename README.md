# Omnichannel Retail Revenue Optimization & Supply Chain Matrix

> **Enterprise Portfolio Project** · Excel Analytics · FY 2024–2025  
> *Diagnostic & Prescriptive Analytics for Compressed EBITDA Margins*

---

## 1. Executive Summary

A multichannel retail corporation is experiencing **compressed net profit margins** despite a year-over-year increase in gross sales volume. This project builds a full end-to-end analytics solution in Excel — from raw messy data ingestion through to executive-ready prescriptive recommendations — to diagnose and reverse the margin erosion.

### Final Business Results

| Business Outcome | Impact |
|---|---|
| Root Cause Identified | Aggressive discounting (>25%) on low-margin SKUs eroding ~12% of net margin |
| Fulfillment Insight | East/Central regions averaging 5.5+ days lead time → 55% higher return rate |
| Prescriptive Action | Discount capping + regional inventory redistribution → projected 14–18% return reduction |
| Automation Savings | Pipeline replaces ~8 hours/week of manual reporting work |

---

## 2. Business Problem Statement

Three operational bottlenecks drive the margin compression:

1. **Unoptimised Discounting** — Blanket promotional tiers applied across all product lines, including low-margin categories (Grocery: 12–22% GM, Electronics: 22–35% GM), erode EBITDA without proportional volume uplift.
2. **Fulfillment Latency** — Delivery delays >5 days in East and Central regions correlate with a 55% increase in product return velocity, compounding logistics cost.
3. **Return Rate Concentration** — Electronics (8.2%) and Apparel (7.1%) are primary return-risk categories, driven by mis-described listings and size/spec mismatches.

**Target Stakeholders:** CFO · VP Supply Chain · Regional Operations Directors

---

## 3. Data Architecture — Star Schema

```
                  ┌─────────────────┐
                  │  Fact_Orders    │  (36,000 Transactions)
                  └────────┬────────┘
                           │
         ┌─────────────────┼──────────────────┐
         ▼                 ▼                  ▼
  ┌─────────────┐  ┌─────────────┐  ┌──────────────────┐
  │ Dim_Products│  │  Dim_Stores │  │  Fact_Returns    │
  │ (200 SKUs)  │  │ (50 Stores) │  │ (2,221 Events)   │
  └─────────────┘  └─────────────┘  └──────────────────┘
                           │
                  ┌────────▼────────┐
                  │  Dim_Calendar   │
                  │ (730 Days)      │
                  └─────────────────┘
```

| Table | Rows | Primary Key | Description |
|---|---|---|---|
| `Fact_Orders` | 36,000 | `Order_ID` | Core transactional log — revenue, COGS, margin, lead times |
| `Dim_Products` | 200 | `Product_ID` | SKU hierarchy, category, brand, margin tier classification |
| `Dim_Stores` | 50 | `Store_ID` | Regional and channel dimension |
| `Fact_Returns` | 2,221 | `Return_ID` | Returns log linked to Fact_Orders via Order_ID |
| `Dim_Calendar` | 730 | `Date` | Corporate fiscal calendar — quarters, weeks, business days |

---

## 4. Workbook Structure (13 Sheets)

| Sheet | Tab Colour | Purpose |
|---|---|---|
| `📋 README` | Blue | Project overview, schema map, Power Query setup guide |
| `RAW_Orders` | Amber | Unprocessed transactions: mixed date formats, text discounts, null regions |
| `RAW_Products` | Amber | Raw product dimension with inconsistent SKU case/spacing |
| `RAW_Stores` | Amber | Store dimension with NULL zone codes (~6%) |
| `RAW_Returns` | Amber | Returns log before foreign key validation |
| `Fact_Orders` | Green | Cleaned, validated 36,000-row orders table |
| `Dim_Products` | Green | Parsed SKU hierarchy with margin tier classification |
| `Dim_Stores` | Green | Standardised store dimension |
| `Fact_Returns` | Green | Validated returns linked to fact table |
| `Dim_Calendar` | Green | Fiscal calendar with period, quarter, week attributes |
| `🔧 Calc_Engine` | Teal | Diagnostic engine — 6 analysis sections, 362 formulas |
| `✅ Data_Audit` | Green | Cross-validation layer — formula results vs expected source totals |
| `📊 Dashboard` | Blue | Executive interface — KPI blocks + 4 embedded charts |

---

## 5. Data Ingestion & Transformation Pipeline (Power Query ETL)

### Setup Instructions

1. Open Excel → **Data** tab → **Get Data → From File → From Workbook**
2. Select each `RAW_*` sheet individually to load into the Power Query Editor
3. Apply the following transformations per table:

**RAW_Orders cleanup steps:**
```
→ Text.Trim on all string columns
→ Date.FromText on Order_Date_Raw (handles dd/mm/yyyy, yyyy-mm-dd, "Jan 15, 2024")
→ Type.Replace on Discount_Pct: remove "%" suffix → divide by 100 → set as Decimal
→ Table.FillDown on Region column (null imputation by last known value)
```

**RAW_Products cleanup steps:**
```
→ Text.Trim + Text.Upper on SKU_Raw → normalised SKU_Clean
→ Text.Split on SKU_Clean using "-" delimiter → Channel | Category | Item_ID
→ Set Unit_Cost and MSRP as Currency type
```

4. Load all cleaned tables to the Excel **Data Model** via *Close & Load To → Only Create Connection → Add to Data Model*
5. Create table relationships in **Power Pivot → Diagram View**:
   - `Fact_Orders[Product_ID]` → `Dim_Products[Product_ID]`
   - `Fact_Orders[Store_ID]` → `Dim_Stores[Store_ID]`
6. Connect slicers on the Dashboard to Pivot Tables built from the Data Model

---

## 6. Analytical Engineering — Key Formula Patterns

### SUMIFS Aggregation (Monthly Revenue Engine)
```excel
=IFERROR(SUMIFS(Fact_Orders!$U:$U,
                Fact_Orders!$L:$L, [Year],
                Fact_Orders!$M:$M, [Month]), 0)
```

### LET Block — GMROI Classification
```excel
=LET(
  GrossProfit,    Fact_Orders!W2,
  AvgInventory,   Fact_Orders!V2,
  GMROI,          IFERROR(GrossProfit / AvgInventory, 0),
  IF(GMROI >= 1.8, "Optimal",
    IF(GMROI >= 1.0, "Acceptable", "Poor Capital Efficiency"))
)
```

### Four-Tier Margin Health Flag (IFERROR + Nested IF)
```excel
=IFERROR(
  IF((W2/U2) < 0.15, "🔴 Critical Margin Risk",
    IF((W2/U2) < 0.30, "⚠️ Below Target",
      IF((W2/U2) < 0.42, "👁 Monitor", "✅ Optimal"))),
  "No Data")
```

### Dynamic Low-Margin SKU Ledger (FILTER + SORT)
```excel
=SORT(
  FILTER(Dim_Products!A:C, Dim_Products!H:H < 35, "No low-margin SKUs"),
  3, 1)
```

### XLOOKUP Cross-Table Dimension Join
```excel
=XLOOKUP(B2, Dim_Products!$A:$A, Dim_Products!$C:$C, "Unknown", 0)
```

---

## 7. Diagnostic Findings & Prescriptive Recommendations

### Finding 1 — Discount Tier Margin Erosion 🔴
**Data evidence:** Orders with Discount_Pct > 25% represent 15.4% of transaction volume but generated disproportionate margin compression. Deep-cut promotions (>30%) in Electronics and Apparel drove gross margin below 20% on affected orders.

**Prescription:** Enforce automatic 20% discount ceiling for Electronics and Apparel via system-level promotion rules. Reserve deep promotional tiers for Beauty and Toys (52–65% baseline GM provides sufficient buffer).

---

### Finding 2 — East & Central Region Fulfillment Latency 🔴
**Data evidence:** East region average lead time: 5.8 days. Central region: 5.1 days. Both exceed the 4.2-day SLA. Orders with lead time > 6 days show a **55% higher return rate** — the strongest single predictor of returns in the dataset.

**Prescription:** Establish micro-fulfilment hubs in Kolkata (East) and Bhopal (Central) stocked with top-50 velocity SKUs. Model projects lead time reduction to 3.2–3.8 days, reducing regional return rates by an estimated 14–18%.

---

### Finding 3 — Grocery Portfolio Margin Dilution ⚠️
**Data evidence:** Grocery operates at 12–22% gross margin — 20+ percentage points below the 42% corporate target — and represents a significant share of transaction volume. This category consistently drags the blended portfolio margin below target.

**Prescription:** Strategic portfolio review for Grocery SKU mix. Introduce bundled cross-sell promotions pairing Grocery staples with high-margin Beauty and Home products to increase blended transaction value per order.

---

## 8. Workbook Governance

- All underlying calculation sheets (`Calc_Engine`, `Data_Audit`) can be protected via **Review → Protect Sheet** with password, leaving Dashboard slicers and data validation drop-downs interactive
- Named Ranges store global business thresholds (target margin: 42%, SLA days: 4.2, max return rate: 4.5%)
- `Data_Audit` sheet cross-validates all calculated totals against raw source row counts — any variance above ₹0.01 triggers a `🔴 FAIL` alert

---

## 9. Repository Structure

```
├── README.md
├── assets/
│   ├── dashboard_main.png          # Executive dashboard screenshot
│   ├── calc_engine_preview.png     # Formula engine overview
│   └── schema_diagram.png          # Star schema data model
├── data/
│   └── data_source_manifest.md     # Dataset provenance notes
└── production/
    └── OMNICHANNEL_RETAIL_OPTIMIZATION_2026.xlsx
```

---

## 10. Excel Skills Demonstrated

| Skill Area | Specific Features Used |
|---|---|
| Power Query ETL | `Text.Trim`, `Text.Clean`, delimiter splits, date format coercion, null imputation |
| Data Modelling | Star Schema, Fact/Dimension relationships, Power Pivot Data Model |
| Dynamic Arrays | `FILTER`, `SORT`, `UNIQUE`, `SEQUENCE` — automated low-margin SKU ledger |
| Business Logic | `LET` variable blocks, `IFERROR`, four-tier nested `IF` classification |
| Cross-Table Joins | `XLOOKUP` replacing legacy `VLOOKUP` chains across normalised tables |
| Period Aggregation | `SUMIFS`, `COUNTIFS`, `AVERAGEIFS` over Year/Month sliding windows |
| Conditional Formatting | Data bars, colour scales, icon sets for performance alerting |
| Dashboard UX | Gridline-off layout, global slicer panels, combo charts, KPI metric blocks |
| Governance | Sheet protection, named ranges for business thresholds, print layout isolation |
| Documentation | Structured GitHub repo, executive README, inline formula commentary |

---

## 11. Data Source

Synthetic retail operations dataset modelled on the **Global Superstore Enterprise Schema** structure. Generated programmatically to simulate realistic enterprise data quality issues (mixed date formats, null regional keys, text-encoded numeric fields) for authentic ETL pipeline demonstration.

- **Volume:** 36,000 orders · 2,221 return events · 200 SKUs · 50 stores
- **Period:** 01 January 2024 – 31 December 2025
- **Geographies:** 5 regions · 25 cities across India
- **Channels:** Online · Offline · Wholesale

---

*Built with openpyxl · pandas · numpy · Python 3.12*  
*Portfolio project — Retail Analytics, June 2026*
