# Manufacturing Performance & OEE Dashboard | Power BI Production Analytics
Power BI + DAX + Power Query + Excel Project | Production Run Data | From Raw Shop-Floor Records to OEE-Driven Business Intelligence

## Table Of Contents
* [Project Overview](#project-overview)
* [Tools & Technologies](#tools--technologies)
* [Dataset Overview](#dataset-overview)
* [Data Preparation](#data-preparation)
* [Data Model](#data-model)
* [DAX Measures](#dax-measures)
* [Power BI Dashboard](#power-bi-dashboard)
* [Key Metrics](#key-metrics)
* [Key Insights](#key-insights)
* [Executive Summary](#executive-summary)
* [Recommendations](#recommendations)
* [Assumptions & Limitations](#assumptions--limitations)

---

### Project Overview
This project delivers a 4-page interactive Power BI dashboard that tracks manufacturing performance across 6 machines, 3 shifts and 4 products from January to June 2026. Built around **Overall Equipment Effectiveness (OEE)** — Availability × Performance × Quality — the dashboard shows management where production capacity is lost, which machines and shifts are underperforming, which products drive rejections, and what causes downtime.

Over the six months, the plant produced **2,203,570 units against a target of 2,364,916 (93.2% achievement)**, with an overall **OEE of 77.3% — 7.7 points below the 85% world-class benchmark**. The analysis traces most of that gap to a single machine (Machine-04), breakdown-driven downtime, Night-shift speed losses, and a product-specific quality issue with Product C.

---

### Tools & Technologies
* **Microsoft Excel** → Source data (`Production_Data_Cleaned.xlsx`)
* **Power Query** → Data loading, type checks and transformation
* **Power BI Desktop** → 4-page interactive dashboard with synced slicers and page navigation
* **DAX** → 25+ measures covering OEE components, targets, rejections, downtime and rankings
* **Data Modelling** → Star schema with a dedicated measures table and date dimension

---

### Dataset Overview

**Production Fact Table (`Fact_Production`)** — one row per production run

| Column | Description |
|--------|-------------|
| Date | Production date (01/01/2026 – 30/06/2026) |
| Machine_ID | Machine identifier (Machine-01 to Machine-06) |
| Shift | Operating shift (Morning, Evening, Night) |
| Product | Product produced on the run (Product A, B, C, D) |
| Target Qty | Planned output for the run |
| Produced Qty | Actual units produced |
| Good Qty | Units that passed quality checks |
| Rejected Qty | Units rejected at quality check |
| Downtime (Min) | Minutes of unplanned/planned stoppage during the run |
| DowntimeReason | Breakdown, Changeover, Material Shortage, No Operator, Power Outage, Preventive Maintenance, Unknown / Not Recorded |

**Dimension Tables**

| Table | Description |
|-------|-------------|
| Dim_Date | Continuous calendar table (Year, Quarter, Month, Day) for time intelligence |
| Dim_Machine | Machine master list used for machine-level slicing and ranking |
| _Measures | Dedicated table holding all DAX measures |

---

### Data Preparation
- Loaded `Production_Data_Cleaned.xlsx` into Power Query and validated data types (dates, whole numbers, decimals)
- Confirmed that `Good Qty + Rejected Qty = Produced Qty` on every run (2,129,931 + 73,639 = 2,203,570)
- Standardised downtime reason labels and grouped missing reasons as **"Unknown / Not Recorded"** rather than dropping rows
- Converted downtime minutes to hours for reporting; kept minutes for per-run averages
- Built a continuous `Dim_Date` table and marked it as the model's date table
- Flagged runs with downtime longer than 2 hours as **Long Stops** for separate analysis

---

### Data Model
The model follows a star schema built around a single fact table:

```
             Dim_Date                     Dim_Machine
           (Date, Month,                  (Machine_ID)
            Quarter, Year)                     │
                 │ 1                           │ 1
                 │                             │
                 └────────► Fact_Production ◄──┘
                        *  (Shift, Product,    *
                            DowntimeReason,
                            quantities, downtime)

                     _Measures (DAX measures only)
```

Shift, Product and DowntimeReason are low-cardinality attributes stored directly on `Fact_Production` (degenerate dimensions), so their slicers and visuals reference the fact table columns. Relationships are single-direction, one-to-many from dimensions to the fact table.

---

### DAX Measures

- Core Volume Measures
```DAX
Total Target = SUM ( Fact_Production[Target Qty] )

Total Production = SUM ( Fact_Production[Produced Qty] )

Good Qty = SUM ( Fact_Production[Good Qty] )

Rejected Qty = SUM ( Fact_Production[Rejected Qty] )

Achievement % = DIVIDE ( [Total Production], [Total Target] )

Rejection % = DIVIDE ( [Rejected Qty], [Total Production] )
```

- Downtime Measures
```DAX
Downtime (Hrs) = DIVIDE ( SUM ( Fact_Production[Downtime (Min)] ), 60 )

Avg Downtime per Run (Min) =
    AVERAGE ( Fact_Production[Downtime (Min)] )

Long Stops (>2 hrs) =
    CALCULATE (
        COUNTROWS ( Fact_Production ),
        Fact_Production[Downtime (Min)] > 120
    )
```

- OEE Components
```DAX
Availability % = 1 - [Downtime %]

Performance % =
    DIVIDE ( [Total Production], [Total Target] * [Availability %] )

Quality % = DIVIDE ( [Good Qty], [Total Production] )

OEE % = [Availability %] * [Performance %] * [Quality %]
```

- Ranking & Callout Measures
```DAX
Machine OEE Rank =
    RANKX ( ALL ( Dim_Machine[Machine_ID] ), [OEE %], , DESC, DENSE )

Top Machine (OEE) =
    FIRSTNONBLANK (
        TOPN ( 1, VALUES ( Dim_Machine[Machine_ID] ), [OEE %], DESC ),
        1
    )

Runs Meeting Target % =
    DIVIDE (
        CALCULATE (
            COUNTROWS ( Fact_Production ),
            FILTER ( Fact_Production,
                Fact_Production[Produced Qty] >= Fact_Production[Target Qty] )
        ),
        COUNTROWS ( Fact_Production )
    )
```

- Pareto Measure
```DAX
Cumulative Downtime % =
VAR CurrentHrs = [Downtime (Hrs)]
VAR Cumulative =
    SUMX (
        FILTER (
            ALL ( Fact_Production[DowntimeReason] ),
            [Downtime (Hrs)] >= CurrentHrs
        ),
        [Downtime (Hrs)]
    )
RETURN
    DIVIDE ( Cumulative, CALCULATE ( [Downtime (Hrs)], ALL ( Fact_Production[DowntimeReason] ) ) )
```

---

### Power BI Dashboard
The dashboard has 4 pages with synced slicers (Date, Shift, Machine, Product, Downtime Reason) and a side navigation panel:

* **Management Overview** — Headline KPIs, OEE gauge vs 85% benchmark, OEE loss waterfall, monthly target vs actual, OEE components trend
* **Machine Analysis** — Machine scorecard, OEE by machine, top/bottom 3 machines, target vs actual by machine
* **Shift & Product Analysis** — Shift efficiency, target vs actual by shift and product, rejection % by product, machine × product rejection heatmap
* **Downtime Analysis** — Downtime Pareto by reason, downtime by machine and reason, long stops, downtime by shift, monthly downtime % by machine

---

#### Dashboard Screenshots

**Management Overview**

<img width="777" height="406" alt="Screenshot 2026-09-29 105208" src="https://github.com/user-attachments/assets/be630233-9d39-42ed-967c-ae2ab0e2b47b" />


**Machine Analysis**

<img width="775" height="407" alt="Screenshot 2026-09-29 105349" src="https://github.com/user-attachments/assets/019d3f67-e442-4a65-8f08-60fb453d9dc5" />


**Shift & Product Analysis**

<img width="776" height="404" alt="Screenshot 2026-09-29 105430" src="https://github.com/user-attachments/assets/4c046d26-cb34-44a8-b0df-694c6bc66e7d" />


**Downtime Analysis**

<img width="778" height="406" alt="Screenshot 2026-09-29 105454" src="https://github.com/user-attachments/assets/309cded7-6695-4f30-8680-d6cf97cc3655" />


---

### Key Metrics
- **Reporting Period:** 01 Jan 2026 – 30 Jun 2026
- **Machines Monitored:** 6
- **Total Target:** 2,364,916 units
- **Total Production:** 2,203,570 units
- **Achievement %:** 93.2%
- **OEE %:** 77.3% (benchmark 85%)
- **Availability %:** 89.5%
- **Performance %:** 89.3%
- **Quality %:** 96.7%
- **Rejection %:** 3.3% (73,639 units rejected)
- **Total Downtime:** 2,481.3 hours (10.5% of planned time)
- **Avg Downtime per Run:** 46.14 minutes
- **Long Stops (>2 hrs):** 214
- **Runs Meeting Target:** 41.9%

---

### Key Insights
1. **OEE is 7.7 points below world-class.** Overall OEE of 77.3% sits below the 85% benchmark, although it improved steadily every month from 75.2% in January to 79.1% in June, alongside achievement rising from 90.5% to 95.5%.
2. **Availability is the biggest source of lost output.** The OEE loss waterfall shows 2.78M units of rated capacity reduced to 2.13M good units — availability loss (0.32M) is the largest single loss, ahead of speed loss (0.26M) and quality loss (0.07M), making downtime the first lever to pull.
3. **Machine-04 is the plant's main bottleneck.** It has the lowest OEE (60.4%), lowest achievement (73.5%) and the most downtime (869.8 hrs — about 35% of the plant total). It also accounts for 156 of the 214 long stops and runs at roughly 22% downtime every month, well above every other machine.
4. **Only one machine beats the benchmark.** Machine-06 is the top performer at 88.3% OEE and 107.5% achievement, with the least downtime (193.6 hrs). Machine-01 (84.3%) is just below the benchmark; the remaining four fall short.
5. **Breakdowns drive downtime.** Breakdown is the top downtime reason by a wide margin, followed by changeover and material shortage — the Pareto shows these three reasons account for the large majority of lost hours.
6. **Night shift loses output through speed, not stoppages.** Night produced the least (705,318 units) and has the lowest OEE, yet downtime is almost identical across shifts (Morning 821, Evening 853, Night 807 hrs). The gap comes from lower performance (running speed) on Nights.
7. **Product C's quality issue is product-driven, not machine-driven.** Product C has a 5.0% rejection rate — roughly double Product D (2.4%) — and accounts for 28,176 rejected units (about 38% of all rejections). Its rejection rate is 5.0–5.1% on every machine, pointing to the product specification or process rather than any single machine.
8. **Most runs miss their target.** Only 41.9% of individual runs met target, even though overall achievement is 93.2% — overproduction on strong machines (Machine-06, Machine-01) is masking frequent shortfalls elsewhere.

---

### Executive Summary
Between January and June 2026, the plant produced 2,203,570 units against a target of 2,364,916, achieving 93.2% of plan with an OEE of 77.3% — 7.7 points below the 85% world-class benchmark. Quality is strong at 96.7%, so the gap is driven mainly by availability (2,481 hours of downtime, 10.5% of planned time) and speed losses. Machine-04 alone accounts for about 35% of all downtime and 156 of 214 long stops, operating at 60.4% OEE, while Machine-06 shows what is achievable at 88.3%. Breakdowns are the leading cause of downtime, Night shift underperforms on running speed, and Product C's 5.0% rejection rate is consistent across all machines, indicating a product or process issue. OEE has improved every month, and targeting Machine-04 reliability alone would close a significant share of the remaining gap to benchmark.

---

### Recommendations
* **Prioritise a reliability review of Machine-04** — root-cause its breakdowns and long stops; bringing it to the plant average would recover the largest single block of lost capacity.
* **Shift breakdowns into planned maintenance** — use the downtime Pareto to build a preventive maintenance schedule targeting the machines with the most breakdown hours.
* **Reduce changeover time** — as the second-largest downtime reason, apply SMED (quick changeover) techniques and sequence products to minimise changeovers.
* **Investigate Night shift running speed** — check staffing levels, supervision and machine settings on Nights, since stoppages are not the cause of the gap.
* **Review Product C's specification and process** — the identical rejection rate on every machine points to materials, tooling standards or process parameters rather than equipment.
* **Improve downtime reason capture** — reduce "Unknown / Not Recorded" entries so future analysis can attribute every lost hour.
* **Replicate Machine-06 practices** — document its settings, maintenance routine and operator practices as a benchmark for the rest of the line.

---

### Assumptions & Limitations
* OEE is calculated as Availability × Performance × Quality, with target quantity used as the ideal output baseline for Performance.
* Downtime includes all recorded stoppage reasons, including preventive maintenance and changeovers.
* Shift, Product and DowntimeReason are modelled as attributes on the fact table rather than separate dimension tables.
* The analysis covers six months only; seasonal effects and longer-term trends cannot be assessed.

---

### Project Files

| File | Description |
|------|-------------|
| `Production_Data_Cleaned.xlsx` | [Source dataset](https://docs.google.com/spreadsheets/d/1nTLpkq1oDrgL4M7WBA6DUl_c9M5_XXA2/edit?usp=sharing&ouid=111035109327406077555&rtpof=true&sd=true) |



---

*By Tunde Adebayo | Data Analyst*
