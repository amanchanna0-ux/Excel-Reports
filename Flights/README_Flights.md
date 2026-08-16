# ✈️ Flights Operations Analysis — Excel Report

Analysis of airline operations data for 280 flights across Pakistani and international routes. Built as a self-directed practice project to apply Excel reporting techniques on a real-world business scenario.

---

## 📋 Dataset Overview

| Field | Description |
|---|---|
| FlightID | Unique flight identifier |
| Airline | PIA, AirBlue, SereneAir, Fly Jinnah |
| Origin / Destination | Departure and arrival cities |
| ScheduledDeparture | Scheduled date and time |
| DelayMinutes | Minutes delayed (0 = on time) |
| AircraftType | Boeing 737, Airbus A320, ATR 72, Airbus A321 |
| SeatsAvailable / SeatsSold | Capacity and occupancy |
| Class | Economy or Business |
| AvgTicketPrice | Average ticket price in PKR |
| DelayReason | Weather, Technical, ATC, Crew, or None |

**Rows:** 280 flights &nbsp;|&nbsp; **Period:** Jul 2025 – Jun 2026 &nbsp;|&nbsp; **Routes:** 8 (domestic + international)

---

## 📊 Reports Built

### Report 1 — On-Time Performance by Airline
Counts total flights and delayed flights (>15 min) per airline, calculates on-time percentage.

**Technique:** Helper column (IsDelayed), COUNTIFS, percentage formula

---

### Report 2 — Route Load Factor
Average seat occupancy (SeatsSold / SeatsAvailable) per route using a helper Route column.

**Technique:** Helper column (Route = Origin & "-" & Destination), AVERAGEIFS

---

### Report 3 — Delay Reason Breakdown
Count of flights by delay reason, excluding "None", visualised as a pie chart.

**Technique:** COUNTIFS, pie chart

---

### Report 4 — Revenue Estimation by Airline & Class
PivotTable summarising total estimated revenue (SeatsSold × AvgTicketPrice) by Airline and Class, with an Origin slicer.

**Technique:** Helper column (Revenue), PivotTable, Slicer

---

### Report 5 — Peak Hour Departure Analysis
Flight volume by hour of day to identify peak and off-peak departure times.

**Technique:** HOUR() helper column, COUNTIFS, column chart

---

### Report 6 — Dynamic Route Lookup Panel
Enter any Origin and Destination to instantly see number of flights, average delay, and average load factor for that route.

**Technique:** COUNTIFS, AVERAGEIFS with two criteria (Origin + Destination input cells)

---

### Report 7 — Aircraft Utilization Heatmap
Matrix of AircraftType × Airline showing flight count per combination, with a color scale to highlight utilization patterns.

**Technique:** COUNTIFS matrix, 3-color scale conditional formatting

---

### Report 8 — Delay Threshold What-If
One input cell (Delay Threshold in minutes) that dynamically recalculates how many flights exceed it and what percentage of total flights that represents.

**Technique:** COUNTIFS with `">"&cell` criteria, COUNTA for total

---

## 🛠️ Tools & Functions Used

`COUNTIFS` `AVERAGEIFS` `MAXIFS` `HOUR()` `XLOOKUP` `IFERROR` `SUMPRODUCT`
PivotTables · Slicers · Conditional Formatting (Color Scale) · Excel Tables

---

## 💡 Key Findings

- **SereneAir** had the highest on-time rate among the four airlines
- **Karachi–Dubai** route showed the highest average load factor
- **Technical issues** were the most common cause of delays after weather
- Digital channel (Mobile + Online) bookings accounted for the majority of Business class revenue
- Peak departure hour was **8 AM**, with a secondary peak around **6 PM**

---

## 📁 File

| File | Description |
|---|---|
| `Flights_Operations_Analysis.xlsx` | Dataset + Brief + 8 Report sheets |

