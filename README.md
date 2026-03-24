# Flow & Pressure Analysis Dashboard

An Excel workbook for analysing flow and pressure sensor data.  
Pick any flow and any pressure from the selection lists, apply scaling/offset adjustments, and view the results in a live dual-axis chart and data table.

---

## Workbook Structure

| Sheet | Purpose |
|---|---|
| **Raw Flow Data** | Paste your wide-format flow data here (Date + flow columns) |
| **Raw Pressure Data** | Paste your wide-format pressure data here (Date + pressure columns) |
| **Dashboard** | Main working area — selection lists, controls, chart, data table |
| **MOD Flow** | Saved adjusted flow data (appended each time you click Save) |
| **MOD Pressure** | Saved adjusted pressure data (appended each time you click Save) |
| **Instructions** | Step-by-step guide, Power Query setup, and VBA Save button code |

---

## Data Format

Paste your data in this exact format into both Raw tabs:

```
Date              | AL012       | AL013       | AL014       | ...
12/01/2026 00:00  | 3.168205    | 2.204250    | 2.665153    | ...
12/01/2026 00:15  | 3.190769    | 2.225250    | 2.681334    | ...
```

- **Column A** — Date/Time values (not text)
- **Columns B onwards** — one column per flow or pressure sensor
- Names can be any combination of letters and numbers (`AL012`, `AM005`, `AF037`, etc.)
- Use **-999** for missing/no-data values — they are automatically excluded from all calculations
- Both tabs can have **different column names** (flows and pressures are selected independently)

---

## How to Use

### First-time setup
1. Open the workbook and go to **Raw Flow Data**
2. Delete the sample rows (keep Row 1 headers), then paste your flow data from Row 2
3. Do the same in **Raw Pressure Data**
4. Go to the **Dashboard** tab

### Daily use
1. **Select a flow** — use the `Select Flow ▼` dropdown (cell **B3**)  
   The chosen flow is highlighted in the list on the left side of the dashboard
2. **Select a pressure** — use the `Select Pressure ▼` dropdown (cell **E3**)  
   The chosen pressure is highlighted in the list on the right side
3. **Adjust values** if needed:
   - `Flow Scaling Factor` (cell **B2**) — multiplies the flow by this value (default 1.0)
   - `Pressure Offset` (cell **E2**) — adds this value to the pressure (default 0)
4. The **dual-axis chart** and **formula table** update instantly
5. When satisfied, run the **SaveToMOD** macro to append the adjusted data to the MOD tabs

### After pasting your own data
The dropdown in B3 and E3 reads directly from the Raw tab headers.  
If your new column range extends beyond the default, update the source:
- Right-click **B3** → Data Validation → change Source to e.g. `'Raw Flow Data'!$B$1:$BZ$1`
- Do the same for **E3** using `'Raw Pressure Data'`

---

## VBA Save Button

The full `SaveToMOD` macro code is in the **Instructions** sheet.

To add it:
1. Press **Alt+F11** to open the VBA editor
2. **Insert → Module** and paste the code from the Instructions sheet
3. **Developer tab → Insert → Button (Form Control)** — draw on the Dashboard and assign `SaveToMOD`

The macro appends rows to MOD Flow and MOD Pressure (history is never overwritten).

---

## Optional Enhancements

### Power Query (for large datasets)
The Raw tabs are already set up as named Excel Tables (`FlowData`, `PressureData`), making Power Query setup one-click:
- **Data → Get Data → From Table/Range** on either Raw tab
- Unpivot the flow/pressure columns from wide to long format
- Merge flow + pressure queries on Date + Name
- Apply scaling/offset as calculated columns
- Load to a sheet and build a **PivotTable + Slicer** for interactive multi-flow comparison

See the **Instructions** sheet for the full step-by-step walkthrough.

---

## Re-generating the Workbook

If you need to regenerate the file (e.g., after code changes):

```bash
pip install openpyxl
python3 generate_dashboard.py
```
