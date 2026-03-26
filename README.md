# Flow & Pressure Analysis Dashboard

An Excel workbook for analysing flow and pressure sensor data.  
Pick any flow and any pressure from the selection lists, apply scaling/offset adjustments,
and view the results in a live dual-axis chart and data table.

---

## Workbook Structure

| Sheet | Purpose |
|---|---|
| **Raw Flow Data** | Paste your wide-format flow data here (Date + flow columns) |
| **Raw Pressure Data** | Paste your wide-format pressure data here (Date + pressure columns) |
| **Point Index** | Reference table of sensor point codes, asset IDs, coordinates, etc. |
| **Dashboard** | Main working area — selection lists, controls, chart, formula table |
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
- Use **-999** for missing/no-data values — they are automatically excluded
- Both tabs can have **different column names** (flows and pressures are selected independently)

---

## How to Use

### First-time setup
1. Open the workbook and go to **Raw Flow Data**
2. Delete the sample rows (keep Row 1 headers), then paste your flow data from Row 2
3. Do the same in **Raw Pressure Data**
4. Go to the **Dashboard** tab

### Daily use
1. **Select flows** — rows 3–22 each have a `Flow N ▼` dropdown (col **B**).  
   Pick up to 20 flow meters; leave unused rows blank.
2. **Select pressures** — rows 3–22 each have a `Pres N ▼` dropdown (col **G**).  
   Pick up to 20 pressure points; leave unused rows blank.
3. **Adjust per-series values**:
   - `Scale` (col **C**, rows 3–22) — multiplies each flow by its own factor (default 1.000)
   - `Offset` (col **H**, rows 3–22) — adds a constant to each pressure (default 0.000)
   - `Δt` (col **D** / col **I**) — integer timestep shift to align sensors in time
4. The **dual-axis chart** (20 flow series on left axis; 20 pressure on right) and
   **formula table** (rows 26+, cols A–AO) update instantly
5. Click a **💾** cell (col **E** for flow, col **J** for pressure, rows 3–22) to apply the
   current Scale / Offset / Δt and write the adjusted values back into the Raw tab.  
   **Keep a backup of your original raw data before clicking Save.**

### After pasting your own data
Each Name dropdown reads directly from the Raw tab headers.  
If your column range extends beyond the default, update the source:
- Right-click any **Flow Name** cell (B3–B22) → Data Validation → change Source to  
  e.g. `'Raw Flow Data'!$B$1:$BZ$1`
- Do the same for any **Pressure Name** cell (G3–G22) using `'Raw Pressure Data'`

---

## VBA Save Macro

The full VBA code is in the **Instructions** sheet (section 4) and in two text files
generated alongside the workbook:

| File | Contents |
|---|---|
| `VBA_Module1_SaveSensor.txt` | `SaveOneSensor` subroutine — writes adjusted values back into the Raw tab |
| `VBA_Dashboard_Sheet.txt` | Sheet event handler — makes 💾 cells clickable |

| Macro | What it does |
|---|---|
| `SaveOneSensor(True, n)` | Flow row n — applies Scale × raw value, writes back to Raw Flow Data |
| `SaveOneSensor(False, n)` | Pressure row n — applies raw value + Offset, writes back to Raw Pressure Data |

To install:
1. Press **Alt+F11** to open the VBA editor
2. **Insert → Module** and paste the contents of `VBA_Module1_SaveSensor.txt`
3. In the Project tree, double-click **Sheet1 (Dashboard)** and paste `VBA_Dashboard_Sheet.txt`
4. Save the file as `.xlsm`

Once installed, clicking a **💾** cell writes that sensor's adjusted data into the Raw tab.

> ⚠️ **Important:** The 💾 buttons overwrite the original column in the Raw tab.
> Always keep a backup before saving.

---

## Re-generating the Workbook

To apply the latest dashboard changes to `Model Build Dashboard v1.21.xlsx`:

```bash
python3 generate_dashboard.py
```

This script removes legacy MOD sheets, clears the Save Rest button, and refreshes the
Instructions sheet with the current VBA — while leaving Raw Data, Point Index, Dashboard
selections, and the chart completely untouched.