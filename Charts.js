/**
 * Reads, filters, sorts, and prepares data for charting
 * with calculated error bars, outputting to a new sheet.
 */
function prepareChartData() {
  // ============================================================
  // 1. CONFIGURATION VARIABLES
  // ============================================================
  const X_COL_NAME = "BMI (kg/m²)"; // Waist (cm) or Waist circumference (cm)
  const Y_COL_NAME = "Systolic blood pressure (mmHg)";
  const ERROR_MULTIPLIER = 0.05; // 5% flat error

  // ============================================================
  // 2. DATA EXTRACTION
  // ============================================================
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sourceSheet = ss.getActiveSheet();
  const sourceSheetName = sourceSheet.getName();
  const allData = sourceSheet.getDataRange().getValues();

  // Locate column indices from headers (row 0)
  const headers = allData[0];
  const xIdx = headers.indexOf(X_COL_NAME);
  const yIdx = headers.indexOf(Y_COL_NAME);

  if (xIdx === -1 || yIdx === -1) {
    throw new Error(
      `Column not found. Looking for "${X_COL_NAME}" (${xIdx === -1 ? "MISSING" : "ok"}) ` +
      `and "${Y_COL_NAME}" (${yIdx === -1 ? "MISSING" : "ok"}). ` +
      `Available headers: ${headers.join(", ")}`
    );
  }

  // Regex to catch "N/A", "n/a", "n.a.", "N.A.", "na", etc.
  const naPattern = /^n[\.\s\/]*a[\.\s]*$/i;

  // Extract rows, filtering out any N/A variations in either target column
  const filteredRows = [];
  for (let i = 1; i < allData.length; i++) {
    const xRaw = allData[i][xIdx];
    const yRaw = allData[i][yIdx];

    // Skip if either value is blank or matches an N/A pattern
    if (xRaw === "" || yRaw === "" || xRaw == null || yRaw == null) continue;
    if (naPattern.test(String(xRaw).trim())) continue;
    if (naPattern.test(String(yRaw).trim())) continue;

    const xVal = Number(xRaw);
    const yVal = Number(yRaw);

    // Skip if conversion to number fails
    if (isNaN(xVal) || isNaN(yVal)) continue;

    filteredRows.push({ x: xVal, y: yVal });
  }

  if (filteredRows.length === 0) {
    throw new Error("No valid data rows remain after filtering out N/A and non-numeric values.");
  }

  // ============================================================
  // 3. SORT by X-axis ascending
  // ============================================================
  filteredRows.sort((a, b) => a.x - b.x);

  // ============================================================
  // 4. ERROR CALCULATION (flat percentage of each Y value)
  // ============================================================
  const outputRows = filteredRows.map(row => [
    row.x,
    row.y,
    Math.round(Math.abs(row.y * ERROR_MULTIPLIER) * 1000) / 1000 // ±error, rounded to 3 dp
  ]);

  // ============================================================
  // 5. SHEET MANAGEMENT — create or replace the output sheet
  // ============================================================
  let newSheetName = `${sourceSheetName} - ${X_COL_NAME} and ${Y_COL_NAME}`;
  if (newSheetName.length > 100) {
    newSheetName = newSheetName.substring(0, 99);
  }

  const existing = ss.getSheetByName(newSheetName);
  if (existing) {
    ss.deleteSheet(existing);
  }

  const outSheet = ss.insertSheet(newSheetName);

  // ============================================================
  // 6. DATA OUTPUT — write headers + data as clean columns
  // ============================================================
  const outHeaders = [[X_COL_NAME, Y_COL_NAME, `${Y_COL_NAME} Error (±${ERROR_MULTIPLIER * 100}%)`]];
  outSheet.getRange(1, 1, 1, 3).setValues(outHeaders);
  outSheet.getRange(2, 1, outputRows.length, 3).setValues(outputRows);

  // Light formatting: bold headers, auto-resize columns
  outSheet.getRange(1, 1, 1, 3).setFontWeight("bold");
  outSheet.autoResizeColumns(1, 3);

  // Activate the new sheet so the user sees it immediately
  ss.setActiveSheet(outSheet);
  // createScatterChart(X_COL_NAME,Y_COL_NAME,ERROR_MULTIPLIER);
  createBinnedMeanChart(X_COL_NAME,Y_COL_NAME)
  SpreadsheetApp.getUi().alert(
    `Done — ${outputRows.length} rows written to "${newSheetName}".`
  );
}


/**
 * Reads the prepared data sheet and inserts a formatted
 * scatter plot with error bars.
 */
function createScatterChart(X_COL_NAME,Y_COL_NAME,ERROR_MULTIPLIER) {
  // ============================================================
  // CONFIGURATION — must match the values in prepareChartData()
  // ============================================================
  // const X_COL_NAME = "HbA1C (%)";
  // const Y_COL_NAME = "Waist (cm)";
  // const ERROR_MULTIPLIER = 0.05;

  // ============================================================
  // 1. DATA RETRIEVAL — locate the prepared sheet
  // ============================================================
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sourceSheetName = ss.getActiveSheet().getName();

  // Build the expected sheet name; if called right after
  // prepareChartData(), the active sheet IS the prepared sheet,
  // so we also accept that directly.
  let targetName = `${sourceSheetName} - ${X_COL_NAME} and ${Y_COL_NAME}`;
  let chartSheet = ss.getSheetByName(targetName);

  // If not found, the active sheet itself may already be the
  // prepared sheet (e.g., user ran this function independently).
  if (!chartSheet) {
    const active = ss.getActiveSheet();
    if (active.getName().indexOf(X_COL_NAME) !== -1 &&
        active.getName().indexOf(Y_COL_NAME) !== -1) {
      chartSheet = active;
    } else {
      throw new Error(
        `Prepared sheet "${targetName}" not found. Run prepareChartData() first.`
      );
    }
  }

  // Read all data (row 1 = headers, rows 2+ = data)
  const allData = chartSheet.getDataRange().getValues();
  const numRows = allData.length; // includes header

  if (numRows < 2) {
    throw new Error("The prepared sheet has no data rows.");
  }

  // ============================================================
  // 2. REMOVE ANY EXISTING CHARTS on this sheet (avoid stacking)
  // ============================================================
  const existingCharts = chartSheet.getCharts();
  for (let i = 0; i < existingCharts.length; i++) {
    chartSheet.removeChart(existingCharts[i]);
  }

  // ============================================================
  // 3. DEFINE DATA RANGES
  //    Col A = X values, Col B = Y values, Col C = Error values
  //    Each range INCLUDES the header row so the chart can pick
  //    up series labels automatically.
  // ============================================================
  const xRange = chartSheet.getRange(1, 1, numRows, 1); // Column A
  const yRange = chartSheet.getRange(1, 2, numRows, 1); // Column B
  const errRange = chartSheet.getRange(1, 3, numRows, 1); // Column C

  // ============================================================
  // 4. BUILD THE SCATTER CHART using EmbeddedChartBuilder
  //    -------------------------------------------------------
  //    NOTE on correct API usage:
  //    - Use .addRange() to attach data ranges (NOT setXAxisColumn).
  //    - The FIRST addRange() call supplies the X-axis domain.
  //    - Subsequent addRange() calls add Y-axis series.
  //    - Use .setOption() for all advanced formatting.
  //    This avoids the "setXAxisColumn is not a function" error.
  // ============================================================
  const chartBuilder = chartSheet.newChart()
    .setChartType(Charts.ChartType.SCATTER)

    // Data ranges: first = domain (X), second = series (Y)
    .addRange(xRange)
    .addRange(yRange)

    // Position the chart to the right of the data (row 1, col E)
    .setPosition(1, 5, 0, 0)

    // ---- General options ----
    .setOption("title", `${Y_COL_NAME} vs ${X_COL_NAME}`)
    .setOption("width", 700)
    .setOption("height", 450)

    // ---- X-axis ----
    .setOption("hAxis.title", X_COL_NAME)
    .setOption("hAxis.titleTextStyle", { italic: false, bold: true, fontSize: 12 })
    .setOption("hAxis.gridlines", { color: "#e0e0e0" })

    // ---- Y-axis ----
    .setOption("vAxis.title", Y_COL_NAME)
    .setOption("vAxis.titleTextStyle", { italic: false, bold: true, fontSize: 12 })
    .setOption("vAxis.gridlines", { color: "#e0e0e0" })

    // ---- Data point styling ----
    .setOption("pointSize", 5)
    .setOption("colors", ["#2980b9"])

    // ---- Legend ----
    .setOption("legend", { position: "top", alignment: "center" })

    // ---- Trendline (polynomial to reveal non-linear paradox) ----
    .setOption("trendlines", {
      0: {
        type: "polynomial",
        degree: 2,
        color: "#e74c3c",
        lineWidth: 2,
        opacity: 0.7,
        showR2: true,
        visibleInLegend: true,
        labelInLegend: "Trend (poly-2)"
      }
    })

    // ---- Error bars from column C ----
    //   series 0 = the Y series; errorBars.value points to
    //   the Error column (C) so Sheets draws ± bars.
    .setOption("series", {
      0: {
        errorBars: {
          errorType: "custom",
          magnitude: ERROR_MULTIPLIER * 100 // percentage magnitude
        }
      }
    });

  // ============================================================
  // 5. INSERT THE CHART
  // ============================================================
  const chart = chartBuilder.build();
  chartSheet.insertChart(chart);
}

/**
 * Creates a binned-mean scatter chart with error bars.
 * Call this after prepareChartData() or independently.
 *
 * SETUP: Paste into the same script file. Update the three
 * config constants to match your column headers.
 */
function createBinnedMeanChart(X_COL_NAME,Y_COL_NAME) {
  // ============================================================
  // CONFIGURATION
  // ============================================================
  const BIN_WIDTH    = 1;                              // each bin spans 3 BMI units
  const MIN_BIN_N    = 5;                              // suppress bins with < 5 obs

  // ============================================================
  // READ SOURCE DATA
  // ============================================================
  const ss    = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getActiveSheet();
  const data  = sheet.getDataRange().getValues();
  const headers = data[0];

  const xIdx = headers.indexOf(X_COL_NAME);
  const yIdx = headers.indexOf(Y_COL_NAME);
  if (xIdx === -1 || yIdx === -1) {
    throw new Error(`Column not found: "${X_COL_NAME}" or "${Y_COL_NAME}"`);
  }

  // Filter to valid numeric rows
  const naPattern = /^n[\.\s\/]*a[\.\s]*$/i;
  const valid = [];
  for (let i = 1; i < data.length; i++) {
    const xRaw = data[i][xIdx];
    const yRaw = data[i][yIdx];
    if (xRaw === "" || yRaw === "" || xRaw == null || yRaw == null) continue;
    if (naPattern.test(String(xRaw).trim()) || naPattern.test(String(yRaw).trim())) continue;
    const x = Number(xRaw);
    const y = Number(yRaw);
    if (!isNaN(x) && !isNaN(y)) valid.push({ x, y });
  }

  // ============================================================
  // BIN AND AGGREGATE
  // ============================================================
  // Determine range
  const xValues = valid.map(r => r.x);
  const xMin = Math.floor(Math.min(...xValues));
  const xMax = Math.ceil(Math.max(...xValues));

  // Create bins
  const bins = {};
  for (let edge = xMin; edge < xMax; edge += BIN_WIDTH) {
    const label = edge + BIN_WIDTH / 2; // midpoint
    bins[label] = [];
  }

  // Assign each observation to a bin
  valid.forEach(row => {
    const binStart = xMin + Math.floor((row.x - xMin) / BIN_WIDTH) * BIN_WIDTH;
    const midpoint = binStart + BIN_WIDTH / 2;
    if (bins[midpoint]) {
      bins[midpoint].push(row.y);
    }
  });

  // Compute mean and standard error per bin
  const summaryRows = [["Bin Midpoint (" + X_COL_NAME + ")",
                         "Mean " + Y_COL_NAME,
                         "SE (±)",
                         "n"]];

  const sortedMidpoints = Object.keys(bins).map(Number).sort((a, b) => a - b);
  sortedMidpoints.forEach(mid => {
    const vals = bins[mid];
    if (vals.length < MIN_BIN_N) return; // suppress small bins
    const n    = vals.length;
    const mean = vals.reduce((a, b) => a + b, 0) / n;
    const sd   = Math.sqrt(vals.reduce((s, v) => s + (v - mean) ** 2, 0) / (n - 1));
    const se   = sd / Math.sqrt(n);
    summaryRows.push([
      Math.round(mid * 100) / 100,
      Math.round(mean * 100) / 100,
      Math.round(se * 100) / 100,
      n
    ]);
  });

  // ============================================================
  // WRITE TO NEW SHEET
  // ============================================================
  let sheetName = `${sheet.getName()} - Binned ${X_COL_NAME} vs ${Y_COL_NAME}`;
  if (sheetName.length > 100) sheetName = sheetName.substring(0, 99);
  let outSheet = ss.getSheetByName(sheetName);
  if (outSheet) ss.deleteSheet(outSheet);
  outSheet = ss.insertSheet(sheetName);

  outSheet.getRange(1, 1, summaryRows.length, 4).setValues(summaryRows);
  outSheet.getRange(1, 1, 1, 4).setFontWeight("bold");
  outSheet.autoResizeColumns(1, 4);

  // ============================================================
  // BUILD SCATTER CHART OF BIN MEANS
  // ============================================================
  const numRows = summaryRows.length;
  const xRange  = outSheet.getRange(1, 1, numRows, 1);
  const yRange  = outSheet.getRange(1, 2, numRows, 1);

  // Remove old charts
  outSheet.getCharts().forEach(c => outSheet.removeChart(c));

  const chart = outSheet.newChart()
    .setChartType(Charts.ChartType.SCATTER)
    .addRange(xRange)
    .addRange(yRange)
    .setPosition(1, 6, 0, 0)
    .setOption("title", `Mean ${Y_COL_NAME} by ${X_COL_NAME} Bin`)
    .setOption("hAxis", { title: X_COL_NAME })
    .setOption("vAxis", { title: `Mean ${Y_COL_NAME}` })
    .setOption("pointSize", 8)
    .setOption("colors", ["#2c3e50"])
    .setOption("legend", { position: "top" })
    .setOption("trendlines", {
      0: {
        type: "polynomial",
        degree: 2,
        color: "#e74c3c",
        lineWidth: 3,
        showR2: true,
        visibleInLegend: true,
        labelInLegend: "Quadratic fit"
      }
    })
    .setOption("width", 700)
    .setOption("height", 450)
    .build();

  outSheet.insertChart(chart);
  ss.setActiveSheet(outSheet);
}
