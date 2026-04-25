// ============================================================
// COMBINED SCRIPT: Cardiovascular Scatter Plots 
// Includes: Mean line overlays, statistical legends, and data cleaning
// ============================================================

const SOURCE_SHEET_NAME = 'The gut microbiome in atherosclerotic'; // Change if needed

/**
 * Main execution function. Reads the source data and iterates through
 * the configured cardiovascular metrics to generate charts.
 */
function generateAllFigures() {
  Logger.log('=== FIGURE GENERATION START ===');

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(SOURCE_SHEET_NAME);
  
  if (!sheet) {
    throw new Error(`Source sheet "${SOURCE_SHEET_NAME}" not found. Please check the tab name.`);
  }

  const data = sheet.getDataRange().getValues();
  const headers = data[0];
  const rows = data.slice(1);
  Logger.log(`Loaded ${rows.length} data rows from source sheet.`);

  // Map header names to their column indices
  const idx = {
    bmi:   headers.indexOf('Body Mass Index (BMI)'),
    waist: headers.indexOf('Waist (cm)'),
    sbp:   headers.indexOf('Systolic Blood Pressure (mmHg)'),
    dbp:   headers.indexOf('Diastolic Blood Pressure (mmHg)'),
    ldl:   headers.indexOf('LDLC (mmol/L)'),
    chol:  headers.indexOf('CHOL (mmol/L)')
  };

  for (const k in idx) {
    if (idx[k] === -1) throw new Error(`Column not found: ${k}. Check exact header spelling.`);
  }
  Logger.log('All column indices successfully resolved.');

  // Configuration Array: [sheetName, xKey, yKey, xLabel, yLabel, figNum, figDesc]
  const figures = [
    ['Fig1a_BMI_vs_SBP',   'bmi',   'sbp',  'BMI (kg/m^2)',             'Systolic Blood Pressure (mmHg)',  'Figure 1a', 'BMI vs Systolic BP'],
    ['Fig1b_BMI_vs_DBP',   'bmi',   'dbp',  'BMI (kg/m^2)',             'Diastolic Blood Pressure (mmHg)', 'Figure 1b', 'BMI vs Diastolic BP'],
    ['Fig2_BMI_vs_LDL',    'bmi',   'ldl',  'BMI (kg/m^2)',             'LDL Cholesterol (mmol/L)',        'Figure 2',  'BMI vs LDL Cholesterol'],
    ['Fig3_BMI_vs_CHOL',   'bmi',   'chol', 'BMI (kg/m^2)',             'Total Cholesterol (mmol/L)',      'Figure 3',  'BMI vs Total Cholesterol'],
    ['Fig4a_Waist_vs_SBP', 'waist', 'sbp',  'Waist Circumference (cm)', 'Systolic Blood Pressure (mmHg)',  'Figure 4a', 'Waist Circumference vs Systolic BP'],
    ['Fig4b_Waist_vs_DBP', 'waist', 'dbp',  'Waist Circumference (cm)', 'Diastolic Blood Pressure (mmHg)', 'Figure 4b', 'Waist Circumference vs Diastolic BP'],
    ['Fig5_Waist_vs_LDL',  'waist', 'ldl',  'Waist Circumference (cm)', 'LDL Cholesterol (mmol/L)',        'Figure 5',  'Waist Circumference vs LDL Cholesterol'],
    ['Fig6_Waist_vs_CHOL', 'waist', 'chol', 'Waist Circumference (cm)', 'Total Cholesterol (mmol/L)',      'Figure 6',  'Waist Circumference vs Total Cholesterol']
  ];

  // Clean up previous runs by deleting old figure sheets
  figures.forEach(f => {
    const existing = ss.getSheetByName(f[0]);
    if (existing) { 
      ss.deleteSheet(existing); 
      Logger.log(`Deleted old sheet: ${f[0]}`); 
    }
  });

  // Build each figure
  figures.forEach(f => {
    Logger.log(`--- Building ${f[5]} ---`);
    buildFigure(ss, rows, idx, f[0], f[1], f[2], f[3], f[4], f[5], f[6]);
  });

  Logger.log('=== ALL FIGURES COMPLETE ===');
}

/**
 * Extracts data, calculates statistics, and renders the scatter plot with a mean line.
 */
function buildFigure(ss, rows, idx, sheetName, xKey, yKey, xLabel, yLabel, figNum, figDesc) {
  const pairs = [];
  const naPattern = /^n[\.\/\s]*a[\.]*$/i; 

  // --- Step 1: Data Extraction & Filtering ---
  rows.forEach(r => {
    const xRaw = r[idx[xKey]];
    const yRaw = r[idx[yKey]];
    
    if (xRaw === '' || yRaw === '' || xRaw == null || yRaw == null) return;
    if (naPattern.test(String(xRaw).trim()) || naPattern.test(String(yRaw).trim())) return;
    
    const x = Number(xRaw);
    const y = Number(yRaw);
    
    if (!isNaN(x) && !isNaN(y)) {
      pairs.push({ x, y });
    }
  });

  if (pairs.length === 0) {
    Logger.log(`  WARNING: No valid pairs for ${sheetName}`);
    return;
  }
  Logger.log(`  n = ${pairs.length} valid pairs extracted.`);

  // --- Step 2: Sorting ---
  // CRITICAL: Sort by X ascending so the secondary mean line draws cleanly from left to right
  pairs.sort((a, b) => a.x - b.x);

  // --- Step 3: Statistical Calculations ---
  const n = pairs.length;
  let sumY = 0, sumX = 0;
  pairs.forEach(p => { sumY += p.y; sumX += p.x; });
  
  const meanY = sumY / n;
  const meanX = sumX / n;

  let sumSqY = 0;
  pairs.forEach(p => { sumSqY += (p.y - meanY) ** 2; });
  const sdY = Math.sqrt(sumSqY / (n - 1));

  let num = 0, denX = 0, denY = 0;
  pairs.forEach(p => {
    num  += (p.x - meanX) * (p.y - meanY);
    denX += (p.x - meanX) ** 2;
    denY += (p.y - meanY) ** 2;
  });
  const r = num / Math.sqrt(denX * denY);

  // --- Step 4: Data Staging Output ---
  const out = ss.insertSheet(sheetName);

  // The headers dictate the text that populates the chart legend
  const pointsLabel = `Individual Participants (n = ${n})`;
  const meanLabel   = `Cohort Mean = ${meanY.toFixed(2)}`;
  
  out.getRange(1, 1, 1, 3).setValues([[xLabel, pointsLabel, meanLabel]]);
  out.getRange(1, 1, 1, 3).setFontWeight('bold');

  // Map the pairs into 3 columns: X, Y, and the constant Mean Y for the solid line
  const outputRows = pairs.map(p => [p.x, p.y, meanY]);
  out.getRange(2, 1, outputRows.length, 3).setValues(outputRows);
  SpreadsheetApp.flush();

  const numRows = outputRows.length + 1; // Include header row

  // --- Step 5: Isolate Ranges (CRITICAL FIX FOR LEGEND LABELS) ---
  // By breaking the data into explicit ranges, Apps Script is forced to properly 
  // identify the column headers and attach them to the scatter legend.
  const xRange = out.getRange(1, 1, numRows, 1);    // Domain (X-Axis)
  const yRange = out.getRange(1, 2, numRows, 1);    // Series 0 (Scatter Points)
  const meanRange = out.getRange(1, 3, numRows, 1); // Series 1 (Mean Line)

  // --- Step 6: Chart Generation ---
  const richTitle = `${figNum}: ${figDesc} (n=${n} | Mean=${meanY.toFixed(2)} | SD=${sdY.toFixed(2)} | r=${r.toFixed(3)})`;

  const chart = out.newChart()
    .setChartType(Charts.ChartType.SCATTER)
    .addRange(xRange)
    .addRange(yRange)
    .addRange(meanRange)
    .setPosition(2, 5, 0, 0)
    .setOption('useFirstColumnAsDomain', true) // Forces X-axis to bind properly
    .setOption('title', richTitle)
    .setOption('titleTextStyle', { fontSize: 14, bold: true, color: '#333333' })
    .setOption('width', 720)
    .setOption('height', 480)
    .setOption('hAxis', {
      title: xLabel,
      titleTextStyle: { fontSize: 12, bold: true, italic: false },
      gridlines: { color: '#e0e0e0' }
    })
    .setOption('vAxis', {
      title: yLabel,
      titleTextStyle: { fontSize: 12, bold: true, italic: false },
      gridlines: { color: '#e0e0e0' }
    })
    .setOption('legend', {
      position: 'bottom',
      alignment: 'center',
      textStyle: { fontSize: 11 }
    })
    .setOption("trendlines", {
      0: {
        type: "linear", // "polynomial" curves; "linear" provides straight line of best fit
        color: "#e74c3c", // Red for the trendline
        lineWidth: 2,
        opacity: 0.7,
        showR2: true,
        visibleInLegend: true,
        labelInLegend: "Trendline (Best Fit)"
      }
    })
    .setOption('series', {
      // Series 0: The actual participant scatter points
      0: {
        pointShape: 'circle',
        pointSize: 6,
        lineWidth: 0,
        color: '#1f77b4',
        dataOpacity: 0.75,
        visibleInLegend: true
      },
      // Series 1: The mean trendline (hidden points, visible solid line)
      // Removed lineDashStyle as it frequently blocks rendering in Scatter Charts
      1: {
        pointSize: 0,
        lineWidth: 3,
        color: '#d62728',
        visibleInLegend: true
      }
    })
    .setOption('backgroundColor', { fill: '#ffffff' })
    .build();

  out.insertChart(chart);
  Logger.log(`  Chart successfully inserted into ${sheetName}`);
}
