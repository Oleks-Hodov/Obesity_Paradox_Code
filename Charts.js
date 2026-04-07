/**
 * ============================================================
 *  CARDIOVASCULAR PHENOTYPE GROUPING — Google Sheets Script
 * ============================================================
 *  Groups:
 *    0 = Missing data
 *    1 = Lean   + Favorable
 *    2 = Obese  + Favorable
 *    3 = Lean   + Unfavorable
 *    4 = Obese  + Unfavorable
 *
 *  MISSING DATA RULES:                                         ← UPDATED
 *    - Waist is ALWAYS required → missing = Group 0
 *    - If HTN column present + valid (0 or 1):
 *        → SBP/DBP become optional; HTN alone determines hypertension
 *        → If SBP+DBP also present, all three are combined (OR logic)
 *    - If HTN column absent or has invalid value:
 *        → Both SBP AND DBP required → missing either = Group 0
 *
 *  FAVORABLE   = SBP <130 AND DBP <85 AND HTN=0  (when all present)
 *              = HTN=0                            (when only HTN present)
 *  UNFAVORABLE = SBP ≥130 OR  DBP ≥85 OR  HTN=1  (when all present)
 *              = HTN=1                            (when only HTN present)
 * ============================================================
 */

// ── Thresholds ─────────────────────────────────────────────────
var WAIST_CUTOFF_CM = 94;
var SBP_CUTOFF      = 130;
var DBP_CUTOFF      = 85;
var OUTPUT_COL_NAME = "Group #";

// ── Column aliases ─────────────────────────────────────────────
var WAIST_ALIASES = [
  "waist circumference", "waist_circumference", "waist circ",
  "waist circ.", "waistcircumference", "waist (cm)", "waist",
  "wc", "waist circumference (cm)", "abdominal circumference"
];

var SBP_ALIASES = [
  "systolic blood pressure", "systolic bp", "sbp", "sys bp",
  "systolic", "blood pressure systolic", "systolicbp",
  "systolic_blood_pressure", "sys", "systolic pressure"
];

var DBP_ALIASES = [
  "diastolic blood pressure", "diastolic bp", "dbp", "dia bp",
  "diastolic", "blood pressure diastolic", "diastolicbp",
  "diastolic_blood_pressure", "dia", "diastolic pressure"
];

var HTN_ALIASES = [
  "hypertension", "htn", "hypertensive", "high blood pressure",
  "hypertension (1/0)", "hypertension_status", "bp_diagnosis",
  "hypertension diagnosis", "has hypertension", "hypertension_flag",
  "hypertension (yes=1)", "hypertension (0/1)", "arterial hypertension"
];


// ── MAIN FUNCTION ──────────────────────────────────────────────
function assignCVGroups() {

  var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  var data  = sheet.getDataRange().getValues();

  if (data.length < 2) {
    SpreadsheetApp.getUi().alert("⚠️ Sheet appears empty or has no data rows.");
    return;
  }

  var headers = data[0].map(function(h) { return String(h).trim(); });

  // ── Locate columns ────────────────────────────────────────────
  var waistIdx = findColumnIndex(headers, WAIST_ALIASES);
  var sbpIdx   = findColumnIndex(headers, SBP_ALIASES);
  var dbpIdx   = findColumnIndex(headers, DBP_ALIASES);
  var htnIdx   = findColumnIndex(headers, HTN_ALIASES);

  // ── Waist is always required ──────────────────────────────────
  if (waistIdx === -1) {
    SpreadsheetApp.getUi().alert(
      "❌ Could not find the Waist Circumference column.\n\n" +
      "Headers found:\n" + headers.join(", ") +
      "\n\nAdd an alias to WAIST_ALIASES at the top of the script."
    );
    return;
  }

  // ← UPDATED: SBP and DBP are only required if HTN column is absent
  var bpMissing = [];
  if (sbpIdx === -1) bpMissing.push("Systolic Blood Pressure (SBP)");
  if (dbpIdx === -1) bpMissing.push("Diastolic Blood Pressure (DBP)");

  if (bpMissing.length > 0 && htnIdx === -1) {
    // Neither BP nor HTN column available — cannot classify
    SpreadsheetApp.getUi().alert(
      "❌ Missing required columns:\n\n" + bpMissing.join("\n") +
      "\n\nNo Hypertension column was found either, so classification " +
      "is impossible without at least one of these data sources.\n\n" +
      "Headers found:\n" + headers.join(", ")
    );
    return;
  }

  // ← UPDATED: Warn (but continue) if only some BP columns are missing
  if (bpMissing.length > 0 && htnIdx !== -1) {
    SpreadsheetApp.getUi().alert(
      "⚠️ Warning: The following BP columns were not found:\n" +
      bpMissing.join("\n") +
      "\n\nA Hypertension column was detected → classification will " +
      "rely on HTN values for rows where BP data is unavailable.\n\n" +
      "Click OK to continue."
    );
  }

  // ── Build HTN column status string for summary ────────────────
  var htnStatus = (htnIdx !== -1)
    ? "✅ Found → " + headers[htnIdx] + " (col " + (htnIdx + 1) + ")"
    : "⚠️ Not found — BP values only";

  // ── Find or create output column ──────────────────────────────
  var groupColIdx = headers.indexOf(OUTPUT_COL_NAME);
  if (groupColIdx === -1) {
    groupColIdx = headers.length;
    sheet.getRange(1, groupColIdx + 1).setValue(OUTPUT_COL_NAME);
  }

  // ── Process rows ──────────────────────────────────────────────
  var groupValues = [];
  var stats = { 0: 0, 1: 0, 2: 0, 3: 0, 4: 0 };

  for (var i = 1; i < data.length; i++) {
    var row = data[i];

    var waist  = parseFloat(row[waistIdx]);
    var sbp    = (sbpIdx  !== -1) ? parseFloat(row[sbpIdx])  : NaN;  // ← UPDATED
    var dbp    = (dbpIdx  !== -1) ? parseFloat(row[dbpIdx])  : NaN;  // ← UPDATED
    var htnRaw = (htnIdx  !== -1) ? row[htnIdx]              : null;

    var group = classifySample(waist, sbp, dbp, htnRaw, htnIdx !== -1);
    groupValues.push([group]);
    stats[group]++;
  }

  // ── Write & format ────────────────────────────────────────────
  sheet.getRange(2, groupColIdx + 1, groupValues.length, 1).setValues(groupValues);
  applyColorCoding(sheet, groupColIdx + 1, groupValues);
  sheet.getRange(1, groupColIdx + 1).setFontWeight("bold");

  // ── Summary popup ─────────────────────────────────────────────
  var total = data.length - 1;

  // ← UPDATED: dynamic rule description based on available columns
  var favorableRule   = buildRuleString(sbpIdx, dbpIdx, htnIdx, false);
  var unfavorableRule = buildRuleString(sbpIdx, dbpIdx, htnIdx, true);

  SpreadsheetApp.getUi().alert(
    "✅ Grouping complete!\n\n" +
    "Columns used:\n" +
    "  Waist → " + headers[waistIdx] + " (col " + (waistIdx + 1) + ")\n" +
    "  SBP   → " + (sbpIdx !== -1 ? headers[sbpIdx] + " (col " + (sbpIdx + 1) + ")" : "Not found") + "\n" +
    "  DBP   → " + (dbpIdx !== -1 ? headers[dbpIdx] + " (col " + (dbpIdx + 1) + ")" : "Not found") + "\n" +
    "  HTN   → " + htnStatus + "\n\n" +
    "Rules applied:\n" +
    "  Favorable   : " + favorableRule   + "\n" +
    "  Unfavorable : " + unfavorableRule + "\n\n" +
    "Results (" + total + " samples):\n" +
    "  Group 0 – Missing data        : " + stats[0] + " (" + pct(stats[0], total) + ")\n" +
    "  Group 1 – Lean   + Favorable  : " + stats[1] + " (" + pct(stats[1], total) + ")\n" +
    "  Group 2 – Obese  + Favorable  : " + stats[2] + " (" + pct(stats[2], total) + ")\n" +
    "  Group 3 – Lean   + Unfavorable: " + stats[3] + " (" + pct(stats[3], total) + ")\n" +
    "  Group 4 – Obese  + Unfavorable: " + stats[4] + " (" + pct(stats[4], total) + ")\n\n" +
    "Thresholds:\n" +
    "  Obese       : Waist ≥ " + WAIST_CUTOFF_CM + " cm\n" +
    "  HTN via BP  : SBP ≥ " + SBP_CUTOFF + " OR DBP ≥ " + DBP_CUTOFF + " mm Hg\n" +
    "  HTN via col : HTN column = 1"
  );
}


// ── CLASSIFICATION LOGIC ──────────────────────────────────────
// ← UPDATED: full rewrite to handle BP-only, HTN-only, or combined
function classifySample(waist, sbp, dbp, htnRaw, htnPresent) {

  // ── Waist always required ─────────────────────────────────────
  if (!isValidNumber(waist)) return 0;

  // ── Parse HTN column value ────────────────────────────────────
  var htnValid = false;
  var htnValue = false;  // true = has hypertension

  if (htnPresent) {
    var htnStr = String(htnRaw).trim();
    if (htnStr === "1") { htnValid = true; htnValue = true;  }
    if (htnStr === "0") { htnValid = true; htnValue = false; }
    // Any other value (blank, NaN, text) → htnValid stays false
  }

  // ── Parse BP values ───────────────────────────────────────────
  var sbpValid = isValidNumber(sbp);
  var dbpValid = isValidNumber(dbp);
  var bpValid  = sbpValid && dbpValid;   // need BOTH for BP-based HTN

  // ── Determine if hypertension status can be established ───────
  // ← UPDATED: HTN column alone is sufficient when BP is missing
  var canClassify = htnValid || bpValid;
  if (!canClassify) return 0;           // Group 0: no usable HTN data

  // ── Compute hypertension flag ─────────────────────────────────
  // ← UPDATED: three scenarios
  var hypertensive;

  if (bpValid && htnValid) {
    // CASE A — both BP and HTN column available → combine (OR logic)
    hypertensive = (sbp >= SBP_CUTOFF) || (dbp >= DBP_CUTOFF) || htnValue;

  } else if (htnValid && !bpValid) {
    // CASE B — BP missing but HTN column valid → HTN column only     ← UPDATED
    hypertensive = htnValue;

  } else {
    // CASE C — BP available but no valid HTN column → BP only
    hypertensive = (sbp >= SBP_CUTOFF) || (dbp >= DBP_CUTOFF);
  }

  // ── Obesity flag ──────────────────────────────────────────────
  var obese = (waist >= WAIST_CUTOFF_CM);

  // ── Assign group ──────────────────────────────────────────────
  if (!obese && !hypertensive) return 1;  // Lean  + Favorable
  if ( obese && !hypertensive) return 2;  // Obese + Favorable
  if (!obese &&  hypertensive) return 3;  // Lean  + Unfavorable
  if ( obese &&  hypertensive) return 4;  // Obese + Unfavorable
}


// ── HELPER: validate a number ─────────────────────────────────
function isValidNumber(val) {
  return (val !== null && val !== "" && !isNaN(val) && isFinite(val));
}


// ← UPDATED: builds a human-readable rule string for the summary popup
function buildRuleString(sbpIdx, dbpIdx, htnIdx, unfavorable) {
  var hasBP  = (sbpIdx !== -1 && dbpIdx !== -1);
  var hasHTN = (htnIdx !== -1);

  if (hasBP && hasHTN) {
    return unfavorable
      ? "SBP ≥130 OR DBP ≥85 OR HTN=1"
      : "SBP <130 AND DBP <85 AND HTN=0";
  }
  if (hasHTN && !hasBP) {
    return unfavorable ? "HTN=1 (no BP columns found)" : "HTN=0 (no BP columns found)";
  }
  return unfavorable ? "SBP ≥130 OR DBP ≥85" : "SBP <130 AND DBP <85";
}


// ── COLUMN FINDER ─────────────────────────────────────────────
function findColumnIndex(headers, aliases) {
  for (var i = 0; i < headers.length; i++) {
    var h = headers[i].toLowerCase().trim();
    for (var j = 0; j < aliases.length; j++) {
      if (h === aliases[j].toLowerCase().trim()) return i;
    }
  }
  for (var i = 0; i < headers.length; i++) {
    var h = headers[i].toLowerCase().trim();
    for (var j = 0; j < aliases.length; j++) {
      var a = aliases[j].toLowerCase().trim();
      if (h.indexOf(a) !== -1 || a.indexOf(h) !== -1) return i;
    }
  }
  return -1;
}


// ── COLOR CODING ──────────────────────────────────────────────
function applyColorCoding(sheet, colNumber, groupValues) {
  var colorMap = {
    0: "#F5F5F5",  // Grey   – Missing
    1: "#C8E6C9",  // Green  – Lean   + Favorable
    2: "#FFF9C4",  // Yellow – Obese  + Favorable
    3: "#FFE0B2",  // Orange – Lean   + Unfavorable
    4: "#FFCDD2"   // Red    – Obese  + Unfavorable
  };
  for (var i = 0; i < groupValues.length; i++) {
    var group = groupValues[i][0];
    sheet.getRange(i + 2, colNumber).setBackground(colorMap[group] || "#FFFFFF");
  }
}


// ── HELPER: percentage ────────────────────────────────────────
function pct(count, total) {
  if (total === 0) return "0%";
  return (count / total * 100).toFixed(1) + "%";
}


// ── CUSTOM MENU ───────────────────────────────────────────────
function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu("🫀 CV Phenotyping")
    .addItem("Assign CV Groups", "assignCVGroups")
    .addToUi();
}
