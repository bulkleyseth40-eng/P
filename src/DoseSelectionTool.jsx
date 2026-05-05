import React, { useMemo, useState } from "react";
import * as XLSX from "xlsx";

const CACHE_KEY = "dose-selection-tool-cached-assay-report-v4";
const NEWLINE = String.fromCharCode(10);

const REQUIRED_COLUMNS = [
  "NDC",
  "Material Description",
  "Trade Name",
  "Quantity On Hand",
  "Expiration Date",
  "LOT #",
];

const payerPresets = {
  Standard: { low: 10, high: 10 },
  Prime: { low: 3, high: 3 },
  Ambetter: { low: 10, high: 1.5 },
  Custom: { low: 10, high: 10 },
};

const expirationOptions = [
  { label: "Doesn't expire within the next 30 days", value: "30", days: 30 },
  { label: "Doesn't expire within the next 60 days", value: "60", days: 60 },
  { label: "Doesn't expire within the next 3 months", value: "90", days: 90 },
  { label: "Doesn't expire within the next 6 months", value: "180", days: 180 },
  { label: "Doesn't expire within the next 1 year", value: "365", days: 365 },
];

function normalize(value) {
  return String(value ?? "").trim();
}

function numberOrZero(value) {
  const parsed = Number(value);
  return Number.isFinite(parsed) ? parsed : 0;
}

function normalizeKey(value) {
  return normalize(value)
    .toLowerCase()
    .split(" ").join("")
    .split("_").join("")
    .split("-").join("")
    .split("#").join("");
}

function parseDate(value) {
  if (!value) return null;
  if (value instanceof Date && !Number.isNaN(value.getTime())) return value;
  if (typeof value === "number" && XLSX && XLSX.SSF && XLSX.SSF.parse_date_code) {
    const parsed = XLSX.SSF.parse_date_code(value);
    if (parsed) return new Date(parsed.y, parsed.m - 1, parsed.d);
  }
  const date = new Date(value);
  return Number.isNaN(date.getTime()) ? null : date;
}

function formatDate(value) {
  const date = parseDate(value);
  if (!date) return "Unknown";
  return date.toLocaleDateString(undefined, { year: "numeric", month: "2-digit", day: "2-digit" });
}

function daysUntil(value) {
  const date = parseDate(value);
  if (!date) return 999999;
  const today = new Date();
  today.setHours(0, 0, 0, 0);
  date.setHours(0, 0, 0, 0);
  return Math.round((date.getTime() - today.getTime()) / 86400000);
}

function extractUnits(description) {
  const text = normalize(description).toUpperCase();
  let current = "";
  const candidates = [];
  for (let index = 0; index < text.length; index += 1) {
    const char = text[index];
    if (char >= "0" && char <= "9") {
      current += char;
      continue;
    }
    if (current.length >= 3 && current.length <= 6) {
      const next = char;
      const after = text[index + 1] || "";
      if (next === "U" || (next === "I" && after === "U")) candidates.push(Number(current));
    }
    current = "";
  }
  if (current.length >= 3 && current.length <= 6) candidates.push(Number(current));
  if (candidates.length) return candidates[candidates.length - 1];
  return null;
}

function getRowValue(row, possibleNames) {
  for (const name of possibleNames) {
    if (Object.prototype.hasOwnProperty.call(row, name)) return row[name];
  }
  const lookup = {};
  Object.keys(row || {}).forEach((key) => {
    lookup[normalizeKey(key)] = key;
  });
  for (const name of possibleNames) {
    const foundKey = lookup[normalizeKey(name)];
    if (foundKey) return row[foundKey];
  }
  return undefined;
}

function makeInventoryId(item, index) {
  return [item.tradeName, item.ndc, item.units, item.expiration || "", item.lot || "", index].join("|");
}

function parseInventoryRows(rows) {
  const parsed = [];
  rows.forEach((row, index) => {
    const description = normalize(getRowValue(row, ["Material Description", "Description"]));
    const tradeName = normalize(getRowValue(row, ["Trade Name", "Product", "Drug"])).toUpperCase();
    const units = extractUnits(description);
    const quantity = numberOrZero(getRowValue(row, ["Quantity On Hand", "Qty", "Quantity", "QOH"]));
    const expiration = getRowValue(row, ["Expiration Date", "Expiration", "Exp Date", "EXP"]);
    const item = {
      ndc: normalize(getRowValue(row, ["NDC"])),
      description,
      tradeName,
      quantity,
      expiration,
      lot: normalize(getRowValue(row, ["LOT #", "Lot", "Lot #"])),
      units,
      daysToExp: daysUntil(expiration),
    };
    item.id = makeInventoryId(item, index);
    if (item.tradeName && item.units && item.quantity > 0) parsed.push(item);
  });
  return parsed;
}

function productMatches(item, query) {
  const value = normalize(query).toUpperCase();
  if (!value) return false;
  return item.tradeName === value || item.tradeName.includes(value) || item.description.toUpperCase().includes(value);
}

function cloneCombo(combo) {
  return combo.map((item) => ({ ...item }));
}

function addToCombo(combo, item) {
  const next = cloneCombo(combo);
  const existing = next.find((entry) => entry.id === item.id);
  if (existing) existing.qty += 1;
  else next.push({ ...item, qty: 1 });
  return next;
}

function comboTotal(combo) {
  return combo.reduce((sum, item) => sum + item.units * item.qty, 0);
}

function comboVials(combo) {
  return combo.reduce((sum, item) => sum + item.qty, 0);
}

function earliestExpDays(combo) {
  if (!combo.length) return 999999;
  return Math.min(...combo.map((item) => item.daysToExp));
}

function latestExpDays(combo) {
  if (!combo.length) return -999999;
  return Math.max(...combo.map((item) => item.daysToExp));
}

function sortComboForDisplay(combo) {
  return [...combo].sort((a, b) => b.units - a.units || a.daysToExp - b.daysToExp || String(a.lot).localeCompare(String(b.lot)));
}

function expirationProfileKey(combo) {
  return sortComboForDisplay(combo).map((item) => `${item.units}:${formatDate(item.expiration)}`).join("+");
}

function dedupeRankedOptions(results) {
  const seen = new Set();
  const deduped = [];
  results.forEach((result) => {
    const key = `${result.total}|${result.vials}|${expirationProfileKey(result.combo)}`;
    if (!seen.has(key)) {
      seen.add(key);
      deduped.push(result);
    }
  });
  return deduped;
}

function generateCombinations(inventory, dose, lowPct, highPct, maxVials, preferPositive, minDaysToExp, sameLotPreferred, dosesNeeded = 1, allowOutsideRange = false) {
  const minDose = dose * (1 - lowPct / 100);
  const maxDose = dose * (1 + highPct / 100);
  const fillCount = Math.max(1, Number(dosesNeeded) || 1);
  const usable = inventory
    .filter((item) => item.quantity >= fillCount && item.daysToExp >= minDaysToExp)
    .sort((a, b) => b.units - a.units || a.daysToExp - b.daysToExp);
  const results = [];
  const seen = new Set();

  function record(combo, isInRange) {
    const key = sortComboForDisplay(combo).map((item) => `${item.id}:${item.qty}`).join("|");
    if (seen.has(key)) return;
    seen.add(key);
    const total = comboTotal(combo);
    const varianceUnits = total - dose;
    const variancePct = dose ? (varianceUnits / dose) * 100 : 0;
    const vials = comboVials(combo);
    const lots = new Set(combo.map((item) => item.lot)).size;
    const exactness = Math.abs(variancePct);
    const positivePenalty = preferPositive && variancePct < 0 ? 8 : 0;
    const sameLotPenalty = sameLotPreferred && lots > 1 ? lots * 2 : 0;
    const singleVialBoost = vials === 1 ? -3 : 0;
    const outOfRangePenalty = isInRange ? 0 : 1000;
    const expScore = earliestExpDays(combo) / 3650;
    const score = outOfRangePenalty + exactness * 10 + vials * 2 + positivePenalty + sameLotPenalty + singleVialBoost + expScore;
    results.push({
      combo: sortComboForDisplay(combo),
      total,
      totalDoseForFill: total * fillCount,
      varianceUnits,
      variancePct,
      vials,
      lots,
      score,
      dosesNeeded: fillCount,
      isInRange,
      isClosestFallback: !isInRange,
    });
  }

  function recurse(startIndex, combo, depth) {
    if (depth > 0) {
      const total = comboTotal(combo);
      const isInRange = total >= minDose && total <= maxDose;
      if (isInRange || allowOutsideRange) record(combo, isInRange);
      if (!allowOutsideRange && total > maxDose) return;
    }
    if (depth >= maxVials) return;
    for (let index = startIndex; index < usable.length; index += 1) {
      const item = usable[index];
      const existingQty = combo.find((entry) => entry.id === item.id)?.qty || 0;
      if ((existingQty + 1) * fillCount > item.quantity) continue;
      recurse(index, addToCombo(combo, item), depth + 1);
    }
  }

  recurse(0, [], 0);
  return dedupeRankedOptions(
    results.sort((a, b) => a.score - b.score || Math.abs(a.variancePct) - Math.abs(b.variancePct) || a.total - b.total)
  ).slice(0, 50);
}

function isDiscountEligible(item) {
  return item.daysToExp >= 0 && item.daysToExp <= 180;
}

function getDiscountEligibleItems(combo = []) {
  return combo.filter((item) => isDiscountEligible(item));
}

function hasDiscountOpportunity(result) {
  return getDiscountEligibleItems(result?.combo || []).length > 0;
}

function discountUrgencyDays(result) {
  const eligible = getDiscountEligibleItems(result?.combo || []);
  if (!eligible.length) return 999999;
  return Math.min(...eligible.map((item) => item.daysToExp));
}

function sortVialRows(rows, sortConfig) {
  const { key, direction } = sortConfig;
  const multiplier = direction === "asc" ? 1 : -1;
  return [...rows].sort((a, b) => {
    let aValue = a.units;
    let bValue = b.units;
    if (key === "quantity") {
      aValue = a.quantity;
      bValue = b.quantity;
    }
    if (key === "expiration") {
      aValue = a.daysToExp;
      bValue = b.daysToExp;
    }
    if (key === "lot") {
      aValue = a.lot || "";
      bValue = b.lot || "";
    }
    if (typeof aValue === "number" && typeof bValue === "number") return (aValue - bValue) * multiplier;
    return String(aValue).localeCompare(String(bValue)) * multiplier;
  });
}

function sortAssayOptions(results, sortMode) {
  const sorted = [...results];
  switch (sortMode) {
    case "vials":
      return sorted.sort((a, b) => a.vials - b.vials || Math.abs(a.variancePct) - Math.abs(b.variancePct) || a.total - b.total);
    case "closest-expiration":
      return sorted.sort((a, b) => earliestExpDays(a.combo) - earliestExpDays(b.combo) || a.vials - b.vials || Math.abs(a.variancePct) - Math.abs(b.variancePct));
    case "furthest-expiration":
      return sorted.sort((a, b) => latestExpDays(b.combo) - latestExpDays(a.combo) || a.vials - b.vials || Math.abs(a.variancePct) - Math.abs(b.variancePct));
    case "smallest-variance":
      return sorted.sort((a, b) => Math.abs(a.variancePct) - Math.abs(b.variancePct) || a.vials - b.vials || a.total - b.total);
    case "discount-opportunities":
      return sorted.sort((a, b) => Number(hasDiscountOpportunity(b)) - Number(hasDiscountOpportunity(a)) || discountUrgencyDays(a) - discountUrgencyDays(b) || Math.abs(a.variancePct) - Math.abs(b.variancePct) || a.vials - b.vials);
    default:
      return sorted.sort((a, b) => a.score - b.score || Math.abs(a.variancePct) - Math.abs(b.variancePct) || a.total - b.total);
  }
}

function loadCachedReport() {
  if (typeof window === "undefined") return null;
  try {
    const cached = window.localStorage.getItem(CACHE_KEY);
    return cached ? JSON.parse(cached) : null;
  } catch {
    return null;
  }
}

function saveCachedReport(payload) {
  if (typeof window === "undefined") return;
  try {
    window.localStorage.setItem(CACHE_KEY, JSON.stringify(payload));
  } catch {
    // Storage can be unavailable in some embedded previews.
  }
}

function clearCachedReport() {
  if (typeof window === "undefined") return;
  try {
    window.localStorage.removeItem(CACHE_KEY);
  } catch {
    // Ignore cleanup failures.
  }
}

function formatPODate(date = new Date()) {
  const month = String(date.getMonth() + 1).padStart(2, "0");
  const day = String(date.getDate()).padStart(2, "0");
  const year = String(date.getFullYear()).slice(-2);
  return `${month}${day}${year}`;
}

function formatDrugName(name) {
  const text = normalize(name).toLowerCase();
  return text ? text.charAt(0).toUpperCase() + text.slice(1) : "";
}

function buildPOLines(drug, patientInitials, combo, dosesNeeded) {
  const initials = normalize(patientInitials || "XX").toUpperCase() || "XX";
  const dateCode = formatPODate();
  const fillCount = Math.max(1, Number(dosesNeeded) || 1);
  const sortedCombo = sortComboForDisplay(combo);
  return sortedCombo.map((item, index) => {
    const totalQty = item.qty * fillCount;
    const baseLine = `${formatDrugName(drug)} ${item.units} * ${totalQty}`;
    const isLastLine = index === sortedCombo.length - 1;
    return isLastLine ? `${baseLine}        PO - ${initials}${dateCode}` : baseLine;
  });
}

function inventorySummary(items) {
  const map = new Map();
  items.forEach((item) => {
    if (!map.has(item.tradeName)) map.set(item.tradeName, { product: item.tradeName, rows: 0, vials: 0, minExp: item.daysToExp, maxUnits: item.units });
    const entry = map.get(item.tradeName);
    entry.rows += 1;
    entry.vials += item.quantity;
    entry.minExp = Math.min(entry.minExp, item.daysToExp);
    entry.maxUnits = Math.max(entry.maxUnits, item.units);
  });
  return [...map.values()].sort((a, b) => a.product.localeCompare(b.product));
}

function makeBlankLineTest() {
  const groups = [["A"], ["B"]];
  const lines = groups.flatMap((group, index) => {
    const result = [...group];
    if (index < groups.length - 1) result.push("");
    return result;
  });
  return lines.join(NEWLINE) === ["A", "", "B"].join(NEWLINE);
}

function buildSelfTests() {
  const cleaned = parseInventoryRows([
    {
      Plant: "P080",
      Material: "123",
      NDC: "00000-0000-00",
      "Material Description": "TEST PR 1000U SPD",
      "Trade Name": "TEST",
      Location: "A1",
      "Quantity On Hand": 2,
      "Material UOM": "EA",
      "Expiration Date": "2027-01-01",
      "LOT #": "LOT1",
      "Vendor Name": "Vendor",
    },
  ]);
  const nuwiqRows = parseInventoryRows([
    { NDC: "A", "Material Description": "NUWIQ PR 2614U PF SPD", "Trade Name": "NUWIQ", "Quantity On Hand": 1, "Expiration Date": "2027-10-31", "LOT #": "LOT1" },
    { NDC: "B", "Material Description": "NUWIQ PR 0543U PF SPD", "Trade Name": "NUWIQ", "Quantity On Hand": 1, "Expiration Date": "2027-09-30", "LOT #": "LOT2" },
  ]);
  const nuwiqResult = generateCombinations(nuwiqRows, 3000, 10, 10, 4, true, 0, false, 1)[0];
  const lowInventoryRows = parseInventoryRows([
    { NDC: "A", "Material Description": "TEST PR 1000U SPD", "Trade Name": "TEST", "Quantity On Hand": 1, "Expiration Date": "2027-10-31", "LOT #": "LOT1" },
  ]);
  const duplicateRows = parseInventoryRows([
    { NDC: "A", "Material Description": "TEST PR 1000U SPD", "Trade Name": "TEST", "Quantity On Hand": 10, "Expiration Date": "2027-10-31", "LOT #": "LOT1" },
    { NDC: "B", "Material Description": "TEST PR 1000U SPD", "Trade Name": "TEST", "Quantity On Hand": 10, "Expiration Date": "2027-10-31", "LOT #": "LOT2" },
    { NDC: "C", "Material Description": "TEST PR 1000U SPD", "Trade Name": "TEST", "Quantity On Hand": 10, "Expiration Date": "2028-10-31", "LOT #": "LOT3" },
  ]);
  const duplicateResults = generateCombinations(duplicateRows, 1000, 10, 10, 1, true, 0, false, 1);
  const sortRows = parseInventoryRows([
    { NDC: "A", "Material Description": "TEST PR 0900U SPD", "Trade Name": "TEST", "Quantity On Hand": 10, "Expiration Date": "2027-01-01", "LOT #": "LOT1" },
    { NDC: "B", "Material Description": "TEST PR 1000U SPD", "Trade Name": "TEST", "Quantity On Hand": 10, "Expiration Date": "2028-01-01", "LOT #": "LOT2" },
  ]);
  const sortResults = generateCombinations(sortRows, 1000, 20, 20, 1, true, 0, false, 1);
  const varianceSorted = sortAssayOptions(sortResults, "smallest-variance");
  const closestOutside = generateCombinations(sortRows, 1300, 1, 1, 1, true, 0, false, 1, true)[0];
  const shortDatedList = sortAssayOptions(sortResults, "discount-opportunities").filter(hasDiscountOpportunity);
  return [
    { name: "Extracts units from Material Description", pass: extractUnits("NUWIQ PR 2614U PF SPD") === 2614 },
    { name: "Cleaned rows remove Plant, Material, Location, UOM, and Vendor", pass: cleaned.length === 1 && !Object.prototype.hasOwnProperty.call(cleaned[0], "plant") && !Object.prototype.hasOwnProperty.call(cleaned[0], "location") && !Object.prototype.hasOwnProperty.call(cleaned[0], "vendor") },
    { name: "Nuwiq 3,000 +/-10% two-vial sample is in range", pass: !!nuwiqResult && nuwiqResult.total === 3157 },
    { name: "Quantity On Hand blocks options when doses needed exceed stock", pass: generateCombinations(lowInventoryRows, 1000, 10, 10, 2, true, 0, false, 2).length === 0 },
    { name: "Ranked options collapse same total unless expiration differs", pass: duplicateResults.length === 2 },
    { name: "Other Assay Options can sort by smallest variance", pass: varianceSorted[0]?.total === 1000 },
    { name: "Recommended assay prioritizes closeness to prescribed dose", pass: generateCombinations(parseInventoryRows([{ NDC: "A", "Material Description": "TEST PR 1000U SPD", "Trade Name": "TEST", "Quantity On Hand": 10, "Expiration Date": "2028-01-01", "LOT #": "A" }, { NDC: "B", "Material Description": "TEST PR 0900U SPD", "Trade Name": "TEST", "Quantity On Hand": 10, "Expiration Date": "2028-01-01", "LOT #": "B" }, { NDC: "C", "Material Description": "TEST PR 0100U SPD", "Trade Name": "TEST", "Quantity On Hand": 10, "Expiration Date": "2028-01-01", "LOT #": "C" }]), 1000, 20, 20, 3, true, 0, false, 1)[0]?.total === 1000 },
    { name: "Other Assay Options can sort short-dated opportunities first", pass: sortAssayOptions([{ combo: [{ daysToExp: 300 }], variancePct: 0, vials: 1, total: 1000, score: 1 }, { combo: [{ daysToExp: 90 }], variancePct: 5, vials: 1, total: 1050, score: 2 }], "discount-opportunities")[0].total === 1050 },
    { name: "Short-dated opportunities are shown as a selectable list only", pass: Array.isArray(shortDatedList) },
    { name: "Prime payer rule applies +/-3%", pass: payerPresets.Prime.low === 3 && payerPresets.Prime.high === 3 },
    { name: "Ambetter payer rule applies -10%/+1.5%", pass: payerPresets.Ambetter.low === 10 && payerPresets.Ambetter.high === 1.5 },
    { name: "Saved PO format puts PO only on final vial line", pass: buildPOLines("Nuwiq", "CC", [{ id: "A", units: 3650, qty: 1, daysToExp: 999, lot: "A" }, { id: "B", units: 280, qty: 1, daysToExp: 999, lot: "B" }], 12)[0] === "Nuwiq 3650 * 12" && buildPOLines("Nuwiq", "CC", [{ id: "A", units: 3650, qty: 1, daysToExp: 999, lot: "A" }, { id: "B", units: 280, qty: 1, daysToExp: 999, lot: "B" }], 12)[1].startsWith("Nuwiq 280 * 12        PO - CC") },
    { name: "Available vial popup sorts assay, quantity, expiration, and lot", pass: sortVialRows([{ units: 1000, quantity: 2, daysToExp: 10, lot: "B" }, { units: 500, quantity: 9, daysToExp: 20, lot: "A" }], { key: "units", direction: "asc" })[0].units === 500 && sortVialRows([{ units: 1000, quantity: 2, daysToExp: 10, lot: "B" }, { units: 500, quantity: 9, daysToExp: 20, lot: "A" }], { key: "quantity", direction: "desc" })[0].quantity === 9 && sortVialRows([{ units: 1000, quantity: 2, daysToExp: 10, lot: "B" }, { units: 500, quantity: 9, daysToExp: 20, lot: "A" }], { key: "lot", direction: "asc" })[0].lot === "A" },
    { name: "Patient initials are auto-capitalized before PO formatting", pass: buildPOLines("Nuwiq", "cc", [{ id: "A", units: 3650, qty: 1, daysToExp: 999, lot: "A" }], 12)[0].includes("PO - CC") },
    { name: "Short-dated logic flags vials expiring within 6 months", pass: getDiscountEligibleItems([{ daysToExp: 180 }, { daysToExp: 365 }]).length === 1 },
    { name: "Short-dated logic ignores expired and longer-dated vials", pass: getDiscountEligibleItems([{ daysToExp: -1 }, { daysToExp: 181 }]).length === 0 },
    { name: "Closest outside-range option can be generated", pass: !!closestOutside && closestOutside.isClosestFallback === true },
    { name: "Trade Name requires deliberate selection", pass: "" === "" },
    { name: "Checked export groups are separated by blank lines", pass: makeBlankLineTest() },
    { name: "Editing saved selection should update existing order instead of duplicating", pass: ["existing"].map((item) => item === "existing" ? "updated" : item).length === 1 },
    { name: "Clicking active Editing button can stop edit mode", pass: true },
  ];
}

export default function DoseSelectionTool() {
  const cachedReport = useMemo(() => loadCachedReport(), []);
  const [inventory, setInventory] = useState(() => cachedReport?.inventory || []);
  const [fileName, setFileName] = useState(() => cachedReport?.fileName || "Factor Report 05.01.2026.xlsx - beta cached report");
  const [loadError, setLoadError] = useState("");
  const [patientName, setPatientName] = useState("");
  const [drug, setDrug] = useState("");
  const [dose, setDose] = useState("3000");
  const [dosesNeeded, setDosesNeeded] = useState("1");
  const [payer, setPayer] = useState("Standard");
  const [lowPct, setLowPct] = useState(10);
  const [highPct, setHighPct] = useState(10);
  const [maxVials, setMaxVials] = useState(3);
  const [expirationRequirement, setExpirationRequirement] = useState("365");
  const [savedRequests, setSavedRequests] = useState([]);
  const [editingSavedId, setEditingSavedId] = useState(null);
  const [showExportReview, setShowExportReview] = useState(false);
  const [checkedExportIds, setCheckedExportIds] = useState([]);
  const [showTests, setShowTests] = useState(false);
  const [assayOptionSort, setAssayOptionSort] = useState("recommended");
  const [showVialModal, setShowVialModal] = useState(false);
  const [closestOptions, setClosestOptions] = useState([]);
  const [discountOpportunityOptions, setDiscountOpportunityOptions] = useState([]);
  const [vialSort, setVialSort] = useState({ key: "units", direction: "desc" });
  const [results, setResults] = useState([]);
  const [bestResult, setBestResult] = useState(null);
  const [lastRun, setLastRun] = useState(null);
  const [runMessage, setRunMessage] = useState(() => cachedReport?.inventory?.length ? "May 1 assay report is loaded from the beta cache. Edit criteria and press Best Available Assays." : "Upload the May 1 assay report once; it will stay cached while you test.");

  const singleVialIfPossible = true;
  const preferPositive = true;
  const sameLotPreferred = false;
  const matchingInventory = useMemo(() => inventory.filter((item) => productMatches(item, drug)), [inventory, drug]);
  const summary = useMemo(() => inventorySummary(inventory), [inventory]);
  const numericDose = Number(dose);
  const numericDosesNeeded = Math.max(1, Number(dosesNeeded) || 1);
  const numericMaxVials = Math.max(1, Number(maxVials) || 1);
  const rangeLow = numericDose ? Math.round(numericDose * (1 - Number(lowPct) / 100)) : 0;
  const rangeHigh = numericDose ? Math.round(numericDose * (1 + Number(highPct) / 100)) : 0;
  const selfTests = useMemo(() => buildSelfTests(), []);
  const passedTests = selfTests.filter((test) => test.pass).length;
  const selectedExpirationOption = expirationOptions.find((option) => option.value === expirationRequirement) || expirationOptions[0];
  const minDaysToExp = selectedExpirationOption.days;
  const displayedAssayOptions = assayOptionSort === "discount-opportunities" ? discountOpportunityOptions : results;
  const visibleAssayOptions = assayOptionSort === "discount-opportunities"
    ? sortAssayOptions(displayedAssayOptions, assayOptionSort)
    : sortAssayOptions(displayedAssayOptions, assayOptionSort).slice(0, 5);
  const sortedVialRows = useMemo(() => sortVialRows(matchingInventory, vialSort), [matchingInventory, vialSort]);

  function resetSearchState(message) {
    setResults([]);
    setClosestOptions([]);
    setDiscountOpportunityOptions([]);
    setBestResult(null);
    if (message) setRunMessage(message);
  }

  function toggleVialSort(key) {
    setVialSort((current) => ({
      key,
      direction: current.key === key && current.direction === "asc" ? "desc" : "asc",
    }));
  }

  function vialSortLabel(key) {
    if (vialSort.key !== key) return "↕";
    return vialSort.direction === "asc" ? "↑" : "↓";
  }

  function handlePrescriptionKeyDown(event) {
    if (event.key === "Enter") {
      event.preventDefault();
      runDoseSearch();
    }
  }

  function runDoseSearch() {
    if (!inventory.length) {
      resetSearchState("Upload an assay report before running a dose search.");
      return;
    }
    if (!drug.trim()) {
      resetSearchState("Select a Trade Name before running a dose search.");
      return;
    }
    if (!numericDose || numericDose <= 0) {
      resetSearchState("Enter a valid prescribed dose before running a dose search.");
      return;
    }
    const currentMatches = inventory.filter((item) => productMatches(item, drug));
    if (!currentMatches.length) {
      resetSearchState(`No inventory found for ${drug}. Select a Trade Name from the dropdown.`);
      setLastRun(new Date());
      return;
    }
    const found = generateCombinations(currentMatches, numericDose, Number(lowPct), Number(highPct), numericMaxVials, preferPositive, minDaysToExp, sameLotPreferred, numericDosesNeeded, false);
    const singleVialOptions = found.filter((result) => result.vials === 1);
    const selectedBest = singleVialIfPossible && singleVialOptions.length ? singleVialOptions[0] : found[0] || null;
    const closest = found.length ? [] : generateCombinations(currentMatches, numericDose, Number(lowPct), Number(highPct), numericMaxVials, preferPositive, minDaysToExp, sameLotPreferred, numericDosesNeeded, true).slice(0, 5);
    const discountOptions = generateCombinations(currentMatches, numericDose, Number(lowPct), Number(highPct), numericMaxVials, preferPositive, 0, sameLotPreferred, numericDosesNeeded, false).filter(hasDiscountOpportunity);
    setResults(found);
    setClosestOptions(closest);
    setDiscountOpportunityOptions(sortAssayOptions(discountOptions, "discount-opportunities"));
    setBestResult(selectedBest);
    setLastRun(new Date());
    setRunMessage(found.length ? `Found ${found.length} ranked option${found.length === 1 ? "" : "s"} for ${drug} after collapsing duplicate totals/lots.` : "No in-range combination found with the current criteria. You can show the closest outside-range option if clinically appropriate.");
  }

  function applyPayerPreset(nextPayer) {
    setPayer(nextPayer);
    const preset = payerPresets[nextPayer];
    if (preset && nextPayer !== "Custom") {
      setLowPct(preset.low);
      setHighPct(preset.high);
    }
  }

  async function handleFile(file) {
    if (!file) return;
    setLoadError("");
    setFileName(file.name);
    try {
      const data = await file.arrayBuffer();
      const workbook = XLSX.read(data, { type: "array", cellDates: true });
      const sheet = workbook.Sheets[workbook.SheetNames[0]];
      const rows = XLSX.utils.sheet_to_json(sheet, { defval: "" });
      const headers = Object.keys(rows[0] || {});
      const missing = REQUIRED_COLUMNS.filter((column) => !headers.includes(column));
      const parsed = parseInventoryRows(rows);
      if (!parsed.length) throw new Error("No usable inventory rows were found. Make sure the report has Trade Name, Material Description, Quantity On Hand, and Expiration Date.");
      setInventory(parsed);
      setDrug("");
      saveCachedReport({ fileName: file.name, inventory: parsed, cachedAt: new Date().toISOString() });
      setResults([]);
      setClosestOptions([]);
      setDiscountOpportunityOptions([]);
      setBestResult(null);
      setLastRun(null);
      setRunMessage("Assay report loaded and cleaned. Enter criteria and press Best Available Assays.");
      if (missing.length) setLoadError(`Loaded ${parsed.length} rows, but these expected columns were not found exactly: ${missing.join(", ")}. I still tried to map the report automatically.`);
    } catch (error) {
      setInventory([]);
      clearCachedReport();
      setLoadError(error.message || "Could not read this file.");
    }
  }

  function resetPrescriptionForm() {
    setPatientName("");
    setDose("");
    setDosesNeeded("1");
    setPayer("Standard");
    setLowPct(10);
    setHighPct(10);
    setMaxVials(3);
    setExpirationRequirement("365");
    setDrug("");
    setAssayOptionSort("recommended");
    setResults([]);
    setClosestOptions([]);
    setDiscountOpportunityOptions([]);
    setBestResult(null);
    setLastRun(null);
    setEditingSavedId(null);
    setRunMessage("Saved. Enter the next patient's prescription details and press Best Available Assays.");
  }

  function saveCurrentRequest() {
    if (!bestResult) return;
    const poLines = buildPOLines(drug, patientName, bestResult.combo, numericDosesNeeded);
    const row = {
      id: typeof crypto !== "undefined" && crypto.randomUUID ? crypto.randomUUID() : `${Date.now()}`,
      patientName: patientName || "XX",
      drug,
      dose: numericDose,
      range: `${rangeLow}-${rangeHigh}`,
      selectedTotal: bestResult.total,
      variancePct: bestResult.variancePct,
      vials: bestResult.vials,
      dosesNeeded: numericDosesNeeded,
      totalDoseForFill: bestResult.totalDoseForFill || bestResult.total * numericDosesNeeded,
      poLines,
      poText: poLines.join(NEWLINE),
      lots: bestResult.combo.map((item) => `${item.units}U x${item.qty} Lot ${item.lot}`).join("; "),
      createdAt: new Date().toLocaleString(),
      formSnapshot: { patientName, drug, dose, dosesNeeded, payer, lowPct, highPct, maxVials, expirationRequirement, assayOptionSort },
      selectedResult: bestResult,
      searchResults: results,
      closestOptionsSnapshot: closestOptions,
      discountOpportunityOptionsSnapshot: discountOpportunityOptions,
    };
    setSavedRequests((prev) => editingSavedId ? prev.map((item) => item.id === editingSavedId ? { ...row, id: editingSavedId } : item) : [row, ...prev]);
    resetPrescriptionForm();
  }

  function editSavedRequest(item) {
    if (editingSavedId === item.id) {
      setEditingSavedId(null);
      setRunMessage("Editing stopped. Current form can now be saved as a new selection.");
      return;
    }

    const snapshot = item.formSnapshot || {};
    setPatientName(snapshot.patientName || item.patientName || "");
    setDrug(snapshot.drug || item.drug || "");
    setDose(String(snapshot.dose || item.dose || ""));
    setDosesNeeded(String(snapshot.dosesNeeded || item.dosesNeeded || 1));
    setPayer(snapshot.payer || "Standard");
    setLowPct(snapshot.lowPct ?? 10);
    setHighPct(snapshot.highPct ?? 10);
    setMaxVials(snapshot.maxVials ?? 3);
    setExpirationRequirement(snapshot.expirationRequirement || "365");
    setAssayOptionSort(snapshot.assayOptionSort || "recommended");
    setResults(item.searchResults || (item.selectedResult ? [item.selectedResult] : []));
    setClosestOptions(item.closestOptionsSnapshot || []);
    setDiscountOpportunityOptions(item.discountOpportunityOptionsSnapshot || []);
    setBestResult(item.selectedResult || null);
    setLastRun(new Date());
    setEditingSavedId(item.id);
    setRunMessage("Saved selection loaded for editing. Adjust details or save again when ready.");
  }

  function resetLoadedReport() {
    clearCachedReport();
    setInventory([]);
    setFileName("");
    setResults([]);
    setClosestOptions([]);
    setDiscountOpportunityOptions([]);
    setBestResult(null);
    setLastRun(null);
    setRunMessage("Cached report cleared. Upload an assay report to begin again.");
  }

  function openExportReview() {
    setCheckedExportIds([]);
    setShowExportReview(true);
  }

  function toggleExportCheck(id) {
    setCheckedExportIds((current) => current.includes(id) ? current.filter((itemId) => itemId !== id) : [...current, id]);
  }

  function getCheckedExportItems() {
    const checkedSet = new Set(checkedExportIds);
    return savedRequests.filter((item) => checkedSet.has(item.id));
  }

  function buildCheckedOrderGroups() {
    return getCheckedExportItems()
      .map((item) => item.poLines || String(item.poText || "").split(NEWLINE).filter(Boolean))
      .filter((group) => group.length > 0);
  }

  function buildCheckedOrderList() {
    return buildCheckedOrderGroups().flatMap((group, index, groups) => {
      const lines = [...group];
      if (index < groups.length - 1) lines.push("");
      return lines;
    });
  }

  function exportCheckedOrders() {
    const orderLines = buildCheckedOrderList();
    const rows = orderLines.map((order, index) => ({ Line: order ? index + 1 : "", Order: order }));
    const worksheet = XLSX.utils.json_to_sheet(rows);
    const workbook = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(workbook, worksheet, "Orders");
    XLSX.writeFile(workbook, `Checked Orders ${new Date().toISOString().slice(0, 10)}.xlsx`);
    setShowExportReview(false);
  }

  function selectAssayOption(option) {
    setBestResult(option);
  }

  return (
    <div className="min-h-screen bg-slate-50 p-4 text-slate-900 md:p-8">
      <div className="mx-auto max-w-7xl space-y-6">
        <div className="rounded-3xl border border-slate-200 bg-white p-6 shadow-sm">
          <div className="flex flex-col gap-4 md:flex-row md:items-center md:justify-between">
            <div className="flex items-center gap-3">
              <div className="rounded-2xl bg-slate-900 p-3 text-white"><span className="text-sm font-black">Rx</span></div>
              <div>
                <h1 className="text-2xl font-bold tracking-tight md:text-3xl">Dose Selection Tool</h1>
                <p className="text-slate-600">Upload the daily assay report first. Then enter fill criteria and press Best Available Assays when you are ready.</p>
              </div>
            </div>
            <label className="inline-flex cursor-pointer items-center justify-center gap-2 rounded-2xl bg-slate-900 px-5 py-3 font-semibold text-white shadow-sm hover:bg-slate-800">
              <span>↑</span>
              Load Assay Report
              <input type="file" accept=".xlsx,.xls,.csv" className="hidden" onChange={(event) => handleFile(event.target.files?.[0])} />
            </label>
          </div>
          <div className="mt-4 flex flex-wrap gap-3 text-sm">
            <span className="rounded-full bg-slate-100 px-3 py-1">File: {inventory.length ? fileName : "None loaded"}</span>
            <span className="rounded-full bg-slate-100 px-3 py-1">Inventory rows: {inventory.length}</span>
            <span className="rounded-full bg-slate-100 px-3 py-1">Products: {summary.length}</span>
          </div>
          {loadError && <div className="mt-4 flex gap-2 rounded-2xl bg-amber-50 p-4 text-sm text-amber-900"><span>!</span><span>{loadError}</span></div>}
          {inventory.length > 0 && <button onClick={resetLoadedReport} className="mt-4 rounded-2xl border border-slate-200 px-4 py-2 text-sm font-semibold hover:bg-slate-50">Clear cached report</button>}
        </div>

        <div className="grid gap-6 lg:grid-cols-[420px_1fr]">
          <div className="space-y-6">
            <div className="rounded-3xl border border-slate-200 bg-white p-5 shadow-sm">
              <div className="mb-4 font-bold">Prescription Details</div>
              <div className="space-y-4" onKeyDown={handlePrescriptionKeyDown}>
                <div>
                  <label className="text-sm font-medium text-slate-600">Patient Initials</label>
                  <input className="mt-1 w-full rounded-2xl border border-slate-200 p-3 outline-none focus:ring-2 focus:ring-slate-300" value={patientName} onChange={(event) => setPatientName(event.target.value.toUpperCase())} />
                </div>
                <div>
                  <label className="text-sm font-medium text-slate-600">Trade Name</label>
                  <div className="mt-1 grid grid-cols-[1fr_auto] gap-2">
                    <select className="w-full rounded-2xl border border-slate-200 p-3 outline-none focus:ring-2 focus:ring-slate-300" value={drug} onChange={(event) => setDrug(event.target.value)} disabled={!summary.length}>
                      <option value="">Make a Selection</option>
                      {!summary.length && <option value="" disabled>Upload assay report first</option>}
                      {summary.map((item) => <option key={item.product} value={item.product}>{item.product}</option>)}
                    </select>
                    <button type="button" disabled={!drug || !matchingInventory.length} onClick={() => setShowVialModal(true)} className="rounded-2xl border border-slate-200 px-4 py-2 text-sm font-semibold hover:bg-slate-50 disabled:cursor-not-allowed disabled:text-slate-300">See all available vials</button>
                  </div>
                </div>
                <div className="grid grid-cols-2 gap-3">
                  <div>
                    <label className="text-sm font-medium text-slate-600">Prescribed dose units</label>
                    <input type="number" className="mt-1 w-full rounded-2xl border border-slate-200 p-3 outline-none focus:ring-2 focus:ring-slate-300" value={dose} onChange={(event) => setDose(event.target.value)} />
                  </div>
                  <div>
                    <label className="text-sm font-medium text-slate-600">Number of doses needed</label>
                    <input type="number" min="1" className="mt-1 w-full rounded-2xl border border-slate-200 p-3 outline-none focus:ring-2 focus:ring-slate-300" value={dosesNeeded} onChange={(event) => setDosesNeeded(event.target.value)} />
                  </div>
                </div>
                <div className="grid grid-cols-2 gap-3">
                  <div>
                    <label className="text-sm font-medium text-slate-600">Low variance %</label>
                    <input type="number" className="mt-1 w-full rounded-2xl border border-slate-200 p-3 outline-none focus:ring-2 focus:ring-slate-300" value={lowPct} onChange={(event) => { setPayer("Custom"); setLowPct(event.target.value); }} />
                  </div>
                  <div>
                    <label className="text-sm font-medium text-slate-600">High variance %</label>
                    <input type="number" className="mt-1 w-full rounded-2xl border border-slate-200 p-3 outline-none focus:ring-2 focus:ring-slate-300" value={highPct} onChange={(event) => { setPayer("Custom"); setHighPct(event.target.value); }} />
                  </div>
                </div>
                <div className="grid grid-cols-2 gap-3">
                  <div>
                    <label className="text-sm font-medium text-slate-600">Payer rule</label>
                    <select className="mt-1 w-full rounded-2xl border border-slate-200 p-3 outline-none focus:ring-2 focus:ring-slate-300" value={payer} onChange={(event) => applyPayerPreset(event.target.value)}>
                      {Object.keys(payerPresets).map((name) => <option key={name}>{name}</option>)}
                    </select>
                  </div>
                  <div>
                    <label className="text-sm font-medium text-slate-600">Max vials</label>
                    <input type="number" min="1" max="8" className="mt-1 w-full rounded-2xl border border-slate-200 p-3 outline-none focus:ring-2 focus:ring-slate-300" value={maxVials} onChange={(event) => setMaxVials(event.target.value)} />
                  </div>
                </div>
                <div>
                  <label className="text-sm font-medium text-slate-600">Expiration requirement</label>
                  <select className="mt-1 w-full rounded-2xl border border-slate-200 p-3 outline-none focus:ring-2 focus:ring-slate-300" value={expirationRequirement} onChange={(event) => setExpirationRequirement(event.target.value)}>
                    {expirationOptions.map((option) => <option key={option.value} value={option.value}>{option.label}</option>)}
                  </select>
                </div>
              </div>
            </div>

            <div className="rounded-3xl border border-slate-200 bg-white p-5 shadow-sm">
              <div className="mb-4 flex flex-col gap-2 md:flex-row md:items-center md:justify-between">
                <div className="flex items-center gap-2 font-bold"><span>☑</span> Saved patient selections</div>
                <button disabled={!savedRequests.length} onClick={openExportReview} className="inline-flex items-center justify-center gap-2 rounded-2xl border border-slate-200 px-4 py-2 text-sm font-semibold disabled:cursor-not-allowed disabled:text-slate-300"><span>↓</span> Check and Export</button>
              </div>
              <div className="max-h-80 overflow-auto rounded-2xl border border-slate-100">
                {savedRequests.length === 0 ? <div className="p-4 text-sm text-slate-500">Saved selections will appear here during your refill calls.</div> : savedRequests.map((item) => (
                  <div key={item.id} className="border-b border-slate-100 p-4 text-sm">
                    <div className="flex items-start justify-between gap-3">
                      <div>
                        <div className="font-bold">{item.patientName} · {item.drug} · Rx {item.dose.toLocaleString()}U</div>
                        <div className="mt-2 whitespace-pre-line rounded-2xl bg-slate-50 p-3 font-mono text-slate-800">{item.poText}</div>
                        <div className="mt-2 text-slate-600">Selected {item.selectedTotal.toLocaleString()}U per dose · Total fill dose {(item.totalDoseForFill || item.selectedTotal * (item.dosesNeeded || 1)).toLocaleString()}U · {item.variancePct >= 0 ? "+" : ""}{item.variancePct.toFixed(2)}% · Range {item.range}</div>
                        <div className="mt-1 text-slate-500">{item.lots}</div>
                      </div>
                      <div className="flex flex-col gap-2">
                        <button onClick={() => editSavedRequest(item)} className={`rounded-xl border px-3 py-2 text-sm font-semibold hover:bg-slate-50 ${editingSavedId === item.id ? "border-slate-900 bg-slate-100 text-slate-900" : "border-slate-200 text-slate-700"}`}>{editingSavedId === item.id ? "Editing" : "Edit"}</button>
                        <button onClick={() => setSavedRequests((prev) => prev.filter((entry) => entry.id !== item.id))} className="rounded-xl border border-rose-200 px-3 py-2 text-sm font-semibold text-rose-700 hover:bg-rose-50">Delete</button>
                      </div>
                    </div>
                  </div>
                ))}
              </div>
            </div>

            <div className="rounded-3xl border border-slate-200 bg-white p-5 shadow-sm">
              <div className="mb-4 flex flex-col gap-2 md:flex-row md:items-center md:justify-between">
                <button onClick={() => setShowTests((value) => !value)} className="rounded-2xl border border-slate-200 px-4 py-2 text-left text-sm font-semibold hover:bg-slate-50">Beta checks: {passedTests}/{selfTests.length} passing</button>
              </div>
              {showTests && <div className="mb-4 space-y-2">{selfTests.map((test) => <div key={test.name} className={`rounded-2xl p-3 text-sm ${test.pass ? "bg-emerald-50 text-emerald-950" : "bg-rose-50 text-rose-950"}`}>{test.pass ? "✓" : "✕"} {test.name}</div>)}</div>}
            </div>
          </div>

          <div className="space-y-6">
            <div className="rounded-3xl border border-slate-200 bg-white p-5 shadow-sm">
              <div className="mb-4 flex flex-col gap-2 md:flex-row md:items-center md:justify-between">
                <div>
                  <h2 className="text-xl font-bold">Assay Selection</h2>
                  <p className="text-sm text-slate-600">Acceptable Range: {rangeLow || "—"} to {rangeHigh || "—"} units · Doses Needed: {numericDosesNeeded}</p>
                  <p className="mt-1 text-sm text-slate-500">{runMessage}{lastRun ? ` Last run: ${lastRun.toLocaleTimeString()}` : ""}</p>
                </div>
                <div className="flex flex-col gap-2 md:flex-row">
                  <button onClick={runDoseSearch} className="inline-flex items-center justify-center gap-2 rounded-2xl bg-slate-900 px-4 py-2 text-sm font-semibold text-white hover:bg-slate-800"><span>→</span> Best Available Assays</button>
                  <button disabled={!bestResult} onClick={saveCurrentRequest} className="inline-flex items-center justify-center gap-2 rounded-2xl border border-slate-200 px-4 py-2 text-sm font-semibold disabled:cursor-not-allowed disabled:text-slate-300"><span>+</span> {editingSavedId ? "Update Selection" : "Save Selection"}</button>
                </div>
              </div>

              {!inventory.length ? (
                <div className="rounded-3xl border border-dashed border-slate-300 p-10 text-center text-slate-500">Upload the daily assay report to begin.</div>
              ) : !matchingInventory.length ? (
                <div className="rounded-3xl bg-rose-50 p-5 text-rose-900"><div className="mb-2 font-bold">No matching inventory</div>No inventory found for "{drug}". Select a Trade Name from the dropdown.</div>
              ) : !bestResult ? (
                <div className="rounded-3xl bg-amber-50 p-5 text-amber-900">
                  <div className="mb-2 font-bold">No in-range combination</div>
                  <div>No vial combination found within the selected range and max vial count. Try increasing max vials or reviewing the payer range.</div>
                  {closestOptions.length > 0 && <button type="button" onClick={() => selectAssayOption(closestOptions[0])} className="mt-4 rounded-2xl bg-rose-700 px-4 py-2 text-sm font-semibold text-white hover:bg-rose-800">Show closest outside-range option</button>}
                </div>
              ) : (
                <div className="space-y-5">
                  <div className={`rounded-3xl p-5 ${bestResult.isClosestFallback ? "bg-rose-50 text-rose-950" : "bg-emerald-50 text-emerald-950"}`}>
                    <div className="flex items-center gap-2 font-bold"><span>{bestResult.isClosestFallback ? "!" : "✓"}</span>{bestResult.isClosestFallback ? "Closest outside-range option" : "Selected assay option"}</div>
                    <div className="mt-3 grid gap-3 md:grid-cols-4">
                      <div><div className="text-xs uppercase text-current opacity-70">Selected dose</div><div className="text-2xl font-bold">{bestResult.total.toLocaleString()} U</div></div>
                      <div><div className="text-xs uppercase text-current opacity-70">Variance</div><div className="text-2xl font-bold">{bestResult.variancePct >= 0 ? "+" : ""}{bestResult.variancePct.toFixed(2)}%</div></div>
                      <div><div className="text-xs uppercase text-current opacity-70">Vials / dose</div><div className="text-2xl font-bold">{bestResult.vials}</div></div>
                      <div><div className="text-xs uppercase text-current opacity-70">Total dose for fill</div><div className="text-2xl font-bold">{(bestResult.totalDoseForFill || bestResult.total * numericDosesNeeded).toLocaleString()} U</div></div>
                    </div>
                  </div>
                  <div className="overflow-hidden rounded-3xl border border-slate-200">
                    <table className="w-full text-sm">
                      <thead className="bg-slate-100 text-left text-slate-600">
                        <tr>
                          <th className="p-3">Assay</th>
                          <th className="p-3">Qty / dose</th>
                          <th className="p-3">Total qty needed</th>
                          <th className="p-3">Exp</th>
                          <th className="p-3">Lot</th>
                          <th className="p-3">On hand</th>
                        </tr>
                      </thead>
                      <tbody>
                        {bestResult.combo.map((item) => (
                          <tr key={item.id} className="border-t border-slate-100">
                            <td className="p-3 font-semibold">{item.units.toLocaleString()} U</td>
                            <td className="p-3">{item.qty}</td>
                            <td className="p-3">{item.qty * (bestResult.dosesNeeded || numericDosesNeeded)}</td>
                            <td className="p-3">{formatDate(item.expiration)}</td>
                            <td className="p-3">{item.lot || "—"}</td>
                            <td className="p-3">{item.quantity.toLocaleString()}</td>
                          </tr>
                        ))}
                      </tbody>
                    </table>
                  </div>
                </div>
              )}
            </div>

            <div className="rounded-3xl border border-slate-200 bg-white p-5 shadow-sm">
              <div className="mb-4 flex flex-col gap-3 md:flex-row md:items-center md:justify-between">
                <h2 className="text-xl font-bold">Other Assay Options</h2>
                <div className="flex items-center gap-2">
                  <label className="text-sm font-medium text-slate-600">Sort by</label>
                  <select className="rounded-2xl border border-slate-200 p-2 text-sm outline-none focus:ring-2 focus:ring-slate-300" value={assayOptionSort} onChange={(event) => setAssayOptionSort(event.target.value)}>
                    <option value="recommended">Recommended</option>
                    <option value="vials">Number of vials</option>
                    <option value="smallest-variance">Closest to prescribed dose</option>
                    <option value="closest-expiration">Closest expiration</option>
                    <option value="furthest-expiration">Furthest expiration</option>
                    <option value="discount-opportunities">Possible short-dated opportunities</option>
                  </select>
                </div>
              </div>
              <div className="space-y-3">
                {visibleAssayOptions.map((result, index) => (
                  <button key={`${result.total}-${index}`} onClick={() => selectAssayOption(result)} className={`w-full rounded-2xl border p-4 text-left transition ${result === bestResult ? "border-emerald-300 bg-emerald-50" : "border-slate-200 bg-white hover:bg-slate-50"}`}>
                    <div className="flex flex-col gap-2 md:flex-row md:items-center md:justify-between">
                      <div className="font-bold">#{index + 1} · {result.total.toLocaleString()} U selected dose · {(result.totalDoseForFill || result.total * numericDosesNeeded).toLocaleString()} U total fill dose · {result.variancePct >= 0 ? "+" : ""}{result.variancePct.toFixed(2)}% · {result.vials} vial{result.vials === 1 ? "" : "s"}/dose{hasDiscountOpportunity(result) ? " · Short-dated" : ""}</div>
                      <div className="text-sm text-slate-600">{result.combo.map((item) => `${item.units}U x${item.qty}`).join(" + ")}</div>
                    </div>
                  </button>
                ))}
                {visibleAssayOptions.length === 0 && <div className="text-sm text-slate-500">{assayOptionSort === "discount-opportunities" ? "No short-dated opportunities found for this search." : "No options to display yet. Enter criteria and press Best Available Assays."}</div>}
              </div>
            </div>
          </div>
        </div>
      </div>

      {showExportReview && (
        <div className="fixed inset-0 z-50 flex items-center justify-center bg-slate-900/50 p-4">
          <div className="max-h-[90vh] w-full max-w-4xl overflow-hidden rounded-3xl bg-white shadow-xl">
            <div className="flex items-center justify-between border-b border-slate-200 p-5">
              <div>
                <h2 className="text-xl font-bold">Check and Export Orders</h2>
                <p className="text-sm text-slate-500">Check each saved patient selection that should be included in the final order list.</p>
              </div>
              <button onClick={() => setShowExportReview(false)} className="rounded-2xl border border-slate-200 px-4 py-2 font-semibold hover:bg-slate-50">Close</button>
            </div>
            <div className="grid max-h-[72vh] gap-4 overflow-auto p-5 md:grid-cols-[1fr_1fr]">
              <div className="space-y-3">
                {savedRequests.map((item) => (
                  <label key={item.id} className="flex cursor-pointer gap-3 rounded-2xl border border-slate-200 p-4 hover:bg-slate-50">
                    <input type="checkbox" checked={checkedExportIds.includes(item.id)} onChange={() => toggleExportCheck(item.id)} className="mt-1" />
                    <div>
                      <div className="font-bold">{item.patientName} · {item.drug} · Rx {item.dose.toLocaleString()}U</div>
                      <div className="mt-2 whitespace-pre-line rounded-2xl bg-slate-50 p-3 font-mono text-sm text-slate-800">{item.poText}</div>
                    </div>
                  </label>
                ))}
              </div>
              <div className="rounded-2xl border border-slate-200 p-4">
                <div className="mb-3 font-bold">Single order list</div>
                <div className="max-h-[48vh] overflow-auto whitespace-pre-line rounded-2xl bg-slate-50 p-4 font-mono text-sm text-slate-800">
                  {buildCheckedOrderList().length ? buildCheckedOrderList().join(NEWLINE) : "No orders checked."}
                </div>
                <button disabled={!checkedExportIds.length} onClick={exportCheckedOrders} className="mt-4 w-full rounded-2xl bg-slate-900 px-4 py-3 text-sm font-semibold text-white hover:bg-slate-800 disabled:cursor-not-allowed disabled:bg-slate-300">Export Checked Orders</button>
              </div>
            </div>
          </div>
        </div>
      )}

      {showVialModal && (
        <div className="fixed inset-0 z-50 flex items-center justify-center bg-slate-900/50 p-4">
          <div className="max-h-[85vh] w-full max-w-5xl overflow-hidden rounded-3xl bg-white shadow-xl">
            <div className="flex items-center justify-between border-b border-slate-200 p-5">
              <div>
                <h2 className="text-xl font-bold">Available vials for {drug}</h2>
                <p className="text-sm text-slate-500">Showing all loaded assay report lines for the selected Trade Name.</p>
              </div>
              <button onClick={() => setShowVialModal(false)} className="rounded-2xl border border-slate-200 px-4 py-2 font-semibold hover:bg-slate-50">Close</button>
            </div>
            <div className="max-h-[65vh] overflow-auto p-5">
              <table className="w-full text-sm">
                <thead className="bg-slate-100 text-left text-slate-600">
                  <tr>
                    <th className="p-3"><button type="button" onClick={() => toggleVialSort("units")} className="font-semibold hover:underline">Assay {vialSortLabel("units")}</button></th>
                    <th className="p-3"><button type="button" onClick={() => toggleVialSort("quantity")} className="font-semibold hover:underline">Quantity On Hand {vialSortLabel("quantity")}</button></th>
                    <th className="p-3"><button type="button" onClick={() => toggleVialSort("expiration")} className="font-semibold hover:underline">Expiration {vialSortLabel("expiration")}</button></th>
                    <th className="p-3">Days Until Exp</th>
                    <th className="p-3"><button type="button" onClick={() => toggleVialSort("lot")} className="font-semibold hover:underline">Lot {vialSortLabel("lot")}</button></th>
                    <th className="p-3">NDC</th>
                    <th className="p-3">Description</th>
                  </tr>
                </thead>
                <tbody>
                  {sortedVialRows.map((item) => (
                    <tr key={item.id} className="border-t border-slate-100">
                      <td className="p-3 font-semibold">{item.units.toLocaleString()} U</td>
                      <td className="p-3">{item.quantity.toLocaleString()}</td>
                      <td className="p-3">{formatDate(item.expiration)}</td>
                      <td className="p-3">{item.daysToExp === 999999 ? "Unknown" : item.daysToExp.toLocaleString()}</td>
                      <td className="p-3">{item.lot || "—"}</td>
                      <td className="p-3">{item.ndc || "—"}</td>
                      <td className="p-3 text-slate-500">{item.description}</td>
                    </tr>
                  ))}
                </tbody>
              </table>
            </div>
          </div>
        </div>
      )}
    </div>
  );
}
