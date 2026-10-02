const CBM_DIVISOR = 1000000;

const containers = [
  { id: "20FT", name: "20 FT", capacity: 33 },
  { id: "40FT", name: "40 FT", capacity: 67 },
  { id: "40HC", name: "40 FT High Cube", capacity: 76 }
];

const excelFile = document.getElementById("excelFile");
const fileName = document.getElementById("fileName");
const dataBody = document.getElementById("dataBody");
const packCount = document.getElementById("packCount");
const packetQuantityTotal = document.getElementById("packetQuantityTotal");
const totalCbmEl = document.getElementById("totalCbm");
const recommendedSmall = document.getElementById("recommendedSmall");
const recommendation = document.getElementById("recommendation");
const statusBadge = document.getElementById("statusBadge");
const detectionPanel = document.getElementById("detectionPanel");
const detectionStatus = document.getElementById("detectionStatus");
const detectedSheet = document.getElementById("detectedSheet");
const detectedHeaderRow = document.getElementById("detectedHeaderRow");
const detectedCbmColumn = document.getElementById("detectedCbmColumn");
const detectedQuantityColumn = document.getElementById("detectedQuantityColumn");
const detectedRecordCount = document.getElementById("detectedRecordCount");
const selectionControls = document.getElementById("selectionControls");
const sheetSelect = document.getElementById("sheetSelect");
const cbmSelect = document.getElementById("cbmSelect");
const defaultLength = document.getElementById("defaultLength");
const defaultWidth = document.getElementById("defaultWidth");
const defaultHeight = document.getElementById("defaultHeight");
const freelanLogo = document.getElementById("freelanLogo");

const state = {
  candidates: [],
  selectedIndex: -1,
  rows: []
};

excelFile.addEventListener("change", handleFile);
sheetSelect.addEventListener("change", handleSheetChange);
cbmSelect.addEventListener("change", event => selectCandidate(Number(event.target.value)));
document.getElementById("clearData").addEventListener("click", clearData);
document.getElementById("applyDimensionsEmpty").addEventListener("click", () => applyDefaultDimensions(false));
document.getElementById("applyDimensionsAll").addEventListener("click", () => applyDefaultDimensions(true));
document.getElementById("downloadPdf").addEventListener("click", () => downloadPdf());

document.getElementById("loadDemo").addEventListener("click", () => {
  const demoCandidate = createDemoCandidate();
  fileName.textContent = "Demo data loaded";
  state.candidates = [demoCandidate];
  state.selectedIndex = 0;
  renderDetection(demoCandidate);
  processCandidate(demoCandidate);
});

async function handleFile(event) {
  const file = event.target.files[0];
  if (!file) return;

  fileName.textContent = file.name;

  try {
    const buffer = await file.arrayBuffer();
    const workbook = XLSX.read(buffer, { type: "array" });
    const candidates = detectCandidates(workbook);
    state.candidates = candidates;

    if (!candidates.length) {
      detectionPanel.hidden = true;
      showError("Packing table could not be detected. Please upload an Excel file with item, carton, and Length / Width / Height columns.");
      return;
    }

    state.selectedIndex = findBestCandidateIndex(candidates);
    populateSelectionControls();
    selectCandidate(state.selectedIndex);
  } catch (error) {
    console.error(error);
    detectionPanel.hidden = true;
    showError("Could not read the Excel file. Please check the file format.");
  }
}

function detectCandidates(workbook) {
  const candidates = [];

  workbook.SheetNames.forEach((sheetName, sheetIndex) => {
    const sheet = workbook.Sheets[sheetName];
    const rows = XLSX.utils.sheet_to_json(sheet, {
      header: 1,
      defval: "",
      blankrows: true,
      raw: true
    });
    const range = sheet["!ref"] ? XLSX.utils.decode_range(sheet["!ref"]) : { s: { r: 0 } };
    const rowOffset = range.s.r || 0;
    const scanLimit = Math.min(rows.length, 150);

    for (let rowIndex = 0; rowIndex < scanLimit; rowIndex += 1) {
      const row = rows[rowIndex] || [];
      const headerInfo = inspectHeader(row);
      const hasDimensions = headerInfo.lengthIndex !== -1 && headerInfo.widthIndex !== -1 && headerInfo.heightIndex !== -1;
      if (!headerInfo.cbmColumns.length && !hasDimensions && headerInfo.quantityIndex === -1) continue;

      const recognizedCount = headerInfo.cbmColumns.length
        + (hasDimensions ? 3 : 0)
        + (headerInfo.quantityIndex === -1 ? 0 : 1)
        + (headerInfo.descriptionIndex === -1 ? 0 : 1)
        + (headerInfo.packetQuantityIndex === -1 ? 0 : 1);

      const sources = headerInfo.cbmColumns.length
        ? headerInfo.cbmColumns
        : [{ index: -1, type: "calculated" }];

      sources.forEach(cbmColumn => {
        const candidate = {
          sheetName,
          sheetIndex,
          rows,
          headerIndex: rowIndex,
          headerRow: rowOffset + rowIndex + 1,
          cbmColumnIndex: cbmColumn.index,
          cbmHeader: cbmColumn.index === -1
            ? "Calculated CBM"
            : displayHeader(row[cbmColumn.index], `Column ${cbmColumn.index + 1}`),
          cbmType: hasDimensions ? "calculated" : cbmColumn.type,
          quantityColumnIndex: headerInfo.quantityIndex,
          quantityHeader: headerInfo.quantityIndex === -1 ? "" : displayHeader(row[headerInfo.quantityIndex], "No. of Ctns"),
          packetQuantityColumnIndex: headerInfo.packetQuantityIndex,
          packetsPerBoxColumnIndex: headerInfo.packetsPerBoxIndex,
          netWeightColumnIndex: headerInfo.netWeightIndex,
          grossWeightColumnIndex: headerInfo.grossWeightIndex,
          batchIndex: headerInfo.batchIndex,
          descriptionColumnIndex: headerInfo.descriptionIndex,
          lengthColumnIndex: headerInfo.lengthIndex,
          widthColumnIndex: headerInfo.widthIndex,
          heightColumnIndex: headerInfo.heightIndex,
          unitColumnIndex: headerInfo.unitColumnIndex,
          hasDimensions,
          recognizedCount
        };
        candidate.previewRows = extractRows(candidate);
        candidate.validCount = candidate.previewRows.length;

        if (recognizedCount >= 2 || (candidate.validCount > 0 && (hasDimensions || headerInfo.cbmColumns.length))) {
          candidates.push(candidate);
        }
      });
    }
  });

  return candidates;
}

function inspectHeader(row) {
  const cbmColumns = [];
  let quantityIndex = -1;
  let quantityScore = -1;
  let packetQuantityIndex = -1;
  let packetsPerBoxIndex = -1;
  let netWeightIndex = -1;
  let grossWeightIndex = -1;
  let batchIndex = -1;
  let descriptionIndex = -1;
  let descriptionScore = -1;
  let lengthIndex = -1;
  let widthIndex = -1;
  let heightIndex = -1;
  let unitColumnIndex = -1;

  row.forEach((cell, index) => {
    const header = String(cell ?? "").trim();
    const cbmType = classifyCbmHeader(header);
    if (cbmType) {
      cbmColumns.push({ index, type: cbmType });
      if (cbmType === "unit" && unitColumnIndex === -1) unitColumnIndex = index;
    }

    const currentQuantityScore = cartonQuantityScore(header);
    if (currentQuantityScore > quantityScore) {
      quantityIndex = index;
      quantityScore = currentQuantityScore;
    }
    if (packetQuantityIndex === -1 && isPacketQuantityHeader(header)) packetQuantityIndex = index;
    if (packetsPerBoxIndex === -1 && isPacketsPerBoxHeader(header)) packetsPerBoxIndex = index;
    if (netWeightIndex === -1 && isNetWeightHeader(header)) netWeightIndex = index;
    if (grossWeightIndex === -1 && isGrossWeightHeader(header)) grossWeightIndex = index;
    if (batchIndex === -1 && isBatchHeader(header)) batchIndex = index;
    if (lengthIndex === -1 && isLengthHeader(header)) lengthIndex = index;
    if (widthIndex === -1 && isWidthHeader(header)) widthIndex = index;
    if (heightIndex === -1 && isHeightHeader(header)) heightIndex = index;

    const currentDescriptionScore = descriptionHeaderScore(header);
    if (currentDescriptionScore > descriptionScore) {
      descriptionIndex = index;
      descriptionScore = currentDescriptionScore;
    }
  });

  return {
    cbmColumns,
    quantityIndex,
    packetQuantityIndex,
    packetsPerBoxIndex,
    netWeightIndex,
    grossWeightIndex,
    batchIndex,
    descriptionIndex,
    lengthIndex,
    widthIndex,
    heightIndex,
    unitColumnIndex
  };
}

function classifyCbmHeader(value) {
  const header = normalize(value);
  const hasCbmName = header.includes("cbm") || header.includes("cubicmeter") || header.includes("cubicmetre") || header.includes("volume");
  if (!hasCbmName) return null;
  if (/total|overall|shipment|grand|line/.test(header)) return "total";
  if (/unit|carton|box|piece|each|per/.test(header)) return "unit";
  return "generic";
}

function cartonQuantityScore(value) {
  const header = normalize(value);
  if (!header || isPacketQuantityHeader(value) || isPacketsPerBoxHeader(value)) return -1;
  if (header.includes("noofctn") || header.includes("noofcarton") || header.includes("numberofcarton") || header === "ctns" || header === "ctn" || header === "cartons") return 12;
  if (header.includes("cartonqty") || header.includes("carton") || header.includes("ctn")) return 10;
  if (header.includes("packqty") || header === "packs" || header === "boxes") return 8;
  if (header === "qty" || header === "quantity") return 6;
  if (header.includes("quantity") && !header.includes("pkt") && !header.includes("packet") && !header.includes("bottle") && !header.includes("weight")) return 4;
  return -1;
}

function isPacketQuantityHeader(value) {
  const header = normalize(value);
  return header.includes("quantitypkt")
    || header.includes("pktbottle")
    || header.includes("packetquantity")
    || header.includes("quantitypacket")
    || header.includes("qtypkt")
    || header.includes("qtypacket")
    || header.includes("noofpkt")
    || header.includes("noofpacket")
    || header === "packets"
    || header === "pkts"
    || header.includes("ballqty")
    || header.includes("quantityball")
    || header.includes("noofball");
}

function isPacketsPerBoxHeader(value) {
  const header = normalize(value);
  return header.includes("pktperbox")
    || header.includes("packetperbox")
    || header.includes("packetsperbox")
    || header.includes("packperbox")
    || header.includes("piecesperbox")
    || header.includes("qtyperbox")
    || header.includes("unitspercarton")
    || header.includes("pktperctn")
    || header.includes("packetperctn")
    || header.includes("ballperbox")
    || header.includes("ballsperbox");
}

function isNetWeightHeader(value) {
  const header = normalize(value);
  if (isGrossWeightHeader(value)) return false;
  return header.includes("netweight") || header === "netwt" || header === "nw" || header === "nwt" || header === "weight" || header === "wt" || header === "weightkg" || header === "wtkg";
}

function isGrossWeightHeader(value) {
  const header = normalize(value);
  return header.includes("grossweight") || header === "grosswt" || header === "gw" || header === "gwt" || header.includes("grosswt");
}

function isBatchHeader(value) {
  const header = normalize(value);
  return header.includes("batch") || header.includes("lotno") || header === "lot";
}

function isLengthHeader(value) {
  const header = normalize(value);
  if (!header || header.includes("cbm")) return false;
  return header === "l"
    || header === "lcm"
    || header.includes("length")
    || header.includes("lenght")
    || header === "long";
}

function isWidthHeader(value) {
  const header = normalize(value);
  if (!header || header.includes("cbm")) return false;
  return header === "w"
    || header === "wcm"
    || header.includes("width")
    || header.includes("wdth")
    || header === "wide";
}

function isHeightHeader(value) {
  const header = normalize(value);
  if (!header || header.includes("cbm") || header.includes("highcube")) return false;
  return header === "h"
    || header === "hcm"
    || header.includes("height")
    || header.includes("heigth")
    || header.includes("hight")
    || header === "depth";
}

function isDescriptionHeader(value) {
  const header = normalize(value);
  if (!header) return false;
  if (isPacketQuantityHeader(value) || isPacketsPerBoxHeader(value) || cartonQuantityScore(value) > 0) return false;
  if (isLengthHeader(value) || isWidthHeader(value) || isHeightHeader(value) || classifyCbmHeader(value)) return false;
  return header.includes("description") || header.includes("product") || header.includes("item") || header.includes("article") || header === "pack" || header === "packs" || header.includes("stylename") || header === "style";
}

function descriptionHeaderScore(value) {
  const header = normalize(value);
  if (!isDescriptionHeader(header)) return -1;
  if (header.includes("productdescription")) return 10;
  if (header.includes("description")) return 8;
  if (header.includes("product")) return 7;
  if (header.includes("item")) return 5;
  return 3;
}

function findBestCandidateIndex(candidates) {
  return candidates.reduce((bestIndex, candidate, index) => {
    const best = candidates[bestIndex];
    return candidateScore(candidate) > candidateScore(best) ? index : bestIndex;
  }, 0);
}

function candidateScore(candidate) {
  const typeScore = candidate.cbmType === "calculated" ? 16 : candidate.cbmType === "total" ? 12 : candidate.cbmType === "unit" ? 5 : 1;
  const dimensionScore = candidate.hasDimensions ? 20 : 0;
  return candidate.validCount * 10 + candidate.recognizedCount * 2 + typeScore + dimensionScore + (candidate.quantityColumnIndex === -1 ? 0 : 2);
}

function populateSelectionControls() {
  const uniqueSheets = [...new Map(state.candidates.map(candidate => [candidate.sheetName, candidate])).values()];
  sheetSelect.innerHTML = uniqueSheets.map(candidate => `<option value="${candidate.sheetIndex}">${escapeHtml(candidate.sheetName)}</option>`).join("");
  sheetSelect.value = String(state.candidates[state.selectedIndex].sheetIndex);
  populateCbmOptions(Number(sheetSelect.value));
  selectionControls.hidden = state.candidates.length <= 1;
}

function populateCbmOptions(sheetIndex) {
  const sheetCandidates = state.candidates
    .map((candidate, index) => ({ candidate, index }))
    .filter(item => item.candidate.sheetIndex === sheetIndex);
  cbmSelect.innerHTML = sheetCandidates.map(item => {
    const candidate = item.candidate;
    const label = `${candidate.cbmHeader} | header row ${candidate.headerRow} | ${candidate.validCount} valid rows`;
    return `<option value="${item.index}">${escapeHtml(label)}</option>`;
  }).join("");
  const selectedForSheet = sheetCandidates.find(item => item.index === state.selectedIndex) || sheetCandidates[0];
  if (selectedForSheet) cbmSelect.value = String(selectedForSheet.index);
}

function handleSheetChange(event) {
  const sheetIndex = Number(event.target.value);
  const sheetCandidates = state.candidates
    .map((candidate, index) => ({ candidate, index }))
    .filter(item => item.candidate.sheetIndex === sheetIndex);
  if (!sheetCandidates.length) return;
  const best = sheetCandidates.reduce((current, item) => candidateScore(item.candidate) > candidateScore(current.candidate) ? item : current, sheetCandidates[0]);
  state.selectedIndex = best.index;
  populateCbmOptions(sheetIndex);
  selectCandidate(best.index);
}

function selectCandidate(index) {
  const candidate = state.candidates[index];
  if (!candidate) return;
  state.selectedIndex = index;
  sheetSelect.value = String(candidate.sheetIndex);
  populateCbmOptions(candidate.sheetIndex);
  cbmSelect.value = String(index);
  renderDetection(candidate);
  processCandidate(candidate);
}

function renderDetection(candidate) {
  detectionPanel.hidden = false;
  detectionStatus.textContent = state.candidates.length > 1 ? "Choose data source" : "Automatic detection";
  detectedSheet.textContent = candidate.sheetName;
  detectedHeaderRow.textContent = candidate.headerRow;
  detectedCbmColumn.textContent = candidate.hasDimensions ? "CBM = (L × W × H × Ctns) ÷ 1,000,000" : candidate.cbmHeader;
  detectedQuantityColumn.textContent = candidate.quantityHeader || "Not detected (using 1)";
  detectedRecordCount.textContent = candidate.validCount;
  selectionControls.hidden = state.candidates.length <= 1;
}

function processCandidate(candidate) {
  const rows = extractRows(candidate);
  candidate.validCount = rows.length;
  detectedRecordCount.textContent = rows.length;
  state.rows = rows;

  if (!rows.length) {
    showError("A packing table was detected, but no valid item rows were found.");
    return;
  }

  renderTable(rows);
  calculateRecommendation(rows);
  setStatus("Data loaded successfully", "#eaf7ef", "#207044");
}

function extractRows(candidate) {
  const processed = [];
  const dataRows = candidate.rows.slice(candidate.headerIndex + 1);

  dataRows.forEach((row, index) => {
    if (!Array.isArray(row) || isSummaryRow(row) || isRepeatedHeaderRow(row, candidate)) return;

    const description = candidate.descriptionColumnIndex === -1
      ? ""
      : String(row[candidate.descriptionColumnIndex] ?? "").trim();

    if (candidate.descriptionColumnIndex !== -1 && !description) return;
    if (looksLikeTotalDescription(description)) return;

    const rawQuantity = candidate.quantityColumnIndex === -1 ? 0 : numberValue(row[candidate.quantityColumnIndex]);
    const length = candidate.lengthColumnIndex === -1 ? 0 : numberValue(row[candidate.lengthColumnIndex]);
    const width = candidate.widthColumnIndex === -1 ? 0 : numberValue(row[candidate.widthColumnIndex]);
    const height = candidate.heightColumnIndex === -1 ? 0 : numberValue(row[candidate.heightColumnIndex]);
    const packetsPerBox = candidate.packetsPerBoxColumnIndex === -1 ? 0 : numberValue(row[candidate.packetsPerBoxColumnIndex]);
    const sourcePacketQuantity = candidate.packetQuantityColumnIndex === -1 ? 0 : numberValue(row[candidate.packetQuantityColumnIndex]);
    const weight = candidate.netWeightColumnIndex === -1 ? 0 : numberValue(row[candidate.netWeightColumnIndex]);
    const batchNo = candidate.batchIndex === -1 ? "" : String(row[candidate.batchIndex] ?? "").trim();

    const hasItemText = Boolean(description) && !/^pack\s+\d+$/i.test(description);
    const excelCbm = candidate.cbmColumnIndex === -1 ? 0 : numberValue(row[candidate.cbmColumnIndex]);
    const hasMeasure = excelCbm > 0 || length > 0 || width > 0 || height > 0 || rawQuantity > 0 || sourcePacketQuantity > 0;
    if (!hasItemText && !hasMeasure) return;
    if (!description && rawQuantity <= 0 && sourcePacketQuantity <= 0 && length <= 0) return;

    const quantity = rawQuantity > 0 ? rawQuantity : (length > 0 && width > 0 && height > 0 ? 1 : 0);
    const packetQuantity = packetsPerBox > 0 && quantity > 0
      ? quantity * packetsPerBox
      : sourcePacketQuantity;
    const unitCbm = excelCbm > 0 && quantity > 0 ? excelCbm / quantity : calculateUnitCbm(length, width, height);
    const totalCbm = excelCbm > 0 ? excelCbm : calculateLineCbm(length, width, height, quantity);
    processed.push({
      description: description || `Pack ${processed.length + 1}`,
      quantity,
      packetQuantity,
      packetsPerBox,
      length,
      width,
      height,
      weight,
      batchNo,
      unitCbm,
      cbmValue: totalCbm,
      usesExcelCbm: excelCbm > 0,
      totalCbm,
      sourceRow: candidate.headerRow + index + 1
    });
  });

  return removeAggregateTail(processed);
}

function calculateUnitCbm(length, width, height) {
  const l = numberValue(length);
  const w = numberValue(width);
  const h = numberValue(height);
  if (!(l > 0 && w > 0 && h > 0)) return 0;
  return roundTo((l * w * h) / CBM_DIVISOR, 6);
}

function calculateLineCbm(length, width, height, cartons) {
  const unitCbm = calculateUnitCbm(length, width, height);
  const qty = numberValue(cartons);
  if (!(unitCbm > 0 && qty > 0)) return 0;
  return roundTo(unitCbm * qty, 4);
}

function roundTo(value, digits) {
  if (!Number.isFinite(value)) return 0;
  const factor = 10 ** digits;
  return Math.round((value + Number.EPSILON) * factor) / factor;
}

function applyDefaultDimensions(overwriteAll) {
  const length = numberValue(defaultLength.value);
  const width = numberValue(defaultWidth.value);
  const height = numberValue(defaultHeight.value);

  if (!(length > 0 && width > 0 && height > 0)) {
    setStatus("Enter Length, Width and Height first. Decimals like 34.5 are allowed.", "#fff6e8", "#9a6700");
    return;
  }

  state.rows.forEach(row => {
    if (overwriteAll || !(row.length > 0 && row.width > 0 && row.height > 0)) {
      if (length > 0) row.length = length;
      if (width > 0) row.width = width;
      if (height > 0) row.height = height;
    }
    refreshRowCbm(row);
  });

  renderTable(state.rows);
  calculateRecommendation(state.rows);
  setStatus("Values applied", "#eaf7ef", "#207044");
}

function refreshRowCbm(row) {
  if (row.usesExcelCbm) return;
  row.unitCbm = calculateUnitCbm(row.length, row.width, row.height);
  row.totalCbm = calculateLineCbm(row.length, row.width, row.height, row.quantity);
  row.cbmValue = row.totalCbm;
}

function updateRowDimension(index, field, value) {
  const row = state.rows[index];
  if (!row) return;
  row[field] = numberValue(value);
  if (field === "quantity" && row.packetsPerBox > 0) {
    row.packetQuantity = roundTo(row.quantity * row.packetsPerBox, 4);
  }
  refreshRowCbm(row);
  const unitCell = document.querySelector(`[data-unit-cbm-index="${index}"]`);
  const cbmCell = document.querySelector(`[data-cbm-index="${index}"]`);
  const packetCell = document.querySelector(`[data-packet-index="${index}"]`);
  if (unitCell) unitCell.textContent = formatCbm(row.unitCbm);
  if (cbmCell) cbmCell.textContent = formatCbm(row.totalCbm);
  if (packetCell) packetCell.textContent = formatNumber(row.packetQuantity);
  calculateRecommendation(state.rows);
}

function removeAggregateTail(rows) {
  let result = rows.slice();
  while (result.length >= 2) {
    const lastRow = result[result.length - 1];
    const previousRows = result.slice(0, -1);
    if (looksLikeTotalDescription(lastRow.description) || isAggregateRow(lastRow, previousRows)) {
      result = previousRows;
      continue;
    }
    break;
  }
  return result;
}

function isAggregateRow(lastRow, previousRows) {
  const previousCbm = previousRows.reduce((sum, row) => sum + row.totalCbm, 0);
  const previousQuantity = previousRows.reduce((sum, row) => sum + row.quantity, 0);
  const previousPackets = previousRows.reduce((sum, row) => sum + row.packetQuantity, 0);
  const cbmMatches = previousCbm > 0 && Math.abs(lastRow.totalCbm - previousCbm) <= Math.max(0.01, previousCbm * 0.002);
  const quantityMatches = previousQuantity > 0 && Math.abs(lastRow.quantity - previousQuantity) <= 0.51;
  const packetMatches = previousPackets > 0 && Math.abs(lastRow.packetQuantity - previousPackets) <= 0.51;
  const generatedAggregate = /^Pack\s+\d+$/i.test(lastRow.description)
    && (lastRow.quantity >= previousQuantity || lastRow.packetQuantity >= previousPackets);
  return cbmMatches || quantityMatches || packetMatches || generatedAggregate;
}

function looksLikeTotalDescription(description) {
  const header = normalize(description);
  return /^(total|subtotal|grandtotal|overalltotal)(packing|quantity|cartons|ctns|cbm|volume|shipment|packets|pkts|qty)?$/.test(header);
}

function isSummaryRow(row) {
  const values = row.map(value => normalize(value)).filter(Boolean);
  if (!values.length) return true;
  return values.some(value => /^(total|subtotal|grandtotal|totalpacking|packingtotal|totalquantity|totalcartons|totalctns|totalcbm|totalvolume|totalshipment|overalltotal|totalpackets|totalpkts)$/.test(value));
}

function isRepeatedHeaderRow(row, candidate) {
  const description = candidate.descriptionColumnIndex === -1 ? "" : normalize(row[candidate.descriptionColumnIndex]);
  return isLengthHeader(description) || isWidthHeader(description) || isHeightHeader(description) || Boolean(classifyCbmHeader(description));
}

function renderTable(rows) {
  dataBody.innerHTML = "";

  rows.forEach((row, index) => {
    const tr = document.createElement("tr");
    tr.innerHTML = `
      <td>${index + 1}</td>
      <td>${escapeHtml(row.description)}</td>
      <td>${dimensionInput(index, "length", row.length)}</td>
      <td>${dimensionInput(index, "width", row.width)}</td>
      <td>${dimensionInput(index, "height", row.height)}</td>
      <td>${dimensionInput(index, "quantity", row.quantity)}</td>
      <td data-packet-index="${index}">${formatNumber(row.packetQuantity)}</td>
      <td data-unit-cbm-index="${index}">${formatCbm(row.unitCbm)}</td>
      <td><strong data-cbm-index="${index}">${formatCbm(row.totalCbm)}</strong></td>
    `;
    dataBody.appendChild(tr);
  });

  dataBody.querySelectorAll("input[data-row]").forEach(input => {
    input.addEventListener("input", event => {
      const target = event.target;
      updateRowDimension(Number(target.dataset.row), target.dataset.field, target.value);
    });
  });
}

function dimensionInput(index, field, value) {
  const shown = value > 0 ? formatInputNumber(value) : "";
  return `<input class="dim-input" type="text" inputmode="decimal" autocomplete="off" data-row="${index}" data-field="${field}" value="${escapeHtml(shown)}" placeholder="0">`;
}

function calculateRecommendation(rows) {
  const total = roundTo(rows.reduce((sum, row) => sum + numberValue(row.totalCbm), 0), 4);
  const packetTotal = rows.reduce((sum, row) => sum + numberValue(row.packetQuantity), 0);
  const usesExcelCbm = rows.some(row => row.usesExcelCbm);
  packCount.textContent = rows.length;
  packetQuantityTotal.textContent = formatNumber(packetTotal);
  totalCbmEl.textContent = `${formatCbm(total)} CBM`;
  const suitable = containers.find(container => total <= container.capacity);

  document.querySelectorAll(".container-option").forEach(card => card.classList.remove("selected"));

  containers.forEach(container => {
    const usage = total / container.capacity * 100;
    const elementId = container.id === "20FT" ? "use20" : container.id === "40FT" ? "use40" : "use40hc";
    document.getElementById(elementId).textContent = `${Math.min(usage, 100).toFixed(1)}% volume used`;
    if (suitable && container.id === suitable.id) {
      document.querySelector(`[data-container="${container.id}"]`).classList.add("selected");
    }
  });

  const missingDimensions = usesExcelCbm
    ? 0
    : rows.filter(row => !(row.length > 0 && row.width > 0 && row.height > 0 && row.quantity > 0)).length;

  if (!suitable) {
    recommendedSmall.textContent = total > 0 ? "No standard option" : "Enter dimensions";
    recommendation.className = "recommendation";
    recommendation.innerHTML = total > 0 ? `
      <h3>No standard container is large enough</h3>
      <p>Total shipment volume: <strong>${formatCbm(total)} CBM</strong></p>
      <p>The current system supports 20 FT, 40 FT and 40 FT High Cube containers.</p>
    ` : `
      <h3>CBM is calculated from box size</h3>
      <p>Formula: <strong>CBM = (Length × Width × Height × No. of Ctns) ÷ 1,000,000</strong></p>
      <p>Enter Length, Width and Height in centimetres for each box, or use the default box size fields above.</p>
    `;
    return;
  }

  const remaining = suitable.capacity - total;
  recommendedSmall.textContent = suitable.name;
  recommendation.className = "recommendation";
  recommendation.innerHTML = `
    <h3>Recommended Container: ${suitable.name}</h3>
    <p>${usesExcelCbm ? "CBM source: <strong>Excel calculated/result values</strong>" : "Formula: <strong>CBM = (Length × Width × Height × No. of Ctns) ÷ 1,000,000</strong>"}</p>
    <div class="recommendation-metrics">
      <div class="recommendation-metric total"><span>Total CBM</span><strong>${formatCbm(total)} CBM</strong></div>
      <div class="recommendation-metric capacity"><span>Capacity</span><strong>${formatCbm(suitable.capacity)} CBM</strong></div>
      <div class="recommendation-metric remaining"><span>Remaining</span><strong>${formatCbm(remaining)} CBM</strong></div>
      <div class="recommendation-metric utilization"><span>Utilization</span><strong>${(total / suitable.capacity * 100).toFixed(1)}%</strong></div>
    </div>
    ${missingDimensions ? `<p>${missingDimensions} row(s) still need Length, Width, Height or carton quantity.</p>` : ""}
  `;
}

async function downloadPdf() {
  if (!state.rows.length) {
    setStatus("Load data before downloading PDF", "#fff6e8", "#9a6700");
    return;
  }

  if (!window.jspdf || !window.jspdf.jsPDF) {
    setStatus("PDF library could not be loaded", "#fff0f0", "#b42318");
    return;
  }

  if (!(await ensureFreelanLogo())) {
    setStatus("The local Freelan logo could not be loaded", "#fff0f0", "#b42318");
    return;
  }

  const { jsPDF } = window.jspdf;
  const doc = new jsPDF({ unit: "mm", format: "a4", orientation: "landscape" });
  const rows = state.rows;
  const cartonTotal = rows.reduce((sum, row) => sum + numberValue(row.quantity), 0);
  const total = roundTo(rows.reduce((sum, row) => sum + row.totalCbm, 0), 4);
  const packetTotal = rows.reduce((sum, row) => sum + numberValue(row.packetQuantity), 0);
  const suitable = containers.find(container => total <= container.capacity);
  const remaining = suitable ? suitable.capacity - total : 0;
  const utilization = suitable ? total / suitable.capacity * 100 : 0;
  const pageWidth = doc.internal.pageSize.getWidth();
  const pageHeight = doc.internal.pageSize.getHeight();
  const reportDate = new Date().toLocaleDateString("en-GB");
  let y = 47;

  doc.setFillColor(16, 42, 67);
  doc.rect(0, 0, pageWidth, 37, "F");
  doc.addImage(freelanLogo, "PNG", 14, 7, 43, 16);
  doc.setTextColor(255, 255, 255);
  doc.setFontSize(15);
  doc.setFont("helvetica", "bold");
  doc.text("FREELAN ENTERPRISES (PVT) LTD", 64, 12);
  doc.setFontSize(9);
  doc.setFont("helvetica", "normal");
  doc.text("No. 38, Gunawardana Mawatha, Navimana South, Matara, Sri Lanka", 64, 18);
  doc.text("www.freelansrilanka.com", 64, 24);
  doc.setFontSize(16);
  doc.setFont("helvetica", "bold");
  doc.text("Container CBM Recommendation Report", 64, 32);
  doc.setFontSize(9);
  doc.setFont("helvetica", "normal");
  doc.text(`Report Date: ${reportDate}`, pageWidth - 14, 12, { align: "right" });

  doc.setTextColor(16, 42, 67);
  doc.setFontSize(10);
  doc.setFont("helvetica", "bold");
  doc.text("SHIPMENT SUMMARY", 14, y);
  y += 4;
  const cards = [
    { label: "TOTAL CARTONS", value: formatNumber(cartonTotal) },
    { label: "TOTAL CBM", value: `${formatCbm(total)} CBM` },
    { label: "RECOMMENDED CONTAINER", value: suitable ? suitable.name.toUpperCase() : "NO STANDARD OPTION" },
    { label: "PACKET QUANTITY", value: formatNumber(packetTotal) }
  ];
  const cardWidth = (pageWidth - 28 - 12) / 4;
  cards.forEach((card, index) => {
    const x = 14 + index * (cardWidth + 4);
    doc.setFillColor(245, 248, 252);
    doc.setDrawColor(207, 219, 230);
    doc.roundedRect(x, y, cardWidth, 20, 2, 2, "FD");
    doc.setTextColor(72, 96, 117);
    doc.setFontSize(7);
    doc.setFont("helvetica", "bold");
    doc.text(card.label, x + 4, y + 7);
    doc.setTextColor(16, 42, 67);
    doc.setFontSize(index === 2 ? 10 : 13);
    doc.text(card.value, x + 4, y + 15);
  });

  y += 27;
  doc.setFillColor(239, 245, 250);
  doc.setDrawColor(180, 198, 214);
  doc.roundedRect(14, y, pageWidth - 28, 48, 2, 2, "FD");
  doc.setTextColor(16, 42, 67);
  doc.setFontSize(9);
  doc.setFont("helvetica", "bold");
  doc.text("CONTAINER RECOMMENDATION", 20, y + 8);
  doc.setFontSize(18);
  doc.text(suitable ? suitable.name.toUpperCase() : "NO STANDARD OPTION", 20, y + 18);
  if (suitable) {
    const metrics = [
      { label: "CONTAINER CAPACITY", value: `${formatCbm(suitable.capacity)} CBM`, x: 123 },
      { label: "REMAINING CAPACITY", value: `${formatCbm(remaining)} CBM`, x: 181 },
      { label: "CONTAINER UTILIZATION", value: `${utilization.toFixed(1)}%`, x: 239 }
    ];
    metrics.forEach(metric => {
      doc.setTextColor(72, 96, 117);
      doc.setFontSize(7);
      doc.text(metric.label, metric.x, y + 9);
      doc.setTextColor(16, 42, 67);
      doc.setFontSize(13);
      doc.setFont("helvetica", "bold");
      doc.text(metric.value, metric.x, y + 18);
    });
    const barX = 20;
    const barY = y + 29;
    const barWidth = 88;
    doc.setFillColor(215, 225, 234);
    doc.roundedRect(barX, barY, barWidth, 7, 1.5, 1.5, "F");
    doc.setFillColor(15, 118, 110);
    doc.roundedRect(barX, barY, barWidth * Math.min(utilization, 100) / 100, 7, 1.5, 1.5, "F");
    doc.setTextColor(72, 96, 117);
    doc.setFontSize(7);
    doc.setFont("helvetica", "normal");
    doc.text(`${utilization.toFixed(1)}% utilized`, barX + barWidth + 4, barY + 5);
  } else {
    doc.setFontSize(10);
    doc.setFont("helvetica", "normal");
    doc.text(`Total shipment volume: ${formatCbm(total)} CBM`, 20, y + 30);
  }
  doc.setTextColor(72, 96, 117);
  doc.setFontSize(7);
  doc.setFont("helvetica", "normal");
  doc.text("Standard capacities: 20 FT = 33 CBM | 40 FT = 67 CBM | 40 FT High Cube = 76 CBM", 20, y + 43);
  y += 56;
  y = drawPdfTableHeader(doc, y);
  rows.forEach((row, index) => {
    if (y > pageHeight - 18) {
      addPdfFooter(doc, pageWidth, pageHeight);
      doc.addPage();
      y = 16;
      y = drawPdfTableHeader(doc, y);
    }
    if (index % 2 === 0) {
      doc.setFillColor(245, 248, 252);
      doc.rect(12, y - 5, pageWidth - 24, 7, "F");
    }
    doc.setTextColor(55, 65, 81);
    doc.setFontSize(7.5);
    doc.text(String(index + 1), 16, y);
    const columns = getPdfTableColumns();
    const values = [
      row.description || `Pack ${index + 1}`,
      formatNumber(row.quantity),
      formatNumber(row.packetsPerBox),
      formatNumber(row.packetQuantity),
      row.batchNo || "-",
      formatNumber(row.length),
      formatNumber(row.width),
      formatNumber(row.height),
      row.weight ? formatNumber(row.weight) : "-",
      formatCbm(row.totalCbm)
    ];
    values.forEach((value, valueIndex) => {
      const column = columns[valueIndex];
      const text = valueIndex === 0 ? doc.splitTextToSize(value, column.width - 2)[0] : value;
      if (column.align === "right") {
        doc.text(text, column.x + column.width, y, { align: "right" });
      } else {
        doc.text(text, column.x, y);
      }
    });
    y += 7;
  });

  doc.setDrawColor(16, 42, 67);
  doc.line(12, y, pageWidth - 12, y);
  doc.setFont("helvetica", "bold");
  doc.setTextColor(16, 42, 67);
  doc.setFontSize(7.5);
  doc.text("TOTAL", 27, y + 7);
  const columns = getPdfTableColumns();
  doc.text(formatNumber(packetTotal), columns[2].x + columns[2].width, y + 7, { align: "right" });
  doc.text(`${formatCbm(total)} CBM`, columns[9].x + columns[9].width, y + 7, { align: "right" });
  addPdfFooter(doc, pageWidth, pageHeight);
  const pdfBlob = doc.output("blob");
  const downloadUrl = URL.createObjectURL(pdfBlob);
  const downloadLink = document.createElement("a");
  downloadLink.href = downloadUrl;
  downloadLink.download = "Container_CBM_Report.pdf";
  document.body.appendChild(downloadLink);
  downloadLink.click();
  downloadLink.remove();
  setTimeout(() => URL.revokeObjectURL(downloadUrl), 1000);
}

function ensureFreelanLogo() {
  if (!freelanLogo) return Promise.resolve(false);
  if (freelanLogo.complete) return Promise.resolve(freelanLogo.naturalWidth > 0);
  return new Promise(resolve => {
    freelanLogo.addEventListener("load", () => resolve(true), { once: true });
    freelanLogo.addEventListener("error", () => resolve(false), { once: true });
  });
}

function getPdfTableColumns() {
  return [
    { label: "BOX DESCRIPTION", x: 27, width: 52, align: "left" },
    { label: "CARTONS", x: 79, width: 17, align: "right" },
    { label: "PKT / BOX", x: 96, width: 18, align: "right" },
    { label: "QTY PKT", x: 114, width: 20, align: "right" },
    { label: "BATCH NO.", x: 134, width: 25, align: "left" },
    { label: "L", x: 159, width: 13, align: "right" },
    { label: "W", x: 172, width: 13, align: "right" },
    { label: "H", x: 185, width: 13, align: "right" },
    { label: "WEIGHT", x: 198, width: 22, align: "right" },
    { label: "CBM", x: 220, width: 24, align: "right" }
  ];
}

function drawPdfTableHeader(doc, y) {
  doc.setFillColor(16, 42, 67);
  doc.rect(12, y - 6, 273, 9, "F");
  doc.setTextColor(255, 255, 255);
  doc.setFont("helvetica", "bold");
  doc.setFontSize(7);
  doc.text("#", 16, y);
  getPdfTableColumns().forEach(column => {
    if (column.align === "right") {
      doc.text(column.label, column.x + column.width, y, { align: "right" });
    } else {
      doc.text(column.label, column.x, y);
    }
  });
  return y + 8;
}

function addPdfFooter(doc, pageWidth, pageHeight) {
  const pageNumber = doc.getNumberOfPages();
  doc.setDrawColor(207, 219, 230);
  doc.line(12, pageHeight - 13, pageWidth - 12, pageHeight - 13);
  doc.setFont("helvetica", "normal");
  doc.setFontSize(7.5);
  doc.setTextColor(107, 114, 128);
  doc.text("FREELAN ENTERPRISES (PVT) LTD | www.freelansrilanka.com", 14, pageHeight - 7);
  doc.text(`Page ${pageNumber}`, pageWidth - 14, pageHeight - 7, { align: "right" });
}

function createDemoCandidate() {
  return {
    sheetName: "Demo data",
    sheetIndex: 0,
    rows: [
      ["Pack", "Length (cm)", "Width (cm)", "Height (cm)", "Weight (kg)", "No. of Ctns", "Pkt per Box", "Quantity (Pkt)"],
      ["Pack 001", 34.5, 34.5, 23.5, 12.5, 100, 12, 1200],
      ["Pack 002", 48.5, 40.25, 25.5, 10.75, 100, 10, 1000],
      ["Pack 003", 55.5, 42.5, 24.25, 14.5, 110, 8, 880],
      ["Pack 004", 52.25, 38.5, 28.5, 11.25, 90, 12, 1080]
    ],
    headerIndex: 0,
    headerRow: 1,
    cbmColumnIndex: -1,
    cbmHeader: "Calculated CBM",
    cbmType: "calculated",
    quantityColumnIndex: 5,
    quantityHeader: "No. of Ctns",
    packetQuantityColumnIndex: 7,
    packetsPerBoxColumnIndex: 6,
    netWeightColumnIndex: 4,
    grossWeightColumnIndex: -1,
    descriptionColumnIndex: 0,
    lengthColumnIndex: 1,
    widthColumnIndex: 2,
    heightColumnIndex: 3,
    unitColumnIndex: -1,
    hasDimensions: true,
    recognizedCount: 6,
    validCount: 4
  };
}

function clearData() {
  excelFile.value = "";
  fileName.textContent = ".xlsx / .xls";
  state.candidates = [];
  state.selectedIndex = -1;
  state.rows = [];
  defaultLength.value = "";
  defaultWidth.value = "";
  defaultHeight.value = "";
  detectionPanel.hidden = true;
  resetCalculatedView();
}

function showError(message) {
  setStatus("Error", "#fff0f0", "#b42318");
  dataBody.innerHTML = `<tr><td colspan="9" class="empty">${escapeHtml(message)}</td></tr>`;
  resetCalculatedView(false);
  recommendation.textContent = message;
}

function resetCalculatedView(resetRecommendation = true) {
  packCount.textContent = "0";
  packetQuantityTotal.textContent = "0";
  totalCbmEl.textContent = "0.0000 CBM";
  recommendedSmall.textContent = "—";
  if (resetRecommendation) {
    dataBody.innerHTML = '<tr><td colspan="9" class="empty">Upload an Excel file to display pack data.</td></tr>';
    recommendation.className = "recommendation empty-result";
    recommendation.textContent = "Upload your Excel data to get a recommendation.";
    setStatus("Waiting for file", "#eef2f6", "#52606d");
  } else {
    recommendation.className = "recommendation empty-result";
  }

  document.querySelectorAll(".container-option").forEach(card => card.classList.remove("selected"));
  ["use20", "use40", "use40hc"].forEach(id => {
    document.getElementById(id).textContent = "—";
  });
}

function setStatus(text, background, color) {
  statusBadge.textContent = text;
  statusBadge.style.background = background;
  statusBadge.style.color = color;
}

function normalize(value) {
  return String(value ?? "").trim().toLowerCase().replace(/[^a-z0-9]/g, "");
}

function numberValue(value) {
  if (typeof value === "number") return Number.isFinite(value) ? value : 0;
  let text = String(value ?? "").trim();
  if (!text) return 0;
  if (text.endsWith(".")) text = text.slice(0, -1);
  const hasDot = text.includes(".");
  const hasComma = text.includes(",");
  if (hasDot && hasComma) {
    text = text.replace(/,/g, "");
  } else if (hasComma && !hasDot) {
    text = text.replace(",", ".");
  }
  const cleaned = text.replace(/[^0-9.-]/g, "");
  if (!cleaned || cleaned === "-" || cleaned === ".") return 0;
  const number = parseFloat(cleaned);
  return Number.isFinite(number) ? number : 0;
}

function formatInputNumber(value) {
  if (!Number.isFinite(value) || value === 0) return "";
  return String(value);
}

function formatNumber(value) {
  const number = numberValue(value);
  return Number.isInteger(number) ? String(number) : number.toFixed(2);
}

function formatCbm(value) {
  return roundTo(numberValue(value), 4).toFixed(4);
}

function displayHeader(value, fallback) {
  const text = String(value ?? "").trim();
  return text || fallback;
}

function formatNumber(value) {
  const number = numberValue(value);
  return String(Number.isInteger(number) ? number : number.toFixed(2));
}

function formatCbm(value) {
  return numberValue(value).toFixed(4).replace(/\.?(0+)$/, "");
}

function escapeHtml(value) {
  return String(value)
    .replace(/&/g, "&amp;")
    .replace(/</g, "&lt;")
    .replace(/>/g, "&gt;")
    .replace(/"/g, "&quot;")
    .replace(/'/g, "&#039;");
}
