import fs from "node:fs/promises";
import path from "node:path";
import { FileBlob, SpreadsheetFile, Workbook } from "@oai/artifact-tool";

const baseDir = "C:/Users/ADOLF/Desktop/Repositorios/Creacion_Dataset/Dataset_V1";
const inputDir = path.join(baseDir, "Excel");
const outputDir = path.join(baseDir, "Analisis");
const outputPath = path.join(outputDir, "propuesta_estructuracion_temporal.xlsx");
const families = ["zusy", "strictor", "vbclone", "eylr", "fragtor", "injector", "razy", "salgorea"];
const goodYearMin = 1992;
const goodYearMax = 2026;

const asString = (value) => (value === null || value === undefined ? "" : String(value));
const asNumber = (value) => {
  if (typeof value === "number") return value;
  const parsed = Number(value);
  return Number.isFinite(parsed) ? parsed : 0;
};
const excelSerialToDate = (serial) => {
  const ms = Math.round((serial - 25569) * 86400 * 1000);
  const date = new Date(ms);
  return Number.isNaN(date.getTime()) ? null : date;
};
const asDate = (value) => {
  if (value instanceof Date && !Number.isNaN(value.getTime())) return value;
  if (typeof value === "number") return excelSerialToDate(value);
  const parsed = new Date(asString(value));
  return Number.isNaN(parsed.getTime()) ? null : parsed;
};
const formatDay = (date) => {
  if (!date) return "";
  const y = date.getUTCFullYear();
  const m = String(date.getUTCMonth() + 1).padStart(2, "0");
  const d = String(date.getUTCDate()).padStart(2, "0");
  return `${y}-${m}-${d}`;
};
const pct = (value, total) => (total ? value / total : 0);
const round = (value, decimals = 2) => Number(value.toFixed(decimals));
const avg = (values) => values.reduce((sum, value) => sum + value, 0) / values.length;
const median = (values) => values.length % 2
  ? values[(values.length - 1) / 2]
  : (values[values.length / 2 - 1] + values[values.length / 2]) / 2;
const countBy = (items, getter) => {
  const map = new Map();
  for (const item of items) {
    const key = getter(item) || "sin dato";
    map.set(key, (map.get(key) ?? 0) + 1);
  }
  return [...map.entries()].sort((a, b) => b[1] - a[1] || String(a[0]).localeCompare(String(b[0])));
};
const lettersToIndex = (letters) => [...letters].reduce((n, ch) => n * 26 + ch.charCodeAt(0) - 64, 0);
const indexToLetters = (n) => {
  let s = "";
  while (n > 0) {
    const mod = (n - 1) % 26;
    s = String.fromCharCode(65 + mod) + s;
    n = Math.floor((n - mod) / 26);
  }
  return s;
};
const writeTable = (sheet, startCell, header, rowsToWrite, tableName) => {
  const match = startCell.match(/^([A-Z]+)(\d+)$/);
  const startCol = lettersToIndex(match[1]);
  const startRow = Number(match[2]);
  const endCol = indexToLetters(startCol + header.length - 1);
  const endRow = startRow + rowsToWrite.length;
  const range = `${startCell}:${endCol}${endRow}`;
  sheet.getRange(range).values = [header, ...rowsToWrite];
  sheet.getRange(`${startCell}:${endCol}${startRow}`).format = {
    fill: "#374151",
    font: { bold: true, color: "#FFFFFF" },
  };
  sheet.tables.add(range, true, tableName);
  return range;
};
const splitCounts = (n) => {
  const train = Math.floor(n * 0.7);
  const val = Math.floor(n * 0.15);
  const test = n - train - val;
  return { train, val, test };
};
const sortTemporal = (items) => [...items].sort((a, b) => {
  const delta = a.creationDate - b.creationDate;
  if (delta !== 0) return delta;
  return a.hash.localeCompare(b.hash);
});
const selectTemporalCoverage = (items, cap) => {
  const sorted = sortTemporal(items);
  if (!cap || sorted.length <= cap) return sorted;
  if (cap <= 1) return sorted.slice(0, cap);
  const selected = [];
  const used = new Set();
  for (let i = 0; i < cap; i += 1) {
    let index = Math.round((i * (sorted.length - 1)) / (cap - 1));
    while (used.has(index) && index < sorted.length - 1) index += 1;
    while (used.has(index) && index > 0) index -= 1;
    used.add(index);
    selected.push(sorted[index]);
  }
  return sortTemporal(selected);
};
const yearRange = (items) => {
  if (items.length === 0) return "";
  const years = items.map((r) => r.creationDate.getUTCFullYear());
  const first = years[0];
  const last = years[years.length - 1];
  return first === last ? String(first) : `${first}-${last}`;
};
const dateRange = (items) => {
  if (items.length === 0) return "";
  return `${formatDay(items[0].creationDate)} a ${formatDay(items[items.length - 1].creationDate)}`;
};
const splitYearPlan = (items, counts) => {
  const trainItems = items.slice(0, counts.train);
  const valItems = items.slice(counts.train, counts.train + counts.val);
  const testItems = items.slice(counts.train + counts.val);
  return {
    trainYears: yearRange(trainItems),
    valYears: yearRange(valItems),
    testYears: yearRange(testItems),
    trainDates: dateRange(trainItems),
    valDates: dateRange(valItems),
    testDates: dateRange(testItems),
    trainCut: trainItems.length ? formatDay(trainItems.at(-1).creationDate) : "",
    valCut: valItems.length ? formatDay(valItems.at(-1).creationDate) : "",
  };
};
const combineYearRanges = (rows, columnIndex) => {
  const years = [];
  for (const row of rows) {
    const matches = String(row[columnIndex] ?? "").match(/\d{4}/g) ?? [];
    for (const match of matches) years.push(Number(match));
  }
  if (years.length === 0) return "";
  const min = Math.min(...years);
  const max = Math.max(...years);
  return min === max ? String(min) : `${min}-${max}`;
};

const records = [];
for (const family of families) {
  const workbook = await SpreadsheetFile.importXlsx(await FileBlob.load(path.join(inputDir, `${family}.xlsx`)));
  const sheet = workbook.worksheets.getItemAt(0);
  const raw = sheet.getUsedRange(true).values;
  const headers = raw[0].map(String);
  const idx = Object.fromEntries(headers.map((header, i) => [header, i]));
  for (const row of raw.slice(1)) {
    if (!row.some((value) => value !== null && value !== "")) continue;
    const scanDate = asDate(row[idx.fecha_escaneo_vt]);
    const creationDate = asDate(row[idx.dia_creacion_archivo]);
    const creationYear = creationDate ? creationDate.getUTCFullYear() : null;
    const validCreation = creationDate && creationYear >= goodYearMin && creationYear <= goodYearMax;
    records.push({
      family,
      hash: asString(row[idx.hash_md5]),
      lote: asString(row[idx.lote_origen]),
      type: asString(row[idx.tipo_probable]),
      extension: asString(row[idx.extension]),
      size: asNumber(row[idx.size]),
      detection: asNumber(row[idx.detection_percent]),
      scanDate,
      scanDay: formatDay(scanDate),
      creationDate,
      creationYear: validCreation ? creationYear : "fuera_rango_o_vacia",
      validCreation: Boolean(validCreation),
    });
  }
}

const familySummary = families.map((family) => {
  const subset = records.filter((r) => r.family === family);
  const valid = subset.filter((r) => r.validCreation);
  const detections = subset.map((r) => r.detection).sort((a, b) => a - b);
  const years = valid.map((r) => r.creationDate).sort((a, b) => a - b);
  return {
    family,
    total: subset.length,
    temporalEligible: valid.length,
    temporalOutlier: subset.length - valid.length,
    eligiblePct: pct(valid.length, subset.length),
    creationMin: years.length ? formatDay(years[0]) : "",
    creationMax: years.length ? formatDay(years[years.length - 1]) : "",
    scanMin: formatDay(subset.map((r) => r.scanDate).sort((a, b) => a - b)[0]),
    scanMax: formatDay(subset.map((r) => r.scanDate).sort((a, b) => a - b).at(-1)),
    detectionAvg: avg(detections),
    detectionMedian: median(detections),
  };
});

const minTotal = Math.min(...familySummary.map((r) => r.total));
const maxTotal = Math.max(...familySummary.map((r) => r.total));
const minEligible = Math.min(...familySummary.map((r) => r.temporalEligible));
const maxEligible = Math.max(...familySummary.map((r) => r.temporalEligible));
const totalRows = familySummary.reduce((sum, r) => sum + r.total, 0);
const totalEligible = familySummary.reduce((sum, r) => sum + r.temporalEligible, 0);
const totalOutlier = familySummary.reduce((sum, r) => sum + r.temporalOutlier, 0);

const strategies = [
  {
    name: "Usar todo como esta",
    sheetName: "Split_Todo",
    tableName: "TablaSplitTodo",
    relationUser: "8.5 : 1",
    cap: null,
    mode: "Sin limite",
    comment: "Maximiza datos, pero conserva el desbalance fuerte y mezcla familias muy dominantes.",
    recommendation: "Usarlo como baseline descriptivo, no como particion principal de entrenamiento.",
  },
  {
    name: "Balancear a 1000",
    sheetName: "Split_1000",
    tableName: "TablaSplit1000",
    relationUser: "1 : 1",
    cap: 1000,
    mode: "Objetivo 1000 por familia",
    comment: "Balance casi perfecto sin duplicar; exacto solo si se completa zusy con 32 muestras o se baja todo a 968.",
    recommendation: "Excelente baseline controlado para comparar modelos.",
  },
  {
    name: "Limitar a 1500",
    sheetName: "Split_1500",
    tableName: "TablaSplit1500",
    relationUser: "1.41 : 1",
    cap: 1500,
    mode: "Cap 1500 por familia",
    comment: "Buen compromiso entre balance y volumen; reduce mucho fragtor/salgorea sin perder variedad.",
    recommendation: "Recomendacion principal para primer dataset temporal.",
  },
  {
    name: "Limitar a 2000",
    sheetName: "Split_2000",
    tableName: "TablaSplit2000",
    relationUser: "1.88 : 1",
    cap: 2000,
    mode: "Cap 2000 por familia",
    comment: "Mas volumen que 1500, con desbalance moderado; util si el modelo tolera clases grandes.",
    recommendation: "Buena segunda corrida si el baseline de 1500 queda estable.",
  },
];

const strategyRows = strategies.map((s) => {
  const perFamilyTotal = familySummary.map((r) => s.cap ? Math.min(r.total, s.cap) : r.total);
  const perFamilyEligible = familySummary.map((r) => s.cap ? Math.min(r.temporalEligible, s.cap) : r.temporalEligible);
  const minSelected = Math.min(...perFamilyTotal);
  const maxSelected = Math.max(...perFamilyTotal);
  const minSelectedEligible = Math.min(...perFamilyEligible);
  const maxSelectedEligible = Math.max(...perFamilyEligible);
  return {
    ...s,
    selectedTotal: perFamilyTotal.reduce((sum, value) => sum + value, 0),
    selectedEligible: perFamilyEligible.reduce((sum, value) => sum + value, 0),
    relationCurrent: `${round(maxSelected / minSelected, 2)} : 1`,
    relationEligible: `${round(maxSelectedEligible / minSelectedEligible, 2)} : 1`,
    minSelected,
    maxSelected,
    minSelectedEligible,
    maxSelectedEligible,
  };
});

const impactRows = [];
for (const familyRow of familySummary) {
  for (const s of strategies) {
    const selectedTotal = s.cap ? Math.min(familyRow.total, s.cap) : familyRow.total;
    const selectedEligible = s.cap ? Math.min(familyRow.temporalEligible, s.cap) : familyRow.temporalEligible;
    impactRows.push([
      s.name,
      familyRow.family,
      selectedTotal,
      selectedEligible,
      selectedTotal - selectedEligible,
      pct(selectedEligible, selectedTotal),
    ]);
  }
}

const splitRows = [];
const splitRowsByStrategy = new Map(strategies.map((s) => [s.name, []]));
for (const familyRow of familySummary) {
  for (const s of strategies) {
    const eligibleRecords = records.filter((r) => r.family === familyRow.family && r.validCreation);
    const selectedRecords = selectTemporalCoverage(eligibleRecords, s.cap);
    const n = selectedRecords.length;
    const split = splitCounts(n);
    const years = splitYearPlan(selectedRecords, split);
    splitRows.push([
      s.name,
      familyRow.family,
      n,
      split.train,
      split.val,
      split.test,
      years.trainYears,
      years.valYears,
      years.testYears,
      years.trainCut,
      years.valCut,
      "Ordenar por dia_creacion_archivo valida; train=antiguo, validacion=intermedio, test=reciente.",
    ]);
    splitRowsByStrategy.get(s.name).push([
      familyRow.family,
      n,
      split.train,
      split.val,
      split.test,
      years.trainYears,
      years.valYears,
      years.testYears,
      years.trainCut,
      years.valCut,
      years.trainDates,
      years.valDates,
      years.testDates,
      "Ordenar por dia_creacion_archivo valida; train=antiguo, validacion=intermedio, test=reciente.",
    ]);
  }
}

const yearRows = [];
const yearCounts = countBy(records.filter((r) => r.validCreation), (r) => r.creationYear)
  .sort((a, b) => Number(a[0]) - Number(b[0]));
for (const [year, count] of yearCounts) {
  yearRows.push([year, count]);
}

const wb = Workbook.create();
const dash = wb.worksheets.add("Dashboard");
const proposals = wb.worksheets.add("Propuestas");
const temporal = wb.worksheets.add("Temporalidad");
const splits = wb.worksheets.add("Splits_Propuestos");
const notes = wb.worksheets.add("Criterios");
const splitDetailSheets = new Map(strategies.map((s) => [s.name, wb.worksheets.add(s.sheetName)]));
for (const sheet of [dash, proposals, temporal, splits, notes, ...splitDetailSheets.values()]) sheet.showGridLines = false;

dash.getRange("A1").values = [["Propuesta de estructuracion temporal del dataset"]];
dash.getRange("A1:L1").merge();
dash.getRange("A1").format = { fill: "#111827", font: { bold: true, color: "#FFFFFF", size: 18 } };
dash.getRange("A3:H4").values = [
  ["Familias", "Muestras", "Elegibles temporales", "Fechas outlier", "% elegibles", "Rango total actual", "Rango elegible", "Recomendacion"],
  [
    families.length,
    totalRows,
    totalEligible,
    totalOutlier,
    pct(totalEligible, totalRows),
    `${round(maxTotal / minTotal, 2)} : 1`,
    `${round(maxEligible / minEligible, 2)} : 1`,
    "Limitar a 1500 + split cronologico 70/15/15",
  ],
];
dash.getRange("A3:H3").format = { fill: "#374151", font: { bold: true, color: "#FFFFFF" } };
dash.getRange("A4:H4").format = { fill: "#F3F4F6", font: { bold: true, color: "#111827" } };
dash.getRange("E4:E4").format.numberFormat = "0.0%";
dash.getRange("A:H").format.columnWidthPx = 155;
dash.getRange("H:H").format.columnWidthPx = 290;

dash.getRange("A7:D15").values = [["Familia", "Total", "Elegibles", "Outlier"], ...familySummary.map((r) => [r.family, r.total, r.temporalEligible, r.temporalOutlier])];
dash.getRange("F7:I11").values = [["Estrategia", "Seleccion total", "Seleccion elegible", "Relacion elegible"], ...strategyRows.map((r) => [r.name, r.selectedTotal, r.selectedEligible, r.relationEligible])];
dash.getRange("AA7:AB15").values = [["Familia", "Outlier"], ...familySummary.map((r) => [r.family, r.temporalOutlier])];
dash.getRange("A19:B40").values = [["Ano creacion", "Muestras elegibles"], ...yearRows];
for (const range of ["A7:D7", "F7:I7", "A19:B19"]) {
  dash.getRange(range).format = { fill: "#374151", font: { bold: true, color: "#FFFFFF" } };
}
dash.getRange("F:I").format.columnWidthPx = 170;
dash.getRange("I:I").format.columnWidthPx = 190;

const dashCharts = [
  ["bar", "A7:C15", "Total vs elegibles por familia", "K3", "R18"],
  ["bar", "F7:H11", "Muestras por estrategia", "K20", "R35"],
  ["line", `A19:B${yearRows.length + 19}`, "Distribucion temporal por ano de creacion", "S3", "Z18"],
  ["bar", "AA7:AB15", "Fechas outlier por familia", "S20", "Z35"],
];
for (const [type, range, title, from, to] of dashCharts) {
  const chart = dash.charts.add(type, dash.getRange(range));
  chart.title = title;
  chart.hasLegend = title.includes("vs") || title.includes("estrategia");
  chart.xAxis = { axisType: "textAxis" };
  chart.yAxis = { numberFormatCode: "#,##0" };
  chart.setPosition(from, to);
}
dash.getRange("A43:L45").merge();
dash.getRange("A43").values = [["Decision sugerida: construir primero un dataset temporal con cap 1500 por familia, usando solo muestras con dia_creacion_archivo valido entre 1992 y 2026 para el split cronologico. Las muestras fuera de rango se guardan como holdout/auditoria y no se mezclan en train/valid/test temporal."]];
dash.getRange("A43").format = { fill: "#FEF3C7", font: { color: "#92400E", bold: true }, wrapText: true };

writeTable(proposals, "A1", [
  "Estrategia",
  "Relacion mayor/menor indicada",
  "Relacion actual con archivos",
  "Relacion elegible temporal",
  "Muestras seleccionadas",
  "Muestras elegibles temporales",
  "Min por familia",
  "Max por familia",
  "Comentario",
  "Uso propuesto",
], strategyRows.map((r) => [
  r.name,
  r.relationUser,
  r.relationCurrent,
  r.relationEligible,
  r.selectedTotal,
  r.selectedEligible,
  r.minSelected,
  r.maxSelected,
  r.comment,
  r.recommendation,
]), "TablaEstrategias");
proposals.getRange("I:J").format.wrapText = true;
proposals.getRange("A:J").format.columnWidthPx = 160;
proposals.getRange("I:J").format.columnWidthPx = 310;
writeTable(proposals, "A9", [
  "Estrategia",
  "Familia",
  "Seleccion total",
  "Seleccion elegible",
  "No temporal/reserva",
  "% elegible dentro seleccion",
], impactRows, "ImpactoPorFamilia");
proposals.getRange(`F10:F${impactRows.length + 9}`).format.numberFormat = "0.0%";

writeTable(temporal, "A1", [
  "Familia",
  "Total",
  "Elegibles temporales",
  "Fuera de rango/vacia",
  "% elegible",
  "Creacion min valida",
  "Creacion max valida",
  "Escaneo min",
  "Escaneo max",
  "Deteccion prom.",
  "Deteccion mediana",
], familySummary.map((r) => [
  r.family,
  r.total,
  r.temporalEligible,
  r.temporalOutlier,
  r.eligiblePct,
  r.creationMin,
  r.creationMax,
  r.scanMin,
  r.scanMax,
  round(r.detectionAvg, 2),
  round(r.detectionMedian, 2),
]), "CoberturaTemporal");
temporal.getRange("E2:E20").format.numberFormat = "0.0%";
temporal.getRange("J2:K20").format.numberFormat = "0.00";
temporal.getRange("A:K").format.columnWidthPx = 150;
writeTable(temporal, "A13", ["Ano creacion", "Muestras elegibles"], yearRows, "DistribucionAnual");

writeTable(splits, "A1", [
  "Estrategia",
  "Pestana",
  "Elegibles usados",
  "Train total",
  "Validacion total",
  "Test total",
  "Relacion elegible",
  "Rango anos propuesto",
  "Uso propuesto",
], strategyRows.map((r) => {
  const detailRows = splitRowsByStrategy.get(r.name);
  const trainYears = combineYearRanges(detailRows, 5);
  const valYears = combineYearRanges(detailRows, 6);
  const testYears = combineYearRanges(detailRows, 7);
  return [
    r.name,
    r.sheetName,
    r.selectedEligible,
    detailRows.reduce((sum, row) => sum + row[2], 0),
    detailRows.reduce((sum, row) => sum + row[3], 0),
    detailRows.reduce((sum, row) => sum + row[4], 0),
    r.relationEligible,
    `${trainYears} / ${valYears} / ${testYears} (ver pestana)`,
    r.recommendation,
  ];
}), "IndiceSplits");
splits.getRange("A:I").format.columnWidthPx = 165;
splits.getRange("H:I").format.columnWidthPx = 300;
splits.getRange("H:I").format.wrapText = true;

for (const strategy of strategies) {
  const sheet = splitDetailSheets.get(strategy.name);
  const row = strategyRows.find((r) => r.name === strategy.name);
  const detailRows = splitRowsByStrategy.get(strategy.name);
  const trainTotal = detailRows.reduce((sum, r) => sum + r[2], 0);
  const valTotal = detailRows.reduce((sum, r) => sum + r[3], 0);
  const testTotal = detailRows.reduce((sum, r) => sum + r[4], 0);
  sheet.getRange("A1").values = [[`${strategy.name} - split temporal por familia`]];
  sheet.getRange("A1:M1").merge();
  sheet.getRange("A1").format = { fill: "#111827", font: { bold: true, color: "#FFFFFF", size: 16 } };
  sheet.getRange("A2:F2").values = [[
    `Elegibles usados: ${row.selectedEligible}`,
    `Relacion elegible: ${row.relationEligible}`,
    `Train total: ${trainTotal}`,
    `Validacion total: ${valTotal}`,
    `Test total: ${testTotal}`,
    strategy.mode,
  ]];
  sheet.getRange("A2:F2").format = { fill: "#F3F4F6", font: { bold: true, color: "#111827" }, wrapText: true };
  writeTable(sheet, "A4", [
    "Familia",
    "Elegibles usados",
    "Train 70% antiguo",
    "Validacion 15% intermedio",
    "Test 15% reciente",
    "Anios train",
    "Anios validacion",
    "Anios test",
    "Corte train hasta",
    "Corte validacion hasta",
    "Fechas train",
    "Fechas validacion",
    "Fechas test",
    "Regla",
  ], detailRows, strategy.tableName);
  sheet.getRange("A:N").format.columnWidthPx = 150;
  sheet.getRange("K:N").format.columnWidthPx = 280;
  sheet.getRange("K:N").format.wrapText = true;
  sheet.getRange("A14:E14").values = [[
    "Total",
    detailRows.reduce((sum, r) => sum + r[1], 0),
    trainTotal,
    valTotal,
    testTotal,
  ]];
  sheet.getRange("A14:E14").format = { fill: "#FEF3C7", font: { bold: true, color: "#92400E" } };
  sheet.getRange("Q4:T12").values = [
    ["Familia", "Train 70% antiguo", "Validacion 15% intermedio", "Test 15% reciente"],
    ...detailRows.map((r) => [r[0], r[2], r[3], r[4]]),
  ];
  const chart = sheet.charts.add("bar", sheet.getRange("Q4:T12"));
  chart.title = `${strategy.name}: train / validacion / test`;
  chart.hasLegend = true;
  chart.xAxis = { axisType: "textAxis" };
  chart.yAxis = { numberFormatCode: "#,##0" };
  chart.setPosition("P3", "W18");
}

notes.getRange("A1:C14").values = [
  ["Criterio", "Propuesta", "Motivo"],
  ["Eje temporal principal", "dia_creacion_archivo valido entre 1992 y 2026", "Representa mejor la antiguedad del artefacto que la fecha de escaneo VT."],
  ["Fallback temporal", "No usar fallback automatico para el split principal", "fecha_escaneo_vt mide recoleccion/escaneo y puede introducir sesgo de lote."],
  ["Fechas outlier/vacias", "Reservar como holdout, auditoria o analisis no temporal", "Evita meter muestras con tiempo incierto en train/test cronologico."],
  ["Orden del split", "Por familia, ordenar ascendente por dia_creacion_archivo", "Train queda con muestras antiguas; test con muestras recientes."],
  ["Anios propuestos", "Se calculan con los cortes acumulados 70/15/15 de cada familia y estrategia", "Dan una guia temporal antes de seleccionar hashes especificos."],
  ["Caps por familia", "Cuando una familia supera el cap, se usa muestreo temporal uniforme para proponer anos", "Evita que el cap se concentre solo en anos antiguos o recientes."],
  ["Proporcion inicial", "70% train, 15% validacion, 15% test", "Es simple, reproducible y conserva evaluacion futura."],
  ["Estrategia recomendada", "Limitar a 1500", "Compromiso entre balance y volumen; evita que fragtor/salgorea dominen."],
  ["Baseline alterno", "Balancear a 1000", "Sirve para medir si el modelo mejora por senal o por distribucion de clases."],
  ["Uso de todo", "Solo baseline descriptivo o entrenamiento con ponderacion", "El desbalance fuerte puede dominar el aprendizaje."],
  ["Cap 2000", "Segunda corrida si 1500 es estable", "Aporta volumen adicional con desbalance moderado."],
  ["Siguiente paso", "Elegir estrategia y generar lista de hashes por split", "Ahi ya seleccionamos muestras especificas."],
  ["Semilla/reproducibilidad", "Desempatar por hash_md5 ordenado", "Hace que el resultado sea determinista si hay fechas repetidas."],
];
notes.getRange("A1:C1").format = { fill: "#1F2937", font: { bold: true, color: "#FFFFFF" } };
notes.getRange("A:C").format.columnWidthPx = 285;
notes.getRange("B:C").format.wrapText = true;

for (const sheet of [dash, proposals, temporal, splits, notes, ...splitDetailSheets.values()]) {
  sheet.freezePanes.freezeRows(1);
}

await fs.mkdir(outputDir, { recursive: true });
for (const [sheetName, range] of Object.entries({
  Dashboard: "A1:Z45",
  Propuestas: "A1:J42",
  Temporalidad: "A1:K38",
  Splits_Propuestos: "A1:I8",
  Split_Todo: "A1:W18",
  Split_1000: "A1:W18",
  Split_1500: "A1:W18",
  Split_2000: "A1:W18",
  Criterios: "A1:C14",
})) {
  const preview = await wb.render({ sheetName, range, scale: 1, format: "png" });
  await fs.writeFile(path.join(outputDir, `propuesta_temporal_${sheetName.toLowerCase()}_preview.png`), new Uint8Array(await preview.arrayBuffer()));
}
const errors = await wb.inspect({
  kind: "match",
  searchTerm: "#REF!|#DIV/0!|#VALUE!|#NAME\\?|#N/A",
  options: { useRegex: true, maxResults: 300 },
  summary: "final formula error scan",
});
await fs.writeFile(path.join(outputDir, "propuesta_temporal_formula_scan.ndjson"), errors.ndjson);
const output = await SpreadsheetFile.exportXlsx(wb);
let savedPath = outputPath;
try {
  await output.save(outputPath);
} catch (error) {
  if (error?.code !== "EBUSY") throw error;
  savedPath = path.join(outputDir, "propuesta_estructuracion_temporal_tablas.xlsx");
  await output.save(savedPath);
}
console.log(JSON.stringify({
  outputPath: savedPath,
  savedAsFallback: savedPath !== outputPath,
  totalRows,
  totalEligible,
  totalOutlier,
  currentRatio: `${round(maxTotal / minTotal, 2)} : 1`,
  eligibleRatio: `${round(maxEligible / minEligible, 2)} : 1`,
  recommended: "Limitar a 1500 + split cronologico por familia 70/15/15",
}, null, 2));
