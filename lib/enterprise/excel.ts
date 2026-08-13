import { EnterpriseAnalysis, RawTable, RawWorksheet, WorkbookSnapshot } from "./types";
import { WORKBOOK_LIMITS } from "./discovery";

function excelValues(values: unknown[][]) {
  return values.map((row) => row.map((value) => {
    if (value === undefined || value === null) return null;
    if (typeof value === "number" || typeof value === "boolean" || typeof value === "string") return value;
    return String(value);
  }));
}

export async function captureWorkbookSnapshot(): Promise<WorkbookSnapshot> {
  if (!window.Excel) throw new Error("Open this page from the Bucephalus task pane in Excel.");
  return Excel.run(async (context) => {
    const worksheets = context.workbook.worksheets;
    worksheets.load("items/name");
    await context.sync();
    if (worksheets.items.length > WORKBOOK_LIMITS.maxWorksheets) {
      throw new Error(`Workbook has ${worksheets.items.length} worksheets; the safety limit is ${WORKBOOK_LIMITS.maxWorksheets}.`);
    }

    const ranges = worksheets.items.map((sheet) => {
      const usedRange = sheet.getUsedRangeOrNullObject(true);
      const tables = sheet.tables;
      usedRange.load("isNullObject,address,rowIndex,columnIndex,rowCount,columnCount");
      tables.load("items/name");
      return { sheet, usedRange, tables };
    });
    await context.sync();

    let scannedCells = 0;
    let truncated = false;
    const reads = ranges.map(({ sheet, usedRange, tables }) => {
      const rowCount = usedRange.isNullObject ? 0 : Math.min(usedRange.rowCount, WORKBOOK_LIMITS.maxRowsPerSheet);
      const columnCount = usedRange.isNullObject ? 0 : Math.min(usedRange.columnCount, WORKBOOK_LIMITS.maxColumnsPerSheet);
      if (!usedRange.isNullObject && (rowCount !== usedRange.rowCount || columnCount !== usedRange.columnCount)) truncated = true;
      scannedCells += rowCount * columnCount;
      if (scannedCells > WORKBOOK_LIMITS.maxCells) {
        throw new Error(`Workbook exceeds the ${WORKBOOK_LIMITS.maxCells.toLocaleString()}-cell safety limit. Narrow the used ranges or exclude oversized raw-data sheets.`);
      }
      const dataRange = rowCount && columnCount
        ? sheet.getRangeByIndexes(usedRange.rowIndex, usedRange.columnIndex, rowCount, columnCount)
        : null;
      dataRange?.load("address,rowIndex,columnIndex,rowCount,columnCount,values,formulas");
      const tableRanges = tables.items.map((table) => {
        const range = table.getRange();
        range.load("address,rowIndex,columnIndex,rowCount,columnCount,values");
        return { table, range };
      });
      return { sheet, usedRange, dataRange, tableRanges };
    });
    await context.sync();

    const rawWorksheets: RawWorksheet[] = reads.map(({ sheet, usedRange, dataRange, tableRanges }) => ({
      name: sheet.name,
      usedRangeAddress: usedRange.isNullObject ? "" : usedRange.address,
      startRow: dataRange?.rowIndex ?? 0,
      startColumn: dataRange?.columnIndex ?? 0,
      values: dataRange ? excelValues(dataRange.values) : [],
      formulas: dataRange ? excelValues(dataRange.formulas) : [],
      tables: tableRanges.map(({ table, range }): RawTable => ({
        name: table.name,
        address: range.address,
        startRow: range.rowIndex,
        startColumn: range.columnIndex,
        values: excelValues(range.values)
      }))
    }));
    return { worksheets: rawWorksheets, scannedCells, truncated };
  });
}

type OutputSheet = {
  name: string;
  headers: string[];
  rows: Array<Array<string | number | boolean | null>>;
  numberFormats?: Record<number, string>;
  blurb: string;
};

function makeOutputSheets(analysis: EnterpriseAnalysis): OutputSheet[] {
  const normalizedItems: OutputSheet = {
    name: "Buc Normalized Items",
    headers: ["Canonical SKU", "Product", "Category", "Source Aliases", "Base UOM", "Units per Case", "Standard Cost", "Currency", "Lifecycle", "Lead Time Days", "Minimum Order Quantity", "NRV per Base Unit", "Valuation Method", "Source Reorder Point", "Source Safety Stock", "Source Sheet", "Source Row", "Transformations"],
    rows: analysis.normalization.items.map((record) => [
      record.canonicalSku, record.productDescription, record.category ?? null, record.sourceAliases.join(" | "),
      record.baseUnitOfMeasure, record.unitsPerCase ?? null, record.standardCost ?? null, record.currency,
      record.lifecycleStatus ?? null, record.leadTimeDays ?? null, record.minimumOrderQuantity ?? null,
      record.netRealizableValue ?? null, record.valuationMethod ?? null, record.sourceReorderPoint ?? null,
      record.sourceSafetyStock ?? null, record.lineage.sourceSheet, record.lineage.sourceRow, record.lineage.transformations.join(" | ")
    ]),
    numberFormats: { 5: "#,##0", 6: "#,##0.00", 9: "#,##0", 10: "#,##0", 11: "#,##0.00", 13: "#,##0", 14: "#,##0" },
    blurb: "ABOUT THIS SHEET — The standardized product catalog used by Bucephalus. It connects source-system aliases to canonical SKUs and supplies the SKU-specific units-per-case, cost, currency, lifecycle, lead time, minimum order quantity, and source lineage used during normalization."
  };
  const normalizedInventory: OutputSheet = {
    name: "Buc Normalized Inventory",
    headers: [
      "Canonical SKU", "Source SKU", "Location", "Inventory Status", "Gross Base Units", "Reserved Base Units",
      "Quality Hold Base Units", "Damaged Base Units", "Net Available Base Units", "Source UOM", "Unit Cost",
      "Currency", "FX Rate", "Gross Cost Value", "Net LCM Value", "LCM Reserve", "Value USD", "As-of Date", "Source Sheet", "Source Table", "Source Row", "Transformations"
    ],
    rows: analysis.normalization.inventory.map((record) => [
      record.canonicalSku, record.sourceSku, record.location, record.inventoryStatus, record.normalizedBaseUnits,
      record.normalizedReservedUnits, record.normalizedQualityHoldUnits, record.normalizedDamagedUnits,
      record.normalizedBaseUnits - record.normalizedReservedUnits - record.normalizedQualityHoldUnits - record.normalizedDamagedUnits,
      record.unitOfMeasure, record.unitCost ?? null, record.currency, record.fxRate ?? null,
      record.valueInReportingCurrency ?? null, record.netInventoryValueInReportingCurrency ?? null,
      record.lcmReserveInReportingCurrency ?? null, record.valueInUsd ?? null, record.asOfDate ?? null,
      record.sourceSheet, record.sourceTable ?? null, record.sourceRow, record.lineage.transformations.join(" | ")
    ]),
    numberFormats: { 4: "#,##0", 5: "#,##0", 6: "#,##0", 7: "#,##0", 8: "#,##0", 10: "#,##0.00", 12: "0.0000", 13: "#,##0.00", 14: "#,##0.00", 15: "#,##0.00", 16: "#,##0.00" },
    blurb: "ABOUT THIS SHEET — The cleaned inventory-position records discovered across the workbook. Quantities are normalized to base units, restrictions remain separate, values are currency-converted, and every row retains its original source sheet, row, values, and transformations for audit purposes."
  };
  const normalizedSales: OutputSheet = {
    name: "Buc Normalized Sales",
    headers: ["Canonical SKU", "Source SKU", "Period", "Location / Channel", "Gross Units", "Returns", "Net Demand", "Source Sheet", "Source Row", "Transformations"],
    rows: analysis.normalization.sales.map((record) => [
      record.canonicalSku, record.sourceSku, record.period, record.locationOrChannel ?? null, record.grossUnits,
      record.returns, record.netDemand, record.lineage.sourceSheet, record.lineage.sourceRow, record.lineage.transformations.join(" | ")
    ]),
    numberFormats: { 4: "#,##0", 5: "#,##0", 6: "#,##0" },
    blurb: "ABOUT THIS SHEET — The normalized sales and returns history used for deterministic forecast backtesting. Negative returns are preserved as valid activity, and net demand is grouped by canonical SKU and period before model selection."
  };
  const normalizedSupply: OutputSheet = {
    name: "Buc Normalized Supply",
    headers: ["Type", "Identifier", "Canonical SKU", "Source SKU", "Origin", "Destination", "Base Units", "Expected Date", "Status", "Unit Price", "Currency", "Commitment Value", "Source Sheet", "Source Row", "Transformations"],
    rows: analysis.normalization.supply.map((record) => [
      record.supplyType, record.supplyIdentifier, record.canonicalSku, record.sourceSku, record.origin ?? null,
      record.destination ?? null, record.normalizedBaseUnits, record.expectedDate ?? null, record.status,
      record.unitPrice ?? null, record.currency, record.commitmentValue ?? null, record.lineage.sourceSheet,
      record.lineage.sourceRow, record.lineage.transformations.join(" | ")
    ]),
    numberFormats: { 6: "#,##0", 9: "#,##0.00", 11: "#,##0.00" },
    blurb: "ABOUT THIS SHEET — The standardized purchase orders and transfers detected in the workbook. It shows canonical SKUs, origins, destinations, expected dates, statuses, normalized quantities, commitment values, and row-level source lineage."
  };
  const normalizedAging: OutputSheet = {
    name: "Buc Normalized Aging",
    headers: ["Canonical SKU", "Source SKU", "Total Quantity", "Total Value", "0-30 Days", "31-90 Days", "91-180 Days", "181-365 Days", ">365 Days", "Reserve Rate", "Required Reserve", "Currency", "Source Sheet", "Source Row", "Transformations"],
    rows: analysis.normalization.agingReserves.map((record) => [
      record.canonicalSku, record.sourceSku, record.totalQuantity ?? null, record.totalValue ?? null,
      record.age0To30 ?? null, record.age31To90 ?? null, record.age91To180 ?? null,
      record.age181To365 ?? null, record.ageOver365 ?? null, record.reserveRate ?? null,
      record.requiredReserve, record.currency, record.lineage.sourceSheet, record.lineage.sourceRow,
      record.lineage.transformations.join(" | ")
    ]),
    numberFormats: { 2: "#,##0", 3: "#,##0.00", 4: "#,##0", 5: "#,##0", 6: "#,##0", 7: "#,##0", 8: "#,##0", 9: "0.0%", 10: "#,##0.00" },
    blurb: "ABOUT THIS SHEET — The reviewed aging and obsolescence-reserve records. Bucephalus preserves the source reserve, validates aging-bucket arithmetic, flags negative or unreconciled buckets, and links the reserve to the canonical SKU without treating this derived report as physical stock."
  };
  const forecast: OutputSheet = {
    name: "Buc Forecast Analysis",
    headers: [
      "Canonical SKU", "Product", "Category", "Location", "Gross On Hand", "Restricted", "Net Available", "In Transit", "Open PO",
      "Gross Cost Value", "Net LCM Value", "LCM Reserve", "Obsolescence Reserve", "Months Cover", "Selected Forecast", "Monthly Forecast", "Prediction Low", "Prediction High", "Lead Time Days",
      "Demand Through Lead Time", "Safety Stock", "Reorder Point", "Reorder Policy Source", "Minimum Order Quantity", "Suggested Order", "Data Quality"
    ],
    rows: analysis.consolidated.map((record) => [
      record.canonicalSku, record.productDescription, record.category ?? null, record.location, record.grossOnHand,
      record.restrictedStock, record.netAvailable, record.inTransit, record.openPoQuantity, record.inventoryValue,
      record.netInventoryValue, record.lcmReserve, record.obsolescenceReserve, record.monthsOfCover ?? null,
      record.forecastModel ?? "Insufficient history", record.forecastMonthlyDemand ?? null,
      record.predictionLower ?? null, record.predictionUpper ?? null, record.leadTimeDays,
      record.forecastDemandThroughLeadTime ?? null, record.safetyStock, record.reorderPoint ?? null,
      record.reorderPolicySource, record.minimumOrderQuantity ?? null, record.suggestedOrder ?? null, record.dataQuality.join(" | ")
    ]),
    numberFormats: {
      4: "#,##0", 5: "#,##0", 6: "#,##0", 7: "#,##0", 8: "#,##0", 9: "#,##0.00",
      10: "#,##0.00", 11: "#,##0.00", 12: "#,##0.00", 13: "0.0", 15: "#,##0.0", 16: "#,##0.0",
      17: "#,##0.0", 18: "#,##0", 19: "#,##0.0", 20: "#,##0.0", 21: "#,##0.0", 23: "#,##0", 24: "#,##0"
    },
    blurb: "ABOUT THIS SHEET — The main decision-support output. Each row represents a canonical SKU and location, combining available and restricted stock, inbound supply, inventory value, forecast coverage, prediction ranges, lead-time demand, safety stock, reorder point, and suggested order. Review the Data Quality column before acting."
  };
  const exceptions: OutputSheet = {
    name: "Buc Import Exceptions",
    headers: ["Severity", "Code", "Message", "Source Sheet", "Source Row", "Source SKU"],
    rows: [
      ...analysis.normalization.issues.map((issue) => [issue.severity, issue.code, issue.message, issue.sourceSheet ?? null, issue.sourceRow ?? null, issue.sourceSku ?? null]),
      ...analysis.normalization.quarantined.map((record) => ["quarantined", "QUARANTINED", record.reason, record.sourceSheet, record.sourceRow, null])
    ],
    blurb: "ABOUT THIS SHEET — The import audit report. Blockers identify records that could not be safely normalized; warnings document approved limitations or fallbacks; quarantined rows were excluded from analysis. Use the source sheet and row columns to investigate the original data."
  };
  return [normalizedItems, normalizedInventory, normalizedSales, normalizedSupply, normalizedAging, forecast, exceptions];
}

export async function writeEnterpriseOutputs(analysis: EnterpriseAnalysis): Promise<string[]> {
  const outputs = makeOutputSheets(analysis);
  if (outputs.some((output) => output.rows.length > WORKBOOK_LIMITS.maxRowsPerSheet)) {
    throw new Error(`An output exceeds the ${WORKBOOK_LIMITS.maxRowsPerSheet.toLocaleString()}-row sheet limit.`);
  }
  return Excel.run(async (context) => {
    const worksheets = context.workbook.worksheets;
    worksheets.load("items/name");
    await context.sync();
    const usedNames = new Set(worksheets.items.map((sheet) => sheet.name));
    const created: string[] = [];
    for (const output of outputs) {
      let name = output.name.slice(0, 31);
      let suffix = 2;
      while (usedNames.has(name)) {
        const ending = ` ${suffix++}`;
        name = `${output.name.slice(0, 31 - ending.length)}${ending}`;
      }
      usedNames.add(name);
      created.push(name);
      const sheet = worksheets.add(name);
      const values = [output.headers, ...output.rows];
      const range = sheet.getRangeByIndexes(0, 0, values.length, output.headers.length);
      range.values = values;
      range.format.autofitColumns();
      range.format.autofitRows();
      range.format.wrapText = false;
      const header = sheet.getRangeByIndexes(0, 0, 1, output.headers.length);
      header.format.fill.color = "#1D135F";
      header.format.font.color = "#FFFFFF";
      header.format.font.bold = true;
      header.format.wrapText = true;
      for (const [column, format] of Object.entries(output.numberFormats ?? {})) {
        if (!output.rows.length) continue;
        sheet.getRangeByIndexes(1, Number(column), output.rows.length, 1).numberFormat = output.rows.map(() => [format]);
      }
      const noteRange = sheet.getRangeByIndexes(output.rows.length + 2, 0, 2, output.headers.length);
      noteRange.merge();
      noteRange.getCell(0, 0).values = [[output.blurb]];
      noteRange.format.fill.color = "#F8EFE2";
      noteRange.format.font.color = "#312E81";
      noteRange.format.font.italic = true;
      noteRange.format.wrapText = true;
      noteRange.format.verticalAlignment = Excel.VerticalAlignment.center;
      noteRange.format.rowHeight = 28;
      sheet.freezePanes.freezeRows(1);
      if (output.rows.length) {
        const tableName = `Buc${created.length}${Date.now().toString().slice(-6)}`;
        const table = sheet.tables.add(range, true);
        table.name = tableName;
        table.style = "TableStyleMedium2";
      }
    }
    worksheets.getItem(created.at(-2) ?? created[0]).activate();
    await context.sync();
    return created;
  });
}
