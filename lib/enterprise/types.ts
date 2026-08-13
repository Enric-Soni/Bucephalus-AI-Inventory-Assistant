export type CellValue = string | number | boolean | null;

export type DatasetRole =
  | "item_master"
  | "inventory"
  | "sales_history"
  | "movement_history"
  | "supply"
  | "fx"
  | "aging_reserve"
  | "ignore"
  | "unknown";

export type CanonicalField =
  | "sku"
  | "canonicalSku"
  | "description"
  | "category"
  | "aliases"
  | "baseUom"
  | "unitsPerCase"
  | "standardCost"
  | "currency"
  | "lifecycleStatus"
  | "leadTimeDays"
  | "minimumOrderQuantity"
  | "netRealizableValue"
  | "valuationMethod"
  | "sourceReorderPoint"
  | "sourceSafetyStock"
  | "location"
  | "inventoryStatus"
  | "onHand"
  | "reserved"
  | "qualityHold"
  | "damaged"
  | "uom"
  | "unitCost"
  | "asOfDate"
  | "batch"
  | "period"
  | "channel"
  | "grossUnits"
  | "returns"
  | "netDemand"
  | "supplyId"
  | "origin"
  | "destination"
  | "quantity"
  | "expectedDate"
  | "status"
  | "unitPrice"
  | "fxRate"
  | "rateDate"
  | "movementType"
  | "totalQuantity"
  | "totalValue"
  | "age0To30"
  | "age31To90"
  | "age91To180"
  | "age181To365"
  | "ageOver365"
  | "reserveRate"
  | "requiredReserve";

export type RawTable = {
  name: string;
  address: string;
  startRow: number;
  startColumn: number;
  values: CellValue[][];
};

export type RawWorksheet = {
  name: string;
  usedRangeAddress: string;
  startRow: number;
  startColumn: number;
  values: CellValue[][];
  formulas?: CellValue[][];
  tables: RawTable[];
};

export type WorkbookSnapshot = {
  workbookName?: string;
  worksheets: RawWorksheet[];
  scannedCells: number;
  truncated: boolean;
};

export type SourceRow = {
  sourceRow: number;
  values: CellValue[];
};

export type ColumnMapping = {
  field: CanonicalField;
  columnIndex: number;
  header: string;
  confidence: number;
  reason: string;
};

export type RoleScore = {
  role: Exclude<DatasetRole, "ignore" | "unknown">;
  confidence: number;
  evidence: string[];
};

export type DatasetCandidate = {
  id: string;
  sourceSheet: string;
  sourceTable?: string;
  sourceAddress: string;
  headerRow: number;
  startColumn: number;
  headers: string[];
  rows: SourceRow[];
  proposedRole: DatasetRole;
  roleConfidence: number;
  roleScores: RoleScore[];
  mappings: Partial<Record<CanonicalField, ColumnMapping>>;
  requiredFields: CanonicalField[];
  missingRequiredFields: CanonicalField[];
  projectionOf?: string;
  projectionLabel?: string;
  metadataDefaults?: DatasetDefaults;
};

export type MappingRule = {
  role: DatasetRole;
  normalizedHeaders: string[];
  mappings: Partial<Record<CanonicalField, string>>;
  defaults?: DatasetDefaults;
};

export type DuplicateResolution = "keep_all" | "exclude_repeats";

export type DatasetDefaults = {
  location?: string;
  destination?: string;
  uom?: string;
  currency?: string;
  asOfDate?: string;
};

export type ImportReviewState = {
  datasetRoles: Record<string, DatasetRole>;
  columnMappings: Record<string, Partial<Record<CanonicalField, number>>>;
  datasetDefaults?: Record<string, DatasetDefaults>;
  excludedDatasets: string[];
  excludedRows: Record<string, number[]>;
  skuOverrides: Record<string, string>;
  duplicateResolutions: Record<string, DuplicateResolution>;
  reportingCurrency: string;
  warningConfirmation: boolean;
  leadTimeDays: number;
  safetyStockDays: number;
};

export type SourceLineage = {
  sourceSystem: string;
  sourceSheet: string;
  sourceTable?: string;
  sourceRow: number;
  sourceValues: Record<string, CellValue>;
  transformations: string[];
};

export type CanonicalItemMaster = {
  canonicalSku: string;
  productDescription: string;
  category?: string;
  sourceAliases: string[];
  baseUnitOfMeasure: string;
  unitsPerCase?: number;
  standardCost?: number;
  currency: string;
  lifecycleStatus?: string;
  leadTimeDays?: number;
  minimumOrderQuantity?: number;
  netRealizableValue?: number;
  valuationMethod?: string;
  sourceReorderPoint?: number;
  sourceSafetyStock?: number;
  lineage: SourceLineage;
};

export type CanonicalInventoryPosition = {
  id: string;
  datasetId: string;
  sourceSystem: string;
  sourceSheet: string;
  sourceTable?: string;
  sourceRow: number;
  sourceSku: string;
  canonicalSku: string;
  location: string;
  inventoryStatus: string;
  onHandQuantity: number;
  reservedQuantity: number;
  qualityHoldQuantity: number;
  damagedQuantity: number;
  unitOfMeasure: string;
  normalizedBaseUnits: number;
  normalizedReservedUnits: number;
  normalizedQualityHoldUnits: number;
  normalizedDamagedUnits: number;
  unitCost?: number;
  currency: string;
  fxRate?: number;
  valueInReportingCurrency?: number;
  valueInUsd?: number;
  netRealizableValuePerBaseUnit?: number;
  netInventoryValueInReportingCurrency?: number;
  lcmReserveInReportingCurrency?: number;
  asOfDate?: string;
  lineage: SourceLineage;
};

export type CanonicalSalesHistory = {
  id: string;
  period: string;
  sourceSku: string;
  canonicalSku: string;
  locationOrChannel?: string;
  grossUnits: number;
  returns: number;
  netDemand: number;
  lineage: SourceLineage;
};

export type CanonicalSupply = {
  id: string;
  supplyIdentifier: string;
  supplyType: "purchase_order" | "transfer" | "unknown";
  sourceSku: string;
  canonicalSku: string;
  origin?: string;
  destination?: string;
  orderedOrInTransitQuantity: number;
  unitOfMeasure: string;
  normalizedBaseUnits: number;
  expectedDate?: string;
  status: string;
  unitPrice?: number;
  currency: string;
  fxRate?: number;
  commitmentValue?: number;
  lineage: SourceLineage;
};

export type CanonicalAgingReserve = {
  id: string;
  sourceSku: string;
  canonicalSku: string;
  totalQuantity?: number;
  totalValue?: number;
  age0To30?: number;
  age31To90?: number;
  age91To180?: number;
  age181To365?: number;
  ageOver365?: number;
  reserveRate?: number;
  requiredReserve: number;
  currency: string;
  lineage: SourceLineage;
};

export type FxRate = {
  currency: string;
  usdPerUnit: number;
  rateDate?: string;
  lineage: SourceLineage;
};

export type IssueSeverity = "blocker" | "warning" | "info";

export type ImportIssue = {
  id: string;
  severity: IssueSeverity;
  code: string;
  message: string;
  datasetId?: string;
  sourceSheet?: string;
  sourceRow?: number;
  sourceSku?: string;
};

export type QuarantinedRecord = {
  id: string;
  datasetId: string;
  role: DatasetRole;
  sourceSheet: string;
  sourceRow: number;
  reason: string;
  sourceValues: Record<string, CellValue>;
};

export type DuplicateGroup = {
  id: string;
  kind: "exact" | "potential";
  datasetId: string;
  businessKey: string;
  recordIds: string[];
  sourceRows: number[];
  reason: string;
};

export type ForecastMethod = "naive" | "moving_average" | "seasonal_naive" | "exponential_smoothing";

export type ForecastResult = {
  canonicalSku: string;
  selectedModel?: ForecastMethod;
  historyPeriods: number;
  forecastMonthlyDemand?: number;
  lowerPrediction?: number;
  upperPrediction?: number;
  errorMae?: number;
  insufficientHistory: boolean;
  explanation: string;
};

export type ConsolidatedInventory = {
  canonicalSku: string;
  productDescription: string;
  category?: string;
  location: string;
  grossOnHand: number;
  restrictedStock: number;
  netAvailable: number;
  inTransit: number;
  openPoQuantity: number;
  inventoryValue: number;
  netInventoryValue: number;
  lcmReserve: number;
  obsolescenceReserve: number;
  averageMonthlyDemand?: number;
  monthsOfCover?: number;
  forecastModel?: ForecastMethod;
  forecastMonthlyDemand?: number;
  predictionLower?: number;
  predictionUpper?: number;
  leadTimeDays: number;
  forecastDemandThroughLeadTime?: number;
  safetyStock: number;
  reorderPoint?: number;
  minimumOrderQuantity?: number;
  suggestedOrder?: number;
  reorderPolicySource: "forecast" | "source_policy" | "none";
  dataQuality: string[];
};

export type NormalizationResult = {
  items: CanonicalItemMaster[];
  inventory: CanonicalInventoryPosition[];
  sales: CanonicalSalesHistory[];
  supply: CanonicalSupply[];
  agingReserves: CanonicalAgingReserve[];
  fxRates: FxRate[];
  issues: ImportIssue[];
  quarantined: QuarantinedRecord[];
  duplicates: DuplicateGroup[];
  includedRecordCount: number;
  excludedRecordCount: number;
  normalizedUnitTotal: number;
  normalizedValueTotal: number;
  normalizedNetValueTotal: number;
  normalizedLcmReserveTotal: number;
  normalizedObsolescenceReserveTotal: number;
};

export type EnterpriseAnalysis = {
  normalization: NormalizationResult;
  forecasts: ForecastResult[];
  consolidated: ConsolidatedInventory[];
};
