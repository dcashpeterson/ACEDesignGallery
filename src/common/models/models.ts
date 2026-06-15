import { ChoiceFieldFormatType } from "@pnp/sp/fields";
import { statusType } from "./enums";

export interface IChoice {
  choice: string;
  value: string;
}
export class Choice implements IChoice {
  constructor(
    public choice: string = "",
    public value: string = ""
  ) { }
}

export interface IFieldList {
  internalName: string;
  displayName: string;
  props: { FieldTypeKind: number, choices?: string[], richText?: boolean, editFormat?: ChoiceFieldFormatType, minValue?: number, maxValue?: number, localID?: number };
}

export const LunchMenuFields: IFieldList[] = [
  { internalName: "Location", displayName: "Location", props: { FieldTypeKind: 2 } },
  { internalName: "Description", displayName: "Description", props: { FieldTypeKind: 3 } },
  { internalName: "Category", displayName: "Category", props: { FieldTypeKind: 6, choices: ["Entree", "Side", "Dessert", "Beverage", "Special"], editFormat: ChoiceFieldFormatType.Dropdown } },
  { internalName: "Price", displayName: "Price", props: { FieldTypeKind: 9 } },
  { internalName: "Image", displayName: "Image", props: { FieldTypeKind: 99 } }
];

export interface IServiceHealth {
  id: string;
  title: string;
  service: string;
  statusType: statusType;
  issues?: IServiceHealthIssue[];
}

export class ServiceHealth implements IServiceHealth {
  constructor(
    public id: string,
    public title: string = "",
    public service: string = "",
    public statusType: statusType = 0 as unknown as statusType,
    public issues?: IServiceHealthIssue[]
  ) { }
}

export interface ILocation {
  name: string;
  imageUrl: string;
}

export interface ILunchMenuItem {
  id: string;
  title: string;
  location: string;
  description: string;
  category: string;
  price: number;
  image: string;
}

export interface IServiceHealthIssue {
  id: string;
  title: string;
  service: string;
  statusType: statusType;
  classification: string;
  feature: string;
  featureGroup: string;
  impactDescription: string;
  isResolved: boolean;
  origin: string;
  startDateTime: string;
  endDateTime: string;
  lastModifiedDateTime: string;
}

export interface ISiteActivityStat {
  label: string;
  date: string;
  viewCount: number;
  visitorCount: number;
}

export interface ISiteAnalyticsSummary {
  last7Views: number;
  last7Visitors: number;
  allTimeViews: number;
  allTimeVisitors: number;
  dailyStats: ISiteActivityStat[];
}

export class SiteAnalyticsSummary implements ISiteAnalyticsSummary{
  constructor(
    public last7Views: number = 0,
    public last7Visitors: number = 0,
    public allTimeViews: number = 0,
    public allTimeVisitors: number = 0,
    public dailyStats: ISiteActivityStat[] = []
  ) { }
}

export interface ISupportTicket {
  id: string;
  title: string;
  status: string;
  priority: string;
  assignedTo: string;
  description: string;
  submittedDate: string;
}

export interface ITicketStatusCounts {
  open: number;
  inProgress: number;
  resolved: number;
  escalated: number;
}

export const SupportTicketFields: IFieldList[] = [
  { internalName: "Status", displayName: "Status", props: { FieldTypeKind: 6, choices: ["Open", "In Progress", "Resolved", "Escalated"], editFormat: ChoiceFieldFormatType.Dropdown } },
  { internalName: "Priority", displayName: "Priority", props: { FieldTypeKind: 6, choices: ["Low", "Medium", "High", "Critical"], editFormat: ChoiceFieldFormatType.Dropdown } },
  { internalName: "AssignedTo", displayName: "Assigned To", props: { FieldTypeKind: 20 } },
  { internalName: "Description", displayName: "Description", props: { FieldTypeKind: 3, richText: false } },
  { internalName: "SubmittedDate", displayName: "Submitted Date", props: { FieldTypeKind: 4 } }
];

export interface ISalesDataItem {
  id: string;
  title: string;
  saleDate: string;
  amount: number;
  salesperson: string;
  region: string;
  product: string;
}

export interface ISalesDataPoint {
  date: Date;
  amount: number;
}

export interface ISalesDataSeries {
  name: string;
  color: string;
  dataPoints: ISalesDataPoint[];
  total: number;
}

export const SalesDataFields: IFieldList[] = [
  { internalName: "SaleDate",    displayName: "Sale Date",   props: { FieldTypeKind: 4 } },
  { internalName: "Amount",      displayName: "Amount",      props: { FieldTypeKind: 9 } },
  { internalName: "Salesperson", displayName: "Salesperson", props: { FieldTypeKind: 2 } },
  { internalName: "Region",      displayName: "Region",      props: { FieldTypeKind: 6, choices: ["North", "South", "East", "West", "Central"], editFormat: ChoiceFieldFormatType.Dropdown } },
  { internalName: "Product",     displayName: "Product",     props: { FieldTypeKind: 6, choices: ["Software", "Hardware", "Services", "Support", "Training"], editFormat: ChoiceFieldFormatType.Dropdown } }
];

export interface IEmployeeSpotlight {
  id: string;
  employeeName: string;
  spotlightType: string;
  description: string;
  employeeImageUrl: string;
  spotlightDate: string;
  spotlightDateFormatted: string;
  fullArticle: string;
  pageUrl: string;
}

export const EmployeeSpotlightFields: IFieldList[] = [
  { internalName: "SpotlightEmployeeName",  displayName: "Employee Name",   props: { FieldTypeKind: 2 } },
  { internalName: "SpotlightType",          displayName: "Spotlight Type",  props: { FieldTypeKind: 6, choices: ["Promotion", "Retirement", "New Baby", "Award", "Work Anniversary"], editFormat: ChoiceFieldFormatType.Dropdown } },
  { internalName: "SpotlightEmployeeImage", displayName: "Employee Image",  props: { FieldTypeKind: 11 } },
  { internalName: "SpotlightDate",          displayName: "Spotlight Date",  props: { FieldTypeKind: 4 } },
  { internalName: "SpotlightArticle",       displayName: "Article",         props: { FieldTypeKind: 3, richText: true } }
];

export interface ICompanyAnnouncement {
  id: string;
  title: string;
  body: string;
  publishedDate: string;
  publishedDateFormatted: string;
  category: string;
}

export const CompanyAnnouncementsFields: IFieldList[] = [
  { internalName: "AnnouncementBody", displayName: "Body",           props: { FieldTypeKind: 3, richText: false } },
  { internalName: "Category",         displayName: "Category",       props: { FieldTypeKind: 6, choices: ["General", "HR", "IT", "Operations", "Leadership"], editFormat: ChoiceFieldFormatType.Dropdown } },
  { internalName: "PublishedDate",    displayName: "Published Date", props: { FieldTypeKind: 4 } }
];

export interface ISafetyTip {
  id: string;
  title: string;
  summary: string;
  guidance: string;
  category: string;
  policyUrl: string;
  policyTitle: string;
}

export const SafetyTipsFields: IFieldList[] = [
  { internalName: "Summary",     displayName: "Summary",               props: { FieldTypeKind: 2 } },
  { internalName: "Guidance",    displayName: "Guidance",              props: { FieldTypeKind: 3, richText: false } },
  { internalName: "Category",    displayName: "Category",              props: { FieldTypeKind: 6, choices: ["Ergonomics", "Fire Safety", "Chemical Safety", "Emergency Preparedness", "General Safety", "Health & Hygiene"], editFormat: ChoiceFieldFormatType.Dropdown } },
  { internalName: "PolicyUrl",   displayName: "Policy Document URL",   props: { FieldTypeKind: 11 } },
  { internalName: "PolicyTitle", displayName: "Policy Document Title", props: { FieldTypeKind: 2 } }
];

export interface IVacationRequest {
  id: string;
  title: string;
  startDate: string;
  endDate: string;
  days: number;
  status: string;
  notes: string;
}

export interface IVacationSummary {
  used: number;
  available: number;
  total: number;
}

export const VacationRequestFields: IFieldList[] = [
  { internalName: "StartDate", displayName: "Start Date",       props: { FieldTypeKind: 4 } },
  { internalName: "EndDate",   displayName: "End Date",         props: { FieldTypeKind: 4 } },
  { internalName: "Days",      displayName: "Number of Days",   props: { FieldTypeKind: 9 } },
  { internalName: "Status",    displayName: "Status",           props: { FieldTypeKind: 6, choices: ["Submitted", "Waiting for Approval", "Approved"], editFormat: ChoiceFieldFormatType.Dropdown } },
  { internalName: "Notes",     displayName: "Notes",            props: { FieldTypeKind: 3, richText: false } }
];

export interface IInventoryItem {
  id: string;
  title: string;
  sku: string;
  category: string;
  quantity: number;
  location: string;
  description: string;
  unitPrice: number;
}

export const InventoryItemFields: IFieldList[] = [
  { internalName: "SKU",         displayName: "SKU",         props: { FieldTypeKind: 2 } },
  { internalName: "Category",    displayName: "Category",    props: { FieldTypeKind: 6, choices: ["Electronics", "Tools", "Office Supplies", "Hardware", "Safety Equipment"], editFormat: ChoiceFieldFormatType.Dropdown } },
  { internalName: "Quantity",    displayName: "Quantity",    props: { FieldTypeKind: 9, minValue: 0 } },
  { internalName: "Location",    displayName: "Location",    props: { FieldTypeKind: 2 } },
  { internalName: "Description", displayName: "Description", props: { FieldTypeKind: 3, richText: false } },
  { internalName: "UnitPrice",   displayName: "Unit Price",  props: { FieldTypeKind: 10, minValue: 0, localID: 1033 } }
];

export interface IWeatherLocation {
  id: string;
  name: string;
  latitude: number;
  longitude: number;
}

export interface IWeatherData {
  temperature: number;
  weatherCode: number;
  windspeed: number;
  isDay: boolean;
  description: string;
  iconUrl: string;
}

export const WeatherLocationFields: IFieldList[] = [
  { internalName: "LocationName", displayName: "Location Name", props: { FieldTypeKind: 2 } },
  { internalName: "Latitude",     displayName: "Latitude",      props: { FieldTypeKind: 9 } },
  { internalName: "Longitude",    displayName: "Longitude",     props: { FieldTypeKind: 9 } }
];

export interface IFacilitiesMaintenanceRequest {
  id: string;
  title: string;
  description: string;
  category: string;
  location: string;
  priority: string;
  status: string;
  requestedByName: string;
  requestedById: number;
  requestedDate: string;
  assignedToName: string;
  assignedToId: number;
  maintenanceNotes: string;
  maintenanceUpdatedDate: string;
  assignedManagerName: string;
  assignedManagerId: number;
  managerNotes: string;
  completedDate: string;
}

export interface IRequestFormDraft {
  title: string;
  category: string;
  location: string;
  priority: string;
  description: string;
  requestedById: number;
  requestedByName: string;
}

// Stock Ticker
export type StockTimeRange = IStockTickerCardAdaptiveCardExtensionStrings['Range7DShort'] | IStockTickerCardAdaptiveCardExtensionStrings['Range30DShort'] | IStockTickerCardAdaptiveCardExtensionStrings['Range1YShort'];

export interface IStockDataPoint {
  date: Date;
  price: number;
}

export interface IStockQuote {
  symbol: string;
  price: number;
  change: number;
  changePercent: number;
  open: number;
  high: number;
  low: number;
  previousClose: number;
  lastUpdated: Date;
}

export const FacilitiesMaintenanceFields: IFieldList[] = [
  { internalName: "Description",            displayName: "Description",              props: { FieldTypeKind: 3, richText: false } },
  { internalName: "Category",               displayName: "Category",                 props: { FieldTypeKind: 6, choices: ["HVAC", "Electrical", "Plumbing", "Structural", "Cleaning", "Safety Hazard", "IT/Network", "Other"], editFormat: ChoiceFieldFormatType.Dropdown } },
  { internalName: "FacilityLocation",       displayName: "Location",                 props: { FieldTypeKind: 2 } },
  { internalName: "Priority",               displayName: "Priority",                 props: { FieldTypeKind: 6, choices: ["Low", "Medium", "High", "Critical"], editFormat: ChoiceFieldFormatType.Dropdown } },
  { internalName: "RequestStatus",          displayName: "Status",                   props: { FieldTypeKind: 6, choices: ["New", "In Progress", "Resolved - Pending Validation", "Completed", "Cancelled"], editFormat: ChoiceFieldFormatType.Dropdown } },
  { internalName: "RequestedBy",            displayName: "Requested By",             props: { FieldTypeKind: 20 } },
  { internalName: "RequestedDate",          displayName: "Requested Date",           props: { FieldTypeKind: 4 } },
  { internalName: "AssignedMaintenanceTo",  displayName: "Assigned To",              props: { FieldTypeKind: 20 } },
  { internalName: "MaintenanceNotes",       displayName: "Maintenance Notes",        props: { FieldTypeKind: 3, richText: false } },
  { internalName: "MaintenanceUpdatedDate", displayName: "Maintenance Updated Date", props: { FieldTypeKind: 4 } },
  { internalName: "AssignedManager",        displayName: "Assigned Manager",         props: { FieldTypeKind: 20 } },
  { internalName: "ManagerNotes",           displayName: "Manager Notes",            props: { FieldTypeKind: 3, richText: false } },
  { internalName: "CompletedDate",          displayName: "Completed Date",           props: { FieldTypeKind: 4 } }
];

