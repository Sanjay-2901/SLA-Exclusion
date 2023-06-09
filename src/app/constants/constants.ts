import * as ExcelJS from 'exceljs';

export const SHEET_HEADING = {
  fill: { type: 'pattern', pattern: 'solid', fgColor: { argb: '6495ED' } },
  font: { color: { argb: 'FFFFFF' }, bold: true },
} as ExcelJS.Style;

export const TABLE_HEADING = {
  fill: { type: 'pattern', pattern: 'solid', fgColor: { argb: 'ffffcc00' } },
  font: { color: { argb: 'FFFFFF' }, bold: true },
} as ExcelJS.Style;

export const BlockSLASummaryPercentHeaders = [
  'Report type',
  'Time span',
  'No of Blocks',
  'UP (%)',
  'No. of UP Blocks',
  'Power Down (%)',
  'Fibre Down (%)',
  'Equipment Down (%)',
  'HRT Down (%)',
  'DCN Down (%)',
  'Planned Maintanance (%)',
  'Down (%) (Exclusive of SLA Exclusions)',
  'No. of Down Blocks',
  'Total SLA Exclusion (%)',
  'Total UP (%) (UP % + SLA Exclusion %)',
];

export const BlockSLASummaryMinutesHeaders = [
  'Report type',
  'Time span',
  'No of Blocks',
  'UP (min)',
  'Power Down (min)',
  'Fibre Down (min)',
  'Equipment Down (min)',
  'HRT Down (min)',
  'DCN Down (min)',
  'Planned Maintanance (min)',
  'Down (min)',
  'Total SLA Exclusion (min)',
  'Total UP (min) (UP (min) + SLA Exclusion (min))',
];

export const BLOCK_GP_SLA_SUMMARY_PERCENT = [
  'Report type',
  'Time Span',
  'IP Address',
  'State',
  'Cluster',
  'District',
  'District LGD Code',
  'Block Name',
  'Block LGD Code',
  'No of GPs in a Block',
  'UP (%)',
  'No. of UP GPs Count',
  'Power Down (%)',
  'Fibre Down (%)',
  'Equipment Down (%)',
  'HRT Down (%)',
  'DCN Down (%)',
  'Planned Maintenance (%)',
  'Down %  (Exclusive of SLA Exclusions)',
  'No. of Down GPs Count',
  'Total Exclusion (%)',
  'Polling Time %',
  'Total UP ( %) ((UP (%) + SLA Exclusion (%))',
];

export const BLOCK_SLA_FINAL_REPORT_COLUMNS = [
  { width: 20 },
  { width: 20 },
  { width: 20 },
  { width: 20 },
  { width: 20 },
  { width: 20 },
  { width: 20 },
  { width: 20 },
  { width: 20 },
  { width: 20 },
  { width: 20 },
  { width: 20 },
  { width: 20 },
  { width: 20 },
  { width: 20 },
  { width: 20 },
  { width: 20 },
  { width: 20 },
  { width: 30 },
  { width: 30 },
  { width: 30 },
  { width: 30 },
  { width: 30 },
  { width: 40 },
];

export interface BlockNMSData {
  created_date: Date;
  departments: string;
  down_percent: number;
  down_time: string;
  ip_address: string;
  maintenance_percent: number;
  maintenance_time: string;
  monitor: string;
  total_up_percent: number;
  total_up_time: string;
  type: string;
  up_percent: number;
  up_time: string;
  total_uptime_in_minutes?: number;
  total_downtime_in_minutes?: number;
}

export interface BlockTTData {
  ageing: string;
  assigned_time: string;
  assigned_to_field: string;
  assigned_to_vendor: string;
  block: string;
  cancelled: string;
  category: string;
  city: string;
  closed: string;
  cluster: string;
  effect_on_services: string;
  entity_subtype_name: string;
  enitity_type_name: string;
  equipment_host: string;
  gp: string;
  hold_time: string;
  ip: string;
  incident_id: string;
  incident_name: string;
  incident_start_on: string;
  incident_type: string;
  incident_creation_time: string;
  incident_created_on: Date;
  mode_of_contact: string;
  open_time: string;
  parent_incident_id: string;
  priority_of_repair: string;
  rfo: string;
  remark_type: string;
  remarks: string;
  reopen_date: string;
  reporting_sla: string;
  resolution_method: string;
  resolution_type_in_min: string;
  resolved_by: string;
  resolved_date_time: Date;
  slab_reach: string;
  sla_ageing: string;
  severity: string;
  status: string;
  total_resolution_time: string;
}

export interface BlockAlertData {
  alarm_clear_time: Date;
  alarm_start_time: string;
  alert: string;
  departments: string;
  duration: string;
  ip_address: string;
  message: string;
  severity: string;
  source: string;
  type: string;
  total_duration_in_minutes: number;
}

export const SEVERITY_CRITICAL = 'Critical';
export const ALERT_DOWN_MESSAGE = 'Status has entered into critical state with value [Down (  ) ]';

export enum RFO_CATEGORIZATION {
  POWER_ISSUE = 'Power Issue',
  JIO_LINK_ISSUE = 'Jio Link Issue',
  SWAN_ISSUE = 'SWAN Issue'
}

export interface RFOCategorizedTimeInMinutes {
  total_dcn_downtime_minutes: number;
  total_power_downtime_minutes: number;

}