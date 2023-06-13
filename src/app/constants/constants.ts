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

export const SHQ_SLQ_FINAL_REPORT_COLUMNS = [
  {
    width: 20,
  },
  {
    width: 20,
  },
  {
    width: 20,
  },
  {
    width: 20,
  },
  {
    width: 20,
  },
  {
    width: 20,
  },
  {
    width: 20,
  },
  {
    width: 20,
  },
  {
    width: 20,
  },
  {
    width: 20,
  },
  {
    width: 20,
  },
  {
    width: 20,
  },
  {
    width: 20,
  },
  {
    width: 20,
  },
  {
    width: 20,
  },
  {
    width: 20,
  },
  {
    width: 20,
  },
  {
    width: 20,
  },
  {
    width: 20,
  },
  {
    width: 20,
  },
  {
    width: 20,
  },
  {
    width: 20,
  },
  {
    width: 20,
  },
  {
    width: 20,
  },
  {
    width: 20,
  },
  {
    width: 20,
  },
  {
    width: 20,
  },
];

export const SEVERITY_CRITICAL = 'Critical';
export const SEVERITY_WARNING = 'Warning';
export const ALERT_DOWN_MESSAGE =
  'Status has entered into critical state with value [Down (  ) ]';

export enum RFO_CATEGORIZATION {
  POWER_ISSUE = 'Power Issue',
  JIO_LINK_ISSUE = 'Jio Link Issue',
  SWAN_ISSUE = 'SWAN Issue',
}

export const borderStyle = {
  style: 'thin',
} as ExcelJS.Border;

export const BORDER_STYLE = {
  top: borderStyle,
  bottom: borderStyle,
  left: borderStyle,
  right: borderStyle,
} as ExcelJS.Borders;

export const VMWAREDEVICE = 'VMware ESX/ESXi';

export const SHQ_SUMMARY_HEADERS = [
  'Report Type',
  'TAG',
  'Time Span',
  'No of SHQ Device',
  'UP',
  'Total Down ( Exclusive of SLA Exclusion)',
  'Power Down',
  'Fibre Down',
  'Equipment Down',
  'HRT Down',
  'DCN Down',
  'Planned Maintenance',
  'Total SLA Exclusion',
  'Total UP ( UP + Total Exclusion)',
];
