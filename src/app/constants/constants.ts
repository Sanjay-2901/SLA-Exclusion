import * as ExcelJS from 'exceljs';

export const SHEET_HEADING = {
  fill: { type: 'pattern', pattern: 'solid', fgColor: { argb: '6495ED' } },
  font: { color: { argb: 'FFFFFF' }, bold: true },
} as ExcelJS.Style;

export const borderStyle = {
  style: 'thin',
} as ExcelJS.Border;

export const BORDER_STYLE = {
  top: borderStyle,
  bottom: borderStyle,
  left: borderStyle,
  right: borderStyle,
} as ExcelJS.Borders;

export const TABLE_HEADING = {
  fill: { type: 'pattern', pattern: 'solid', fgColor: { argb: 'fffffd8d' } },
  font: { bold: true },
  border: BORDER_STYLE,
} as ExcelJS.Style;

export const TABLE_HEADERS = {
  border: BORDER_STYLE,
  fill: { type: 'pattern', pattern: 'solid', fgColor: { argb: 'ffb9d3ee' } },
  font: { bold: true },
  alignment: { wrapText: true, horizontal: 'center' },
} as ExcelJS.Style;

export const PERCENT_STYLE = {
  fill: { type: 'pattern', pattern: 'solid', fgColor: { argb: 'ffffffdc' } },
} as ExcelJS.Style;

export const MINUTE_STYLE = {
  fill: { type: 'pattern', pattern: 'solid', fgColor: { argb: 'fff0f0f0' } },
} as ExcelJS.Style;

export enum VALUES {
  PERCENT = '%',
  MINUTES = 'Min',
}

export const IP_ADDRESS_PATTERN =
  /^((25[0-5]|2[0-4][0-9]|[01]?[0-9][0-9]?).){3}(25[0-5]|2[0-4][0-9]|[01]?[0-9][0-9]?)$/;

export const BlockSLASummarytHeaders = [
  'Report type',
  'TAG',
  'Time span',
  'No of Blocks',
  'No of GPs',
  'UP',
  'Total Down (Exclusive of SLA Exclusion)',
  'Power Down',
  'Fibre Down',
  'Equipment Down',
  'HRT Down',
  'DCN Down',
  'Planned Maintenance',
  'UnKnown',
  'Total SLA Exclusion',
  'Total UP (UP + Total Exclusion)',
];

export const BlockSLAFinalReportHeaders = [
  'Report type',
  'Host Name',
  'IP Address',
  'State',
  'Cluster',
  'District',
  'District LGD Code',
  'Block Name',
  'Block LGD Code',
  'No of GPs in a Block',
  'UP',
  'Total Down (Exclusive of SLA Exclusion)',
  'Power Down',
  'Fibre Down',
  'Equipment Down',
  'HRT Down',
  'DCN Down',
  'Planned Maintenance',
  'UnKnown',
  'Total SLA Exclusion',
  'Polling Time',
  'Total UP ((UP + Total Exclusion)',
];

export const BLOCK_SLA_REPORT_HEADERS = [
  'Monitor',
  'IP Address',
  'Departments',
  'Type',
  'Up (%)',
  'Up Time',
  'Down (%)',
  'Down Time',
  'Maintenance (%)',
  'Maintenance Time',
  'Total Up (%)',
  'Total Up Time',
  'Created Date',
];

export const SHQ_SLA_REPORT_HEADERS = [
  'Monitor',
  'Departments',
  'Type',
  'Up (%)',
  'Up Time',
  'Down (%)',
  'Down Time',
  'Created Date',
];

export const TT_REPORT_HEADERS = [
  'IncidentID',
  'ParentIncidentID',
  'EntityTypeName',
  'EntitySubTypeName',
  'IncidentName',
  'EquipmnetHost',
  'IP',
  'Severity',
  'Status',
  'PriorityOfRepair',
  'EffectOnServices',
  'IncidentType',
  'ModeOfContact',
  'IncidentcreationType',
  'RemarkType',
  'Remarks',
  'Cluster',
  'City',
  'Block',
  'GP',
  'SLABreach',
  'ResolutionMethod',
  'RFO',
  'IncidentStartOn',
  'IncidnetCreatedOn',
  'Ageing',
  'Open_Time',
  'Assigned_Time',
  'Assigned_To_Field',
  'Assigned_to_Vendor',
  'Cancelled',
  'Closed',
  'Hold_time',
  'Resolved_DateTime',
  'Resolved_By',
  'Total_Resolution_Time',
  'ResolutionTimeInMin',
  'SLAageing',
  'Reporting_SLA',
  'ReopenDate',
  'Category',
  'Changeid',
  'ExclusionName',
  'ExclusionRemark',
  'ExclusionType',
  'Pendency',
  'VendorName',
];

export const BLOCK_ALERT_REPORT_HEADERS = [
  'Alert',
  'Source',
  'IP Address',
  'Departments',
  'Type',
  'Severity',
  'Message',
  'Alarm Start Time',
  'Duration',
  'Alarm Clear Time',
];

export const SHQ_ALERT_REPORT_HEADERS = [
  'Alert',
  'Source',
  'Type',
  'Severity',
  'Message',
  'Last Poll Time',
  'Duration',
  'Duration Time',
];

export const BLOCK_INPUT_FILE_NAMES = [
  'block_sla_report',
  'block_noc_tt_report',
  'block_alert_report',
];

export const SHQ_INPUT_FILE_NAMES = [
  'shq_sla_report',
  'shq_noc_tt_report',
  'shq_alert_report',
];

export const BLOCK_SLA_FINAL_REPORT_COLUMNS = [
  { width: 20 },
  { width: 28 },
  { width: 20 },
  { width: 20 },
  { width: 20 },
  { width: 28 },
  { width: 20 },
  { width: 30 },
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
  { width: 20 },
  { width: 20 },
  { width: 20 },
  { width: 20 },
  { width: 20 },
  { width: 20 },
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

export interface BlockDeviceDetail {
  report_type: string;
  host_name: string;
  ip_address: string;
  state: string;
  cluster: string;
  district: string;
  district_lgd_code: number;
  block_name: string;
  block_lgd_code: string;
  no_of_gp_in_block: number;
}

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
  'UnKnown',
  'Total SLA Exclusion',
  'Total UP ( UP + Total Exclusion)',
];

export const SHQ_DEVICE_LEVEL_HEADERS = [
  'Report Type',
  'TAG',
  'Host Name',
  'IP Address',
  'Device Type',
  'UP',
  'Total Down (Exclusive of SLA Exclusion)',
  'Power Down',
  'Fibre Down',
  'Equipment Down',
  'HRT Down',
  'DCN Down',
  'Planned Maintenance',
  'Unknown',
  'Total SLA Exclusion',
  'Polling Time',
  'Total UP (UP+Total Exclusion)',
];

export const BLOCK_DEVICE_DETAILS: BlockDeviceDetail[] = [
  {
    report_type: 'Block - SLA',
    host_name: 'BNB-JCA-AATA-LBR-001',
    ip_address: '10.128.0.35',
    state: 'Chhattisgarh',
    cluster: 'B',
    district: 'Janjgir - Champa',
    district_lgd_code: 379,
    block_name: 'AKALTARA',
    block_lgd_code: 'B03641',
    no_of_gp_in_block: 53,
  },
  {
    report_type: 'Block - SLA',
    host_name: 'BNA-RNN-ABGH-LBR-001',
    ip_address: '10.128.0.69',
    state: 'Chhattisgarh',
    cluster: 'A1',
    district: 'Rajnandgaon',
    district_lgd_code: 388,
    block_name: 'AMBAGARH ',
    block_lgd_code: 'B03708',
    no_of_gp_in_block: 58,
  },
  {
    report_type: 'Block - SLA',
    host_name: 'BNC-UTR-ATGH-LBR-001',
    ip_address: '10.128.0.83',
    state: 'Chhattisgarh',
    cluster: 'C',
    district: 'Uttar Bastar Kanker',
    district_lgd_code: 381,
    block_name: 'ANTAGARH',
    block_lgd_code: 'B03658',
    no_of_gp_in_block: 47,
  },
  {
    report_type: 'Block - SLA',
    host_name: 'BNA-RPR-ARNG-HBR-001',
    ip_address: '10.128.0.67',
    state: 'Chhattisgarh',
    cluster: 'A2',
    district: 'Raipur',
    district_lgd_code: 387,
    block_name: 'ARANG  ',
    block_lgd_code: 'B03694',
    no_of_gp_in_block: 125,
  },
  {
    report_type: 'Block - SLA',
    host_name: 'BNC-BSR-BKVD-HBR-001',
    ip_address: '10.128.0.10',
    state: 'Chhattisgarh',
    cluster: 'C',
    district: 'Bastar',
    district_lgd_code: 374,
    block_name: 'BAKAWAND',
    block_lgd_code: 'B03591',
    no_of_gp_in_block: 64,
  },
  {
    report_type: 'Block - SLA',
    host_name: 'BNA-BBZ-BDBZ-HBR-001',
    ip_address: '10.128.0.6',
    state: 'Chhattisgarh',
    cluster: 'A2',
    district: 'Baloda Bazar',
    district_lgd_code: 644,
    block_name: 'BALODA BAZAR',
    block_lgd_code: 'B03695',
    no_of_gp_in_block: 96,
  },
  {
    report_type: 'Block - SLA',
    host_name: 'BNB-RGH-BAKA-LBR-001',
    ip_address: '10.128.0.62',
    state: 'Chhattisgarh',
    cluster: 'B',
    district: 'Raigarh',
    district_lgd_code: 386,
    block_name: 'BARAMKELA',
    block_lgd_code: 'B03684',
    no_of_gp_in_block: 69,
  },
  {
    report_type: 'Block - SLA',
    host_name: 'BNA-MSD-BANA-LBR-001',
    ip_address: '10.128.0.56',
    state: 'Chhattisgarh',
    cluster: 'A2',
    district: 'Mahasumnd',
    district_lgd_code: 385,
    block_name: 'BASNA',
    block_lgd_code: 'B03680',
    no_of_gp_in_block: 89,
  },
  {
    report_type: 'Block - SLA',
    host_name: 'BNC-BSR-BSAR-LBR-001',
    ip_address: '10.128.0.11',
    state: 'Chhattisgarh',
    cluster: 'C',
    district: 'Bastar',
    district_lgd_code: 374,
    block_name: 'BASTANAR',
    block_lgd_code: 'B03592',
    no_of_gp_in_block: 22,
  },
  {
    report_type: 'Block - SLA',
    host_name: 'BND-SGA-BTUI-LBR-001',
    ip_address: '10.128.0.79',
    state: 'Chhattisgarh',
    cluster: 'D',
    district: 'Surguja',
    district_lgd_code: 389,
    block_name: 'BATOULI',
    block_lgd_code: 'B03719',
    no_of_gp_in_block: 38,
  },
  {
    report_type: 'Block - SLA',
    host_name: 'BNA-BEA-BMTA-HBR-001',
    ip_address: '10.128.0.15',
    state: 'Chhattisgarh',
    cluster: 'A2',
    district: 'Bemetara',
    district_lgd_code: 650,
    block_name: 'BEMETARA',
    block_lgd_code: 'B03630',
    no_of_gp_in_block: 90,
  },
  {
    report_type: 'Block - SLA',
    host_name: 'BNA-BEA-BELA-LBR-001',
    ip_address: '10.128.0.16',
    state: 'Chhattisgarh',
    cluster: 'A2',
    district: 'Bemetara',
    district_lgd_code: 650,
    block_name: 'BERLA',
    block_lgd_code: 'B03631',
    no_of_gp_in_block: 83,
  },
  {
    report_type: 'Block - SLA',
    host_name: 'BNC-BJR-BRMH-HBR-001',
    ip_address: '10.128.0.19',
    state: 'Chhattisgarh',
    cluster: 'C',
    district: 'Bijapur',
    district_lgd_code: 636,
    block_name: 'BHAIRAMGARH',
    block_lgd_code: 'B03614',
    no_of_gp_in_block: 12,
  },
  {
    report_type: 'Block - SLA',
    host_name: 'BND-SAR-BIAN-HBR-001',
    ip_address: '10.128.0.76',
    state: 'Chhattisgarh',
    cluster: 'D',
    district: 'Surajpur',
    district_lgd_code: 648,
    block_name: 'BHAIYATHAN',
    block_lgd_code: 'B03720',
    no_of_gp_in_block: 65,
  },
  {
    report_type: 'Block - SLA',
    host_name: 'BND-KRA-BRTR-LBR-001',
    ip_address: '10.128.0.53',
    state: 'Chhattisgarh',
    cluster: 'D',
    district: 'Koriya',
    district_lgd_code: 384,
    block_name: 'BHARATPUR',
    block_lgd_code: 'B03675',
    no_of_gp_in_block: 10,
  },
  {
    report_type: 'Block - SLA',
    host_name: 'BNA-GYD-BAAH-LBR-001',
    ip_address: '10.128.0.31',
    state: 'Chhattisgarh',
    cluster: 'A2',
    district: 'Gariyaband',
    district_lgd_code: 645,
    block_name: 'BINDRANAVAGARH(GARIYABAND)',
    block_lgd_code: 'B03702',
    no_of_gp_in_block: 55,
  },
  {
    report_type: 'Block - SLA',
    host_name: 'BNA-KEM-BOLA-LBR-001',
    ip_address: '10.128.0.45',
    state: 'Chhattisgarh',
    cluster: 'A1',
    district: 'Kabeerdham',
    district_lgd_code: 382,
    block_name: 'BODLA ',
    block_lgd_code: 'B03665',
    no_of_gp_in_block: 88,
  },
  {
    report_type: 'Block - SLA',
    host_name: 'BNB-JCA-CHMA-LBR-001',
    ip_address: '10.128.0.36',
    state: 'Chhattisgarh',
    cluster: 'B',
    district: 'Janjgir - Champa',
    district_lgd_code: 379,
    block_name: 'CHAMPA',
    block_lgd_code: 'B03642',
    no_of_gp_in_block: 51,
  },
  {
    report_type: 'Block - SLA',
    host_name: 'BNC-UTR-CAAA-LBR-001',
    ip_address: '10.128.0.84',
    state: 'Chhattisgarh',
    cluster: 'C',
    district: 'Uttar Bastar Kanker',
    district_lgd_code: 381,
    block_name: 'CHARAMA ',
    block_lgd_code: 'B03660',
    no_of_gp_in_block: 59,
  },
  {
    report_type: 'Block - SLA',
    host_name: 'BNA-RNN-CIHN-LBR-001',
    ip_address: '10.128.0.70',
    state: 'Chhattisgarh',
    cluster: 'A1',
    district: 'Rajnandgaon',
    district_lgd_code: 388,
    block_name: 'CHHUIKHADAN',
    block_lgd_code: 'B03709',
    no_of_gp_in_block: 90,
  },
  {
    report_type: 'Block - SLA',
    host_name: 'BNA-GYD-CHUA-HBR-001',
    ip_address: '10.128.0.32',
    state: 'Chhattisgarh',
    cluster: 'A2',
    district: 'Gariyaband',
    district_lgd_code: 645,
    block_name: 'CHHURA',
    block_lgd_code: 'B03698',
    no_of_gp_in_block: 64,
  },
  {
    report_type: 'Block - SLA',
    host_name: 'BNA-RNN-CHRA-HBR-001',
    ip_address: '10.128.0.71',
    state: 'Chhattisgarh',
    cluster: 'A2',
    district: 'Rajnandgaon',
    district_lgd_code: 388,
    block_name: 'CHHURIYA',
    block_lgd_code: 'B03710',
    no_of_gp_in_block: 96,
  },
  {
    report_type: 'Block - SLA',
    host_name: 'BNB-JCA-DAHA-HBR-001',
    ip_address: '10.128.0.37',
    state: 'Chhattisgarh',
    cluster: 'B',
    district: 'Janjgir - Champa',
    district_lgd_code: 379,
    block_name: 'DABHRA',
    block_lgd_code: 'B03644',
    no_of_gp_in_block: 77,
  },
  {
    report_type: 'Block - SLA',
    host_name: 'BNA-GYD-DABA-LBR-001',
    ip_address: '10.128.0.12',
    state: 'Chhattisgarh',
    cluster: 'C',
    district: 'Bastar',
    district_lgd_code: 374,
    block_name: 'DARBHA',
    block_lgd_code: 'B03594',
    no_of_gp_in_block: 8,
  },
  {
    report_type: 'Block - SLA',
    host_name: 'BNA-GYD-DOHG-LBR-001',
    ip_address: '10.128.0.33',
    state: 'Chhattisgarh',
    cluster: 'A2',
    district: 'Gariyaband',
    district_lgd_code: 645,
    block_name: 'DEOBHOG',
    block_lgd_code: 'B03699',
    no_of_gp_in_block: 47,
  },
  {
    report_type: 'Block - SLA',
    host_name: 'BNA-DUG-DADA-HBR-001',
    ip_address: '10.128.0.30',
    state: 'Chhattisgarh',
    cluster: 'A1',
    district: 'Durg',
    district_lgd_code: 378,
    block_name: 'DHAMDHA',
    block_lgd_code: 'B03632',
    no_of_gp_in_block: 108,
  },
  {
    report_type: 'Block - SLA',
    host_name: 'BNA-BLD-DODI-LBR-001',
    ip_address: '10.128.0.3',
    state: 'Chhattisgarh',
    cluster: 'A1',
    district: 'Balod',
    district_lgd_code: 646,
    block_name: 'DONDI',
    block_lgd_code: 'B03633',
    no_of_gp_in_block: 50,
  },
  {
    report_type: 'Block - SLA',
    host_name: 'BNA-BLD-DILA-LBR-001',
    ip_address: '10.128.0.4',
    state: 'Chhattisgarh',
    cluster: 'A1',
    district: 'Balod',
    district_lgd_code: 646,
    block_name: 'DONDI LUHARA',
    block_lgd_code: 'B03634',
    no_of_gp_in_block: 111,
  },
  {
    report_type: 'Block - SLA',
    host_name: 'BNA-RNN-DGRN-LBR-001',
    ip_address: '10.128.0.72',
    state: 'Chhattisgarh',
    cluster: 'A1',
    district: 'Rajnandgaon',
    district_lgd_code: 388,
    block_name: 'DONGARGAON ',
    block_lgd_code: 'B03711',
    no_of_gp_in_block: 63,
  },
  {
    report_type: 'Block - SLA',
    host_name: 'BNA-RNN-DGRH-LBR-001',
    ip_address: '10.128.0.73',
    state: 'Chhattisgarh',
    cluster: 'A1',
    district: 'Rajnandgaon',
    district_lgd_code: 388,
    block_name: 'DONGARGARH',
    block_lgd_code: 'B03712',
    no_of_gp_in_block: 89,
  },
  {
    report_type: 'Block - SLA',
    host_name: 'BNC-UTR-DGOL-LBR-001',
    ip_address: '10.128.0.85',
    state: 'Chhattisgarh',
    cluster: 'C',
    district: 'Uttar Bastar Kanker',
    district_lgd_code: 381,
    block_name: 'DURGKONDAL',
    block_lgd_code: 'B03661',
    no_of_gp_in_block: 34,
  },
  {
    report_type: 'Block - SLA',
    host_name: 'BNC-KAN-FAGN-HBR-001',
    ip_address: '10.128.0.48',
    state: 'Chhattisgarh',
    cluster: 'C',
    district: 'Kondagaon',
    district_lgd_code: 643,
    block_name: 'FARASGAON',
    block_lgd_code: 'B03602',
    no_of_gp_in_block: 62,
  },
  {
    report_type: 'Block - SLA',
    host_name: 'BND-JHR-FSBR-HBR-001',
    ip_address: '10.128.0.43',
    state: 'Chhattisgarh',
    cluster: 'D',
    district: 'Jashpur',
    district_lgd_code: 380,
    block_name: 'FARSABAHAR',
    block_lgd_code: 'B03657',
    no_of_gp_in_block: 55,
  },
  {
    report_type: 'Block - SLA',
    host_name: 'BNC-DTA-GIAM-HBR-001',
    ip_address: '10.128.0.26',
    state: 'Chhattisgarh',
    cluster: 'C',
    district: 'Dakshin Bastar Dantewada',
    district_lgd_code: 376,
    block_name: 'GIDAM',
    block_lgd_code: 'B03619',
    no_of_gp_in_block: 30,
  },
  {
    report_type: 'Block - SLA',
    host_name: 'BNA-BLD-GDRI-HBR-001',
    ip_address: '10.128.0.5',
    state: 'Chhattisgarh',
    cluster: 'A1',
    district: 'Balod',
    district_lgd_code: 646,
    block_name: 'GUNDERDEHI',
    block_lgd_code: 'B03636',
    no_of_gp_in_block: 109,
  },
  {
    report_type: 'Block - SLA',
    host_name: 'BNC-BSR-JDLR-LBR-001',
    ip_address: '10.128.0.13',
    state: 'Chhattisgarh',
    cluster: 'C',
    district: 'Bastar',
    district_lgd_code: 374,
    block_name: 'JAGDALPUR',
    block_lgd_code: 'B03595',
    no_of_gp_in_block: 54,
  },
  {
    report_type: 'Block - SLA',
    host_name: 'BNB-KRB-KRAA-LBR-001',
    ip_address: '10.128.0.51',
    state: 'Chhattisgarh',
    cluster: 'B',
    district: 'Korba',
    district_lgd_code: 383,
    block_name: 'KARTALA',
    block_lgd_code: 'B03669',
    no_of_gp_in_block: 63,
  },
  {
    report_type: 'Block - SLA',
    host_name: 'BNA-BBZ-KADL-LBR-001',
    ip_address: '10.128.0.7',
    state: 'Chhattisgarh',
    cluster: 'A2',
    district: 'Baloda Bazar',
    district_lgd_code: 644,
    block_name: 'KASDOL',
    block_lgd_code: 'B03703',
    no_of_gp_in_block: 80,
  },
  {
    report_type: 'Block - SLA',
    host_name: 'BNC-DTA-KEAN-LBR-001',
    ip_address: '10.128.0.27',
    state: 'Chhattisgarh',
    cluster: 'C',
    district: 'Dakshin Bastar Dantewada',
    district_lgd_code: 376,
    block_name: 'KATEKALYAN',
    block_lgd_code: 'B03620',
    no_of_gp_in_block: 14,
  },
  {
    report_type: 'Block - SLA',
    host_name: 'BNA-KEM-KWRA-LBR-001',
    ip_address: '10.128.0.46',
    state: 'Chhattisgarh',
    cluster: 'A1',
    district: 'Kabeerdham',
    district_lgd_code: 382,
    block_name: 'KAWARDHA ',
    block_lgd_code: 'B03666',
    no_of_gp_in_block: 84,
  },
  {
    report_type: 'Block - SLA',
    host_name: 'BNC-KAN-KEKL-LBR-001',
    ip_address: '10.128.0.49',
    state: 'Chhattisgarh',
    cluster: 'C',
    district: 'Kondagaon',
    district_lgd_code: 643,
    block_name: 'KESKAL',
    block_lgd_code: 'B03596',
    no_of_gp_in_block: 23,
  },
  {
    report_type: 'Block - SLA',
    host_name: 'BND-KRA-KDAA-HBR-001',
    ip_address: '10.128.0.54',
    state: 'Chhattisgarh',
    cluster: 'D',
    district: 'Koriya',
    district_lgd_code: 384,
    block_name: 'KHADGANVA',
    block_lgd_code: 'B03676',
    no_of_gp_in_block: 48,
  },
  {
    report_type: 'Block - SLA',
    host_name: 'BNB-BAR-KKTA-LBR-001',
    ip_address: '10.128.0.22',
    state: 'Chhattisgarh',
    cluster: 'B',
    district: 'Bilaspur',
    district_lgd_code: 375,
    block_name: 'KOTA',
    block_lgd_code: 'B03607',
    no_of_gp_in_block: 82,
  },
  {
    report_type: 'Block - SLA',
    host_name: 'BNB-RGH-LIUA-LBR-001',
    ip_address: '10.128.0.63',
    state: 'Chhattisgarh',
    cluster: 'B',
    district: 'Raigarh',
    district_lgd_code: 386,
    block_name: 'LAILUNGA',
    block_lgd_code: 'B03688',
    no_of_gp_in_block: 69,
  },
  {
    report_type: 'Block - SLA',
    host_name: 'BND-SGA-LHNR-LBR-001',
    ip_address: '10.128.0.80',
    state: 'Chhattisgarh',
    cluster: 'D',
    district: 'Surguja',
    district_lgd_code: 389,
    block_name: 'LAKHANPUR',
    block_lgd_code: 'B03722',
    no_of_gp_in_block: 58,
  },
  {
    report_type: 'Block - SLA',
    host_name: 'BNC-BSR-LNIA-LBR-001',
    ip_address: '10.128.0.14',
    state: 'Chhattisgarh',
    cluster: 'C',
    district: 'Bastar',
    district_lgd_code: 374,
    block_name: 'LOHANDIGUDA',
    block_lgd_code: 'B03598',
    no_of_gp_in_block: 29,
  },
  {
    report_type: 'Block - SLA',
    host_name: 'BNB-MGI-LOMI-HBR-001',
    ip_address: '10.128.0.59',
    state: 'Chhattisgarh',
    cluster: 'B',
    district: 'Mungeli',
    district_lgd_code: 647,
    block_name: 'LORMI ',
    block_lgd_code: 'B03608',
    no_of_gp_in_block: 112,
  },
  {
    report_type: 'Block - SLA',
    host_name: 'BND-SGA-LUDA-HBR-001',
    ip_address: '10.128.0.81',
    state: 'Chhattisgarh',
    cluster: 'D',
    district: 'Surguja',
    district_lgd_code: 389,
    block_name: 'LUNDRA',
    block_lgd_code: 'B03723',
    no_of_gp_in_block: 63,
  },
  {
    report_type: 'Block - SLA',
    host_name: 'BNA-DMI-MGRD-LBR-001',
    ip_address: '10.128.0.28',
    state: 'Chhattisgarh',
    cluster: 'A2',
    district: 'Dhamtari',
    district_lgd_code: 377,
    block_name: 'MAGARLOD ',
    block_lgd_code: 'B03627',
    no_of_gp_in_block: 48,
  },
  {
    report_type: 'Block - SLA',
    host_name: 'BND-SGA-MIPT-LBR-001',
    ip_address: '10.128.0.82',
    state: 'Chhattisgarh',
    cluster: 'D',
    district: 'Surguja',
    district_lgd_code: 389,
    block_name: 'MAINPAT',
    block_lgd_code: 'B03724',
    no_of_gp_in_block: 25,
  },
  {
    report_type: 'Block - SLA',
    host_name: 'BNA-GYD-MIPR-LBR-001',
    ip_address: '10.128.0.34',
    state: 'Chhattisgarh',
    cluster: 'A2',
    district: 'Gariyaband',
    district_lgd_code: 645,
    block_name: 'MAINPUR',
    block_lgd_code: 'B03704',
    no_of_gp_in_block: 10,
  },
  {
    report_type: 'Block - SLA',
    host_name: 'BNC-KAN-MADI-LBR-001',
    ip_address: '10.128.0.50',
    state: 'Chhattisgarh',
    cluster: 'C',
    district: 'Kondagaon',
    district_lgd_code: 643,
    block_name: 'MAKDI',
    block_lgd_code: 'B03599',
    no_of_gp_in_block: 43,
  },
  {
    report_type: 'Block - SLA',
    host_name: 'BNB-JCA-MKAA-LBR-001',
    ip_address: '10.128.0.39',
    state: 'Chhattisgarh',
    cluster: 'B',
    district: 'Janjgir - Champa',
    district_lgd_code: 379,
    block_name: 'MALKHARODA',
    block_lgd_code: 'B03646',
    no_of_gp_in_block: 71,
  },
  {
    report_type: 'Block - SLA',
    host_name: 'BNB-BAR-MRAI-LBR-001',
    ip_address: '10.128.0.23',
    state: 'Chhattisgarh',
    cluster: 'B',
    district: 'Bilaspur',
    district_lgd_code: 375,
    block_name: 'MARWAHI ',
    block_lgd_code: 'B03609',
    no_of_gp_in_block: 56,
  },
  {
    report_type: 'Block - SLA',
    host_name: 'BNB-BAR-MSUI-HBR-001',
    ip_address: '10.128.0.24',
    state: 'Chhattisgarh',
    cluster: 'B',
    district: 'Bilaspur',
    district_lgd_code: 375,
    block_name: 'MASTURI',
    block_lgd_code: 'B03610',
    no_of_gp_in_block: 111,
  },
  {
    report_type: 'Block - SLA',
    host_name: 'BNA-RNN-MOLA-LBR-001',
    ip_address: '10.128.0.74',
    state: 'Chhattisgarh',
    cluster: 'A1',
    district: 'Rajnandgaon',
    district_lgd_code: 388,
    block_name: 'MOHLA',
    block_lgd_code: 'B03715',
    no_of_gp_in_block: 50,
  },
  {
    report_type: 'Block - SLA',
    host_name: 'BNA-DMI-NARI-HBR-001',
    ip_address: '10.128.0.29',
    state: 'Chhattisgarh',
    cluster: 'A2',
    district: 'Dhamtari',
    district_lgd_code: 377,
    block_name: 'NAGRI',
    block_lgd_code: 'B03628',
    no_of_gp_in_block: 58,
  },
  {
    report_type: 'Block - SLA',
    host_name: 'BNC-UTR-NHRR-LBR-001',
    ip_address: '10.128.0.86',
    state: 'Chhattisgarh',
    cluster: 'C',
    district: 'Uttar Bastar Kanker',
    district_lgd_code: 381,
    block_name: 'NARHARPUR',
    block_lgd_code: 'B03664',
    no_of_gp_in_block: 64,
  },
  {
    report_type: 'Block - SLA',
    host_name: 'BNB-JCA-NWGH-LBR-001',
    ip_address: '10.128.0.40',
    state: 'Chhattisgarh',
    cluster: 'B',
    district: 'Janjgir - Champa',
    district_lgd_code: 379,
    block_name: 'NAWAGARH',
    block_lgd_code: 'B03647',
    no_of_gp_in_block: 94,
  },
  {
    report_type: 'Block - SLA',
    host_name: 'BND-SAR-OUGI-LBR-001',
    ip_address: '10.128.0.77',
    state: 'Chhattisgarh',
    cluster: 'D',
    district: 'Surajpur',
    district_lgd_code: 648,
    block_name: 'OUDGI',
    block_lgd_code: 'B03725',
    no_of_gp_in_block: 12,
  },
  {
    report_type: 'Block - SLA',
    host_name: 'BNC-UTR-PHNR-HBR-001',
    ip_address: '10.128.0.87',
    state: 'Chhattisgarh',
    cluster: 'C',
    district: 'Uttar Bastar Kanker',
    district_lgd_code: 381,
    block_name: 'PAKHANJUR',
    block_lgd_code: 'B03663',
    no_of_gp_in_block: 74,
  },
  {
    report_type: 'Block - SLA',
    host_name: 'BNA-BBZ-PAAI-LBR-001',
    ip_address: '10.128.0.8',
    state: 'Chhattisgarh',
    cluster: 'A2',
    district: 'Baloda Bazar',
    district_lgd_code: 644,
    block_name: 'PALARI ',
    block_lgd_code: 'B03705',
    no_of_gp_in_block: 76,
  },
  {
    report_type: 'Block - SLA',
    host_name: 'BNB-JCA-PMAH-LBR-001',
    ip_address: '10.128.0.41',
    state: 'Chhattisgarh',
    cluster: 'B',
    district: 'Janjgir - Champa',
    district_lgd_code: 379,
    block_name: 'PAMGARH ',
    block_lgd_code: 'B03648',
    no_of_gp_in_block: 31,
  },
  {
    report_type: 'Block - SLA',
    host_name: 'BNA-KEM-PDRA-HBR-001',
    ip_address: '10.128.0.47',
    state: 'Chhattisgarh',
    cluster: 'A1',
    district: 'Kabeerdham',
    district_lgd_code: 382,
    block_name: 'PANDARIYA ',
    block_lgd_code: 'B03667',
    no_of_gp_in_block: 111,
  },
  {
    report_type: 'Block - SLA',
    host_name: 'BND-JHR-PHLN-LBR-001',
    ip_address: '10.128.0.44',
    state: 'Chhattisgarh',
    cluster: 'D',
    district: 'Jashpur',
    district_lgd_code: 380,
    block_name: 'PATHALGAON ',
    block_lgd_code: 'B03656',
    no_of_gp_in_block: 41,
  },
  {
    report_type: 'Block - SLA',
    host_name: 'BNB-MGI-PHRA-LBR-001',
    ip_address: '10.128.0.60',
    state: 'Chhattisgarh',
    cluster: 'B',
    district: 'Mungeli',
    district_lgd_code: 647,
    block_name: 'PATHARIYA',
    block_lgd_code: 'B03612',
    no_of_gp_in_block: 82,
  },
  {
    report_type: 'Block - SLA',
    host_name: 'BNA-MSD-PTOA-HBR-001',
    ip_address: '10.128.0.57',
    state: 'Chhattisgarh',
    cluster: 'A2',
    district: 'Mahasumnd',
    district_lgd_code: 385,
    block_name: 'PITHORA',
    block_lgd_code: 'B03682',
    no_of_gp_in_block: 106,
  },
  {
    report_type: 'Block - SLA',
    host_name: 'BNB-KRB-PPDI-HBR-001',
    ip_address: '10.128.0.52',
    state: 'Chhattisgarh',
    cluster: 'B',
    district: 'Korba',
    district_lgd_code: 383,
    block_name: 'PODI',
    block_lgd_code: 'B03673',
    no_of_gp_in_block: 10,
  },
  {
    report_type: 'Block - SLA',
    host_name: 'BND-SAR-PMAR-LBR-001',
    ip_address: '10.128.0.78',
    state: 'Chhattisgarh',
    cluster: 'D',
    district: 'Surajpur',
    district_lgd_code: 648,
    block_name: 'PREMNAGAR',
    block_lgd_code: 'B03727',
    no_of_gp_in_block: 28,
  },
  {
    report_type: 'Block - SLA',
    host_name: 'BNB-RGH-RIAH-HBR-001',
    ip_address: '10.128.0.64',
    state: 'Chhattisgarh',
    cluster: 'B',
    district: 'Raigarh',
    district_lgd_code: 386,
    block_name: 'RAIGARH',
    block_lgd_code: 'B03690',
    no_of_gp_in_block: 38,
  },
  {
    report_type: 'Block - SLA',
    host_name: 'BNA-BEA-SSJA-LBR-001',
    ip_address: '10.128.0.17',
    state: 'Chhattisgarh',
    cluster: 'A1',
    district: 'Bemetara',
    district_lgd_code: 650,
    block_name: 'SAJA ',
    block_lgd_code: 'B03640',
    no_of_gp_in_block: 94,
  },
  {
    report_type: 'Block - SLA',
    host_name: 'BNB-JCA-SATI-LBR-001',
    ip_address: '10.128.0.42',
    state: 'Chhattisgarh',
    cluster: 'B',
    district: 'Janjgir - Champa',
    district_lgd_code: 379,
    block_name: 'SAKTI',
    block_lgd_code: 'B03649',
    no_of_gp_in_block: 67,
  },
  {
    report_type: 'Block - SLA',
    host_name: 'BNA-MSD-SAPI-LBR-001',
    ip_address: '10.128.0.58',
    state: 'Chhattisgarh',
    cluster: 'A2',
    district: 'Mahasumnd',
    district_lgd_code: 385,
    block_name: 'SARAIPALI',
    block_lgd_code: 'B03683',
    no_of_gp_in_block: 95,
  },
  {
    report_type: 'Block - SLA',
    host_name: 'BNA-BBZ-SIGA-LBR-001',
    ip_address: '10.128.0.9',
    state: 'Chhattisgarh',
    cluster: 'A2',
    district: 'Baloda Bazar',
    district_lgd_code: 644,
    block_name: 'SIMGA',
    block_lgd_code: 'B03706',
    no_of_gp_in_block: 90,
  },
  {
    report_type: 'Block - SLA',
    host_name: 'BND-KRA-SOHT-LBR-001',
    ip_address: '10.128.0.55',
    state: 'Chhattisgarh',
    cluster: 'D',
    district: 'Koriya',
    district_lgd_code: 384,
    block_name: 'SONHAT ',
    block_lgd_code: 'B03678',
    no_of_gp_in_block: 22,
  },
  {
    report_type: 'Block - SLA',
    host_name: 'BNB-BAR-THTR-LBR-001',
    ip_address: '10.128.0.25',
    state: 'Chhattisgarh',
    cluster: 'B',
    district: 'Bilaspur',
    district_lgd_code: 375,
    block_name: 'TAKHATPUR ',
    block_lgd_code: 'B03613',
    no_of_gp_in_block: 106,
  },
  {
    report_type: 'Block - SLA',
    host_name: 'BNB-RGH-TANR-LBR-001',
    ip_address: '10.128.0.65',
    state: 'Chhattisgarh',
    cluster: 'B',
    district: 'Raigarh',
    district_lgd_code: 386,
    block_name: 'TAMNAR',
    block_lgd_code: 'B03692',
    no_of_gp_in_block: 51,
  },
  {
    report_type: 'Block - SLA',
    host_name: 'BNA-RPR-TIDA-HBR-001',
    ip_address: '10.128.0.68',
    state: 'Chhattisgarh',
    cluster: 'A2',
    district: 'Raipur',
    district_lgd_code: 387,
    block_name: 'TILDA',
    block_lgd_code: 'B03707',
    no_of_gp_in_block: 95,
  },
  {
    report_type: 'Block - SLA',
    host_name: 'BNB-RGH-UAPR-LBR-001',
    ip_address: '10.128.0.66',
    state: 'Chhattisgarh',
    cluster: 'B',
    district: 'Raigarh',
    district_lgd_code: 386,
    block_name: 'UDAIPUR (DHARAMJAIGARH)',
    block_lgd_code: 'B03685',
    no_of_gp_in_block: 96,
  },
];
