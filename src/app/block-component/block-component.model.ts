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
  up_time: number;
  total_uptime_in_minutes?: number;
  total_downtime_in_minutes?: number;
}

export interface ManipulatedNMSData extends BlockNMSData {
  total_uptime_in_minutes: number;
  total_downtime_in_minutes: number;
  total_time_exclusive_of_sla_exclusions_in_min: number;
  total_time_exclusive_of_sla_exclusions_in_percent: number;
  alert_downtime_in_minutes: number;
  alert_downtime_in_percent: number;
  power_downtime_in_minutes: number;
  dcn_downtime_in_minutes: number;
  power_downtime_in_percent: number;
  dcn_downtime_in_percent: number;
  planned_maintenance_in_minutes: number;
  planned_maintenance_in_percent: number;
  unknown_downtime_in_minutes: number;
  unknown_downtime_in_percent: number;
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
  resolution_time_in_min: string;
  resolved_by: string;
  resolved_date_time: Date;
  slab_reach: string;
  sla_ageing: string;
  severity: string;
  status: string;
  total_resolution_time: string;
  change_id: string;
  exclusion_name: string;
  exclusion_remark: string;
  exclusion_type: string;
  pendency: string;
  vendor_name: string;
}

export interface BlockAlertData {
  alarm_clear_time: string;
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

export interface BlockSLASummary {
  report_type: string;
  time_span: string;
  no_of_blocks: number;
  up_percent: string;
  no_of_up_blocks: number | string;
  power_down_percent: string;
  fibre_down_percent: string;
  equipment_down_percent: string;
  hrt_down_percent: string;
  dcn_down_percent: string;
  planned_maintenance_percent: string;
  down_percent_exclusive_of_sla: string;
  total_sla_exclusion_percent: string;
  total_sla_exclusion_minutes: string;
  up_minutes: string;
  power_down_minutes: string;
  dcn_down_minutes: string;
  fibre_down_minutes: string;
  equipment_down_minutes: string;
  hrt_down_minutes: string;
  planned_maintenance_minutes: string;
  unknown_downtime_in_percent: string;
  unknown_downtime_in_minutes: string;
  total_down_percent: string;
  total_down_minutes: string;
  total_up_percent_exclusion: string;
  total_up_minutes_exclusion: string;
}

export interface RFOCategorizedTimeInMinutes {
  total_power_downtime_minutes: number;
  total_dcn_downtime_minutes: number;
  alert_report_empty?: boolean;
}

export type AOA = [][];
