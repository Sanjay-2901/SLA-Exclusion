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

export interface BlockSLASummaryPercent {
  report_type: string;
  time_span: string;
  no_of_blocks: number;
  up_percent: number;
  no_of_up_blocks: number | string;
  power_down_percent: number;
  fibre_down_percent: number;
  equipment_down_percent: number;
  hrt_down_percent: number;
  dcn_down_percent: number;
  planned_maintenance_percent: number;
  down_percent_exclusive_of_sla: number;
  no_of_down_blocks: number | string;
  total_sla_exclusion_percent: number;
  total_sla_exclusion_minutes: number;
  total_up_percent: number;
  up_minutes: number;
  power_down_minutes: number;
  dcn_down_minutes: number;
  fibre_down_minutes: number;
  equipment_down_minutes: number;
  hrt_down_minutes: number;
  planned_maintenance_minutes: number;
  total_up_minutes: number;
  total_down_percent: number;
  total_down_minutes: number;
  total_up_percent_exclusion: number;
  total_up_minutes_exclusion: number;
}

export interface RFOCategorizedTimeInMinutes {
  total_dcn_downtime_minutes: number;
  total_power_downtime_minutes: number;
}

export type AOA = [][];
