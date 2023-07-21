export interface GpNMSData {
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

export interface ManipulatedGpNMSData extends GpNMSData {
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
  polling_time_in_minutes: number;
  polling_time_in_percent: number;
}

export interface GpAlertData {
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

export interface GpTTData {
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

export interface GpSLASummary {
  report_type: string;
  tag: string;
  time_span: string;
  no_of_gp_devices: number;
  up_percent: string;
  up_minutes: string;
  total_down_percent: string;
  total_down_minutes: string;
  power_down_percent: string;
  power_down_minutes: string;
  fibre_down_percent: string;
  fibre_down_minutes: string;
  equipment_down_percent: string;
  equipment_down_minutes: string;
  hrt_down_percent: string;
  hrt_down_minutes: string;
  dcn_down_percent: string;
  dcn_down_minutes: string;
  planned_maintenance_percent: string;
  planned_maintenance_minutes: string;
  unknown_downtime_in_percent: string;
  unknown_downtime_in_minutes: string;
  total_sla_exclusion_percent: string;
  total_sla_exclusion_minutes: string;
  total_up_percent_exclusion: string;
  total_up_minutes_exclusion: string;
}

export interface GpDeviceDetails {
  report_type: string;
  host_name: string;
  gp_ip_address: string;
  state: string;
  cluster: string;
  district: string;
  district_lgd_code: number;
  block_name: string;
  block_ip_address: string;
  block_lgd_code: string;
  gp_name: string;
  gp_lgd_code: number;
}
