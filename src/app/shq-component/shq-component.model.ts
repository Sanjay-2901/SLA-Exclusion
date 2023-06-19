export interface ShqNMSData {
  monitor: string;
  departments: string;
  type: string;
  up_percent: number;
  up_time: string;
  down_percent: number;
  down_time: string;
  created_date: Date;
  ip_address: string;
  total_uptime_in_minutes?: number;
  total_downtime_in_minutes?: number;
}

export interface ManipulatedShqNmsData extends ShqNMSData {
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
  unknown_downtime_in_minutes: number;
  unknown_downtime_in_percent: number;
}

export interface ShqAlertData {
  alert: string;
  source: string;
  type: string;
  severity: string;
  message: string;
  last_poll_time: string;
  duration: string;
  duration_time: string;
  ip_address: string;
  total_duration_in_minutes: number;
}

export interface ShqTTData {
  incident_id: string;
  parent_incident_id: string;
  enitity_type_name: string;
  entity_subtype_name: string;
  incident_name: string;
  equipment_host: string;
  ip: string;
  severity: string;
  status: string;
  priority_of_repair: string;
  effect_on_services: string;
  incident_type: string;
  mode_of_contact: string;
  incident_creation_time: string;
  remark_type: string;
  remarks: string;
  cluster: string;
  city: string;
  block: string;
  gp: string;
  slab_reach: string;
  resolution_method: string;
  rfo: string;
  incident_start_on: string;
  incident_created_on: Date;
  ageing: string;
  open_time: string;
  assigned_time: string;
  assigned_to_field: string;
  assigned_to_vendor: string;
  cancelled: string;
  closed: string;
  hold_time: string;
  resolved_date_time: Date;
  resolved_by: string;
  total_resolution_time: string;
  resolution_type_in_min: string;
  sla_ageing: string;
  reporting_sla: string;
  reopen_date: string;
  category: string;
  change_id: string;
  exclusion_name: string;
  exclusion_remark: string;
  exclusion_type: string;
  pendency: string;
  vendor_name: string;
}

export interface ShqSlaSummary {
  report_type: string;
  tag: string;
  time_span: string;
  no_of_shq_devices: number;
  up_percent: number;
  up_minutes: number;
  total_down_exclusive_of_sla_exclusion_percent: number;
  total_down_exclusive_of_sla_exclusion_minute: number;
  power_down_percent: number;
  power_dowm_minute: number;
  fibre_down_percent: number;
  fiber_down_minute: number;
  equipment_down_percent: number;
  equipment_down_minute: number;
  hrt_down_percent: number;
  hrt_down_minute: number;
  dcn_down_percent: number;
  dcn_down_minute: number;
  planned_maintenance_percent: number;
  planned_maintenance_minute: number;
  unknown_downtime_in_percent: number;
  unknown_downtime_in_minutes: number;
  total_sla_exclusion_percent: number;
  total_sla_exclusion_minute: number;
  total_up_percent: number;
  total_up_minute: number;
}
