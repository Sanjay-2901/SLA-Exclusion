import { Component, OnInit } from '@angular/core';
import * as ExcelJS from 'exceljs';
import {
  ManipulatedShqNmsData,
  ShqAlertData,
  ShqNMSData,
  ShqTTData,
} from './shq-component.model';
import * as moment from 'moment';
import { ShqService } from './shq-service.service';
import { AOA } from '../block-component/block-component.model';

@Component({
  selector: 'app-shq-component',
  templateUrl: './shq-component.component.html',
  styleUrls: ['./shq-component.component.scss'],
})
export class ShqComponentComponent implements OnInit {
  shqNMSData: any = [];
  shqTTData: any = [];
  shqAlertData: any = [];
  manipulatedNMSData: any = [];
  worksheet!: ExcelJS.Worksheet;

  constructor(private ShqService: ShqService) {}

  ngOnInit(): void {}

  onFileChange(event: any) {
    const file = event.target.files[0];
    const workbook = new ExcelJS.Workbook();
    const reader = new FileReader();

    reader.onload = (e: any) => {
      const buffer = e.target.result;

      workbook.xlsx.load(buffer).then(() => {
        workbook.worksheets.forEach((_, index) => {
          console.log('worksheet', index);
          this.worksheet = workbook.getWorksheet(index + 1);
          this.readWorksheet(this.worksheet);
        });
        this.manipulateShqNmsData();
      });
    };
    reader.readAsArrayBuffer(file);
  }

  readWorksheet(worksheet: ExcelJS.Worksheet): void {
    let workSheetName = worksheet.name;
    let data: AOA = [];
    this.worksheet.eachRow({ includeEmpty: true }, (row: ExcelJS.Row) => {
      const rowData: any = [];
      row.eachCell({ includeEmpty: true }, (cell: ExcelJS.Cell) => {
        rowData.push(cell.value);
      });
      data.push(rowData);
    });

    let result: any = [];
    data.shift();
    data.forEach((data: any, index: number) => {
      if (workSheetName === 'SHQ-SLA-Report') {
        let obj: ShqNMSData = {
          monitor: data[0].trim(),
          departments: data[1],
          ip_address: data[0].match(/\((.*?)\)/)[1].trim(),
          type: data[2],
          up_percent: data[3],
          up_time: data[4],
          down_percent: data[5],
          down_time: data[6],
          created_date: data[7],
        };
        result.push(obj);
      } else if (workSheetName === 'SHQ-Alert Report') {
        let obj: ShqAlertData = {
          alert: data[0],
          source: data[1].trim(),
          type: data[2],
          ip_address: data[1].match(/\((.*?)\)/)[1].trim(),
          severity: data[3].trim(),
          message: data[4].trim(),
          last_poll_time: moment(data[5]).format(),
          duration: data[6].trim(),
          duration_time: moment(data[7]).format(),
          total_duration_in_minutes: this.ShqService.CalucateTimeInMinutes(
            data[6]
          ),
        };
        result.push(obj);
      } else if (workSheetName === 'SHQ-NOC TT Report') {
        let obj: ShqTTData = {
          incident_id: data[0],
          parent_incident_id: data[1],
          enitity_type_name: data[2],
          entity_subtype_name: data[3],
          incident_name: data[4],
          equipment_host: data[5],
          ip: data[6] ? data[6].trim() : data[6],
          severity: data[7],
          status: data[8],
          priority_of_repair: data[9],
          effect_on_services: data[10],
          incident_type: data[11],
          mode_of_contact: data[12],
          incident_creation_time: data[13],
          remark_type: data[14],
          remarks: data[15],
          cluster: data[16],
          city: data[17],
          block: data[18],
          gp: data[19],
          slab_reach: data[20],
          resolution_method: data[21],
          rfo: data[22] ? data[22].trim() : data[22],
          incident_start_on: moment(data[23]).format(),
          incident_created_on: data[24],
          ageing: data[25],
          open_time: data[26],
          assigned_time: data[27],
          assigned_to_field: data[28],
          assigned_to_vendor: data[29],
          cancelled: data[30],
          closed: data[31],
          hold_time: data[32],
          resolved_date_time: data[33],
          resolved_by: data[34],
          total_resolution_time: data[35],
          resolution_type_in_min: data[36],
          sla_ageing: data[37],
          reporting_sla: data[38],
          reopen_date: data[39],
          category: data[40],
          change_id: data[41],
          exclusion_name: data[42],
          exclusion_remark: data[43],
          exclusion_type: data[44],
          pendency: data[45],
          vendor_name: data[46],
        };
        result.push(obj);
      }
    });

    if (workSheetName === 'SHQ-SLA-Report') {
      this.shqNMSData = result;
    } else if (workSheetName === 'SHQ-Alert Report') {
      this.shqAlertData = result;
    } else if (workSheetName === 'SHQ-NOC TT Report') {
      this.shqTTData = result;
    }
  }

  manipulateShqNmsData() {
    let manipulatedShqNmsData: ManipulatedShqNmsData[] = [];
    this.shqNMSData.forEach((nmsData: ShqNMSData) => {
      let totalUpTimeInMinutes = this.ShqService.CalucateTimeInMinutes(
        nmsData.up_time
      );
      let totalDownTimeInMinutes = this.ShqService.CalucateTimeInMinutes(
        nmsData.down_time
      );

      let totalTimeSlaExclusionInMinutes =
        totalUpTimeInMinutes + totalDownTimeInMinutes;
      let totalTimeSlaExclusionInPercent =
        nmsData.up_percent + nmsData.down_percent;
      let alertDownTimeInMinutes =
        this.ShqService.calculateAlertDownTimeInMinutes(
          nmsData.ip_address,
          this.shqAlertData
        );
      let alertDownTimeInPercent = +(
        (alertDownTimeInMinutes / totalTimeSlaExclusionInMinutes) *
        100
      ).toFixed(2);
      let rfoCategorizedData = this.ShqService.categorizeRFO(
        nmsData.ip_address,
        this.shqAlertData,
        this.shqTTData
      );
      let powerDownTimeInpercent = +(
        (rfoCategorizedData.total_power_downtime_minutes /
          totalTimeSlaExclusionInMinutes) *
        100
      ).toFixed(2);
      let dcnDownTimeInPercent = +(
        (rfoCategorizedData.total_dcn_downtime_minutes /
          totalTimeSlaExclusionInMinutes) *
        100
      ).toFixed(2);

      let newNmsData: ManipulatedShqNmsData = {
        ...nmsData,
        total_uptime_in_minutes: totalUpTimeInMinutes,
        total_downtime_in_minutes: totalDownTimeInMinutes,
        total_time_exclusive_of_sla_exclusions_in_min:
          totalTimeSlaExclusionInMinutes,
        total_time_exclusive_of_sla_exclusions_in_percent:
          totalTimeSlaExclusionInPercent,
        alert_downtime_in_minutes: alertDownTimeInMinutes,
        alert_downtime_in_percent: alertDownTimeInPercent,
        power_downtime_in_minutes:
          rfoCategorizedData.total_power_downtime_minutes,
        dcn_downtime_in_minutes: rfoCategorizedData.total_dcn_downtime_minutes,
        power_downtime_in_percent: powerDownTimeInpercent,
        dcn_downtime_in_percent: dcnDownTimeInPercent,
      };
      manipulatedShqNmsData.push(newNmsData);
    });

    this.manipulatedNMSData = manipulatedShqNmsData;
    console.log('manipulatedNMSData', this.manipulatedNMSData);
  }
}
