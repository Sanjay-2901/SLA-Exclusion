import { Component } from '@angular/core';
import * as ExcelJS from 'exceljs';
import {
  ManipulatedShqNmsData,
  ShqAlertData,
  ShqNMSData,
  ShqSlaSummary,
  ShqTTData,
} from './shq-component.model';
import * as moment from 'moment';
import { ShqService } from './shq-service.service';
import { AOA } from '../block-component/block-component.model';
import {
  IP_ADDRESS_PATTERN,
  SHQ_ALERT_REPORT_HEADERS,
  SHQ_INPUT_FILE_NAMES,
  SHQ_SLA_REPORT_HEADERS,
  TT_REPORT_HEADERS,
} from '../constants/constants';
import { ToastrService } from 'ngx-toastr';

@Component({
  selector: 'app-shq-component',
  templateUrl: './shq-component.component.html',
  styleUrls: ['./shq-component.component.scss', '../../styles.scss'],
})
export class ShqComponentComponent {
  shqNMSData: ShqNMSData[] = [];
  shqTTData: ShqTTData[] = [];
  shqAlertData: ShqAlertData[] = [];
  manipulatedNMSData: ManipulatedShqNmsData[] = [];
  worksheet!: ExcelJS.Worksheet;
  file!: any;
  shqSlaSummary!: ShqSlaSummary;
  isLoading: boolean = false;

  constructor(
    private ShqService: ShqService,
    private toastrService: ToastrService
  ) {}

  onFileChange(event: any) {
    this.isLoading = true;
    this.file = event.target.files[0];
    const workbook = new ExcelJS.Workbook();
    const reader = new FileReader();

    reader.onload = (e: any) => {
      const buffer = e.target.result;

      workbook.xlsx.load(buffer).then(() => {
        for (let index = 1; index <= workbook.worksheets.length; index++) {
          this.worksheet = workbook.getWorksheet(index);
          try {
            this.validateWorksheets(this.worksheet);
          } catch (error: any) {
            this.isLoading = false;
            this.resetInputFile();
            this.toastrService.error(error.message);
            break;
          }
        }
        if (
          this.shqNMSData.length > 0 &&
          this.shqAlertData.length > 0 &&
          this.shqTTData.length > 0
        ) {
          this.manipulateShqNmsData();
        }
      });
    };
    reader.readAsArrayBuffer(this.file);
  }

  resetInputFile(): void {
    this.isLoading = false;
    this.file = null;
    const fileInput = document.getElementById(
      'shqFileInput'
    ) as HTMLInputElement;
    if (fileInput) {
      fileInput.value = '';
    }
    this.shqAlertData = [];
    this.shqNMSData = [];
    this.shqTTData = [];
    this.ShqService.ttCorelation = [];
  }

  validateWorksheets(worksheet: ExcelJS.Worksheet) {
    let workSheetName = worksheet.name;
    if (!SHQ_INPUT_FILE_NAMES.includes(workSheetName)) {
      throw new Error(
        'SHQ - Invalid sheet name of the input file. Kindly provide the valid sheet names.'
      );
    } else {
      let data: AOA = [];
      this.worksheet.eachRow({ includeEmpty: true }, (row: ExcelJS.Row) => {
        const rowData: any = [];
        row.eachCell({ includeEmpty: true }, (cell: ExcelJS.Cell) => {
          rowData.push(cell.value);
        });
        data.push(rowData);
      });

      const headers = JSON.stringify(data[0]);

      if (workSheetName === 'shq_sla_report') {
        if (headers !== JSON.stringify(SHQ_SLA_REPORT_HEADERS)) {
          throw new Error(
            'SHQ - Invalid template of the SLA report. Kindly provide the valid column names.'
          );
        } else {
          try {
            this.validateEachRowsInSlaReport(data, workSheetName);
          } catch (error: any) {
            this.toastrService.error(error.message);
            this.resetInputFile();
          }
        }
      } else if (workSheetName === 'shq_noc_tt_report') {
        if (headers !== JSON.stringify(TT_REPORT_HEADERS)) {
          throw new Error(
            'SHQ - Invalid template of the TT report. Kindly provide the valid column names.'
          );
        } else {
          this.storeDataAsObject(workSheetName, data);
        }
      } else if (workSheetName === 'shq_alert_report') {
        if (headers !== JSON.stringify(SHQ_ALERT_REPORT_HEADERS)) {
          throw new Error(
            'SHQ - Invalid template of the Alert report. Kindly provide the valid column names.'
          );
        } else {
          this.storeDataAsObject(workSheetName, data);
        }
      }
    }
  }

  validateEachRowsInSlaReport(data: AOA, workSheetName: string) {
    for (let index = 1; index < data.length; index++) {
      let row: any = data[index];
      if (row[0] === null || row[0] === undefined) {
        throw new Error(`SHQ - ${
          SHQ_SLA_REPORT_HEADERS[0]
        } is not available in SLA report in row number :
          ${index + 1}`);
      } else if (row[1] === null || row[1] === undefined) {
        throw new Error(`SHQ - ${
          SHQ_SLA_REPORT_HEADERS[1]
        } is not available in SLA report in row number :
          ${index + 1}`);
      } else if (row[4] === null || row[4] === undefined) {
        throw new Error(`SHQ - ${
          SHQ_SLA_REPORT_HEADERS[4]
        } is not available in SLA report in row number :
          ${index + 1}`);
      } else if (row[5] === null || row[5] === undefined) {
        throw new Error(`SHQ - ${
          SHQ_SLA_REPORT_HEADERS[5]
        } is not available in SLA report in row number :
          ${index + 1}`);
      } else if (row[6] === null || row[6] === undefined) {
        throw new Error(`SHQ - ${
          SHQ_SLA_REPORT_HEADERS[6]
        } is not available in SLA report in row number :
          ${index + 1}`);
      } else if (row[7] === null || row[7] === undefined) {
        throw new Error(`SHQ - ${
          SHQ_SLA_REPORT_HEADERS[7]
        } is not available in SLA report in row number :
          ${index + 1}`);
      } else {
        if (!IP_ADDRESS_PATTERN.test(row[1].trim())) {
          throw new Error(
            `SHQ - ${
              SHQ_SLA_REPORT_HEADERS[1]
            } is invalid in SLA report in row number : ${index + 1}`
          );
        }
      }
    }

    this.storeDataAsObject(workSheetName, data);
  }

  storeDataAsObject(workSheetName: string, data: any[]): void {
    let result: any = [];
    data.forEach((data: any, index: number) => {
      if (index >= 1) {
        if (workSheetName === 'shq_sla_report') {
          let obj: ShqNMSData = {
            monitor: data[0] ? data[0].trim() : data[0],
            ip_address: data[1] ? data[1].trim() : data[1],
            departments: data[2],
            type: data[3],
            up_percent: data[4],
            up_time: data[5],
            down_percent: data[6],
            down_time: data[7],
            created_date: data[8],
          };
          result.push(obj);
        } else if (workSheetName === 'shq_alert_report') {
          let obj: ShqAlertData = {
            alert: data[0],
            source: data[1] ? data[1].trim() : data[1],
            ip_address: data[2] ? data[2].trim() : data[2],
            type: data[3],
            severity: data[4] ? data[4].trim() : data[4],
            message: data[5] ? data[5].trim() : data[5],
            alarm_start_time: moment(data[6]).format(),
            duration: data[7] ? data[7].trim() : data[7],
            alarm_clear_time: moment(data[8]).format(),
            total_duration_in_minutes: this.ShqService.CalucateTimeInMinutes(
              data[7]
            ),
          };
          result.push(obj);
        } else if (workSheetName === 'shq_noc_tt_report') {
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
      }
    });

    if (workSheetName === 'shq_sla_report') {
      this.shqNMSData = this.ShqService.shqNMSDatawithoutVmwareDevices(result);
    } else if (workSheetName === 'shq_alert_report') {
      this.shqAlertData = result;
    } else if (workSheetName === 'shq_noc_tt_report') {
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
        nmsData,
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
      let unknownDownTimeInMinutes =
        rfoCategorizedData.alert_report_empty === true
          ? totalDownTimeInMinutes
          : totalDownTimeInMinutes - alertDownTimeInMinutes <= 15
          ? 0
          : totalDownTimeInMinutes - alertDownTimeInMinutes;

      let unknownDownTimeInPercent = +(
        (unknownDownTimeInMinutes / totalTimeSlaExclusionInMinutes) *
        100
      ).toFixed(2);

      let pollingTimeMinutes = 0;

      if (
        alertDownTimeInMinutes < totalDownTimeInMinutes &&
        totalDownTimeInMinutes - alertDownTimeInMinutes <= 15
      ) {
        pollingTimeMinutes = totalDownTimeInMinutes - alertDownTimeInMinutes;
      }

      if (alertDownTimeInMinutes > totalDownTimeInMinutes) {
        pollingTimeMinutes = alertDownTimeInMinutes - totalDownTimeInMinutes;
      }

      let pollingTimePercent =
        pollingTimeMinutes > 0
          ? +(
              (pollingTimeMinutes / totalTimeSlaExclusionInMinutes) *
              100
            ).toFixed(2)
          : 0;

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
        unknown_downtime_in_minutes: unknownDownTimeInMinutes,
        power_downtime_in_percent: powerDownTimeInpercent,
        dcn_downtime_in_percent: dcnDownTimeInPercent,
        unknown_downtime_in_percent: unknownDownTimeInPercent,
        polling_time_in_minutes: pollingTimeMinutes,
        polling_time_in_percent: pollingTimePercent,
      };
      manipulatedShqNmsData.push(newNmsData);
    });

    this.manipulatedNMSData = manipulatedShqNmsData;
    this.shqSlaSummary = this.ShqService.calculateShqSlaSummary(
      this.manipulatedNMSData
    );
    this.generateFinalBlockReport();
  }

  generateFinalBlockReport() {
    const workbook = new ExcelJS.Workbook();
    const worksheet = workbook.addWorksheet('SHQ-SLA-Exclusion-Report');
    this.ShqService.FrameShqFinalSlaReportWorkbook(
      workbook,
      worksheet,
      this.shqSlaSummary,
      this.manipulatedNMSData
    );
    workbook.xlsx.writeBuffer().then((buffer) => {
      this.ShqService.downloadFinalReport(buffer, 'SHQ-SLA-Exclusion-Report');
      this.isLoading = false;
      this.resetInputFile();
    });
  }
}
