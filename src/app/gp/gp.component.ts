import { Component, EventEmitter, Input, Output } from '@angular/core';
import * as ExcelJS from 'exceljs';
import {
  BLOCK_ALERT_REPORT_HEADERS,
  BLOCK_TT_CO_RELATION_HEADERS,
  DEVICES_COUNT,
  GP_ALERT_REPORT_HEADERS,
  GP_INPUT_FILE_NAMES,
  GP_SLA_REPORT_HEADERS,
  IP_ADDRESS_PATTERN,
  TIME_SPAN_REGEX_PATTERN,
  TT_REPORT_HEADERS,
} from '../constants/constants';
import { AOA } from '../shared/shared-model';
import { ToastrService } from 'ngx-toastr';
import {
  GpAlertData,
  GpNMSData,
  GpSLASummary,
  GpTTData,
  ManipulatedGpNMSData,
} from './gp.model';
import { SharedService } from '../shared/shared.service';
import { GpService } from './gp.service';
import * as moment from 'moment';
import {
  BlockAlertData,
  BlockDeviceLevelHeaders,
  TTCorelation,
} from '../block-component/block-component.model';
import * as lodash from 'lodash';

@Component({
  selector: 'app-gp',
  templateUrl: './gp.component.html',
  styleUrls: ['./gp.component.scss'],
})
export class GpComponent {
  isLoading: boolean = false;
  file!: any;
  worksheet!: ExcelJS.Worksheet;
  gpAlertData: GpAlertData[] = [];
  gpNMSData: GpNMSData[] = [];
  gpTTData: GpTTData[] = [];
  gpSlaSummary!: GpSLASummary;
  gpSlaSummaryWithAlerts!: GpSLASummary;
  gpSlaSummaryWithoutAlerts!: GpSLASummary;
  manipulatedNMSData: ManipulatedGpNMSData[] = [];
  blockFinalReport: BlockDeviceLevelHeaders[] = [];
  blockTTCorelationReport: TTCorelation[] = [];
  blockAlertData: BlockAlertData[] = [];
  timeSpanValue: string = '';
  isAllFilesValid: boolean = true;
  @Output() isGpLoading = new EventEmitter<boolean>();
  @Input() shouldDisable!: boolean;

  constructor(
    private sharedService: SharedService,
    private gpService: GpService,
    private toastrService: ToastrService
  ) {}

  onFileChange(event: any): void {
    this.isLoading = true;
    this.isGpLoading.emit(true);
    this.file = event.target.files[0];
    const workbook = new ExcelJS.Workbook();
    const reader = new FileReader();

    reader.onload = (e: any) => {
      const buffer = e.target.result;

      workbook.xlsx.load(buffer).then(() => {
        if (workbook.worksheets.length !== 6) {
          this.toastrService.error(
            'Some files cannot be found. Please upload all necessary files'
          );
          this.resetInputFile();
        } else {
          for (let index = 1; index <= workbook.worksheets.length; index++) {
            this.worksheet = workbook.getWorksheet(index);
            try {
              this.validateWorksheets(this.worksheet);
            } catch (error: any) {
              this.isLoading = false;
              this.isAllFilesValid = false;
              this.isGpLoading.emit(false);
              this.resetInputFile();
              this.toastrService.error(error.message);
              break;
            }
          }
        }

        if (this.isAllFilesValid) {
          if (this.gpNMSData.length === DEVICES_COUNT.GP) {
            this.manipulateGpNMSData();
          } else {
            this.resetInputFile();
            this.toastrService.error(
              'NMS data is insufficient. Please provide the correct data.'
            );
          }
        } else {
          this.isAllFilesValid = true;
        }
      });
    };
    reader.readAsArrayBuffer(this.file);
  }

  validateWorksheets(worksheet: ExcelJS.Worksheet) {
    let workSheetName = worksheet.name;
    if (!GP_INPUT_FILE_NAMES.includes(workSheetName)) {
      throw new Error(
        'GP - Invalid sheet name of the input file. Kindly provide the valid sheet names.'
      );
    } else {
      let data: AOA = [];
      this.worksheet.eachRow({ includeEmpty: false }, (row: ExcelJS.Row) => {
        const rowData: any = [];
        row.eachCell({ includeEmpty: true }, (cell: ExcelJS.Cell) => {
          rowData.push(cell.value);
        });
        data.push(rowData);
      });

      const headers = JSON.stringify(data[0]);

      if (workSheetName === GP_INPUT_FILE_NAMES[0]) {
        const timeSpanRow: string[] = data[0];
        this.timeSpanValue = timeSpanRow[0];
        const slaReportHeader = JSON.stringify(data[1]);

        if (!TIME_SPAN_REGEX_PATTERN.test(this.timeSpanValue)) {
          throw new Error(
            'GP - The Time Span value in the first column is either incorrect or unavailable. Please provide a valid Time Span.'
          );
        }
        if (slaReportHeader !== JSON.stringify(GP_SLA_REPORT_HEADERS)) {
          throw new Error(
            'GP - Invalid template of the SLA report. Kindly provide the valid column names.'
          );
        } else {
          try {
            this.validateEachRowsOfSlaReport(data, workSheetName);
          } catch (error: any) {
            this.toastrService.error(error.message);
            this.resetInputFile();
          }
        }
      } else if (workSheetName === GP_INPUT_FILE_NAMES[1]) {
        if (headers !== JSON.stringify(TT_REPORT_HEADERS)) {
          throw new Error(
            'GP - Invalid template of the TT report. Kindly provide the valid column names.'
          );
        } else {
          this.storeDataAsObject(workSheetName, data);
        }
      } else if (workSheetName === GP_INPUT_FILE_NAMES[2]) {
        if (headers !== JSON.stringify(GP_ALERT_REPORT_HEADERS)) {
          throw new Error(
            'GP - Invalid template of the Alert report. Kindly provide the valid column names.'
          );
        } else {
          this.storeDataAsObject(workSheetName, data);
        }
      } else if (workSheetName === GP_INPUT_FILE_NAMES[3]) {
        this.storeDataAsObject(workSheetName, data);
      } else if (workSheetName === GP_INPUT_FILE_NAMES[4]) {
        if (headers !== JSON.stringify(BLOCK_TT_CO_RELATION_HEADERS)) {
          throw new Error(
            'GP - Invalid template of the Co-Relation report. Kindly provide the valid column names.'
          );
        } else {
          this.storeDataAsObject(workSheetName, data);
        }
      } else if (workSheetName === GP_INPUT_FILE_NAMES[5]) {
        if (headers !== JSON.stringify(BLOCK_ALERT_REPORT_HEADERS)) {
          throw new Error(
            'Block - Invalid template of the  Alert report. Kindly provide the valid column names.'
          );
        } else {
          try {
            this.validateEachRowsOfAlertReport(data, workSheetName);
          } catch (error: any) {
            this.toastrService.error(error.message);
            this.resetInputFile();
          }
        }
      }
    }
  }

  validateEachRowsOfAlertReport(data: AOA, workSheetName: string) {
    data.forEach((row: any, index: number) => {
      if (index >= 1) {
        if (!IP_ADDRESS_PATTERN.test(row[2])) {
          throw new Error(
            ` BLOCK - ${
              BLOCK_ALERT_REPORT_HEADERS[2]
            } is invalid in Alert report in row number : ${index + 1}`
          );
        }
      }
    });
    this.storeDataAsObject(workSheetName, data);
  }

  validateEachRowsOfSlaReport(data: AOA, workSheetName: string) {
    for (let index = 2; index < data.length; index++) {
      let row: any = data[index];
      if (row[1] === null || row[1] === undefined) {
        throw new Error(`GP - ${
          GP_SLA_REPORT_HEADERS[1]
        } is not available in SLA report in row number:
        ${index + 1}`);
      } else if (row[4] === null || row[4] === undefined) {
        throw new Error(`GP - ${
          GP_SLA_REPORT_HEADERS[4]
        } is not available in SLA report in row number:
          ${index + 1}`);
      } else if (row[5] === null || row[5] === undefined) {
        throw new Error(`GP - ${
          GP_SLA_REPORT_HEADERS[5]
        } is not available in SLA report in row number:
          ${index + 1}`);
      } else if (row[6] === null || row[6] === undefined) {
        throw new Error(`GP - ${
          GP_SLA_REPORT_HEADERS[6]
        } is not available in SLA report in row number:
          ${index + 1}`);
      } else if (row[7] === null || row[7] === undefined) {
        throw new Error(`GP - ${
          GP_SLA_REPORT_HEADERS[7]
        } is not available in SLA report in row number:
          ${index + 1}`);
      } else if (row[8] === null || row[8] === undefined) {
        throw new Error(`GP - ${
          GP_SLA_REPORT_HEADERS[8]
        } is not available in SLA report in row number:
          ${index + 1}`);
      } else if (row[9] === null || row[9] === undefined) {
        throw new Error(`GP - ${
          GP_SLA_REPORT_HEADERS[5]
        } is not available in SLA report in row number:
          ${index + 1}`);
      } else {
        if (!IP_ADDRESS_PATTERN.test(row[1].trim())) {
          throw new Error(
            ` GP - ${
              GP_SLA_REPORT_HEADERS[1]
            } is not valid in SLA report in row number : ${index + 1}`
          );
        }
      }
    }
    this.storeDataAsObject(workSheetName, data);
  }

  storeDataAsObject(workSheetName: string, data: any) {
    let result: any = [];
    data.forEach((data: any, index: number) => {
      if (workSheetName === GP_INPUT_FILE_NAMES[0] && index >= 2) {
        let obj: GpNMSData = {
          monitor: data[0],
          ip_address: data[1] ? data[1].trim() : data[1],
          departments: data[2],
          type: data[3],
          up_percent: data[4],
          up_time: this.sharedService.formatTimeInSlaReport(data[5]),
          down_percent: data[6],
          down_time: this.sharedService.formatTimeInSlaReport(data[7]),
          maintenance_percent: data[8],
          maintenance_time: this.sharedService.formatTimeInSlaReport(data[9]),
          total_up_percent: data[10],
          total_up_time: this.sharedService.formatTimeInSlaReport(data[11]),
          created_date: data[12],
        };
        result.push(obj);
      } else if (workSheetName === GP_INPUT_FILE_NAMES[1] && index >= 1) {
        let obj: GpTTData = {
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
          incident_start_on: this.sharedService.setStandardTime(data[23]),
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
          resolution_time_in_min: data[36],
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
      } else if (workSheetName === GP_INPUT_FILE_NAMES[2] && index >= 1) {
        let obj: GpAlertData = {
          alert: data[0],
          source: data[1],
          ip_address: data[2],
          departments: data[3],
          type: data[4],
          severity: data[5] ? data[5].trim() : data[5],
          message: data[6] ? data[6].trim() : data[6],
          alarm_start_time: this.sharedService.setStandardTime(data[7]),
          duration: this.sharedService.setDuration(
            this.timeSpanValue,
            this.sharedService.setStandardTime(data[7]),
            this.sharedService.setStandardTime(data[9]),
            data[8]
          ),
          alarm_clear_time: this.sharedService.setStandardTime(data[9]),
          total_duration_in_minutes: this.sharedService.calculateTimeInMinutes(
            this.sharedService.setDuration(
              this.timeSpanValue,
              this.sharedService.setStandardTime(data[7]),
              this.sharedService.setStandardTime(data[9]),
              data[8]
            )
          ),
        };
        result.push(obj);
      } else if (workSheetName === GP_INPUT_FILE_NAMES[3] && index >= 11) {
        let obj: BlockDeviceLevelHeaders = {
          report_type: data[0],
          host_name: data[1],
          ip_address: data[2],
          state: data[3],
          cluster: data[4],
          district: data[5],
          district_lgd_code: data[6],
          block_name: data[7],
          block_lgd_code: data[8],
          no_of_gp_in_block: data[9],
          up_percent: data[10],
          up_minute: data[11],
          total_down_exclusive_of_sla_exclusion_percent: data[12],
          total_down_exclusive_of_sla_exclusion_minute: data[13],
          power_down_percent: data[14],
          power_down_minute: data[15],
          fibre_down_percent: data[16],
          fibre_down_minute: data[17],
          equipement_down_percent: data[18],
          equipement_down_minute: data[19],
          hrt_down_percent: data[20],
          hrt_down_minute: data[21],
          dcn_down_percent: data[22],
          dcn_down_miniute: data[23],
          planned_maintenance_percent: data[24],
          planned_maintenance_minute: data[25],
          unknown_percent: data[26],
          unknown_minute: data[27],
          total_sla_exclusion_percent: data[28],
          total_sla_exclusion_minute: data[29],
          polling_time_percent: data[30],
          polling_time_minute: data[31],
          total_up_percent: data[32],
          total_up_minute: data[33],
        };
        result.push(obj);
      } else if (workSheetName === GP_INPUT_FILE_NAMES[4] && index >= 1) {
        let obj: TTCorelation = {
          ip: data[1],
          powerIssueTT: [data[3]],
          linkIssueTT: [data[4]],
          otherTT: [data[5]],
        };
        result.push(obj);
      } else if (workSheetName === GP_INPUT_FILE_NAMES[5] && index >= 1) {
        let obj: BlockAlertData = {
          alert: data[0],
          source: data[1],
          ip_address: data[2] ? data[2].trim() : data[2],
          departments: data[3],
          type: data[4],
          severity: data[5] ? data[5].trim() : data[5],
          message: data[6] ? data[6].trim() : data[6],
          alarm_start_time: this.sharedService.setStandardTime(data[7]),
          duration: this.sharedService.setDuration(
            this.timeSpanValue,
            this.sharedService.setStandardTime(data[7]),
            this.sharedService.setStandardTime(data[9]),
            data[8]
          ),
          alarm_clear_time: this.sharedService.setStandardTime(data[9]),
          total_duration_in_minutes: this.sharedService.calculateTimeInMinutes(
            this.sharedService.setDuration(
              this.timeSpanValue,
              this.sharedService.setStandardTime(data[7]),
              this.sharedService.setStandardTime(data[9]),
              data[8]
            )
          ),
        };
        result.push(obj);
      }
    });

    if (workSheetName === GP_INPUT_FILE_NAMES[0]) {
      this.gpNMSData = result;
    } else if (workSheetName === GP_INPUT_FILE_NAMES[1]) {
      this.gpTTData = result;
    } else if (workSheetName === GP_INPUT_FILE_NAMES[2]) {
      this.gpAlertData = result;
    } else if (workSheetName === GP_INPUT_FILE_NAMES[3]) {
      this.blockFinalReport = result;
    } else if (workSheetName === GP_INPUT_FILE_NAMES[4]) {
      this.blockTTCorelationReport = result;
    } else if (workSheetName === GP_INPUT_FILE_NAMES[5]) {
      this.blockAlertData = result;
    }
  }

  manipulateGpNMSData(): void {
    let manipulatedGpNMSData: ManipulatedGpNMSData[] = [];
    this.gpNMSData.forEach((nmsData: GpNMSData) => {
      let totalUpTimeInMinutes = this.sharedService.calculateTimeInMinutes(
        nmsData.total_up_time
      );
      let totalDownTimeInMinutes = this.sharedService.calculateTimeInMinutes(
        nmsData.down_time
      );
      let plannedMaintenanceInMinutes =
        this.sharedService.calculateTimeInMinutes(nmsData.maintenance_time);
      let totalTimeExclusiveOfSLAExclusionInMinutes =
        totalUpTimeInMinutes + totalDownTimeInMinutes;
      let totalTimeExclusiveOfSLAExclusionInPercent =
        nmsData.up_percent + nmsData.down_percent;
      let alertDownTimeInMinutes =
        this.gpService.calculateAlertDownTimeInMinutes(
          nmsData.ip_address,
          this.gpAlertData
        );
      let alertDownTimeInPercent = +(
        (alertDownTimeInMinutes / totalTimeExclusiveOfSLAExclusionInMinutes) *
        100
      ).toFixed(2);
      let rfoCategorizedData = this.gpService.categorizeRFO(
        nmsData,
        this.gpAlertData,
        this.gpTTData,
        this.blockAlertData
      );
      let powerDownTimeInpercent = +(
        (rfoCategorizedData.total_power_downtime_minutes /
          totalTimeExclusiveOfSLAExclusionInMinutes) *
        100
      ).toFixed(2);
      let dcnDownTimeInPercent = +(
        (rfoCategorizedData.total_dcn_downtime_minutes /
          totalTimeExclusiveOfSLAExclusionInMinutes) *
        100
      ).toFixed(2);
      let unknownDownTimeInMinutes =
        rfoCategorizedData.alert_report_empty === true
          ? totalDownTimeInMinutes
          : totalDownTimeInMinutes - alertDownTimeInMinutes <= 15
          ? 0
          : totalDownTimeInMinutes - alertDownTimeInMinutes;

      let unknownDownTimeInPercent = +(
        (unknownDownTimeInMinutes / totalTimeExclusiveOfSLAExclusionInMinutes) *
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
              (pollingTimeMinutes / totalTimeExclusiveOfSLAExclusionInMinutes) *
              100
            ).toFixed(2)
          : 0;

      let newNMSData: ManipulatedGpNMSData = {
        ...nmsData,
        total_uptime_in_minutes: totalUpTimeInMinutes,
        total_downtime_in_minutes: totalDownTimeInMinutes,
        total_time_exclusive_of_sla_exclusions_in_min:
          totalTimeExclusiveOfSLAExclusionInMinutes,
        total_time_exclusive_of_sla_exclusions_in_percent:
          totalTimeExclusiveOfSLAExclusionInPercent,
        alert_downtime_in_minutes: alertDownTimeInMinutes,
        alert_downtime_in_percent: alertDownTimeInPercent,
        power_downtime_in_minutes:
          rfoCategorizedData.total_power_downtime_minutes,
        dcn_downtime_in_minutes: rfoCategorizedData.total_dcn_downtime_minutes,
        planned_maintenance_in_minutes: plannedMaintenanceInMinutes,
        unknown_downtime_in_minutes: unknownDownTimeInMinutes,
        power_downtime_in_percent: powerDownTimeInpercent,
        dcn_downtime_in_percent: dcnDownTimeInPercent,
        planned_maintenance_in_percent: nmsData.maintenance_percent,
        unknown_downtime_in_percent: unknownDownTimeInPercent,
        polling_time_in_minutes: pollingTimeMinutes,
        polling_time_in_percent: pollingTimePercent,
      };
      manipulatedGpNMSData.push(newNMSData);
    });
    this.manipulatedNMSData = manipulatedGpNMSData;

    this.gpSlaSummary = this.gpService.calculateGpSlaSummary(
      this.manipulatedNMSData,
      this.timeSpanValue
    );
    let gpNmsDataWithoutAlerts = this.manipulatedNMSData.filter(
      (gpNmsData: ManipulatedGpNMSData) =>
        gpNmsData.down_percent == 100 &&
        gpNmsData.alert_downtime_in_minutes == 0 &&
        gpNmsData.unknown_downtime_in_percent == 100
    );

    let gpNmsDataWithAlerts = this.manipulatedNMSData.filter(
      (gpNmsData: ManipulatedGpNMSData) =>
        !lodash.some(gpNmsDataWithoutAlerts, gpNmsData)
    );

    this.gpSlaSummaryWithoutAlerts = this.gpService.calculateGpSlaSummary(
      gpNmsDataWithoutAlerts,
      this.timeSpanValue
    );
    this.gpSlaSummaryWithAlerts = this.gpService.calculateGpSlaSummary(
      gpNmsDataWithAlerts,
      this.timeSpanValue
    );

    this.generateFinalGpReport();
  }

  resetInputFile(): void {
    this.isLoading = false;
    this.isGpLoading.emit(false);
    this.file = null;
    const fileInput = document.getElementById(
      'gpFileInput'
    ) as HTMLInputElement;
    if (fileInput) {
      fileInput.value = '';
    }
    this.timeSpanValue = '';
    this.gpAlertData = [];
    this.gpNMSData = [];
    this.gpTTData = [];
    this.gpService.ttCorelation = [];
  }

  generateFinalGpReport() {
    const workbook = new ExcelJS.Workbook();
    const worksheet = workbook.addWorksheet('GP-SLA-Exclusion-Report');
    this.gpService.FrameGpFinalSlaReportWorkbook(
      workbook,
      worksheet,
      this.gpSlaSummary,
      this.gpSlaSummaryWithAlerts,
      this.gpSlaSummaryWithoutAlerts,
      this.manipulatedNMSData,
      this.blockFinalReport,
      this.blockTTCorelationReport
    );
    workbook.xlsx.writeBuffer().then((buffer) => {
      this.sharedService.downloadFinalReport(buffer, 'GP-SLA-Exclusion-Report');
      this.isLoading = false;
      this.isGpLoading.emit(false);
      this.resetInputFile();
    });
  }
}
