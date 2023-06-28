import { Component } from '@angular/core';
import * as ExcelJS from 'exceljs';
import {
  GP_ALERT_REPORT_HEADERS,
  GP_INPUT_FILE_NAMES,
  GP_SLA_REPORT_HEADERS,
  IP_ADDRESS_PATTERN,
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
  manipulatedNMSData: ManipulatedGpNMSData[] = [];

  constructor(
    private sharedService: SharedService,
    private gpService: GpService,
    private toastrService: ToastrService
  ) {}

  onFileChange(event: any): void {
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
          this.gpNMSData.length > 0 &&
          this.gpAlertData.length > 0 &&
          this.gpTTData.length > 0
        ) {
          this.manipulateGpNMSData();
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
      this.worksheet.eachRow({ includeEmpty: true }, (row: ExcelJS.Row) => {
        const rowData: any = [];
        row.eachCell({ includeEmpty: true }, (cell: ExcelJS.Cell) => {
          rowData.push(cell.value);
        });
        data.push(rowData);
      });

      const headers = JSON.stringify(data[0]);

      if (workSheetName === GP_INPUT_FILE_NAMES[0]) {
        if (headers !== JSON.stringify(GP_SLA_REPORT_HEADERS)) {
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
      }
    }
  }

  validateEachRowsOfSlaReport(data: AOA, workSheetName: string) {
    for (let index = 1; index < data.length; index++) {
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

  storeDataAsObject(workSheetName: string, data: any) {}

  manipulateGpNMSData(): void {
    let manipulatedGpNMSData: ManipulatedGpNMSData[] = [];
    this.gpNMSData.forEach((nmsData: GpNMSData) => {
      let totalUpTimeInMinutes = this.sharedService.CalucateTimeInMinutes(
        nmsData.total_up_time
      );
      let totalDownTimeInMinutes = this.sharedService.CalucateTimeInMinutes(
        nmsData.down_time
      );
      let plannedMaintenanceInMinutes =
        this.sharedService.CalucateTimeInMinutes(nmsData.maintenance_time);
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
        this.gpTTData
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
          : totalDownTimeInMinutes - alertDownTimeInMinutes;

      let unknownDownTimeInPercent = +(
        (unknownDownTimeInMinutes / totalTimeExclusiveOfSLAExclusionInMinutes) *
        100
      ).toFixed(2);

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
      };
      manipulatedGpNMSData.push(newNMSData);
    });
    this.manipulatedNMSData = manipulatedGpNMSData;
    this.gpSlaSummary = this.gpService.calculateGpSlaSummary(
      this.manipulatedNMSData
    );
    this.generateFinalBlockReport();
  }

  resetInputFile(): void {
    this.isLoading = false;
    this.file = null;
    const fileInput = document.getElementById(
      'gpFileInput'
    ) as HTMLInputElement;
    if (fileInput) {
      fileInput.value = '';
    }
    this.gpAlertData = [];
    this.gpNMSData = [];
    this.gpTTData = [];
  }

  generateFinalBlockReport() {
    const workbook = new ExcelJS.Workbook();
    const worksheet = workbook.addWorksheet('GP-SLA-Exclusion-Report');
    this.gpService.FrameGpFinalSlaReportWorkbook(
      worksheet,
      this.gpSlaSummary,
      this.manipulatedNMSData
    );
    workbook.xlsx.writeBuffer().then((buffer) => {
      this.sharedService.downloadFinalReport(buffer, 'GP-SLA-Exclusion-Report');
      this.isLoading = false;
      this.resetInputFile();
    });
  }
}
