import { Component, EventEmitter, Input, Output } from '@angular/core';
import * as ExcelJS from 'exceljs';
import * as moment from 'moment';
import * as lodash from 'lodash';
import {
  BlockAlertData,
  BlockNMSData,
  BlockSLASummary,
  BlockTTData,
  ManipulatedNMSData,
  RFOCategorizedTimeInMinutes,
  TTCorelation,
} from './block-component.model';
import {
  SEVERITY_CRITICAL,
  ALERT_DOWN_MESSAGE,
  SEVERITY_WARNING,
  RFO_CATEGORIZATION,
  BLOCK_SLA_REPORT_HEADERS,
  TT_REPORT_HEADERS,
  BLOCK_ALERT_REPORT_HEADERS,
  BLOCK_INPUT_FILE_NAMES,
  IP_ADDRESS_PATTERN,
  DEVICES_COUNT,
  TIME_SPAN_REGEX_PATTERN,
} from '../constants/constants';
import { ToastrService } from 'ngx-toastr';
import { AOA } from '../shared/shared-model';
import { SharedService } from '../shared/shared.service';

import { BlockService } from './block.service';

@Component({
  selector: 'app-block-component',
  templateUrl: './block-component.component.html',
  styleUrls: ['./block-component.component.scss', '../../styles.scss'],
})
export class BlockComponentComponent {
  blockNMSData: BlockNMSData[] = [];
  blockTTData: BlockTTData[] = [];
  blockAlertData: BlockAlertData[] = [];

  manipulatedNMSData: ManipulatedNMSData[] = [];
  blockSLASummary!: BlockSLASummary;
  blockSLASummaryWithoutAlerts!: BlockSLASummary;
  blockSLASummaryWithAlerts!: BlockSLASummary;

  ttCorelation: TTCorelation[] = [];

  worksheet!: ExcelJS.Worksheet;
  file!: any;
  isSheetNamesValid: boolean = true;
  isLoading: boolean = false;
  timeSpanValue: string = '';
  isAllFilesValid: boolean = true;

  @Output() isBlockLoading = new EventEmitter<boolean>();
  @Input() shouldDisable!: boolean;

  constructor(
    private toastrService: ToastrService,
    private sharedService: SharedService,
    private blockService: BlockService
  ) {}

  // Getting the input file (excel workbook containing the required sheets)
  onFileChange(event: any): void {
    this.isLoading = true;
    this.isBlockLoading.emit(true);
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
            this.isAllFilesValid = false;
            this.isBlockLoading.emit(false);
            this.resetInputFile();
            this.toastrService.error(error.message);
            break;
          }
        }

        if (this.isAllFilesValid) {
          if (this.blockNMSData.length === DEVICES_COUNT.BLOCK) {
            this.manipulateBlockNMSData();
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

  resetInputFile(): void {
    this.isLoading = false;
    this.isBlockLoading.emit(false);
    this.file = null;
    const fileInput = document.getElementById(
      'blockFileInput'
    ) as HTMLInputElement;
    if (fileInput) {
      fileInput.value = '';
    }
    this.timeSpanValue = '';
    this.blockAlertData = [];
    this.blockNMSData = [];
    this.blockTTData = [];
    this.ttCorelation = [];
  }

  validateWorksheets(worksheet: ExcelJS.Worksheet) {
    let workSheetName = worksheet.name;
    if (!BLOCK_INPUT_FILE_NAMES.includes(workSheetName)) {
      throw new Error(
        'Block - Invalid sheet name of the input file. Kindly provide the valid sheet names.'
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

      if (workSheetName === 'block_sla_report') {
        const timeSpanRow: string[] = data[0];
        this.timeSpanValue = timeSpanRow[0];
        const slaReportHeader = JSON.stringify(data[1]);
        if (!TIME_SPAN_REGEX_PATTERN.test(this.timeSpanValue)) {
          throw new Error(
            'BLOCK - The Time Span value in the first column is either incorrect or unavailable. Please provide a valid Time Span.'
          );
        }
        if (slaReportHeader !== JSON.stringify(BLOCK_SLA_REPORT_HEADERS)) {
          throw new Error(
            'Block - Invalid template of the SLA report. Kindly provide the valid column names.'
          );
        } else {
          try {
            this.validateEachRowsOfSlaReport(data, workSheetName);
          } catch (error: any) {
            this.toastrService.error(error.message);
            this.resetInputFile();
          }
        }
      } else if (workSheetName === 'block_noc_tt_report') {
        if (headers !== JSON.stringify(TT_REPORT_HEADERS)) {
          throw new Error(
            'Block - Invalid template of the TT report. Kindly provide the valid column names.'
          );
        } else {
          this.storeDataAsObject(workSheetName, data);
        }
      } else if (workSheetName === 'block_alert_report') {
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

  validateEachRowsOfSlaReport(data: AOA, workSheetName: string) {
    for (let index = 2; index < data.length; index++) {
      let row: any = data[index];
      if (row[1] === null || row[1] === undefined) {
        throw new Error(`Block - ${
          BLOCK_SLA_REPORT_HEADERS[1]
        } is not available in SLA report in row number:
        ${index + 1}`);
      } else if (row[4] === null || row[4] === undefined) {
        throw new Error(`Block - ${
          BLOCK_SLA_REPORT_HEADERS[4]
        } is not available in SLA report in row number:
          ${index + 1}`);
      } else if (row[5] === null || row[5] === undefined) {
        throw new Error(`Block - ${
          BLOCK_SLA_REPORT_HEADERS[5]
        } is not available in SLA report in row number:
          ${index + 1}`);
      } else if (row[6] === null || row[6] === undefined) {
        throw new Error(`Block - ${
          BLOCK_SLA_REPORT_HEADERS[6]
        } is not available in SLA report in row number:
          ${index + 1}`);
      } else if (row[7] === null || row[7] === undefined) {
        throw new Error(`Block - ${
          BLOCK_SLA_REPORT_HEADERS[7]
        } is not available in SLA report in row number:
          ${index + 1}`);
      } else if (row[8] === null || row[8] === undefined) {
        throw new Error(`Block - ${
          BLOCK_SLA_REPORT_HEADERS[8]
        } is not available in SLA report in row number:
          ${index + 1}`);
      } else if (row[9] === null || row[9] === undefined) {
        throw new Error(`Block - ${
          BLOCK_SLA_REPORT_HEADERS[5]
        } is not available in SLA report in row number:
          ${index + 1}`);
      } else {
        if (!IP_ADDRESS_PATTERN.test(row[1].trim())) {
          throw new Error(
            ` Block - ${
              BLOCK_SLA_REPORT_HEADERS[1]
            } is not valid in SLA report in row number : ${index + 1}`
          );
        }
      }
    }
    this.storeDataAsObject(workSheetName, data);
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

  // Reading the worksheets individually and storing the data as Array of Objects
  storeDataAsObject(workSheetName: string, data: any): void {
    let result: any = [];
    data.forEach((data: any, index: number) => {
      if (workSheetName === 'block_sla_report' && index >= 2) {
        let obj: BlockNMSData = {
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
      } else if (workSheetName === 'block_noc_tt_report' && index >= 1) {
        let obj: BlockTTData = {
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
      } else if (workSheetName === 'block_alert_report' && index >= 1) {
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

    if (workSheetName === 'block_sla_report') {
      this.blockNMSData = result;
    } else if (workSheetName === 'block_noc_tt_report') {
      this.blockTTData = result;
    } else if (workSheetName === 'block_alert_report') {
      this.blockAlertData = result;
    }
  }

  // Alert report
  calculateAlertDownTimeInMinutes(ip: string) {
    let filteredAlertData = this.blockAlertData.filter(
      (alert: BlockAlertData) => {
        return (
          alert.ip_address == ip &&
          alert.severity == SEVERITY_CRITICAL &&
          alert.message == ALERT_DOWN_MESSAGE
        );
      }
    );

    let alertDownTimeInMinutes: number = 0;
    filteredAlertData.forEach((filteredAlertData: BlockAlertData) => {
      alertDownTimeInMinutes += filteredAlertData.total_duration_in_minutes;
    });
    return alertDownTimeInMinutes;
  }

  // Alert report and TT report
  categorizeRFO(nmsData: BlockNMSData) {
    if (nmsData.up_percent !== 100) {
      let totalPowerDownTimeInMinutes = 0;
      let totalDCNDownTimeInMinutes = 0;
      let isAlertReportEmpty: boolean = false;

      let powerDownArray: BlockAlertData[] = [];
      let DCNDownArray: BlockAlertData[] = [];
      let criticalAlertAndTTDataTimeMismatch: BlockAlertData[] = [];

      let powerIssueTT: string[] = [];
      let linkIssueTT: string[] = [];
      let otherTT: string[] = [];

      const filteredCriticalAlertData = this.blockAlertData.filter(
        (alertData: BlockAlertData) => {
          return (
            alertData.ip_address == nmsData.ip_address &&
            alertData.severity == SEVERITY_CRITICAL &&
            alertData.message == ALERT_DOWN_MESSAGE
          );
        }
      );

      const filteredWarningAlertData = this.blockAlertData.filter(
        (alertData: BlockAlertData) => {
          return (
            alertData.ip_address == nmsData.ip_address &&
            alertData.severity == SEVERITY_WARNING &&
            alertData.message.includes('reboot')
          );
        }
      );

      const filteredTTData = this.blockTTData.filter((ttData: BlockTTData) => {
        return ttData.ip == nmsData.ip_address;
      });

      if (filteredCriticalAlertData.length) {
        filteredCriticalAlertData.forEach(
          (alertCriticalData: BlockAlertData) => {
            filteredTTData.forEach((ttData: BlockTTData) => {
              if (
                moment(alertCriticalData.alarm_start_time).isSame(
                  ttData.incident_start_on,
                  'minute'
                )
              ) {
                if (ttData.rfo == RFO_CATEGORIZATION.POWER_ISSUE) {
                  if (
                    !lodash.some(powerDownArray, alertCriticalData) &&
                    !lodash.some(DCNDownArray, alertCriticalData)
                  ) {
                    powerDownArray.push(alertCriticalData);
                    powerIssueTT.push(ttData.incident_id);
                  }
                } else if (
                  ttData.rfo == RFO_CATEGORIZATION.JIO_LINK_ISSUE ||
                  ttData.rfo == RFO_CATEGORIZATION.SWAN_ISSUE
                ) {
                  if (
                    !lodash.some(powerDownArray, alertCriticalData) &&
                    !lodash.some(DCNDownArray, alertCriticalData)
                  ) {
                    DCNDownArray.push(alertCriticalData);
                    linkIssueTT.push(ttData.incident_id);
                  }
                } else {
                  otherTT.push(ttData.incident_id);
                }
              }
            });

            if (
              !lodash.some(powerDownArray, alertCriticalData) &&
              !lodash.some(DCNDownArray, alertCriticalData)
            ) {
              criticalAlertAndTTDataTimeMismatch.push(alertCriticalData);
            }
          }
        );
      } else {
        isAlertReportEmpty = true;
      }

      this.ttCorelation.push({
        ip: nmsData.ip_address,
        powerIssueTT: powerIssueTT,
        linkIssueTT: linkIssueTT,
        otherTT: otherTT,
      });

      if (criticalAlertAndTTDataTimeMismatch) {
        criticalAlertAndTTDataTimeMismatch.forEach(
          (alertCriticalData: BlockAlertData) => {
            filteredWarningAlertData.forEach(
              (alertWarningData: BlockAlertData) => {
                if (
                  moment(alertCriticalData.alarm_clear_time).isSame(
                    alertWarningData.alarm_start_time,
                    'minute'
                  ) &&
                  !lodash.some(powerDownArray, alertCriticalData) &&
                  !lodash.some(DCNDownArray, alertCriticalData)
                ) {
                  powerDownArray.push(alertCriticalData);
                }
              }
            );

            if (
              !lodash.some(powerDownArray, alertCriticalData) &&
              !lodash.some(DCNDownArray, alertCriticalData)
            ) {
              DCNDownArray.push(alertCriticalData);
            }
          }
        );
      }

      powerDownArray.forEach((powerDownAlert: BlockAlertData) => {
        totalPowerDownTimeInMinutes += powerDownAlert.total_duration_in_minutes;
      });

      DCNDownArray.forEach((dcnDownAlert: BlockAlertData) => {
        totalDCNDownTimeInMinutes += dcnDownAlert.total_duration_in_minutes;
      });

      const rfoCategorizedTimeInMinutes: RFOCategorizedTimeInMinutes = {
        total_dcn_downtime_minutes: +totalDCNDownTimeInMinutes.toFixed(2),
        total_power_downtime_minutes: +totalPowerDownTimeInMinutes.toFixed(2),
        alert_report_empty: isAlertReportEmpty,
      };
      return rfoCategorizedTimeInMinutes;
    } else {
      const rfoCategorizedTimeInMinutes: RFOCategorizedTimeInMinutes = {
        total_dcn_downtime_minutes: 0,
        total_power_downtime_minutes: 0,
        alert_report_empty: true,
      };
      return rfoCategorizedTimeInMinutes;
    }
  }

  manipulateBlockNMSData(): void {
    let manipulatedBlockNMSData: ManipulatedNMSData[] = [];
    this.blockNMSData.forEach((nmsData: BlockNMSData) => {
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
      let alertDownTimeInMinutes = this.calculateAlertDownTimeInMinutes(
        nmsData.ip_address
      );
      let alertDownTimeInPercent = +(
        (alertDownTimeInMinutes / totalTimeExclusiveOfSLAExclusionInMinutes) *
        100
      ).toFixed(2);
      let rfoCategorizedData = this.categorizeRFO(nmsData);
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

      let newNMSData: ManipulatedNMSData = {
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
      manipulatedBlockNMSData.push(newNMSData);
    });
    this.manipulatedNMSData = manipulatedBlockNMSData;

    this.blockSLASummary = this.blockService.calculateBlockSlaSummary(
      this.manipulatedNMSData,
      this.timeSpanValue
    );

    let blockNMSDataWithoutAlerts = this.manipulatedNMSData.filter(
      (blockNmsData: ManipulatedNMSData) =>
        blockNmsData.down_percent == 100 &&
        blockNmsData.alert_downtime_in_minutes == 0 &&
        blockNmsData.unknown_downtime_in_percent == 100
    );

    let blockNMSDataWithAlerts = this.manipulatedNMSData.filter(
      (blockNmsData: ManipulatedNMSData) =>
        !lodash.some(blockNMSDataWithoutAlerts, blockNmsData)
    );

    this.blockSLASummaryWithoutAlerts =
      this.blockService.calculateBlockSlaSummary(
        blockNMSDataWithoutAlerts,
        this.timeSpanValue
      );

    this.blockSLASummaryWithAlerts = this.blockService.calculateBlockSlaSummary(
      blockNMSDataWithAlerts,
      this.timeSpanValue
    );

    this.generateFinalBlockReport();
  }

  // Generating the final report as excel-workbook using the calculated data.

  generateFinalBlockReport() {
    const workbook = new ExcelJS.Workbook();
    const worksheet = workbook.addWorksheet('Block-SLA-Exclusion-Report');
    this.blockService.generateFinalBlockReport(
      workbook,
      worksheet,
      this.blockSLASummary,
      this.blockSLASummaryWithAlerts,
      this.blockSLASummaryWithoutAlerts,
      this.manipulatedNMSData,
      this.ttCorelation
    );
    workbook.xlsx.writeBuffer().then((buffer) => {
      this.sharedService.downloadFinalReport(
        buffer,
        'Block-SLA-Exclusion-Report'
      );
      this.isLoading = false;
      this.isBlockLoading.emit(false);
      this.resetInputFile();
    });
  }
}
