import { Component, ViewChild } from '@angular/core';
import * as ExcelJS from 'exceljs';
import * as moment from 'moment';
import * as lodash from 'lodash';
import {
  AOA,
  BlockAlertData,
  BlockNMSData,
  BlockSLASummaryPercent,
  BlockTTData,
  ManipulatedNMSData,
  RFOCategorizedTimeInMinutes,
} from './block-component.model';
import {
  SEVERITY_CRITICAL,
  ALERT_DOWN_MESSAGE,
  SEVERITY_WARNING,
  RFO_CATEGORIZATION,
  BLOCK_SLA_FINAL_REPORT_COLUMNS,
  SHEET_HEADING,
  TABLE_HEADING,
  BlockSLASummaryPercentHeaders,
  BORDER_STYLE,
  BlockSLAFinalReportHeaders,
  TABLE_HEADERS,
  PERCENT_STYLE,
  MINUTE_STYLE,
  VALUES,
  BLOCK_DEVICE_DETAILS,
  BlockDeviceDetail,
  BLOCK_SLA_REPORT_HEADERS,
  TT_REPORT_HEADERS,
  BLOCK_ALERT_REPORT_HEADERS,
  BLOCK_INPUT_FILE_NAMES,
} from '../constants/constants';
import { ToastrService } from 'ngx-toastr';

@Component({
  selector: 'app-block-component',
  templateUrl: './block-component.component.html',
  styleUrls: ['./block-component.component.scss'],
})
export class BlockComponentComponent {
  blockNMSData: any = [];
  blockTTData: any = [];
  blockAlertData: any = [];
  manipulatedNMSData: any = [];
  blockSLASummaryPercent!: BlockSLASummaryPercent;
  worksheet!: ExcelJS.Worksheet;
  file!: any;
  isSheetNamesValid: boolean = true;
  @ViewChild('blockInput') blockInput: any;

  constructor(private toastrService: ToastrService) {}

  // Getting the input file (excel workbook containing the required sheets)
  onFileChange(event: any): void {
    this.file = event.target.files[0];
    const workbook = new ExcelJS.Workbook();
    const reader = new FileReader();

    reader.onload = (e: any) => {
      const buffer = e.target.result;

      workbook.xlsx.load(buffer).then(() => {
        workbook.worksheets.forEach((_, index) => {
          this.worksheet = workbook.getWorksheet(index + 1);
          this.validateWorksheets(this.worksheet);
        });
        if (
          this.blockNMSData.length > 0 &&
          this.blockTTData.length > 0 &&
          this.blockAlertData.length > 0
        ) {
          this.manipulateBlockNMSData();
        }
      });
    };

    reader.readAsArrayBuffer(this.file);
  }

  resetInputFile(): void {
    this.file = null;
    const fileInput = document.getElementById('fileInput') as HTMLInputElement;
    fileInput.value = '';
    this.blockAlertData = [];
    this.blockNMSData = [];
    this.blockTTData = [];
  }

  validateWorksheets(worksheet: ExcelJS.Worksheet) {
    let workSheetName = worksheet.name;
    if (!BLOCK_INPUT_FILE_NAMES.includes(workSheetName)) {
      this.toastrService.error(
        'Sheet name is incorrect. Please provide valid sheet names'
      );
      this.resetInputFile();
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

      if (workSheetName === 'Block-SLA-Report') {
        if (headers !== JSON.stringify(BLOCK_SLA_REPORT_HEADERS)) {
          this.toastrService.error('Block SLA report format is incorrect');
          this.resetInputFile();
        } else {
          this.readWorksheet(worksheet, data);
        }
      } else if (workSheetName === 'Block-NOC TT Report') {
        if (headers !== JSON.stringify(TT_REPORT_HEADERS)) {
          this.toastrService.error('TT report format is incorrect');
          this.resetInputFile();
        } else {
          this.readWorksheet(worksheet, data);
        }
      } else if (workSheetName === 'Block-Alert Report') {
        if (headers !== JSON.stringify(BLOCK_ALERT_REPORT_HEADERS)) {
          this.toastrService.error('Block alert report format is incorrect');
          this.resetInputFile();
        } else {
          this.readWorksheet(worksheet, data);
        }
      }
    }
  }

  // Reading the worksheets individually and storing the data as Array of Objects
  readWorksheet(worksheet: ExcelJS.Worksheet, data: any): void {
    let workSheetName = worksheet.name;
    let result: any = [];
    data.shift();
    data.forEach((data: any, index: number) => {
      if (workSheetName === 'Block-SLA-Report') {
        let obj: BlockNMSData = {
          monitor: data[0],
          ip_address: data[1].trim(),
          departments: data[2],
          type: data[3],
          up_percent: data[4],
          up_time: data[5],
          down_percent: data[6],
          down_time: data[7],
          maintenance_percent: data[8],
          maintenance_time: data[9],
          total_up_percent: data[10],
          total_up_time: data[11],
          created_date: data[12],
        };
        result.push(obj);
      } else if (workSheetName === 'Block-NOC TT Report') {
        let obj: BlockTTData = {
          incident_id: data[0],
          parent_incident_id: data[1],
          enitity_type_name: data[2],
          entity_subtype_name: data[3],
          incident_name: data[4],
          equipment_host: data[5],
          ip: data[6].trim(),
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
          rfo: data[22].trim(),
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
      } else if (workSheetName === 'Block-Alert Report') {
        let obj: BlockAlertData = {
          alert: data[0],
          source: data[1],
          ip_address: data[2].trim(),
          departments: data[3],
          type: data[4],
          severity: data[5].trim(),
          message: data[6].trim(),
          alarm_start_time: moment(data[7]).format(),
          duration: data[8].trim(),
          alarm_clear_time: moment(data[9]).format(),
          total_duration_in_minutes: this.calculateTimeInMinutes(data[8]),
        };
        result.push(obj);
      }
    });

    if (workSheetName === 'Block-SLA-Report') {
      this.blockNMSData = result;
      // console.log('Block NMS Report', this.blockNMSData);
    } else if (workSheetName === 'Block-NOC TT Report') {
      this.blockTTData = result;
      // console.log('Block TT Report', this.blockTTData);
    } else if (workSheetName === 'Block-Alert Report') {
      this.blockAlertData = result;
      // console.log('Block Alert Report', this.blockAlertData);
    }
  }

  calculateTimeInMinutes(timePeriod: string): number {
    let totalTimeinMinutes = timePeriod.trim().split(' ');
    if (timePeriod.includes('days')) {
      return +(
        parseInt(totalTimeinMinutes[0]) * 1440 +
        parseInt(totalTimeinMinutes[2]) * 60 +
        parseInt(totalTimeinMinutes[4]) +
        parseInt(totalTimeinMinutes[6]) / 60
      ).toFixed(2);
    } else {
      return +(
        parseInt(totalTimeinMinutes[0]) * 60 +
        parseInt(totalTimeinMinutes[2]) +
        parseInt(totalTimeinMinutes[4]) / 60
      ).toFixed(2);
    }
  }

  // Alert report
  calculateAlertDownTimeInMinutes(ip: string) {
    let filteredAlertData = this.blockAlertData.filter(
      (alert: BlockAlertData) => {
        return (
          alert.ip_address.trim() == ip &&
          alert.severity.trim() == SEVERITY_CRITICAL &&
          alert.message.trim() == ALERT_DOWN_MESSAGE
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
  categorizeRFO(ip: string) {
    let totalPowerDownTimeInMinutes = 0;
    let totalDCNDownTimeInMinutes = 0;

    let powerDownArray: BlockAlertData[] = [];
    let DCNDownArray: BlockAlertData[] = [];
    let criticalAlertAndTTDataTimeMismatch: BlockAlertData[] = [];

    const filteredCriticalAlertData = this.blockAlertData.filter(
      (alertData: BlockAlertData) => {
        return (
          alertData.ip_address.trim() == ip &&
          alertData.severity.trim() == SEVERITY_CRITICAL &&
          alertData.message.trim() == ALERT_DOWN_MESSAGE
        );
      }
    );

    const filteredWarningAlertData = this.blockAlertData.filter(
      (alertData: BlockAlertData) => {
        return (
          alertData.ip_address.trim() == ip &&
          alertData.severity.trim() == SEVERITY_WARNING &&
          alertData.message.trim().includes('reboot')
        );
      }
    );

    const filteredTTData = this.blockTTData.filter((ttData: BlockTTData) => {
      return ttData.ip == ip;
    });

    filteredCriticalAlertData.forEach((alertCriticalData: BlockAlertData) => {
      filteredTTData.forEach((ttData: BlockTTData) => {
        if (
          moment(alertCriticalData.alarm_start_time).isSame(
            ttData.incident_start_on,
            'minute'
          )
        ) {
          if (ttData.rfo == RFO_CATEGORIZATION.POWER_ISSUE) {
            powerDownArray.push(alertCriticalData);
          } else if (
            ttData.rfo == RFO_CATEGORIZATION.JIO_LINK_ISSUE ||
            ttData.rfo == RFO_CATEGORIZATION.SWAN_ISSUE
          ) {
            DCNDownArray.push(alertCriticalData);
          }
        }
      });

      if (
        !lodash.some(powerDownArray, alertCriticalData) &&
        !lodash.some(DCNDownArray, alertCriticalData)
      ) {
        criticalAlertAndTTDataTimeMismatch.push(alertCriticalData);
      }
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
                )
              ) {
                powerDownArray.push(alertCriticalData);
              }
            }
          );

          if (!lodash.some(powerDownArray, alertCriticalData)) {
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
    };

    return rfoCategorizedTimeInMinutes;
  }

  manipulateBlockNMSData(): void {
    let manipulatedBlockNMSData: any = [];
    this.blockNMSData.forEach((nmsData: BlockNMSData) => {
      let totalUpTimeInMinutes = this.calculateTimeInMinutes(
        nmsData.total_up_time
      );
      let totalDownTimeInMinutes = this.calculateTimeInMinutes(
        nmsData.down_time
      );
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
      let rfoCategorizedData = this.categorizeRFO(nmsData.ip_address);
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

      let newNMSData = {
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
        power_downtime_in_percent: powerDownTimeInpercent,
        dcn_downtime_in_percent: dcnDownTimeInPercent,
      };
      manipulatedBlockNMSData.push(newNMSData);
    });
    this.manipulatedNMSData = manipulatedBlockNMSData;
    this.calcluateBlockSLASummary();
    this.generateFinalBlockReport();
  }

  calcluateBlockSLASummary() {
    let upPercent = 0;
    let upMinutes = 0;
    let powerDownPercent = 0;
    let powerDownMinutes = 0;
    let dcnDownPercent = 0;
    let dcnDownMinutes = 0;
    let plannedMaintenance = 0;
    let dcnAndPowerDownPercent = 0;
    let dcnAndPowerDownMinutes = 0;
    let totalDownPercent = 0;
    let totalDownMinutes = 0;
    let totalExclusionPercent = 0;
    let totalExclusionMinutes = 0;
    let pollingTimePercent = 0;
    let pollingTimeMinutes = 0;

    this.manipulatedNMSData.forEach((nmsData: ManipulatedNMSData) => {
      upPercent += nmsData.up_percent;
      powerDownPercent += nmsData.power_downtime_in_percent;
      powerDownMinutes += nmsData.power_downtime_in_minutes;
      dcnDownPercent += nmsData.dcn_downtime_in_percent;
      dcnDownMinutes += nmsData.dcn_downtime_in_minutes;
      dcnAndPowerDownPercent +=
        nmsData.power_downtime_in_percent + nmsData.dcn_downtime_in_percent;
      upMinutes += nmsData.total_uptime_in_minutes;
      dcnAndPowerDownMinutes +=
        nmsData.power_downtime_in_minutes + nmsData.dcn_downtime_in_minutes;
      totalDownMinutes += nmsData.total_downtime_in_minutes;
      totalExclusionPercent +=
        nmsData.power_downtime_in_percent + nmsData.dcn_downtime_in_percent;
      totalExclusionMinutes +=
        nmsData.power_downtime_in_minutes + nmsData.dcn_downtime_in_minutes;
      pollingTimePercent +=
        nmsData.down_percent -
        (nmsData.power_downtime_in_percent + nmsData.dcn_downtime_in_percent);
      pollingTimeMinutes +=
        nmsData.total_downtime_in_minutes -
        (nmsData.power_downtime_in_minutes + nmsData.dcn_downtime_in_minutes);
    });

    this.blockSLASummaryPercent = {
      report_type: 'BLOCK-SLA',
      time_span: '',
      no_of_blocks: 79,
      up_percent: +(upPercent / 79).toFixed(2),
      up_minutes: +(upMinutes / 79).toFixed(2),
      no_of_up_blocks: '',
      power_down_percent: +(powerDownPercent / 79).toFixed(2),
      power_down_minutes: +(powerDownMinutes / 79).toFixed(2),
      fibre_down_percent: 0.0,
      fibre_down_minutes: 0.0,
      equipment_down_percent: 0.0,
      equipment_down_minutes: 0.0,
      hrt_down_percent: 0.0,
      hrt_down_minutes: 0.0,
      dcn_down_percent: +(dcnDownPercent / 79).toFixed(2),
      dcn_down_minutes: +(dcnDownMinutes / 79).toFixed(2),
      planned_maintenance_percent: 0.0,
      planned_maintenance_minutes: 0.0,
      down_percent_exclusive_of_sla: 100 - upPercent / 79,
      no_of_down_blocks: '',
      total_sla_exclusion_percent: +(dcnAndPowerDownPercent / 79).toFixed(2),
      total_sla_exclusion_minutes: +(dcnAndPowerDownMinutes / 79).toFixed(2),
      total_up_percent: 0,
      total_up_minutes: 0,
      total_down_minutes: +(totalDownMinutes / 79).toFixed(2),
      total_down_percent: 100 - +(upPercent / 79).toFixed(2),
      total_up_percent_exclusion: +(
        (upPercent + pollingTimePercent + totalExclusionPercent) /
        79
      ).toFixed(2),
      total_up_minutes_exclusion: +(
        (upMinutes + pollingTimeMinutes + totalExclusionMinutes) /
        79
      ).toFixed(2),
    };
  }

  // Generating the final report as excel-workbook using the calculated data.
  generateFinalBlockReport(): void {
    const workbook = new ExcelJS.Workbook();
    const worksheet = workbook.addWorksheet('Block-Final-Report');
    worksheet.columns = BLOCK_SLA_FINAL_REPORT_COLUMNS;
    // worksheet.views = [{ state: 'frozen', xSplit: 10, ySplit: 0 }];

    worksheet.mergeCells('A1:B1');
    let cellA1 = worksheet.getCell('A1');
    cellA1.value = '1. Daily Network availability report';
    cellA1.style = SHEET_HEADING;

    worksheet.mergeCells('C1:D1');
    let cellC1 = worksheet.getCell('C1');
    cellC1.value = 'Report-Frequency- Daily';
    cellC1.style = {
      font: { bold: true },
      alignment: { horizontal: 'center' },
    };

    worksheet.mergeCells('A3:B3');
    let cellA3 = worksheet.getCell('A3');
    cellA3.value = 'Block - SLA Summary (%) & (Min)';
    cellA3.style = TABLE_HEADING;
    worksheet.getCell('B3').style = TABLE_HEADING;

    worksheet.mergeCells('C4:G4');
    worksheet.mergeCells('I4:J4');
    worksheet.mergeCells('K4:L4');
    worksheet.mergeCells('M4:N4');
    worksheet.mergeCells('O4:P4');
    worksheet.mergeCells('Q4:R4');
    worksheet.mergeCells('S4:T4');
    worksheet.mergeCells('U4:V4');
    worksheet.mergeCells('W4:X4');
    worksheet.mergeCells('Y4:Z4');
    worksheet.mergeCells('AA4:AB4');
    worksheet.mergeCells('AC4:AD4');

    worksheet.getCell('A4').value = BlockSLASummaryPercentHeaders[0];
    worksheet.getCell('B4').value = BlockSLASummaryPercentHeaders[1];
    worksheet.getCell('C4').value = BlockSLASummaryPercentHeaders[2];
    worksheet.getCell('H4').value = BlockSLASummaryPercentHeaders[3];
    worksheet.getCell('I4').value = BlockSLASummaryPercentHeaders[4];
    worksheet.getCell('K4').value = BlockSLASummaryPercentHeaders[5];
    worksheet.getCell('M4').value = BlockSLASummaryPercentHeaders[6];
    worksheet.getCell('O4').value = BlockSLASummaryPercentHeaders[7];
    worksheet.getCell('Q4').value = BlockSLASummaryPercentHeaders[8];
    worksheet.getCell('S4').value = BlockSLASummaryPercentHeaders[9];
    worksheet.getCell('U4').value = BlockSLASummaryPercentHeaders[10];
    worksheet.getCell('W4').value = BlockSLASummaryPercentHeaders[11];
    worksheet.getCell('Y4').value = BlockSLASummaryPercentHeaders[12];
    worksheet.getCell('AA4').value = BlockSLASummaryPercentHeaders[13];
    worksheet.getCell('AC4').value = BlockSLASummaryPercentHeaders[14];

    let blockSummaryHeadersRow = worksheet.getRow(4);
    blockSummaryHeadersRow.eachCell((cell) => {
      cell.style = TABLE_HEADERS;
    });

    worksheet.mergeCells('A5:A6');
    worksheet.mergeCells('B5:B6');
    worksheet.mergeCells('C5:G6');
    worksheet.mergeCells('H5:H6');
    worksheet.mergeCells('I5:J6');

    worksheet.getCell('A5').value = 'Block - SLA';
    worksheet.getCell('B5').value = 'Q&M Block';
    worksheet.getCell('H5').value = '79';
    worksheet.getCell('I5').value = '5001';
    let k5 = worksheet.getCell('K5');
    k5.value = VALUES.PERCENT;
    k5.style = PERCENT_STYLE;
    let l5 = worksheet.getCell('L5');
    l5.value = VALUES.MINUTES;
    l5.style = MINUTE_STYLE;
    let m5 = worksheet.getCell('M5');
    m5.value = VALUES.PERCENT;
    m5.style = PERCENT_STYLE;
    let n5 = worksheet.getCell('N5');
    n5.value = VALUES.MINUTES;
    n5.style = MINUTE_STYLE;
    let o5 = worksheet.getCell('O5');
    o5.value = VALUES.PERCENT;
    o5.style = PERCENT_STYLE;
    let p5 = worksheet.getCell('P5');
    p5.value = VALUES.MINUTES;
    p5.style = MINUTE_STYLE;
    let q5 = worksheet.getCell('Q5');
    q5.value = VALUES.PERCENT;
    q5.style = PERCENT_STYLE;
    let r5 = worksheet.getCell('R5');
    r5.value = VALUES.MINUTES;
    r5.style = MINUTE_STYLE;
    let s5 = worksheet.getCell('S5');
    s5.value = VALUES.PERCENT;
    s5.style = PERCENT_STYLE;
    let t5 = worksheet.getCell('T5');
    t5.value = VALUES.MINUTES;
    t5.style = MINUTE_STYLE;
    let u5 = worksheet.getCell('U5');
    u5.value = VALUES.PERCENT;
    u5.style = PERCENT_STYLE;
    let v5 = worksheet.getCell('V5');
    v5.value = VALUES.MINUTES;
    v5.style = MINUTE_STYLE;
    let w5 = worksheet.getCell('W5');
    w5.value = VALUES.PERCENT;
    w5.style = PERCENT_STYLE;
    let x5 = worksheet.getCell('X5');
    x5.value = VALUES.MINUTES;
    x5.style = MINUTE_STYLE;
    let y5 = worksheet.getCell('Y5');
    y5.value = VALUES.PERCENT;
    y5.style = PERCENT_STYLE;
    let z5 = worksheet.getCell('Z5');
    z5.value = VALUES.MINUTES;
    z5.style = MINUTE_STYLE;
    let aa5 = worksheet.getCell('AA5');
    aa5.value = VALUES.PERCENT;
    aa5.style = PERCENT_STYLE;
    let ab5 = worksheet.getCell('AB5');
    ab5.value = VALUES.MINUTES;
    ab5.style = MINUTE_STYLE;
    let ac5 = worksheet.getCell('AC5');
    ac5.value = VALUES.PERCENT;
    ac5.style = PERCENT_STYLE;
    let ad5 = worksheet.getCell('AD5');
    ad5.value = VALUES.MINUTES;
    ad5.style = MINUTE_STYLE;

    worksheet.getCell('K6').value = this.blockSLASummaryPercent.up_percent;
    worksheet.getCell('L6').value = this.blockSLASummaryPercent.up_minutes;
    worksheet.getCell('M6').value =
      this.blockSLASummaryPercent.total_down_percent;
    worksheet.getCell('N6').value =
      this.blockSLASummaryPercent.total_down_minutes;
    worksheet.getCell('O6').value =
      this.blockSLASummaryPercent.power_down_percent;
    worksheet.getCell('P6').value =
      this.blockSLASummaryPercent.power_down_minutes;
    worksheet.getCell('Q6').value =
      this.blockSLASummaryPercent.fibre_down_percent;
    worksheet.getCell('R6').value =
      this.blockSLASummaryPercent.fibre_down_minutes;
    worksheet.getCell('S6').value =
      this.blockSLASummaryPercent.equipment_down_percent;
    worksheet.getCell('T6').value =
      this.blockSLASummaryPercent.equipment_down_minutes;
    worksheet.getCell('U6').value =
      this.blockSLASummaryPercent.hrt_down_percent;
    worksheet.getCell('V6').value =
      this.blockSLASummaryPercent.hrt_down_minutes;
    worksheet.getCell('W6').value =
      this.blockSLASummaryPercent.dcn_down_percent;
    worksheet.getCell('X6').value =
      this.blockSLASummaryPercent.dcn_down_minutes;
    worksheet.getCell('Y6').value =
      this.blockSLASummaryPercent.planned_maintenance_percent;
    worksheet.getCell('Z6').value =
      this.blockSLASummaryPercent.planned_maintenance_minutes;
    worksheet.getCell('AA6').value =
      this.blockSLASummaryPercent.total_sla_exclusion_percent;
    worksheet.getCell('AB6').value =
      this.blockSLASummaryPercent.total_sla_exclusion_minutes;
    worksheet.getCell('AC6').value =
      this.blockSLASummaryPercent.total_up_percent_exclusion;
    worksheet.getCell('AD6').value =
      this.blockSLASummaryPercent.total_up_minutes_exclusion;

    let row5 = worksheet.getRow(5);
    row5.eachCell((cell) => {
      cell.border = BORDER_STYLE;
      cell.alignment = { horizontal: 'center' };
      cell.font = { bold: true };
    });

    let row6 = worksheet.getRow(6);
    row6.eachCell((cell) => {
      cell.border = BORDER_STYLE;
      cell.alignment = { horizontal: 'center' };
    });

    worksheet.addRow('');

    worksheet.mergeCells('A9:B9');
    let cellA11 = worksheet.getCell('A9');
    cellA11.value = 'Block - SLA Device Level (%) & (Min)';
    cellA11.style = TABLE_HEADING;
    worksheet.getCell('B9').style = TABLE_HEADERS;

    worksheet.mergeCells('A10:A11');
    worksheet.mergeCells('B10:B11');
    worksheet.mergeCells('C10:C11');
    worksheet.mergeCells('D10:D11');
    worksheet.mergeCells('E10:E11');
    worksheet.mergeCells('F10:F11');
    worksheet.mergeCells('G10:G11');
    worksheet.mergeCells('H10:H11');
    worksheet.mergeCells('I10:I11');
    worksheet.mergeCells('J10:J11');
    worksheet.mergeCells('K10:L10');
    worksheet.mergeCells('M10:N10');
    worksheet.mergeCells('O10:P10');
    worksheet.mergeCells('Q10:R10');
    worksheet.mergeCells('S10:T10');
    worksheet.mergeCells('U10:V10');
    worksheet.mergeCells('W10:X10');
    worksheet.mergeCells('Y10:Z10');
    worksheet.mergeCells('AA10:AB10');
    worksheet.mergeCells('AC10:AD10');
    worksheet.mergeCells('AE10:AF10');

    worksheet.getCell('A10').value = BlockSLAFinalReportHeaders[0];
    worksheet.getCell('B10').value = BlockSLAFinalReportHeaders[1];
    worksheet.getCell('C10').value = BlockSLAFinalReportHeaders[2];
    worksheet.getCell('D10').value = BlockSLAFinalReportHeaders[3];
    worksheet.getCell('E10').value = BlockSLAFinalReportHeaders[4];
    worksheet.getCell('F10').value = BlockSLAFinalReportHeaders[5];
    worksheet.getCell('G10').value = BlockSLAFinalReportHeaders[6];
    worksheet.getCell('H10').value = BlockSLAFinalReportHeaders[7];
    worksheet.getCell('I10').value = BlockSLAFinalReportHeaders[8];
    worksheet.getCell('J10').value = BlockSLAFinalReportHeaders[9];
    worksheet.getCell('K10').value = BlockSLAFinalReportHeaders[10];
    worksheet.getCell('M10').value = BlockSLAFinalReportHeaders[11];
    worksheet.getCell('O10').value = BlockSLAFinalReportHeaders[12];
    worksheet.getCell('Q10').value = BlockSLAFinalReportHeaders[13];
    worksheet.getCell('S10').value = BlockSLAFinalReportHeaders[14];
    worksheet.getCell('U10').value = BlockSLAFinalReportHeaders[15];
    worksheet.getCell('W10').value = BlockSLAFinalReportHeaders[16];
    worksheet.getCell('Y10').value = BlockSLAFinalReportHeaders[17];
    worksheet.getCell('AA10').value = BlockSLAFinalReportHeaders[18];
    worksheet.getCell('AC10').value = BlockSLAFinalReportHeaders[19];
    worksheet.getCell('AE10').value = BlockSLAFinalReportHeaders[20];

    let finalReportHeaders = worksheet.getRow(10);

    finalReportHeaders.eachCell((cell) => {
      cell.style = TABLE_HEADERS;
    });

    let k11 = worksheet.getCell('K11');
    k11.value = VALUES.PERCENT;
    k11.style = PERCENT_STYLE;
    let l11 = worksheet.getCell('L11');
    l11.value = VALUES.MINUTES;
    l11.style = MINUTE_STYLE;

    let m11 = worksheet.getCell('M11');
    m11.value = VALUES.PERCENT;
    m11.style = PERCENT_STYLE;
    let n11 = worksheet.getCell('N11');
    n11.value = VALUES.MINUTES;
    n11.style = MINUTE_STYLE;

    let o11 = worksheet.getCell('O11');
    o11.value = VALUES.PERCENT;
    o11.style = PERCENT_STYLE;
    let p11 = worksheet.getCell('P11');
    p11.value = VALUES.MINUTES;
    p11.style = MINUTE_STYLE;

    let q11 = worksheet.getCell('Q11');
    q11.value = VALUES.PERCENT;
    q11.style = PERCENT_STYLE;
    let r11 = worksheet.getCell('R11');
    r11.value = VALUES.MINUTES;
    r11.style = MINUTE_STYLE;

    let s11 = worksheet.getCell('S11');
    s11.value = VALUES.PERCENT;
    s11.style = PERCENT_STYLE;
    let t11 = worksheet.getCell('T11');
    t11.value = VALUES.MINUTES;
    t11.style = MINUTE_STYLE;

    let u11 = worksheet.getCell('U11');
    u11.value = VALUES.PERCENT;
    u11.style = PERCENT_STYLE;
    let v11 = worksheet.getCell('V11');
    v11.value = VALUES.MINUTES;
    v11.style = MINUTE_STYLE;

    let w11 = worksheet.getCell('W11');
    w11.value = VALUES.PERCENT;
    w11.style = PERCENT_STYLE;
    let x11 = worksheet.getCell('X11');
    x11.value = VALUES.MINUTES;
    x11.style = MINUTE_STYLE;

    let y11 = worksheet.getCell('Y11');
    y11.value = VALUES.PERCENT;
    y11.style = PERCENT_STYLE;
    let z11 = worksheet.getCell('Z11');
    z11.value = VALUES.MINUTES;
    z11.style = MINUTE_STYLE;

    let aa11 = worksheet.getCell('AA11');
    aa11.value = VALUES.PERCENT;
    aa11.style = PERCENT_STYLE;
    let ab11 = worksheet.getCell('AB11');
    ab11.value = VALUES.MINUTES;
    ab11.style = MINUTE_STYLE;

    let ac11 = worksheet.getCell('AC11');
    ac11.value = VALUES.PERCENT;
    ac11.style = PERCENT_STYLE;
    let ad11 = worksheet.getCell('AD11');
    ad11.value = VALUES.MINUTES;
    ad11.style = MINUTE_STYLE;

    let ae11 = worksheet.getCell('AE11');
    ae11.value = VALUES.PERCENT;
    ae11.style = PERCENT_STYLE;
    let af11 = worksheet.getCell('AF11');
    af11.value = VALUES.MINUTES;
    af11.style = MINUTE_STYLE;

    let row11 = worksheet.getRow(11);
    row11.eachCell((cell) => {
      cell.border = BORDER_STYLE;
      cell.font = { bold: true };
      cell.alignment = { horizontal: 'center' };
    });

    this.manipulatedNMSData.forEach((row: any) => {
      let block_device_detail = BLOCK_DEVICE_DETAILS.filter(
        (device: BlockDeviceDetail) => {
          return device.ip_address == row.ip_address;
        }
      );
      let [blockDeviceDetail] = block_device_detail;
      let reportType: string = blockDeviceDetail.report_type;
      let hostName: string = blockDeviceDetail.host_name;
      let ipAddress: string = blockDeviceDetail.ip_address;
      let state: string = blockDeviceDetail.state;
      let cluster: string = blockDeviceDetail.cluster;
      let district: string = blockDeviceDetail.district;
      let districtLGDCode: number = blockDeviceDetail.district_lgd_code;
      let blockName: string = blockDeviceDetail.block_name;
      let blockLGDCode: string = blockDeviceDetail.block_lgd_code;
      let noOfGPinBlock: number = blockDeviceDetail.no_of_gp_in_block;
      let upPercent: number = row.up_percent;
      let upMinute: number = row.total_uptime_in_minutes;
      let downPercent: number = upPercent == 100 ? 0 : row.down_percent;
      let downMinute: number = row.total_downtime_in_minutes;
      let powerDownPercent: number =
        upPercent == 100 ? 0 : row.power_downtime_in_percent;
      let powerDownMinutes = row.power_downtime_in_minutes;
      let fiberDownPercent: number = upPercent == 100 ? 0 : 0;
      let fiberDownMinutes: number = upPercent == 100 ? 0 : 0;
      let equipmentDownPercent: number = upPercent == 100 ? 0 : 0;
      let equipmentDownMinutes: number = upPercent == 100 ? 0 : 0;
      let hrtDownPercent: number = upPercent == 100 ? 0 : 0;
      let hrtDownMinutes: number = upPercent == 100 ? 0 : 0;
      let dcnDownPercent: number =
        upPercent == 100 ? 0 : row.dcn_downtime_in_percent;
      let dcnDownMinutes: number = row.dcn_downtime_in_minutes;
      let plannedMaintanancePercent: number = upPercent == 100 ? 0 : 0;
      let plannedMaintananceMinutes: number = upPercent == 100 ? 0 : 0;
      let totalExclusionPercent: number =
        upPercent == 100
          ? 0
          : row.power_downtime_in_percent + row.dcn_downtime_in_percent;
      let totalExclusionMinutes: number =
        row.power_downtime_in_minutes + row.dcn_downtime_in_minutes;
      let pollingTimePercent: number =
        upPercent == 100 ? 0 : row.down_percent - +totalExclusionPercent;
      let pollingTimeMinutes: number =
        row.total_downtime_in_minutes - totalExclusionMinutes;

      let totalUpPercentSLAExclusion: number =
        upPercent + totalExclusionPercent + pollingTimePercent;
      let totalUpMinutesSLAExclusion: number =
        upMinute + totalExclusionMinutes + pollingTimeMinutes;

      const blockSummaryPercentRowValues = worksheet.addRow([
        reportType,
        hostName,
        ipAddress,
        state,
        cluster,
        district,
        districtLGDCode,
        blockName,
        blockLGDCode,
        noOfGPinBlock,
        upPercent.toFixed(2),
        upMinute.toFixed(2),
        downPercent.toFixed(2),
        downMinute.toFixed(2),
        powerDownPercent.toFixed(2),
        powerDownMinutes.toFixed(2),
        fiberDownPercent.toFixed(2),
        fiberDownMinutes.toFixed(2),
        equipmentDownPercent.toFixed(2),
        equipmentDownMinutes.toFixed(2),
        hrtDownPercent.toFixed(2),
        hrtDownMinutes.toFixed(2),
        dcnDownPercent.toFixed(2),
        dcnDownMinutes.toFixed(2),
        plannedMaintanancePercent.toFixed(2),
        plannedMaintananceMinutes.toFixed(2),
        totalExclusionPercent.toFixed(2),
        totalExclusionMinutes.toFixed(2),
        pollingTimePercent.toFixed(2),
        pollingTimeMinutes.toFixed(2),
        totalUpPercentSLAExclusion.toFixed(2),
        totalUpMinutesSLAExclusion.toFixed(2),
      ]);

      blockSummaryPercentRowValues.eachCell((cell) => {
        cell.border = BORDER_STYLE;
        cell.alignment = { horizontal: 'left' };
      });
    });

    workbook.xlsx.writeBuffer().then((buffer) => {
      this.downloadFinalReport(buffer, 'sample');
    });
  }

  // Downloading the generated final excel workbook
  downloadFinalReport(buffer: ArrayBuffer, fileName: string): void {
    const data = new Blob([buffer], {
      type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
    });
    const link = document.createElement('a');
    link.href = window.URL.createObjectURL(data);
    link.download = fileName + '.xlsx';
    link.click();
    this.resetInputFile();
  }
}
