import { Injectable } from '@angular/core';
import {
  GpAlertData,
  GpDeviceDetails,
  GpNMSData,
  GpSLASummary,
  GpTTData,
  ManipulatedGpNMSData,
} from './gp.model';
import {
  ALERT_DOWN_MESSAGE,
  BORDER_STYLE,
  GP_DEVICE_DETAILS,
  GP_FINAL_REPORT_DEVICE_LEVEL_HEADERS,
  GP_SLA_FINAL_REPORT_COLUMN_WIDTHS,
  GP_SUMMARY_HEADERS,
  MINUTE_STYLE,
  PERCENT_STYLE,
  RFO_CATEGORIZATION,
  SEVERITY_CRITICAL,
  SEVERITY_WARNING,
  SHEET_HEADING,
  TABLE_HEADERS,
  TABLE_HEADING,
  VALUES,
} from '../constants/constants';
import * as moment from 'moment';
import * as lodash from 'lodash';
import * as ExcelJS from 'exceljs';

import { RFOCategorizedTimeInMinutes } from '../block-component/block-component.model';

@Injectable({
  providedIn: 'root',
})
export class GpService {
  constructor() {}

  calculateAlertDownTimeInMinutes(
    ipAddress: string,
    gpAlertData: GpAlertData[]
  ) {
    let filteredAlertData = gpAlertData.filter((alert: GpAlertData) => {
      return (
        alert.ip_address == ipAddress &&
        alert.severity == SEVERITY_CRITICAL &&
        alert.message == ALERT_DOWN_MESSAGE
      );
    });

    let alertDownTimeInMinutes: number = 0;
    filteredAlertData.forEach((filteredAlertData: GpAlertData) => {
      alertDownTimeInMinutes += filteredAlertData.total_duration_in_minutes;
    });
    return alertDownTimeInMinutes;
  }

  categorizeRFO(
    nmsData: GpNMSData,
    gpAlertData: GpAlertData[],
    gpTTData: GpTTData[]
  ) {
    if (nmsData.up_percent !== 100) {
      let totalPowerDownTimeInMinutes = 0;
      let totalDCNDownTimeInMinutes = 0;
      let isAlertReportEmpty: boolean = false;

      let powerDownArray: GpAlertData[] = [];
      let DCNDownArray: GpAlertData[] = [];
      let criticalAlertAndTTDataTimeMismatch: GpAlertData[] = [];

      const filteredCriticalAlertData = gpAlertData.filter(
        (alertData: GpAlertData) => {
          return (
            alertData.ip_address == nmsData.ip_address &&
            alertData.severity == SEVERITY_CRITICAL &&
            alertData.message == ALERT_DOWN_MESSAGE
          );
        }
      );

      const filteredWarningAlertData = gpAlertData.filter(
        (alertData: GpAlertData) => {
          return (
            alertData.ip_address == nmsData.ip_address &&
            alertData.severity == SEVERITY_WARNING &&
            alertData.message.includes('reboot')
          );
        }
      );

      const filteredTTData = gpTTData.filter((ttData: GpTTData) => {
        return ttData.ip == nmsData.ip_address;
      });

      if (filteredCriticalAlertData.length) {
        filteredCriticalAlertData.forEach((alertCriticalData: GpAlertData) => {
          filteredTTData.forEach((ttData: GpTTData) => {
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
      } else {
        isAlertReportEmpty = true;
      }

      if (criticalAlertAndTTDataTimeMismatch) {
        criticalAlertAndTTDataTimeMismatch.forEach(
          (alertCriticalData: GpAlertData) => {
            filteredWarningAlertData.forEach(
              (alertWarningData: GpAlertData) => {
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

      powerDownArray.forEach((powerDownAlert: GpAlertData) => {
        totalPowerDownTimeInMinutes += powerDownAlert.total_duration_in_minutes;
      });

      DCNDownArray.forEach((dcnDownAlert: GpAlertData) => {
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

  calculateGpSlaSummary(
    manipulatedGpNMSData: ManipulatedGpNMSData[]
  ): GpSLASummary {
    let upPercent = 0;
    let upMinutes = 0;
    let powerDownPercent = 0;
    let powerDownMinutes = 0;
    let fiberDownPercent = 0;
    let fiberDownMinute = 0;
    let equipmentDownPercent = 0;
    let equipmentDownMinute = 0;
    let hrtDownPercent = 0;
    let hrtDownMinute = 0;
    let dcnDownPercent = 0;
    let dcnDownMinutes = 0;
    let plannedMaintenancePercent = 0;
    let plannedMaintenanceMinutes = 0;
    let unKnownDownMinutes = 0;
    let unKnownDownPercent = 0;
    let cumulativeRfoDownInPercent = 0;
    let cumulativeRfoDownInMinutes = 0;
    let totalDownMinutes = 0;
    let totalExclusionPercent = 0;
    let totalExclusionMinutes = 0;
    let pollingTimePercent = 0;
    let pollingTimeMinutes = 0;

    manipulatedGpNMSData.forEach((nmsData: ManipulatedGpNMSData) => {
      upPercent += nmsData.up_percent;
      powerDownPercent += nmsData.power_downtime_in_percent;
      powerDownMinutes += nmsData.power_downtime_in_minutes;
      dcnDownPercent += nmsData.dcn_downtime_in_percent;
      dcnDownMinutes += nmsData.dcn_downtime_in_minutes;
      plannedMaintenancePercent += nmsData.maintenance_percent;
      plannedMaintenanceMinutes += nmsData.planned_maintenance_in_minutes;
      unKnownDownMinutes += nmsData.unknown_downtime_in_minutes;
      unKnownDownPercent += nmsData.unknown_downtime_in_percent;
      cumulativeRfoDownInPercent +=
        nmsData.power_downtime_in_percent +
        nmsData.dcn_downtime_in_percent +
        nmsData.maintenance_percent +
        nmsData.unknown_downtime_in_percent;
      upMinutes += nmsData.total_uptime_in_minutes;
      cumulativeRfoDownInMinutes +=
        nmsData.power_downtime_in_minutes +
        nmsData.dcn_downtime_in_minutes +
        nmsData.planned_maintenance_in_minutes +
        nmsData.unknown_downtime_in_minutes;
      totalDownMinutes += nmsData.total_downtime_in_minutes;
      totalExclusionPercent +=
        nmsData.power_downtime_in_percent +
        nmsData.dcn_downtime_in_percent +
        nmsData.planned_maintenance_in_percent +
        nmsData.unknown_downtime_in_percent;
      totalExclusionMinutes +=
        nmsData.power_downtime_in_minutes +
        nmsData.dcn_downtime_in_minutes +
        nmsData.planned_maintenance_in_minutes +
        nmsData.unknown_downtime_in_minutes;
      pollingTimePercent +=
        nmsData.down_percent -
        (nmsData.power_downtime_in_percent +
          nmsData.dcn_downtime_in_percent +
          nmsData.unknown_downtime_in_percent);
      pollingTimeMinutes +=
        nmsData.total_downtime_in_minutes -
        (nmsData.power_downtime_in_minutes +
          nmsData.dcn_downtime_in_minutes +
          nmsData.unknown_downtime_in_minutes);
    });

    return {
      report_type: 'BLOCK-SLA',
      time_span: '',
      no_of_blocks: 79,
      up_percent: (upPercent / 79).toFixed(2),
      up_minutes: upMinutes.toFixed(2),
      no_of_up_blocks: '',
      down_percent_exclusive_of_sla: (100 - upPercent / 79).toFixed(2),
      power_down_percent: (powerDownPercent / 79).toFixed(2),
      power_down_minutes: powerDownMinutes.toFixed(2),
      fibre_down_percent: (fiberDownPercent / 79).toFixed(2),
      fibre_down_minutes: fiberDownMinute.toFixed(2),
      equipment_down_percent: (equipmentDownPercent / 79).toFixed(2),
      equipment_down_minutes: equipmentDownMinute.toFixed(2),
      hrt_down_percent: (hrtDownPercent / 79).toFixed(2),
      hrt_down_minutes: hrtDownMinute.toFixed(2),
      dcn_down_percent: (dcnDownPercent / 79).toFixed(2),
      dcn_down_minutes: dcnDownMinutes.toFixed(2),
      planned_maintenance_percent: (plannedMaintenancePercent / 79).toFixed(2),
      planned_maintenance_minutes: plannedMaintenanceMinutes.toFixed(2),
      unknown_downtime_in_percent: (unKnownDownPercent / 79).toFixed(2),
      unknown_downtime_in_minutes: unKnownDownMinutes.toFixed(2),
      total_sla_exclusion_percent: (cumulativeRfoDownInPercent / 79).toFixed(2),
      total_sla_exclusion_minutes: cumulativeRfoDownInMinutes.toFixed(2),
      total_down_minutes: totalDownMinutes.toFixed(2),
      total_down_percent: (100 - +(upPercent / 79)).toFixed(2),
      total_up_percent_exclusion: (
        (upPercent + pollingTimePercent + totalExclusionPercent) /
        79
      ).toFixed(2),

      total_up_minutes_exclusion: (
        upMinutes +
        pollingTimeMinutes +
        totalExclusionMinutes
      ).toFixed(2),
    };
  }

  FrameGpFinalSlaReportWorkbook(
    workSheet: ExcelJS.Worksheet,
    gpSlaSummary: GpSLASummary,
    manipulatedGpNmsData: ManipulatedGpNMSData[]
  ): void {
    const workbook = new ExcelJS.Workbook();
    const worksheet = workbook.addWorksheet('GP-SLA-Exclusion-Report');
    worksheet.columns = GP_SLA_FINAL_REPORT_COLUMN_WIDTHS;

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
    cellA3.value = 'SHQ - SLA Summary (%) & (Min)';
    cellA3.style = TABLE_HEADING;
    worksheet.getCell('B3').style = TABLE_HEADING;

    worksheet.mergeCells('C4:J4');
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
    worksheet.mergeCells('AE4:AF4');
    worksheet.mergeCells('AG4:AH4');

    workSheet.getCell('A4').value = GP_SUMMARY_HEADERS[0];
    workSheet.getCell('B4').value = GP_SUMMARY_HEADERS[1];
    workSheet.getCell('C4').value = GP_SUMMARY_HEADERS[2];
    workSheet.getCell('K4').value = GP_SUMMARY_HEADERS[3];
    workSheet.getCell('M4').value = GP_SUMMARY_HEADERS[4];
    workSheet.getCell('O4').value = GP_SUMMARY_HEADERS[5];
    workSheet.getCell('Q4').value = GP_SUMMARY_HEADERS[6];
    workSheet.getCell('S4').value = GP_SUMMARY_HEADERS[7];
    workSheet.getCell('U4').value = GP_SUMMARY_HEADERS[8];
    workSheet.getCell('W4').value = GP_SUMMARY_HEADERS[9];
    workSheet.getCell('Y4').value = GP_SUMMARY_HEADERS[10];
    workSheet.getCell('Y4').value = GP_SUMMARY_HEADERS[10];
    workSheet.getCell('AA4').value = GP_SUMMARY_HEADERS[11];
    workSheet.getCell('AC4').value = GP_SUMMARY_HEADERS[12];
    workSheet.getCell('AE4').value = GP_SUMMARY_HEADERS[13];
    workSheet.getCell('AG4').value = GP_SUMMARY_HEADERS[14];

    let gpSummaryHeadersRow = worksheet.getRow(4);
    gpSummaryHeadersRow.eachCell((cell) => {
      cell.style = TABLE_HEADERS;
    });

    workSheet.mergeCells('A5:A6');
    workSheet.mergeCells('B5:B6');
    workSheet.mergeCells('C5:J6');
    workSheet.mergeCells('K5:L6');

    workSheet.getCell('A5').value = 'GP - SLA';
    workSheet.getCell('B5').value = 'O&M GP';
    workSheet.getCell('C5').value = '';
    workSheet.getCell('K5').value = '5001';

    let M5 = worksheet.getCell('M5');
    M5.value = VALUES.PERCENT;
    M5.style = PERCENT_STYLE;
    let N5 = worksheet.getCell('N5');
    N5.value = VALUES.MINUTES;
    N5.style = MINUTE_STYLE;

    let O5 = worksheet.getCell('O5');
    O5.value = VALUES.PERCENT;
    O5.style = PERCENT_STYLE;
    let P5 = worksheet.getCell('P5');
    P5.value = VALUES.MINUTES;
    P5.style = MINUTE_STYLE;

    let Q5 = worksheet.getCell('Q5');
    Q5.value = VALUES.PERCENT;
    Q5.style = PERCENT_STYLE;
    let R5 = worksheet.getCell('R5');
    R5.value = VALUES.MINUTES;
    R5.style = MINUTE_STYLE;

    let S5 = worksheet.getCell('S5');
    S5.value = VALUES.PERCENT;
    S5.style = PERCENT_STYLE;
    let T5 = worksheet.getCell('T5');
    T5.value = VALUES.MINUTES;
    T5.style = MINUTE_STYLE;

    let U5 = worksheet.getCell('U5');
    U5.value = VALUES.PERCENT;
    U5.style = PERCENT_STYLE;
    let V5 = worksheet.getCell('V5');
    V5.value = VALUES.MINUTES;
    V5.style = MINUTE_STYLE;

    let W5 = worksheet.getCell('W5');
    W5.value = VALUES.PERCENT;
    W5.style = PERCENT_STYLE;
    let X5 = worksheet.getCell('X5');
    X5.value = VALUES.MINUTES;
    X5.style = MINUTE_STYLE;

    let Y5 = worksheet.getCell('Y5');
    Y5.value = VALUES.PERCENT;
    Y5.style = PERCENT_STYLE;
    let Z5 = worksheet.getCell('Z5');
    Z5.value = VALUES.MINUTES;
    Z5.style = MINUTE_STYLE;

    let AA5 = worksheet.getCell('AA5');
    AA5.value = VALUES.PERCENT;
    AA5.style = PERCENT_STYLE;
    let AB5 = worksheet.getCell('AB5');
    AB5.value = VALUES.MINUTES;
    AB5.style = MINUTE_STYLE;

    let AC5 = worksheet.getCell('AC5');
    AC5.value = VALUES.PERCENT;
    AC5.style = PERCENT_STYLE;
    let AD5 = worksheet.getCell('AD5');
    AD5.value = VALUES.MINUTES;
    AD5.style = MINUTE_STYLE;

    let AE5 = worksheet.getCell('AE5');
    AE5.value = VALUES.PERCENT;
    AE5.style = PERCENT_STYLE;
    let AF5 = worksheet.getCell('AF5');
    AF5.value = VALUES.MINUTES;
    AF5.style = MINUTE_STYLE;

    let AG5 = worksheet.getCell('AG5');
    AG5.value = VALUES.PERCENT;
    AG5.style = PERCENT_STYLE;
    let AH5 = worksheet.getCell('AH5');
    AH5.value = VALUES.MINUTES;
    AH5.style = MINUTE_STYLE;

    workSheet.getCell('M6').value = gpSlaSummary.up_percent;
    workSheet.getCell('N6').value = gpSlaSummary.up_minutes;
    workSheet.getCell('O6').value = gpSlaSummary.total_down_percent;
    workSheet.getCell('P6').value = gpSlaSummary.total_down_minutes;
    workSheet.getCell('Q6').value = gpSlaSummary.power_down_percent;
    workSheet.getCell('R6').value = gpSlaSummary.power_down_minutes;
    workSheet.getCell('S6').value = gpSlaSummary.fibre_down_percent;
    workSheet.getCell('T6').value = gpSlaSummary.fibre_down_minutes;
    workSheet.getCell('U6').value = gpSlaSummary.equipment_down_percent;
    workSheet.getCell('V6').value = gpSlaSummary.equipment_down_minutes;
    workSheet.getCell('W6').value = gpSlaSummary.hrt_down_percent;
    workSheet.getCell('X6').value = gpSlaSummary.hrt_down_minutes;
    workSheet.getCell('Y6').value = gpSlaSummary.dcn_down_percent;
    workSheet.getCell('Z6').value = gpSlaSummary.dcn_down_minutes;
    workSheet.getCell('Z6').value = gpSlaSummary.dcn_down_minutes;
    workSheet.getCell('AA6').value = gpSlaSummary.planned_maintenance_percent;
    workSheet.getCell('AB6').value = gpSlaSummary.planned_maintenance_minutes;
    workSheet.getCell('AC6').value = gpSlaSummary.unknown_downtime_in_percent;
    workSheet.getCell('AD6').value = gpSlaSummary.unknown_downtime_in_minutes;
    workSheet.getCell('AE6').value = gpSlaSummary.total_sla_exclusion_percent;
    workSheet.getCell('AF6').value = gpSlaSummary.total_sla_exclusion_minutes;
    workSheet.getCell('AG6').value = gpSlaSummary.total_up_percent_exclusion;
    workSheet.getCell('AH6').value = gpSlaSummary.total_up_minutes_exclusion;

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

    worksheet.mergeCells('A9:B9');
    let cellA11 = worksheet.getCell('A9');
    cellA11.value = 'GP - SLA Device Level (%) & (Min)';
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
    worksheet.mergeCells('K10:K11');
    worksheet.mergeCells('L10:L11');

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
    worksheet.mergeCells('AG10:AH10');
    worksheet.mergeCells('AI10:AJ10');

    worksheet.getCell('A10').value = GP_FINAL_REPORT_DEVICE_LEVEL_HEADERS[0];
    worksheet.getCell('B10').value = GP_FINAL_REPORT_DEVICE_LEVEL_HEADERS[1];
    worksheet.getCell('C10').value = GP_FINAL_REPORT_DEVICE_LEVEL_HEADERS[2];
    worksheet.getCell('D10').value = GP_FINAL_REPORT_DEVICE_LEVEL_HEADERS[3];
    worksheet.getCell('E10').value = GP_FINAL_REPORT_DEVICE_LEVEL_HEADERS[4];
    worksheet.getCell('F10').value = GP_FINAL_REPORT_DEVICE_LEVEL_HEADERS[5];
    worksheet.getCell('G10').value = GP_FINAL_REPORT_DEVICE_LEVEL_HEADERS[6];
    worksheet.getCell('H10').value = GP_FINAL_REPORT_DEVICE_LEVEL_HEADERS[7];
    worksheet.getCell('I10').value = GP_FINAL_REPORT_DEVICE_LEVEL_HEADERS[8];
    worksheet.getCell('J10').value = GP_FINAL_REPORT_DEVICE_LEVEL_HEADERS[9];
    worksheet.getCell('K10').value = GP_FINAL_REPORT_DEVICE_LEVEL_HEADERS[10];
    worksheet.getCell('L10').value = GP_FINAL_REPORT_DEVICE_LEVEL_HEADERS[11];

    worksheet.getCell('M10').value = GP_FINAL_REPORT_DEVICE_LEVEL_HEADERS[12];
    worksheet.getCell('O10').value = GP_FINAL_REPORT_DEVICE_LEVEL_HEADERS[13];
    worksheet.getCell('Q10').value = GP_FINAL_REPORT_DEVICE_LEVEL_HEADERS[14];
    worksheet.getCell('S10').value = GP_FINAL_REPORT_DEVICE_LEVEL_HEADERS[15];
    worksheet.getCell('U10').value = GP_FINAL_REPORT_DEVICE_LEVEL_HEADERS[16];
    worksheet.getCell('W10').value = GP_FINAL_REPORT_DEVICE_LEVEL_HEADERS[17];
    worksheet.getCell('Y10').value = GP_FINAL_REPORT_DEVICE_LEVEL_HEADERS[18];
    worksheet.getCell('AA10').value = GP_FINAL_REPORT_DEVICE_LEVEL_HEADERS[19];
    worksheet.getCell('AC10').value = GP_FINAL_REPORT_DEVICE_LEVEL_HEADERS[20];
    worksheet.getCell('AE10').value = GP_FINAL_REPORT_DEVICE_LEVEL_HEADERS[21];
    worksheet.getCell('AG10').value = GP_FINAL_REPORT_DEVICE_LEVEL_HEADERS[22];
    worksheet.getCell('AI10').value = GP_FINAL_REPORT_DEVICE_LEVEL_HEADERS[23];

    let finalReportHeaders = worksheet.getRow(10);

    finalReportHeaders.eachCell((cell) => {
      cell.style = TABLE_HEADERS;
    });

    let M11 = worksheet.getCell('M11');
    M11.value = VALUES.PERCENT;
    M11.style = PERCENT_STYLE;
    let N11 = worksheet.getCell('N11');
    N11.value = VALUES.MINUTES;
    N11.style = MINUTE_STYLE;

    let O11 = worksheet.getCell('O11');
    O11.value = VALUES.PERCENT;
    O11.style = PERCENT_STYLE;
    let P11 = worksheet.getCell('P11');
    P11.value = VALUES.MINUTES;
    P11.style = MINUTE_STYLE;

    let Q11 = worksheet.getCell('Q11');
    Q11.value = VALUES.PERCENT;
    Q11.style = PERCENT_STYLE;
    let R11 = worksheet.getCell('R11');
    R11.value = VALUES.MINUTES;
    R11.style = MINUTE_STYLE;

    let S11 = worksheet.getCell('S11');
    S11.value = VALUES.PERCENT;
    S11.style = PERCENT_STYLE;
    let T11 = worksheet.getCell('T11');
    T11.value = VALUES.MINUTES;
    T11.style = MINUTE_STYLE;

    let U11 = worksheet.getCell('U11');
    U11.value = VALUES.PERCENT;
    U11.style = PERCENT_STYLE;
    let V11 = worksheet.getCell('V11');
    V11.value = VALUES.MINUTES;
    V11.style = MINUTE_STYLE;

    let W11 = worksheet.getCell('W11');
    W11.value = VALUES.PERCENT;
    W11.style = PERCENT_STYLE;
    let X11 = worksheet.getCell('X11');
    X11.value = VALUES.MINUTES;
    X11.style = MINUTE_STYLE;

    let Y11 = worksheet.getCell('Y11');
    Y11.value = VALUES.PERCENT;
    Y11.style = PERCENT_STYLE;
    let Z11 = worksheet.getCell('Z11');
    Z11.value = VALUES.MINUTES;
    Z11.style = MINUTE_STYLE;

    let AA11 = worksheet.getCell('AA11');
    AA11.value = VALUES.PERCENT;
    AA11.style = PERCENT_STYLE;
    let AB11 = worksheet.getCell('AB11');
    AB11.value = VALUES.MINUTES;
    AB11.style = MINUTE_STYLE;

    let AC11 = worksheet.getCell('AC11');
    AC11.value = VALUES.PERCENT;
    AC11.style = PERCENT_STYLE;
    let AD11 = worksheet.getCell('AD11');
    AD11.value = VALUES.MINUTES;
    AD11.style = MINUTE_STYLE;

    let AE11 = worksheet.getCell('AE11');
    AE11.value = VALUES.PERCENT;
    AE11.style = PERCENT_STYLE;
    let AF11 = worksheet.getCell('AF11');
    AF11.value = VALUES.MINUTES;
    AF11.style = MINUTE_STYLE;

    let AG11 = worksheet.getCell('AG11');
    AG11.value = VALUES.PERCENT;
    AG11.style = PERCENT_STYLE;
    let AH11 = worksheet.getCell('AH11');
    AH11.value = VALUES.MINUTES;
    AH11.style = MINUTE_STYLE;

    let AI11 = worksheet.getCell('AI11');
    AI11.value = VALUES.PERCENT;
    AI11.style = PERCENT_STYLE;
    let AJ11 = worksheet.getCell('AJ11');
    AJ11.value = VALUES.MINUTES;
    AJ11.style = MINUTE_STYLE;

    let row11 = worksheet.getRow(11);
    row11.eachCell((cell) => {
      cell.border = BORDER_STYLE;
      cell.font = { bold: true };
      cell.alignment = { horizontal: 'center' };
    });

    manipulatedGpNmsData.forEach((row: any) => {
      let gp_device_details = GP_DEVICE_DETAILS.filter(
        (deviceDetails: GpDeviceDetails) =>
          deviceDetails.gp_ip_address == row.ip_address
      )[0];

      let reportType: string = gp_device_details.report_type;
      let hostName: string = gp_device_details.host_name;
      let gpIpAddress: string = gp_device_details.gp_ip_address;
      let state: string = gp_device_details.gp_ip_address;
      let cluster: string = gp_device_details.cluster;
      let district: string = gp_device_details.district;
      let distrctlgdCode: number = gp_device_details.district_lgd_code;
      let blockName: string = gp_device_details.block_name;
      let blockIpAddress: string = gp_device_details.block_ip_address;
      let blockLgdCode: string = gp_device_details.block_lgd_code;
      let gpName: string = gp_device_details.gp_name;
      let gpLgdCode: number = gp_device_details.gp_lgd_code;

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
      let plannedMaintanancePercent: number =
        upPercent == 100 ? 0 : row.planned_maintenance_in_percent;
      let plannedMaintananceMinutes: number =
        upPercent == 100 ? 0 : row.planned_maintenance_in_minutes;
      let unKnownDownPercent: number =
        upPercent == 100 ? 0 : row.unknown_downtime_in_percent;
      let unKnownDownMinutes: number =
        upPercent == 100 ? 0 : row.unknown_downtime_in_minutes;
      let totalExclusionPercent: number =
        upPercent == 100
          ? 0
          : row.power_downtime_in_percent +
            row.dcn_downtime_in_percent +
            row.planned_maintenance_in_percent;
      let totalExclusionMinutes: number =
        row.power_downtime_in_minutes +
        row.dcn_downtime_in_minutes +
        row.planned_maintenance_in_minutes;
      let pollingTimePercent: number =
        upPercent == 100 ? 0 : row.down_percent - +totalExclusionPercent;
      let pollingTimeMinutes: number =
        upPercent == 100
          ? 0
          : row.total_downtime_in_minutes - totalExclusionMinutes;

      let totalUpPercentSLAExclusion: number =
        upPercent + totalExclusionPercent + pollingTimePercent;
      let totalUpMinutesSLAExclusion: number =
        upMinute + totalExclusionMinutes + pollingTimeMinutes;

      const gpDeviceLevelRowValues = worksheet.addRow([
        reportType,
        hostName,
        gpIpAddress,
        state,
        cluster,
        district,
        distrctlgdCode,
        blockName,
        blockIpAddress,
        blockLgdCode,
        gpName,
        gpLgdCode,
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
        unKnownDownPercent.toFixed(2),
        unKnownDownMinutes.toFixed(2),
        totalExclusionPercent.toFixed(2),
        totalExclusionMinutes.toFixed(2),
        pollingTimePercent.toFixed(2),
        pollingTimeMinutes.toFixed(2),
        totalUpPercentSLAExclusion.toFixed(2),
        totalUpMinutesSLAExclusion.toFixed(2),
      ]);

      gpDeviceLevelRowValues.eachCell((cell) => {
        cell.border = BORDER_STYLE;
        cell.alignment = { horizontal: 'left' };
      });
    });
  }
}
