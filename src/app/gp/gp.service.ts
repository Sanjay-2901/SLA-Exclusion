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
  GP_TT_CO_RELATION_COLUMN_WIDTHS,
  GP_TT_CO_RELATION_HEADERS,
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

import {
  BlockAlertData,
  BlockDeviceLevelHeaders,
  RFOCategorizedTimeInMinutes,
  TTCorelation,
} from '../block-component/block-component.model';

@Injectable({
  providedIn: 'root',
})
export class GpService {
  ttCorelation: TTCorelation[] = [];

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

  checkBlockAlarmDeviation(
    gpAlarmStartTime: any,
    blockAlarmStartTime: any
  ): boolean {
    let gp_time = new Date(gpAlarmStartTime);
    let block_time = new Date(blockAlarmStartTime);
    let tenMinutesInMilliseconds = 600000;

    let lowerBound = new Date(gp_time.getTime() - tenMinutesInMilliseconds);
    let upperBound = new Date(gp_time.getTime() + tenMinutesInMilliseconds);

    if (block_time >= lowerBound && block_time <= upperBound) {
      return true;
    } else {
      return false;
    }
  }

  categorizeRFO(
    nmsData: GpNMSData,
    gpAlertData: GpAlertData[],
    gpTTData: GpTTData[],
    blockAlertData: BlockAlertData[]
  ) {
    if (nmsData.up_percent !== 100) {
      let totalPowerDownTimeInMinutes = 0;
      let totalDCNDownTimeInMinutes = 0;
      let isAlertReportEmpty: boolean = false;

      let powerDownArray: GpAlertData[] = [];
      let DCNDownArray: GpAlertData[] = [];
      let criticalAlertAndTTDataTimeMismatch: GpAlertData[] = [];

      let powerIssueTT: string[] = [];
      let linkIssueTT: string[] = [];
      let otherTT: string[] = [];

      const filteredCriticalGpAlertData = gpAlertData.filter(
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

      const correspondingBlockIpForGp = GP_DEVICE_DETAILS.filter(
        (gpDeviceDetail: GpDeviceDetails) => {
          return nmsData.ip_address === gpDeviceDetail.gp_ip_address;
        }
      )[0].block_ip_address;

      const filteredBlockAlertData = blockAlertData.filter(
        (blockAlertData: BlockAlertData) => {
          return (
            blockAlertData.ip_address == correspondingBlockIpForGp &&
            blockAlertData.severity == SEVERITY_CRITICAL &&
            blockAlertData.message == ALERT_DOWN_MESSAGE
          );
        }
      );

      if (filteredCriticalGpAlertData.length) {
        // scenario 1 : Checking whether GP Down due to BLOCK Down
        filteredCriticalGpAlertData.forEach(
          (gpAlertCriticalData: GpAlertData) => {
            let matchingGpAlerts: GpAlertData[] = [];
            filteredBlockAlertData.forEach(
              (blockAlertCriticalData: BlockAlertData) => {
                let isTenMinutesDeviationFoundForStartTime: boolean =
                  this.checkBlockAlarmDeviation(
                    gpAlertCriticalData.alarm_start_time,
                    blockAlertCriticalData.alarm_start_time
                  );

                if (isTenMinutesDeviationFoundForStartTime) {
                  matchingGpAlerts.push(gpAlertCriticalData);

                  let gpAlertAndBlockAlertDifference = +(
                    gpAlertCriticalData.total_duration_in_minutes -
                    blockAlertCriticalData.total_duration_in_minutes
                  ).toFixed(0);

                  if (
                    lodash.countBy(matchingGpAlerts, gpAlertCriticalData)[
                      'true'
                    ] === 1
                  ) {
                    if (
                      // (gpAlertAndBlockAlertDifference >= 0 &&
                      //   gpAlertAndBlockAlertDifference <= 10) ||
                      gpAlertAndBlockAlertDifference <= 0
                    ) {
                      totalDCNDownTimeInMinutes +=
                        gpAlertCriticalData.total_duration_in_minutes;
                      gpAlertCriticalData.total_duration_in_minutes = 0;
                    } else {
                      totalDCNDownTimeInMinutes +=
                        blockAlertCriticalData.total_duration_in_minutes;
                      gpAlertCriticalData.total_duration_in_minutes =
                        gpAlertAndBlockAlertDifference;
                    }
                  }
                }
              }
            );
            matchingGpAlerts = [];
          }
        );

        // scenario 2 : Checking with NOC TT RFO
        filteredCriticalGpAlertData.forEach(
          (alertCriticalData: GpAlertData) => {
            filteredTTData.forEach((ttData: GpTTData) => {
              if (
                moment(alertCriticalData.alarm_start_time).isSame(
                  ttData.incident_start_on,
                  'minute'
                )
              ) {
                if (ttData.rfo == RFO_CATEGORIZATION.POWER_ISSUE) {
                  if (
                    !lodash.some(DCNDownArray, alertCriticalData) &&
                    !lodash.some(powerDownArray, alertCriticalData)
                  ) {
                    powerIssueTT.push(ttData.incident_id);
                    powerDownArray.push(alertCriticalData);
                  }
                } else if (
                  ttData.rfo == RFO_CATEGORIZATION.JIO_LINK_ISSUE ||
                  ttData.rfo == RFO_CATEGORIZATION.SWAN_ISSUE
                ) {
                  if (
                    !lodash.some(DCNDownArray, alertCriticalData) &&
                    !lodash.some(powerDownArray, alertCriticalData)
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

      // scenario :3  Checking with warning Alerts
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
                  if (
                    !lodash.some(powerDownArray, alertCriticalData) &&
                    !lodash.some(DCNDownArray, alertCriticalData)
                  ) {
                    powerDownArray.push(alertCriticalData);
                  }
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
    let totalDownPercent = 0;
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
      totalDownPercent += nmsData.down_percent;
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
      pollingTimePercent += nmsData.polling_time_in_percent;
      pollingTimeMinutes += nmsData.polling_time_in_minutes;
    });

    return {
      report_type: 'GP-SLA',
      time_span: '',
      no_of_blocks: 79,
      up_percent: (upPercent / 5001).toFixed(2),
      up_minutes: upMinutes.toFixed(2),
      no_of_up_blocks: '',
      down_percent_exclusive_of_sla: (100 - upPercent / 5001).toFixed(2),
      power_down_percent: (powerDownPercent / 5001).toFixed(2),
      power_down_minutes: powerDownMinutes.toFixed(2),
      fibre_down_percent: (fiberDownPercent / 5001).toFixed(2),
      fibre_down_minutes: fiberDownMinute.toFixed(2),
      equipment_down_percent: (equipmentDownPercent / 5001).toFixed(2),
      equipment_down_minutes: equipmentDownMinute.toFixed(2),
      hrt_down_percent: (hrtDownPercent / 5001).toFixed(2),
      hrt_down_minutes: hrtDownMinute.toFixed(2),
      dcn_down_percent: (dcnDownPercent / 5001).toFixed(2),
      dcn_down_minutes: dcnDownMinutes.toFixed(2),
      planned_maintenance_percent: (plannedMaintenancePercent / 5001).toFixed(
        2
      ),
      planned_maintenance_minutes: plannedMaintenanceMinutes.toFixed(2),
      unknown_downtime_in_percent: (unKnownDownPercent / 5001).toFixed(2),
      unknown_downtime_in_minutes: unKnownDownMinutes.toFixed(2),
      total_sla_exclusion_percent: (cumulativeRfoDownInPercent / 5001).toFixed(
        2
      ),
      total_sla_exclusion_minutes: cumulativeRfoDownInMinutes.toFixed(2),
      total_down_minutes: totalDownMinutes.toFixed(2),
      total_down_percent: (100 - +(upPercent / 5001)).toFixed(2),
      total_up_percent_exclusion: (
        (upPercent + totalDownPercent) /
        5001
      ).toFixed(2),

      total_up_minutes_exclusion: (upMinutes + totalDownMinutes).toFixed(2),
    };
  }

  FrameGpFinalSlaReportWorkbook(
    workbook: ExcelJS.Workbook,
    workSheet: ExcelJS.Worksheet,
    timeSpan: string,
    gpSlaSummary: GpSLASummary,
    manipulatedGpNmsData: ManipulatedGpNMSData[],
    blockFinalreport: BlockDeviceLevelHeaders[],
    blockTTCorelationReport: TTCorelation[]
  ): void {
    workSheet.columns = GP_SLA_FINAL_REPORT_COLUMN_WIDTHS;

    workSheet.mergeCells('A1:B1');
    let cellA1 = workSheet.getCell('A1');
    cellA1.value = '1. Daily Network availability report';
    cellA1.style = SHEET_HEADING;

    workSheet.mergeCells('C1:D1');
    let cellC1 = workSheet.getCell('C1');
    cellC1.value = 'Report-Frequency- Daily';
    cellC1.style = {
      font: { bold: true },
      alignment: { horizontal: 'center' },
    };

    workSheet.mergeCells('A3:B3');
    let cellA3 = workSheet.getCell('A3');
    cellA3.value = 'GP- SLA Summary (%) & (Min)';
    cellA3.style = TABLE_HEADING;
    workSheet.getCell('B3').style = TABLE_HEADING;

    workSheet.mergeCells('C4:J4');
    workSheet.mergeCells('K4:L4');
    workSheet.mergeCells('M4:N4');
    workSheet.mergeCells('O4:P4');
    workSheet.mergeCells('Q4:R4');
    workSheet.mergeCells('S4:T4');
    workSheet.mergeCells('U4:V4');
    workSheet.mergeCells('W4:X4');
    workSheet.mergeCells('Y4:Z4');
    workSheet.mergeCells('AA4:AB4');
    workSheet.mergeCells('AC4:AD4');
    workSheet.mergeCells('AE4:AF4');
    workSheet.mergeCells('AG4:AH4');

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

    let gpSummaryHeadersRow = workSheet.getRow(4);
    gpSummaryHeadersRow.eachCell((cell) => {
      cell.style = TABLE_HEADERS;
    });

    workSheet.mergeCells('A5:A6');
    workSheet.mergeCells('B5:B6');
    workSheet.mergeCells('C5:J6');
    workSheet.mergeCells('K5:L6');

    workSheet.getCell('A5').value = 'GP - SLA';
    workSheet.getCell('B5').value = 'O&M GP';
    workSheet.getCell('C5').value = timeSpan.replace(/Time Span: /, '');
    workSheet.getCell('K5').value = '5001';

    let M5 = workSheet.getCell('M5');
    M5.value = VALUES.PERCENT;
    M5.style = PERCENT_STYLE;
    let N5 = workSheet.getCell('N5');
    N5.value = VALUES.MINUTES;
    N5.style = MINUTE_STYLE;

    let O5 = workSheet.getCell('O5');
    O5.value = VALUES.PERCENT;
    O5.style = PERCENT_STYLE;
    let P5 = workSheet.getCell('P5');
    P5.value = VALUES.MINUTES;
    P5.style = MINUTE_STYLE;

    let Q5 = workSheet.getCell('Q5');
    Q5.value = VALUES.PERCENT;
    Q5.style = PERCENT_STYLE;
    let R5 = workSheet.getCell('R5');
    R5.value = VALUES.MINUTES;
    R5.style = MINUTE_STYLE;

    let S5 = workSheet.getCell('S5');
    S5.value = VALUES.PERCENT;
    S5.style = PERCENT_STYLE;
    let T5 = workSheet.getCell('T5');
    T5.value = VALUES.MINUTES;
    T5.style = MINUTE_STYLE;

    let U5 = workSheet.getCell('U5');
    U5.value = VALUES.PERCENT;
    U5.style = PERCENT_STYLE;
    let V5 = workSheet.getCell('V5');
    V5.value = VALUES.MINUTES;
    V5.style = MINUTE_STYLE;

    let W5 = workSheet.getCell('W5');
    W5.value = VALUES.PERCENT;
    W5.style = PERCENT_STYLE;
    let X5 = workSheet.getCell('X5');
    X5.value = VALUES.MINUTES;
    X5.style = MINUTE_STYLE;

    let Y5 = workSheet.getCell('Y5');
    Y5.value = VALUES.PERCENT;
    Y5.style = PERCENT_STYLE;
    let Z5 = workSheet.getCell('Z5');
    Z5.value = VALUES.MINUTES;
    Z5.style = MINUTE_STYLE;

    let AA5 = workSheet.getCell('AA5');
    AA5.value = VALUES.PERCENT;
    AA5.style = PERCENT_STYLE;
    let AB5 = workSheet.getCell('AB5');
    AB5.value = VALUES.MINUTES;
    AB5.style = MINUTE_STYLE;

    let AC5 = workSheet.getCell('AC5');
    AC5.value = VALUES.PERCENT;
    AC5.style = PERCENT_STYLE;
    let AD5 = workSheet.getCell('AD5');
    AD5.value = VALUES.MINUTES;
    AD5.style = MINUTE_STYLE;

    let AE5 = workSheet.getCell('AE5');
    AE5.value = VALUES.PERCENT;
    AE5.style = PERCENT_STYLE;
    let AF5 = workSheet.getCell('AF5');
    AF5.value = VALUES.MINUTES;
    AF5.style = MINUTE_STYLE;

    let AG5 = workSheet.getCell('AG5');
    AG5.value = VALUES.PERCENT;
    AG5.style = PERCENT_STYLE;
    let AH5 = workSheet.getCell('AH5');
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

    let row5 = workSheet.getRow(5);
    row5.eachCell((cell) => {
      cell.border = BORDER_STYLE;
      cell.alignment = { horizontal: 'center' };
      cell.font = { bold: true };
    });

    let row6 = workSheet.getRow(6);
    row6.eachCell((cell) => {
      cell.border = BORDER_STYLE;
      cell.alignment = { horizontal: 'center' };
    });

    workSheet.mergeCells('A9:B9');
    let cellA11 = workSheet.getCell('A9');
    cellA11.value = 'GP - SLA Device Level (%) & (Min)';
    cellA11.style = TABLE_HEADING;
    workSheet.getCell('B9').style = TABLE_HEADERS;

    workSheet.mergeCells('A10:A11');
    workSheet.mergeCells('B10:B11');
    workSheet.mergeCells('C10:C11');
    workSheet.mergeCells('D10:D11');
    workSheet.mergeCells('E10:E11');
    workSheet.mergeCells('F10:F11');
    workSheet.mergeCells('G10:G11');
    workSheet.mergeCells('H10:H11');
    workSheet.mergeCells('I10:I11');
    workSheet.mergeCells('J10:J11');
    workSheet.mergeCells('K10:K11');
    workSheet.mergeCells('L10:L11');

    workSheet.mergeCells('M10:N10');
    workSheet.mergeCells('O10:P10');
    workSheet.mergeCells('Q10:R10');
    workSheet.mergeCells('S10:T10');
    workSheet.mergeCells('U10:V10');
    workSheet.mergeCells('W10:X10');
    workSheet.mergeCells('Y10:Z10');
    workSheet.mergeCells('AA10:AB10');
    workSheet.mergeCells('AC10:AD10');
    workSheet.mergeCells('AE10:AF10');
    workSheet.mergeCells('AG10:AH10');
    workSheet.mergeCells('AI10:AJ10');

    workSheet.getCell('A10').value = GP_FINAL_REPORT_DEVICE_LEVEL_HEADERS[0];
    workSheet.getCell('B10').value = GP_FINAL_REPORT_DEVICE_LEVEL_HEADERS[1];
    workSheet.getCell('C10').value = GP_FINAL_REPORT_DEVICE_LEVEL_HEADERS[2];
    workSheet.getCell('D10').value = GP_FINAL_REPORT_DEVICE_LEVEL_HEADERS[3];
    workSheet.getCell('E10').value = GP_FINAL_REPORT_DEVICE_LEVEL_HEADERS[4];
    workSheet.getCell('F10').value = GP_FINAL_REPORT_DEVICE_LEVEL_HEADERS[5];
    workSheet.getCell('G10').value = GP_FINAL_REPORT_DEVICE_LEVEL_HEADERS[6];
    workSheet.getCell('H10').value = GP_FINAL_REPORT_DEVICE_LEVEL_HEADERS[7];
    workSheet.getCell('I10').value = GP_FINAL_REPORT_DEVICE_LEVEL_HEADERS[8];
    workSheet.getCell('J10').value = GP_FINAL_REPORT_DEVICE_LEVEL_HEADERS[9];
    workSheet.getCell('K10').value = GP_FINAL_REPORT_DEVICE_LEVEL_HEADERS[10];
    workSheet.getCell('L10').value = GP_FINAL_REPORT_DEVICE_LEVEL_HEADERS[11];

    workSheet.getCell('M10').value = GP_FINAL_REPORT_DEVICE_LEVEL_HEADERS[12];
    workSheet.getCell('O10').value = GP_FINAL_REPORT_DEVICE_LEVEL_HEADERS[13];
    workSheet.getCell('Q10').value = GP_FINAL_REPORT_DEVICE_LEVEL_HEADERS[14];
    workSheet.getCell('S10').value = GP_FINAL_REPORT_DEVICE_LEVEL_HEADERS[15];
    workSheet.getCell('U10').value = GP_FINAL_REPORT_DEVICE_LEVEL_HEADERS[16];
    workSheet.getCell('W10').value = GP_FINAL_REPORT_DEVICE_LEVEL_HEADERS[17];
    workSheet.getCell('Y10').value = GP_FINAL_REPORT_DEVICE_LEVEL_HEADERS[18];
    workSheet.getCell('AA10').value = GP_FINAL_REPORT_DEVICE_LEVEL_HEADERS[19];
    workSheet.getCell('AC10').value = GP_FINAL_REPORT_DEVICE_LEVEL_HEADERS[20];
    workSheet.getCell('AE10').value = GP_FINAL_REPORT_DEVICE_LEVEL_HEADERS[21];
    workSheet.getCell('AG10').value = GP_FINAL_REPORT_DEVICE_LEVEL_HEADERS[22];
    workSheet.getCell('AI10').value = GP_FINAL_REPORT_DEVICE_LEVEL_HEADERS[23];

    let finalReportHeaders = workSheet.getRow(10);

    finalReportHeaders.eachCell((cell) => {
      cell.style = TABLE_HEADERS;
    });

    let M11 = workSheet.getCell('M11');
    M11.value = VALUES.PERCENT;
    M11.style = PERCENT_STYLE;
    let N11 = workSheet.getCell('N11');
    N11.value = VALUES.MINUTES;
    N11.style = MINUTE_STYLE;

    let O11 = workSheet.getCell('O11');
    O11.value = VALUES.PERCENT;
    O11.style = PERCENT_STYLE;
    let P11 = workSheet.getCell('P11');
    P11.value = VALUES.MINUTES;
    P11.style = MINUTE_STYLE;

    let Q11 = workSheet.getCell('Q11');
    Q11.value = VALUES.PERCENT;
    Q11.style = PERCENT_STYLE;
    let R11 = workSheet.getCell('R11');
    R11.value = VALUES.MINUTES;
    R11.style = MINUTE_STYLE;

    let S11 = workSheet.getCell('S11');
    S11.value = VALUES.PERCENT;
    S11.style = PERCENT_STYLE;
    let T11 = workSheet.getCell('T11');
    T11.value = VALUES.MINUTES;
    T11.style = MINUTE_STYLE;

    let U11 = workSheet.getCell('U11');
    U11.value = VALUES.PERCENT;
    U11.style = PERCENT_STYLE;
    let V11 = workSheet.getCell('V11');
    V11.value = VALUES.MINUTES;
    V11.style = MINUTE_STYLE;

    let W11 = workSheet.getCell('W11');
    W11.value = VALUES.PERCENT;
    W11.style = PERCENT_STYLE;
    let X11 = workSheet.getCell('X11');
    X11.value = VALUES.MINUTES;
    X11.style = MINUTE_STYLE;

    let Y11 = workSheet.getCell('Y11');
    Y11.value = VALUES.PERCENT;
    Y11.style = PERCENT_STYLE;
    let Z11 = workSheet.getCell('Z11');
    Z11.value = VALUES.MINUTES;
    Z11.style = MINUTE_STYLE;

    let AA11 = workSheet.getCell('AA11');
    AA11.value = VALUES.PERCENT;
    AA11.style = PERCENT_STYLE;
    let AB11 = workSheet.getCell('AB11');
    AB11.value = VALUES.MINUTES;
    AB11.style = MINUTE_STYLE;

    let AC11 = workSheet.getCell('AC11');
    AC11.value = VALUES.PERCENT;
    AC11.style = PERCENT_STYLE;
    let AD11 = workSheet.getCell('AD11');
    AD11.value = VALUES.MINUTES;
    AD11.style = MINUTE_STYLE;

    let AE11 = workSheet.getCell('AE11');
    AE11.value = VALUES.PERCENT;
    AE11.style = PERCENT_STYLE;
    let AF11 = workSheet.getCell('AF11');
    AF11.value = VALUES.MINUTES;
    AF11.style = MINUTE_STYLE;

    let AG11 = workSheet.getCell('AG11');
    AG11.value = VALUES.PERCENT;
    AG11.style = PERCENT_STYLE;
    let AH11 = workSheet.getCell('AH11');
    AH11.value = VALUES.MINUTES;
    AH11.style = MINUTE_STYLE;

    let AI11 = workSheet.getCell('AI11');
    AI11.value = VALUES.PERCENT;
    AI11.style = PERCENT_STYLE;
    let AJ11 = workSheet.getCell('AJ11');
    AJ11.value = VALUES.MINUTES;
    AJ11.style = MINUTE_STYLE;

    let row11 = workSheet.getRow(11);
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

      let block_output_report_details = blockFinalreport.filter(
        (blockDetails: BlockDeviceLevelHeaders) =>
          blockDetails.ip_address == row.ip_address
      )[0];

      let reportType: string = gp_device_details.report_type;
      let hostName: string = gp_device_details.host_name;
      let gpIpAddress: string = gp_device_details.gp_ip_address;
      let state: string = gp_device_details.state;
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
        upPercent == 100 ? 0 : row.polling_time_in_percent;
      let pollingTimeMinutes: number =
        upPercent == 100 ? 0 : row.polling_time_in_minutes;

      let totalUpPercentSLAExclusion: number = upPercent + downPercent;
      let totalUpMinutesSLAExclusion: number = upMinute + downMinute;

      const gpDeviceLevelRowValues = workSheet.addRow([
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

    //Generating TT co-relation report for GP

    const gpTtCoRelationWorkSheet = workbook.addWorksheet('GP TT co-relation');
    gpTtCoRelationWorkSheet.columns = GP_TT_CO_RELATION_COLUMN_WIDTHS;
    gpTtCoRelationWorkSheet
      .addRow(GP_TT_CO_RELATION_HEADERS)
      .eachCell((cell) => {
        cell.style = TABLE_HEADERS;
      });

    this.ttCorelation.forEach(
      (ttCorelationData: TTCorelation, index: number) => {
        let gpDevicedetails = GP_DEVICE_DETAILS.filter(
          (gpDevice: GpDeviceDetails) =>
            gpDevice.gp_ip_address === ttCorelationData.ip
        )[0];
        let blockTTCoRelation: TTCorelation = blockTTCorelationReport.filter(
          (blockTTCoRelationData: TTCorelation) => {
            return (
              blockTTCoRelationData.ip === gpDevicedetails.block_ip_address
            );
          }
        )[0];
        gpTtCoRelationWorkSheet
          .addRow([
            index + 1,
            ttCorelationData.ip,
            gpDevicedetails.block_name,
            gpDevicedetails.gp_name,
            blockTTCoRelation?.powerIssueTT
              ? blockTTCoRelation.powerIssueTT[0]
              : '',
            blockTTCoRelation?.linkIssueTT
              ? blockTTCoRelation.linkIssueTT[0]
              : '',
            ttCorelationData?.powerIssueTT
              ? ttCorelationData.powerIssueTT.toString().split(',').join(', ')
              : '',
            ttCorelationData?.linkIssueTT
              ? ttCorelationData.linkIssueTT.toString().split(',').join(', ')
              : '',
            ttCorelationData?.otherTT
              ? ttCorelationData.otherTT.toString().split(',').join(', ')
              : '',
          ])
          .eachCell((cell) => {
            cell.border = BORDER_STYLE;
            cell.alignment = { horizontal: 'left' };
          });
      }
    );
  }
}
