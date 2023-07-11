import { Injectable } from '@angular/core';
import {
  ManipulatedShqNmsData,
  ShqAlertData,
  ShqNMSData,
  ShqSlaSummary,
  ShqTTData,
} from './shq-component.model';
import {
  ALERT_DOWN_MESSAGE,
  BORDER_STYLE,
  MINUTE_STYLE,
  PERCENT_STYLE,
  RFO_CATEGORIZATION,
  SEVERITY_CRITICAL,
  SEVERITY_WARNING,
  SHEET_HEADING,
  SHQ_DEVICE_LEVEL_HEADERS,
  SHQ_SLQ_FINAL_REPORT_COLUMN_WIDTHS,
  SHQ_SUMMARY_HEADERS,
  SHQ_TT_CO_RELATION_COLUMNS,
  SHQ_TT_CO_RELATION_HEADERS,
  TABLE_HEADERS,
  TABLE_HEADING,
  VALUES,
  VMWAREDEVICE,
} from '../constants/constants';
import * as moment from 'moment';
import * as lodash from 'lodash';
import * as ExcelJS from 'exceljs';
import {
  RFOCategorizedTimeInMinutes,
  TTCorelation,
} from '../block-component/block-component.model';

@Injectable({
  providedIn: 'root',
})
export class ShqService {
  ttCorelation: TTCorelation[] = [];

  constructor() {}

  shqNMSDatawithoutVmwareDevices(
    AllShqDevicesArray: ShqNMSData[]
  ): ShqNMSData[] {
    return AllShqDevicesArray.filter((nmsData) => {
      return nmsData.type.trim() !== VMWAREDEVICE;
    });
  }

  calculateAlertDownTimeInMinutes(
    ipAddress: string,
    shqAlertData: ShqAlertData[]
  ) {
    let filteredAlertData = shqAlertData.filter((alert: ShqAlertData) => {
      return (
        alert.ip_address == ipAddress &&
        alert.severity == SEVERITY_CRITICAL &&
        alert.message == ALERT_DOWN_MESSAGE
      );
    });

    let alertDownTimeInMinutes: number = 0;
    filteredAlertData.forEach((filteredAlertData: ShqAlertData) => {
      alertDownTimeInMinutes += filteredAlertData.total_duration_in_minutes;
    });
    return alertDownTimeInMinutes;
  }

  categorizeRFO(
    nmsData: ShqNMSData,
    shqAlertData: ShqAlertData[],
    shqTTData: ShqTTData[]
  ) {
    if (nmsData.up_percent !== 100) {
      let totalPowerDownTimeInMinutes = 0;
      let totalDCNDownTimeInMinutes = 0;
      let isAlertReportEmpty: boolean = false;

      let powerDownArray: ShqAlertData[] = [];
      let DCNDownArray: ShqAlertData[] = [];
      let criticalAlertAndTTDataTimeMismatch: ShqAlertData[] = [];

      let powerIssueTT: string[] = [];
      let linkIssueTT: string[] = [];
      let otherTT: string[] = [];

      const filteredCriticalAlertData = shqAlertData.filter(
        (alertData: ShqAlertData) => {
          return (
            alertData.ip_address.trim() == nmsData.ip_address &&
            alertData.severity.trim() == SEVERITY_CRITICAL &&
            alertData.message.trim() == ALERT_DOWN_MESSAGE
          );
        }
      );

      const filteredWarningAlertData = shqAlertData.filter(
        (alertData: ShqAlertData) => {
          return (
            alertData.ip_address.trim() == nmsData.ip_address &&
            alertData.severity.trim() == SEVERITY_WARNING &&
            alertData.message.trim().includes('reboot')
          );
        }
      );

      const filteredTTData = shqTTData.filter((ttData: ShqTTData) => {
        return ttData.ip == nmsData.ip_address;
      });

      if (filteredCriticalAlertData.length) {
        filteredCriticalAlertData.forEach((alertCriticalData: ShqAlertData) => {
          filteredTTData.forEach((ttData: ShqTTData) => {
            if (
              moment(alertCriticalData.alarm_start_time).isSame(
                ttData.incident_start_on,
                'minute'
              )
            ) {
              if (ttData.rfo == RFO_CATEGORIZATION.POWER_ISSUE) {
                powerDownArray.push(alertCriticalData);
                powerIssueTT.push(ttData.incident_id);
              } else if (
                ttData.rfo == RFO_CATEGORIZATION.JIO_LINK_ISSUE ||
                ttData.rfo == RFO_CATEGORIZATION.SWAN_ISSUE
              ) {
                DCNDownArray.push(alertCriticalData);
                linkIssueTT.push(ttData.incident_id);
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
        });
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
          (alertCriticalData: ShqAlertData) => {
            filteredWarningAlertData.forEach(
              (alertWarningData: ShqAlertData) => {
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

      powerDownArray.forEach((powerDownAlert: ShqAlertData) => {
        totalPowerDownTimeInMinutes += powerDownAlert.total_duration_in_minutes;
      });

      DCNDownArray.forEach((dcnDownAlert: ShqAlertData) => {
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

  calculateCumulativePercentage(value: number): string {
    return (value / 22).toFixed(2);
  }

  calculateCumulativeMinutes(value: number): string {
    return value.toFixed(2);
  }

  calculateShqSlaSummary(
    manipulatedNMSData: ManipulatedShqNmsData[]
  ): ShqSlaSummary {
    let upPercent = 0;
    let upMinutes = 0;
    let downPercent = 0;
    let downMinutes = 0;
    let totalDownExclusiveOfSlaExclusionInMinute = 0;
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
    let plannedMaitenanceMinute = 0;
    let unKnownDownMinutes = 0;
    let unKnownDownPercent = 0;
    let cumulativeRfoDownInPercent = 0;
    let cumulativeRfoDownInMinutes = 0;
    let totalSlaExclusionPercent = 0;
    let totalSlaExclusionMinute = 0;
    let pollingTimePercent = 0;
    let pollingTimeMinutes = 0;

    manipulatedNMSData.forEach((nmsData: ManipulatedShqNmsData) => {
      upPercent += nmsData.up_percent;
      upMinutes += nmsData.total_uptime_in_minutes;
      downPercent += nmsData.down_percent;
      downMinutes += nmsData.total_downtime_in_minutes;
      totalDownExclusiveOfSlaExclusionInMinute +=
        nmsData.total_downtime_in_minutes;
      powerDownPercent += nmsData.power_downtime_in_percent;
      powerDownMinutes += nmsData.power_downtime_in_minutes;
      dcnDownPercent += nmsData.dcn_downtime_in_percent;
      dcnDownMinutes += nmsData.dcn_downtime_in_minutes;
      unKnownDownMinutes += nmsData.unknown_downtime_in_minutes;
      unKnownDownPercent += nmsData.unknown_downtime_in_percent;
      cumulativeRfoDownInPercent +=
        nmsData.power_downtime_in_percent +
        nmsData.dcn_downtime_in_percent +
        nmsData.unknown_downtime_in_percent;
      cumulativeRfoDownInMinutes +=
        nmsData.power_downtime_in_minutes +
        nmsData.dcn_downtime_in_minutes +
        nmsData.unknown_downtime_in_minutes;
      totalSlaExclusionPercent +=
        nmsData.power_downtime_in_percent +
        nmsData.dcn_downtime_in_percent +
        nmsData.unknown_downtime_in_percent;
      totalSlaExclusionMinute +=
        nmsData.power_downtime_in_minutes +
        nmsData.dcn_downtime_in_minutes +
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
      report_type: 'SHQ-SLA',
      tag: 'SHQ Core Device',
      time_span: '',
      no_of_shq_devices: 22,
      up_percent: this.calculateCumulativePercentage(upPercent),
      up_minutes: this.calculateCumulativeMinutes(upMinutes),
      total_down_exclusive_of_sla_exclusion_percent: (
        100 - parseInt(this.calculateCumulativePercentage(upPercent))
      ).toFixed(2),
      total_down_exclusive_of_sla_exclusion_minute:
        this.calculateCumulativeMinutes(
          totalDownExclusiveOfSlaExclusionInMinute
        ),
      power_down_percent: this.calculateCumulativePercentage(powerDownPercent),
      power_dowm_minute: this.calculateCumulativeMinutes(powerDownMinutes),
      fibre_down_percent: this.calculateCumulativePercentage(fiberDownPercent),
      fiber_down_minute: this.calculateCumulativeMinutes(fiberDownMinute),
      equipment_down_percent:
        this.calculateCumulativePercentage(equipmentDownPercent),
      equipment_down_minute:
        this.calculateCumulativeMinutes(equipmentDownMinute),
      hrt_down_percent: this.calculateCumulativePercentage(hrtDownPercent),
      hrt_down_minute: this.calculateCumulativeMinutes(hrtDownMinute),
      dcn_down_percent: this.calculateCumulativePercentage(dcnDownPercent),
      dcn_down_minute: this.calculateCumulativeMinutes(dcnDownMinutes),
      planned_maintenance_percent: this.calculateCumulativePercentage(
        plannedMaintenancePercent
      ),
      planned_maintenance_minute: this.calculateCumulativeMinutes(
        plannedMaitenanceMinute
      ),
      unknown_downtime_in_percent:
        this.calculateCumulativePercentage(unKnownDownPercent),
      unknown_downtime_in_minutes:
        this.calculateCumulativeMinutes(unKnownDownMinutes),
      total_sla_exclusion_percent: this.calculateCumulativePercentage(
        cumulativeRfoDownInPercent
      ),
      total_sla_exclusion_minute: this.calculateCumulativeMinutes(
        cumulativeRfoDownInMinutes
      ),
      total_up_percent: this.calculateCumulativePercentage(
        upPercent + downPercent
      ),
      total_up_minute: this.calculateCumulativeMinutes(upMinutes + downMinutes),
    };
  }

  getHostName(monitor: string): string {
    return monitor.split(' ')[0];
  }

  FrameShqFinalSlaReportWorkbook(
    workbook: ExcelJS.Workbook,
    workSheet: ExcelJS.Worksheet,
    timeSpan: string,
    shqSlaSummary: ShqSlaSummary,
    manipulatedShqNmsData: ManipulatedShqNmsData[]
  ): void {
    workSheet.columns = SHQ_SLQ_FINAL_REPORT_COLUMN_WIDTHS;

    workSheet.mergeCells('A1:B1:C1');
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

    workSheet.mergeCells('A3:C3');
    let cellA3 = workSheet.getCell('A3');
    cellA3.value = 'SHQ - SLA Summary (%) & (Min)';
    cellA3.style = TABLE_HEADING;
    workSheet.getCell('B3').style = TABLE_HEADING;
    workSheet.getCell('C3').style = TABLE_HEADING;

    workSheet.mergeCells('C4:D4');
    workSheet.mergeCells('C5:D6');
    workSheet.mergeCells('F4:G4');
    workSheet.mergeCells('H4:I4');
    workSheet.mergeCells('J4:K4');
    workSheet.mergeCells('L4:M4');
    workSheet.mergeCells('N4:O4');
    workSheet.mergeCells('P4:Q4');
    workSheet.mergeCells('R4:S4');
    workSheet.mergeCells('T4:U4');
    workSheet.mergeCells('V4:W4');
    workSheet.mergeCells('X4:Y4');
    workSheet.mergeCells('Z4:AA4');

    workSheet.getCell('A4').value = SHQ_SUMMARY_HEADERS[0];
    workSheet.getCell('B4').value = SHQ_SUMMARY_HEADERS[1];
    workSheet.getCell('C4').value = SHQ_SUMMARY_HEADERS[2];
    workSheet.getCell('E4').value = SHQ_SUMMARY_HEADERS[3];
    workSheet.getCell('F4').value = SHQ_SUMMARY_HEADERS[4];
    workSheet.getCell('H4').value = SHQ_SUMMARY_HEADERS[5];
    workSheet.getCell('J4').value = SHQ_SUMMARY_HEADERS[6];
    workSheet.getCell('L4').value = SHQ_SUMMARY_HEADERS[7];
    workSheet.getCell('N4').value = SHQ_SUMMARY_HEADERS[8];
    workSheet.getCell('P4').value = SHQ_SUMMARY_HEADERS[9];
    workSheet.getCell('R4').value = SHQ_SUMMARY_HEADERS[10];
    workSheet.getCell('T4').value = SHQ_SUMMARY_HEADERS[11];
    workSheet.getCell('V4').value = SHQ_SUMMARY_HEADERS[12];
    workSheet.getCell('X4').value = SHQ_SUMMARY_HEADERS[13];
    workSheet.getCell('Z4').value = SHQ_SUMMARY_HEADERS[14];

    let shqSummaryHeaderRow = workSheet.getRow(4);
    shqSummaryHeaderRow.eachCell((cell) => {
      cell.style = TABLE_HEADERS;
    });

    workSheet.mergeCells('A5:A6');
    workSheet.mergeCells('B5:B6');
    workSheet.mergeCells('E5:E6');

    workSheet.getCell('A5').value = 'SHQ - SLA';
    workSheet.getCell('B5').value = 'SHQ Core Device';
    workSheet.getCell('C5').value = timeSpan.replace(/Time Span: /, '');
    workSheet.getCell('E5').value = '22';

    let F5 = workSheet.getCell('F5');
    F5.value = VALUES.PERCENT;
    F5.style = PERCENT_STYLE;

    let G5 = workSheet.getCell('G5');
    G5.value = VALUES.MINUTES;
    G5.style = MINUTE_STYLE;

    let H5 = workSheet.getCell('H5');
    H5.value = VALUES.PERCENT;
    H5.style = PERCENT_STYLE;

    let I5 = workSheet.getCell('I5');
    I5.value = VALUES.MINUTES;
    I5.style = MINUTE_STYLE;

    let J5 = workSheet.getCell('J5');
    J5.value = VALUES.PERCENT;
    J5.style = PERCENT_STYLE;

    let K5 = workSheet.getCell('K5');
    K5.value = VALUES.MINUTES;
    K5.style = MINUTE_STYLE;

    let L5 = workSheet.getCell('L5');
    L5.value = VALUES.PERCENT;
    L5.style = PERCENT_STYLE;

    let M5 = workSheet.getCell('M5');
    M5.value = VALUES.MINUTES;
    M5.style = MINUTE_STYLE;

    let N5 = workSheet.getCell('N5');
    N5.value = VALUES.PERCENT;
    N5.style = PERCENT_STYLE;

    let O5 = workSheet.getCell('O5');
    O5.value = VALUES.MINUTES;
    O5.style = MINUTE_STYLE;

    let P5 = workSheet.getCell('P5');
    P5.value = VALUES.PERCENT;
    P5.style = PERCENT_STYLE;

    let Q5 = workSheet.getCell('Q5');
    Q5.value = VALUES.MINUTES;
    Q5.style = MINUTE_STYLE;

    let R5 = workSheet.getCell('R5');
    R5.value = VALUES.PERCENT;
    R5.style = PERCENT_STYLE;

    let S5 = workSheet.getCell('S5');
    S5.value = VALUES.MINUTES;
    S5.style = MINUTE_STYLE;

    let T5 = workSheet.getCell('T5');
    T5.value = VALUES.PERCENT;
    T5.style = PERCENT_STYLE;

    let U5 = workSheet.getCell('U5');
    U5.value = VALUES.MINUTES;
    U5.style = MINUTE_STYLE;

    let V5 = workSheet.getCell('V5');
    V5.value = VALUES.PERCENT;
    V5.style = PERCENT_STYLE;

    let W5 = workSheet.getCell('W5');
    W5.value = VALUES.MINUTES;
    W5.style = MINUTE_STYLE;

    let X5 = workSheet.getCell('X5');
    X5.value = VALUES.PERCENT;
    X5.style = PERCENT_STYLE;

    let Y5 = workSheet.getCell('Y5');
    Y5.value = VALUES.MINUTES;
    Y5.style = MINUTE_STYLE;

    let Z5 = workSheet.getCell('Z5');
    Z5.value = VALUES.PERCENT;
    Z5.style = PERCENT_STYLE;

    let AA5 = workSheet.getCell('AA5');
    AA5.value = VALUES.MINUTES;
    AA5.style = MINUTE_STYLE;

    workSheet.getCell('F6').value = shqSlaSummary.up_percent;
    workSheet.getCell('G6').value = shqSlaSummary.up_minutes;
    workSheet.getCell('H6').value =
      shqSlaSummary.total_down_exclusive_of_sla_exclusion_percent;
    workSheet.getCell('I6').value =
      shqSlaSummary.total_down_exclusive_of_sla_exclusion_minute;
    workSheet.getCell('J6').value = shqSlaSummary.power_down_percent;
    workSheet.getCell('K6').value = shqSlaSummary.power_dowm_minute;
    workSheet.getCell('L6').value = shqSlaSummary.fibre_down_percent;
    workSheet.getCell('M6').value = shqSlaSummary.fiber_down_minute;
    workSheet.getCell('N6').value = shqSlaSummary.equipment_down_percent;
    workSheet.getCell('O6').value = shqSlaSummary.equipment_down_minute;
    workSheet.getCell('P6').value = shqSlaSummary.hrt_down_percent;
    workSheet.getCell('Q6').value = shqSlaSummary.hrt_down_minute;
    workSheet.getCell('R6').value = shqSlaSummary.dcn_down_percent;
    workSheet.getCell('S6').value = shqSlaSummary.dcn_down_minute;
    workSheet.getCell('T6').value = shqSlaSummary.planned_maintenance_percent;
    workSheet.getCell('U6').value = shqSlaSummary.planned_maintenance_minute;
    workSheet.getCell('V6').value = shqSlaSummary.unknown_downtime_in_percent;
    workSheet.getCell('W6').value = shqSlaSummary.unknown_downtime_in_minutes;
    workSheet.getCell('X6').value = shqSlaSummary.total_sla_exclusion_percent;
    workSheet.getCell('Y6').value = shqSlaSummary.total_sla_exclusion_minute;
    workSheet.getCell('Z6').value = shqSlaSummary.total_up_percent;
    workSheet.getCell('AA6').value = shqSlaSummary.total_up_minute;

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

    workSheet.addRow('');

    workSheet.mergeCells('A8:B8:C8');
    let A8 = workSheet.getCell('A8');
    A8.value = 'SHQ - SLA Device Level (%) & (Min)';
    A8.style = TABLE_HEADING;
    workSheet.getCell('B8').style = TABLE_HEADING;

    workSheet.mergeCells('A9:A10');
    workSheet.mergeCells('B9:B10');
    workSheet.mergeCells('C9:C10');
    workSheet.mergeCells('D9:D10');
    workSheet.mergeCells('E9:E10');

    workSheet.mergeCells('F9:G9');
    workSheet.mergeCells('H9:I9');
    workSheet.mergeCells('J9:K9');
    workSheet.mergeCells('L9:M9');
    workSheet.mergeCells('N9:O9');
    workSheet.mergeCells('P9:Q9');
    workSheet.mergeCells('R9:S9');
    workSheet.mergeCells('T9:U9');
    workSheet.mergeCells('V9:W9');
    workSheet.mergeCells('X9:Y9');
    workSheet.mergeCells('Z9:AA9');
    workSheet.mergeCells('AB9:AC9');

    workSheet.getCell('A9').value = SHQ_DEVICE_LEVEL_HEADERS[0];
    workSheet.getCell('B9').value = SHQ_DEVICE_LEVEL_HEADERS[1];
    workSheet.getCell('C9').value = SHQ_DEVICE_LEVEL_HEADERS[2];
    workSheet.getCell('D9').value = SHQ_DEVICE_LEVEL_HEADERS[3];
    workSheet.getCell('E9').value = SHQ_DEVICE_LEVEL_HEADERS[4];
    workSheet.getCell('F9').value = SHQ_DEVICE_LEVEL_HEADERS[5];
    workSheet.getCell('H9').value = SHQ_DEVICE_LEVEL_HEADERS[6];
    workSheet.getCell('J9').value = SHQ_DEVICE_LEVEL_HEADERS[7];
    workSheet.getCell('L9').value = SHQ_DEVICE_LEVEL_HEADERS[8];
    workSheet.getCell('N9').value = SHQ_DEVICE_LEVEL_HEADERS[9];
    workSheet.getCell('P9').value = SHQ_DEVICE_LEVEL_HEADERS[10];
    workSheet.getCell('R9').value = SHQ_DEVICE_LEVEL_HEADERS[11];
    workSheet.getCell('T9').value = SHQ_DEVICE_LEVEL_HEADERS[12];
    workSheet.getCell('V9').value = SHQ_DEVICE_LEVEL_HEADERS[13];
    workSheet.getCell('X9').value = SHQ_DEVICE_LEVEL_HEADERS[14];
    workSheet.getCell('Z9').value = SHQ_DEVICE_LEVEL_HEADERS[15];
    workSheet.getCell('AB9').value = SHQ_DEVICE_LEVEL_HEADERS[16];

    let finalReportHeader = workSheet.getRow(9);

    finalReportHeader.eachCell((cell) => {
      cell.style = TABLE_HEADERS;
    });

    let F10 = workSheet.getCell('F10');
    F10.value = VALUES.PERCENT;
    F10.style = PERCENT_STYLE;

    let G10 = workSheet.getCell('G10');
    G10.value = VALUES.MINUTES;
    G10.style = MINUTE_STYLE;

    let H10 = workSheet.getCell('H10');
    H10.value = VALUES.PERCENT;
    H10.style = PERCENT_STYLE;

    let I10 = workSheet.getCell('I10');
    I10.value = VALUES.MINUTES;
    I10.style = MINUTE_STYLE;

    let J10 = workSheet.getCell('J10');
    J10.value = VALUES.PERCENT;
    J10.style = PERCENT_STYLE;

    let K10 = workSheet.getCell('K10');
    K10.value = VALUES.MINUTES;
    K10.style = MINUTE_STYLE;

    let L10 = workSheet.getCell('L10');
    L10.value = VALUES.PERCENT;
    L10.style = PERCENT_STYLE;

    let M10 = workSheet.getCell('M10');
    M10.value = VALUES.MINUTES;
    M10.style = MINUTE_STYLE;

    let N10 = workSheet.getCell('N10');
    N10.value = VALUES.PERCENT;
    N10.style = PERCENT_STYLE;

    let O10 = workSheet.getCell('O10');
    O10.value = VALUES.MINUTES;
    O10.style = MINUTE_STYLE;

    let P10 = workSheet.getCell('P10');
    P10.value = VALUES.PERCENT;
    P10.style = PERCENT_STYLE;

    let Q10 = workSheet.getCell('Q10');
    Q10.value = VALUES.MINUTES;
    Q10.style = MINUTE_STYLE;

    let R10 = workSheet.getCell('R10');
    R10.value = VALUES.PERCENT;
    R10.style = PERCENT_STYLE;

    let S10 = workSheet.getCell('S10');
    S10.value = VALUES.MINUTES;
    S10.style = MINUTE_STYLE;

    let T10 = workSheet.getCell('T10');
    T10.value = VALUES.PERCENT;
    T10.style = PERCENT_STYLE;

    let U10 = workSheet.getCell('U10');
    U10.value = VALUES.MINUTES;
    U10.style = MINUTE_STYLE;

    let V10 = workSheet.getCell('V10');
    V10.value = VALUES.PERCENT;
    V10.style = PERCENT_STYLE;

    let W10 = workSheet.getCell('W10');
    W10.value = VALUES.MINUTES;
    W10.style = MINUTE_STYLE;

    let X10 = workSheet.getCell('X10');
    X10.value = VALUES.PERCENT;
    X10.style = PERCENT_STYLE;

    let Y10 = workSheet.getCell('Y10');
    Y10.value = VALUES.MINUTES;
    Y10.style = MINUTE_STYLE;

    let Z10 = workSheet.getCell('Z10');
    Z10.value = VALUES.PERCENT;
    Z10.style = PERCENT_STYLE;

    let AA10 = workSheet.getCell('AA10');
    AA10.value = VALUES.MINUTES;
    AA10.style = MINUTE_STYLE;

    let AB10 = workSheet.getCell('AB10');
    AB10.value = VALUES.PERCENT;
    AB10.style = PERCENT_STYLE;

    let AC10 = workSheet.getCell('AC10');
    AC10.value = VALUES.MINUTES;
    AC10.style = MINUTE_STYLE;

    let row10 = workSheet.getRow(10);
    row10.eachCell((cell) => {
      cell.border = BORDER_STYLE;
      cell.font = { bold: true };
      cell.alignment = { horizontal: 'center' };
    });

    manipulatedShqNmsData.forEach((nmsData: ManipulatedShqNmsData) => {
      let reportType = 'SHQ';
      let tag = 'SHQ Core Device';
      let hostName: string = this.getHostName(nmsData.monitor);
      let ipAddress: string = nmsData.ip_address;
      let deviceType: string = nmsData.type;
      let upPercent: number = nmsData.up_percent;
      let upMinute: number = nmsData.total_uptime_in_minutes;
      let totalDownExclusionOfSlaExclusionInPercent: number =
        upPercent === 100
          ? 0
          : nmsData.total_time_exclusive_of_sla_exclusions_in_percent;
      let totalDownExclusivceOfSlaExclusionInMinutes: number =
        upPercent === 100
          ? 0
          : nmsData.total_time_exclusive_of_sla_exclusions_in_min;
      let powerDownInPercent: number =
        upPercent === 100 ? 0 : nmsData.power_downtime_in_percent;
      let powerDownInMinute: number =
        upPercent === 100 ? 0 : nmsData.power_downtime_in_minutes;
      let fiberDownPercent: number = upPercent == 100 ? 0 : 0;
      let fiberDownMinutes: number = upPercent == 100 ? 0 : 0;
      let equipmentDownPercent: number = upPercent == 100 ? 0 : 0;
      let equipmentDownMinutes: number = upPercent == 100 ? 0 : 0;
      let hrtDownPercent: number = upPercent == 100 ? 0 : 0;
      let hrtDownMinutes: number = upPercent == 100 ? 0 : 0;
      let dcnDownPercent: number =
        upPercent == 100 ? 0 : nmsData.dcn_downtime_in_percent;
      let dcnDownMinutes =
        upPercent == 100 ? 0 : nmsData.dcn_downtime_in_minutes;
      let plannedMaintanancePercent: number = upPercent == 100 ? 0 : 0;
      let plannedMaintananceMinutes: number = upPercent == 100 ? 0 : 0;
      let unKnownDownPercent: number =
        upPercent == 100 ? 0 : nmsData.unknown_downtime_in_percent;
      let unKnownDownMinutes: number =
        upPercent == 100 ? 0 : nmsData.unknown_downtime_in_minutes;
      let totalSlaExclusionInPercent: number =
        upPercent == 100
          ? 0
          : nmsData.power_downtime_in_percent + nmsData.dcn_downtime_in_percent;
      let totalSlaExclusionInMinute: number =
        upPercent == 100
          ? 0
          : nmsData.power_downtime_in_minutes + nmsData.dcn_downtime_in_minutes;
      let pollingTimeInPercent: number =
        upPercent == 100 ? 0 : nmsData.polling_time_in_percent;
      let pollingTimeInMinute: number =
        upPercent == 100 ? 0 : nmsData.polling_time_in_minutes;
      let totalUpInPercent: number =
        upPercent == 100 ? 100 : upPercent + nmsData.down_percent;
      let totalUpInMinutes: number =
        upMinute + nmsData.total_downtime_in_minutes;

      const ShqDeviceLevelRowValues = workSheet.addRow([
        reportType,
        tag,
        hostName,
        ipAddress,
        deviceType,
        upPercent.toFixed(2),
        upMinute.toFixed(2),
        totalDownExclusionOfSlaExclusionInPercent.toFixed(2),
        totalDownExclusivceOfSlaExclusionInMinutes.toFixed(2),
        powerDownInPercent.toFixed(2),
        powerDownInMinute.toFixed(2),
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
        totalSlaExclusionInPercent.toFixed(2),
        totalSlaExclusionInMinute.toFixed(2),
        pollingTimeInPercent.toFixed(2),
        pollingTimeInMinute.toFixed(2),
        totalUpInPercent.toFixed(2),
        totalUpInMinutes.toFixed(2),
      ]);

      ShqDeviceLevelRowValues.eachCell((cell) => {
        cell.border = BORDER_STYLE;
        cell.alignment = { horizontal: 'left' };
      });
    });

    // Generating Sheet 2
    const ttCorelationWorkSheet = workbook.addWorksheet('SHQ-TT co-relation');
    ttCorelationWorkSheet.columns = SHQ_TT_CO_RELATION_COLUMNS;
    ttCorelationWorkSheet
      .addRow(SHQ_TT_CO_RELATION_HEADERS)
      .eachCell((cell) => {
        cell.style = TABLE_HEADERS;
      });
    this.ttCorelation.forEach(
      (ttCorelationData: TTCorelation, index: number) => {
        let [shqData] = manipulatedShqNmsData.filter(
          (nmsData: ManipulatedShqNmsData) => {
            return nmsData.ip_address == ttCorelationData.ip;
          }
        );
        ttCorelationWorkSheet
          .addRow([
            index + 1,
            this.getHostName(shqData.monitor),
            ttCorelationData.ip,
            shqData.type,
            ttCorelationData.powerIssueTT.toString().split(',').join(', '),
            ttCorelationData.linkIssueTT.toString().split(',').join(', '),
            ttCorelationData.otherTT.toString().split(',').join(', '),
          ])
          .eachCell({ includeEmpty: true }, (cell) => {
            cell.border = BORDER_STYLE;
            cell.alignment = { horizontal: 'left' };
          });
      }
    );
  }
}
