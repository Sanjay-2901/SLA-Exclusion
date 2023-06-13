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
  RFO_CATEGORIZATION,
  SEVERITY_CRITICAL,
  SEVERITY_WARNING,
  SHEET_HEADING,
  SHQ_SLQ_FINAL_REPORT_COLUMNS,
  SHQ_SUMMARY_HEADERS,
  TABLE_HEADING,
  VMWAREDEVICE,
} from '../constants/constants';
import * as moment from 'moment';
import * as lodash from 'lodash';
import * as ExcelJS from 'exceljs';
import { RFOCategorizedTimeInMinutes } from '../block-component/block-component.model';

@Injectable({
  providedIn: 'root',
})
export class ShqService {
  constructor() {}

  shqNMSDatawithoutVmwareDevices(
    AllShqDevicesArray: ShqNMSData[]
  ): ShqNMSData[] {
    return AllShqDevicesArray.filter((nmsData) => {
      return nmsData.type.trim() !== VMWAREDEVICE;
    });
  }

  CalucateTimeInMinutes(timePeriod: string) {
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

  calculateAlertDownTimeInMinutes(
    ipAddress: string,
    shqAlertData: ShqAlertData[]
  ) {
    let filteredAlertData = shqAlertData.filter((alert: ShqAlertData) => {
      return (
        alert.ip_address.trim() == ipAddress &&
        alert.severity.trim() == SEVERITY_CRITICAL &&
        alert.message.trim() == ALERT_DOWN_MESSAGE
      );
    });

    let alertDownTimeInMinutes: number = 0;
    filteredAlertData.forEach((filteredAlertData: ShqAlertData) => {
      alertDownTimeInMinutes += filteredAlertData.total_duration_in_minutes;
    });
    return alertDownTimeInMinutes;
  }

  categorizeRFO(
    ipAddress: string,
    shqAlertData: ShqAlertData[],
    shqTTData: ShqTTData[]
  ) {
    let totalPowerDownTimeInMinutes = 0;
    let totalDCNDownTimeInMinutes = 0;

    let powerDownArray: ShqAlertData[] = [];
    let DCNDownArray: ShqAlertData[] = [];
    let criticalAlertAndTTDataTimeMismatch: ShqAlertData[] = [];

    const filteredCriticalAlertData = shqAlertData.filter(
      (alertData: ShqAlertData) => {
        return (
          alertData.ip_address.trim() == ipAddress &&
          alertData.severity.trim() == SEVERITY_CRITICAL &&
          alertData.message.trim() == ALERT_DOWN_MESSAGE
        );
      }
    );

    const filteredWarningAlertData = shqAlertData.filter(
      (alertData: ShqAlertData) => {
        return (
          alertData.ip_address.trim() == ipAddress &&
          alertData.severity.trim() == SEVERITY_WARNING &&
          alertData.message.trim().includes('reboot')
        );
      }
    );

    const filteredTTData = shqTTData.filter((ttData: ShqTTData) => {
      return ttData.ip == ipAddress;
    });

    filteredCriticalAlertData.forEach((alertCriticalData: ShqAlertData) => {
      filteredTTData.forEach((ttData: ShqTTData) => {
        if (
          moment(alertCriticalData.last_poll_time).isSame(
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
        (alertCriticalData: ShqAlertData) => {
          filteredWarningAlertData.forEach((alertWarningData: ShqAlertData) => {
            if (
              moment(alertCriticalData.duration_time).isSame(
                alertWarningData.last_poll_time,
                'minute'
              )
            ) {
              powerDownArray.push(alertCriticalData);
            }
          });

          if (!lodash.some(powerDownArray, alertCriticalData)) {
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
    };

    return rfoCategorizedTimeInMinutes;
  }

  calculateCumulativeValue(value: number): number {
    return parseInt((value / 22).toFixed(2));
  }

  calculateShqSlaSummary(
    manipulatedNMSData: ManipulatedShqNmsData[]
  ): ShqSlaSummary {
    let upPercent = 0;
    let upMinutes = 0;
    let totalDownExclusiveOfSlaExclusionInPercent = 0;
    let totalDownExclusiveOfSlaExclusionInMinute = 0;
    let powerDownPercent = 0;
    let powerDownMinutes = 0;
    let dcnDownPercent = 0;
    let dcnDownMinutes = 0;
    let plannedMaintenance = 0;
    let dcnAndPowerDownPercent = 0;
    let dcnAndPowerDownMinutes = 0;
    let totalSlaExclusionPercent = 0;
    let totalSlaExclusionMinute = 0;
    let pollingTimePercent = 0;
    let pollingTimeMinutes = 0;
    let totalUpWithExclusionPercent = 0;
    let totalUpWithExclusionMinute = 0;

    manipulatedNMSData.forEach((nmsData: ManipulatedShqNmsData) => {
      upPercent += nmsData.up_percent;
      upMinutes += nmsData.total_uptime_in_minutes;
      totalDownExclusiveOfSlaExclusionInMinute +=
        nmsData.total_downtime_in_minutes;
      powerDownPercent += nmsData.power_downtime_in_percent;
      powerDownMinutes += nmsData.power_downtime_in_minutes;
      dcnDownPercent += nmsData.dcn_downtime_in_percent;
      dcnDownMinutes += nmsData.dcn_downtime_in_minutes;
      dcnAndPowerDownPercent +=
        nmsData.power_downtime_in_percent + nmsData.dcn_downtime_in_percent;
      dcnAndPowerDownMinutes +=
        nmsData.power_downtime_in_minutes + nmsData.dcn_downtime_in_minutes;
      totalSlaExclusionPercent +=
        nmsData.power_downtime_in_percent + nmsData.dcn_downtime_in_percent;
      totalSlaExclusionMinute +=
        nmsData.power_downtime_in_minutes + nmsData.dcn_downtime_in_minutes;
      pollingTimePercent +=
        nmsData.down_percent -
        (nmsData.power_downtime_in_percent + nmsData.dcn_downtime_in_percent);
      pollingTimeMinutes +=
        nmsData.total_downtime_in_minutes -
        (nmsData.power_downtime_in_minutes + nmsData.dcn_downtime_in_minutes);
    });

    return {
      report_type: 'SHQ-SLA',
      tag: 'SHQ Core Device',
      time_span: '',
      no_of_shq_devices: 22,
      up_percent: this.calculateCumulativeValue(upPercent),
      up_minutes: this.calculateCumulativeValue(upMinutes),
      total_down_exclusive_of_sla_exclusion_percent:
        100 - this.calculateCumulativeValue(upPercent),
      total_down_exclusive_of_sla_exclusion_minute:
        this.calculateCumulativeValue(totalDownExclusiveOfSlaExclusionInMinute),
      power_down_percent: this.calculateCumulativeValue(powerDownPercent),
      power_dowm_minute: this.calculateCumulativeValue(powerDownMinutes),
      fibre_down_percent: 0,
      fiber_down_minute: 0,
      equipment_down_percent: 0,
      equipment_down_minute: 0,
      hrt_down_percent: 0,
      hrt_down_minute: 0,
      dcn_down_percent: this.calculateCumulativeValue(dcnDownPercent),
      dcn_down_minute: this.calculateCumulativeValue(dcnDownMinutes),
      planned_maintenance_percent: 0,
      planned_maintenance_minute: 0,
      total_sla_exclusion_percent: this.calculateCumulativeValue(
        dcnAndPowerDownPercent
      ),
      total_sla_exclusion_minute: this.calculateCumulativeValue(
        dcnAndPowerDownMinutes
      ),
      total_up_percent: this.calculateCumulativeValue(
        upPercent + pollingTimePercent + totalSlaExclusionPercent
      ),
      total_up_minute: this.calculateCumulativeValue(
        upMinutes + pollingTimeMinutes + totalSlaExclusionMinute
      ),
    };
  }

  FrameShqFinalSlaReportWorkbook(workSheet: ExcelJS.Worksheet): void {
    workSheet.columns = SHQ_SLQ_FINAL_REPORT_COLUMNS;

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

    workSheet.mergeCells('A3:c3');
    let cellA3 = workSheet.getCell('A3');
    cellA3.value = 'SHQ - SLA Summary (%) & (Min)';
    cellA3.style = TABLE_HEADING;
    workSheet.getCell('B3').style = TABLE_HEADING;

    workSheet.mergeCells('C4:D4');
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

    let shqSummaryHeaderRow = workSheet.getRow(4);
    shqSummaryHeaderRow.eachCell((cell) => {
      // cell.style = TABLE_HEADERS;
    });
  }

  downloadFinalReport(buffer: ArrayBuffer, fileName: string) {
    const data = new Blob([buffer], {
      type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
    });
    const link = document.createElement('a');
    link.href = window.URL.createObjectURL(data);
    link.download = fileName + '.xlsx';
    link.click();
  }
}
