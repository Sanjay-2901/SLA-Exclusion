import { Injectable } from '@angular/core';
import * as ExcelJS from 'exceljs';
import {
  BlockSLASummary,
  ManipulatedNMSData,
  TTCorelation,
} from './block-component.model';
import {
  BLOCK_ALERT_REPORT_HEADERS,
  BLOCK_DEVICE_DETAILS,
  BLOCK_INPUT_FILE_NAMES,
  BLOCK_SLA_FINAL_REPORT_COLUMN_WIDTHS,
  BLOCK_SLA_REPORT_HEADERS,
  BLOCK_TT_CO_RELATION_COLUMNS_WIDTHS,
  BLOCK_TT_CO_RELATION_HEADERS,
  BORDER_STYLE,
  BlockDeviceDetail,
  BlockSLAFinalReportHeaders,
  BlockSLASummarytHeaders,
  MINUTE_STYLE,
  PERCENT_STYLE,
  SHEET_HEADING,
  TABLE_HEADERS,
  TABLE_HEADING,
  TT_REPORT_HEADERS,
  UNKNOWN_COLUMN_COLOR,
  VALUES,
} from '../constants/constants';
import { SharedService } from '../shared/shared.service';

@Injectable({
  providedIn: 'root',
})
export class BlockService {
  constructor(private sharedService: SharedService) {}

  calculateBlockSlaSummary(
    manipulatedNMSData: ManipulatedNMSData[],
    timeSpan: string
  ): BlockSLASummary {
    let upPercent = 0;
    let upMinutes = 0;
    let totalDownMinutes = 0;
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
    let pollingTimePercent = 0;
    let pollingTimeMinutes = 0;
    const blockCount = manipulatedNMSData.length;
    let gpCount = 0;

    manipulatedNMSData.forEach((row: ManipulatedNMSData) => {
      let block_device_details = BLOCK_DEVICE_DETAILS.filter(
        (device: BlockDeviceDetail) => {
          return row.ip_address == device.ip_address;
        }
      )[0];
      gpCount += block_device_details.no_of_gp_in_block;
    });

    manipulatedNMSData.forEach((nmsData: ManipulatedNMSData) => {
      upPercent += nmsData.up_percent;
      upMinutes += nmsData.total_uptime_in_minutes;
      totalDownMinutes += nmsData.total_downtime_in_minutes;
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
        hrtDownPercent +
        nmsData.maintenance_percent +
        nmsData.polling_time_in_percent;
      cumulativeRfoDownInMinutes +=
        nmsData.power_downtime_in_minutes +
        nmsData.dcn_downtime_in_minutes +
        hrtDownMinute +
        nmsData.planned_maintenance_in_minutes +
        nmsData.polling_time_in_minutes;
      pollingTimePercent += nmsData.polling_time_in_percent;
      pollingTimeMinutes += nmsData.polling_time_in_minutes;
    });

    return {
      report_type: 'BLOCK-SLA',
      tag: 'Q&M Block',
      time_span: timeSpan.replace(/Time Span: /, ''),
      no_of_blocks: blockCount,
      no_of_gps: gpCount,
      up_percent: this.sharedService.CaloculateSummaryPercentageValue(
        blockCount,
        upPercent
      ),
      up_minutes: upMinutes,
      total_down_percent:
        blockCount !== 0
          ? 100 - upPercent / blockCount > 100
            ? 100
            : 100 - +(upPercent / blockCount)
          : 0,
      total_down_minutes: totalDownMinutes,
      power_down_percent: this.sharedService.CaloculateSummaryPercentageValue(
        blockCount,
        powerDownPercent
      ),
      power_down_minutes: powerDownMinutes,
      fibre_down_percent: this.sharedService.CaloculateSummaryPercentageValue(
        blockCount,
        fiberDownPercent
      ),
      fibre_down_minutes: fiberDownMinute,
      equipment_down_percent:
        this.sharedService.CaloculateSummaryPercentageValue(
          blockCount,
          equipmentDownPercent
        ),
      equipment_down_minutes: equipmentDownMinute,
      hrt_down_percent: this.sharedService.CaloculateSummaryPercentageValue(
        blockCount,
        hrtDownPercent
      ),
      hrt_down_minutes: hrtDownMinute,
      dcn_down_percent: this.sharedService.CaloculateSummaryPercentageValue(
        blockCount,
        dcnDownPercent
      ),
      dcn_down_minutes: dcnDownMinutes,
      planned_maintenance_percent:
        this.sharedService.CaloculateSummaryPercentageValue(
          blockCount,
          plannedMaintenancePercent
        ),
      planned_maintenance_minutes: plannedMaintenanceMinutes,
      unknown_downtime_in_percent:
        this.sharedService.CaloculateSummaryPercentageValue(
          blockCount,
          unKnownDownPercent
        ),
      unknown_downtime_in_minutes: unKnownDownMinutes,
      total_sla_exclusion_percent:
        this.sharedService.CaloculateSummaryPercentageValue(
          blockCount,
          cumulativeRfoDownInPercent - pollingTimePercent
        ),
      total_sla_exclusion_minutes:
        cumulativeRfoDownInMinutes - pollingTimeMinutes,
      total_up_percent_exclusion:
        this.sharedService.CaloculateSummaryPercentageValue(
          blockCount,
          upPercent + cumulativeRfoDownInPercent
        ),
      total_up_minutes_exclusion: upMinutes + cumulativeRfoDownInMinutes,
    };
  }

  generateFinalBlockReport(
    workbook: ExcelJS.Workbook,
    workSheet: ExcelJS.Worksheet,
    blockSlaSummary: BlockSLASummary,
    blockSLASummaryWithAlerts: BlockSLASummary,
    blockSLASummaryWithoutAlerts: BlockSLASummary,
    manipulatedNmsData: ManipulatedNMSData[],
    ttCorelation: TTCorelation[]
  ): void {
    workSheet.columns = BLOCK_SLA_FINAL_REPORT_COLUMN_WIDTHS;

    workSheet.mergeCells('A1:B1');
    let cellA1 = workSheet.getCell('A1');
    cellA1.value = '1. Daily Network availability report';
    cellA1.style = SHEET_HEADING;

    workSheet.mergeCells('C1:D1');
    let cellC1 = workSheet.getCell('C1');
    cellC1.value = 'Report-Frequency: ';
    cellC1.style = {
      font: { bold: true },
      alignment: { horizontal: 'center' },
    };

    workSheet.mergeCells('A3:B3');
    let cellA3 = workSheet.getCell('A3');
    cellA3.value = 'Block - SLA Summary (%) & (Min)';
    cellA3.style = TABLE_HEADING;
    workSheet.getCell('B3').style = TABLE_HEADING;

    workSheet.mergeCells('C4:G4');
    workSheet.mergeCells('I4:J4');
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

    workSheet.getCell('A4').value = BlockSLASummarytHeaders[0];
    workSheet.getCell('B4').value = BlockSLASummarytHeaders[1];
    workSheet.getCell('C4').value = BlockSLASummarytHeaders[2];
    workSheet.getCell('H4').value = BlockSLASummarytHeaders[3];
    workSheet.getCell('I4').value = BlockSLASummarytHeaders[4];
    workSheet.getCell('K4').value = BlockSLASummarytHeaders[5];
    workSheet.getCell('M4').value = BlockSLASummarytHeaders[6];
    workSheet.getCell('O4').value = BlockSLASummarytHeaders[7];
    workSheet.getCell('Q4').value = BlockSLASummarytHeaders[8];
    workSheet.getCell('S4').value = BlockSLASummarytHeaders[9];
    workSheet.getCell('U4').value = BlockSLASummarytHeaders[10];
    workSheet.getCell('W4').value = BlockSLASummarytHeaders[11];
    workSheet.getCell('Y4').value = BlockSLASummarytHeaders[12];
    workSheet.getCell('AA4').value = BlockSLASummarytHeaders[13];
    workSheet.getCell('AC4').value = BlockSLASummarytHeaders[14];
    workSheet.getCell('AE4').value = BlockSLASummarytHeaders[15];

    let blockSummaryHeadersRow = workSheet.getRow(4);
    blockSummaryHeadersRow.eachCell((cell) => {
      cell.style = TABLE_HEADERS;
    });

    // Common columns in Block Summary Section

    workSheet.mergeCells('A5:A6');
    workSheet.mergeCells('B5:B6');
    workSheet.mergeCells('C5:G6');

    workSheet.getCell('A5').value = blockSlaSummary.report_type;
    workSheet.getCell('B5').value = blockSlaSummary.tag;
    workSheet.getCell('C5').value = blockSlaSummary.time_span;

    // Block SLA Summary section Framing

    workSheet.mergeCells('H5:H6');
    workSheet.mergeCells('I5:J6');

    workSheet.getCell('H5').value = blockSlaSummary.no_of_blocks;
    workSheet.getCell('I5').value = blockSlaSummary.no_of_gps;

    let k5 = workSheet.getCell('K5');
    k5.value = VALUES.PERCENT;
    k5.style = PERCENT_STYLE;
    let l5 = workSheet.getCell('L5');
    l5.value = VALUES.MINUTES;
    l5.style = MINUTE_STYLE;
    let m5 = workSheet.getCell('M5');
    m5.value = VALUES.PERCENT;
    m5.style = PERCENT_STYLE;
    let n5 = workSheet.getCell('N5');
    n5.value = VALUES.MINUTES;
    n5.style = MINUTE_STYLE;
    let o5 = workSheet.getCell('O5');
    o5.value = VALUES.PERCENT;
    o5.style = PERCENT_STYLE;
    let p5 = workSheet.getCell('P5');
    p5.value = VALUES.MINUTES;
    p5.style = MINUTE_STYLE;
    let q5 = workSheet.getCell('Q5');
    q5.value = VALUES.PERCENT;
    q5.style = PERCENT_STYLE;
    let r5 = workSheet.getCell('R5');
    r5.value = VALUES.MINUTES;
    r5.style = MINUTE_STYLE;
    let s5 = workSheet.getCell('S5');
    s5.value = VALUES.PERCENT;
    s5.style = PERCENT_STYLE;
    let t5 = workSheet.getCell('T5');
    t5.value = VALUES.MINUTES;
    t5.style = MINUTE_STYLE;
    let u5 = workSheet.getCell('U5');
    u5.value = VALUES.PERCENT;
    u5.style = PERCENT_STYLE;
    let v5 = workSheet.getCell('V5');
    v5.value = VALUES.MINUTES;
    v5.style = MINUTE_STYLE;
    let w5 = workSheet.getCell('W5');
    w5.value = VALUES.PERCENT;
    w5.style = PERCENT_STYLE;
    let x5 = workSheet.getCell('X5');
    x5.value = VALUES.MINUTES;
    x5.style = MINUTE_STYLE;
    let y5 = workSheet.getCell('Y5');
    y5.value = VALUES.PERCENT;
    y5.style = PERCENT_STYLE;
    let z5 = workSheet.getCell('Z5');
    z5.value = VALUES.MINUTES;
    z5.style = MINUTE_STYLE;
    let aa5 = workSheet.getCell('AA5');
    aa5.value = VALUES.PERCENT;
    aa5.style = PERCENT_STYLE;
    let ab5 = workSheet.getCell('AB5');
    ab5.value = VALUES.MINUTES;
    ab5.style = MINUTE_STYLE;
    let ac5 = workSheet.getCell('AC5');
    ac5.value = VALUES.PERCENT;
    ac5.style = PERCENT_STYLE;
    let ad5 = workSheet.getCell('AD5');
    ad5.value = VALUES.MINUTES;
    ad5.style = MINUTE_STYLE;
    let ae5 = workSheet.getCell('AE5');
    ae5.value = VALUES.PERCENT;
    ae5.style = PERCENT_STYLE;
    let af5 = workSheet.getCell('AF5');
    af5.value = VALUES.MINUTES;
    af5.style = MINUTE_STYLE;

    workSheet.getCell('K6').value = blockSlaSummary.up_percent;
    workSheet.getCell('L6').value = blockSlaSummary.up_minutes;
    workSheet.getCell('M6').value = blockSlaSummary.total_down_percent;
    workSheet.getCell('N6').value = blockSlaSummary.total_down_minutes;
    workSheet.getCell('O6').value = blockSlaSummary.power_down_percent;
    workSheet.getCell('P6').value = blockSlaSummary.power_down_minutes;
    workSheet.getCell('Q6').value = blockSlaSummary.fibre_down_percent;
    workSheet.getCell('R6').value = blockSlaSummary.fibre_down_minutes;
    workSheet.getCell('S6').value = blockSlaSummary.equipment_down_percent;
    workSheet.getCell('T6').value = blockSlaSummary.equipment_down_minutes;
    workSheet.getCell('U6').value = blockSlaSummary.hrt_down_percent;
    workSheet.getCell('V6').value = blockSlaSummary.hrt_down_minutes;
    workSheet.getCell('W6').value = blockSlaSummary.dcn_down_percent;
    workSheet.getCell('X6').value = blockSlaSummary.dcn_down_minutes;
    workSheet.getCell('Y6').value = blockSlaSummary.planned_maintenance_percent;
    workSheet.getCell('Z6').value = blockSlaSummary.planned_maintenance_minutes;
    workSheet.getCell('AA6').value =
      blockSlaSummary.unknown_downtime_in_percent;
    workSheet.getCell('AB6').value =
      blockSlaSummary.unknown_downtime_in_minutes;
    workSheet.getCell('AC6').value =
      blockSlaSummary.total_sla_exclusion_percent;
    workSheet.getCell('AD6').value =
      blockSlaSummary.total_sla_exclusion_minutes;
    workSheet.getCell('AE6').value = blockSlaSummary.total_up_percent_exclusion;
    workSheet.getCell('AF6').value = blockSlaSummary.total_up_minutes_exclusion;

    let row5 = workSheet.getRow(5);
    row5.eachCell((cell) => {
      cell.border = BORDER_STYLE;
      cell.alignment = { horizontal: 'center' };
      cell.font = { bold: true };
    });

    let row6 = workSheet.getRow(6);
    row6.eachCell((cell, colNumber: number) => {
      cell.border = BORDER_STYLE;
      cell.alignment = { horizontal: 'center' };
      if (colNumber > 10) {
        cell.numFmt = '0.00';
      }
    });

    // Block SLA Summary With Alerts section Framing

    workSheet.mergeCells('F7:G7');
    let F7 = workSheet.getCell('F7');
    F7.value = 'Blocks with Alerts';

    workSheet.getCell('H7').value = blockSLASummaryWithAlerts.no_of_blocks;

    workSheet.mergeCells('I7:J7');
    let I7 = workSheet.getCell('I7');
    I7.value = blockSLASummaryWithAlerts.no_of_gps;

    workSheet.getCell('K7').value = blockSLASummaryWithAlerts.up_percent;
    workSheet.getCell('L7').value = blockSLASummaryWithAlerts.up_minutes;
    workSheet.getCell('M7').value =
      blockSLASummaryWithAlerts.total_down_percent;
    workSheet.getCell('N7').value =
      blockSLASummaryWithAlerts.total_down_minutes;
    workSheet.getCell('O7').value =
      blockSLASummaryWithAlerts.power_down_percent;
    workSheet.getCell('P7').value =
      blockSLASummaryWithAlerts.power_down_minutes;
    workSheet.getCell('Q7').value =
      blockSLASummaryWithAlerts.fibre_down_percent;
    workSheet.getCell('R7').value =
      blockSLASummaryWithAlerts.fibre_down_minutes;
    workSheet.getCell('S7').value =
      blockSLASummaryWithAlerts.equipment_down_percent;
    workSheet.getCell('T7').value =
      blockSLASummaryWithAlerts.equipment_down_minutes;
    workSheet.getCell('U7').value = blockSLASummaryWithAlerts.hrt_down_percent;
    workSheet.getCell('V7').value = blockSLASummaryWithAlerts.hrt_down_minutes;
    workSheet.getCell('W7').value = blockSLASummaryWithAlerts.dcn_down_percent;
    workSheet.getCell('X7').value = blockSLASummaryWithAlerts.dcn_down_minutes;
    workSheet.getCell('Y7').value =
      blockSLASummaryWithAlerts.planned_maintenance_percent;
    workSheet.getCell('Z7').value =
      blockSLASummaryWithAlerts.planned_maintenance_minutes;
    workSheet.getCell('AA7').value =
      blockSLASummaryWithAlerts.unknown_downtime_in_percent;
    workSheet.getCell('AB7').value =
      blockSLASummaryWithAlerts.unknown_downtime_in_minutes;
    workSheet.getCell('AC7').value =
      blockSLASummaryWithAlerts.total_sla_exclusion_percent;
    workSheet.getCell('AD7').value =
      blockSLASummaryWithAlerts.total_sla_exclusion_minutes;
    workSheet.getCell('AE7').value =
      blockSLASummaryWithAlerts.total_up_percent_exclusion;
    workSheet.getCell('AF7').value =
      blockSLASummaryWithAlerts.total_up_minutes_exclusion;

    let row7 = workSheet.getRow(7);
    row7.eachCell((cell: ExcelJS.Cell, colNumber: number) => {
      cell.style = {
        border: BORDER_STYLE,
        alignment: { horizontal: 'center' },
      };
      if (colNumber > 10) {
        cell.numFmt = '0.00';
      }
    });
    F7.font = { bold: true };

    // Block SLA Summary Without Alerts Section Framing

    workSheet.mergeCells('F8:G8');
    let F8 = workSheet.getCell('F8');
    F8.value = 'Blocks without Alerts';

    workSheet.getCell('H8').value = blockSLASummaryWithoutAlerts.no_of_blocks;

    workSheet.mergeCells('I8:J8');
    let I8 = workSheet.getCell('I8');
    I8.value = blockSLASummaryWithoutAlerts.no_of_gps;

    workSheet.getCell('K8').value = blockSLASummaryWithoutAlerts.up_percent;
    workSheet.getCell('L8').value = blockSLASummaryWithoutAlerts.up_minutes;
    workSheet.getCell('M8').value =
      blockSLASummaryWithoutAlerts.total_down_percent;
    workSheet.getCell('N8').value =
      blockSLASummaryWithoutAlerts.total_down_minutes;
    workSheet.getCell('O8').value =
      blockSLASummaryWithoutAlerts.power_down_percent;
    workSheet.getCell('P8').value =
      blockSLASummaryWithoutAlerts.power_down_minutes;
    workSheet.getCell('Q8').value =
      blockSLASummaryWithoutAlerts.fibre_down_percent;
    workSheet.getCell('R8').value =
      blockSLASummaryWithoutAlerts.fibre_down_minutes;
    workSheet.getCell('S8').value =
      blockSLASummaryWithoutAlerts.equipment_down_percent;
    workSheet.getCell('T8').value =
      blockSLASummaryWithoutAlerts.equipment_down_minutes;
    workSheet.getCell('U8').value =
      blockSLASummaryWithoutAlerts.hrt_down_percent;
    workSheet.getCell('V8').value =
      blockSLASummaryWithoutAlerts.hrt_down_minutes;
    workSheet.getCell('W8').value =
      blockSLASummaryWithoutAlerts.dcn_down_percent;
    workSheet.getCell('X8').value =
      blockSLASummaryWithoutAlerts.dcn_down_minutes;
    workSheet.getCell('Y8').value =
      blockSLASummaryWithoutAlerts.planned_maintenance_percent;
    workSheet.getCell('Z8').value =
      blockSLASummaryWithoutAlerts.planned_maintenance_minutes;
    workSheet.getCell('AA8').value =
      blockSLASummaryWithoutAlerts.unknown_downtime_in_percent;
    workSheet.getCell('AB8').value =
      blockSLASummaryWithoutAlerts.unknown_downtime_in_minutes;
    workSheet.getCell('AC8').value =
      blockSLASummaryWithoutAlerts.total_sla_exclusion_percent;
    workSheet.getCell('AD8').value =
      blockSLASummaryWithoutAlerts.total_sla_exclusion_minutes;
    workSheet.getCell('AE8').value =
      blockSLASummaryWithoutAlerts.total_up_percent_exclusion;
    workSheet.getCell('AF8').value =
      blockSLASummaryWithoutAlerts.total_up_minutes_exclusion;

    let row8 = workSheet.getRow(8);
    row8.eachCell((cell: ExcelJS.Cell, colNumber: number) => {
      cell.style = {
        border: BORDER_STYLE,
        alignment: { horizontal: 'center' },
      };
      if (colNumber > 10) {
        cell.numFmt = '0.00';
      }
    });
    F8.font = { bold: true };

    F7.style = TABLE_HEADERS;
    F8.style = TABLE_HEADERS;

    workSheet.getCell('AA6').style = UNKNOWN_COLUMN_COLOR;
    workSheet.getCell('AB6').style = UNKNOWN_COLUMN_COLOR;
    workSheet.getCell('AA7').style = UNKNOWN_COLUMN_COLOR;
    workSheet.getCell('AB7').style = UNKNOWN_COLUMN_COLOR;
    workSheet.getCell('AA8').style = UNKNOWN_COLUMN_COLOR;
    workSheet.getCell('AB8').style = UNKNOWN_COLUMN_COLOR;

    workSheet.mergeCells('A9:B9');
    let cellA11 = workSheet.getCell('A9');
    cellA11.value = 'Block - SLA Device Level (%) & (Min)';
    cellA11.style = TABLE_HEADING;
    workSheet.getCell('B9').style = TABLE_HEADERS;

    // Block SLA Device Level Section Framing

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

    workSheet.mergeCells('K10:L10');
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

    workSheet.getCell('A10').value = BlockSLAFinalReportHeaders[0];
    workSheet.getCell('B10').value = BlockSLAFinalReportHeaders[1];
    workSheet.getCell('C10').value = BlockSLAFinalReportHeaders[2];
    workSheet.getCell('D10').value = BlockSLAFinalReportHeaders[3];
    workSheet.getCell('E10').value = BlockSLAFinalReportHeaders[4];
    workSheet.getCell('F10').value = BlockSLAFinalReportHeaders[5];
    workSheet.getCell('G10').value = BlockSLAFinalReportHeaders[6];
    workSheet.getCell('H10').value = BlockSLAFinalReportHeaders[7];
    workSheet.getCell('I10').value = BlockSLAFinalReportHeaders[8];
    workSheet.getCell('J10').value = BlockSLAFinalReportHeaders[9];
    workSheet.getCell('K10').value = BlockSLAFinalReportHeaders[10];
    workSheet.getCell('M10').value = BlockSLAFinalReportHeaders[11];
    workSheet.getCell('O10').value = BlockSLAFinalReportHeaders[12];
    workSheet.getCell('Q10').value = BlockSLAFinalReportHeaders[13];
    workSheet.getCell('S10').value = BlockSLAFinalReportHeaders[14];
    workSheet.getCell('U10').value = BlockSLAFinalReportHeaders[15];
    workSheet.getCell('W10').value = BlockSLAFinalReportHeaders[16];
    workSheet.getCell('Y10').value = BlockSLAFinalReportHeaders[17];
    workSheet.getCell('AA10').value = BlockSLAFinalReportHeaders[18];
    workSheet.getCell('AC10').value = BlockSLAFinalReportHeaders[19];
    workSheet.getCell('AE10').value = BlockSLAFinalReportHeaders[20];
    workSheet.getCell('AG10').value = BlockSLAFinalReportHeaders[21];

    let finalReportHeaders = workSheet.getRow(10);

    finalReportHeaders.eachCell((cell) => {
      cell.style = TABLE_HEADERS;
    });

    let k11 = workSheet.getCell('K11');
    k11.value = VALUES.PERCENT;
    k11.style = PERCENT_STYLE;
    let l11 = workSheet.getCell('L11');
    l11.value = VALUES.MINUTES;
    l11.style = MINUTE_STYLE;

    let m11 = workSheet.getCell('M11');
    m11.value = VALUES.PERCENT;
    m11.style = PERCENT_STYLE;
    let n11 = workSheet.getCell('N11');
    n11.value = VALUES.MINUTES;
    n11.style = MINUTE_STYLE;

    let o11 = workSheet.getCell('O11');
    o11.value = VALUES.PERCENT;
    o11.style = PERCENT_STYLE;
    let p11 = workSheet.getCell('P11');
    p11.value = VALUES.MINUTES;
    p11.style = MINUTE_STYLE;

    let q11 = workSheet.getCell('Q11');
    q11.value = VALUES.PERCENT;
    q11.style = PERCENT_STYLE;
    let r11 = workSheet.getCell('R11');
    r11.value = VALUES.MINUTES;
    r11.style = MINUTE_STYLE;

    let s11 = workSheet.getCell('S11');
    s11.value = VALUES.PERCENT;
    s11.style = PERCENT_STYLE;
    let t11 = workSheet.getCell('T11');
    t11.value = VALUES.MINUTES;
    t11.style = MINUTE_STYLE;

    let u11 = workSheet.getCell('U11');
    u11.value = VALUES.PERCENT;
    u11.style = PERCENT_STYLE;
    let v11 = workSheet.getCell('V11');
    v11.value = VALUES.MINUTES;
    v11.style = MINUTE_STYLE;

    let w11 = workSheet.getCell('W11');
    w11.value = VALUES.PERCENT;
    w11.style = PERCENT_STYLE;
    let x11 = workSheet.getCell('X11');
    x11.value = VALUES.MINUTES;
    x11.style = MINUTE_STYLE;

    let y11 = workSheet.getCell('Y11');
    y11.value = VALUES.PERCENT;
    y11.style = PERCENT_STYLE;
    let z11 = workSheet.getCell('Z11');
    z11.value = VALUES.MINUTES;
    z11.style = MINUTE_STYLE;

    let aa11 = workSheet.getCell('AA11');
    aa11.value = VALUES.PERCENT;
    aa11.style = PERCENT_STYLE;
    let ab11 = workSheet.getCell('AB11');
    ab11.value = VALUES.MINUTES;
    ab11.style = MINUTE_STYLE;

    let ac11 = workSheet.getCell('AC11');
    ac11.value = VALUES.PERCENT;
    ac11.style = PERCENT_STYLE;
    let ad11 = workSheet.getCell('AD11');
    ad11.value = VALUES.MINUTES;
    ad11.style = MINUTE_STYLE;

    let ae11 = workSheet.getCell('AE11');
    ae11.value = VALUES.PERCENT;
    ae11.style = PERCENT_STYLE;
    let af11 = workSheet.getCell('AF11');
    af11.value = VALUES.MINUTES;
    af11.style = MINUTE_STYLE;

    let AG11 = workSheet.getCell('AG11');
    AG11.value = VALUES.PERCENT;
    AG11.style = PERCENT_STYLE;
    let AH11 = workSheet.getCell('AH11');
    AH11.value = VALUES.MINUTES;
    AH11.style = MINUTE_STYLE;

    let row11 = workSheet.getRow(11);
    row11.eachCell((cell) => {
      cell.border = BORDER_STYLE;
      cell.font = { bold: true };
      cell.alignment = { horizontal: 'center' };
    });

    manipulatedNmsData.forEach((row: any) => {
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
            hrtDownPercent +
            row.planned_maintenance_in_percent;
      let totalExclusionMinutes: number =
        row.power_downtime_in_minutes +
        row.dcn_downtime_in_minutes +
        hrtDownMinutes +
        row.planned_maintenance_in_minutes;
      let pollingTimePercent: number = row.polling_time_in_percent;
      let pollingTimeMinutes: number = row.polling_time_in_minutes;
      let totalUpPercentSLAExclusion: number =
        upPercent +
        row.power_downtime_in_percent +
        row.dcn_downtime_in_percent +
        hrtDownPercent +
        row.planned_maintenance_in_percent +
        row.polling_time_in_percent;
      let totalUpMinutesSLAExclusion: number =
        upMinute +
        row.power_downtime_in_minutes +
        row.dcn_downtime_in_minutes +
        hrtDownMinutes +
        row.planned_maintenance_in_minutes +
        row.polling_time_in_minutes;

      const blockDeviceLevelRowValues = workSheet.addRow([
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
        upPercent,
        upMinute,
        downPercent,
        downMinute,
        powerDownPercent,
        powerDownMinutes,
        fiberDownPercent,
        fiberDownMinutes,
        equipmentDownPercent,
        equipmentDownMinutes,
        hrtDownPercent,
        hrtDownMinutes,
        dcnDownPercent,
        dcnDownMinutes,
        plannedMaintanancePercent,
        plannedMaintananceMinutes,
        unKnownDownPercent,
        unKnownDownMinutes,
        totalExclusionPercent,
        totalExclusionMinutes,
        pollingTimePercent,
        pollingTimeMinutes,
        +totalUpPercentSLAExclusion > 100 ? 100 : totalUpPercentSLAExclusion,
        totalUpMinutesSLAExclusion,
      ]);

      const unknownPercentColumn = blockDeviceLevelRowValues.getCell(27);
      unknownPercentColumn.style = UNKNOWN_COLUMN_COLOR;

      const unknownMinuteColumn = blockDeviceLevelRowValues.getCell(28);
      unknownMinuteColumn.style = UNKNOWN_COLUMN_COLOR;

      blockDeviceLevelRowValues.eachCell((cell, colNumber: number) => {
        cell.border = BORDER_STYLE;
        cell.alignment = { horizontal: 'left' };
        if (colNumber > 10) {
          cell.numFmt = '0.00';
        }
      });
    });

    // Generating Sheet 2
    const ttCorelationWorkSheet = workbook.addWorksheet('Block-TT co-relation');
    ttCorelationWorkSheet.columns = BLOCK_TT_CO_RELATION_COLUMNS_WIDTHS;
    ttCorelationWorkSheet
      .addRow(BLOCK_TT_CO_RELATION_HEADERS)
      .eachCell((cell) => {
        cell.style = TABLE_HEADERS;
      });
    ttCorelation.forEach((ttCorelationData: TTCorelation, index: number) => {
      let block_device_detail = BLOCK_DEVICE_DETAILS.filter(
        (device: BlockDeviceDetail) => {
          return device.ip_address == ttCorelationData.ip;
        }
      );
      let [blockDeviceDetail] = block_device_detail;
      ttCorelationWorkSheet
        .addRow([
          index + 1,
          ttCorelationData.ip,
          blockDeviceDetail.block_name,
          ttCorelationData.powerIssueTT.toString().split(',').join(', '),
          ttCorelationData.linkIssueTT.toString().split(',').join(', '),
          ttCorelationData.otherTT.toString().split(',').join(', '),
        ])
        .eachCell({ includeEmpty: true }, (cell) => {
          cell.border = BORDER_STYLE;
          cell.alignment = { horizontal: 'left' };
        });
    });
  }

  downloadBlockInputTemplate(): void {
    const workbook = new ExcelJS.Workbook();
    const slaWorksheet = workbook.addWorksheet(BLOCK_INPUT_FILE_NAMES[0]);
    BLOCK_SLA_REPORT_HEADERS.forEach((_, index) => {
      slaWorksheet.getColumn(index + 1).width = 40;
    });
    slaWorksheet.getColumn(1).width = 80;
    slaWorksheet.getCell('A1').value =
      'Time Span: From 01 May  2023 12:00:00 AM To 01 May  2023 11:59:59 PM (Example)';
    slaWorksheet.addRow(BLOCK_SLA_REPORT_HEADERS);

    const blockAlertWorksheet = workbook.addWorksheet(
      BLOCK_INPUT_FILE_NAMES[2]
    );
    BLOCK_ALERT_REPORT_HEADERS.forEach((_, index) => {
      blockAlertWorksheet.getColumn(index + 1).width = 40;
    });
    blockAlertWorksheet.addRow(BLOCK_ALERT_REPORT_HEADERS);

    const blockTTWorksheet = workbook.addWorksheet(BLOCK_INPUT_FILE_NAMES[1]);
    blockTTWorksheet.addRow(TT_REPORT_HEADERS);
    TT_REPORT_HEADERS.forEach((_, index) => {
      blockTTWorksheet.getColumn(index + 1).width = 30;
    });

    workbook.xlsx.writeBuffer().then((buffer) => {
      this.sharedService.downloadFinalReport(
        buffer,
        'Block_Input_Template',
        true
      );
    });
  }
}
