import { Component, OnInit } from '@angular/core';
import * as ExcelJS from 'exceljs';
import * as moment from 'moment';
import * as lodash from 'lodash';

import { ALERT_DOWN_MESSAGE, BLOCK_GP_SLA_SUMMARY_PERCENT, BLOCK_SLA_FINAL_REPORT_COLUMNS, BORDER_STYLE, BlockSLASummaryMinutesHeaders, BlockSLASummaryPercentHeaders,RFO_CATEGORIZATION, SEVERITY_CRITICAL, SEVERITY_WARNING, SHEET_HEADING, TABLE_HEADING } from '../constants/constants';
import { BlockAlertData, BlockNMSData, BlockSLASummaryPercent, BlockTTData, ManipulatedNMSData, RFOCategorizedTimeInMinutes } from './block-component.model';

export type AOA = [][];


@Component({
  selector: 'app-block-component',
  templateUrl: './block-component.component.html',
  styleUrls: ['./block-component.component.scss']
})
export class BlockComponentComponent implements OnInit {

  blockNMSData: any = [];
  blockTTData: any = [];
  blockAlertData: any = [];
  manipulatedNMSData: any = [];
  blockSLASummaryPercent!: BlockSLASummaryPercent;
  worksheet!: ExcelJS.Worksheet;

  constructor() { }

  ngOnInit(): void {
  }

    // Getting the input file (excel workbook containing the required sheets)
    onFileChange(event: any): void {
      const file = event.target.files[0];
      const workbook = new ExcelJS.Workbook();
      const reader = new FileReader();
  
      reader.onload = (e: any) => {
        const buffer = e.target.result;
  
        workbook.xlsx.load(buffer).then(() => {
          workbook.worksheets.forEach((_, index) => {
            this.worksheet = workbook.getWorksheet(index + 1);
            this.readWorksheet(this.worksheet);
          });
          this.manipulateBlockNMSData();
          this.categorizeRFO('10.128.0.32');
        });
      };
  
      reader.readAsArrayBuffer(file);
    }
  
    // Reading the worksheets individually and storing the data as Array of Objects
    readWorksheet(worksheet: ExcelJS.Worksheet): void {
      let workSheetName = worksheet.name;
      let data: AOA = [];
      this.worksheet.eachRow({ includeEmpty: true }, (row: ExcelJS.Row) => {
        const rowData: any = [];
        row.eachCell({ includeEmpty: true }, (cell: ExcelJS.Cell) => {
          rowData.push(cell.value);
        });
        data.push(rowData);
      });
  
      const headers = data[0];
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
            resolution_type_in_min: data[36],
            sla_ageing: data[37],
            reporting_sla: data[38],
            reopen_date: data[39],
            category: data[40],
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
          parseInt(totalTimeinMinutes[0]) * 24 +
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
      this.calcluateBlockSLASummaryinPercent();
      this.generateFinalBlockReport();
    }
  
    calcluateBlockSLASummaryinPercent() {
      let upPercent = 0;
      let powerDownPercent = 0;
      let dcnDownPercent = 0;
      let plannedMaintenance = 0;
      let dcnAndPowerDownPercent = 0;
  
      this.manipulatedNMSData.forEach((nmsData: ManipulatedNMSData) => {
        upPercent += nmsData.up_percent;
        powerDownPercent += nmsData.power_downtime_in_percent;
        dcnDownPercent += nmsData.dcn_downtime_in_percent;
        dcnAndPowerDownPercent +=
          nmsData.power_downtime_in_percent + nmsData.dcn_downtime_in_percent;
      });
  
      this.blockSLASummaryPercent = {
        report_type: 'BLOCK-SLA',
        time_span: '',
        no_of_blocks: 79,
        up_percent: upPercent / 79,
        no_of_up_blocks: '',
        power_down_percent: powerDownPercent / 79,
        fibre_down_percent: 0.0,
        equipment_down_percent: 0.0,
        hrt_down_percent: 0.0,
        dcn_down_percent: dcnDownPercent / 79,
        planned_maintenance_percent: 0.0,
        down_percent_exclusive_of_sla: 100 - upPercent / 79,
        no_of_down_blocks: '',
        total_sla_exclusion_percent: dcnAndPowerDownPercent / 79,
        total_up_percent: 0,
      };
    }
  
    getBlockNames(department: string) {
      let splittedValue = department.trim().split('-');
      return splittedValue[splittedValue.length - 1];
    }
  
    // Generating the final report as excel-workbook using the calculated data.
    generateFinalBlockReport(): void {
      const workbook = new ExcelJS.Workbook();
      const worksheet = workbook.addWorksheet('Block-Final-Report');
      worksheet.columns = BLOCK_SLA_FINAL_REPORT_COLUMNS;
  
      worksheet.mergeCells('A1:B1');
      let cellA1 = worksheet.getCell('A1');
      cellA1.value = '1. Daily Network availability report';
      cellA1.style = SHEET_HEADING;
  
      worksheet.mergeCells('D1:E1');
      let cellE1 = worksheet.getCell('D1');
      cellE1.value = 'Report-Frequency- Daily';
      cellE1.style = { font: { bold: true } };
  
      worksheet.mergeCells('A3:B3');
      let cellA3 = worksheet.getCell('A3');
      cellA3.value = 'Block - SLA Summary (%)';
      cellA3.style = TABLE_HEADING;
      let blockSLASummaryPercentHeaders = worksheet.addRow(
        BlockSLASummaryPercentHeaders
      );
  
      blockSLASummaryPercentHeaders.eachCell((cell) => {
        cell.border = BORDER_STYLE;
        cell.font = { bold: true };
      });
      blockSLASummaryPercentHeaders.alignment = {
        horizontal: 'center',
        wrapText: true,
      };
      let blockSLASummaryPercentArray = Object.values(
        this.blockSLASummaryPercent
      );
      let blockSLASummaryPercentValues = worksheet.addRow(
        blockSLASummaryPercentArray
      );
  
      blockSLASummaryPercentValues.eachCell((cell) => {
        cell.alignment = { horizontal: 'left' };
        cell.border = BORDER_STYLE;
      });
  
      worksheet.addRow('');
  
      worksheet.mergeCells('A7:B7');
      let cellA7 = worksheet.getCell('A7');
      cellA7.value = 'Block - SLA Summary (min)';
      cellA7.style = TABLE_HEADING;
      worksheet.addRow(BlockSLASummaryMinutesHeaders);
  
      worksheet.addRow('');
  
      worksheet.mergeCells('A11:B11');
      let cellA11 = worksheet.getCell('A11');
      cellA11.value = 'Block - GP  SLA Summary (%)';
      cellA11.style = TABLE_HEADING;
      const blockSummaryPercentRow = worksheet.addRow(
        BLOCK_GP_SLA_SUMMARY_PERCENT
      );
      blockSummaryPercentRow.font = { bold: true };
      blockSummaryPercentRow.alignment = { wrapText: true, horizontal: 'center' };
      blockSummaryPercentRow.eachCell((cell) => {
        if (cell.value) {
          cell.border = BORDER_STYLE;
        }
      });
  
      this.manipulatedNMSData.forEach((row: any) => {
        let reportType: string = 'Block - SLA';
        let timeSpan: string = '';
        let ipAddress: string = row.ip_address;
        let state: string = 'Chhattisgarh';
        let cluster: string = '';
        let district: string = '';
        let districtLGDCode: string = '';
        let blockName: string = this.getBlockNames(row.departments);
        let blockLGDCode: string = '';
        let noOfGPinBlock: string = '';
        let upPercent: number = row.up_percent;
        let noOfUpGPCount: string = '';
        let powerDown: string =
          upPercent == 100
            ? '0'
            : `${row.power_downtime_in_percent}, min: ${row.power_downtime_in_minutes}`;
        let fiberDown: string = upPercent == 100 ? '0' : '0';
        let equipmentDown: string = upPercent == 100 ? '0' : '0';
        let hrtDownPercent: string = upPercent == 100 ? '0' : '0';
        let dcnDownPercent: string =
          upPercent == 100
            ? '0'
            : `${row.dcn_downtime_in_percent}, min: ${row.dcn_downtime_in_minutes}`;
        let plannedMaintanance: string = upPercent == 100 ? '0' : '0';
        let downPercentSLAExclusions: string =
          upPercent == 100 ? '0' : row.down_percent;
        let noOfDownGPCount: string = '';
        let totalExclusionPercent: number =
          upPercent == 100
            ? 0
            : row.power_downtime_in_percent + row.dcn_downtime_in_percent;
        // : `${
        //     row.power_downtime_in_percent + row.dcn_downtime_in_percent
        //   }, min: ${
        //     row.power_downtime_in_minutes + row.dcn_downtime_in_minutes
        //   }`;
        let pollingTimePercent: number =
          upPercent == 100 ? 0 : row.down_percent - +totalExclusionPercent;
  
        let totalUpPercentSLAExclusion: number = 100;
  
        const blockSummaryPercentRowValues = worksheet.addRow([
          reportType,
          timeSpan,
          ipAddress,
          state,
          cluster,
          district,
          districtLGDCode,
          blockName,
          blockLGDCode,
          noOfGPinBlock,
          upPercent,
          noOfUpGPCount,
          powerDown,
          fiberDown,
          equipmentDown,
          hrtDownPercent,
          dcnDownPercent,
          plannedMaintanance,
          downPercentSLAExclusions,
          noOfDownGPCount,
          totalExclusionPercent,
          pollingTimePercent,
          totalUpPercentSLAExclusion,
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
    }

}
