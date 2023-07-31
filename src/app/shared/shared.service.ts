import { Injectable } from '@angular/core';
import * as moment from 'moment';
import { FROM_DATE_REGEX, TO_DATE_REGEX } from '../constants/constants';

@Injectable({
  providedIn: 'root',
})
export class SharedService {
  constructor() {}

  calculateTimeInMinutes(timePeriod: string): number {
    if (timePeriod) {
      let totalTimeinMinutes = timePeriod.trim().split(' ');

      let averageDaysPerYear = 365.25;
      let averageDaysPerMonth = 30.44;

      return +(
        +parseInt(totalTimeinMinutes[0]) * averageDaysPerYear * 24 * 60 +
        +parseInt(totalTimeinMinutes[2]) * averageDaysPerMonth * 24 * 60 +
        +parseInt(totalTimeinMinutes[4]) * 1440 +
        +parseInt(totalTimeinMinutes[6]) * 60 +
        +parseInt(totalTimeinMinutes[8]) +
        +parseInt(totalTimeinMinutes[10]) / 60
      );
    } else {
      return 0;
    }
  }

  setStandardTime(time: any) {
    return typeof time === 'object'
      ? moment(time).subtract(5, 'hours').subtract(30, 'minutes').format()
      : moment(time).format();
  }

  extractFromDateFromTimeSpan(timeSpan: string) {
    const matches = timeSpan.match(FROM_DATE_REGEX);
    return matches ? matches[1] : null;
  }

  extractToDateFromTimeSpan(timeSpan: string) {
    const matches = timeSpan.match(TO_DATE_REGEX);
    return matches ? matches[1] : null;
  }

  extractYearFromTime(time: string) {
    const matches = time.match(/(\d+)\s+year/);
    return matches ? matches[1] : null;
  }

  extractMonthFromTime(time: string) {
    const matches = time.match(/(\d+)\s+month/);
    return matches ? matches[1] : null;
  }

  extractDayFromTime(time: string) {
    const matches = time.match(/(\d+)\s+day/);
    return matches ? matches[1] : null;
  }

  extractHourFromTime(time: string) {
    const matches = time.match(/(\d+)\s+hour/);
    return matches ? matches[1] : null;
  }

  extractMinuteFromTime(time: string) {
    const matches = time.match(/(\d+)\s+minute/);
    return matches ? matches[1] : null;
  }

  extractSecondFromTime(time: string) {
    const matches = time.match(/(\d+)\s+second/);
    return matches ? matches[1] : null;
  }

  formatTimeInSlaReport(timeFromSlaReport: string): string {
    let years = timeFromSlaReport.includes('year')
      ? this.extractYearFromTime(timeFromSlaReport)
      : 0;
    let months = timeFromSlaReport.includes('month')
      ? this.extractMonthFromTime(timeFromSlaReport)
      : 0;
    let days = timeFromSlaReport.includes('day')
      ? this.extractDayFromTime(timeFromSlaReport)
      : 0;
    let hours = timeFromSlaReport.includes('hour')
      ? this.extractHourFromTime(timeFromSlaReport)
      : 0;
    let minutes = timeFromSlaReport.includes('minute')
      ? this.extractMinuteFromTime(timeFromSlaReport)
      : 0;
    let seconds = timeFromSlaReport.includes('second')
      ? this.extractSecondFromTime(timeFromSlaReport)
      : 0;
    return `${years} year(s) ${months} month(s) ${days} day(s) ${hours} hour(s) ${minutes} minute(s) ${seconds} second(s)`;
  }

  foramatDuration(diffInMilliSeconds: number): string {
    let years = moment.duration(diffInMilliSeconds).years();
    let months = moment.duration(diffInMilliSeconds).months();
    let days = moment.duration(diffInMilliSeconds).days();
    let hours = moment.duration(diffInMilliSeconds).hours();
    let minutes = moment.duration(diffInMilliSeconds).minutes();
    let seconds = moment.duration(diffInMilliSeconds).seconds();

    return `${years} year(s) ${months} month(s) ${days} day(s) ${hours} hour(s) ${minutes} minute(s) ${seconds} second(s)`;
  }

  setDuration(
    timeSpan: string,
    alarmStartTime: any,
    alarmClearTime: any,
    duration: string
  ): string {
    let fromDate = this.extractFromDateFromTimeSpan(timeSpan);
    let toDate = this.extractToDateFromTimeSpan(timeSpan);

    if (
      moment(alarmStartTime).isAfter(moment(fromDate)) &&
      moment(alarmClearTime).isBefore(moment(toDate))
    ) {
      return this.formatTimeInSlaReport(duration);
    } else if (
      moment(alarmStartTime).isBefore(moment(fromDate)) &&
      moment(alarmClearTime).isBefore(moment(toDate))
    ) {
      // let diffInMilliSeconds = 0;
      // if (moment(alarmClearTime).isBefore(fromDate)) {
      //   diffInMilliSeconds = moment(alarmClearTime).diff(
      //     moment(alarmStartTime)
      //   );
      //   diffInMilliSeconds = 0;
      // } else {
      //   diffInMilliSeconds = moment(alarmClearTime).diff(moment(fromDate));
      // }
      let diffInMilliSeconds = moment(alarmClearTime).diff(moment(fromDate));
      return this.foramatDuration(diffInMilliSeconds);
    } else if (
      moment(alarmStartTime).isAfter(fromDate) &&
      moment(alarmClearTime).isAfter(toDate)
    ) {
      let diffInMilliSeconds = moment(toDate).diff(moment(alarmStartTime));
      return this.foramatDuration(diffInMilliSeconds);
    } else {
      let diffInMilliSeconds = moment(toDate).diff(moment(fromDate));
      return this.foramatDuration(diffInMilliSeconds);
    }
  }

  downloadFinalReport(buffer: ArrayBuffer, fileName: string) {
    const data = new Blob([buffer], {
      type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
    });
    const link = document.createElement('a');
    link.href = window.URL.createObjectURL(data);
    link.download =
      fileName + ' ' + moment().format('DD/MM/YYYY, hh:mm') + '.xlsx';
    link.click();
  }
}
