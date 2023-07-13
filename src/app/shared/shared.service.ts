import { Injectable } from '@angular/core';
import * as moment from 'moment';

@Injectable({
  providedIn: 'root',
})
export class SharedService {
  constructor() {}

  calculateTimeInMinutes(timePeriod: string): number {
    if (timePeriod) {
      let totalTimeinMinutes = timePeriod.trim().split(' ');
      if (timePeriod.includes('day')) {
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
    } else {
      return 0;
    }
  }

  setStandardTime(time: any) {
    return typeof time === 'object'
      ? moment(time).subtract(5, 'hours').subtract(30, 'minutes').format()
      : moment(time).format();
  }

  setDuration(
    timeSpan: string,
    alarmStartTime: any,
    alarmClearTime: any,
    duration: string
  ): string {
    let toDate = moment(
      timeSpan.substring(timeSpan.indexOf('To ') + 3)
    ).format();

    if (moment(alarmClearTime).isAfter(moment(toDate))) {
      let diffInMilliSeconds = moment(toDate).diff(moment(alarmStartTime));
      let days = moment.duration(diffInMilliSeconds).days();
      let hours = moment.duration(diffInMilliSeconds).hours();
      let minutes = moment.duration(diffInMilliSeconds).minutes();
      let seconds = moment.duration(diffInMilliSeconds).seconds();
      return `${days} day(s) ${hours} hour(s) ${minutes} minute(s) ${seconds} second(s)`;
    } else {
      return duration;
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
