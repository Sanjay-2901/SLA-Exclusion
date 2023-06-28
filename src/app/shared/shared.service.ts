import { Injectable } from '@angular/core';
import * as moment from 'moment';

@Injectable({
  providedIn: 'root',
})
export class SharedService {
  constructor() {}

  CalucateTimeInMinutes(timePeriod: string): number {
    if (timePeriod) {
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
    } else {
      return 0;
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
