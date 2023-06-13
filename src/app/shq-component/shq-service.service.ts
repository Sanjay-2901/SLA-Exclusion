import { Injectable } from '@angular/core';
import { ShqAlertData, ShqTTData } from './shq-component.model';
import {
  ALERT_DOWN_MESSAGE,
  RFO_CATEGORIZATION,
  SEVERITY_CRITICAL,
  SEVERITY_WARNING,
} from '../constants/constants';
import * as moment from 'moment';
import * as lodash from 'lodash';
import { RFOCategorizedTimeInMinutes } from '../block-component/block-component.model';

@Injectable({
  providedIn: 'root',
})
export class ShqService {
  constructor() {}

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
}
