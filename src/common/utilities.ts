import { isEmpty } from '@microsoft/sp-lodash-subset';
import {  format, getDate, isValid, isBefore } from 'date-fns';


export default class Utilities {

  public static getDayFromDate(date: Date): number {
    return isValid(date) ? getDate(date) : 0;
  }
  public static getMonthFromDate(date: Date): string {
    const month = isValid(date) ? date.toLocaleString('default', { month: 'short' }) : "";
    return month;
  }
  public static getLocaleDateString(dateStr: string): string {
    return !isEmpty(dateStr) ? format(new Date(dateStr), 'yyyy-MM-dd') : "";
  }
  public static getCurrentDate(): string {
    let result = format(new Date(), 'yyyy-MM-dd');
    return result;
  }
  public static isDateBeforeToday(dateStr :any) {
    let result = isBefore(new Date(dateStr), new Date())
    return result;
  }
  public static IsNullOrEmpty(value: any): boolean {
    return isEmpty(value);
  }
  public static GetStatus(percentComplete : number) {
    // let status = {
    //   '0': 'Pending',
    //   '50': 'In Progress',
    //   '100': 'Completed',
    //   'default': 'Pending'
    // };
    let stusofpercent = percentComplete == 0 ? "Pending" : percentComplete == 50 ? "In Progress" :  percentComplete == 100 ? "Completed" : "Pending"
    return ( stusofpercent );
  }

  public static GetSelectedTypeName(type: string)
  {
    // let TasksTypes = {
    //   'due': 'Upcoming Tasks',
    //   'overdue': 'Overdue Tasks',
    //   'inprogress': 'In Progress Tasks',
    //   'pending': 'Pending Tasks',
    //   'completed': 'Completed Tasks',
    //   'default': 'Upcoming Tasks'
    // };
    let getSelectedTypeName =  type == 'urgent' ? "Urgent Tasks" : type == 'due' ? "Upcoming Tasks" : type == 'overdue' ? "Overdue Tasks" :  type == 'inprogress' ? "In Progress Tasks" : type == 'pending' ? "Pending Tasks" : type == 'completed' ? "Completed Tasks" : "Upcoming Tasks";
    return (getSelectedTypeName);
  }
}