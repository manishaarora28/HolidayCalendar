import {
  Component,
  OnInit
} from '@angular/core';
import * as moment from 'moment';
import * as XLSX from 'xlsx';

@Component({
  selector: 'app-calendar',
  templateUrl: './calendar.component.html',
  styleUrls: ['./calendar.component.scss']
})
export class CalendarComponent implements OnInit {
  days = [{
          val: 'Sunday',
          index: 0
      },
      {
          val: 'Monday',
          index: 1
      },
      {
          val: 'Tuesday',
          index: 2
      },
      {
          val: 'Wednesday',
          index: 3
      },
      {
          val: 'Thursday',
          index: 4
      },
      {
          val: 'Friday',
          index: 5
      },
      {
          val: 'Saturday',
          index: 6
      }
  ];

  totalDays: number;
  dayNum = [];
  date: any;
  yearInput: any;
  monthInput: number;
  monthNum: number;
  currentMonthYear: string;
  file: any;
  filelist = [];
  arrayBuffer: any;
  arrayList = [];
  listOfDates = [];
  listOfOptionalDates = [];
  holidays = [];
  
  constructor() {}

  ngOnInit(): void {
      this.date = new Date();
      this.yearInput = this.date.getFullYear();
      this.monthInput = this.date.getMonth();
      this.monthNum = this.monthInput + 1;
      this.currentMonthYear = moment().month(this.monthInput).format('MMMM, YYYY');
      this.addfile();
  }

  calendarDays() {
      this.holidays = [];      
      this.dayNum = [];
      this.totalDays = moment().month(this.monthInput).daysInMonth();
      var firstDay = new Date(this.yearInput, this.monthInput, 1);
      var lastDay = new Date(this.yearInput, this.monthInput + 1, 0);
      this.listOfDates.forEach(element => {
          let monthnum = moment(element.date, 'DD-MMM-YYYY').month();
          let monthday = moment(element.date, 'DD-MMM-YYYY').date();
          if (this.monthInput == monthnum) {
              this.holidays.push({
                  date: monthday,
                  holiday: element.holiday,
                  optional: false
              });
          }
      });
      this.listOfOptionalDates.forEach(element => {
          let monthnum = moment(element.date, 'DD-MMM-YYYY').month();
          let monthday = moment(element.date, 'DD-MMM-YYYY').date();
          if (this.monthInput == monthnum) {
              this.holidays.push({
                  date: monthday,
                  holiday: element.holiday,
                  optional: true
              });
          }
      });

      // sort by date
      this.sortHolidays();
      
      let count = firstDay.getDay();
      let holidayCount = 0;

      if (this.holidays.length > 1) {
          holidayCount = 1;
      } else if (this.holidays.length == 1) {
          holidayCount = 0;
      }
      if (count == 0) {
          for (var i = 1; i <= this.totalDays; i++) {
              if (this.holidays.length > 0 && holidayCount == 0) {
                  if (this.holidays[0].date == i) {
                      this.dayNum.push({
                          val: i,
                          day: this.holidays[0].holiday,
                          optional: this.holidays[0].optional
                      });
                  } else {
                      this.dayNum.push({
                          val: i,
                          day: '',
                          optional: false
                      });
                  }
              } else if (this.holidays.length > 0 && holidayCount == 1) {
                  for (var k = 0; k < this.holidays.length; k++) {
                      if (this.holidays[k].date == i) {
                          this.dayNum.push({
                              val: i,
                              day: this.holidays[k].holiday,
                              optional: this.holidays[k].optional
                          });
                          this.holidays = this.holidays.splice(1);
                          break;
                         
                      } else {
                          this.dayNum.push({
                              val: i,
                              day: '',
                              optional: false
                          });
                          break;
                      }
                  }
              } else {
                  this.dayNum.push({
                      val: i,
                      day: '',
                      optional: false
                  });
              }
          }
      } else {
          for (var i = 1; i <= count; i++) {
              this.dayNum.push('');
          }
          for (var j = 1; j <= this.totalDays; j++) {
              if (this.holidays.length > 0 && holidayCount == 0) {
                  if (this.holidays[0].date == j) {
                      this.dayNum.push({
                          val: j,
                          day: this.holidays[0].holiday,
                          optional: this.holidays[0].optional
                      });
                  } else {
                      this.dayNum.push({
                          val: j,
                          day: '',
                          optional: false
                      });
                  }
              } else if (this.holidays.length > 0 && holidayCount == 1) {
                  for (var k = 0; k < this.holidays.length; k++) {
                      if (this.holidays[k].date == j) {
                          this.dayNum.push({
                              val: j,
                              day: this.holidays[k].holiday,
                              optional: this.holidays[k].optional
                          });
                         this.holidays = this.holidays.splice(1);
                         break;
                        
                      } else {
                          this.dayNum.push({
                              val: j,
                              day: '',
                              optional: false
                          });
                          break;
                      }
                  }
              } else {
                  this.dayNum.push({
                      val: j,
                      day: '',
                      optional: false
                  });
              }
          }
      }
  }

  previousMonth() {
      this.monthNum = this.monthNum - 1;
      var previousMonth = new Date(this.yearInput, this.monthNum, 1).getMonth();     
      this.currentMonthYear = moment(previousMonth.toString()).format('MMMM') + `, ` + this.yearInput;
      this.monthInput = this.monthNum - 1;
      this.dayNum = [];
      this.calendarDays();
  }

  nextMonth() {
      this.monthNum = this.monthNum + 1;
      var nextMonth = new Date(this.yearInput, this.monthNum, 1).getMonth();
      if (this.monthNum == 12) {
          nextMonth = 12;
      }
      this.currentMonthYear = moment(nextMonth.toString()).format('MMMM') + `, ` + this.yearInput;
      this.monthInput = this.monthNum - 1;
      this.dayNum = [];
      this.calendarDays();
  }

  addfile() {
      this.arrayList = [];
      let url = "/assets/files/PL_Holiday Calendar 2020.xlsx";
      let req = new XMLHttpRequest();
      req.open("GET", url, true);
      req.responseType = "arraybuffer";
      req.onload = (e) => {
          let data = new Uint8Array(req.response);
          var arr = new Array();
          for (var i = 0; i != data.length; ++i) arr[i] = String.fromCharCode(data[i]);
          var bstr = arr.join("");
          var workbook = XLSX.read(bstr, {
              type: "binary",
              cellDates: true
          });
          var first_sheet_name = workbook.SheetNames[0];
          var worksheet = workbook.Sheets[first_sheet_name];
          this.arrayList = XLSX.utils.sheet_to_json(worksheet, {
              raw: false
          });
          this.filelist = [];
          this.getDates();
          this.calendarDays();
      };
      req.send();
  }

  getDates() {
      var arr = this.arrayList.splice(2);
      let index = 0;
      arr.forEach(element => {
          if (element['LIST OF HOLIDAYS'] != "ADDITIONAL OPTIONAL LEAVES-2020 - Pick Any Two") {
              this.listOfDates.push({
                  'date': element.__EMPTY_1,
                  'holiday': element.__EMPTY_2
              });
          } else {
              var optionalarr = arr.splice(index + 1); //
              for (var i = 1; i < optionalarr.length; i++) {
                  if (optionalarr[i].__EMPTY_1 != "DATE") {
                      if (optionalarr[i].__EMPTY_2 != "Your Birthday") {
                          if (optionalarr[i].__EMPTY_2 != "Your Spouse's Birthday") {
                              if (optionalarr[i].__EMPTY_2 != "Your Child's Birthday") {
                                  if (optionalarr[i].__EMPTY_2 != "Your Anniversary") {
                                      if (optionalarr[i]['LIST OF HOLIDAYS'] != "*SUBJECT TO VISBILITY OF MOON") {
                                          this.listOfOptionalDates.push({
                                              'date': optionalarr[i].__EMPTY_1,
                                              'holiday': optionalarr[i].__EMPTY_2
                                          });

                                      } else {
                                          continue;
                                      }
                                  }
                              }
                          }
                      }
                  }
              }
          }
          ++index;
      });
  }

  sortHolidays(){
    this.holidays.sort(function (a, b) {
        return a.date - b.date;
    });
  }
}