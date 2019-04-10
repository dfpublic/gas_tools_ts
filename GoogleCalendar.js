function GoogleCalendar(calendarId) {
  this.calendar = CalendarApp.getCalendarById(calendarId);
}
GoogleCalendar.SpreadsheetHeaders = ['id', 'calendar_name', 'title', 'start_time', 'end_time', 'duration', 'is_all_day', 'is_recurring'];
GoogleCalendar.FormatDate = function (year, month, day) {
  return new Date(parseInt(year), parseInt(month) - 1, parseInt(day));
}
GoogleCalendar.prototype.getEvents = function (startTime, endTime) {
  var calendar_name = this.calendar.getName();
  return this.calendar.getEvents(startTime, endTime).map(
    function (event) {
      var event_start = event.getStartTime();
      var event_end = event.getEndTime();
      var start_day = formatDay(event_start);
      var end_day = formatDay(event_end);
      return {
        id: event.getId(),
        calendar_name: calendar_name,
        title: event.getTitle(),
        description: event.getDescription(),
        start_time: event_start,
        end_time: event_end,
        duration: start_day === end_day ? start_day : start_day + ' - ' + end_day,
        //month: formatMonth(event_start),
        is_all_day: event.isAllDayEvent() ? 1 : 0,
        is_recurring: event.isRecurringEvent() ? 1 : 0,
        toShortString: function () {
          if (this.is_all_day) {
            return '[' + formatDate(this.start_time) + '] | ' + this.title;
          } else {
            return '[' + formatDate(this.start_time) + ' - ' + formatDate(this.end_time) + ']' + ' | ' + this.title;
          }
          //return '[' + this.start_time.toDateString() + ' - ' + this.end_time.toDateString() + ']' + ' | ' + this.title;
        }
      }
    }
  );
}

/**
 * 
 * @param {{id: string, calendar_name: string, title: string, description: string, start_time: Date, end_time: Date}} event 
 */
GoogleCalendar.prototype.updateEvent = function(event) {
  var self = this;
  if(event.id) {
    var calendar_event = CalendarApp.getEventById(event.id);
    if(calendar_event) {
      calendar_event.setTitle(event.title);
      calendar_event.setDescription(event.description);
      calendar_event.setTime(event.start_time, event.end_time);
    }
  }
  else {
    var new_event = self.createEvent(event.title, event.start_time, event.end_time);
    new_event.setDescription(event.description);
  }
}

GoogleCalendar.prototype.deleteAllEvents = function (startTime, endTime) {
  return this.calendar.getEvents(startTime, endTime).map(
    function (event) {
      event.deleteEvent();
    }
  );
}

GoogleCalendar.prototype.createEvent = function (title, startTime, endTime) {
  var formattedStartTime = new Date(startTime);
  var formattedEndTime = new Date(endTime);
  return this.calendar.createEvent(title, formattedStartTime, formattedEndTime);
}

/**
 * @param {Date} date_raw 
 */
function formatDate(date_raw) {
  var date_formatted = {
    year: date_raw.getFullYear(),
    month: date_raw.getMonth() + 1,
    date: date_raw.getDate()
  };
  var year = date_formatted.year,
    month = date_formatted.month,
    date = date_formatted.date;
  month = month < 10 ? '0' + month : '' + month;
  date = date < 10 ? '0' + date : '' + date;
  return year + "/" + month + "/" + date;
  // return '' + date.getFullYear() + '/' + (date.getMonth() + 1) + '/' + date.getDate();
}

/**
 * @param {Date} date_raw 
 */
function formatDateTime(date_raw) {
  var date_formatted = {
    year: date_raw.getFullYear(),
    month: date_raw.getMonth() + 1,
    date: date_raw.getDate(),
    hours: date_raw.getHours(),
    minutes: date_raw.getMinutes(),
    seconds: date_raw.getSeconds()

  };
  var year = date_formatted.year,
    month = date_formatted.month,
    date = date_formatted.date,
    hours = date_formatted.hours,
    minutes = date_formatted.minutes,
    seconds = date_formatted.seconds;
  month = month < 10 ? '0' + month : '' + month;
  date = date < 10 ? '0' + date : '' + date;
  hours = hours < 10 ? '0' + hours : '' + hours;
  minutes = minutes < 10 ? '0' + minutes : '' + minutes;
  seconds = seconds < 10 ? '0' + seconds : '' + seconds;
  return year + "/" + month + "/" + date + " " + hours + ":" + minutes + ":" + seconds;
}

function formatMonth(date) {
  switch (date.getMonth()) {
    case 0:
      return 'January';
    case 1:
      return 'February';
    case 2:
      return 'March';
    case 3:
      return 'April';
    case 4:
      return 'May';
    case 5:
      return 'June';
    case 6:
      return 'July';
    case 7:
      return 'August';
    case 8:
      return 'September';
    case 9:
      return 'October';
    case 10:
      return 'November';
    case 11:
      return 'December';
  }
}

/**
 * @param {Date} date 
 */
function formatDay(date) {
  switch (date.getDay()) {
    case 0:
      return 'Sunday'
    case 1:
      return 'Monday'
    case 2:
      return 'Tuesday'
    case 3:
      return 'Wednesday'
    case 4:
      return 'Thursday'
    case 5:
      return 'Friday'
    case 6:
      return 'Saturday';
    default:
      return 'Unknown'
  }
}