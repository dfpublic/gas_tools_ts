function GoogleCalendar(calendarId) {
  this.calendar = CalendarApp.getCalendarById(calendarId);
}
GoogleCalendar.SpreadsheetHeaders = ['id', 'calendar_name', 'title', 'start_time', 'end_time', 'duration', 'is_all_day'];
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
        start_time: event_start,
        end_time: event_end,
        duration: start_day === end_day ? start_day : start_day + ' - ' + end_day,
        //month: formatMonth(event_start),
        is_all_day: event.isAllDayEvent() ? 1 : 0,
        toShortString: function(){
          return '[' + this.start_time.toDateString() + ' - ' + this.end_time.toDateString() + ']' + ' | ' + this.title;
        }
      }
    }
  );
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
 * @param {Date} date 
 */
function formatDate(date) {
  return '' + date.getFullYear() + '/' + (date.getMonth() + 1) + '/' + date.getDate();
}

function formatMonth(date){
  switch(date.getMonth()) {
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