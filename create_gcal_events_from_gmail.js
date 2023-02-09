//forked from https://github.com/al-codaio/events-from-gmail/blob/main/create_gcal_events_from_gmail.js
//adapted by github.com/0xblurr
GMAIL_LABEL = 'EventsFromGmail'
DEFAULT_EVENT_TIME = 30
DATE_FORMAT = 'ROW' // US (m/d/y) or ROW (d/m/y)

//////////////////////////////////////////////////////////////////////////// 

function getEmail() {
  var label = GmailApp.getUserLabelByName(GMAIL_LABEL);
  var threads = label.getThreads();
  
  // Only create events for unread messages in the GMAIL_LABEL
  for (var i = 0; i < threads.length; i++) {    
    if (threads[i].isUnread()) {
      var emailSubject = threads[i].getFirstMessageSubject()
      var emailMessage = threads[i].getMessages()[0].getPlainBody()
      message = []
      message = Array.from(emailMessage)
      message_start = emailMessage.search("RESERVATION APPROVED")
      emailMessage = emailMessage.slice(message_start, message.length)
      console.log(emailMessage)
     
      //get time from email
      time_index = emailMessage.search("Time: ")
      if(emailMessage[time_index] != 'T'){
        time_index = emailMessage.search("\\*Time\\*: ")
        time = emailMessage.slice(time_index+8, time_index + 25)
      }else{
        time = emailMessage.slice(time_index+6, time_index + 25)
      }
      
      //process time chars
      if(time[1] != ':'){
        start_time = time.slice(0,5) + time.slice(6, 8)
        if(time[10] != ':'){
        end_time = time.slice(9,14) + time.slice(15, 17)
        }else{
          end_time = time.slice(9,13) + time.slice(14, 16)
        }
      }else{
        start_time = time.slice(0,4) + time.slice(5, 7)
        if(time[9] != ':'){
        end_time = time.slice(8,13) + time.slice(14, 16)
          }else{
            end_time = time.slice(8,12) + time.slice(13, 15)
          }
      }
      
      //get date from email
      date_index = (emailMessage.search("Date: "))+7
      console.log(emailMessage[date_index-7])
      date = emailMessage.slice(date_index-1, date_index + 2)
      console.log('date' + date)
      if (emailMessage[date_index-7] != 'D'){
        date_index = (emailMessage.search("\\*Date\\*: "))+8
        date = emailMessage.slice(date_index, date_index + 2)
        console.log('format')
      }else{
        date_index -= 1
      }
    
      if(emailMessage[date_index+1] != '/'){
        date = emailMessage.slice(date_index, date_index + 3)
        console.log(1)
        if(emailMessage[date_index+5] != '/'){
          date = emailMessage.slice(date_index, date_index + 4)
          console.log(2)
        }
      }if(emailMessage[date_index+4] != '/'){
        date = emailMessage.slice(date_index, date_index + 3)
        console.log(3)
      }
      console.log(date)
      date = date.split('/').reverse().join('/')
      console.log(date)
    
      court_index = emailMessage.search("P&A: ")
      eventTitle = emailMessage.slice(court_index, court_index+37)
      if(court_index == -1){
        court_index = emailMessage.search("\\*P&A\\*: ")
        eventTitle = emailMessage.slice(court_index, court_index+37)
      }

      event = calcDateTime(date,start_time,false)
      startTime = event[0]
      event1 = calcDateTime(date, end_time, false)
      endTime = event1[0]
      console.log(event + '\n' + event1)

      isAllDay = false
      optionalParams = ''
      createEvent(eventTitle, startTime, endTime, isAllDay, optionalParams) 
      threads[i].markRead()
      
    } 
  }
}

function createEvent(eventTitle, startTime, endTime, isAllDay, optionalParams) {
  var calPayload
  
  // For single-day events don't need an end date

  calPayload = [eventTitle, startTime, endTime, optionalParams]

  
  if (isAllDay) {
    var event = CalendarApp.getDefaultCalendar().createAllDayEvent(...calPayload)
  } else {
    var event = CalendarApp.getDefaultCalendar().createEvent(...calPayload)
  }
  Logger.log('Event Added: ' + eventTitle + ', ' + startTime + '(ID: ' + event.getId() + ')');
}

function calcDateTime(rawDate, rawTime, isAllDay) {
  var dashModifierIndex = rawDate.indexOf("-")
  var isMultiDay = dashModifierIndex == -1 ? false : true 
  
  // Get time if event is not an all-day event
  if (isAllDay == false && isMultiDay == false) {
    var [month, day, year] = parseDate(rawDate)
    var [hours, minutes] = convertTime12to24(rawTime)
    var newDateStart = new Date(year, month - 1, day, hours, minutes)
    var newDateEnd = new Date(newDateStart)
    newDateEnd.setMinutes(newDateStart.getMinutes() + DEFAULT_EVENT_TIME)
  }  else {
    // Set start and end date if event is multi-day
    if (isMultiDay) {
      var startDate = rawDate.substring(0, dashModifierIndex)
      var endDate = rawDate.substring(dashModifierIndex + 1, rawDate.length)
      var [startMonth, startDay, startYear] = parseDate(startDate)
      var [endMonth, endDay, endYear] = parseDate(endDate) 
      var newDateStart = new Date(startYear, startMonth - 1, startDay)
      var newDateEnd = new Date(endYear, endMonth - 1, endDay)
      isAllDay = true
    } else {
      var [month, day, year] = parseDate(rawDate)
      var newDateStart = newDateEnd = new Date(year, month - 1, day)
    }    
  }
  return [newDateStart, newDateEnd, isAllDay]
}

// Get month, day, year from date with slash
function parseDate(date) {
  var dateDetails = date.split("/")  
  const d = new Date();
  let thisYear = d.getFullYear();
  
  // US or ROW date format
  if (DATE_FORMAT == 'US') {
    var month = parseInt(dateDetails[0].trim())
    var day = parseInt(dateDetails[1].trim())  
  } else {
    var month = parseInt(dateDetails[1].trim())
    var day = parseInt(dateDetails[0].trim())
  }  
  var year = dateDetails[2] ? parseInt(dateDetails[2].trim()) : thisYear;  
  return [month, day, year]
}

// Convert AM/PM to 24-hour time
const convertTime12to24 = (time12h) => {
  var upperCaseTime = time12h.toUpperCase()
  
  // Figure out if time contains AM or PM 
  if (upperCaseTime.indexOf("AM") != -1) {
    var index = upperCaseTime.indexOf("AM")
  } else {
    var index = upperCaseTime.indexOf("PM")
  }
   
  // Get time and modifier (AM/PM) based on position of modifier
  var newTime = upperCaseTime.substring(0, index)
  var modifier = upperCaseTime.substring(index, index + 2)
  
  // Get hours/minutes depending on if there is a : in time
  if (upperCaseTime.indexOf(':') != -1) {
    var [hours, minutes] = newTime.split(':');
  } else {
    var [hours, minutes] = [newTime, 0]
  }  
  if (hours === '12') { hours = '00' }
  if (modifier === 'PM') { hours = parseInt(hours, 10) + 12 }
  return [hours, minutes];
}
