const calendar = CalendarApp.getCalendarById(
  "<your-google-calendar-id>"
);
const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Travel Desk (Incoming)");

const TransportMode = {
  TRAIN: "Train",
  FLIGHT: "Flight",
  BUS: "Bus",
  CAR: "Car",
};

// Helper to combine date and time
function makeStartTime(date, time) {
  const parsedDate = new Date(date);
  const parsedTime = new Date(time);
  parsedDate.setHours(parsedTime.getHours(), parsedTime.getMinutes(), parsedTime.getSeconds());
  return parsedDate;
}

// Mark row as invalid
function markRowAsInvalid(sheet, rowIndex, reason) {
  const range = sheet.getRange(rowIndex + 1, 1, 1, sheet.getLastColumn());
  range.setBackground("red");
  sheet.getRange(rowIndex + 1, sheet.getLastColumn() + 1).setValue(reason); // Add reason in a new column
}

// Create or update a grouped calendar event
function createOrUpdateGroupedEvent(groupKey, startDateTime, transport, destination, description) {
  const endDateTime = new Date(startDateTime.getTime() + 15 * 60 * 1000); // 15-minute duration
  const searchStart = new Date(startDateTime.getTime() - 24 * 60 * 60 * 1000); // 24 hours before
  const searchEnd = new Date(startDateTime.getTime() + 24 * 60 * 60 * 1000); // 24 hours after

  const events = calendar.getEvents(searchStart, searchEnd);
  let existingEvent = events.find(event =>
    event.getDescription().includes(`Idempotence Key: ${groupKey}`)
  );

  // Determine event color
  let color = CalendarApp.EventColor.GREEN;
  if (transport.includes("Flight")) color = CalendarApp.EventColor.BLUE;
  else if (transport.includes("Car")) color = CalendarApp.EventColor.ORANGE;

  const title = `${transport} :: Arrival :: ${destination}`;
  const fullDescription = `Idempotence Key: ${groupKey}\n${description}`;

  if (existingEvent) {
    existingEvent.setTitle(title);
    existingEvent.setTime(startDateTime, endDateTime);
    existingEvent.setDescription(fullDescription);
    existingEvent.setLocation(destination);
    existingEvent.setColor(color);
    Logger.log(`Event updated: ${existingEvent.getTitle()}`);
  } else {
    const newEvent = calendar.createEvent(title, startDateTime, endDateTime, {
      description: fullDescription,
      location: destination,
    });
    newEvent.setColor(color);
    Logger.log(`Event created: ${newEvent.getTitle()}`);
  }
}

// Process sheet data and create events
function processSheetAndCombineEventsWithValidation() {
  const rows = sheet.getDataRange().getValues();
  const groupedEvents = {};

  for (let i = 1; i < rows.length; i++) {
    const [date, name, secondName, contactNo, guests, time, transportDetails, destination] = rows[i];
    if (!date || !time || !(time instanceof Date)) {
      markRowAsInvalid(sheet, i, "Invalid Date or Time");
      continue;
    }

    const startTime = makeStartTime(date, time);
    const groupKey = `${+startTime}-${destination}`;

    if (!groupedEvents[groupKey]) {
      groupedEvents[groupKey] = {
        startTime,
        destination,
        transportDetails,
        travelers: [],
      };
    }

    const fullName = secondName ? `${name}, ${secondName}` : name;
    groupedEvents[groupKey].travelers.push(`${fullName} (${guests})`);
  }

  for (const groupKey in groupedEvents) {
    const group = groupedEvents[groupKey];
    const description = `Travelers: ${group.travelers.join(", ")}`;
    createOrUpdateGroupedEvent(groupKey, group.startTime, group.transportDetails, group.destination, description);
  }
}

// Delete all calendar events
function deleteAllCalendarEvents() {
  const events = calendar.getEvents(new Date('2000-01-01'), new Date('2100-01-01'));
  events.forEach(event => event.deleteEvent());
  Logger.log(`Deleted all events.`);
}
