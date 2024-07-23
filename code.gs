/**
 * This script generates agenda notes for 1:1 meetings.
 * It takes the content from a 1:1 meeting template and add it to documents
 * based on a dictionary.
 * It search for recurrences in the calendar and if there is a meeting scheduled,
 * the document is update with the notes from the template
 */

// Set up the required constants
const CALENDAR_ID = 'replace for your calendar id'; 
const TEMPLATE_DOC_ID = 'replace for your template doc id';

// Array of objects with the event id and doc id for each one for reports.
// eventId can be obtained using the identifier from the event url and pass it through a base64 decode 
// like https://www.base64decode.org/ but taking only the first part, before _time
const AGENDA_EVENTS = [
  {
    eventId: 'replace for eventId', 
    targetDoc: 'replace for agenda doc for that event'
  }
];


// Main function
function run() {
  const eventsToProcess = getAgendaEventsInNextWeekEvents();

  eventsToProcess.forEach(event => {
    addMeetingNotesToDoc(event.calendarEvent, event.settings.targetDoc)
  })
}

// Function to add meeting notes to the target document
function addMeetingNotesToDoc(event, targetDocId) {
  const targetDoc = DocumentApp.openById(targetDocId);
  const targetBody = targetDoc.getBody();

  // Index to insert elements in the doc based on the last meeting - secont title
  let index = findLastMeetingTitle(DocumentApp.ParagraphHeading.HEADING2,targetDocId)

  // Insert Title
  const formattedDate = Utilities.formatDate(event.getStartTime(),'GMT','yyyy-MM-dd');
  const titleParagraph = `${formattedDate} | ${event.getTitle()}`;
  const title = targetBody.insertParagraph(index++,titleParagraph).setHeading(DocumentApp.ParagraphHeading.HEADING2);
  const notesParagraph = targetBody.insertParagraph(index++, 'Notes');

  // Get notes template
  const templateDoc = DocumentApp.openById(TEMPLATE_DOC_ID);
  const notesBody = templateDoc.getBody();

  // Find the list in the template document
  const lists = notesBody.getListItems();

  // Iterate through each list item and copy with formatting to the destination document
  lists.forEach(listItem => {
    // Get the text and attributes of the list item
    const text = listItem.getText();
    const itemAttributes = listItem.getAttributes();

    const newItem = targetBody.insertListItem(index++, listItem.getText());

    // Apply attributes to the new paragraph in the destination document
    //applyAttributes(newItem, attributes);
    newItem.setAttributes(itemAttributes)
  });

  Logger.log(`Agenda for meeting on ${formattedDate} for ${event.getTitle()} was added`)
}

function applyAttributes(element, attributes) {
  // Apply each attribute to the element
  if (attributes) {
    Logger.log(attributes)
    if (attributes.NESTING_LEVEL) {
      element.setNestingLevel(attributes.NESTING_LEVEL);
    }
    if (attributes.BOLD) {
      element.setItalic(attributes.BOLD);
    }
    //if (attributes.GLYPH_TYPE) {
    //  element.setGlyphType(attributes.GLYPH_TYPE);
    //}
  }
}

// Function to identify the index where to start adding new notes
function findLastMeetingTitle(searchCriteria, targetDocId) {
  const targetDoc = DocumentApp.openById(targetDocId);
  const body = targetDoc.getBody();
  const elements = body.getNumChildren();

  // start searching after the first element because we normally have an stack section
  for (let i = 1; i < elements; i++) {
    let item = body.getChild(i);
    if (item.getHeading() == searchCriteria) {
      return i;
    }
  }
}

// Find agenda events to process
function getAgendaEventsInNextWeekEvents() {
  const now = new Date();
  const startOfNextWeek = new Date(now.getFullYear(), now.getMonth(), now.getDate() + (7 - now.getDay()));
  const endOfNextWeek = new Date(startOfNextWeek.getFullYear(), startOfNextWeek.getMonth(), startOfNextWeek.getDate() + 7);

  const nextWeekEvents = CalendarApp.getCalendarById(CALENDAR_ID).getEvents(startOfNextWeek, endOfNextWeek);

  const foundEvents = [];

  // Iterate over nextWeekEvents and find matching agendaEvents
  for (let nextWeekEvent of nextWeekEvents) {
    for (let agendaEvent of AGENDA_EVENTS) {
      if (nextWeekEvent.getId().startsWith(agendaEvent.eventId)) {
        // If the event was declined by the organizer, ignore it
        const creator = nextWeekEvent.getGuestList(true)[0]
        const isEventDeclined = (creator.getGuestStatus() !== 'NO')
        if (isEventDeclined) {
          Logger.log(`Skiping event ${nextWeekEvent.getTitle()} because creator declined`)
          continue;
        }
        
        foundEvents.push({
          settings: agendaEvent,
          calendarEvent: nextWeekEvent
        });
        break; // Break inner loop once a match is found
      }
    }
  }

  Logger.log(`Found ${foundEvents.length} events to process`)

  return foundEvents;
}
