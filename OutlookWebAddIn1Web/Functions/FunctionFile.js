// Loads the Office.js library.
Office.onReady();

// Helper function to add a status message to the notification bar.
function statusUpdate(icon, text, event) {
  const details = {
    type: Office.MailboxEnums.ItemNotificationMessageType.InformationalMessage,
    icon: icon,
    message: text,
    persistent: false
  };
  Office.context.mailbox.item.notificationMessages.replaceAsync("status", details, { asyncContext: event }, asyncResult => {
    const event = asyncResult.asyncContext;
    event.completed();
  });
}
// Displays a notification bar.
//function defaultStatus(event) {
//  statusUpdate("icon16" , "Hello World!", event);
//}

function basicICS(event) {
    const item = Office.context.mailbox.item;
    const subject = item.subject || "New Event";
    const now = new Date();
    const start = new Date(now.getTime() + 15 * 60000);
    const end = new Date(now.getTime() + 45 * 60000);

    const formatDate = (date) => {
        return date.toISOString().replace(/[-:]/g, '').split('.')[0] + 'Z';
    };

    const uid = `${Date.now()}@yourdomain.com`; // Unique ID for the event

    const icsContent = `BEGIN:VCALENDAR
VERSION:2.0
PRODID:-//YourCompany//YourApp//EN
METHOD:PUBLISH
BEGIN:VEVENT
UID:${uid}
SUMMARY:${subject}
DTSTART:${formatDate(start)}
DTEND:${formatDate(end)}
LOCATION:Auto-generated
DESCRIPTION:Created from email: ${subject}
STATUS:CONFIRMED
SEQUENCE:0
TRANSP:OPAQUE
BEGIN:VALARM
TRIGGER:-PT10M
ACTION:DISPLAY
DESCRIPTION:Reminder
END:VALARM
END:VEVENT
END:VCALENDAR`;

    const blob = new Blob([icsContent], { type: 'text/calendar;charset=utf-8' });
    const url = window.URL.createObjectURL(blob);

    const anchor = document.createElement('a');
    anchor.href = url;
    anchor.download = 'event.ics';
    document.body.appendChild(anchor);
    anchor.click();
    document.body.removeChild(anchor);
    window.URL.revokeObjectURL(url);


    statusUpdate("icon16", "...ICS file ready...!", event); //Notification
}
function defaultStatus(event) {
    basicICS(event);
    event.completed();
}



// Maps the function name specified in the manifest to its JavaScript counterpart.
Office.actions.associate("defaultStatus", defaultStatus);