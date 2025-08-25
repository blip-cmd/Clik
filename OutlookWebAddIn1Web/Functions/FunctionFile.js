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
async function defaultStatus(event) {
    try {
        const ics = await aiICS(event);
        console.log("Received ICS content:", ics);
    } catch (error) {
        console.error("Error in defaultStatus:", error);
        statusUpdate("icon16", "Error processing email", event);
    } finally {
        event.completed();
    }
}

function aiICS(event) {
    const item = Office.context.mailbox.item;
    const msgBody = item.body; // Get the email body content

    // Updated prompt logic as per new requirements
    const prompt = [
        'You are an expert assistant that extracts calendar event details from emails and formats them into valid ICS (iCalendar) strings.',
        '',
        'Your job:',
        '- Generate one or more well-formed ICS events based on the email content provided.',
        '- For each event, include ONLY these fields:',
        '  - SUMMARY',
        '  - DTSTART (in UTC, with "Z" suffix)',
        '  - DTEND (if available; otherwise omit)',
        '  - LOCATION (if available; otherwise use "Auto-generated")',
        '  - DESCRIPTION (summarize the email content)',
        '  - UID (must be unique per event; use a timestamp or hash)',
        '  - STATUS (always CONFIRMED)',
        '  - SEQUENCE (always 0)',
        '  - TRANSP (always OPAQUE)',
        '  - VALARM (10-minute display reminder)',
        '',
        'Wrap all events in a single VCALENDAR block with:',
        '  - VERSION:2.0',
        '  - PRODID:-//YourCompany//YourApp//EN',
        '  - METHOD:PUBLISH',
        '',
        'Smart NLP Expectations:',
        '- Parse natural language time expressions (e.g., "next Friday at 2pm", "Aug 20 at 11:30 a.m.") into UTC timestamps',
        '- If no explicit time is given, default to 00:00 UTC on the date mentioned',
        '- If no end time is given, omit DTEND entirely',
        '',
        'Multi-Event Support:',
        '- If multiple events are described in the email, generate a separate VEVENT block for each',
        '- Ensure each VEVENT has a unique UID and correct DTSTART/DTEND values',
        '',
        'Email content:',
        '"""',
        msgBody,
        '"""',
        '',
        'Example output:',
        'BEGIN:VCALENDAR',
        'VERSION:2.0',
        'PRODID:-//YourCompany//YourApp//EN',
        'METHOD:PUBLISH',
        'BEGIN:VEVENT',
        'UID:1756059623998@yourdomain.com',
        'SUMMARY:New recommendations available on ScienceDirect',
        'DTSTART:20250824T183523Z',
        'DTEND:20250824T190523Z',
        'LOCATION:Auto-generated',
        'DESCRIPTION:Created from email: New recommendations available on ScienceDirect',
        'STATUS:CONFIRMED',
        'SEQUENCE:0',
        'TRANSP:OPAQUE',
        'BEGIN:VALARM',
        'TRIGGER:-PT10M',
        'ACTION:DISPLAY',
        'DESCRIPTION:Reminder',
        'END:VALARM',
        'END:VEVENT',
        'END:VCALENDAR'
    ].join('\n');

    const apiKey = "AIzaSyD20df_EXFV7qIaZi-wHk-9IWNPhXyQB30"; //dev's API key

    return fetch(`https://generativelanguage.googleapis.com/v1beta/models/gemini-1.5-flash:generateContent?key=${apiKey}`, {
        method: "POST",
        headers: {
            "Content-Type": "application/json"
        },
        body: JSON.stringify({
            contents: [{ parts: [{ text: prompt }] }]
        })
    })
        .then(res => res.json())
        .then(data => {
            let ics = data.candidates?.[0]?.content?.parts?.[0]?.text;
            console.log("Raw AI response:\n" + ics);
            
            if (ics) {
                // Remove markdown code blocks (```ics...``` or ```...```)
                ics = ics.replace(/```[\w]*\n?/g, '').trim();
                
                // Ensure it starts with BEGIN:VCALENDAR and ends with END:VCALENDAR
                if (!ics.startsWith('BEGIN:VCALENDAR')) {
                    const startIndex = ics.indexOf('BEGIN:VCALENDAR');
                    if (startIndex !== -1) {
                        ics = ics.substring(startIndex);
                    }
                }
                
                if (!ics.endsWith('END:VCALENDAR')) {
                    const endIndex = ics.lastIndexOf('END:VCALENDAR');
                    if (endIndex !== -1) {
                        ics = ics.substring(0, endIndex + 'END:VCALENDAR'.length);
                    }
                }
                
                console.log("Cleaned ICS content:\n" + ics);
                dub(ics);
                statusUpdate("icon16", "ICS file downloaded successfully!", event);
                return ics; // Return the clean ICS content
            } else {
                statusUpdate("icon16", "Failed to generate ICS file", event);
                throw new Error("No ICS content generated");
            }
        })
        .catch(err => {
            console.error("Error calling Gemini API:", err);
            statusUpdate("icon16", "Error generating ICS file", event);
            throw err; // Re-throw to be caught by defaultStatus
        });
}

function dub(ics) {
    // Download ICS content as a file
    if (!ics || typeof ics !== 'string') {
        console.error('Invalid ICS content provided to dub function');
        return;
    }

    try {
        // Create a blob with the ICS content
        const blob = new Blob([ics], { type: 'text/calendar;charset=utf-8' });
        const url = window.URL.createObjectURL(blob);

        // Create a temporary anchor element for download
        const anchor = document.createElement('a');
        anchor.href = url;
        
        // Generate filename with timestamp
        const timestamp = new Date().toISOString().replace(/[:.]/g, '-').split('T')[0];
        anchor.download = `calendar-event-${timestamp}.ics`;
        
        // Add to DOM, click, and remove
        document.body.appendChild(anchor);
        anchor.click();
        document.body.removeChild(anchor);
        
        // Clean up the object URL
        window.URL.revokeObjectURL(url);
        
        console.log('ICS file download initiated successfully');
    } catch (error) {
        console.error('Error downloading ICS file:', error);
    }
}

// Maps the function name specified in the manifest to its JavaScript counterpart.
Office.actions.associate("defaultStatus", defaultStatus);