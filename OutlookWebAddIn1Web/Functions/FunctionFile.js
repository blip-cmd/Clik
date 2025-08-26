// Loads the Office.js library.
Office.onReady();

// Helper function to add a status message to the notificati        // Create Outlook deep link
        const outlookLink = createOutlookDeepLink(eventData);
        console.log("Generated Outlook link:", outlookLink);
        
        // Display the link in notification and console
        displayLink(outlookLink, event);
        function statusUpdate(icon, text, event) {
  // Check if Office.context.mailbox.item is available
  if (!Office.context || !Office.context.mailbox || !Office.context.mailbox.item) {
    console.log("Status update:", text);
    return;
  }
  
  const details = {
    type: Office.MailboxEnums.ItemNotificationMessageType.InformationalMessage,
    icon: icon,
    message: text,
    persistent: true  // Make the notification persistent so user can copy the link
  };
  Office.context.mailbox.item.notificationMessages.replaceAsync("status", details, (asyncResult) => {
    // Don't call event.completed() here - only call it in the main function
    if (asyncResult.status === Office.AsyncResultStatus.Failed) {
      console.error("Failed to update status:", asyncResult.error);
    }
  });
}

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
    event.completed();
}

async function defaultStatus(event) {
    try {
        statusUpdate("icon16", "Processing email...", event);
        const ics = await aiICS();
        console.log("Received ICS content:", ics);
        
        if (!ics || ics.trim() === '') {
            statusUpdate("icon16", "Generated ICS is empty", event);
        } else {
            dub(ics); // Download the ICS file
            statusUpdate("icon16", "ICS file downloaded successfully!", event);
        }
    } catch (error) {
        console.error("Error in defaultStatus:", error);
        statusUpdate("icon16", `Error: ${error.message}`, event);
    } finally {
        // Always complete the event
        if (event && typeof event.completed === 'function') {
            event.completed();
        }
    }
}

async function webLinkWithPreview(event) {
    try {
        statusUpdate("icon16", "Generating Outlook web link...", event);
        const ics = await aiICS();
        console.log("Received ICS content for web link:", ics);
        
        if (!ics || ics.trim() === '') {
            statusUpdate("icon16", "Cannot generate link - ICS is empty", event);
            return;
        }

        // Parse ICS content to extract event details
        const eventData = parseICSContent(ics);
        console.log("Parsed event data:", eventData);
        
        // Create Outlook deep link
        const outlookLink = createOutlookDeepLink(eventData);
        console.log("Generated Outlook link:", outlookLink);
        
        // Display the link in notification and console
        displayLink(outlookLink, event);
        
    } catch (error) {
        console.error("Error in webLinkWithPreview:", error);
        statusUpdate("icon16", `Error: ${error.message}`, event);
    } finally {
        // Always complete the event
        if (event && typeof event.completed === 'function') {
            event.completed();
        }
    }
}

function displayLink(url, event) {
    // Show full URL in console with clear formatting
    console.log("======================================");
    console.log("OUTLOOK CALENDAR LINK:");
    console.log(url);
    console.log("======================================");
    console.log("Copy the above URL to open the event in Outlook Calendar");
    
    // Show link in notification with instruction
    statusUpdate("icon16", "Calendar link ready! Check browser console to copy the link.", event);
    
    // Open task pane to display the link nicely
    openTaskPane(url);
}

function openTaskPane(url) {
    try {
        if (Office.context && Office.context.ui && Office.context.ui.displayDialogAsync) {
            // Create a simple task pane URL
            const taskPaneUrl = `https://localhost:44300/CalendarLinkTaskPane.html?url=${encodeURIComponent(url)}`;
            
            Office.context.ui.displayDialogAsync(taskPaneUrl, {
                height: 70,
                width: 60,
                requireHTTPS: false
            }, (result) => {
                if (result.status === Office.AsyncResultStatus.Failed) {
                    console.error("Failed to open task pane:", result.error);
                    console.log("Task pane failed - link is available in browser console above");
                } else {
                    console.log("Task pane opened successfully with calendar link");
                }
            });
        } else {
            console.log("Task pane not available - link is available in browser console above");
        }
    } catch (error) {
        console.error("Error opening task pane:", error);
        console.log("Task pane error - link is available in browser console above");
    }
}

function parseICSContent(ics) {
    const eventData = {
        summary: '',
        dtstart: '',
        dtend: '',
        location: '',
        description: '',
        uid: ''
    };
    
    const lines = ics.split('\n');
    
    lines.forEach(line => {
        const [key, ...valueParts] = line.split(':');
        const value = valueParts.join(':').trim();
        
        switch (key) {
            case 'SUMMARY':
                eventData.summary = value;
                break;
            case 'DTSTART':
                eventData.dtstart = value;
                break;
            case 'DTEND':
                eventData.dtend = value;
                break;
            case 'LOCATION':
                eventData.location = value;
                break;
            case 'DESCRIPTION':
                eventData.description = value;
                break;
            case 'UID':
                eventData.uid = value;
                break;
        }
    });
    
    return eventData;
}

function createOutlookDeepLink(eventData) {
    console.log("Creating Outlook deep link with data:", eventData);
    
    // Try different base URLs in case one doesn't work
    const baseUrls = [
        'https://outlook.office.com/calendar/action/compose',
        'https://outlook.live.com/calendar/action/compose',
        'https://outlook.office365.com/calendar/action/compose'
    ];
    
    // Use the first URL as primary
    const baseUrl = baseUrls[0];
    
    // Prepare parameters
    const params = new URLSearchParams();
    
    // Add subject - ensure it's properly encoded
    if (eventData.summary && eventData.summary.trim()) {
        params.append('subject', eventData.summary.trim());
    } else {
        params.append('subject', 'Event from Email');
    }
    
    // Add start date/time
    if (eventData.dtstart) {
        const startDate = convertICSDateToISO(eventData.dtstart);
        console.log("Converted start date:", eventData.dtstart, "->", startDate);
        if (startDate) {
            params.append('startdt', startDate);
        }
    } else {
        // Default to current time if no start date
        const now = new Date();
        params.append('startdt', now.toISOString());
        console.log("Using default start date:", now.toISOString());
    }
    
    // Add end date/time
    if (eventData.dtend) {
        const endDate = convertICSDateToISO(eventData.dtend);
        console.log("Converted end date:", eventData.dtend, "->", endDate);
        if (endDate) {
            params.append('enddt', endDate);
        }
    } else if (eventData.dtstart) {
        // Default to 1 hour after start if no end date
        const startDate = convertICSDateToISO(eventData.dtstart);
        if (startDate) {
            const endDate = new Date(startDate);
            endDate.setHours(endDate.getHours() + 1);
            params.append('enddt', endDate.toISOString());
            console.log("Using calculated end date:", endDate.toISOString());
        }
    }
    
    // Add location
    if (eventData.location && eventData.location.trim()) {
        params.append('location', eventData.location.trim());
    }
    
    // Add description/body - limit length to avoid URL length issues
    if (eventData.description && eventData.description.trim()) {
        let description = eventData.description.trim();
        // Limit description to 500 characters to avoid URL length issues
        if (description.length > 500) {
            description = description.substring(0, 497) + '...';
        }
        params.append('body', description);
    }
    
    // Construct the full URL
    const fullUrl = `${baseUrl}?${params.toString()}`;
    
    console.log("Generated Outlook deep link:", fullUrl);
    console.log("URL length:", fullUrl.length);
    
    // Also log alternative URLs in case the primary doesn't work
    baseUrls.slice(1).forEach((altUrl, index) => {
        const altFullUrl = `${altUrl}?${params.toString()}`;
        console.log(`Alternative URL ${index + 1}:`, altFullUrl);
    });
    
    // Validate URL length (most browsers support URLs up to ~2000 characters)
    if (fullUrl.length > 2000) {
        console.warn("URL might be too long:", fullUrl.length, "characters");
    }
    
    return fullUrl;
}

function convertICSDateToISO(icsDate) {
    if (!icsDate || icsDate.length < 15) {
        return null;
    }
    
    try {
        // Parse YYYYMMDDTHHMMSSZ format
        const year = icsDate.substring(0, 4);
        const month = icsDate.substring(4, 6);
        const day = icsDate.substring(6, 8);
        const hour = icsDate.substring(9, 11);
        const minute = icsDate.substring(11, 13);
        const second = icsDate.substring(13, 15);
        
        // Construct ISO format: YYYY-MM-DDTHH:MM:SSZ
        const isoDate = `${year}-${month}-${day}T${hour}:${minute}:${second}Z`;
        
        // Validate by creating a Date object
        const date = new Date(isoDate);
        if (isNaN(date.getTime())) {
            console.error("Invalid date:", icsDate);
            return null;
        }
        
        return isoDate;
    } catch (error) {
        console.error("Error converting ICS date:", icsDate, error);
        return null;
    }
}

async function aiICS() {
    return new Promise((resolve, reject) => {
        const item = Office.context.mailbox.item;
        
        // Properly get the email body content
        item.body.getAsync(Office.CoercionType.Text, (result) => {
            if (result.status === Office.AsyncResultStatus.Failed) {
                console.error("Failed to get email body:", result.error);
                reject(new Error("Failed to get email body: " + result.error.message));
                return;
            }
            
            const msgBody = result.value;
            console.log("Email body retrieved:", msgBody ? msgBody.substring(0, 200) + "..." : "Empty body");
            
            if (!msgBody || msgBody.trim() === '') {
                reject(new Error("Email body is empty"));
                return;
            }

            // Updated prompt logic as per new requirements
            const prompt = [
                'You are an expert assistant that extracts calendar event details from emails and formats them into valid ICS (iCalendar) strings.',
                "",
                "Your job:",
                "- Generate one or more well-formed ICS events based on the email content provided.",
                "- For each event, include ONLY these fields:",
                "  - SUMMARY",
                "  - DTSTART (in UTC, with 'Z' suffix)",
                "  - DTEND (if available; otherwise omit)",
                "  - LOCATION (if available; otherwise use 'Auto-generated')",
                "  - DESCRIPTION (summarize the email content)",
                "  - UID (must be unique per event; use a timestamp or hash)",
                "  - STATUS (always CONFIRMED)",
                "  - SEQUENCE (always 0)",
                "  - TRANSP (always OPAQUE)",
                "  - VALARM (10-minute display reminder)",
                "",
                "Wrap all events in a single VCALENDAR block with:",
                "  - VERSION:2.0",
                "  - PRODID:-//YourCompany//YourApp//EN",
                "  - METHOD:PUBLISH",
                "",
                "Smart NLP Expectations:",
                "- Parse natural language time expressions (e.g., 'next Friday at 2pm', 'Aug 20 at 11:30 a.m.') into UTC timestamps",
                "- If no explicit time is given, default to 00:00 UTC on the date mentioned",
                "- If no end time is given, omit DTEND entirely",
                "",
                "Multi-Event Support:",
                "- If multiple events are described in the email, generate a separate VEVENT block for each",
                "- Ensure each VEVENT has a unique UID and correct DTSTART/DTEND values",
                "",
                "Email content:",
                '"""',
                msgBody,
                '"""',
                "",
                "Example output:",
                "BEGIN:VCALENDAR",
                "VERSION:2.0",
                "PRODID:-//YourCompany//YourApp//EN",
                "METHOD:PUBLISH",
                "BEGIN:VEVENT",
                "UID:1756059623998@yourdomain.com",
                "SUMMARY:New recommendations available on ScienceDirect",
                "DTSTART:20250824T183523Z",
                "DTEND:20250824T190523Z",
                "LOCATION:Auto-generated",
                "DESCRIPTION:Created from email: New recommendations available on ScienceDirect",
                "STATUS:CONFIRMED",
                "SEQUENCE:0",
                "TRANSP:OPAQUE",
                "BEGIN:VALARM",
                "TRIGGER:-PT10M",
                "ACTION:DISPLAY",
                "DESCRIPTION:Reminder",
                "END:VALARM",
                "END:VEVENT",
                "END:VCALENDAR"
            ].join('\n');

            const apiKey = "AIzaSyD20df_EXFV7qIaZi-wHk-9IWNPhXyQB30"; //dev's API key

            console.log("Making API call to Gemini...");
            
            fetch(`https://generativelanguage.googleapis.com/v1beta/models/gemini-1.5-flash:generateContent?key=${apiKey}`, {
                method: "POST",
                headers: {
                    "Content-Type": "application/json"
                },
                body: JSON.stringify({
                    contents: [{ parts: [{ text: prompt }] }]
                })
            })
            .then(res => {
                console.log("API response status:", res.status, res.statusText);
                if (!res.ok) {
                    throw new Error(`HTTP error! status: ${res.status} - ${res.statusText}`);
                }
                return res.json();
            })
            .then(data => {
                console.log("Full API response:", JSON.stringify(data, null, 2));
                
                // Check for API errors
                if (data.error) {
                    throw new Error(`API Error: ${data.error.message || 'Unknown API error'}`);
                }
                
                if (!data.candidates || data.candidates.length === 0) {
                    throw new Error("No candidates in API response");
                }
                
                let ics = data.candidates?.[0]?.content?.parts?.[0]?.text;
                console.log("Raw AI response:\n" + ics);
                
                if (!ics) {
                    throw new Error("No text content in API response");
                }
                
                // Remove markdown code blocks (```ics...``` or ```...```
                ics = ics.replace(/```[\w]*\n?/g, '').trim();
                
                // Ensure it starts with BEGIN:VCALENDAR and ends with END:VCALENDAR
                if (!ics.startsWith('BEGIN:VCALENDAR')) {
                    const startIndex = ics.indexOf('BEGIN:VCALENDAR');
                    if (startIndex !== -1) {
                        ics = ics.substring(startIndex);
                    } else {
                        throw new Error("No valid VCALENDAR found in response");
                    }
                }
                
                if (!ics.endsWith('END:VCALENDAR')) {
                    const endIndex = ics.lastIndexOf('END:VCALENDAR');
                    if (endIndex !== -1) {
                        ics = ics.substring(0, endIndex + 'END:VCALENDAR'.length);
                    } else {
                        throw new Error("No valid VCALENDAR end found in response");
                    }
                }
                
                console.log("Cleaned ICS content:\n" + ics);
                
                if (ics.trim() === '') {
                    throw new Error("Cleaned ICS content is empty");
                }
                
                resolve(ics); // Return the clean ICS content
            })
            .catch(err => {
                console.error("Error calling Gemini API:", err);
                reject(err);
            });
        });
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

// Test function to verify redirect functionality
function testOutlookRedirect() {
    console.log("Testing Outlook redirect...");
    
    // Create test event data
    const testEventData = {
        summary: "Test Event from Add-in",
        dtstart: "20250125T140000Z", // Today at 2 PM UTC
        dtend: "20250125T150000Z",   // Today at 3 PM UTC
        location: "Test Location",
        description: "This is a test event created by the Outlook add-in"
    };
    
    console.log("Test event data:", testEventData);
    
    // Create the deep link
    const outlookLink = createOutlookDeepLink(testEventData);
    
    // Try to open it
    const success = tryOpenLink(outlookLink);
    
    console.log("Test redirect result:", success ? "Success" : "Failed");
    
    return outlookLink;
}

// You can call this function from the browser console to test:
// testOutlookRedirect()

// Maps the function name specified in the manifest to its JavaScript counterpart.
Office.actions.associate("defaultStatus", defaultStatus);
Office.actions.associate("webLinkWithPreview", webLinkWithPreview);