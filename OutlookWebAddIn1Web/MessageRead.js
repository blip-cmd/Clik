(function () {
  "use strict";

  let messageBanner;
  let currentCalendarLink = '';

  // Loads the Office.js library.
  Office.onReady(function (reason) {
    $(() => {
      const element = document.querySelector('.MessageBanner');
      messageBanner = new components.MessageBanner(element);
      messageBanner.hideBanner();
      
      // Auto-generate calendar link when task pane opens
      generateNewLink();
    });
  });

  // Generate a new calendar link from the current email
  window.generateNewLink = async function() {
    try {
      // Show loading state
      $('#loading-section').show();
      $('#calendar-section').hide();
      $('#error-section').hide();
      
      showNotification("Processing", "Generating calendar link...");
      
      // Call the AI function to get calendar link
      const ics = await aiICS();
      console.log("Received ICS content for task pane:", ics);
      
      if (!ics || ics.trim() === '') {
        throw new Error("Generated ICS is empty");
      }

      // Parse ICS content to extract event details
      const eventData = parseICSContent(ics);
      console.log("Parsed event data:", eventData);
      
      // Create Outlook deep link
      const outlookLink = createOutlookDeepLink(eventData);
      console.log("Generated Outlook link:", outlookLink);
      
      // Display the link in the task pane
      displayCalendarLink(outlookLink);
      
      showNotification("Success", "Calendar link generated successfully!");
      
    } catch (error) {
      console.error("Error generating calendar link:", error);
      showError(`Error: ${error.message}`);
      showNotification("Error", `Failed to generate link: ${error.message}`);
    }
  };

  // Display the calendar link in the task pane
  function displayCalendarLink(url) {
    currentCalendarLink = url;
    
    // Hide loading, show calendar section
    $('#loading-section').hide();
    $('#error-section').hide();
    $('#calendar-section').show();
    
    // Set the link text and href
    $('#calendar-link').text(url);
    $('#open-outlook-btn').attr('href', url).show();
    
    // Also log to console
    console.log("======================================");
    console.log("OUTLOOK CALENDAR LINK:");
    console.log(url);
    console.log("======================================");
    console.log("Copy the above URL to open the event in Outlook Calendar");
  }

  // Show error state
  function showError(message) {
    $('#loading-section').hide();
    $('#calendar-section').hide();
    $('#error-section').show();
    $('#error-message').text(message);
  }

  // Copy calendar link to clipboard
  window.copyCalendarLink = function() {
    if (!currentCalendarLink) {
      showStatus('No calendar link available', 'status-error');
      return;
    }
    
    // Try modern clipboard API first
    if (navigator.clipboard && navigator.clipboard.writeText) {
      navigator.clipboard.writeText(currentCalendarLink)
        .then(() => {
          showStatus('✅ Link copied to clipboard!', 'status-success');
        })
        .catch((err) => {
          console.error('Modern clipboard failed:', err);
          fallbackCopy();
        });
    } else {
      fallbackCopy();
    }
  };

  // Fallback copy method
  function fallbackCopy() {
    try {
      // Create temporary textarea
      const textarea = document.createElement('textarea');
      textarea.value = currentCalendarLink;
      textarea.style.position = 'fixed';
      textarea.style.opacity = '0';
      document.body.appendChild(textarea);
      textarea.select();
      
      const successful = document.execCommand('copy');
      document.body.removeChild(textarea);
      
      if (successful) {
        showStatus('✅ Link copied to clipboard!', 'status-success');
      } else {
        showStatus('Please select and copy the link manually', 'status-error');
      }
    } catch (err) {
      console.error('Fallback copy failed:', err);
      showStatus('Please select and copy the link manually', 'status-error');
    }
  }

  // Select all text in the link display
  window.selectLinkText = function() {
    const linkElement = document.getElementById('calendar-link');
    if (window.getSelection) {
      const selection = window.getSelection();
      const range = document.createRange();
      range.selectNodeContents(linkElement);
      selection.removeAllRanges();
      selection.addRange(range);
    }
  };

  // Show status message
  function showStatus(message, className) {
    const statusDiv = $('#status-message');
    statusDiv.text(message);
    statusDiv.attr('class', 'status-message ' + className);
    statusDiv.show();
    
    setTimeout(() => {
      statusDiv.hide();
    }, 4000);
  }

  // Helper function for displaying notifications
  function showNotification(header, content) {
    $("#notificationHeader").text(header);
    $("#notificationBody").text(content);
    messageBanner.showBanner();
    messageBanner.toggleExpansion();
  }

  // Parse ICS content - copied from FunctionFile.js
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

  // Create Outlook deep link - copied from FunctionFile.js
  function createOutlookDeepLink(eventData) {
    console.log("Creating Outlook deep link with data:", eventData);
    
    const baseUrl = 'https://outlook.office.com/calendar/action/compose';
    const params = new URLSearchParams();
    
    // Add subject
    if (eventData.summary && eventData.summary.trim()) {
        params.append('subject', eventData.summary.trim());
    } else {
        params.append('subject', 'Event from Email');
    }
    
    // Add start date/time
    if (eventData.dtstart) {
        const startDate = convertICSDateToISO(eventData.dtstart);
        if (startDate) {
            params.append('startdt', startDate);
        }
    } else {
        const now = new Date();
        params.append('startdt', now.toISOString());
    }
    
    // Add end date/time
    if (eventData.dtend) {
        const endDate = convertICSDateToISO(eventData.dtend);
        if (endDate) {
            params.append('enddt', endDate);
        }
    } else if (eventData.dtstart) {
        const startDate = convertICSDateToISO(eventData.dtstart);
        if (startDate) {
            const endDate = new Date(startDate);
            endDate.setHours(endDate.getHours() + 1);
            params.append('enddt', endDate.toISOString());
        }
    }
    
    // Add location
    if (eventData.location && eventData.location.trim()) {
        params.append('location', eventData.location.trim());
    }
    
    // Add description
    if (eventData.description && eventData.description.trim()) {
        let description = eventData.description.trim();
        if (description.length > 500) {
            description = description.substring(0, 497) + '...';
        }
        params.append('body', description);
    }
    
    const fullUrl = `${baseUrl}?${params.toString()}`;
    console.log("Generated Outlook deep link:", fullUrl);
    
    return fullUrl;
  }

  // Convert ICS date to ISO - copied from FunctionFile.js
  function convertICSDateToISO(icsDate) {
    if (!icsDate || icsDate.length < 15) {
        return null;
    }
    
    try {
        const year = icsDate.substring(0, 4);
        const month = icsDate.substring(4, 6);
        const day = icsDate.substring(6, 8);
        const hour = icsDate.substring(9, 11);
        const minute = icsDate.substring(11, 13);
        const second = icsDate.substring(13, 15);
        
        const isoDate = `${year}-${month}-${day}T${hour}:${minute}:${second}Z`;
        
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

  // AI ICS function - copied from FunctionFile.js
  async function aiICS() {
    return new Promise((resolve, reject) => {
        const item = Office.context.mailbox.item;
        
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
                "Email content:",
                '"""',
                msgBody,
                '"""'
            ].join('\n');

            const apiKey = "AIzaSyD20df_EXFV7qIaZi-wHk-9IWNPhXyQB30";

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
                
                // Clean up the response
                ics = ics.replace(/```[\w]*\n?/g, '').trim();
                
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
                
                resolve(ics);
            })
            .catch(err => {
                console.error("Error calling Gemini API:", err);
                reject(err);
            });
        });
    });
  }

})();
