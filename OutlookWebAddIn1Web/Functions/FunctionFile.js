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
function defaultStatus(event) {
  statusUpdate("icon16" , "Hello World!", event);
}

// Maps the function name specified in the manifest to its JavaScript counterpart.
Office.actions.associate("defaultStatus", defaultStatus);