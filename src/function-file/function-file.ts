(() => {

  let config: any;
  let settingsDialog: any;

  // The initialize function must be run each time a new page is loaded
  Office.initialize = () => {
    config = getConfig();
  };

  // Add any ui-less function here
  function showError(error) {
    const mailItem = Office.context.mailbox.item as any;
    mailItem.notificationMessages.replaceAsync("github-error", {
      message: error,
      type: "errorMessage",
    }, (result) => {
      // do stuff
    });
  }

  function insertJiraLink(event) {
    // Check if the add-in has been configured
    if (config && config.jiraUrl) {
      // do stuff
    } else {
      const url = new URI("../settings/dialog.html?warn=1").absoluteTo(window.location.href).toString();
      const dialogOptions = { width: 20, height: 40 };

      Office.context.ui.displayDialogAsync(url, dialogOptions, (result) => {
        settingsDialog = result.value;
        settingsDialog.addEventHandler(Office.EventType.DialogMessageReceived, receiveMessage);
        settingsDialog.addEventHandler(Office.EventType.DialogEventReceived, dialogClosed);
        event.completed();
      });
  }
  }

  function receiveMessage(message) {
    config = JSON.parse(message.message);
    setConfig(config, (result) => {
      settingsDialog.close();
      settingsDialog = null;
    });
  }

  function dialogClosed(message) {
    settingsDialog = null;
  }
})();
