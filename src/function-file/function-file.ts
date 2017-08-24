/*
 * Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.
 * See LICENSE in the project root for license information.
 */

/// <reference path="../../node_modules/@types/urijs/index.d.ts" />

 
(() => {

  var config: any;
  var settingsDialog: any;

  // The initialize function must be run each time a new page is loaded
  Office.initialize = () => {
    config = getConfig();
  };

  // Add any ui-less function here
  function showError(error) {
    
    
      (<any>Office.context.mailbox.item).notificationMessages.replaceAsync('github-error', {
      type: 'errorMessage',
      message: error
    }, function(result){
    });
  }

  function insertJiraLink(event) {
    // Check if the add-in has been configured
    if (config && config.jiraUrl) {
    }
    else {
      var url = new URI('../settings/dialog.html?warn=1').absoluteTo(window.location.href).toString();
      var dialogOptions = { width: 20, height: 40 };

      Office.context.ui.displayDialogAsync(url, dialogOptions, function(result) {
        settingsDialog = result.value;
        settingsDialog.addEventHandler(Office.EventType.DialogMessageReceived, receiveMessage);
        settingsDialog.addEventHandler(Office.EventType.DialogEventReceived, dialogClosed);
        event.completed();
      });
  }
  }


  function receiveMessage(message) {
    config = JSON.parse(message.message);
    setConfig(config, function(result) {
      settingsDialog.close();
      settingsDialog = null;
    });
  }

  function dialogClosed(message) {
    settingsDialog = null;
  }
})();
