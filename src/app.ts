/*
 * Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.
 * See LICENSE in the project root for license information.
 */

(() => {
  // The initialize function must be run each time a new page is loaded
  Office.initialize = (reason) => {
    $(document).ready(() => {
      $('#run').click(run);
    });
  };

  async function run() {
    
    
    var videos = (<any>(<Office.Types.ItemRead>Office.context.mailbox.item).getRegExMatches()).JiraCasenumber;
    alert(videos);
    
    /**
     * Insert your Outlook code here
     */
    
  }
})();
