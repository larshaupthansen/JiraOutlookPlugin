/*
 * Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.
 * See LICENSE in the project root for license information.
 */

(() => {
  // The initialize function must be run each time a new page is loaded
  Office.initialize = (reason) => {
    $(document).ready(() => {
      $("#run").click(run);
    });
  };

  async function run() {
    const mailItem = Office.context.mailbox.item as Office.Types.ItemRead;
    const videos = ( mailItem as any).getRegExMatches().JiraCasenumber;
    alert(videos);
    /**
     * Insert your Outlook code here
     */
  }
})();
