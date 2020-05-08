/*
 * Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.
 * See LICENSE in the project root for license information.
 */
/* global console, document, Excel, Office */

// The initialize function must be run each time a new page is loaded
Office.initialize = () => {
  document.getElementById("sideload-msg").style.display = "none";
  document.getElementById("app-body").style.display = "flex";
  document.getElementById("run").onclick = run;
};

async function run() {
  try {
    await Excel.run(async context => {
      /**
       * Insert your Excel code here
       */
      const range = context.workbook.getSelectedRange();

      // Update the fill color
      range.format.fill.color = "#19b1a3";
      range.values = [["Welcome to the ServerlessDays Amsterdam"]];
      range.format.autofitColumns();

      await context.sync();
    });
  } catch (error) {
    console.error(error);
  }
}
