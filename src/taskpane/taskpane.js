/*
 * Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.
 * See LICENSE in the project root for license information.
 */

/* global document, Office, Word */

Office.onReady(info => {
  if (info.host === Office.HostType.Word) {
    document.getElementById("sideload-msg").style.display = "none";
    document.getElementById("app-body").style.display = "flex";
    document.getElementById("run").onclick = run;
  }
});

export async function run() {
  return Word.run(async context => {
    /**
     * Insert your Word code here
     */

    // insert a paragraph at the end of the document.
    let document = context.document.body;
    document.insertParagraph("Hello wod", Word.InsertLocation.end);
    const paragraph = document.insertParagraph("aea", Word.InsertLocation.end);
    setTimeout(async () => {
      document.clear();
      await context.sync();
      console.log("aea brys");
    }, 3000);

    // change the paragraph color to blue.
    paragraph.font.color = "blue";

  });
}