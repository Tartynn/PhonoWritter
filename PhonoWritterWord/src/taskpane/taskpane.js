/*
 * Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.
 * See LICENSE in the project root for license information.
 */

/* global document, Office, Word */

Office.onReady((info) => {
  if (info.host === Office.HostType.Word) {
    document.getElementById("sideload-msg").style.display = "none";
    document.getElementById("app-body").style.display = "flex";
    document.getElementById("run").onclick = run;
    Console.log("fdp");
    // Get the current Word document object
    // Get the current selection object
    const selection = document.getSelection();
    // Get the range object for the current paragraph
    const paragraphRange = selection.getRange("Paragraph");
    // Get the text of the current paragraph
    Console.log(`The current paragraph is: ${paragraphRange}`);
  }
});

export async function run() {
  return Word.run(async (context) => {
    /**
     * Insert your Word code here
     */
    // insert a paragraph at the end of the document.
    //context.document.body.paragraphs.
    const paragraph = context.document.body.insertParagraph("Hello World", Word.InsertLocation.end);
    // change the paragraph color to blue.
    paragraph.font.color = "blue";

    await context.sync();
  });
}
