/*
 * Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.
 * See LICENSE in the project root for license information.
 */

/* global document, Office, Word */
//const PONCTUATION = [" ", ",", ".", "!", "?"];
//var currentWord = "";
Office.onReady((info) => {
  if (info.host === Office.HostType.Word) {
    document.getElementById("sideload-msg").style.display = "none";
    document.getElementById("app-body").style.display = "flex";
    document.getElementById("selection").onclick = select;
    document.getElementById("lastWord").onclick = lastWord;
    /*
    // Get the current Word document object
    // Get the current selection object
    const selection = document.getSelection();
    // Get the range object for the current paragraph
    const paragraphRange = selection.getRange("Paragraph");
    // Get the text of the current paragraph
    Console.log(`The current paragraph is: ${paragraphRange}`);
    */
  }
});

export function lastWord() {
  return Word.run(function (context) {
    /*
    // Retrieve the first text range in the document
    var textRange = context.document.body.getRange("Whole");

    // Create a new binding for the text range
    var binding = context.document.bindings;

    // Set up an event handler for the BindingDataChanged event
    binding.addHandlerAsync(Office.EventType.BindingDataChanged, onDataChanged);

    // Function to handle the BindingDataChanged event
    function onDataChanged(eventArgs) {
      // Get the binding data
      var bindingData = binding.getDataAsync(function (result) {
        if (result.status === "succeeded") {
          // Log the binding data to the console
          console.log("User input: " + result.value);
        }
      });
    }
    // Synchronize the document state with the Office.js runtime
    return context.sync();*/
  });
}

export async function select() {
  return Word.run(async function (context) {
    var selection = context.document.getSelection();
    context.load(selection);
    var html = selection.getHtml();
    await context.sync();
    document.getElementById("mySelection").innerHTML = html.value; //Get the selected text in HTML
  });
}
