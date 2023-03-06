/*
 * Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.
 * See LICENSE in the project root for license information.
 *
 * Tutorial's link : https://learn.microsoft.com/fr-fr/office/dev/add-ins/tutorials/word-tutorial
 */

/* global document, Office, Word */

Office.onReady((info) => {
  if (info.host === Office.HostType.Word) {
    document.getElementById("sideload-msg").style.display = "none";
    document.getElementById("app-body").style.display = "flex";
    // document.getElementById("run").onclick = run;

    // Determine if the user's version of Office supports all the Office.js APIs that are used in the tutorial.
    if (!Office.context.requirements.isSetSupported("WordApi", "1.3")) {
      console.log("Sorry. The tutorial add-in uses Word.js APIs that are not available in your version of Office.");
    }

    // Assign event handlers and other initialization logic.

    // Add event handler for document selection change event
    Office.context.document.addHandlerAsync(
      Office.EventType.DocumentSelectionChanged,
      detectSelection,
      function (result) {
        if (result.status !== Office.AsyncResultStatus.Succeeded) {
          console.log("Could not add event handler: " + result.error.message);
        }
      }
    );

    // Office.context.document.addHandlerAsync(
    //   Office.EventType.ContentControlAdded,
    //   detectWrittenText,
    //   function (result) {
    //     if (result.status !== Office.AsyncResultStatus.Succeeded) {
    //       console.log("Could not add event handler: " + result.error.message);
    //     }
    //   }
    // );
  }
});

function detectSelection() {
  var selection = Office.context.document.getSelectedDataAsync(Office.CoercionType.Text, function (result) {
    if (result.status !== Office.AsyncResultStatus.Succeeded) {
      console.log("Error: " + result.error.message);
    } else {
      var selectedText = result.value;
      document.getElementById("userText").innerHTML = selectedText;
      var listContainer = document.getElementById("AlternativesList"); // retrieves the element from the dedicated element
      listContainer.innerHTML = ""; // Erase previous content
      var list = createAlternativesList(); // create new list - Method to adapt with PhonoWriter source code...
      listContainer.appendChild(list); // Display (add) the list
    }
  });
}

// //DOESN'T WORK CORRECTLY FOR NOW
// function detectWrittenText(eventArgs) {
//   // Retrieves the text of the last added content control element
//   var contentControl = eventArgs.contentControl;
//   contentControl.addHandlerAsync(Office.EventType.Input, function (eventArgs) {
//       var text = eventArgs.target.text;
//       console.log(text);
//       // display the written text in dedicated text zone
//       document.getElementById("userWrittenText").innerHTML = text;
//     });
// }

function createAlternativesList() {
  var list = document.createElement("ul");
  var items = ["Alternative1", "Alternative2", "Alternative3"]; // List 'example'
  for (var i = 0; i < items.length; i++) {
    var li = document.createElement("li");
    li.appendChild(document.createTextNode(items[i]));
    // add listener to replace user's selection by selected word from the list (item)
    li.addEventListener("dblclick", function () {
      // document.getElementById("userText").innerHTML = this.innerHTML;
      replaceText(this.innerHTML);
    });
    list.appendChild(li);
  }
  return list;
}

async function replaceText(newText) {
  await Word.run(async (context) => {
    const doc = context.document;
    const originalRange = doc.getSelection();
    originalRange.insertText(newText, "Replace");

    await context.sync();
  }).catch(function (error) {
    console.log("Error: " + error);
    if (error instanceof OfficeExtension.Error) {
      console.log("Debug info: " + JSON.stringify(error.debugInfo));
    }
  });
}
