/* eslint-disable prettier/prettier */
/* eslint-disable office-addins/call-sync-before-read */
/* eslint-disable office-addins/load-object-before-read */
/* eslint-disable no-undef */
/*
 * Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.
 * See LICENSE in the project root for license information.
 */

/* global document, Office, Word */

Office.onReady((info) => {
  if (info.host === Office.HostType.Word) {
    // Determine if the user's version of Office supports all the Office.js APIs that are used in the tutorial.
    if (!Office.context.requirements.isSetSupported("WordApi", "1.3")) {
      console.log("Sorry. The tutorial add-in uses Word.js APIs that are not available in your version of Office.");
    }

    // Assign event handlers and other initialization logic.
    document.getElementById("insert-paragraph").onclick = insertParagraph;
    document.getElementById("apply-style").onclick = applyStyle;
    document.getElementById("apply-custom-style").onclick = applyCustomStyle;
    document.getElementById("change-font").onclick = changeFont;
    document.getElementById("insert-text-into-range").onclick = insertTextIntoRange;
    document.getElementById("insert-text-outside-range").onclick = insertTextBeforeRange;
    document.getElementById("replace-text").onclick = replaceText;

    document.getElementById("sideload-msg").style.display = "none";
    document.getElementById("app-body").style.display = "flex";
    //document.getElementById("run").onclick = run;
  }
});

async function replaceText() {
  await Word.run(async (context) => {

    const doc = context.document;
    const originalRange = doc.getSelection();
    originalRange.insertText("many", "Replace");

      await context.sync();
  })
  .catch(function (error) {
      console.log("Error: " + error);
      if (error instanceof OfficeExtension.Error) {
          console.log("Debug info: " + JSON.stringify(error.debugInfo));
      }
  });
}

async function insertTextBeforeRange() {
  await Word.run(async (context) => {

    const doc = context.document;
    const originalRange = doc.getSelection();
    originalRange.insertText("Office 2019, ", "Before");

    originalRange.load("text");
    await context.sync();

    doc.body.insertParagraph("Current text of original range: " + originalRange.text, "End");

  })
  .catch(function (error) {
      console.log("Error: " + error);
      if (error instanceof OfficeExtension.Error) {
          console.log("Debug info: " + JSON.stringify(error.debugInfo));
      }
  });
}

async function insertTextIntoRange() {
  await Word.run(async (context) => {
    const doc = context.document;
    const originalRange = doc.getSelection();
    originalRange.insertText(" (C2R)", "End");

    originalRange.load("text");
     await context.sync();

    doc.body.insertParagraph("Original range: " + originalRange.text, "End");

    await context.sync();
  }).catch(function (error) {
    console.log("Error: " + error);
    if (error instanceof OfficeExtension.Error) {
      console.log("Debug info: " + JSON.stringify(error.debugInfo));
    }
  });
}

async function changeFont() {
  await Word.run(async (context) => {
    const secondParagraph = context.document.body.paragraphs.getFirst().getNext();
    secondParagraph.font.set({
      name: "Courier New",
      bold: true,
      size: 18,
    });
    await context.sync();
  }).catch(function (error) {
    console.log("Error: " + error);
    if (error instanceof OfficeExtension.Error) {
      console.log("Debug info: " + JSON.stringify(error.debugInfo));
    }
  });
}

async function applyCustomStyle() {
  await Word.run(async (context) => {
    const lastParagraph = context.document.body.paragraphs.getLast();
    lastParagraph.style = "MyCustomStyle1";

    await context.sync();
  }).catch(function (error) {
    console.log("Error: " + error);
    if (error instanceof OfficeExtension.Error) {
      console.log("Debug info: " + JSON.stringify(error.debugInfo));
    }
  });
}

async function applyStyle() {
  await Word.run(async (context) => {
    const firstParagraph = context.document.body.paragraphs.getFirst();
    firstParagraph.styleBuiltIn = Word.Style.intenseReference;

    await context.sync();
  }).catch(function (error) {
    console.log("Error: " + error);
    if (error instanceof OfficeExtension.Error) {
      console.log("Debug info: " + JSON.stringify(error.debugInfo));
    }
  });
}

async function insertParagraph() {
  await Word.run(async (context) => {
    const docBody = context.document.body;
    docBody.insertParagraph(
      "Office has several versions, including Office 2016, Microsoft 365 subscription, and Office on the web.",
      "Start"
    );

    await context.sync();
  }).catch(function (error) {
    console.log("Error: " + error);
    if (error instanceof OfficeExtension.Error) {
      console.log("Debug info: " + JSON.stringify(error.debugInfo));
    }
  });
}

// export async function run() {
//   return Word.run(async (context) => {
//     /**
//      * Insert your Word code here
//      */

//     // insert a paragraph at the end of the document.
//     const paragraph = context.document.body.insertParagraph("Hello World", Word.InsertLocation.end);

//     // change the paragraph color to blue.
//     paragraph.font.color = "blue";

//     await context.sync();
//   });
// }
