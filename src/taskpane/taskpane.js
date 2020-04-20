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
    document.querySelector("#clearDocument").onclick = clearDocument;
    document.querySelector("#test").onclick = test;
  }
});

function test() {
  return Word.run(async context => {
    let document = context.document.body;
    console.log("funcion: ", document);
    document.insertTable();
    await context.sync();
  });
}

function clearDocument() {
  return Word.run(async context => {
    let document = context.document.body;
    document.clear();

  });
}

export async function run() {
  return Word.run(async context => {
    let document = context.document.body;
    document.insertParagraph("Hello wod", Word.InsertLocation.end);
    console.log("testeando");
    const paragraph = document.insertParagraph("123123", Word.InsertLocation.end);
    await context.sync();
    paragraph.font.color = "red";

  });
}