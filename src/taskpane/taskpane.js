/*
 * Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.
 * See LICENSE in the project root for license information.
 */

/* global document, Office */

Office.onReady((info) => {
  if (info.host === Office.HostType.Outlook) {
    document.getElementById("sideload-msg").style.display = "none";
    document.getElementById("app-body").style.display = "flex";

    if (!localStorage.getItem("showedFRE")) {
      showFirstRunExperience();
    }

    document.getElementById("run").onclick = run;
    document.getElementById("fre-video-button").onclick = showFre;
  }
});

export async function showFirstRunExperience() {
  document.getElementById("first-run-experience").style.display = "flex";
  localStorage.setItem("showedFRE", true);
}

export async function showFre() {
  Office.context.ui.displayDialogAsync('https://localhost:3000/fre-video.html', { height: 30, width: 20, displayInIframe: true });
}

export async function run() {
  // Get a reference to the current message
  const item = Office.context.mailbox.item;

  // Write message property value to the task pane
  document.getElementById("item-subject").innerHTML = "<b>Subject:</b> <br/>" + item.subject;
}
