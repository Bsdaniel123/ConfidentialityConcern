/* eslint-disable no-undef */
/*
 * Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.
 * See LICENSE in the project root for license information.
 */

// images references in the manifest
import "../../assets/icon-16.png";
import "../../assets/icon-32.png";
import "../../assets/icon-80.png";

("use strict");
/* global document, Office */

Office.onReady((info) => {
  if (info.host === Office.HostType.Outlook) {
    document.getElementById("sideload-msg").style.display = "none";
    document.getElementById("app-body").style.display = "flex";
    document.getElementById("submitconcern").onclick = submitconcern;
    document.getElementById("reviewconcern").onclick = reviewconcern;
    document.getElementById("Codeofethics").onclick = Codeofethics;
  }
});

export async function reviewconcern() {
  var fromRecip = Office.context.mailbox.item.from;
  if (fromRecip.emailAddress === undefined) {
    window.open("https://goteam.avanade.com/", "_blank");
  } else {
    window.open("https://www.accenture.com/us-en/about/company/business-ethics");
  }
}

export async function Codeofethics() {
  var fromRecip = Office.context.mailbox.item.from;
  if (fromRecip.emailAddress === undefined) {
    window.open("https://goteam.avanade.com/", "_blank");
  } else {
    window.open("https://www.accenture.com/us-en/about/company/business-ethics");
  }
}

export async function submitconcern() {
  // Get a reference to the current message
  // Write message property value to the task pane
  window.location.href = "commands.html";
}
