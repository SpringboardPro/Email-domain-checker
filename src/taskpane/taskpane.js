/*
 * Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.
 * See LICENSE in the project root for license information.
 */

// images references in the manifest
import "../../assets/icon-16.png";
import "../../assets/icon-32.png";
import "../../assets/icon-80.png";

/* global document, Office */

Office.onReady((info) => {
  if (info.host === Office.HostType.Outlook) {
    document.getElementById("sideload-msg").style.display = "none";
    document.getElementById("app-body").style.display = "flex";
    document.getElementById("run").onclick = run;
  }
});

export async function run() {
  //var url = new URI('taskpane.html').absoluteTo(window.location.origin).toString();
  var url2 = window.location.origin+'/dialogbox.html'
  console.log(window.location)
  Office.context.ui.displayDialogAsync(url2);
  console.log(window.location)
}
