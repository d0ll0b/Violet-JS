/*
 * Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.
 * See LICENSE in the project root for license information.
 */

Office.onReady(info => {
    if (info.host === Office.HostType.Word) {
        // Determine if the user's version of Office supports all the Office.js APIs that are used in the tutorial.
        if (!Office.context.requirements.isSetSupported('WordApi', '1.3')) {
            console.log('Sorry. The tutorial add-in uses Word.js APIs that are not available in your version of Office.');
        }

        // Assign event handlers and other initialization logic.
        document.getElementById("line").onclick = openwin;
        document.getElementById("sideload-msg").style.display = "none";
        document.getElementById("app-body").style.display = "flex";
  }
});

function openwin() {
    window.open ("file:///C:/Users/qwuoo/my%20office2/src/taskpane/Dialog.html", "newwindow", "height=500, width=500, top=50, left=200, toolbar=no, scrollbars=yes, resizable=yes, location=no, status=no")
}

