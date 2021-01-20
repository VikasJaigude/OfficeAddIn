/*
 * Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.
 * See LICENSE in the project root for license information.
 */

 /* global document, Office */
if (!window.Promise) {
  window.Promise = Office.Promise;
}
// let bootstrapToken;

//Office.onReady(info => {
Office.initialize = function () { 
  mailboxItem = Office.context.mailbox.item;

    console.log('Template addin');
    $(document).ready(function () {
        // After the DOM is loaded, add-in-specific code can run.
        document.getElementById("set-subject-data").addEventListener("click", setSubjectData);
        document.getElementById("set-body-data").addEventListener("click", setBodyData);
    
       setCustomHeaders();
  });

};

function setCustomHeaders() {
  Office.context.mailbox.item.internetHeaders.setAsync(
    { "x-custom-addin": "smart" },
    function(asyncResult) {
      if (asyncResult.status === Office.AsyncResultStatus.Succeeded) {
        console.log("Successfully set headers");
      } else {
        console.log("Error setting headers: " + JSON.stringify(asyncResult.error));
      }
    }
  );
}

function setBodyData() {
  Office.context.mailbox.item.setSelectedDataAsync($('#txtBody').val(), function(asyncResult) {
    if (asyncResult.status === Office.AsyncResultStatus.Succeeded) {
      console.log("Successfully added to body");
    } else {
      console.error(asyncResult.error);
    }
  });

}

function setSubjectData() {
  Office.context.mailbox.item.subject.setAsync($('#txtSubject').val(), function(asyncResult) {
    if (asyncResult.status === Office.AsyncResultStatus.Succeeded) {
      console.log("Successfully added to subject");
    } else {
      console.error(asyncResult.error);
    }
  });
}

