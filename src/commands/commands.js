/*
 * Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.
 * See LICENSE in the project root for license information.
 */

/* global global, Office, self, window */

Office.onReady(() => {
  // If needed, Office.js is ready to be called
});

var dialog

function openDialog(event) {
  //get email compose information from Outlook (using promised since they are asynchronous functions)
  var promise1 = getToEmails();
  var promise2 = getCCEmails();
  var promise3 = Promise.all([promise1, promise2]).then(function(result){
    return recipients = getRecipients(result)
  })

  //check if recipients are only internal or not
  promise3.then(function(result){
    var internal_bool = (check_if_internal(result.toRecipients) && check_if_internal(result.ccRecipients))
    if (internal_bool){
      event.completed({allowEvent: true});
    } else {
      //event.completed({allowEvent: false});
      //display dialog box (callback function in dialog is to create event handler in host page to recieve info from dialog page)
      var url ='https://hamish-atkins-sb.github.io/Email-domain-checker/src/dialogbox/dialogbox.html'
      
      Office.context.ui.displayDialogAsync(url, {height: 50, width: 50, displayInIframe: true},
        function (asyncResult) {

          dialog = asyncResult.value;
          dialog.addEventHandler(Office.EventType.DialogMessageReceived, processMessage(recipients));
          //console.log(JSON.stringify(recipients));
          
                    
        });
    }
  })

};


function processMessage(arg,recipients){
  console.log(recipients)
  dialog.messageChild("hello from host", { targetOrigin: "*" })
}


function getToEmails() {
  return new Promise(function (resolve, reject) {
      try {
        Office.context.mailbox.item.to.getAsync(function (asyncResult) {
              resolve(asyncResult.value);
          });
      }
      catch (error) {
          reject('Error');
      }
  })
}

function getCCEmails() {
  return new Promise(function (resolve, reject) {
      try {
        Office.context.mailbox.item.cc.getAsync(function (asyncResult) {
              resolve(asyncResult.value);
          });
      }
      catch (error) {
          reject('Error');
      }
  })
}

//gets 'to' and 'cc' recipients and returns as an object
function getRecipients(result){
  var toRecipients = processEmails(result[0])
  var ccRecipients = processEmails(result[1])
  return {toRecipients, ccRecipients}
}

//gets emails using Outlook API and formats and returns in an array
function processEmails(result){
  var emails = new Array()
  for (var i = 0; i < result.length; i++) {
    var Email = result[i].emailAddress
    emails.push(Email)
  }
  return emails
}



//function that takes list of email address domains and returns boolean value based on if any external domains are present
function check_if_internal(emails){
  if (emails.length == 0){
    var SendBool = true;
  } 
  else{
    for (var i = 0; i < emails.length; i++) {
      if ((emails[i].slice(emails[i].indexOf('@'),emails[i].length))=='@springboard.pro'){
        var SendBool = true;
      }
      else {
        var SendBool = false;
        break;
      }
    }
  }
  return SendBool;
}
