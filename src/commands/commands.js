/*
 * Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.
 * See LICENSE in the project root for license information.
 */

/* global global, Office, self, window */

Office.onReady(() => {
  // If needed, Office.js is ready to be called
});

var dialog
var recipients
var all_recipient_data
var item
var send_event

function openDialog(event) {
  //get email compose information from Outlook (using promised since they are asynchronous functions)
  var promise1 = getToEmails();
  var promise2 = getCCEmails();
  var promise3 = Promise.all([promise1, promise2]).then(function(result){
    all_recipient_data = result
    recipients = getRecipients(result)
    return recipients
  })

  //check if recipients are only internal or not
  promise3.then(function(result){
    console.log(all_recipient_data)
    send_event = event
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
          //Once dialog box has sent message to confirm it is ready. Send dialog box the recipient emails
          dialog.addEventHandler(Office.EventType.DialogMessageReceived, sendEmailsToDialog);
        });
    }
  })

};


function sendEmailsToDialog(arg){
  if (JSON.parse(arg.message).messageType == 'initialise') {
    dialog.messageChild(JSON.stringify(all_recipient_data), { targetOrigin: "*" })
    dialog.addEventHandler(Office.EventType.DialogMessageReceived, sendEmailwithUpdatedRecipients);}
}

function sendEmailwithUpdatedRecipients(arg){
  var message = JSON.parse(arg.message)
  if (message.messageType == 'form_output'){
    setRecipients(message.toRecipients, message.ccRecipients)
    dialog.close()
    console.log(send_event)
    send_event.completed({allowEvent: false});
  }
}

function setRecipients(toRecipients, ccRecipients) {
    // Local objects to point to recipients of either
    // the appointment or message that is being composed.
    // bccRecipients applies to only messages, not appointments.
    var Recipients_to, Recipients_cc;
    item = Office.context.mailbox.item;
    // Verify if the composed item is an appointment or message.
    if (item.itemType == Office.MailboxEnums.ItemType.Appointment) {
        Recipients_to = item.requiredAttendees;
        Recipients_cc = item.optionalAttendees;
    }
    else {
        Recipients_to = item.to;
        Recipients_cc = item.cc;
    }
    
    // Use asynchronous method setAsync to set each type of recipients
    // of the composed item. Each time, this example passes a set of
    // names and email addresses to set, and an anonymous 
    // callback function that doesn't take any parameters. 
    Recipients_to.setAsync(toRecipients,
        function (asyncResult) {
            if (asyncResult.status == Office.AsyncResultStatus.Failed){
                write(asyncResult.error.message);
            }
            else {
                // Async call to set to-recipients of the item completed.

            }    
    }); // End to setAsync.


    // Set any cc-recipients.
    Recipients_cc.setAsync(ccRecipients,
        function (asyncResult) {
            if (asyncResult.status == Office.AsyncResultStatus.Failed){
                write(asyncResult.error.message);
            }
            else {
                // Async call to set cc-recipients of the item completed.
            }
    }); // End cc setAsync.
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
