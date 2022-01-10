/* © 2021 Springboard Pro Ltd. */

var dialog
var recipients
var all_recipient_data
var item
var send_event


Office.onReady(() => {
  // Initialise Office JS
});

/**
 * Function that is run when the send button is pressed by the user.
 * @param {object} event - The email send event that is to be controlled.
 */
function openDialog(event) {
  //Get email compose information from Outlook (using promises since they are asynchronous functions).
  
  item = Office.context.mailbox.item;
 
    // Verify if the composed item is an appointment or message.
  if (item.itemType == Office.MailboxEnums.ItemType.Appointment){
    var promise1 = getToEmails_appointment();
    var promise2 = getCCEmails_appointment();
  } else {
    var promise1 = getToEmails();
    var promise2 = getCCEmails();
  }
  
  
  var promise3 = Promise.all([promise1, promise2]).then(function(result){
    all_recipient_data = result
    recipients = getRecipients(result)
    return recipients
  })

  //Check if recipients are only internal or not.
  promise3.then(function(result){
    console.log(all_recipient_data)
    send_event = event
    var multiple_external_bool = check_multiple_external(result.toRecipients, result.ccRecipients) 
    
    if (!multiple_external_bool){
      event.completed({allowEvent: true});
    } else {
      //Display dialog box (callback function in dialog is to create event handler in host page to recieve info from dialog page).
      var url ='https://springboardpro.github.io/Email-domain-checker/src/dialogbox/dialogbox.html'
      console.log(Office.context.ui)
      Office.context.ui.displayDialogAsync(url, {height: 50, width: 50, displayInIframe: true}, 
        function (asyncResult) {
            //If dialog failed to open (probably popup blocker) then do 'dialogClosed' function.
              if (asyncResult.status === Office.AsyncResultStatus.Failed) {
                 Office.context.ui.closeContainer()
                 console.log(asyncResult.error.message)
                 console.log(asyncResult.value)
                 if (asyncResult.error.code == 12007){
                    dialog.addEventHandler(Office.EventType.DialogMessageReceived, sendEmailwithUpdatedRecipients);
                 }
                 event.completed({allowEvent: false});
            } else {
                dialog = asyncResult.value;
                console.log(dialog)
                window.addEventListener("beforeunload", function(event) {console.log('unloading2')});
                //Once dialog box has sent message to confirm it is ready. Send dialog box the recipient emails
                dialog.addEventHandler(Office.EventType.DialogMessageReceived, sendEmailsToDialog);
                //If dialog  sends event (probably user closes), then do 'dialogClosed' function.
                dialog.addEventHandler(Office.EventType.DialogEventReceived, dialogClosed);
      };
    })
  }})
};

/**
 * Function that prevents the message from being sent when the dialog is closed without correct input from the user.
 */
function dialogClosed(){
  send_event.completed({allowEvent: false})
}

/**
 * Function that sends the initial recipient data to the dialog box.
 * @param {object} arg - The message object that is passed from the host to the dialog that contains the emails .
 */
function sendEmailsToDialog(arg){
  if (JSON.parse(arg.message).messageType == 'initialise') {
    dialog.messageChild(JSON.stringify(all_recipient_data), { targetOrigin: "*" })
    dialog.addEventHandler(Office.EventType.DialogMessageReceived, sendEmailwithUpdatedRecipients);}
}

/**
 * Function that sends the message with the updated recipients from the checkbox form. The message doesn't send if  there are no recipients.
 * @param {object} arg - A message object from the dialog that contains selected recipient data from the checkbox form.
 */
function sendEmailwithUpdatedRecipients(arg){
  $(window).bind('resize', function(e){dialog.close()});
  var message = JSON.parse(arg.message)
  if (message.messageType == 'form_output'){
    if ((message.toRecipients.length + message.ccRecipients.length) == 0){
       dialog.close()
       send_event.completed({allowEvent: false});
    } else{
      setRecipients(message.toRecipients, message.ccRecipients)
      send_event.completed({allowEvent: true});
      dialog.close()
    }
    
    
  } else if (message.messageType == 'cancel') {
      dialog.close()
      send_event.completed({allowEvent: false});
  }
  
}


/**
 * Function that updates the 'to' and 'cc' (or 'required' and 'optional') fields in Outlook.
 * @param {object} toRecipients - Object that contains the 'to' or 'required' recipients data.
 * @param {object} ccRecipients - Object that contains the 'cc' or 'optional' recipients data.
 */
function setRecipients(toRecipients, ccRecipients) {
    // Local objects to point to recipients of either the appointment or message that is being composed.
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

/**
 * A function that gets the 'to' recipient data from the email.
 */
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

/**
 * A function that gets the 'cc' recipient data from the email.
 */
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

/**
 * A function that gets the 'required' recipient data from the meeting request.
 */
function getToEmails_appointment() {
  return new Promise(function (resolve, reject) {
      try {
        Office.context.mailbox.item.requiredAttendees.getAsync(function (asyncResult) {
              resolve(asyncResult.value);
          });
      }
      catch (error) {
          reject('Error');
      }
  })
}

/**
 * A function that gets the 'optional' recipient data from the meeting request.
 */
function getCCEmails_appointment() {
  return new Promise(function (resolve, reject) {
      try {
        Office.context.mailbox.item.optionalAttendees.getAsync(function (asyncResult) {
              resolve(asyncResult.value);
          });
      }
      catch (error) {
          reject('Error');
      }
  })
}

/**
 * Function that takes the two categories of recipient data and stores the emails of each in an object together.
 * @param {array} result - An array containing the 'to' and 'cc' recipient data.
 */
function getRecipients(result){
  var toRecipients = processEmails(result[0])
  var ccRecipients = processEmails(result[1])
  return {toRecipients, ccRecipients}
}


/**
 * Function that a recipient data object and returns just the email addresses as an array.
 * @param {array} result - An array containing the recipient data.
 */
function processEmails(result){
  var emails = new Array()
  for (var i = 0; i < result.length; i++) {
    var Email = result[i].emailAddress
    emails.push(Email)
  }
  return emails
}



/**
 * Function that returns a boolean value based on if any passed emails are external or not.
 * @param {array} emails - An array containing the emails to be checked.
 */
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

function check_multiple_external(to_emails, cc_emails){
  var emails = to_emails.concat(cc_emails)
  var external_emails = []
  for (var i = 0; i < emails.length; i++) {
    var domain = emails.slice(emails[i].indexOf('@'), emails[i].length)
    if (emails.slice != '@springboard.pro'){
      external_emails.push(domain)
    }
  }
  print(external_emails)
  number_external_domains = new Set(external_emails).size;
  console.log(number_external_domains)
  if (number_external_domains > 1){
    return true
  } else {
  return false
  }
} 
