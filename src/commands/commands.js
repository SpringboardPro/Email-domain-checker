/* 
© 2021 Springboard Pro Ltd.
Version 1.0.0
Author: Hamish Atkins
*/

let dialog
let recipients
let allRecipientData
let item
let sendEvent

Office.onReady(() => {
  // Initialise Office JS
})

/**
 * Function that is run when the send button is pressed by the user.
 * @param {object} event - The email send event that is to be controlled.
 */
function openDialog (event) {
  //  Get email compose information from Outlook (using promises since they are asynchronous functions).
  item = Office.context.mailbox.item

  // Verify if the composed item is an appointment or message.
  let promise1
  let promise2
  let promise3
  let promise4
  if (item.itemType === Office.MailboxEnums.ItemType.Appointment) {
    promise1 = getToEmailsAppointment()
    promise2 = getCCEmails_appointment()
    // Use promises to ensure required and optional attendees have been fetched.
    promise4 = Promise.all([promise1, promise2]).then(function (result) {
      allRecipientData = result
      return allRecipientData
    })
  } else {
    promise1 = getToEmails()
    promise2 = getCCEmails()
    promise3 = getBCCEmails()
    // Use promises to ensure bcc, cc and to recipients have been fetched.
    promise4 = Promise.all([promise1, promise2, promise3]).then(function (result) {
      allRecipientData = result
      return allRecipientData
    })
  }

  //  Check if multiple external recipients are present to decide to display dialog box.
  promise4.then(function (result) {
    sendEvent = event
    const multipleExternalBool = checkMultipleExternal(processEmails(allRecipientData))
    if (!multipleExternalBool) {
      event.completed({ allowEvent: true })
    } else {
      //  Display dialog box (callback function in dialog is to create event handler in host page to recieve info from dialog page).
      const url = 'https://springboardpro.github.io/Email-domain-checker/src/dialogbox/dialogbox.html'
      Office.context.ui.displayDialogAsync(url, { height: 50, width: 50, displayInIframe: true },
        function (asyncResult) {
          //  If dialog failed to open (probably popup blocker) then do 'dialogClosed' function.
          console.log(asyncResult.status)
          if (asyncResult.status === Office.AsyncResultStatus.Failed) {
            Office.context.ui.closeContainer()
            // If dialog box already open, close dialog and do not send email.
            if (asyncResult.error.code === 12007) {
              dialog.addEventHandler(Office.EventType.DialogMessageReceived, sendEmailwithUpdatedRecipients)
            }
            event.completed({ allowEvent: false })
          } else {
            dialog = asyncResult.value
            //  Once dialog box has sent message to confirm it is ready. Send dialog box the recipient emails.
            dialog.addEventHandler(Office.EventType.DialogMessageReceived, sendEmailsToDialog)
            //  If dialog  sends event (probably user closes), then do not send the email.
            dialog.addEventHandler(Office.EventType.DialogEventReceived, dialogClosed)
          }
        })
    }
  })
}

/**
 * Function that prevents the message from being sent when the dialog is closed without correct input from the user.
 */
function dialogClosed () {
  sendEvent.completed({ allowEvent: false })
}

/**
 * Function that sends the initial recipient data to the dialog box.
 * @param {object} arg - The message object that is passed from the host to the dialog that contains the emails .
 */
function sendEmailsToDialog (arg) {
  if (JSON.parse(arg.message).messageType === 'initialise') {
    dialog.messageChild(JSON.stringify(allRecipientData.concat(Office.context.mailbox.item.itemType)), { targetOrigin: '*' })
    dialog.addEventHandler(Office.EventType.DialogMessageReceived, sendEmailwithUpdatedRecipients)
  }
}

/**
 * Function that sends the message with the updated recipients from the checkbox form. The message doesn't send if  there are no recipients.
 * @param {object} arg - A message object from the dialog that contains selected recipient data from the checkbox form.
 */
function sendEmailwithUpdatedRecipients (arg) {
  $(window).bind('resize', function (e) { dialog.close() })
  const message = JSON.parse(arg.message)
  // If checkbox form results recieved from dialog, send with selected recipients otherwise do not send email and close dialog.
  if (message.messageType === 'form_output') {
    if ((message.toRecipients.length + message.ccRecipients.length + message.bccRecipients) === 0) {
      dialog.close()
      sendEvent.completed({ allowEvent: false })
    } else {
      setRecipients(message.toRecipients, message.ccRecipients, message.bccRecipients)
      sendEvent.completed({ allowEvent: true })
      dialog.close()
    }
  } else if (message.messageType === 'cancel') {
    dialog.close()
    sendEvent.completed({ allowEvent: false })
  }
}

/**
 * Function that updates the 'to' and 'cc' (or 'required' and 'optional') fields in Outlook.
 * @param {object} toRecipients - Object that contains the 'to' or 'required' recipients data.
 * @param {object} ccRecipients - Object that contains the 'cc' or 'optional' recipients data.
 */
function setRecipients (toRecipients, ccRecipients, bccRecipients) {
  // Local objects to point to recipients of either the appointment or message that is being composed.
  // bccRecipients applies to only messages, not appointments.
  let RecipientsTo, RecipientsCC, RecipientsBCC
  item = Office.context.mailbox.item
  // Verify if the composed item is an appointment or message.
  if (item.itemType === Office.MailboxEnums.ItemType.Appointment) {
    RecipientsTo = item.requiredAttendees
    RecipientsCC = item.optionalAttendees
  } else {
    RecipientsTo = item.to
    RecipientsCC = item.cc
    RecipientsBCC = item.bcc
  }

  // Use asynchronous method setAsync to set each type of recipients
  // of the composed item. Each time, this example passes a set of
  // names and email addresses to set, and an anonymous
  // callback function that doesn't take any parameters.
  RecipientsTo.setAsync(toRecipients,
    function (asyncResult) {
      if (asyncResult.status === Office.AsyncResultStatus.Failed) {
        write(asyncResult.error.message)
      } else {
        // Async call to set to-recipients of the item completed.
      }
    }) // End to setAsync.

  // Set any cc-recipients.
  RecipientsCC.setAsync(ccRecipients,
    function (asyncResult) {
      if (asyncResult.status === Office.AsyncResultStatus.Failed) {
        write(asyncResult.error.message)
      } else {
        // Async call to set cc-recipients of the item completed.
      }
    }) // End cc setAsync.
  
  if (item.itemType !== Office.MailboxEnums.ItemType.Appointment) {
    // Set any cc-recipients.
    RecipientsBCC.setAsync(bccRecipients,
      function (asyncResult) {
        if (asyncResult.status === Office.AsyncResultStatus.Failed) {
          write(asyncResult.error.message)
        } else {
          // Async call to set cc-recipients of the item completed.
        }
      }) // End bcc setAsync.
  }
}

/**
 * A function that gets the 'to' recipient data from the email.
 */
function getToEmails () {
  return new Promise(function (resolve, reject) {
    try {
      Office.context.mailbox.item.to.getAsync(function (asyncResult) {
        resolve(asyncResult.value)
      })
    }
    catch (error) {
      reject('Error')
    }
  })
}

/**
 * A function that gets the 'cc' recipient data from the email.
 */
function getCCEmails () {
  return new Promise(function (resolve, reject) {
    try {
      Office.context.mailbox.item.cc.getAsync(function (asyncResult) {
        resolve(asyncResult.value)
      })
    }
    catch (error) {
      reject('Error')
    }
  })
}

/**
 * A function that gets the 'bcc' recipient data from the email.
 */
function getBCCEmails () {
  return new Promise(function (resolve, reject) {
    try {
      Office.context.mailbox.item.bcc.getAsync(function (asyncResult) {
        resolve(asyncResult.value)
      })
    }
    catch (error) {
      reject('Error')
    }
  })
}

/**
 * A function that gets the 'required' recipient data from the meeting request.
 */
function getToEmailsAppointment() {
  return new Promise(function (resolve, reject) {
    try {
      Office.context.mailbox.item.requiredAttendees.getAsync(function (asyncResult) {
        resolve(asyncResult.value)
      })
    }
    catch (error) {
      reject('Error')
    }
  })
}

/**
 * A function that gets the 'optional' recipient data from the meeting request.
 */
function getCCEmails_appointment () {
  return new Promise(function (resolve, reject) {
    try {
      Office.context.mailbox.item.optionalAttendees.getAsync(function (asyncResult) {
        resolve(asyncResult.value)
      })
    }
    catch (error) {
      reject('Error')
    }
  })
}

/**
 * Function that a recipient data object and returns just the email addresses as an array.
 * @param {array} result - An array containing the recipient data.
 */
function processEmails (result) {
  // Combine cc and to recipients if needed.
  let recipientData
  if (result.length > 2) {
    recipientData = result[0].concat(result[1]).concat(result[2])
  } else if (result.length > 1) {
    recipientData = result[0].concat(result[1])
  } else {
    recipientData = result[0]
  }
  // Add email address information to a list.
  let emails = []
  for (let i = 0; i < recipientData.length; i++) {
    let Email = recipientData[i].emailAddress
    emails.push(Email)
  }
  return emails
}

/**
 * Function that returns a boolean value based on if the number of external emails is larger than
 * @param {array} emails - An array containing the emails to be checked.
 */
function checkMultipleExternal (emails) {
  // Create list of external emails.
  let externalEmails = []
  for (let i = 0; i < emails.length; i++) {
    let domain = emails[i].slice(emails[i].indexOf('@'), emails[i].length)
    if (domain !== '@springboard.pro') {
      externalEmails.push(domain)
    }
  }
  // Return true if number of unique external domains is more than 1.
  const numberExternalDomains = new Set(externalEmails).size
  return (numberExternalDomains > 1)
}
