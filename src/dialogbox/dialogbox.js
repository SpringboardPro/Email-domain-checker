/* 
© 2021 Springboard Pro Ltd.
Version 1.0.0
Author: Hamish Atkins
*/

Office.onReady().then(() => {
  //  Office JS in the dialog might not be initiallised by the time the host tries to send the email data so send a confirmation message to confirm it is ready.
  Office.context.ui.messageParent(JSON.stringify({ messageType: 'initialise', message: 'Dialog is ready' }))
  
  //  Recieve emails from host page.
  Office.context.ui.addHandlerAsync(
    Office.EventType.DialogParentMessageReceived, createEmailCheckBoxList)

  //  Get form results from dialog box.
  getFormValues = function () {
    const selectedToValues = Array.from(document.querySelectorAll("input[type='checkbox']:checked.toCheckBox")).map(item => JSON.parse(item.name))
    const selectedCCValues = Array.from(document.querySelectorAll("input[type='checkbox']:checked.ccCheckBox")).map(item => JSON.parse(item.name))
    const selectedBCCValues = Array.from(document.querySelectorAll("input[type='checkbox']:checked.bccCheckBox")).map(item => JSON.parse(item.name))
    const toValues = Array.from(document.querySelectorAll("input[type='checkbox'].toCheckBox")).map(item => JSON.parse(item.name))
    const ccValues = Array.from(document.querySelectorAll("input[type='checkbox'].ccCheckBox")).map(item => JSON.parse(item.name))
    const bccValues = Array.from(document.querySelectorAll("input[type='checkbox'].bccCheckBox")).map(item => JSON.parse(item.name))
    // Display warning message if decoy email selected or if not all of the recipients are selected, otherwise send selected email recipients to host.
    if ((selectedToValues.some(e => e.displayName === 'Decoy email unselect')) || (selectedToValues.length !== toValues.length - 1) || (selectedCCValues.length !== ccValues.length) || (selectedBCCValues.length !== bccValues.length)) {
      document.getElementById('warning').style.display = 'block'
    } else {
      document.getElementById('warning').style.display = 'none'
      const selectedEmails = { messageType: 'form_output', toRecipients: selectedToValues, ccRecipients: selectedCCValues, bccRecipients: selectedBCCValues }
      Office.context.ui.messageParent(JSON.stringify(selectedEmails))
    }
  }
  //  Send cancel message to host if cancel button is pressed.
  cancel = function () {
    const cancelMessage = { messageType: 'cancel' }
    Office.context.ui.messageParent(JSON.stringify(cancelMessage))
  }
})

/**
 * Function that creates the check box form from the recipients.
 * @param {object} arg - The message object from the host pages that contains the recipient data.
 */
function createEmailCheckBoxList (arg) {
  const unstringifiedMessage = JSON.parse(arg.message)
  let recipientsTo
  let recipientsCC
  let recipientsBCC
  let messageType
  if (unstringifiedMessage.length > 3) {
    recipientsTo = unstringifiedMessage[0]
    recipientsCC = unstringifiedMessage[1]
    recipientsBCC = unstringifiedMessage[2]
    messageType = unstringifiedMessage[3]
  } else {
    recipientsTo = unstringifiedMessage[0]
    recipientsCC = unstringifiedMessage[1]
    messageType = unstringifiedMessage[2]
  }

  let toLabel
  let ccLabel
  let bccLabel = 'Bcc Recipients'
  if (messageType === 'appointment') {
    toLabel = 'Required Attendees'
    ccLabel = 'Optional Attendees'
  } else {
    toLabel = 'To Recipients'
    ccLabel = 'Cc Recipients'
  }
  // Create html checkbox list for 'to' recipients for email or 'required' attendees if meeting request.
  let decoyEmail = createDecoyEmail(unstringifiedMessage)
  recipientsTo.splice(Math.floor(Math.random() * (recipientsTo.length + 1)), 0, { displayName: 'Decoy email unselect', emailAddress: decoyEmail, recipientType: 'other' })
  if (recipientsTo.length > 0) {
    // Set list title depending on if email or meeting request.
    document.getElementById("toListTitle").innerHTML = toLabel 
    // Create checkbox for each email address.
    for (let i = 0; i < recipientsTo.length; i++) {
      $('#toContainer').append(
        $(document.createElement('input')).prop({
          id: 'emailTo' + String(i),
          name: JSON.stringify(recipientsTo[i]),
          class: 'toCheckBox',
          type: 'checkbox'
        })
      ).append(
        $(document.createElement('label')).prop({
          for: 'emailTo' + String(i)
        }).html(String(recipientsTo[i].emailAddress))
      ).append(document.createElement('br'))
    }
  } else {
    // Do not display list if no 'to' recipients are present.
    const toListTitle = document.getElementById('toListTitle')
    toListTitle.style.display = 'none'
    const toListContainer = document.getElementById('toContainer')
    toListContainer.style.display = 'none'
  }
  
  // Create html checkbox list for 'cc' recipients for email or 'optional' attendees if meeting request.
  if (recipientsCC.length > 0) {
    // Set list title depending on if email or meeting request.
    document.getElementById("ccListTitle").innerHTML = ccLabel
    // Create checkbox for each email address. 
    for (let i = 0; i < recipientsCC.length; i++) {
      $('#ccContainer').append(
        $(document.createElement('input')).prop({
          id: 'emailCc' + String(i),
          name: JSON.stringify(recipientsCC[i]),
          class: 'ccCheckBox',
          type: 'checkbox'
        })
      ).append(
        $(document.createElement('label')).prop({
          for: 'emailCc' + String(i)
        }).html(String(recipientsCC[i].emailAddress))
      ).append(document.createElement('br'))
    }
  } else {
    // Do not display list if no 'cc' recipients are present.
    const ccListTitle = document.getElementById('ccListTitle')
    ccListTitle.style.display = 'none'
    const ccListContainer = document.getElementById('ccContainer')
    ccListContainer.style.display = 'none'
    const ccEmailList = document.getElementById('ccEmailList')
    ccEmailList.style.display = 'none'
  }
  
// Create html checkbox list for 'bcc' recipients for email.
    if (recipientsBCC.length > 0) {
    // Set list title depending on if email or meeting request.
    document.getElementById("bccListTitle").innerHTML = 'Bcc Recipients'
    // Create checkbox for each email address. 
    for (let i = 0; i < recipientsBCC.length; i++) {
      $('#bccContainer').append(
        $(document.createElement('input')).prop({
          id: 'emailBcc' + String(i),
          name: JSON.stringify(recipientsBCC[i]),
          class: 'bccCheckBox',
          type: 'checkbox'
        })
      ).append(
        $(document.createElement('label')).prop({
          for: 'emailBcc' + String(i)
        }).html(String(recipientsBCC[i].emailAddress))
      ).append(document.createElement('br'))
    }
  } else {
    // Do not display list if no 'cc' recipients are present.
    const bccListTitle = document.getElementById('bccListTitle')
    bccListTitle.style.display = 'none'
    const bccListContainer = document.getElementById('bccContainer')
    bccListContainer.style.display = 'none'
    const bccEmailList = document.getElementById('bccEmailList')
    bccEmailList.style.display = 'none'
  }
}

/**
 * Function that creates a decoy email by taking the name part of any of the external emails and adding it to the internal domain.
 * @param {array} unstringifiedEmails - An array containing the email recipient objects.
 */
function createDecoyEmail (unstringifiedEmails) {
  const emails = unstringifiedEmails[0].concat(unstringifiedEmails[1])
  let i = Math.floor(Math.random() * (emails.length))
  let domain = '@springboard.pro'
  while (domain === '@springboard.pro') {
    i = Math.floor(Math.random() * (emails.length))
    domain = emails[i].emailAddress.slice(emails[i].emailAddress.indexOf('@'), emails[i].emailAddress.length)
  }
  const name = emails[i].emailAddress.slice(0, emails[i].emailAddress.indexOf('@'))
  return name + '@example.com'
}
