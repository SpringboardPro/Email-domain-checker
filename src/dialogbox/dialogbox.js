/* © 2021 Springboard Pro Ltd. */

Office.onReady().then(() => {
  //  Office JS in the dialog might not be initiallised by the time the host tries to send the email data so send a confirmation message to confirm it is ready.
  var decoyEmail
  Office.context.ui.messageParent(JSON.stringify({ messageType: 'initialise', message: 'Dialog is ready' }))
  
  //  Recieve emails from host page.
  Office.context.ui.addHandlerAsync(
    Office.EventType.DialogParentMessageReceived, createEmailCheckBoxList)

  //  Get form results from dialog box.
  getFormValues = function () {
    const toValues = Array.from(document.querySelectorAll("input[type='checkbox']:checked.toCheckBox")).map(item => JSON.parse(item.name))
    const ccValues = Array.from(document.querySelectorAll("input[type='checkbox']:checked.ccCheckBox")).map(item => JSON.parse(item.name))
    if (toValues.some(e => e.emailAddress === decoyEmail)) {
      document.getElementById('warning').style.display = 'block'
    } else {
      document.getElementById('warning').style.display = 'none'
      const selectedEmails = { messageType: 'form_output', toRecipients: toValues, ccRecipients: ccValues }
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
  const recipientsTo = unstringifiedMessage[0]
  const recipientsCC = unstringifiedMessage[1]
  decoyEmail = createDecoyEmail(unstringifiedMessage)
  console.log(decoyEmail)
  recipientsTo.splice(Math.floor(Math.random() * (recipientsTo.length + 1)), 0, { displayName: 'Deselect This', emailAddress: decoyEmail, recipientType: 'other' })
  if (recipientsTo.length > 0) {
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
    const toListTitle = document.getElementById('toListTitle')
    toListTitle.style.display = 'none'
    const toListContainer = document.getElementById('toContainer')
    toListContainer.style.display = 'none'
  }

  if (recipientsCC.length > 0) {
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
    const ccListTitle = document.getElementById('ccListTitle')
    ccListTitle.style.display = 'none'
    const ccListContainer = document.getElementById('ccContainer')
    ccListContainer.style.display = 'none'
    const ccEmailList = document.getElementById('ccEmailList')
    ccEmailList.style.display = 'none'
  }
}

function createDecoyEmail (unstringifiedEmails) {
  const emails = unstringifiedEmails[0].concat(unstringifiedEmails[1])
  let i = math.floor(Math.random() * (emails.length + 1))
  let domain = '@springboard.pro'
  while (domain === '@springboard.pro') {
    i = math.floor(Math.random() * (emails.length + 1))
    domain = emails[i].slice(emails[i].indexOf('@'), emails[i].length)
    console.log(domain)
  }
  const name = emails[i].slice(0, emails[i].indexOf('@'))
  return name + '@springboard.pro'
}
