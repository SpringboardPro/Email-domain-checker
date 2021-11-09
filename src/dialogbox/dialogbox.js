Office.onReady().then(()=> {
        
        Office.context.ui.messageParent(JSON.stringify({messageType:'initialise', message: "Dialog is ready"}))
        //OFFICE MIGHT NOT BE READY BY THE TIME IT TRIES TO SEND INFORMATION TO THE DIALOG- GET DIALOG TO SEND BACK FIRST THAT IT IS READY THEN USE meesageChild
          
        
        //Recieve emails from host page
        Office.context.ui.addHandlerAsync(
                Office.EventType.DialogParentMessageReceived,
                createEmailCheckBoxList);     
       
      get_form_values = function(){
                var toValues = Array.from(document.querySelectorAll("input[type='checkbox']:checked.toCheckBox")).map(item => JSON.parse(item.name))
                var ccValues = Array.from(document.querySelectorAll("input[type='checkbox']:checked.ccCheckBox")).map(item => JSON.parse(item.name))
                if (toValues.some(e => e.emailAddress =="deselect.this@springboard.pro")) {
                        document.getElementById("warning").style.display = "block";
                } else{
                        document.getElementById("warning").style.display = "none";
                        console.log('SEND IT')
                        let selected_emails = {messageType: 'form_output', toRecipients: toValues, ccRecipients: ccValues}
                        Office.context.ui.messageParent(JSON.stringify(selected_emails))
                }
      }
         
        
        //Get form results from dialog box
        

                
        
        
    });

function createEmailCheckBoxList(arg){
     
    unstringified_message = JSON.parse(arg.message)
    to_recipients = unstringified_message[0]
    cc_recipients = unstringified_message[1]
    //to_recipients = unstringified_message.toRecipients
    //cc_recipients = unstringified_message.ccRecipients
    
    
    to_recipients.splice(Math.floor(Math.random()*(to_recipients.length+1)),0,{displayName: 'Deselect This', emailAddress: 'deselect.this@springboard.pro', recipientType: 'other'})
        
    if (to_recipients.length > 0){
        for (let i = 0; i < to_recipients.length; i++) { 
                
            $('#toContainer').append(
                $(document.createElement('input')).prop({
                    id: 'email'+String(i),
                    name: JSON.stringify(to_recipients[i]),
                    class: 'toCheckBox',
                    type: 'checkbox'
                })
            ).append(
                $(document.createElement('label')).prop({
                    for: 'email'+String(i)
                }).html(String(to_recipients[i].emailAddress))
                ).append(document.createElement('br'));
                
                }
    } else {
        var to_list_title = document.getElementById('toListTitle')
        to_list_title.style.display = "none";
        var to_list_container = document.getElementById('toContainer')
        to_list_container.style.display = "none";
    }
        
    if(cc_recipients.length >0){
       for (let i = 0; i < cc_recipients.length; i++) { 
               
            $('#ccContainer').append(
                $(document.createElement('input')).prop({
                    id: 'email'+String(i),
                    name: JSON.stringify(cc_recipients[i]),
                    class: 'ccCheckBox',
                    type: 'checkbox'
                })
            ).append(
                $(document.createElement('label')).prop({
                    for: 'email'+String(i)
                }).html(String(cc_recipients[i].emailAddress))
                ).append(document.createElement('br'));
                
                }
       } else {
             var cc_list_title = document.getElementById('ccListTitle')
             cc_list_title.style.display = "none";
             var cc_list_container = document.getElementById('ccContainer')
             cc_list_container.style.display = "none";
    }
        
}
