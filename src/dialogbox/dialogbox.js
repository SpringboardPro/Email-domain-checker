Office.onReady().then(()=> {
        Office.context.ui.messageParent("Dialog is ready")
        //OFFICE MIGHT NOT BE READY BY THE TIME IT TRIES TO SEND INFORMATION TO THE DIALOG- GET DIALOG TO SEND BACK FIRST THAT IT IS READY THEN USE meesageChild
          
        
        //Recieve emails from host page
        Office.context.ui.addHandlerAsync(
                Office.EventType.DialogParentMessageReceived,
                createEmailCheckBoxList);     
       
      get_form_values = function(){
                //const toForm = document.querySelector('toEmailList');
               // const toValues = Array.from(document.querySelector('toCheckBox').checked).map(item => item.value).join(',');
                var toCheckedBoxes = []
                var toValues = document.querySelectorAll("input[type='checkbox']:checked.toCheckBox")
                for (let i = 0; i < toValues.length; i++) {
                        if (toValues[i].checked) {
                              toCheckedBoxes.push(toValues[i])
                        }
                }
                console.log(toValues)
              /*
                //const ccForm = document.querySelector('ccEmailList');
                const ccValues = Array.from(document.querySelector('ccCheckBox').checked).map(item => item.value).join(',');
                console.log(`${ccValues}`);
               */
      
      }
         
        
        //Get form results from dialog box
        

                
        
        
    });

function createEmailCheckBoxList(arg){
     
    unstringified_message = JSON.parse(arg.message)
    to_recipients = unstringified_message.toRecipients
   
    cc_recipients = unstringified_message.ccRecipients
    
    
    if (to_recipients.length > 0){
        for (let i = 0; i < to_recipients.length; i++) { 
                
            $('#toContainer').append(
                $(document.createElement('input')).prop({
                    id: 'email'+String(i),
                    name: String(to_recipients[i]),
                    class: 'toCheckBox',
                    type: 'checkbox'
                })
            ).append(
                $(document.createElement('label')).prop({
                    for: 'email'+String(i)
                }).html(String(to_recipients[i]))
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
                    name: String(cc_recipients[i]),
                    class: 'ccCheckBox',
                    type: 'checkbox'
                })
            ).append(
                $(document.createElement('label')).prop({
                    for: 'email'+String(i)
                }).html(String(cc_recipients[i]))
                ).append(document.createElement('br'));
                
                }
       } else {
             var cc_list_title = document.getElementById('ccListTitle')
             cc_list_title.style.display = "none";
             var cc_list_container = document.getElementById('ccContainer')
             cc_list_container.style.display = "none";
    }
        
}

        

