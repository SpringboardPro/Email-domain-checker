Office.onReady().then(()=> {
        console.log(Office.context.ui)
        Office.context.ui.messageParent("Dialog is ready")
        //OFFICE MIGHT NOT BE READY BY THE TIME IT TRIES TO SEND INFORMATION TO THE DIALOG- GET DIALOG TO SEND BACK FIRST THAT IT IS READY THEN USE meesageChild
          
        
        //Recieve emails from host page
        Office.context.ui.addHandlerAsync(
                Office.EventType.DialogParentMessageReceived,
                createEmailCheckBoxList);     
       
      
         
        
        //Get form results from dialog box
        const form = document.querySelector('form');
        form.addEventListener('submit', e => {
                e.preventDefault();
                  const values = Array.from(document.querySelectorAll('input[type=checkbox]:checked'))
                    .map(item => item.value)
                    .join(',');
                console.log(`${values}`);
                });
        
        
    });

function createEmailCheckBoxList(arg){
     
    unstringified_message = JSON.parse(arg.message)
    to_recipients = unstringified_message.toRecipients
    console.log(to_recipients)
    cc_recipients = unstringified_message.ccRecipients
    console.log(cc_recipients)
    
    if (to_recipients.length > 0){
        for (let i = 0; i < to_recipients.length; i++) { 
                
            $('#container').append(
                $(document.createElement('input')).prop({
                    id: 'email'+String(i),
                    name: String(to_recipients[i]),
                    value: String(to_recipients[i]),
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
               
            $('#container').append(
                $(document.createElement('input')).prop({
                    id: 'email'+String(i),
                    name: String(cc_recipients[i]),
                    value: String(cc_recipients[i]),
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

        

