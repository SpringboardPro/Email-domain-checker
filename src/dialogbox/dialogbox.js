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
        
    all_recipients = create_list_of_recipients(arg)
    console.log(all_recipients)
        
    for (let i = 0; i < all_recipients.length; i++) { 
            var x = document.createElement("INPUT");
            x.setAttribute("type", "checkbox");
            x.setAttribute("id", "email"+String(i))
            x.setAttribute("value", "email"+String(i))

            var y = document.createElement("LABEL");
            y.setAttribute("for", "email"+String(i))    
            y.innerHTML = all_recipients[i];


            if (i == 0){
               var element = document.getElementById("dummyElement");
            } else{
                var element = document.getElementById("email"+String(i-1));
            }
            element.appendChild(x);
            
            
            

            //element.appendChild(y)
    }
}

function create_list_of_recipients(arg){
        
        var all_recipients = []
        recipients_object = JSON.parse(arg.message)
        for (let i = 0; i < recipients_object.ccRecipients.length; i++) { 
                all_recipients.push(recipients_object.ccRecipients[i])}
        for (let i = 0; i < recipients_object.toRecipients.length; i++) { 
                all_recipients.push(recipients_object.toRecipients[i])}
        return all_recipients
}
        
        

