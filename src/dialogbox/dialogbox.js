Office.onReady().then(()=> {
        console.log(Office.context.ui)
        //Office.context.ui.messageParent("Hello from the dialog")
        //OFFICE MIGHT NOT BE READY BY THE TIME IT TRIES TO SEND INFORMATION TO THE DIALOG- GET DIALOG TO SEND BACK FIRST THAT IT IS READY THEN USE meesageChild
        console.log('in dialog')    
        
        //Recieve emails from host page
          Office.context.ui.addHandlerAsync(
                Office.EventType.DialogParentMessageReceived,
                onMessageFromParent);     
       
      
         
        
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

function onMessageFromParent(arg){
    console.log('hello')
}

