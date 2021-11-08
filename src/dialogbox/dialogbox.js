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
    console.log(arg.message)
    const para = document.createElement("input");
    const para2 = document.createElement("label");
    const node = document.createTextNode("This is new.");
    para2.appendChild(onde);

    const element = document.getElementById("emailList");
    element.appendChild(para);
}

