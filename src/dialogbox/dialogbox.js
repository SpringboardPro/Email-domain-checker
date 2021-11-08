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
        /*
    console.log(arg.message)
    const para = document.createElement("p");
    const node = document.createTextNode("This is new.");
    para.appendChild(node);

    const element = document.getElementById("dummyElement");
    element.appendChild(para);
        */
    //DO THIS ON THE PREVIOUS PAGE???
    for (let i = 0; i < 3; i++) { 
            var x = document.createElement("INPUT");
            x.setAttribute("type", "checkbox");
            x.setAttribute("id", "email4")
            x.setAttribute("value", "email4")

            var y = document.createElement("LABEL");
            y.setAttribute("for", "email4")    
            y.innerHTML = "Email 4";


            const element = document.getElementById("dummyElement");
            element.appendChild(x);
            element.appendChild(y)
    }
}

