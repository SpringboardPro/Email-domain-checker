Office.onReady().then(()=> {
        //console.log(Office.context.ui)
        Office.context.ui.messageParent("READY")
        //OFFICE MIGHT NOT BE READY BY THE TIME IT TRIES TO SEND INFORMATION TO THE DIALOG- GET DIALOG TO SEND BACK FIRST THAT IT IS READY THEN USE meesageChild
        console.log('in dialog')
        form.addEventListener('submit', e => {
                e.preventDefault();

                  const values = Array.from(document.querySelectorAll('input[type=checkbox]:checked'))
                    .map(item => item.value)
                    .join(',');

                console.log(`?${values}`);
                });
});
        
        
        
        Office.context.ui.addHandlerAsync(
            Office.EventType.DialogParentMessageReceived,
            onMessageFromParent);     
    });

function onMessageFromParent(arg){
    console.log(arg.message)
    document.getElementById('ID').style.display = 'none';
}

