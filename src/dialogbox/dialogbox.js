Office.onReady().then(()=> {
        //console.log(Office.context.ui)
        Office.context.ui.messageParent("READY")
        //OFFICE MIGHT NOT BE READY BY THE TIME IT TRIES TO SEND INFORMATION TO THE DIALOG- GET DIALOG TO SEND BACK FIRST THAT IT IS READY THEN USE meesageChild
        console.log('in dialog')
        Office.context.ui.addHandlerAsync(
            Office.EventType.DialogParentMessageReceived,
            onMessageFromParent);     
    });

function onMessageFromParent(arg){
    console.log(arg.message)
    document.getElementById('ID').style.display = 'none';
}

