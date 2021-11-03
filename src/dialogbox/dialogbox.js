Office.onReady().then(()=> {
        //console.log(Office.context.ui)
        console.log(Office.context.ui.addHandlerAsync)
        //Office.context.ui.addHandlerAsync(
          //  Office.EventType.DialogParentMessageReceived,
            //onMessageFromParent);     
    });

function onMessageFromParent(arg){
    console.log(arg.message)
    document.getElementById('ID').style.display = 'none';
}

