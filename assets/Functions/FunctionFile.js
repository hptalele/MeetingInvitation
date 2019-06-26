//var htmlBodyInvite = '<div id="tblInvite">' +
//    '< div class="inviteRow"><label for="purpose" class="inviteColTitle">Purpose</label><textarea class="inputInvite" rows="3"></textarea></div>' +
//    '<div class="inviteRow"><label for="process" class="inviteColTitle">Process</label><textarea class="inputInvite" rows="3"></textarea></div>' +
//    '<div class="inviteRow"><label for="product" class="inviteColTitle">Product</label><textarea class="inputInvite" rows="2"></textarea></div>' +
//    '</div>';

var htmlBodyInvite = '<table id="tblInvite" Style="width:400px; table-layout:fixed; border-collapse:collapse; border:Orange 2px solid; overflow:hidden; color: Black;"><tr class="inviteRow" Style="width:100%;"><th>Team Meeting Objective</th></tr><tr class="inviteRow" Style="width:100%;"><td class="inviteColTitle" Style="width:100%;padding: 5px 5px;background-color: Orange;color: white;font-size: 20px; overflow:hidden;">Purpose (Why)</td></tr><tr class="inviteRow" Style="width:100%;padding:5px;"><td Style="height:30px;">Type here...</td></tr><tr class="inviteRow" Style="width:100%;"><td class="inviteColTitle" Style="width:100%;padding: 5px 5px;background-color: Orange;color: white;font-size: 20px;">Product (Outcome)</td></tr><tr class="inviteRow" Style="width:100%;padding:5px;"><td Style="height:30px;">Type here...</td></tr></table>';
var item;

Office.initialize = function () {
    item = Office.context.mailbox.item;
    // Checks for the DOM to load using the jQuery ready function.
    $(document).ready(function () {
        // After the DOM is loaded, app-specific code can run.
        // Insert data in the top of the body of the composed 
        // item.
        //prependItemBody();
    });
}

// Helper function to add a status message to the info bar.
function statusUpdate(icon, text) {
  Office.context.mailbox.item.notificationMessages.replaceAsync("status", {
    type: "informationalMessage",
    icon: icon,
    message: text,
    persistent: false
  });
}

function defaultStatus(event) {
  statusUpdate("icon16" , "Hello World!");
}

function prependItemBody() {
    item.body.getTypeAsync(
        function (result) {
            if (result.status == Office.AsyncResultStatus.Failed) {
                write(asyncResult.error.message);
            }
            else {
                // Successfully got the type of item body.
                // Prepend data of the appropriate type in body.
                if (result.value == Office.MailboxEnums.BodyType.Html) {
                    // Body is of HTML type.
                    // Specify HTML in the coercionType parameter
                    // of prependAsync.
                    item.body.prependAsync(
                        '<b>Greetings!</b>' + htmlBodyInvite,
                        {
                            coercionType: Office.CoercionType.Html,
                            asyncContext: { var3: 1, var4: 2 }
                        },
                        function (asyncResult) {
                            if (asyncResult.status == Office.AsyncResultStatus.Failed) {
                                write(asyncResult.error.message);
                            }
                            else {
                                // Successfully prepended data in item body.
                                // Do whatever appropriate for your scenario,
                                // using the arguments var3 and var4 as applicable.
                            }
                        });
                }
                //else {
                //    // Body is of text type. 
                //    item.body.prependAsync(
                //        'Greetings!',
                //        {
                //            coercionType: Office.CoercionType.Text,
                //            asyncContext: { var3: 1, var4: 2 }
                //        },
                //        function (asyncResult) {
                //            if (asyncResult.status == Office.AsyncResultStatus.Failed) {
                //                write(asyncResult.error.message);
                //            }
                //            else {
                //                // Successfully prepended data in item body.
                //                // Do whatever appropriate for your scenario,
                //                // using the arguments var3 and var4 as applicable.
                //            }
                //        });
                //}
            }
        });

}

// Writes to a div with id='message' on the page.
function write(message) {
    document.getElementById('message').innerText += message;
    $("td").css("color: Black");
}