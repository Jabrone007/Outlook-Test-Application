/// <reference path="../App.js" />

(function () {
    "use strict";

    // The Office initialize function must be run each time a new page is loaded
    Office.initialize = function (reason) {
        $(document).ready(function () {
            app.initialize();

            var emailList = [["Brett Simmonds", "Dave Royce", "Dave Boster", "Toni Bennet"], ["bsimmonds@omahait.com", "droyce@omahait.com", "dboster@omahait.com", "tbennet@omahait.com"]];

            //var item = Office.cast.item.toItemRead(Office.context.mailbox.item);
            
            for (var i = 0, emailListLength = emailList[0].length; i < emailListLength; i++) {
                //$('#senderDisplayName').text = emailList[0][i];
                //$('#senderEmailAddress').text = emailList[1][i];
                //$('#senderChoose').val('Send email to ' + emailList[0][i]);
                $('#senderDisplayName').text(emailList[0][i]);
                $('#senderEmailAddress').text(emailList[1][i]);
                $('#senderChoose').val('Send email to ' + emailList[0][i]);
            }
            
            $('#senderChoose').click(selectRecipient(emailList[1][3]));
        });
    };

    function selectRecipient(emailAddress) {
        var item = Outlook.Application.ActiveInspector;
        item.recipientField = "tbennet@omahait.com";

    }
})();