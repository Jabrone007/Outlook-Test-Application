/// <reference path="../App.js" />

(function () {
    "use strict";

    // The Office initialize function must be run each time a new page is loaded
    Office.initialize = function (reason) {
        $(document).ready(function () {
            app.initialize();

            var item = Office.cast.item.toItemRead(Office.context.mailbox.item);

            $('#senderDisplayName').text(item.from.displayName);
            $('#senderEmailAddress').text(item.from.emailAddress);
            $('#recipientDisplayName').text(item.to[0].displayName);
            $('#recipientEmailAddress').text(item.to[0].emailAddress);

            $('#searchLinkedIn').val('search LinkedIn for ' + item.from.displayName);
            $('#searchLinkedIn').click(searchLinkedIn);
        });
    };

    function searchLinkedIn() {
        var from = Office.context.mailbox.item.sender.displayName;
        var split = from.split(' ');
        var firstName = '';
        var lastName = '';

        // Assumes first string in name is first name.
        firstName = split[0];

        if (split.length > 0) {
            lastName = split[split.length - 1];
        }
        var src = "https://www.linkedin.com/pub/dir/?first=" + firstName + "&last=" + lastName + "&search=Search";
        window.location.href(src);

    }
})();