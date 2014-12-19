<%@ Page Language="C#" AutoEventWireup="true" CodeBehind="Home.aspx.cs" Inherits="OutlookTestAppWeb.AppCompose.Home.Home" %>

<!DOCTYPE html>

<html xmlns="http://www.w3.org/1999/xhtml">
<head>
    <meta charset="UTF-8" />
    <meta http-equiv="X-UA-Compatible" content="IE=Edge" />
    <title></title>
    <script src="../../Scripts/jquery-1.9.1.js" type="text/javascript"></script>

    <link href="../../Content/Office.css" rel="stylesheet" type="text/css" />
    <script src="https://appsforoffice.microsoft.com/lib/1.1/hosted/office.js" type="text/javascript"></script>

    <!-- To enable offline debugging using a local reference to Office.js, use:                        -->
    <!--  <script src="../../Scripts/Office/MicrosoftAjax.js" type="text/javascript"></script>  -->
    <!--  <script src="../../Scripts/Office/1.1/office.js" type="text/javascript"></script>  -->

    <link href="../App.css" rel="stylesheet" type="text/css" />
    <script src="../App.js" type="text/javascript"></script>

    <link href="Home.css" rel="stylesheet" type="text/css" />
    <script src="Home.js" type="text/javascript"></script>
</head>
<body style="margin: 4px">
    <form id="form1" runat="server">
        <h3>Address Book</h3>
        <table id="tblAddressBook">
            <tr><td>Display Name</td><td>Email Address</td></tr>
            <tr><td id="senderDisplayName" /><td id="senderEmailAddress" /></tr>
            <tr><td>&nbsp</td><td><asp:Button runat="server" ID="senderChoose" OnClick="selectRecipient"/></td></tr>
            <tr><td>&nbsp</td><td><asp:Button runat="server" ID="addToContacts" OnClick="addToContacts_OnClick"/></td></tr>
        </table>
    </form>
</body>
</html>
