using System;
using System.Collections.Generic;
using System.Drawing;
using System.Linq;
using System.Web;
using System.Web.UI;
using System.Web.UI.WebControls;
using System.Windows.Forms;
using Microsoft.Office.Interop.Outlook;
using Exception = System.Exception;

namespace OutlookTestAppWeb.AppRead.Home
{
    public partial class Home : System.Web.UI.Page
    {
        protected void Page_Load(object sender, EventArgs e)
        {

        }

        protected void addToContacts_OnClick(object sender, EventArgs e)
        {
            var oApp = new Microsoft.Office.Interop.Outlook.Application();

            MAPIFolder contactsFolder = oApp.ActiveExplorer().Session.GetDefaultFolder(OlDefaultFolders.olFolderContacts);

            Items contactItems = contactsFolder.Items;

            string sFirstName = "Test";
            string sLastName = "McTesterson";

            string sSearch = String.Format("[FirstName]='{0}' and " + "[LastName]='{1}'", sFirstName, sLastName);

            ContactItem contact = (ContactItem) contactItems.Find(sSearch);

            if (contact != null)
            {
                DialogResult updateContact = MessageBox.Show("Would you like to update?", "This contact already exists!",
                    MessageBoxButtons.YesNoCancel);
                if (updateContact == DialogResult.Yes)
                {
                    contact.MailingAddressStreet = "1234 Test St.";
                    contact.Display();
                }
                else if (updateContact == DialogResult.No)
                    contact.Display();
            }

            else
            {
                var newContact = (ContactItem) oApp.CreateItem(OlItemType.olContactItem);
                try
                {
                    newContact.FirstName = "Test";
                    newContact.LastName = "Testerson";
                    newContact.Email1Address = "test@test.com";
                    newContact.CustomerID = "123456";
                    newContact.PrimaryTelephoneNumber = "(425)555-0111";
                    newContact.MailingAddressStreet = "123 Test St.";
                    newContact.MailingAddressCity = "Redmond";
                    newContact.MailingAddressState = "WA";
                    newContact.Save();
                    newContact.Display(true);
                }
                catch (Exception)
                {
                    MessageBox.Show("The new contact was not saved.");
                }
            }
        }
    }
}