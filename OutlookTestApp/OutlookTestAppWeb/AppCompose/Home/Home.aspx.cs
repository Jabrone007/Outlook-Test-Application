using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Web.UI;
using System.Web.UI.WebControls;
using Outlook = Microsoft.Office.Interop.Outlook;

namespace OutlookTestAppWeb.AppCompose.Home
{
    public partial class Home : System.Web.UI.Page
    {
        protected void Page_Load(object sender, EventArgs e)
        {

        }

        public void selectRecipient(object sender, EventArgs e)
        {
            var oApp = new Outlook.Application();
            var oMsg = (Outlook.MailItem)oApp.ActiveInspector().CurrentItem;
            //var oMsg = (Outlook.MailItem) oApp.ActiveWindow();
            oMsg.Recipients.Add("tbennet@omahait.com");
        }

        
        void Inspectors_NewInspector(Outlook.Inspector Inspector)
        {
            var mailItem = Inspector.CurrentItem as Outlook.MailItem;
            mailItem.To = "tbennet@omahait.com";
        }

        protected void addToContacts_OnClick(object sender, EventArgs e)
        {
            throw new NotImplementedException();
        }
    }
}