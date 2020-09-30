using System;
using System.Collections;
using System.Configuration;
using System.Data;
using System.Linq;
using System.Web;
using System.Web.Security;
using System.Web.UI;
using System.Web.UI.HtmlControls;
using System.Web.UI.WebControls;
using System.Web.UI.WebControls.WebParts;
using System.Xml.Linq;
using System.Net;
using System.Net.Mail;

namespace CardPerso.Administration
{
    public partial class EMail : System.Web.UI.Page
    {
        protected void Page_Load(object sender, EventArgs e)
        {
            if (Page.IsPostBack)
                return;

        }

        protected void bSend_Click(object sender, EventArgs e)
        {
            SmtpClient sc = new SmtpClient(ConfigurationSettings.AppSettings["SmtpServer"]);
            sc.Credentials = new NetworkCredential(ConfigurationSettings.AppSettings["EMailFrom"], ConfigurationSettings.AppSettings["Pwd"]);
            MailAddress mailFrom = new MailAddress(ConfigurationSettings.AppSettings["EMailFrom"],"CardPerso");
            MailAddress mailTo = new MailAddress(tbTo.Text);
            MailMessage mm = new MailMessage(mailFrom, mailTo);
            mm.Subject = "CardPerso TestMessage";
            mm.Body = "This is CardPerso test message. If you get it, all is fine";
            try
            {
                sc.Send(mm);
            }
            catch (Exception ex)
            {
                string str = ex.Message;
            }
        }
    }
}
