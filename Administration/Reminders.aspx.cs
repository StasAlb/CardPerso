using System;
using System.Collections;
using System.Configuration;
using System.Data;
using System.Data.SqlClient;
using System.Linq;
using System.Web;
using System.Web.Security;
using System.Web.Profile;
using System.Web.UI;
using System.Web.UI.WebControls;
using System.Web.UI.WebControls.WebParts;
using System.Web.UI.HtmlControls;
using System.Xml.Linq;
using System.Net;
using System.Net.Mail;
using OstCard.Data;

namespace CardPerso.Administration
{
    public partial class Reminders : System.Web.UI.Page
    {
        protected void Page_Load(object sender, EventArgs e)
        {
            lock (Database.lockObjectDB)
            {
                ServiceClass sc = new ServiceClass();
                if (!sc.UserAction(User.Identity.Name, Restrictions.Events))
                    Response.Redirect("~\\Account\\Restricted.aspx", true);
                if (Page.IsPostBack)
                    return;
                LoadRepeater();
            }
        }
        private void LoadRepeater()
        {
            lock (Database.lockObjectDB)
            {
                SqlCommand comm = new SqlCommand();
                comm.CommandText = "select id, date_doc, branch, id_branch, priz_gen from V_StorageDocs where (type=@type) and (id not in (select id_act from StorageDocs where priz_gen=1))";
                comm.Parameters.Add("@type", SqlDbType.Int).Value = (int)TypeDoc.SendToFilial;
                DataSet ds = new DataSet();
                object obj = Database.ExecuteCommand(comm, ref ds, null);
                if (ds.Tables.Count == 0)
                    return;
                DataTable dt = new DataTable();
                dt.Columns.Add("id");
                dt.Columns.Add("Message");
                dt.Columns.Add("FIO");
                dt.Columns.Add("branchID");
                dt.Columns.Add("ToButton");
                dt.Columns.Add("Enbl", Type.GetType("System.Boolean"));
                for (int i = 0; i < ds.Tables[0].Rows.Count; i++)
                {
                    DateTime dateT = Convert.ToDateTime(ds.Tables[0].Rows[i]["date_doc"]).Date;
                    if (Convert.ToInt32((DateTime.Now.Date - dateT).TotalDays) < Convert.ToInt32(ConfigurationSettings.AppSettings["ReminderFilialDay"]))
                        continue;
                    string mails = "";
                    int branchID = Convert.ToInt32(ds.Tables[0].Rows[i]["id_branch"]);
                    string res = Database.ExecuteScalar(String.Format("select id from V_StorageDocs where id_act={0}", ds.Tables[0].Rows[i]["id"].ToString()), ref obj, null);
                    DataRow dr = dt.NewRow();
                    dr["FIO"] = (string)"";
                    dr["id"] = ds.Tables[0].Rows[i]["id"].ToString();
                    if (obj == null)
                        dr["Message"] = String.Format("{0} - Прием по акту от {1} не сформирован", ds.Tables[0].Rows[i]["branch"].ToString(), Convert.ToDateTime(ds.Tables[0].Rows[i]["date_doc"]).ToShortDateString());
                    else
                        dr["Message"] = String.Format("{0} - Прием по акту от {1} не подтвержден", ds.Tables[0].Rows[i]["branch"].ToString(), Convert.ToDateTime(ds.Tables[0].Rows[i]["date_doc"]).ToShortDateString());
                    DataSet ds1 = new DataSet();
                    res = Database.ExecuteQuery(String.Format("select UserName from V_UserAction where ActionID={0}", (int)Restrictions.ReminderReceiveFilial), ref ds1, null);
                    if (ds1.Tables.Count == 0)
                        continue;
                    for (int t = 0; t < ds1.Tables[0].Rows.Count; t++)
                    {
                        ProfileBase pb = ProfileBase.Create(ds1.Tables[0].Rows[t]["UserName"].ToString());
                        UserClass uc = (UserClass)pb.GetPropertyValue("UserData");
                        if (uc.BranchId == branchID)
                        {
                            obj = null;
                            Database.ExecuteScalar(String.Format("select EMail from vw_aspnet_MembershipUsers where UserName='{0}'", ds1.Tables[0].Rows[t]["UserName"].ToString()), ref obj, null);
                            dr["FIO"] = String.Format("{0}, {1}", (string)dr["FIO"], uc.Name1);
                            mails += ";" + (string)obj;
                        }
                    }
                    if (((string)dr["FIO"]).Length != 0)
                    {
                        dr["FIO"] = ((string)dr["FIO"]).Substring(2);
                        dr["Enbl"] = true;
                    }
                    else
                    {
                        dr["FIO"] = "не определено";
                        dr["Enbl"] = false;
                    }
                    if (mails.StartsWith(";"))
                        mails = mails.Substring(1);
                    dr["ToButton"] = String.Format("{0}\t{1}\t{2}", mails, "CardPerso - напоминание", dr["Message"].ToString());
                    dr["branchID"] = branchID;
                    dt.Rows.Add(dr);
                }
                rReceiveInFilial.DataSource = dt;
                rReceiveInFilial.DataBind();
            }
        }
        protected void bSendMessage_Click(object sender, EventArgs e)
        {
            SmtpClient sc = new SmtpClient(ConfigurationSettings.AppSettings["SmtpServer"]);
            sc.Credentials = new NetworkCredential(ConfigurationSettings.AppSettings["EMailFrom"], ConfigurationSettings.AppSettings["Pwd"]);
            MailAddress mailFrom = new MailAddress(ConfigurationSettings.AppSettings["EMailFrom"], "CardPerso");
            string mails = ((Button)sender).CommandArgument.Split('\t')[0];
            MailMessage mm = new MailMessage();
            foreach(string mail in mails.Split(';'))
            {
                if (mail.Length > 0)
                {
                    MailAddress mailTo = new MailAddress(mail);
                    mm.Bcc.Add(mailTo);
                }
            }
            mm.From = mailFrom;
            mm.Subject = ((Button)sender).CommandArgument.Split('\t')[1];
            mm.Body = ((Button)sender).CommandArgument.Split('\t')[2];
            sc.Send(mm);
        }
    }
}
