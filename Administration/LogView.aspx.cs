using System;
using System.Collections;
using System.Configuration;
using System.Data;
using System.Data.SqlClient;
using System.Linq;
using System.Web;
using System.Web.Security;
using System.Web.UI;
using System.Web.UI.HtmlControls;
using System.Web.UI.WebControls;
using System.Web.UI.WebControls.WebParts;
using System.Xml.Linq;
using OstCard.WebControls;
using System.Web.Configuration;
using OstCard.Data;

namespace CardPerso.Administration
{
    public partial class LogView : System.Web.UI.Page
    {
        protected void Page_Load(object sender, EventArgs e)
        {
            ServiceClass sc = new ServiceClass();
            if (!sc.UserAction(User.Identity.Name, Restrictions.Events))
                Response.Redirect("~\\Account\\Restricted.aspx", true);
            if (Page.IsPostBack)
                return;
            DatePickerStart.FirstDayOfWeek = DatePicker.Day.Monday;
            DatePickerEnd.FirstDayOfWeek = DatePicker.Day.Monday;
            DatePickerStart.SelectedDate = DateTime.Now.AddDays(-7);
            DatePickerEnd.SelectedDate = DateTime.Now;
            //LoadLog();
        }
        protected void LoadLog()
        {
            lock (Database.lockObjectDB)
            {
                DataSet ds = new DataSet();
                SqlCommand comm = new SqlCommand();
                comm.CommandText = "select UserName, ActionDate, Description from V_Log where ActionDate >= @dateStart and ActionDate <= @dateEnd and UserName like @UserName and Description like @Event order by ActionDate desc";
                comm.Parameters.Add("@dateStart", SqlDbType.DateTime).Value = DatePickerStart.SelectedDate;
                comm.Parameters.Add("@dateEnd", SqlDbType.DateTime).Value = DatePickerEnd.SelectedDate;
                comm.Parameters.Add("@UserName", SqlDbType.VarChar, 30).Value = String.Format("%{0}%", tbLogin.Text.Trim());
                comm.Parameters.Add("@Event", SqlDbType.VarChar, 500).Value = String.Format("%{0}%", tbEvent.Text.Trim());
                Database.ExecuteCommand(comm, ref ds, null);
                gvLog.DataSource = ds.Tables[0];
                gvLog.DataBind();
            }
        }
        protected void bExcel_Click(object sender, ImageClickEventArgs e)
        {
            lock (Database.lockObjectDB)
            {
                System.Globalization.CultureInfo oldCI = System.Threading.Thread.CurrentThread.CurrentCulture;
                System.Threading.Thread.CurrentThread.CurrentCulture = new System.Globalization.CultureInfo("en-US");
                string doc = "";
                ExcelAp ep = new ExcelAp();
                if (ep.RunApp(ConfigurationSettings.AppSettings["DocPath"] + "Empty.xls"))
                {
                    ep.SetWorkSheet(1);
                    ep.ExportGridExcel(gvLog);
                    if (WebConfigurationManager.AppSettings["DocPath"] != null)
                    {
                        doc = String.Format("{0}Temp\\catalog.xls", WebConfigurationManager.AppSettings["DocPath"]);
                        ep.SaveAsDoc(doc, false);
                    }
                }
                ep.Close();
                GC.Collect();
                System.Threading.Thread.CurrentThread.CurrentCulture = oldCI;
                if (doc.Length > 0)
                    ep.ReturnXls(Response, doc);
            }
        }
        protected void bRefresh_Click(object sender, ImageClickEventArgs e)
        {
            LoadLog();
        }
    }
}
