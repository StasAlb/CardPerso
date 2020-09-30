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
using OstCard.Data;

namespace CardPerso.Administration
{
    public partial class FltUsers : System.Web.UI.Page
    {
        protected void Page_Load(object sender, EventArgs e)
        {
            if (IsPostBack)
                return;
            LoadData();
        }
        private void LoadData()
        {
            lock (Database.lockObjectDB)
            {
                DataSet ds = new DataSet();
                ds.Clear();
                string res = Database.ExecuteQuery("select id, department from Branchs order by department", ref ds, null);
                dListBranch.Items.Add(new ListItem("Все", "-1"));
                for (int i = 0; i < ds.Tables[0].Rows.Count; i++)
                    dListBranch.Items.Add(new ListItem(ds.Tables[0].Rows[i]["department"].ToString(), ds.Tables[0].Rows[i]["id"].ToString()));
                ds.Clear();
                res = Database.ExecuteQuery("select RoleId, RoleName from vw_aspnet_Roles where ApplicationName = '" + Membership.ApplicationName + "' order by RoleName", ref ds, null);
                dListRole.Items.Add(new ListItem("Все", "-1"));
                for (int i = 0; i < ds.Tables[0].Rows.Count; i++)
                    dListRole.Items.Add(new ListItem(ds.Tables[0].Rows[i]["RoleName"].ToString(), ds.Tables[0].Rows[i]["RoleId"].ToString()));
            }
        }
        protected void bSave_Click(object sender, ImageClickEventArgs e)
        {
            string[] strs = new string[5];
            strs[0] = tbLogin.Text.Trim();
            strs[1] = tbFio.Text.Trim();
            strs[2] = tbProf.Text.Trim();
            strs[3] = dListBranch.SelectedItem.Value;
            strs[4] = dListRole.SelectedItem.Value;

            Response.Write("<script language=javascript>window.returnValue='" + String.Join("=", strs) + "'; window.close();</script>");
        }
    }
}
