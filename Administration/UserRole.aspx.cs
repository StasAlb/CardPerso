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
    public partial class UserRole : System.Web.UI.Page
    {
        protected void Page_Load(object sender, EventArgs e)
        {
            ServiceClass sc = new ServiceClass();
            if (!sc.UserAction(User.Identity.Name, Restrictions.UserRolesEdit))
                Response.Redirect("~\\Account\\Restricted.aspx", true);
            if (Page.IsPostBack)
                return;
            lock (Database.lockObjectDB)
            {
                DataSet ds = new DataSet();
                Database.ExecuteQuery("select RoleId, RoleName from vw_aspnet_Roles where ApplicationName = '" + Membership.ApplicationName + "' order by RoleName", ref ds, null);
                cblRoles.DataSource = ds.Tables[0];
                cblRoles.DataTextField = "RoleName";
                cblRoles.DataValueField = "RoleId";
                cblRoles.DataBind();
                ds.Clear();
                string UserId = Request.QueryString["id"].ToString();
                Database.ExecuteQuery(String.Format("select RoleId from V_UsersRoles where UserId='{0}'", UserId), ref ds, null);
                foreach (DataRow dr in ds.Tables[0].Rows)
                {
                    foreach (ListItem li in cblRoles.Items)
                        if (li.Value == dr["RoleId"].ToString())
                            li.Selected = true;
                }
            }
        }
        protected void bSave_Click(object sender, EventArgs e)
        {
            lock (Database.lockObjectDB)
            {
                object obj = null;
                string UserId = Request.QueryString["id"].ToString();
                Database.ExecuteScalar(String.Format("select UserName from aspnet_Users where UserId='{0}'", UserId), ref obj, null);
                foreach (ListItem li in cblRoles.Items)
                {
                    if (li.Selected)
                    {
                        if (!Roles.IsUserInRole((string)obj, li.Text))
                            Roles.AddUserToRole((string)obj, li.Text);
                    }
                    else
                    {
                        if (Roles.IsUserInRole((string)obj, li.Text))
                            Roles.RemoveUserFromRole((string)obj, li.Text);
                    }
                }
                Response.Write(String.Format("<script language=javascript>window.returnValue='{0}'; window.close();</script>", UserId));
            }
        }
    }
}
