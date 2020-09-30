using System;
using System.Collections;
using System.Configuration;
using System.Data;
using System.Data.SqlClient;
using System.Linq;
using System.Web;
using System.Web.Profile;
using System.Web.Security;
using System.Web.UI;
using System.Web.UI.HtmlControls;
using System.Web.UI.WebControls;
using System.Web.UI.WebControls.WebParts;
using System.Xml.Linq;
using OstCard.Data;
namespace CardPerso.Administration
{
    public partial class RoleEdit : System.Web.UI.Page
    {
        protected void Page_Load(object sender, EventArgs e)
        {
            if (Page.IsPostBack)
                return;
            lock (Database.lockObjectDB)
            {
                ServiceClass sc = new ServiceClass();
                if (!sc.UserAction(User.Identity.Name, Restrictions.UserRolesEdit))
                    Response.Redirect("~\\Account\\Restricted.aspx", true);
                if (Request.QueryString["mode"] == "2")
                    LoadData();
                tbRoleName.Focus();
            }
        }
        private void LoadData()
        {
            string str = Convert.ToString(Request.QueryString["id"]);
            object res = null;
            Database.ExecuteScalar("select RoleName from aspnet_Roles where RoleId='" + str + "'", ref res, null);
            tbRoleName.Text = (string)res;
        }

        protected void bSave_Click(object sender, ImageClickEventArgs e)
        {
            lock (Database.lockObjectDB)
            {
                if (tbRoleName.Text.Trim().Length == 0)
                {
                    lbInform.Text = "Введите наименование роли";
                    tbRoleName.Focus();
                    return;
                }
                string ret = "";
                SqlCommand comm = new SqlCommand();
                if (Request.QueryString["mode"] == "1")
                {
                    Roles.CreateRole(tbRoleName.Text.Trim());
                    object obj = null;
                    Database.ExecuteScalar(String.Format("select RoleId from aspnet_Roles where RoleName='{0}'", tbRoleName.Text.Trim()), ref obj, null);
                    ret = Convert.ToString(obj);
                }
                if (Request.QueryString["mode"] == "2")
                {
                    ret = Request.QueryString["id"];
                    comm.CommandText = "update aspnet_Roles set RoleName=@RoleName, LoweredRoleName=@LowName where RoleId=@id";
                    comm.Parameters.Add("@id", SqlDbType.UniqueIdentifier).Value = new Guid(ret);
                    comm.Parameters.Add("@RoleName", SqlDbType.NVarChar, 255).Value = tbRoleName.Text.Trim();
                    comm.Parameters.Add("@LowName", SqlDbType.NVarChar, 255).Value = tbRoleName.Text.Trim().ToLower();

                    Database.ExecuteNonQuery(comm, null);
                }
                Response.Write(String.Format("<script language=javascript>window.returnValue='{0}'; window.close();</script>", ret));
            }
        }
    }
}
