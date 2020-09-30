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
using OstCard.Data;

namespace CardPerso.Administration
{
    public partial class ManageRoles : System.Web.UI.Page
    {
        ServiceClass sc = new ServiceClass();
        protected void Page_Load(object sender, EventArgs e)
        {
            ClientScript.RegisterHiddenField("MyHiddenField", "");
            if (!sc.UserAction(User.Identity.Name, Restrictions.UserRolesEdit))
                Response.Redirect("~\\Account\\Restricted.aspx", true);
            if (Page.IsPostBack)
                return;
            lock (Database.lockObjectDB)
            {
                LoadActions();
                LoadRoles();
                SelectRow();
            }
        }
        private void LoadRoles()
        {
            DataTable dt = new DataTable();
            dt.Columns.Add("RoleId");
            dt.Columns.Add("RoleName");
            dt.Columns.Add("UserCnt");
            foreach (string str in Roles.GetAllRoles())
            {
                DataRow dr = dt.NewRow();
                dr["RoleName"] = str;
                dr["UserCnt"] = Roles.GetUsersInRole(str).Length;
                object obj = null;
                Database.ExecuteScalar("select RoleId from vw_aspnet_Roles where ApplicationName = '" + Membership.ApplicationName + "' and RoleName='" + str + "'", ref obj, null);
                dr["RoleId"] = obj.ToString();
                dt.Rows.Add(dr);
            }
            gvRoles.DataSource = dt;
//            DataSet ds = new DataSet();
//            Database.ExecuteQuery("select * from V_Roles order by RoleName", ref ds, null);
//            gvRoles.DataSource = ds.Tables[0];
            gvRoles.DataBind();
        }
        private void SelectRow()
        {
            if (gvRoles.Rows.Count > 0)
            {
                gvRoles.SelectedIndex = 0;
                gvRoles.Rows[gvRoles.SelectedIndex].Focus();
                lAction.Text = String.Format("Перечень действий для роли {0}", gvRoles.DataKeys[Convert.ToInt32(gvRoles.SelectedIndex)].Values["RoleName"].ToString());
                LoadActions(gvRoles.DataKeys[Convert.ToInt32(gvRoles.SelectedIndex)].Values["RoleId"].ToString());
            }
            SetButton();
        }
        private void SelectRow(string RoleId)
        {
            for (int i = 0; i < gvRoles.Rows.Count; i++)
                if (gvRoles.DataKeys[i].Values["RoleId"].ToString() == RoleId)
                {
                    gvRoles.SelectedIndex = i;
                    gvRoles.Rows[i].Focus();
                    lAction.Text = String.Format("Перечень действий для роли {0}", gvRoles.DataKeys[i].Values["RoleName"].ToString());
                    LoadActions(gvRoles.DataKeys[i].Values["RoleId"].ToString());
                    break;
                }
            SetButton();
        }
        private void LoadActions()
        {
            DataSet ds = new DataSet();
            Database.ExecuteQuery("select * from Actions order by priority", ref ds, null);
            clbAction.DataSource = ds.Tables[0];
            clbAction.DataBind();
        }
        private void SetButton()
        {
            if (gvRoles.Rows.Count > 0)
            {
                bEdit.Visible = true;
                bDelete.Visible = true;
                int userCnt = Convert.ToInt32(gvRoles.DataKeys[Convert.ToInt32(gvRoles.SelectedIndex)].Values["UserCnt"]);
                if (userCnt == 0)
                {
                    bDelete.Enabled = true;
                    bDelete.Attributes.Add("OnClick", String.Format("return confirm('Удалить роль {0}?');", gvRoles.DataKeys[Convert.ToInt32(gvRoles.SelectedIndex)].Values["RoleName"].ToString()));
                    bDelete.ToolTip = "Удалить";
                }
                else
                {
                    bDelete.Enabled = false;
                    bDelete.ToolTip = "Удаление невозможно";
                }
                bEdit.Attributes.Add("OnClick", String.Format("return show_role('mode=2&id={0}')", gvRoles.DataKeys[Convert.ToInt32(gvRoles.SelectedIndex)].Values["RoleId"].ToString()));
            }
            else
            {
                bEdit.Visible = false;
                bDelete.Visible = false;
            }
        }

        protected void bNew_Click(object sender, ImageClickEventArgs e)
        {
            lock (Database.lockObjectDB)
            {
                LoadRoles();
                SelectRow(Request.Form["MyHiddenField"]);
            }
        }

        protected void bEdit_Click(object sender, ImageClickEventArgs e)
        {
            lock (Database.lockObjectDB)
            {
                LoadRoles();
                string str = Request.Form["MyHiddenField"];
                SelectRow(Request.Form["MyHiddenField"]);
            }
        }

        protected void bDelete_Click(object sender, ImageClickEventArgs e)
        {
            lock (Database.lockObjectDB)
            {
                string RoleId = gvRoles.DataKeys[Convert.ToInt32(gvRoles.SelectedIndex)].Values["RoleId"].ToString();
                Database.ExecuteNonQuery(String.Format("delete from RoleAction where RoleId={0}", RoleId), null);
                Roles.DeleteRole(gvRoles.DataKeys[Convert.ToInt32(gvRoles.SelectedIndex)].Values["RoleName"].ToString());
                LoadRoles();
                SelectRow();
            }
        }

        protected void bExcel_Click(object sender, ImageClickEventArgs e)
        {

        }

        protected void gvRoles_SelectedIndexChanged(object sender, EventArgs e)
        {
            lock (Database.lockObjectDB)
            {
                gvRoles.Rows[gvRoles.SelectedIndex].Focus();
                lAction.Text = String.Format("Перечень действий для роли {0}", gvRoles.DataKeys[Convert.ToInt32(gvRoles.SelectedIndex)].Values["RoleName"].ToString());
                LoadActions(gvRoles.DataKeys[Convert.ToInt32(gvRoles.SelectedIndex)].Values["RoleId"].ToString());
                SetButton();
            }
        }
        private void LoadActions(string roleid)
        {
            foreach (ListItem li in clbAction.Items)
                li.Selected = false;
            DataSet ds = new DataSet();
            Database.ExecuteQuery("select ActionId from RoleAction where RoleId='"+roleid.ToString()+"'", ref ds, null);
            foreach (DataRow dr in ds.Tables[0].Rows)
            {
                foreach(ListItem li in clbAction.Items)
                    if (li.Value == dr["ActionId"].ToString())
                        li.Selected = true;
            }
        }

        protected void clbAction_SelectedIndexChanged(object sender, EventArgs e)
        {
            lock (Database.lockObjectDB)
            {
                SqlCommand comm = new SqlCommand();
                comm.CommandText = "delete from RoleAction where RoleId=@RoleId";
                comm.Parameters.Add("@RoleId", SqlDbType.UniqueIdentifier).Value = new Guid(gvRoles.DataKeys[Convert.ToInt32(gvRoles.SelectedIndex)].Values["RoleId"].ToString());
                Database.ExecuteNonQuery(comm, null);
                comm.CommandText = "insert into RoleAction (ActionId, RoleId) values (@ActionId, @RoleId)";
                comm.Parameters.Add("@ActionId", SqlDbType.Int);
                foreach (ListItem li in clbAction.Items)
                {
                    if (li.Selected)
                    {
                        comm.Parameters["@ActionId"].Value = Convert.ToInt32(li.Value);
                        Database.ExecuteNonQuery(comm, null);
                    }
                }
            }
        }
    }
}
