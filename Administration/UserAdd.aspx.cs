using System;
using System.Collections;
using System.Configuration;
using System.Data;
using System.Linq;
using System.Web;
using System.Web.Security;
using System.Web.Profile;
using System.Web.UI;
using System.Web.UI.HtmlControls;
using System.Web.UI.WebControls;
using System.Web.UI.WebControls.WebParts;
using System.Xml.Linq;
using OstCard.Data;

namespace CardPerso.Administration
{
    public partial class UserEdit : System.Web.UI.Page
    {
        protected void Page_Load(object sender, EventArgs e)
        {
            lock (Database.lockObjectDB)
            {
                ServiceClass sc = new ServiceClass();
                if (!sc.UserAction(User.Identity.Name, Restrictions.UserRolesEdit))
                    Response.Redirect("~\\Account\\Restricted.aspx", true);
                if (IsPostBack)
                    return;
                NewUserWizard.RequireEmail = false;
                DataSet ds = new DataSet();
                Database.ExecuteQuery("select id, department from Branchs order by department", ref ds, null);
                DropDownList ddl = (DropDownList)NewUserWizard.CreateUserStep.ContentTemplateContainer.FindControl("BranchDDL");
                ddl.DataSource = ds.Tables[0];
                ddl.DataTextField = "department";
                ddl.DataValueField = "id";
                ddl.DataBind();
            }
        }

        protected void NewUserWizard_CreatedUser(object sender, EventArgs e)
        {
            lock (Database.lockObjectDB)
            {
                UserClass uc = new UserClass();
                uc.FirstName = ((TextBox)NewUserWizard.CreateUserStep.ContentTemplateContainer.FindControl("UserFirstName")).Text;
                uc.SecondName = ((TextBox)NewUserWizard.CreateUserStep.ContentTemplateContainer.FindControl("UserSecondName")).Text;
                uc.LastName = ((TextBox)NewUserWizard.CreateUserStep.ContentTemplateContainer.FindControl("UserLastName")).Text;
                uc.Position = ((TextBox)NewUserWizard.CreateUserStep.ContentTemplateContainer.FindControl("UserPosition")).Text;
                uc.BranchId = Convert.ToInt32(((DropDownList)NewUserWizard.CreateUserStep.ContentTemplateContainer.FindControl("BranchDDL")).SelectedItem.Value);
                uc.SetPassport(((TextBox)NewUserWizard.CreateUserStep.ContentTemplateContainer.FindControl("tbPassportSeries")).Text.Trim(),
                    ((TextBox)NewUserWizard.CreateUserStep.ContentTemplateContainer.FindControl("tbPassportNumber")).Text.Trim());
                ProfileBase pb = ProfileBase.Create(NewUserWizard.UserName, true);
                pb.SetPropertyValue("UserData", uc);
                pb.Save();
                CheckBoxList cbl = (CheckBoxList)NewUserWizard.CompleteStep.ContentTemplateContainer.FindControl("cblRoles");
                cbl.Items.Clear();
                DataTable dt = new DataTable();
                dt.Columns.Add("RoleId");
                dt.Columns.Add("RoleName");
                foreach (string str in Roles.GetAllRoles())
                {
                    DataRow dr = dt.NewRow();
                    dr["RoleName"] = str;
                    object obj = null;
                    Database.ExecuteScalar("select RoleId from  vw_aspnet_Roles where ApplicationName='" + Membership.ApplicationName + "' and RoleName='" + str + "'", ref obj, null);
                    dr["RoleId"] = obj.ToString();
                    dt.Rows.Add(dr);
                }
                /*            DataSet ds = new DataSet();
                            Database.ExecuteQuery("select RoleId, RoleName from aspnet_Roles order by RoleName", ref ds, null);
                            cbl.DataSource = ds.Tables[0];*/
                cbl.DataTextField = "RoleName";
                cbl.DataValueField = "RoleId";
                cbl.DataSource = dt;
                cbl.DataBind();
            }
        }

        protected void NewUserWizard_CreatingUser(object sender, LoginCancelEventArgs e)
        {
            
        }

        protected void bRoles_Click(object sender, EventArgs e)
        {
            lock (Database.lockObjectDB)
            {
                CheckBoxList cbl = (CheckBoxList)NewUserWizard.CompleteStep.ContentTemplateContainer.FindControl("cblRoles");
                foreach (ListItem li in cbl.Items)
                {
                    if (li.Selected)
                        Roles.AddUserToRole(NewUserWizard.UserName, li.Text);
                }
                object obj = null;
                Database.ExecuteScalar(String.Format("select UserId from aspnet_Users where UserName='{0}'", NewUserWizard.UserName), ref obj, null);
                Response.Write(String.Format("<script language=javascript>window.returnValue='{0}'; window.close();</script>", Convert.ToString(obj)));
            }
        }
    }
}
