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
    public partial class UserEdit1 : System.Web.UI.Page
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
                Database.ExecuteQuery("select id, department from Branchs order by department", ref ds, null);
                BranchDDL.DataSource = ds.Tables[0];
                BranchDDL.DataTextField = "department";
                BranchDDL.DataValueField = "id";
                BranchDDL.DataBind();
                UserData();
            }
        }
        private void UserData()
        {
            DataSet ds = new DataSet();
            string str = Convert.ToString(Request.QueryString["id"]);
            Database.ExecuteQuery(String.Format("select UserName, EMail from vw_aspnet_MembershipUsers where UserId='{0}'", str), ref ds, null);
            if (ds != null && ds.Tables.Count > 0 && ds.Tables[0].Rows.Count > 0)
            {
                UserName.Text = ds.Tables[0].Rows[0]["UserName"].ToString(); ;
                Email.Text = ds.Tables[0].Rows[0]["EMail"].ToString();
                ProfileBase pb = ProfileBase.Create(UserName.Text, false);
                UserClass uc = (UserClass)pb.GetPropertyValue("UserData");
                UserLastName.Text = uc.LastName;
                UserFirstName.Text = uc.FirstName;
                UserSecondName.Text = uc.SecondName;
                UserPosition.Text = uc.Position;
                tbPassportSeries.Text = uc.PassportSeries;
                tbPassportNumber.Text = uc.PassportNumber;
                foreach (ListItem li in BranchDDL.Items)
                {
                    if (li.Value == uc.BranchId.ToString())
                        li.Selected = true;
                }
            }
        }
        protected void bSave_Click(object sender, ImageClickEventArgs e)
        {
            lock (Database.lockObjectDB)
            {
                string str = Convert.ToString(Request.QueryString["id"]);
                Database.ExecuteNonQuery(String.Format("update aspnet_membership set email='{0}' where UserId='{1}'", Email.Text, str), null);
                ProfileBase pb = ProfileBase.Create(UserName.Text, true);
                UserClass uc = new UserClass();
                uc.FirstName = UserFirstName.Text;
                uc.SecondName = UserSecondName.Text;
                uc.LastName = UserLastName.Text;
                uc.Position = UserPosition.Text;
                uc.BranchId = Convert.ToInt32(BranchDDL.SelectedItem.Value);
                uc.SetPassport(tbPassportSeries.Text.Trim(), tbPassportNumber.Text.Trim());
                pb.SetPropertyValue("UserData", uc);
                pb.Save();
                Response.Write(String.Format("<script language=javascript>window.returnValue='{0}'; window.close();</script>", str));
            }
        }
    }
}
