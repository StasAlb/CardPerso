using System;
using System.Data;
using System.Diagnostics;
using System.Web.Profile;
using System.Web.Security;
using System.Web.UI;
using System.Web.UI.WebControls;
using System.Data.SqlClient;
using OstCard.Data;

namespace CardPerso.Administration
{
    public partial class ManageUser : Page
    {
        ServiceClass sc = new ServiceClass();
        protected void Page_Load(object sender, EventArgs e)
        {
            ClientScript.RegisterHiddenField("MyHiddenField", "");
            if (!sc.UserAction(User.Identity.Name, Restrictions.UserRolesEdit))
                Response.Redirect(@"~\Account\Restricted.aspx", true);
            if (Page.IsPostBack)
                return;
            lbInform.Text = "";
            lock (Database.lockObjectDB)
            {
                //LoadUsers();
            }
        }
        private void LoadUsers()
        {
            Debug.WriteLine($"{DateTime.Now:HH:mm:ss.fff} Start load user");
            string[] strs = new string[]{"","","","",""};
            if (lbFilter.Text.Trim().Length > 0)
                strs = lbFilter.Text.Trim().Split('=');
            DataSet ds = new DataSet();
            ds.Tables.Add(new DataTable());
            ds.Tables[0].Columns.Add("UserId");
            ds.Tables[0].Columns.Add("UserName");
            ds.Tables[0].Columns.Add("FIO");
            ds.Tables[0].Columns.Add("Position");
            ds.Tables[0].Columns.Add("Branch");
            ds.Tables[0].Columns.Add("Roles");
            var users = Membership.GetAllUsers();
            foreach (MembershipUser mu in users)
                //foreach (MembershipUser mu in Membership.FindUsersByName("%"))
            {
                DataRow dr = ds.Tables[0].NewRow();
                dr["UserName"] = mu.UserName;
                dr["UserId"] = mu.ProviderUserKey;
                ProfileBase pb = ProfileBase.Create(dr["UserName"].ToString());
                UserClass uc = (UserClass)pb.GetPropertyValue("UserData");
                dr["FIO"] = (uc.LastName.Length > 0 && uc.FirstName.Length + uc.SecondName.Length > 0) ? $"{uc.LastName}, {uc.FirstName} {uc.SecondName}"
                    .Trim() : $"{uc.LastName} {uc.FirstName} {uc.SecondName}".Trim();
                dr["Position"] = uc.Position;
                object obj = null;
                Database.ExecuteScalar($"select department from Branchs where id={uc.BranchId}", ref obj, null);
                dr["Branch"] = (obj == DBNull.Value) ? "" : (string)obj;
                DataSet dsr = new DataSet();
                Database.ExecuteQuery($"select RoleId, RoleName from V_UsersRoles where UserId='{dr["UserId"]}'", ref dsr, null);
                string str = "";
                bool fl = false;
                foreach (DataRow dr1 in dsr.Tables[0].Rows)
                {
                    str = String.Format("{0}, {1}", str, dr1["RoleName"].ToString().Trim());
                    if (strs[4].Length > 0 && dr1["RoleId"].ToString() == strs[4])
                        fl = true;
                }
                if (str.StartsWith(", "))
                    str = str.Remove(0, 2);
                dr["Roles"] = str;
                if (strs[0].Length > 0 && mu.UserName.ToLower().IndexOf(strs[0].ToLower(), StringComparison.Ordinal) < 0)
                    continue;
                if (strs[1].Length > 0 && dr["FIO"].ToString().ToLower().IndexOf(strs[1].ToLower(), StringComparison.Ordinal) < 0)
                    continue;
                if (strs[2].Length > 0 && uc.Position.ToLower().IndexOf(strs[2].ToLower(), StringComparison.Ordinal) < 0)
                    continue;
                if (strs[3].Length > 0 && strs[3] != "-1" && Convert.ToInt32(strs[3]) != uc.BranchId)
                    continue;
                if (strs[4].Length > 0 && strs[4] != "-1" && !fl)
                    continue;
                ds.Tables[0].Rows.Add(dr);
            }
            //string res = Database.ExecuteQuery("select id,UserId,UserName from aspnet_Users", ref ds, null);
            //foreach (DataRow dr in ds.Tables[0].Rows)
            //{
            //    ProfileBase pb = ProfileBase.Create(dr["UserName"].ToString());
            //    UserClass uc = (UserClass)pb.GetPropertyValue("UserData");
            //    dr["FIO"] = (uc.LastName.Length > 0 && uc.FirstName.Length+uc.SecondName.Length > 0) ? String.Format("{0}, {1} {2}", uc.LastName, uc.FirstName, uc.SecondName).Trim() : String.Format("{0} {1} {2}", uc.LastName, uc.FirstName, uc.SecondName).Trim();
            //    dr["Position"] = uc.Position;
            //    object obj = null;
            //    Database.ExecuteScalar(String.Format("select department from Branchs where id={0}", uc.BranchId), ref obj, null);
            //    dr["Branch"] = (obj == null) ? "" : (string)obj;
            //    DataSet dsr = new DataSet();
            //    Database.ExecuteQuery(String.Format("select RoleName from V_UsersRoles where id={0}", dr["id"]), ref dsr, null);
            //    string str = "";
            //    foreach (DataRow dr1 in dsr.Tables[0].Rows)
            //        str = String.Format("{0}, {1}", str, dr1["RoleName"].ToString().Trim());
            //    if (str.StartsWith(", "))
            //        str = str.Remove(0, 2);
            //    dr["Roles"] = str;
            //}
            gvUsers.DataSource = ds.Tables[0];
            gvUsers.DataBind();
            lbCount.Text = "Кол-во: " + gvUsers.Rows.Count.ToString();
            SelectRow();
            Debug.WriteLine($"{DateTime.Now:HH:mm:ss.fff} End load user");
        }
        private void SelectRow()
        {
            if (gvUsers.Rows.Count > 0)
            {
                gvUsers.SelectedIndex = 0;
                gvUsers.Rows[gvUsers.SelectedIndex].Focus();
            }
            SetButton();
        }
        private void SelectRow(string UserId)
        {
            for (int i = 0; i < gvUsers.Rows.Count; i++)
                if (gvUsers.DataKeys[i].Values["UserId"].ToString() == UserId)
                {
                    gvUsers.SelectedIndex = i;
                    gvUsers.Rows[i].Focus();
                    break;
                }
            SetButton();
        }
        private void SetButton()
        {
            lock (Database.lockObjectDB)
            {
                if (gvUsers.Rows.Count > 0)
                {
                    bEdit.Visible = true;
                    bDelete.Visible = true;
                    bActivate.Visible = true;
                    bActivate.Enabled = true;
                    bRoles.Visible = true;
                    object obj = null;
                    string str = gvUsers.DataKeys[Convert.ToInt32(gvUsers.SelectedIndex)].Values["UserId"].ToString();
                    Database.ExecuteScalar(String.Format("select count(*) from LogAction where UserId='{0}'", str), ref obj, null);
                    if ((int)obj > 0)
                    {
                        bDelete.ToolTip = "Сделать неактивным";
                        bDelete.Attributes.Add("OnClick", String.Format("return confirm('Сделать неактивным пользователя {0}?');", gvUsers.DataKeys[Convert.ToInt32(gvUsers.SelectedIndex)].Values["UserName"].ToString()));
                    }
                    else
                    {
                        bDelete.ToolTip = "Удалить пользователя";
                        bDelete.Attributes.Add("OnClick", String.Format("return confirm('Удалить пользователя {0}?');", gvUsers.DataKeys[Convert.ToInt32(gvUsers.SelectedIndex)].Values["UserName"].ToString()));
                    }
                    str = gvUsers.DataKeys[Convert.ToInt32(gvUsers.SelectedIndex)].Values["UserName"].ToString();
                    if (!String.IsNullOrEmpty(str) && (Membership.FindUsersByName(str)[str].IsLockedOut || !Membership.FindUsersByName(str)[str].IsApproved))
                        bActivate.Attributes.Add("OnClick", String.Format("return confirm('Активировать пользователя {0}?');", gvUsers.DataKeys[Convert.ToInt32(gvUsers.SelectedIndex)].Values["UserName"].ToString()));
                    else
                        bActivate.Enabled = false;
                    bEdit.Attributes.Add("OnClick", String.Format("return edit_user('id={0}')", gvUsers.DataKeys[Convert.ToInt32(gvUsers.SelectedIndex)].Values["UserId"].ToString()));
                    bRoles.Attributes.Add("OnClick", String.Format("return edit_role('id={0}')", gvUsers.DataKeys[Convert.ToInt32(gvUsers.SelectedIndex)].Values["UserId"].ToString()));
                }
                else
                {
                    bEdit.Visible = false;
                    bRoles.Visible = false;
                    bDelete.Visible = false;
                    bActivate.Visible = false;
                }
            }
        }
        protected void bNew_Click(object sender, ImageClickEventArgs e)
        {
            lock (Database.lockObjectDB)
            {
                //LoadUsers();
                SelectRow(Request.Form["MyHiddenField"]);
            }
        }

        protected void bEdit_Click(object sender, ImageClickEventArgs e)
        {
            lock (Database.lockObjectDB)
            {
                LoadUsers();
                SelectRow(Request.Form["MyHiddenField"]);
            }
        }

        protected void bDelete_Click(object sender, ImageClickEventArgs e)
        {
            lock (Database.lockObjectDB)
            {
                object obj = null;
                Database.ExecuteScalar(String.Format("select count(*) from LogAction where UserId='{0}'", gvUsers.DataKeys[Convert.ToInt32(gvUsers.SelectedIndex)].Values["UserId"].ToString()), ref obj, null);
                if ((int)obj > 0)
                {
                    MembershipUser mu = Membership.GetUser(gvUsers.DataKeys[Convert.ToInt32(gvUsers.SelectedIndex)].Values["UserName"].ToString());
                    mu.IsApproved = false;
                    Membership.UpdateUser(mu);
                }
                else
                {
                    using (SqlCommand comm = new SqlCommand())
                    {
                        comm.Parameters.Add("@uid", SqlDbType.UniqueIdentifier).Value =
                            new Guid(gvUsers.DataKeys[Convert.ToInt32(gvUsers.SelectedIndex)].Values["UserId"].ToString());
                        comm.CommandText =
                            "delete a from AccountablePersonAccount a left join AccountablePerson b on a.id_accountableperson=b.id where b.UserId=(select id from aspnet_Users where UserID=@uid)";
                        SingleQuery.ExecuteNonQuery(comm);
                        comm.CommandText =
                            "delete a from [AccountablePerson_StorageDocs] a left join AccountablePerson b on a.id_person=b.id where b.UserId=((select id from aspnet_Users where UserID=@uid))";
                        SingleQuery.ExecuteNonQuery(comm);
                        comm.CommandText = "delete from AccountablePerson where UserId=(select id from aspnet_Users where UserID=@uid)";
                        SingleQuery.ExecuteNonQuery(comm);
                    }
                    Membership.DeleteUser(gvUsers.DataKeys[Convert.ToInt32(gvUsers.SelectedIndex)].Values["UserName"].ToString());
                }
                LoadUsers();
                SelectRow();
            }
        }

        protected void bExcel_Click(object sender, ImageClickEventArgs e)
        {
        }

        protected void gvUsers_SelectedIndexChanged(object sender, EventArgs e)
        {
            gvUsers.Rows[gvUsers.SelectedIndex].Focus();
            SetButton();
        }

        protected void bActivate_Click(object sender, ImageClickEventArgs e)
        {
            lock (Database.lockObjectDB)
            {
                MembershipUser mu = Membership.GetAllUsers()[gvUsers.DataKeys[Convert.ToInt32(gvUsers.SelectedIndex)].Values["UserName"].ToString()];
                string UserId = gvUsers.DataKeys[Convert.ToInt32(gvUsers.SelectedIndex)].Values["UserId"].ToString();
                bool bl = mu.UnlockUser();
                mu.IsApproved = true;
                string pwd = mu.ResetPassword();
                mu.ChangePassword(pwd, "123");
                Membership.UpdateUser(mu);
                string str = Database.ExecuteNonQuery(String.Format("update aspnet_Users set ActivePassword=0 where UserID='{0}'", UserId), null);
                str = Database.ExecuteNonQuery(String.Format("update aspnet_Membership set IsApproved=1 where UserID='{0}'", UserId), null);
                LoadUsers();
                SelectRow(UserId);
            }
        }

        protected void gvUsers_RowDataBound(object sender, GridViewRowEventArgs e)
        {
            lock (Database.lockObjectDB)
            {
                if (e.Row.RowType == DataControlRowType.DataRow)
                {
                    string str = DataBinder.Eval(e.Row.DataItem, "UserName").ToString();
                    if (!String.IsNullOrEmpty(str) && Membership.FindUsersByName(str).Count > 0)
                        if (Membership.FindUsersByName(str)[str].IsLockedOut || !Membership.FindUsersByName(str)[str].IsApproved)
                            e.Row.Font.Italic = true;
                }
            }
        }
        protected void bRoles_Click(object sender, ImageClickEventArgs e)
        {
            lock (Database.lockObjectDB)
            {
                lbFilter.Text = "";
                LoadUsers();
                SelectRow(Request.Form["MyHiddenField"]);
            }
        }
        protected void bSetFilter_Click(object sender, ImageClickEventArgs e)
        {
            lock (Database.lockObjectDB)
            {
                lbFilter.Text = Request.Form["MyHiddenField"];
                LoadUsers();
            }
        }
        protected void bResetFilter_Click(object sender, ImageClickEventArgs e)
        {
            lock (Database.lockObjectDB)
            {
                lbFilter.Text = "";
                LoadUsers();
            }
        }

        protected void gvUsers_Sorting(object sender, GridViewSortEventArgs e)
        {
            //int i = FuncClass.GetSortIndex(e.SortExpression, gvUsers);

            //if (lbSortIndex.Text != "")
            //{
            //    int ind = Convert.ToInt32(lbSortIndex.Text);
            //    gvUsers.Columns[ind].HeaderText = gvUsers.Columns[ind].HeaderText.Replace("^", "");
            //    gvUsers.Columns[ind].HeaderStyle.BackColor = System.Drawing.Color.FromArgb(0, 102, 153);
            //}

            //lbSortIndex.Text = i.ToString();

            //gvUsers.Columns[i].HeaderStyle.BackColor = System.Drawing.Color.FromArgb(102, 153, 153);

            //if ("order by " + e.SortExpression + " asc" == lbSort.Text)
            //{
            //    lbSort.Text = "order by " + e.SortExpression + " desc";
            //    gvUsers.Columns[i].HeaderText = gvUsers.Columns[i].HeaderText + "^";
            //}
            //else
            //    lbSort.Text = "order by " + e.SortExpression + " asc";
        }
    }
}
