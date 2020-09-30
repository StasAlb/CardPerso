using System;
using System.Collections.Generic;
using System.Linq;
using System.Data;
using System.Web;
using System.Web.UI;
using System.Configuration;
using System.Web.UI.WebControls;
using System.Collections;
using CardPerso;

namespace CardPerso.Account
{
    public partial class Login : System.Web.UI.Page
    {
        CardPerso.Administration.ServiceClass sc = new CardPerso.Administration.ServiceClass();
        string superusername = "";
        protected void Page_Load(object sender, EventArgs e)
        {
            superusername = "AkBarsAdmin";
            try
            {
                string str = ConfigurationManager.AppSettings["ClientType"];
                if (str.ToLower() == "uzcard")
                    superusername = "UzcardAdmin";
            }
            catch
            {
            }

            LoginUser.Focus();
            if (LoginUser.FailureText == "Данный логин/пароль не найден")
            {
                getBranchLabel().Visible = false;
                getBranch().Visible = false;
            }
        }
        protected Label getBranchLabel() 
        {
            return (System.Web.UI.WebControls.Label)LoginUser.FindControl("BranchLabel");
        }
        protected DropDownList getBranch()
        {
            return (System.Web.UI.WebControls.DropDownList)LoginUser.FindControl("Branch");
        }
        protected void LoginUser_LoggingIn(object sender, LoginCancelEventArgs e)
        {

        }
        protected void LoginUser_Authenticate(object sender, AuthenticateEventArgs e)
        {
            LoginUser.FailureText = "Данный логин/пароль не найден";
            HttpContext.Current.Session["BranchId"] = null;
            HttpContext.Current.Session["UserActions"] = null;
            if (LoginUser.UserName == superusername && LoginUser.Password == superusername.ToLower())
            {
                e.Authenticated = true;
                OstCard.Data.Database.Log(sc.UserGuid(superusername), "Вход в систему", null);
                return;
            }
            else
            {
                bool auth = System.Web.Security.Membership.ValidateUser(LoginUser.UserName, LoginUser.Password);
                if (auth == true)
                {
                    // Проверяем на Казанский филиал и может ли он логинится в любой филиал
                    int branchId = sc.BranchId(LoginUser.UserName);
                    int branchIdMain = BranchStore.getBranchMainFilial(branchId, false);

                    string abu = System.Configuration.ConfigurationManager.AppSettings["AllBranchsUser"];
                    string[] allbranchsusers = abu?.Split(',');

                    ArrayList al = new ArrayList();
                    DataSet ds = new DataSet();

                    if (allbranchsusers.Contains(LoginUser.UserName))
                    {
                        lock (OstCard.Data.Database.lockObjectDB)
                        {
                            OstCard.Data.Database.ExecuteQuery("select id, department from branchs order by department", ref ds, null);
                            if (ds.Tables.Count > 0 && ds.Tables[0].Rows.Count > 0)
                            {
                                for (int i = 0; i < ds.Tables[0].Rows.Count; i++)
                                {
                                    BranchStore branchItem = new BranchStore(Convert.ToInt32(ds.Tables[0].Rows[i]["id"]), "", ds.Tables[0].Rows[i]["department"].ToString());
                                    if (branchItem.id == branchId) branchItem.ident_dep = "select";
                                    al.Add(branchItem);
                                }
                            }
                        }
                    }
                    else
                    {
                        //if (branchId == 106) // Казанский филиал - можно выбрать все подчиненные
                        if (branchIdMain > 0)
                        {

                            lock (OstCard.Data.Database.lockObjectDB)
                            {
                                //OstCard.Data.Database.ExecuteQuery("select id, department from branchs where id=106 or id_parent=106 order by id", ref ds, null);
                                OstCard.Data.Database.ExecuteQuery(string.Format("select id, department from branchs where id={0} or id_parent={0} order by id", branchIdMain), ref ds, null);
                                if (ds.Tables.Count > 0 && ds.Tables[0].Rows.Count > 0)
                                {
                                    //if (branchId == 106)
                                    if (branchId == branchIdMain)
                                    {
                                        BranchStore branchItem = new BranchStore(-1, "", "Выберите филиал");
                                        branchItem.ident_dep = "select";
                                        al.Add(branchItem);
                                    }
                                    for (int i = 0; i < ds.Tables[0].Rows.Count; i++)
                                    {
                                        BranchStore branchItem = new BranchStore(Convert.ToInt32(ds.Tables[0].Rows[i]["id"]), "", ds.Tables[0].Rows[i]["department"].ToString());
                                        //if (branchItem.id == branchId && branchId!=106) branchItem.ident_dep = "select";
                                        if (branchItem.id == branchId && branchId != branchIdMain) branchItem.ident_dep = "select";
                                        al.Add(branchItem);
                                    }
                                }
                            }
                        }
                        else // Ищем подчиненные Казанского филиала
                        {
                            lock (OstCard.Data.Database.lockObjectDB)
                            {
                                //OstCard.Data.Database.ExecuteQuery("select id, department from branchs where id_parent=106 and id=" + branchId, ref ds, null);
                                OstCard.Data.Database.ExecuteQuery("select id, department from branchs where id_parent=" + branchIdMain + " and id=" + branchId, ref ds, null);
                                if (ds.Tables.Count > 0 && ds.Tables[0].Rows.Count > 0)
                                {
                                    //OstCard.Data.Database.ExecuteQuery("select id, department from branchs where id_parent=106 order by id", ref ds, null);
                                    OstCard.Data.Database.ExecuteQuery("select id, department from branchs where id_parent=" + branchIdMain + " order by id", ref ds, null);
                                    if (ds.Tables.Count > 0 && ds.Tables[0].Rows.Count > 0)
                                    {
                                        for (int i = 0; i < ds.Tables[0].Rows.Count; i++)
                                        {
                                            BranchStore branchItem = new BranchStore(Convert.ToInt32(ds.Tables[0].Rows[i]["id"]), "", ds.Tables[0].Rows[i]["department"].ToString());
                                            if (branchItem.id == branchId) branchItem.ident_dep = "select";
                                            al.Add(branchItem);
                                        }
                                    }
                                }
                            }
                        }
                    }

                    DropDownList dList = getBranch();
                    if (dList.Items.Count > 1 && dList.SelectedIndex >= 0)
                    {
                        if (Convert.ToInt32(dList.Items[dList.SelectedIndex].Value) > 0)
                        {

                            HttpContext.Current.Session["BranchId"] = dList.Items[dList.SelectedIndex].Value;
                            HttpContext.Current.Session["CurrentUserName"] = LoginUser.UserName.ToLower();
                            HttpContext.Current.Session["CurrentUserId"] = sc.UserId(LoginUser.UserName);
                        }
                        else
                        {
                            dList.Items.Clear();
                        }
                    }
                    
                    if (dList.Items.Count < 1 && al.Count > 0 && dList.SelectedIndex < 0)
                    {
                        for (int i = 0; i < al.Count; i++)
                        {
                            BranchStore b = (BranchStore)al[i];
                            dList.Items.Add(new ListItem(b.department, b.id.ToString()));
                            if (b.ident_dep == "select") dList.SelectedIndex = i;
                        }
                        getBranchLabel().Visible = true;
                        getBranch().Visible = true;
                        auth = false;
                        LoginUser.FailureText = "Необходимо указать подразделение";
                        String id_pasword=LoginUser.FindControl("Password").ClientID;
                        String id_paswordLabel=LoginUser.FindControl("PasswordLabel").ClientID;
                        String id_branch=LoginUser.FindControl("Branch").ClientID;
                        ClientScript.RegisterClientScriptBlock(GetType(), "focus42", 
                            "<script type='text/javascript'>$(document).ready(function(){ " +
                            "$('#" + id_pasword + "').val('" + LoginUser.Password + "');" +
                            "$('#" + id_branch + "').focus();" +
                            //"$('#" + id_paswordLabel + "').css('display','none');" + 
                            //"$('#" + id_pasword + "').css('display','none');" + 
                            //"$('#" + id_pasword + "').attr('disabled', true);" + 
                            "});</script>");
                    }
                }
                else
                {
                    getBranchLabel().Visible = false;
                    getBranch().Visible = false;
                }
                e.Authenticated = auth;
                
            }
             
        }

        protected void LoginUser_LoggedIn(object sender, EventArgs e)
        {
            lock (OstCard.Data.Database.lockObjectDB)
            {
                DataSet ds = new DataSet();
                OstCard.Data.Database.ExecuteQuery(String.Format("select UserId, ActivePassword, id from aspnet_Users where UserName='{0}'", LoginUser.UserName), ref ds, null);
                if (ds.Tables.Count > 0 && ds.Tables[0].Rows.Count > 0)
                {
                    OstCard.Data.Database.Log(sc.UserGuid(LoginUser.UserName), "Вход в систему", null);
                    if (!Convert.ToBoolean(ds.Tables[0].Rows[0]["ActivePassword"]))
                        Response.Redirect("~\\Account\\ChangePassword.aspx");
                }
            }
        }


        /*
        [System.Web.Services.WebMethod]
        string getBranchesForUser(string user)
        {
            return "";
        }
        */
         


    }
}
