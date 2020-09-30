using System;
using System.Collections;
using System.Configuration;
using System.Data;
using System.Globalization;
using System.Linq;
using System.Web;
using System.Web.Security;
using System.Web.UI;
using System.Web.UI.WebControls;
using System.Web.UI.WebControls.WebParts;
using System.Web.UI.HtmlControls;
using System.Xml.Linq;
using System.Web.Configuration;
using CardPerso.Administration;

namespace CardPerso
{
    public partial class AccountablePerson : System.Web.UI.Page
    {
        private CommonDataBase db;
        ServiceClass sc = new ServiceClass();

        private int selectID()
        {
            int id = Convert.ToInt32(gvAccountablePersons.DataKeys[gvAccountablePersons.SelectedIndex].Values["id"]);
            AccountablePersonData apd = db.getAccountablePersonData(id);
            fillToForm(apd);
            return gvAccountablePersons.SelectedIndex;
        }

        protected void Page_Load(object sender, EventArgs e)
        {
            if (!sc.UserAction(User.Identity.Name, Restrictions.LibraryPodotchet))
                Response.Redirect("~\\Account\\Restricted.aspx", true);
            db = new CommonDataBase(null);
            if (!IsPostBack)
            {
               RefreshList(0);
               if (gvAccountablePersons.Rows.Count > 0)
               {
                   try
                   {
                       selectID();
                   }
                   catch (Exception e1)
                   {
                       ClientScript.RegisterClientScriptBlock(GetType(), "bPage_Load", "<script type='text/javascript'>$(document).ready(function(){ " +
                            "ShowError('" + System.Security.SecurityElement.Escape(e1.Message) + "'); });</script>");
                   }
               }
            }
        }

        private void RefreshList(int rowindex)
        {
            lbInform.Text = "";
            gvAccountablePersons.DataSource = db.getAccountablePersonsDataset().Tables[0];
            gvAccountablePersons.DataBind();

            if (gvAccountablePersons.Rows.Count > 0)
            {
                gvAccountablePersons.SelectedIndex = rowindex;
                gvAccountablePersons.Rows[gvAccountablePersons.SelectedIndex].Focus();
                bEdit.Visible = true;
                bDelete.Visible = true;
                bExcel.Visible = true;
            }
            else
            {
                bEdit.Visible = false;
                bDelete.Visible = false;
                bExcel.Visible = false;
            }

            lbCount.Text = "Кол-во: " + gvAccountablePersons.Rows.Count.ToString();
        }

        protected void gvAccountablePersons_SelectedIndexChanged(object sender, EventArgs e)
        {
            try
            {

                gvAccountablePersons.Rows[gvAccountablePersons.SelectedIndex].Focus();
                selectID();
            }
            catch (Exception e1)
            {
                ClientScript.RegisterClientScriptBlock(GetType(), "bSelectedIndexChanged", "<script type='text/javascript'>$(document).ready(function(){ " +
                     "ShowError('" + System.Security.SecurityElement.Escape(e1.Message) + "'); });</script>");
            }
        }
      
        protected void bNewAccountablePerson_Click(object sender, EventArgs e)
        {
            try
            {
                AccountablePersonData apd = fillFromForm();
                int idBranch = sc.BranchId(Page.User.Identity.Name);
                db.saveAccountablePersonData(apd, idBranch);
                RefreshList(0);
                selectID();
            }
            catch (Exception e1)
            {
                ClientScript.RegisterClientScriptBlock(GetType(), "bNew_ReClick", "<script type='text/javascript'>$(document).ready(function(){ " +
                     "clickNew(); ShowError('" + System.Security.SecurityElement.Escape(e1.Message) + "'); });</script>");
            }
             
        }

        protected void bSaveAccountablePerson_Click(object sender, EventArgs e)
        {
            try
            {
                AccountablePersonData apd = fillFromForm();
                if (gvAccountablePersons.Rows.Count > 0)
                {
                    int id = Convert.ToInt32(gvAccountablePersons.DataKeys[gvAccountablePersons.SelectedIndex].Values["id"]);
                    AccountablePersonData apds = db.getAccountablePersonData(id);
                    for (int i = 0; i < apd.accountdata.Length; i++)
                    {
                        AccountablePersonAccountData apad = apds.accountdata.FirstOrDefault(r => r.id > 0 && r.product_type_enum == apd.accountdata[i].product_type_enum &&
                                                                                                  r.account_type_enum == apd.accountdata[i].account_type_enum &&
                                                                                                  r.issafe == apd.accountdata[i].issafe);
                        if (apad != null) apd.accountdata[i].id = apad.id;
                    }
                    apd.id = id;
                }
                int idBranch = sc.BranchId(Page.User.Identity.Name);
                db.saveAccountablePersonData(apd, idBranch);
                RefreshList(gvAccountablePersons.SelectedIndex);
                selectID();
            }
            catch (Exception e1)
            {
                ClientScript.RegisterClientScriptBlock(GetType(), "bSave_ReClick", "<script type='text/javascript'>$(document).ready(function(){ " +
                     "clickEdit(); ShowError('" + System.Security.SecurityElement.Escape(e1.Message) + "'); });</script>");
            } 
        }

        
        protected void bCancelAccountablePerson_Click(object sender, EventArgs e)
        {
            if (gvAccountablePersons.Rows.Count > 0)
            {
                try
                {
                    selectID();
                }
                catch (Exception e1)
                {
                    ClientScript.RegisterClientScriptBlock(GetType(), "bCancel_Click", "<script type='text/javascript'>$(document).ready(function(){ " +
                         "ShowError('" + System.Security.SecurityElement.Escape(e1.Message) + "'); });</script>");
                }
            }  
        }

        protected void bNew_Click(object sender, ImageClickEventArgs e)
        {
            AccountablePersonData apd = new AccountablePersonData();
            fillToForm(apd);
            ClientScript.RegisterClientScriptBlock(GetType(), "bNew_Click", "<script type='text/javascript'>$(document).ready(function(){ clickNew();});</script>");
        }

        protected AccountablePersonData fillFromForm()
        {
            AccountablePersonData apd = new AccountablePersonData();
            if (secondname.Text.Trim().Length < 1)
                throw new Exception("Поле 'Фамилия' не может быть пустым");
            if (firstname.Text.Trim().Length < 1)
                throw new Exception("Поле 'Имя' не может быть пустым");
            apd.secondname = secondname.Text;
            apd.firstname = firstname.Text;
            apd.patronymic = patronymic.Text;
            apd.passport = passport.Text;
            apd.userLogin = userLogin.Text;
            for (int i = 0; i < apd.accountdata.Length; i++)
            {
                string inpName = getDefaultName(apd.accountdata[i]);
                TextBox txt = getControlCredit(apd.accountdata[i]);
                if (txt != null) apd.accountdata[i].account_credit = txt.Text;
                txt = getControlDebet(apd.accountdata[i]);
                if (txt != null) apd.accountdata[i].account_debet = txt.Text;
            }
            return apd;
        }

        protected void fillToForm(AccountablePersonData apd)
        {
            secondname.Text = apd.secondname;
            firstname.Text = apd.firstname; 
            patronymic.Text = apd.patronymic;
            passport.Text = apd.passport;
            userLogin.Text = apd.userLogin;
            for (int i = 0; i < apd.accountdata.Length; i++)
            {
                string inpName = getDefaultName(apd.accountdata[i]);
                TextBox txt = getControlCredit(apd.accountdata[i]);
                if(txt!=null) txt.Text = apd.accountdata[i].account_credit;
                txt = getControlDebet(apd.accountdata[i]);
                if(txt!=null) txt.Text = apd.accountdata[i].account_debet;
            }
        }

        protected string getDefaultName(AccountablePersonAccountData aData)
        {
            return ((BaseProductType)(aData.product_type_enum)).ToString() + "_" + ((AccountBranchType)(aData.account_type_enum)).ToString() + "_" + aData.issafe.ToString();
        }

        protected TextBox getControlDebet(AccountablePersonAccountData aData)
        {
            return getControl(null, getDefaultName(aData) + "_Debet");
        }

        protected TextBox getControlCredit(AccountablePersonAccountData aData)
        {
            return getControl(null, getDefaultName(aData) + "_Credit");
        }


        protected TextBox getControl(ControlCollection c, string tName)
        {
            if(c==null) return getControl(Form.Controls, tName);
            if (c.Count < 1) return null;
            for(int i=0;i<c.Count;i++)
            {
                Control cc = c[i].FindControl(tName);
                if (cc != null) return (TextBox)cc;
                cc = getControl(c[i].Controls,tName);
                if (cc != null) return (TextBox)cc;
            }
            return null;
            
        }

        protected void bExcel_Click(object sender, ImageClickEventArgs e)
        {
            System.Globalization.CultureInfo oldCI = System.Threading.Thread.CurrentThread.CurrentCulture;
            System.Threading.Thread.CurrentThread.CurrentCulture = new System.Globalization.CultureInfo("en-US");
            string doc = "";
            ExcelAp ep = new ExcelAp();
            if (ep.RunApp(ConfigurationSettings.AppSettings["DocPath"] + "Empty.xls"))
            {
                ep.SetWorkSheet(1);
                ep.ExportGridExcel(gvAccountablePersons);
                if (WebConfigurationManager.AppSettings["DocPath"] != null)
                {
                    doc = String.Format("{0}Temp\\AccountablePersons.xls", WebConfigurationManager.AppSettings["DocPath"]);
                    ep.SaveAsDoc(doc, false);
                }
            }
            ep.Close();
            GC.Collect();
            System.Threading.Thread.CurrentThread.CurrentCulture = oldCI;
            if (doc.Length > 0)
                ep.ReturnXls(Response, doc);
        }

        
        protected void bDelAccountablePerson_Click(object sender, EventArgs e)
        {
            if (gvAccountablePersons.Rows.Count > 0)
            {
                try
                {
                    int id = Convert.ToInt32(gvAccountablePersons.DataKeys[gvAccountablePersons.SelectedIndex].Values["id"]);
                    db.delAccountablePersonData(id);
                    RefreshList(0);
                    selectID();
                }
                catch (Exception e1)
                {
                    ClientScript.RegisterClientScriptBlock(GetType(), "bDel_Click", "<script type='text/javascript'>$(document).ready(function(){ " +
                         "ShowError('" + System.Security.SecurityElement.Escape(e1.Message) + "'); });</script>");

                }
            }
        }



    }
}
