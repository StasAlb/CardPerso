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
using System.Data.SqlClient;
using System.Web.Configuration;
using CardPerso.Administration;
using System.Collections.Generic;

namespace CardPerso
{
    public partial class Branch : System.Web.UI.Page
    {
        string res = "";
        DataSet ds = new DataSet();
        ServiceClass sc = new ServiceClass();
        bool isAccountBranchView = false;
        bool isAccountBranchEdit = false;
        int branch_main_filial = 0;
        int branch_current = 0;
        bool perso = false;
        public String oprdaystart;
        public String oprdayend;
        public String raschet;

        public bool isShowAccount()
        {
            return bShowAccount.Checked;
        }

        public bool isAccountSave()
        {
            return (bShowAccount.Checked && isAccountBranchEdit);
        }

        public string getDepartment()
        {
            if (gvBranchs.Rows.Count < 1) return "";
            if(Convert.ToInt32(gvBranchs.DataKeys[Convert.ToInt32(gvBranchs.SelectedIndex)].Values["id_parent"])<1)
                return gvBranchs.DataKeys[Convert.ToInt32(gvBranchs.SelectedIndex)].Values["department"].ToString();
            else
                return gvBranchs.DataKeys[Convert.ToInt32(gvBranchs.SelectedIndex)].Values["office"].ToString();
        }

        public bool isMainFilial()
        {
            if (gvBranchs.Rows.Count < 1) return false;
            if (Convert.ToInt32(gvBranchs.DataKeys[Convert.ToInt32(gvBranchs.SelectedIndex)].Values["id_parent"]) > 0)
            {
                return false;
            }
            return true;
        }

        protected void Page_Load(object sender, EventArgs e)
        {


            isAccountBranchView = sc.UserAction(User.Identity.Name, Restrictions.AccountBranchView);
            
            if (!sc.UserAction(User.Identity.Name, Restrictions.LibraryView) && isAccountBranchView==false)
                Response.Redirect("~\\Account\\Restricted.aspx", true);
            ClientScript.RegisterHiddenField("resd", "");
            if (FuncClass.ClientType != ClientType.AkBars)
            {
                isAccountBranchView = false;
                isAccountBranchEdit = false;
            }

            perso = sc.UserAction(User.Identity.Name, Restrictions.Perso);
            branch_current = sc.BranchId(User.Identity.Name);

            branch_main_filial = BranchStore.getBranchMainFilial(branch_current, perso);

            isAccountBranchEdit = sc.UserAction(User.Identity.Name, Restrictions.AccountBranchEdit);

            if (isAccountBranchView == false) isAccountBranchEdit = false;

            lbInform.Text = "";

            oprdaystart = operDayStart.ClientID;
            oprdayend = operDayEnd.ClientID;
            
            if (IsPostBack)
                return;
            
            lock(Database.lockObjectDB)
            {
                lbInform.Text = "";
                Refr(0);
            }
            raschet = "";
        }

        private void Refr(int rowindex)
        {
                lbInform.Text = "";
                ds.Clear();
   
                if(branch_main_filial>0)
                    res = Database.ExecuteQuery("select * from V_Branchs where (id_parent=0 and id=" + branch_main_filial.ToString() + ") or id_parent=" + branch_main_filial.ToString() + " order by ident_dep_sort,office", ref ds, null);
                else
                    res = Database.ExecuteQuery("select * from V_Branchs order by ident_dep_sort,office", ref ds, null);
                gvBranchs.DataSource = ds.Tables[0];
                gvBranchs.DataBind();

                if (gvBranchs.Rows.Count > 0)
                {
                    gvBranchs.SelectedIndex = rowindex;
                    gvBranchs.Rows[gvBranchs.SelectedIndex].Focus();
                    if (bShowAccount.Checked == true)
                    {
                        int id = Convert.ToInt32(gvBranchs.DataKeys[Convert.ToInt32(gvBranchs.SelectedIndex)].Values["id"]);
                        getAccountBranch(getAccountBranch(id));
                        getOperDay(id);
                    }
                }
                else clearAccountFields();
                
                SetButton();

                lbCount.Text = "Кол-во: " + gvBranchs.Rows.Count.ToString();
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
                    ep.ExportGridExcel(gvBranchs);
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

        protected void bNew_Click(object sender, ImageClickEventArgs e)
        {
            lock (Database.lockObjectDB)
            {
                Refr(0);
            }
        }

        protected void gvBranchs_SelectedIndexChanged(object sender, EventArgs e)
        {
            lock (Database.lockObjectDB)
            {
                gvBranchs.Rows[gvBranchs.SelectedIndex].Focus();
                SetButton();
                if (bShowAccount.Checked == true)
                {
                    int id = Convert.ToInt32(gvBranchs.DataKeys[Convert.ToInt32(gvBranchs.SelectedIndex)].Values["id"]);
                    getAccountBranch(getAccountBranch(id));
                    getOperDay(id);
                }
            }
        }

        private void SetButton()
        {
            bAddOffice.Visible = false;
            bDelOffice.Visible = false;
            {
                bShowAccount.Visible = isAccountBranchView;
                mastercardInDebet.Enabled = isAccountBranchEdit;
                mastercardInCredit.Enabled = isAccountBranchEdit;
                mastercardOutDebet.Enabled = isAccountBranchEdit;
                mastercardOutCredit.Enabled = isAccountBranchEdit;
                mastercardRetDebet.Enabled = isAccountBranchEdit;
                mastercardRetCredit.Enabled = isAccountBranchEdit;
                visaInDebet.Enabled = isAccountBranchEdit;
                visaInCredit.Enabled = isAccountBranchEdit;
                visaOutDebet.Enabled = isAccountBranchEdit;
                visaOutCredit.Enabled = isAccountBranchEdit;
                visaRetDebet.Enabled = isAccountBranchEdit;
                visaRetCredit.Enabled = isAccountBranchEdit;
                nfcInDebet.Enabled = isAccountBranchEdit;
                nfcInCredit.Enabled = isAccountBranchEdit;
                nfcOutDebet.Enabled = isAccountBranchEdit;
                nfcOutCredit.Enabled = isAccountBranchEdit;
                nfcRetDebet.Enabled = isAccountBranchEdit;
                nfcRetCredit.Enabled = isAccountBranchEdit;
                srvInDebet.Enabled = isAccountBranchEdit;
                srvInCredit.Enabled = isAccountBranchEdit;
                srvOutDebet.Enabled = isAccountBranchEdit;
                srvOutCredit.Enabled = isAccountBranchEdit;
                srvRetDebet.Enabled = isAccountBranchEdit;
                srvRetCredit.Enabled = isAccountBranchEdit;
                pinInDebet.Enabled = isAccountBranchEdit;
                pinInCredit.Enabled = isAccountBranchEdit;
                pinOutDebet.Enabled = isAccountBranchEdit;
                pinOutCredit.Enabled = isAccountBranchEdit;
                pinRetDebet.Enabled = isAccountBranchEdit;
                pinRetCredit.Enabled = isAccountBranchEdit;
                mirInDebet.Enabled = isAccountBranchEdit;
                mirInCredit.Enabled = isAccountBranchEdit;
                mirOutDebet.Enabled = isAccountBranchEdit;
                mirOutCredit.Enabled = isAccountBranchEdit;
                mirRetDebet.Enabled = isAccountBranchEdit;
                mirRetCredit.Enabled = isAccountBranchEdit;
            }

            if (!sc.UserAction(User.Identity.Name, Restrictions.LibraryEdit))
            {
                bNew.Visible = false;
                bEdit.Visible = false;
                bDelete.Visible = false;
                bExcel.Visible = false;
                return;
            }

            if (gvBranchs.Rows.Count > 0)
            {
                bEdit.Visible = true;
                bDelete.Visible = true;
                bExcel.Visible = true;
                bShowAccount.Visible = isAccountBranchView;

                if (Convert.ToInt32(gvBranchs.DataKeys[gvBranchs.SelectedIndex].Values["id_parent"]) == 0)
                {
                    bAddOffice.Visible = true;
                    bAddOffice.Attributes.Add("OnClick", String.Format("return show_branch_office('id={0}')", gvBranchs.DataKeys[Convert.ToInt32(gvBranchs.SelectedIndex)].Values["id"].ToString()));
                }
                else
                {
                    bDelOffice.Visible = true;
                    bDelOffice.Attributes.Add("OnClick", String.Format("return confirm('Отвязать офис {0}?');", gvBranchs.DataKeys[Convert.ToInt32(gvBranchs.SelectedIndex)].Values["office"].ToString()));
                }

                bDelete.Attributes.Add("OnClick", String.Format("return confirm('Удалить подразделение {0}?');", gvBranchs.DataKeys[Convert.ToInt32(gvBranchs.SelectedIndex)].Values["department"].ToString()));
                bEdit.Attributes.Add("OnClick", String.Format("return show_branch('mode=2&id={0}')", gvBranchs.DataKeys[Convert.ToInt32(gvBranchs.SelectedIndex)].Values["id"].ToString()));
            }
            else
            {
                bEdit.Visible = false;
                bDelete.Visible = false;
                bExcel.Visible = false;
                bShowAccount.Visible = false;
            }
        }

        protected void bEdit_Click(object sender, ImageClickEventArgs e)
        {
            lock (Database.lockObjectDB)
            {
                Refr(gvBranchs.SelectedIndex);
            }
        }

        protected void bDelete_Click(object sender, ImageClickEventArgs e)
        {
            lock (Database.lockObjectDB)
            {
                int id = Convert.ToInt32(gvBranchs.DataKeys[Convert.ToInt32(gvBranchs.SelectedIndex)].Values["id"]);

                if (!Database.CheckDelBranch(id, null))
                {
                    lbInform.Text = "Невозможно удалить филиал, так как существуют карты.";
                    return;
                }

                SqlCommand sqCom = new SqlCommand();
                try
                {
                    sqCom.CommandText = "delete from Branchs where id=@id";
                    sqCom.Parameters.Add("@id", SqlDbType.Int).Value = id;
                    res = Database.ExecuteNonQuery(sqCom, null);

                    sqCom.CommandText = "update Branchs set id_parent=0 where id_parent=@id";
                    sqCom.Parameters["@id"].Value = id;
                    Database.ExecuteNonQuery(sqCom, null);
                }
                catch
                {
                }
                lbInform.Text = "";
                Refr(0);
            }
        }

        protected void bDelOffice_Click(object sender, ImageClickEventArgs e)
        {
            lock (Database.lockObjectDB)
            {
                int id = Convert.ToInt32(gvBranchs.DataKeys[Convert.ToInt32(gvBranchs.SelectedIndex)].Values["id"]);
                SqlCommand sqCom = new SqlCommand();
                sqCom.CommandText = "update Branchs set id_parent=0 where id=@id";
                sqCom.Parameters.Add("@id", SqlDbType.Int).Value = id;
                res = Database.ExecuteNonQuery(sqCom, null);

                lbInform.Text = "";
                Refr(0);
            }
        }

        protected void bAddOffice_Click(object sender, ImageClickEventArgs e)
        {
            lock (Database.lockObjectDB)
            {
                lbInform.Text = "";
                Refr(gvBranchs.SelectedIndex);
            }
        }

        protected void bAcountNumberSave_Click(object sender, ImageClickEventArgs e)
        {
            lbInform.Text = "";
            int id = Convert.ToInt32(gvBranchs.DataKeys[Convert.ToInt32(gvBranchs.SelectedIndex)].Values["id"]);
            List<AccountBranch> list = getAccountBranch();
            saveAccountBranch(id,list);
            if(isMainFilial()==true || (isMainFilial()==false && bUseMainBranch.Checked==true) || (isMainFilial()==false && bUseMainBranch.Checked==false))
            {
                try
                {
                    OperationDay op = new OperationDay();
                    //if (bUseMainBranch.Checked == true) op.parent = false;
                    //else op.parent = true;
                    op.operdays=operDayStart.Text;
                    op.operdaye=operDayEnd.Text;
                    if (tomorrow.Checked == true) op.isTomorrow = true;
                    else op.isTomorrow = false;
                    if (isMainFilial() == false && bUseMainBranch.Checked == false)
                    {
                        op.operdays = "";
                        op.operdaye = "";
                        op.isTomorrow = true;
                    }
                    op.isShift = bUseShift.Checked;
                    if (bUseShift.Checked == true)
                    {
                        op.shift1s = shift1s.Text;
                        op.shift1e = shift1e.Text;
                        op.shift2s = shift2s.Text;
                        op.shift2e = shift2e.Text;
                    }
                    op.write(id);
                    getOperDay(id);
                }
                catch (Exception e1)
                {
                    ClientScript.RegisterClientScriptBlock(GetType(), "errOperDay", "<script type='text/javascript'>$(document).ready(function(){ ShowError('" + e1.Message + "');});</script>");
                }
            }
            
        }

        protected void bShowAccount_CheckedChanged(object sender, EventArgs e)
        {
            if (bShowAccount.Checked == true)
            {
                if (gvBranchs.Rows.Count > 0)
                {
                    int id = Convert.ToInt32(gvBranchs.DataKeys[Convert.ToInt32(gvBranchs.SelectedIndex)].Values["id"]);
                    getAccountBranch(getAccountBranch(id));
                    getOperDay(id);
                }
            }
        }

        protected bool getOperDay(int id)
        {
            OperationDay op = new OperationDay();
            op.read(id);
            if (op.isParent() == true) bUseMainBranch.Checked = false;
            else bUseMainBranch.Checked = true;
            operDayStart.Text = (op.operdays == null) ? "" : op.operdays;
            operDayEnd.Text = (op.operdaye == null) ? "" : op.operdaye;
            if (op.isTomorrow == true)
            {
                today.Checked = false;
                tomorrow.Checked = true;
            }
            else
            {
                today.Checked = true;
                tomorrow.Checked = false;
            }
            if (isMainFilial() == true)
            {
                operDayStart.Enabled = true;
                operDayEnd.Enabled = true;
            }
            else
            {
                operDayStart.Enabled = bUseMainBranch.Checked;
                operDayEnd.Enabled = bUseMainBranch.Checked;
            }
            DateTime dt1=op.getDateTimeStart(DateTime.Now);
            DateTime dt2 = op.getDateTimeEnd(DateTime.Now);

            bUseShift.Checked = op.isShift;
            shift1s.Text = (op.shift1s == null) ? "" : op.shift1s;
            shift1e.Text = (op.shift1e == null) ? "" : op.shift1e;
            shift2s.Text = (op.shift2s == null) ? "" : op.shift2s;
            shift2e.Text = (op.shift2e == null) ? "" : op.shift2e;
            shift1s.Enabled = op.isShift;
            shift1e.Enabled = op.isShift;
            shift2s.Enabled = op.isShift;
            shift2e.Enabled = op.isShift;

            raschet = op.getMessagePart();
            
            return (dt1 < dt2);

        }


        protected List<AccountBranch> getAccountBranch(int branchId)
        {
            List<AccountBranch> listAB = new List<AccountBranch>();
            ds.Clear();
            res = Database.ExecuteQuery(String.Format("select * from BranchAccount where id_branch={0}",branchId), ref ds, null);
            if (ds != null && ds.Tables.Count > 0)
            {
                for (int i = 0; i < ds.Tables[0].Rows.Count; i++)
                {
                    int pt=Convert.ToInt32(ds.Tables[0].Rows[i]["product_type"]);
                    if (pt!=0 && Enum.IsDefined(typeof(BaseProductType), pt))
                    {
                        int at = Convert.ToInt32(ds.Tables[0].Rows[i]["account_type"]);
                        if (at != 0 && Enum.IsDefined(typeof(AccountBranchType), at))
                        {
                            AccountBranch ab = new AccountBranch();
                            ab.productType = (BaseProductType)pt;
                            ab.accountType = (AccountBranchType)at;
                            ab.accountDebet = ds.Tables[0].Rows[i]["account_debet"].ToString();
                            ab.accountCredit = ds.Tables[0].Rows[i]["account_credit"].ToString();
                            listAB.Add(ab);
                        }
                    }
                }
            }
            return listAB;
        }

        protected bool saveAccountBranch(int branchId, List<AccountBranch> listAB)
        {
            if (listAB == null || listAB.Count < 1) return true;
            lock (Database.lockObjectDB)
            {
                SqlCommand sqCom = new SqlCommand();
                for (int i = 0; i < listAB.Count; i++)
                {
                    sqCom.Parameters.Clear();
                    object obj = null;
                    Database.ExecuteScalar(String.Format("select top 1 id from BranchAccount where id_branch={0} and product_type={1} and account_type={2}", 
                        branchId,(int)listAB[i].productType,(int)listAB[i].accountType), ref obj, null);
                    if (obj != null)
                    {
                        int id = Convert.ToInt32(obj);
                        try
                        {
                            sqCom.CommandText = "update BranchAccount set account_debet=@debet, account_credit=@credit where id=@id";
                            sqCom.Parameters.Add("@id",SqlDbType.Int).Value = id;
                            sqCom.Parameters.Add("@debet", SqlDbType.VarChar,32).Value = listAB[i].accountDebet;
                            sqCom.Parameters.Add("@credit", SqlDbType.VarChar,32).Value = listAB[i].accountCredit;
                            Database.ExecuteNonQuery(sqCom, null);
                        }
                        catch { }

                    }
                    else
                    {
                        try
                        {
                            sqCom.CommandText = "insert into BranchAccount (id_branch,product_type,account_type,account_debet,account_credit) values (@branch,@product,@account,@debet,@credit)";
                            sqCom.Parameters.Add("@branch", SqlDbType.Int).Value = branchId;
                            sqCom.Parameters.Add("@product", SqlDbType.Int).Value = (int)listAB[i].productType;
                            sqCom.Parameters.Add("@account", SqlDbType.Int).Value = (int)listAB[i].accountType;
                            sqCom.Parameters.Add("@debet", SqlDbType.VarChar).Value = listAB[i].accountDebet;
                            sqCom.Parameters.Add("@credit", SqlDbType.VarChar).Value = listAB[i].accountCredit;
                            Database.ExecuteNonQuery(sqCom, null);
                        }
                        catch { }
                    
                    }
                }
            }
            return true;
        }

        protected void clearAccountFields()
        {
            mastercardInDebet.Text="";
            mastercardInCredit.Text="";
            mastercardOutDebet.Text="";
            mastercardOutCredit.Text="";
            mastercardRetDebet.Text = "";
            mastercardRetCredit.Text = "";
            visaInDebet.Text = "";
            visaInCredit.Text = "";
            visaOutDebet.Text = "";
            visaOutCredit.Text = "";
            visaRetDebet.Text = "";
            visaRetCredit.Text = "";
            nfcInDebet.Text = "";
            nfcInCredit.Text = "";
            nfcOutDebet.Text = "";
            nfcOutCredit.Text = "";
            nfcRetDebet.Text = "";
            nfcRetCredit.Text = "";
            srvInDebet.Text = "";
            srvInCredit.Text = "";
            srvOutDebet.Text = "";
            srvOutCredit.Text = "";
            srvRetDebet.Text = "";
            srvRetCredit.Text = "";
            pinInDebet.Text = "";
            pinInCredit.Text = "";
            pinOutDebet.Text = "";
            pinOutCredit.Text = "";
            pinRetDebet.Text = "";
            pinRetCredit.Text = "";
            mirInDebet.Text = "";
            mirInCredit.Text = "";
            mirOutDebet.Text = "";
            mirOutCredit.Text = "";
            mirRetDebet.Text = "";
            mirRetCredit.Text = "";
        }

        protected void setAccountBranch(AccountBranch ab)
        {
            if (ab.productType == BaseProductType.MasterCard)
            {
                if (ab.accountType == AccountBranchType.In)
                {
                    mastercardInDebet.Text=ab.accountDebet;
                    mastercardInCredit.Text=ab.accountCredit;
                }
                if (ab.accountType == AccountBranchType.Out)
                {
                    mastercardOutDebet.Text=ab.accountDebet;
                    mastercardOutCredit.Text=ab.accountCredit;
                }
                if (ab.accountType == AccountBranchType.Return)
                {
                    mastercardRetDebet.Text=ab.accountDebet;
                    mastercardRetCredit.Text=ab.accountCredit;
                }
            }
            if (ab.productType == BaseProductType.VisaCard)
            {
                if (ab.accountType == AccountBranchType.In)
                {
                    visaInDebet.Text=ab.accountDebet;
                    visaInCredit.Text=ab.accountCredit;
                }
                if (ab.accountType == AccountBranchType.Out)
                {
                    visaOutDebet.Text=ab.accountDebet;
                    visaOutCredit.Text=ab.accountCredit;
                }
                if (ab.accountType == AccountBranchType.Return)
                {
                    visaRetDebet.Text=ab.accountDebet;
                    visaRetCredit.Text=ab.accountCredit;
                }
            }
            if (ab.productType == BaseProductType.NFCCard)
            {
                if (ab.accountType == AccountBranchType.In)
                {
                    nfcInDebet.Text=ab.accountDebet;
                    nfcInCredit.Text=ab.accountCredit;
                }
                if (ab.accountType == AccountBranchType.Out)
                {
                    nfcOutDebet.Text=ab.accountDebet;
                    nfcOutCredit.Text=ab.accountCredit;
                }
                if (ab.accountType == AccountBranchType.Return)
                {
                    nfcRetDebet.Text=ab.accountDebet;
                    nfcRetCredit.Text=ab.accountCredit;
                }
            }
            if (ab.productType == BaseProductType.ServiceCard)
            {
                if (ab.accountType == AccountBranchType.In)
                {
                    srvInDebet.Text=ab.accountDebet;
                    srvInCredit.Text=ab.accountCredit;
                }
                if (ab.accountType == AccountBranchType.Out)
                {
                    srvOutDebet.Text=ab.accountDebet;
                    srvOutCredit.Text = ab.accountCredit;
                }
                if (ab.accountType == AccountBranchType.Return)
                {
                    srvRetDebet.Text=ab.accountDebet;
                    srvRetCredit.Text=ab.accountCredit;
                }
            }
            if (ab.productType == BaseProductType.PinConvert)
            {
                if (ab.accountType == AccountBranchType.In)
                {
                    pinInDebet.Text=ab.accountDebet;
                    pinInCredit.Text=ab.accountCredit;
                }
                if (ab.accountType == AccountBranchType.Out)
                {
                    pinOutDebet.Text=ab.accountDebet;
                    pinOutCredit.Text=ab.accountCredit;
                }
                if (ab.accountType == AccountBranchType.Return)
                {
                    pinRetDebet.Text=ab.accountDebet;
                    pinRetCredit.Text=ab.accountCredit;
                }
                
            }
            if (ab.productType == BaseProductType.MirCard)
            {
                if (ab.accountType == AccountBranchType.In)
                {
                    mirInDebet.Text = ab.accountDebet;
                    mirInCredit.Text = ab.accountCredit;
                }
                if (ab.accountType == AccountBranchType.Out)
                {
                    mirOutDebet.Text = ab.accountDebet;
                    mirOutCredit.Text = ab.accountCredit;
                }
                if (ab.accountType == AccountBranchType.Return)
                {
                    mirRetDebet.Text = ab.accountDebet;
                    mirRetCredit.Text = ab.accountCredit;
                }
            }
        }

        protected void getAccountBranch(List<AccountBranch> listAB)
        {
            clearAccountFields();
            if (listAB != null)
            {
                for (int i = 0; i < listAB.Count; i++)
                {
                    setAccountBranch(listAB[i]);
                }
            }
        }

        protected List<AccountBranch> getAccountBranch()
        {
            List<AccountBranch> listAB = new List<AccountBranch>();
            
            AccountBranch ab = new AccountBranch();
            ab.productType = BaseProductType.MasterCard;
            ab.accountType = AccountBranchType.In;
            ab.accountDebet = mastercardInDebet.Text;
            ab.accountCredit = mastercardInCredit.Text;
            listAB.Add(ab);
            ab = new AccountBranch();
            ab.productType = BaseProductType.MasterCard;
            ab.accountType = AccountBranchType.Out;
            ab.accountDebet = mastercardOutDebet.Text;
            ab.accountCredit = mastercardOutCredit.Text;
            listAB.Add(ab);
            ab = new AccountBranch();
            ab.productType = BaseProductType.MasterCard;
            ab.accountType = AccountBranchType.Return;
            ab.accountDebet = mastercardRetDebet.Text;
            ab.accountCredit = mastercardRetCredit.Text;
            listAB.Add(ab);

            ab = new AccountBranch();
            ab.productType = BaseProductType.VisaCard;
            ab.accountType = AccountBranchType.In;
            ab.accountDebet = visaInDebet.Text;
            ab.accountCredit = visaInCredit.Text;
            listAB.Add(ab);
            ab = new AccountBranch();
            ab.productType = BaseProductType.VisaCard;
            ab.accountType = AccountBranchType.Out;
            ab.accountDebet = visaOutDebet.Text;
            ab.accountCredit = visaOutCredit.Text;
            listAB.Add(ab);
            ab = new AccountBranch();
            ab.productType = BaseProductType.VisaCard;
            ab.accountType = AccountBranchType.Return;
            ab.accountDebet = visaRetDebet.Text;
            ab.accountCredit = visaRetCredit.Text;
            listAB.Add(ab);

            ab = new AccountBranch();
            ab.productType = BaseProductType.NFCCard;
            ab.accountType = AccountBranchType.In;
            ab.accountDebet = nfcInDebet.Text;
            ab.accountCredit = nfcInCredit.Text;
            listAB.Add(ab);
            ab = new AccountBranch();
            ab.productType = BaseProductType.NFCCard;
            ab.accountType = AccountBranchType.Out;
            ab.accountDebet = nfcOutDebet.Text;
            ab.accountCredit = nfcOutCredit.Text;
            listAB.Add(ab);
            ab = new AccountBranch();
            ab.productType = BaseProductType.NFCCard;
            ab.accountType = AccountBranchType.Return;
            ab.accountDebet = nfcRetDebet.Text;
            ab.accountCredit = nfcRetCredit.Text;
            listAB.Add(ab);

            ab = new AccountBranch();
            ab.productType = BaseProductType.ServiceCard;
            ab.accountType = AccountBranchType.In;
            ab.accountDebet = srvInDebet.Text;
            ab.accountCredit = srvInCredit.Text;
            listAB.Add(ab);
            ab = new AccountBranch();
            ab.productType = BaseProductType.ServiceCard;
            ab.accountType = AccountBranchType.Out;
            ab.accountDebet = srvOutDebet.Text;
            ab.accountCredit = srvOutCredit.Text;
            listAB.Add(ab);
            ab = new AccountBranch();
            ab.productType = BaseProductType.ServiceCard;
            ab.accountType = AccountBranchType.Return;
            ab.accountDebet = srvRetDebet.Text;
            ab.accountCredit = srvRetCredit.Text;
            listAB.Add(ab);

            ab = new AccountBranch();
            ab.productType = BaseProductType.PinConvert;
            ab.accountType = AccountBranchType.In;
            ab.accountDebet = pinInDebet.Text;
            ab.accountCredit = pinInCredit.Text;
            listAB.Add(ab);
            ab = new AccountBranch();
            ab.productType = BaseProductType.PinConvert;
            ab.accountType = AccountBranchType.Out;
            ab.accountDebet = pinOutDebet.Text;
            ab.accountCredit = pinOutCredit.Text;
            listAB.Add(ab);
            ab = new AccountBranch();
            ab.productType = BaseProductType.PinConvert;
            ab.accountType = AccountBranchType.Return;
            ab.accountDebet = pinRetDebet.Text;
            ab.accountCredit = pinRetCredit.Text;
            listAB.Add(ab);

            ab = new AccountBranch();
            ab.productType = BaseProductType.MirCard;
            ab.accountType = AccountBranchType.In;
            ab.accountDebet = mirInDebet.Text;
            ab.accountCredit = mirInCredit.Text;
            listAB.Add(ab);
            ab = new AccountBranch();
            ab.productType = BaseProductType.MirCard;
            ab.accountType = AccountBranchType.Out;
            ab.accountDebet = mirOutDebet.Text;
            ab.accountCredit = mirOutCredit.Text;
            listAB.Add(ab);
            ab = new AccountBranch();
            ab.productType = BaseProductType.MirCard;
            ab.accountType = AccountBranchType.Return;
            ab.accountDebet = mirRetDebet.Text;
            ab.accountCredit = mirRetCredit.Text;
            listAB.Add(ab);


            return listAB;
        }

        protected void bUseMainBranch_CheckedChanged(object sender, EventArgs e)
        {
            operDayStart.Text = "";
            operDayEnd.Text = "";
            today.Checked = false;
            tomorrow.Checked = true;
            if (isMainFilial() == true)
            {
                operDayStart.Enabled = true;
                operDayEnd.Enabled = true;
            }
            else
            {
                operDayStart.Enabled = bUseMainBranch.Checked;
                operDayEnd.Enabled = bUseMainBranch.Checked;
            }
            
        }

        protected void bUseShift_CheckedChanged(object sender, EventArgs e)
        {
            shift1s.Enabled = bUseShift.Checked;
            shift1e.Enabled = bUseShift.Checked;
            shift2s.Enabled = bUseShift.Checked;
            shift2e.Enabled = bUseShift.Checked;
        }
       
    }
}
