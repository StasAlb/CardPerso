using System;
using System.Collections;
using System.Configuration;
using System.Data;
using System.Linq;
using System.Net;
using System.Web;
using System.Web.Security;
using System.Web.UI;
using System.Web.UI.HtmlControls;
using System.Web.UI.WebControls;
using System.Web.UI.WebControls.WebParts;
using System.Xml.Linq;
using OstCard.Data;
using System.Data.SqlClient;
using CardPerso.Administration;
 
namespace CardPerso
{
    public partial class FltCard : System.Web.UI.Page
    {
        DataSet ds = new DataSet();
        string res = "";
        ServiceClass sc = new ServiceClass();
        bool alldata;
        int current_branch_id;
        int branch_main_filial;

        public bool getViewBranchCurrent()
        {
            int currentStatus=Convert.ToInt32(dListStatus.SelectedItem.Value);
            if(currentStatus==4 && alldata==true) return true;
            return false;
        }

        public string FIO_state
        {
            get
            {
                object o = ViewState["fio_state"];
                return (o == null) ? String.Empty : (string)o;
            }
            set
            {
                ViewState["fio_state"] = value;
            }
        }

        protected void Page_Load(object sender, EventArgs e)
        {
          

            lock (Database.lockObjectDB)
            {
                //alldata = sc.UserAction(User.Identity.Name, Restrictions.FilialDeliver);
                current_branch_id = sc.BranchId(User.Identity.Name);
                //if (current_branch_id == 106) alldata = true;
                //else alldata = false;
                int branchIdMain = BranchStore.getBranchMainFilial(current_branch_id, false);
                if (branchIdMain > 0 && branchIdMain == current_branch_id) alldata = true;
                else alldata = false;
                dListStatus.AutoPostBack = alldata;

                branch_main_filial = BranchStore.getBranchMainFilial(current_branch_id, sc.UserAction(User.Identity.Name, Restrictions.Perso)); 

                if(FIO_state.Trim().Length>0) tbFio.Text = FIO_state;
                
                if (!IsPostBack)
                {
                    ZapCombo();
                    tbNumber.Focus();
                    
                }
            }
        }


        private String getSchoolProducts()
        {
            
            String sqlSchoolProducts = "";
            if (ConfigurationSettings.AppSettings["SchoolProducts"] == null) return sqlSchoolProducts;
            String[] schoolProducts = ConfigurationSettings.AppSettings["SchoolProducts"].Split(',');
            for (int i = 0; i < schoolProducts.Length; i++)
            {
                String[] bin_prefix = schoolProducts[i].Split('_');
                if (sqlSchoolProducts.Length>0) sqlSchoolProducts += " or ";
                sqlSchoolProducts += "(bin='" + bin_prefix[0] + "' and prefix_ow='" + bin_prefix[1] + "')";
            }
            ds.Clear();
            
            if (sqlSchoolProducts.Length > 0)
                res = Database.ExecuteQuery("SELECT pb.id as id_prb FROM Products p join Products_Banks pb on p.id=pb.id_prod where " + sqlSchoolProducts, ref ds, null);
            String id_prb = "";
            for (int i = 0; ds.Tables.Count > 0 && i < ds.Tables[0].Rows.Count; i++)
            {
                if (id_prb.Length > 0) id_prb += ",";
                id_prb += (Convert.ToInt32(ds.Tables[0].Rows[i]["id_prb"])).ToString();
            }
            return id_prb;
        }

        private void ZapCombo()
        {
            bool action = sc.UserAction(User.Identity.Name, Restrictions.AllData);
                        
            ds.Clear();
            res = Database.ExecuteQuery("select id,name from Status", ref ds, null);
            dListStatus.Items.Add(new ListItem("Все", "-1"));
            for (int i = 0; i < ds.Tables[0].Rows.Count; i++)
                dListStatus.Items.Add(new ListItem(ds.Tables[0].Rows[i]["name"].ToString(), ds.Tables[0].Rows[i]["id"].ToString()));

          //  dListStatus.SelectedIndex = -1;
            ds.Clear();
            res = Database.ExecuteQuery("select id,prop from CardProperty", ref ds, null);
            ddlProp.Items.Add(new ListItem("Все", "-1"));
            for (int i = 0; i < ds.Tables[0].Rows.Count; i++)
                ddlProp.Items.Add(new ListItem(ds.Tables[0].Rows[i]["prop"].ToString(), ds.Tables[0].Rows[i]["id"].ToString()));

            string branchS = "";
            dListBranchProd.Items.Add(new ListItem("Все", "-1")); // филиалы где заведено, все 
            if (!action/*sc.UserAction(User.Identity.Name, Restrictions.AllData)*/)
                branchS = String.Format(" where (id={0} or id in (select id from Branchs where id_parent={0} or id_parenttr={0}))", sc.BranchId(User.Identity.Name));
            /*else*/
                dListBranchCard.Items.Add(new ListItem("Все", "-1"));
            
            ds.Clear();
            res = Database.ExecuteQuery("select id,department from Branchs " + branchS + " order by department", ref ds, null);
            for (int i = 0; i < ds.Tables[0].Rows.Count; i++)
            {
                dListBranchCard.Items.Add(new ListItem(ds.Tables[0].Rows[i]["department"].ToString(), ds.Tables[0].Rows[i]["id"].ToString()));
                dListBranchProd.Items.Add(new ListItem(ds.Tables[0].Rows[i]["department"].ToString(), ds.Tables[0].Rows[i]["id"].ToString()));
            }
            ds.Clear();
            res = Database.ExecuteQuery("select id, name from banks", ref ds, null);
            dListBank.Items.Add(new ListItem("Все", "-1"));
            for(int i=0;i<ds.Tables[0].Rows.Count;i++)
                dListBank.Items.Add(new ListItem(ds.Tables[0].Rows[i]["name"].ToString(), ds.Tables[0].Rows[i]["id"].ToString()));

            ds.Clear();
            res = Database.ExecuteQuery("select id_prb,(prod_name + ' (' + bank_name+')') as prod from V_ProductsBanks_T order by id_sort", ref ds, null);
            dListProd.Items.Add(new ListItem("Все", "-1"));
            for (int i = 0; i < ds.Tables[0].Rows.Count; i++)
                dListProd.Items.Add(new ListItem(ds.Tables[0].Rows[i]["prod"].ToString(), ds.Tables[0].Rows[i]["id_prb"].ToString()));

            String id_prb = getSchoolProducts();

            ds.Clear();
            //01.11.2018 отключил заполнение организаций. Дюже долго
            //res = Database.ExecuteQuery("select distinct company from Cards order by company", ref ds, null);
            //String sqlOrg = "select distinct company from Cards";
            //if (sc.UserAction(User.Identity.Name, Restrictions.AllData) == false)
            //{
            //    sqlOrg += String.Format(" where (id_branchCard={0} or id_branchCard in (select id from Branchs where id_parent={0} or id_parentTr={0}))", sc.BranchId(User.Identity.Name));
            //    if (id_prb.Length > 0) sqlOrg += " and id_prb not in(" + id_prb + ")";
            //}
            //else
            //{
            //    if (id_prb.Length > 0) sqlOrg += " where id_prb not in(" + id_prb + ")";
            //}
            //sqlOrg += " order by company";
            //res = Database.ExecuteQuery(sqlOrg, ref ds, null);
            dListOrgan.Items.Add(new ListItem("Все", "-1"));
            //for (int i = 0; i < ds.Tables[0].Rows.Count; i++)
            ////    dListOrgan.Items.Add(new ListItem(ds.Tables[0].Rows[i]["company"].ToString(), ds.Tables[0].Rows[i]["company"].ToString()));
            //    if(ds.Tables[0].Rows[i]["company"].ToString().Trim().Length>0)    
            //        dListOrgan.Items.Add(new ListItem(ds.Tables[0].Rows[i]["company"].ToString(), i.ToString()));

            DListSchoolProduct.Items.Add(new ListItem("Все", "-1"));
            if (id_prb.Length > 0)
            {
                ds.Clear();
                res = Database.ExecuteQuery("SELECT pb.id as id_prb,p.id,p.[name] as productname,p.prefix_ow  FROM Products p join Products_Banks pb on p.id=pb.id_prod where pb.id in(" + id_prb + ")", ref ds, null);
                for (int i = 0; ds.Tables.Count>0 && i<ds.Tables[0].Rows.Count; i++)
                    DListSchoolProduct.Items.Add(new ListItem(ds.Tables[0].Rows[i]["productname"].ToString(),ds.Tables[0].Rows[i]["id_prb"].ToString()));
            }

            ds.Clear();
            res = Database.ExecuteQuery("select distinct courier_name from Cards order by courier_name", ref ds, null);
            ddlWorker.Items.Add(new ListItem("Все", "-1"));
            for (int i = 0; i < ds.Tables[0].Rows.Count; i++)
                if (ds.Tables[0].Rows[i]["courier_name"].ToString().Trim().Length > 0)
                    ddlWorker.Items.Add(new ListItem(ds.Tables[0].Rows[i]["courier_name"].ToString(), ds.Tables[0].Rows[i]["courier_name"].ToString()));

            ds.Clear();
            res = Database.ExecuteQuery("select id,name from Couriers order by name", ref ds, null);
            ddlCourier.Items.Add(new ListItem("Все", "-1"));
            for (int i = 0; i < ds.Tables[0].Rows.Count; i++)
                ddlCourier.Items.Add(new ListItem(ds.Tables[0].Rows[i]["name"].ToString(), ds.Tables[0].Rows[i]["id"].ToString()));

            ds.Clear();
            res = Database.ExecuteQuery("select UserName from aspnet_Users order by UserName", ref ds, null);
            ddlClientWorker.Items.Add(new ListItem("Все", "-1"));
            int brId = sc.BranchId(User.Identity.Name);
            
            for(int i=0;i<ds.Tables[0].Rows.Count;i++)
            {
                if (action/*sc.UserAction(User.Identity.Name, Restrictions.AllData)*/ || sc.BranchId(ds.Tables[0].Rows[i]["UserName"].ToString()) == brId)
                    ddlClientWorker.Items.Add(new ListItem(ds.Tables[0].Rows[i]["UserName"].ToString(), ds.Tables[0].Rows[i]["UserName"].ToString()));
            }

         //   dListBranchCard.SelectedIndex = -1;

            Session["ZAPCOMBO"] = true;
        }
        private bool CheckDate(OstCard.WebControls.DatePicker tb, string lb)
        {
            if (tb.DatePickerText.Trim().Length > 0)
            {
                try
                {
                    Convert.ToDateTime(tb.DatePickerText.Trim());
                }
                catch
                {
                    lbInform.Text = "Неправильно введена дата " + lb;
                    tb.Focus();
                    return false;
                }

            }
            return true;
        }

        protected void dListStatus_SelectedIndexChanged(object sender, EventArgs e)
        {
            int currentStatus = Convert.ToInt32(dListStatus.SelectedItem.Value);
            if (currentStatus == 4)
            {
                dListBranchCurrent.Items.Clear();
                dListBranchCurrent.Items.Add(new ListItem("Все", "-1"));
                ds.Clear();
                //res = Database.ExecuteQuery("select id,department from branchs where id in (select id_branchcurrent from cards where id_branchcurrent is not null) order by ident_dep", ref ds, null);
                //res = Database.ExecuteQuery("select id,department from branchs where id=106", ref ds, null);
                res = Database.ExecuteQuery("select id,department from branchs where id in (select id_branchcurrent from cards where id_branchcurrent is not null) order by ident_dep", ref ds, null);
                for (int i = 0; i < ds.Tables[0].Rows.Count; i++)
                {
                    dListBranchCurrent.Items.Add(new ListItem(ds.Tables[0].Rows[i]["department"].ToString(), ds.Tables[0].Rows[i]["id"].ToString()));
                }
            }
            else dListBranchCurrent.Items.Clear();
        }

        protected void bSave_Click(object sender, ImageClickEventArgs e)
        {
            lock (Database.lockObjectDB)
            {
                ArrayList al = new ArrayList();

                string s = "";

                if (!CheckDate(DatePickerProd1, "изготовления с"))
                    return;
                if (!CheckDate(DatePickerProd2, "изготовления по"))
                    return;
                if (!CheckDate(DatePickerBegin1, "начала действия с"))
                    return;
                if (!CheckDate(DatePickerBegin2, "начала действия по"))
                    return;
                if (!CheckDate(DatePickerEnd1, "окончание действия с"))
                    return;
                if (!CheckDate(DatePickerEnd2, "окончание действия по"))
                    return;
                if (!CheckDate(DatePickerReceive1, "получения с"))
                    return;
                if (!CheckDate(DatePickerReceive1, "получения по"))
                    return;
                if (!CheckDate(DatePickerClient1, "выдачи с"))
                    return;
                if (!CheckDate(DatePickerClient2, "выдачи по"))
                    return;
                if (!CheckDate(DatePickerPered1, "передачи с"))
                    return;
                if (!CheckDate(DatePickerPered2, "передачи по"))
                    return;
                if (!CheckDate(DatePickerDestroy11, "об уничтожении с"))
                    return;
                if (!CheckDate(DatePickerDestroy12, "об уничтожении по"))
                    return;
                if (!CheckDate(DatePickerFilial1, "получения (филиал) с"))
                    return;
                if (!CheckDate(DatePickerFilial2, "получения (филиал) по"))
                    return;
                if (!CheckDate(DatePickerDestroy21, "уничтожения с"))
                    return;
                if (!CheckDate(DatePickerDestroy22, "уничтожения по"))
                    return;

                FIO_state = tbFio.Text;


                if (tbNumber.Text != "")
                    al.Add(String.Format("(pan like [%{0}%])", tbNumber.Text));
                if (tbFio.Text != "")
                    al.Add(String.Format("(fio like [%{0}%])", tbFio.Text));
                if (DatePickerProd1.DatePickerText.Trim() != "")
                    al.Add(String.Format("(dateProd>=[{0:" + ConfigurationSettings.AppSettings["DateFormat"] + "}])", DatePickerProd1.SelectedDate));
                if (DatePickerProd2.DatePickerText.Trim() != "")
                    al.Add(String.Format("(dateProd<=[{0:" + ConfigurationSettings.AppSettings["DateFormat"] + "}])", DatePickerProd2.SelectedDate));
                if (tbPassport.Text.Trim().Length > 0)
                    al.Add(String.Format("(passport like [%{0}%])", tbPassport.Text.Trim()));
                if (tbAccount.Text.Trim().Length > 0)
                    al.Add(String.Format("(account like [%{0}%])", tbAccount.Text.Trim()));
                if (tbInvoice.Text.Trim().Length > 0)
                    al.Add(String.Format("(invoice like [%{0}%])", tbInvoice.Text.Trim()));
                if (tbComment.Text != "")
                    al.Add(String.Format("(comment like [%{0}%])", tbComment.Text));

                if (DatePickerBegin1.DatePickerText.Trim() != "")
                    al.Add(String.Format("(dateStart>=[{0:" + ConfigurationSettings.AppSettings["DateFormat"] + "}])", DatePickerBegin1.SelectedDate));
                if (DatePickerBegin2.DatePickerText.Trim() != "")
                    al.Add(String.Format("(dateStart<=[{0:" + ConfigurationSettings.AppSettings["DateFormat"] + "}])", DatePickerBegin2.SelectedDate));
                if (DatePickerEnd1.DatePickerText.Trim() != "")
                    al.Add(String.Format("(dateEnd>=[{0:" + ConfigurationSettings.AppSettings["DateFormat"] + "}])", DatePickerEnd1.SelectedDate));
                if (DatePickerEnd2.DatePickerText.Trim() != "")
                    al.Add(String.Format("(dateEnd<=[{0:" + ConfigurationSettings.AppSettings["DateFormat"] + "}])", DatePickerEnd2.SelectedDate));
                if (DatePickerReceive1.DatePickerText.Trim() != "")
                    al.Add(String.Format("(dateReceipt>=[{0:" + ConfigurationSettings.AppSettings["DateFormat"] + "}])", DatePickerReceive1.SelectedDate));
                if (DatePickerReceive2.DatePickerText.Trim() != "")
                    al.Add(String.Format("(dateReceipt<=[{0:" + ConfigurationSettings.AppSettings["DateFormat"] + "}])", DatePickerReceive2.SelectedDate));
                if (DatePickerClient1.DatePickerText.Trim() != "")
                    al.Add(String.Format("(dateClient>=[{0:" + ConfigurationSettings.AppSettings["DateFormat"] + "}])", DatePickerClient1.SelectedDate));
                if (DatePickerClient2.DatePickerText.Trim() != "")
                    al.Add(String.Format("(dateClient<=[{0:" + ConfigurationSettings.AppSettings["DateFormat"] + "}])", DatePickerClient2.SelectedDate));
                if (DatePickerPered1.DatePickerText.Trim() != "")
                    al.Add(String.Format("(date_courier>=[{0:" + ConfigurationSettings.AppSettings["DateFormat"] + "}])", DatePickerPered1.SelectedDate));
                if (DatePickerPered2.DatePickerText.Trim() != "")
                    al.Add(String.Format("(date_courier<=[{0:" + ConfigurationSettings.AppSettings["DateFormat"] + "}])", DatePickerPered2.SelectedDate));
                if (DatePickerDestroy11.DatePickerText.Trim() != "")
                    al.Add(String.Format("(datesendterminate>=[{0:" + ConfigurationSettings.AppSettings["DateFormat"] + "}])", DatePickerDestroy11.SelectedDate));
                if (DatePickerDestroy12.DatePickerText.Trim() != "")
                    al.Add(String.Format("(datesendterminate<=[{0:" + ConfigurationSettings.AppSettings["DateFormat"] + "}])", DatePickerDestroy12.SelectedDate));
                if (DatePickerDestroy21.DatePickerText.Trim() != "")
                    al.Add(String.Format("(dateterminated>=[{0:" + ConfigurationSettings.AppSettings["DateFormat"] + "}])", DatePickerDestroy21.SelectedDate));
                if (DatePickerDestroy22.DatePickerText.Trim() != "")
                    al.Add(String.Format("(dateterminated<=[{0:" + ConfigurationSettings.AppSettings["DateFormat"] + "}])", DatePickerDestroy22.SelectedDate));
                if (DatePickerFilial1.DatePickerText.Trim() != "")
                    al.Add(String.Format("(dategetterminate>=[{0:" + ConfigurationSettings.AppSettings["DateFormat"] + "}])", DatePickerFilial1.SelectedDate));
                if (DatePickerFilial2.DatePickerText.Trim() != "")
                    al.Add(String.Format("(dategetterminate<=[{0:" + ConfigurationSettings.AppSettings["DateFormat"] + "}])", DatePickerFilial2.SelectedDate));


                string id_list = dListStatus.SelectedItem.Value;
                if (id_list != "-1")
                    al.Add(String.Format("(id_stat={0})", id_list));
                id_list = ddlProp.SelectedItem.Value;
                if (id_list != "-1")
                    al.Add(String.Format("(id_prop={0})", id_list));
                id_list = dListProd.SelectedItem.Value;
                if (id_list != "-1")
                    al.Add(String.Format("(id_prb={0})", id_list));

                id_list = dListOrgan.SelectedItem.Text; //.Value;
                //if (id_list != "-1")
                if (id_list != "Все")
                    al.Add(String.Format("(company=[{0}])", id_list));
                if (tbOrganization.Text.Trim().Length > 0)
                    al.Add($"(company like [%{tbOrganization.Text.Trim()}%])");
                id_list = (dListBranchCard.SelectedItem == null) ? "-1" : dListBranchCard.SelectedItem.Value;
                if (id_list != "-1")
                    al.Add(String.Format("(id_branchCard={0} or id_branchCard in (select id from Branchs where id_parent={0}))", id_list));
               
                id_list = dListBranchProd.SelectedItem.Value;
                if (id_list != "-1")
                    al.Add(String.Format("(id_branchInit={0} or id_branchInit in (select id from Branchs where id_parent={0}))", id_list));
                id_list = ddlClientWorker.SelectedItem.Value;
                if (id_list != "-1")
                    al.Add(String.Format("(clientWorker=[{0}])", id_list));

                id_list = dListBank.SelectedItem.Value;
                if (id_list != "-1")
                    al.Add(String.Format("(id_bank={0})", id_list));

                id_list = ddlWorker.SelectedItem.Value;
                if (id_list != "-1")
                    al.Add(String.Format("(courier_name like [{0}])", id_list));

                id_list = ddlCourier.SelectedItem.Value;
                if (id_list != "-1")
                    al.Add(String.Format("(id_courier={0})", id_list));

                if (getViewBranchCurrent() == true)
                {
                    id_list = dListBranchCurrent.SelectedItem.Value;
                    if (id_list != "-1")
                        al.Add(String.Format("(id_branchCurrent={0})", id_list));
                }
                
                if (branch_main_filial > 0 && branch_main_filial != current_branch_id)
                {
                    al.Add(String.Format("(id_branchCurrent={0})", current_branch_id));
                }

                id_list = DListSchoolProduct.SelectedItem.Value;
                if (id_list != "-1")
                    al.Add(String.Format("(id_prb={0})", id_list));
                id_list = ddlPins.SelectedItem.Value;
                if (id_list == "2")
                    al.Add($"(isPin=1)");
                if (id_list == "3")
                    al.Add($"(isPin=0)");


                if (al.Count > 0)
                {
                    string[] all = Array.CreateInstance(typeof(string), al.Count) as string[];
                    al.CopyTo(all, 0);
                    s = "where " + String.Join(" and ", all);
                }
                
                                
                Response.Write("<script language=javascript>window.returnValue='" + s + "'; window.close();</script>");
            }
        }
            
    }
}
