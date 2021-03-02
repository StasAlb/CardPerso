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
using System.Web.Configuration;
using System.Xml.Linq;
using System.Net;
using System.Net.Mail;

using System.IO;
using OstCard.Data;
using System.Data.SqlClient;
using CardPerso.Administration;
using OstCard.WebControls;
using System.Threading;
using System.Collections.Generic;
using System.Diagnostics;
using System.Web.Mvc;
using System.Web.Profile;
using System.Web.Services.Description;
using Ionic.Zip;
//using System.Data.OleDb;
using SocialExplorer.IO.FastDBF;

namespace CardPerso
{
    public partial class GenDoc : System.Web.UI.Page
    {


        public bool AllData = true;
        ExcelAp ep = null;
        ServiceClass sc = new ServiceClass();

        //string tableName = "eq_1_vneb20";
        string fileDBF = "eq_1_vneb20.dbf";
        string zipArch = "eq_1_vneb20.zip";


        int branch_main_filial = 0;
        int current_branch = 0;
        bool perso = false;
        int branchIdMain = 0; 

       

        private String oneDateMessage;
        public String OneDataMessage()
        {
            //if (isMemoricType())
            //{
                if (oneDateMessage == null) return "";
                else
                {
                    if (oneDateMessage.Length > 0) return "Операционный день:&nbsp; " + oneDateMessage;
                    else return oneDateMessage;
                }
            //}
            //else return "";
        }

        public String excludeMemoricProduct = "";

        public bool isTwoDay = false;
        
        public bool isMemoricType()
        {
            if (dListDoc.SelectedIndex >= 0 && (dListDoc.SelectedItem.Value == "42")) return true;
            return false;
        }
        
        protected void Page_Load(object sender, EventArgs e)
        {
            DialogUtils.SetScriptSubmit(this);

            HttpCookie cookie = Request.Cookies.Get("downLoadEnd");

            //if (!Page.IsCallback)
            //    ltCallback.Text =
            //        ClientScript.GetCallbackEventReference(this, "'bindgrid'", "EndGetData", "'asyncgrid'", false);


                if (!sc.UserAction(User.Identity.Name, Restrictions.Reports))
                //Response.Redirect("~\\Account\\Restricted.aspx", true);
                    Server.Transfer("~\\Account\\Restricted.aspx");
                
            perso = sc.UserAction(User.Identity.Name, Restrictions.Perso);
            current_branch = sc.BranchId(User.Identity.Name);


            branch_main_filial = BranchStore.getBranchMainFilial(current_branch, perso);
            branchIdMain = BranchStore.getBranchMainFilial(current_branch, false);

            
            OperationDay op = new OperationDay();
            op.read(current_branch);
            
            if (dListDoc.SelectedIndex >= 0 && (dListDoc.SelectedItem.Value == "42" || dListDoc.SelectedItem.Value == "142"))
            {
                lock (Database2.lockObjectDB)
                {
                    excludeMemoricProduct = getExcludeMemoricNameProducts();
                }
            }

            if (Page.IsPostBack==false)
            {
                userradio.Checked = false;
                branchradio.Checked = true;
            }
            


            if (Page.IsPostBack)
            {
                if (dListDoc.SelectedIndex >= 0 && dListDoc.SelectedItem.Value != "42")
                {
                    oneDateMessage = "";
                    isTwoDay = false;
                }
                else
                
                if(branch_main_filial>0)
                {
                    oneDateMessage = op.getMessage(DatePickerOne.SelectedDate);
                    isTwoDay = op.isShift;
                    allrangeradio.Text = "две смены";//op.getMessage(DatePickerOne.SelectedDate);
                    firstrangeradio.Text = op.getMessagePart(DatePickerOne.SelectedDate,true);
                    nextrangeradio.Text = op.getMessagePart(DatePickerOne.SelectedDate, false);
                }
                if (dListDoc.SelectedIndex >= 0 && dListDoc.SelectedItem.Value == "42")
                {
                    userradio.Text = "текущему пользователю";
                    branchradio.Text = "текущему подразделению";
                }
                if (dListDoc.SelectedIndex >= 0 && dListDoc.SelectedItem.Value == "142")
                {
                    userradio.Text = "текущему пользователю и подразделению";
                    branchradio.Text = "подразделениям";
                }
                return;
            }
          

            lock (Database2.lockObjectDB)
            {
                AllData = sc.UserAction(User.Identity.Name, Restrictions.AllData);
                dListDoc.Items.Clear();
                if (AllData && !sc.UserAction(User.Identity.Name, Restrictions.Transport))
                {
                    dListDoc.Items.Add(new ListItem("Сводный акт получения ценностей", "12"));
                    //dListDoc.Items.Add(new ListItem("Движение ценностей", "13"));   // старый вариант
                    dListDoc.Items.Add(new ListItem("Движение ценностей", "15")); // новый вариант
                    dListDoc.Items.Add(new ListItem("Хранилище: текущее состояние", "6")); // старый вариант
                    dListDoc.Items.Add(new ListItem("Хранилище: текущее состояние_1", "18")); // новый вариант
                    dListDoc.Items.Add(new ListItem("Хранилище: ценности с количеством ниже порогового значения", "7"));
                    dListDoc.Items.Add(new ListItem("Выпущенные карты: по продуктам", "3"));
                    dListDoc.Items.Add(new ListItem("Выпущенные карты: динамика", "4"));
                    dListDoc.Items.Add(new ListItem("Уничтоженные карты: по продуктам", "8"));
                    dListDoc.Items.Add(new ListItem("Уничтоженные карты: по подразделениям", "9"));
                    dListDoc.Items.Add(new ListItem("Бракованные карты: по продуктам-", "10"));
                    dListDoc.Items.Add(new ListItem("Закупочные договора", "11"));
                    dListDoc.Items.Add(new ListItem("Выданные ценности", "14"));
                    if (sc.UserAction(User.Identity.Name, Restrictions.ReportGO))
                        dListDoc.Items.Add(new ListItem("Выданные карты по подразделению", "16"));
                    dListDoc.Items.Add(new ListItem("Списание карт/пин конвертов с подотчета МОЛ", "43"));
                }
                else
                {
                    dListDoc.Items.Add(new ListItem("Выпущенные карты: по продуктам", "3"));
                    dListDoc.Items.Add(new ListItem("Выпущенные карты: динамика", "4"));
                    if (sc.UserAction(User.Identity.Name, Restrictions.ReportGO))
                        dListDoc.Items.Add(new ListItem("Выданные карты по подразделению", "16"));
                    if(branch_main_filial>0)
                        dListDoc.Items.Add(new ListItem("Передача между сменами", "41"));
                }
                if (sc.UserAction(User.Identity.Name, Restrictions.Filial))
                    if (branchIdMain == 0)
                        dListDoc.Items.Add(new ListItem("Выданные карты по филиалу", "17"));
                dListDoc.Items.Add(new ListItem("Филиал итог день", "19"));
                if (sc.UserAction(User.Identity.Name, Restrictions.Report4502))
                    dListDoc.Items.Add(new ListItem("Консолидированный отчет по Казанскому филиалу", "20"));
                if (branch_main_filial > 0)
                {
                    dListDoc.Items.Add(new ListItem("Мемориальные ордера", "42"));
                    dListDoc.Items.Add(new ListItem("Экспорт мемориальных ордеров", "142"));
                }
                if (sc.UserAction(User.Identity.Name, Restrictions.ReportKGRT))
                    dListDoc.Items.Add(new ListItem("Карта жителя", "44"));
                /// RESTRICTIONS
                dListDoc.Items.Add(new ListItem("Отчет выданных карт по книге 124", "45"));
                dListDoc.Items.Add(new ListItem("Отчет по движению карт по книге 124", "46"));
                dListDoc.Items.Add(new ListItem("Ведомость подотчет", "47"));
                if (sc.UserAction(User.Identity.Name, Restrictions.Filial))
                    dListDoc.Items.Add(new ListItem("Отчет приходов-расходов завкассы", "48"));
                if (sc.UserAction(User.Identity.Name, Restrictions.Filial))
                    dListDoc.Items.Add(new ListItem("Отчет приходов-расходов бригадира ГОЗ", "49"));

                if (sc.UserAction(User.Identity.Name, Restrictions.GetGoz) &&
                    sc.UserAction(User.Identity.Name, Restrictions.FromGoz))
                {
                    dListDoc.Items.Clear();
                    dListDoc.Items.Add(new ListItem("Отчет приходов-расходов бригадира ГОЗ", "49"));
                }


                DatePickerStart.FirstDayOfWeek = DatePicker.Day.Monday;
                DatePickerEnd.FirstDayOfWeek = DatePicker.Day.Monday;
                DatePickerOne.FirstDayOfWeek = DatePicker.Day.Monday;
                DatePickerStart.SelectedDate = DateTime.Now; //.AddYears(-1);
                DatePickerEnd.SelectedDate = DateTime.Now;
                DatePickerOne.SelectedDate = DateTime.Now;
                DatePickerStart.Culture = System.Globalization.CultureInfo.GetCultureInfo("ru-RU");
                DatePickerEnd.Culture = System.Globalization.CultureInfo.GetCultureInfo("ru-RU");
                DatePickerOne.Culture = System.Globalization.CultureInfo.GetCultureInfo("ru-RU");

                oneDateMessage = ""; // op.getMessage(DatePickerOne.SelectedDate);

                string criter = " (1=1) ";
                string temp = "Все продукты";
                if (sc.UserAction(User.Identity.Name, Restrictions.Transport))
                {
                    string[] strs = ConfigurationSettings.AppSettings["Transport"].Split(',');
                    for (int t = 0; t < strs.Length; t++)
                        strs[t] = String.Format("prefix_ow like '{0}'", strs[t]);
                    criter = String.Format("({0})", String.Join(" or ", strs));
                    temp = "Все продукты (транспортные)";
                }

                DataSet ds = new DataSet();
                Database2.ExecuteQuery($"select id_prb, prod_name, bank_name from V_ProductsBanks_T where id_type=1 and {criter} order by id_sort", ref ds, null);
                DataRow dr = ds.Tables[0].NewRow();
                dr["id_prb"] = 0;
                dr["prod_name"] = temp;
                dr["bank_name"] = "";
                ds.Tables[0].Rows.InsertAt(dr, 0);
                ddlProducts.DataSource = ds.Tables[0];
                ddlProducts.DataTextField = "prod_name";
                ddlProducts.DataValueField = "id_prb";
                ddlProducts.DataBind();
                ds.Clear();
                string str = "select id, ident_dep, department from Branchs where id=" + sc.BranchId(User.Identity.Name).ToString();
                if (sc.UserAction(User.Identity.Name, Restrictions.AllData))
                    str = "select id, ident_dep, department from Branchs order by department";
                Database2.ExecuteQuery(str, ref ds, null);
                ddlBranch.DataSource = ds.Tables[0];
                ddlBranch.DataTextField = "department";
                ddlBranch.DataValueField = "ident_dep";
                ddlBranch.DataBind();
                Database2.ExecuteQuery($"set NUMERIC_ROUNDABORT off set ANSI_NULLS, ANSI_PADDING, ANSI_WARNINGS, ARITHABORT, CONCAT_NULL_YIELDS_NULL, QUOTED_IDENTIFIER on select ap.id, ap.secondname + ' ' + ap.firstname + ' ' + case when ap.patronymic is null then '' else ' ' + ap.patronymic end as fio, ap.personnelnumber from AccountablePerson ap inner join aspnet_users u on ap.userId=u.id where dbo.UserMainBranch(u.UserName)={branchIdMain} order by secondname, firstname asc", ref ds, null);
                Database2.Log(sc.UserGuid(User.Identity.Name), $"report {ds.Tables?.Count}", null);
                ddlPersons.Items.Add(new ListItem("Не определено", "-1"));
                Database2.Log(sc.UserGuid(User.Identity.Name), $"0000", null);
                for (int i = 0; i < ds.Tables[0].Rows.Count; i++)
                {
                    Database2.Log(sc.UserGuid(User.Identity.Name), $"{ds.Tables[0].Rows[i]["fio"]}", null);
                    ddlPersons.Items.Add(new ListItem($"{ds.Tables[0].Rows[i]["fio"].ToString()} ({ds.Tables[0].Rows[i]["personnelnumber"].ToString()})", ds.Tables[0].Rows[i]["id"].ToString()));
                }
                dListDoc_SelectedIndexChanged(this, null);
            }
        }

        private void SetButton()
        {
            //lock (Database2.lockObjectDB)
            //{
            //    if (gvUsers.Rows.Count > 0)
            //    {
            //        bEdit.Visible = true;
            //        bDelete.Visible = true;
            //        bActivate.Visible = true;
            //        bActivate.Enabled = true;
            //        bRoles.Visible = true;
            //        object obj = null;
            //        string str = gvUsers.DataKeys[Convert.ToInt32(gvUsers.SelectedIndex)].Values["UserId"].ToString();
            //        Database2.ExecuteScalar(String.Format("select count(*) from LogAction where UserId='{0}'", str), ref obj, null);
            //        if ((int)obj > 0)
            //        {
            //            bDelete.ToolTip = "Сделать неактивным";
            //            bDelete.Attributes.Add("OnClick", String.Format("return confirm('Сделать неактивным пользователя {0}?');", gvUsers.DataKeys[Convert.ToInt32(gvUsers.SelectedIndex)].Values["UserName"].ToString()));
            //        }
            //        else
            //        {
            //            bDelete.ToolTip = "Удалить пользователя";
            //            bDelete.Attributes.Add("OnClick", String.Format("return confirm('Удалить пользователя {0}?');", gvUsers.DataKeys[Convert.ToInt32(gvUsers.SelectedIndex)].Values["UserName"].ToString()));
            //        }
            //        str = gvUsers.DataKeys[Convert.ToInt32(gvUsers.SelectedIndex)].Values["UserName"].ToString();
            //        if (!String.IsNullOrEmpty(str) && (Membership.FindUsersByName(str)[str].IsLockedOut || !Membership.FindUsersByName(str)[str].IsApproved))
            //            bActivate.Attributes.Add("OnClick", String.Format("return confirm('Активировать пользователя {0}?');", gvUsers.DataKeys[Convert.ToInt32(gvUsers.SelectedIndex)].Values["UserName"].ToString()));
            //        else
            //            bActivate.Enabled = false;
            //        bEdit.Attributes.Add("OnClick", String.Format("return edit_user('id={0}')", gvUsers.DataKeys[Convert.ToInt32(gvUsers.SelectedIndex)].Values["UserId"].ToString()));
            //        bRoles.Attributes.Add("OnClick", String.Format("return edit_role('id={0}')", gvUsers.DataKeys[Convert.ToInt32(gvUsers.SelectedIndex)].Values["UserId"].ToString()));
            //    }
            //    else
            //    {
            //        bEdit.Visible = false;
            //        bRoles.Visible = false;
            //        bDelete.Visible = false;
            //        bActivate.Visible = false;
            //    }
            //}
        }




        protected List<AccountBranch> getAccountBranch(String ident_dep)
        {
            DataSet ds = new DataSet();
            ds.Clear();
            lock (Database2.lockObjectDB)
            {
                Database2.ExecuteQuery(String.Format("select id from Branchs where ident_dep='{0}'", ident_dep), ref ds, null);
            }
            List<AccountBranch> ret = null;
            if (ds.Tables.Count > 0 && ds.Tables[0].Rows.Count > 0)
            {
                ret = getAccountBranch(Convert.ToInt32(ds.Tables[0].Rows[0]["id"]));
            }
            else ret = new List<AccountBranch>();
            ds = null;
            return ret;
        }

        protected List<AccountBranch> getAccountBranch(int branchId)
        {
            List<AccountBranch> listAB = new List<AccountBranch>();
            DataSet ds = new DataSet();
            ds.Clear();
            lock (Database2.lockObjectDB)
            {
                Database2.ExecuteQuery(String.Format("select * from BranchAccount where id_branch={0}", branchId), ref ds, null);
            }
            if (ds != null && ds.Tables.Count > 0)
            {
                for (int i = 0; i < ds.Tables[0].Rows.Count; i++)
                {
                    int pt = Convert.ToInt32(ds.Tables[0].Rows[i]["product_type"]);
                    if (pt != 0 && Enum.IsDefined(typeof(BaseProductType), pt))
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
            ds = null;
            return listAB;
        }
        protected void bExcel_Click(object sender, ImageClickEventArgs e)
        {
            lbInform.Text = "";
            lock (Database2.lockObjectDB)
            {
                
                string res = "";

                if (dListDoc.SelectedIndex < 0) return;
                try
                {
                    if (rbPeriods.SelectedValue == "2")
                        Convert.ToInt32(tbDays.Text);
                    lDaysWarning.Visible = false;
                }
                catch
                {
                    lDaysWarning.Visible = true;
                    return;
                }
                int id_type = Convert.ToInt32(dListDoc.SelectedItem.Value);
                if (id_type == 12 || id_type == 13 || id_type == 15 || id_type == 16 || id_type == 17 || id_type == 42 || id_type == 19 || (id_type == 142 && userradio.Checked == false) || id_type == 41 || id_type == 48 || id_type == 49)
                {
                    try
                    {
                        DateTime d1 = DateTime.ParseExact(DatePickerOne.DatePickerText, "dd.MM.yyyy", null);
                        //Convert.ToDateTime(DatePickerOne.DatePickerText);
                        lDatePickerOneError.Visible = false;
                        OperationDay op = new OperationDay();
                        op.read(current_branch);
                        oneDateMessage = op.getMessage(DatePickerOne.SelectedDate);
                        if ((id_type == 48 || id_type == 49)
                            && (DateTime.Now.Date - d1.Date).TotalDays > 14)

                        {
                            ClientScript.RegisterClientScriptBlock(GetType(), "err43", "<script type='text/javascript'>$(document).ready(function(){ ShowError('данный отчет недоступен для дат более чем на 2 недели назад');});</script>");
                            return;
                        }
                    }
                    catch
                    {
                        lDatePickerOneError.Visible = true;
                        return;
                    }
                }
                if (id_type == 3 || id_type == 4 || id_type == 5 || id_type == 8 || id_type == 9 || id_type == 10 || id_type == 11 || id_type == 43 || id_type == 44 || id_type == 47)
                {
                    try
                    {
                        //Convert.ToDateTime(DatePickerStart.DatePickerText);
                        //Convert.ToDateTime(DatePickerEnd.DatePickerText);
                        DateTime d1 = DateTime.ParseExact(DatePickerStart.DatePickerText, "dd.MM.yyyy", null);
                        DateTime d2 = DateTime.ParseExact(DatePickerEnd.DatePickerText, "dd.MM.yyyy", null);
                        if (d2 < d1)
                        {
                            ClientScript.RegisterClientScriptBlock(GetType(), "err43", "<script type='text/javascript'>$(document).ready(function(){ ShowError('дата окончания не может быть меньше дата начала');});</script>");
                            return;
                        }
                        //ведомость подотчет не разрешаем откатывать назад более чем на 2 недели, поскольку считается от сегодняшнего дня назад. Чтобы не нагружать систему
                        if (id_type == 47 && (DateTime.Now.Date - d1.Date).TotalDays > 14)
                        {
                            ClientScript.RegisterClientScriptBlock(GetType(), "err43", "<script type='text/javascript'>$(document).ready(function(){ ShowError('данный отчет недоступен для дат более чем на 2 недели назад');});</script>");
                            return;
                        }

                        lDatePickerTwoError.Visible = false;
                    }
                    catch (Exception ex)
                    {
                        lDatePickerTwoError.Visible = true;
                        return;
                    }
                }
                if (id_type == 47 && !(ddlPersons.SelectedIndex > 0))
                {
                    ClientScript.RegisterClientScriptBlock(GetType(), "err43", "<script type='text/javascript'>$(document).ready(function(){ ShowError('не выбрано подотчетное лицо');});</script>");
                    return;
                }
                string doc = "";
                string dt_doc = "";

                if (id_type == 1) doc = "Attachment1.xls";
                if (id_type == 2) doc = "Attachment7.xls";
                if (id_type == 3) doc = "Report1.xls";
                if (id_type == 4) doc = "Report1.xls";
                if (id_type == 6) doc = "Report2.xls";
                if (id_type == 7) doc = "Report2.xls";
                if (id_type == 8) doc = "Report1.xls";
                if (id_type == 9) doc = "Report1.xls";
                if (id_type == 10) doc = "Report1.xls";
                if (id_type == 11) doc = "Report1.xls";
                if (id_type == 12) doc = "Attachment9.xls";
                if (id_type == 13) doc = "Attachment10.xls";
                if (id_type == 14) doc = "Report1.xls";
                if (id_type == 15) doc = "Attachment10_1.xls";
                if (id_type == 16) doc = "Attachment15.xls";
                if (id_type == 17) doc = "Attachment15_1.xls";
                if (id_type == 18) doc = "Attachment23.xls";
                if (id_type == 19) doc = "Attachment24.xls";
                if (id_type == 20) doc = "Attachment1_4502.xls";
                if (id_type == 41) doc = "Attachment41.xls";
                if (id_type == 42) doc = "AttMemoric.xls";
                if (id_type == 43) doc = "Attachment10_2.xls";
                if (id_type == 44) doc = "Attachment44.xls";
                if (id_type == 45) doc = "Attachment124_r.xls";
                if (id_type == 46) doc = "Attachment_OP_KP.xls";
                if (id_type == 47) doc = "Attachment_AccVed.xls";
                if (id_type == 48) doc = "Attachment_InOut.xls";
                if (id_type == 49) doc = "Attachment_InOut.xls";


                System.Globalization.CultureInfo oldCI = System.Threading.Thread.CurrentThread.CurrentCulture;
                System.Threading.Thread.CurrentThread.CurrentCulture = new System.Globalization.CultureInfo("en-US");
                ep = new ExcelAp();
                
                string criter = "(1=1)";
                string branchStr = "";
                string branchName = "";
                int i = 0;
                bool make = true; // для движения ценностей - отображение архивного
                string fname = ""; // для движения ценностей - отображение архивного
                if (sc.UserAction(User.Identity.Name, Restrictions.Transport))
                {
                    string[] strs = ConfigurationSettings.AppSettings["Transport"].Split(',');
                    for (int t = 0; t < strs.Length; t++)
                        strs[t] = String.Format("prefix_ow like '{0}'", strs[t]);
                    criter = String.Format("({0})", String.Join(" or ", strs));
                }
                
                if (ep.RunApp(ConfigurationSettings.AppSettings["DocPath"] + doc))
                {
                    try
                    {
                        ep.SaveAsDoc(ConfigurationSettings.AppSettings["DocPath"] + "Temp/" + doc, false);
                        ep.SetWorkSheet(1);
                        DataSet ds = new DataSet();

                        #region записка на выпуск карт
                        if (id_type == 1)
                        {
                            dt_doc = String.Format("{0:dd.MM.yyyy}", DateTime.Now);
                            ep.SetText_Name("dt_doc", dt_doc);

                            ds.Clear();
                            res = Database2.ExecuteQuery("select prefix_file, count(id) as cnt from V_Cards where (id_stat = 1) group by prefix_file", ref ds, null);
                            for (i = 0; i < ds.Tables[0].Rows.Count; i++)
                                ep.SetText_Name(ds.Tables[0].Rows[i]["prefix_file"].ToString(), ds.Tables[0].Rows[i]["cnt"].ToString());
                        }

                        #endregion
                        #region сорт лист
                        if (id_type == 2)
                        {
                            dt_doc = String.Format("{0:dd.MM.yyyy HH:mm}", DateTime.Now);
                            ep.SetText(6, 5, dt_doc);

                            string dep_c = "";
                            string dep = "";
                            int cnt_sh = 0;
                            int cnt_row = 0;

                            ds.Clear();
                            res = Database2.ExecuteQuery("select DepBranchCard, prod_name, count(prod_name) AS cnt from V_Cards where ((id_stat = 2) and (DepBranchCard is not null)) group by DepBranchCard, prod_name order by DepBranchCard, prod_name", ref ds, null);
                            for (i = 0; i < ds.Tables[0].Rows.Count; i++)
                            {
                                dep = ds.Tables[0].Rows[i]["DepBranchCard"].ToString();

                                if (dep_c != dep)
                                {
                                    if (i != 0)
                                    {
                                        ep.SetRangeBorders(8 + i + cnt_sh - cnt_row, 1, 8 + i + cnt_sh - 1, 2);
                                        ep.SetRangeAutoFit(8 + i + cnt_sh - cnt_row, 1, 8 + i + cnt_sh - 1, 2);
                                        cnt_sh = cnt_sh + 2;
                                    }
                                    ep.SetText(8 + i + cnt_sh, 1, dep);
                                    ep.SetRangeBold(8 + i + cnt_sh, 1, 8 + i + cnt_sh, 2);
                                    cnt_sh++;
                                    cnt_row = 0;

                                }

                                ep.SetText(8 + i + cnt_sh, 1, ds.Tables[0].Rows[i]["prod_name"].ToString());
                                ep.SetText(8 + i + cnt_sh, 2, ds.Tables[0].Rows[i]["cnt"].ToString());

                                cnt_row++;
                                dep_c = ds.Tables[0].Rows[i]["DepBranchCard"].ToString();

                                ep.SetRangeBorders(8 + i + cnt_sh - cnt_row + 1, 1, 8 + i + cnt_sh, 2);
                                ep.SetRangeAutoFit(8 + i + cnt_sh - cnt_row + 1, 1, 8 + i + cnt_sh, 2);
                            }
                        }
                        #endregion
                        #region Выпущенные карты по продуктам
                        if (id_type == 3)
                        {
                            object obj = null;
                            SqlCommand comm = new SqlCommand();
                            comm.Parameters.Add("dateStart", SqlDbType.DateTime).Value = DatePickerStart.SelectedDate;
                            comm.Parameters.Add("dateEnd", SqlDbType.DateTime).Value = DatePickerEnd.SelectedDate;
                            if (!sc.UserAction(User.Identity.Name, Restrictions.Transport))
                            {
                                comm.CommandType = CommandType.StoredProcedure;
                                comm.CommandText = "Rep_GivenCards";
                                comm.Parameters.Add("branchCard", SqlDbType.Int);
                                if (!sc.UserAction(User.Identity.Name, Restrictions.AllData))
                                {
                                    SqlCommand comm1 = new SqlCommand();
                                    comm.Parameters["branchCard"].Value = sc.BranchId(User.Identity.Name);
                                    comm1.CommandText = "select department from Branchs where id=" + sc.BranchId(User.Identity.Name).ToString();
                                    string sss = Database2.ExecuteScalar(comm1, ref obj, null);
                                    branchName = " в " + obj.ToString();
                                }
                                else
                                    comm.Parameters["branchCard"].Value = -1;
                            }
                            else
                            {
                                //транспортные карты
                                res = comm.CommandText = String.Format("SELECT dbo.Products.name, dbo.Products_Banks.bin,(SELECT COUNT(*) AS Expr1 FROM dbo.V_Cards WHERE (id_prb = dbo.Products_Banks.id and {0}) and DateProd >= @dateStart and DateProd <= @dateEnd) AS Cnt FROM dbo.Products RIGHT OUTER JOIN dbo.Products_Banks ON dbo.Products.id = dbo.Products_Banks.id_prod WHERE (dbo.Products.id_type = 1)", criter);
                                branchName = "(транспортные карты)";
                            }
                            Database2.ExecuteCommand(comm, ref ds, null);
                            int t = 0;
                            for (i = 0; i < ds.Tables[0].Rows.Count; i++)
                            {
                                if (ds.Tables[0].Rows[i][2] != DBNull.Value && Convert.ToInt32(ds.Tables[0].Rows[i][2]) > 0)
                                {
                                    ep.SetText(t + 5, 2, ds.Tables[0].Rows[i][0].ToString());
                                    ep.SetText(t + 5, 3, ds.Tables[0].Rows[i][2].ToString());
                                    t++;
                                }
                            }
                            ep.SetChartData(3, 2, 5, 5 + t - 1, Microsoft.Office.Interop.Excel.XlChartType.xlPie, true);
                            ep.SetText_Name("title", "Выпущенные карты по продуктам " + branchName);
                            ep.SetText_Name("date_string", String.Format("за период с {0} по {1}", DatePickerStart.SelectedDate.ToString("dd.MM.yyyy"), DatePickerEnd.SelectedDate.ToString("dd.MM.yyyy")));
                            ep.SetText_Name("sost_string", String.Format("подготовлено {0}", DateTime.Now.ToString("dd.MM.yyyy")));
                        }
                        #endregion
                        #region Динамика выпуска карт
                        // Параметры DateStart, DateEnd, BranchId, Idprb
                        if (id_type == 4)
                        {
                            object obj = null;
                            SqlCommand comm = new SqlCommand();
                            if (!sc.UserAction(User.Identity.Name, Restrictions.AllData))
                            {
                                branchStr = " and id_branchCard=" + sc.BranchId(User.Identity.Name).ToString();
                                comm.CommandText = "select department from Branchs where id=" + sc.BranchId(User.Identity.Name).ToString();
                                string sss = Database2.ExecuteScalar(comm, ref obj, null);
                                branchName = obj.ToString();
                            }
                            DateTime dt = DatePickerStart.SelectedDate;

                            if (sc.UserAction(User.Identity.Name, Restrictions.Transport))
                                comm.CommandText = "select count(*) from V_Cards where dateProd >= @dateStart and dateProd < @dateEnd and " + criter;
                            else
                                comm.CommandText = "select count(*) from Cards where dateProd >= @dateStart and dateProd < @dateEnd" + branchStr;
                            comm.Parameters.Add("@dateStart", SqlDbType.DateTime);
                            comm.Parameters.Add("@dateEnd", SqlDbType.DateTime);
                            if (Convert.ToInt32(ddlProducts.SelectedValue) > 0)
                            {
                                comm.CommandText += " and id_prb = @idprb";
                                comm.Parameters.Add("@idprb", SqlDbType.Int).Value = Convert.ToInt32(ddlProducts.SelectedValue);
                            }
                            else
                                comm.CommandText += " and id_prb > 0";
                            i = 0;
                            int days = 0;
                            if (rbPeriods.SelectedValue == "2")
                                days = Convert.ToInt32(tbDays.Text);

                            do
                            {
                                comm.Parameters["@dateStart"].Value = dt;
                                if (rbPeriods.SelectedValue == "2")
                                {
                                    if (days > 1)
                                        ep.SetText(i + 5, 2, String.Format("{0:dd.MM} - {1:dd.MM.yyyy}", dt, dt.AddDays(days - 1)));
                                    else
                                        ep.SetText(i + 5, 2, String.Format("{0:dd.MM.yyyy}", dt));
                                    dt = dt.AddDays(days);

                                }
                                else
                                {
                                    ep.SetText(i + 5, 2, String.Format("{0:dd.MM} - {1:dd.MM.yyyy}", dt, dt.AddMonths(1).AddDays(-1)));
                                    dt = dt.AddMonths(1);
                                }
                                comm.Parameters["@dateEnd"].Value = dt;
                                Database2.ExecuteScalar(comm, ref obj, null);
                                ep.SetText(i + 5, 3, Convert.ToString(obj));
                                i++;
                            }
                            while (dt <= DatePickerEnd.SelectedDate);
                            ep.SetChartData(3, 2, 5, 5 + i - 1, Microsoft.Office.Interop.Excel.XlChartType.xlLineMarkers, false);
                            ep.SetText_Name("title", String.Format("Динамика выпуска карт ({0}) {1}", ddlProducts.SelectedItem.Text, branchName));
                            ep.SetText_Name("date_string", String.Format("за период с {0} по {1}", DatePickerStart.SelectedDate.ToString("dd.MM.yyyy"), DatePickerEnd.SelectedDate.ToString("dd.MM.yyyy")));
                            ep.SetText_Name("sost_string", String.Format("подготовлено {0}", DateTime.Now.ToString("dd.MM.yyyy")));
                        }
                        #endregion
                        #region Состояние хранилища
                        if (id_type == 6)
                        {
                            Database2.ExecuteQuery("select name, bank_name, cnt_new, cnt_perso, cnt_brak, bin from V_Storage where id_type=1 order by id_sort", ref ds, null);
                            int rw = 4;
                            ep.SetText(rw, 2, "Карты");
                            ep.SetText(rw, 3, "Банк");
                            ep.SetText(rw, 4, "bin");
                            ep.SetText(rw, 5, "Новые");
                            ep.SetText(rw, 6, "Персо");
                            ep.SetText(rw, 7, "Брак");
                            ep.SetRangeBold(rw, 2, rw, 8);
                            rw++;
                            for (i = 0; i < ds.Tables[0].Rows.Count; i++)
                            {
                                ep.SetText(rw, 2, ds.Tables[0].Rows[i]["name"].ToString());
                                ep.SetText(rw, 3, ds.Tables[0].Rows[i]["bank_name"].ToString());
                                ep.SetText(rw, 4, ds.Tables[0].Rows[i]["bin"].ToString());
                                ep.SetText(rw, 5, ds.Tables[0].Rows[i]["cnt_new"].ToString());
                                ep.SetText(rw, 6, ds.Tables[0].Rows[i]["cnt_perso"].ToString());
                                ep.SetText(rw, 7, ds.Tables[0].Rows[i]["cnt_brak"].ToString());
                                rw++;
                            }
                            rw++;
                            ds.Clear();
                            Database2.ExecuteQuery("select name, bank_name, cnt_new, cnt_perso, cnt_brak from V_Storage where id_type=2 order by id_sort", ref ds, null);
                            ep.SetText(rw, 2, "Пин-конверты");
                            ep.SetText(rw, 3, "Банк");
                            ep.SetText(rw, 5, "Новые");
                            ep.SetText(rw, 6, "Персо");
                            ep.SetText(rw, 7, "Брак");
                            ep.SetRangeBold(rw, 2, rw, 8);
                            rw++;
                            for (i = 0; i < ds.Tables[0].Rows.Count; i++)
                            {
                                //ep.SetText(rw, 2, ds.Tables[0].Rows[i]["name"].ToString());
                                ep.SetText(rw, 3, ds.Tables[0].Rows[i]["bank_name"].ToString());
                                ep.SetText(rw, 5, ds.Tables[0].Rows[i]["cnt_new"].ToString());
                                ep.SetText(rw, 6, ds.Tables[0].Rows[i]["cnt_perso"].ToString());
                                ep.SetText(rw, 7, ds.Tables[0].Rows[i]["cnt_brak"].ToString());
                                rw++;
                            }
                            rw++;
                            ds.Clear();
                            Database2.ExecuteQuery("select name, bank_name, cnt_new, cnt_perso from V_Storage where id_type=3 order by id_sort", ref ds, null);
                            ep.SetText(rw, 2, "Доп. продукция");
                            ep.SetText(rw, 3, "Банк");
                            ep.SetText(rw, 5, "Новые");
                            ep.SetText(rw, 6, "Персо");
                            ep.SetRangeBold(rw, 2, rw, 8);
                            rw++;
                            for (i = 0; i < ds.Tables[0].Rows.Count; i++)
                            {
                                ep.SetText(rw, 2, ds.Tables[0].Rows[i]["name"].ToString());
                                ep.SetText(rw, 3, ds.Tables[0].Rows[i]["bank_name"].ToString());
                                ep.SetText(rw, 5, ds.Tables[0].Rows[i]["cnt_new"].ToString());
                                ep.SetText(rw, 6, ds.Tables[0].Rows[i]["cnt_perso"].ToString());
                                rw++;
                            }
                            rw++;
                            ep.SetText_Name("title", "Состояние хранилища");
                            ep.SetText_Name("sost_string", String.Format("подготовлено {0}", DateTime.Now.ToString("dd.MM.yyyy")));
                        }
                        #endregion
                        #region Состояние хранилища - новый вариант
                        if (id_type == 18)
                        {
                            ep.SetText_Name("date_doc", DateTime.Now.ToString("dd.MM.yyyy"));
                            ds.Clear();
                            Database2.ExecuteQuery("select name, cnt_new, cnt_perso, id_type from V_Storage where id_type=1 order by id_sort", ref ds, null);
                            int[] visa = new int[2];
                            int[] mc = new int[2];
                            int[] srv = new int[2];
                            int[] mir = new int[2];
                            visa[0] = 0; visa[1] = 0;
                            mc[0] = 0; mc[1] = 0;
                            srv[0] = 0; srv[1] = 0;
                            mir[0] = 0; mir[1] = 0;
                            int col = 16;
                            ep.SetText(col++, 1, "Персонализированные карты");
                            col++;
                            int frs = col;
                            ep.SetText(col, 1, "Тип карт");
                            ep.SetText(col++, 2, "Количество");
                            for (i = 0; i < ds.Tables[0].Rows.Count; i++)
                            {
                                int cnt = Convert.ToInt32(ds.Tables[0].Rows[i]["cnt_perso"]);
                                if (cnt > 0)
                                {
                                    string name = ds.Tables[0].Rows[i]["name"].ToString();
                                    ep.SetText(col, 1, name);
                                    ep.SetText(col, 2, ds.Tables[0].Rows[i]["cnt_perso"].ToString());
                                    name = name.ToLower();
                                    /*
                                    if (name.StartsWith("mc") || name.StartsWith("master") || name.StartsWith("mifare standart"))
                                        mc[0] += cnt;
                                    if (name.StartsWith("visa"))
                                        visa[0] += cnt;
                                    if (name.StartsWith("белый") || name.StartsWith("карты типа восток"))
                                        srv[0] += cnt;
                                    */
                                    BaseProductType tp = BranchStore.codeFromTypeAndProdName(1, name);
                                    switch (tp)
                                    {
                                        case BaseProductType.MasterCard: mc[0] += cnt; break;
                                        case BaseProductType.VisaCard: visa[0] += cnt; break;
                                        case BaseProductType.ServiceCard: srv[0] += cnt; break;
                                        case BaseProductType.MirCard: mir[0] += cnt; break;

                                    }
                                    col++;
                                }
                            }
                            ep.SetRangeBorders(frs, 1, col - 1, 2);
                            col++;
                            ep.SetText(col++, 1, "Заготовки пластиковых карт");
                            col++;
                            frs = col;
                            ep.SetText(col, 1, "Тип карт");
                            ep.SetText(col++, 2, "Количество");
                            for (i = 0; i < ds.Tables[0].Rows.Count; i++)
                            {
                                int cnt = Convert.ToInt32(ds.Tables[0].Rows[i]["cnt_new"]);
                                if (cnt > 0)
                                {
                                    string name = ds.Tables[0].Rows[i]["name"].ToString();
                                    ep.SetText(col, 1, name);
                                    ep.SetText(col, 2, ds.Tables[0].Rows[i]["cnt_new"].ToString());
                                    name = name.ToLower();
                                    /*
                                    if (name.StartsWith("mc") || name.StartsWith("master") || name.StartsWith("mifare standart"))
                                        mc[1] += cnt;
                                    if (name.StartsWith("visa"))
                                        visa[1] += cnt;
                                    if (name.StartsWith("белый") || name.StartsWith("карты типа восток"))
                                        srv[1] += cnt;
                                    */
                                    BaseProductType tp = BranchStore.codeFromTypeAndProdName(1, name);
                                    switch (tp)
                                    {
                                        case BaseProductType.MasterCard: mc[1] += cnt; break;
                                        case BaseProductType.VisaCard: visa[1] += cnt; break;
                                        case BaseProductType.ServiceCard: srv[1] += cnt; break;
                                        case BaseProductType.MirCard: mir[1] += cnt; break;
                                    }
                                    col++;
                                }
                            }
                            ep.SetRangeBorders(frs, 1, col - 1, 2);
                            col++;
                            frs = col;
                            ep.SetText(col, 1, "Виды карт");
                            ep.SetText(col, 2, "Заготовки");
                            ep.SetText(col++, 3, "Персонализированные");
                            ep.SetText(col, 1, "MasterCard");
                            ep.SetText(col, 2, mc[1].ToString());
                            ep.SetText(col++, 3, mc[0].ToString());
                            ep.SetText(col, 1, "Visa");
                            ep.SetText(col, 2, visa[1].ToString());
                            ep.SetText(col++, 3, visa[0].ToString());
                            ep.SetText(col, 1, "Сервисные карты");
                            ep.SetText(col, 2, srv[1].ToString());
                            ep.SetText(col++, 3, srv[0].ToString());
                            ep.SetText(col, 1, "Карты МИР");
                            ep.SetText(col, 2, mir[1].ToString());
                            ep.SetText(col++, 3, mir[0].ToString());
                            ds.Clear();
                            Database2.ExecuteQuery("select name, cnt_new, cnt_perso, id_type from V_Storage where id_type=2 order by id_sort", ref ds, null);
                            if (ds.Tables.Count > 0 && ds.Tables[0].Rows.Count > 0)
                            {
                                ep.SetText(col, 1, "Пин-конверты");
                                ep.SetText(col, 3, Convert.ToInt32(ds.Tables[0].Rows[0]["cnt_perso"]).ToString());
                                ep.SetText(col, 2, Convert.ToInt32(ds.Tables[0].Rows[0]["cnt_new"]).ToString());
                                col++;
                            }
                            ep.SetRangeBorders(frs, 1, col - 1, 3);

                            ep.SetRangeBold(16, 1, col, 3);

                        }
                        #endregion
                        #region Ценности ниже порогового значения
                        if (id_type == 7)
                        {
                            Database2.ExecuteQuery("select name, bank_name, cnt_new, bin, min_cnt, id_type from V_Storage where cnt_new < min_cnt order by id_type, id_sort", ref ds, null);
                            int rw = 3;
                            int tp = 0;
                            rw++;
                            for (i = 0; i < ds.Tables[0].Rows.Count; i++)
                            {
                                if (Convert.ToInt32(ds.Tables[0].Rows[i]["id_type"]) != tp)
                                {
                                    rw++;
                                    tp = Convert.ToInt32(ds.Tables[0].Rows[i]["id_type"]);
                                    if (tp == 1)
                                    {
                                        ep.SetText(rw, 2, "Карты");
                                        ep.SetText(rw, 4, "bin");
                                    }
                                    if (tp == 2)
                                        ep.SetText(rw, 2, "Пин-конверты");
                                    if (tp == 3)
                                        ep.SetText(rw, 2, "Доп. продукция");
                                    ep.SetText(rw, 3, "Банк");
                                    ep.SetText(rw, 5, "Порог");
                                    ep.SetText(rw, 6, "Новые");
                                    ep.SetRangeBold(rw, 2, rw, 6);
                                    rw++;
                                }
                                ep.SetText(rw, 2, ds.Tables[0].Rows[i]["name"].ToString());
                                ep.SetText(rw, 3, ds.Tables[0].Rows[i]["bank_name"].ToString());
                                ep.SetText(rw, 4, ds.Tables[0].Rows[i]["bin"].ToString());
                                ep.SetText(rw, 5, ds.Tables[0].Rows[i]["min_cnt"].ToString());
                                ep.SetText(rw, 6, ds.Tables[0].Rows[i]["cnt_new"].ToString());
                                rw++;
                            }
                            ep.SetText_Name("title", "Ценности с количеством ниже порогового значения");
                            ep.SetText_Name("sost_string", String.Format("подготовлено {0}", DateTime.Now.ToString("dd.MM.yyyy")));
                        }
                        #endregion
                        #region Уничтоженные карты по продуктам
                        if (id_type == 8)
                        {
                            SqlCommand comm = new SqlCommand();
                            comm.CommandType = CommandType.StoredProcedure;
                            comm.CommandText = "Rep_DestroyedCards";
                            comm.Parameters.Add("dateStart", SqlDbType.DateTime).Value = DatePickerStart.SelectedDate;
                            comm.Parameters.Add("dateEnd", SqlDbType.DateTime).Value = DatePickerEnd.SelectedDate;
                            Database2.ExecuteCommand(comm, ref ds, null);
                            for (i = 0; i < ds.Tables[0].Rows.Count; i++)
                            {
                                ep.SetText(i + 5, 2, ds.Tables[0].Rows[i][0].ToString());
                                ep.SetText(i + 5, 3, ds.Tables[0].Rows[i][2].ToString());
                            }
                            ep.SetChartData(3, 2, 5, 5 + ds.Tables[0].Rows.Count - 1, Microsoft.Office.Interop.Excel.XlChartType.xlPie, true);
                            ep.SetText_Name("title", "Уничтоженные карты по продуктам");
                            ep.SetText_Name("date_string", String.Format("за период с {0} по {1}", DatePickerStart.SelectedDate.ToString("dd.MM.yyyy"), DatePickerEnd.SelectedDate.ToString("dd.MM.yyyy")));
                            ep.SetText_Name("sost_string", String.Format("подготовлено {0}", DateTime.Now.ToString("dd.MM.yyyy")));
                        }
                        #endregion
                        #region Уничтоженные карты по подразделениям
                        if (id_type == 9)
                        {
                            SqlCommand comm = new SqlCommand();
                            comm.CommandText = "Rep_DestroyedCardsB";
                            comm.CommandType = CommandType.StoredProcedure;
                            comm.Parameters.Add("dateStart", SqlDbType.DateTime).Value = DatePickerStart.SelectedDate;
                            comm.Parameters.Add("dateEnd", SqlDbType.DateTime).Value = DatePickerEnd.SelectedDate;
                            Database2.ExecuteCommand(comm, ref ds, null);
                            int t = 0;
                            for (i = 0; i < ds.Tables[0].Rows.Count; i++)
                            {
                                if (ds.Tables[0].Rows[i][2] != DBNull.Value && Convert.ToInt32(ds.Tables[0].Rows[i][2]) > 0)
                                {
                                    ep.SetText(t + 5, 2, ds.Tables[0].Rows[i][1].ToString());
                                    ep.SetText(t + 5, 3, ds.Tables[0].Rows[i][2].ToString());
                                    t++;
                                }
                            }
                            ep.SetChartData(3, 2, 5, 5 + t - 1, Microsoft.Office.Interop.Excel.XlChartType.xlColumnStacked, true);
                            ep.SetText_Name("title", "Уничтоженные карты по подразделениям");
                            ep.SetText_Name("date_string", String.Format("за период с {0} по {1}", DatePickerStart.SelectedDate.ToString("dd.MM.yyyy"), DatePickerEnd.SelectedDate.ToString("dd.MM.yyyy")));
                            ep.SetText_Name("sost_string", String.Format("подготовлено {0}", DateTime.Now.ToString("dd.MM.yyyy")));
                        }
                        #endregion
                        #region Бракованные карты по продуктам
                        if (id_type == 10)
                        {
                            SqlCommand comm = new SqlCommand();
                            comm.CommandText = "select prod_name, bank_name, cnt_brak from V_Rep_Moving where priz_gen=1 and type = 8 and cnt_brak > 0 and date_doc >= @dateStart and date_doc <= @dateEnd order by id_type, id_sort";
                            comm.Parameters.Add("dateStart", SqlDbType.DateTime).Value = DatePickerStart.SelectedDate;
                            comm.Parameters.Add("dateEnd", SqlDbType.DateTime).Value = DatePickerEnd.SelectedDate;
                            Database2.ExecuteCommand(comm, ref ds, null);
                            for (i = 0; i < ds.Tables[0].Rows.Count; i++)
                            {
                                ep.SetText(i + 5, 2, String.Format("{0} ({1})", ds.Tables[0].Rows[i]["prod_name"], ds.Tables[0].Rows[i]["bank_name"]));
                                ep.SetText(i + 5, 3, ds.Tables[0].Rows[i]["cnt_brak"].ToString());
                            }
                            ep.SetChartData(3, 2, 5, 5 + ds.Tables[0].Rows.Count - 1, Microsoft.Office.Interop.Excel.XlChartType.xlBarClustered, true);
                            ep.SetText_Name("title", "Брак по продуктам");
                            ep.SetText_Name("date_string", String.Format("за период с {0} по {1}", DatePickerStart.SelectedDate.ToString("dd.MM.yyyy"), DatePickerEnd.SelectedDate.ToString("dd.MM.yyyy")));
                            ep.SetText_Name("sost_string", String.Format("подготовлено {0}", DateTime.Now.ToString("dd.MM.yyyy")));
                        }
                        #endregion
                        #region Закупочные договора
                        if (id_type == 11)
                        {
                            SqlCommand comm = new SqlCommand();
                            comm.CommandText = "select date_dog, supplier, number_dog, summa from V_Rep_Dog where date_dog >= @dateStart and date_dog <= @dateEnd order by date_dog";
                            comm.Parameters.Add("dateStart", SqlDbType.DateTime).Value = DatePickerStart.SelectedDate;
                            comm.Parameters.Add("dateEnd", SqlDbType.DateTime).Value = DatePickerEnd.SelectedDate;
                            Database2.ExecuteCommand(comm, ref ds, null);
                            ep.SetText(5, 2, "Поставщик");
                            ep.SetText(5, 3, "Дата");
                            ep.SetText(5, 4, "Номер");
                            ep.SetText(5, 5, "Сумма");
                            ep.SetRangeBold(5, 2, 2, 5);
                            for (i = 0; i < ds.Tables[0].Rows.Count; i++)
                            {
                                ep.SetText(i + 6, 2, ds.Tables[0].Rows[i]["supplier"].ToString());
                                ep.SetText(i + 6, 3, Convert.ToDateTime(ds.Tables[0].Rows[i]["date_dog"]).ToString("dd.MM.yyyy"));
                                ep.SetFormat(i + 6, 4, "@");
                                ep.SetText(i + 6, 4, ds.Tables[0].Rows[i]["number_dog"].ToString());
                                if (ds.Tables[0].Rows[i]["summa"] != DBNull.Value)
                                    ep.SetText(i + 6, 5, Convert.ToDouble(ds.Tables[0].Rows[i]["summa"]).ToString("0.00") + "p.");
                            }
                            ep.SetText_Name("title", "Закупочные договора");
                            ep.SetText_Name("date_string", String.Format("за период с {0} по {1}", DatePickerStart.SelectedDate.ToString("dd.MM.yyyy"), DatePickerEnd.SelectedDate.ToString("dd.MM.yyyy")));
                            ep.SetText_Name("sost_string", String.Format("подготовлено {0}", DateTime.Now.ToString("dd.MM.yyyy")));
                        }
                        #endregion
                        #region Сводный акт
                        if (id_type == 12)
                        {
                            int col = 0; int col_t = 0; int row_t = 0; int row = 0;
                            string dop_prod = "";
                            SqlCommand comm = new SqlCommand();
                            DataSet ds1 = new DataSet();
                            int cnt = 0;
                            dt_doc = String.Format("{0:dd.MM.yyyy}", DatePickerOne.SelectedDate);
                            ep.SetText_Name("dt_doc", dt_doc);
                            ds.Clear();
                            res = Database2.ExecuteQuery(String.Format("select distinct id_branch,branch from V_StorageDocs where type=5 and date_doc='{0:" + ConfigurationSettings.AppSettings["DateFormat"] + "}'", DatePickerOne.SelectedDate), ref ds, null);
                            if (ep.GetNameRC("branch", ref row, ref col))
                            {
                                for (i = 0; i < ds.Tables[0].Rows.Count; i++)
                                {
                                    ep.SetText(row + i + 1, col, ds.Tables[0].Rows[i]["branch"].ToString());
                                    //передача
                                    ds1.Clear();
                                    comm.Parameters.Clear();
                                    comm.CommandType = CommandType.StoredProcedure;
                                    comm.CommandText = "Rep_SummaryActSend";
                                    comm.Parameters.Add("id_type_s", SqlDbType.Int).Value = (int)TypeDoc.SendToFilial;
                                    comm.Parameters.Add("id_branch", SqlDbType.Int).Value = Convert.ToInt32(ds.Tables[0].Rows[i]["id_branch"]);
                                    comm.Parameters.Add("dateDoc", SqlDbType.DateTime).Value = DatePickerOne.SelectedDate.Date;
                                    Database2.ExecuteCommand(comm, ref ds1, null);

                                    dop_prod = "";
                                    for (int j = 0; j < ds1.Tables[0].Rows.Count; j++)
                                    {
                                        if (Convert.ToInt32(ds1.Tables[0].Rows[j]["id_type"]) == 1)
                                            if (ep.GetNameRC(ds1.Tables[0].Rows[j]["prefix_file"].ToString(), ref row_t, ref col_t))
                                            {
                                                ep.SetFormat(row_t + i + 1, col_t, "@");
                                                ep.SetText(row_t + i + 1, col_t, ds1.Tables[0].Rows[j]["cnt_perso"].ToString());
                                            }

                                        if (Convert.ToInt32(ds1.Tables[0].Rows[j]["id_type"]) == 2)
                                            if (ep.GetNameRC("pin", ref row_t, ref col_t))
                                            {
                                                ep.SetFormat(row_t + i + 1, col_t, "@");
                                                ep.SetText(row_t + i + 1, col_t, ds1.Tables[0].Rows[j]["cnt_perso"].ToString());
                                            }

                                        if (Convert.ToInt32(ds1.Tables[0].Rows[j]["id_type"]) == 3)
                                            dop_prod = dop_prod + ds1.Tables[0].Rows[j]["name"].ToString() + ": " + ds1.Tables[0].Rows[j]["cnt_new"].ToString() + " ";
                                    }

                                    if ((ep.GetNameRC("dop_send", ref row_t, ref col_t)) && (dop_prod != ""))
                                    {
                                        ep.SetFormat(row_t + i + 1, col_t, "@");
                                        ep.SetText(row_t + i + 1, col_t, dop_prod);
                                    }
                                    //прием
                                    ds1.Clear();
                                    comm.Parameters.Clear();
                                    comm.CommandType = CommandType.StoredProcedure;
                                    comm.CommandText = "Rep_SummaryActReceive";
                                    comm.Parameters.Add("id_type_s", SqlDbType.Int).Value = (int)TypeDoc.SendToFilial;
                                    comm.Parameters.Add("id_type_r", SqlDbType.Int).Value = (int)TypeDoc.ReceiveToFilial;
                                    comm.Parameters.Add("id_branch", SqlDbType.Int).Value = Convert.ToInt32(ds.Tables[0].Rows[i]["id_branch"]);
                                    comm.Parameters.Add("dateDoc", SqlDbType.DateTime).Value = DatePickerOne.SelectedDate.Date;
                                    Database2.ExecuteCommand(comm, ref ds1, null);

                                    dop_prod = "";
                                    for (int j = 0; j < ds1.Tables[0].Rows.Count; j++)
                                    {
                                        if (Convert.ToInt32(ds1.Tables[0].Rows[j]["id_type"]) == 1)
                                            if (ep.GetNameRC(ds1.Tables[0].Rows[j]["prefix_file"].ToString(), ref row_t, ref col_t))
                                            {
                                                ep.SetFormat(row_t + i + 1, col_t, "@");
                                                ep.AddText(row_t + i + 1, col_t, "/" + ds1.Tables[0].Rows[j]["cnt_perso"].ToString());
                                            }

                                        if (Convert.ToInt32(ds1.Tables[0].Rows[j]["id_type"]) == 2)
                                            if (ep.GetNameRC("pin", ref row_t, ref col_t))
                                            {
                                                ep.SetFormat(row_t + i + 1, col_t, "@");
                                                ep.AddText(row_t + i + 1, col_t, "/" + ds1.Tables[0].Rows[j]["cnt_perso"].ToString());
                                            }

                                        if (Convert.ToInt32(ds1.Tables[0].Rows[j]["id_type"]) == 3)
                                            dop_prod = dop_prod + ds1.Tables[0].Rows[j]["name"].ToString() + ": " + ds1.Tables[0].Rows[j]["cnt_new"].ToString() + " ";

                                        if (j == ds1.Tables[0].Rows.Count - 1)
                                        {
                                            if (ep.GetNameRC("fio", ref row_t, ref col_t))
                                                ep.SetText(row_t + i + 1, col_t, ds1.Tables[0].Rows[j]["UserName"].ToString());

                                            if (ep.GetNameRC("comment", ref row_t, ref col_t))
                                                ep.SetText(row_t + i + 1, col_t, ds1.Tables[0].Rows[j]["comment"].ToString());
                                        }

                                    }

                                    if ((ep.GetNameRC("dop_receive", ref row_t, ref col_t)) && (dop_prod != ""))
                                    {
                                        ep.SetFormat(row_t + i + 1, col_t, "@");
                                        ep.SetText(row_t + i + 1, col_t, dop_prod);
                                    }

                                    cnt = i;
                                }

                                if (cnt > 0)
                                {
                                    if (ep.GetNameRC("comment", ref row_t, ref col_t))
                                    {
                                        ep.SetRangeAutoFit(row, 1, row + cnt + 1, col_t);
                                        ep.SetRangeBorders(row, 1, row + cnt + 1, col_t);
                                    }
                                }
                            }
                        }
                        #endregion
                        #region Распоряжение на движение ценностей (старый вариант)
                        //!!!!!!!!!!!!! если правишь здесь - поправь и в посылке отчета по почте
                        if (id_type == 13)
                        {
                            fname = String.Format("{0}RP_{1:ddMMyyyy}.xls", ConfigurationSettings.AppSettings["ArchivePath"], DatePickerOne.SelectedDate);
                            if (DatePickerOne.SelectedDate.Date != DateTime.Now.Date)
                            {
                                if (!File.Exists(fname))
                                {
                                    fname = "";
                                    lInfo.Text = "В архиве нет отчета за данное число. Состояние хранилища будет на текущую дату.";
                                }
                                else
                                    make = false;
                            }
                            if (make) // не делать если есть архивный
                            {
                                SqlCommand comm = new SqlCommand();
                                //проверяем, что есть неподтвержденные документы на персонализацию или отправку в филиал
                                comm.CommandText = "select number_doc from StorageDocs where date_doc=@d and priz_gen=0 and (type=5 or type=8)";
                                comm.Parameters.Add("@d", SqlDbType.DateTime).Value = DatePickerOne.SelectedDate.Date;
                                Database2.ExecuteCommand(comm, ref ds, null);
                                if (ds.Tables.Count > 0 && ds.Tables[0].Rows.Count > 0)
                                {
                                    string str = "";
                                    for (i = 0; i < ds.Tables[0].Rows.Count; i++)
                                        str += ", " + ds.Tables[0].Rows[i][0].ToString();
                                    str = str.Substring(2);
                                    lInfo.Text = String.Format("Создание отчета невозможно. За данное число есть неподтвержденные документы ({0})", str);
                                    return;
                                }
                                ds.Clear();
                                comm.CommandText = "select distinct id_prb, prod_name, bank_name,id_sort from V_Rep_Moving where priz_gen=1 and type=@type and (id_type=1 or id_type=2) and date_doc=@dateDoc order by id_sort";
                                comm.Parameters.Add("@type", SqlDbType.Int).Value = (int)TypeDoc.PersoCard;
                                comm.Parameters.Add("@dateDoc", SqlDbType.DateTime).Value = DatePickerOne.SelectedDate.Date;
                                res = (string)Database2.ExecuteCommand(comm, ref ds, null);
                                ArrayList al = new ArrayList();
                                foreach (DataRow dr in ds.Tables[0].Rows)
                                    al.Add(new MyProd(Convert.ToInt32(dr["id_prb"]), dr["prod_name"].ToString(), dr["bank_name"].ToString()));
                                comm.CommandText = "select sum(cnt_brak) as cntbrak, sum(cnt_perso) as cntperso from V_Rep_Moving where priz_gen=1 and type=@type and (id_type=1 or id_type=2) and date_doc=@dateDoc and id_prb=@prb";
                                comm.Parameters.Add("@prb", SqlDbType.Int).Value = 0;
                                i = 11;
                                foreach (MyProd mp in al)
                                {
                                    ds.Clear();
                                    comm.Parameters["@prb"].Value = mp.ID;
                                    res = (string)Database2.ExecuteCommand(comm, ref ds, null);
                                    if (ds.Tables[0].Rows.Count == 1)
                                    {
                                        int br = (ds.Tables[0].Rows[0]["cntbrak"] == DBNull.Value) ? 0 : Convert.ToInt32(ds.Tables[0].Rows[0]["cntbrak"]);
                                        int pr = (ds.Tables[0].Rows[0]["cntperso"] == DBNull.Value) ? 0 : Convert.ToInt32(ds.Tables[0].Rows[0]["cntperso"]);
                                        ep.SetText(i, 2, String.Format("{0}({1})", mp.Name, mp.Bank));
                                        ep.SetText(i, 3, String.Format("{0}шт.", br + pr));
                                        i++;
                                    }
                                }
                                ep.ShowRows(11, i);
                                al.Clear();
                                ds.Clear();
                                comm.CommandText = "select distinct id_prb, prod_name, bank_name,id_sort from V_Rep_Moving where priz_gen=1 and type=@type and (id_type=1 or id_type=2) and date_doc=@dateDoc order by id_sort";
                                comm.Parameters["@type"].Value = (int)TypeDoc.SendToFilial;
                                res = (string)Database2.ExecuteCommand(comm, ref ds, null);
                                foreach (DataRow dr in ds.Tables[0].Rows)
                                    al.Add(new MyProd(Convert.ToInt32(dr["id_prb"]), dr["prod_name"].ToString(), dr["bank_name"].ToString()));
                                comm.CommandText =
                                comm.CommandText = "select sum(cnt_perso) as cntperso from V_Rep_Moving where priz_gen=1 and type=@type and (id_type=1 or id_type=2) and date_doc=@dateDoc and id_prb=@prb and is_head=@ishead";
                                comm.Parameters.Add("@ishead", SqlDbType.Bit).Value = true;
                                i = 63;
                                foreach (MyProd mp in al)
                                {
                                    ds.Clear();
                                    comm.Parameters["@prb"].Value = mp.ID;
                                    res = (string)Database2.ExecuteCommand(comm, ref ds, null);
                                    if (ds.Tables[0].Rows.Count == 1)
                                    {
                                        int pr = (ds.Tables[0].Rows[0]["cntperso"] == DBNull.Value) ? 0 : Convert.ToInt32(ds.Tables[0].Rows[0]["cntperso"]);
                                        if (pr > 0)
                                        {
                                            ep.SetText(i, 2, String.Format("{0}({1})", mp.Name, mp.Bank));
                                            ep.SetText(i, 3, String.Format("{0}шт.", pr));
                                            i++;
                                        }
                                    }
                                }
                                ep.ShowRows(63, i);
                                comm.CommandText = "select sum(cnt_perso) as cntperso from V_Rep_Moving where priz_gen=1 and type=@type and (id_type=1 or id_type=2) and date_doc=@dateDoc and id_prb=@prb and (is_head=@ishead or is_head is null) and (is_rkc=@ishead or is_rkc is null) and (is_trans=@ishead or is_trans is null)";
                                comm.Parameters["@ishead"].Value = false;
                                i = 116;
                                foreach (MyProd mp in al)
                                {
                                    ds.Clear();
                                    comm.Parameters["@prb"].Value = mp.ID;
                                    res = (string)Database2.ExecuteCommand(comm, ref ds, null);
                                    if (ds.Tables[0].Rows.Count == 1)
                                    {
                                        int pr = (ds.Tables[0].Rows[0]["cntperso"] == DBNull.Value) ? 0 : Convert.ToInt32(ds.Tables[0].Rows[0]["cntperso"]);
                                        if (pr > 0)
                                        {
                                            ep.SetText(i, 2, String.Format("{0}({1})", mp.Name, mp.Bank));
                                            ep.SetText(i, 3, String.Format("{0}шт.", pr));
                                            i++;
                                        }
                                    }
                                }
                                ep.ShowRows(116, i);
                                comm.CommandText = "select sum(cnt_perso) as cntperso from V_Rep_Moving where priz_gen=1 and type=@type and (id_type=1 or id_type=2) and date_doc=@dateDoc and id_prb=@prb and is_rkc=@ishead";
                                comm.Parameters["@ishead"].Value = true;
                                i = 169;
                                foreach (MyProd mp in al)
                                {
                                    ds.Clear();
                                    comm.Parameters["@prb"].Value = mp.ID;
                                    res = (string)Database2.ExecuteCommand(comm, ref ds, null);
                                    if (ds.Tables[0].Rows.Count == 1)
                                    {
                                        int pr = (ds.Tables[0].Rows[0]["cntperso"] == DBNull.Value) ? 0 : Convert.ToInt32(ds.Tables[0].Rows[0]["cntperso"]);
                                        if (pr > 0)
                                        {
                                            ep.SetText(i, 2, String.Format("{0}({1})", mp.Name, mp.Bank));
                                            ep.SetText(i, 3, String.Format("{0}шт.", pr));
                                            i++;
                                        }
                                    }
                                }
                                ep.ShowRows(169, i);
                                comm.CommandText = "select sum(cnt_perso) as cntperso from V_Rep_Moving where priz_gen=1 and type=@type and (id_type=1 or id_type=2) and date_doc=@dateDoc and id_prb=@prb and is_trans=@ishead";
                                comm.Parameters["@ishead"].Value = true;
                                i = 223;
                                foreach (MyProd mp in al)
                                {
                                    ds.Clear();
                                    comm.Parameters["@prb"].Value = mp.ID;
                                    res = (string)Database2.ExecuteCommand(comm, ref ds, null);
                                    if (ds.Tables[0].Rows.Count == 1)
                                    {
                                        int pr = (ds.Tables[0].Rows[0]["cntperso"] == DBNull.Value) ? 0 : Convert.ToInt32(ds.Tables[0].Rows[0]["cntperso"]);
                                        if (pr > 0)
                                        {
                                            ep.SetText(i, 2, String.Format("{0}({1})", mp.Name, mp.Bank));
                                            ep.SetText(i, 3, String.Format("{0}шт.", pr));
                                            i++;
                                        }
                                    }
                                }
                                ep.ShowRows(223, i);

                                ep.SetText_Name("date_string", String.Format("от {0:dd MM yyyy}", DatePickerOne.SelectedDate));
                                ep.SetText_Name("date_string1", String.Format("{0:dd.MM.yyyy}", DatePickerOne.SelectedDate));
                                /////////////////////////////////
                                //второй лист - хранилище
                                /////////////////////////////////
                                ds.Clear();
                                ep.SetWorkSheet(2);
                                Database2.ExecuteQuery("select name, bank_name, cnt_new, cnt_perso from V_Storage where id_type=1 order by id_sort", ref ds, null);
                                int rw = 4;
                                for (i = 0; i < ds.Tables[0].Rows.Count; i++)
                                {
                                    string str = String.Format("{0} ({1})", ds.Tables[0].Rows[i]["name"].ToString(), ds.Tables[0].Rows[i]["bank_name"].ToString());
                                    ep.SetText(rw, 2, str);
                                    ep.SetText(rw, 3, ds.Tables[0].Rows[i]["cnt_new"].ToString());
                                    ep.SetText(rw, 4, ds.Tables[0].Rows[i]["cnt_perso"].ToString());
                                    rw++;
                                }
                                rw++;
                                ds.Clear();
                                Database2.ExecuteQuery("select name, bank_name, cnt_new, cnt_perso from V_Storage where id_type=2 order by id_sort", ref ds, null);
                                ep.SetText(rw, 3, "Бланки (шт)");
                                ep.SetText(rw, 4, "Персо");
                                rw++;
                                for (i = 0; i < ds.Tables[0].Rows.Count; i++)
                                {
                                    string str = String.Format("{0} ({1})", ds.Tables[0].Rows[i]["name"].ToString(), ds.Tables[0].Rows[i]["bank_name"].ToString());
                                    ep.SetText(rw, 2, str);
                                    ep.SetText(rw, 3, ds.Tables[0].Rows[i]["cnt_new"].ToString());
                                    ep.SetText(rw, 4, ds.Tables[0].Rows[i]["cnt_perso"].ToString());
                                    rw++;
                                }
                                ds.Clear();
                                ep.SetRangeBorders(3, 2, rw - 1, 4);
                                ep.SetText_Name("date_string2", String.Format("на {0:dd.MM.yyyy}", DateTime.Now));
                            }
                        }
                        #endregion
                        #region Распоряжение на движение ценностей (новый вариант)
                        //!!!!!!!!!!!!! если правишь здесь - поправь и в посылке отчета по почте
                        //!!!!!!!!!!!!! а еще поправь теперь и в автоформировании отчета
                        if (id_type == 15)
                        {
                            fname = String.Format("{0}RPn_{1:ddMMyyyy}.xls", ConfigurationSettings.AppSettings["ArchivePath"], DatePickerOne.SelectedDate);
                            if (DatePickerOne.SelectedDate.Date != DateTime.Now.Date)
                            {
                                if (!File.Exists(fname))
                                {
                                    fname = "";
                                    lInfo.Text = "В архиве нет отчета за данное число. Состояние хранилища будет на текущую дату.";
                                }
                                else
                                    make = false;
                            }
                            if (make) // не делать если есть архивный
                            {
                                MovingStock(ep);
                            }
                        }
                        #endregion
                        #region Выданные ценности
                        if (id_type == 14)
                        {
                            SqlCommand comm = new SqlCommand();
                            comm.CommandText = "select prod_name, bank_name, dbo.MovingCnt(id_prb, 1, @dateStart, @dateEnd) as perso_cnt, dbo.MovingCnt(id_prb, 2, @dateStart, @dateEnd) as new_cnt from V_ProductsBanks_T where id_type=3";
                            comm.Parameters.Add("dateStart", SqlDbType.DateTime).Value = DatePickerStart.SelectedDate;
                            comm.Parameters.Add("dateEnd", SqlDbType.DateTime).Value = DatePickerEnd.SelectedDate;
                            object obj = Database2.ExecuteCommand(comm, ref ds, null);
                            int t = 0;
                            for (i = 0; i < ds.Tables[0].Rows.Count; i++)
                            {
                                if (ds.Tables[0].Rows[i]["new_cnt"] != DBNull.Value && Convert.ToInt32(ds.Tables[0].Rows[i]["new_cnt"]) > 0)
                                {
                                    string str = String.Format("{0} ({1})", ds.Tables[0].Rows[i]["prod_name"], ds.Tables[0].Rows[i]["bank_name"]);
                                    ep.SetText(t + 5, 2, str);
                                    ep.SetText(t + 5, 3, Convert.ToInt32(ds.Tables[0].Rows[i]["new_cnt"]).ToString());
                                    t++;
                                }
                            }
                            ep.SetText_Name("title", "Выданные доп. ценности");
                            ep.SetText_Name("date_string", String.Format("за период с {0} по {1}", DatePickerStart.SelectedDate.ToString("dd.MM.yyyy"), DatePickerEnd.SelectedDate.ToString("dd.MM.yyyy")));
                            ep.SetText_Name("sost_string", String.Format("подготовлено {0}", DateTime.Now.ToString("dd.MM.yyyy")));
                        }
                        #endregion
                        #region карты выданные клиентам подразделением ГО
                        if (id_type == 16)
                        {
                            GivenCardGO(ep);
                        }
                        #endregion
                        #region карты выданные клиентам подраздиления
                        if (id_type == 17)
                        {
                            GivenCardBranch(ep, ddlBranch.SelectedValue);
                        }
                        #endregion
                        #region Отчет МЕМОРИАЛЬНЫЙ ОРДЕР

                        if (id_type == 42)
                        {

                            int idUser = sc.UserId(User.Identity.Name);
                            int idBranch = sc.BranchId(User.Identity.Name);
                            String sqlUserV = " " + idUser.ToString() + "=(select [user_id] from StorageDocs where V_Cards_StorageDocs1.id=StorageDocs.id) and ";
                            String sqlUser = " " + idUser.ToString() + "=[user_id] and ";

                            //if (idUser < 1 || userradio.Checked==false)
                            //{
                            sqlUserV = " ";
                            sqlUser = " ";
                            //}

                            SqlCommand comm = Database2.Conn.CreateCommand();
                            DateTime sDate = DatePickerOne.SelectedDate.Date;

                            OperationDay op = new OperationDay();
                            op.read(current_branch);


                            DateTime curDateS = op.getDateTimeStart(sDate);
                            DateTime curDateE = op.getDateTimeEnd(sDate);

                            if (op.isShift && (firstrangeradio.Checked == true || nextrangeradio.Checked == true))
                            {
                                DateTime[] tm = op.getDateTimePart(sDate, firstrangeradio.Checked);
                                if (tm != null)
                                {
                                    curDateS = tm[0];
                                    curDateE = tm[1];
                                    tm = null;
                                }
                            }
                            if (op.isShift && allrangeradio.Checked == true)
                            {
                                DateTime[] tm = op.getDateTimePart(sDate, false);
                                if (tm != null)
                                {
                                    curDateS = tm[0];
                                    curDateE = tm[1];
                                    tm = null;
                                }
                            }

                            //op.isShift = false;                                                 

                            ds.Clear();
                            comm.Parameters.Clear();
                            comm.CommandText = "select top 1 * from Branchs where id=@branch";
                            comm.Parameters.Add("@branch", SqlDbType.Int).Value = idBranch;
                            Database2.ExecuteCommand(comm, ref ds, null);
                            try
                            {
                                if (ds.Tables.Count > 0 && ds.Tables[0].Rows.Count > 0)
                                {
                                    BranchStore branch = new BranchStore(ds.Tables[0].Rows[0]["id"], ds.Tables[0].Rows[0]["ident_dep"], ds.Tables[0].Rows[0]["department"]);
                                    BranchStore branch_prev = new BranchStore(ds.Tables[0].Rows[0]["id"], ds.Tables[0].Rows[0]["ident_dep"], ds.Tables[0].Rows[0]["department"]);
                                    DataSet dsProduct = new DataSet();
                                    string id_prbs = getExcludeMemoricProducts();
                                    if (id_prbs.Length < 1)
                                        Database2.ExecuteQuery("select id_prb, id_prod, prod_name, id_type from V_ProductsBanks_T where id_type in(1,2) order by id_type, id_sort", ref dsProduct, null);
                                    else
                                        Database2.ExecuteQuery("select id_prb, id_prod, prod_name, id_type from V_ProductsBanks_T where id_type in(1,2) and id_prb not in (" + id_prbs + ") order by id_type, id_sort", ref dsProduct, null);
                                    //Database2.ExecuteQuery("select id_prb, id_prod, prod_name, id_type from V_ProductsBanks_T where id_type in(1,2) order by id_type, id_sort", ref dsProduct, null);
                                    ds.Clear();
                                    comm.Parameters.Clear();
                                    ArrayList al = new ArrayList();
                                    if (dsProduct.Tables.Count > 0)
                                    {
                                        //String sql = "select id_prb, count(*) as count_prb from Cards  where id_stat=4 and id_branchCurrent=" + branch.id.ToString() + " group by id_prb";
                                        //if(branch_main_filial<1)
                                        //       sql = "select id_prb, count(*) as count_prb from Cards  where id_stat=4 and id_branchCard=" + branch.id.ToString() + " group by id_prb";
                                        //Database2.ExecuteQuery(sql, ref ds, null);
                                        for (i = 0; i < dsProduct.Tables[0].Rows.Count; i++)
                                        {
                                            MyProd mp = new MyProd(Convert.ToInt32(dsProduct.Tables[0].Rows[i]["id_prb"]),
                                                Convert.ToString(dsProduct.Tables[0].Rows[i]["prod_name"]), "", Convert.ToInt32(dsProduct.Tables[0].Rows[i]["id_type"]));
                                            al.Add(mp);
                                        }
                                    }
                                    if (op.isShift == true && (nextrangeradio.Checked == true || allrangeradio.Checked == true))
                                    {
                                        DateTime[] tm = op.getDateTimePart(sDate, true);
                                        if (tm != null) setMemoric(branch_prev, al, tm[0], tm[1], sqlUserV, sqlUser);
                                    }
                                    setMemoric(branch, al, curDateS, curDateE, sqlUserV, sqlUser);
                                   
                                    if (branch.isEmpty() && branch_prev.isEmpty()) throw new Exception("На " + sDate.ToString("dd.MM.yyyy") + " нет операций по продуктам для подразделения: " + branch.department);
                                    // формирование
                                    int count = branch.countNonEmpty();
                                    if (op.isShift == true)
                                    {
                                        if (allrangeradio.Checked == true) count += branch_prev.countNonEmpty();
                                    }

                                    for (int pp = 0; pp < count; pp++)
                                    {
                                        if (pp < count - 1)
                                        {
                                            ep.SetWorkSheet(1);
                                            ep.AddWorkSheet(1 + pp);
                                        }
                                    }

                                    List<AccountBranch> listAB = getAccountBranch(idBranch);
                                    string isTrans = "99999";
                                    string trans = "Транзитный счет";
                                    string sCurDate = sDate.ToString("dd.MM.yyyy");
                                    int doc_num = 1; // OperationDay.getBranchStartNumber(idBranch);
                                    //if (doc_num < 0) doc_num = 1;
                                    //string doc_prefix = OperationDay.getBranchStartNumber(idBranch,sDate,true);
                                    int smena = 1;
                                                                        
                                    //else
                                    //{
                                    //    if (nextrangeradio.Checked == true) doc_num += OperationDay.getShiftNumber();
                                    //}
                                    if (nextrangeradio.Checked == true) smena = 2;//doc_prefix = OperationDay.getBranchStartNumber(idBranch, sDate, false);
                                    int p = 0;
                                    for (int k = 0; k < 2; k++)
                                    {
                                        BranchStore bs;
                                        if (op.isShift == false || (op.isShift == true && firstrangeradio.Checked == true)) k = 1;
                                        if (k == 0) bs = branch_prev;
                                        else bs = branch;
                                        if (op.isShift == true)
                                        {
                                            if (k == 0 && nextrangeradio.Checked == true) continue;
                                            if (k == 1 && allrangeradio.Checked == true)
                                            {
                                                //doc_num = OperationDay.getBranchStartNumber(idBranch);
                                                //if (doc_num < 0) doc_num = 1;
                                                //doc_num +=OperationDay.getShiftNumber();
                                                doc_num = 1;
                                                //doc_prefix = OperationDay.getBranchStartNumber(idBranch, sDate, false);
                                                smena = 2;
                                            }
                                        }

                                        
                                        int maxN = Enum.GetNames(typeof(BaseProductType)).Length - 1;
                                        int maxE = maxN * 3;

                                        for (int n = 0; n < maxE; n++, doc_num++)
                                        {
                                            int countP = 0;
                                            String nameP = "";
                                            String nameS = "";

                                            BaseProductType bp = BaseProductType.None;
                                            AccountBranchType ab = AccountBranchType.None;

                                            int abt = n / maxN;

                                            if (n % maxN == 0) { countP = bs.countMasterCard[abt]; nameS = "MC"; nameP = "ПК Master Card"; bp = BaseProductType.MasterCard; }
                                            if (n % maxN == 1) { countP = bs.countVisaCard[abt]; nameS = "VISA"; nameP = "ПК Visa"; bp = BaseProductType.VisaCard; }
                                            if (n % maxN == 2) { countP = bs.countNFCCard[abt]; nameS = "NFC"; nameP = "ПК NFC"; bp = BaseProductType.NFCCard; }
                                            if (n % maxN == 3) { countP = bs.countServiceCard[abt]; nameS = "Service"; nameP = "ПК Сервисные"; bp = BaseProductType.ServiceCard; }
                                            if (n % maxN == 4) { countP = bs.countMirCard[abt]; nameS = "MIR"; nameP = "ПК МИР"; bp = BaseProductType.MirCard; }
                                            if (n % maxN == 5) { countP = bs.countPinConvert[abt]; nameS = "PIN"; nameP = "Бланки ПИН-конвертов"; bp = BaseProductType.PinConvert; }

                                            if (abt == 0) ab = AccountBranchType.In;
                                            if (abt == 1) ab = AccountBranchType.Out;
                                            if (abt == 2) ab = AccountBranchType.Return;

                                            if (countP < 1) continue;
                                            p++;
                                            ep.SetWorkSheet(p);
                                            if (ab == AccountBranchType.In) ep.SetWorkSheetName(p, p.ToString() + "_Приняты_" + nameS);
                                            if (ab == AccountBranchType.Out) ep.SetWorkSheetName(p, p.ToString() + "_Переданы_" + nameS);
                                            if (ab == AccountBranchType.Return) ep.SetWorkSheetName(p, p.ToString() + "_Инкассация_" + nameS);
                                            ep.SetText_Name("CURDATE", sCurDate);
                                            ep.SetText_Name("CARDCOUNT", (countP).ToString() + "=");
                                            ep.SetText_Name("PRODDEBET", nameP);
                                            ep.SetText_Name("PRODCREDIT", nameP);
                                            int row = 0, col = 0;
                                            String txt;

                                            row = 31; col = 2;
                                            txt = ep.GetCell(row, col);
                                            txt = txt.Replace("X", sCurDate);

                                            if (ab == AccountBranchType.Out)
                                            {
                                                txt = txt.Substring(0, txt.IndexOf("Y"));
                                                txt = txt.Replace("W", "Переданы");

                                            }
                                            if (ab == AccountBranchType.In)
                                            {
                                                txt = txt.Replace("W", "Приняты");
                                                txt = txt.Replace("Y", "из");
                                                txt = txt.Replace("Z", "а");


                                            }
                                            if (ab == AccountBranchType.Return)
                                            {
                                                txt = txt.Replace("W", "Инкассация");
                                                txt = txt.Replace("Y", "в");
                                                txt = txt.Replace("Z", "е");

                                            }
                                            ep.SetText(row, col, txt);
                                            row = 12; col = 2;
                                            txt = ep.GetCell(row, col);
                                            //txt += " " + p.ToString();
                                            //txt += doc_prefix + doc_num.ToString("D03");
                                            txt += OperationDay.getBranchStartNumber(idBranch, sDate, doc_num, smena);
                                            ep.SetText(row, col, txt);

                                            AccountBranch AB = listAB.FirstOrDefault(ai => ai.accountType == ab && ai.productType == bp);
                                            String debet = "";
                                            String credit = "";
                                            if (AB != null)
                                            {
                                                debet = AB.accountDebet;
                                                credit = AB.accountCredit;
                                            }
                                            ep.SetText_Name("ACCOUNTDEBET", debet);
                                            ep.SetText_Name("ACCOUNTCREDIT", credit);
                                            if (debet.StartsWith(isTrans)) ep.SetText_Name("BRANCHDEBET", trans);
                                            else ep.SetText_Name("BRANCHDEBET", bs.department);
                                            if (credit.StartsWith(isTrans)) ep.SetText_Name("BRANCHCREDIT", trans);
                                            else ep.SetText_Name("BRANCHCREDIT", bs.department);
                                            decimal money = Convert.ToDecimal(countP);
                                            ep.SetText_Name("SUMMA", bs.RurPhrase(money));

                                            if (ab == AccountBranchType.Out) doc_num++; // в экспорте dbf еще одна запись на выдачу
                                        }
                                    }
                                    ep.SetWorkSheet(1);
                                }
                                else throw new Exception("На " + sDate.ToString("dd.MM.yyyy") + " не найдено подразделение " + idBranch.ToString());

                            }
                            catch (Exception e1)
                            {
                                ep.Close();
                                GC.Collect();
                                System.Threading.Thread.CurrentThread.CurrentCulture = oldCI;
                                ClientScript.RegisterClientScriptBlock(GetType(), "err42", "<script type='text/javascript'>$(document).ready(function(){ ShowMessage('" + e1.Message + "');});</script>");
                                return;
                            }

                        }

                        #endregion
                        #region Отчет для филиала на конец дня или передачи между сменами
                        if (id_type == 19 || id_type == 41)
                        {


                            OperationDay op = new OperationDay();
                            op.read(current_branch);

                            DateTime curDate = DateTime.Now.Date;
                            DateTime toDate = DateTime.Now.Date;

                            //if (id_type == 19) 
                            //{
                                curDate = DatePickerOne.SelectedDate.Date;
                            //}

                            DateTime curDateS = op.getDateTimeStart(curDate);
                            DateTime curDateE = op.getDateTimeEnd(curDate);
                            DateTime toDateS = op.getDateTimeStart(toDate);
                            DateTime toDateE = op.getDateTimeEnd(toDate);

                            SqlCommand comm = Database2.Conn.CreateCommand();

                            ds.Clear();
                            comm.Parameters.Clear();

                            comm.CommandText = "select top 1 * from Branchs where id=@branch";

                            comm.Parameters.Add("@branch", SqlDbType.Int).Value = sc.BranchId(User.Identity.Name);
                            Database2.ExecuteCommand(comm, ref ds, null);

                            if (ds.Tables.Count > 0 && ds.Tables[0].Rows.Count > 0)
                            {//--------------------------------

                                BranchStore branch = new BranchStore(ds.Tables[0].Rows[0]["id"], ds.Tables[0].Rows[0]["ident_dep"], ds.Tables[0].Rows[0]["department"]);
                                DataSet dsProduct = new DataSet();
                                Database2.ExecuteQuery("select id_prb, id_prod, prod_name, id_type from V_ProductsBanks_T where id_type in(1,2) order by id_type, id_sort", ref dsProduct, null);

                                ds.Clear();
                                comm.Parameters.Clear();
                                ArrayList al = new ArrayList();
                                if (dsProduct.Tables.Count > 0)
                                {
                                    String sql = "select id_prb, count(*) as count_prb from Cards  where id_stat in (4,6,11) and id_branchCurrent=" + branch.id.ToString() + " group by id_prb";
                                    //~!if (branch_main_filial < 1)
                                    //~!    sql = "select id_prb, count(*) as count_prb from Cards  where id_stat in (4,6) and id_branchCard=" + branch.id.ToString() + " group by id_prb";
                                    Database2.ExecuteQuery(sql, ref ds, null);
                                    for (i = 0; i < dsProduct.Tables[0].Rows.Count; i++)
                                    {
                                        MyProd mp = new MyProd(Convert.ToInt32(dsProduct.Tables[0].Rows[i]["id_prb"]),
                                            Convert.ToString(dsProduct.Tables[0].Rows[i]["prod_name"]), "", Convert.ToInt32(dsProduct.Tables[0].Rows[i]["id_type"]));
                                        for (int n = 0; mp.Type != 2 && ds.Tables.Count > 0 && n < ds.Tables[0].Rows.Count; n++)
                                        {
                                            if (mp.ID == Convert.ToInt32(ds.Tables[0].Rows[n]["id_prb"]))
                                            {
                                                mp.cnts[0] = Convert.ToInt32(ds.Tables[0].Rows[n]["count_prb"]);
                                                break;
                                            }
                                        }
                                        al.Add(mp);
                                        if (mp.Type == 2) // Пин конверты на конец дня
                                        {
                                            comm.CommandText = "select count(*) from Cards where id_stat in (4,6) and id_branchCurrent=" + branch.id.ToString() + " and isPin=1";
                                            //~!if (branch_main_filial < 1)
                                            //~!    comm.CommandText = "select count(*) from Cards where id_stat in (4,6) and id_branchCard=" + branch.id.ToString() + " and isPin=1";
                                            object obj = comm.ExecuteScalar();
                                            mp.cnts[0] = (obj == null || obj == DBNull.Value) ? 0 : Convert.ToInt32(obj);
                                        }
                                        branch.addCount(mp.Type, mp.Name, 0, mp.cnts[0]);
                                    }
                                }
                                if (/*id_type == 19 && */curDateS < toDateS)
                                {
                                    ds.Clear();
                                    comm.Parameters.Clear();
                                    //корректировка выдачи
                                    comm.CommandText = "select dbo.Cards.id_prb,count(*) as count_prb" +
                                                   " from dbo.Branchs RIGHT OUTER JOIN" +
                                                   " dbo.StorageDocs ON dbo.Branchs.id = dbo.StorageDocs.id_branch LEFT OUTER JOIN" +
                                                   " dbo.Cards RIGHT OUTER JOIN" +
                                                   " dbo.Cards_StorageDocs ON dbo.Cards.id = dbo.Cards_StorageDocs.id_card ON dbo.StorageDocs.id = dbo.Cards_StorageDocs.id_doc" +
                                                   " where date_time>=@dateS and date_time<@dateP and" +
                                                   " (id_branch=@id_branch and (type=7 or type=9 or type=11 or type=13 or type = 24 or type = 21 or type = 30) or (id_act=@id_branch and (type=19 or type=5)))" +
                                                   " and priz_gen=1 group by dbo.Cards.id_prb";
                                    comm.Parameters.Add("@dateS", SqlDbType.DateTime).Value = curDateE;
                                    comm.Parameters.Add("@dateP", SqlDbType.DateTime).Value = toDateE;
                                    comm.Parameters.Add("@id_branch", SqlDbType.Int).Value = branch.id;
                                    Database2.ExecuteCommand(comm, ref ds, null);
                                    foreach (MyProd mp in al)
                                    {
                                        int countMinus = 0;
                                        for (int n = 0; mp.Type != 2 && ds.Tables.Count > 0 && n < ds.Tables[0].Rows.Count; n++)
                                        {
                                            if (mp.ID == Convert.ToInt32(ds.Tables[0].Rows[n]["id_prb"]))
                                            {
                                                countMinus = Convert.ToInt32(ds.Tables[0].Rows[n]["count_prb"]);
                                                break;
                                            }
                                        }
                                        if (mp.Type == 2)
                                        {
                                            comm.Parameters.Clear();
                                            //comm.CommandText = "select count(*) from cards where id in (select id_card from cards_storagedocs where id_doc in (select id from storagedocs where date_time>=@dateS and date_time<@dateP and ((id_branch=@id_branch and (type=7 or type=9 or type=11 or type=13)) or (id_act=@id_branch and type=19)) and priz_gen=1)) and isPin=1";
                                            comm.CommandText = "select count(id_card) from cards_storagedocs cs join cards c on c.id=cs.id_card where id_doc in (select id from storagedocs where date_time>=@dateS and date_time<@dateP and ((id_branch=@id_branch and (type=7 or type=9 or type=11 or type=13 or type=24)) or (id_act=@id_branch and (type=19 or type=5))) and priz_gen=1) and isPin=1";
                                            comm.Parameters.Add("@dateS", SqlDbType.DateTime).Value = curDateE;
                                            comm.Parameters.Add("@dateP", SqlDbType.DateTime).Value = toDateE;
                                            comm.Parameters.Add("@id_branch", SqlDbType.Int).Value = branch.id;
                                            object obj = comm.ExecuteScalar();
                                            countMinus = (obj == null || obj == DBNull.Value) ? 0 : Convert.ToInt32(obj);
                                        }
                                        branch.addCount(mp.Type, mp.Name, 0, countMinus);
                                        mp.cnts[0] += countMinus;
                                    }
                                    //корректировка получения
                                    ds.Clear();
                                    comm.Parameters.Clear();
                                    comm.CommandText = "select id_prb,count(*) as count_prb from V_Cards_StorageDocs1 where date_time>=@dateS and date_time<@dateP and id_branch=@id_branch and (type=6 or type=10 or type=28) and priz_gen=1 group by id_prb";
                                    comm.Parameters.Add("@dateS", SqlDbType.DateTime).Value = curDateE;
                                    comm.Parameters.Add("@dateP", SqlDbType.DateTime).Value = toDateE;
                                    comm.Parameters.Add("@id_branch", SqlDbType.Int).Value = branch.id;
                                    Database2.ExecuteCommand(comm, ref ds, null);
                                    foreach (MyProd mp in al)
                                    {

                                        int countPlus = 0;
                                        for (int n = 0; mp.Type != 2 && ds.Tables.Count > 0 && n < ds.Tables[0].Rows.Count; n++)
                                        {
                                            if (mp.ID == Convert.ToInt32(ds.Tables[0].Rows[n]["id_prb"]))
                                            {
                                                countPlus = Convert.ToInt32(ds.Tables[0].Rows[n]["count_prb"]);
                                                break;
                                            }
                                        }
                                        if (mp.Type == 2) // Пин конверты получены
                                        {

                                            comm.Parameters.Clear();
                                            //comm.CommandText = "select count(*) from cards where id in (select id_card from cards_storagedocs where id_doc in (select id from storagedocs where date_time>=@dateS and date_time<@dateP and (type=6 or type=10) and priz_gen=1 and id_branch=@id_branch)) and isPin=1";
                                            comm.CommandText = "select count(id_card) from cards_storagedocs cs join cards c on c.id=cs.id_card where id_doc in (select id from storagedocs where date_time>=@dateS and date_time<@dateP and (type=6 or type=10 or type=28) and priz_gen=1 and id_branch=@id_branch) and isPin=1";
                                            comm.Parameters.Add("@dateS", SqlDbType.DateTime).Value = curDateE;
                                            comm.Parameters.Add("@dateP", SqlDbType.DateTime).Value = toDateE;
                                            comm.Parameters.Add("@id_branch", SqlDbType.Int).Value = branch.id;
                                            object obj = comm.ExecuteScalar();
                                            countPlus = (obj == null || obj == DBNull.Value) ? 0 : Convert.ToInt32(obj);
                                        }
                                        branch.addCount(mp.Type, mp.Name, 0, -countPlus);
                                        mp.cnts[0] -= countPlus;
                                    }

                                }

                                ds.Clear();
                                comm.Parameters.Clear();
                                //выдача
                                //comm.CommandText = "select id_prb,count(*) as count_prb from V_Cards_StorageDocs1 where date_doc=@dateS and id_branch=@id_branch and (type=7 or type=9 or type=11 or type=19) and priz_gen=1 group by id_prb";
                                comm.CommandText = "select dbo.Cards.id_prb,count(*) as count_prb" +
                                                   " from dbo.Branchs RIGHT OUTER JOIN" +
                                                   " dbo.StorageDocs ON dbo.Branchs.id = dbo.StorageDocs.id_branch LEFT OUTER JOIN" +
                                                   " dbo.Cards RIGHT OUTER JOIN" +
                                                   " dbo.Cards_StorageDocs ON dbo.Cards.id = dbo.Cards_StorageDocs.id_card ON dbo.StorageDocs.id = dbo.Cards_StorageDocs.id_doc" +
                                                   " where date_time>=@dateS and date_time<@dateE and" +
                                                   " (id_branch=@id_branch and (type=7 or type=9 or type=11 or type=13  or type = 24 or type = 21 or type = 30) or (id_act=@id_branch and (type=19 or type=5)))" +
                                                   " and priz_gen=1 group by dbo.Cards.id_prb";
                                comm.Parameters.Add("@dateS", SqlDbType.DateTime).Value = curDateS;
                                comm.Parameters.Add("@dateE", SqlDbType.DateTime).Value = curDateE;
                                comm.Parameters.Add("@id_branch", SqlDbType.Int).Value = branch.id;
                                Database2.ExecuteCommand(comm, ref ds, null);
                                foreach (MyProd mp in al)
                                {
                                    for (int n = 0; mp.Type != 2 && ds.Tables.Count > 0 && n < ds.Tables[0].Rows.Count; n++)
                                    {
                                        if (mp.ID == Convert.ToInt32(ds.Tables[0].Rows[n]["id_prb"]))
                                        {
                                            mp.cnts[1] = Convert.ToInt32(ds.Tables[0].Rows[n]["count_prb"]);
                                            break;
                                        }
                                    }
                                    if (mp.Type == 2) // Пин конверты выданы
                                    {
                                        comm.Parameters.Clear();
                                        //comm.CommandText = "select count(*) from cards where id in (select id_card from cards_storagedocs where id_doc in (select id from storagedocs where date_time>=@dateS and date_time<@dateE and ((id_branch=@id_branch and (type=7 or type=9 or type=11 or type=13)) or (id_act=@id_branch and type=19)) and priz_gen=1)) and isPin=1";
                                        comm.CommandText = "select count(id_card) from cards_storagedocs cs join cards c on c.id=cs.id_card where id_doc in (select id from storagedocs where date_time>=@dateS and date_time<@dateE and ((id_branch=@id_branch and (type=7 or type=9 or type=11 or type=13  or type = 24)) or (id_act=@id_branch and (type=19 or type=5))) and priz_gen=1) and isPin=1";
                                        comm.Parameters.Add("@dateS", SqlDbType.DateTime).Value = curDateS;
                                        comm.Parameters.Add("@dateE", SqlDbType.DateTime).Value = curDateE;
                                        comm.Parameters.Add("@id_branch", SqlDbType.Int).Value = branch.id;
                                        object obj = comm.ExecuteScalar();
                                        mp.cnts[1] = (obj == null || obj == DBNull.Value) ? 0 : Convert.ToInt32(obj);
                                    }
                                    branch.addCount(mp.Type, mp.Name, 1, mp.cnts[1]);
                                }
                                //получение
                                ds.Clear();
                                comm.Parameters.Clear();
                                comm.CommandText = "select id_prb,count(*) as count_prb from V_Cards_StorageDocs1 where date_time>=@dateS and date_time<@dateE and id_branch=@id_branch and (type=6 or type=10 or type=28) and priz_gen=1 group by id_prb";
                                comm.Parameters.Add("@dateS", SqlDbType.DateTime).Value = curDateS;
                                comm.Parameters.Add("@dateE", SqlDbType.DateTime).Value = curDateE;
                                comm.Parameters.Add("@id_branch", SqlDbType.Int).Value = branch.id;
                                Database2.ExecuteCommand(comm, ref ds, null);
                                foreach (MyProd mp in al)
                                {

                                    for (int n = 0; mp.Type != 2 && ds.Tables.Count > 0 && n < ds.Tables[0].Rows.Count; n++)
                                    {
                                        if (mp.ID == Convert.ToInt32(ds.Tables[0].Rows[n]["id_prb"]))
                                        {
                                            mp.cnts[2] = Convert.ToInt32(ds.Tables[0].Rows[n]["count_prb"]);
                                            break;
                                        }
                                    }
                                    if (mp.Type == 2) // Пин конверты получены
                                    {

                                        comm.Parameters.Clear();
                                        //comm.CommandText = "select count(*) from cards where id in (select id_card from cards_storagedocs where id_doc in (select id from storagedocs where date_time>=@dateS and date_time<@dateE and (type=6 or type=10) and priz_gen=1 and id_branch=@id_branch)) and isPin=1";
                                        comm.CommandText = "select count(id_card) from cards_storagedocs cs join cards c on c.id=cs.id_card where id_doc in (select id from storagedocs where date_time>=@dateS and date_time<@dateE and (type=6 or type=10 or type=28) and priz_gen=1 and id_branch=@id_branch) and isPin=1";
                                        comm.Parameters.Add("@dateS", SqlDbType.DateTime).Value = curDateS;
                                        comm.Parameters.Add("@dateE", SqlDbType.DateTime).Value = curDateE;
                                        comm.Parameters.Add("@id_branch", SqlDbType.Int).Value = branch.id;
                                        object obj = comm.ExecuteScalar();
                                        mp.cnts[2] = (obj == null || obj == DBNull.Value) ? 0 : Convert.ToInt32(obj);
                                    }
                                    branch.addCount(mp.Type, mp.Name, 2, mp.cnts[2]);
                                }
                                if (id_type == 19)
                                {
                                    //выдача только карт
                                    ds.Clear();
                                    comm.Parameters.Clear();
                                    comm.CommandText = "select id_prb,count(*) as count_prb from V_Cards_StorageDocs1 where date_time>=@dateS and date_time<@dateE and id_branch=@id_branch and type=7 and priz_gen=1 group by id_prb";
                                    comm.Parameters.Add("@dateS", SqlDbType.DateTime).Value = curDateS;
                                    comm.Parameters.Add("@dateE", SqlDbType.DateTime).Value = curDateE;
                                    comm.Parameters.Add("@id_branch", SqlDbType.Int).Value = branch.id;
                                    Database2.ExecuteCommand(comm, ref ds, null);
                                    foreach (MyProd mp in al)
                                    {

                                        for (int n = 0; mp.Type != 2 && ds.Tables.Count > 0 && n < ds.Tables[0].Rows.Count; n++)
                                        {
                                            if (mp.ID == Convert.ToInt32(ds.Tables[0].Rows[n]["id_prb"])) // карты выданы
                                            {
                                                mp.cnts[4] = Convert.ToInt32(ds.Tables[0].Rows[n]["count_prb"]);
                                                break;
                                            }
                                        }
                                        if (mp.Type == 2) // Пин конверты выданы
                                        {
                                            comm.Parameters.Clear();
                                            //comm.CommandText = "select count(*) from cards where id in (select id_card from cards_storagedocs where id_doc in (select id from storagedocs where date_time>=@dateS and date_time<@dateE and type=7 and priz_gen=1 and id_branch=@id_branch)) and isPin=1";
                                            comm.CommandText = "select count(id_card) from cards_storagedocs cs join cards c on c.id=cs.id_card where id_doc in (select id from storagedocs where date_time>=@dateS and date_time<@dateE and type=7 and priz_gen=1 and id_branch=@id_branch) and isPin=1";
                                            comm.Parameters.Add("@dateS", SqlDbType.DateTime).Value = curDateS;
                                            comm.Parameters.Add("@dateE", SqlDbType.DateTime).Value = curDateE;
                                            comm.Parameters.Add("@id_branch", SqlDbType.Int).Value = branch.id;
                                            object obj = comm.ExecuteScalar();
                                            mp.cnts[4] = (obj == null || obj == DBNull.Value) ? 0 : Convert.ToInt32(obj);
                                        }
                                        branch.addCount(mp.Type, mp.Name, 4, mp.cnts[4]);
                                    }
                                }

                                i = 0;
                                ep.SetWorkSheet(1);

                                if (id_type == 19) ep.SetText(5, 1, branch.department);

                                string fio = "";
                                int curt = 0;
                                int curNum = 0;
                                int[] curSum = new int[10];
                                int[] cardSum = new int[10];
                                int sL = 8, sC = 3;
                                if (id_type != 19)
                                {
                                    sL = 31; sC = 1;
                                }

                                foreach (MyProd mp in al)
                                {
                                    if (mp.cnts[0] + mp.cnts[1] + mp.cnts[2] > 0)
                                    {

                                        int s = mp.cnts[0] + mp.cnts[1] - mp.cnts[2];
                                        //s = (s > 0) ? s : 0;

                                        int t = branch.addCount(mp.Type, mp.Name, 3, s);
                                        if (t < 1) continue;

                                        if (curt == 0) curt = t;
                                        else
                                        {
                                            if (curt != t)
                                            {
                                                string itogo = branch.getNameForCard(curt);
                                                if (itogo.Length > 0)
                                                {
                                                    ep.InsertRow(sL + 1 + i);
                                                    ep.SetRangeBold(i + sL, 1, i + sL, 10);
                                                    ep.SetRangeAlignment(i + sL, sC, i + sL, sC, Microsoft.Office.Interop.Excel.Constants.xlCenter);
                                                    ep.SetText(i + sL, sC, itogo);
                                                    if (id_type == 19)
                                                    {
                                                        ep.SetText(i + sL, sC + 1, curSum[3].ToString());
                                                        ep.SetText(i + sL, sC + 2, curSum[2].ToString());
                                                        ep.SetText(i + sL, sC + 3, curSum[1].ToString());
                                                        ep.SetText(i + sL, sC + 4, curSum[0].ToString());
                                                    }
                                                    else
                                                    {
                                                        ep.SetRangeAlignment(i + sL, sC + 1, i + sL, sC + 1, Microsoft.Office.Interop.Excel.Constants.xlCenter);
                                                        ep.SetText(i + sL, sC + 1, "шт.");
                                                        ep.SetText(i + sL, sC + 2, curSum[0].ToString());
                                                    }
                                                    i++;
                                                }
                                                curt = t;
                                                curNum = 0;
                                                cardSum[0] += curSum[0];
                                                cardSum[1] += curSum[1];
                                                cardSum[2] += curSum[2];
                                                cardSum[3] += curSum[3];
                                                if (curt == (int)BaseProductType.PinConvert && id_type == 19)
                                                {
                                                    ep.InsertRow(sL + 1 + i);
                                                    ep.SetText(i + sL, 2, curDate.ToShortDateString());
                                                    ep.SetText(i + sL, sC, "Итого по картам:");
                                                    ep.SetText(i + sL, sC + 1, cardSum[3].ToString());
                                                    ep.SetText(i + sL, sC + 2, cardSum[2].ToString());
                                                    ep.SetText(i + sL, sC + 3, cardSum[1].ToString());
                                                    ep.SetText(i + sL, sC + 4, cardSum[0].ToString());
                                                    i++;
                                                }
                                                curSum[0] = curSum[1] = curSum[2] = curSum[3] = 0;
                                            }
                                        }

                                        ep.InsertRow(sL + 1 + i);

                                        curSum[0] += mp.cnts[0];
                                        curSum[1] += mp.cnts[1];
                                        curSum[2] += mp.cnts[2];
                                        curSum[3] += s;

                                        if (id_type == 19)
                                        {
                                            if (curt != (int)BaseProductType.PinConvert/* < BranchStore.PINCODE*/) ep.SetText(i + sL, 1, (curNum + 1).ToString());
                                            ep.SetText(i + sL, 2, curDate.ToShortDateString());
                                        }

                                        ep.SetText(i + sL, sC, mp.Name);
                                        if (id_type == 19)
                                        {
                                            ep.SetText(i + sL, sC + 1, s.ToString());
                                            ep.SetText(i + sL, sC + 2, mp.cnts[2].ToString());
                                            ep.SetText(i + sL, sC + 3, mp.cnts[1].ToString());
                                            ep.SetText(i + sL, sC + 4, mp.cnts[0].ToString());
                                            ep.SetText(i + sL, sC + 5, fio);
                                        }
                                        else
                                        {
                                            ep.SetRangeAlignment(i + sL, sC + 1, i + sL, sC + 1, Microsoft.Office.Interop.Excel.Constants.xlCenter);
                                            ep.SetText(i + sL, sC + 1, "шт.");
                                            ep.SetText(i + sL, sC + 2, mp.cnts[0].ToString());
                                        }


                                        i++;
                                        curNum++;
                                    }
                                }


                                ep.SetRangeBorders(sL, 1, sL + i - 1, (id_type == 19) ? 10 : 3);
                                if (id_type != 19) ep.SetRangeXlMedium(sL, 1, sL + i - 1, 3);


                                //второй лист
                                if (id_type == 19)
                                {
                                    ep.SetWorkSheet(2);
                                    ep.SetText(4, 2, "Дата: " + curDate.ToString("dd.MM.yyyy"));
                                    ep.SetText(5, 2, branch.department);
                                    i = 0;
                                    curt = 0;
                                    curSum[0] = 0;
                                    branch.countVisaCard[4] = branch.countMasterCard[4] = branch.countNFCCard[4] = branch.countServiceCard[4] = branch.countPinConvert[4] = branch.countMirCard[4] = 0;
                                    foreach (MyProd mp in al)
                                    {
                                        if (mp.cnts[4] > 0)
                                        {

                                            int t = branch.addCount(mp.Type, mp.Name, 4, mp.cnts[4]);

                                            if (curt == 0) curt = t;
                                            else
                                            {
                                                if (curt != t)
                                                {
                                                    string itogo = branch.getNameForCard(curt);
                                                    if (itogo.Length > 0)
                                                    {
                                                        ep.InsertRow(8 + i);
                                                        //if (curt == 1) ep.SetRangeColorIndex(i + 7, 2, i + 9, 3, 3);
                                                        //if (curt == 2) ep.SetRangeColorIndex(i + 7, 2, i + 9, 3, 41);
                                                        //if (curt == 3) ep.SetRangeColorIndex(i + 7, 2, i + 9, 3, 10);
                                                        //if (curt == 4) ep.SetRangeColorIndex(i + 7, 2, i + 9, 3, 15);
                                                        ep.SetRangeBold(i + 7, 2, i + 7, 3);
                                                        ep.SetRangeAlignment(i + 7, 2, i + 7, 3, Microsoft.Office.Interop.Excel.Constants.xlCenter);
                                                        ep.SetText(i + 7, 2, itogo);
                                                        ep.SetText(i + 7, 3, curSum[4].ToString());
                                                        i++;
                                                    }
                                                    curt = t;
                                                    curSum[4] = 0;
                                                }
                                            }

                                            ep.InsertRow(8 + i);

                                            curSum[4] += mp.cnts[4];
                                            ep.SetText(i + 7, 2, mp.Name);
                                            ep.SetText(i + 7, 3, mp.cnts[4].ToString());
                                            if (t == 5)
                                            {
                                                ep.SetRangeBold(i + 7, 2, i + 7, 3);
                                                ep.SetRangeAlignment(i + 7, 2, i + 7, 3, Microsoft.Office.Interop.Excel.Constants.xlCenter);
                                                //ep.SetRangeColorIndex(i + 7, 2, i + 7, 3, 27);
                                            }
                                            i++;

                                        }
                                    }
                                    ep.SetRangeBorders(7, 2, 7 + i - 1, 3);
                                }

                            }
                        }

                        #endregion
                        #region Отчет для Казанского филиала (консолидированный отчет)
                        if (id_type == 20)
                        {
                            ep.SetWorkSheetName(2, "NULL");
                            
                            DateTime curDate = DatePickerOne.SelectedDate.Date;
                            DateTime toDate = DateTime.Now.Date;

                            OperationDay op = new OperationDay();
                            op.read(current_branch);

                            DateTime curDateS = op.getDateTimeStart(curDate);
                            DateTime curDateE = op.getDateTimeEnd(curDate);
                            DateTime toDateS = op.getDateTimeStart(toDate);
                            DateTime toDateE = op.getDateTimeEnd(toDate);


                            SqlCommand comm = Database2.Conn.CreateCommand();

                            ds.Clear();
                            comm.Parameters.Clear();

                            comm.CommandText = "select * from Branchs where id_parent=@id_parent or id=@id_parent order by ident_dep";
                            //comm.CommandText = "select * from Branchs where id=176 or id=106 order by ident_dep";
                            comm.Parameters.Add("@id_parent", SqlDbType.Int).Value = 106;
                            Database2.ExecuteCommand(comm, ref ds, null);

                            if (ds.Tables.Count > 0 && ds.Tables[0].Rows.Count > 0)
                            {//--------------------------------
                                ArrayList branches = new ArrayList();

                                for (int b = 0; b < ds.Tables[0].Rows.Count; b++)
                                {
                                    BranchStore branch = new BranchStore(ds.Tables[0].Rows[b]["id"], ds.Tables[0].Rows[b]["ident_dep"], ds.Tables[0].Rows[b]["department"]);
                                    branches.Add(branch);
                                }

                                DataSet dsProduct = new DataSet();
                                Database2.ExecuteQuery("select id_prb, id_prod, prod_name, id_type from V_ProductsBanks_T where id_type in(1,2) order by id_type, id_sort", ref dsProduct, null);
                                int page = 0;
                                for (int b = 0; b < branches.Count; b++)
                                {

                                    BranchStore branch = (BranchStore)branches[b];

                                    op.read(branch.id);

                                    curDateS = op.getDateTimeStart(curDate);
                                    curDateE = op.getDateTimeEnd(curDate);
                                    toDateS = op.getDateTimeStart(toDate);
                                    toDateE = op.getDateTimeEnd(toDate);

                                    ds.Clear();
                                    comm.Parameters.Clear();
                                    ArrayList al = new ArrayList();

                                    if (dsProduct.Tables.Count > 0)
                                    {
                                        String sql = "select id_prb, count(*) as count_prb from Cards  where id_stat in (4,6,11) and id_branchCurrent=" + branch.id.ToString() + " group by id_prb";
                                        if (branch_main_filial < 1)
                                            sql = "select id_prb, count(*) as count_prb from Cards  where id_stat in (4,6,11) and id_branchCard=" + branch.id.ToString() + " group by id_prb";
                                        Database2.ExecuteQuery(sql, ref ds, null);
                                        for (i = 0; i < dsProduct.Tables[0].Rows.Count; i++)
                                        {
                                            MyProd mp = new MyProd(Convert.ToInt32(dsProduct.Tables[0].Rows[i]["id_prb"]),
                                                Convert.ToString(dsProduct.Tables[0].Rows[i]["prod_name"]), "", Convert.ToInt32(dsProduct.Tables[0].Rows[i]["id_type"]));
                                            for (int n = 0; mp.Type != 2 && ds.Tables.Count > 0 && n < ds.Tables[0].Rows.Count; n++)
                                            {
                                                if (mp.ID == Convert.ToInt32(ds.Tables[0].Rows[n]["id_prb"]))
                                                {
                                                    mp.cnts[0] = Convert.ToInt32(ds.Tables[0].Rows[n]["count_prb"]);
                                                    break;
                                                }
                                            }
                                            al.Add(mp);
                                            if (mp.Type == 2) // Пин конверты на конец дня
                                            {
                                                //не учитывем карты статуса филиал, которые до этого побывали на экспертизе. они без пина
                                                comm.CommandText = "select count(*) from Cards where id_stat in (4,6) and id_branchCurrent=" + branch.id.ToString() + " and isPin=1 and id not in (select id_card from V_CardsTypeDocs where type=23)";
                                                if (branch_main_filial < 1)
                                                    comm.CommandText = "select count(*) from Cards where id_stat in (4,6) and id_branchCard=" + branch.id.ToString() + " and isPin=1 and id not in (select id_card from V_CardsTypeDocs where type=23)";
                                                object obj = comm.ExecuteScalar();
                                                mp.cnts[0] = (obj == null || obj == DBNull.Value) ? 0 : Convert.ToInt32(obj);
                                            }
                                            branch.addCount(mp.Type, mp.Name, 0, mp.cnts[0]);

                                        }
                                    }

                                    if (curDateS < toDateS)
                                    {
                                        ds.Clear();
                                        comm.Parameters.Clear();
                                        //корректировка выдачи
                                        comm.CommandText = "select dbo.Cards.id_prb,count(*) as count_prb" +
                                                       " from dbo.Branchs RIGHT OUTER JOIN" +
                                                       " dbo.StorageDocs ON dbo.Branchs.id = dbo.StorageDocs.id_branch LEFT OUTER JOIN" +
                                                       " dbo.Cards RIGHT OUTER JOIN" +
                                                       " dbo.Cards_StorageDocs ON dbo.Cards.id = dbo.Cards_StorageDocs.id_card ON dbo.StorageDocs.id = dbo.Cards_StorageDocs.id_doc" +
                                                       " where date_time>=@dateS and date_time<@dateP and" +
                                                       " (id_branch=@id_branch and (type=7 or type=9 or type=11 or type=13 or type=24 or type=21) or (id_act=@id_branch and (type=19 or type=5)))" +
                                                       " and priz_gen=1 group by dbo.Cards.id_prb";
                                        comm.Parameters.Add("@dateS", SqlDbType.DateTime).Value = curDateE;
                                        comm.Parameters.Add("@dateP", SqlDbType.DateTime).Value = toDateE;
                                        comm.Parameters.Add("@id_branch", SqlDbType.Int).Value = branch.id;
                                        Database2.ExecuteCommand(comm, ref ds, null);
                                        foreach (MyProd mp in al)
                                        {
                                            int countMinus = 0;
                                            for (int n = 0; mp.Type != 2 && ds.Tables.Count > 0 && n < ds.Tables[0].Rows.Count; n++)
                                            {
                                                if (mp.ID == Convert.ToInt32(ds.Tables[0].Rows[n]["id_prb"]))
                                                {
                                                    countMinus = Convert.ToInt32(ds.Tables[0].Rows[n]["count_prb"]);
                                                    break;
                                                }
                                            }
                                            if (mp.Type == 2)
                                            {
                                                //отпраку на экспертизу не учитываем. она идет без пина
                                                comm.Parameters.Clear();
                                                //comm.CommandText = "select count(*) from cards where id in (select id_card from cards_storagedocs where id_doc in (select id from storagedocs where date_time>=@dateS and date_time<@dateP and ((id_branch=@id_branch and (type=7 or type=9 or type=11 or type=13)) or (id_act=@id_branch and type=19)) and priz_gen=1)) and isPin=1";
                                                comm.CommandText = "select count(id_card) from cards_storagedocs cs join cards c on c.id=cs.id_card where id_doc in (select id from storagedocs where date_time>=@dateS and date_time<@dateP and ((id_branch=@id_branch and (type=7 or type=9 or type=11 or type=13 or type=24)) or (id_act=@id_branch and (type=19 or type=5))) and priz_gen=1) and isPin=1";
                                                comm.Parameters.Add("@dateS", SqlDbType.DateTime).Value = curDateE;
                                                comm.Parameters.Add("@dateP", SqlDbType.DateTime).Value = toDateE;
                                                comm.Parameters.Add("@id_branch", SqlDbType.Int).Value = branch.id;
                                                object obj = comm.ExecuteScalar();
                                                countMinus = (obj == null || obj == DBNull.Value) ? 0 : Convert.ToInt32(obj);
                                            }
                                            branch.addCount(mp.Type, mp.Name, 0, countMinus);
                                            mp.cnts[0] += countMinus;
                                        }
                                        //корректировка получения
                                        ds.Clear();
                                        comm.Parameters.Clear();
                                        comm.CommandText = "select id_prb,count(*) as count_prb from V_Cards_StorageDocs1 where date_time>=@dateS and date_time<@dateP and id_branch=@id_branch and (type=6 or type=10 or type=28 or type=20) and priz_gen=1 group by id_prb";
                                        comm.Parameters.Add("@dateS", SqlDbType.DateTime).Value = curDateE;
                                        comm.Parameters.Add("@dateP", SqlDbType.DateTime).Value = toDateE;
                                        comm.Parameters.Add("@id_branch", SqlDbType.Int).Value = branch.id;
                                        Database2.ExecuteCommand(comm, ref ds, null);
                                        foreach (MyProd mp in al)
                                        {

                                            int countPlus = 0;
                                            for (int n = 0; mp.Type != 2 && ds.Tables.Count > 0 && n < ds.Tables[0].Rows.Count; n++)
                                            {
                                                if (mp.ID == Convert.ToInt32(ds.Tables[0].Rows[n]["id_prb"]))
                                                {
                                                    countPlus = Convert.ToInt32(ds.Tables[0].Rows[n]["count_prb"]);
                                                    break;
                                                }
                                            }
                                            if (mp.Type == 2) // Пин конверты получены
                                            {
                                                comm.Parameters.Clear();
                                                //comm.CommandText = "select count(*) from cards where id in (select id_card from cards_storagedocs where id_doc in (select id from storagedocs where date_time>=@dateS and date_time<@dateP and (type=6 or type=10) and priz_gen=1 and id_branch=@id_branch)) and isPin=1";
                                                comm.CommandText = "select count(id_card) from cards_storagedocs cs join cards c on c.id=cs.id_card where id_doc in (select id from storagedocs where date_time>=@dateS and date_time<@dateP and (type=6 or type=10 or type=28) and priz_gen=1 and id_branch=@id_branch) and isPin=1 and c.id not in (select id_card from V_CardsTypeDocs where type=23)";
                                                comm.Parameters.Add("@dateS", SqlDbType.DateTime).Value = curDateE;
                                                comm.Parameters.Add("@dateP", SqlDbType.DateTime).Value = toDateE;
                                                comm.Parameters.Add("@id_branch", SqlDbType.Int).Value = branch.id;
                                                object obj = comm.ExecuteScalar();
                                                countPlus = (obj == null || obj == DBNull.Value) ? 0 : Convert.ToInt32(obj);
                                            }
                                            branch.addCount(mp.Type, mp.Name, 0, -countPlus);
                                            mp.cnts[0] -= countPlus;
                                        }
                                    }
                                    ds.Clear();
                                    comm.Parameters.Clear();
                                    //выдача
                                    comm.CommandText = "select dbo.Cards.id_prb,count(*) as count_prb" +
                                                   " from dbo.Branchs RIGHT OUTER JOIN" +
                                                   " dbo.StorageDocs ON dbo.Branchs.id = dbo.StorageDocs.id_branch LEFT OUTER JOIN" +
                                                   " dbo.Cards RIGHT OUTER JOIN" +
                                                   " dbo.Cards_StorageDocs ON dbo.Cards.id = dbo.Cards_StorageDocs.id_card ON dbo.StorageDocs.id = dbo.Cards_StorageDocs.id_doc" +
                                                   " where date_time>=@dateS and date_time<@dateE and" +
                                                   " (id_branch=@id_branch and (type=7 or type=9 or type=11 or type=13  or type=24 or type=21) or (id_act=@id_branch and (type=19 or type=5)))" +
                                                   " and priz_gen=1 group by dbo.Cards.id_prb";
                                    comm.Parameters.Add("@dateS", SqlDbType.DateTime).Value = curDateS;
                                    comm.Parameters.Add("@dateE", SqlDbType.DateTime).Value = curDateE;
                                    comm.Parameters.Add("@id_branch", SqlDbType.Int).Value = branch.id;
                                    Database2.ExecuteCommand(comm, ref ds, null);
                                    foreach (MyProd mp in al)
                                    {

                                        for (int n = 0; mp.Type != 2 && ds.Tables.Count > 0 && n < ds.Tables[0].Rows.Count; n++)
                                        {
                                            if (mp.ID == Convert.ToInt32(ds.Tables[0].Rows[n]["id_prb"]))
                                            {
                                                mp.cnts[1] = Convert.ToInt32(ds.Tables[0].Rows[n]["count_prb"]);
                                                break;
                                            }
                                        }
                                        if (mp.Type == 2) // Пин конверты выданы
                                        {
                                            //отправку на экспертизу не учитываем. она без пина
                                            comm.Parameters.Clear();
                                            //comm.CommandText = "select count(*) from cards where id in (select id_card from cards_storagedocs where id_doc in (select id from storagedocs where date_time>=@dateS and date_time<@dateE and ((id_branch=@id_branch and (type=7 or type=9 or type=11 or type=13)) or (id_act=@id_branch and type=19)) and priz_gen=1)) and isPin=1";
                                            comm.CommandText = "select count(id_card) from cards_storagedocs cs join cards c on c.id=cs.id_card where id_doc in (select id from storagedocs where date_time>=@dateS and date_time<@dateE and ((id_branch=@id_branch and (type=7 or type=9 or type=11 or type=13 or type=24)) or (id_act=@id_branch and (type=19 or type=5))) and priz_gen=1) and isPin=1";
                                            comm.Parameters.Add("@dateS", SqlDbType.DateTime).Value = curDateS;
                                            comm.Parameters.Add("@dateE", SqlDbType.DateTime).Value = curDateE;
                                            comm.Parameters.Add("@id_branch", SqlDbType.Int).Value = branch.id;
                                            object obj = comm.ExecuteScalar();
                                            mp.cnts[1] = (obj == null || obj == DBNull.Value) ? 0 : Convert.ToInt32(obj);
                                        }
                                        branch.addCount(mp.Type, mp.Name, 1, mp.cnts[1]);
                                    }
                                    //получение
                                    ds.Clear();
                                    comm.Parameters.Clear();
                                    comm.CommandText = "select id_prb,count(*) as count_prb from V_Cards_StorageDocs1 where date_time>=@dateS and date_time<@dateE and id_branch=@id_branch and (type=6 or type=10 or type=28 or type=20) and priz_gen=1 group by id_prb";
                                    comm.Parameters.Add("@dateS", SqlDbType.DateTime).Value = curDateS;
                                    comm.Parameters.Add("@dateE", SqlDbType.DateTime).Value = curDateE;
                                    comm.Parameters.Add("@id_branch", SqlDbType.Int).Value = branch.id;
                                    Database2.ExecuteCommand(comm, ref ds, null);
                                    foreach (MyProd mp in al)
                                    {

                                        for (int n = 0; mp.Type != 2 && ds.Tables.Count > 0 && n < ds.Tables[0].Rows.Count; n++)
                                        {
                                            if (mp.ID == Convert.ToInt32(ds.Tables[0].Rows[n]["id_prb"]))
                                            {
                                                mp.cnts[2] = Convert.ToInt32(ds.Tables[0].Rows[n]["count_prb"]);
                                                break;
                                            }
                                        }
                                        if (mp.Type == 2) // Пин конверты получены
                                        {

                                            comm.Parameters.Clear();
                                            //comm.CommandText = "select count(*) from cards where id in (select id_card from cards_storagedocs where id_doc in (select id from storagedocs where date_time>=@dateS and date_time<@dateE and (type=6 or type=10) and priz_gen=1 and id_branch=@id_branch)) and isPin=1";
                                            comm.CommandText = "select count(id_card) from cards_storagedocs cs join cards c on c.id=cs.id_card where id_doc in (select id from storagedocs where date_time>=@dateS and date_time<@dateE and (type=6 or type=10 or type=28) and priz_gen=1 and id_branch=@id_branch) and isPin=1 and c.id not in (select id_card from V_CardsTypeDocs where type=23)";
                                            comm.Parameters.Add("@dateS", SqlDbType.DateTime).Value = curDateS;
                                            comm.Parameters.Add("@dateE", SqlDbType.DateTime).Value = curDateE;
                                            comm.Parameters.Add("@id_branch", SqlDbType.Int).Value = branch.id;
                                            object obj = comm.ExecuteScalar();
                                            mp.cnts[2] = (obj == null || obj == DBNull.Value) ? 0 : Convert.ToInt32(obj);
                                        }
                                        branch.addCount(mp.Type, mp.Name, 2, mp.cnts[2]);
                                    }

                                    i = 0;

                                    Boolean isPage = false;
                                    foreach (MyProd mp in al)
                                    {
                                        if (mp.cnts[1] + mp.cnts[2] > 0)
                                        {
                                            isPage = true;
                                            break;
                                        }
                                    }

                                    if (isPage)
                                    {

                                        ep.SetWorkSheet(2 + page);
                                        //if (b < branches.Count - 1)
                                        //{
                                        ep.AddWorkSheet(2 + page);
                                        //}
                                        ep.SetWorkSheetName(2 + page, branch.ident_dep);

                                        page++;

                                        ep.SetText(5, 1, branch.department);
                                    }

                                    string fio = "";// sc.UserFIO(User.Identity.Name);
                                    int curt = 0;
                                    int curNum = 0;
                                    int[] curSum = new int[4];
                                    int[] cardSum = new int[4];
                                    foreach (MyProd mp in al)
                                    {
                                        if (mp.cnts[0] + mp.cnts[1] + mp.cnts[2] > 0)
                                        {

                                            int s = mp.cnts[0] + mp.cnts[1] - mp.cnts[2];
                                            s = (s > 0) ? s : 0;

                                            int t = branch.addCount(mp.Type, mp.Name, 3, s);

                                            if (!isPage) continue;

                                            if (curt == 0) curt = t;
                                            else
                                            {
                                                if (curt != t)
                                                {
                                                    string itogo = branch.getNameForCard(curt);
                                                    if (itogo.Length > 0)
                                                    {
                                                        ep.InsertRow(9 + i);
                                                        ep.SetRangeBold(i + 8, 1, i + 8, 10);
                                                        ep.SetRangeAlignment(i + 8, 3, i + 8, 3, Microsoft.Office.Interop.Excel.Constants.xlCenter);
                                                        ep.SetText(i + 8, 3, itogo);
                                                        ep.SetText(i + 8, 4, curSum[3].ToString());
                                                        ep.SetText(i + 8, 5, curSum[2].ToString());
                                                        ep.SetText(i + 8, 6, curSum[1].ToString());
                                                        ep.SetText(i + 8, 7, curSum[0].ToString());
                                                        i++;
                                                    }
                                                    curt = t;
                                                    curNum = 0;
                                                    cardSum[0] += curSum[0];
                                                    cardSum[1] += curSum[1];
                                                    cardSum[2] += curSum[2];
                                                    cardSum[3] += curSum[3];
                                                    if (curt == (int)BaseProductType.PinConvert)
                                                    {
                                                        ep.InsertRow(9 + i);
                                                        ep.SetText(i + 8, 2, curDate.ToShortDateString());
                                                        ep.SetText(i + 8, 3, "Итого по картам:");
                                                        ep.SetText(i + 8, 4, cardSum[3].ToString());
                                                        ep.SetText(i + 8, 5, cardSum[2].ToString());
                                                        ep.SetText(i + 8, 6, cardSum[1].ToString());
                                                        ep.SetText(i + 8, 7, cardSum[0].ToString());
                                                        i++;
                                                    }
                                                    curSum[0] = curSum[1] = curSum[2] = curSum[3] = 0;
                                                }
                                            }

                                            ep.InsertRow(9 + i);

                                            curSum[0] += mp.cnts[0];
                                            curSum[1] += mp.cnts[1];
                                            curSum[2] += mp.cnts[2];
                                            curSum[3] += s;
                                            if (curt != (int)BaseProductType.PinConvert) ep.SetText(i + 8, 1, (curNum + 1).ToString());
                                            ep.SetText(i + 8, 2, curDate.ToShortDateString());
                                            ep.SetText(i + 8, 3, mp.Name);

                                            ep.SetText(i + 8, 4, s.ToString());
                                            ep.SetText(i + 8, 5, mp.cnts[2].ToString());
                                            ep.SetText(i + 8, 6, mp.cnts[1].ToString());
                                            ep.SetText(i + 8, 7, mp.cnts[0].ToString());
                                            ep.SetText(i + 8, 8, fio);


                                            i++;
                                            curNum++;
                                        }
                                    }
                                    if (isPage) ep.SetRangeBorders(8, 1, 8 + i - 1, 10);


                                    ep.SetWorkSheet(1);

                                    ep.SetText(3 + b, 1, branch.department);
                                    ep.SetText(3 + b, 2, branch.ident_dep);
                                    ep.SetText(3 + b, 3, branch.countMasterCard[3].ToString());
                                    ep.SetText(3 + b, 4, branch.countMasterCard[2].ToString());
                                    ep.SetText(3 + b, 5, branch.countMasterCard[1].ToString());
                                    ep.SetText(3 + b, 6, branch.countMasterCard[0].ToString());
                                    ep.SetText(3 + b, 7, branch.countVisaCard[3].ToString());
                                    ep.SetText(3 + b, 8, branch.countVisaCard[2].ToString());
                                    ep.SetText(3 + b, 9, branch.countVisaCard[1].ToString());
                                    ep.SetText(3 + b, 10, branch.countVisaCard[0].ToString());
                                    //ep.SetText(3 + b, 11, branch.countNFCCard[3].ToString());
                                    //ep.SetText(3 + b, 12, branch.countNFCCard[2].ToString());
                                    //ep.SetText(3 + b, 13, branch.countNFCCard[1].ToString());
                                    //ep.SetText(3 + b, 14, branch.countNFCCard[0].ToString());
                                    ep.SetText(3 + b, 11, branch.countMirCard[3].ToString());
                                    ep.SetText(3 + b, 12, branch.countMirCard[2].ToString());
                                    ep.SetText(3 + b, 13, branch.countMirCard[1].ToString());
                                    ep.SetText(3 + b, 14, branch.countMirCard[0].ToString());
                                    //ep.SetText(3 + b, 15, branch.countPinConvert[3].ToString());
                                    //ep.SetText(3 + b, 16, branch.countPinConvert[2].ToString());
                                    //ep.SetText(3 + b, 17, branch.countPinConvert[1].ToString());
                                    //ep.SetText(3 + b, 18, branch.countPinConvert[0].ToString());
                                    ep.SetText(3 + b, 15, branch.countNFCCard[3].ToString());
                                    ep.SetText(3 + b, 16, branch.countNFCCard[2].ToString());
                                    ep.SetText(3 + b, 17, branch.countNFCCard[1].ToString());
                                    ep.SetText(3 + b, 18, branch.countNFCCard[0].ToString());
                                    ep.SetText(3 + b, 19, branch.countServiceCard[3].ToString());
                                    ep.SetText(3 + b, 20, branch.countServiceCard[2].ToString());
                                    ep.SetText(3 + b, 21, branch.countServiceCard[1].ToString());
                                    ep.SetText(3 + b, 22, branch.countServiceCard[0].ToString());
                                    //ep.SetText(3 + b, 23, branch.countMirCard[3].ToString());
                                    //ep.SetText(3 + b, 24, branch.countMirCard[2].ToString());
                                    //ep.SetText(3 + b, 25, branch.countMirCard[1].ToString());
                                    //ep.SetText(3 + b, 26, branch.countMirCard[0].ToString());
                                    ep.SetText(3 + b, 23, branch.countPinConvert[3].ToString());
                                    ep.SetText(3 + b, 24, branch.countPinConvert[2].ToString());
                                    ep.SetText(3 + b, 25, branch.countPinConvert[1].ToString());
                                    ep.SetText(3 + b, 26, branch.countPinConvert[0].ToString());
                                }

                                string wname = ep.GetWorkSheetName(2 + page);
                                if (wname.Contains("NULL") == true)
                                {
                                    ep.DelWorkSheet(2 + page);
                                }
                                ep.SetWorkSheet(1);
                                ep.SetRangeBorders(3, 1, 3 + branches.Count - 1, 26);



                            }//--------------------------------
                        }
                        #endregion
                        #region Отчет для списания с подотчета МОЛ
                        if (id_type == 43)
                        {
                            try
                            {
                                BranchStore brstore = new BranchStore(0, "", "");
                                DateTime dtStart = DatePickerStart.SelectedDate.Date;
                                DateTime dtEnd = DatePickerEnd.SelectedDate.Date;
                                DateTime dtCur = DateTime.Now.Date;
                                if (dtEnd < dtStart) throw new Exception("Неверно задан диапазон дат");
                                SqlCommand comm = Database2.Conn.CreateCommand();
                                ds.Clear();
                                comm.Parameters.Clear();
                                //проверяем, что есть неподтвержденные документы на персонализацию или отправку в филиал
                                comm.CommandText = "select number_doc from StorageDocs where date_doc>=@dtStart and date_doc<=@dtEnd and priz_gen=0 and (type=5 or type=8)";
                                comm.Parameters.Add("@dtStart", SqlDbType.DateTime).Value = dtStart.Date;
                                comm.Parameters.Add("@dtEnd", SqlDbType.DateTime).Value = dtEnd.Date;
                                Database2.ExecuteCommand(comm, ref ds, null);
                                if (ds.Tables.Count > 0 && ds.Tables[0].Rows.Count > 0)
                                {
                                    string str = "";
                                    for (i = 0; i < ds.Tables[0].Rows.Count; i++)
                                        str += ", " + ds.Tables[0].Rows[i][0].ToString();
                                    str = str.Substring(2);
                                    throw new Exception(String.Format("Создание отчета невозможно. За данное число есть неподтвержденные документы ({0})", str));
                                }
                                ds.Clear();
                                comm.Parameters.Clear();
                                comm.CommandText = "select id_prb, prod_name, bank_name, id_type from V_ProductsBanks_T order by id_sort";
                                res = (string)Database2.ExecuteCommand(comm, ref ds, null);
                                ArrayList al = new ArrayList();
                                foreach (DataRow dr in ds.Tables[0].Rows)
                                    al.Add(new MyProd(Convert.ToInt32(dr["id_prb"]), dr["prod_name"].ToString(), dr["bank_name"].ToString(), Convert.ToInt32(dr["id_type"])));
                                //выдача на персонализацию
                                comm.Parameters.Clear();
                                ds.Clear();
                                comm.Parameters.Add("@type", SqlDbType.Int).Value = (int)TypeDoc.PersoCard;
                                comm.Parameters.Add("@dtStart", SqlDbType.DateTime).Value = dtStart.Date;
                                comm.Parameters.Add("@dtEnd", SqlDbType.DateTime).Value = dtEnd.Date;
                                comm.CommandText = "select id_prb,sum(cnt_brak) as cntbrak, sum(cnt_perso) as cntperso from V_Rep_Moving where priz_gen=1 and type=@type and (id_type=1 or id_type=2) and date_doc>=@dtStart and date_doc<=@dtEnd group by id_prb";
                                ds.Clear();
                                Database2.ExecuteCommand(comm, ref ds, null);
                                foreach (MyProd mp in al)
                                {
                                    for (i = 0; ds.Tables.Count > 0 && i < ds.Tables[0].Rows.Count; i++)
                                    {
                                        if (mp.ID == Convert.ToInt32(ds.Tables[0].Rows[i]["id_prb"]))
                                        {
                                            int br = (ds.Tables[0].Rows[i]["cntbrak"] == DBNull.Value) ? 0 : Convert.ToInt32(ds.Tables[0].Rows[i]["cntbrak"]);
                                            int pr = (ds.Tables[0].Rows[i]["cntperso"] == DBNull.Value) ? 0 : Convert.ToInt32(ds.Tables[0].Rows[i]["cntperso"]);
                                            mp.cnts[0] = br + pr;
                                            brstore.addCount(mp.Type, mp.Name, 0, mp.cnts[0]);
                                            break;
                                        }
                                    }
                                }
                                
                                //------------------------------------------------------
                                //выдача на рекламу и тесты
                                comm.Parameters.Clear();
                                ds.Clear();
                                comm.Parameters.Add("@type", SqlDbType.Int).Value = (int)TypeDoc.Reklama;
                                comm.Parameters.Add("@dtStart", SqlDbType.DateTime).Value = dtStart.Date;
                                comm.Parameters.Add("@dtEnd", SqlDbType.DateTime).Value = dtEnd.Date;
                                comm.CommandText = "select id_prb,sum(cnt_new) as cntnew from V_Rep_Moving where priz_gen=1 and type=@type and (id_type=1 or id_type=2) and date_doc>=@dtStart and date_doc<=@dtEnd group by id_prb";
                                ds.Clear();
                                Database2.ExecuteCommand(comm, ref ds, null);
                                foreach (MyProd mp in al)
                                {
                                    for (i = 0; ds.Tables.Count > 0 && i < ds.Tables[0].Rows.Count; i++)
                                    {
                                        if (mp.ID == Convert.ToInt32(ds.Tables[0].Rows[i]["id_prb"]))
                                        {
                                            int nw = (ds.Tables[0].Rows[i]["cntnew"] == DBNull.Value) ? 0 : Convert.ToInt32(ds.Tables[0].Rows[i]["cntnew"]);
                                            if (nw != 0)
                                            {
                                                mp.cnts[0] += nw;
                                                brstore.addCount(mp.Type, mp.Name, 0, nw);
                                                break;
                                            }
                                        }
                                    }
                                }
                                //------------------------------------------------------

                                ds.Clear();
                                Database2.ExecuteQuery("select name, cnt_new from V_Storage where id_type=1 order by id_sort", ref ds, null);
                                for (i = 0; i < ds.Tables[0].Rows.Count; i++)
                                {
                                    brstore.addCount(1, ds.Tables[0].Rows[i]["name"].ToString(), 1, Convert.ToInt32(ds.Tables[0].Rows[i]["cnt_new"]));
                                }
                                ds.Clear();
                                Database2.ExecuteQuery("select name, cnt_new from V_Storage where id_type=2 order by id_sort", ref ds, null);
                                for (i = 0; i < ds.Tables[0].Rows.Count; i++)
                                {
                                    brstore.addCount(2, ds.Tables[0].Rows[i]["name"].ToString(), 1, Convert.ToInt32(ds.Tables[0].Rows[i]["cnt_new"]));
                                }
                                int itogo1 = brstore.countMasterCard[0] + brstore.countVisaCard[0] + /*brstore.countNFCCard[0]*/ brstore.countServiceCard[0] + brstore.countMirCard[0];
                                int itogo2 = brstore.countMasterCard[1] + brstore.countVisaCard[1] + /*brstore.countNFCCard[1]*/ brstore.countServiceCard[1] + brstore.countMirCard[1];

                                String strH = ep.GetCell(19, 4);
                                strH += String.Format(" c {0:dd.MM.yyyy} по {1:dd.MM.yyyy}", dtStart, dtEnd);
                                ep.SetText(19, 4, strH);
                                strH = ep.GetCell(19, 5);
                                strH += String.Format(" {0:dd.MM.yyyy}", dtCur);
                                ep.SetText(19, 5, strH);
                                ep.SetText(15, 2, String.Format(" от {0:dd.MM.yyyy}", dtCur));


                                ep.SetText(21, 4, (brstore.countMasterCard[0] > 0) ? brstore.countMasterCard[0].ToString() : "");
                                ep.SetText(21, 5, (brstore.countMasterCard[1] > 0) ? brstore.countMasterCard[1].ToString() : "");
                                ep.SetText(22, 4, (brstore.countVisaCard[0] > 0) ? brstore.countVisaCard[0].ToString() : "");
                                ep.SetText(22, 5, (brstore.countVisaCard[1] > 0) ? brstore.countVisaCard[1].ToString() : "");
                                //ep.SetText(23, 4, (brstore.countNFCCard[0] > 0) ? brstore.countNFCCard[0].ToString() : "");
                                //ep.SetText(23, 5, (brstore.countNFCCard[1] > 0) ? brstore.countNFCCard[1].ToString() : "");
                                
                                ep.SetText(23, 4, (brstore.countServiceCard[0] > 0) ? brstore.countServiceCard[0].ToString() : "");
                                ep.SetText(23, 5, (brstore.countServiceCard[1] > 0) ? brstore.countServiceCard[1].ToString() : "");

                                ep.SetText(24, 4, (brstore.countMirCard[0] > 0) ? brstore.countMirCard[0].ToString() : "");
                                ep.SetText(24, 5, (brstore.countMirCard[1] > 0) ? brstore.countMirCard[1].ToString() : "");
                                
                                ep.SetText(25, 4, (brstore.countPinConvert[0] > 0) ? brstore.countPinConvert[0].ToString() : "");
                                ep.SetText(25, 5, (brstore.countPinConvert[1] > 0) ? brstore.countPinConvert[1].ToString() : "");
                                ep.SetText(26, 4, (itogo1 > 0) ? itogo1.ToString() : "");
                                ep.SetText(26, 5, (itogo2 > 0) ? itogo2.ToString() : "");

                                ep.SetText(30, 2, sc.UserFIO(User.Identity.Name));

                            }
                            catch (Exception e1)
                            {
                                ep.Close();
                                GC.Collect();
                                System.Threading.Thread.CurrentThread.CurrentCulture = oldCI;
                                ClientScript.RegisterClientScriptBlock(GetType(), "err43", "<script type='text/javascript'>$(document).ready(function(){ ShowError('" + e1.Message + "');});</script>");
                                return;
                            }
                        }
                        #endregion
                        #region Отчет по карте жителя
                        if (id_type == 44)
                        {
                            try
                            {
                                GenReportKgrt(ep);          
                            }
                            catch (Exception e44)
                            {
                                ep.Close();
                                GC.Collect();
                                System.Threading.Thread.CurrentThread.CurrentCulture = oldCI;
                                ClientScript.RegisterClientScriptBlock(GetType(), "err44", "<script type='text/javascript'>$(document).ready(function(){ ShowError('" + e44.Message + "');});</script>");
                                return;
                            }
                        }
                        #endregion
                        #region отчет выданных карт по книге 124
                        if (id_type == 45)
                        {
                            try
                            {
                                using (SqlConnection conn = new SqlConnection(FuncClass.ConnectionString))
                                {
                                    ServiceClass sc = new ServiceClass();
                                    conn.Open();
                                    using (SqlCommand comm = conn.CreateCommand())
                                    {
                                        comm.CommandText = "select distinct id_person from Cards where id_branchCard=@branch and dateClient=@dateClient and id_person is not NULL";
                                        comm.Parameters.Add("@branch", SqlDbType.Int).Value = current_branch;
                                        comm.Parameters.Add("@dateClient", SqlDbType.DateTime).Value = DatePickerOne.SelectedDate.Date;
                                        ArrayList users = new ArrayList();
                                        SqlDataReader dr = comm.ExecuteReader();
                                        while (dr.Read())
                                            users.Add(Convert.ToInt32(dr["id_person"]));
                                        dr.Close();
                                        comm.Parameters.Add("@id_person", SqlDbType.Int);
                                        for (i = 0; i < users.Count - 1; i++)
                                        {
                                            ep.SetWorkSheet(1);
                                            ep.AddWorkSheet();
                                        }
                                        int index = 1;
                                        List<string[]> rows = new List<string[]>();
                                        foreach (int person_id in users)
                                        {
                                            int cntpin = 0, row = 9, cnt = 0;
                                            comm.CommandText = "select v.fio, v.pan, v.prod_name, v.company, v.isPin from V_Cards v inner join Cards on v.id=Cards.id where v.id_branchCard=@branch and v.dateClient=@dateClient and v.id_person=@id_person order by prod_name";
                                            comm.Parameters["@id_person"].Value = person_id;
                                            dr = comm.ExecuteReader();
                                            while (dr.Read())
                                            {
                                                rows.Add(new string[] { dr["fio"].ToString(), dr["pan"].ToString(), dr["prod_name"].ToString(), dr["company"].ToString() });
                                                if (Convert.ToBoolean(dr["isPin"]))
                                                    cntpin++;
                                            }
                                            dr.Close();
                                            ep.SetWorkSheet(index);
                                            cnt = 0;
                                            for (i = 0; i < rows.Count; i++)
                                            {
                                                if (BranchStore.codeFromTypeAndProdName(1, rows[i][2]) == BaseProductType.MasterCard)
                                                {
                                                    cnt++;
                                                    ep.SetText(row, 1, $"{cnt}");
                                                    ep.SetText(row, 2, rows[i][0]);
                                                    ep.SetText(row, 3, rows[i][1]);
                                                    ep.SetText(row, 4, rows[i][2]);
                                                    ep.SetText(row, 5, rows[i][3]);
                                                    row++;
                                                }
                                            }
                                            if (cnt > 0)
                                            {
                                                ep.SetRangeBold(row, 2, row, 3);
                                                ep.SetText(row, 2, "MasterCard (ИТОГО):");
                                                ep.SetText(row, 3, $"{cnt}");
                                                row++;
                                            }
                                            cnt = 0;
                                            for (i = 0; i < rows.Count; i++)
                                            {
                                                if (BranchStore.codeFromTypeAndProdName(1, rows[i][2]) == BaseProductType.VisaCard)
                                                {
                                                    cnt++;
                                                    ep.SetText(row, 1, $"{cnt}");
                                                    ep.SetText(row, 2, rows[i][0]);
                                                    ep.SetText(row, 3, rows[i][1]);
                                                    ep.SetText(row, 4, rows[i][2]);
                                                    ep.SetText(row, 5, rows[i][3]);
                                                    row++;
                                                }
                                            }
                                            if (cnt > 0)
                                            {
                                                ep.SetRangeBold(row, 2, row, 3);
                                                ep.SetText(row, 2, "Visa (ИТОГО):");
                                                ep.SetText(row, 3, $"{cnt}");
                                                row++;
                                            }
                                            cnt = 0;
                                            for (i = 0; i < rows.Count; i++)
                                            {
                                                if (BranchStore.codeFromTypeAndProdName(1, rows[i][2]) == BaseProductType.MirCard)
                                                {
                                                    cnt++;
                                                    ep.SetText(row, 1, $"{cnt}");
                                                    ep.SetText(row, 2, rows[i][0]);
                                                    ep.SetText(row, 3, rows[i][1]);
                                                    ep.SetText(row, 4, rows[i][2]);
                                                    ep.SetText(row, 5, rows[i][3]);
                                                    row++;
                                                }
                                            }
                                            if (cnt > 0)
                                            {
                                                ep.SetRangeBold(row, 2, row, 3);
                                                ep.SetText(row, 2, "Мир (ИТОГО):");
                                                ep.SetText(row, 3, $"{cnt}");
                                                row++;
                                            }
                                            cnt = 0;
                                            for (i = 0; i < rows.Count; i++)
                                            {
                                                if (BranchStore.codeFromTypeAndProdName(1, rows[i][2]) == BaseProductType.NFCCard)
                                                {
                                                    cnt++;
                                                    ep.SetText(row, 1, $"{cnt}");
                                                    ep.SetText(row, 2, rows[i][0]);
                                                    ep.SetText(row, 3, rows[i][1]);
                                                    ep.SetText(row, 4, rows[i][2]);
                                                    ep.SetText(row, 5, rows[i][3]);
                                                    row++;
                                                }
                                            }
                                            if (cnt > 0)
                                            {
                                                ep.SetRangeBold(row, 2, row, 3);
                                                ep.SetText(row, 2, "NFC (ИТОГО):");
                                                ep.SetText(row, 3, $"{cnt}");
                                                row++;
                                            }
                                            if (cntpin > 0)
                                            {
                                                ep.SetRangeBold(row, 2, row, 3);
                                                ep.SetText(row, 2, "Пинконверты:");
                                                ep.SetText(row, 3, $"{cntpin}");
                                                row++;
                                            }
                                            ep.SetRangeBorders(9, 1, row, 5);
                                            ep.SetText(row + 2, 2, "Выдачу карт/пинов подтверждаю");
                                            ep.SetRangeBottom(row + 2, 3, row + 2, 5);
                                            ep.SetText(row + 3, 3, "должность");
                                            ep.SetText(row + 3, 4, "подпись");
                                            ep.SetText(row + 3, 5, "фамилия и инициалы");
                                            string nm = sc.UserFIO(User.Identity.Name);
                                            ep.SetText(row + 2, 3, sc.UserPosition(User.Identity.Name));
                                            ep.SetText(row + 2, 5, nm);
                                            ep.SetText(row + 5, 2, "Сверено");
                                            ep.SetRangeBottom(row + 5, 3, row + 5, 5);
                                            ep.SetText(row + 6, 3, "должность");
                                            ep.SetText(row + 6, 4, "подпись");
                                            ep.SetText(row + 6, 5, "фамилия и инициалы");
                                            ep.SetText(5, 3, $"{DatePickerOne.SelectedDate.Date:dd.MM.yyyy}");
                                            ep.SetWorkSheetName(index, nm);
                                            index++;
                                        }
                                    }
                                    conn.Close();
                                }
                            }
                            catch (Exception e45)
                            {
                                ep.Close();
                                GC.Collect();
                                System.Threading.Thread.CurrentThread.CurrentCulture = oldCI;
                                ClientScript.RegisterClientScriptBlock(GetType(), "err45", "<script type='text/javascript'>$(document).ready(function(){ ShowError('" + e45.Message + "');});</script>");
                                return;
                            }
                        }
                        #endregion
                        #region движение карт по книге 124 (отчет ОР_КР)
                        if (id_type == 46)
                        {
                            using (SqlConnection conn = new SqlConnection(FuncClass.ConnectionString))
                            {
                                try
                                {
                                    ArrayList soks = new ArrayList();
                                    ArrayList users = new ArrayList();
                                    conn.Open();
                                    using (SqlCommand comm = conn.CreateCommand())
                                    {
                                        int row = 0;
                                        comm.CommandText = $"select [department] from Branchs where id={current_branch}";
                                        branchName = comm.ExecuteScalar().ToString();
                                        ep.SetWorkSheet(1);
                                        ep.SetText(4, 3, branchName);
                                        ep.SetText(6, 4, $"{DatePickerOne.SelectedDate:dd.MM.yyyy}");
                                        #region получение карт
                                        comm.Parameters.Clear();
                                        comm.CommandText =
                                            $"select prod_name, cnt_perso, id_type from V_Products_StorageDocs where id_doc in (select id from StorageDocs where type = {(int)TypeDoc.ToBook124} and id_branch=@br and priz_gen = 1 and date_doc = @dt)";
                                        comm.Parameters.Add("@br", SqlDbType.Int).Value = current_branch;
                                        comm.Parameters.Add("@dt", SqlDbType.Date).Value = DatePickerOne.SelectedDate;
                                        SqlDataReader dr = comm.ExecuteReader();
                                        int mc0 = 0, vis0 = 0, mir0 = 0, pin0 = 0;
                                        int mc = 0, vis = 0, mir = 0, pin = 0;
                                        while (dr.Read())
                                        {
                                            int cnt = Convert.ToInt32(dr["cnt_perso"]);
                                            BaseProductType tp = BranchStore.codeFromTypeAndProdName(Convert.ToInt32(dr["id_type"]), Convert.ToString(dr["prod_name"]));
                                            switch (tp)
                                            {
                                                case BaseProductType.MasterCard: mc0 += cnt; break;
                                                case BaseProductType.VisaCard: vis0 += cnt; break;
                                                case BaseProductType.MirCard: mir0 += cnt; break;
                                                case BaseProductType.PinConvert: pin0 += cnt; break;
                                            }
                                        }
                                        dr.Close();
                                        ep.SetText(13, 3, mc0.ToString());
                                        ep.SetText(13, 4, vis0.ToString());
                                        ep.SetText(13, 5, mir0.ToString());
                                        ep.SetText(13, 6, pin0.ToString());
                                        #endregion
                                        #region выдача карт через сервис
                                        comm.Parameters.Clear();
                                        //сразу определеям сколько и обычной выдачи было, чтобы листов надобавлять для реестров
                                        comm.CommandText = $"select distinct LoweredUserName from V_StorageDocs where type={(int)TypeDoc.SendToClient} and id_branch = @br and priz_gen=1 and date_doc = @dt";
                                        comm.Parameters.Add("@br", SqlDbType.Int).Value = current_branch;
                                        comm.Parameters.Add("@dt", SqlDbType.Date).Value = DatePickerOne.SelectedDate;
                                        dr = comm.ExecuteReader();
                                        while (dr.Read())
                                        {
                                            users.Add(dr["LoweredUserName"]?.ToString());
                                        }
                                        dr.Close();

                                        comm.CommandText =
                                            $"select distinct comment from StorageDocs where type = {(int)TypeDoc.SendToClientService} and id_branch = @br and priz_gen = 1 and date_doc = @dt";
                                        dr = comm.ExecuteReader();
                                        while (dr.Read())
                                        {
                                            soks.Add(dr["comment"]?.ToString());
                                        }
                                        dr.Close();
                                        comm.CommandText =
                                            $"select c.pan, c.fio, c.company, c.isPin, p.name from Cards c " +
                                            "left join Cards_StorageDocs s on c.id = s.id_card " +
                                            "left join Products_Banks pb on c.id_prb = pb.id " +
                                            "left join Products p on pb.id_prod = p.id " +
                                            "where s.id_doc in " +
                                            $"(select id from StorageDocs where type = {(int)TypeDoc.SendToClientService} and id_branch = @br and priz_gen = 1 and date_doc = @dt and comment = @SokName)";
                                        comm.Parameters.Add("@SokName", SqlDbType.VarChar, 100);
                                        int index = 0;
                                        for (i = 0; i < (soks.Count + users.Count) - 1; i++)
                                        {
                                            ep.SetWorkSheet(3);
                                            ep.AddWorkSheet();
                                        }
                                        foreach (string sok in soks)
                                        {
                                            mc = 0; vis = 0; mir = 0; pin = 0;
                                            comm.Parameters["@SokName"].Value = sok;
                                            dr = comm.ExecuteReader();
                                            ep.SetWorkSheet(3+index);
                                            ep.SetWorkSheetName(3+index, $"Реестр {sok}");
                                            row = 0;
                                            while (dr.Read())
                                            {
                                                
                                                BaseProductType tp = BranchStore.codeFromTypeAndProdName(1, Convert.ToString(dr["name"]));
                                                switch (tp)
                                                {
                                                    case BaseProductType.MasterCard: mc++; break;
                                                    case BaseProductType.VisaCard: vis++; break;
                                                    case BaseProductType.MirCard: mir++; break;
                                                }

                                                if (Convert.ToBoolean(dr["isPin"]))
                                                {
                                                    pin++;
                                                    ep.SetText(10 + row, 5, "+");
                                                }
                                                ep.SetText(10 + row, 1, $"{row+1}");
                                                ep.SetText(10 + row, 2, dr["fio"].ToString());
                                                ep.SetText(10 + row, 3, dr["pan"].ToString());
                                                ep.SetText(10 + row, 4, dr["company"].ToString());
                                                row++;
                                            }
                                            dr.Close();
                                            ep.SetRangeBorders(10, 1, 10 + row - 1, 5);
                                            ep.SetText(2, 2, branchName);
                                            ep.SetText(6, 2, $"{DatePickerOne.SelectedDate:dd.MM.yyyy}");
                                            ep.SetRangeBottom(10 + row + 2, 2, 10 + row + 1, 2);
                                            ep.SetText(10 + row + 2, 2, sok);
                                            ep.SetText(10 + row + 3, 2, "Фамилия");
                                            ep.SetRangeBottom(10 + row + 2, 4, 10 + row + 1, 4);
                                            ep.SetText(10 + row + 3, 4, "Подпись");
                                            index++;
                                            ep.SetWorkSheet(1);
                                            ep.SetText(13 + index + 1, 2, sok);
                                            ep.SetText(13 + index + 1, 3, mc.ToString());
                                            ep.SetText(13 + index + 1, 4, vis.ToString());
                                            ep.SetText(13 + index + 1, 5, mir.ToString());
                                            ep.SetText(13 + index + 1, 6, pin.ToString());
                                            ep.ShowRows(13, 13 + index + 1);
                                            mc0 -= mc;
                                            vis0 -= vis;
                                            mir0 -= mir;
                                            pin0 -= pin;
                                        }
                                        #endregion
                                        #region обычная выдача
                                        
                                        comm.Parameters.Clear();
                                        comm.CommandText =
                                            $"select c.pan, c.fio, c.company, c.isPin, p.name from Cards c " +
                                            "left join Cards_StorageDocs s on c.id = s.id_card " +
                                            "left join Products_Banks pb on c.id_prb = pb.id " +
                                            "left join Products p on pb.id_prod = p.id " +
                                            "where s.id_doc in " +
                                            $"(select id from V_StorageDocs where type = {(int)TypeDoc.SendToClient} and id_branch = @br and priz_gen = 1 and date_doc = @dt and LoweredUserName=@uname)";
                                        comm.Parameters.Add("@br", SqlDbType.Int).Value = current_branch;
                                        comm.Parameters.Add("@dt", SqlDbType.Date).Value = DatePickerOne.SelectedDate;
                                        comm.Parameters.Add("@uname", SqlDbType.NChar, 255);
                                        foreach (string user in users)
                                        {
                                            string name = sc.UserFIO(user);
                                            mc = 0; vis = 0; mir = 0; pin = 0;
                                            comm.Parameters["@uname"].Value = user;
                                            dr = comm.ExecuteReader();
                                            ep.SetWorkSheet(3 + index);
                                            ep.SetWorkSheetName(3 + index, $"Реестр {name}");
                                            row = 0;
                                            while (dr.Read())
                                            {

                                                BaseProductType tp =
                                                    BranchStore.codeFromTypeAndProdName(1,
                                                        Convert.ToString(dr["name"]));
                                                switch (tp)
                                                {
                                                    case BaseProductType.MasterCard:
                                                        mc++;
                                                        break;
                                                    case BaseProductType.VisaCard:
                                                        vis++;
                                                        break;
                                                    case BaseProductType.MirCard:
                                                        mir++;
                                                        break;
                                                }

                                                if (Convert.ToBoolean(dr["isPin"]))
                                                {
                                                    pin++;
                                                    ep.SetText(10 + row, 5, "+");
                                                }
                                                ep.SetText(10 + row, 1, $"{row + 1}");
                                                ep.SetText(10 + row, 2, dr["fio"].ToString());
                                                ep.SetText(10 + row, 3, dr["pan"].ToString());
                                                ep.SetText(10 + row, 4, dr["company"].ToString());
                                                row++;
                                            }
                                            dr.Close();
                                            ep.SetRangeBorders(10, 1, 10 + row - 1, 5);
                                            ep.SetText(2, 2, branchName);
                                            ep.SetText(6, 2, $"{DatePickerOne.SelectedDate:dd.MM.yyyy}");
                                            ep.SetRangeBottom(10 + row + 2, 2, 10 + row + 1, 2);
                                            ep.SetText(10 + row + 2, 2, name);
                                            ep.SetText(10 + row + 3, 2, "Фамилия");
                                            ep.SetRangeBottom(10 + row + 2, 4, 10 + row + 1, 4);
                                            ep.SetText(10 + row + 3, 4, "Подпись");
                                            index++;
                                            ep.SetWorkSheet(1);
                                            ep.SetText(13 + index + 1, 2, name);
                                            ep.SetText(13 + index + 1, 3, mc.ToString());
                                            ep.SetText(13 + index + 1, 4, vis.ToString());
                                            ep.SetText(13 + index + 1, 5, mir.ToString());
                                            ep.SetText(13 + index + 1, 6, pin.ToString());
                                            ep.ShowRows(13, 13 + index + 1);
                                            mc0 -= mc;
                                            vis0 -= vis;
                                            mir0 -= mir;
                                            pin0 -= pin;
                                        }

                                        #endregion
                                        #region остатки на конец дня, сданные карты
                                        //28.02.21 - остатки теперь считаем, а не берем из сданных карт
                                        comm.Parameters.Clear();
                                        comm.CommandText =
                                            $"select c.pan, c.fio, c.company, c.isPin, p.name from Cards c " +
                                            "left join Cards_StorageDocs s on c.id = s.id_card " +
                                            "left join Products_Banks pb on c.id_prb = pb.id " +
                                            "left join Products p on pb.id_prod = p.id " +
                                            "where s.id_doc in " +
                                            $"(select id from StorageDocs where type = {(int)TypeDoc.ReceiveBook124} and id_branch = @br and date_doc = @dt)";
                                        comm.Parameters.Add("@br", SqlDbType.Int).Value = current_branch;
                                        comm.Parameters.Add("@dt", SqlDbType.Date).Value = DatePickerOne.SelectedDate;
                                        dr = comm.ExecuteReader();
                                        mc = 0; vis = 0; mir = 0; pin = 0;
                                        ep.SetWorkSheet(2);
                                        row = 0;
                                        while (dr.Read())
                                        {
                                            BaseProductType tp = BranchStore.codeFromTypeAndProdName(1, Convert.ToString(dr["name"]));
                                            switch (tp)
                                            {
                                                case BaseProductType.MasterCard: mc++; break;
                                                case BaseProductType.VisaCard: vis++; break;
                                                case BaseProductType.MirCard: mir++; break;
                                            }

                                            if (Convert.ToBoolean(dr["isPin"]))
                                            {
                                                pin++;
                                                ep.SetText(10 + row, 5, "+");
                                            }

                                            ep.SetText(10 + row, 1, $"{row + 1}");
                                            ep.SetText(10 + row, 2, dr["fio"].ToString());
                                            ep.SetText(10 + row, 3, dr["pan"].ToString());
                                            ep.SetText(10 + row, 4, dr["company"].ToString());
                                            row++;
                                        }
                                        dr.Close();
                                        ep.SetRangeBorders(10, 1, 10 + row - 1, 5);
                                        ep.SetText(2, 2, branchName);
                                        ep.SetText(6, 2, $"{DatePickerOne.SelectedDate:dd.MM.yyyy}");
                                        ep.SetRangeBottom(10 + row + 2, 2, 10 + row + 1, 2);
                                        ep.SetText(10 + row + 3, 2, "Фамилия");
                                        ep.SetRangeBottom(10 + row + 2, 4, 10 + row + 1, 4);
                                        ep.SetText(10 + row + 3, 4, "Подпись");
                                        ep.SetWorkSheet(1);
                                        ep.SetText(81, 3, mc0.ToString());
                                        ep.SetText(81, 4, vis0.ToString());
                                        ep.SetText(81, 5, mir0.ToString());
                                        ep.SetText(81, 6, pin0.ToString());
                                        #endregion
                                    }
                                    ep.SetWorkSheet(1);
                                }
                                catch (Exception e46)
                                {
                                    ep.Close();
                                    GC.Collect();
                                    Thread.CurrentThread.CurrentCulture = oldCI;
                                    ClientScript.RegisterClientScriptBlock(GetType(), "err46", "<script type='text/javascript'>$(document).ready(function(){ ShowError('" + e46.Message + "');});</script>");
                                    return;
                                }
                                finally
                                {
                                    conn?.Close();
                                }

                            }
                        }
                        #endregion
                        #region ведомость подотчетников
                        if (id_type == 47)
                        {
                            //26.02.21 все запросы стали делать только по текущему отделению
                            DateTime date_start = DateTime.ParseExact(DatePickerStart.DatePickerText, "dd.MM.yyyy", null);
                            DateTime date_end = DateTime.ParseExact(DatePickerEnd.DatePickerText, "dd.MM.yyyy", null);

                            using (SqlConnection conn = new SqlConnection(FuncClass.ConnectionString))
                            {
                                conn.Open();
                                int personid = Convert.ToInt32(ddlPersons.SelectedValue);
                                SqlCommand comm = conn.CreateCommand();
                                comm.CommandText = $"select department from Branchs where id={current_branch}";
                                string branch_name = comm.ExecuteScalar()?.ToString();
                                comm.CommandText =
                                    $"select ap.secondname + ' ' + ap.firstname + ' ' + case when ap.patronymic is null then '' else ' ' + ap.patronymic end as fio from AccountablePerson ap where id={personid}";
                                string acc_name = comm.ExecuteScalar()?.ToString();
                                int mc = 0, vis = 0, mir = 0, pin = 0;
                                #region получаем сколько карт у него сегодня, заодно заполняем реестр оставшихся у него карт
                                comm.CommandText = $@"select c.pan, c.fio, c.company, p.name, c.isPin from cards c 
                                                        left join Products_Banks bp on c.id_prb = bp.id
                                                        left join Products p on bp.id_prod = p.id
                                                        where id_stat = {(int)CardStatus.Accountable} and id_person = {personid}
                                                        and (id_branchCard = {current_branch} or id_branchCard in (select id from Branchs where id_parent={current_branch}))";
                                ep.SetWorkSheet(2);
                                int row = 0;
                                using (SqlDataReader dr = comm.ExecuteReader())
                                {
                                    while (dr.Read())
                                    {
                                        ep.SetText(10 + row, 1, $"{row+1}");
                                        ep.SetText(10 + row, 2, dr["fio"]?.ToString());
                                        ep.SetText(10 + row, 3, dr["pan"]?.ToString());
                                        ep.SetText(10 + row, 4, dr["company"]?.ToString());
                                        BaseProductType tp = BranchStore.codeFromTypeAndProdName(1, Convert.ToString(dr["name"]));
                                        switch (tp)
                                        {
                                            case BaseProductType.MasterCard: mc++; break;
                                            case BaseProductType.VisaCard: vis++; break;
                                            case BaseProductType.MirCard: mir++; break;
                                        }
                                        if (Convert.ToBoolean(dr["isPin"]))
                                            pin++;
                                        row++;
                                    }
                                    dr.Close();
                                }
                                if (row > 0)
                                    row--;
                                ep.SetRangeBorders(10, 1, 10 + row, 4);
                                ep.SetRangeBottom(10 + row + 2, 2, 10 + row + 1, 2);
                                ep.SetText(10 + row + 3, 2, "Фамилия");
                                ep.SetRangeBottom(10 + row + 2, 4, 10 + row + 1, 4);
                                ep.SetText(10 + row + 3, 4, "Подпись");
                                ep.SetText("date2", $"{DateTime.Now:dd.MM.yyyy}");
                                ep.SetText("name2", acc_name);
                                #endregion
                                #region начинаем откатываться назад считаем сколько карт получили и сколько отдали
                                int reestr_row = 0;
                                DateTime index_date = DateTime.Now.Date;
                                //выданные карты через сервис
                                comm.CommandText =
                                    $@"select c.pan, c.fio, c.company, p.name, c.isPin from storagedocs sd
                                                        left join Cards_StorageDocs cs on cs.id_doc = sd.id
                                                        left join Cards c on cs.id_card = c.id
                                                        left join Products_Banks pb on c.id_prb=pb.id
                                                        left join Products p on pb.id_prod=p.id
                                                        where sd.type = {(int) TypeDoc.SendToClientService} and sd.priz_gen = 1 
                                                        and (sd.id_branch = {current_branch} or sd.id_branch in (select id from Branchs where id_parent={current_branch}))
                                                        and invoice_courier like '{(int) CardStatus.Accountable},%' and c.id_person = {personid}
                                                        and date_doc=@dt";
                                comm.Parameters.Add("@dt", SqlDbType.Date);
                                //выданные карты через кардперсо
                                SqlCommand sel_cp = conn.CreateCommand();
                                sel_cp.CommandText =
                                    $@"select c.pan, c.fio, c.company, p.name, c.isPin from storagedocs sd
                                                        left join Cards_StorageDocs cs on cs.id_doc = sd.id
                                                        left join Cards c on cs.id_card = c.id
                                                        left join Products_Banks pb on c.id_prb=pb.id
                                                        left join Products p on pb.id_prod=p.id
                                                        left join AccountablePerson_StorageDocs a on sd.id = a.id_doc
                                                        where sd.type = {(int)TypeDoc.SendToClientFromPodotchet} and sd.priz_gen = 1 
                                                        and (sd.id_branch = {current_branch} or sd.id_branch in (select id from Branchs where id_parent={current_branch}))
                                                        and a.id_person={personid}
                                                        and date_doc=@dt";
                                sel_cp.Parameters.Add("@dt", SqlDbType.Date);
                                //отданные другим образом (вернулись в хранилище или в гоз
                                SqlCommand sel_m = conn.CreateCommand();
                                sel_m.CommandText = 
                                    $@"select p.name, c.isPin from storagedocs sd
                                                        left join AccountablePerson_StorageDocs a on a.id_doc=sd.id 
                                                        left join Cards_StorageDocs cs on cs.id_doc = sd.id
                                                        left join Cards c on cs.id_card = c.id
                                                        left join Products_Banks pb on c.id_prb=pb.id
                                                        left join Products p on pb.id_prod=p.id
                                                        where (sd.type = {(int)TypeDoc.ReceiveFromPodotchet} or sd.type = {(int)TypeDoc.ToGozFromPodotchet})
                                                        and (sd.id_branch = {current_branch} or sd.id_branch in (select id from Branchs where id_parent={current_branch}))
                                                        and sd.priz_gen = 1 and a.id_person = {personid}
                                                        and date_doc=@dt";
                                sel_m.Parameters.Add("@dt", SqlDbType.Date);
                                //полученные в подотчет (из хранилища, из гоза)
                                SqlCommand sel_p = conn.CreateCommand();
                                sel_p.CommandText =
                                    $@"select p.name, c.isPin from storagedocs sd
                                                        left join AccountablePerson_StorageDocs a on a.id_doc=sd.id 
                                                        left join Cards_StorageDocs cs on cs.id_doc = sd.id
                                                        left join Cards c on cs.id_card = c.id
                                                        left join Products_Banks pb on c.id_prb=pb.id
                                                        left join Products p on pb.id_prod=p.id
                                                        where (sd.type = {(int)TypeDoc.SendToPodotchet} or sd.type = {(int)TypeDoc.FromGozToPodotchet})
                                                        and (sd.id_branch = {current_branch} or sd.id_branch in (select id from Branchs where id_parent={current_branch}))
                                                        and sd.priz_gen = 1 and a.id_person = {personid}
                                                        and date_doc=@dt";
                                sel_p.Parameters.Add("@dt", SqlDbType.Date);
                                int days_cnt = (date_end - date_start).Days + 1;
                                int d = 1;
                                reestr_row = 0;
                                while (index_date >= date_start)
                                {
                                    int mcp = 0, visp = 0, mirp = 0, pinp = 0;
                                    int mcm = 0, vism = 0, mirm = 0, pinm = 0;
                                    comm.Parameters["@dt"].Value = index_date;
                                    sel_cp.Parameters["@dt"].Value = index_date;
                                    sel_p.Parameters["@dt"].Value = index_date;
                                    sel_m.Parameters["@dt"].Value = index_date;
                                    using (SqlDataReader dr = comm.ExecuteReader())
                                    {
                                        while (dr.Read())
                                        {
                                            BaseProductType tp =
                                                BranchStore.codeFromTypeAndProdName(1, Convert.ToString(dr["name"]));
                                            switch (tp)
                                            {
                                                case BaseProductType.MasterCard:
                                                    mcm++;
                                                    break;
                                                case BaseProductType.VisaCard:
                                                    vism++;
                                                    break;
                                                case BaseProductType.MirCard:
                                                    mirm++;
                                                    break;
                                            }

                                            if (Convert.ToBoolean(dr["isPin"]))
                                                pinm++;
                                            if (index_date <= date_end)
                                            {
                                                //идем на реестр выданных
                                                ep.SetWorkSheet(3);
                                                ep.SetText(10 + reestr_row, 1, $"{reestr_row + 1}");
                                                ep.SetText(10 + reestr_row, 2, dr["fio"]?.ToString());
                                                ep.SetText(10 + reestr_row, 3, dr["pan"]?.ToString());
                                                ep.SetText(10 + reestr_row, 4, dr["company"]?.ToString());
                                                ep.SetText(10 + reestr_row, 5, $"{index_date:dd.MM.yyyy}");
                                                reestr_row++;
                                            }
                                        }
                                        dr.Close();
                                    }
                                    using (SqlDataReader dr = sel_cp.ExecuteReader())
                                    {
                                        while (dr.Read())
                                        {
                                            BaseProductType tp =
                                                BranchStore.codeFromTypeAndProdName(1, Convert.ToString(dr["name"]));
                                            switch (tp)
                                            {
                                                case BaseProductType.MasterCard:
                                                    mcm++;
                                                    break;
                                                case BaseProductType.VisaCard:
                                                    vism++;
                                                    break;
                                                case BaseProductType.MirCard:
                                                    mirm++;
                                                    break;
                                            }

                                            if (Convert.ToBoolean(dr["isPin"]))
                                                pinm++;
                                            if (index_date <= date_end)
                                            {
                                                //идем на реестр выданных
                                                ep.SetWorkSheet(3);
                                                ep.SetText(10 + reestr_row, 1, $"{reestr_row + 1}");
                                                ep.SetText(10 + reestr_row, 2, dr["fio"]?.ToString());
                                                ep.SetText(10 + reestr_row, 3, dr["pan"]?.ToString());
                                                ep.SetText(10 + reestr_row, 4, dr["company"]?.ToString());
                                                ep.SetText(10 + reestr_row, 5, $"{index_date:dd.MM.yyyy}");
                                                reestr_row++;
                                            }
                                        }
                                        dr.Close();
                                    }
                                    using (SqlDataReader dr = sel_m.ExecuteReader())
                                    {
                                        while (dr.Read())
                                        {
                                            BaseProductType tp =
                                                BranchStore.codeFromTypeAndProdName(1, Convert.ToString(dr["name"]));
                                            switch (tp)
                                            {
                                                case BaseProductType.MasterCard:
                                                    mcm++;
                                                    break;
                                                case BaseProductType.VisaCard:
                                                    vism++;
                                                    break;
                                                case BaseProductType.MirCard:
                                                    mirm++;
                                                    break;
                                            }

                                            if (Convert.ToBoolean(dr["isPin"]))
                                                pinm++;
                                        }

                                        dr.Close();
                                    }

                                    using (SqlDataReader dr = sel_p.ExecuteReader())
                                    {
                                        while (dr.Read())
                                        {
                                            BaseProductType tp =
                                                BranchStore.codeFromTypeAndProdName(1, Convert.ToString(dr["name"]));
                                            switch (tp)
                                            {
                                                case BaseProductType.MasterCard:
                                                    mcp++;
                                                    break;
                                                case BaseProductType.VisaCard:
                                                    visp++;
                                                    break;
                                                case BaseProductType.MirCard:
                                                    mirp++;
                                                    break;
                                            }

                                            if (Convert.ToBoolean(dr["isPin"]))
                                                pinp++;
                                        }
                                        dr.Close();
                                    }
                                    if (index_date <= date_end)
                                    {
                                        ep.SetWorkSheet(1);
                                        ep.SetText(12 + days_cnt * 5 - d * 5, 2, $"{index_date:dd.MM.yyyy}");
                                        ep.SetText(12 + days_cnt * 5 - d * 5, 3, $"MasterCard (ИТОГО)");
                                        ep.SetText(12 + days_cnt * 5 - d * 5, 4, $"{mc+mcm-mcp}");
                                        ep.SetText(12 + days_cnt * 5 - d * 5, 5, $"{mcp}");
                                        ep.SetText(12 + days_cnt * 5 - d * 5, 6, $"{mcm}");
                                        ep.SetText(12 + days_cnt * 5 - d * 5, 7, $"{mc}");
                                        ep.SetText(12 + days_cnt * 5 - d * 5 + 1, 2, $"{index_date:dd.MM.yyyy}");
                                        ep.SetText(12 + days_cnt * 5 - d * 5 + 1, 3, $"Visa (ИТОГО)");
                                        ep.SetText(12 + days_cnt * 5 - d * 5 + 1, 4, $"{vis+vism-visp}");
                                        ep.SetText(12 + days_cnt * 5 - d * 5 + 1, 5, $"{visp}");
                                        ep.SetText(12 + days_cnt * 5 - d * 5 + 1, 6, $"{vism}");
                                        ep.SetText(12 + days_cnt * 5 - d * 5 + 1, 7, $"{vis}");
                                        ep.SetText(12 + days_cnt * 5 - d * 5 + 2, 2, $"{index_date:dd.MM.yyyy}");
                                        ep.SetText(12 + days_cnt * 5 - d * 5 + 2, 3, $"Карты МИР (ИТОГО)");
                                        ep.SetText(12 + days_cnt * 5 - d * 5 + 2, 4, $"{mir+mirm-mirp}");
                                        ep.SetText(12 + days_cnt * 5 - d * 5 + 2, 5, $"{mirp}");
                                        ep.SetText(12 + days_cnt * 5 - d * 5 + 2, 6, $"{mirm}");
                                        ep.SetText(12 + days_cnt * 5 - d * 5 + 2, 7, $"{mir}");
                                        ep.SetText(12 + days_cnt * 5 - d * 5 + 3, 2, $"{index_date:dd.MM.yyyy}");
                                        ep.SetText(12 + days_cnt * 5 - d * 5 + 3, 3, $"Пин конверт");
                                        ep.SetText(12 + days_cnt * 5 - d * 5 + 3, 4, $"{pin+pinm-pinp}");
                                        ep.SetText(12 + days_cnt * 5 - d * 5 + 3, 5, $"{pinp}");
                                        ep.SetText(12 + days_cnt * 5 - d * 5 + 3, 6, $"{pinm}");
                                        ep.SetText(12 + days_cnt * 5 - d * 5 + 3, 7, $"{pin}");
                                        d++;
                                    }
                                    mc = mc + mcm - mcp;
                                    vis = vis + vism - visp;
                                    mir = mir + mirm - mirp;
                                    pin = pin + pinm - pinp;
                                    index_date = index_date.AddDays(-1);
                                }
                                ep.SetWorkSheet(3);
                                reestr_row--;
                                ep.SetRangeBorders(10, 1, 10 + reestr_row, 5);
                                ep.SetRangeBottom(10 + reestr_row + 2, 2, 10 + reestr_row + 1, 2);
                                ep.SetText(10 + reestr_row + 3, 2, "Фамилия");
                                ep.SetRangeBottom(10 + reestr_row + 2, 4, 11 + reestr_row + 1, 4);
                                ep.SetText(10 + reestr_row + 3, 4, "Подпись");
                                ep.SetText("date3", $"{DateTime.Now:dd.MM.yyyy}");
                                ep.SetText("name3", acc_name);
                                ep.SetWorkSheet(1);
                                ep.SetRangeBorders(12, 1, 12 + days_cnt * 5, 7);
                                ep.SetRangeBottom(12 + days_cnt * 5 + 2, 3, 12 + days_cnt * 5 + 2, 6);
                                ep.SetText(12 + days_cnt * 5 + 3, 3, "должность");
                                ep.SetText(12 + days_cnt * 5 + 3, 4, "фамилия и инициалы");
                                ep.SetText(12 + days_cnt * 5 + 3, 6, "подпись");
                                ep.SetRangeBottom(12 + days_cnt * 5 + 4, 3, 12 + days_cnt * 5 + 4, 6);
                                ep.SetText(12 + days_cnt * 5 + 5, 3, "должность");
                                ep.SetText(12 + days_cnt * 5 + 5, 4, "фамилия и инициалы");
                                ep.SetText(12 + days_cnt * 5 + 5, 6, "подпись");
                                ep.SetText("branch", branch_name);
                                ep.SetText("date1", $"{DateTime.Now:dd.MM.yyyy}");
                                ep.SetText("fio", acc_name);
                                comm.Parameters.Clear();
                                comm.CommandText = $"select dbo.UserPosition('{User.Identity.Name}')";
                                ep.SetText(12 + days_cnt * 5 + 2, 3, comm.ExecuteScalar()?.ToString());
                                comm.CommandText = $"select dbo.UserFio('{User.Identity.Name}')";
                                ep.SetText(12 + days_cnt * 5 + 2, 4, comm.ExecuteScalar()?.ToString());
                                #endregion
                                conn.Close();
                            }
                        }
                        #endregion
                        #region отчет приход расход завкассы
                        if (id_type == 48)
                        {
                            DateTime repdate = DateTime.ParseExact(DatePickerOne.DatePickerText, "dd.MM.yyyy", null);
                            using (SqlConnection conn = new SqlConnection(FuncClass.ConnectionString))
                            {
                                conn.Open();                                
                                SqlCommand comm = conn.CreateCommand();
                                comm.CommandText = $"select department from Branchs where id={current_branch}";
                                string branch_name = comm.ExecuteScalar()?.ToString();
                                // заполняем все продукт
                                comm.CommandText = "select id_prb, prod_name, bank_name, id_type from V_ProductsBanks_T order by id_type, id_sort";
                                Hashtable prods = new Hashtable();
                                int pinid = 0;
                                using (SqlDataReader dr = comm.ExecuteReader())
                                {
                                    while (dr.Read())
                                    {
                                        prods.Add(Convert.ToInt32(dr["id_prb"]),
                                            new MyProd(Convert.ToInt32(dr["id_prb"]), dr["prod_name"].ToString(), dr["bank_name"].ToString(), Convert.ToInt32(dr["id_type"])));
                                        if (Convert.ToInt32(dr["id_type"]) == 2)
                                            pinid = Convert.ToInt32(dr["id_prb"]);
                                    }
                                    dr.Read();
                                }
                                #region получаем сколько карт у него сегодня, заодно заполняем реестр оставшихся у него карт
                                comm.CommandText = $@"select c.pan, c.fio, c.company, bp.id as pid, p.name, c.isPin from cards c 
                                                            left join Products_Banks bp on c.id_prb = bp.id
                                                            left join Products p on bp.id_prod = p.id
                                                            where id_stat = {(int)CardStatus.Filial} and id_branchCurrent={current_branch}";
                                ep.SetWorkSheet(2);
                                int row = 0;
                                
                                using (SqlDataReader dr = comm.ExecuteReader())
                                {
                                    while (dr.Read())
                                    {
                                        ep.SetText(10 + row, 1, $"{row + 1}");
                                        ep.SetText(10 + row, 2, dr["fio"]?.ToString());
                                        ep.SetText(10 + row, 3, dr["pan"]?.ToString());
                                        ep.SetText(10 + row, 4, dr["company"]?.ToString());
                                        if (prods.ContainsKey(Convert.ToInt32(dr["pid"])))
                                            ((MyProd) prods[Convert.ToInt32(dr["pid"])]).cnts[3]++;
                                        if (Convert.ToBoolean(dr["isPin"]) && prods.ContainsKey(pinid))
                                            ((MyProd)prods[pinid]).cnts[3]++;
                                        row++;
                                    }
                                    dr.Close();
                                }
                                if (row > 0)
                                    row--;
                                ep.SetRangeBorders(10, 1, 10 + row, 4);
                                ep.SetRangeBottom(10 + row + 2, 2, 10 + row + 1, 2);
                                ep.SetText(10 + row + 3, 2, "Фамилия");
                                ep.SetRangeBottom(10 + row + 2, 4, 10 + row + 1, 4);
                                ep.SetText("name2", branch_name);
                                ep.SetText("date2", $"{DateTime.Now:dd.MM.yyyy}");
                                #endregion
                                #region начинаем откатываться назад считаем сколько карт получили и сколько отдали
                                DateTime index_date = DateTime.Now.Date;
                                //выданные карты через сервис
                                comm.CommandText =
                                    $@"select pb.id as pid, c.isPin from storagedocs sd
                                                                left join Cards_StorageDocs cs on cs.id_doc = sd.id
                                                                left join Cards c on cs.id_card = c.id
                                                                left join Products_Banks pb on c.id_prb=pb.id
                                                                left join Products p on pb.id_prod=p.id
                                                                where sd.type = {(int)TypeDoc.SendToClientService} and sd.priz_gen = 1 
                                                                and invoice_courier like '{(int)CardStatus.Filial},%' and sd.id_branch = {current_branch}
                                                                and date_doc=@dt";
                                comm.Parameters.Add("@dt", SqlDbType.Date);//        
                                //отданные другим образом (обычная выдача, возврат ценностей, передача филиал-филиал, отказ от получения, передача ДО,
                                //отправка на экспертизу в ПЦ, передача в подотчет, передача в книгу 124, уничтожение карт)
                                SqlCommand sel_m = conn.CreateCommand();
                                sel_m.CommandText =
                                    $@"select pb.id as pid, c.isPin from storagedocs sd                                                                
                                                                left join Cards_StorageDocs cs on cs.id_doc = sd.id
                                                                left join Cards c on cs.id_card = c.id
                                                                left join Products_Banks pb on c.id_prb=pb.id
                                                                left join Products p on pb.id_prod=p.id
                                                                where (sd.type = {(int)TypeDoc.SendToClient} or sd.type = {(int)TypeDoc.ReceiveToBank}
                                                                    or sd.type = {(int)TypeDoc.FilialFilial} or sd.type = {(int)TypeDoc.DontReceiveToFilial}
                                                                    or sd.type = {(int)TypeDoc.SendToFilialFilial} or sd.type = {(int)TypeDoc.SendToExpertiza}
                                                                    or sd.type = {(int)TypeDoc.SendToPodotchet} or sd.type = {(int)TypeDoc.ToBook124}
                                                                    or sd.type = {(int)TypeDoc.ToGoz} or sd.type = {(int)TypeDoc.KillingCard})
                                                                and sd.priz_gen = 1 and sd.id_branch = {current_branch}
                                                                and date_doc=@dt";
                                sel_m.Parameters.Add("@dt", SqlDbType.Date);
                                //полученные в хранилище (из пц или иного филиала, получение на экспертизу в филиале, из подотчета, получение пакетно (Сейчас не используется)
                                SqlCommand sel_p = conn.CreateCommand();
                                sel_p.CommandText =
                                    $@"select pb.id as pid, c.isPin from storagedocs sd
                                                                left join Cards_StorageDocs cs on cs.id_doc = sd.id
                                                                left join Cards c on cs.id_card = c.id
                                                                left join Products_Banks pb on c.id_prb=pb.id
                                                                left join Products p on pb.id_prod=p.id
                                                                where (sd.type = {(int)TypeDoc.ReceiveToFilial} or sd.type = {(int)TypeDoc.ReceiveToFilialExpertiza}
                                                                    or sd.type = {(int)TypeDoc.ReceiveFromPodotchet} or sd.type = {(int)TypeDoc.ReceiveBook124}
                                                                    or sd.type = {(int)TypeDoc.ReceiveToFilialPacket} or sd.type = {(int)TypeDoc.ReceiveGoz})
                                                                and sd.priz_gen = 1 and sd.id_branch = {current_branch}
                                                                and date_doc=@dt";
                                sel_p.Parameters.Add("@dt", SqlDbType.Date);
                                int d = 1;
                                while (index_date >= repdate)
                                {
                                    comm.Parameters["@dt"].Value = index_date;
                                    sel_p.Parameters["@dt"].Value = index_date;
                                    sel_m.Parameters["@dt"].Value = index_date;
                                    using (SqlDataReader dr = comm.ExecuteReader())
                                    {
                                        while (dr.Read())
                                        {
                                            if (prods.ContainsKey(Convert.ToInt32(dr["pid"])))
                                                ((MyProd)prods[Convert.ToInt32(dr["pid"])]).cnts[2]++;
                                            if (Convert.ToBoolean(dr["isPin"]) && prods.ContainsKey(pinid))
                                                ((MyProd)prods[pinid]).cnts[2]++;
                                        }
                                        dr.Close();
                                    }
                                    using (SqlDataReader dr = sel_m.ExecuteReader())
                                    {
                                        while (dr.Read())
                                        {
                                            if (prods.ContainsKey(Convert.ToInt32(dr["pid"])))
                                                ((MyProd)prods[Convert.ToInt32(dr["pid"])]).cnts[2]++;
                                            if (Convert.ToBoolean(dr["isPin"]) && prods.ContainsKey(pinid))
                                                ((MyProd)prods[pinid]).cnts[2]++;
                                        }
                                        dr.Close();
                                    }

                                    using (SqlDataReader dr = sel_p.ExecuteReader())
                                    {
                                        while (dr.Read())
                                        {
                                            if (prods.ContainsKey(Convert.ToInt32(dr["pid"])))
                                                ((MyProd)prods[Convert.ToInt32(dr["pid"])]).cnts[1]++;
                                            if (Convert.ToBoolean(dr["isPin"]) && prods.ContainsKey(pinid))
                                                ((MyProd)prods[pinid]).cnts[1]++;
                                        }
                                        dr.Close();
                                    }
                                    if (index_date == repdate) //выводим на лист
                                    {
                                        row = 0;
                                        ep.SetWorkSheet(1);
                                        int c1, c2, c3, c11, c22, c33;
                                        c11 = 0; c22 = 0; c33 = 0;
                                        //подсчет мастеркард
                                        c1 = 0; c2 = 0; c3 = 0;
                                        foreach (int key in prods.Keys)
                                        {
                                            MyProd mp = (MyProd)prods[key];
                                            if (BranchStore.codeFromTypeAndProdName(mp.Type, mp.Name) != BaseProductType.MasterCard)
                                                continue;
                                            if (mp.cnts[3] + mp.cnts[2] + mp.cnts[1] != 0)
                                            {
                                                c1 += mp.cnts[1];
                                                c2 += mp.cnts[2];
                                                c3 += mp.cnts[3];
                                                ep.SetText(10 + row, 3, mp.Name);
                                                ep.SetText(10 + row, 4, (mp.cnts[3] - mp.cnts[1] + mp.cnts[2]).ToString());
                                                ep.SetText(10 + row, 5, mp.cnts[1].ToString());
                                                ep.SetText(10 + row, 6, mp.cnts[2].ToString());
                                                ep.SetText(10 + row, 7, mp.cnts[3].ToString());
                                                row++;
                                            }
                                        }
                                        if ((c1 + c2 + c3) > 0)
                                        {
                                            c11 += c1; c22 += c2; c33 += c3;
                                            ep.SetText(10 + row, 3, "MasterCard (итого):");
                                            ep.SetText(10 + row, 4, (c3 - c1 + c2).ToString());
                                            ep.SetText(10 + row, 5, c1.ToString());
                                            ep.SetText(10 + row, 6, c2.ToString());
                                            ep.SetText(10 + row, 7, c3.ToString());
                                            ep.SetRangeBold(10 + row, 1, 10 + row, 7);
                                            row++;
                                        }

                                        //подсчет виза
                                        c1 = 0; c2 = 0; c3 = 0;
                                        foreach (int key in prods.Keys)
                                        {
                                            MyProd mp = (MyProd)prods[key];
                                            if (BranchStore.codeFromTypeAndProdName(mp.Type, mp.Name) != BaseProductType.VisaCard)
                                                continue;
                                            if (mp.cnts[3] + mp.cnts[2] + mp.cnts[1] != 0)
                                            {
                                                c1 += mp.cnts[1];
                                                c2 += mp.cnts[2];
                                                c3 += mp.cnts[3];
                                                ep.SetText(10 + row, 3, mp.Name);
                                                ep.SetText(10 + row, 4, (mp.cnts[3] - mp.cnts[1] + mp.cnts[2]).ToString());
                                                ep.SetText(10 + row, 5, mp.cnts[1].ToString());
                                                ep.SetText(10 + row, 6, mp.cnts[2].ToString());
                                                ep.SetText(10 + row, 7, mp.cnts[3].ToString());
                                                row++;
                                            }
                                        }
                                        if ((c1 + c2 + c3) > 0)
                                        {
                                            c11 += c1; c22 += c2; c33 += c3;
                                            ep.SetText(10 + row, 3, "Visa (итого):");
                                            ep.SetText(10 + row, 4, (c3 - c1 + c2).ToString());
                                            ep.SetText(10 + row, 5, c1.ToString());
                                            ep.SetText(10 + row, 6, c2.ToString());
                                            ep.SetText(10 + row, 7, c3.ToString());
                                            ep.SetRangeBold(10+row, 1, 10+row, 7);
                                            row++;
                                        }
                                        //подсчет нспк
                                        c1 = 0; c2 = 0; c3 = 0;
                                        foreach (int key in prods.Keys)
                                        {
                                            MyProd mp = (MyProd)prods[key];
                                            if (BranchStore.codeFromTypeAndProdName(mp.Type, mp.Name) != BaseProductType.MirCard)
                                                continue;
                                            if (mp.cnts[3] + mp.cnts[2] + mp.cnts[1] != 0)
                                            {
                                                c1 += mp.cnts[1];
                                                c2 += mp.cnts[2];
                                                c3 += mp.cnts[3];
                                                ep.SetText(10 + row, 3, mp.Name);
                                                ep.SetText(10 + row, 4, (mp.cnts[3] - mp.cnts[1] + mp.cnts[2]).ToString());
                                                ep.SetText(10 + row, 5, mp.cnts[1].ToString());
                                                ep.SetText(10 + row, 6, mp.cnts[2].ToString());
                                                ep.SetText(10 + row, 7, mp.cnts[3].ToString());
                                                row++;
                                            }
                                        }

                                        if ((c1 + c2 + c3) > 0)
                                        {
                                            c11 += c1; c22 += c2; c33 += c3;
                                            ep.SetText(10 + row, 3, "Карты МИР (итого):");
                                            ep.SetText(10 + row, 4, (c3 - c1 + c2).ToString());
                                            ep.SetText(10 + row, 5, c1.ToString());
                                            ep.SetText(10 + row, 6, c2.ToString());
                                            ep.SetText(10 + row, 7, c3.ToString());
                                            ep.SetRangeBold(10 + row, 1, 10 + row, 7);
                                            row++;
                                        }
                                        //подсчет nfc
                                        c1 = 0; c2 = 0; c3 = 0;
                                        foreach (int key in prods.Keys)
                                        {
                                            MyProd mp = (MyProd)prods[key];
                                            if (BranchStore.codeFromTypeAndProdName(mp.Type, mp.Name) != BaseProductType.NFCCard)
                                                continue;
                                            if (mp.cnts[3] + mp.cnts[2] + mp.cnts[1] != 0)
                                            {
                                                c1 += mp.cnts[1];
                                                c2 += mp.cnts[2];
                                                c3 += mp.cnts[3];
                                                ep.SetText(10 + row, 3, mp.Name);
                                                ep.SetText(10 + row, 4, (mp.cnts[3] - mp.cnts[1] + mp.cnts[2]).ToString());
                                                ep.SetText(10 + row, 5, mp.cnts[1].ToString());
                                                ep.SetText(10 + row, 6, mp.cnts[2].ToString());
                                                ep.SetText(10 + row, 7, mp.cnts[3].ToString());
                                                row++;
                                            }
                                        }

                                        if ((c1 + c2 + c3) > 0)
                                        {
                                            c11 += c1; c22 += c2; c33 += c3;
                                            ep.SetText(10 + row, 3, "Карты NFC (итого):");
                                            ep.SetText(10 + row, 4, (c3 - c1 + c2).ToString());
                                            ep.SetText(10 + row, 5, c1.ToString());
                                            ep.SetText(10 + row, 6, c2.ToString());
                                            ep.SetText(10 + row, 7, c3.ToString());
                                            ep.SetRangeBold(10 + row, 1, 10 + row, 7);
                                            row++;
                                        }
                                        //подсчет сервисных
                                        c1 = 0; c2 = 0; c3 = 0;
                                        foreach (int key in prods.Keys)
                                        {
                                            MyProd mp = (MyProd)prods[key];
                                            if (BranchStore.codeFromTypeAndProdName(mp.Type, mp.Name) != BaseProductType.ServiceCard)
                                                continue;
                                            if (mp.cnts[3] + mp.cnts[2] + mp.cnts[1] != 0)
                                            {
                                                c1 += mp.cnts[1];
                                                c2 += mp.cnts[2];
                                                c3 += mp.cnts[3];
                                                ep.SetText(10 + row, 3, mp.Name);
                                                ep.SetText(10 + row, 4, (mp.cnts[3] - mp.cnts[1] + mp.cnts[2]).ToString());
                                                ep.SetText(10 + row, 5, mp.cnts[1].ToString());
                                                ep.SetText(10 + row, 6, mp.cnts[2].ToString());
                                                ep.SetText(10 + row, 7, mp.cnts[3].ToString());
                                                row++;
                                            }
                                        }
                                        if ((c1 + c2 + c3) > 0)
                                        {
                                            c11 += c1; c22 += c2; c33 += c3;
                                            ep.SetText(10 + row, 3, "Сервисные карт (итого):");
                                            ep.SetText(10 + row, 4, (c3 - c1 + c2).ToString());
                                            ep.SetText(10 + row, 5, c1.ToString());
                                            ep.SetText(10 + row, 6, c2.ToString());
                                            ep.SetText(10 + row, 7, c3.ToString());
                                            ep.SetRangeBold(10 + row, 1, 10 + row, 7);
                                            row++;
                                        }
                                        ep.SetText(10 + row, 3, "Все карты (итого):");
                                        ep.SetText(10 + row, 4, (c33 - c11 + c22).ToString());
                                        ep.SetText(10 + row, 5, c11.ToString());
                                        ep.SetText(10 + row, 6, c22.ToString());
                                        ep.SetText(10 + row, 7, c33.ToString());
                                        ep.SetRangeBold(10 + row, 1, 10 + row, 7);
                                        row++;
                                        //пины
                                        foreach (int key in prods.Keys)
                                        {
                                            MyProd mp = (MyProd)prods[key];
                                            if (BranchStore.codeFromTypeAndProdName(mp.Type, mp.Name) != BaseProductType.PinConvert)
                                                continue;
                                            if (mp.cnts[3] + mp.cnts[2] + mp.cnts[1] != 0)
                                            {
                                                ep.SetText(10 + row, 3, mp.Name);
                                                ep.SetText(10 + row, 4, (mp.cnts[3] - mp.cnts[1] + mp.cnts[2]).ToString());
                                                ep.SetText(10 + row, 5, mp.cnts[1].ToString());
                                                ep.SetText(10 + row, 6, mp.cnts[2].ToString());
                                                ep.SetText(10 + row, 7, mp.cnts[3].ToString());
                                                ep.SetRangeBold(10 + row, 1, 10 + row, 7);
                                                row++;
                                            }
                                        }
                                        ep.SetRangeBorders(10, 1, 10 + row - 1, 7);
                                        ep.SetText(10 + row + 3, 3, "должность");
                                        ep.SetText(10 + row + 3, 4, "фамилия и инициалы");
                                        ep.SetText(10 + row + 3, 6, "подпись");
                                        ep.SetRangeAlignment(10 + row + 3, 1, 10 + row + 3, 7, Microsoft.Office.Interop.Excel.Constants.xlLeft);
                                        ep.SetRangeBottom(10 + row + 2, 3, 10 + row + 2, 6);
                                        ep.SetText(10 + row + 5, 3, "должность");
                                        ep.SetText(10 + row + 5, 4, "фамилия и инициалы");
                                        ep.SetText(10 + row + 5, 6, "подпись");
                                        ep.SetRangeAlignment(10 + row + 5, 1, 10 + row + 5, 7, Microsoft.Office.Interop.Excel.Constants.xlLeft);
                                        ep.SetRangeBottom(10 + row + 4, 3, 10 + row + 4, 6);

                                    }
                                    else // либо актуализируем кол-во на конец дня
                                    {
                                        foreach (int key in prods.Keys)
                                        {
                                            ((MyProd) prods[key]).cnts[3] =
                                                ((MyProd) prods[key]).cnts[3] - ((MyProd) prods[key]).cnts[1] +
                                                ((MyProd) prods[key]).cnts[2];
                                            ((MyProd) prods[key]).cnts[1] = 0;
                                            ((MyProd) prods[key]).cnts[2] = 0;
                                        }
                                    }
                                    index_date = index_date.AddDays(-1);
                                }
                                #endregion
                                ep.SetText("name1", branch_name);
                                ep.SetText("date1", $"{repdate:dd.MM.yyyy}");
                                conn.Close();
                            }
                        }
                        #endregion
                        #region отчет приход расход гоз
                        if (id_type == 49)
                        {
                            DateTime repdate = DateTime.ParseExact(DatePickerOne.DatePickerText, "dd.MM.yyyy", null);
                            using (SqlConnection conn = new SqlConnection(FuncClass.ConnectionString))
                            {
                                conn.Open();
                                SqlCommand comm = conn.CreateCommand();
                                comm.CommandText = $"select department from Branchs where id={current_branch}";
                                string branch_name = comm.ExecuteScalar()?.ToString();
                                // заполняем все продукт
                                comm.CommandText = "select id_prb, prod_name, bank_name, id_type from V_ProductsBanks_T order by id_type, id_sort";
                                Hashtable prods = new Hashtable();
                                int pinid = 0;
                                using (SqlDataReader dr = comm.ExecuteReader())
                                {
                                    while (dr.Read())
                                    {
                                        prods.Add(Convert.ToInt32(dr["id_prb"]),
                                            new MyProd(Convert.ToInt32(dr["id_prb"]), dr["prod_name"].ToString(), dr["bank_name"].ToString(), Convert.ToInt32(dr["id_type"])));
                                        if (Convert.ToInt32(dr["id_type"]) == 2)
                                            pinid = Convert.ToInt32(dr["id_prb"]);
                                    }
                                    dr.Read();
                                }
                                #region получаем сколько карт у него сегодня, заодно заполняем реестр оставшихся у него карт
                                comm.CommandText = $@"select c.pan, c.fio, c.company, bp.id as pid, p.name, c.isPin from cards c 
                                                            left join Products_Banks bp on c.id_prb = bp.id
                                                            left join Products p on bp.id_prod = p.id
                                                            where id_stat = {(int)CardStatus.Goz} and id_branchCurrent={current_branch}";
                                ep.SetWorkSheet(2);
                                int row = 0;

                                using (SqlDataReader dr = comm.ExecuteReader())
                                {
                                    while (dr.Read())
                                    {
                                        ep.SetText(10 + row, 1, $"{row + 1}");
                                        ep.SetText(10 + row, 2, dr["fio"]?.ToString());
                                        ep.SetText(10 + row, 3, dr["pan"]?.ToString());
                                        ep.SetText(10 + row, 4, dr["company"]?.ToString());
                                        if (prods.ContainsKey(Convert.ToInt32(dr["pid"])))
                                            ((MyProd)prods[Convert.ToInt32(dr["pid"])]).cnts[3]++;
                                        if (Convert.ToBoolean(dr["isPin"]) && prods.ContainsKey(pinid))
                                            ((MyProd)prods[pinid]).cnts[3]++;
                                        row++;
                                    }
                                    dr.Close();
                                }
                                if (row > 0)
                                    row--;
                                ep.SetRangeBorders(10, 1, 10 + row, 4);
                                ep.SetRangeBottom(10 + row + 2, 2, 10 + row + 1, 2);
                                ep.SetText(10 + row + 3, 2, "Фамилия");
                                ep.SetRangeBottom(10 + row + 2, 4, 10 + row + 1, 4);
                                ep.SetText("name2", branch_name + " (ГОЗ)");
                                ep.SetText("date2", $"{DateTime.Now:dd.MM.yyyy}");
                                #endregion
                                #region начинаем откатываться назад считаем сколько карт получили и сколько отдали
                                DateTime index_date = DateTime.Now.Date;
                                //отданные: вернуть из гоз, передать в подотчет 
                                SqlCommand sel_m = conn.CreateCommand();
                                sel_m.CommandText =
                                    $@"select pb.id as pid, c.isPin from storagedocs sd                                                                
                                                                left join Cards_StorageDocs cs on cs.id_doc = sd.id
                                                                left join Cards c on cs.id_card = c.id
                                                                left join Products_Banks pb on c.id_prb=pb.id
                                                                left join Products p on pb.id_prod=p.id
                                                                where (sd.type = {(int)TypeDoc.FromGoz} or sd.type = {(int)TypeDoc.FromGozToPodotchet})
                                                                and sd.priz_gen = 1 and sd.id_branch = {current_branch}
                                                                and date_doc=@dt";
                                sel_m.Parameters.Add("@dt", SqlDbType.Date);
                                //полученные: приянть в гоз, принять из подотчета
                                SqlCommand sel_p = conn.CreateCommand();
                                sel_p.CommandText =
                                    $@"select pb.id as pid, c.isPin from storagedocs sd
                                                                left join Cards_StorageDocs cs on cs.id_doc = sd.id
                                                                left join Cards c on cs.id_card = c.id
                                                                left join Products_Banks pb on c.id_prb=pb.id
                                                                left join Products p on pb.id_prod=p.id
                                                                where (sd.type = {(int)TypeDoc.GetGoz} or sd.type = {(int)TypeDoc.ToGozFromPodotchet})
                                                                and sd.priz_gen = 1 and sd.id_branch = {current_branch}
                                                                and date_doc=@dt";
                                sel_p.Parameters.Add("@dt", SqlDbType.Date);
                                int d = 1;
                                while (index_date >= repdate)
                                {
                                    sel_p.Parameters["@dt"].Value = index_date;
                                    sel_m.Parameters["@dt"].Value = index_date;
                                    using (SqlDataReader dr = sel_m.ExecuteReader())
                                    {
                                        while (dr.Read())
                                        {
                                            if (prods.ContainsKey(Convert.ToInt32(dr["pid"])))
                                                ((MyProd)prods[Convert.ToInt32(dr["pid"])]).cnts[2]++;
                                            if (Convert.ToBoolean(dr["isPin"]) && prods.ContainsKey(pinid))
                                                ((MyProd)prods[pinid]).cnts[2]++;
                                        }
                                        dr.Close();
                                    }

                                    using (SqlDataReader dr = sel_p.ExecuteReader())
                                    {
                                        while (dr.Read())
                                        {
                                            if (prods.ContainsKey(Convert.ToInt32(dr["pid"])))
                                                ((MyProd)prods[Convert.ToInt32(dr["pid"])]).cnts[1]++;
                                            if (Convert.ToBoolean(dr["isPin"]) && prods.ContainsKey(pinid))
                                                ((MyProd)prods[pinid]).cnts[1]++;
                                        }
                                        dr.Close();
                                    }
                                    if (index_date == repdate) //выводим на лист
                                    {
                                        row = 0;
                                        ep.SetWorkSheet(1);
                                        int c1, c2, c3, c11, c22, c33;
                                        c11 = 0; c22 = 0; c33 = 0;
                                        //подсчет мастеркард
                                        c1 = 0; c2 = 0; c3 = 0;
                                        foreach (int key in prods.Keys)
                                        {
                                            MyProd mp = (MyProd)prods[key];
                                            if (BranchStore.codeFromTypeAndProdName(mp.Type, mp.Name) != BaseProductType.MasterCard)
                                                continue;
                                            if (mp.cnts[3] + mp.cnts[2] + mp.cnts[1] != 0)
                                            {
                                                c1 += mp.cnts[1];
                                                c2 += mp.cnts[2];
                                                c3 += mp.cnts[3];
                                                ep.SetText(10 + row, 3, mp.Name);
                                                ep.SetText(10 + row, 4, (mp.cnts[3] - mp.cnts[1] + mp.cnts[2]).ToString());
                                                ep.SetText(10 + row, 5, mp.cnts[1].ToString());
                                                ep.SetText(10 + row, 6, mp.cnts[2].ToString());
                                                ep.SetText(10 + row, 7, mp.cnts[3].ToString());
                                                row++;
                                            }
                                        }
                                        if ((c1 + c2 + c3) > 0)
                                        {
                                            c11 += c1; c22 += c2; c33 += c3;
                                            ep.SetText(10 + row, 3, "MasterCard (итого):");
                                            ep.SetText(10 + row, 4, (c3 - c1 + c2).ToString());
                                            ep.SetText(10 + row, 5, c1.ToString());
                                            ep.SetText(10 + row, 6, c2.ToString());
                                            ep.SetText(10 + row, 7, c3.ToString());
                                            ep.SetRangeBold(10 + row, 1, 10 + row, 7);
                                            row++;
                                        }

                                        //подсчет виза
                                        c1 = 0; c2 = 0; c3 = 0;
                                        foreach (int key in prods.Keys)
                                        {
                                            MyProd mp = (MyProd)prods[key];
                                            if (BranchStore.codeFromTypeAndProdName(mp.Type, mp.Name) != BaseProductType.VisaCard)
                                                continue;
                                            if (mp.cnts[3] + mp.cnts[2] + mp.cnts[1] != 0)
                                            {
                                                c1 += mp.cnts[1];
                                                c2 += mp.cnts[2];
                                                c3 += mp.cnts[3];
                                                ep.SetText(10 + row, 3, mp.Name);
                                                ep.SetText(10 + row, 4, (mp.cnts[3] - mp.cnts[1] + mp.cnts[2]).ToString());
                                                ep.SetText(10 + row, 5, mp.cnts[1].ToString());
                                                ep.SetText(10 + row, 6, mp.cnts[2].ToString());
                                                ep.SetText(10 + row, 7, mp.cnts[3].ToString());
                                                row++;
                                            }
                                        }
                                        if ((c1 + c2 + c3) > 0)
                                        {
                                            c11 += c1; c22 += c2; c33 += c3;
                                            ep.SetText(10 + row, 3, "Visa (итого):");
                                            ep.SetText(10 + row, 4, (c3 - c1 + c2).ToString());
                                            ep.SetText(10 + row, 5, c1.ToString());
                                            ep.SetText(10 + row, 6, c2.ToString());
                                            ep.SetText(10 + row, 7, c3.ToString());
                                            ep.SetRangeBold(10 + row, 1, 10 + row, 7);
                                            row++;
                                        }
                                        //подсчет нспк
                                        c1 = 0; c2 = 0; c3 = 0;
                                        foreach (int key in prods.Keys)
                                        {
                                            MyProd mp = (MyProd)prods[key];
                                            if (BranchStore.codeFromTypeAndProdName(mp.Type, mp.Name) != BaseProductType.MirCard)
                                                continue;
                                            if (mp.cnts[3] + mp.cnts[2] + mp.cnts[1] != 0)
                                            {
                                                c1 += mp.cnts[1];
                                                c2 += mp.cnts[2];
                                                c3 += mp.cnts[3];
                                                ep.SetText(10 + row, 3, mp.Name);
                                                ep.SetText(10 + row, 4, (mp.cnts[3] - mp.cnts[1] + mp.cnts[2]).ToString());
                                                ep.SetText(10 + row, 5, mp.cnts[1].ToString());
                                                ep.SetText(10 + row, 6, mp.cnts[2].ToString());
                                                ep.SetText(10 + row, 7, mp.cnts[3].ToString());
                                                row++;
                                            }
                                        }

                                        if ((c1 + c2 + c3) > 0)
                                        {
                                            c11 += c1; c22 += c2; c33 += c3;
                                            ep.SetText(10 + row, 3, "Карты МИР (итого):");
                                            ep.SetText(10 + row, 4, (c3 - c1 + c2).ToString());
                                            ep.SetText(10 + row, 5, c1.ToString());
                                            ep.SetText(10 + row, 6, c2.ToString());
                                            ep.SetText(10 + row, 7, c3.ToString());
                                            ep.SetRangeBold(10 + row, 1, 10 + row, 7);
                                            row++;
                                        }
                                        //подсчет nfc
                                        c1 = 0; c2 = 0; c3 = 0;
                                        foreach (int key in prods.Keys)
                                        {
                                            MyProd mp = (MyProd)prods[key];
                                            if (BranchStore.codeFromTypeAndProdName(mp.Type, mp.Name) != BaseProductType.NFCCard)
                                                continue;
                                            if (mp.cnts[3] + mp.cnts[2] + mp.cnts[1] != 0)
                                            {
                                                c1 += mp.cnts[1];
                                                c2 += mp.cnts[2];
                                                c3 += mp.cnts[3];
                                                ep.SetText(10 + row, 3, mp.Name);
                                                ep.SetText(10 + row, 4, (mp.cnts[3] - mp.cnts[1] + mp.cnts[2]).ToString());
                                                ep.SetText(10 + row, 5, mp.cnts[1].ToString());
                                                ep.SetText(10 + row, 6, mp.cnts[2].ToString());
                                                ep.SetText(10 + row, 7, mp.cnts[3].ToString());
                                                row++;
                                            }
                                        }

                                        if ((c1 + c2 + c3) > 0)
                                        {
                                            c11 += c1; c22 += c2; c33 += c3;
                                            ep.SetText(10 + row, 3, "Карты NFC (итого):");
                                            ep.SetText(10 + row, 4, (c3 - c1 + c2).ToString());
                                            ep.SetText(10 + row, 5, c1.ToString());
                                            ep.SetText(10 + row, 6, c2.ToString());
                                            ep.SetText(10 + row, 7, c3.ToString());
                                            ep.SetRangeBold(10 + row, 1, 10 + row, 7);
                                            row++;
                                        }
                                        //подсчет сервисных
                                        c1 = 0; c2 = 0; c3 = 0;
                                        foreach (int key in prods.Keys)
                                        {
                                            MyProd mp = (MyProd)prods[key];
                                            if (BranchStore.codeFromTypeAndProdName(mp.Type, mp.Name) != BaseProductType.ServiceCard)
                                                continue;
                                            if (mp.cnts[3] + mp.cnts[2] + mp.cnts[1] != 0)
                                            {
                                                c1 += mp.cnts[1];
                                                c2 += mp.cnts[2];
                                                c3 += mp.cnts[3];
                                                ep.SetText(10 + row, 3, mp.Name);
                                                ep.SetText(10 + row, 4, (mp.cnts[3] - mp.cnts[1] + mp.cnts[2]).ToString());
                                                ep.SetText(10 + row, 5, mp.cnts[1].ToString());
                                                ep.SetText(10 + row, 6, mp.cnts[2].ToString());
                                                ep.SetText(10 + row, 7, mp.cnts[3].ToString());
                                                row++;
                                            }
                                        }
                                        if ((c1 + c2 + c3) > 0)
                                        {
                                            c11 += c1; c22 += c2; c33 += c3;
                                            ep.SetText(10 + row, 3, "Сервисные карт (итого):");
                                            ep.SetText(10 + row, 4, (c3 - c1 + c2).ToString());
                                            ep.SetText(10 + row, 5, c1.ToString());
                                            ep.SetText(10 + row, 6, c2.ToString());
                                            ep.SetText(10 + row, 7, c3.ToString());
                                            ep.SetRangeBold(10 + row, 1, 10 + row, 7);
                                            row++;
                                        }
                                        ep.SetText(10 + row, 3, "Все карты (итого):");
                                        ep.SetText(10 + row, 4, (c33 - c11 + c22).ToString());
                                        ep.SetText(10 + row, 5, c11.ToString());
                                        ep.SetText(10 + row, 6, c22.ToString());
                                        ep.SetText(10 + row, 7, c33.ToString());
                                        ep.SetRangeBold(10 + row, 1, 10 + row, 7);
                                        row++;
                                        //пины
                                        foreach (int key in prods.Keys)
                                        {
                                            MyProd mp = (MyProd)prods[key];
                                            if (BranchStore.codeFromTypeAndProdName(mp.Type, mp.Name) != BaseProductType.PinConvert)
                                                continue;
                                            if (mp.cnts[3] + mp.cnts[2] + mp.cnts[1] != 0)
                                            {
                                                ep.SetText(10 + row, 3, mp.Name);
                                                ep.SetText(10 + row, 4, (mp.cnts[3] - mp.cnts[1] + mp.cnts[2]).ToString());
                                                ep.SetText(10 + row, 5, mp.cnts[1].ToString());
                                                ep.SetText(10 + row, 6, mp.cnts[2].ToString());
                                                ep.SetText(10 + row, 7, mp.cnts[3].ToString());
                                                ep.SetRangeBold(10 + row, 1, 10 + row, 7);
                                                row++;
                                            }
                                        }
                                        ep.SetRangeBorders(10, 1, 10 + row - 1, 7);
                                        ep.SetText(10 + row + 3, 3, "должность");
                                        ep.SetText(10 + row + 3, 4, "фамилия и инициалы");
                                        ep.SetText(10 + row + 3, 6, "подпись");
                                        ep.SetRangeAlignment(10 + row + 3, 1, 10 + row + 3, 7, Microsoft.Office.Interop.Excel.Constants.xlLeft);
                                        ep.SetRangeBottom(10 + row + 2, 3, 10 + row + 2, 6);
                                        ep.SetText(10 + row + 5, 3, "должность");
                                        ep.SetText(10 + row + 5, 4, "фамилия и инициалы");
                                        ep.SetText(10 + row + 5, 6, "подпись");
                                        ep.SetRangeAlignment(10 + row + 5, 1, 10 + row + 5, 7, Microsoft.Office.Interop.Excel.Constants.xlLeft);
                                        ep.SetRangeBottom(10 + row + 4, 3, 10 + row + 4, 6);

                                    }
                                    else // либо актуализируем кол-во на конец дня
                                    {
                                        foreach (int key in prods.Keys)
                                        {
                                            ((MyProd)prods[key]).cnts[3] =
                                                ((MyProd)prods[key]).cnts[3] - ((MyProd)prods[key]).cnts[1] +
                                                ((MyProd)prods[key]).cnts[2];
                                            ((MyProd)prods[key]).cnts[1] = 0;
                                            ((MyProd)prods[key]).cnts[2] = 0;
                                        }
                                    }
                                    index_date = index_date.AddDays(-1);
                                }
                                #endregion
                                ep.SetText("name1", branch_name + " (ГОЗ)");
                                ep.SetText("date1", $"{repdate:dd.MM.yyyy}");
                                conn.Close();
                            }
                        }
                        #endregion
                        if (make) // чтобы не сохранять движение ценностей архивный
                        {
                            if (doc.Length > 0 && WebConfigurationManager.AppSettings["DocPath"] != null)
                            {
                                doc = String.Format("{0}Temp\\{1}", WebConfigurationManager.AppSettings["DocPath"], doc);
                                ep.SaveAsDoc(doc, false);
                                if (make && (id_type == 13 || id_type == 15) && fname.Length > 0)
                                {
                                    File.Copy(doc, fname, true);
                                }
                            }
                        }
                    }
                    catch (Exception eg)
                    {
                        ep.Close();
                        GC.Collect();
                        System.Threading.Thread.CurrentThread.CurrentCulture = oldCI;
                        String m = eg.Message;
                        WebLog.LogClass.WriteToLog(eg.Message);
                        WebLog.LogClass.WriteToLog(eg.StackTrace);
                        
                        m = m.Replace("\n", "");
                        m = m.Replace("\r", " ");
                        m = m.Replace("\t", "");
                        ClientScript.RegisterClientScriptBlock(GetType(), "err43", "<script type='text/javascript'>$(document).ready(function(){ ShowError('" + m + "');});</script>");
                        return;
                    }
                }
                ep.Close();
                GC.Collect();
                System.Threading.Thread.CurrentThread.CurrentCulture = oldCI;
                if (make && doc.Length > 0)
                {
                    Database2.Log(sc.UserGuid(User.Identity.Name), String.Format("Сформирован отчет '{0}'", dListDoc.SelectedItem.Text), null);
                    ep.ReturnXls(Response, doc);
                    return;
                }
                if (!make && fname.Length > 0)
                {
                    Database2.Log(sc.UserGuid(User.Identity.Name), String.Format("Показан архивный отчет '{0}' {1:dd.MM.yyyy}", dListDoc.SelectedItem.Text, DatePickerOne.SelectedDate), null);
                    ep.ReturnXls(Response, fname);
                }
            }
        }
        protected void dListDoc_SelectedIndexChanged(object sender, EventArgs e)
        {
            lbInform.Text = "";
            PanelDinamyc.Visible = false;
            PanelTreasures.Visible = false;
            PanelOnePeriod.Visible = false;
            PanelTwoPeriod.Visible = false;
            PanelPersons.Visible = false;
            bEMail.Visible = false;
            bExportDBF.Visible = false;
            bExcel.Visible = true;
            if (dListDoc.SelectedItem.Value == "3")
                PanelTwoPeriod.Visible = true;
            if (dListDoc.SelectedItem.Value == "4")
            {
                PanelTwoPeriod.Visible = true;
                PanelDinamyc.Visible = true;
            }
            if (dListDoc.SelectedItem.Value == "5")
            {
                PanelTwoPeriod.Visible = true;
                PanelTreasures.Visible = true;
            }
            if (dListDoc.SelectedItem.Value == "6")
            { }
            if (dListDoc.SelectedItem.Value == "7")
            { }
            if (dListDoc.SelectedItem.Value == "8")
                PanelTwoPeriod.Visible = true;
            if (dListDoc.SelectedItem.Value == "9")
                PanelTwoPeriod.Visible = true;
            if (dListDoc.SelectedItem.Value == "10")
                PanelTwoPeriod.Visible = true;
            if (dListDoc.SelectedItem.Value == "11")
                PanelTwoPeriod.Visible = true;
            if (dListDoc.SelectedItem.Value == "12")
                PanelOnePeriod.Visible = true;
            if (dListDoc.SelectedItem.Value == "13")
            {
                PanelOnePeriod.Visible = true;
                bEMail.Visible = true;
            }
            if (dListDoc.SelectedItem.Value == "14")
                PanelTwoPeriod.Visible = true;
            if (dListDoc.SelectedItem.Value == "15")
            {
                PanelOnePeriod.Visible = true;
                bEMail.Visible = true;
                bExportDBF.Visible = true;
            }
            if (dListDoc.SelectedItem.Value == "16")
            {
                PanelOnePeriod.Visible = true;
                bEMail.Visible = true;
            }
            if (dListDoc.SelectedItem.Value == "17")
            {
                PanelOnePeriod.Visible = true;
                PanelTreasures.Visible = true;
                bEMail.Visible = true;
            }

            if (dListDoc.SelectedItem.Value == "42")
            {
                PanelOnePeriod.Visible = true;
            }
            if (dListDoc.SelectedItem.Value == "19")
            {
                PanelOnePeriod.Visible = true;
            }
            if (dListDoc.SelectedItem.Value == "20")
            {
                PanelOnePeriod.Visible = true;
            }
            if (dListDoc.SelectedItem.Value == "43")
            {
                PanelTwoPeriod.Visible = true;
            }
            if(dListDoc.SelectedItem.Value == "142")
            {
                //PanelOnePeriod.Visible = true;
                PanelTwoPeriod.Visible = true;
                bExportDBF.Visible = true;
                bExcel.Visible = false;
            }
            if (dListDoc.SelectedItem.Value == "41")
            {
                PanelOnePeriod.Visible = true;
            }
            if (dListDoc.SelectedItem.Value == "44")
            {
                PanelTwoPeriod.Visible = true;
                bEMail.Visible = true;
            }
            if (dListDoc.SelectedItem.Value == "45")
                PanelOnePeriod.Visible = true;
            //движение карт по 124 книге
            if (dListDoc.SelectedItem.Value == "46")
            {
                PanelOnePeriod.Visible = true;
            }
            //ведомость подотчет
            if (dListDoc.SelectedItem.Value == "47")
            {
                PanelTwoPeriod.Visible = true;
                PanelPersons.Visible = true;
            }
            //отчет завкассы или гоз
            if (dListDoc.SelectedItem.Value == "48"
                || dListDoc.SelectedItem.Value == "49")
            {
                PanelOnePeriod.Visible = true;
            }
        }
        protected void DatePickerOne_OnSelectedDateChanged(object sender, EventArgs e)
        {
            lbInform.Text = "";
            if (DateTime.Parse(DatePickerOne.SelectedDate.ToString()) == DateTime.Now)
                lDatePickerOneError.Visible = true;
            oneDateMessage = "";
            if (dListDoc.SelectedIndex >= 0 && dListDoc.SelectedItem.Value == "42")
            {
                OperationDay op = new OperationDay();
                op.read(current_branch);
                oneDateMessage = op.getMessage(DatePickerOne.SelectedDate);
                isTwoDay = op.isShift;
                allrangeradio.Text = "две смены"; //op.getMessage(DatePickerOne.SelectedDate);
                firstrangeradio.Text = op.getMessagePart(DatePickerOne.SelectedDate, true);
                nextrangeradio.Text = op.getMessagePart(DatePickerOne.SelectedDate, false);
            }
        }
        protected void bOkDirDBF_click(object sender, EventArgs e)
        {
            lbInform.Text = "";
            String fn = dirDBF.Text;
            if (fn.Length < 1) return;
            try
            {
                dirDBF.Text = "";
                System.IO.FileInfo f = new System.IO.FileInfo(fn);
                Response.ClearHeaders();
                Response.ClearContent();
                Response.HeaderEncoding = System.Text.Encoding.Default;
                Response.AddHeader("Content-Disposition", "attachment; filename=" + f.Name);
                Response.AddHeader("Content-Length", f.Length.ToString());
                Response.ContentType = "application/zip";
                Response.Cache.SetCacheability(HttpCacheability.NoCache);
                Response.BufferOutput = true;

                byte[] b = new byte[f.Length];
                FileStream fs = f.OpenRead();
                fs.Read(b, 0, b.Length);
                fs.Close();
                fs = null;
                f = null;
                Response.BinaryWrite(b);
                Response.Flush();
                b = null;

                //Response.BinaryWrite(File.ReadAllBytes(f.FullName));

                Response.End();
            }
            catch (Exception) { }
            finally
            {
                //FileHelper.DeleteFiles(fn, true);
            }            
        }
        private String getExcludeMemoricNameProducts()
        {
            String sqlProducts = "";
            if (ConfigurationSettings.AppSettings["ExcludeMemoricProducts"] == null) return "";
            DataSet ds = new DataSet();
            String[] Products = ConfigurationSettings.AppSettings["ExcludeMemoricProducts"].Split(';');
            for (int i = 0; i < Products.Length; i++)
            {
                String[] bin_prefix = Products[i].Split('_');
                if (sqlProducts.Length > 0) sqlProducts += " or ";
                sqlProducts += "(bin='" + bin_prefix[0] + "' and prefix_ow='" + bin_prefix[1] + "')";
            }
            ds.Clear();
            if (sqlProducts.Length > 0)
                Database2.ExecuteQuery("SELECT p.name as prbname FROM Products p join Products_Banks pb on p.id=pb.id_prod where " + sqlProducts, ref ds, null);
            String id_prb = "";
            for (int i = 0; ds.Tables.Count > 0 && i < ds.Tables[0].Rows.Count; i++)
            {
                if (id_prb.Length > 0) id_prb += ", ";
                id_prb += ds.Tables[0].Rows[i]["prbname"].ToString();
            }
            ds = null;
            if (id_prb.Length > 0) return "Будут исключены продукты: " + id_prb;
            return id_prb;
        }
        private String getExcludeMemoricProducts()
        {
            String sqlProducts = "";
            if (ConfigurationSettings.AppSettings["ExcludeMemoricProducts"] == null) return "";
            DataSet ds = new DataSet();
            String[] Products = ConfigurationSettings.AppSettings["ExcludeMemoricProducts"].Split(';');
            for (int i = 0; i < Products.Length; i++)
            {
                String[] bin_prefix = Products[i].Split('_');
                if (sqlProducts.Length > 0) sqlProducts += " or ";
                sqlProducts += "(bin='" + bin_prefix[0] + "' and prefix_ow='" + bin_prefix[1] + "')";
            }
            ds.Clear();
            if (sqlProducts.Length > 0)
                Database2.ExecuteQuery("SELECT pb.id as id_prb FROM Products p join Products_Banks pb on p.id=pb.id_prod where " + sqlProducts, ref ds, null);
            String id_prb = "";
            for (int i = 0; ds.Tables.Count > 0 && i < ds.Tables[0].Rows.Count; i++)
            {
                if (id_prb.Length > 0) id_prb += ",";
                id_prb += (Convert.ToInt32(ds.Tables[0].Rows[i]["id_prb"])).ToString();
            }
            ds = null;
            return id_prb;
        }
        private int[] getExcludeExportDBF(String key)
        {
            if (ConfigurationSettings.AppSettings[key] == null) return null;
            String[] Branchs = ConfigurationSettings.AppSettings[key].Split(';');
            if (Branchs.Length > 0)
            {
                DataSet ds = new DataSet();
                String id_branchs = "";
                for (int i = 0; i < Branchs.Length; i++)
                {
                    if (id_branchs.Length > 0) id_branchs += " or ";
                    id_branchs += "ident_dep='" + Branchs[i] + "'";
                }
                Database2.ExecuteQuery("SELECT id FROM Branchs where " + id_branchs, ref ds, null);
                if (ds.Tables.Count > 0 && ds.Tables[0].Rows.Count > 0)
                {
                    int[] id = new int[ds.Tables[0].Rows.Count];
                    for (int i = 0; i < ds.Tables[0].Rows.Count; i++)
                    {
                        id[i]=Convert.ToInt32(ds.Tables[0].Rows[i]["id"]);
                    }
                    return id;
                }
                ds = null;
            }
            return null;
        }
        private void exportDbf()
        {
            dirDBF.Text = "";
            zipArch = "eq_1_vneb20.zip";
            /*
            try
            {
                Convert.ToDateTime(DatePickerOne.DatePickerText);
                lDatePickerOneError.Visible = false;
                oneDateMessage = "";
            }
            catch
            {
                lDatePickerOneError.Visible = true;
                return;
            }
            */
            try
            {
                DateTime.ParseExact(DatePickerStart.DatePickerText, "dd.MM.yyyy", null);
                DateTime.ParseExact(DatePickerEnd.DatePickerText, "dd.MM.yyyy", null);
                //Convert.ToDateTime(DatePickerStart.DatePickerText);
                //Convert.ToDateTime(DatePickerEnd.DatePickerText);
                lDatePickerTwoError.Visible = false;
            }
            catch
            {
                lDatePickerTwoError.Visible = true;
                return;
            }

            bool isOpenDbf = false;
            DataSet ds = new DataSet();
            SqlCommand comm = Database2.Conn.CreateCommand();

            ds.Clear();

            DateTime sDate = DatePickerOne.SelectedDate.Date;


            ArrayList al = new ArrayList();
            ArrayList br = new ArrayList();
            ArrayList br_prev = new ArrayList();

            bool isError = false;
            bool isEmpty = true;

            DbfFile odbf = new DbfFile(System.Text.Encoding.GetEncoding(866));
            DbfRecord orec = null;

            int[] id_exclude_sendclient = null;
            int[] id_exclude_branch = null;

            int idUser = sc.UserId(User.Identity.Name);
            int idBranch = sc.BranchId(User.Identity.Name);
            String sqlUserV = " " + idUser.ToString() + "=(select [user_id] from StorageDocs where V_Cards_StorageDocs1.id=StorageDocs.id) and ";
            String sqlUser = " " + idUser.ToString() + "=[user_id] and ";


            //if (idUser < 1 || userradio.Checked == false)
            //{
            sqlUserV = " ";
            sqlUser = " ";
            //}

            OperationDay op = new OperationDay();

            try
            {
                DateTime dtS = op.roundDateTime(DatePickerStart.SelectedDate.Date);
                DateTime dtE = op.roundDateTime(DatePickerEnd.SelectedDate.Date);
                if (dtS > dtE) throw new Exception("Неверно задан период.");
                TimeSpan left = dtE - dtS;
                int dleft = left.Days + 1;
                if (dleft > 14) throw new Exception("Период должен быть не более 14 дней.");
                for (sDate = dtS; sDate <= dtE; sDate = sDate.AddDays(1))
                {
                    lock (Database2.lockObjectDB)
                    {
                        if (br.Count == 0)
                        {
                            id_exclude_sendclient = getExcludeExportDBF("ExcludeExportDBFSendClent");
                            id_exclude_branch = getExcludeExportDBF("ExcludeExportDBFBranch");
                            string id_prbs = getExcludeMemoricProducts();
                            if (id_prbs.Length < 1)
                                Database2.ExecuteQuery("select id_prb, id_prod, prod_name, id_type from V_ProductsBanks_T where id_type in(1,2) order by id_type, id_sort", ref ds, null);
                            else
                                Database2.ExecuteQuery("select id_prb, id_prod, prod_name, id_type from V_ProductsBanks_T where id_type in(1,2) and id_prb not in (" + id_prbs + ") order by id_type, id_sort", ref ds, null);
                            if (ds.Tables.Count > 0)
                            {
                                for (int i = 0; i < ds.Tables[0].Rows.Count; i++)
                                {
                                    MyProd mp = new MyProd(Convert.ToInt32(ds.Tables[0].Rows[i]["id_prb"]),
                                        Convert.ToString(ds.Tables[0].Rows[i]["prod_name"]), "", Convert.ToInt32(ds.Tables[0].Rows[i]["id_type"]));
                                    al.Add(mp);
                                }
                            }
                            if (al.Count < 1) throw new Exception("Продукты для экспорта не найдены.");
                            ds.Clear();
                            comm.Parameters.Clear();
                            if (branchradio.Checked == true)
                            {
                                comm.CommandText = "select * from Branchs where id=@branch or id_parent=@branch order by id";
                                //comm.Parameters.Add("@branch", SqlDbType.Int).Value = 106;
                                comm.Parameters.Add("@branch", SqlDbType.Int).Value = branchIdMain;
                            }
                            else
                            {
                                comm.CommandText = "select * from Branchs where id=@branch";
                                comm.Parameters.Add("@branch", SqlDbType.Int).Value = idBranch;
                            }
                            Database2.ExecuteCommand(comm, ref ds, null);
                            if (ds.Tables.Count > 0 && ds.Tables[0].Rows.Count > 0)
                            {
                                for (int i = 0; i < ds.Tables[0].Rows.Count; i++)
                                {
                                    int id = Convert.ToInt32(ds.Tables[0].Rows[i]["id"]);
                                    if (id_exclude_branch != null && Array.IndexOf(id_exclude_branch, id) >= 0) continue;
                                    BranchStore branch = new BranchStore(ds.Tables[0].Rows[i]["id"], ds.Tables[0].Rows[i]["ident_dep"], ds.Tables[0].Rows[i]["department"]);
                                    br.Add(branch);
                                    BranchStore branch_prev = new BranchStore(branch.id, branch.ident_dep, branch.department);
                                    br_prev.Add(branch_prev);
                                }
                            }
                            else throw new Exception("Подразделения для экспорта не найдены.");
                        }
                        else
                        {
                            for (int i = 0; i < br.Count; i++)
                            {
                                ((BranchStore)br[i]).Clear();
                                ((BranchStore)br_prev[i]).Clear();
                            }
                        }

                        isEmpty = true;

                        for (int i = 0; i < br.Count; i++)
                        {
                            BranchStore branch = (BranchStore)br[i];
                            BranchStore branch_prev = (BranchStore)br_prev[i];
                            op.read(branch.id);
                            DateTime curDateS = op.getDateTimeStart(sDate);
                            DateTime curDateE = op.getDateTimeEnd(sDate);

                            if (op.isShift)
                            {
                                DateTime[] tm = op.getDateTimePart(sDate, false);
                                if (tm != null)
                                {
                                    curDateS = tm[0];
                                    curDateE = tm[1];
                                    tm = null;
                                }
                                tm = op.getDateTimePart(sDate, true);
                                if (tm != null)
                                {
                                    setMemoric(branch_prev, al, tm[0], tm[1], sqlUserV, sqlUser);
                                    tm = null;
                                }
                            }

                            setMemoric(branch, al, curDateS, curDateE, sqlUserV, sqlUser);


                            if (branch.isEmpty() == false || branch_prev.isEmpty() == false) isEmpty = false;
                        }
                    }


                    // формирование

                    if (isEmpty == false && isOpenDbf == false)
                    {
                        FileHelper.initDefault();
                        FileHelper.SetDirectoryApp(ConfigurationSettings.AppSettings["DocPath"]);
                        FileHelper.SetDataFolder("temp\\expdbf\\" + Path.GetFileNameWithoutExtension(Path.GetRandomFileName()), true);
                        FileHelper.setCode866(fileDBF);
                        //DataBaseHelper.ExecuteQuery("delete from " + tableName, null);
                        odbf.Open(FileHelper.getFullName(fileDBF), FileMode.Open);
                        orec = new DbfRecord(odbf.Header);
                        for (int i = 0; i < odbf.Header.RecordCount; i++)
                        {
                            if (!odbf.Read(i, orec)) break;
                            orec.IsDeleted = true;
                            odbf.Write(orec);
                        }
                        isOpenDbf = true;

                        //string isTrans = "99999";
                        //string trans = "Транзитный счет";
                        //string sCurDate = sDate.ToString("dd.MM.yyyy");
                        //string mquery = "INSERT INTO " + tableName + 
                        //                " (MFO_PAYER,ACC_PAYER,MFO_REC,ACC_REC,[SUM],SUMDEB,SUMKRED,CURDEB,CURKRED,SROK,DATE_CARRY,DATE_DOCUM,NUMB_DOCUM,KIND_OPER,GROUND) " +
                        //                " VALUES ('049205805',@acc_payer,'049205805',@acc_rec,@sum,@sumdeb,@sumkred,'810','810',@srok,@date_carry,@date_docum,@numb_docum,'09',@ground)";

                        //OleDbParameter[] mpr = new OleDbParameter[10];
                    }

                    int numb_docum = 0;

                    for (int i = 0; i < br.Count; i++)
                    {
                        BranchStore branch = (BranchStore)br[i];
                        BranchStore branch_prev = (BranchStore)br_prev[i];
                        //OperationDay op = new OperationDay();
                        op.read(branch.id);
                        //string numb_prefix = OperationDay.getBranchStartNumber(branch.id,sDate,true); // .getBranchStartNumber(branch.id);
                        int smena = 1;
                        //if (numb_docum < 1) 
                        numb_docum = 1;
                        if (branch.isEmpty() == true && branch_prev.isEmpty() == true) continue;
                        isEmpty = false;
                        List<AccountBranch> listAB = getAccountBranch(branch.id);
                        for (int k = 0; k < 2; k++)
                        {
                            BranchStore bs;
                            if (op.isShift == false) k = 1;
                            if (k == 0) bs = branch_prev;
                            else bs = branch;
                            if (op.isShift == true && k == 1)
                            {
                                //numb_prefix = OperationDay.getBranchStartNumber(branch.id,sDate,false);
                                smena = 2;
                                //if (numb_docum < 1) 
                                numb_docum = 1;
                                //numb_docum += OperationDay.getShiftNumber();
                            }

                            int maxN = Enum.GetNames(typeof(BaseProductType)).Length - 1;
                            int maxE = maxN * 3;

                            for (int n = 0; n < maxE; n++, numb_docum++)
                            {
                                int countP = 0;
                                String nameP = "";
                                String nameS = "";
                                BaseProductType bp = BaseProductType.None;
                                AccountBranchType ab = AccountBranchType.None;
                                int abt = n / maxN;
                                if (n % maxN == 0) { countP = bs.countMasterCard[abt]; nameS = "MC"; nameP = "ПК Master Card"; bp = BaseProductType.MasterCard; }
                                if (n % maxN == 1) { countP = bs.countVisaCard[abt]; nameS = "VISA"; nameP = "ПК Visa"; bp = BaseProductType.VisaCard; }
                                if (n % maxN == 2) { countP = bs.countNFCCard[abt]; nameS = "NFC"; nameP = "ПК NFC"; bp = BaseProductType.NFCCard; }
                                if (n % maxN == 3) { countP = bs.countServiceCard[abt]; nameS = "Service"; nameP = "ПК Сервисные"; bp = BaseProductType.ServiceCard; }
                                if (n % maxN == 4) { countP = bs.countMirCard[abt]; nameS = "MIR"; nameP = "ПК MIR"; bp = BaseProductType.MirCard; }
                                if (n % maxN == 5) { countP = bs.countPinConvert[abt]; nameS = "PIN"; nameP = "Бланки ПИН-конвертов"; bp = BaseProductType.PinConvert; }
                                if (abt == 0) ab = AccountBranchType.In;
                                if (abt == 1) ab = AccountBranchType.Out;
                                if (abt == 2) ab = AccountBranchType.Return;
                                if (countP < 1) continue;
                                //numb_docum++;
                                string ground = "";
                                string txt = "W " + nameP + " согласно акта " + sDate.ToString("dd.MM.yyyy") + " Y хр-щZ Амирхана, 21";
                                if (ab == AccountBranchType.Out)
                                {
                                    txt = txt.Substring(0, txt.IndexOf("Y"));
                                    txt = txt.Replace("W", "Переданы");
                                    ground = txt.Trim();
                                }
                                if (ab == AccountBranchType.In)
                                {
                                    txt = txt.Replace("W", "Приняты");
                                    txt = txt.Replace("Y", "из");
                                    txt = txt.Replace("Z", "а");
                                    ground = txt.Trim();
                                }
                                if (ab == AccountBranchType.Return)
                                {
                                    txt = txt.Replace("W", "Инкассация");
                                    txt = txt.Replace("Y", "в");
                                    txt = txt.Replace("Z", "е");
                                    ground = txt.Trim();
                                }
                                AccountBranch AB = listAB.FirstOrDefault(ai => ai.accountType == ab && ai.productType == bp);
                                String debet = "";
                                String credit = "";
                                if (AB != null)
                                {
                                    debet = AB.accountDebet;
                                    credit = AB.accountCredit;
                                }
                                /*
                                mpr[0] = new OleDbParameter("@acc_payer", debet);
                                mpr[1] = new OleDbParameter("@acc_rec", credit);
                                mpr[2] = new OleDbParameter("@sum",countP);
                                mpr[3] = new OleDbParameter("@sumdeb",countP);
                                mpr[4] = new OleDbParameter("@sumkred",countP);
                                mpr[5] = new OleDbParameter("@srok",sDate);
                                mpr[6] = new OleDbParameter("@date_carry",sDate);
                                mpr[7] = new OleDbParameter("@date_docum",sDate);
                                mpr[8] = new OleDbParameter("@numb_docum",numb_docum);
                                mpr[9] = new OleDbParameter("@ground", ground);
                                DataBaseHelper.ExecuteQuery(mquery, mpr);
                                for (int j = 0; j < mpr.Length; j++) mpr[j] = null;
                                */
                                orec.Clear();
                                orec[orec.FindColumn("mfo_payer")] = "049205805";
                                orec[orec.FindColumn("acc_payer")] = debet;
                                orec[orec.FindColumn("mfo_rec")] = "049205805";
                                orec[orec.FindColumn("acc_rec")] = credit;
                                orec[orec.FindColumn("sum")] = String.Format("{0}.00", countP);
                                orec[orec.FindColumn("sumdeb")] = String.Format("{0}.00", countP);
                                orec[orec.FindColumn("sumkred")] = String.Format("{0}.00", countP);
                                orec[orec.FindColumn("curdeb")] = "810";
                                orec[orec.FindColumn("curkred")] = "810";
                                orec[orec.FindColumn("srok")] = sDate.ToString("yyyy-MM-dd");
                                orec[orec.FindColumn("date_carry")] = sDate.ToString("yyyy-MM-dd");
                                orec[orec.FindColumn("date_docum")] = sDate.ToString("yyyy-MM-dd");
                                orec[orec.FindColumn("numb_docum")] = OperationDay.getBranchStartNumber(branch.id, sDate, numb_docum, smena); //numb_prefix + numb_docum.ToString("D03"); //numb_docum.ToString(); 
                                orec[orec.FindColumn("kind_oper")] = "09";
                                orec[orec.FindColumn("ground")] = ground;
                                odbf.Write(orec);

                                if (ab == AccountBranchType.Out) numb_docum++;

                                if (ab == AccountBranchType.Out && (id_exclude_sendclient == null || (id_exclude_sendclient != null && Array.IndexOf(id_exclude_sendclient, branch.id) < 0)))
                                {
                                    orec.Clear();
                                    ground = "Выданы клиенту " + nameP + " согласно отчета " + sDate.ToString("dd.MM.yyyy");
                                    orec[orec.FindColumn("acc_rec")] = debet;

                                    string ident_dep = "";

                                    if (debet.Length >= 14) ident_dep = debet.Substring(9, 4);

                                    debet = "";
                                    credit = "";

                                    if (ident_dep.Length > 0)
                                    {
                                        List<AccountBranch> listab = getAccountBranch(ident_dep);
                                        AB = listab.FirstOrDefault(ai => ai.accountType == AccountBranchType.Return && ai.productType == bp);
                                        if (AB != null)
                                        {
                                            debet = AB.accountDebet;
                                            credit = AB.accountCredit;
                                        }
                                        listab = null;
                                    }

                                    orec[orec.FindColumn("mfo_payer")] = "049205805";
                                    orec[orec.FindColumn("acc_payer")] = debet;
                                    orec[orec.FindColumn("mfo_rec")] = "049205805";
                                    orec[orec.FindColumn("sum")] = String.Format("{0}.00", countP);
                                    orec[orec.FindColumn("sumdeb")] = String.Format("{0}.00", countP);
                                    orec[orec.FindColumn("sumkred")] = String.Format("{0}.00", countP);
                                    orec[orec.FindColumn("curdeb")] = "810";
                                    orec[orec.FindColumn("curkred")] = "810";
                                    orec[orec.FindColumn("srok")] = sDate.ToString("yyyy-MM-dd");
                                    orec[orec.FindColumn("date_carry")] = sDate.ToString("yyyy-MM-dd");
                                    orec[orec.FindColumn("date_docum")] = sDate.ToString("yyyy-MM-dd");
                                    orec[orec.FindColumn("numb_docum")] = OperationDay.getBranchStartNumber(branch.id, sDate, numb_docum, smena);//numb_prefix + numb_docum.ToString("D03"); //numb_docum.ToString();
                                    orec[orec.FindColumn("kind_oper")] = "09";
                                    orec[orec.FindColumn("ground")] = ground;
                                    odbf.Write(orec);
                                }

                            }
                        }
                        listAB = null;
                    }

                    //if (isEmpty == true) throw new Exception("На " + sDate.ToString("dd.MM.yyyy") + " нет операций по продуктам.");
                }
                if (isOpenDbf == false) throw new Exception("Операции по продуктам не найдены.");
            }
            catch (Exception e1)
            {
                isError = true;
                String m = e1.Message;
                m = m.Replace("\n", "");
                m = m.Replace("\r", " ");
                m = m.Replace("\t", "");

                ClientScript.RegisterClientScriptBlock(GetType(), "err142", "<script type='text/javascript'>$(document).ready(function(){ ShowError('" + m + "');});</script>");

            }
            finally
            {
                odbf.Close();
                if (isError == false)
                {
                    dirDBF.Text = FileHelper.GetDirectoryData() + zipArch;
                    //ClientScript.RegisterClientScriptBlock(GetType(), "ok142", "<script type='text/javascript'>$(document).ready(function(){ setTimeout(function(){$('#" + bOkDirDBF.ClientID +  "').click();},500); });</script>");
                    ClientScript.RegisterClientScriptBlock(GetType(), "ok142", "<script type='text/javascript'>$(document).ready(function(){ $('#" + bOkDirDBF.ClientID + "').click(); });</script>");
                    FileHelper.setCodeDefault(fileDBF);
                    using (var zip = new ZipFile())
                    {
                        zip.AddDirectory(FileHelper.GetDirectoryData());
                        zip.Save(FileHelper.getFullName(zipArch));
                    }

                    /*
                    System.IO.FileInfo f = new System.IO.FileInfo(FileHelper.getFullName(zipArch));
                    Response.ClearHeaders();
                    Response.ClearContent();
                    Response.HeaderEncoding = System.Text.Encoding.Default;
                    Response.AddHeader("Content-Disposition", "attachment; filename=" + f.Name);
                    Response.AddHeader("Content-Length", f.Length.ToString());
                    Response.ContentType = "application/zip";
                    Response.Cache.SetCacheability(HttpCacheability.NoCache);
                    Response.BufferOutput = true;
                    Response.WriteFile(f.FullName);
                    Response.End();
                    */


                }
                //FileHelper.DeleteFiles(FileHelper.GetDirectoryData(), false);
            }
            //GC.Collect();
        }

        
        protected void bExportDBF_Click(object sender, ImageClickEventArgs e)
        {
            lbInform.Text = "";
            current_branch = sc.BranchId(User.Identity.Name);
            int id_type = Convert.ToInt32(dListDoc.SelectedItem.Value);
            if (id_type == 15)
            {
                try
                {
                    //MovingStock(Convert.ToDateTime(DatePickerOne.DatePickerText));//MovingStock(new DateTime(2015, 6, 23));
                    MovingStock(DatePickerOne.SelectedDate.Date);
                    lDatePickerOneError.Visible = false;
                }
                catch
                {
                    lDatePickerOneError.Visible = true;
                    return;
                }
            }
            else
                exportDbf();
            
            
        }

        protected void bMail_Click(object sender, ImageClickEventArgs e)
        {
            lbInform.Text = "";
            lock (Database2.lockObjectDB)
            {
                string res = "";

                if (dListDoc.SelectedIndex < 0) return;
                try
                {
                    if (rbPeriods.SelectedValue == "2")
                        Convert.ToInt32(tbDays.Text);
                    lDaysWarning.Visible = false;
                }
                catch
                {
                    lDaysWarning.Visible = true;
                    return;
                }
                int id_type = Convert.ToInt32(dListDoc.SelectedItem.Value);
                if (id_type == 12 || id_type == 13 || id_type == 15 || id_type == 16 || id_type == 17 || id_type == 44)
                {
                    try
                    {
                        DateTime.ParseExact(DatePickerOne.DatePickerText, "dd.MM.yyyy", null);
                        //Convert.ToDateTime(DatePickerOne.DatePickerText);
                        lDatePickerOneError.Visible = false;
                    }
                    catch
                    {
                        lDatePickerOneError.Visible = true;
                        return;
                    }
                }
                if (id_type == 3 || id_type == 4 || id_type == 5 || id_type == 8 || id_type == 9 || id_type == 10 || id_type == 11)
                {
                    try
                    {
                        DateTime.ParseExact(DatePickerStart.DatePickerText, "dd.MM.yyyy", null);
                        DateTime.ParseExact(DatePickerEnd.DatePickerText, "dd.MM.yyyy", null);
                        //Convert.ToDateTime(DatePickerStart.DatePickerText);
                        //Convert.ToDateTime(DatePickerEnd.DatePickerText);
                        lDatePickerTwoError.Visible = false;
                    }
                    catch
                    {
                        lDatePickerTwoError.Visible = true;
                        return;
                    }
                }
                string doc = "";
                //            string dt_doc = "";

                if (id_type == 13) doc = "Attachment10.xls";
                if (id_type == 15) doc = "Attachment10_1.xls";
                if (id_type == 16) doc = "Attachment15.xls";
                if (id_type == 17) doc = "Attachment15_1.xls";
                if (id_type == 44) doc = "Attachment44.xls";
                if (doc.Length == 0)
                    return;


                System.Globalization.CultureInfo oldCI = System.Threading.Thread.CurrentThread.CurrentCulture;
                System.Threading.Thread.CurrentThread.CurrentCulture = new System.Globalization.CultureInfo("en-US");

                ep = new ExcelAp();
                //            string branchStr = "";
                //            string branchName = "";
                bool make = true; // для движения ценностей - отображение архивного
                string fname = ""; // для движения ценностей - отображение архивного
                int i = 0;

                if (ep.RunApp(ConfigurationSettings.AppSettings["DocPath"] + doc))
                {
                    ep.SaveAsDoc(ConfigurationSettings.AppSettings["DocPath"] + "Temp/" + doc, false);
                    ep.SetWorkSheet(1);

                    DataSet ds = new DataSet();
                    #region Распоряжение на движение ценностей (старый вариант)
                    if (id_type == 13)
                    {
                        fname = String.Format("{0}RP_{1:ddMMyyyy}.xls", ConfigurationSettings.AppSettings["ArchivePath"], DatePickerOne.SelectedDate);
                        if (DatePickerOne.SelectedDate.Date != DateTime.Now.Date)
                        {
                            if (!File.Exists(fname))
                            {
                                fname = "";
                                lInfo.Text = "В архиве нет отчета за данное число. Состояние хранилища будет на текущую дату.";
                            }
                            else
                                make = false;
                        }
                        if (make) // не делать если есть архивный
                        {
                            SqlCommand comm = new SqlCommand();
                            //проверяем, что есть неподтвержденные документы на персонализацию или отправку в филиал
                            comm.CommandText = "select number_doc from StorageDocs where date_doc=@d and priz_gen=0 and (type=5 or type=8)";
                            comm.Parameters.Add("@d", SqlDbType.DateTime).Value = DatePickerOne.SelectedDate.Date;
                            Database2.ExecuteCommand(comm, ref ds, null);
                            if (ds.Tables.Count > 0 && ds.Tables[0].Rows.Count > 0)
                            {
                                string str = "";
                                for (i = 0; i < ds.Tables[0].Rows.Count; i++)
                                    str += ", " + ds.Tables[0].Rows[i][0].ToString();
                                str = str.Substring(2);
                                lInfo.Text = String.Format("Создание отчета невозможно. За данное число есть неподтвержденные документы ({0})", str);
                                return;
                            }
                            ds.Clear();
                            comm.CommandText = "select distinct id_prb, prod_name, bank_name,id_sort from V_Rep_Moving where priz_gen=1 and type=@type and (id_type=1 or id_type=2) and date_doc=@dateDoc order by id_sort";
                            comm.Parameters.Add("@type", SqlDbType.Int).Value = (int)TypeDoc.PersoCard;
                            comm.Parameters.Add("@dateDoc", SqlDbType.DateTime).Value = DatePickerOne.SelectedDate.Date;
                            res = (string)Database2.ExecuteCommand(comm, ref ds, null);
                            ArrayList al = new ArrayList();
                            foreach (DataRow dr in ds.Tables[0].Rows)
                                al.Add(new MyProd(Convert.ToInt32(dr["id_prb"]), dr["prod_name"].ToString(), dr["bank_name"].ToString()));
                            comm.CommandText = "select sum(cnt_brak) as cntbrak, sum(cnt_perso) as cntperso from V_Rep_Moving where priz_gen=1 and type=@type and (id_type=1 or id_type=2) and date_doc=@dateDoc and id_prb=@prb";
                            comm.Parameters.Add("@prb", SqlDbType.Int).Value = 0;
                            i = 11;
                            foreach (MyProd mp in al)
                            {
                                ds.Clear();
                                comm.Parameters["@prb"].Value = mp.ID;
                                res = (string)Database2.ExecuteCommand(comm, ref ds, null);
                                if (ds.Tables[0].Rows.Count == 1)
                                {
                                    int br = (ds.Tables[0].Rows[0]["cntbrak"] == DBNull.Value) ? 0 : Convert.ToInt32(ds.Tables[0].Rows[0]["cntbrak"]);
                                    int pr = (ds.Tables[0].Rows[0]["cntperso"] == DBNull.Value) ? 0 : Convert.ToInt32(ds.Tables[0].Rows[0]["cntperso"]);
                                    ep.SetText(i, 2, String.Format("{0}({1})", mp.Name, mp.Bank));
                                    ep.SetText(i, 3, String.Format("{0}шт.", br + pr));
                                    i++;
                                }
                            }
                            ep.ShowRows(11, i);
                            al.Clear();
                            ds.Clear();
                            comm.CommandText = "select distinct id_prb, prod_name, bank_name,id_sort from V_Rep_Moving where priz_gen=1 and type=@type and (id_type=1 or id_type=2) and date_doc=@dateDoc order by id_sort";
                            comm.Parameters["@type"].Value = (int)TypeDoc.SendToFilial;
                            res = (string)Database2.ExecuteCommand(comm, ref ds, null);
                            foreach (DataRow dr in ds.Tables[0].Rows)
                                al.Add(new MyProd(Convert.ToInt32(dr["id_prb"]), dr["prod_name"].ToString(), dr["bank_name"].ToString()));
                            comm.CommandText =
                            comm.CommandText = "select sum(cnt_perso) as cntperso from V_Rep_Moving where priz_gen=1 and type=@type and (id_type=1 or id_type=2) and date_doc=@dateDoc and id_prb=@prb and is_head=@ishead";
                            comm.Parameters.Add("@ishead", SqlDbType.Bit).Value = true;
                            i = 63;
                            foreach (MyProd mp in al)
                            {
                                ds.Clear();
                                comm.Parameters["@prb"].Value = mp.ID;
                                res = (string)Database2.ExecuteCommand(comm, ref ds, null);
                                if (ds.Tables[0].Rows.Count == 1)
                                {
                                    int pr = (ds.Tables[0].Rows[0]["cntperso"] == DBNull.Value) ? 0 : Convert.ToInt32(ds.Tables[0].Rows[0]["cntperso"]);
                                    if (pr > 0)
                                    {
                                        ep.SetText(i, 2, String.Format("{0}({1})", mp.Name, mp.Bank));
                                        ep.SetText(i, 3, String.Format("{0}шт.", pr));
                                        i++;
                                    }
                                }
                            }
                            ep.ShowRows(63, i);
                            comm.CommandText = "select sum(cnt_perso) as cntperso from V_Rep_Moving where priz_gen=1 and type=@type and (id_type=1 or id_type=2) and date_doc=@dateDoc and id_prb=@prb and (is_head=@ishead or is_head is null) and (is_rkc=@ishead or is_rkc is null) and (is_trans=@ishead or is_trans is null)";
                            comm.Parameters["@ishead"].Value = false;
                            i = 116;
                            foreach (MyProd mp in al)
                            {
                                ds.Clear();
                                comm.Parameters["@prb"].Value = mp.ID;
                                res = (string)Database2.ExecuteCommand(comm, ref ds, null);
                                if (ds.Tables[0].Rows.Count == 1)
                                {
                                    int pr = (ds.Tables[0].Rows[0]["cntperso"] == DBNull.Value) ? 0 : Convert.ToInt32(ds.Tables[0].Rows[0]["cntperso"]);
                                    if (pr > 0)
                                    {
                                        ep.SetText(i, 2, String.Format("{0}({1})", mp.Name, mp.Bank));
                                        ep.SetText(i, 3, String.Format("{0}шт.", pr));
                                        i++;
                                    }
                                }
                            }
                            ep.ShowRows(116, i);
                            comm.CommandText = "select sum(cnt_perso) as cntperso from V_Rep_Moving where priz_gen=1 and type=@type and (id_type=1 or id_type=2) and date_doc=@dateDoc and id_prb=@prb and is_rkc=@ishead";
                            comm.Parameters["@ishead"].Value = true;
                            i = 169;
                            foreach (MyProd mp in al)
                            {
                                ds.Clear();
                                comm.Parameters["@prb"].Value = mp.ID;
                                res = (string)Database2.ExecuteCommand(comm, ref ds, null);
                                if (ds.Tables[0].Rows.Count == 1)
                                {
                                    int pr = (ds.Tables[0].Rows[0]["cntperso"] == DBNull.Value) ? 0 : Convert.ToInt32(ds.Tables[0].Rows[0]["cntperso"]);
                                    if (pr > 0)
                                    {
                                        ep.SetText(i, 2, String.Format("{0}({1})", mp.Name, mp.Bank));
                                        ep.SetText(i, 3, String.Format("{0}шт.", pr));
                                        i++;
                                    }
                                }
                            }
                            ep.ShowRows(169, i);
                            comm.CommandText = "select sum(cnt_perso) as cntperso from V_Rep_Moving where priz_gen=1 and type=@type and (id_type=1 or id_type=2) and date_doc=@dateDoc and id_prb=@prb and is_trans=@ishead";
                            comm.Parameters["@ishead"].Value = true;
                            i = 223;
                            foreach (MyProd mp in al)
                            {
                                ds.Clear();
                                comm.Parameters["@prb"].Value = mp.ID;
                                res = (string)Database2.ExecuteCommand(comm, ref ds, null);
                                if (ds.Tables[0].Rows.Count == 1)
                                {
                                    int pr = (ds.Tables[0].Rows[0]["cntperso"] == DBNull.Value) ? 0 : Convert.ToInt32(ds.Tables[0].Rows[0]["cntperso"]);
                                    if (pr > 0)
                                    {
                                        ep.SetText(i, 2, String.Format("{0}({1})", mp.Name, mp.Bank));
                                        ep.SetText(i, 3, String.Format("{0}шт.", pr));
                                        i++;
                                    }
                                }
                            }
                            ep.ShowRows(223, i);

                            ep.SetText_Name("date_string", String.Format("от {0:dd MM yyyy}", DatePickerOne.SelectedDate));
                            ep.SetText_Name("date_string1", String.Format("{0:dd.MM.yyyy}", DatePickerOne.SelectedDate));
                            /////////////////////////////////
                            //второй лист - хранилище
                            /////////////////////////////////
                            ds.Clear();
                            ep.SetWorkSheet(2);
                            Database2.ExecuteQuery("select name, bank_name, cnt_new, cnt_perso from V_Storage where id_type=1 order by id_sort", ref ds, null);
                            int rw = 4;
                            for (i = 0; i < ds.Tables[0].Rows.Count; i++)
                            {
                                string str = String.Format("{0} ({1})", ds.Tables[0].Rows[i]["name"].ToString(), ds.Tables[0].Rows[i]["bank_name"].ToString());
                                ep.SetText(rw, 2, str);
                                ep.SetText(rw, 3, ds.Tables[0].Rows[i]["cnt_new"].ToString());
                                ep.SetText(rw, 4, ds.Tables[0].Rows[i]["cnt_perso"].ToString());
                                rw++;
                            }
                            rw++;
                            ds.Clear();
                            Database2.ExecuteQuery("select name, bank_name, cnt_new, cnt_perso from V_Storage where id_type=2 order by id_sort", ref ds, null);
                            ep.SetText(rw, 3, "Бланки (шт)");
                            ep.SetText(rw, 4, "Персо");
                            rw++;
                            for (i = 0; i < ds.Tables[0].Rows.Count; i++)
                            {
                                string str = String.Format("{0} ({1})", ds.Tables[0].Rows[i]["name"].ToString(), ds.Tables[0].Rows[i]["bank_name"].ToString());
                                ep.SetText(rw, 2, str);
                                ep.SetText(rw, 3, ds.Tables[0].Rows[i]["cnt_new"].ToString());
                                ep.SetText(rw, 4, ds.Tables[0].Rows[i]["cnt_perso"].ToString());
                                rw++;
                            }
                            ds.Clear();
                            ep.SetRangeBorders(3, 2, rw - 1, 4);
                            ep.SetText_Name("date_string2", String.Format("на {0:dd.MM.yyyy}", DateTime.Now));
                        }
                    }
                    #endregion
                    #region Распоряжение на движение ценностей (новый вариант)
                    if (id_type == 15)
                    {
                        fname = String.Format("{0}RPn_{1:ddMMyyyy}.xls", ConfigurationSettings.AppSettings["ArchivePath"], DatePickerOne.SelectedDate);
                        if (DatePickerOne.SelectedDate.Date != DateTime.Now.Date)
                        {
                            if (!File.Exists(fname))
                            {
                                fname = "";
                                lInfo.Text = "В архиве нет отчета за данное число. Состояние хранилища будет на текущую дату.";
                            }
                            else
                                make = false;
                        }
                        if (make) // не делать если есть архивный
                        {
                            MovingStock(ep);
                            /*                        SqlCommand comm = new SqlCommand();
                                                    //проверяем, что есть неподтвержденные документы на персонализацию или отправку в филиал
                                                    comm.CommandText = "select number_doc from StorageDocs where date_doc=@d and priz_gen=0 and (type=5 or type=8)";
                                                    comm.Parameters.Add("@d", SqlDbType.DateTime).Value = DatePickerOne.SelectedDate.Date;
                                                    Database2.ExecuteCommand(comm, ref ds, null);
                                                    if (ds.Tables.Count > 0 && ds.Tables[0].Rows.Count > 0)
                                                    {
                                                        string str = "";
                                                        for (i = 0; i < ds.Tables[0].Rows.Count; i++)
                                                            str += ", " + ds.Tables[0].Rows[i][0].ToString();
                                                        str = str.Substring(2);
                                                        lInfo.Text = String.Format("Создание отчета невозможно. За данное число есть неподтвержденные документы ({0})", str);
                                                        return;
                                                    }
                                                    ds.Clear();
                                                    comm.CommandText = "select id_prb, prod_name, bank_name, id_type from V_ProductsBanks_T order by id_sort";
                                                    res = (string)Database2.ExecuteCommand(comm, ref ds, null);
                                                    ArrayList al = new ArrayList();
                                                    foreach (DataRow dr in ds.Tables[0].Rows)
                                                        al.Add(new MyProd(Convert.ToInt32(dr["id_prb"]), dr["prod_name"].ToString(), dr["bank_name"].ToString(), Convert.ToInt32(dr["id_type"])));
                                                    //выдача на персонализацию
                                                    comm.Parameters.Add("@type", SqlDbType.Int).Value = (int)TypeDoc.PersoCard;
                                                    comm.Parameters.Add("@dateDoc", SqlDbType.DateTime).Value = DatePickerOne.SelectedDate.Date;
                                                    comm.CommandText = "select sum(cnt_brak) as cntbrak, sum(cnt_perso) as cntperso from V_Rep_Moving priz_gen=1 and where type=@type and (id_type=1 or id_type=2) and date_doc=@dateDoc and id_prb=@prb";
                                                    comm.Parameters.Add("@prb", SqlDbType.Int).Value = 0;
                                                    i = 0;
                                                    foreach (MyProd mp in al)
                                                    {
                            //                            if (mp.Type != 1)
                            //                                continue;
                                                        ds.Clear();
                                                        comm.Parameters["@prb"].Value = mp.ID;
                                                        res = (string)Database2.ExecuteCommand(comm, ref ds, null);
                                                        if (ds.Tables[0].Rows.Count == 1)
                                                        {
                                                            int br = (ds.Tables[0].Rows[0]["cntbrak"] == DBNull.Value) ? 0 : Convert.ToInt32(ds.Tables[0].Rows[0]["cntbrak"]);
                                                            int pr = (ds.Tables[0].Rows[0]["cntperso"] == DBNull.Value) ? 0 : Convert.ToInt32(ds.Tables[0].Rows[0]["cntperso"]);
                                                            mp.cnts[0] = br + pr;
                                                        }
                                                    }
                                                    // выдача головному офису
                                                    ds.Clear();
                                                    comm.CommandText = "select sum(cnt_perso) as cntperso from V_Rep_Moving where priz_gen=1 and type=@type and (id_type=1 or id_type=2) and date_doc=@dateDoc and id_prb=@prb and is_head=@ishead";
                                                    comm.Parameters["@type"].Value = (int)TypeDoc.SendToFilial;
                                                    comm.Parameters.Add("@ishead", SqlDbType.Bit).Value = true;
                                                    foreach (MyProd mp in al)
                                                    {
                            //                            if (mp.Type != 1)
                            //                                continue;
                                                        ds.Clear();
                                                        comm.Parameters["@prb"].Value = mp.ID;
                                                        res = (string)Database2.ExecuteCommand(comm, ref ds, null);
                                                        if (ds.Tables[0].Rows.Count == 1)
                                                        {
                                                            int pr = (ds.Tables[0].Rows[0]["cntperso"] == DBNull.Value) ? 0 : Convert.ToInt32(ds.Tables[0].Rows[0]["cntperso"]);
                                                            if (pr > 0)
                                                                mp.cnts[1] = pr;
                                                        }
                                                    }
                                                    // выдача территориальным подразделениям
                                                    comm.CommandText = "select sum(cnt_perso) as cntperso from V_Rep_Moving where priz_gen=1 and type=@type and (id_type=1 or id_type=2) and date_doc=@dateDoc and id_prb=@prb and (is_head=@ishead or is_head is null) and (is_rkc=@ishead or is_rkc is null) and (is_trans=@ishead or is_trans is null)";
                                                    comm.Parameters["@ishead"].Value = false;
                                                    foreach (MyProd mp in al)
                                                    {
                            //                            if (mp.Type != 1)
                            //                                continue;
                                                        ds.Clear();
                                                        comm.Parameters["@prb"].Value = mp.ID;
                                                        res = (string)Database2.ExecuteCommand(comm, ref ds, null);
                                                        if (ds.Tables[0].Rows.Count == 1)
                                                        {
                                                            int pr = (ds.Tables[0].Rows[0]["cntperso"] == DBNull.Value) ? 0 : Convert.ToInt32(ds.Tables[0].Rows[0]["cntperso"]);
                                                            if (pr > 0)
                                                                mp.cnts[3] = pr;
                                                        }
                                                    }
                                                    // выдача ГО ДРБ РКЦ
                                                    comm.CommandText = "select sum(cnt_perso) as cntperso from V_Rep_Moving where priz_gen=1 and type=@type and (id_type=1 or id_type=2) and date_doc=@dateDoc and id_prb=@prb and is_rkc=@ishead";
                                                    comm.Parameters["@ishead"].Value = true;
                                                    foreach (MyProd mp in al)
                                                    {
                            //                            if (mp.Type != 1)
                            //                                continue;
                                                        ds.Clear();
                                                        comm.Parameters["@prb"].Value = mp.ID;
                                                        res = (string)Database2.ExecuteCommand(comm, ref ds, null);
                                                        if (ds.Tables[0].Rows.Count == 1)
                                                        {
                                                            int pr = (ds.Tables[0].Rows[0]["cntperso"] == DBNull.Value) ? 0 : Convert.ToInt32(ds.Tables[0].Rows[0]["cntperso"]);
                                                            if (pr > 0)
                                                                mp.cnts[2] = pr;
                                                        }
                                                    }
                                                    // выдача транспортных карт
                                                    comm.CommandText = "select sum(cnt_perso) as cntperso from V_Rep_Moving where priz_gen=1 and type=@type and (id_type=1 or id_type=2) and date_doc=@dateDoc and id_prb=@prb and is_trans=@ishead";
                                                    comm.Parameters["@ishead"].Value = true;
                                                    foreach (MyProd mp in al)
                                                    {
                            //                            if (mp.Type != 1)
                            //                                continue;
                                                        ds.Clear();
                                                        comm.Parameters["@prb"].Value = mp.ID;
                                                        res = (string)Database2.ExecuteCommand(comm, ref ds, null);
                                                        if (ds.Tables[0].Rows.Count == 1)
                                                        {
                                                            int pr = (ds.Tables[0].Rows[0]["cntperso"] == DBNull.Value) ? 0 : Convert.ToInt32(ds.Tables[0].Rows[0]["cntperso"]);
                                                            if (pr > 0)
                                                                mp.cnts[4] = pr;
                                                        }
                                                    }
                                                    // прием в хранилище персонализированных карт, вернувшихся из филиалов
                                                    comm.CommandText = "select sum(cnt_perso) as cntperso from V_Rep_Moving where priz_gen=1 and type=@type and id_type=1 and date_doc=@dateDoc and id_prb=@prb";
                                                    comm.Parameters["@type"].Value = (int)TypeDoc.ReceiveToBank;
                                                    foreach (MyProd mp in al)
                                                    {
                                                        if (mp.Type != 1)
                                                            continue;
                                                        ds.Clear();
                                                        comm.Parameters["@prb"].Value = mp.ID;
                                                        res = (string)Database2.ExecuteCommand(comm, ref ds, null);
                                                        if (ds.Tables[0].Rows.Count == 1)
                                                        {
                                                            int pr = (ds.Tables[0].Rows[0]["cntperso"] == DBNull.Value) ? 0 : Convert.ToInt32(ds.Tables[0].Rows[0]["cntperso"]);
                                                            if (pr > 0)
                                                                mp.cnts[6] = pr;
                                                        }
                                                    }
                                                    // уничтожено брака (берется все в том числе не востребованные и с истекшим сроком, которые потом вычитаются из этих цифр
                                                    // прием в хранилище персонализированных карт, вернувшихся из филиалов
                                                    comm.CommandText = "select sum(cnt_brak) as cntbrak from V_Rep_Moving where priz_gen=1 and type=@type and id_type=1 and date_doc=@dateDoc and id_prb=@prb";
                                                    comm.Parameters["@type"].Value = (int)TypeDoc.DeleteBrak;
                                                    foreach (MyProd mp in al)
                                                    {
                                                        if (mp.Type != 1)
                                                            continue;
                                                        ds.Clear();
                                                        comm.Parameters["@prb"].Value = mp.ID;
                                                        res = (string)Database2.ExecuteCommand(comm, ref ds, null);
                                                        if (ds.Tables[0].Rows.Count == 1)
                                                        {
                                                            int pr = (ds.Tables[0].Rows[0]["cntbrak"] == DBNull.Value) ? 0 : Convert.ToInt32(ds.Tables[0].Rows[0]["cntbrak"]);
                                                            if (pr > 0)
                                                                mp.cnts[7] = pr;
                                                        }
                                                    }
                                                    // уничтожено карт не востребованных и с истекшим сроком годности
                                                    object obj = 0;
                                                    comm.Parameters.Clear();
                                                    comm.CommandText = "select count(*) as cnt from Cards where id_prb=@prb and dateTerminated=@date and id_prop=@prop";
                                                    comm.Parameters.Add("@prb", SqlDbType.Int);
                                                    comm.Parameters.Add("@prop", SqlDbType.Int);
                                                    comm.Parameters.Add("@date", SqlDbType.DateTime).Value = DatePickerOne.SelectedDate.Date;
                                                    foreach (MyProd mp in al)
                                                    {
                                                        if (mp.Type != 1)
                                                            continue;
                                                        comm.Parameters["@prb"].Value = mp.ID;
                                                        comm.Parameters["@prop"].Value = 7;
                                                        res = Database2.ExecuteScalar(comm, ref obj, null);
                                                        mp.cnts[8] = Convert.ToInt32(obj);
                                                        mp.cnts[7] -= Convert.ToInt32(obj);
                                                        comm.Parameters["@prop"].Value = 8;
                                                        res = Database2.ExecuteScalar(comm, ref obj, null);
                                                        mp.cnts[9] = Convert.ToInt32(obj);
                                                        mp.cnts[7] -= Convert.ToInt32(obj);
                                                    }
                        
                                                    // всего карт с истекшим сроком годности
                                                    //comm.CommandText = "select count(*) as cnt from Cards where id_prb=@prb and dateEnd<@date and id_stat<>@prop";
                                                    //comm.Parameters["@date"].Value = DateTime.Now.Date;
                                                    //comm.Parameters["@prop"].Value = 7;
                                                    //foreach (MyProd mp in al)
                                                    //{
                                                    //  if (mp.Type != 1)
                                                    //      continue;
                                                    //    comm.Parameters["@prb"].Value = mp.ID;
                                                    //    res = Database2.ExecuteScalar(comm, ref obj, null);
                                                    //    mp.cnts[10] = Convert.ToInt32(obj);
                                                    //}
                                                    // закупленные карты и все остальное
                                                    ds.Clear();
                                                    comm.Parameters.Clear();
                                                    comm.CommandText = "select id from PurchDogs where date_stor=@date";
                                                    comm.Parameters.Add("@date", SqlDbType.DateTime).Value = DatePickerOne.SelectedDate.Date;
                                                    Database2.ExecuteCommand(comm, ref ds, null);
                                                    if (ds != null && ds.Tables.Count > 0 && ds.Tables[0].Rows.Count > 0)
                                                    {
                                                        string[] strs = new string[ds.Tables[0].Rows.Count];
                                                        for (i = 0; i < ds.Tables[0].Rows.Count; i++)
                                                            strs[i] = ds.Tables[0].Rows[i]["id"].ToString();
                                                        comm.CommandText = String.Format("select cnt, id_prb from Products_PurchDogs where id_dog in ({0})", String.Join(",", strs));
                                                        ds.Clear();
                                                        Database2.ExecuteCommand(comm, ref ds, null);
                                                        if (ds != null && ds.Tables.Count > 0)
                                                        {
                                                            for (i = 0; i < ds.Tables[0].Rows.Count; i++)
                                                            {
                                                                t = 0;
                                                                for (t = 0; t < al.Count; t++)
                                                                {
                                                                    if (((MyProd)al[t]).ID == Convert.ToInt32(ds.Tables[0].Rows[i]["id_prb"]))
                                                                    {
                                                                        ((MyProd)al[t]).cnts[5] += Convert.ToInt32(ds.Tables[0].Rows[i]["cnt"]);
                                                                        break;
                                                                    }
                                                                }
                                                            }
                                                        }
                                                    }
                                                    comm.Parameters.Clear();
                                                    bool flag = false;
                                                    t = 20;
                                                    int[] cnts = new int[11];
                                                    for (i = 0; i < 11; i++)
                                                        cnts[i] = 0;
                                                    foreach (MyProd mp in al)
                                                    {
                                                        flag = false;
                                                        for (i = 0; i < 11; i++)
                                                            if (mp.cnts[i] > 0)
                                                            {
                                                                flag = true;
                                                                break;
                                                            }
                                                        if (flag)
                                                        {
                                                            ep.SetText(t, 1, String.Format("{0}({1})", mp.Name, mp.Bank));
                                                            for (i = 0; i < 11; i++)
                                                                if (mp.cnts[i] > 0)
                                                                {
                                                                    if (mp.Type == 1)
                                                                        cnts[i] += mp.cnts[i];
                                                                    ep.SetText(t, 3 + i, String.Format("{0}", mp.cnts[i]));
                                                                }
                                                            t++;
                                                        }
                                                    }
                                                    ep.ShowRows(20, t);
                                                    for (i = 0; i < 11; i++)
                                                    {
                                                        if (cnts[i] > 0)
                                                            ep.SetText(100, 3 + i, String.Format("{0}", cnts[i]));
                                                    }
                                                    ep.SetText_Name("date_string", String.Format("от {0:dd MM yyyy}", DatePickerOne.SelectedDate));
                                                    ep.SetText_Name("date_string1", String.Format("{0:dd.MM.yyyy}", DatePickerOne.SelectedDate));
                                                    /////////////////////////////////
                                                    //второй лист - хранилище
                                                    /////////////////////////////////
                                                    ds.Clear();
                                                    ep.SetWorkSheet(2);
                                                    Database2.ExecuteQuery("select name, bank_name, cnt_new, cnt_perso from V_Storage where id_type=1 order by id_sort", ref ds, null);
                                                    int rw = 4;
                                                    int cnt1 = 0, cnt2 = 0;
                                                    for (i = 0; i < ds.Tables[0].Rows.Count; i++)
                                                    {
                                                        string str = String.Format("{0} ({1})", ds.Tables[0].Rows[i]["name"].ToString(), ds.Tables[0].Rows[i]["bank_name"].ToString());
                                                        ep.SetText(rw, 2, str);
                                                        cnt1 += Convert.ToInt32(ds.Tables[0].Rows[i]["cnt_new"]);
                                                        cnt2 += Convert.ToInt32(ds.Tables[0].Rows[i]["cnt_perso"]);

                                                        ep.SetText(rw, 3, ds.Tables[0].Rows[i]["cnt_new"].ToString());
                                                        ep.SetText(rw, 4, ds.Tables[0].Rows[i]["cnt_perso"].ToString());
                                                        rw++;
                                                    }
                                                    ep.SetText(rw, 2, "Итого карт");
                                                    ep.SetText(rw, 3, cnt1.ToString());
                                                    ep.SetText(rw, 4, cnt2.ToString());
                                                    rw++;
                                                    ds.Clear();
                                                    Database2.ExecuteQuery("select name, bank_name, cnt_new, cnt_perso from V_Storage where id_type=2 order by id_sort", ref ds, null);
                                                    ep.SetText(rw, 3, "Бланки (шт)");
                                                    ep.SetText(rw, 4, "Персо");
                                                    rw++;
                                                    for (i = 0; i < ds.Tables[0].Rows.Count; i++)
                                                    {
                                                        string str = String.Format("{0} ({1})", ds.Tables[0].Rows[i]["name"].ToString(), ds.Tables[0].Rows[i]["bank_name"].ToString());
                                                        ep.SetText(rw, 2, str);
                                                        ep.SetText(rw, 3, ds.Tables[0].Rows[i]["cnt_new"].ToString());
                                                        ep.SetText(rw, 4, ds.Tables[0].Rows[i]["cnt_perso"].ToString());
                                                        rw++;
                                                    }
                                                    ds.Clear();
                                                    ep.SetRangeBorders(3, 2, rw - 1, 4);
                                                    ep.SetText_Name("date_string2", String.Format("на {0:dd.MM.yyyy}", DateTime.Now));*/
                        }
                    }
                    #endregion
                    #region карты выданные клиентам подразделением ГО
                    if (id_type == 16)
                    {
                        GivenCardGO(ep);
                    }
                    #endregion
                    #region карты выданные клиентам филиалом
                    if (id_type == 17)
                    {
                        GivenCardBranch(ep, ddlBranch.SelectedValue);
                    }
                    #endregion
                    #region Отчет по карте жителя
                    if (id_type == 44)
                    {
                      GenReportKgrt(ep);
                    }
                    #endregion
                    if (make) // чтобы не сохранять движение ценностей архивный
                    {
                        if (doc.Length > 0 && WebConfigurationManager.AppSettings["DocPath"] != null)
                        {
                            doc = String.Format("{0}Temp\\{1}", WebConfigurationManager.AppSettings["DocPath"], doc);
                            ep.SaveAsDoc(doc, false);
                            if (make && id_type == 13 && fname.Length > 0)
                            {
                                File.Copy(doc, fname, true);
                            }
                        }
                    }

                }
                ep.Close();
                GC.Collect();
                System.Threading.Thread.CurrentThread.CurrentCulture = oldCI;
                if (make && doc.Length > 0)
                {
                    Database2.Log(sc.UserGuid(User.Identity.Name), String.Format("Сформирован отчет '{0}'", dListDoc.SelectedItem.Text), null);
                    if (id_type != 44)
                    {
                        SendConfirm(Restrictions.ConfirmSendingValues, doc);
                    }
                    else
                    {
                        SendConfirm(Restrictions.ConfirmKGRT, doc);
                    }
                    return;
                }
                if (!make && fname.Length > 0)
                {
                    Database2.Log(sc.UserGuid(User.Identity.Name), String.Format("Показан архивный отчет '{0}' {1:dd.MM.yyyy}", dListDoc.SelectedItem.Text, DatePickerOne.SelectedDate), null);
                    if (id_type != 44)
                    {
                        SendConfirm(Restrictions.ConfirmSendingValues, fname);
                    }
                    else
                    {
                        SendConfirm(Restrictions.ConfirmKGRT, fname);
                    }
                    ep.ReturnXls(Response, fname);
                }
            }
        }
        protected void bReportQuery_Click(object sender, ImageClickEventArgs e)
        {
            int id_type = Convert.ToInt32(dListDoc.SelectedItem.Value);
            string rparams = "";
            switch (id_type)
            {
                case (1):
                    break;
                case (2):
                    break;
                case 3: //выпущенные карты по продуктам
                    rparams = String.Format("{0:dd.MM.yyyy};{1:dd.MM.yyyy};{2}", DatePickerStart.SelectedDate, DatePickerEnd.SelectedDate, sc.BranchId(User.Identity.Name));
                    break;
                case 4:
                    string productId = ddlProducts.SelectedValue;
                    string productName = ddlProducts.SelectedItem.Text;
                    string type = rbPeriods.SelectedValue;
                    string days = tbDays.Text;

                    rparams = String.Format("{0:dd.MM.yyyy};{1:dd.MM.yyyy};{2};{3};{4};{5};{6}", DatePickerStart.SelectedDate, DatePickerEnd.SelectedDate, productId, productName, type, days, sc.BranchId(User.Identity.Name));
                    break;
                case 6:
                    break;

                case (7): //ценности ниже порогового значения
                    break;

                case 8:
                    rparams = String.Format("{0:dd.MM.yyyy};{1:dd.MM.yyyy};{2}", DatePickerStart.SelectedDate, DatePickerEnd.SelectedDate, sc.BranchId(User.Identity.Name));
                    break;

                case 9:
                    rparams = String.Format("{0:dd.MM.yyyy};{1:dd.MM.yyyy};{2}", DatePickerStart.SelectedDate, DatePickerEnd.SelectedDate, sc.BranchId(User.Identity.Name));
                    break;

                case 10:
                    rparams = String.Format("{0:dd.MM.yyyy};{1:dd.MM.yyyy};{2}", DatePickerStart.SelectedDate, DatePickerEnd.SelectedDate, sc.BranchId(User.Identity.Name));
                    break;

                case 11:
                    rparams = String.Format("{0:dd.MM.yyyy};{1:dd.MM.yyyy};{2}", DatePickerStart.SelectedDate, DatePickerEnd.SelectedDate, sc.BranchId(User.Identity.Name));
                    break;
                case (12): //сводный акт
                    rparams = String.Format("{0:dd.MM.yyyy}", DatePickerOne.SelectedDate);

                    break;
                case 14:
                    rparams = String.Format("{0:dd.MM.yyyy};{1:dd.MM.yyyy};{2}", DatePickerStart.SelectedDate, DatePickerEnd.SelectedDate, sc.BranchId(User.Identity.Name));
                    break;

                case (18): //состояние хранилища (новый вариант)


                case (13): //движение ценностей (старый вариант)
                    rparams = String.Format("{0:dd.MM.yyyy}", DatePickerOne.SelectedDate);
                    break;
                case (15): //движение ценностей (новый вариант)
                    rparams = String.Format("{0:dd.MM.yyyy}", DatePickerOne.SelectedDate);

                    break;
                case (16): //движение ценностей (новый вариант)
                    rparams = String.Format("{0:dd.MM.yyyy}", DatePickerOne.SelectedDate);

                    break;
                case (17):
                    var sel = ddlBranch.SelectedValue;
                    rparams = String.Format("{0:dd.MM.yyyy};{1}", DatePickerOne.SelectedDate, ddlBranch.SelectedValue);
                    break;
                case (19):
                    rparams = String.Format("{0:dd.MM.yyyy};{1}", DatePickerOne.SelectedDate, sc.BranchId(User.Identity.Name));
                    break;
                case (20):

                    //splittedParams = reportParameters.Split(';');

                    //DateTime SelectedDate = DateTime.Parse(splittedParams[0]);
                    //int current_branch = Convert.ToInt32(splittedParams[1]);

                    rparams = String.Format("{0:dd.MM.yyyy};{1};{2}", DatePickerOne.SelectedDate, sc.BranchId(User.Identity.Name), branch_main_filial);
                    break;
                case (43):

                    rparams = String.Format("{0:dd.MM.yyyy};{1:dd.MM.yyyy}", DatePickerStart.SelectedDate, DatePickerEnd.SelectedDate);
                    break;
                case (42):
                    rparams = String.Format("{0:dd.MM.yyyy};{1};{2};{3};{4}", DatePickerOne.SelectedDate, current_branch, this.firstrangeradio.Checked, this.nextrangeradio.Checked, this.allrangeradio.Checked);
                    break;
                case (44):

                    rparams = String.Format("{0:dd.MM.yyyy};{1:dd.MM.yyyy}", DatePickerStart.SelectedDate, DatePickerEnd.SelectedDate);
                    break;
                case (45):

                    rparams = String.Format("{0:dd.MM.yyyy};{1};{2}", DatePickerOne.SelectedDate, sc.BranchId(User.Identity.Name), User.Identity.Name);
                    break;
            }
            using (SqlConnection conn = new SqlConnection(FuncClass.ConnectionString))
            {
                conn.Open();

                int reportLimit = Convert.ToInt32(ConfigurationManager.AppSettings["ReportLimit"]);
                int count;
                // Проверка количества заявок у этого пользователя
                using (SqlCommand comm = Database2.Conn.CreateCommand())
                {
                    object obj = null;

                    comm.CommandText = @"select count (*) from ReportQuery where UserId = @userid";

                    string userId = sc.UserGuid(User.Identity.Name);
                    comm.Parameters.AddWithValue("@userid", userId);
                    var result = Database2.ExecuteScalar(comm, ref obj, null);

                    count = (int)obj;


                }

                if (count >= reportLimit)
                {
                    using (SqlCommand comm = Database2.Conn.CreateCommand())
                    {
                        object obj = null;


                        string s = sc.UserGuid(User.Identity.Name);

                        comm.CommandText =
                            "DELETE FROM ReportQuery WHERE id not in(SELECT TOP(@limit) id FROM ReportQuery WHERE UserId = @userid ORDER BY ReportDate DESC) and UserId = @userid";
                        comm.Parameters.AddWithValue("@userid", sc.UserGuid(User.Identity.Name));
                        comm.Parameters.AddWithValue("@limit", reportLimit);

                        var result = Database2.ExecuteScalar(comm, ref obj, null);
                    }
                }




                using (SqlCommand comm = conn.CreateCommand())
                {
                    object obj = null;


                    comm.CommandText = @"insert into ReportQuery (UserID, ReportType, ReportParameters, ReportStatus, ReportDate) values
                                            (@userid, @rtype, @rparams, @rstatus, @reportDate)";

                    //comm.Parameters.Add("@userid", SqlDbType.UniqueIdentifier).Value = Session["CurrentUserId"];
                    comm.Parameters.Add("@userid", SqlDbType.UniqueIdentifier).Value = Guid.Parse(sc.UserGuid(User.Identity.Name));
                    comm.Parameters.Add("@rtype", SqlDbType.Int).Value = id_type;
                    comm.Parameters.Add("@rparams", SqlDbType.NText).Value = rparams;
                    comm.Parameters.AddWithValue("@rstatus", 1);
                    comm.Parameters.Add("@reportDate", SqlDbType.DateTime).Value = DateTime.Now;

                    var result = Database2.ExecuteScalar(comm, ref obj, null);

                }
                conn.Close();
            }

        }
        private void SendConfirm(Restrictions r, string docTempName)
        {
            if (r == Restrictions.ConfirmSendingValues || r == Restrictions.ConfirmKGRT)
            {
                if (docTempName.Length > 0)
                {
                    MailMessage mm = new MailMessage();
                    DataSet ds1 = new DataSet();
                    string res = Database2.ExecuteQuery(String.Format("select UserName from V_UserAction where ActionID={0}", (int)r) , ref ds1, null); //(int)Restrictions.ConfirmSendingValues), ref ds1, null);
                    if (ds1.Tables.Count == 0 || ds1.Tables[0].Rows.Count<1)
                    {
                        lbInform.Text = "Не найдены пользователи с ролью " + r.ToString();
                        return;
                    }
                    for (int t = 0; t < ds1.Tables[0].Rows.Count; t++)
                    {
                        object obj = null;
                        Database2.ExecuteScalar(String.Format("select EMail from vw_aspnet_MembershipUsers where UserName='{0}'", ds1.Tables[0].Rows[t]["UserName"].ToString()), ref obj, null);
                        if (obj != null)
                        {
                            try
                            {
                                MailAddress mTo = new MailAddress((string)obj);
                                mm.Bcc.Add(mTo);
                            }
                            catch { }
                        }
                    }
                    try
                    {
                        int id_type = Convert.ToInt32(dListDoc.SelectedItem.Value);
                        SmtpClient sc = new SmtpClient(ConfigurationSettings.AppSettings["SmtpServer"]);
                        sc.Credentials = new NetworkCredential(ConfigurationSettings.AppSettings["EMailFrom"], ConfigurationSettings.AppSettings["Pwd"]);
                        MailAddress mailFrom = new MailAddress(ConfigurationSettings.AppSettings["EMailFrom"], "CardPerso");
                        mm.From = mailFrom;
                        if (id_type == 13 || id_type == 15)
                        {
                            mm.Subject = "CardPerso - движение ценностей";
                            mm.Body = String.Format("В системе CardPerso сформирован отчет Движение ценностей.\n{0:HH.mm dd.MM.yyyy}", DateTime.Now);
                        }
                        if (id_type == 16)
                        {
                            mm.Subject = "CardPerso - выданные карты";
                            mm.Body = String.Format("В системе CardPerso сформирован отчет Выданные карты подразделением.\n{0:HH.mm dd.MM.yyyy}", DateTime.Now);
                        }
                        if (id_type == 17)
                        {
                            mm.Subject = "CardPerso - выданные карты";
                            mm.Body = String.Format("В системе CardPerso сформирован отчет Выданные карты филиалом {1}.\n{0:HH.mm dd.MM.yyyy}", DateTime.Now, ddlBranch.SelectedValue.ToString());
                        }
                        if (id_type == 44)
                        {
                            mm.Subject = "CardPerso - карты жителя";
                            mm.Body = String.Format("В системе CardPerso сформирован отчет Карты жителя.\n{0:HH.mm dd.MM.yyyy}", DateTime.Now);
                        }
                        mm.Attachments.Add(new Attachment(docTempName));
                        if (mm.Bcc.Count > 0)
                        {
                            sc.Send(mm);
                            lbInform.Text = "Служебная записка отправлена по рассылке";
                        }
                    }
                    catch (Exception  Ex)
                    {
                        lbInform.Text = Ex.Message;
                    }
                }
            }
        }
        private void MovingStock(DateTime sDate)
        {
            string n_fileDBF = "eq_1_vneb10.dbf";
            zipArch = "eq_1_vneb10.zip";

            dirDBF.Text = "";

            String fname = String.Format("{0}RPn_{1:ddMMyyyy}.zip", ConfigurationSettings.AppSettings["ArchivePath"], sDate);
            if (sDate != DateTime.Now.Date)
            {
                if (!File.Exists(fname))
                {
                    fname = "";
                    lInfo.Text = "В архиве нет экспортного файла за данное число. Состояние хранилища будет на текущую дату.";
                }
                else
                {
                    dirDBF.Text = fname;
                    ClientScript.RegisterClientScriptBlock(GetType(), "okMovSt", "<script type='text/javascript'>$(document).ready(function(){ $('#" + bOkDirDBF.ClientID + "').click(); });</script>");
                    return;
                }
            }
            

            DataSet ds = new DataSet();
            SqlCommand comm = new SqlCommand();
            String res;
            ArrayList al = new ArrayList();
            bool isError = false;
            DbfFile odbf = new DbfFile(System.Text.Encoding.GetEncoding(866));
            DbfRecord orec = null;
            BranchStore branch = new BranchStore(current_branch, "", "");
            MovingStockSection mvConfiguration = null;

            try
            {
                lock (Database2.lockObjectDB)
                {
                    int i = 0;
                    comm.CommandText = "select number_doc from StorageDocs where date_doc=@d and priz_gen=0 and (type=5 or type=8)";
                    comm.Parameters.Add("@d", SqlDbType.DateTime).Value = sDate;
                    Database2.ExecuteCommand(comm, ref ds, null);
                    if (ds.Tables.Count > 0 && ds.Tables[0].Rows.Count > 0)
                    {
                        string str = "";
                        for (i = 0; i < ds.Tables[0].Rows.Count; i++)
                            str += ", " + ds.Tables[0].Rows[i][0].ToString();
                        str = str.Substring(2);
                        throw new Exception("Создание экспортного файла невозможно. За данное число есть неподтвержденные документы (" + str + ")");
                    }
                    ds.Clear();
                    comm.CommandText = "select id_prb, prod_name, bank_name, id_type from V_ProductsBanks_T order by id_sort";
                    res = (string)Database2.ExecuteCommand(comm, ref ds, null);
                    foreach (DataRow dr in ds.Tables[0].Rows)
                    {
                        al.Add(new MyProd(Convert.ToInt32(dr["id_prb"]), dr["prod_name"].ToString(), dr["bank_name"].ToString(), Convert.ToInt32(dr["id_type"])));
                    }
                    if (al.Count < 1)
                    {
                        throw new Exception("Не найдены продукты");
                    }
                    try
                    {
                        mvConfiguration = (MovingStockSection)ConfigurationManager.GetSection("movingstock");
                    }
                    catch (Exception ee)
                    {
                        throw new Exception("Ошибка при загрузке секции movingstock: " + ee.Message);
                    }
                    if (mvConfiguration == null || mvConfiguration.Fields.Count < 1)
                    {
                        throw new Exception("Не заданы продукты в секции movingstock");
                    }

                    //выдача на персонализацию
                    comm.Parameters.Add("@type", SqlDbType.Int).Value = (int)TypeDoc.PersoCard;
                    comm.Parameters.Add("@dateDoc", SqlDbType.DateTime).Value = sDate;
                    comm.CommandText = "select id_prb,sum(cnt_brak) as cntbrak, sum(cnt_perso) as cntperso from V_Rep_Moving where priz_gen=1 and type=@type and (id_type=1 or id_type=2) and date_doc=@dateDoc group by id_prb";
                    comm.Parameters.Add("@prb", SqlDbType.Int).Value = 0;
                    i = 0;
                    ds.Clear();
                    Database2.ExecuteCommand(comm, ref ds, null);
                    foreach (MyProd mp in al)
                    {
                        for (i = 0; ds.Tables.Count > 0 && i < ds.Tables[0].Rows.Count; i++)
                        {
                            if (mp.ID == Convert.ToInt32(ds.Tables[0].Rows[i]["id_prb"]))
                            {
                                int br = (ds.Tables[0].Rows[i]["cntbrak"] == DBNull.Value) ? 0 : Convert.ToInt32(ds.Tables[0].Rows[i]["cntbrak"]);
                                int pr = (ds.Tables[0].Rows[i]["cntperso"] == DBNull.Value) ? 0 : Convert.ToInt32(ds.Tables[0].Rows[i]["cntperso"]);
                                branch.addCount(mp.Type, mp.Name,0, br + pr);
                                break;
                            }
                        }
                    }
                    // выдача головному офису
                    ds.Clear();
                    comm.CommandText = "select id_prb,sum(cnt_perso) as cntperso from V_Rep_Moving where priz_gen=1 and type=@type and (id_type=1 or id_type=2) and date_doc=@dateDoc and is_head=@ishead group by id_prb";
                    comm.Parameters["@type"].Value = (int)TypeDoc.SendToFilial;
                    comm.Parameters.Add("@ishead", SqlDbType.Bit).Value = true;
                    Database2.ExecuteCommand(comm, ref ds, null);
                    foreach (MyProd mp in al)
                    {
                        for (i = 0; ds.Tables.Count > 0 && i < ds.Tables[0].Rows.Count; i++)
                        {
                            if (mp.ID == Convert.ToInt32(ds.Tables[0].Rows[i]["id_prb"]))
                            {
                                int pr = (ds.Tables[0].Rows[i]["cntperso"] == DBNull.Value) ? 0 : Convert.ToInt32(ds.Tables[0].Rows[i]["cntperso"]);
                                if (pr > 0)
                                    branch.addCount(mp.Type, mp.Name, 1, pr);
                                break;
                            }
                        }
                    }
                    // выдача территориальным подразделениям
                    comm.CommandText = "select id_prb,sum(cnt_perso) as cntperso from V_Rep_Moving where priz_gen=1 and type=@type and (id_type=1 or id_type=2) and date_doc=@dateDoc and (is_head=@ishead or is_head is null) and (is_rkc=@ishead or is_rkc is null) and (is_trans=@ishead or is_trans is null) group by id_prb";
                    comm.Parameters["@ishead"].Value = false;
                    ds.Clear();
                    Database2.ExecuteCommand(comm, ref ds, null);
                    foreach (MyProd mp in al)
                    {
                        for (i = 0; ds.Tables.Count > 0 && i < ds.Tables[0].Rows.Count; i++)
                        {
                            if (mp.ID == Convert.ToInt32(ds.Tables[0].Rows[i]["id_prb"]))
                            {
                                int pr = (ds.Tables[0].Rows[i]["cntperso"] == DBNull.Value) ? 0 : Convert.ToInt32(ds.Tables[0].Rows[i]["cntperso"]);
                                if (pr > 0)
                                    branch.addCount(mp.Type, mp.Name, 3, pr);
                                break;
                            }
                        }
                    }
                    // выдача ГО ДРБ РКЦ
                    comm.CommandText = "select id_prb,sum(cnt_perso) as cntperso from V_Rep_Moving where priz_gen=1 and type=@type and (id_type=1 or id_type=2) and date_doc=@dateDoc and is_rkc=@ishead group by id_prb";
                    comm.Parameters["@ishead"].Value = true;
                    ds.Clear();
                    Database2.ExecuteCommand(comm, ref ds, null);
                    foreach (MyProd mp in al)
                    {
                        for (i = 0; ds.Tables.Count > 0 && i < ds.Tables[0].Rows.Count; i++)
                        {
                            if (mp.ID == Convert.ToInt32(ds.Tables[0].Rows[i]["id_prb"]))
                            {
                                int pr = (ds.Tables[0].Rows[i]["cntperso"] == DBNull.Value) ? 0 : Convert.ToInt32(ds.Tables[0].Rows[i]["cntperso"]);
                                if (pr > 0)
                                    branch.addCount(mp.Type, mp.Name, 2, pr);
                                break;
                            }
                        }
                    }
                    // выдача транспортных карт
                    comm.CommandText = "select id_prb,sum(cnt_perso) as cntperso from V_Rep_Moving where priz_gen=1 and type=@type and (id_type=1 or id_type=2) and date_doc=@dateDoc and is_trans=@ishead group by id_prb";
                    comm.Parameters["@ishead"].Value = true;
                    ds.Clear();
                    Database2.ExecuteCommand(comm, ref ds, null);
                    foreach (MyProd mp in al)
                    {
                        for (i = 0; ds.Tables.Count > 0 && i < ds.Tables[0].Rows.Count; i++)
                        {
                            if (mp.ID == Convert.ToInt32(ds.Tables[0].Rows[i]["id_prb"]))
                            {
                                int pr = (ds.Tables[0].Rows[i]["cntperso"] == DBNull.Value) ? 0 : Convert.ToInt32(ds.Tables[0].Rows[i]["cntperso"]);
                                if (pr > 0)
                                    branch.addCount(mp.Type, mp.Name, 4, pr);
                                break;
                            }
                        }
                    }
                    // прием в хранилище персонализированных карт, вернувшихся из филиалов и бракованных карт
                    /*
                    if (branch_main_filial > 0)
                    {
                        comm.CommandText = String.Format("select id_prb,sum(cnt_perso) as cntperso, sum(cnt_brak) as cntbrak from V_Rep_Moving where priz_gen=1 and type=@type and (id_type=1 or id_type=2) and date_doc=@dateDoc and (id_branch={0} or id_branch in (select id from Branchs where id_parent={0})) group by id_prb", branch_main_filial);
                    }
                    else
                    {
                        comm.CommandText = String.Format("select id_prb,sum(cnt_perso) as cntperso, sum(cnt_brak) as cntbrak from V_Rep_Moving where priz_gen=1 and type=@type and (id_type=1 or id_type=2) and date_doc=@dateDoc and (id_branch!={0} and id_branch not in (select id from Branchs where id_parent={0})) group by id_prb", 106);
                    }
                    */
                    comm.CommandText = "select id_prb,sum(cnt_perso) as cntperso, sum(cnt_brak) as cntbrak from V_Rep_Moving where priz_gen=1 and type=@type and (id_type=1 or id_type=2) and date_doc=@dateDoc group by id_prb";

                    comm.Parameters["@type"].Value = (int)TypeDoc.ReceiveToBank;
                    ds.Clear();
                    Database2.ExecuteCommand(comm, ref ds, null);
                    foreach (MyProd mp in al)
                    {
                        if (mp.Type != 1 && mp.Type != 2)
                            continue;
                        for (i = 0; ds.Tables.Count > 0 && i < ds.Tables[0].Rows.Count; i++)
                        {
                            if (mp.ID == Convert.ToInt32(ds.Tables[0].Rows[i]["id_prb"]))
                            {
                                int pr = (ds.Tables[0].Rows[i]["cntperso"] == DBNull.Value) ? 0 : Convert.ToInt32(ds.Tables[0].Rows[i]["cntperso"]);
                                if (pr > 0)
                                    branch.addCount(mp.Type, mp.Name, 8, pr);
                                pr = (ds.Tables[0].Rows[i]["cntbrak"] == DBNull.Value) ? 0 : Convert.ToInt32(ds.Tables[0].Rows[i]["cntbrak"]);
                                if (pr > 0)
                                    branch.addCount(mp.Type, mp.Name, 9, pr);
                                break;
                            }
                        }

                    }
                    // прием в хранилище персонализированных карт на экспертизу                    
                    comm.CommandText = "select id_prb,sum(cnt_perso) as cntperso, sum(cnt_brak) as cntbrak from V_Rep_Moving where priz_gen=1 and type=@type and (id_type=1 or id_type=2) and date_doc=@dateDoc group by id_prb";
                    comm.Parameters["@type"].Value = (int)TypeDoc.ReceiveToExpertiza;
                    ds.Clear();
                    Database2.ExecuteCommand(comm, ref ds, null);
                    foreach (MyProd mp in al)
                    {
                        if (mp.Type != 1)
                            continue;
                        for (i = 0; ds.Tables.Count > 0 && i < ds.Tables[0].Rows.Count; i++)
                        {
                            if (mp.ID == Convert.ToInt32(ds.Tables[0].Rows[i]["id_prb"]))
                            {
                                int pr = (ds.Tables[0].Rows[i]["cntperso"] == DBNull.Value) ? 0 : Convert.ToInt32(ds.Tables[0].Rows[i]["cntperso"]);
                                if (pr > 0)
                                    branch.addCount(mp.Type, mp.Name, 8, pr);
                                break;
                            }
                        }
                    }
                    // выдача из хранилища карт на рекламу и тесты
                    comm.CommandText = "select id_prb,sum(cnt_new) as cntnew from V_Rep_Moving where priz_gen=1 and type=@type and date_doc=@dateDoc group by id_prb";
                    comm.Parameters["@type"].Value = (int)TypeDoc.Reklama;
                    ds.Clear();
                    Database2.ExecuteCommand(comm, ref ds, null);
                    foreach (MyProd mp in al)
                    {
                        for (i = 0; ds.Tables.Count > 0 && i < ds.Tables[0].Rows.Count; i++)
                        {
                            if (mp.ID == Convert.ToInt32(ds.Tables[0].Rows[i]["id_prb"]))
                            {
                                int pr = (ds.Tables[0].Rows[i]["cntnew"] == DBNull.Value) ? 0 : Convert.ToInt32(ds.Tables[0].Rows[i]["cntnew"]);
                                if (pr > 0)
                                    branch.addCount(mp.Type, mp.Name, 5, pr);
                                break;
                            }
                        }
                    }
                    // уничтожено брака (берется все в том числе не востребованные и с истекшим сроком, которые потом вычитаются из этих цифр
                    // прием в хранилище персонализированных карт, вернувшихся из филиалов
                    //comm.CommandText = "select sum(cnt_brak) as cntbrak from V_Rep_Moving where priz_gen=1 and type=@type and (id_type=1 or id_type=2) and date_doc=@dateDoc and id_prb=@prb";
                    ////comm.CommandText = "select id_prb,sum(cnt_brak) as cntbrak from V_Rep_Moving where priz_gen=1 and type=@type and (id_type=1 or id_type=2) and date_doc=@dateDoc group by id_prb";
                    /*
                    if (branch_main_filial > 0)
                    {
                        comm.CommandText = String.Format("select id_prb,sum(cnt_brak) as cntbrak from V_Rep_Moving where priz_gen=1 and type=@type and (id_type=1 or id_type=2) and date_doc=@dateDoc and (id_branch={0} or id_branch in (select id from Branchs where id_parent={0})) group by id_prb", branch_main_filial);
                    }
                    else
                    {
                        comm.CommandText = String.Format("select id_prb,sum(cnt_brak) as cntbrak from V_Rep_Moving where priz_gen=1 and type=@type and (id_type=1 or id_type=2) and date_doc=@dateDoc and (id_branch is null or (id_branch!={0} and id_branch not in (select id from Branchs where id_parent={0}))) group by id_prb", 106);
                    }
                    */
                    comm.CommandText = "select id_prb,sum(cnt_brak) as cntbrak from V_Rep_Moving where priz_gen=1 and type=@type and (id_type=1 or id_type=2) and date_doc=@dateDoc group by id_prb";
                    comm.Parameters["@type"].Value = (int)TypeDoc.DeleteBrak;
                    ds.Clear();
                    Database2.ExecuteCommand(comm, ref ds, null);
                    foreach (MyProd mp in al)
                    {
                        if (mp.Type != 1 && mp.Type != 2)
                            continue;
                        for (i = 0; ds.Tables.Count > 0 && i < ds.Tables[0].Rows.Count; i++)
                        {
                            if (mp.ID == Convert.ToInt32(ds.Tables[0].Rows[i]["id_prb"]))
                            {
                                int pr = (ds.Tables[0].Rows[i]["cntbrak"] == DBNull.Value) ? 0 : Convert.ToInt32(ds.Tables[0].Rows[i]["cntbrak"]);
                                if (pr > 0)
                                    branch.addCount(mp.Type, mp.Name, 11, pr);
                                break;
                            }
                        }
                    }
                    // уничтожено карт не востребованных и с истекшим сроком годности
                    object obj = 0;
                    comm.Parameters.Clear();
                    /*
                    if (branch_main_filial > 0)
                    {
                        comm.CommandText = String.Format("select count(*) as cnt from Cards where id_prb=@prb and dateTerminated=@date and id_prop=@prop and id_BranchCurrent={0}", branch_main_filial);
                    }
                    else
                    {
                        comm.CommandText = String.Format("select count(*) as cnt from Cards where id_prb=@prb and dateTerminated=@date and id_prop=@prop and id_BranchCurrent!={0}", 106);
                    }
                    */

                    // Изменено 25.11.2019, потому что стало сказываться уничтожение карт в филиале
                    /*
                    comm.CommandText = "select count(*) as cnt from Cards where id_prb=@prb and dateTerminated=@date and id_prop=@prop";
                    comm.Parameters.Add("@prb", SqlDbType.Int);
                    comm.Parameters.Add("@prop", SqlDbType.Int);
                    comm.Parameters.Add("@date", SqlDbType.DateTime).Value = sDate;
                    */

                    comm.CommandText = $"select count(*) from cards c inner join Cards_StorageDocs cs on c.id=cs.id_card inner join StorageDocs s on cs.id_doc=s.id " +
                        "where s.date_doc = @date and(s.type = 11) and s.priz_gen = 1 and id_prop=@prop and id_prb=@prb";

                    comm.Parameters.Add("@prb", SqlDbType.Int);
                    comm.Parameters.Add("@prop", SqlDbType.Int);
                    comm.Parameters.Add("@date", SqlDbType.DateTime).Value = sDate;

                    foreach (MyProd mp in al)
                    {
                        if (mp.Type != 1)
                            continue;
                        comm.Parameters["@prb"].Value = mp.ID;
                        comm.Parameters["@prop"].Value = 7;
                        res = Database2.ExecuteScalar(comm, ref obj, null);
                        branch.addCount(mp.Type, mp.Name, 12, Convert.ToInt32(obj));
                        branch.addCount(mp.Type, mp.Name, 11, -Convert.ToInt32(obj));
                        comm.Parameters["@prop"].Value = 8;
                        res = Database2.ExecuteScalar(comm, ref obj, null);
                        branch.addCount(mp.Type, mp.Name, 13, Convert.ToInt32(obj));
                        branch.addCount(mp.Type, mp.Name, 11, -Convert.ToInt32(obj));
                    }
                    comm.Parameters.Clear();
                    // и отдельно считаем пины
                    comm.CommandText = $"select count(*) from cards c inner join Cards_StorageDocs cs on c.id=cs.id_card inner join StorageDocs s on cs.id_doc=s.id " +
                        "where s.date_doc = @date and (s.type = 11) and s.priz_gen = 1 and c.id_prop=@prop and c.isPin=1";

                    comm.Parameters.Add("@prop", SqlDbType.Int);
                    comm.Parameters.Add("@date", SqlDbType.DateTime).Value = DatePickerOne.SelectedDate.Date;

                    foreach (MyProd mp in al)
                    {
                        if (mp.Type != 2)
                            continue;
                        comm.Parameters["@prop"].Value = 7;
                        res = Database2.ExecuteScalar(comm, ref obj, null);
                        mp.cnts[12] = Convert.ToInt32(obj);
                        mp.cnts[11] -= Convert.ToInt32(obj);
                        comm.Parameters["@prop"].Value = 8;
                        res = Database2.ExecuteScalar(comm, ref obj, null);
                        mp.cnts[13] = Convert.ToInt32(obj);
                        mp.cnts[11] -= Convert.ToInt32(obj);
                    }
                    // карты ожидающие уничтожения
                    /*
                    if (branch_main_filial > 0)
                    {
                        comm.CommandText = String.Format("select id_prb,count(*) as ccard,case when sum(cast(ispin as integer)) is null then 0 else sum(cast(ispin as integer)) end as cpin from Cards where id_stat=6 and id_branchCurrent={0} group by id_prb", branch_main_filial);
                    }
                    else
                    {
                        comm.CommandText = String.Format("select id_prb,count(*) as ccard,case when sum(cast(ispin as integer)) is null then 0 else sum(cast(ispin as integer)) end as cpin from Cards where id_stat=6 and id_branchCurrent not in ({0}) group by id_prb", 106);
                    }
                    */
                    comm.CommandText = "select id_prb,count(*) as ccard,case when sum(cast(ispin as integer)) is null then 0 else sum(cast(ispin as integer)) end as cpin from Cards where id_stat=6 and (id_BranchCurrent = 0 or id_BranchCurrent is NULL) group by id_prb";
                    ds.Clear();
                    Database2.ExecuteCommand(comm, ref ds, null);
                    int pinc = 0;
                    foreach (MyProd mp in al)
                    {
                        if (mp.Type != 1) continue;
                        for (i = 0; ds.Tables.Count > 0 && i < ds.Tables[0].Rows.Count; i++)
                        {
                            if (mp.ID == Convert.ToInt32(ds.Tables[0].Rows[i]["id_prb"]))
                            {
                                int pr = (ds.Tables[0].Rows[i]["ccard"] == DBNull.Value) ? 0 : Convert.ToInt32(ds.Tables[0].Rows[i]["ccard"]);
                                if (pr > 0)
                                    branch.addCount(mp.Type, mp.Name, 14, pr); 
                                pr = (ds.Tables[0].Rows[i]["cpin"] == DBNull.Value) ? 0 : Convert.ToInt32(ds.Tables[0].Rows[i]["cpin"]);
                                pinc += pr;
                                break;
                            }
                        }
                    }
                    foreach (MyProd mp in al)
                    {
                        if (mp.Type == 2)
                        {
                            branch.addCount(mp.Type, mp.Name, 14, pinc); break;
                        }
                    }
                    // закупленные карты и все остальное
                    ds.Clear();
                    comm.Parameters.Clear();
                    comm.CommandText = "select id from PurchDogs where date_stor=@date";
                    comm.Parameters.Add("@date", SqlDbType.DateTime).Value = sDate;
                    Database2.ExecuteCommand(comm, ref ds, null);
                    if (ds != null && ds.Tables.Count > 0 && ds.Tables[0].Rows.Count > 0)
                    {
                        string[] strs = new string[ds.Tables[0].Rows.Count];
                        for (i = 0; i < ds.Tables[0].Rows.Count; i++)
                            strs[i] = ds.Tables[0].Rows[i]["id"].ToString();
                        comm.CommandText = String.Format("select cnt, id_prb from Products_PurchDogs where id_dog in ({0})", String.Join(",", strs));
                        ds.Clear();
                        Database2.ExecuteCommand(comm, ref ds, null);
                        if (ds != null && ds.Tables.Count > 0)
                        {
                            for (i = 0; i < ds.Tables[0].Rows.Count; i++)
                            {
                                int t = 0;
                                for (t = 0; t < al.Count; t++)
                                {
                                    if (((MyProd)al[t]).ID == Convert.ToInt32(ds.Tables[0].Rows[i]["id_prb"]))
                                    {
                                        branch.addCount(((MyProd)al[t]).Type,((MyProd)al[t]).Name, 7, Convert.ToInt32(ds.Tables[0].Rows[i]["cnt"]));
                                        break;
                                    }
                                }
                            }
                        }
                    }
                    // на упаковку
                    ds.Clear();
                    comm.Parameters.Clear();
                    comm.CommandText = "select id_prb,sum(cnt_perso) as cntperso from V_Rep_Moving where priz_gen=1 and type=@type and date_doc=@dateDoc group by id_prb";
                    comm.Parameters.Add("@type", SqlDbType.Int).Value = (int)TypeDoc.ToWrapping;
                    comm.Parameters.Add("@dateDoc", SqlDbType.DateTime).Value = sDate;
                    ds.Clear();
                    Database2.ExecuteCommand(comm, ref ds, null);
                    foreach (MyProd mp in al)
                    {
                        for (i = 0; ds.Tables.Count > 0 && i < ds.Tables[0].Rows.Count; i++)
                        {
                            if (mp.ID == Convert.ToInt32(ds.Tables[0].Rows[i]["id_prb"]))
                            {
                                int pr = (ds.Tables[0].Rows[i]["cntperso"] == DBNull.Value) ? 0 : Convert.ToInt32(ds.Tables[0].Rows[i]["cntperso"]);
                                branch.addCount(mp.Type, mp.Name, 6, pr);
                                break;
                            }
                        }
                    }
                    // c упаковки
                    ds.Clear();
                    comm.Parameters.Clear();
                    comm.CommandText = "select id_prb,sum(cnt_perso) as cntperso from V_Rep_Moving where priz_gen=1 and type=@type and date_doc=@dateDoc group by id_prb";
                    comm.Parameters.Add("@type", SqlDbType.Int).Value = (int)TypeDoc.FromWrapping;
                    comm.Parameters.Add("@dateDoc", SqlDbType.DateTime).Value = sDate;
                    ds.Clear();
                    Database2.ExecuteCommand(comm, ref ds, null);
                    foreach (MyProd mp in al)
                    {
                        for (i = 0; ds.Tables.Count > 0 && i < ds.Tables[0].Rows.Count; i++)
                        {
                            if (mp.ID == Convert.ToInt32(ds.Tables[0].Rows[i]["id_prb"]))
                            {
                                int pr = (ds.Tables[0].Rows[i]["cntperso"] == DBNull.Value) ? 0 : Convert.ToInt32(ds.Tables[0].Rows[i]["cntperso"]);
                                 branch.addCount(mp.Type, mp.Name, 10, pr);
                                break;
                            }
                        }
                    }


                    //-----------------------------------------
                    if (branch.isEmpty() == true) throw new Exception("За данное число нет движения по продуктам");
                    //-----------------------------------------


                    FileHelper.initDefault();
                    FileHelper.SetDirectoryApp(ConfigurationSettings.AppSettings["DocPath"]);
                    FileHelper.SetDataFolder("temp\\movstdbf\\" + Path.GetFileNameWithoutExtension(Path.GetRandomFileName()), true);
                    FileHelper.setCode866(fileDBF);
                    FileHelper.rename(fileDBF, n_fileDBF);
                    odbf.Open(FileHelper.getFullName(n_fileDBF), FileMode.Open);
                    orec = new DbfRecord(odbf.Header);
                    for (i = 0; i < odbf.Header.RecordCount; i++)
                    {
                        if (!odbf.Read(i, orec)) break;
                        orec.IsDeleted = true;
                        odbf.Write(orec);
                    }

                    int numb_docum = 1;

                    foreach (FieldElement field in mvConfiguration.Fields)
                    {
                        int[] count = null;
                        if (field.Product.Equals(BaseProductType.MasterCard.ToString())==true) { if (branch.isEmptyMasterCard() == false) count = branch.countMasterCard; }
                        if (field.Product.Equals(BaseProductType.VisaCard.ToString()) == true) { if (branch.isEmptyVisaCard() == false) count = branch.countVisaCard; }
                        if (field.Product.Equals(BaseProductType.MirCard.ToString()) == true) { if (branch.isEmptyMirCard() == false) count = branch.countMirCard; }
                        if (field.Product.Equals(BaseProductType.NFCCard.ToString()) == true) { if (branch.isEmptyNFCCard() == false) count = branch.countNFCCard; }
                        if (field.Product.Equals(BaseProductType.ServiceCard.ToString()) == true) { if (branch.isEmptyServiceCard() == false) count = branch.countServiceCard; }
                        if (field.Product.Equals(BaseProductType.PinConvert.ToString()) == true) { if (branch.isEmptyPinConvert() == false) count = branch.countPinConvert; }
                         
                        foreach (ColumnElement column in field.Columns)
                        {
                                                        
                            if (count != null)
                            {
                                int countP = 0;
                                String[] nums = column.Nums.Split('+');
                                for (int n = 0; n < nums.Length; n++)
                                {
                                    int c=0;
                                    if (true == Int32.TryParse(nums[n], out c))
                                    {
                                        c--;
                                        if (c >= 0 && c < count.Length) countP += count[c];
                                    }
                                }
                                if (countP > 0)
                                {
                                    String debet = column.Debet, credit = column.Credit, ground = column.Ground;
                                    orec.Clear();
                                    orec[orec.FindColumn("mfo_payer")] = "049205805";
                                    orec[orec.FindColumn("acc_payer")] = debet;
                                    orec[orec.FindColumn("mfo_rec")] = "049205805";
                                    orec[orec.FindColumn("acc_rec")] = credit;
                                    orec[orec.FindColumn("sum")] = String.Format("{0}.00", countP);
                                    orec[orec.FindColumn("sumdeb")] = String.Format("{0}.00", countP);
                                    orec[orec.FindColumn("sumkred")] = String.Format("{0}.00", countP);
                                    orec[orec.FindColumn("curdeb")] = "810";
                                    orec[orec.FindColumn("curkred")] = "810";
                                    orec[orec.FindColumn("srok")] = sDate.ToString("yyyy-MM-dd");
                                    orec[orec.FindColumn("date_carry")] = sDate.ToString("yyyy-MM-dd");
                                    orec[orec.FindColumn("date_docum")] = sDate.ToString("yyyy-MM-dd");
                                    orec[orec.FindColumn("numb_docum")] = OperationDay.getBranchStartNumber(current_branch, sDate, numb_docum, 0);
                                    orec[orec.FindColumn("kind_oper")] = "09";
                                    orec[orec.FindColumn("ground")] = ground;
                                    odbf.Write(orec);
                                }
                            }
                            numb_docum++;
                        }
                    }
                }
            }
            catch (Exception e)
            {
                isError = true;
                String m = e.Message;
                m = m.Replace("\n", "");
                m = m.Replace("\r", " ");
                m = m.Replace("\t", "");
                ClientScript.RegisterClientScriptBlock(GetType(), "errMovSt", "<script type='text/javascript'>$(document).ready(function(){ ShowError('" + m + "');});</script>");
            }
            finally
            {
                ds.Clear();
                comm.Parameters.Clear();
                al.Clear();
                ds = null;
                comm = null;
                al = null;
                branch = null;
                odbf.Close();
                if (isError == false)
                {
                    dirDBF.Text = FileHelper.GetDirectoryData() + zipArch;
                    ClientScript.RegisterClientScriptBlock(GetType(), "okMovSt", "<script type='text/javascript'>$(document).ready(function(){ $('#" + bOkDirDBF.ClientID + "').click(); });</script>");
                    FileHelper.setCodeDefault(n_fileDBF);
                    using (var zip = new ZipFile())
                    {
                        zip.AddDirectory(FileHelper.GetDirectoryData());
                        zip.Save(FileHelper.getFullName(zipArch));
                    }
                    if (fname.Length > 0)
                    {
                        File.Copy(dirDBF.Text, fname, true);
                        dirDBF.Text = fname;
                    }
                    // Отправляем по почте
                    {
                        lbInform.Text = "";
                        MailMessage mm = new MailMessage();
                        DataSet ds1 = new DataSet();
                        res = Database2.ExecuteQuery(String.Format("select UserName from V_UserAction where ActionID={0}", (int)Restrictions.ConfirmSendingValues), ref ds1, null);
                        if (ds1.Tables.Count == 0)
                        {
                            lbInform.Text = "Данные рассылки не найдены";

                        }
                        else
                        {
                            for (int t = 0; t < ds1.Tables[0].Rows.Count; t++)
                            {
                                object obj = null;
                                Database2.ExecuteScalar(String.Format("select EMail from vw_aspnet_MembershipUsers where UserName='{0}'", ds1.Tables[0].Rows[t]["UserName"].ToString()), ref obj, null);
                                if (obj != null)
                                {
                                    try
                                    {
                                        MailAddress mTo = new MailAddress((string)obj);
                                        mm.Bcc.Add(mTo);
                                    }
                                    catch { }
                                }
                            }
                            try
                            {
                                SmtpClient sc = new SmtpClient(ConfigurationSettings.AppSettings["SmtpServer"]);
                                sc.Credentials = new NetworkCredential(ConfigurationSettings.AppSettings["EMailFrom"], ConfigurationSettings.AppSettings["Pwd"]);
                                MailAddress mailFrom = new MailAddress(ConfigurationSettings.AppSettings["EMailFrom"], "CardPerso");
                                mm.From = mailFrom;
                                mm.Subject = "Экспортный архив - движение ценностей";
                                mm.Body = String.Format("В системе CardPerso сформирован экспортный архив Движение ценностей.\n{0:HH.mm dd.MM.yyyy}", DateTime.Now);
                                mm.Attachments.Add(new Attachment(dirDBF.Text));
                                if (mm.Bcc.Count > 0)
                                {
                                    sc.Send(mm);
                                    lbInform.Text = "Экспортный архив Движение ценностей отправлен по рассылке";
                                }
                            }
                            catch (Exception Ex)
                            {
                                lbInform.Text = Ex.Message;
                            }
                        }
                    }
                }
            }
        }
        
        private void MovingStock(ExcelAp ep)
        {
            DataSet ds = new DataSet();
            int t = 0, i = 0;
            string res = "";
            SqlCommand comm = new SqlCommand();
            //проверяем, что есть неподтвержденные документы на персонализацию или отправку в филиал
            comm.CommandText = "select number_doc from StorageDocs where date_doc=@d and priz_gen=0 and (type=5 or type=8)";
            comm.Parameters.Add("@d", SqlDbType.DateTime).Value = DatePickerOne.SelectedDate.Date;
            Database2.ExecuteCommand(comm, ref ds, null);
            if (ds.Tables.Count > 0 && ds.Tables[0].Rows.Count > 0)
            {
                string str = "";
                for (i = 0; i < ds.Tables[0].Rows.Count; i++)
                    str += ", " + ds.Tables[0].Rows[i][0].ToString();
                str = str.Substring(2);
                lInfo.Text = String.Format("Создание отчета невозможно. За данное число есть неподтвержденные документы ({0})", str);
                throw new Exception(lInfo.Text);
                return;
            }
            ds.Clear();
            comm.CommandText = "select id_prb, prod_name, bank_name, id_type from V_ProductsBanks_T order by id_sort";
            res = (string)Database2.ExecuteCommand(comm, ref ds, null);
            ArrayList al = new ArrayList();
            foreach (DataRow dr in ds.Tables[0].Rows)
                al.Add(new MyProd(Convert.ToInt32(dr["id_prb"]), dr["prod_name"].ToString(), dr["bank_name"].ToString(), Convert.ToInt32(dr["id_type"])));
            //выдача на персонализацию
            comm.Parameters.Add("@type", SqlDbType.Int).Value = (int)TypeDoc.PersoCard;
            comm.Parameters.Add("@dateDoc", SqlDbType.DateTime).Value = DatePickerOne.SelectedDate.Date;
            //comm.CommandText = "select sum(cnt_brak) as cntbrak, sum(cnt_perso) as cntperso from V_Rep_Moving where priz_gen=1 and type=@type and (id_type=1 or id_type=2) and date_doc=@dateDoc and id_prb=@prb";
            comm.CommandText = "select id_prb,sum(cnt_brak) as cntbrak, sum(cnt_perso) as cntperso from V_Rep_Moving where priz_gen=1 and type=@type and (id_type=1 or id_type=2) and date_doc=@dateDoc group by id_prb";
            comm.Parameters.Add("@prb", SqlDbType.Int).Value = 0;
            i = 0;
            
            ds.Clear();
            Database2.ExecuteCommand(comm, ref ds, null);

            foreach (MyProd mp in al)
            {
                //                            if (mp.Type != 1)
                //                                continue;
                /*
                ds.Clear();
                comm.Parameters["@prb"].Value = mp.ID;
                res = (string)Database2.ExecuteCommand(comm, ref ds, null);
                if (ds.Tables[0].Rows.Count == 1)
                {
                    int br = (ds.Tables[0].Rows[0]["cntbrak"] == DBNull.Value) ? 0 : Convert.ToInt32(ds.Tables[0].Rows[0]["cntbrak"]);
                    int pr = (ds.Tables[0].Rows[0]["cntperso"] == DBNull.Value) ? 0 : Convert.ToInt32(ds.Tables[0].Rows[0]["cntperso"]);
                    mp.cnts[0] = br + pr;
                }
                */
                for(i=0;ds.Tables.Count>0 && i< ds.Tables[0].Rows.Count;i++)
                {
                    if(mp.ID==Convert.ToInt32(ds.Tables[0].Rows[i]["id_prb"]))
                    {
                        int br = (ds.Tables[0].Rows[i]["cntbrak"] == DBNull.Value) ? 0 : Convert.ToInt32(ds.Tables[0].Rows[i]["cntbrak"]);
                        int pr = (ds.Tables[0].Rows[i]["cntperso"] == DBNull.Value) ? 0 : Convert.ToInt32(ds.Tables[0].Rows[i]["cntperso"]);
                        mp.cnts[0] = br + pr;
                        break;
                    }
                }
 
            }
            // выдача головному офису
            ds.Clear();
            //comm.CommandText = "select sum(cnt_perso) as cntperso from V_Rep_Moving where priz_gen=1 and type=@type and (id_type=1 or id_type=2) and date_doc=@dateDoc and id_prb=@prb and is_head=@ishead";
            comm.CommandText = "select id_prb,sum(cnt_perso) as cntperso from V_Rep_Moving where priz_gen=1 and type=@type and (id_type=1 or id_type=2) and date_doc=@dateDoc and is_head=@ishead group by id_prb";
            comm.Parameters["@type"].Value = (int)TypeDoc.SendToFilial;
            comm.Parameters.Add("@ishead", SqlDbType.Bit).Value = true;

            Database2.ExecuteCommand(comm, ref ds, null);

            foreach (MyProd mp in al)
            {
                //                            if (mp.Type != 1)
                //                                continue;
                /*
                ds.Clear();
                comm.Parameters["@prb"].Value = mp.ID;
                res = (string)Database2.ExecuteCommand(comm, ref ds, null);
                if (ds.Tables[0].Rows.Count == 1)
                {
                    int pr = (ds.Tables[0].Rows[0]["cntperso"] == DBNull.Value) ? 0 : Convert.ToInt32(ds.Tables[0].Rows[0]["cntperso"]);
                    if (pr > 0)
                        mp.cnts[1] = pr;
                }
                */
                for (i = 0; ds.Tables.Count > 0 && i < ds.Tables[0].Rows.Count; i++)
                {
                    if (mp.ID == Convert.ToInt32(ds.Tables[0].Rows[i]["id_prb"]))
                    {
                        int pr = (ds.Tables[0].Rows[i]["cntperso"] == DBNull.Value) ? 0 : Convert.ToInt32(ds.Tables[0].Rows[i]["cntperso"]);
                        if (pr > 0)
                            mp.cnts[1] = pr;
                        break;
                    }
                }
            }
            // выдача территориальным подразделениям
            //comm.CommandText = "select sum(cnt_perso) as cntperso from V_Rep_Moving where priz_gen=1 and type=@type and (id_type=1 or id_type=2) and date_doc=@dateDoc and id_prb=@prb and (is_head=@ishead or is_head is null) and (is_rkc=@ishead or is_rkc is null) and (is_trans=@ishead or is_trans is null)";
            comm.CommandText = "select id_prb,sum(cnt_perso) as cntperso from V_Rep_Moving where priz_gen=1 and type=@type and (id_type=1 or id_type=2) and date_doc=@dateDoc and (is_head=@ishead or is_head is null) and (is_rkc=@ishead or is_rkc is null) and (is_trans=@ishead or is_trans is null) group by id_prb";
            comm.Parameters["@ishead"].Value = false;

            ds.Clear();
            Database2.ExecuteCommand(comm, ref ds, null);
            
            foreach (MyProd mp in al)
            {
                //                            if (mp.Type != 1)
                //                                continue;
                /*
                ds.Clear();
                comm.Parameters["@prb"].Value = mp.ID;
                res = (string)Database2.ExecuteCommand(comm, ref ds, null);
                if (ds.Tables[0].Rows.Count == 1)
                {
                    int pr = (ds.Tables[0].Rows[0]["cntperso"] == DBNull.Value) ? 0 : Convert.ToInt32(ds.Tables[0].Rows[0]["cntperso"]);
                    if (pr > 0)
                        mp.cnts[3] = pr;
                }
                */
                for (i = 0; ds.Tables.Count > 0 && i < ds.Tables[0].Rows.Count; i++)
                {
                    if (mp.ID == Convert.ToInt32(ds.Tables[0].Rows[i]["id_prb"]))
                    {
                        int pr = (ds.Tables[0].Rows[i]["cntperso"] == DBNull.Value) ? 0 : Convert.ToInt32(ds.Tables[0].Rows[i]["cntperso"]);
                        if (pr > 0)
                            mp.cnts[3] = pr;
                        break;
                    }
                }
            }
            // выдача ГО ДРБ РКЦ
            //comm.CommandText = "select sum(cnt_perso) as cntperso from V_Rep_Moving where priz_gen=1 and type=@type and (id_type=1 or id_type=2) and date_doc=@dateDoc and id_prb=@prb and is_rkc=@ishead";
            comm.CommandText = "select id_prb,sum(cnt_perso) as cntperso from V_Rep_Moving where priz_gen=1 and type=@type and (id_type=1 or id_type=2) and date_doc=@dateDoc and is_rkc=@ishead group by id_prb";
            comm.Parameters["@ishead"].Value = true;

            ds.Clear();
            Database2.ExecuteCommand(comm, ref ds, null);

            foreach (MyProd mp in al)
            {
                //                            if (mp.Type != 1)
                //                                continue;
                /*
                ds.Clear();
                comm.Parameters["@prb"].Value = mp.ID;
                res = (string)Database2.ExecuteCommand(comm, ref ds, null);
                if (ds.Tables[0].Rows.Count == 1)
                {
                    int pr = (ds.Tables[0].Rows[0]["cntperso"] == DBNull.Value) ? 0 : Convert.ToInt32(ds.Tables[0].Rows[0]["cntperso"]);
                    if (pr > 0)
                        mp.cnts[2] = pr;
                }
                */
                for (i = 0; ds.Tables.Count > 0 && i < ds.Tables[0].Rows.Count; i++)
                {
                    if (mp.ID == Convert.ToInt32(ds.Tables[0].Rows[i]["id_prb"]))
                    {
                        int pr = (ds.Tables[0].Rows[i]["cntperso"] == DBNull.Value) ? 0 : Convert.ToInt32(ds.Tables[0].Rows[i]["cntperso"]);
                        if (pr > 0)
                            mp.cnts[2] = pr;
                        break;
                    }
                }

            }
            // выдача транспортных карт
            //comm.CommandText = "select sum(cnt_perso) as cntperso from V_Rep_Moving where priz_gen=1 and type=@type and (id_type=1 or id_type=2) and date_doc=@dateDoc and id_prb=@prb and is_trans=@ishead";
            comm.CommandText = "select id_prb,sum(cnt_perso) as cntperso from V_Rep_Moving where priz_gen=1 and type=@type and (id_type=1 or id_type=2) and date_doc=@dateDoc and is_trans=@ishead group by id_prb";
            comm.Parameters["@ishead"].Value = true;

            ds.Clear();
            Database2.ExecuteCommand(comm, ref ds, null);


            foreach (MyProd mp in al)
            {
                //                            if (mp.Type != 1)
                //                                continue;
                /*
                ds.Clear();
                comm.Parameters["@prb"].Value = mp.ID;
                res = (string)Database2.ExecuteCommand(comm, ref ds, null);
                if (ds.Tables[0].Rows.Count == 1)
                {
                    int pr = (ds.Tables[0].Rows[0]["cntperso"] == DBNull.Value) ? 0 : Convert.ToInt32(ds.Tables[0].Rows[0]["cntperso"]);
                    if (pr > 0)
                        mp.cnts[4] = pr;
                }
                */
                for (i = 0; ds.Tables.Count > 0 && i < ds.Tables[0].Rows.Count; i++)
                {
                    if (mp.ID == Convert.ToInt32(ds.Tables[0].Rows[i]["id_prb"]))
                    {
                        int pr = (ds.Tables[0].Rows[i]["cntperso"] == DBNull.Value) ? 0 : Convert.ToInt32(ds.Tables[0].Rows[i]["cntperso"]);
                        if (pr > 0)
                            mp.cnts[4] = pr;
                        break;
                    }
                }
            }
            // прием в хранилище персонализированных карт, вернувшихся из филиалов и бракованных карт
            //comm.CommandText = "select sum(cnt_perso) as cntperso, sum(cnt_brak) as cntbrak from V_Rep_Moving where priz_gen=1 and type=@type and (id_type=1 or id_type=2) and date_doc=@dateDoc and id_prb=@prb";
            ////comm.CommandText = "select id_prb,sum(cnt_perso) as cntperso, sum(cnt_brak) as cntbrak from V_Rep_Moving where priz_gen=1 and type=@type and (id_type=1 or id_type=2) and date_doc=@dateDoc group by id_prb";
            /*
            if (branch_main_filial > 0)
            {
                comm.CommandText = String.Format("select id_prb,sum(cnt_perso) as cntperso, sum(cnt_brak) as cntbrak from V_Rep_Moving where priz_gen=1 and type=@type and (id_type=1 or id_type=2) and date_doc=@dateDoc and (id_branch={0} or id_branch in (select id from Branchs where id_parent={0})) group by id_prb", branch_main_filial);
            }
            else
            {
                comm.CommandText = String.Format("select id_prb,sum(cnt_perso) as cntperso, sum(cnt_brak) as cntbrak from V_Rep_Moving where priz_gen=1 and type=@type and (id_type=1 or id_type=2) and date_doc=@dateDoc and (id_branch!={0} and id_branch not in (select id from Branchs where id_parent={0})) group by id_prb", 106);
            }
            */
            comm.CommandText = "select id_prb,sum(cnt_perso) as cntperso, sum(cnt_brak) as cntbrak from V_Rep_Moving where priz_gen=1 and type=@type and (id_type=1 or id_type=2) and date_doc=@dateDoc group by id_prb";

            comm.Parameters["@type"].Value = (int)TypeDoc.ReceiveToBank;

            ds.Clear();
            Database2.ExecuteCommand(comm, ref ds, null);

            foreach (MyProd mp in al)
            {
                if (mp.Type != 1 && mp.Type != 2)
                    continue;
                /*
                ds.Clear();
                comm.Parameters["@prb"].Value = mp.ID;
                res = (string)Database2.ExecuteCommand(comm, ref ds, null);
                if (ds.Tables[0].Rows.Count == 1)
                {
                    int pr = (ds.Tables[0].Rows[0]["cntperso"] == DBNull.Value) ? 0 : Convert.ToInt32(ds.Tables[0].Rows[0]["cntperso"]);
                    if (pr > 0)
                        mp.cnts[8] = pr;
                    pr = (ds.Tables[0].Rows[0]["cntbrak"] == DBNull.Value) ? 0 : Convert.ToInt32(ds.Tables[0].Rows[0]["cntbrak"]);
                    if (pr > 0)
                        mp.cnts[9] = pr;
                }
                */
                for (i = 0; ds.Tables.Count > 0 && i < ds.Tables[0].Rows.Count; i++)
                {
                    if (mp.ID == Convert.ToInt32(ds.Tables[0].Rows[i]["id_prb"]))
                    {
                        int pr = (ds.Tables[0].Rows[i]["cntperso"] == DBNull.Value) ? 0 : Convert.ToInt32(ds.Tables[0].Rows[i]["cntperso"]);
                        if (pr > 0)
                            mp.cnts[8] = pr;
                        pr = (ds.Tables[0].Rows[i]["cntbrak"] == DBNull.Value) ? 0 : Convert.ToInt32(ds.Tables[0].Rows[i]["cntbrak"]);
                        if (pr > 0)
                            mp.cnts[9] = pr;
                        break;
                    }
                }
 
            }
            // выдача из хранилища карт на рекламу и тесты
            //comm.CommandText = "select sum(cnt_new) as cntnew from V_Rep_Moving where priz_gen=1 and type=@type and date_doc=@dateDoc and id_prb=@prb";
            comm.CommandText = "select id_prb,sum(cnt_new) as cntnew from V_Rep_Moving where priz_gen=1 and type=@type and date_doc=@dateDoc group by id_prb";
            comm.Parameters["@type"].Value = (int)TypeDoc.Reklama;

            ds.Clear();
            Database2.ExecuteCommand(comm, ref ds, null);

            foreach (MyProd mp in al)
            {
                /*
                ds.Clear();
                comm.Parameters["@prb"].Value = mp.ID;
                res = (string)Database2.ExecuteCommand(comm, ref ds, null);
                if (ds.Tables[0].Rows.Count == 1)
                {
                    int pr = (ds.Tables[0].Rows[0]["cntnew"] == DBNull.Value) ? 0 : Convert.ToInt32(ds.Tables[0].Rows[0]["cntnew"]);
                    if (pr > 0)
                        mp.cnts[5] = pr;
                }
                */
                for (i = 0; ds.Tables.Count > 0 && i < ds.Tables[0].Rows.Count; i++)
                {
                    if (mp.ID == Convert.ToInt32(ds.Tables[0].Rows[i]["id_prb"]))
                    {
                        int pr = (ds.Tables[0].Rows[i]["cntnew"] == DBNull.Value) ? 0 : Convert.ToInt32(ds.Tables[0].Rows[i]["cntnew"]);
                        if (pr > 0)
                            mp.cnts[5] = pr;
                        break;
                    }
                }
            }            
            // уничтожено брака (берется все в том числе не востребованные и с истекшим сроком, которые потом вычитаются из этих цифр
            // прием в хранилище персонализированных карт, вернувшихся из филиалов
            //comm.CommandText = "select sum(cnt_brak) as cntbrak from V_Rep_Moving where priz_gen=1 and type=@type and (id_type=1 or id_type=2) and date_doc=@dateDoc and id_prb=@prb";
            ////comm.CommandText = "select id_prb,sum(cnt_brak) as cntbrak from V_Rep_Moving where priz_gen=1 and type=@type and (id_type=1 or id_type=2) and date_doc=@dateDoc group by id_prb";
            /*
            if (branch_main_filial > 0)
            {
                comm.CommandText = String.Format("select id_prb,sum(cnt_brak) as cntbrak from V_Rep_Moving where priz_gen=1 and type=@type and (id_type=1 or id_type=2) and date_doc=@dateDoc and (id_branch={0} or id_branch in (select id from Branchs where id_parent={0})) group by id_prb", branch_main_filial);
            }
            else
            {
                comm.CommandText = String.Format("select id_prb,sum(cnt_brak) as cntbrak from V_Rep_Moving where priz_gen=1 and type=@type and (id_type=1 or id_type=2) and date_doc=@dateDoc and (id_branch is null or (id_branch!={0} and id_branch not in (select id from Branchs where id_parent={0}))) group by id_prb", 106);
            }
            */
            comm.CommandText = "select id_prb,sum(cnt_brak) as cntbrak from V_Rep_Moving where priz_gen=1 and type=@type and (id_type=1 or id_type=2) and date_doc=@dateDoc group by id_prb";

            comm.Parameters["@type"].Value = (int)TypeDoc.DeleteBrak;

            ds.Clear();
            Database2.ExecuteCommand(comm, ref ds, null);

            foreach (MyProd mp in al)
            {
                if (mp.Type != 1 && mp.Type != 2)
                    continue;
                /*
                ds.Clear();
                comm.Parameters["@prb"].Value = mp.ID;
                res = (string)Database2.ExecuteCommand(comm, ref ds, null);
                if (ds.Tables[0].Rows.Count == 1)
                {
                    int pr = (ds.Tables[0].Rows[0]["cntbrak"] == DBNull.Value) ? 0 : Convert.ToInt32(ds.Tables[0].Rows[0]["cntbrak"]);
                    if (pr > 0)
                        mp.cnts[11] = pr;
                }
                */
                for (i = 0; ds.Tables.Count > 0 && i < ds.Tables[0].Rows.Count; i++)
                {
                    if (mp.ID == Convert.ToInt32(ds.Tables[0].Rows[i]["id_prb"]))
                    {
                        int pr = (ds.Tables[0].Rows[i]["cntbrak"] == DBNull.Value) ? 0 : Convert.ToInt32(ds.Tables[0].Rows[i]["cntbrak"]);
                        if (pr > 0)
                            mp.cnts[11] = pr;
                        break;
                    }
                }

            }
            // уничтожено карт не востребованных и с истекшим сроком годности
            object obj = 0;
            comm.Parameters.Clear();
            //comm.CommandText = "select count(*) as cnt from Cards where id_prb=@prb and dateTerminated=@date and id_prop=@prop";
            /*
            if (branch_main_filial > 0)
            {
                comm.CommandText = String.Format("select count(*) as cnt from Cards where id_prb=@prb and dateTerminated=@date and id_prop=@prop and id_BranchCurrent={0}", branch_main_filial);
            }
            else
            {
                comm.CommandText = String.Format("select count(*) as cnt from Cards where id_prb=@prb and dateTerminated=@date and id_prop=@prop and id_BranchCurrent!={0}", 106);            
            }
            */
            // Изменено 25.11.2019, потому что стало сказываться уничтожение карт в филиале
            /*
            comm.CommandText = "select count(*) as cnt from Cards where id_prb=@prb and dateTerminated=@date and id_prop=@prop";

            comm.Parameters.Add("@prb", SqlDbType.Int);
            comm.Parameters.Add("@prop", SqlDbType.Int);
            comm.Parameters.Add("@date", SqlDbType.DateTime).Value = DatePickerOne.SelectedDate.Date;
            */
            comm.CommandText = $"select count(*) from cards c inner join Cards_StorageDocs cs on c.id=cs.id_card inner join StorageDocs s on cs.id_doc=s.id " +
                "where s.date_doc = @date and (s.type = 11) and s.priz_gen = 1 and id_prop=@prop and id_prb=@prb";

            comm.Parameters.Add("@prb", SqlDbType.Int);
            comm.Parameters.Add("@prop", SqlDbType.Int);
            comm.Parameters.Add("@date", SqlDbType.DateTime).Value = DatePickerOne.SelectedDate.Date;

            foreach (MyProd mp in al)
            {
                if (mp.Type != 1)
                    continue;
                comm.Parameters["@prb"].Value = mp.ID;
                comm.Parameters["@prop"].Value = 7;
                res = Database2.ExecuteScalar(comm, ref obj, null);
                mp.cnts[12] = Convert.ToInt32(obj);
                mp.cnts[11] -= Convert.ToInt32(obj);
                comm.Parameters["@prop"].Value = 8;
                res = Database2.ExecuteScalar(comm, ref obj, null);
                mp.cnts[13] = Convert.ToInt32(obj);
                mp.cnts[11] -= Convert.ToInt32(obj);
            }
            comm.Parameters.Clear();
            // и отдельно считаем пины
            comm.CommandText = $"select count(*) from cards c inner join Cards_StorageDocs cs on c.id=cs.id_card inner join StorageDocs s on cs.id_doc=s.id " +
                "where s.date_doc = @date and (s.type = 11) and s.priz_gen = 1 and c.id_prop=@prop and c.isPin=1";

            comm.Parameters.Add("@prop", SqlDbType.Int);
            comm.Parameters.Add("@date", SqlDbType.DateTime).Value = DatePickerOne.SelectedDate.Date;

            foreach (MyProd mp in al)
            {
                if (mp.Type != 2)
                    continue;
                comm.Parameters["@prop"].Value = 7;
                res = Database2.ExecuteScalar(comm, ref obj, null);
                mp.cnts[12] = Convert.ToInt32(obj);
                mp.cnts[11] -= Convert.ToInt32(obj);
                comm.Parameters["@prop"].Value = 8;
                res = Database2.ExecuteScalar(comm, ref obj, null);
                mp.cnts[13] = Convert.ToInt32(obj);
                mp.cnts[11] -= Convert.ToInt32(obj);
            }
            comm.Parameters.Clear();
            // карты ожидающие уничтожения
            /*
            if (branch_main_filial > 0)
            {
                comm.CommandText = String.Format("select id_prb,count(*) as ccard,case when sum(cast(ispin as integer)) is null then 0 else sum(cast(ispin as integer)) end as cpin from Cards where id_stat=6 and id_branchCurrent={0} group by id_prb", branch_main_filial);
            }
            else
            {
                comm.CommandText = String.Format("select id_prb,count(*) as ccard,case when sum(cast(ispin as integer)) is null then 0 else sum(cast(ispin as integer)) end as cpin from Cards where id_stat=6 and id_branchCurrent not in ({0}) group by id_prb", 106);
            }
            */
            comm.CommandText = "select id_prb,count(*) as ccard,case when sum(cast(ispin as integer)) is null then 0 else sum(cast(ispin as integer)) end as cpin from Cards where id_stat=6 and (id_BranchCurrent = 0 or id_BranchCurrent is NULL) group by id_prb";

            ds.Clear();
            Database2.ExecuteCommand(comm, ref ds, null);

            int pinc = 0;

            foreach (MyProd mp in al)
            {
                if (mp.Type != 1) continue;
                for (i = 0; ds.Tables.Count > 0 && i < ds.Tables[0].Rows.Count; i++)
                {
                    if (mp.ID == Convert.ToInt32(ds.Tables[0].Rows[i]["id_prb"]))
                    {
                        int pr = (ds.Tables[0].Rows[i]["ccard"] == DBNull.Value) ? 0 : Convert.ToInt32(ds.Tables[0].Rows[i]["ccard"]);
                        if (pr > 0)
                            mp.cnts[14] = pr;
                        pr = (ds.Tables[0].Rows[i]["cpin"] == DBNull.Value) ? 0 : Convert.ToInt32(ds.Tables[0].Rows[i]["cpin"]);
                        pinc += pr;
                        break;
                    }
                }
            }
            foreach (MyProd mp in al)
            {
                if (mp.Type == 2)
                {
                    mp.cnts[14] = pinc; break;
                }
            }
            

            // всего карт с истекшим сроком годности
            //comm.CommandText = "select count(*) as cnt from Cards where id_prb=@prb and dateEnd<@date and id_stat<>@prop";
            //comm.Parameters["@date"].Value = DateTime.Now.Date;
            //comm.Parameters["@prop"].Value = 7;
            //foreach (MyProd mp in al)
            //{
            //  if (mp.Type != 1)
            //      continue;
            //    comm.Parameters["@prb"].Value = mp.ID;
            //    res = Database2.ExecuteScalar(comm, ref obj, null);
            //    mp.cnts[10] = Convert.ToInt32(obj);
            //}
            // закупленные карты и все остальное
            ds.Clear();
            comm.Parameters.Clear();
            comm.CommandText = "select id from PurchDogs where date_stor=@date";
            comm.Parameters.Add("@date", SqlDbType.DateTime).Value = DatePickerOne.SelectedDate.Date;
            Database2.ExecuteCommand(comm, ref ds, null);
            if (ds != null && ds.Tables.Count > 0 && ds.Tables[0].Rows.Count > 0)
            {
                string[] strs = new string[ds.Tables[0].Rows.Count];
                for (i = 0; i < ds.Tables[0].Rows.Count; i++)
                    strs[i] = ds.Tables[0].Rows[i]["id"].ToString();
                comm.CommandText = String.Format("select cnt, id_prb from Products_PurchDogs where id_dog in ({0})", String.Join(",", strs));
                ds.Clear();
                Database2.ExecuteCommand(comm, ref ds, null);
                if (ds != null && ds.Tables.Count > 0)
                {
                    for (i = 0; i < ds.Tables[0].Rows.Count; i++)
                    {
                        t = 0;
                        for (t = 0; t < al.Count; t++)
                        {
                            if (((MyProd)al[t]).ID == Convert.ToInt32(ds.Tables[0].Rows[i]["id_prb"]))
                            {
                                ((MyProd)al[t]).cnts[7] += Convert.ToInt32(ds.Tables[0].Rows[i]["cnt"]);
                                break;
                            }
                        }
                    }
                }
            }
            // на упаковку
            ds.Clear();
            comm.Parameters.Clear();
            //comm.CommandText = "select sum(cnt_perso) as cntperso from V_Rep_Moving where priz_gen=1 and type=@type and date_doc=@dateDoc and id_prb=@prb";
            comm.CommandText = "select id_prb,sum(cnt_perso) as cntperso from V_Rep_Moving where priz_gen=1 and type=@type and date_doc=@dateDoc group by id_prb";
            //comm.CommandText = "select count(*) as cnt from Cards where id_prb=@prb and dateToWrapping=@date";
            comm.Parameters.Add("@type", SqlDbType.Int).Value = (int)TypeDoc.ToWrapping;
            comm.Parameters.Add("@dateDoc", SqlDbType.DateTime).Value = DatePickerOne.SelectedDate.Date;
            //comm.Parameters.Add("@prb", SqlDbType.Int);

            ds.Clear();
            Database2.ExecuteCommand(comm, ref ds, null);
            
            foreach (MyProd mp in al)
            {
//                if (mp.Type != 1)
//                    continue;
                /*
                comm.Parameters["@prb"].Value = mp.ID;
                res = Database2.ExecuteScalar(comm, ref obj, null);
                if (obj == DBNull.Value)
                    continue;
                mp.cnts[6] = Convert.ToInt32(obj);
               */
                for (i = 0; ds.Tables.Count > 0 && i < ds.Tables[0].Rows.Count; i++)
                {
                    if (mp.ID == Convert.ToInt32(ds.Tables[0].Rows[i]["id_prb"]))
                    {
                        int pr = (ds.Tables[0].Rows[i]["cntperso"] == DBNull.Value) ? 0 : Convert.ToInt32(ds.Tables[0].Rows[i]["cntperso"]);
                        mp.cnts[6] = pr;
                        break;
                    }
                }
            }
            // c упаковку
            ds.Clear();
            comm.Parameters.Clear();
            //comm.CommandText = "select sum(cnt_perso) as cntperso from V_Rep_Moving where priz_gen=1 and type=@type and date_doc=@dateDoc and id_prb=@prb";
            comm.CommandText = "select id_prb,sum(cnt_perso) as cntperso from V_Rep_Moving where priz_gen=1 and type=@type and date_doc=@dateDoc group by id_prb";
            //comm.CommandText = "select count(*) as cnt from Cards where id_prb=@prb and dateFromWrapping=@date";
            comm.Parameters.Add("@type", SqlDbType.Int).Value = (int)TypeDoc.FromWrapping;
            comm.Parameters.Add("@dateDoc", SqlDbType.DateTime).Value = DatePickerOne.SelectedDate.Date;
            //comm.Parameters.Add("@prb", SqlDbType.Int);
            ds.Clear();
            Database2.ExecuteCommand(comm, ref ds, null);

            foreach (MyProd mp in al)
            {
//                if (mp.Type != 1)
//                    continue;
                /*
                comm.Parameters["@prb"].Value = mp.ID;
                res = Database2.ExecuteScalar(comm, ref obj, null);
                if (obj == DBNull.Value)
                    continue;
                mp.cnts[10] = Convert.ToInt32(obj);
                */
                for (i = 0; ds.Tables.Count > 0 && i < ds.Tables[0].Rows.Count; i++)
                {
                    if (mp.ID == Convert.ToInt32(ds.Tables[0].Rows[i]["id_prb"]))
                    {
                        int pr = (ds.Tables[0].Rows[i]["cntperso"] == DBNull.Value) ? 0 : Convert.ToInt32(ds.Tables[0].Rows[i]["cntperso"]);
                        mp.cnts[10] = pr;
                        break;
                    }
                }
            }

            // прием на экспертизу плюсуем к 8 полю (прием персонализированных карт)
            ds.Clear();
            comm.Parameters.Clear();
            //comm.CommandText = "select sum(cnt_perso) as cntperso from V_Rep_Moving where priz_gen=1 and type=@type and date_doc=@dateDoc and id_prb=@prb";
            comm.CommandText = "select id_prb,sum(cnt_perso) as cntperso from V_Rep_Moving where priz_gen=1 and type=@type and date_doc=@dateDoc group by id_prb";
            //comm.CommandText = "select count(*) as cnt from Cards where id_prb=@prb and dateFromWrapping=@date";
            comm.Parameters.Add("@type", SqlDbType.Int).Value = (int)TypeDoc.ReceiveToExpertiza;
            comm.Parameters.Add("@dateDoc", SqlDbType.DateTime).Value = DatePickerOne.SelectedDate.Date;
            //comm.Parameters.Add("@prb", SqlDbType.Int);
            ds.Clear();
            Database2.ExecuteCommand(comm, ref ds, null);

            foreach (MyProd mp in al)
            {                
                for (i = 0; ds.Tables.Count > 0 && i < ds.Tables[0].Rows.Count; i++)
                {
                    if (mp.ID == Convert.ToInt32(ds.Tables[0].Rows[i]["id_prb"]))
                    {
                        int pr = (ds.Tables[0].Rows[i]["cntperso"] == DBNull.Value) ? 0 : Convert.ToInt32(ds.Tables[0].Rows[i]["cntperso"]);
                        mp.cnts[8] += pr;
                        break;
                    }
                }
            }
            comm.Parameters.Clear();
            bool flag = false;
            t = 20;
            int[] cnts = new int[MyProd.cntMax];
            int[] mc = new int[MyProd.cntMax];
            int[] visa = new int[MyProd.cntMax];
            int[] pins = new int[MyProd.cntMax];
            int[] serv = new int[MyProd.cntMax];
            int[] nfc = new int[MyProd.cntMax];
            int[] mir = new int[MyProd.cntMax];
            for (i = 0; i < MyProd.cntMax; i++)
            {
                cnts[i] = 0;
                mc[i] = 0;
                visa[i] = 0;
                pins[i] = 0;
                serv[i] = 0;
                nfc[i] = 0;
                mir[i] = 0;
            }
            
            foreach (MyProd mp in al)
            {
                if (mp.Type != 1 && mp.Type != 2)
                    continue;
                flag = false;
                for (i = 0; i < MyProd.cntMax; i++)
                    if (mp.cnts[i] > 0)
                    {
                        flag = true;
                        break;
                    }
                if (flag)
                {
                    ep.SetText(t, 1, String.Format("{0}({1})", mp.Name, mp.Bank));
                    for (i = 0; i < MyProd.cntMax; i++)
                        if (mp.cnts[i] > 0)
                        {
                            if (mp.Type == 1)
                            {
                                cnts[i] += mp.cnts[i];
                                /*
                                if (mp.Name.ToLower().StartsWith("mc") || mp.Name.ToLower().StartsWith("master") || mp.Name.ToLower().StartsWith("mifare standart"))
                                    mc[i] += mp.cnts[i];
                                if (mp.Name.ToLower().StartsWith("visa"))
                                    visa[i] += mp.cnts[i];
                                if (mp.Name.ToLower().StartsWith("белый") || mp.Name.ToLower().IndexOf("карты типа восток") >= 0)
                                    serv[i] += mp.cnts[i];
                                if (mp.Name.ToLower().StartsWith("nfs") | mp.Name.ToLower().StartsWith("nfc"))
                                    nfc[i] += mp.cnts[i];
                                */
                                BaseProductType tp = BranchStore.codeFromTypeAndProdName(1, mp.Name);
                                switch (tp)
                                {
                                    case BaseProductType.MasterCard: mc[i] += mp.cnts[i]; break;
                                    case BaseProductType.VisaCard: visa[i] += mp.cnts[i]; break;
                                    case BaseProductType.NFCCard: nfc[i] += mp.cnts[i]; break;
                                    case BaseProductType.ServiceCard: serv[i] += mp.cnts[i]; break;
                                    case BaseProductType.MirCard: mir[i] += mp.cnts[i]; break;

                                }
                           
                            }
                            if (mp.Type == 2)
                            {
                                pins[i] += mp.cnts[i];
                            }
                            ep.SetText(t, 3 + i, String.Format("{0}", mp.cnts[i]));
                        }
                    t++;
                }
            }
            ep.ShowRows(20, t-1);
            for (i = 0; i < MyProd.cntMax; i++)
            {
                if (cnts[i] > 0)
                    ep.SetText(100, 3 + i, String.Format("{0}", cnts[i]));
            }           
/*            //пин-конверты
            foreach (MyProd mp in al)
            {
                if (mp.Type != 2)
                    continue;
                if (mp.cnts[7] > 0)
                {
                    ep.SetText(101, 10, String.Format("{0}", mp.cnts[7]));
                    ep.ShowRows(101,101);
                }
            }*/
            ep.SetText_Name("date_string", String.Format("от {0:dd MM yyyy}", DatePickerOne.SelectedDate));
            ep.SetText_Name("date_string1", String.Format("{0:dd.MM.yyyy}", DatePickerOne.SelectedDate));
            ///////////////////////////
            //третий лист - отчет сборный по MC, Visa, NFC
            ///////////////////////////
            ep.SetWorkSheet(3);
            t = 21;
            ep.SetText(t, 1, "MasterCard");
            for (i = 0; i < MyProd.cntMax; i++)
            {
                if (mc[i] > 0)
                    ep.SetText(t, 3 + i, String.Format("{0}", mc[i]));
            }
            t++;
            ep.SetText(t, 1, "Visa");
            for (i = 0; i < MyProd.cntMax; i++)
            {
                if (visa[i] > 0)
                    ep.SetText(t, 3 + i, String.Format("{0}", visa[i]));
            }

            t++;
            ep.SetText(t, 1, "Сервисные карты");
            for (i = 0; i < MyProd.cntMax; i++)
            {
                if (serv[i] > 0)
                    ep.SetText(t, 3 + i, String.Format("{0}", serv[i]));
            }

            t++;
            ep.SetText(t, 1, "NFC карты");
            for (i = 0; i < MyProd.cntMax; i++)
            {
                if (nfc[i] > 0)
                    ep.SetText(t, 3 + i, String.Format("{0}", nfc[i]));
                    
            }
            
            t++;

            ep.SetText(t, 1, "Карты МИР");
            for (i = 0; i < MyProd.cntMax; i++)
            {
                if (mir[i] > 0)
                    ep.SetText(t, 3 + i, String.Format("{0}", mir[i]));
            }

            t++;

            ep.SetText(t, 1, "Пин-конверты");
            for (i = 0; i < MyProd.cntMax; i++)
            {
                if (pins[i] > 0)
                    ep.SetText(t, 3 + i, String.Format("{0}", pins[i]));
            }

            t++;
            ep.ShowRows(20, t-1);
            for (i = 0; i < MyProd.cntMax; i++)
            {
                if (mc[i]+visa[i]+serv[i]+nfc[i]+mir[i] > 0)
                    ep.SetText(101, 3 + i, String.Format("{0}", mc[i] + visa[i] + serv[i] + nfc[i] + mir[i]));
            }
            ep.SetText_Name("date_string3", String.Format("от {0:dd MM yyyy}", DatePickerOne.SelectedDate));
            ep.SetText_Name("date_string4", String.Format("{0:dd.MM.yyyy}", DatePickerOne.SelectedDate));
            #region Storage
            /////////////////////////////////
            //второй лист - хранилище
            /////////////////////////////////
            int mc1 = 0, mc2 = 0;
            int v1 = 0, v2 = 0;
            int ser1 = 0, ser2 = 0;
            int pin1 = 0, pin2 = 0;
            int nfc1 = 0, nfc2 = 0;
            int mir1 = 0, mir2 = 0;
            ds.Clear();
            ep.SetWorkSheet(2);
            Database2.ExecuteQuery("select name, bank_name, cnt_new, cnt_perso from V_Storage where id_type=1 order by id_sort", ref ds, null);
            int rw = 4;
            int cnt1 = 0, cnt2 = 0;
            for (i = 0; i < ds.Tables[0].Rows.Count; i++)
            {
                string str = String.Format("{0} ({1})", ds.Tables[0].Rows[i]["name"].ToString(), ds.Tables[0].Rows[i]["bank_name"].ToString());
                ep.SetText(rw, 2, str);
                /*
                if (str.ToLower().StartsWith("mc") || str.ToLower().StartsWith("master") || str.ToLower().StartsWith("mifare standart"))
                {
                    mc1 += Convert.ToInt32(ds.Tables[0].Rows[i]["cnt_new"]);
                    mc2 += Convert.ToInt32(ds.Tables[0].Rows[i]["cnt_perso"]);
                }
                if (str.ToLower().StartsWith("visa"))
                {
                    v1 += Convert.ToInt32(ds.Tables[0].Rows[i]["cnt_new"]);
                    v2 += Convert.ToInt32(ds.Tables[0].Rows[i]["cnt_perso"]);
                }
                if (str.ToLower().StartsWith("белый") || str.ToLower().IndexOf("карты типа восток") >= 0)
                {
                    ser1 += Convert.ToInt32(ds.Tables[0].Rows[i]["cnt_new"]);
                    ser2 += Convert.ToInt32(ds.Tables[0].Rows[i]["cnt_perso"]);
                }
                if (str.ToLower().StartsWith("nfs") || str.ToLower().StartsWith("nfc"))
                {
                    nfc1 += Convert.ToInt32(ds.Tables[0].Rows[i]["cnt_new"]);
                    nfc2 += Convert.ToInt32(ds.Tables[0].Rows[i]["cnt_perso"]);
                }
                */
                int cc1 = Convert.ToInt32(ds.Tables[0].Rows[i]["cnt_new"]);
                int cc2 = Convert.ToInt32(ds.Tables[0].Rows[i]["cnt_perso"]);
                BaseProductType tp = BranchStore.codeFromTypeAndProdName(1, str);
                switch (tp)
                {
                    case BaseProductType.MasterCard: mc1 += cc1; mc2 += cc2; break;
                    case BaseProductType.VisaCard: v1 += cc1; v2 += cc2; break;
                    case BaseProductType.NFCCard: nfc1 += cc1; nfc2 += cc2; break;
                    case BaseProductType.ServiceCard: ser1 += cc1; ser2 += cc2; break;
                    case BaseProductType.MirCard: mir1 += cc1; mir2 += cc2; break; break;

                }


                cnt1 += cc1; // Convert.ToInt32(ds.Tables[0].Rows[i]["cnt_new"]);
                cnt2 += cc2; // Convert.ToInt32(ds.Tables[0].Rows[i]["cnt_perso"]);
                ep.SetText(rw, 3, cc1.ToString()); // ds.Tables[0].Rows[i]["cnt_new"].ToString());
                ep.SetText(rw, 4, cc2.ToString()); // ds.Tables[0].Rows[i]["cnt_perso"].ToString());
                rw++;
            }
            ep.SetText(rw, 2, "Итого карт");
            ep.SetText(rw, 3, cnt1.ToString());
            ep.SetText(rw, 4, cnt2.ToString());
            rw++;
            ds.Clear();
            Database2.ExecuteQuery("select name, bank_name, cnt_new, cnt_perso from V_Storage where id_type=2 order by id_sort", ref ds, null);
            ep.SetText(rw, 3, "Бланки (шт)");
            ep.SetText(rw, 4, "Персо");
            rw++;
            for (i = 0; i < ds.Tables[0].Rows.Count; i++)
            {
                string str = String.Format("{0} ({1})", ds.Tables[0].Rows[i]["name"].ToString(), ds.Tables[0].Rows[i]["bank_name"].ToString());
                pin1 += Convert.ToInt32(ds.Tables[0].Rows[i]["cnt_new"]);
                pin2 += Convert.ToInt32(ds.Tables[0].Rows[i]["cnt_perso"]);
                ep.SetText(rw, 2, str);
                ep.SetText(rw, 3, ds.Tables[0].Rows[i]["cnt_new"].ToString());
                ep.SetText(rw, 4, ds.Tables[0].Rows[i]["cnt_perso"].ToString());
                rw++;
            }
            ds.Clear();
            ep.SetRangeBorders(3, 2, rw - 1, 4);
            ep.SetText_Name("date_string2", String.Format("на {0:dd.MM.yyyy}", DateTime.Now));            
            #endregion
            //////////////////////////////
            // 3 лист - возвращаемся на 3 лист, проставляем хранилище
            //////////////////////////////
            ep.SetWorkSheet(3);
            ep.SetText(21, 18, mc1.ToString());
            ep.SetText(21, 19, mc2.ToString());
            ep.SetText(22, 18, v1.ToString());
            ep.SetText(22, 19, v2.ToString());
            ep.SetText(23, 18, ser1.ToString());
            ep.SetText(23, 19, ser2.ToString());
            ep.SetText(101, 18, (mc1+v1+ser1+nfc1+mir1).ToString());
            ep.SetText(101, 19, (mc2+v2+ser2+nfc2+mir2).ToString());
            ep.SetText(24, 18, nfc1.ToString());
            ep.SetText(24, 19, nfc2.ToString());
            
            ep.SetText(25, 18, mir1.ToString());
            ep.SetText(25, 19, mir2.ToString());
            
            ep.SetText(26, 18, pin1.ToString());
            ep.SetText(26, 19, pin2.ToString());
//            ep.SetRangeBorders(3, 2, 9, 4);
            ep.SetText(19,18, String.Format("состояние хранилища на конец дня {0:dd.MM.yyyy}", DateTime.Now));

        }

        public class ProductComparer : IComparer
        {
            public int Compare(Object prod1, Object prod2)
            {
                MyProd p1 = (MyProd)prod1;
                MyProd p2 = (MyProd)prod2;
                int code1 = (int)BranchStore.codeFromTypeAndProdName(p1.Type, p1.Name);
                int code2 = (int)BranchStore.codeFromTypeAndProdName(p2.Type, p2.Name);
                if (code1 == code2) return 0;
                if (code1 < code2) return -1;
                return 1;
            }
        }
        private void GivenCardGO(ExcelAp ep)
        {
            DataSet ds = new DataSet();
            int i = 0;
            string res = "";
            SqlCommand comm = new SqlCommand();
            ds.Clear();
            comm.CommandText = "select id_prb, prod_name, bank_name, id_type from V_ProductsBanks_T order by id_sort";
            res = (string)Database2.ExecuteCommand(comm, ref ds, null);
            ArrayList al = new ArrayList();
            foreach (DataRow dr in ds.Tables[0].Rows)
                al.Add(new MyProd(Convert.ToInt32(dr["id_prb"]), dr["prod_name"].ToString(), dr["bank_name"].ToString(), Convert.ToInt32(dr["id_type"])));
            ds.Clear();
            //comm.CommandText = "select id_prb, ident_dep from V_Cards_StorageDocs1 where priz_gen=1 and type=7 and date_doc=@dateS and (ident_dep='0053' or ident_dep='0000')";
            comm.CommandText = "select vcs.id_prb, vcs.ident_dep,c.ispin from V_Cards_StorageDocs1 vcs inner join cards c on c.id=vcs.id_card where vcs.priz_gen=1 and vcs.type=7 and vcs.date_doc=@dateS and (ident_dep='0053' or ident_dep='0000')";
            comm.Parameters.Add("@dateS", SqlDbType.DateTime).Value = DatePickerOne.SelectedDate.Date;
            Database2.ExecuteCommand(comm, ref ds, null);
            if (ds == null || ds.Tables.Count == 0)
                return;
            //-------------------------------------------------------
            al.Sort(new ProductComparer());
            BranchStore branch = new BranchStore(0, "0000", "0000");
            //-------------------------------------------------------

            int pin1 = 0, pin2 = 0; int pin3 = 0, pin4 = 0, pin5 = 0, pin6 = 0;
            foreach (DataRow dr in ds.Tables[0].Rows)
            {
                foreach (MyProd mp in al)
                {
                    if (mp.ID == Convert.ToInt32(dr["id_prb"]))
                    {
                        if (dr["ident_dep"].ToString() == "0000")
                        {
                            mp.cnts[0]++;
                            if (Convert.ToInt32(dr["ispin"]) > 0) pin1++;
                        }
                        if (dr["ident_dep"].ToString() == "0053")
                        {
                            mp.cnts[1]++;
                            if (Convert.ToInt32(dr["ispin"]) > 0) pin2++;
                        }
                        break;
                    }
                }
            }
            ds.Clear();
            //comm.CommandText = "select id_prb, ident_dep from V_Cards_StorageDocs1 where priz_gen=1 and type=9 and date_doc=@dateS and (ident_dep='0053' or ident_dep='0000')";
            comm.CommandText = "select vcs.id_prb, vcs.ident_dep,c.ispin from V_Cards_StorageDocs1 vcs inner join cards c on c.id=vcs.id_card where vcs.priz_gen=1 and vcs.type=9 and vcs.date_doc=@dateS and (ident_dep='0053' or ident_dep='0000')";
            Database2.ExecuteCommand(comm, ref ds, null);
            if (ds == null || ds.Tables.Count == 0)
                return;
            foreach (DataRow dr in ds.Tables[0].Rows)
            {
                foreach (MyProd mp in al)
                {
                    if (mp.ID == Convert.ToInt32(dr["id_prb"]))
                    {
                        if (dr["ident_dep"].ToString() == "0000")
                        {
                            mp.cnts[4]++;
                            if (Convert.ToInt32(dr["ispin"]) > 0) pin5++;
                        }
                        if (dr["ident_dep"].ToString() == "0053")
                        {
                            mp.cnts[5]++;
                            if (Convert.ToInt32(dr["ispin"]) > 0) pin6++;
                        }
                    }
                }
            }
            object obj = null;
            comm.Parameters.Clear();
            comm.CommandText = "select id from Branchs where ident_dep='0000'";
            Database2.ExecuteScalar(comm, ref obj, null);
            int id0 = Convert.ToInt32(obj);
            comm.CommandText = "select id from Branchs where ident_dep='0053'";
            Database2.ExecuteScalar(comm, ref obj, null);
            int id53 = Convert.ToInt32(obj);

            DateTime toDate  = DateTime.Now.Date;
            DateTime curDate = DatePickerOne.SelectedDate.Date;
            ds.Clear();
            comm.Parameters.Clear();
            //String sql = "select id_prb, count(*) as count_prb from Cards  where (id_stat=4 or id_stat=6) and id_branchCurrent=" + id0.ToString() + " group by id_prb";
            //String sql = "select id_prb, count(*) as count_prb from Cards  where id_stat=4 and id_branchCard=" + id0.ToString() + " group by id_prb";
            String sql = "select id_prb, count(*) as count_prb, case when sum(cast(ispin as integer)) is null then 0 else sum(cast(ispin as integer)) end as pin from Cards  where id_stat=4 and id_branchCard=" + id0.ToString() + " group by id_prb";
            Database2.ExecuteQuery(sql, ref ds, null);
            foreach (MyProd mp in al)
            {
                for (int n = 0; n < ds.Tables[0].Rows.Count; n++)
                {
                    if (mp.ID == Convert.ToInt32(ds.Tables[0].Rows[n]["id_prb"]))
                    {
                        mp.cnts[2] = Convert.ToInt32(ds.Tables[0].Rows[n]["count_prb"]);
                        pin3 += Convert.ToInt32(ds.Tables[0].Rows[n]["pin"]);
                        break;
                    }
                }
            }
            ds.Clear();
            //sql = "select id_prb, count(*) as count_prb from Cards  where (id_stat=4 or id_stat=6) and id_branchCurrent=" + id53.ToString() + " group by id_prb";
            //sql = "select id_prb, count(*) as count_prb from Cards  where id_stat=4 and id_branchCard=" + id53.ToString() + " group by id_prb";
            sql = "select id_prb, count(*) as count_prb, case when sum(cast(ispin as integer)) is null then 0 else sum(cast(ispin as integer)) end as pin from Cards  where id_stat=4 and id_branchCard=" + id53.ToString() + " group by id_prb";
            Database2.ExecuteQuery(sql, ref ds, null);
            foreach (MyProd mp in al)
            {
                for (int n = 0; n < ds.Tables[0].Rows.Count; n++)
                {
                    if (mp.ID == Convert.ToInt32(ds.Tables[0].Rows[n]["id_prb"]))
                    {
                        mp.cnts[3] = Convert.ToInt32(ds.Tables[0].Rows[n]["count_prb"]);
                        pin4 += Convert.ToInt32(ds.Tables[0].Rows[n]["pin"]);
                        break;
                    }
                }
            }
            if (curDate < toDate)
            {
               ds.Clear();
               comm.Parameters.Clear();
               //корректировка выдачи
               comm.CommandText = "select dbo.Cards.id_prb,count(*) as count_prb," +
                                  " case when sum(cast(dbo.Cards.ispin as integer)) is null then 0 else sum(cast(dbo.Cards.ispin as integer)) end as pin" +
                                  " from dbo.Branchs RIGHT OUTER JOIN" +
                                  " dbo.StorageDocs ON dbo.Branchs.id = dbo.StorageDocs.id_branch LEFT OUTER JOIN" +
                                  " dbo.Cards RIGHT OUTER JOIN" +
                                  " dbo.Cards_StorageDocs ON dbo.Cards.id = dbo.Cards_StorageDocs.id_card ON dbo.StorageDocs.id = dbo.Cards_StorageDocs.id_doc" +
                                  " where date_doc>@dateS and date_doc<=@dateP and" +
                                  " (id_branch=@id_branch and (type=7 or type=9 or type=11 or type=13) or (id_act=@id_branch and (type=19 or type=5)))" +
                                  " and priz_gen=1 group by dbo.Cards.id_prb";
                comm.Parameters.Add("@dateS", SqlDbType.DateTime).Value = curDate;
                comm.Parameters.Add("@dateP", SqlDbType.DateTime).Value = toDate;
                comm.Parameters.Add("@id_branch", SqlDbType.Int).Value = id0;
                Database2.ExecuteCommand(comm, ref ds, null);
                foreach (MyProd mp in al)
                {
                    int countMinus = 0;
                    int pinMinus = 0;
                    for (int n = 0; n < ds.Tables[0].Rows.Count; n++)
                    {
                        if (mp.ID == Convert.ToInt32(ds.Tables[0].Rows[n]["id_prb"]))
                        {
                            countMinus = Convert.ToInt32(ds.Tables[0].Rows[n]["count_prb"]);
                            pinMinus = Convert.ToInt32(ds.Tables[0].Rows[n]["pin"]);
                            break;
                        }
                    }
                    mp.cnts[2] += countMinus;
                    pin3 += pinMinus;
                }
                ds.Clear();
                comm.Parameters["@id_branch"].Value = id53;
                Database2.ExecuteCommand(comm, ref ds, null);
                foreach (MyProd mp in al)
                {
                    int countMinus = 0;
                    int pinMinus = 0;
                    for (int n = 0; n < ds.Tables[0].Rows.Count; n++)
                    {
                        if (mp.ID == Convert.ToInt32(ds.Tables[0].Rows[n]["id_prb"]))
                        {
                            countMinus = Convert.ToInt32(ds.Tables[0].Rows[n]["count_prb"]);
                            pinMinus = Convert.ToInt32(ds.Tables[0].Rows[n]["pin"]);
                            break;
                        }
                    }
                    mp.cnts[3] += countMinus;
                    pin4 += pinMinus;
                }
                //корректировка получения
                ds.Clear();
                comm.Parameters.Clear();
                //comm.CommandText = "select id_prb,count(*) as count_prb from V_Cards_StorageDocs1 where date_doc>@dateS and date_doc<=@dateP and id_branch=@id_branch and (type=6 or type=10) and priz_gen=1 group by id_prb";
                comm.CommandText = "select vcs.id_prb,count(*) as count_prb,case when sum(cast(c.ispin as integer)) is null then 0 else sum(cast(c.ispin as integer)) end as pin from V_Cards_StorageDocs1 vcs inner join cards c on c.id=vcs.id_card where vcs.date_doc>@dateS and vcs.date_doc<=@dateP and vcs.id_branch=@id_branch and (vcs.type=6 or vcs.type=10) and vcs.priz_gen=1 group by vcs.id_prb";
                comm.Parameters.Add("@dateS", SqlDbType.DateTime).Value = curDate;
                comm.Parameters.Add("@dateP", SqlDbType.DateTime).Value = toDate;
                comm.Parameters.Add("@id_branch", SqlDbType.Int).Value = id0;
                Database2.ExecuteCommand(comm, ref ds, null);
                foreach (MyProd mp in al)
                {

                    int countPlus = 0;
                    int pinPlus = 0;
                    for (int n = 0;n < ds.Tables[0].Rows.Count; n++)
                    {
                        if (mp.ID == Convert.ToInt32(ds.Tables[0].Rows[n]["id_prb"]))
                        {
                            countPlus = Convert.ToInt32(ds.Tables[0].Rows[n]["count_prb"]);
                            pinPlus = Convert.ToInt32(ds.Tables[0].Rows[n]["pin"]);
                            break;
                        }
                    }
                    mp.cnts[2] -= countPlus;
                    pin3 -= pinPlus;
                }
                ds.Clear();
                comm.Parameters["@id_branch"].Value = id53;
                Database2.ExecuteCommand(comm, ref ds, null);
                foreach (MyProd mp in al)
                {

                    int countPlus = 0;
                    int pinPlus = 0;
                    for (int n = 0; n < ds.Tables[0].Rows.Count; n++)
                    {
                        if (mp.ID == Convert.ToInt32(ds.Tables[0].Rows[n]["id_prb"]))
                        {
                            countPlus = Convert.ToInt32(ds.Tables[0].Rows[n]["count_prb"]);
                            pinPlus = Convert.ToInt32(ds.Tables[0].Rows[n]["pin"]);
                            break;
                        }
                    }
                    mp.cnts[3] -= countPlus;
                    pin4 -= pinPlus;
                }
            }

            i = 0;
            int[] mc = new int[MyProd.cntMax];
            int[] vis = new int[MyProd.cntMax];
            //            int[] pin = new int[MyProd.cntMax];
            int[] ser = new int[MyProd.cntMax];
            int[] nfc = new int[MyProd.cntMax];
            int[] mir = new int[MyProd.cntMax];

            for (int t = 0; t < MyProd.cntMax; t++)
            {
                mc[t] = 0;
                vis[t] = 0;
                //                pin[t] = 0;
                ser[t] = 0;
                nfc[t] = 0;
                mir[t] = 0;
            }
            int code = -1;
            foreach (MyProd mp in al)
            {
                BaseProductType curCode = BranchStore.codeFromTypeAndProdName(mp.Type, mp.Name);
                if (curCode == BaseProductType.MasterCard)
                //if (mp.Name.ToLower().StartsWith("mc") || mp.Name.ToLower().StartsWith("master"))
                {
                    mc[0] += mp.cnts[0];
                    mc[1] += mp.cnts[1];
                    mc[2] += mp.cnts[2];
                    mc[3] += mp.cnts[3];
                    mc[4] += mp.cnts[4];
                    mc[5] += mp.cnts[5];

                }
                else
                if (curCode == BaseProductType.VisaCard)
                //if (mp.Name.ToLower().StartsWith("visa"))
                {
                    vis[0] += mp.cnts[0];
                    vis[1] += mp.cnts[1];
                    vis[2] += mp.cnts[2];
                    vis[3] += mp.cnts[3];
                    vis[4] += mp.cnts[4];
                    vis[5] += mp.cnts[5];
                }
                else
                if (curCode == BaseProductType.ServiceCard)
                //if (mp.Name.ToLower().StartsWith("белый"))
                {
                    ser[0] += mp.cnts[0];
                    ser[1] += mp.cnts[1];
                    ser[2] += mp.cnts[2];
                    ser[3] += mp.cnts[3];
                    ser[4] += mp.cnts[4];
                    ser[5] += mp.cnts[5];
                }
                else
                if (curCode == BaseProductType.NFCCard)
                //if (mp.Name.StartsWith("nfs") == true || mp.Name.StartsWith("nfc") == true)
                {
                    nfc[0] += mp.cnts[0];
                    nfc[1] += mp.cnts[1];
                    nfc[2] += mp.cnts[2];
                    nfc[3] += mp.cnts[3];
                    nfc[4] += mp.cnts[4];
                    nfc[5] += mp.cnts[5];
                }
                else
                if (curCode == BaseProductType.MirCard)
                {
                    mir[0] += mp.cnts[0];
                    mir[1] += mp.cnts[1];
                    mir[2] += mp.cnts[2];
                    mir[3] += mp.cnts[3];
                    mir[4] += mp.cnts[4];
                    mir[5] += mp.cnts[5];
                }
                if (mp.Type == 1)
                {
                    branch.addCount(mp.Type, mp.Name, 0, mp.cnts[0]);
                    branch.addCount(mp.Type, mp.Name, 1, mp.cnts[1]);
                    branch.addCount(mp.Type, mp.Name, 2, mp.cnts[2]);
                    branch.addCount(mp.Type, mp.Name, 3, mp.cnts[3]);
                    branch.addCount(mp.Type, mp.Name, 4, mp.cnts[4]);
                    branch.addCount(mp.Type, mp.Name, 5, mp.cnts[5]);
                    branch.addCount(mp.Type, mp.Name, 6, mp.cnts[5]);
                }
                if (code < 0) code = (int)curCode;
                else
                {
                    if (code != (int)curCode && code > 0)
                    {
                        int[] count0145 = null;
                        if (code == (int)BaseProductType.MasterCard) count0145 = branch.countMasterCard;
                        if (code == (int)BaseProductType.VisaCard) count0145 = branch.countVisaCard;
                        if (code == (int)BaseProductType.NFCCard) count0145 = branch.countNFCCard;
                        if (code == (int)BaseProductType.ServiceCard) count0145 = branch.countServiceCard;
                        if (code == (int)BaseProductType.MirCard) count0145 = branch.countMirCard;
                        if (count0145 != null && count0145[0] + count0145[1] + count0145[4] + count0145[5] > 0)
                        {
                            ep.SetText(26 + i, 1, branch.getNameForCard(code));
                            ep.SetRangeBold(26 + i, 1, 26 + i, 1);
                            if (count0145[0] != 0)
                                ep.SetText(26 + i, 6, count0145[0].ToString());
                            if (count0145[1] != 0)
                                ep.SetText(26 + i, 8, count0145[1].ToString());
                            if (count0145[4] != 0)
                                ep.SetText(26 + i, 7, count0145[4].ToString());
                            if (count0145[5] != 0)
                                ep.SetText(26 + i, 9, count0145[5].ToString());
                            i++;
                        }
                        code = (int)curCode;
                    }
                }
                if (mp.cnts[0] + mp.cnts[1] + mp.cnts[4] + mp.cnts[5] != 0)
                {
                    ep.SetText(26 + i, 1, mp.Name);
                    if (mp.cnts[0] != 0)
                        ep.SetText(26 + i, 6, mp.cnts[0].ToString());
                    if (mp.cnts[1] != 0)
                        ep.SetText(26 + i, 8, mp.cnts[1].ToString());
                    if (mp.cnts[4] != 0)
                        ep.SetText(26 + i, 7, mp.cnts[4].ToString());
                    if (mp.cnts[5] != 0)
                        ep.SetText(26 + i, 9, mp.cnts[5].ToString());
                    i++;
                }
            }
            ep.SetText(26 + i, 1, "Пин-конверты");
            if (pin1 > 0)
                ep.SetText(26 + i, 6, pin1.ToString());
            if (pin2 > 0)
                ep.SetText(26 + i, 8, pin2.ToString());
            if (pin5 > 0)
                ep.SetText(26 + i, 7, pin5.ToString());
            if (pin6 > 0)
                ep.SetText(26 + i, 9, pin6.ToString());
            ep.ShowRows(26, 26 + i);
            if (pin1 > 0)
                ep.SetText(90, 6, pin1.ToString());
            if (pin2 > 0)
                ep.SetText(90, 8, pin2.ToString());
            if (pin5 > 0)
                ep.SetText(90, 7, pin5.ToString());
            if (pin6 > 0)
                ep.SetText(90, 9, pin6.ToString());


            ep.SetText_Name("date_string", String.Format("Дата {0:dd.MM.yyyy}", DatePickerOne.SelectedDate));
            ep.SetText_Name("date_string1", String.Format("от {0: dd MM yyyy}", DatePickerOne.SelectedDate));
            ///////////////////////////
            // лист 3 распоряжение сводный
            ///////////////////////////
            ep.SetWorkSheet(3);
            ep.SetText(25, 2, "MasterCard");
            if (mc[0] != 0)
                ep.SetText(25, 7, mc[0].ToString());
            //            if (mc[1] != 0)
            //                ep.SetText(25, 8, mc[1].ToString());
            if (mc[4] != 0)
                ep.SetText(25, 8, mc[4].ToString());
            //            if (mc[5] != 0)
            //                ep.SetText(25, 9, mc[5].ToString());
            ep.SetText(26, 2, "Visa");
            if (vis[0] != 0)
                ep.SetText(26, 7, vis[0].ToString());
            //            if (vis[1] != 0)
            //                ep.SetText(26, 8, vis[1].ToString());
            if (vis[4] != 0)
                ep.SetText(26, 8, vis[4].ToString());
            //            if (vis[5] != 0)
            //                ep.SetText(26, 9, vis[5].ToString());

            ep.SetText(27, 2, "NFC");
            if (nfc[0] != 0)
                ep.SetText(27, 7, nfc[0].ToString());
            //            if (nfc[1] != 0)
            //                ep.SetText(27, 8, nfc[1].ToString());
            if (nfc[4] != 0)
                ep.SetText(27, 8, nfc[4].ToString());
            //          


            ep.SetText(28, 2, "Сервисные карты");
            if (ser[0] != 0)
                ep.SetText(28, 7, ser[0].ToString());
            //            if (ser[1] != 0)
            //                ep.SetText(28, 8, ser[1].ToString());
            if (ser[4] != 0)
                ep.SetText(28, 8, ser[4].ToString());
            //            if (ser[5] != 0)
            //                ep.SetText(28, 9, ser[5].ToString());

            ep.SetText(29, 2, "Карты МИР");
            if (mir[0] != 0)
                ep.SetText(29, 7, mir[0].ToString());
            //            if (mir[1] != 0)
            //                ep.SetText(29, 8, mir[1].ToString());
            if (mir[4] != 0)
                ep.SetText(29, 8, mir[4].ToString());
            //            if (mir[5] != 0)
            //                ep.SetText(29, 9, mir[5].ToString());

            ep.SetText(30, 2, "Пин-конверты");
            if (pin1 > 0)
                ep.SetText(30, 7, pin1.ToString());
            //            if (pin2 > 0)
            //                ep.SetText(28, 8, pin2.ToString());
            if (pin5 > 0)
                ep.SetText(30, 8, pin5.ToString());
            //            if (pin6 > 0)
            //                ep.SetText(28, 9, pin6.ToString());
            ep.ShowRows(25, 30);

            if (mc[0] + vis[0] + ser[0] + nfc[0] + mir[0]> 0)
                ep.SetText(89, 7, (mc[0] + vis[0] + ser[0] + nfc[0] + mir[0]).ToString());
            //            if (mc[1] + vis[1] + ser[1] > 0)
            //                ep.SetText(89, 8, (mc[1]+vis[1]+ser[1]).ToString());
            if (mc[4] + vis[4] + ser[4] + nfc[4] + mir[4]> 0)
                ep.SetText(89, 8, (mc[4] + vis[4] + ser[4] + nfc[4] + mir[4]).ToString());
            //            if (mc[5] + vis[5] + ser[5] + nfc[5] > 0)
            //                ep.SetText(89, 9, (mc[5]+vis[5]+ser[5]+nfc[5]).ToString());
            ep.SetText_Name("date_string2", String.Format("Дата {0:dd.MM.yyyy}", DatePickerOne.SelectedDate));
            ep.SetText_Name("date_string3", String.Format("от {0: dd MM yyyy}", DatePickerOne.SelectedDate));
            ///////////////////////////
            // лист 2 остатки
            ///////////////////////////
            ep.SetWorkSheet(2);
            i = 0;
            ep.SetText(5, 5, "ГО РКЦ");
            ep.SetText(5, 6, "ГО БТА");
            int card0 = 0, card53 = 0;
            foreach (MyProd mp in al)
            {
                if (mp.cnts[2] != 0 || mp.cnts[3] != 0)
                {
                    ep.SetText(6 + i, 1, mp.Name);
                    if (mp.cnts[2] != 0)
                        ep.SetText(6 + i, 5, mp.cnts[2].ToString());
                    card0 += mp.cnts[2];
                    if (mp.cnts[3] != 0)
                        ep.SetText(6 + i, 6, mp.cnts[3].ToString());
                    card53 += mp.cnts[3];
                    i++;
                }
            }
            ep.SetText(6 + i, 1, "Пин-конверты");
            if (pin3 > 0)
                ep.SetText(6 + i, 5, pin3.ToString());
            if (pin4 > 0)
                ep.SetText(6 + i, 6, pin4.ToString());
            i++;
            ep.SetText(6 + i, 1, "Итого карт");
            /*if (pin3 > 0)
                ep.SetText(6 + i, 5, pin3.ToString());
            if (pin4 > 0)
                ep.SetText(6 + i, 6, pin4.ToString());*/
            if (card0 > 0)
                ep.SetText(6 + i, 5, card0.ToString());
            if (card53 > 0)
                ep.SetText(6 + i, 6, card53.ToString());
            i++;
            ep.ShowRows(5, i + 6 - 1);
            //////////////////////////////////
            // лист 3 остатки сводные - возвращаемся на лист 3
            //////////////////////////////////
            ep.SetWorkSheet(3);
            i = 0;
            //            ep.SetText(5, 5, "ГО РКЦ");
            //            ep.SetText(5, 6, "ГО БТА");
            //            i = 0;
            //            ep.SetText(6 + i, 1, "MasterCard");
            if (mc[2] > 0)
                ep.SetText(25, 9, mc[2].ToString());
            //            if (mc[3] > 0)
            //                ep.SetText(6 + i, 6, mc[3].ToString());
            i++;
            //            ep.SetText(6 + i, 1, "Visa");
            if (vis[2] > 0)
                ep.SetText(26, 9, vis[2].ToString());
            //            if (vis[3] > 0)
            //                ep.SetText(6 + i, 6, vis[3].ToString());
            i++;
            //            ep.SetText(6 + i, 1, "Сервисные карты");

            if (nfc[2] > 0)
                ep.SetText(27, 9, nfc[2].ToString());

            if (ser[2] > 0)
                ep.SetText(28, 9, ser[2].ToString());
            //            if (ser[3] > 0)
            //                ep.SetText(6 + i, 6, ser[3].ToString());
            i++;
            //            ep.SetText(6 + i, 1, "Пин-конверты");

            if (mir[2] > 0)
                ep.SetText(29, 9, mir[2].ToString());
            //            if (mir[3] > 0)
            //                ep.SetText(6 + i, 6, mir[3].ToString());
            i++;
            //            ep.SetText(6 + i, 1, "Пин-конверты");

            if (pin3 > 0)
                ep.SetText(30, 9, pin3.ToString());
            //            if (pin4 > 0)
            //                ep.SetText(6 + i, 6, pin4.ToString());
            //          i++;
            //            ep.SetText(6 + i, 1, "Итого карт");
            if (mc[2] + vis[2] + ser[2] + nfc[2] + mir[2] > 0)
                ep.SetText(89, 9, (mc[2] + vis[2] + ser[2] + nfc[2] + mir[2]).ToString());
            //            if (mc[3] + vis[3] + ser[3] > 0)
            //                ep.SetText(6 + i, 6, (mc[3] + vis[3] + ser[3]).ToString());
            i++;
            //ep.ShowRows(5, i + 6 - 1);
        }
        
        private void GivenCardBranch(ExcelAp ep, string identDep)
        {
            DataSet ds = new DataSet();
            int i = 0;
            string res = "";
            SqlCommand comm = new SqlCommand();
            ds.Clear();
            comm.CommandText = "select id_prb, prod_name, bank_name, id_type from V_ProductsBanks_T order by id_sort";
            res = (string)Database2.ExecuteCommand(comm, ref ds, null);
            ArrayList al = new ArrayList();
            foreach (DataRow dr in ds.Tables[0].Rows)
                al.Add(new MyProd(Convert.ToInt32(dr["id_prb"]), dr["prod_name"].ToString(), dr["bank_name"].ToString(), Convert.ToInt32(dr["id_type"])));
            ds.Clear();
            //comm.CommandText = String.Format("select id_prb, ident_dep from V_Cards_StorageDocs1 where priz_gen=1 and type=7 and date_doc=@dateS and ident_dep='{0}'", identDep);
            comm.CommandText = String.Format("select vcs.id_prb, vcs.ident_dep,c.ispin from V_Cards_StorageDocs1 vcs inner join cards c on c.id=vcs.id_card where vcs.priz_gen=1 and vcs.type=7 and vcs.date_doc=@dateS and vcs.ident_dep='{0}'", identDep);
            comm.Parameters.Add("@dateS", SqlDbType.DateTime).Value = DatePickerOne.SelectedDate.Date;
            Database2.ExecuteCommand(comm, ref ds, null);
            if (ds == null || ds.Tables.Count == 0)
                return;
            int pin1 = 0, pin2 = 0; int pin3 = 0, pin4 = 0, pin5 = 0, pin6 = 0;
            foreach (DataRow dr in ds.Tables[0].Rows)
            {
                foreach (MyProd mp in al)
                {
                    if (mp.ID == Convert.ToInt32(dr["id_prb"]))
                    {
                        mp.cnts[0]++; 
                        if(Convert.ToInt32(dr["ispin"])>0) pin1++;
                        break;
                    }
                }
            }
            ds.Clear();
            //comm.CommandText = String.Format("select id_prb, ident_dep from V_Cards_StorageDocs1 where priz_gen=1 and type=9 and date_doc=@dateS and ident_dep='{0}'", identDep);
            comm.CommandText = String.Format("select vcs.id_prb, vcs.ident_dep,c.ispin from V_Cards_StorageDocs1 vcs inner join cards c on c.id=vcs.id_card where vcs.priz_gen=1 and vcs.type=9 and vcs.date_doc=@dateS and vcs.ident_dep='{0}'", identDep);
            Database2.ExecuteCommand(comm, ref ds, null);
            if (ds == null || ds.Tables.Count == 0)
                return;
            foreach (DataRow dr in ds.Tables[0].Rows)
            {
                foreach (MyProd mp in al)
                {
                    if (mp.ID == Convert.ToInt32(dr["id_prb"]))
                    {
                        mp.cnts[4]++;
                        if (Convert.ToInt32(dr["ispin"])>0) pin5++;
                        break;
                    }
                }
            }
            object obj = null;
            comm.Parameters.Clear();
            comm.CommandText = String.Format("select id from Branchs where ident_dep='{0}'", identDep);
            Database2.ExecuteScalar(comm, ref obj, null);
            int id0 = Convert.ToInt32(obj);
            comm.CommandText = String.Format("select department from Branchs where ident_dep='{0}'", identDep);
            Database2.ExecuteScalar(comm, ref obj, null);
            string depName = Convert.ToString(obj);

            DateTime toDate = DateTime.Now.Date;
            DateTime curDate = DatePickerOne.SelectedDate.Date;
            ds.Clear();
            comm.Parameters.Clear();
            //String sql = "select id_prb, count(*) as count_prb from Cards  where (id_stat=4 or id_stat=6) and id_branchCurrent=" + id0.ToString() + " group by id_prb";
            String sql = "select id_prb, count(*) as count_prb, case when sum(cast(ispin as integer)) is null then 0 else sum(cast(ispin as integer)) end as pin from Cards  where id_stat=4 and id_branchCard=" + id0.ToString() + " group by id_prb";
            Database2.ExecuteQuery(sql, ref ds, null);
            foreach (MyProd mp in al)
            {
                for (int n = 0; n < ds.Tables[0].Rows.Count; n++)
                {
                    if (mp.ID == Convert.ToInt32(ds.Tables[0].Rows[n]["id_prb"]))
                    {
                        mp.cnts[2] = Convert.ToInt32(ds.Tables[0].Rows[n]["count_prb"]);
                        pin3 += Convert.ToInt32(ds.Tables[0].Rows[n]["pin"]);
                        break;
                    }
                }
            }
            if (curDate < toDate)
            {
                ds.Clear();
                comm.Parameters.Clear();
                //корректировка выдачи
                comm.CommandText = "select dbo.Cards.id_prb,count(*) as count_prb," +
                                   " case when sum(cast(dbo.Cards.ispin as integer)) is null then 0 else sum(cast(dbo.Cards.ispin as integer)) end as pin" +
                                   " from dbo.Branchs RIGHT OUTER JOIN" +
                                   " dbo.StorageDocs ON dbo.Branchs.id = dbo.StorageDocs.id_branch LEFT OUTER JOIN" +
                                   " dbo.Cards RIGHT OUTER JOIN" +
                                   " dbo.Cards_StorageDocs ON dbo.Cards.id = dbo.Cards_StorageDocs.id_card ON dbo.StorageDocs.id = dbo.Cards_StorageDocs.id_doc" +
                                   " where date_doc>@dateS and date_doc<=@dateP and" +
                                   " (id_branch=@id_branch and (type=7 or type=9 or type=11 or type=13) or (id_act=@id_branch and (type=19 or type=5)))" +
                                   " and priz_gen=1 group by dbo.Cards.id_prb";
                comm.Parameters.Add("@dateS", SqlDbType.DateTime).Value = curDate;
                comm.Parameters.Add("@dateP", SqlDbType.DateTime).Value = toDate;
                comm.Parameters.Add("@id_branch", SqlDbType.Int).Value = id0;
                Database2.ExecuteCommand(comm, ref ds, null);
                foreach (MyProd mp in al)
                {
                    int countMinus = 0;
                    int pinMinus = 0;
                    for (int n = 0; n < ds.Tables[0].Rows.Count; n++)
                    {
                        if (mp.ID == Convert.ToInt32(ds.Tables[0].Rows[n]["id_prb"]))
                        {
                            countMinus = Convert.ToInt32(ds.Tables[0].Rows[n]["count_prb"]);
                            pinMinus = Convert.ToInt32(ds.Tables[0].Rows[n]["pin"]);
                            break;
                        }
                    }
                    mp.cnts[2] += countMinus;
                    pin3 += pinMinus;
                }
                //корректировка получения
                ds.Clear();
                comm.Parameters.Clear();
                //comm.CommandText = "select id_prb,count(*) as count_prb from V_Cards_StorageDocs1 where date_doc>@dateS and date_doc<=@dateP and id_branch=@id_branch and (type=6 or type=10) and priz_gen=1 group by id_prb";
                comm.CommandText = "select vcs.id_prb,count(*) as count_prb,case when sum(cast(c.ispin as integer)) is null then 0 else sum(cast(c.ispin as integer)) end as pin from V_Cards_StorageDocs1 vcs inner join cards c on c.id=vcs.id_card where vcs.date_doc>@dateS and vcs.date_doc<=@dateP and vcs.id_branch=@id_branch and (vcs.type=6 or vcs.type=10) and vcs.priz_gen=1 group by vcs.id_prb";
                comm.Parameters.Add("@dateS", SqlDbType.DateTime).Value = curDate;
                comm.Parameters.Add("@dateP", SqlDbType.DateTime).Value = toDate;
                comm.Parameters.Add("@id_branch", SqlDbType.Int).Value = id0;
                Database2.ExecuteCommand(comm, ref ds, null);
                foreach (MyProd mp in al)
                {

                    int countPlus = 0;
                    int pinPlus = 0;
                    for (int n = 0; n < ds.Tables[0].Rows.Count; n++)
                    {
                        if (mp.ID == Convert.ToInt32(ds.Tables[0].Rows[n]["id_prb"]))
                        {
                            countPlus = Convert.ToInt32(ds.Tables[0].Rows[n]["count_prb"]);
                            pinPlus=Convert.ToInt32(ds.Tables[0].Rows[n]["pin"]);
                            break;
                        }
                    }
                    mp.cnts[2] -= countPlus;
                    pin3 -= pinPlus;
                }
            }

            i = 0;
            int[] mc = new int[MyProd.cntMax];
            int[] vis = new int[MyProd.cntMax];
            //            int[] pin = new int[MyProd.cntMax];
            int[] ser = new int[MyProd.cntMax];
            int[] nfc = new int[MyProd.cntMax];
            int[] mir = new int[MyProd.cntMax];

            for (int t = 0; t < MyProd.cntMax; t++)
            {
                mc[t] = 0;
                vis[t] = 0;
                //                pin[t] = 0;
                ser[t] = 0;
                nfc[t] = 0;
                mir[t] = 0;
            }
            foreach (MyProd mp in al)
            {
                BaseProductType tp = BranchStore.codeFromTypeAndProdName(mp.Type, mp.Name);
                //if (mp.Name.ToLower().StartsWith("mc") || mp.Name.ToLower().StartsWith("master"))
                if(tp==BaseProductType.MasterCard)
                {
                    mc[0] += mp.cnts[0];
                    mc[1] += mp.cnts[1];
                    mc[2] += mp.cnts[2];
                    mc[3] += mp.cnts[3];
                    mc[4] += mp.cnts[4];
                    mc[5] += mp.cnts[5];
                }
                else
                //if (mp.Name.ToLower().StartsWith("visa"))
                if (tp == BaseProductType.VisaCard)
                {
                    vis[0] += mp.cnts[0];
                    vis[1] += mp.cnts[1];
                    vis[2] += mp.cnts[2];
                    vis[3] += mp.cnts[3];
                    vis[4] += mp.cnts[4];
                    vis[5] += mp.cnts[5];
                }
                else
                //if (mp.Name.ToLower().StartsWith("nfc") || mp.Name.ToLower().StartsWith("nfs"))
                if (tp == BaseProductType.NFCCard)
                {
                    nfc[0] += mp.cnts[0];
                    nfc[1] += mp.cnts[1];
                    nfc[2] += mp.cnts[2];
                    nfc[3] += mp.cnts[3];
                    nfc[4] += mp.cnts[4];
                    nfc[5] += mp.cnts[5];
                }
                else
                //if (mp.Name.ToLower().StartsWith("белый"))
                if (tp == BaseProductType.ServiceCard)
                {
                    ser[0] += mp.cnts[0];
                    ser[1] += mp.cnts[1];
                    ser[2] += mp.cnts[2];
                    ser[3] += mp.cnts[3];
                    ser[4] += mp.cnts[4];
                    ser[5] += mp.cnts[5];
                }
                else
                if (tp == BaseProductType.MirCard)
                {
                    mir[0] += mp.cnts[0];
                    mir[1] += mp.cnts[1];
                    mir[2] += mp.cnts[2];
                    mir[3] += mp.cnts[3];
                    mir[4] += mp.cnts[4];
                    mir[5] += mp.cnts[5];
                }
                else continue;
                if (mp.cnts[0] + mp.cnts[1] + mp.cnts[4] + mp.cnts[5] != 0)
                {
                    ep.SetText(26 + i, 1, mp.Name);
                    if (mp.cnts[0] != 0)
                        ep.SetText(26 + i, 6, mp.cnts[0].ToString());
                    //                    if (mp.cnts[1] != 0)
                    //                        ep.SetText(26 + i, 8, mp.cnts[1].ToString());
                    if (mp.cnts[4] != 0)
                        ep.SetText(26 + i, 7, mp.cnts[4].ToString());
                    //                    if (mp.cnts[5] != 0)
                    //                        ep.SetText(26 + i, 9, mp.cnts[5].ToString());
                    i++;
                }
            }
            ep.SetText(26 + i, 1, "Пин-конверты");
            if (pin1 > 0)
                ep.SetText(26 + i, 6, pin1.ToString());
            //            if (pin2 > 0)
            //                ep.SetText(26 + i, 8, pin2.ToString());
            if (pin5 > 0)
                ep.SetText(26 + i, 7, pin5.ToString());
            //            if (pin6 > 0)
            //                ep.SetText(26 + i, 9, pin6.ToString());
            ep.ShowRows(26, 26 + i);
            if (pin1 > 0)
                ep.SetText(90, 6, pin1.ToString());
            //            if (pin2 > 0)
            //                ep.SetText(90, 8, pin2.ToString());
            if (pin5 > 0)
                ep.SetText(90, 7, pin5.ToString());
            //            if (pin6 > 0)
            //                ep.SetText(90, 9, pin6.ToString());

            ep.SetText(15, 1, depName);
            ep.SetText_Name("date_string", String.Format("Дата {0:dd.MM.yyyy}", DatePickerOne.SelectedDate));
            ep.SetText_Name("date_string1", String.Format("от {0: dd MM yyyy}", DatePickerOne.SelectedDate));
            ///////////////////////////
            // лист 3 распоряжение сводный
            ///////////////////////////
            ep.SetWorkSheet(3);
            ep.SetText(25, 2, "MasterCard");
            if (mc[0] != 0)
                ep.SetText(25, 7, mc[0].ToString());
            //            if (mc[1] != 0)
            //                ep.SetText(25, 8, mc[1].ToString());
            if (mc[4] != 0)
                ep.SetText(25, 8, mc[4].ToString());
            //            if (mc[5] != 0)
            //                ep.SetText(25, 9, mc[5].ToString());
            ep.SetText(26, 2, "Visa");
            if (vis[0] != 0)
                ep.SetText(26, 7, vis[0].ToString());
            //            if (vis[1] != 0)
            //                ep.SetText(26, 8, vis[1].ToString());
            if (vis[4] != 0)
                ep.SetText(26, 8, vis[4].ToString());
            //            if (vis[5] != 0)
            //                ep.SetText(26, 9, vis[5].ToString());

            ep.SetText(27, 2, "NFC");
            if (nfc[0] != 0)
                ep.SetText(27, 7, nfc[0].ToString());
            //            if (ser[1] != 0)
            //                ep.SetText(27, 8, ser[1].ToString());
            if (nfc[4] != 0)
                ep.SetText(27, 8, nfc[4].ToString());
            //            if (nfcr[5] != 0)
            //                ep.SetText(27, 9, nfc[5].ToString());
            
            ep.SetText(28, 2, "Сервисные карты");
            if (ser[0] != 0)
                ep.SetText(28, 7, ser[0].ToString());
            //            if (ser[1] != 0)
            //                ep.SetText(28, 8, ser[1].ToString());
            if (ser[4] != 0)
                ep.SetText(28, 8, ser[4].ToString());
            //            if (ser[5] != 0)
            //                ep.SetText(28, 9, ser[5].ToString());

            ep.SetText(29, 2, "Карты МИР");
            if (mir[0] != 0)
                ep.SetText(29, 7, mir[0].ToString());
            //            if (mir[1] != 0)
            //                ep.SetText(29, 8, mir[1].ToString());
            if (mir[4] != 0)
                ep.SetText(29, 8, mir[4].ToString());
            //            if (mir[5] != 0)
            //                ep.SetText(29, 9, mir[5].ToString());
            
            ep.SetText(30, 2, "Пин-конверты");
            if (pin1 > 0)
                ep.SetText(30, 7, pin1.ToString());
            //            if (pin2 > 0)
            //                ep.SetText(29, 8, pin2.ToString());
            if (pin5 > 0)
                ep.SetText(30, 8, pin5.ToString());
            //            if (pin6 > 0)
            //                ep.SetText(29, 9, pin6.ToString());
            ep.ShowRows(25, 30);
            if (mc[0] + vis[0] + ser[0] + nfc[0] + mir[0] > 0)
                ep.SetText(89, 7, (mc[0] + vis[0] + ser[0] + nfc[0] + mir[0]).ToString());
            //            if (mc[1] + vis[1] + ser[1] > 0)
            //                ep.SetText(89, 8, (mc[1]+vis[1]+ser[1]).ToString());
            if (mc[4] + vis[4] + ser[4] + nfc[4] + mir[4] > 0)
                ep.SetText(89, 8, (mc[4] + vis[4] + ser[4] + nfc[4] + mir[4]).ToString());
            //            if (mc[5] + vis[5] + ser[5] > 0)
            //                ep.SetText(89, 9, (mc[5]+vis[5]+ser[5]).ToString());
            ep.SetText(15, 1, depName);
            ep.SetText_Name("date_string2", String.Format("Дата {0:dd.MM.yyyy}", DatePickerOne.SelectedDate));
            ep.SetText_Name("date_string3", String.Format("от {0: dd MM yyyy}", DatePickerOne.SelectedDate));
            ///////////////////////////
            // лист 2 остатки
            ///////////////////////////
            ep.SetWorkSheet(2);
            i = 0;
            ep.SetText(5, 1, "Указать имя филиала");
            //            ep.SetText(5, 6, "ГО БТА");
            int cards = 0;
            foreach (MyProd mp in al)
            {
                if (mp.cnts[2] != 0 || mp.cnts[3] != 0)
                {
                    ep.SetText(6 + i, 1, mp.Name);
                    if (mp.cnts[2] != 0)
                        ep.SetText(6 + i, 5, mp.cnts[2].ToString());
                    cards += mp.cnts[2];
                    //                    if (mp.cnts[3] != 0)
                    //                        ep.SetText(6 + i, 6, mp.cnts[3].ToString());
                    i++;
                }
            }
            ep.SetText(6 + i, 1, "Пин-конверты");
            if (pin3 > 0)
                ep.SetText(6 + i, 5, pin3.ToString());
            //            if (pin4 > 0)
            //                ep.SetText(6 + i, 6, pin4.ToString());
            i++;
            ep.SetText(6 + i, 1, "Итого карт");
            //if (pin3 > 0)
            //    ep.SetText(6 + i, 5, pin3.ToString());
            if (cards > 0)
                ep.SetText(6 + i, 5, cards.ToString());

            //            if (pin4 > 0)
            //                ep.SetText(6 + i, 6, pin4.ToString());
            i++;
            ep.ShowRows(5, i + 6 - 1);
            ep.SetText(5, 1, depName);
            //////////////////////////////////
            // лист 3 остатки сводные - возвращаемся на лист 3
            //////////////////////////////////
            ep.SetWorkSheet(3);
            i = 0;
            //            ep.SetText(5, 5, "ГО РКЦ");
            //            ep.SetText(5, 6, "ГО БТА");
            //            i = 0;
            //            ep.SetText(6 + i, 1, "MasterCard");
            if (mc[2] > 0)
                ep.SetText(25, 9, mc[2].ToString());
            //            if (mc[3] > 0)
            //                ep.SetText(6 + i, 6, mc[3].ToString());
            i++;
            //            ep.SetText(6 + i, 1, "Visa");
            if (vis[2] > 0)
                ep.SetText(26, 9, vis[2].ToString());
            //            if (vis[3] > 0)
            //                ep.SetText(6 + i, 6, vis[3].ToString());
            i++;
            //            ep.SetText(6 + i, 1, "Сервисные карты");
            if (nfc[2] > 0)
                ep.SetText(27, 9, nfc[2].ToString());
            //            if (nfc[3] > 0)
            //                ep.SetText(6 + i, 6, nfc[3].ToString());
            i++;
            if (ser[2] > 0)
                ep.SetText(28, 9, ser[2].ToString());
            //            if (ser[3] > 0)
            //                ep.SetText(6 + i, 6, ser[3].ToString());
            i++;
            //            ep.SetText(6 + i, 1, "Карты МИР");
            if (mir[2] > 0)
                ep.SetText(29, 9, mir[2].ToString());
            //            if (mir[3] > 0)
            //                ep.SetText(6 + i, 6, mir[3].ToString());
            i++;
            //            ep.SetText(6 + i, 1, "Пин-конверты");
            if (pin3 > 0)
                ep.SetText(30, 9, pin3.ToString());
            //            if (pin4 > 0)
            //                ep.SetText(6 + i, 6, pin4.ToString());
            //          i++;
            //            ep.SetText(6 + i, 1, "Итого карт");
            if (mc[2] + vis[2] + ser[2] + nfc[2] + mir[2]> 0)
                ep.SetText(89, 9, (mc[2] + vis[2] + ser[2] +nfc[2] + mir[2]).ToString());
            //            if (mc[3] + vis[3] + ser[3] > 0)
            //                ep.SetText(6 + i, 6, (mc[3] + vis[3] + ser[3] + nfc[3]).ToString());
            i++;
            //ep.ShowRows(5, i + 6 - 1);
        }

        /*
        private void GivenCardBranch_old(ExcelAp ep, string identDep)
        {
            DataSet ds = new DataSet();
            int i = 0;
            string res = "";
            SqlCommand comm = new SqlCommand();
            ds.Clear();
            comm.CommandText = "select id_prb, prod_name, bank_name, id_type from V_ProductsBanks_T order by id_sort";
            res = (string)Database2.ExecuteCommand(comm, ref ds, null);
            ArrayList al = new ArrayList();
            foreach (DataRow dr in ds.Tables[0].Rows)
                al.Add(new MyProd(Convert.ToInt32(dr["id_prb"]), dr["prod_name"].ToString(), dr["bank_name"].ToString(), Convert.ToInt32(dr["id_type"])));
            ds.Clear();
            comm.CommandText = String.Format("select id_prb, ident_dep from V_Cards_StorageDocs1 where priz_gen=1 and type=7 and date_doc=@dateS and ident_dep='{0}'", identDep);
            comm.Parameters.Add("@dateS", SqlDbType.DateTime).Value = DatePickerOne.SelectedDate.Date;
            Database2.ExecuteCommand(comm, ref ds, null);
            if (ds == null || ds.Tables.Count == 0)
                return;
            int pin1 = 0, pin2 = 0; int pin3 = 0, pin4 = 0, pin5 = 0, pin6 = 0;
            foreach (DataRow dr in ds.Tables[0].Rows)
            {
                foreach (MyProd mp in al)
                {
                    if (mp.ID == Convert.ToInt32(dr["id_prb"]))
                    {
                        mp.cnts[0]++; pin1++;
                        break;
                    }
                }
            }
            ds.Clear();
            comm.CommandText = String.Format("select id_prb, ident_dep from V_Cards_StorageDocs1 where priz_gen=1 and type=9 and date_doc=@dateS and ident_dep='{0}'", identDep);
            Database2.ExecuteCommand(comm, ref ds, null);
            if (ds == null || ds.Tables.Count == 0)
                return;
            foreach (DataRow dr in ds.Tables[0].Rows)
            {
                foreach (MyProd mp in al)
                {
                    if (mp.ID == Convert.ToInt32(dr["id_prb"]))
                    {
                        mp.cnts[4]++; pin5++;
                    }
                }
            }
            object obj = null;
            comm.Parameters.Clear();
            comm.CommandText = String.Format("select id from Branchs where ident_dep='{0}'", identDep);
            Database2.ExecuteScalar(comm, ref obj, null);
            int id0 = Convert.ToInt32(obj);
            comm.CommandText = String.Format("select department from Branchs where ident_dep='{0}'", identDep);
            Database2.ExecuteScalar(comm, ref obj, null);
            string depName = Convert.ToString(obj);

            comm.CommandText = "select count(*) from Cards where id_prb=@prb and id_branchCard=@bid and dateReceipt<=@dateS and (dateClient is null or dateClient>@dateS) and (dateSendTerminate is null or dateSendTerminate>@dateS)";
            comm.Parameters.Add("@prb", SqlDbType.Int);
            comm.Parameters.Add("@dateS", SqlDbType.DateTime).Value = DatePickerOne.SelectedDate.Date;
            comm.Parameters.Add("@bid", SqlDbType.Int);
            foreach (MyProd mp in al)
            {
                comm.Parameters["@prb"].Value = mp.ID;
                comm.Parameters["@bid"].Value = id0;
                res = Database2.ExecuteScalar(comm, ref obj, null);
                mp.cnts[2] = Convert.ToInt32(obj);
                pin3 += Convert.ToInt32(obj);
            }
            i = 0;
            int[] mc = new int[MyProd.cntMax];
            int[] vis = new int[MyProd.cntMax];
            //            int[] pin = new int[MyProd.cntMax];
            int[] ser = new int[MyProd.cntMax];
            for (int t = 0; t < MyProd.cntMax; t++)
            {
                mc[t] = 0;
                vis[t] = 0;
                //                pin[t] = 0;
                ser[t] = 0;
            }
            foreach (MyProd mp in al)
            {
                if (mp.Name.ToLower().StartsWith("mc") || mp.Name.ToLower().StartsWith("master"))
                {
                    mc[0] += mp.cnts[0];
                    mc[1] += mp.cnts[1];
                    mc[2] += mp.cnts[2];
                    mc[3] += mp.cnts[3];
                    mc[4] += mp.cnts[4];
                    mc[5] += mp.cnts[5];
                }
                if (mp.Name.ToLower().StartsWith("visa"))
                {
                    vis[0] += mp.cnts[0];
                    vis[1] += mp.cnts[1];
                    vis[2] += mp.cnts[2];
                    vis[3] += mp.cnts[3];
                    vis[4] += mp.cnts[4];
                    vis[5] += mp.cnts[5];
                }
                if (mp.Name.ToLower().StartsWith("белый"))
                {
                    ser[0] += mp.cnts[0];
                    ser[1] += mp.cnts[1];
                    ser[2] += mp.cnts[2];
                    ser[3] += mp.cnts[3];
                    ser[4] += mp.cnts[4];
                    ser[5] += mp.cnts[5];
                }
                if (mp.cnts[0] + mp.cnts[1] + mp.cnts[4] + mp.cnts[5] != 0)
                {
                    ep.SetText(26 + i, 1, mp.Name);
                    if (mp.cnts[0] != 0)
                        ep.SetText(26 + i, 6, mp.cnts[0].ToString());
//                    if (mp.cnts[1] != 0)
//                        ep.SetText(26 + i, 8, mp.cnts[1].ToString());
                    if (mp.cnts[4] != 0)
                        ep.SetText(26 + i, 7, mp.cnts[4].ToString());
//                    if (mp.cnts[5] != 0)
//                        ep.SetText(26 + i, 9, mp.cnts[5].ToString());
                    i++;
                }
            }
            ep.SetText(26 + i, 1, "Пин-конверты");
            if (pin1 > 0)
                ep.SetText(26 + i, 6, pin1.ToString());
//            if (pin2 > 0)
//                ep.SetText(26 + i, 8, pin2.ToString());
            if (pin5 > 0)
                ep.SetText(26 + i, 7, pin5.ToString());
//            if (pin6 > 0)
//                ep.SetText(26 + i, 9, pin6.ToString());
            ep.ShowRows(26, 26 + i);
            if (pin1 > 0)
                ep.SetText(90, 6, pin1.ToString());
//            if (pin2 > 0)
//                ep.SetText(90, 8, pin2.ToString());
            if (pin5 > 0)
                ep.SetText(90, 7, pin5.ToString());
//            if (pin6 > 0)
//                ep.SetText(90, 9, pin6.ToString());

            ep.SetText(15, 1, depName);
            ep.SetText_Name("date_string", String.Format("Дата {0:dd.MM.yyyy}", DatePickerOne.SelectedDate));
            ep.SetText_Name("date_string1", String.Format("от {0: dd MM yyyy}", DatePickerOne.SelectedDate));
            ///////////////////////////
            // лист 3 распоряжение сводный
            ///////////////////////////
            ep.SetWorkSheet(3);
            ep.SetText(25, 2, "MasterCard");
            if (mc[0] != 0)
                ep.SetText(25, 7, mc[0].ToString());
            //            if (mc[1] != 0)
            //                ep.SetText(25, 8, mc[1].ToString());
            if (mc[4] != 0)
                ep.SetText(25, 8, mc[4].ToString());
            //            if (mc[5] != 0)
            //                ep.SetText(25, 9, mc[5].ToString());
            ep.SetText(26, 2, "Visa");
            if (vis[0] != 0)
                ep.SetText(26, 7, vis[0].ToString());
            //            if (vis[1] != 0)
            //                ep.SetText(26, 8, vis[1].ToString());
            if (vis[4] != 0)
                ep.SetText(26, 8, vis[4].ToString());
            //            if (vis[5] != 0)
            //                ep.SetText(26, 9, vis[5].ToString());
            ep.SetText(27, 2, "Сервисные карты");
            if (ser[0] != 0)
                ep.SetText(27, 7, ser[0].ToString());
            //            if (ser[1] != 0)
            //                ep.SetText(27, 8, ser[1].ToString());
            if (ser[4] != 0)
                ep.SetText(27, 8, ser[4].ToString());
            //            if (ser[5] != 0)
            //                ep.SetText(27, 9, ser[5].ToString());
            ep.SetText(28, 2, "Пин-конверты");
            if (pin1 > 0)
                ep.SetText(28, 7, pin1.ToString());
            //            if (pin2 > 0)
            //                ep.SetText(28, 8, pin2.ToString());
            if (pin5 > 0)
                ep.SetText(28, 8, pin5.ToString());
            //            if (pin6 > 0)
            //                ep.SetText(28, 9, pin6.ToString());
            ep.ShowRows(25, 28);
            if (mc[0] + vis[0] + ser[0] > 0)
                ep.SetText(89, 7, (mc[0] + vis[0] + ser[0]).ToString());
            //            if (mc[1] + vis[1] + ser[1] > 0)
            //                ep.SetText(89, 8, (mc[1]+vis[1]+ser[1]).ToString());
            if (mc[4] + vis[4] + ser[4] > 0)
                ep.SetText(89, 8, (mc[4] + vis[4] + ser[4]).ToString());
            //            if (mc[5] + vis[5] + ser[5] > 0)
            //                ep.SetText(89, 9, (mc[5]+vis[5]+ser[5]).ToString());
            ep.SetText(15, 1, depName);
            ep.SetText_Name("date_string2", String.Format("Дата {0:dd.MM.yyyy}", DatePickerOne.SelectedDate));
            ep.SetText_Name("date_string3", String.Format("от {0: dd MM yyyy}", DatePickerOne.SelectedDate));
            ///////////////////////////
            // лист 2 остатки
            ///////////////////////////
            ep.SetWorkSheet(2);
            i = 0;
            ep.SetText(5, 1, "Указать имя филиала");
//            ep.SetText(5, 6, "ГО БТА");
            foreach (MyProd mp in al)
            {
                if (mp.cnts[2] != 0 || mp.cnts[3] != 0)
                {
                    ep.SetText(6 + i, 1, mp.Name);
                    if (mp.cnts[2] != 0)
                        ep.SetText(6 + i, 5, mp.cnts[2].ToString());
//                    if (mp.cnts[3] != 0)
//                        ep.SetText(6 + i, 6, mp.cnts[3].ToString());
                    i++;
                }
            }
            ep.SetText(6 + i, 1, "Пин-конверты");
            if (pin3 > 0)
                ep.SetText(6 + i, 5, pin3.ToString());
//            if (pin4 > 0)
//                ep.SetText(6 + i, 6, pin4.ToString());
            i++;
            ep.SetText(6 + i, 1, "Итого карт");
            if (pin3 > 0)
                ep.SetText(6 + i, 5, pin3.ToString());
//            if (pin4 > 0)
//                ep.SetText(6 + i, 6, pin4.ToString());
            i++;
            ep.ShowRows(5, i + 6 - 1);
            ep.SetText(5, 1, depName);
            //////////////////////////////////
            // лист 3 остатки сводные - возвращаемся на лист 3
            //////////////////////////////////
            ep.SetWorkSheet(3);
            i = 0;
            //            ep.SetText(5, 5, "ГО РКЦ");
            //            ep.SetText(5, 6, "ГО БТА");
            //            i = 0;
            //            ep.SetText(6 + i, 1, "MasterCard");
            if (mc[2] > 0)
                ep.SetText(25, 9, mc[2].ToString());
            //            if (mc[3] > 0)
            //                ep.SetText(6 + i, 6, mc[3].ToString());
            i++;
            //            ep.SetText(6 + i, 1, "Visa");
            if (vis[2] > 0)
                ep.SetText(26, 9, vis[2].ToString());
            //            if (vis[3] > 0)
            //                ep.SetText(6 + i, 6, vis[3].ToString());
            i++;
            //            ep.SetText(6 + i, 1, "Сервисные карты");
            if (ser[2] > 0)
                ep.SetText(27, 9, ser[2].ToString());
            //            if (ser[3] > 0)
            //                ep.SetText(6 + i, 6, ser[3].ToString());
            i++;
            //            ep.SetText(6 + i, 1, "Пин-конверты");
            if (pin3 > 0)
                ep.SetText(28, 9, pin3.ToString());
            //            if (pin4 > 0)
            //                ep.SetText(6 + i, 6, pin4.ToString());
            //          i++;
            //            ep.SetText(6 + i, 1, "Итого карт");
            if (mc[2] + vis[2] + ser[2] > 0)
                ep.SetText(89, 9, (mc[2] + vis[2] + ser[2]).ToString());
            //            if (mc[3] + vis[3] + ser[3] > 0)
            //                ep.SetText(6 + i, 6, (mc[3] + vis[3] + ser[3]).ToString());
            i++;
            //ep.ShowRows(5, i + 6 - 1);
        }
        */

        void setMemoric(BranchStore branch, ArrayList al, DateTime curDateS, DateTime curDateE, String sqlUserV, String sqlUser)
        {

            DataSet ds = new DataSet();
            SqlCommand comm = new SqlCommand();

            //получение
            ds.Clear();
            comm.Parameters.Clear();
            comm.CommandText = "select id_prb,count(*) as count_prb from V_Cards_StorageDocs1 where" + sqlUserV + "date_time>=@dateS and date_time<@dateE and id_branch=@id_branch and (type=6 or type=10) and priz_gen=1 group by id_prb";
            comm.Parameters.Add("@dateS", SqlDbType.DateTime).Value = curDateS;
            comm.Parameters.Add("@dateE", SqlDbType.DateTime).Value = curDateE;
            comm.Parameters.Add("@id_branch", SqlDbType.Int).Value = branch.id;
            Database2.ExecuteCommand(comm, ref ds, null);
            foreach (MyProd mp in al)
            {
                for (int n = 0; mp.Type != 2 && ds.Tables.Count > 0 && n < ds.Tables[0].Rows.Count; n++)
                {
                    if (mp.ID == Convert.ToInt32(ds.Tables[0].Rows[n]["id_prb"])) // карты получены
                    {
                        mp.cnts[0] = Convert.ToInt32(ds.Tables[0].Rows[n]["count_prb"]);
                        break;
                    }
                }
                if (mp.Type == 2) // Пин конверты получены
                {
                    comm.Parameters.Clear();
                    //comm.CommandText = "select count(*) from cards where id in (select id_card from cards_storagedocs where id_doc in (select id from storagedocs where" + sqlUser + "date_time>=@dateS and date_time<@dateE and (type=6 or type=10) and priz_gen=1 and id_branch=@id_branch)) and isPin=1";
                    comm.CommandText = "select count(id_card) from cards_storagedocs cs join cards c on c.id=cs.id_card where id_doc in (select id from storagedocs where" + sqlUser + "date_time>=@dateS and date_time<@dateE and (type=6 or type=10) and priz_gen=1 and id_branch=@id_branch) and isPin=1";
                    comm.Parameters.Add("@dateS", SqlDbType.DateTime).Value = curDateS;
                    comm.Parameters.Add("@dateE", SqlDbType.DateTime).Value = curDateE;
                    comm.Parameters.Add("@id_branch", SqlDbType.Int).Value = branch.id;
                    object obj = comm.ExecuteScalar();
                    mp.cnts[0] = (obj == null || obj == DBNull.Value) ? 0 : Convert.ToInt32(obj);
                }
                branch.addCount(mp.Type, mp.Name, 0, mp.cnts[0]);
                mp.cnts[0] = 0;

            }
            //выдача
            ds.Clear();
            comm.Parameters.Clear();
            comm.CommandText = "select id_prb,count(*) as count_prb from V_Cards_StorageDocs1 where" + sqlUserV + "date_time>=@dateS and date_time<@dateE and id_branch=@id_branch and (type=7 or type=13 or type=19) and priz_gen=1 group by id_prb";
            comm.Parameters.Add("@dateS", SqlDbType.DateTime).Value = curDateS;
            comm.Parameters.Add("@dateE", SqlDbType.DateTime).Value = curDateE;
            comm.Parameters.Add("@id_branch", SqlDbType.Int).Value = branch.id;
            Database2.ExecuteCommand(comm, ref ds, null);
            foreach (MyProd mp in al)
            {
                for (int n = 0; mp.Type != 2 && ds.Tables.Count > 0 && n < ds.Tables[0].Rows.Count; n++)
                {
                    if (mp.ID == Convert.ToInt32(ds.Tables[0].Rows[n]["id_prb"])) // карты выданы
                    {
                        mp.cnts[1] = Convert.ToInt32(ds.Tables[0].Rows[n]["count_prb"]);
                        break;
                    }
                }
                if (mp.Type == 2) // Пин конверты выданы
                {
                    comm.Parameters.Clear();
                    //comm.CommandText = "select count(*) from cards where id in (select id_card from cards_storagedocs where id_doc in (select id from storagedocs where" + sqlUser + "date_time>=@dateS and date_time<@dateE and (type=7 or type=13 or type=19) and priz_gen=1 and id_branch=@id_branch)) and isPin=1";
                    comm.CommandText = "select count(id_card) from cards_storagedocs cs join cards c on c.id=cs.id_card where id_doc in (select id from storagedocs where" + sqlUser + "date_time>=@dateS and date_time<@dateE and (type=7 or type=13 or type=19) and priz_gen=1 and id_branch=@id_branch) and isPin=1";
                    comm.Parameters.Add("@dateS", SqlDbType.DateTime).Value = curDateS;
                    comm.Parameters.Add("@dateE", SqlDbType.DateTime).Value = curDateE;
                    comm.Parameters.Add("@id_branch", SqlDbType.Int).Value = branch.id;
                    object obj = comm.ExecuteScalar();
                    mp.cnts[1] = (obj == null || obj == DBNull.Value) ? 0 : Convert.ToInt32(obj);
                }
                branch.addCount(mp.Type, mp.Name, 1, mp.cnts[1]);
                mp.cnts[1] = 0;
            }
            //возврат
            ds.Clear();
            comm.Parameters.Clear();
            comm.CommandText = "select id_prb,count(*) as count_prb from V_Cards_StorageDocs1 where" + sqlUserV + "date_time>=@dateS and date_time<@dateE and id_branch=@id_branch and (type=9 or type=11) and priz_gen=1 group by id_prb";
            comm.Parameters.Add("@dateS", SqlDbType.DateTime).Value = curDateS;
            comm.Parameters.Add("@dateE", SqlDbType.DateTime).Value = curDateE;
            comm.Parameters.Add("@id_branch", SqlDbType.Int).Value = branch.id;
            Database2.ExecuteCommand(comm, ref ds, null);
            foreach (MyProd mp in al)
            {
                for (int n = 0; mp.Type != 2 && ds.Tables.Count > 0 && n < ds.Tables[0].Rows.Count; n++)
                {
                    if (mp.ID == Convert.ToInt32(ds.Tables[0].Rows[n]["id_prb"])) // возврат карт
                    {
                        mp.cnts[2] = Convert.ToInt32(ds.Tables[0].Rows[n]["count_prb"]);
                        break;
                    }
                }
                if (mp.Type == 2) // Возврат пин конвертов
                {
                    comm.Parameters.Clear();
                    //comm.CommandText = "select count(*) from cards where id in (select id_card from cards_storagedocs where id_doc in (select id from storagedocs where" + sqlUser + "date_time>=@dateS and date_time<@dateE and (type=9 or type=11) and priz_gen=1 and id_branch=@id_branch)) and isPin=1";
                    comm.CommandText = "select count(id_card) from cards_storagedocs cs join cards c on c.id=cs.id_card where id_doc in (select id from storagedocs where" + sqlUser + "date_time>=@dateS and date_time<@dateE and (type=9 or type=11) and priz_gen=1 and id_branch=@id_branch) and isPin=1";
                    comm.Parameters.Add("@dateS", SqlDbType.DateTime).Value = curDateS;
                    comm.Parameters.Add("@dateE", SqlDbType.DateTime).Value = curDateE;
                    comm.Parameters.Add("@id_branch", SqlDbType.Int).Value = branch.id;
                    object obj = comm.ExecuteScalar();
                    mp.cnts[2] = (obj == null || obj == DBNull.Value) ? 0 : Convert.ToInt32(obj);
                }
                branch.addCount(mp.Type, mp.Name, 2, mp.cnts[2]);
                mp.cnts[2] = 0;
            }
            ds.Clear();
            comm.Parameters.Clear();
            ds = null;
            comm = null;
        }

    

        private class MovingStockDebetCredit
        {
            public int num = -1;
            public String debet = String.Empty;
            public String credit = String.Empty;
            public String ground = String.Empty;
        }
        private MovingStockDebetCredit[] loadMovingStock()
        {
            if (ConfigurationSettings.AppSettings["MovingStock"] == null) return null;
            String[] mv = ConfigurationSettings.AppSettings["MovingStock"].Split(';');
            if (mv.Length < 1) return null;
            MovingStockDebetCredit[] mvdc = new MovingStockDebetCredit[mv.Length];
            for (int i = 0; i < mv.Length; i++)
            {
                mvdc[i] = new MovingStockDebetCredit();
                String[] ndc = mv[i].Split(',');
                if (ndc != null && ndc.Length > 0) 
                {
                    String ns = ndc[0];
                    ns = ns.Replace("\n", "");
                    ns = ns.Replace("\r", "");
                    ns = ns.Replace("\t", "");
                    mvdc[i].num = Convert.ToInt32(ns.Trim());
                    if (ndc.Length > 1) mvdc[i].debet = ndc[1].Trim();
                    if (ndc.Length > 2) mvdc[i].credit = ndc[2].Trim();
                    if (ndc.Length > 3) mvdc[i].ground = ndc[3].Trim(); ;
                }
            }
            return mvdc;
        }

        private void GenReportKgrt(ExcelAp ep)
        {
            DataSet ds = new DataSet();
            SqlCommand comm = new SqlCommand();
            DateTime dt = DatePickerStart.SelectedDate.Date;
            DateTime dte = DatePickerEnd.SelectedDate.Date;

            comm.CommandText = "select count(c.id) as [count], p.prefix_ow," +
                               " stat =" +
                               " case" +
                               " when c.id_stat = 3 or c.id_stat = 4 then 0" +
                               " when  c.id_stat = 8 then 1" +
                               " else 2" +
                               " end" +
                               " from StorageDocs sd" +
                               " join Cards_StorageDocs csd on csd.id_doc = sd.id" +
                               " join Cards c on c.id = csd.id_card" +
                               " join Products_Banks pb on pb.id = c.id_prb" +
                               " join Products p on p.id = pb.id_prod" +
                               " where" +
                               " sd.priz_gen!=0" +
                               " and p.prefix_ow in ('KGF','KGP','KGN')" +
                               " and ((sd.type = 5 and c.id_stat = 3) or (sd.type = 6 and c.id_stat =4) or (sd.type = 7 and c.id_stat =8))" +
                               " and sd.date_doc >= @dt and and sd.date_doc <= @dte" +
                               " group by p.prefix_ow, c.id_stat";
            
            comm.Parameters.Add("@dt", SqlDbType.DateTime).Value = dt;
            comm.Parameters.Add("@dte", SqlDbType.DateTime).Value = dte;

            int[] countKGF = new int[] { 0, 0 };
            int[] countKGP = new int[] { 0, 0 };
            int[] countKGN = new int[] { 0, 0 };

            Database2.ExecuteCommand(comm, ref ds, null);
            
            
            for (int n = 0;ds.Tables.Count > 0 && n < ds.Tables[0].Rows.Count; n++)
            {
                int count = Convert.ToInt32(ds.Tables[0].Rows[n]["count"]);
                int stat = Convert.ToInt32(ds.Tables[0].Rows[n]["stat"]);
                string prefix_ow = Convert.ToString(ds.Tables[0].Rows[n]["prefix_ow"]);
                int[] c = null;
                if (prefix_ow == "KGF") c = countKGF;
                if (prefix_ow == "KGP") c = countKGP;
                if (prefix_ow == "KGN") c = countKGN;
                if (c == null) continue;
                if (stat == 0) // || stat == 1)
                {
                    c[stat] += count;
                }
            }
            ds.Clear();
            comm.CommandText = "select count(c.id) as [count], p.prefix_ow" +
                               " from cards c" +
                               " join Products_Banks pb on pb.id = c.id_prb" +
                               " join Products p on p.id = pb.id_prod" +
                               " where" +
                               " c.id_stat = 8" +
                               " and p.prefix_ow in ('KGF','KGP','KGN')" +
                               " and c.dateClient >= @dt and c.dateClient <= @dte" +
                               " group by  p.prefix_ow";

            Database2.ExecuteCommand(comm, ref ds, null);
            for (int n = 0; ds.Tables.Count > 0 && n < ds.Tables[0].Rows.Count; n++)
            {
                int count = Convert.ToInt32(ds.Tables[0].Rows[n]["count"]);
                string prefix_ow = Convert.ToString(ds.Tables[0].Rows[n]["prefix_ow"]);
                int[] c = null;
                if (prefix_ow == "KGF") c = countKGF;
                if (prefix_ow == "KGP") c = countKGP;
                if (prefix_ow == "KGN") c = countKGN;
                if (c == null) continue;
                c[1] += count;
            }

            ep.SetText(2, 4, String.Format(" {0:dd.MM.yyyy} - {0:dd.MM.yyyy}", dt));
            if (countKGF[0] != 0) ep.SetText(5, 4, countKGF[0].ToString());
            if (countKGF[1] != 0) ep.SetText(5, 5, countKGF[1].ToString());
            if (countKGP[0] != 0) ep.SetText(8, 4, countKGP[0].ToString());
            if (countKGP[1] != 0) ep.SetText(8, 5, countKGP[1].ToString());
            if (countKGN[0] != 0) ep.SetText(10, 4, countKGN[0].ToString());
            if (countKGN[1] != 0) ep.SetText(10, 5, countKGN[1].ToString());


        }

        //public void ReturnFile(string fileName, byte[] fileInBytes)
        //{
        //    System.Web.HttpResponse response = System.Web.HttpContext.Current.Response;
        //    response.Clear();
        //    response.Buffer = true;
        //    response.Charset = "";
        //    response.ContentType = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet";
        //    response.AddHeader("content-disposition", $"attachment;filename={fileName}");
        //    response.BinaryWrite(fileInBytes);
        //    //add following code
        //    response.Flush();
        //    response.End();
        //}
        public bool ReturnFile(string fileName, byte[] fileInBytes)
        {
            System.Web.HttpResponse response = System.Web.HttpContext.Current.Response;

            response.ClearHeaders();
            response.ClearContent();

            DialogUtils.SetCookieResponse(response);

            response.HeaderEncoding = System.Text.Encoding.Default;
            response.AddHeader("Content-Disposition", "attachment; filename=" + fileName);
            response.AddHeader("Content-Length", fileInBytes.Length.ToString());
            response.ContentType = "application/octet-stream";
            response.Cache.SetCacheability(HttpCacheability.NoCache);

            response.BufferOutput = true;


            response.BinaryWrite(fileInBytes);
            //add following code
            response.Flush();
            response.End();
            return true;
        }

        public bool ReturnXls(System.Web.HttpResponse resp, string fileName, byte[] fileInBytes)
        {
            System.IO.FileInfo f = new System.IO.FileInfo(fileName);
            resp.ClearHeaders();
            resp.ClearContent();

            DialogUtils.SetCookieResponse(resp);

            resp.HeaderEncoding = System.Text.Encoding.Default;
            resp.AddHeader("Content-Disposition", "attachment; filename=" + f.Name);
            resp.AddHeader("Content-Length", f.Length.ToString());
            resp.ContentType = "application/octet-stream";
            resp.Cache.SetCacheability(HttpCacheability.NoCache);
            /*
            resp.BufferOutput = false;
            resp.WriteFile(f.FullName);
            resp.Flush();
            resp.End();
            */

            resp.BufferOutput = true;

            //resp.BinaryWrite();
            resp.BinaryWrite(fileInBytes);
            //resp.WriteFile(f.FullName);
            //resp.End();

            return true;
        }

        public string GetReportNameByType(int reportType)
        {
            switch (reportType)
            {
                case 1: return "Записка на выпуск карт";
                case 2: return "Cорт лист";
                case 3: return "Выпущенные карты по продуктам";
                case 4: return "Динамика выпуска карт";
                case 6: return "Состояние хранилища старое";
                case 7: return "Ценности ниже порогового значения";
                case 8: return "Уничтоженные карты по продуктам";
                case 9: return "Уничтоженные карты по подразделениям";
                case 10: return "Бракованные карты по продуктам";
                case 11: return "Закупочные договора";
                case 12: return "Сводный акт";
                case 14: return "Выданные ценности";
                case 15: return "Распоряжение на движение ценностей (новый вариант)";
                case 16: return "Карты выданные клиентам подразделением ГО";
                case 17: return "Карты выданные клиентам подразделения";
                case 18: return "Состояние хранилища новое";
                case 19: case 41: return "Отчет для филиала на конец дня или передачи между сменами";
                case 42: return "Отчет МЕМОРИАЛЬНЫЙ ОРДЕР";
                case 20: return "Отчет для Казанского филиала (консолидированный отчет)";
                case 43: return "Отчет для списания с подотчета МОЛ";
                case 44: return "Отчет по карте жителя";
                case 45: return "Отчет выданных карт по книге 124";
                default: return "Неизвестный тип отчета";
            }


        }

    }
}
