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
using System.Security.Cryptography;
using System.Text;
using System.Xml.Linq;
using OstCard.Data;
using System.Data.SqlClient;
using CardPerso.Administration;

namespace CardPerso
{
    public partial class StorDocCardRec : System.Web.UI.Page
    {
        DataSet ds = new DataSet();
        string res = "";
        int id_doc = 0; int id_type = 0; int id_branch=0;
        string stat = "";

        int branch_main_filial = 0;
        int branch_current = 0;
        bool perso = false;

        bool first = true;        

        public int cntSelectedCard = 0;

        String id_prb = "";
        
        ServiceClass sc = new ServiceClass();

        protected void CardCreating(object sender, ObjectDataSourceEventArgs e)
        {
            e.ObjectInstance = Session["StorDocCardRec"];
        }

        private String getExcludeProductsCardsSelect()
        {
            String sqlProducts = "";
            if (ConfigurationSettings.AppSettings["ExcludeProductsCardsSelect"] == null) return "";
            DataSet ds = new DataSet();
            String[] Products = ConfigurationSettings.AppSettings["ExcludeProductsCardsSelect"].Split(';');
            for (int i = 0; i < Products.Length; i++)
            {
                String[] bin_prefix = Products[i].Split('_');
                if (sqlProducts.Length > 0) sqlProducts += " or ";
                sqlProducts += "(bin='" + bin_prefix[0] + "' and prefix_ow='" + bin_prefix[1] + "')";
            }
            ds.Clear();
            if (sqlProducts.Length > 0)
                Database.ExecuteQuery("SELECT pb.id as id_prb FROM Products p join Products_Banks pb on p.id=pb.id_prod where " + sqlProducts, ref ds, null);
            String id_prb = "";
            for (int i = 0; ds.Tables.Count > 0 && i < ds.Tables[0].Rows.Count; i++)
            {
                if (id_prb.Length > 0) id_prb += ",";
                id_prb += (Convert.ToInt32(ds.Tables[0].Rows[i]["id_prb"])).ToString();
            }
            ds = null;
            return id_prb;
        }

        protected void Page_Load(object sender, EventArgs e)
        {
            if (!User.Identity.IsAuthenticated)
            {
                Response.Write("<script language=javascript>window.returnValue='1'; window.close();</script>");
                return;
            }
            Session["First"] = true;
            
            lock (Database.lockObjectDB)
            {

                perso = sc.UserAction(User.Identity.Name, Restrictions.Perso);

                branch_current = sc.BranchId(User.Identity.Name);
                branch_main_filial = BranchStore.getBranchMainFilial(branch_current, perso);
                
                id_doc = Convert.ToInt32(Request.QueryString["id_doc"]);
                id_type = Convert.ToInt32(Request.QueryString["type_doc"]);
                if (Request.QueryString["id_branch"] != "")
                    id_branch = Convert.ToInt32(Request.QueryString["id_branch"]);
                //bSave.Enabled = false;
                if (id_type == (int)TypeDoc.ToBook124)
                {
                    stat = "(id_stat=4) and (DateClient is null) and id_prop=1";
                    if (branch_main_filial > 0 && branch_main_filial != branch_current)
                    {
                        stat += " and id_BranchCurrent=" + branch_current.ToString();
                    }
                }
                if (id_type == (int)TypeDoc.FromBook124)
                {
                    stat = "(id_stat=19)";
                    if (branch_main_filial > 0 && branch_main_filial != branch_current)
                    {
                        stat += " and id_BranchCurrent=" + branch_current.ToString();
                    }
                }
                if (id_type == (int)TypeDoc.ToGoz)
                {
                    stat = "(id_stat=4) and (DateClient is null) and id_prop=1";
                    if (branch_main_filial > 0 && branch_main_filial != branch_current)
                    {
                        stat += " and id_BranchCurrent=" + branch_current.ToString();
                    }
                }
                if (id_type == (int)TypeDoc.FromGoz)
                {
                    stat = "(id_stat=24)";
                    if (branch_main_filial > 0 && branch_main_filial != branch_current)
                    {
                        stat += " and id_BranchCurrent=" + branch_current.ToString();
                    }
                }
                if (id_type == (int)TypeDoc.SendToClient)
                {
                    // 24.07.20 - для книги 124 было что выдавать может только тот, на ком числится. Теперь выдавать может любой сотрудник филиала
                    //stat = $"(id_stat=4 or ((id_stat=19 or id_stat=24) and id_person={sc.UserId(User.Identity.Name)})) and (DateClient is null) and id_prop=1";
                    stat = $"((id_stat=4 or id_stat=19) or (id_stat=24 and id_person={sc.UserId(User.Identity.Name)})) and (DateClient is null) and id_prop=1";
                    if (branch_main_filial > 0 && branch_main_filial != branch_current)
                    {
                        stat += " and id_BranchCurrent=" + branch_current.ToString();
                    }
                }

                if (id_type == (int) TypeDoc.SendToBank)
                {
                    stat = "id_stat=4";
                    if (branch_main_filial > 0 && branch_main_filial != branch_current)
                    {
                        stat += " and id_BranchCurrent=" + branch_current.ToString();
                    }
                }

                if (id_type == (int) TypeDoc.KillingCard) stat = "(id_stat=4 or id_stat=6)";
                if (id_type == (int)TypeDoc.ReturnToFilial) stat = "id_stat=8";
                if (id_type == (int)TypeDoc.FilialFilial) stat = "id_stat=4";
                if (id_type == (int)TypeDoc.SendToPodotchet) stat = "id_stat=4";
                if (id_type == (int)TypeDoc.FromGozToPodotchet) stat = "id_stat=24";
                if (id_type == (int)TypeDoc.SendToClientFromPodotchet)
                {
                    stat = $"id_stat=15 and id_person=(select id_person from [AccountablePerson_StorageDocs] where id_doc={id_doc})";
                }
                if (id_type == (int)TypeDoc.ReceiveFromPodotchet)
                {
                    stat = $"id_stat=15 and id_person=(select id_person from [AccountablePerson_StorageDocs] where id_doc={id_doc})";
                }
                if (id_type == (int)TypeDoc.ToGozFromPodotchet)
                {
                    stat = $"id_stat=15 and id_person=(select id_person from [AccountablePerson_StorageDocs] where id_doc={id_doc})";
                }


                if (NotSend.Checked == true)
                {
                    id_prb = getExcludeProductsCardsSelect();
                    if (stat.Length > 0) stat += " and ";
                    //stat += "datestart is not null and datediff(month,datestart,getdate())>=3";
                    stat += "datestart is not null and datediff(day,datestart,getdate())>90";
                    if(id_prb.Length>0) stat+=" and id_prb not in (" + id_prb + ")";
                }

                if (!IsPostBack)
                {
                    ZapCombo();
                    tbNumber.Focus();
                    //Refr();
                    NotSend.Text = " (более 3-х месяцев)";
                }
                
                Session["StorDocCardRec"] = null;
                Session["StorDocCardRec"] = this;
            }
        }

        private void ZapCombo()
        {
            ds.Clear();
            // везде вместо V_Cards2 было V_Cards
            // 30.10.18 убрал выборку для всех документов, тормозит
            //if (id_branch == 0)
            //if (id_type != 7) //убрал для выдачи поскольку сильно тормозил
            //    res = Database.ExecuteQuery(String.Format("select distinct company from V_Cards2 where({0}) and (id not in (select id_card from V_CardsTypeDocs where type={1}))", stat, id_type.ToString()), ref ds, null);
            //else
                //res = Database.ExecuteQuery(String.Format("select distinct company from V_Cards2 where (id_branchCard={0} or id_BranchCard_parent={0}) and ({1}) and (id not in (select id_card from V_CardsTypeDocs where type={2}))", id_branch.ToString(), stat, id_type.ToString()), ref ds, null);
            dListCompany.Items.Add("Все");
            //if (ds.Tables.Count > 0)
            //    for (int i = 0; i < ds.Tables[0].Rows.Count; i++)
            //        dListCompany.Items.Add(ds.Tables[0].Rows[i]["company"].ToString());

            ds.Clear();
            
            
                if (id_branch == 0)
                    res = Database.ExecuteQuery(String.Format("select distinct prod_name from V_Cards2 where({0}) and (id not in (select id_card from V_CardsTypeDocs where type={1})) order by prod_name", stat, id_type.ToString()), ref ds, null);
                else
                    res = Database.ExecuteQuery(String.Format("select distinct prod_name from V_Cards2 where (id_branchCard={0} or id_BranchCard_parent={0}) and ({1}) and (id not in (select id_card from V_CardsTypeDocs where type={2})) order by prod_name", id_branch.ToString(), stat, id_type.ToString()), ref ds, null);
            
            dListProduct.Items.Add("Все");
            if (ds.Tables.Count > 0)
                for (int i = 0; i < ds.Tables[0].Rows.Count; i++)
                    dListProduct.Items.Add(ds.Tables[0].Rows[i]["prod_name"].ToString());
        }

        private void Refr()
        {
            /*
            ds.Clear();
            if (id_branch==0)
                res = Database.ExecuteQuery(String.Format("select id,pan,fio,company from V_Cards2 where ({0}) and (id not in (select id_card from V_CardsTypeDocs where type={1})) {2}", stat, id_type.ToString(), lbSearch.Text), ref ds, null);
            else
                res = Database.ExecuteQuery(String.Format("select id,pan,fio,company from V_Cards2 where (id_branchCard={0} or id_BranchCard_parent={0}) and ({1}) and (id not in (select id_card from V_CardsTypeDocs where type={2})) {3}", id_branch.ToString(), stat, id_type.ToString() , lbSearch.Text), ref ds, null);

            if (ds.Tables.Count > 0)
            {
                gvCards.DataSource = ds.Tables[0];
                gvCards.DataBind();
            }
            lbCount.Text = "Карт: " + gvCards.Rows.Count.ToString();
            */
            
            Session["StorDocCardRec"] = null;
            Session["StorDocCardRec"] = this;
            gvCards.DataBind();
         }

        protected void bSave_Click(object sender, ImageClickEventArgs e)
        {
            if (!User.Identity.IsAuthenticated)
            {
                Response.Write("<script language=javascript>window.returnValue='1'; window.close();</script>");
                return;
            }
            lock (Database.lockObjectDB)
            {
                SqlCommand sqCom = Database.Conn.CreateCommand();
                int id_card = 0;

                sqCom.CommandText = "select priz_gen from StorageDocs where id=" + id_doc.ToString();
                object obj = sqCom.ExecuteScalar();
                if (Convert.ToBoolean(obj))
                {
                    lbInform.Text = "Нельзя добавить карты в подтвержденный документ";
                    return;
                }
                /*
                for (int i = 0; i < gvCards.Rows.Count; i++)
                {
                    id_card = Convert.ToInt32(gvCards.DataKeys[i].Values["id"]);

                    if (Database.CheckCardInDoc(id_doc, id_card, null))
                    {
                        sqCom.Parameters.Clear();
                        sqCom.CommandText = "insert into Cards_StorageDocs (id_doc,id_card,comment) values (@id_doc,@id_card,@comment)";
                        sqCom.Parameters.Add("@id_doc", SqlDbType.Int).Value = id_doc;
                        sqCom.Parameters.Add("@id_card", SqlDbType.Int).Value = id_card;
                        sqCom.Parameters.Add("@comment", SqlDbType.VarChar, 50).Value = tbComment.Text;
                        Database.ExecuteNonQuery(sqCom, null);
                    }
                }
                */
                ds.Clear();
                if (id_branch == 0)
                    res = Database.ExecuteQuery(String.Format("select id,pan,fio,company from V_Cards2 where ({0}) and (id not in (select id_card from V_CardsTypeDocs where type={1})) {2}", stat, id_type.ToString(), lbSearch.Text), ref ds, null);
                else
                    //res = Database.ExecuteQuery(String.Format("select id,pan,fio,company from V_Cards2 where (id_branchCard={0} or id_BranchCard_parent={0}) and ({1}) /*and (id not in (select id_card from V_CardsTypeDocs where type={2} and id_branch={0}))*/ {3}", id_branch.ToString(), stat, id_type.ToString(), lbSearch.Text), ref ds, null);
                    res = Database.ExecuteQuery(String.Format("select id,pan,fio,company from V_Cards2 where (id_branchCard={0} or id_BranchCard_parent={0}) and ({1}) {3}", id_branch.ToString(), stat, id_type.ToString(), lbSearch.Text), ref ds, null);
                if (ds.Tables.Count > 0)
                {
                    for (int i = 0; i < ds.Tables[0].Rows.Count; i++)
                    {
                        id_card = Convert.ToInt32(ds.Tables[0].Rows[i]["id"]);

                        if (Database.CheckCardInDoc(id_doc, id_card, null))
                        {
                            WebLog.LogClass.WriteToLog("StorDocCardRec.InsertCard Start id_doc={0}, id_card={1}, comment={2}", id_doc, id_card, tbComment.Text);
                            sqCom.Parameters.Clear();
                            sqCom.CommandText = "insert into Cards_StorageDocs (id_doc,id_card,comment) values (@id_doc,@id_card,@comment)";
                            sqCom.Parameters.Add("@id_doc", SqlDbType.Int).Value = id_doc;
                            sqCom.Parameters.Add("@id_card", SqlDbType.Int).Value = id_card;
                            sqCom.Parameters.Add("@comment", SqlDbType.VarChar, 50).Value = tbComment.Text;
                            Database.ExecuteNonQuery(sqCom, null);
                            WebLog.LogClass.WriteToLog("StorDocCardRec.InsertCard End id_doc={0}, id_card={1}", id_doc, id_card);
                        }
                    }
                }


                Response.Write("<script language=javascript>window.returnValue='1'; window.close();</script>");
            }
        }

        protected void bFilter_Click(object sender, ImageClickEventArgs e)
        {
            if (!User.Identity.IsAuthenticated)
            {
                Response.Write("<script language=javascript>window.returnValue='1'; window.close();</script>");
                return;
            }
            Session["First"] = false;
            first = false;
            lock (Database.lockObjectDB)
            {
                                
                if (cbDateProd.Checked)
                {
                    try
                    {
                        Convert.ToDateTime(DatePickerStart.DatePickerText);
                        Convert.ToDateTime(DatePickerEnd.DatePickerText);
                    }
                    catch
                    {
                        cbDateProd.Checked = false;
                        DatePickerStart.SelectedDate = DateTime.Now.Date;
                        DatePickerEnd.SelectedDate = DateTime.Now.Date;
                        return;
                    }
                }
                ArrayList al = new ArrayList();

                bSave.Visible = true;
                if (tbNumber.Text != "")
                {
                    ArrayList panArrayList = new ArrayList();
                    string[] pans = tbNumber.Text.Split(',');
                    foreach (string p in pans)
                    {
                        if (p.Trim().Length >= 16)
                        {
                            SHA1Managed sha = new SHA1Managed();
                            panArrayList.Add(String.Format("(panhash like '%{0}%')",
                                Utilities.Bin2AHex(sha.ComputeHash(Encoding.ASCII.GetBytes(p.Trim())))));
                        }
                        else
                            panArrayList.Add(String.Format("(pan like '%{0}%')", p.Trim()));
                    }
                    string[] pa = Array.CreateInstance(typeof(string), panArrayList.Count) as string[];
                    panArrayList.CopyTo(pa, 0);
                    al.Add($"({String.Join(" or ", pa)})");
                }

                if (tbFio.Text != "")
                {
                    ArrayList fioList = new ArrayList();
                    string[] fios = tbFio.Text.Split(',');
                    foreach (string f in fios)
                    {

                        fioList.Add($"(fio like '%{f.Trim()}%')");
                    }

                    string[] fa = Array.CreateInstance(typeof(string), fioList.Count) as string[];
                    fioList.CopyTo(fa, 0);
                    al.Add($"({String.Join(" or ", fa)})");
                }

                if (tbCompany.Text != "")
                    al.Add(String.Format("(company like '%{0}%')", tbCompany.Text));
                if (cbDateProd.Checked)
                {
                    al.Add(String.Format("(dateProd >= '{0:MM/dd/yyyy}')", DatePickerStart.SelectedDate));
                    al.Add(String.Format("(dateProd < '{0:MM/dd/yyyy}')", DatePickerEnd.SelectedDate.AddDays(1)));
                }
                if (dListCompany.SelectedIndex != 0)
                    //al.Add(String.Format("(company like '%{0}%')", dListCompany.SelectedItem.Value));
                    al.Add(String.Format("(company='{0}')", dListCompany.SelectedItem.Value));
                if (dListProduct.SelectedIndex != 0)
                    al.Add(String.Format("(prod_name like '{0}')", dListProduct.SelectedItem.Value));

                if (al.Count > 0)
                {
                    string[] all = Array.CreateInstance(typeof(string), al.Count) as string[];
                    al.CopyTo(all, 0);
                    string exlude = (cbExclude.Checked) ? "NOT" : "";
                    string search = String.Join(" and ", all);
                    lbSearch.Text = $" and {exlude} ({search})";
                }
                else
                    lbSearch.Text = "";
                gvCards.PageIndex = 0;
                Refr();
            }
        }

        protected void cbDateProd_CheckedChanged(object sender, EventArgs e)
        {
            //DatePickerEnd.Enabled = cbDateProd.Checked;
            //DatePickerStart.Enabled = cbDateProd.Checked;
        }

        
        public DataTable GetCards(int StartRowIndex, int MaximumRows, String SortExpression)
        {
            if (!IsPostBack)
                return null;
//            if ((bool)Session["First"])
//            {                
//                return null;
//            }
            try
            {
                lock (Database.lockObjectDB)
                {
                    ds.Clear();
                    string sqlstr="";
                    if (id_branch==0)
                    {
                        sqlstr += String.Format("WITH CardsDetails AS (SELECT id, ROW_NUMBER() OVER (ORDER BY id) AS RowNo FROM V_Cards2 where ({0}) and (id not in (select id_card from V_CardsTypeDocs where type={1})) {2})\r\n",stat, id_type.ToString(), lbSearch.Text);
                        sqlstr += String.Format("select CardsDetails.id,pan,fio,company from CardsDetails\r\n");
                        sqlstr += String.Format("inner join  V_Cards2 ON CardsDetails.id = V_Cards2.id\r\n");
                        sqlstr += String.Format("where ({0}) and (RowNo BETWEEN {3} AND {4}) and (CardsDetails.id not in (select id_card from V_CardsTypeDocs where type={1})) {2} ORDER BY  CardsDetails.id", stat, id_type.ToString(), lbSearch.Text, StartRowIndex + 1, StartRowIndex + 1 + MaximumRows);
                    }
                    else
                    {
                        sqlstr += String.Format("WITH CardsDetails AS (SELECT id, ROW_NUMBER() OVER (ORDER BY id) AS RowNo FROM V_Cards2 where (id_branchCard={0} or id_BranchCard_parent={0}) and ({1}) " + /*and (id not in (select id_card from V_CardsTypeDocs where type={2} and id_branch={0}))*/  "{3})\r\n", id_branch.ToString(), stat, id_type.ToString(), lbSearch.Text);
                        sqlstr += String.Format("select CardsDetails.id,pan,fio,company from CardsDetails\r\n");
                        sqlstr += String.Format("inner join  V_Cards2 ON CardsDetails.id = V_Cards2.id\r\n");
                        sqlstr += String.Format("where (RowNo BETWEEN {4} AND {5}) and (id_branchCard={0} or id_BranchCard_parent={0}) and ({1}) " + /*and (CardsDetails.id not in (select id_card from V_CardsTypeDocs where type={2} and id_branch={0}))*/ "{3}", id_branch.ToString(), stat, id_type.ToString(), lbSearch.Text, StartRowIndex + 1, StartRowIndex + 1 + MaximumRows);
                    }
                    res = Database.ExecuteQuery(sqlstr, ref ds, null);
                    
                    return (ds != null && ds.Tables.Count > 0) ? ds.Tables[0]:null;
                 }
            }
            catch (Exception ex)
            {
                lbCount.Text = "Ошибка: " + ex.Message;
            }
            return null;
        }



        public int GetCardsCount()
        {
            //lbCount.Text = "Карт: " + cardscount;
            //cntSelectedCard = cardscount;
            //if (cardscount < 1) bSave.Visible = false;
            
            //return cardscount;
            try
            {
                lock (Database.lockObjectDB)
                {
                    if (id_type != 99)
                    {
                        string sql = "";
                        if (id_branch == 0)
                        {
                            sql = String.Format("select COUNT(*) from V_Cards2 where ({0}) and (id not in (select id_card from V_CardsTypeDocs where type={1})) {2}", stat, id_type.ToString(), lbSearch.Text);
                        }
                        else
                        {
                            sql = String.Format("select COUNT(*) from V_Cards2 where (id_branchCard={0} or id_BranchCard_parent={0}) and ({1}) " + /*and (id not in (select id_card from V_CardsTypeDocs where type={2} and id_branch={0}))*/  "{3}", id_branch.ToString(), stat, id_type.ToString(), lbSearch.Text);
                        }
                        SqlCommand comm = new SqlCommand();
                        comm.CommandText = sql;
                        object obj = 0;
                        Database.ExecuteScalar(comm, ref obj, null);
                        int cnt = Convert.ToInt32(obj);
                        lbCount.Text = "Карт: " + cnt;
                        cntSelectedCard = cnt;
                        if (cnt < 1) bSave.Visible = false;
                        return cnt;
                    }
                 }
            }
            catch (Exception)
            {
                cntSelectedCard = 0;
                bSave.Visible = false;
            }
            return 0;
        }
    }
         

}
