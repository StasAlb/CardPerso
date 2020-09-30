using System;
using System.Collections;
using System.Configuration;
using System.Data;
using System.Linq;
using System.Web;
using System.Text;
using System.Web.Security;
using System.Security.Cryptography;
using System.Web.UI;
using System.Web.UI.HtmlControls;
using System.Web.UI.WebControls;
using System.Web.UI.WebControls.WebParts;
using System.Web.Configuration;
using System.Xml;
using System.Xml.Linq;
using OstCard.Data;
using System.Data.SqlClient;
using CardPerso.Administration;

namespace CardPerso
{
    public partial class Card : System.Web.UI.Page
    {
        private string res = "";
        private string branchSearch = "";
        private DataSet ds = new DataSet();
        ServiceClass sc = new ServiceClass();
        
        int branch_main_filial = 0;
        int branch_current = 0;
        bool perso = false;

    
        public int getBF() 
        { 
            return branch_main_filial;  
        }
        public int getBC()
        {
            return branch_current;
        }

        protected void CardCreating(object sender, ObjectDataSourceEventArgs e)
        {
            e.ObjectInstance = Session["Card"];
        }


        protected void Page_Load(object sender, EventArgs e)
        {

            DialogUtils.SetScriptSubmit(this);            
            
            ClientScript.RegisterHiddenField("resd", "");


            perso = sc.UserAction(User.Identity.Name, Restrictions.Perso);
            
            branch_current = sc.BranchId(User.Identity.Name);
            branch_main_filial = BranchStore.getBranchMainFilial(branch_current, perso);

            if (sc.UserAction(User.Identity.Name, Restrictions.AllData))
                branchSearch = " (1=1) ";
            else
                branchSearch = String.Format(" (id_branchCard={0} or id_branchCard in (select id from Branchs where id_parent={0} or id_parentTr={0}))", sc.BranchId(User.Identity.Name));
            if (!sc.UserAction(User.Identity.Name, Restrictions.CardsView))
                Response.Redirect("~\\Account\\Restricted.aspx", true);

            

            if (IsPostBack)
                return;

            lock(Database.lockObjectDB)
            {
                gvCard.PageSize = 50;
                
                // Решили оставить страницу с картами певоначально пустой
                lbSearch.Text = "where 1=0";
                
                //if (ServiceClass.UserAction(User.Identity.Name, Restrictions.Filial))
                //    lbSearch.Text = " where id_stat=4 ";
                //else
                //    lbSearch.Text = " where id_stat=2 ";
                lbSort.Text="order by dateProd desc,DepBranchCard,company,fio";
                Refr(true);
            }
            if (!sc.UserAction(User.Identity.Name, Restrictions.CardsDelete))
                bDeleteCards.Visible = false;
            if (!sc.UserAction(User.Identity.Name, Restrictions.CardsGive))
                bFilCards.Visible = false;
            GC.Collect();

            Session["Card"] = this;
            
        }
        private void Refr(bool conf)
        {
            /*
            ds.Clear();
            string str = lbSearch.Text.Trim();
            if (str.Length > 0)
                str += " and " + branchSearch;
            else
                str += " where " + branchSearch;
            if (sc.UserAction(User.Identity.Name, Restrictions.Transport))
            {
                string[] strs = ConfigurationSettings.AppSettings["Transport"].Split(',');
                for (int t = 0; t < strs.Length; t++)
                    strs[t] = String.Format("prefix_ow like '{0}'", strs[t]);
                str += String.Format(" and ({0})", String.Join(" or ", strs));
            }
            res = Database.ExecuteQuery(String.Format("select * from V_Cards {0} {1}", str, lbSort.Text), ref ds, null);
            if (ds != null && ds.Tables.Count > 0)
            {
                gvCard.DataSource = ds.Tables[0];
                gvCard.DataBind();
                bExcel.Enabled = ds.Tables[0].Rows.Count < 100;
            }
            if (conf)
            {
                string s_fld = Database.GetFiledsByUser(sc.UserId(User.Identity.Name), "card", null);
                if (s_fld != "") FuncClass.HideFields(s_fld, gvCard);
            }
            bResetFilter.Visible = (lbSearch.Text != "");
            SetButtonDoc();
            if (ds.Tables.Count > 0)
                lbCount.Text = "Кол-во: " + ds.Tables[0].Rows.Count.ToString();
            */
            Session["Card"] = null;
            Session["Card"] = this;
            if (conf)
            {
                string s_fld = Database.GetFiledsByUser(sc.UserId(User.Identity.Name), "card", null);
                if (s_fld != "") FuncClass.HideFields(s_fld, gvCard);
            }
            else gvCard.DataBind();
            
        }
              
        protected void bSetFilter_Click(object sender, ImageClickEventArgs e)
        {
            lock (Database.lockObjectDB)
            {
                lbSearch.Text = Request.Form["resd"].Replace("[", "'");
                lbSearch.Text = lbSearch.Text.Replace("]", "'");
                Refr(false);
            }
        }

        protected void bResetFilter_Click(object sender, ImageClickEventArgs e)
        {
            lock (Database.lockObjectDB)
            {
                lbSearch.Text = "";
                Refr(false);
            }
        }
        protected void bHistory_Click(object sender, ImageClickEventArgs e)
        {
        }


        protected void gvCard_Sorting(object sender, GridViewSortEventArgs e)
        {
            lock (Database.lockObjectDB)
            {
                int i = FuncClass.GetSortIndex(e.SortExpression, gvCard);

                if (lbSortIndex.Text != "")
                {
                    int ind = Convert.ToInt32(lbSortIndex.Text);
                    gvCard.Columns[ind].HeaderText = gvCard.Columns[ind].HeaderText.Replace("^", "");
                    gvCard.Columns[ind].HeaderStyle.BackColor = System.Drawing.Color.FromArgb(0, 102, 153);
                }

                lbSortIndex.Text = i.ToString();

                gvCard.Columns[i].HeaderStyle.BackColor = System.Drawing.Color.FromArgb(102, 153, 153);

                if ("order by " + e.SortExpression + " asc" == lbSort.Text)
                {
                    lbSort.Text = "order by " + e.SortExpression + " desc";
                    gvCard.Columns[i].HeaderText = gvCard.Columns[i].HeaderText + "^";
                }
                else
                    lbSort.Text = "order by " + e.SortExpression + " asc";

                Refr(false);
            }
        }

        protected void bExcel_Click(object sender, ImageClickEventArgs e)
        {
            lock (Database.lockObjectDB)
            {
                DataSet ds1 = new DataSet();
                string str = lbSearch.Text.Trim();
                if (str.Length > 0)
                    str += "and " + branchSearch;
                else
                    str += "where " + branchSearch;

                res = Database.ExecuteQuery(String.Format("select * from V_Cards {0} {1}", str, lbSort.Text), ref ds1, null);

                if (ds1 == null || ds1.Tables.Count == 0)
                    return;
                System.Globalization.CultureInfo oldCI = System.Threading.Thread.CurrentThread.CurrentCulture;
                System.Threading.Thread.CurrentThread.CurrentCulture = new System.Globalization.CultureInfo("en-US");
                string doc = "";
                ExcelAp ep = new ExcelAp();
                if (ep.RunApp(ConfigurationSettings.AppSettings["DocPath"] + "Empty.xls"))
                {
                    ep.SetWorkSheet(1);
                    ep.ExportGridExcel_Cards(gvCard, ds1.Tables[0]);
                    if (WebConfigurationManager.AppSettings["DocPath"] != null)
                    {
                        doc = String.Format("{0}Temp\\cards.xls", WebConfigurationManager.AppSettings["DocPath"]);
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
        protected void bExcelXML_Click(object sender, ImageClickEventArgs e)
        {
            lock (Database.lockObjectDB)
            {
                DataSet ds1 = new DataSet();
                string str = lbSearch.Text.Trim();
                if (str.Length > 0)
                    str += "and " + branchSearch;
                else
                    str += "where " + branchSearch;

                //res = Database.ExecuteQuery(String.Format("select pan,fio,company,prod_name,status,datestart,dateend,passport,account,dateprod,invoice,date_courier,depbranchinit,depbranchcard,datereceipt from V_Cards {0} {1}", str, lbSort.Text), ref ds1, null);
                res = Database.ExecuteQuery(String.Format("select * from V_Cards {0} {1}", str, lbSort.Text), ref ds1, null);

                if (ds1 == null || ds1.Tables.Count == 0)
                    return;
                Response.ClearHeaders();
                Response.ClearContent();

                
                DialogUtils.SetCookieResponse(Response);

                Response.Buffer = true;
                Response.HeaderEncoding = System.Text.Encoding.Default;
                Response.AddHeader("Content-Disposition", "attachment; filename=cards.xls");
                Response.ContentType = "application/vnd.ms_excel";
                Response.Cache.SetCacheability(HttpCacheability.NoCache);
                Response.BufferOutput = false;

                using (XmlTextWriter x = new XmlTextWriter(Response.OutputStream, Encoding.UTF8))
                {
                    x.WriteRaw("<?xml version=\"1.0\"?><?mso-application progid=\"Excel.Sheet\"?>");
                    x.WriteRaw("<Workbook xmlns=\"urn:schemas-microsoft-com:office:spreadsheet\" ");
                    x.WriteRaw("xmlns:o=\"urn:schemas-microsoft-com:office:office\" ");
                    x.WriteRaw("xmlns:x=\"urn:schemas-microsoft-com:office:excel\">");
                    x.WriteRaw("<Styles><Style ss:ID='sText'><NumberFormat ss:Format='@'/></Style>");
                    x.WriteRaw("<Style ss:ID='sDate'><NumberFormat ss:Format='[$-409]dd.mm.yyyy;@'/></Style></Styles>");
                    x.WriteRaw("<Worksheet ss:Name='Cards'>");
                    x.WriteRaw("<Table>");
                    //string[] colTypes = new string[ds1.Tables[0].Columns.Count];
                    //for (int i = 0; i < gvCard.Columns.Count; i++)
                    //{
                    //    BoundField bf = gvCard.Columns[i] as BoundField;
                    //    string coltype = bf.ty

                    //    if (coltype.Contains("datetime"))
                    //    {
                    //        colTypes[i] = "DateTime";
                    //        x.WriteRaw("<Column ss:StyleID='sText'/>");
                    //    }
                    //    else if (coltype.Contains("string"))
                    //    {
                    //        colTypes[i] = "String";
                    //        x.WriteRaw("<Column ss:StyleID='sText'/>");
                    //    }
                    //    else
                    //    {
                    //        x.WriteRaw("<Column/>");
                    //        if (coltype.Contains("boolean"))
                    //            colTypes[i] = "Boolean";
                    //        else
                    //            colTypes[i] = "Number";
                    //    }
                    //}
                    x.WriteRaw("<Row>");
                    for (int i = 0; i < gvCard.Columns.Count; i++)
                    {
                        BoundField bf = gvCard.Columns[i] as BoundField;
                        if (bf != null && bf.Visible)
                        {
                            x.WriteRaw("<Cell ss:StyleID='sText'><Data ss:Type='String'>");
                            x.WriteRaw(bf.HeaderText);
                            x.WriteRaw("</Data></Cell>");
                        }
                        TemplateField tf = gvCard.Columns[i] as TemplateField;
                        if (tf != null && tf.Visible && tf.HeaderText.Length > 0)
                        {
                            x.WriteRaw("<Cell ss:StyleID='sText'><Data ss:Type='String'>");
                            x.WriteRaw(tf.HeaderText);
                            x.WriteRaw("</Data></Cell>");
                        }
                    }
                    x.WriteRaw("</Row>");
                    foreach (DataRow dr in ds1.Tables[0].Rows)
                    {
                        x.WriteRaw("<Row>");
                        for (int i = 0; i < gvCard.Columns.Count; i++)
                        {
                            BoundField bf = gvCard.Columns[i] as BoundField;
                            if (bf != null && bf.Visible)
                            {
                                x.WriteRaw("<Cell><Data ss:Type='String'>");
                                if (ds1.Tables[0].Columns[bf.DataField].DataType.ToString().ToLower().Contains("datetime"))
                                    try
                                    {
                                        x.WriteRaw(((DateTime)dr[bf.DataField]).ToShortDateString());
                                    }
                                    catch
                                    {
                                        x.WriteRaw((dr.IsNull(bf.DataField)) ? "" : (dr[bf.DataField]).ToString().Trim());

                                    }
                                else
                                    x.WriteRaw((dr.IsNull(bf.DataField)) ? "" : (dr[bf.DataField]).ToString().Trim());
                                x.WriteRaw("</Data></Cell>");
                            }
                            TemplateField tf = gvCard.Columns[i] as TemplateField;
                            if (tf != null && tf.Visible && tf.HeaderText.Length > 0)
                            {
                                x.WriteRaw("<Cell><Data ss:Type='String'>");
                                x.WriteRaw((dr.IsNull(tf.SortExpression)) ? "" : (dr[tf.SortExpression]).ToString().Trim());
                                x.WriteRaw("</Data></Cell>");
                            }
                        }
                        x.WriteRaw("</Row>");
                    }
                    x.WriteRaw("</Table></Worksheet></Workbook>");
                }
                
 
                Response.Flush();
                
                Response.End();
            }
        }

        protected void bConfFieldD_Click(object sender, ImageClickEventArgs e)
        {
            Response.Redirect("~/Card.aspx");
        }

        protected void gvCard_PageIndexChanging(object sender, GridViewPageEventArgs e)
        {
            lock (Database.lockObjectDB)
            {
                /*
                gvCard.PageIndex = e.NewPageIndex;
                Refr(false);
                */ 
            }
        }
        protected void bPanSearch_Click(object sender, EventArgs e)
        {
            lock (Database.lockObjectDB)
            {
                ArrayList al = new ArrayList();

                SHA1Managed sha = new SHA1Managed();
                //обычные карты (16,19) или карты priority pass (18)
                if (tbPanSearch.Text.Trim().Length == 16 || tbPanSearch.Text.Trim().Length == 19 || tbPanSearch.Text.Trim().Length == 18)
                    al.Add(String.Format("(PanHash='{0}')", Utilities.Bin2AHex(sha.ComputeHash(Encoding.ASCII.GetBytes(tbPanSearch.Text.Trim())))));
                else
                    al.Add(String.Format("(Replace(Pan,' ','') like '%{0}%')", tbPanSearch.Text.Trim()));

                if (tbFioSearch.Text != "")
                    al.Add(String.Format("(fio like '%{0}%')", tbFioSearch.Text));

                if (al.Count > 0)
                {
                    string[] all = Array.CreateInstance(typeof(string), al.Count) as string[];
                    al.CopyTo(all, 0);
                    lbSearch.Text = "where " + String.Join(" and ", all);
                }

                Refr(false);
            }
        }
        protected void bDeleteCard_Click(object sender, ImageClickEventArgs e)
        {
            lock (Database.lockObjectDB)
            {
                string pan = "";
                CheckBox cb;
                SqlCommand comm = new SqlCommand();
                CheckBox cbA = (CheckBox)gvCard.HeaderRow.FindControl("cbSelectAll");
                if (cbA.Checked)
                {
                    string str = lbSearch.Text.Trim();
                    if (str.Length > 0)
                        str += " and " + branchSearch;
                    else
                        str += " where " + branchSearch;
                    str += " and (id_stat=1 or id_stat=2)";
                    comm.CommandText = "delete from Cards " + str;
                    Database.ExecuteNonQuery(comm, null);
                }
                else
                {
                    comm.Parameters.Add("@id", SqlDbType.Int);
                    foreach (GridViewRow gvr in gvCard.Rows)
                    {
                        cb = (CheckBox)gvr.FindControl("cbSelect");
                        if (cb.Checked && (int)gvCard.DataKeys[gvr.RowIndex].Values["id_stat"] == 1)
                        {
                            comm.Parameters["@id"].Value = (int)gvCard.DataKeys[gvr.RowIndex].Values["id"];
                            pan = gvCard.DataKeys[gvr.RowIndex].Values["pan"].ToString();
                            comm.CommandText = "delete from Cards where id=@id";
                            Database.ExecuteNonQuery(comm, null);
                            Database.Log(sc.UserGuid(User.Identity.Name), String.Format("Удалена карта {0} в статусе обработка", pan), null);
                        }
                    }
                }
                Refr(false);
            }
        }
/*
        protected void bGiveToClient_Click(object sender, ImageClickEventArgs e)
        {
            CheckBox cb;
            object obj = null;
            SqlCommand comm = new SqlCommand();
            comm.CommandText = "update Cards set dateClient=@dt, id_stat=@st where id=@id";
            comm.Parameters.Add("@dt", SqlDbType.DateTime).Value = DateTime.Now;
            comm.Parameters.Add("@st", SqlDbType.Int).Value = 8;
            comm.Parameters.Add("@id", SqlDbType.Int);
            foreach (GridViewRow gvr in gvCard.Rows)
            {
                cb = (CheckBox)gvr.FindControl("cbSelect");
                if (cb.Checked && (int)gvCard.DataKeys[gvr.RowIndex][1] == 4)
                {
                    comm.Parameters["@id"].Value = (int)gvCard.DataKeys[gvr.RowIndex][0];
                    comm.CommandText = "select pan from Cards where id=@id";
                    Database.ExecuteScalar(comm, ref obj, null);
                    comm.CommandText = "update Cards set dateClient=@dt, id_stat=@st where id=@id";
                    Database.ExecuteNonQuery(comm, null);
                    Database.Log(CardPerso.Administration.ServiceClass.UserGuid, String.Format("Карта {0} выдана клиенту", (obj == null) ? "" : (string)obj));
                }
            }
            Refr(false);
        }*/
        protected void cbSelect_CheckedChanged(object sender, EventArgs e)
        {
            lock (Database.lockObjectDB)
            {
                CheckBox cbA = (CheckBox)gvCard.HeaderRow.FindControl("cbSelectAll");
                CheckBox cb;
                foreach (GridViewRow gvr in gvCard.Rows)
                {
                    cb = (CheckBox)gvr.FindControl("cbSelect");
                    if (!cb.Checked)
                        cbA.Checked = false;
                }
                SetButtonDoc();
            }
        }
        protected void cbSelectAll_CheckedChanged(object sender, EventArgs e)
        {
            lock (Database.lockObjectDB)
            {
                CheckBox cbA = (CheckBox)gvCard.HeaderRow.FindControl("cbSelectAll");
                CheckBox cb;
                foreach (GridViewRow gvr in gvCard.Rows)
                {
                    cb = (CheckBox)gvr.FindControl("cbSelect");
                    if (cb.Visible)
                    cb.Checked = cbA.Checked;
                }
            }
            SetButtonDoc();
        }

        private void SetButtonDoc()
        {
            lock (Database.lockObjectDB)
            {
                bDeleteCards.Visible = false;
                bFilCards.Visible = false;

                CheckBox cb;
                int cnt = 0;
                int cntFil = 0;
                int cntNoFil = 0;
                foreach (GridViewRow gvr in gvCard.Rows)
                {
                    cb = (CheckBox)gvr.FindControl("cbSelect");
                    if (cb.Checked)
                    {
                        cnt++;
                        if (branch_main_filial > 0 && branch_main_filial == branch_current)
                        {
                            int curbranch = Convert.ToInt32(gvCard.DataKeys[gvr.RowIndex].Values["id_branchcurrent"]);
                            if (curbranch == branch_current) cntFil++;
                            else cntNoFil++;
                        }

                    }
                }
                if (gvCard.HeaderRow != null)
                {
                    CheckBox cbA = (CheckBox)gvCard.HeaderRow.FindControl("cbSelectAll");
                    if (cbA.Checked)
                    {
                        string str = lbSearch.Text.Trim();
                        if (str.Length > 0)
                            str += " and " + branchSearch;
                        else
                            str += " where " + branchSearch;
                        
                        String str1 = str + " and (id_stat=1 or id_stat=2)";
                        SqlCommand comm = new SqlCommand();
                        comm.CommandText = "select count(*) from Cards " + str1;
                        object obj = 0;
                        Database.ExecuteScalar(comm, ref obj, null);
                        cnt = Convert.ToInt32(obj);

                        cntFil = 0;
                        cntNoFil = 0;
                        
                        if (branch_main_filial > 0 && branch_main_filial == branch_current) 
                        {
                            str1 =str + " and (id_stat=4 and id_BranchCurrent=" + branch_current.ToString() + ")";
                            comm.CommandText = "select count(*) from Cards " + str1;
                            obj = 0;
                            Database.ExecuteScalar(comm, ref obj, null);
                            cntFil = Convert.ToInt32(obj);
                        }
                        
                    }
                }
                if (cnt > 0)
                {
                    if (sc.UserAction(User.Identity.Name, Restrictions.CardsDelete))
                        bDeleteCards.Visible = true;
                    if (branch_main_filial < 1)
                    {
                        if (sc.UserAction(User.Identity.Name, Restrictions.CardsGive))
                            bFilCards.Visible = true;
                    }
                    string str = (cnt == 1) ? "карту" : (cnt > 1 && cnt < 6) ? "карты" : "карт";
                    bDeleteCards.Attributes.Add("OnClick", String.Format("return confirm('Удалить {0} {1}');", cnt, str));
                }
                if (cntFil > 0 && cntNoFil == 0)
                {
                    if (sc.UserAction(User.Identity.Name, Restrictions.CardsGive))
                        bFilCards.Visible = true;
                }
            }
        }

        protected void bFilCards_Click(object sender, ImageClickEventArgs e)
        {
            lock (Database.lockObjectDB)
            {
                CheckBox cb, cbA;
                int stat = 0;
                SqlCommand comm = new SqlCommand();
                cbA = (CheckBox)gvCard.HeaderRow.FindControl("cbSelectAll");
                if (cbA.Checked)
                {
                    string str = lbSearch.Text.Trim();
                    if (str.Length > 0)
                        str += " and " + branchSearch;
                    else
                        str += " where " + branchSearch;
                    if (branch_main_filial <1)
                    {
                        str += " and (id_stat=1 or id_stat=2)";
                    }
                    else
                    if (branch_main_filial > 0 && branch_main_filial == branch_current) // по желанию Рустема для Казанского филиала
                    {
                        str += " and (id_stat=4 and id_BranchCurrent=" + branch_current.ToString() + ")";
                    }
                    else return; // в других филиалах запрещено
                    comm.CommandText = "update Cards set id_branchCard=@id_branch, id_prop=1 " + str;
                    comm.Parameters.Add("@id_branch", SqlDbType.Int).Value = Convert.ToInt32(Request.Form["resd"].ToString());
                    Database.ExecuteNonQuery(comm, null);
                }
                else
                {
                    comm.Parameters.Add("@id", SqlDbType.Int);
                    comm.Parameters.Add("@id_branch", SqlDbType.Int);
                    foreach (GridViewRow gvr in gvCard.Rows)
                    {
                        cb = (CheckBox)gvr.FindControl("cbSelect");
                        stat = (int)gvCard.DataKeys[gvr.RowIndex].Values["id_stat"];
                        if (cb.Checked && (stat == 1 || stat == 2 || 
                            (stat==4 && branch_main_filial>0 && branch_main_filial==branch_current))) // по желанию Рустема для Казанского филиала
                        {
                            comm.Parameters["@id"].Value = (int)gvCard.DataKeys[gvr.RowIndex].Values["id"];
                            comm.Parameters["@id_branch"].Value = Convert.ToInt32(Request.Form["resd"].ToString());
                            comm.CommandText = "update Cards set id_BranchCard=@id_branch, id_prop=1 where id=@id";
                            Database.ExecuteNonQuery(comm, null);
                        }
                    }
                }
                Refr(false);
            }
        }

        public DataTable GetCards(int StartRowIndex, int MaximumRows, String SortExpression)
        {
            try
            {
                lock (Database.lockObjectDB)
                {
                    ds.Clear();
                    string str = lbSearch.Text.Trim();
                    if (str.Length > 0)
                        str += " and " + branchSearch;
                    else
                        str += " where " + branchSearch;
                    if (sc.UserAction(User.Identity.Name, Restrictions.Transport))
                    {
                        string[] strs = ConfigurationSettings.AppSettings["Transport"].Split(',');
                        for (int t = 0; t < strs.Length; t++)
                            strs[t] = String.Format("prefix_ow like '{0}'", strs[t]);
                        str += String.Format(" and ({0})", String.Join(" or ", strs));
                    }
                    string sqlstr=String.Format("WITH CardsDetails AS (SELECT id, ROW_NUMBER() OVER ("+ lbSort.Text + ") AS RowNo FROM V_Cards " + str+ ") " +
                                                "select * from CardsDetails " +
                                                "inner join  V_Cards ON CardsDetails.id = V_Cards.id " +
                                                "{0} and (RowNo BETWEEN {2} AND {3}) {1}", str, lbSort.Text, StartRowIndex + 1, StartRowIndex + 1 + MaximumRows);
                    res = Database.ExecuteQuery(sqlstr, ref ds, null);
                    if (ds != null && ds.Tables.Count > 0)
                    {
                        for (int i = 0; i < ds.Tables[0].Rows.Count; i++)
                        {
                            if (ds.Tables[0].Rows[i]["id_branchcurrent"] == DBNull.Value)
                            {
                                ds.Tables[0].Rows[i]["id_branchcurrent"] = 0;
                            }
                        }
                    }
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
            try
            {
                lock (Database.lockObjectDB)
                {
                    string str = lbSearch.Text.Trim();
                    if (str.Length > 0)
                        str += " and " + branchSearch;
                    else
                        str += " where " + branchSearch;
                    if (sc.UserAction(User.Identity.Name, Restrictions.Transport))
                    {
                        string[] strs = ConfigurationSettings.AppSettings["Transport"].Split(',');
                        for (int t = 0; t < strs.Length; t++)
                            strs[t] = String.Format("prefix_ow like '{0}'", strs[t]);
                        str += String.Format(" and ({0})", String.Join(" or ", strs));
                    }
                    string sql=String.Format("select Count(id) from V_Cards {0} {1}", str,"");
                    SqlCommand comm = new SqlCommand();
                    comm.CommandText = sql;
                    object obj = 0;
                    Database.ExecuteScalar(comm, ref obj, null);
                    int cnt = Convert.ToInt32(obj);
                    
                    bExcel.Enabled = cnt>0 && cnt <= 100;
                    
                    bExcelXML.Enabled = cnt > 0 && cnt <= 30000;

                    bResetFilter.Visible = (lbSearch.Text != "");
                    SetButtonDoc();
                    if(cnt > 0) lbCount.Text = "Кол-во: " + cnt;
                    else lbCount.Text = "";
                    return cnt;
                 }
            }
            catch (Exception ex)
            {
            }
            return 0;
        }

        protected void gvCard_OnRowDataBound(object sender, GridViewRowEventArgs e)
        {
            if (e.Row.RowType == DataControlRowType.DataRow)
            {
                e.Row.Attributes["onclick"] = Page.ClientScript.GetPostBackClientHyperlink(gvCard, "Select$" + e.Row.RowIndex);

            }
        }

        protected void gvCard_OnSelectedIndexChanged(object sender, EventArgs e)
        {
            foreach (GridViewRow row in gvCard.Rows)
            {
                if (row.RowIndex == gvCard.SelectedIndex)
                    row.BackColor = System.Drawing.ColorTranslator.FromHtml("#A1DCF2");
                else
                    row.BackColor = System.Drawing.ColorTranslator.FromHtml("#FFFFFF");
            }
        }

        protected void bHistory_OnClick(object sender, ImageClickEventArgs e)
        {
            if (gvCard.SelectedRow == null)
                return;
            lCardHistory.Text = gvCard.SelectedRow.Cells[GetColumnIndexByName(gvCard.SelectedRow, "pan")].Text;
            using (SqlConnection conn = new SqlConnection(FuncClass.ConnectionString))
            {
                conn.Open();
                using (SqlCommand comm = conn.CreateCommand())
                {
                    comm.CommandText = $@"select d.number_doc, d.date_doc, tsd.name, b.department from Cards c inner join Cards_StorageDocs csd on c.id=csd.id_card 
                        inner join StorageDocs d on csd.id_doc = d.id inner join TypeStorageDocs tsd on d.type = tsd.id
                        left join Branchs b on d.id_branch = b.id
                        where c.id = {((HiddenField)gvCard.SelectedRow.FindControl("cardid")).Value} order by d.number_doc";
                    SqlDataAdapter da = new SqlDataAdapter(comm);
                    DataTable dt = new DataTable();
                    da.Fill(dt);
                    dgCardHistory.DataSource = dt;
                    dgCardHistory.DataBind();
                }
                conn.Close();
            }
            ClientScript.RegisterClientScriptBlock(GetType(), "bHistory_Click", "<script type='text/javascript'>$(document).ready(function(){ clickHistory();});</script>");
        }
        int GetColumnIndexByName(GridViewRow row, string columnName)
        {
            int columnIndex = 0;
            int foundIndex = -1;
            foreach (DataControlFieldCell cell in row.Cells)
            {
                if (cell.ContainingField is BoundField)
                {
                    if (((BoundField)cell.ContainingField).DataField.Equals(columnName))
                    {
                        foundIndex = columnIndex;
                        break;
                    }
                }
                columnIndex++; // keep adding 1 while we don't have the correct name
            }
            return foundIndex;
        }
    }

    public static class Utilities
    {
        public static string Bin2AHex(byte[] bytes)
        {
            if (bytes == null)
                return "";
            string str = "";
            foreach (byte b in bytes)
                str = String.Format("{0}{1:X2}", str, b);
            return str;
        }
        public static byte[] AHex2Bin(string str)
        {
            str = str.ToUpper();
            byte[] res = new byte[str.Length / 2];
            int i = 0, c1 = 0, c2 = 0;
            while (i < str.Length)
            {
                c1 = (Char.IsDigit(str, i)) ? str[i] - '0' : str[i] - 'A' + 10;
                if (i + 1 < str.Length)
                    c2 = (Char.IsDigit(str, i + 1)) ? str[i + 1] - '0' : str[i + 1] - 'A' + 10;
                else
                {
                    c2 = c1;
                    c1 = 0;
                }
                res[i / 2] = (byte)(c1 * 16 + c2);
                i += 2;
            }
            return res;
        }
    }
}
