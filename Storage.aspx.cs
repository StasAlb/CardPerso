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
using System.Data.SqlClient;
using OstCard.Data;
using CardPerso.Administration;

namespace CardPerso
{
    public partial class Storage : System.Web.UI.Page
    {
        private string res = "";
        private DataSet ds = new DataSet();
        ServiceClass sc = new ServiceClass();
        int branch_main_filial = 0;
        int current_branch_id = 0;
        int branchIdMain = 0;

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


            if (!sc.UserAction(User.Identity.Name, Restrictions.StorageView) && !sc.UserAction(User.Identity.Name, Restrictions.StorageFilial))
                Response.Redirect("~\\Account\\Restricted.aspx", true);

            ClientScript.RegisterHiddenField("resd", "");
          

            current_branch_id = sc.BranchId(User.Identity.Name);
            branch_main_filial = BranchStore.getBranchMainFilial(current_branch_id, sc.UserAction(User.Identity.Name, Restrictions.Perso));
            //showChild.Visible = current_branch_id==106;
            branchIdMain = BranchStore.getBranchMainFilial(current_branch_id, false);
            showChild.Visible = current_branch_id == branchIdMain;

            lock (Database.lockObjectDB)
            {
                if (!IsPostBack)
                {
                    lbSort.Text = "order by id_type,id_sort";
                    Refr(true);
                }
            }
            GC.Collect();
        }

        private void Refr(bool conf)
        {
            ds.Reset();
            ds.Clear();

            //bool isMainFilial = false;
            //if (branch_main_filial > 0 && branch_main_filial == current_branch_id) isMainFilial = true;// Казанский

            res = Database.ExecuteQuery("select id_prb, id_bank, id_prod, id_sort, prod_name as name, type_prod, bank_name, bin, prefix_ow, prefix_file, min_cnt, id_type, parent from V_ProductsBanks_T where parent is NULL order by id_sort,prod_name desc", ref ds, null);

            if (ds.Tables == null || ds.Tables.Count == 0)
                return;

            SqlCommand comm = new SqlCommand();
            comm.Connection = Database.Conn;
            comm.CommandText = "select id_prb, id_bank, id_prod, id_sort, prod_name as name, type_prod, bank_name, bin, prefix_ow, prefix_file, min_cnt, id_type, parent from V_ProductsBanks_T where parent=@parent order by id_sort,prod_name desc";
            comm.Parameters.Add("@parent", SqlDbType.Int);

            int t = 0;
            for (int i = 0; i < ds.Tables[0].Rows.Count; i++)
            {
                comm.Parameters["@parent"].Value = ds.Tables[0].Rows[i]["id_prod"];
                lock (Database.lockObjectDB)
                {
                    SqlDataReader sdr = comm.ExecuteReader();
                    int cnt = 0;
                    while (sdr.Read())
                    {
                        ds.Tables[0].Rows[i]["parent"] = -1;
                        DataRow ddr = ds.Tables[0].NewRow();
                        for (t = 0; t < sdr.FieldCount; t++)
                            ddr[ds.Tables[0].Columns[t].ColumnName] = sdr[ds.Tables[0].Columns[t].ColumnName];
                        ddr["name"] = "&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;" + ddr["name"];
                        ds.Tables[0].Rows.InsertAt(ddr, i + cnt + 1);
                        cnt++;
                    }

                    i += cnt;
                    sdr.Close();
                }
            }
            ds.Tables[0].Columns.Add("id", System.Type.GetType("System.Int32"));
            ds.Tables[0].Columns.Add("cnt_new", System.Type.GetType("System.Int32"));
            ds.Tables[0].Columns.Add("cnt_wrk", System.Type.GetType("System.Int32"));
            ds.Tables[0].Columns.Add("cnt_perso", System.Type.GetType("System.String"));
            ds.Tables[0].Columns.Add("cnt_brak", System.Type.GetType("System.String"));
            ds.Tables[0].Columns.Add("cnt_del", System.Type.GetType("System.String"));
            ds.Tables[0].Columns.Add("cnt_expire", System.Type.GetType("System.Int32"));
            ds.Tables[0].Columns.Add("cnt_notaskedfor", System.Type.GetType("System.Int32"));
            ds.Tables[0].Columns.Add("image", System.Type.GetType("System.String"));
            ds.Tables[0].Columns.Add("childs", System.Type.GetType("System.Int32"));

            DataRow dr = ds.Tables[0].NewRow();
            dr["Name"] = "Итого по картам:";
            dr["cnt_new"] = 0;
            dr["cnt_wrk"] = 0;
            dr["cnt_perso"] = 0;
            dr["cnt_brak"] = 0;
            dr["cnt_expire"] = 0;
            dr["cnt_notaskedfor"] = 0;
            dr["image"] = "";
            dr["childs"] = 0;
            
            /*
            DataRow drpin = null;
            if(isMainFilial)
            {
                drpin = ds.Tables[0].NewRow();
                drpin["Name"] = "Итого по пин-конвертам:";
                drpin["cnt_new"] = 0;
                drpin["cnt_wrk"] = 0;
                drpin["cnt_perso"] = 0;
                drpin["cnt_brak"] = 0;
                drpin["cnt_expire"] = 0;
                drpin["cnt_notaskedfor"] = 0;
                drpin["image"] = "";
                drpin["childs"] = 0;
            }
            */
            

            //if (sc.UserAction(User.Identity.Name, Restrictions.Perso))
            if (User.Identity.Name == superusername || sc.UserAction(User.Identity.Name, Restrictions.StorageView))
            {
                //ds.Tables[0].Columns.Add("id", System.Type.GetType("System.Int32"));
                //ds.Tables[0].Columns.Add("cnt_new", System.Type.GetType("System.Int32"));
                //ds.Tables[0].Columns.Add("cnt_wrk", System.Type.GetType("System.Int32"));
                //ds.Tables[0].Columns.Add("cnt_perso", System.Type.GetType("System.String"));
                //ds.Tables[0].Columns.Add("cnt_brak", System.Type.GetType("System.String"));
                //ds.Tables[0].Columns.Add("cnt_del", System.Type.GetType("System.String"));
                //ds.Tables[0].Columns.Add("cnt_expire", System.Type.GetType("System.Int32"));
                //ds.Tables[0].Columns.Add("cnt_notaskedfor", System.Type.GetType("System.Int32"));
                //ds.Tables[0].Columns.Add("image", System.Type.GetType("System.String"));

                SqlCommand sel = new SqlCommand();
                sel.Connection = Database.Conn;
                sel.CommandText = String.Format("select * from V_Storage where id_prod=@idProd {0} {1}", lbSearch.Text.Replace("where", "and"), lbSort.Text);
                sel.Parameters.Add("@idProd", SqlDbType.Int);

                comm = new SqlCommand();
                //comm.CommandText = "select count(*) from Cards where id_prop=@prop and id_stat=6 and id_prb=@prb and id_branchcurrent not in (select id from branchs where id=106 or id_parent=106)";
                
                comm.CommandText = String.Format("select count(*) from Cards where id_prop=@prop and id_stat=6 and id_prb=@prb and id_branchcurrent not in (select id from branchs where id={0} or id_parent={0})", branchIdMain);
                
                comm.Parameters.Add("@prop", SqlDbType.Int);
                comm.Parameters.Add("@prb", SqlDbType.Int);

                object obj = 0;
                t = 0;

                for (int i = 0; i < ds.Tables[0].Rows.Count; i++)
                {
                    ds.Tables[0].Rows[i]["cnt_new"] = 0;
                    ds.Tables[0].Rows[i]["cnt_wrk"] = 0;
                    ds.Tables[0].Rows[i]["cnt_perso"] = 0;
                    ds.Tables[0].Rows[i]["cnt_brak"] = 0;
                    ds.Tables[0].Rows[i]["cnt_del"] = 0;
                    ds.Tables[0].Rows[i]["cnt_notaskedfor"] = 0;
                    ds.Tables[0].Rows[i]["childs"] = 0;

                    ds.Tables[0].Rows[i]["image"] = "";
                    sel.Parameters["@idProd"].Value = Convert.ToInt32(ds.Tables[0].Rows[i]["id_prod"]);
                    lock (Database.lockObjectDB)
                    {
                        SqlDataReader sdr = sel.ExecuteReader();
                        if (sdr.Read())
                        {
                            ds.Tables[0].Rows[i]["id"] = sdr["id"];
                            ds.Tables[0].Rows[i]["cnt_new"] = sdr["cnt_new"];
                            ds.Tables[0].Rows[i]["cnt_wrk"] = sdr["cnt_wrk"];
                            ds.Tables[0].Rows[i]["cnt_perso"] = sdr["cnt_perso"];
                            ds.Tables[0].Rows[i]["cnt_brak"] = sdr["cnt_brak"];
                            ds.Tables[0].Rows[i]["cnt_del"] = sdr["cnt_del"];
                        }

                        sdr.Close();
                    }

                    if (ds.Tables[0].Rows[i]["id_type"] != DBNull.Value && Convert.ToInt32(ds.Tables[0].Rows[i]["id_type"]) == 1)
                    {

                        dr["cnt_new"] = Convert.ToInt32(dr["cnt_new"]) + Convert.ToInt32(ds.Tables[0].Rows[i]["cnt_new"]);
                        dr["cnt_perso"] = Convert.ToInt32(dr["cnt_perso"]) + Convert.ToInt32(ds.Tables[0].Rows[i]["cnt_perso"]);
                        dr["cnt_brak"] = Convert.ToInt32(dr["cnt_brak"]) + Convert.ToInt32(ds.Tables[0].Rows[i]["cnt_brak"]);

                        obj = 0;
                        // истек срок годности
                        comm.Parameters["@prop"].Value = 8;
                        comm.Parameters["@prb"].Value = (int)ds.Tables[0].Rows[i]["id_prb"];
                        Database.ExecuteScalar(comm, ref obj, null);
                        ds.Tables[0].Rows[i]["cnt_expire"] = Convert.ToInt32(obj);
                        ds.Tables[0].Rows[i]["cnt_brak"] = Convert.ToInt32(ds.Tables[0].Rows[i]["cnt_brak"]) - Convert.ToInt32(obj);
                        dr["cnt_brak"] = Convert.ToInt32(dr["cnt_brak"]) - Convert.ToInt32(obj);
                        dr["cnt_expire"] = Convert.ToInt32(dr["cnt_expire"]) + Convert.ToInt32(obj);
                        // не востребовано
                        obj = 0;
                        comm.Parameters["@prop"].Value = 7;
                        comm.Parameters["@prb"].Value = Convert.ToInt32(ds.Tables[0].Rows[i]["id_prb"]);
                        Database.ExecuteScalar(comm, ref obj, null);
                        ds.Tables[0].Rows[i]["cnt_notaskedfor"] = Convert.ToInt32(obj);
                        ds.Tables[0].Rows[i]["cnt_brak"] = Convert.ToInt32(ds.Tables[0].Rows[i]["cnt_brak"]) - Convert.ToInt32(obj);
                        dr["cnt_brak"] = Convert.ToInt32(dr["cnt_brak"]) - Convert.ToInt32(obj);
                        dr["cnt_notaskedfor"] = Convert.ToInt32(dr["cnt_notaskedfor"]) + Convert.ToInt32(obj);
                        t = i;
                        string fname = String.Format("{0}{1}_{2}.jpg", ConfigurationSettings.AppSettings["ImagePath"].ToString(), ds.Tables[0].Rows[i]["bin"], ds.Tables[0].Rows[i]["prefix_ow"]);
                        if (System.IO.File.Exists(fname))
                            //ds.Tables[0].Rows[i]["image"] = String.Format("images/{0}_{1}.jpg", ds.Tables[0].Rows[i]["bin"], ds.Tables[0].Rows[i]["prefix_ow"]);
                            ds.Tables[0].Rows[i]["image"] = String.Format("{0}_{1}.jpg", ds.Tables[0].Rows[i]["bin"], ds.Tables[0].Rows[i]["prefix_ow"]);
                        //                    ds.Tables[0].Rows[i]["name"] += " <IMG ALT='' SRC='images/del.bmp' onmouseover='tooltip1(this, '<b>message</b>', 'hint')'/>";

                    }
                }
            }
            if (User.Identity.Name != superusername && sc.UserAction(User.Identity.Name, Restrictions.StorageFilial))
            {
                SqlCommand sel = new SqlCommand();
                sel.Connection = Database.Conn;
                //sel.CommandText = "select count(*) from Cards where id_prb=@idProd and id_stat=4 and id_branchCurrent=@branch";
                /*
                if (branch_main_filial > 0) // Казанский или его подчиненные
                {
                    if (branch_main_filial > 0 && branch_main_filial == current_branch_id) // Казанский
                    {
                        if (showChild.Checked == true)
                            sel.CommandText = "select count(*),case when sum(cast(ispin as integer)) is null then 0 else sum(cast(ispin as integer)) end from Cards where id_prb=@idProd and id_stat=4 and (id_branchCard=@branch or id_branchCard in (select id from Branchs where id_parent=@branch or id_parentTr=@branch))";
                        else
                            sel.CommandText = "select count(*),case when sum(cast(ispin as integer)) is null then 0 else sum(cast(ispin as integer)) end from Cards where id_prb=@idProd and id_stat=4 and id_branchCurrent=@branch";
                        //sel.CommandText = "select count(*) from Cards where id_prb=@idProd and id_stat=4 and id_branchCard=@branch";
                    }
                    else // Подчиненные
                    {
                        sel.CommandText = "select count(*),case when sum(cast(ispin as integer)) is null then 0 else sum(cast(ispin as integer)) end from Cards where id_prb=@idProd and id_stat=4 and id_branchCurrent=@branch";
                    }
                }
                else
                {
                    // Все остальные
                    sel.CommandText = "select count(*),case when sum(cast(ispin as integer)) is null then 0 else sum(cast(ispin as integer)) end from Cards where id_prb=@idProd and id_stat=4 and (id_branchCard=@branch or id_branchCard in (select id from Branchs where id_parent=@branch or id_parentTr=@branch))";
                }
                sel.Parameters.Add("@idProd", SqlDbType.Int);
                sel.Parameters.Add("@branch", SqlDbType.Int).Value = sc.BranchId(User.Identity.Name);
                int indx_pin = -1;
                int sum_pin = 0;
                for (int i = 0; i < ds.Tables[0].Rows.Count; i++)
                {
                    ds.Tables[0].Rows[i]["cnt_new"] = 0;
                    ds.Tables[0].Rows[i]["cnt_wrk"] = 0;
                    ds.Tables[0].Rows[i]["cnt_perso"] = 0;
                    ds.Tables[0].Rows[i]["cnt_brak"] = 0;
                    ds.Tables[0].Rows[i]["cnt_del"] = 0;
                    ds.Tables[0].Rows[i]["cnt_notaskedfor"] = 0;
                    ds.Tables[0].Rows[i]["childs"] = 0;

                    ds.Tables[0].Rows[i]["image"] = "";
                    sel.Parameters["@idProd"].Value = Convert.ToInt32(ds.Tables[0].Rows[i]["id_prb"]);
                    SqlDataReader sdr = sel.ExecuteReader();
                    if (sdr.Read())
                    {
                        ds.Tables[0].Rows[i]["cnt_perso"] = Convert.ToInt32(sdr[0]);
                        if (ds.Tables[0].Rows[i]["id_type"] != DBNull.Value && Convert.ToInt32(ds.Tables[0].Rows[i]["id_type"]) == 1)
                        {
                            dr["cnt_perso"] = Convert.ToInt32(dr["cnt_perso"]) + Convert.ToInt32(ds.Tables[0].Rows[i]["cnt_perso"]);
                            sum_pin += Convert.ToInt32(sdr[1]);
                        }
                        if (Convert.ToInt32(ds.Tables[0].Rows[i]["id_type"]) == 2)  indx_pin = i;
                    }
                    sdr.Close();
                    string fname = String.Format("{0}{1}_{2}.jpg", ConfigurationSettings.AppSettings["ImagePath"].ToString(), ds.Tables[0].Rows[i]["bin"], ds.Tables[0].Rows[i]["prefix_ow"]);
                    if (System.IO.File.Exists(fname))
                        ds.Tables[0].Rows[i]["image"] = String.Format("{0}_{1}.jpg", ds.Tables[0].Rows[i]["bin"], ds.Tables[0].Rows[i]["prefix_ow"]);
                }
                if (indx_pin >= 0)
                {
                    ds.Tables[0].Rows[indx_pin]["cnt_perso"] = sum_pin;
                }
            }
            */
            
            if (branch_main_filial > 0) // Казанский или его подчиненные
            {
                if (branch_main_filial > 0 && branch_main_filial == current_branch_id) // Казанский
                {
                    if (showChild.Checked == true)
                        sel.CommandText = "select id_prb,id_stat,count(*) as ccard,case when sum(cast(ispin as integer)) is null then 0 else sum(cast(ispin as integer)) end as cpin from Cards where id_stat in (4,6) and (id_branchCard=@branch or id_branchCard in (select id from Branchs where id_parent=@branch or id_parentTr=@branch)) group by id_prb,id_stat";
                    else
                        sel.CommandText = "select id_prb,id_stat,count(*) as ccard,case when sum(cast(ispin as integer)) is null then 0 else sum(cast(ispin as integer)) end as cpin from Cards where id_stat in (4,6) and id_branchCurrent=@branch group by id_prb,id_stat";
                    //sel.CommandText = "select count(*) from Cards where id_prb=@idProd and id_stat=4 and id_branchCard=@branch";
                }
                else // Подчиненные
                {
                    sel.CommandText = "select id_prb,id_stat,count(*) as ccard,case when sum(cast(ispin as integer)) is null then 0 else sum(cast(ispin as integer)) end as cpin from Cards where id_stat in (4,6) and id_branchCurrent=@branch group by id_prb,id_stat";
                }
            }
            else
            {
                // Все остальные
                sel.CommandText = "select id_prb,id_stat,count(*) as ccard,case when sum(cast(ispin as integer)) is null then 0 else sum(cast(ispin as integer)) end as cpin from Cards where id_stat in (4,6) and (id_branchCard=@branch or id_branchCard in (select id from Branchs where id_parent=@branch or id_parentTr=@branch)) group by id_prb,id_stat";
            }
            sel.Parameters.Add("@branch", SqlDbType.Int).Value = sc.BranchId(User.Identity.Name);
            int indx_pin_p = -1, indx_pin_b = -1;
            int sum_pin_p = 0, sum_pin_b = 0;

            DataSet dsCount = new DataSet();
            Database.ExecuteCommand(sel, ref dsCount, null);
            for (int i = 0; i < ds.Tables[0].Rows.Count; i++)
            {
                ds.Tables[0].Rows[i]["cnt_new"] = 0;
                ds.Tables[0].Rows[i]["cnt_wrk"] = 0;
                ds.Tables[0].Rows[i]["cnt_perso"] = 0;
                ds.Tables[0].Rows[i]["cnt_brak"] = 0;
                ds.Tables[0].Rows[i]["cnt_del"] = 0;
                ds.Tables[0].Rows[i]["cnt_notaskedfor"] = 0;
                ds.Tables[0].Rows[i]["childs"] = 0;

                ds.Tables[0].Rows[i]["image"] = "";

                
                for (int j = 0; dsCount.Tables.Count > 0 && j < dsCount.Tables[0].Rows.Count; j++)
                {
                    if (Convert.ToInt32(ds.Tables[0].Rows[i]["id_prb"])==Convert.ToInt32(dsCount.Tables[0].Rows[j]["id_prb"]))
                    {

                        String fieldName="cnt_perso";
                        int tp = 0;
                        if (Convert.ToInt32(dsCount.Tables[0].Rows[j]["id_stat"]) == 6)
                        {
                            fieldName = "cnt_brak";
                            tp = 1;
                        }
                        ds.Tables[0].Rows[i][fieldName] = Convert.ToInt32(dsCount.Tables[0].Rows[j]["ccard"]);
                        if (ds.Tables[0].Rows[i]["id_type"] != DBNull.Value && Convert.ToInt32(ds.Tables[0].Rows[i]["id_type"]) == 1)
                        {
                            dr[fieldName] = Convert.ToInt32(dr[fieldName]) + Convert.ToInt32(ds.Tables[0].Rows[i][fieldName]);
                            if(tp==0) sum_pin_p += Convert.ToInt32(dsCount.Tables[0].Rows[j]["cpin"]);
                            else sum_pin_b += Convert.ToInt32(dsCount.Tables[0].Rows[j]["cpin"]);
                        }
                        //break;
                    }
                    if (Convert.ToInt32(ds.Tables[0].Rows[i]["id_type"]) == 2)
                    {
                        if (Convert.ToInt32(dsCount.Tables[0].Rows[j]["id_stat"]) == 4) indx_pin_p = i;
                        else indx_pin_b = i;
                    }

                   
                }
                string fname = String.Format("{0}{1}_{2}.jpg", ConfigurationSettings.AppSettings["ImagePath"].ToString(), ds.Tables[0].Rows[i]["bin"], ds.Tables[0].Rows[i]["prefix_ow"]);
                if (System.IO.File.Exists(fname))
                    ds.Tables[0].Rows[i]["image"] = String.Format("{0}_{1}.jpg", ds.Tables[0].Rows[i]["bin"], ds.Tables[0].Rows[i]["prefix_ow"]);
            }
            if (indx_pin_p >= 0)
            {
                ds.Tables[0].Rows[indx_pin_p]["cnt_perso"] = sum_pin_p;
            }

            if (indx_pin_b >= 0)
            {
                ds.Tables[0].Rows[indx_pin_b]["cnt_brak"] = sum_pin_b;
            }
            
            }
            
//!!!!!!!!!!!!!!!!!!!!!!!!!!!
            t = 0;
            int parid = -1, parind = -1;
            for (int i = 0; i < ds.Tables[0].Rows.Count; i++)
            {
                if (Convert.ToInt32(ds.Tables[0].Rows[i]["id_type"]) == 1 && (ds.Tables[0].Rows[i]["Parent"] == DBNull.Value || Convert.ToInt32(ds.Tables[0].Rows[i]["Parent"]) < 0))
                {
                    parid = Convert.ToInt32(ds.Tables[0].Rows[i]["id_prod"]);
                    parind = i;
                }                
                int allcnt = Convert.ToInt32(ds.Tables[0].Rows[i]["cnt_new"]) + Convert.ToInt32(ds.Tables[0].Rows[i]["cnt_wrk"]) +
                    Convert.ToInt32(ds.Tables[0].Rows[i]["cnt_perso"]) + Convert.ToInt32(ds.Tables[0].Rows[i]["cnt_brak"]) +
                    Convert.ToInt32(ds.Tables[0].Rows[i]["cnt_del"]) + Convert.ToInt32(ds.Tables[0].Rows[i]["cnt_notaskedfor"]);
                if (allcnt > 0)
                {
                    ds.Tables[0].Rows[i]["childs"] = allcnt;
                    if (parind > 0 && ds.Tables[0].Rows[i]["parent"] != DBNull.Value && Convert.ToInt32(ds.Tables[0].Rows[i]["parent"]) == parid)
                        ds.Tables[0].Rows[parind]["childs"] = allcnt;
                }
            }
            for (int i = 0; i < ds.Tables[0].Rows.Count; i++)
            {
                //if (Convert.ToInt32(ds.Tables[0].Rows[i]["id_type"]) != 1 || Convert.ToInt32(ds.Tables[0].Rows[i]["childs"]) == 0)
                if (Convert.ToInt32(ds.Tables[0].Rows[i]["childs"]) == 0)
                {
                    ds.Tables[0].Rows.RemoveAt(i);
                    i--;
                }
            }
            /*
            for (int i = 0; i < ds.Tables[0].Rows.Count; i++)
            {
                if (ds.Tables[0].Rows[i]["id_type"] != DBNull.Value && Convert.ToInt32(ds.Tables[0].Rows[i]["id_type"]) == 1)
                    t = i;
            }
            if (t < ds.Tables[0].Rows.Count + 1)
                ds.Tables[0].Rows.InsertAt(dr, t + 1);
            
            else
            */ 
                ds.Tables[0].Rows.Add(dr);

            //if (drpin != null) ds.Tables[0].Rows.Add(drpin);
            
            gvStorage.DataSource = ds.Tables[0];
            gvStorage.DataBind();

            if (User.Identity.Name != superusername && sc.UserAction(User.Identity.Name, Restrictions.StorageFilial))
            {
                gvStorage.Columns[3].Visible = false;
                gvStorage.Columns[4].Visible = false;
                //gvStorage.Columns[6].Visible = false;
                gvStorage.Columns[7].Visible = false;
                gvStorage.Columns[8].Visible = false;
            }

            if (conf)
            {
                string s_fld = Database.GetFiledsByUser(sc.UserId(User.Identity.Name), "storage", null);
                if (s_fld!="") FuncClass.HideFields(s_fld,gvStorage);
            }

            bResetFilter.Visible = (lbSearch.Text != "");
            lbCount.Text = "Кол-во: " + gvStorage.Rows.Count.ToString();
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
                    ep.ExportGridExcel(gvStorage);
                    if (WebConfigurationManager.AppSettings["DocPath"] != null)
                    {
                        doc = String.Format("{0}Temp\\storage.xls", WebConfigurationManager.AppSettings["DocPath"]);
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

        protected void gvStorage_Sorting(object sender, GridViewSortEventArgs e)
        {
            lock (Database.lockObjectDB)
            {
                int i = FuncClass.GetSortIndex(e.SortExpression, gvStorage);

                if (lbSortIndex.Text != "")
                {
                    int ind = Convert.ToInt32(lbSortIndex.Text);
                    gvStorage.Columns[ind].HeaderText = gvStorage.Columns[ind].HeaderText.Replace("^", "");
                    gvStorage.Columns[ind].HeaderStyle.BackColor = System.Drawing.Color.FromArgb(0, 102, 153);
                }
                lbSortIndex.Text = i.ToString();
                gvStorage.Columns[i].HeaderStyle.BackColor = System.Drawing.Color.FromArgb(102, 153, 153);
                if ("order by " + e.SortExpression + " asc" == lbSort.Text)
                {
                    lbSort.Text = "order by " + e.SortExpression + " desc";
                    gvStorage.Columns[i].HeaderText = gvStorage.Columns[i].HeaderText + "^";
                }
                else
                    lbSort.Text = "order by " + e.SortExpression + " asc";

                Refr(false);
            }
        }

        protected void bConfField_Click(object sender, ImageClickEventArgs e)
        {
            Response.Redirect("~/Storage.aspx");
        }

        protected void showChild_Changed(Object sender, EventArgs e)
        {
            lock (Database.lockObjectDB)
            {
                //string ctrlName = ((Control)sender).ID;
                Refr(false);
            }
        }
    }
}
