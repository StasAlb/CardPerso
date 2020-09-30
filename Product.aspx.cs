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

namespace CardPerso
{
    public partial class Product : System.Web.UI.Page
    {
        string res = "";
        ServiceClass sc = new ServiceClass();
        DataSet ds = new DataSet();
        protected void Page_Load(object sender, EventArgs e)
        {
            lock (Database.lockObjectDB)
            {
                ClientScript.RegisterHiddenField("resd", "");
                if (!IsPostBack)
                {
                    lbInform.Text = "";
                    Refr(0);
                }
            }
        }

        private void Refr(int rowindex)
        {
            Refr(rowindex, -1);
        }
        private void Refr(int rowindex, int prod_id)
        {
            lbInform.Text = "";
            ds.Reset();
            ds.Clear();

            res = Database.ExecuteQuery("select id_prb, id_bank, id_prod, id_sort, prod_name, prod_name_h, type_prod, bank_name, bin, prefix_ow, prefix_file, min_cnt, parent from V_ProductsBanks where parent is NULL order by id_sort,prod_name desc", ref ds, null);
            SqlCommand comm = new SqlCommand();
            comm.Connection = Database.Conn;
            comm.CommandText = "select id_prb, id_bank, id_prod, id_sort, prod_name, prod_name_h, type_prod, bank_name, bin, prefix_ow, prefix_file, min_cnt, parent from V_ProductsBanks where parent=@parent order by id_sort,prod_name desc";
            comm.Parameters.Add("@parent", SqlDbType.Int);
            for (int i = 0; i < ds.Tables[0].Rows.Count; i++)
            {
                lock (Database.lockObjectDB)
                {
                    comm.Parameters["@parent"].Value = ds.Tables[0].Rows[i]["id_prod"]; // == DBNull.Value) Convert.ToInt32(ds.Tables[0].Rows[i]["id_prod"]);
                    SqlDataReader sdr = comm.ExecuteReader();
                    int cnt = 0;
                    while (sdr.Read())
                    {
                        DataRow dr = ds.Tables[0].NewRow();
                        for (int t = 0; t < sdr.FieldCount; t++)
                            dr[ds.Tables[0].Columns[t].ColumnName] = sdr[ds.Tables[0].Columns[t].ColumnName];
                        dr["prod_name"] = "&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;" +
                                          dr["prod_name"];
                        ds.Tables[0].Rows.InsertAt(dr, i + cnt + 1);
                        cnt++;
                    }

                    i += cnt;
                    sdr.Close();
                }
            }
            gvProducts.DataSource = ds.Tables[0];
            gvProducts.DataBind();

            if (prod_id > 0)
            {
                for (int i = 0; i < ds.Tables[0].Rows.Count; i++)
                    if (prod_id == Convert.ToInt32(ds.Tables[0].Rows[i]["id_prb"]))
                    {
                        gvProducts.SelectedIndex = i;
                        gvProducts.Rows[gvProducts.SelectedIndex].Focus();

                    }
            }
            else
            {
                if (gvProducts.Rows.Count > 0)
                {
                    gvProducts.SelectedIndex = rowindex;
                    gvProducts.Rows[gvProducts.SelectedIndex].Focus();
                }
            }


            SetButton();

            lbCount.Text = "Кол-во: " + gvProducts.Rows.Count.ToString();

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
                    ep.ExportGridExcel(gvProducts);
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
                Refr(gvProducts.Rows.Count);
            }
        }

        protected void gvProducts_SelectedIndexChanged(object sender, EventArgs e)
        {
            lock (Database.lockObjectDB)
            {
                gvProducts.Rows[gvProducts.SelectedIndex].Focus();
                SetButton();
            }
        }

        private void SetButton()
        {
            lbInform.Text = "";
            if (!sc.UserAction(User.Identity.Name, Restrictions.LibraryEdit))
            {
                bNew.Visible = false;
                bEdit.Visible = false;
                bDelete.Visible = false;
                bExcel.Visible = false;
                bDelBank.Visible = false;
                bAddBank.Visible = false;
                bLinkProd.Visible = false;
                bUnlinkProd.Visible = false;
                bSortUp.Visible = false;
                bSortDown.Visible = false;
                return;
            }


            if (gvProducts.Rows.Count > 0)
            {
                bEdit.Visible = true;
                bDelete.Visible = true;
                bExcel.Visible = true;
                bDelBank.Visible = true;
                bAddBank.Visible = true;

                bDelete.Attributes.Add("OnClick", String.Format("return confirm('Удалить продукт {0}?');", gvProducts.DataKeys[Convert.ToInt32(gvProducts.SelectedIndex)].Values["prod_name_h"].ToString()));
                bDelBank.Attributes.Add("OnClick", String.Format("return confirm('Отвязать банк {0} от продукта {1}?');", gvProducts.DataKeys[Convert.ToInt32(gvProducts.SelectedIndex)].Values["bank_name"].ToString(),gvProducts.DataKeys[Convert.ToInt32(gvProducts.SelectedIndex)].Values["prod_name_h"].ToString()));
                bEdit.Attributes.Add("OnClick", String.Format("return show_product('mode=2&id_prb={0}&id_prod={1}')", gvProducts.DataKeys[Convert.ToInt32(gvProducts.SelectedIndex)].Values["id_prb"].ToString(), gvProducts.DataKeys[Convert.ToInt32(gvProducts.SelectedIndex)].Values["id_prod"].ToString()));
                bAddBank.Attributes.Add("OnClick", String.Format("return show_product('mode=3&id_prb=0&id_prod={0}')", gvProducts.DataKeys[Convert.ToInt32(gvProducts.SelectedIndex)].Values["id_prod"].ToString()));
                if (gvProducts.DataKeys[Convert.ToInt32(gvProducts.SelectedIndex)].Values["parent"] != DBNull.Value && Convert.ToInt32(gvProducts.DataKeys[Convert.ToInt32(gvProducts.SelectedIndex)].Values["parent"]) > 0)
                {
                    bUnlinkProd.Visible = true;
                    bLinkProd.Visible = false;
                }
                else
                {
                    bLinkProd.Visible = true;
                    bLinkProd.Attributes.Add("OnClick", String.Format("return show_product_link('id={0}')", gvProducts.DataKeys[Convert.ToInt32(gvProducts.SelectedIndex)].Values["id_prod"].ToString()));
                    bUnlinkProd.Visible = false;
                }
            }
            else
            {
                bEdit.Visible = false;
                bDelete.Visible = false;
                bExcel.Visible = false;
                bDelBank.Visible = false;
                bAddBank.Visible = false;
                bLinkProd.Visible = false;
                bUnlinkProd.Visible = false;
            }
        }

        protected void bEdit_Click(object sender, ImageClickEventArgs e)
        {
            lock (Database.lockObjectDB)
            {
                Refr(gvProducts.SelectedIndex);
            }
        }

        protected void bDelete_Click(object sender, ImageClickEventArgs e)
        {
            lock (Database.lockObjectDB)
            {
                int id = Convert.ToInt32(gvProducts.DataKeys[Convert.ToInt32(gvProducts.SelectedIndex)].Values["id_prod"]);
                int id_prb = Convert.ToInt32(gvProducts.DataKeys[Convert.ToInt32(gvProducts.SelectedIndex)].Values["id_prb"]);

                if (!Database.CheckDelProduct(id_prb, null))
                {
                    lbInform.Text = "Невозможно удалить продукцию, так как существуют связанные документы.";
                    return;
                }
                SqlCommand sqCom = null;
                try
                {
                    sqCom = new SqlCommand();
                    sqCom.CommandText = "delete from Products_Banks where id_prod=@id_prod";
                    sqCom.Parameters.Add("@id_prod", SqlDbType.Int).Value = id;
                    res = Database.ExecuteNonQuery(sqCom, null);
                    sqCom.Parameters.Clear();
                    sqCom.CommandText = "delete from Products where id=@id";
                    sqCom.Parameters.Add("@id", SqlDbType.Int).Value = id;
                    Database.ExecuteNonQuery(sqCom, null);
                }
                catch
                {
                }
                Refr(0);
            }
        }

        protected void bDelBank_Click(object sender, ImageClickEventArgs e)
        {
            lock (Database.lockObjectDB)
            {
                int id_prod = Convert.ToInt32(gvProducts.DataKeys[Convert.ToInt32(gvProducts.SelectedIndex)].Values["id_prod"]);
                if (Database.CntBanksInProduct(id_prod, null) < 2)
                {
                    lbInform.Text = "Невозможно отвязать банк";
                    return;
                }

                int id = Convert.ToInt32(gvProducts.DataKeys[Convert.ToInt32(gvProducts.SelectedIndex)].Values["id_prb"]);
                SqlCommand sqCom = new SqlCommand();
                sqCom.CommandText = "delete from Products_Banks where id=@id";
                sqCom.Parameters.Add("@id", SqlDbType.Int).Value = id;
                try
                {
                    res = Database.ExecuteNonQuery(sqCom, null);
                }
                catch
                {
                }
                Refr(0);
            }
        }

        protected void bAddBank_Click(object sender, ImageClickEventArgs e)
        {
            lock (Database.lockObjectDB)
            {
                Refr(gvProducts.SelectedIndex);
            }
        }

        protected void bSortUp_Click(object sender, ImageClickEventArgs e)
        {
            lock (Database.lockObjectDB)
            {
                UpdateSort(true);
            }
        }

        protected void bSortDown_Click(object sender, ImageClickEventArgs e)
        {
            lock (Database.lockObjectDB)
            {

                // int id2 = Database.GetSortId(Convert.ToInt32(gvProducts.DataKeys[Convert.ToInt32(gvProducts.SelectedIndex)].Values["id_sort"]), false);
                // int id2_sort = Database.GetSortIdSort(Convert.ToInt32(gvProducts.DataKeys[Convert.ToInt32(gvProducts.SelectedIndex)].Values["id_sort"]), false);
                UpdateSort(false);
            }
        }

        private void UpdateSort(bool up)
        {
            int id1 = Convert.ToInt32(gvProducts.DataKeys[Convert.ToInt32(gvProducts.SelectedIndex)].Values["id_prod"]);
            int id1_sort = Convert.ToInt32(gvProducts.DataKeys[Convert.ToInt32(gvProducts.SelectedIndex)].Values["id_sort"]);
            string par = (gvProducts.DataKeys[Convert.ToInt32(gvProducts.SelectedIndex)].Values["parent"] == DBNull.Value) ? "parent is null" : "parent = " + Convert.ToString(gvProducts.DataKeys[Convert.ToInt32(gvProducts.SelectedIndex)].Values["parent"]);

            ds.Clear();
            if (up)
                res = Database.ExecuteQuery(String.Format("select top 1 id_prod as id,id_sort from V_ProductsBanks where {1} and id_sort<{0} order by id_sort desc", id1_sort, par), ref ds, null);
            else
                res = Database.ExecuteQuery(String.Format("select top 1 id_prod as id,id_sort from V_ProductsBanks where {1} and id_sort>{0} order by id_sort", id1_sort, par), ref ds, null);

            if (ds.Tables[0].Rows.Count == 0)
                return;

            int id2 = Convert.ToInt32(ds.Tables[0].Rows[0]["id"]);
            int id2_sort = Convert.ToInt32(ds.Tables[0].Rows[0]["id_sort"]);

            SqlCommand sqCom = new SqlCommand();
            
            sqCom.CommandText = "update Products set id_sort=@id_sort where id=@id";
            sqCom.Parameters.Add("@id", SqlDbType.Int).Value = id2;
            sqCom.Parameters.Add("@id_sort", SqlDbType.Int).Value = id1_sort;
            res = Database.ExecuteNonQuery(sqCom, null);

            sqCom.Parameters.Clear();
            sqCom.CommandText = "update Products set id_sort=@id_sort where id=@id";
            sqCom.Parameters.Add("@id", SqlDbType.Int).Value = id1;
            sqCom.Parameters.Add("@id_sort", SqlDbType.Int).Value = id2_sort;
            res = Database.ExecuteNonQuery(sqCom, null);

            int id=Convert.ToInt32(gvProducts.DataKeys[Convert.ToInt32(gvProducts.SelectedIndex)].Values["id_prb"]);
            ds.Clear();
            Refr(0, id);
        }
        protected void bLinkProd_Click(object sender, ImageClickEventArgs e)        
        {
            int id = Convert.ToInt32(gvProducts.DataKeys[Convert.ToInt32(gvProducts.SelectedIndex)].Values["id_prb"]);
            Refr(0, id);
        }
        protected void bUnlinkProd_Click(object sender, ImageClickEventArgs e)
        {
            int id = Convert.ToInt32(gvProducts.DataKeys[Convert.ToInt32(gvProducts.SelectedIndex)].Values["id_prb"]);
            SqlCommand sqCom = new SqlCommand();

            sqCom.CommandText = "update Products_Banks set parent=null where id=@id";
            sqCom.Parameters.Add("@id", SqlDbType.Int).Value = id;
            Database.ExecuteNonQuery(sqCom, null);
            Refr(0, id);
        }

    }
}
