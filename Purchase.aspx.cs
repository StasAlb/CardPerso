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
    public partial class Purchase : System.Web.UI.Page
    {
        private string res="";
        private DataSet ds = new DataSet();
        ServiceClass sc = new ServiceClass();
        protected void Page_Load(object sender, EventArgs e)
        {
            lock (Database.lockObjectDB)
            {
                ClientScript.RegisterHiddenField("resd", "");
                if (!sc.UserAction(User.Identity.Name, Restrictions.PurchaseView))
                    Response.Redirect("~\\Account\\Restricted.aspx", true);

                if (!IsPostBack)
                {
                    lbInform.Text = "";
                    Refr(0, true);
                }
            }
        }

        private void Refr(int rowindex,bool conf)
        {
                ds.Clear();
                res = Database.ExecuteQuery(String.Format("select * from V_PurchDogs {0} {1}", lbSearch.Text, lbSort.Text), ref ds, null);
                gvPurchDogs.DataSource = ds.Tables[0];
                gvPurchDogs.DataBind();
                lbCountD.Text = "Договоров: " + ds.Tables[0].Rows.Count.ToString();
                if (gvPurchDogs.Rows.Count > 0)
                {
                    gvPurchDogs.SelectedIndex = rowindex;
                    gvPurchDogs.Rows[gvPurchDogs.SelectedIndex].Focus();
                    ViewProducts(Convert.ToInt32(gvPurchDogs.DataKeys[gvPurchDogs.SelectedIndex].Values["id"]), 0);
                }
                else ViewProducts(-1, 0);

                if (conf)
                {
                    string s_fld = Database.GetFiledsByUser(sc.UserId(User.Identity.Name), "purchase_dog", null);
                    if (s_fld != "") FuncClass.HideFields(s_fld, gvPurchDogs);
                    s_fld = Database.GetFiledsByUser(sc.UserId(User.Identity.Name), "purchase_product", null);
                    if (s_fld != "") FuncClass.HideFields(s_fld, gvProducts);
                }

                SetButtonD();
        }

        private void ViewProducts(int id,int rowindex)
        {
            ds.Clear();
            res = Database.ExecuteQuery(String.Format("select * from V_Products_PurchDogs where id_dog={0} order by prod_name",id), ref ds, null);
            gvProducts.DataSource = ds.Tables[0];
            gvProducts.DataBind();

            if (gvProducts.Rows.Count > 0)
            {
                gvProducts.SelectedIndex = rowindex;
                gvProducts.Rows[gvProducts.SelectedIndex].Focus();
            }

            SetButtonP();
            lbCountP.Text = "Продуктов: " + gvProducts.Rows.Count.ToString();
        }

        private void SetButtonD()
        {
            if (!sc.UserAction(User.Identity.Name, Restrictions.PurchaseEdit))
            {
                bNewD.Visible = false;
                bEditD.Visible = false;
                bDeleteD.Visible = false;
                bExcel.Visible = false;
                bNewP.Visible = false;
                bDeleteP.Visible = false;
                bEditP.Visible = false;
                return;
            }

            if (gvPurchDogs.Rows.Count > 0)
            {
                bEditD.Visible = true;
                bDeleteD.Visible = true;
                bExcel.Visible = true;
                bNewP.Visible = true;

                bDeleteD.Attributes.Add("OnClick", String.Format("return confirm('Удалить договор № {0}?');", gvPurchDogs.DataKeys[Convert.ToInt32(gvPurchDogs.SelectedIndex)].Values["number_dog"].ToString()));
                bEditD.Attributes.Add("OnClick", String.Format("return show_purchase_dog('mode=2&id={0}')", gvPurchDogs.DataKeys[Convert.ToInt32(gvPurchDogs.SelectedIndex)].Values["id"].ToString()));
                bNewP.Attributes.Add("OnClick", String.Format("return show_purchase_product('mode=1&id_dog={0}','1')", gvPurchDogs.DataKeys[Convert.ToInt32(gvPurchDogs.SelectedIndex)].Values["id"].ToString()));
            }
            else
            {
                bEditD.Visible = false;
                bDeleteD.Visible = false;
                bExcel.Visible = false;
                bNewP.Visible = false;
            }

            bResetFilter.Visible = (lbSearch.Text != "");
        }

        private void SetButtonP()
        {
            if (gvProducts.Rows.Count > 0)
            {
                bEditP.Visible = true;
                bDeleteP.Visible = true;
                bEditP.Attributes.Add("OnClick", String.Format("return show_purchase_product('mode=2&id_dog={0}&id_prod={1}','2')", gvPurchDogs.DataKeys[Convert.ToInt32(gvPurchDogs.SelectedIndex)].Values["id"].ToString(), gvProducts.DataKeys[Convert.ToInt32(gvProducts.SelectedIndex)].Values["id"].ToString()));
                bDeleteP.Attributes.Add("OnClick", String.Format("return confirm('Удалить продукт {0} ({1})?');", gvProducts.DataKeys[Convert.ToInt32(gvProducts.SelectedIndex)].Values["prod_name"].ToString(), gvProducts.DataKeys[Convert.ToInt32(gvProducts.SelectedIndex)].Values["bank_name"].ToString()));
            }
            else
            {
                bEditP.Visible = false;
                bDeleteP.Visible = false;
            }
        }

        protected void gvPurchDogs_SelectedIndexChanged(object sender, EventArgs e)
        {
            lock (Database.lockObjectDB)
            {
                gvPurchDogs.Rows[gvPurchDogs.SelectedIndex].Focus();
                ViewProducts(Convert.ToInt32(gvPurchDogs.DataKeys[gvPurchDogs.SelectedIndex].Values["id"]), 0);
                SetButtonD();
                lbInform.Text = "";
            }
        }

        protected void gvPurchDogs_Sorting(object sender, GridViewSortEventArgs e)
        {
            lock (Database.lockObjectDB)
            {
                int i = FuncClass.GetSortIndex(e.SortExpression, gvPurchDogs);

                if (lbSortIndex.Text != "")
                {
                    int ind = Convert.ToInt32(lbSortIndex.Text);
                    gvPurchDogs.Columns[ind].HeaderText = gvPurchDogs.Columns[ind].HeaderText.Replace("^", "");
                    gvPurchDogs.Columns[ind].HeaderStyle.BackColor = System.Drawing.Color.FromArgb(0, 102, 153);
                }

                lbSortIndex.Text = i.ToString();

                gvPurchDogs.Columns[i].HeaderStyle.BackColor = System.Drawing.Color.FromArgb(102, 153, 153);

                if ("order by " + e.SortExpression + " asc" == lbSort.Text)
                {
                    lbSort.Text = "order by " + e.SortExpression + " desc";
                    gvPurchDogs.Columns[i].HeaderText = gvPurchDogs.Columns[i].HeaderText + "^";
                }
                else
                    lbSort.Text = "order by " + e.SortExpression + " asc";

                Refr(0, false);
            }
        }

        private bool CheckProducts(int id)
        {
            bool chp = true;
            int cnt = 0;
            int id_prb = 0;

            ds.Clear();
            res = Database.ExecuteQuery(String.Format("select id,id_prb,cnt,prod_name,bank_name from V_Products_PurchDogs where id_dog={0}",id.ToString()), ref ds, null);
            for (int i = 0; i < ds.Tables[0].Rows.Count; i++)
            {
                cnt = Convert.ToInt32(ds.Tables[0].Rows[i]["cnt"]);
                id_prb = Convert.ToInt32(ds.Tables[0].Rows[i]["id_prb"]);

                if (cnt > Database.CntStorage(id_prb, "new",null))
                {
                    lbInform.Text = "Продукт: " + ds.Tables[0].Rows[i]["prod_name"] + " (" + ds.Tables[0].Rows[i]["bank_name"] + ")."+" Нет достаточного количества в хранилище";
                    chp = false;
                }
            }
            return chp;
        }

        private void DelProducts(int id)
        {
            int cnt = 0;
            int id_prb = 0;

            ds.Clear();
            res = Database.ExecuteQuery(String.Format("select id,id_prb,cnt from Products_PurchDogs where id_dog={0}", id.ToString()), ref ds, null);
            for (int i = 0; i < ds.Tables[0].Rows.Count; i++)
            {
                cnt = Convert.ToInt32(ds.Tables[0].Rows[i]["cnt"]);
                id_prb = Convert.ToInt32(ds.Tables[0].Rows[i]["id_prb"]);
                Database.StorageM(cnt, id_prb,"new",null);
            }
         }

        protected void bDeleteP_Click(object sender, ImageClickEventArgs e)
        {
            lock (Database.lockObjectDB)
            {
                int id = Convert.ToInt32(gvProducts.DataKeys[gvProducts.SelectedIndex].Values["id"]);
                int id_prb = Convert.ToInt32(gvProducts.DataKeys[gvProducts.SelectedIndex].Values["id_prb"]);
                int cnt = Convert.ToInt32(gvProducts.DataKeys[gvProducts.SelectedIndex].Values["cnt"]);

                if (cnt > Database.CntStorage(id_prb, "new", null))
                {
                    lbInform.Text = "Продукт: " + gvProducts.DataKeys[gvProducts.SelectedIndex].Values["prod_name"] + " (" + gvProducts.DataKeys[gvProducts.SelectedIndex].Values["bank_name"] + ")." + " Нет достаточного количества в хранилище";
                    return;
                }

                SqlCommand sqCom = new SqlCommand();

                sqCom.CommandText = "delete from Products_PurchDogs where id=@id";
                sqCom.Parameters.Add("@id", SqlDbType.Int).Value = id;
                Database.ExecuteNonQuery(sqCom, null);

                Database.StorageM(cnt, id_prb, "new", null);

                ViewProducts(Convert.ToInt32(gvPurchDogs.DataKeys[gvPurchDogs.SelectedIndex].Values["id"]), 0);
                lbInform.Text = "";
            }
        }

        protected void gvProducts_SelectedIndexChanged(object sender, EventArgs e)
        {
            lock (Database.lockObjectDB)
            {
                gvProducts.Rows[gvProducts.SelectedIndex].Focus();
                SetButtonP();
                lbInform.Text = "";
            }
        }

        protected void bNewP_Click(object sender, ImageClickEventArgs e)
        {
            lock (Database.lockObjectDB)
            {
                ViewProducts(Convert.ToInt32(gvPurchDogs.DataKeys[gvPurchDogs.SelectedIndex].Values["id"]), 0);
                lbInform.Text = "";
            }
        }

        protected void bEditP_Click(object sender, ImageClickEventArgs e)
        {
            lock (Database.lockObjectDB)
            {
                ViewProducts(Convert.ToInt32(gvPurchDogs.DataKeys[gvPurchDogs.SelectedIndex].Values["id"]), gvProducts.SelectedIndex);
                lbInform.Text = "";
            }
        }

        protected void bDeleteD_Click(object sender, ImageClickEventArgs e)
        {
            lock (Database.lockObjectDB)
            {
                int id = Convert.ToInt32(gvPurchDogs.DataKeys[gvPurchDogs.SelectedIndex].Values["id"]);
                if (!CheckProducts(id)) return;
                DelProducts(id);
                SqlCommand sqCom = new SqlCommand();
                object obj = null;
                sqCom.CommandText = "delete from Products_PurchDogs where id_dog=@id";
                sqCom.Parameters.Add("@id", SqlDbType.Int).Value = id;
                Database.ExecuteNonQuery(sqCom, null);
                sqCom.CommandText = "select number_dog from PurchDogs where id=@id";
                res = Database.ExecuteScalar(sqCom, ref obj, null);
                sqCom.CommandText = "delete from PurchDogs where id=@id";
                Database.ExecuteNonQuery(sqCom, null);
                Database.Log(sc.UserGuid(User.Identity.Name), String.Format("Удален договор на закупку {0}", (obj == null) ? "" : (string)obj), null);

                Refr(0, false);
                lbInform.Text = "";
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
                    ep.ExportGridExcel(gvPurchDogs);
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

        protected void bNewD_Click(object sender, ImageClickEventArgs e)
        {
            lock (Database.lockObjectDB)
            {
                Refr(0, false);
                lbInform.Text = "";
            }
        }

        protected void bEditD_Click(object sender, ImageClickEventArgs e)
        {
            lock (Database.lockObjectDB)
            {
                Refr(gvPurchDogs.SelectedIndex, false);
                lbInform.Text = "";
            }
        }

        protected void bSetFilter_Click(object sender, ImageClickEventArgs e)
        {
            lock (Database.lockObjectDB)
            {
                lbSearch.Text = Request.Form["resd"].Replace("[", "'");
                lbSearch.Text = lbSearch.Text.Replace("]", "'");
                Refr(0, false);
            }
        }

        protected void bResetFilter_Click(object sender, ImageClickEventArgs e)
        {
            lock (Database.lockObjectDB)
            {
                lbSearch.Text = "";
                Refr(0, false);
            }
        }

        protected void bConfFieldP_Click(object sender, ImageClickEventArgs e)
        {
            Response.Redirect("~/Purchase.aspx");
        }

        protected void bConfFieldD_Click(object sender, ImageClickEventArgs e)
        {
            Response.Redirect("~/Purchase.aspx");
        }

        protected void gvPurchDogs_PageIndexChanging(object sender, GridViewPageEventArgs e)
        {
            lock (Database.lockObjectDB)
            {
                gvPurchDogs.PageIndex = e.NewPageIndex;
                Refr(0, false);
            }
        }
    }
}
