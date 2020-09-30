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
    public partial class ListDeliver : System.Web.UI.Page
    {
        string res = "";
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
            lbInform.Text = "";
            ds.Clear();

            res = Database.ExecuteQuery("select * from V_DeliversBranchs order by deliver_h,branch", ref ds, null);
            gvDelivers.DataSource = ds.Tables[0];
            gvDelivers.DataBind();

            if (gvDelivers.Rows.Count > 0)
            {
                gvDelivers.SelectedIndex = rowindex;
                gvDelivers.Rows[gvDelivers.SelectedIndex].Focus();
            }
            SetButton();

            lbCount.Text = "Кол-во: " + gvDelivers.Rows.Count.ToString();
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
                    ep.ExportGridExcel(gvDelivers);
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
                Refr(gvDelivers.Rows.Count);
            }
        }

        protected void gvDelivers_SelectedIndexChanged(object sender, EventArgs e)
        {
            lock (Database.lockObjectDB)
            {
                gvDelivers.Rows[gvDelivers.SelectedIndex].Focus();
                SetButton();
            }
        }

        private void SetButton()
        {
            lbInform.Text = "";
            ServiceClass sc = new ServiceClass();
            if (!sc.UserAction(User.Identity.Name, Restrictions.LibraryEdit))
            {
                bNew.Visible = false;
                bEdit.Visible = false;
                bDelete.Visible = false;
                bExcel.Visible = false;
                bDelFilial.Visible = false;
                bAddFilial.Visible = false;
                return;
            }

            if (gvDelivers.Rows.Count > 0)
            {
                bEdit.Visible = true;
                bDelete.Visible=(gvDelivers.DataKeys[gvDelivers.SelectedIndex].Values["deliver"].ToString()!="");
                //bDelete.Visible = true;
                bExcel.Visible = true;
                bDelFilial.Visible = true;
                bAddFilial.Visible = true;

                bDelete.Attributes.Add("OnClick", String.Format("return confirm('Удалить рассылку {0}?');", gvDelivers.DataKeys[Convert.ToInt32(gvDelivers.SelectedIndex)].Values["deliver_h"].ToString()));
                bDelFilial.Attributes.Add("OnClick", String.Format("return confirm('Отвязать филиал {0} от рассылки {1}?');", gvDelivers.DataKeys[Convert.ToInt32(gvDelivers.SelectedIndex)].Values["branch"].ToString(),gvDelivers.DataKeys[Convert.ToInt32(gvDelivers.SelectedIndex)].Values["deliver_h"].ToString()));
                bEdit.Attributes.Add("OnClick", String.Format("return show_listdeliver('mode=2&id_db={0}&id_deliv={1}')", gvDelivers.DataKeys[Convert.ToInt32(gvDelivers.SelectedIndex)].Values["id_db"].ToString(), gvDelivers.DataKeys[Convert.ToInt32(gvDelivers.SelectedIndex)].Values["id_deliv"].ToString()));
                bAddFilial.Attributes.Add("OnClick", String.Format("return show_listdeliver('mode=3&id_db=0&id_deliv={0}')", gvDelivers.DataKeys[Convert.ToInt32(gvDelivers.SelectedIndex)].Values["id_deliv"].ToString()));
            }
            else
            {
                bEdit.Visible = false;
                bDelete.Visible = false;
                bExcel.Visible = false;
                bDelFilial.Visible = false;
                bAddFilial.Visible = false;
            }
        }

        protected void bEdit_Click(object sender, ImageClickEventArgs e)
        {
            lock (Database.lockObjectDB)
            {
                Refr(gvDelivers.SelectedIndex);
            }
        }

        protected void bDelete_Click(object sender, ImageClickEventArgs e)
        {
            lock (Database.lockObjectDB)
            {
                int id = Convert.ToInt32(gvDelivers.DataKeys[Convert.ToInt32(gvDelivers.SelectedIndex)].Values["id_deliv"]);

                SqlCommand sqCom = new SqlCommand();
                sqCom.CommandText = "delete from Delivers_Branchs where id_deliv=@id_deliv";
                sqCom.Parameters.Add("@id_deliv", SqlDbType.Int).Value = id;
                Database.ExecuteNonQuery(sqCom, null);
                sqCom.Parameters.Clear();
                sqCom.CommandText = "delete from Delivers where id=@id";
                sqCom.Parameters.Add("@id", SqlDbType.Int).Value = id;
                Database.ExecuteNonQuery(sqCom, null);
                Refr(0);
            }
        }

        protected void bDelFilial_Click(object sender, ImageClickEventArgs e)
        {
            lock (Database.lockObjectDB)
            {
                int id_deliver = Convert.ToInt32(gvDelivers.DataKeys[Convert.ToInt32(gvDelivers.SelectedIndex)].Values["id_deliv"]);
                if (Database.CntBranchsInDeliver(id_deliver, null) < 2)
                {
                    lbInform.Text = "Невозможно отвязать филиал";
                    return;
                }

                int id = Convert.ToInt32(gvDelivers.DataKeys[Convert.ToInt32(gvDelivers.SelectedIndex)].Values["id_db"]);
                SqlCommand sqCom = new SqlCommand();

                sqCom.CommandText = "delete from Delivers_Branchs where id=@id";
                sqCom.Parameters.Add("@id", SqlDbType.Int).Value = id;
                Database.ExecuteNonQuery(sqCom, null);
                Refr(0);
            }
        }

        protected void bAddFilial_Click(object sender, ImageClickEventArgs e)
        {
            lock (Database.lockObjectDB)
            {
                Refr(gvDelivers.SelectedIndex);
            }
        }
    }
}
