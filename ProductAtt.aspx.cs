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
    public partial class ProductAtt : System.Web.UI.Page
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

            res = Database.ExecuteQuery("select * from V_ProductsAttachments order by prod_name_p_h", ref ds, null);
            gvAttachments.DataSource = ds.Tables[0];
            gvAttachments.DataBind();

            if (gvAttachments.Rows.Count > 0)
            {
                gvAttachments.SelectedIndex = rowindex;
                gvAttachments.Rows[gvAttachments.SelectedIndex].Focus();
            }
            SetButton();

            lbCount.Text = "Кол-во: " + gvAttachments.Rows.Count.ToString();
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
                    ep.ExportGridExcel(gvAttachments);
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

        protected void gvAttachments_SelectedIndexChanged(object sender, EventArgs e)
        {
            lock (Database.lockObjectDB)
            {
                gvAttachments.Rows[gvAttachments.SelectedIndex].Focus();
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
                return;
            }

            if (gvAttachments.Rows.Count > 0)
            {
                bEdit.Visible = true;
                bDelete.Visible = true;
                bExcel.Visible = true;

                bDelete.Attributes.Add("OnClick", String.Format("return confirm('Удалить вложение {0}?');", gvAttachments.DataKeys[Convert.ToInt32(gvAttachments.SelectedIndex)].Values["prod_name_at"].ToString()));
                bEdit.Attributes.Add("OnClick", String.Format("return show_productatt('mode=2&id_pa={0}&id_prb={1}')", gvAttachments.DataKeys[Convert.ToInt32(gvAttachments.SelectedIndex)].Values["id_pa"].ToString(), gvAttachments.DataKeys[Convert.ToInt32(gvAttachments.SelectedIndex)].Values["id_prb_p"].ToString()));
            }
            else
            {
                bEdit.Visible = false;
                bDelete.Visible = false;
                bExcel.Visible = false;
            }
        }

        protected void bEdit_Click(object sender, ImageClickEventArgs e)
        {
            lock (Database.lockObjectDB)
            {
                Refr(gvAttachments.SelectedIndex);
            }
        }

        protected void bDelete_Click(object sender, ImageClickEventArgs e)
        {
            lock (Database.lockObjectDB)
            {
                int id = Convert.ToInt32(gvAttachments.DataKeys[Convert.ToInt32(gvAttachments.SelectedIndex)].Values["id_pa"]);

                SqlCommand sqCom = new SqlCommand();
                sqCom.CommandText = "delete from Products_Attachments where id=@id";
                sqCom.Parameters.Add("@id", SqlDbType.Int).Value = id;
                res = Database.ExecuteNonQuery(sqCom, null);
                Refr(0);
            }
        }
    }
}
