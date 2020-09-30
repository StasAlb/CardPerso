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
    public partial class Bank : System.Web.UI.Page
    {
        string res = "";
        DataSet ds = new DataSet();
        ServiceClass sc = new ServiceClass();
        protected void Page_Load(object sender, EventArgs e)
        {
            ClientScript.RegisterHiddenField("resd", "");
            if (IsPostBack)
                return;            
            lock (Database.lockObjectDB)
            {
                lbInform.Text = "";
                Refr(0);
            }
        }

        private void Refr(int rowindex)
        {
            lbInform.Text = "";
            ds.Clear();

            res = Database.ExecuteQuery("select * from Banks", ref ds, null);
            gvBanks.DataSource = ds.Tables[0];
            gvBanks.DataBind();

            if (gvBanks.Rows.Count > 0)
            {
                gvBanks.SelectedIndex = rowindex;
                gvBanks.Rows[gvBanks.SelectedIndex].Focus();
            }
            SetButton();

            lbCount.Text = "Кол-во: " + gvBanks.Rows.Count.ToString();
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
                    ep.ExportGridExcel(gvBanks);
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
                Refr(gvBanks.Rows.Count);
            }
        }

        protected void gvBanks_SelectedIndexChanged(object sender, EventArgs e)
        {
            lock (Database.lockObjectDB)
            {
                gvBanks.Rows[gvBanks.SelectedIndex].Focus();
                SetButton();
            }
        }

        private void SetButton()
        {
            if (!sc.UserAction(User.Identity.Name, Restrictions.LibraryEdit))
            {
                bNew.Visible = false;
                bEdit.Visible = false;
                bDelete.Visible = false;
                bExcel.Visible = false;
                return;
            }
            if (gvBanks.Rows.Count > 0)
            {
                bEdit.Visible = true;
                bDelete.Visible = true;
                bExcel.Visible = true;

                bDelete.Attributes.Add("OnClick", String.Format("return confirm('Удалить банк {0}?');", gvBanks.DataKeys[Convert.ToInt32(gvBanks.SelectedIndex)].Values["name"].ToString()));
                bEdit.Attributes.Add("OnClick", String.Format("return show_catalog('type=bank&mode=2&id={0}')", gvBanks.DataKeys[Convert.ToInt32(gvBanks.SelectedIndex)].Values["id"].ToString()));
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
                Refr(gvBanks.SelectedIndex);
            }
        }

        protected void bDelete_Click(object sender, ImageClickEventArgs e)
        {
            lock (Database.lockObjectDB)
            {
                int id = Convert.ToInt32(gvBanks.DataKeys[Convert.ToInt32(gvBanks.SelectedIndex)].Values["id"]);
                if (!Database.CheckDelBank(id, null))
                {
                    lbInform.Text = "Невозможно удалить банк, так как существует связанная продукция.";
                    return;
                }
                SqlCommand sqCom = new SqlCommand();
                SqlTransaction trans = Database.Conn.BeginTransaction(User.Identity.Name);
                sqCom.CommandText = "delete from Banks where id=@id";
                sqCom.Parameters.Add("@id", SqlDbType.Int).Value = id;
                Database.ExecuteNonQuery(sqCom, null);
                lbInform.Text = "";
                Refr(0);
            }
        }
    }
}
