﻿using System;
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
    public partial class Courier : System.Web.UI.Page
    {
        string res = "";
        DataSet ds = new DataSet();
        ServiceClass sc = new ServiceClass();
        protected void Page_Load(object sender, EventArgs e)
        {
            ClientScript.RegisterHiddenField("resd", "");
            lock (Database.lockObjectDB)
            {
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

            res = Database.ExecuteQuery("select * from Couriers", ref ds, null);
            gvCouriers.DataSource = ds.Tables[0];
            gvCouriers.DataBind();

            if (gvCouriers.Rows.Count > 0)
            {
                gvCouriers.SelectedIndex = rowindex;
                gvCouriers.Rows[gvCouriers.SelectedIndex].Focus();
            }
            SetButton();

            lbCount.Text = "Кол-во: " + gvCouriers.Rows.Count.ToString();
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
                    ep.ExportGridExcel(gvCouriers);
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
                Refr(gvCouriers.Rows.Count);
            }
        }

        protected void gvCouriers_SelectedIndexChanged(object sender, EventArgs e)
        {
            lock (Database.lockObjectDB)
            {
                gvCouriers.Rows[gvCouriers.SelectedIndex].Focus();
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
            if (gvCouriers.Rows.Count > 0)
            {
                bEdit.Visible = true;
                bDelete.Visible = true;
                bExcel.Visible = true;

                bDelete.Attributes.Add("OnClick", String.Format("return confirm('Удалить курьерскую службу {0}?');", gvCouriers.DataKeys[Convert.ToInt32(gvCouriers.SelectedIndex)].Values["name"].ToString()));
                bEdit.Attributes.Add("OnClick", String.Format("return show_catalog('type=courier&mode=2&id={0}')", gvCouriers.DataKeys[Convert.ToInt32(gvCouriers.SelectedIndex)].Values["id"].ToString()));
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
                Refr(gvCouriers.SelectedIndex);
            }
        }

        protected void bDelete_Click(object sender, ImageClickEventArgs e)
        {
            lock (Database.lockObjectDB)
            {
                int id = Convert.ToInt32(gvCouriers.DataKeys[Convert.ToInt32(gvCouriers.SelectedIndex)].Values["id"]);

                if (!Database.CheckDelCourier(id, null))
                {
                    lbInform.Text = "Невозможно удалить курьерскую службу, так как существуют связанные документы.";
                    return;
                }

                SqlCommand sqCom = new SqlCommand();
                sqCom.CommandText = "delete from Couriers where id=@id";
                sqCom.Parameters.Add("@id", SqlDbType.Int).Value = id;
                Database.ExecuteNonQuery(sqCom, null);
                lbInform.Text = "";
                Refr(0);
            }
        }
    }
}
