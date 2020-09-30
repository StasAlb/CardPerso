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

namespace CardPerso
{
    public partial class FltPurchase : System.Web.UI.Page
    {
        DataSet ds = new DataSet();
        string res = "";

        protected void Page_Load(object sender, EventArgs e)
        {
            lock (Database.lockObjectDB)
            {
                if (!IsPostBack)
                {
                    ZapCombo();
                    tbNumber.Focus();
                }
            }
        }

        private void ZapCombo()
        {
            ds.Clear();
            res = Database.ExecuteQuery("select id,name from Suppliers", ref ds, null);
            dListSup.Items.Add(new ListItem("Все", "-1"));
            for (int i = 0; i < ds.Tables[0].Rows.Count; i++)
                dListSup.Items.Add(new ListItem(ds.Tables[0].Rows[i]["name"].ToString(), ds.Tables[0].Rows[i]["id"].ToString()));

            dListSup.SelectedIndex = 0;

            ds.Clear();
            res = Database.ExecuteQuery("select id,name from Manufacturers", ref ds, null);
            dListManuf.Items.Add(new ListItem("Все", "-1"));
            for (int i = 0; i < ds.Tables[0].Rows.Count; i++)
                dListManuf.Items.Add(new ListItem(ds.Tables[0].Rows[i]["name"].ToString(), ds.Tables[0].Rows[i]["id"].ToString()));

            dListManuf.SelectedIndex = 0;
        }

        protected void bSave_Click(object sender, ImageClickEventArgs e)
        {
            lock (Database.lockObjectDB)
            {
                ArrayList al = new ArrayList();

                string s = "";

                if (tbDataSt.Text != "")
                {
                    try
                    {
                        Convert.ToDateTime(tbDataSt.Text);
                    }
                    catch
                    {
                        lbInform.Text = "Неправильно введена дата договора с";
                        tbDataSt.Focus();
                        return;
                    }
                }
                if (tbDataEnd.Text != "")
                {
                    try
                    {
                        Convert.ToDateTime(tbDataEnd.Text);
                    }
                    catch
                    {
                        lbInform.Text = "Неправильно введена дата договора по";
                        tbDataEnd.Focus();
                        return;
                    }
                }

                if (tbNumber.Text != "")
                    al.Add(String.Format("(number_dog like [%{0}%])", tbNumber.Text));
                if (tbDataSt.Text != "")
                    al.Add(String.Format("(date_dog>=[{0:" + ConfigurationSettings.AppSettings["DateFormat"] + "}])", Convert.ToDateTime(tbDataSt.Text)));
                if (tbDataEnd.Text != "")
                    al.Add(String.Format("(date_dog<=[{0:" + ConfigurationSettings.AppSettings["DateFormat"] + "}])", Convert.ToDateTime(tbDataEnd.Text)));
                string id_list = dListSup.SelectedItem.Value;
                if (id_list != "-1")
                    al.Add(String.Format("(id_sup={0})", id_list));
                id_list = dListManuf.SelectedItem.Value;
                if (id_list != "-1")
                    al.Add(String.Format("(id_manuf={0})", id_list));
                if (tbProd.Text != "")
                    al.Add(String.Format("(id in (select id_dog from V_Products_PurchDogs where prod_name like [%{0}%]))", tbProd.Text));

                if (al.Count > 0)
                {
                    string[] all = Array.CreateInstance(typeof(string), al.Count) as string[];
                    al.CopyTo(all, 0);
                    s = "where " + String.Join(" and ", all);
                }

                Response.Write("<script language=javascript>window.returnValue='" + s + "'; window.close();</script>");
            }
        }
    }
}
