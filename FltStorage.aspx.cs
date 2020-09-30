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
    public partial class FltStorage : System.Web.UI.Page
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
                    tbName.Focus();
                }
            }
        }

        private void ZapCombo()
        {
            ds.Clear();

            dListBank.Items.Add(new ListItem("Все", "-1"));
            res = Database.ExecuteQuery("select id,name from Banks", ref ds, null);

            //dListBank.DataSource = ds.Tables[0];
            //dListBank.DataTextField = "name";
            //dListBank.DataValueField = "id";
            //dListBank.DataBind();

            for (int i = 0; i < ds.Tables[0].Rows.Count; i++)
                dListBank.Items.Add(new ListItem(ds.Tables[0].Rows[i]["name"].ToString(), ds.Tables[0].Rows[i]["id"].ToString()));
            dListBank.SelectedIndex = -1;
        }

        protected void bSave_Click(object sender, ImageClickEventArgs e)
        {
            lock (Database.lockObjectDB)
            {
                ArrayList al = new ArrayList();
                string s = "";

                if (tbName.Text != "")
                    al.Add(String.Format("(name like [%{0}%])", tbName.Text));

                string id_bank = dListBank.SelectedItem.Value;
                if (id_bank != "-1")
                    al.Add(String.Format("(id_bank={0})", id_bank));

                if (tbBin.Text != "")
                    al.Add(String.Format("(bin like [%{0}%])", tbBin.Text));

                if (chMin.Checked)
                    al.Add("(min_cnt>=cnt_new)");

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
