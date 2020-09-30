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

namespace CardPerso
{
    public partial class CardProperties : System.Web.UI.Page
    {
        string idcard = "-1";
        int tpi = -1;
        string iddoc = "-1";
        protected void Page_Load(object sender, EventArgs e)
        {
            pRename.Visible = false;
            if (IsPostBack)
                return;
            lock (Database.lockObjectDB)
            {
                DataSet ds = new DataSet();
                string res = Database.ExecuteQuery("select id,prop from CardProperty", ref ds, null);
                if (ds == null || ds.Tables.Count == 0)
                    return;
                ddlProperty.DataSource = ds.Tables[0];
                ddlProperty.DataTextField = "prop";
                ddlProperty.DataValueField = "id";
                ddlProperty.DataBind();
                ds.Clear();
                idcard = Request.QueryString["id_card"];
                iddoc = Request.QueryString["id_doc"];
                try
                {
                    tpi = Convert.ToInt32(Request.QueryString["tp"]);
                }
                catch
                {
                    tpi = 1;
                }
                pRename.Visible = (tpi == 2);
                pProperty.Visible = (tpi == 1);
                if (idcard.Length > 0)
                {
                    res = Database.ExecuteQuery("select pan, fio, id_prop, id_stat, passport, unemb from cards where id=" + idcard.ToString(), ref ds, null);
                    if (ds != null && ds.Tables.Count > 0 && ds.Tables[0].Rows.Count > 0)
                    {
                        ddlProperty.SelectedValue = Convert.ToString(ds.Tables[0].Rows[0]["id_prop"]);
                        lTitle.Text = String.Format("{0} ({1})", ds.Tables[0].Rows[0]["pan"], ds.Tables[0].Rows[0]["fio"]);
                        tbFio.Text = ds.Tables[0].Rows[0]["fio"].ToString();
                        tbPass.Text = ds.Tables[0].Rows[0]["passport"].ToString();
                        bool unemb = (ds.Tables[0].Rows[0]["unemb"] == DBNull.Value) ? false : Convert.ToBoolean(ds.Tables[0].Rows[0]["unemb"]);
                        if (Convert.ToInt32(ds.Tables[0].Rows[0]["id_stat"]) == 4 && unemb && tpi == 2)
                            pRename.Visible = true;
                    }
                }
            }
        }
        protected void bSave_Click(object sender, ImageClickEventArgs e)
        {
            lock (Database.lockObjectDB)
            {
                idcard = Request.QueryString["id_card"];
                iddoc = Request.QueryString["id_doc"];
                try
                {
                    tpi = Convert.ToInt32(Request.QueryString["tp"]);
                }
                catch
                {
                    tpi = 1;
                }
                if (tpi == 1)
                {
                    if (selectallcard.Checked == false)
                    {
                        Database.ExecuteNonQuery(String.Format("update cards set id_prop={0} where id={1}", ddlProperty.SelectedValue, idcard, tbFio.Text.Trim(), tbPass.Text.Trim()), null);
                    }
                    else
                    {
                        Database.ExecuteNonQuery(String.Format("update cards set id_prop={0} where id in (select id_card from cards_storagedocs where id_doc={1})", ddlProperty.SelectedValue, iddoc), null);
                    }
                }
                if (tpi == 2)
                    Database.ExecuteNonQuery(String.Format("update cards set fio='{2}', passport='{3}' where id={1}", ddlProperty.SelectedValue, idcard, tbFio.Text.Trim(), tbPass.Text.Trim()), null);
                Response.Write("<script language=javascript>window.returnValue='" + ddlProperty.SelectedValue + "'; window.close();</script>");
            }
        }
    }
}
