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
    public partial class AddComment : System.Web.UI.Page
    {
        DataSet ds = new DataSet();
        string res = "";
        int id_card = 0;

        protected void Page_Load(object sender, EventArgs e)
        {
            if (IsPostBack)
                return;
            id_card = Convert.ToInt32(Request.QueryString["id_card"]);


            lock (Database.lockObjectDB)
            {
                ZapFields();
                tbComment.Focus();
            }
        }

        private void ZapFields()
        {
            lbInform.Text = "";

            res = Database.ExecuteQuery(String.Format("select comment from Cards_StorageDocs where id={0}",id_card), ref ds, null);
            tbComment.Text = ds.Tables[0].Rows[0]["comment"].ToString();
        }

        protected void bSave_Click(object sender, ImageClickEventArgs e)
        {
            lock (Database.lockObjectDB)
            {
                id_card = Convert.ToInt32(Request.QueryString["id_card"]);
                SqlCommand sqCom = new SqlCommand();
                sqCom.CommandText = "update Cards_StorageDocs set comment=@comment where id=@id";
                sqCom.Parameters.Add("@id", SqlDbType.Int).Value = id_card;
                sqCom.Parameters.Add("@comment", SqlDbType.VarChar, 50).Value = tbComment.Text;
                Database.ExecuteNonQuery(sqCom, null);
                Response.Write("<script language=javascript>window.returnValue='1'; window.close();</script>");
            }
        }
    }
}
