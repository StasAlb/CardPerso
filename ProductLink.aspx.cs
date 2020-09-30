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
    public partial class ProductLink : System.Web.UI.Page
    {
        DataSet ds = new DataSet();
        string res = "";
        int id_prod = 0;

        protected void Page_Load(object sender, EventArgs e)
        {
            id_prod = Convert.ToInt32(Request.QueryString["id"]);
            if (IsPostBack)
                return;

            lock(Database.lockObjectDB)
            {
                RefrOffice();
            }
        }

        private void RefrOffice()
        {
            ds.Clear();

            res = Database.ExecuteQuery(String.Format("select id_prb,prod_name from V_ProductsBanks_T where parent is null and id_type=1 and id_prod<>{0} and id_prod not in (select parent from Products_Banks where parent>0) order by prod_name", id_prod), ref ds, null);

            dListProd.DataSource = ds.Tables[0];
            dListProd.DataTextField = "prod_name";
            dListProd.DataValueField = "id_prb";
            dListProd.DataBind();
        }


        protected void bSave_Click(object sender, ImageClickEventArgs e)
        {
            lock (Database.lockObjectDB)
            {
                SqlCommand sqCom = new SqlCommand();

                sqCom.CommandText = "update Products_Banks set parent=@parent where id=@id";
                sqCom.Parameters.Add("@parent", SqlDbType.Int).Value = id_prod;
                sqCom.Parameters.Add("@id", SqlDbType.Int).Value = Convert.ToInt32(dListProd.SelectedItem.Value);
                Database.ExecuteNonQuery(sqCom, null);

                lbInform.Text = "Продукт \"" + dListProd.SelectedItem.Text + "\" привязан. Обновление после закрытия формы.";

                RefrOffice();
            }
        }
    }
}
