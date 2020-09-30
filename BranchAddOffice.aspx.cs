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
    public partial class BranchAddOffice : System.Web.UI.Page
    {
        DataSet ds = new DataSet();
        string res = "";
        int id_branch = 0;

        protected void Page_Load(object sender, EventArgs e)
        {
            id_branch = Convert.ToInt32(Request.QueryString["id"]);
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

            res = Database.ExecuteQuery(String.Format("select id,department from Branchs where (id_parent = 0) and (id not in (select id_parent from Branchs)) and (id<>{0}) order by ident_dep",id_branch.ToString()), ref ds, null);

            dListOffice.DataSource = ds.Tables[0];
            dListOffice.DataTextField = "department";
            dListOffice.DataValueField = "id";
            dListOffice.DataBind();
        }


        protected void bSave_Click(object sender, ImageClickEventArgs e)
        {
            lock (Database.lockObjectDB)
            {
                SqlCommand sqCom = new SqlCommand();

                sqCom.CommandText = "update Branchs set id_parent=@id_parent where id=@id";
                sqCom.Parameters.Add("@id_parent", SqlDbType.Int).Value = id_branch;
                sqCom.Parameters.Add("@id", SqlDbType.Int).Value = Convert.ToInt32(dListOffice.SelectedItem.Value);
                Database.ExecuteNonQuery(sqCom, null);

                lbInform.Text = "Офис \"" + dListOffice.SelectedItem.Text + "\" привязан к подразделению. Обновление после закрытия формы.";

                RefrOffice();
            }
        }
    }
}
