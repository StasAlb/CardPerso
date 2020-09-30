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
    public partial class ListDeliverEdit : System.Web.UI.Page
    {
        DataSet ds = new DataSet();
        int mode = 0; int id_deliv = 0; int id_db=0;
        string res = "";

        protected void Page_Load(object sender, EventArgs e)
        {
            lock (Database.lockObjectDB)
            {
                mode = Convert.ToInt32(Request.QueryString["mode"]);
                id_deliv = Convert.ToInt32(Request.QueryString["id_deliv"]);
                id_db = Convert.ToInt32(Request.QueryString["id_db"]);

                if (mode == 1) Title = "Добавление рассылки";
                if (mode == 2) Title = "Редактирование";
                if (mode == 3) Title = "Привязка филиала";

                if (!IsPostBack)
                {
                    ZapCombo();
                    if (mode == 2 || mode == 3) ZapFields();
                    if (mode == 3) pDeliver.Enabled = false;
                }
            }
        }

        private void ZapCombo()
        {
            ds.Clear();
            res = Database.ExecuteQuery("select id,department from Branchs order by department", ref ds, null);
            for (int i = 0; i < ds.Tables[0].Rows.Count; i++)
                dListFilial.Items.Add(new ListItem(ds.Tables[0].Rows[i]["department"].ToString(), ds.Tables[0].Rows[i]["id"].ToString()));
            dListFilial.SelectedIndex = 0;

            tbName.Focus();
        }

        private void ZapFields()
        {
            ds.Clear();
            res = Database.ExecuteQuery(String.Format("select * from Delivers where id={0}",id_deliv), ref ds, null);

            tbName.Text = ds.Tables[0].Rows[0]["name"].ToString();

            if (mode == 2)
            {
                ds.Clear();
                res = Database.ExecuteQuery(String.Format("select * from Delivers_Branchs where id={0}", id_db), ref ds, null);
                dListFilial.SelectedIndex = dListFilial.Items.IndexOf(dListFilial.Items.FindByValue(ds.Tables[0].Rows[0]["id_branch"].ToString()));
            }

        }

        protected void bSave_Click(object sender, ImageClickEventArgs e)
        {
            lock (Database.lockObjectDB)
            {
                if (tbName.Text == "")
                {
                    lbInform.Text = "Введите наименование рассылки";
                    tbName.Focus();
                    return;
                }

                if (dListFilial.SelectedIndex < 0)
                {
                    lbInform.Text = "Выберите филиал";
                    dListFilial.Focus();
                    return;
                }

                if (mode == 2 || mode == 3)
                    if (!Database.CheckBranchInDeliver(id_db, id_deliv, Convert.ToInt32(dListFilial.SelectedItem.Value), null))
                    {
                        lbInform.Text = "Такой филиал уже привязан";
                        dListFilial.Focus();
                        return;
                    }

                SqlCommand sqCom = new SqlCommand();

                if (mode == 1)
                {
                    sqCom.CommandText = "insert into Delivers (name) values(@name) select @@identity as lid";
                    sqCom.Parameters.Add("@name", SqlDbType.VarChar, 50).Value = tbName.Text;
                    object lid = 0;
                    res = Database.ExecuteScalar(sqCom, ref lid, null);
                    sqCom.Parameters.Clear();

                    sqCom.CommandText = "insert into Delivers_Branchs (id_deliv,id_branch) values(@id_deliv,@id_branch)";
                    sqCom.Parameters.Add("@id_deliv", SqlDbType.Int);//.Value = (int)lid;
                    sqCom.Parameters.Add("@id_branch", SqlDbType.Int);//.Value = dListFilial.SelectedItem.Value;
                    sqCom.Parameters["@id_deliv"].Value = Convert.ToInt32(lid);
                    sqCom.Parameters["@id_branch"].Value = dListFilial.SelectedItem.Value;
                    res = Database.ExecuteNonQuery(sqCom, null);
                }

                if (mode == 2)
                {
                    sqCom.CommandText = "update Delivers set name=@name where id=@id";
                    sqCom.Parameters.Add("@id", SqlDbType.Int).Value = id_deliv;
                    sqCom.Parameters.Add("@name", SqlDbType.VarChar, 50).Value = tbName.Text;
                    res = Database.ExecuteNonQuery(sqCom, null);
                    sqCom.Parameters.Clear();

                    sqCom.CommandText = "update Delivers_Branchs set id_branch=@id_branch where id=@id";
                    sqCom.Parameters.Add("@id", SqlDbType.Int).Value = id_db;
                    sqCom.Parameters.Add("@id_branch", SqlDbType.Int).Value = dListFilial.SelectedItem.Value;

                    res = Database.ExecuteNonQuery(sqCom, null);
                }

                if (mode == 3)
                {
                    sqCom.CommandText = "insert into Delivers_Branchs (id_deliv,id_branch) values(@id_deliv,@id_branch)";
                    sqCom.Parameters.Add("@id_deliv", SqlDbType.Int).Value = id_deliv;
                    sqCom.Parameters.Add("@id_branch", SqlDbType.Int).Value = dListFilial.SelectedItem.Value;
                    res = Database.ExecuteNonQuery(sqCom, null);
                }

                Response.Write("<script language=javascript>window.returnValue='1'; window.close();</script>");
            }
        }
    }
}
