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
    public partial class ProductAttEdit : System.Web.UI.Page
    {
        DataSet ds = new DataSet();
        int mode = 0; int id_prb = 0; int id_pa=0;
        string res = "";

        protected void Page_Load(object sender, EventArgs e)
        {
            lock (Database.lockObjectDB)
            {
                mode = Convert.ToInt32(Request.QueryString["mode"]);
                id_prb = Convert.ToInt32(Request.QueryString["id_prb"]);
                id_pa = Convert.ToInt32(Request.QueryString["id_pa"]);

                if (mode == 1) Title = "Добавление вложения";
                if (mode == 2) Title = "Редактирование";

                if (!IsPostBack)
                {
                    ZapProduct();
                    if (mode == 2)
                    {
                        ZapFields();
                        dListProd.Enabled = false;
                        dListAtt.Enabled = false;
                    }
                }
            }
        }

        private void ZapProduct()
        {
            ds.Clear();
            res = Database.ExecuteQuery("select id_prb,(prod_name + ' (' + bank_name+')') as prod from V_ProductsBanks_T where (id_type=1) order by prod", ref ds, null);
            for (int i = 0; i < ds.Tables[0].Rows.Count; i++)
                dListProd.Items.Add(new ListItem(ds.Tables[0].Rows[i]["prod"].ToString(), ds.Tables[0].Rows[i]["id_prb"].ToString()));
            if (dListProd.Items.Count > 0)
            {
                dListProd.SelectedIndex = 0;
                RefrAtt(Convert.ToInt32(dListProd.SelectedItem.Value));
            }
        }

        private void RefrAtt(int id_prod)
        {
            dListAtt.Items.Clear();
            ds.Clear();
            if (mode == 1)
                res = Database.ExecuteQuery(String.Format("select id_prb,(prod_name + ' (' + bank_name+')') as prod from V_ProductsBanks_T where (id_type=3) and (id_prb not in (select id_prb from Products_Attachments where id_prb_parent={0})) order by prod",id_prod.ToString()), ref ds, null);
            if (mode==2)
                res = Database.ExecuteQuery("select id_prb,(prod_name + ' (' + bank_name+')') as prod from V_ProductsBanks_T where (id_type=3) order by prod", ref ds, null);
            for (int i = 0; i < ds.Tables[0].Rows.Count; i++)
                dListAtt.Items.Add(new ListItem(ds.Tables[0].Rows[i]["prod"].ToString(), ds.Tables[0].Rows[i]["id_prb"].ToString()));
            dListAtt.SelectedIndex = 0;
        }

        private void ZapFields()
        {
            dListProd.SelectedIndex = dListProd.Items.IndexOf(dListProd.Items.FindByValue(id_prb.ToString()));

            ds.Clear();
            res = Database.ExecuteQuery(String.Format("select * from Products_Attachments where id={0}", id_pa), ref ds, null);
            dListAtt.SelectedIndex = dListAtt.Items.IndexOf(dListAtt.Items.FindByValue(ds.Tables[0].Rows[0]["id_prb"].ToString()));
            tbCnt.Text = ds.Tables[0].Rows[0]["cnt"].ToString();
        }

        protected void bSave_Click(object sender, ImageClickEventArgs e)
        {
            lock (Database.lockObjectDB)
            {
                int cnt = 0;

                if (dListProd.SelectedIndex < 0)
                {
                    lbInform.Text = "Выберите продукцию";
                    dListProd.Focus();
                    return;
                }

                if (dListAtt.SelectedIndex < 0)
                {
                    lbInform.Text = "Выберите вложение";
                    dListAtt.Focus();
                    return;
                }

                if (tbCnt.Text == "")
                {
                    lbInform.Text = "Введите кол-во продукции";
                    tbCnt.Focus();
                    return;
                }
                else
                {
                    try
                    {
                        cnt = Convert.ToInt32(tbCnt.Text);
                    }
                    catch
                    {
                        lbInform.Text = "Количество продукции должно быть целым";
                        tbCnt.Focus();
                        return;
                    }
                }

                SqlCommand sqCom = new SqlCommand();

                if (mode == 1)
                {
                    sqCom.CommandText = "insert into Products_Attachments (id_prb_parent,id_prb,cnt) values(@id_prb_parent,@id_prb,@cnt)";
                    sqCom.Parameters.Add("@id_prb_parent", SqlDbType.Int).Value = Convert.ToInt32(dListProd.SelectedItem.Value);
                    sqCom.Parameters.Add("@id_prb", SqlDbType.Int).Value = Convert.ToInt32(dListAtt.SelectedItem.Value);
                    sqCom.Parameters.Add("@cnt", SqlDbType.Int).Value = cnt;
                    Database.ExecuteNonQuery(sqCom, null);
                }

                if (mode == 2)
                {
                    sqCom.CommandText = "update Products_Attachments set cnt=@cnt where id=@id";
                    sqCom.Parameters.Add("@id", SqlDbType.Int).Value = id_pa;
                    sqCom.Parameters.Add("@cnt", SqlDbType.Int).Value = cnt;
                    Database.ExecuteNonQuery(sqCom, null);
                }

                Response.Write("<script language=javascript>window.returnValue='1'; window.close();</script>");
            }
        }
        protected void dListProd_SelectedIndexChanged(object sender, EventArgs e)
        {
            lock (Database.lockObjectDB)
            {
                RefrAtt(Convert.ToInt32(dListProd.SelectedItem.Value));
            }
        }
    }
}
