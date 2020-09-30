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
using CardPerso.Administration;

namespace CardPerso
{
    public partial class ProductEdit : System.Web.UI.Page
    {
        DataSet ds = new DataSet();
        int mode = 0; int id_prod = 0; int id_prb=0;
        string res = "";

        protected void Page_Load(object sender, EventArgs e)
        {
            lock (Database.lockObjectDB)
            {
                ServiceClass sc = new ServiceClass();
                if (!sc.UserAction(User.Identity.Name, Restrictions.LibraryEdit))
                    Response.Redirect("~\\Account\\Restricted.aspx", true);

                mode = Convert.ToInt32(Request.QueryString["mode"]);
                id_prod = Convert.ToInt32(Request.QueryString["id_prod"]);
                id_prb = Convert.ToInt32(Request.QueryString["id_prb"]);

                if (mode == 1) Title = "Добавление продукта";
                if (mode == 2) Title = "Редактирование";
                if (mode == 3) Title = "Привязка банка";

                if (!IsPostBack)
                {
                    ZapCombo();
                    if (mode == 2 || mode == 3) ZapFields();
                    if (mode == 3) pProduct.Enabled = false;
                }
            }
        }

        private void ZapCombo()
        {
            ds.Clear();
            res = Database.ExecuteQuery("select id,name from TypeProducts", ref ds, null);
            for (int i = 0; i < ds.Tables[0].Rows.Count; i++)
                dListKind.Items.Add(new ListItem(ds.Tables[0].Rows[i]["name"].ToString(), ds.Tables[0].Rows[i]["id"].ToString()));
            dListKind.SelectedIndex = 0;

            ds.Clear();
            res = Database.ExecuteQuery("select id,name from Banks", ref ds, null);
            for (int i = 0; i < ds.Tables[0].Rows.Count; i++)
                dListBank.Items.Add(new ListItem(ds.Tables[0].Rows[i]["name"].ToString(), ds.Tables[0].Rows[i]["id"].ToString()));
            dListBank.SelectedIndex = 0;

            tbName.Focus();
        }

        private void ZapFields()
        {
            ds.Clear();
            res = Database.ExecuteQuery(String.Format("select * from Products where id={0}",id_prod), ref ds, null);

            tbName.Text = ds.Tables[0].Rows[0]["name"].ToString();
            dListKind.SelectedIndex = dListKind.Items.IndexOf(dListKind.Items.FindByValue(ds.Tables[0].Rows[0]["id_type"].ToString()));
            tbType.Text = ds.Tables[0].Rows[0]["prefix_ow"].ToString();

            if (mode == 2)
            {
                ds.Clear();
                res = Database.ExecuteQuery(String.Format("select * from Products_Banks where id={0}", id_prb), ref ds, null);

                dListBank.SelectedIndex = dListBank.Items.IndexOf(dListBank.Items.FindByValue(ds.Tables[0].Rows[0]["id_bank"].ToString()));
                tbBin.Text = ds.Tables[0].Rows[0]["bin"].ToString();
                tbPrefix.Text = ds.Tables[0].Rows[0]["prefix_file"].ToString();
                tbMinimum.Text = ds.Tables[0].Rows[0]["min_cnt"].ToString();
                cbWrapping.Checked = (ds.Tables[0].Rows[0]["wrapping"] != DBNull.Value && Convert.ToBoolean(ds.Tables[0].Rows[0]["wrapping"]));
                cbInformProduction.Checked = (ds.Tables[0].Rows[0]["inform_production"] != DBNull.Value && Convert.ToBoolean(ds.Tables[0].Rows[0]["inform_production"]));
            }
        }

        protected void bSave_Click(object sender, ImageClickEventArgs e)
        {
            lock (Database.lockObjectDB)
            {
                if (tbName.Text == "")
                {
                    lbInform.Text = "Введите наименование продукта";
                    tbName.Focus();
                    return;
                }

                if (dListBank.SelectedIndex < 0)
                {
                    lbInform.Text = "Выберите банк";
                    dListBank.Focus();
                    return;
                }

                if (tbMinimum.Text != "")
                {
                    try
                    {
                        Convert.ToInt32(tbMinimum.Text);
                    }
                    catch
                    {
                        lbInform.Text = "Критический минимум должно быть целым";
                        tbMinimum.Focus();
                        return;
                    }
                }

                if (mode == 2 || mode == 3)
                    if (!Database.CheckBankInProduct(id_prb, id_prod, Convert.ToInt32(dListBank.SelectedItem.Value), null))
                    {
                        lbInform.Text = "Такой банк уже привязан";
                        dListBank.Focus();
                        return;
                    }

                SqlCommand sqCom = new SqlCommand();

                if (mode == 1)
                {
                    sqCom.CommandText = "insert into Products (name,id_type,prefix_ow) values(@name,@id_type,@prefix_ow) select @@identity as lid";
                    sqCom.Parameters.Add("@name", SqlDbType.VarChar, 100).Value = tbName.Text.Trim();
                    sqCom.Parameters.Add("@id_type", SqlDbType.Int).Value = dListKind.SelectedItem.Value;
                    sqCom.Parameters.Add("@prefix_ow", SqlDbType.VarChar, 10).Value = tbType.Text;
                    object obj = null;
                    Database.ExecuteScalar(sqCom, ref obj, null);
                    int lid = Convert.ToInt32(obj);
                    sqCom.Parameters.Clear();
                    sqCom.CommandText = "update Products set id_sort=@id_sort where id=@id";
                    sqCom.Parameters.Add("@id", SqlDbType.Int).Value = lid;
                    sqCom.Parameters.Add("@id_sort", SqlDbType.Int).Value = lid;
                    Database.ExecuteNonQuery(sqCom, null);
                    sqCom.Parameters.Clear();

                    sqCom.CommandText = "insert into Products_Banks (id_prod,id_bank,bin,prefix_file,min_cnt,wrapping,inform_production) values(@id_prod,@id_bank,@bin,@prefix_file,@min_cnt,@wrapping,@inform)";
                    sqCom.Parameters.Add("@id_prod", SqlDbType.Int).Value = lid;
                    sqCom.Parameters.Add("@id_bank", SqlDbType.Int).Value = dListBank.SelectedItem.Value;
                    sqCom.Parameters.Add("@bin", SqlDbType.VarChar, 20).Value = tbBin.Text;
                    sqCom.Parameters.Add("@prefix_file", SqlDbType.VarChar, 10).Value = tbPrefix.Text;
                    if (tbMinimum.Text == "")
                        sqCom.Parameters.Add("@min_cnt", SqlDbType.Int).Value = DBNull.Value;
                    else
                        sqCom.Parameters.Add("@min_cnt", SqlDbType.Int).Value = Convert.ToInt32(tbMinimum.Text);
                    sqCom.Parameters.Add("@wrapping", SqlDbType.Bit).Value = (cbWrapping.Checked) ? 1 : 0;
                    sqCom.Parameters.Add("@inform", SqlDbType.Bit).Value = (cbInformProduction.Checked) ? 1 : 0;
                    Database.ExecuteNonQuery(sqCom, null);
                }

                if (mode == 2)
                {
                    sqCom.CommandText = "update Products set name=@name,id_type=@id_type,prefix_ow=@prefix_ow where id=@id";
                    sqCom.Parameters.Add("@id", SqlDbType.Int).Value = id_prod;
                    sqCom.Parameters.Add("@name", SqlDbType.VarChar, 100).Value = tbName.Text.Trim();
                    sqCom.Parameters.Add("@id_type", SqlDbType.Int).Value = dListKind.SelectedItem.Value;
                    sqCom.Parameters.Add("@prefix_ow", SqlDbType.VarChar, 10).Value = tbType.Text;
                    Database.ExecuteNonQuery(sqCom, null);
                    sqCom.Parameters.Clear();
                    sqCom.CommandText = "update Products_Banks set id_bank=@id_bank,bin=@bin,prefix_file=@prefix_file,min_cnt=@min_cnt,wrapping=@wrapping, inform_production=@inform where id=@id";
                    sqCom.Parameters.Add("@id", SqlDbType.Int).Value = id_prb;
                    sqCom.Parameters.Add("@id_bank", SqlDbType.Int).Value = dListBank.SelectedItem.Value;
                    sqCom.Parameters.Add("@bin", SqlDbType.VarChar, 20).Value = tbBin.Text;
                    sqCom.Parameters.Add("@prefix_file", SqlDbType.VarChar, 10).Value = tbPrefix.Text;
                    if (tbMinimum.Text == "")
                        sqCom.Parameters.Add("@min_cnt", SqlDbType.Int).Value = DBNull.Value;
                    else
                        sqCom.Parameters.Add("@min_cnt", SqlDbType.Int).Value = Convert.ToInt32(tbMinimum.Text);
                    sqCom.Parameters.Add("@wrapping", SqlDbType.Bit).Value = (cbWrapping.Checked) ? 1 : 0;
                    sqCom.Parameters.Add("@inform", SqlDbType.Bit).Value = (cbInformProduction.Checked) ? 1 : 0;
                    Database.ExecuteNonQuery(sqCom, null);
                }

                if (mode == 3)
                {
                    sqCom.CommandText = "insert into Products_Banks (id_prod,id_bank,bin,prefix_file,min_cnt,wrapping,inform_production) values(@id_prod,@id_bank,@bin,@prefix_file,@min_cnt,@wrapping,@inform)";
                    sqCom.Parameters.Add("@id_prod", SqlDbType.Int).Value = id_prod;
                    sqCom.Parameters.Add("@id_bank", SqlDbType.Int).Value = dListBank.SelectedItem.Value;
                    sqCom.Parameters.Add("@bin", SqlDbType.VarChar, 20).Value = tbBin.Text;
                    sqCom.Parameters.Add("@prefix_file", SqlDbType.VarChar, 10).Value = tbPrefix.Text;
                    if (tbMinimum.Text == "")
                        sqCom.Parameters.Add("@min_cnt", SqlDbType.Int).Value = DBNull.Value;
                    else
                        sqCom.Parameters.Add("@min_cnt", SqlDbType.Int).Value = Convert.ToInt32(tbMinimum.Text);
                    sqCom.Parameters.Add("@wrapping", SqlDbType.Bit).Value = (cbWrapping.Checked) ? 1 : 0;
                    sqCom.Parameters.Add("@inform", SqlDbType.Bit).Value = (cbInformProduction.Checked) ? 1 : 0;
                    Database.ExecuteNonQuery(sqCom, null);
                }

                Response.Write("<script language=javascript>window.returnValue='1'; window.close();</script>");
            }
        }

    }
}
