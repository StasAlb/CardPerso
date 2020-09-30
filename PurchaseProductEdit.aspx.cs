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
    public partial class PurchaseProductEdit : System.Web.UI.Page
    {
        DataSet ds = new DataSet();
        string res = "";
        int id_dog = 0; int mode = 0; int id_prod = 0;
        int cnt_last = 0;

        protected void Page_Load(object sender, EventArgs e)
        {
            lock (Database.lockObjectDB)
            {
                mode = Convert.ToInt32(Request.QueryString["mode"]);
                id_dog = Convert.ToInt32(Request.QueryString["id_dog"]);
                id_prod = Convert.ToInt32(Request.QueryString["id_prod"]);

                if (mode == 1) Title = "Добавление";
                else Title = "Редактирование";

                if (!IsPostBack)
                {
                    ZapCombo();
                    if (mode == 2) ZapFields();
                }
            }
        }

        private void ZapCombo()
        {
            ds.Clear();
            if (mode==1)
                res = Database.ExecuteQuery(String.Format("select id_prb,(prod_name + ' (' + bank_name+')') as prod from V_ProductsBanks_T where (id_prb not in (select id_prb from Products_PurchDogs where id_dog = {0})) order by id_sort",id_dog), ref ds, null);
            else
                res = Database.ExecuteQuery("select id_prb,(prod_name + ' (' + bank_name+')') as prod from V_ProductsBanks_T order by id_sort", ref ds, null);

            dListProd.DataSource = ds.Tables[0];
            dListProd.DataTextField = "prod";
            dListProd.DataValueField = "id_prb";
            dListProd.DataBind();
        }

        private void ZapFields()
        {
            dListProd.Enabled = false;

            ds.Clear();
            res = Database.ExecuteQuery(String.Format("select * from Products_PurchDogs where id={0}", id_prod), ref ds, null);
            dListProd.SelectedIndex = dListProd.Items.IndexOf(dListProd.Items.FindByValue(ds.Tables[0].Rows[0]["id_prb"].ToString()));
            tbCount.Text = ds.Tables[0].Rows[0]["cnt"].ToString();
            tbPrice.Text = String.Format("{0:f2}", ds.Tables[0].Rows[0]["price"]);
            lbCount.Text = tbCount.Text;
        }

        protected void bSave_Click(object sender, ImageClickEventArgs e)
        {
            lock (Database.lockObjectDB)
            {
                int cnt = 0;
                double price = 0;

                if (tbCount.Text == "")
                {
                    lbInform.Text = "Введите количество товара";
                    tbCount.Focus();
                    return;
                }
                else
                {
                    try
                    {
                        cnt = Convert.ToInt32(tbCount.Text);
                        if (mode == 2) cnt_last = Convert.ToInt32(lbCount.Text) - cnt;
                    }
                    catch
                    {
                        lbInform.Text = "Количество товара должно быть целым";
                        tbCount.Focus();
                        return;
                    }
                }

                if (tbPrice.Text != "")
                {
                    try
                    {
                        price = Convert.ToDouble(String.Format("{0:f2}", Convert.ToDouble(tbPrice.Text)));
                    }
                    catch
                    {
                        lbInform.Text = "Цена товара должна быть десятичной";
                        tbPrice.Focus();
                        return;
                    }
                }

                SqlCommand sqCom = new SqlCommand();
                //добавление
                if (mode == 1)
                {
                    if (!Database.CheckProductInPurchDog(id_dog, Convert.ToInt32(dListProd.SelectedItem.Value), null))
                    {
                        lbInform.Text = "Такой товар уже есть в договоре";
                        return;
                    }

                    sqCom.CommandText = "insert into Products_PurchDogs (id_dog,id_prb,cnt,price,summa) values(@id_dog,@id_prb,@cnt,@price,@summa)";
                    sqCom.Parameters.Add("@id_dog", SqlDbType.Int).Value = id_dog;
                    sqCom.Parameters.Add("@id_prb", SqlDbType.Int).Value = Convert.ToInt32(dListProd.SelectedItem.Value);
                    sqCom.Parameters.Add("@cnt", SqlDbType.Int).Value = cnt;
                    sqCom.Parameters.Add("@price", SqlDbType.Money).Value = price;
                    sqCom.Parameters.Add("@summa", SqlDbType.Money).Value = cnt * price;
                    res = Database.ExecuteNonQuery(sqCom, null);

                    Database.StorageP(cnt, Convert.ToInt32(dListProd.SelectedItem.Value), "new", null);

                    lbInform.Text = "Продукция " + dListProd.SelectedItem.Text + " добавлена в договор. Обновление после закрытия формы.";
                    ZapCombo();
                    tbCount.Text = "";
                    tbPrice.Text = "";
                }

                if (mode == 2)
                {
                    if (cnt_last > 0)
                    {
                        if (cnt_last > Database.CntStorage(Convert.ToInt32(dListProd.SelectedItem.Value), "new", null))
                        {
                            lbInform.Text = "Нет достаточного количества продукции в хранилище.";
                            return;
                        }
                        Database.StorageM(cnt_last, Convert.ToInt32(dListProd.SelectedItem.Value), "new", null);
                    }
                    else
                    {
                        if (cnt_last < 0)
                        {
                            cnt_last = -cnt_last;
                            Database.StorageP(cnt_last, Convert.ToInt32(dListProd.SelectedItem.Value), "new", null);
                        }
                    }

                    sqCom.CommandText = "update Products_PurchDogs set cnt=@cnt,price=@price,summa=@summa where id=@id";

                    sqCom.Parameters.Add("@id", SqlDbType.Int).Value = id_prod;
                    sqCom.Parameters.Add("@cnt", SqlDbType.Int).Value = cnt;
                    sqCom.Parameters.Add("@price", SqlDbType.Money).Value = price;
                    sqCom.Parameters.Add("@summa", SqlDbType.Money).Value = cnt * price;
                    res = Database.ExecuteNonQuery(sqCom, null);

                    Response.Write("<script language=javascript>window.returnValue='1'; window.close();</script>");
                }
            }
        }
    }
}
