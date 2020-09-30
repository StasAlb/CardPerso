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
using CardPerso.Administration;
using System.Data.SqlClient;

namespace CardPerso
{
    public partial class StorDocProductEdit : System.Web.UI.Page
    {
        DataSet ds = new DataSet();
        string res = "";
        int id_doc = 0; int id_type = 0; int mode = 0; int id_prod = 0;

        protected void Page_Load(object sender, EventArgs e)
        {
            lock (Database.lockObjectDB)
            {
                mode = Convert.ToInt32(Request.QueryString["mode"]);
                id_doc = Convert.ToInt32(Request.QueryString["id_doc"]);
                id_type = Convert.ToInt32(Request.QueryString["type_doc"]);
                id_prod = Convert.ToInt32(Request.QueryString["id_prod"]);

                if (mode == 1) Title = "Добавление";
                else Title = "Редактирование";

                if (IsPostBack)
                    return;
                ServiceClass sc = new ServiceClass();
                if (id_type == 6 && sc.UserAction(User.Identity.Name, Restrictions.Filial))
                {
                    tbCnt_New.Enabled = false;
                    tbCnt_Perso.Enabled = false;
                }
                if (id_type == 14 || id_type == 15) //расходники; на рекламные цели
                {
                    tbCnt_Brak.Enabled = false;
                    tbCnt_Perso.Enabled = false;
                }
                RefrProduct();
                if (mode == 2) ZapFields();
                BlockCnt();
            }
        }

        private void BlockCnt()
        {
            if (mode == 1)
            {
                tbCnt_New.Text = "0";
                tbCnt_Perso.Text = "0";
                tbCnt_Brak.Text = "0";
            }

            if (id_type == (int)TypeDoc.CardFromStorage || id_type == (int)TypeDoc.PinFromStorage)
            {
                tbCnt_Perso.Enabled = false;
                tbCnt_Brak.Enabled = false;
            }

            if (id_type == (int)TypeDoc.SendToFilial)
                tbCnt_Brak.Enabled = false;

            if (id_type == (int)TypeDoc.ReceiveToFilial)
                tbCnt_Brak.Enabled = false;

            if (id_type == (int)TypeDoc.PersoCard)
                tbCnt_New.Enabled = false;

            if (id_type == (int)TypeDoc.SendToBank)
            {
                tbCnt_New.Enabled = false;
                tbCnt_Perso.Enabled = false;
            }

            if (id_type == (int)TypeDoc.ReceiveToBank)
            {
                tbCnt_New.Enabled = false;
                tbCnt_Perso.Enabled = false;
            }

            if (id_type == (int)TypeDoc.DeleteBrak)
            {
                tbCnt_New.Enabled = false;
                tbCnt_Perso.Enabled = false;
            }
        }

        private void RefrProduct()
        {
            ds.Clear();

            if (mode == 1)
            {
                if (id_type == (int)TypeDoc.CardFromStorage || id_type == (int)TypeDoc.CardToStorage)
                    res = Database.ExecuteQuery(String.Format("select id_prb,(prod_name + ' (' + bank_name+')') as prod from V_ProductsBanks_T where (id_type=1) and (id_prb not in (select id_prb from Products_StorageDocs where id_doc = {0})) order by prod", id_doc), ref ds, null);
                if (id_type == (int)TypeDoc.PinFromStorage || id_type == (int)TypeDoc.PinToStorage)
                    res = Database.ExecuteQuery(String.Format("select id_prb,(prod_name + ' (' + bank_name+')') as prod from V_ProductsBanks_T where (id_type=2) and (id_prb not in (select id_prb from Products_StorageDocs where id_doc = {0})) order by prod", id_doc), ref ds, null);
                if (id_type == (int)TypeDoc.PersoCard)
                    res = Database.ExecuteQuery(String.Format("select id_prb,(prod_name + ' (' + bank_name+')') as prod from V_ProductsBanks_T where ((id_type=1) or (id_type=2)) and (id_prb not in (select id_prb from Products_StorageDocs where id_doc = {0})) order by prod", id_doc), ref ds, null);
                if (id_type == (int)TypeDoc.SendToFilial)
                    res = Database.ExecuteQuery(String.Format("select id_prb,(prod_name + ' (' + bank_name+')') as prod from V_ProductsBanks_T where (id_prb not in (select id_prb from Products_StorageDocs where id_doc = {0})) order by prod", id_doc), ref ds, null);
                if (id_type == (int)TypeDoc.SendToBank)
                    res = Database.ExecuteQuery(String.Format("select id_prb,(prod_name + ' (' + bank_name+')') as prod from V_ProductsBanks_T where ((id_type=2) or (id_type=3)) and (id_prb not in (select id_prb from Products_StorageDocs where id_doc = {0})) order by prod", id_doc), ref ds, null);
                if (id_type == (int)TypeDoc.ReceiveToBank)
                    res = Database.ExecuteQuery(String.Format("select id_prb,(prod_name + ' (' + bank_name+')') as prod from V_ProductsBanks_T where ((id_type=2) or (id_type=3)) and (id_prb not in (select id_prb from Products_StorageDocs where id_doc = {0})) order by prod", id_doc), ref ds, null);
                if (id_type == (int)TypeDoc.Expendables)
                    res = Database.ExecuteQuery(String.Format("select id_prb,(prod_name + ' (' + bank_name+')') as prod from V_ProductsBanks_T where id_type=4 and (id_prb not in (select id_prb from Products_StorageDocs where id_doc = {0})) order by prod", id_doc), ref ds, null);
                if (id_type == (int)TypeDoc.Reklama)
                    res = Database.ExecuteQuery(String.Format("select id_prb,(prod_name + ' (' + bank_name+')') as prod from V_ProductsBanks_T where (id_prb not in (select id_prb from Products_StorageDocs where id_doc = {0})) order by prod", id_doc), ref ds, null);
            }
            else
                res = Database.ExecuteQuery(String.Format("select id_prb,(prod_name + ' (' + bank_name+')') as prod from V_ProductsBanks_T order by prod", id_doc), ref ds, null);

            dListProd.DataSource = ds.Tables[0];
            dListProd.DataTextField = "prod";
            dListProd.DataValueField = "id_prb";
            dListProd.DataBind();
        }

        private void ZapFields()
        {
            dListProd.Enabled = false;

            ds.Clear();
            res = Database.ExecuteQuery(String.Format("select * from Products_StorageDocs where id={0}", id_prod), ref ds, null);
            dListProd.SelectedIndex = dListProd.Items.IndexOf(dListProd.Items.FindByValue(ds.Tables[0].Rows[0]["id_prb"].ToString()));
            tbCnt_New.Text = ds.Tables[0].Rows[0]["cnt_new"].ToString();
            tbCnt_Perso.Text = ds.Tables[0].Rows[0]["cnt_perso"].ToString();
            tbCnt_Brak.Text = ds.Tables[0].Rows[0]["cnt_brak"].ToString();
        }

        protected void bSave_Click(object sender, ImageClickEventArgs e)
        {
            lock (Database.lockObjectDB)
            {
                int cnt_new, cnt_perso, cnt_brak = 0;

                if (tbCnt_New.Text == "")
                {
                    lbInform.Text = "Введите кол-во продукции";
                    tbCnt_New.Focus();
                    return;
                }
                else
                {
                    try
                    {
                        cnt_new = Convert.ToInt32(tbCnt_New.Text);
                    }
                    catch
                    {
                        lbInform.Text = "Количество продукции должно быть целым";
                        tbCnt_New.Focus();
                        return;
                    }
                }

                if (tbCnt_Perso.Text == "")
                {
                    lbInform.Text = "Введите кол-во продукции";
                    tbCnt_Perso.Focus();
                    return;
                }
                else
                {
                    try
                    {
                        cnt_perso = Convert.ToInt32(tbCnt_Perso.Text);
                    }
                    catch
                    {
                        lbInform.Text = "Количество продукции должно быть целым";
                        tbCnt_Perso.Focus();
                        return;
                    }
                }

                if (tbCnt_Brak.Text == "")
                {
                    lbInform.Text = "Введите кол-во продукции";
                    tbCnt_Brak.Focus();
                    return;
                }
                else
                {
                    try
                    {
                        cnt_brak = Convert.ToInt32(tbCnt_Brak.Text);
                    }
                    catch
                    {
                        lbInform.Text = "Количество продукции должно быть целым";
                        tbCnt_Brak.Focus();
                        return;
                    }
                }
                SqlCommand sqCom = new SqlCommand();
                //добавление
                if (mode == 1)
                {
                    if (dListProd.SelectedIndex < 0)
                    {
                        lbInform.Text = "Не выбран тип продукции";
                        return;
                    }
                    if (!Database.CheckProductInStorDoc(id_doc, Convert.ToInt32(dListProd.SelectedItem.Value), null))
                    {
                        lbInform.Text = "Такая продукция уже есть в документе";
                        return;
                    }

                    sqCom.CommandText = "insert into Products_StorageDocs (id_doc,id_prb,cnt_new,cnt_perso,cnt_brak) values(@id_doc,@id_prb,@cnt_new,@cnt_perso,@cnt_brak)";
                    sqCom.Parameters.Add("@id_doc", SqlDbType.Int).Value = id_doc;
                    sqCom.Parameters.Add("@id_prb", SqlDbType.Int).Value = Convert.ToInt32(dListProd.SelectedItem.Value);
                    sqCom.Parameters.Add("@cnt_new", SqlDbType.Int).Value = cnt_new;
                    sqCom.Parameters.Add("@cnt_perso", SqlDbType.Int).Value = cnt_perso;
                    sqCom.Parameters.Add("@cnt_brak", SqlDbType.Int).Value = cnt_brak;
                    Database.ExecuteNonQuery(sqCom, null);

                    lbInform.Text = "Продукция " + dListProd.SelectedItem.Text + " добалена в документ. Обновление после закрытия формы.";
                    RefrProduct();
                    BlockCnt();
                }

                if (mode == 2)
                {
                    sqCom.CommandText = "update Products_StorageDocs set cnt_new=@cnt_new,cnt_perso=@cnt_perso,cnt_brak=@cnt_brak where id=@id";

                    sqCom.Parameters.Add("@id", SqlDbType.Int).Value = id_prod;
                    sqCom.Parameters.Add("@cnt_new", SqlDbType.Int).Value = cnt_new;
                    sqCom.Parameters.Add("@cnt_perso", SqlDbType.Int).Value = cnt_perso;
                    sqCom.Parameters.Add("@cnt_brak", SqlDbType.Int).Value = cnt_brak;
                    Database.ExecuteNonQuery(sqCom, null);

                    Response.Write("<script language=javascript>window.returnValue='1'; window.close();</script>");
                }
            }
        }
    }
}
