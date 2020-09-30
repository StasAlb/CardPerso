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
    public partial class PurchaseDogEdit : System.Web.UI.Page
    {
        DataSet ds = new DataSet();
        string res = "";
        int mode = 0; int id = 0;
        CardPerso.Administration.ServiceClass sc = new CardPerso.Administration.ServiceClass();
        protected void Page_Load(object sender, EventArgs e)
        {
            lock (Database.lockObjectDB)
            {
                mode = Convert.ToInt32(Request.QueryString["mode"]);
                id = Convert.ToInt32(Request.QueryString["id"]);

                if (!IsPostBack)
                {
                    ZapCombo();

                    if (mode == 2)
                    {
                        Title = "Редактирование";
                        ZapFields();
                    }
                    else
                        Title = "Добавление";

                    tbNumber.Focus();
                }
            }
        }

        private void ZapCombo()
        {
            ds.Clear();
            res = Database.ExecuteQuery("select id,name from Suppliers", ref ds, null);
            dListSup.Items.Add(new ListItem("", "-1"));
            for (int i = 0; i < ds.Tables[0].Rows.Count; i++)
                dListSup.Items.Add(new ListItem(ds.Tables[0].Rows[i]["name"].ToString(), ds.Tables[0].Rows[i]["id"].ToString()));

            dListSup.SelectedIndex = 0;

            ds.Clear();
            res = Database.ExecuteQuery("select id,name from Manufacturers", ref ds, null);
            dListManuf.Items.Add(new ListItem("", "-1"));
            for (int i = 0; i < ds.Tables[0].Rows.Count; i++)
                dListManuf.Items.Add(new ListItem(ds.Tables[0].Rows[i]["name"].ToString(), ds.Tables[0].Rows[i]["id"].ToString()));

            dListManuf.SelectedIndex = 0;
        }

        private void ZapFields()
        {
            ds.Clear();
            res = Database.ExecuteQuery(String.Format("select * from PurchDogs where id={0}",id), ref ds, null);

            tbNumber.Text = ds.Tables[0].Rows[0]["number_dog"].ToString();
            tbData.Text = String.Format("{0:d}", ds.Tables[0].Rows[0]["date_dog"]);
            tbDataSt.Text = String.Format("{0:d}", ds.Tables[0].Rows[0]["date_stor"]);
            dListSup.SelectedIndex = dListSup.Items.IndexOf(dListSup.Items.FindByValue(ds.Tables[0].Rows[0]["id_sup"].ToString()));
            dListManuf.SelectedIndex = dListManuf.Items.IndexOf(dListManuf.Items.FindByValue(ds.Tables[0].Rows[0]["id_manuf"].ToString()));
            tbDataR.Text = String.Format("{0:d}", ds.Tables[0].Rows[0]["date_record"]);
            tbComment.Text = ds.Tables[0].Rows[0]["comment"].ToString();
        }

        protected void bSave_Click(object sender, ImageClickEventArgs e)
        {
            lock (Database.lockObjectDB)
            {
                if (tbNumber.Text == "")
                {
                    lbInform.Text = "Введите номер договора";
                    tbNumber.Focus();
                    return;
                }

                if (tbData.Text != "")
                {
                    try
                    {
                        Convert.ToDateTime(tbData.Text);
                    }
                    catch
                    {
                        lbInform.Text = "Неправильно введена дата договора";
                        tbData.Focus();
                        return;
                    }
                }
                if (tbDataSt.Text != "")
                {
                    try
                    {
                        Convert.ToDateTime(tbDataSt.Text);
                    }
                    catch
                    {
                        lbInform.Text = "Неправильно введена дата поступления";
                        tbDataSt.Focus();
                        return;
                    }
                }
                if (tbDataR.Text != "")
                {
                    try
                    {
                        Convert.ToDateTime(tbDataR.Text);
                    }
                    catch
                    {
                        lbInform.Text = "Неправильно введена дата выписки";
                        tbDataR.Focus();
                        return;
                    }
                }

                SqlCommand sqCom = new SqlCommand();
                int id_list = 0;
                ////new
                if (mode == 1)
                    sqCom.CommandText = "insert into PurchDogs (number_dog,date_dog,date_stor,id_sup,id_manuf,date_record,comment) values(@number_dog,@date_dog,@date_stor,@id_sup,@id_manuf,@date_record,@comment)";
                //edit
                if (mode == 2)
                {
                    sqCom.CommandText = "update PurchDogs set number_dog=@number_dog,date_dog=@date_dog,date_stor=@date_stor,id_sup=@id_sup,id_manuf=@id_manuf,date_record=@date_record,comment=@comment where id=@id";
                    sqCom.Parameters.Add("@id", SqlDbType.Int).Value = id;
                }
                sqCom.Parameters.Add("@number_dog", SqlDbType.VarChar, 20).Value = tbNumber.Text;
                if (tbData.Text != "")
                    sqCom.Parameters.Add("@date_dog", SqlDbType.DateTime).Value = Convert.ToDateTime(tbData.Text);
                else
                    sqCom.Parameters.Add("@date_dog", SqlDbType.DateTime).Value = DBNull.Value;
                if (tbDataSt.Text != "")
                    sqCom.Parameters.Add("@date_stor", SqlDbType.DateTime).Value = Convert.ToDateTime(tbDataSt.Text);
                else
                    sqCom.Parameters.Add("@date_stor", SqlDbType.DateTime).Value = DBNull.Value;

                id_list = Convert.ToInt32(dListSup.SelectedItem.Value);

                if (id_list != -1)
                    sqCom.Parameters.Add("@id_sup", SqlDbType.Int).Value = id_list;
                else
                    sqCom.Parameters.Add("@id_sup", SqlDbType.Int).Value = DBNull.Value;

                id_list = Convert.ToInt32(dListManuf.SelectedItem.Value);

                if (id_list != -1)
                    sqCom.Parameters.Add("@id_manuf", SqlDbType.Int).Value = id_list;
                else
                    sqCom.Parameters.Add("@id_manuf", SqlDbType.Int).Value = DBNull.Value;

                if (tbDataR.Text != "")
                    sqCom.Parameters.Add("@date_record", SqlDbType.DateTime).Value = Convert.ToDateTime(tbDataR.Text);
                else
                    sqCom.Parameters.Add("@date_record", SqlDbType.DateTime).Value = DBNull.Value;

                sqCom.Parameters.Add("@comment", SqlDbType.VarChar, 150).Value = tbComment.Text;

                res = Database.ExecuteNonQuery(sqCom, null);
                if (mode == 1)
                    Database.Log(sc.UserGuid(User.Identity.Name), String.Format("Добавлен договор по закупкам {0}", tbNumber.Text), null);
                if (mode == 2)
                    Database.Log(sc.UserGuid(User.Identity.Name), String.Format("Отредактирован договор по закупкам {0}", tbNumber.Text), null);
                Response.Write("<script language=javascript>window.returnValue='1'; window.close();</script>");
            }
        }
    }
}
