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
    public partial class CatalogEdit : System.Web.UI.Page
    {
        DataSet ds = new DataSet();
        string res = "";
        ServiceClass sc = new ServiceClass();

        protected void Page_Load(object sender, EventArgs e)
        {
            lock (Database.lockObjectDB)
            {
                if (!sc.UserAction(User.Identity.Name, Restrictions.LibraryEdit))
                    Response.Redirect("~\\Account\\Restricted.aspx", true);
                if (!IsPostBack)
                {
                    if (Request.QueryString["mode"] == "2")
                        ZapFields();
                    tbName.Focus();
                }
            }
        }

        private void ZapFields()
        {
            int id = Convert.ToInt32(Request.QueryString["id"]);
            lbInform.Text = "";

            ds.Clear();

            if (Request.QueryString["type"] == "courier")
                res = Database.ExecuteQuery(String.Format("select * from Couriers where id={0}",id), ref ds, null);
            if (Request.QueryString["type"] == "bank")
                res = Database.ExecuteQuery(String.Format("select * from Banks where id={0}", id), ref ds, null);
            if (Request.QueryString["type"] == "manufacturer")
                res = Database.ExecuteQuery(String.Format("select * from Manufacturers where id={0}", id), ref ds, null);
            if (Request.QueryString["type"] == "supplier")
                res = Database.ExecuteQuery(String.Format("select * from Suppliers where id={0}", id), ref ds, null);
            if (Request.QueryString["type"] == "expendable")
                res = Database.ExecuteQuery(String.Format("select * from Expendables where id={0}", id), ref ds, null);

            tbName.Text = ds.Tables[0].Rows[0]["name"].ToString();
        }

        protected void bSave_Click(object sender, ImageClickEventArgs e)
        {
            lock (Database.lockObjectDB)
            {
                if (tbName.Text == "")
                {
                    lbInform.Text = "Введите наименование";
                    tbName.Focus();
                    return;
                }

                SqlCommand sqCom = new SqlCommand();

                if (Request.QueryString["type"] == "courier")
                {
                    if (Request.QueryString["mode"] == "1")
                        sqCom.CommandText = "insert into Couriers (name) values(@name)";
                    if (Request.QueryString["mode"] == "2")
                    {
                        sqCom.CommandText = "update Couriers set name=@name where id=@id";
                        sqCom.Parameters.Add("@id", SqlDbType.Int).Value = Convert.ToInt32(Request.QueryString["id"]);
                    }

                    sqCom.Parameters.Add("@name", SqlDbType.VarChar, 50).Value = tbName.Text;
                }

                if (Request.QueryString["type"] == "bank")
                {
                    if (Request.QueryString["mode"] == "1")
                        sqCom.CommandText = "insert into Banks (name) values(@name)";
                    if (Request.QueryString["mode"] == "2")
                    {
                        sqCom.CommandText = "update Banks set name=@name where id=@id";
                        sqCom.Parameters.Add("@id", SqlDbType.Int).Value = Convert.ToInt32(Request.QueryString["id"]);
                    }

                    sqCom.Parameters.Add("@name", SqlDbType.VarChar, 50).Value = tbName.Text;
                }

                if (Request.QueryString["type"] == "manufacturer")
                {
                    if (Request.QueryString["mode"] == "1")
                        sqCom.CommandText = "insert into Manufacturers (name) values(@name)";
                    if (Request.QueryString["mode"] == "2")
                    {
                        sqCom.CommandText = "update Manufacturers set name=@name where id=@id";
                        sqCom.Parameters.Add("@id", SqlDbType.Int).Value = Convert.ToInt32(Request.QueryString["id"]);
                    }

                    sqCom.Parameters.Add("@name", SqlDbType.VarChar, 100).Value = tbName.Text;
                }

                if (Request.QueryString["type"] == "supplier")
                {
                    if (Request.QueryString["mode"] == "1")
                        sqCom.CommandText = "insert into Suppliers (name) values(@name)";
                    if (Request.QueryString["mode"] == "2")
                    {
                        sqCom.CommandText = "update Suppliers set name=@name where id=@id";
                        sqCom.Parameters.Add("@id", SqlDbType.Int).Value = Convert.ToInt32(Request.QueryString["id"]);
                    }

                    sqCom.Parameters.Add("@name", SqlDbType.VarChar, 100).Value = tbName.Text;
                }
                if (Request.QueryString["type"] == "expendable")
                {
                    if (Request.QueryString["mode"] == "1")
                        sqCom.CommandText = "insert into expendables (name) values (@name)";
                    if (Request.QueryString["mode"] == "2")
                    {
                        sqCom.CommandText = "update expendables set name=@name where id=@id";
                        sqCom.Parameters.Add("@id", SqlDbType.Int).Value = Convert.ToInt32(Request.QueryString["id"]);
                    }
                    sqCom.Parameters.Add("@name", SqlDbType.VarChar, 50).Value = tbName.Text;
                }
                Database.ExecuteNonQuery(sqCom, null);

                Response.Write("<script language=javascript>window.returnValue='1'; window.close();</script>");
            }
        }
    }
}
