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
    public partial class BranchEdit : System.Web.UI.Page
    {
        DataSet ds = new DataSet();
        string res = "";
        ServiceClass sc = new ServiceClass();
        protected void Page_Load(object sender, EventArgs e)
        {
            if (!sc.UserAction(User.Identity.Name, Restrictions.LibraryEdit))
                Response.Redirect("~\\Account\\Restricted.aspx", true);
            
            if (IsPostBack)
                return;
            if (FuncClass.ClientType != ClientType.AkBars)
            {
                pAkBarsProperties.Visible = false;
            }



            lock (Database.lockObjectDB)
            {
                if (Request.QueryString["mode"] == "2")
                    ZapFields();
                else cbIsolated.Visible = true;
                tbKodBank.Focus();
            }
        }

        private void ZapFields()
        {
            int id = Convert.ToInt32(Request.QueryString["id"]);
            lbInform.Text = "";

            ds.Clear();
            res = Database.ExecuteQuery(String.Format("select * from Branchs where id={0}",id), ref ds, null);

            if (ds.Tables.Count == 0 || ds.Tables[0].Rows.Count == 0)
                return;
            tbKodBank.Text = ds.Tables[0].Rows[0]["ident_bank"].ToString();
            tbKodDep.Text = ds.Tables[0].Rows[0]["ident_dep"].ToString();
            tbDep.Text = ds.Tables[0].Rows[0]["department"].ToString();
            tbAdress.Text = ds.Tables[0].Rows[0]["adress"].ToString();
            tbPeople.Text = ds.Tables[0].Rows[0]["people"].ToString();
            tbEmail.Text = ds.Tables[0].Rows[0]["email"].ToString();
            chHead.Checked = (ds.Tables[0].Rows[0]["is_head"] == DBNull.Value) ? false : Convert.ToBoolean(ds.Tables[0].Rows[0]["is_head"]);
            cbRKC.Checked = (ds.Tables[0].Rows[0]["is_rkc"] == DBNull.Value) ? false : Convert.ToBoolean(ds.Tables[0].Rows[0]["is_rkc"]);
            cbTrans.Checked = (ds.Tables[0].Rows[0]["is_trans"] == DBNull.Value) ? false : Convert.ToBoolean(ds.Tables[0].Rows[0]["is_trans"]);
            cbIsolated.Checked = (ds.Tables[0].Rows[0]["isolated"] == DBNull.Value) ? false : Convert.ToBoolean(ds.Tables[0].Rows[0]["isolated"]);
            cbIsolated.Visible = (ds.Tables[0].Rows[0]["id_parent"] == DBNull.Value || 
                Convert.ToInt32(ds.Tables[0].Rows[0]["id_parent"])==0) ? false : true; 
        }

        protected void bSave_Click(object sender, ImageClickEventArgs e)
        {
            lock (Database.lockObjectDB)
            {
                if (tbDep.Text == "")
                {
                    lbInform.Text = "Введите наименование";
                    tbDep.Focus();
                    return;
                }

                SqlCommand sqCom = new SqlCommand();

                if (Request.QueryString["mode"] == "1")
                {
                    sqCom.CommandText = "insert into Branchs (id_parent,ident_bank,ident_dep,department,email,adress,people,is_head,is_rkc,is_trans,isolated) values(@id_parent,@ident_bank,@ident_dep,@department,@email,@adress,@people,@is_head,@is_rkc,@is_trans,@isolated)";
                    sqCom.Parameters.Add("@id_parent", SqlDbType.Int).Value = 0;
                }
                if (Request.QueryString["mode"] == "2")
                {
                    sqCom.CommandText = "update Branchs set ident_bank=@ident_bank,ident_dep=@ident_dep,department=@department,email=@email,adress=@adress,people=@people,is_head=@is_head,is_rkc=@is_rkc, is_trans=@is_trans, isolated=@isolated where id=@id";
                    sqCom.Parameters.Add("@id", SqlDbType.Int).Value = Convert.ToInt32(Request.QueryString["id"]);
                }
                sqCom.Parameters.Add("@ident_bank", SqlDbType.VarChar, 15).Value = tbKodBank.Text;
                sqCom.Parameters.Add("@ident_dep", SqlDbType.VarChar, 15).Value = tbKodDep.Text;
                sqCom.Parameters.Add("@department", SqlDbType.VarChar, 100).Value = tbDep.Text;
                sqCom.Parameters.Add("@adress", SqlDbType.VarChar, 150).Value = tbAdress.Text;
                sqCom.Parameters.Add("@people", SqlDbType.VarChar, 50).Value = tbPeople.Text;
                sqCom.Parameters.Add("@email", SqlDbType.VarChar, 30).Value = tbEmail.Text;
                sqCom.Parameters.Add("@is_head", SqlDbType.Bit).Value = chHead.Checked;
                sqCom.Parameters.Add("@is_rkc", SqlDbType.Bit).Value = cbRKC.Checked;
                sqCom.Parameters.Add("@is_trans", SqlDbType.Bit).Value = cbTrans.Checked;

                sqCom.Parameters.Add("@isolated", SqlDbType.Bit).Value = cbIsolated.Checked;
                Database.ExecuteNonQuery(sqCom, null);
                Response.Write("<script language=javascript>window.returnValue='1'; window.close();</script>");
            }
        }

    }
}
