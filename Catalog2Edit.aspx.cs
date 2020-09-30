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
    public partial class Catalog2Edit : System.Web.UI.Page
    {
        DataSet ds = new DataSet();
        string res = "";
        ServiceClass sc = new ServiceClass();

        int branch_main_filial = 0;
        int current_branch = 0;
        bool perso = false;
       
        protected void Page_Load(object sender, EventArgs e)
        {
           
            lock (Database.lockObjectDB)
            {
                if (!sc.UserAction(User.Identity.Name, Restrictions.LibraryOrgEdit))
                    Response.Redirect("~\\Account\\Restricted.aspx", true);
                
                perso = sc.UserAction(User.Identity.Name, Restrictions.Perso);
                current_branch = sc.BranchId(User.Identity.Name);
                branch_main_filial = BranchStore.getBranchMainFilial(current_branch, perso);
                
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

            res = Database.ExecuteQuery(String.Format("select * from Org where id={0}", id), ref ds,null);

            tbName.Text = ds.Tables[0].Rows[0]["title"].ToString();
            tbEmboss.Text = ds.Tables[0].Rows[0]["EmbossTitle"].ToString();
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
                if (tbEmboss.Text == "")
                {
                    lbInform.Text = "Введите название, эмбоссируемое на карте";
                    tbEmboss.Focus();
                    return;
                }

                SqlCommand sqCom = new SqlCommand();
                //sqCom.CommandText = "select count(*) from Org where embosstitle like '" + tbEmboss.Text.PadRight(50) + "'";
                
                if (branch_main_filial <= 0)
                    sqCom.CommandText = "select count(*) from Org where BranchMainFilialId is null and embosstitle like '" + tbEmboss.Text.PadRight(50) + "'";
                else
                    sqCom.CommandText = "select count(*) from Org where BranchMainFilialId=" + branch_main_filial.ToString() + " and embosstitle like '" + tbEmboss.Text.PadRight(50) + "'";
                  
                object obj = null;
                Database.ExecuteScalar(sqCom, ref obj, null);
                if ((Request.QueryString["mode"] == "1" && Convert.ToInt32(obj) > 0) ||
                       (Request.QueryString["mode"] == "2" && Convert.ToInt32(obj) > 1))
                {
                    lbInform.Text = "Ошибка: такое эмбоссированное название уже есть";
                    tbEmboss.Focus();
                    return;
                }


                if (Request.QueryString["mode"] == "1")
                {
                    //sqCom.CommandText = "insert into Org (title, embosstitle) values (@title, @embosstitle)";
                    
                     if (branch_main_filial <= 0) 
                        sqCom.CommandText = "insert into Org (title, embosstitle) values (@title, @embosstitle)";
                    else
                        sqCom.CommandText = "insert into Org (title, embosstitle,BranchMainFilialId) values (@title, @embosstitle," + branch_main_filial.ToString() +")";
                      
                }
                if (Request.QueryString["mode"] == "2")
                {
                    sqCom.CommandText = "update Org set title=@title, embosstitle=@embosstitle where id=@id";
                    sqCom.Parameters.Add("@id", SqlDbType.Int).Value = Convert.ToInt32(Request.QueryString["id"]);
                }
                sqCom.Parameters.Add("@title", SqlDbType.NChar, 150).Value = tbName.Text;
                sqCom.Parameters.Add("@embosstitle", SqlDbType.NChar, 50).Value = tbEmboss.Text;
                Database.ExecuteNonQuery(sqCom, null);

                Response.Write("<script language=javascript>window.returnValue='1'; window.close();</script>");
            }
        }
    }
}
