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
    public partial class OPerson : System.Web.UI.Page
    {
        DataSet ds = new DataSet();
        ServiceClass sc = new ServiceClass();

        int branch_main_filial = 0;
        int current_branch = 0;
        bool perso = false;
                
        protected void Page_Load(object sender, EventArgs e)
        {
        

            lock (Database.lockObjectDB)
            {

                perso = sc.UserAction(User.Identity.Name, Restrictions.Perso);
                current_branch = sc.BranchId(User.Identity.Name);
                branch_main_filial = BranchStore.getBranchMainFilial(current_branch, perso);
                
                /*
                if (!sc.UserAction(User.Identity.Name, Restrictions.LibraryOrgEdit)
                                    && branch_main_filial != 106) // Костыль для казанского филиала!!!
                Response.Redirect("~\\Account\\Restricted.aspx", true);
                */ 
                    
                if (!IsPostBack)
                {
                    SqlCommand comm = new SqlCommand();
                    //comm.CommandText = String.Format("select idP, person from V_Org where embosstitle='{0}'", Request.QueryString["etitle"].ToString());
                    if (branch_main_filial > 0)
                    {
                        comm.CommandText = String.Format("select idP, person from V_Org where embosstitle='{0}' and BranchMainFilialId={1}", Request.QueryString["etitle"].ToString(), branch_main_filial);
                    }
                    else
                    {
                        comm.CommandText = String.Format("select idP, person from V_Org where embosstitle='{0}' and BranchMainFilialId is null", Request.QueryString["etitle"].ToString());
                    }
                    
                    DataSet ds = new DataSet();
                    Database.ExecuteCommand(comm, ref ds, null);
                    rblList.DataTextField = "Person";
                    rblList.DataValueField = "idP";
                    rblList.DataSource = ds.Tables[0];
                    rblList.DataBind();
                    rblList.SelectedIndex = 0;
                }
            }
        }

        protected void bSave_Click(object sender, ImageClickEventArgs e)
        {
            Response.Write(String.Format("<script language=javascript>window.returnValue='{0}'; window.close();</script>", rblList.SelectedValue));
        }
    }
}
