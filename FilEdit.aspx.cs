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
    public partial class FilEdit : System.Web.UI.Page
    {
        DataSet ds = new DataSet();
        string res = "";
        int branch_main_filial = 0;
        int branch_current = 0;
        ServiceClass sc = new ServiceClass();
        bool perso = false;
        

        protected void Page_Load(object sender, EventArgs e)
        {
           
            lock (Database.lockObjectDB)
            {
                perso = sc.UserAction(User.Identity.Name, Restrictions.Perso);
                branch_current = sc.BranchId(User.Identity.Name);
                branch_main_filial = BranchStore.getBranchMainFilial(branch_current,perso);
                if (!IsPostBack)
                {
                    ZapCombo();
                }
            }
        }

        private void ZapCombo()
        {
            ds.Clear();
            if (branch_main_filial > 0)
            {
                if (branch_main_filial > 0 && branch_main_filial==branch_current)
                    res = Database.ExecuteQuery(String.Format("select id,department from Branchs where id={0} or id_parent={0} order by department", branch_main_filial), ref ds, null);
            }
            else
            if (branch_main_filial == 0)
            {
               res = Database.ExecuteQuery("select id,department from Branchs order by department", ref ds, null);
            }

            if (ds.Tables.Count > 0)
            {
                dListBranch.DataSource = ds.Tables[0];
                dListBranch.DataTextField = "department";
                dListBranch.DataValueField = "id";
                dListBranch.DataBind();
            }
        }


        protected void bSave_Click(object sender, ImageClickEventArgs e)
        {
            Response.Write("<script language=javascript>window.returnValue='" + dListBranch.SelectedItem.Value + "'; window.close();</script>");
        }

    }
}
