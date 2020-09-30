using System;
using System.Collections;
using System.Configuration;
using System.Data;
using System.Data.SqlClient;
using System.Linq;
using System.Web;
using System.Web.Security;
using System.Web.UI;
using System.Web.UI.HtmlControls;
using System.Web.UI.WebControls;
using System.Web.UI.WebControls.WebParts;
using System.Web.Configuration;
using System.Xml.Linq;
using CardPerso.Administration;
using OstCard.Data;

namespace CardPerso
{
    public partial class Organization : System.Web.UI.Page
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
                ClientScript.RegisterHiddenField("resd", "");

                perso = sc.UserAction(User.Identity.Name, Restrictions.Perso);
                current_branch = sc.BranchId(User.Identity.Name);
                branch_main_filial = BranchStore.getBranchMainFilial(current_branch, perso);

                if (!IsPostBack)
                {
                    lbInform.Text = "";
                    Refr(0);
                }
            }
        }
        private void Refr(int rowindex)
        {
            lbInform.Text = "";
            ds.Clear();
            
            //res = Database.ExecuteQuery("select * from V_Org order by title, person", ref ds, null);
            
            
            if (branch_main_filial <= 0)
                res = Database.ExecuteQuery("select * from V_Org where BranchMainFilialId is null order by title, person", ref ds, null);
            else
                res = Database.ExecuteQuery("select * from V_Org  where BranchMainFilialId=" + branch_main_filial.ToString() + " order by title, person", ref ds, null); 
            
 
            if (ds == null || ds.Tables.Count == 0)
                return;
            ds.Tables[0].Columns.Add("WarrentS");
            string last_org = "";
            foreach (DataRow dr in ds.Tables[0].Rows)
            {
                if (dr["WStart"] != DBNull.Value && dr["WEnd"] != DBNull.Value)
                    dr["WarrentS"] = String.Format("{0} ({1:dd.MM.yyyy} - {2:dd.MM.yyyy})", dr["Warrent"], Convert.ToDateTime(dr["WStart"]), Convert.ToDateTime(dr["WEnd"]));
                if (dr["title"].ToString() == last_org)
                {
                    dr["title"] = "";
                    dr["embosstitle"] = "";
                }
                else
                    last_org = dr["title"].ToString();

            }
            gvOrg.DataSource = ds.Tables[0];
            gvOrg.DataBind();

            if (gvOrg.Rows.Count > 0)
            {
                gvOrg.SelectedIndex = rowindex;
                gvOrg.Rows[gvOrg.SelectedIndex].Focus();
            }
            
            SetButton();

            lbCount.Text = "Кол-во: " + gvOrg.Rows.Count.ToString();
        }
        private void SetButton()
        {
            lbInform.Text = "";
            if (!sc.UserAction(User.Identity.Name, Restrictions.LibraryOrgEdit))
            {
                bNew.Visible = false;
                bEdit.Visible = false;
                bDelete.Visible = false;
                bExcel.Visible = false;
                return;
            }
            if (gvOrg.Rows.Count > 0)
            {
                bEdit.Visible = true;
                bDelete.Visible = true;
                bExcel.Visible = true;
                bDelete.Attributes.Add("OnClick", String.Format("return confirm('Удалить организацию {0}?');", gvOrg.DataKeys[Convert.ToInt32(gvOrg.SelectedIndex)].Values["Title"].ToString().Trim()));                
                bEdit.Attributes.Add("OnClick", String.Format("return show_org('mode=2&id={0}')", gvOrg.DataKeys[Convert.ToInt32(gvOrg.SelectedIndex)].Values["idO"].ToString()));
                if (gvOrg.DataKeys[Convert.ToInt32(gvOrg.SelectedIndex)].Values["idP"] != DBNull.Value)
                {
                    bEditP.Attributes.Add("OnClick", String.Format("return show_person('mode=2&ido={0}&idp={1}');", gvOrg.DataKeys[Convert.ToInt32(gvOrg.SelectedIndex)].Values["idO"].ToString(), gvOrg.DataKeys[Convert.ToInt32(gvOrg.SelectedIndex)].Values["idP"].ToString()));
                    bEditP.Visible = true;
                }
                else
                {
                    bEditP.Visible = false;
                }
                bNewP.Attributes.Add("OnClick", String.Format("return show_person('mode=1&ido={0}');", gvOrg.DataKeys[Convert.ToInt32(gvOrg.SelectedIndex)].Values["idO"].ToString()));
                bDelP.Attributes.Add("OnClick", String.Format("return confirm('Удалить сотрудника {0}?');",gvOrg.DataKeys[Convert.ToInt32(gvOrg.SelectedIndex)].Values["Person"].ToString()));
            }
            else
            {
                bEdit.Visible = false;
                bDelete.Visible = false;
                bExcel.Visible = false;
            }
        }
        protected void bExcel_Click(object sender, ImageClickEventArgs e)
        {
            lock (Database.lockObjectDB)
            {
                System.Globalization.CultureInfo oldCI = System.Threading.Thread.CurrentThread.CurrentCulture;
                System.Threading.Thread.CurrentThread.CurrentCulture = new System.Globalization.CultureInfo("en-US");
                string doc = "";
                ExcelAp ep = new ExcelAp();
                if (ep.RunApp(ConfigurationSettings.AppSettings["DocPath"] + "Empty.xls"))
                {
                    ep.SetWorkSheet(1);
                    ep.ExportGridExcel(gvOrg);
                    if (WebConfigurationManager.AppSettings["DocPath"] != null)
                    {
                        doc = String.Format("{0}Temp\\catalog.xls", WebConfigurationManager.AppSettings["DocPath"]);
                        ep.SaveAsDoc(doc, false);
                    }
                }
                ep.Close();
                GC.Collect();
                System.Threading.Thread.CurrentThread.CurrentCulture = oldCI;
                if (doc.Length > 0)
                    ep.ReturnXls(Response, doc);
            }
        }
        protected void bNew_Click(object sender, ImageClickEventArgs e)
        {
            lock (Database.lockObjectDB)
            {
                Refr(0);
            }
        }
        protected void gvOrg_SelectedIndexChanged(object sender, EventArgs e)
        {
            lock (Database.lockObjectDB)
            {
                gvOrg.Rows[gvOrg.SelectedIndex].Focus();
                SetButton();
            }
        }
        protected void bEdit_Click(object sender, ImageClickEventArgs e)
        {
            lock (Database.lockObjectDB)
            {
                Refr(gvOrg.SelectedIndex);
            }
        }
        protected void bEditP_Click(object sender, ImageClickEventArgs e)
        {
            lock (Database.lockObjectDB)
            {
                Refr(gvOrg.SelectedIndex);
            }
        }
        protected void bDelete_Click(object sender, ImageClickEventArgs e)
        {
            lock (Database.lockObjectDB)
            {
                int id = Convert.ToInt32(gvOrg.DataKeys[Convert.ToInt32(gvOrg.SelectedIndex)].Values["idO"]);
                SqlCommand comm = new SqlCommand();
                comm.CommandText = "delete from Org where id=@id";
                comm.Parameters.Add("@id", SqlDbType.Int).Value = id;
                Database.ExecuteNonQuery(comm, null);
                Refr(0);
            }
        }
        protected void bNewP_Click(object sender, ImageClickEventArgs e)
        {
            lock (Database.lockObjectDB)
            {
                Refr(gvOrg.SelectedIndex);
            }
        }
        protected void bDelP_Click(object sender, ImageClickEventArgs e)
        {
            lock (Database.lockObjectDB)
            {

                try
                {
                    int id = Convert.ToInt32(gvOrg.DataKeys[Convert.ToInt32(gvOrg.SelectedIndex)].Values["idP"]);
                    SqlCommand comm = new SqlCommand();
                    comm.CommandText = "delete from OrgP where id=@id";
                    comm.Parameters.Add("@id", SqlDbType.Int).Value = id;
                    Database.ExecuteNonQuery(comm, null);
                    Refr(0);
                }
                catch (Exception)
                {
                }
            }
        }
    }
}
