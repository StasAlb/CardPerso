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
using System.Web.Configuration;
using System.Xml.Linq;
using OstCard.Data;
using System.Data.SqlClient;

namespace CardPerso
{
    public partial class Catalog : System.Web.UI.Page
    {
        DataSet ds = new DataSet();
        protected void Page_Load(object sender, EventArgs e)
        {
/*            if (!IsPostBack)
            {
                pEditBranchs.Visible = false;
                pEditSuppliers.Visible = false;
                pEditManufacturers.Visible = false;
                pEditProducts.Visible = false;
                pEditCouriers.Visible = false;
                pSave.Visible = false;
                lbID.Visible = false;
                lbEdit.Visible = false;
                lbInform.Text = "";
                ZapCombo();
                Refr();
            }

            if (Request.QueryString["type"] == "branch")
                 Title = "Справочники: Подразделения";
            if (Request.QueryString["type"] == "supplier")
                 Title = "Справочники: Поставщики";
            if (Request.QueryString["type"] == "manufacturer")
                 Title = "Справочники: Производители";
            if (Request.QueryString["type"] == "product")
                 Title = "Справочники: Продукты";
            if (Request.QueryString["type"] == "courier")
                 Title = "Справочники: Курьерские службы";
          
            GC.Collect();*/
        }
    }
}
