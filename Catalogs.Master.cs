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
using CardPerso.Administration;

namespace CardPerso
{
        
    public partial class Catalogs : System.Web.UI.MasterPage
    {
        ServiceClass sc = new ServiceClass();
        protected void Page_Load(object sender, EventArgs e)
        {
            if (!sc.UserAction(Page.User.Identity.Name, Restrictions.LibraryView) &&
                    !sc.UserAction(Page.User.Identity.Name, Restrictions.LibraryOrgEdit) && 
                        !sc.UserAction(Page.User.Identity.Name, Restrictions.AccountBranchView))
                Response.Redirect("~\\Account\\Restricted.aspx", true);
            if (sc.UserAction(Page.User.Identity.Name, Restrictions.LibraryView))
            {
                CatalogMenu.Items.Add(new MenuItem("Банки", "", "", "~//Bank.aspx"));
                CatalogMenu.Items.Add(new MenuItem("Подразделения", "", "", "~//Branch.aspx"));
                CatalogMenu.Items.Add(new MenuItem("Поставщики", "", "", "~//Supplier.aspx"));
                CatalogMenu.Items.Add(new MenuItem("Производители", "", "", "~//Manufacturer.aspx"));
                CatalogMenu.Items.Add(new MenuItem("Продукция", "", "", "~//Product.aspx"));
                CatalogMenu.Items.Add(new MenuItem("Списки вложений", "", "", "~//ProductAtt.aspx"));
                CatalogMenu.Items.Add(new MenuItem("Курьерские службы", "", "", "~//Courier.aspx"));
                CatalogMenu.Items.Add(new MenuItem("Список рассылок", "", "", "~//ListDeliver.aspx"));
                //CatalogMenu.Items.Add(new MenuItem("Расходные материалы", "", "", "~//Expendables.aspx"));
            }
            else
            {
                if (sc.UserAction(Page.User.Identity.Name, Restrictions.AccountBranchView))
                {
                    CatalogMenu.Items.Add(new MenuItem("Подразделения", "", "", "~//Branch.aspx"));
                }
            }
            if (sc.UserAction(Page.User.Identity.Name, Restrictions.LibraryOrgEdit))
                CatalogMenu.Items.Add(new MenuItem("Организации", "","","~//Organization.aspx"));
            if (sc.UserAction(Page.User.Identity.Name, Restrictions.LibraryPodotchet))
                CatalogMenu.Items.Add(new MenuItem("Подотчетные лица", "", "", "~//AccountablePerson.aspx"));
        }
    }
}