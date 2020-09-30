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
using CardPerso;


//using System.Collections.Generic;



namespace OstCard.CardPerso
{
    public partial class CardPerso : System.Web.UI.MasterPage
    {

        public String getDepName()
        {
            lock (OstCard.Data.Database.lockObjectDB)
            {
                ServiceClass sc = new ServiceClass();
                int idBranch = sc.BranchId(Page.User.Identity.Name);
                if (idBranch < 1) return "";
                DataSet ds = new DataSet();
                OstCard.Data.Database.ExecuteQuery("select ident_dep from branchs where id="+idBranch.ToString(), ref ds, null);
                if (ds.Tables.Count > 0 && ds.Tables[0].Rows.Count > 0)
                    return "(" + ds.Tables[0].Rows[0]["ident_dep"].ToString() + ")";
                return "";
            }
        }
        
        ServiceClass sc = new ServiceClass();
        protected void Page_Load(object sender, EventArgs e)
        {
            Page.Header.DataBind();
            
            /*
            MovingStockSection mvConfiguration = (MovingStockSection)ConfigurationManager.GetSection("movingstock"); 
            foreach (FieldElement field in mvConfiguration.Fields) 
            {
                Response.Write(string.Format("Продукт: {0}<BR>", field.Product));
                foreach (ColumnElement column in field.Columns)
                {
                    Response.Write(string.Format("{0}, {1}, {2}, {3}<BR>", column.Nums, column.Debet, column.Credit, column.Ground));
                }
            } 
            */
            /*SmsInfoSection smss = (SmsInfoSection) ConfigurationManager.GetSection("sms_info");
            Response.Write($"company: {smss.SmsInfo.CompanyField}<BR>");
            foreach (SmsInfoElement sms in smss.SmsInfo)
            {
                Response.Write($"bin: {sms.Bin}, code: {sms.Code}, name: {sms.ShortName}, flag: {sms.AllCards}<BR> ");
            }
            */

            OstCard.Data.Database.Connect(ConfigurationManager.ConnectionStrings["ConString"].ConnectionString);
            OstCard.Data.Database2.Connect(ConfigurationManager.ConnectionStrings["ConString"].ConnectionString);

            /*
            Ozeki O = new Ozeki(OstCard.Data.Database.Conn);
            Ozekidata[] od = O.newData(149662);
            
            List<Ozekidata> lod = od.ToList<Ozekidata>();
            lod.RemoveAll(m => ((m.bin.Trim() == "4244368") && (m.prefix_ow.Trim() == "PWU") && (m.company.Trim() == "OOO ARTA")) == true);
            od = lod.ToArray();  
            
            O.saveData(od);
            O.readData(149662);
            od = O.newData(149665);
          
            O.saveData(od);
            O.deleteData(149662);
            O.deleteData(149665);
            */

            MenuItem mi = null;
            string nm = "";
            if (Page.User != null)
                nm = Page.User.Identity.Name;
            if (sc.UserAction(nm, Restrictions.StorageView) || sc.UserAction(nm, Restrictions.StorageFilial))
            {
                mi = new MenuItem("Хранилище");
                mi.NavigateUrl = "~/Storage.aspx";
                NavigationMenu.Items.Add(mi);
            }
            if (sc.UserAction(nm, Restrictions.CardsView))
            {
                mi = new MenuItem("Карты");
                mi.NavigateUrl = "~/Card.aspx";
                NavigationMenu.Items.Add(mi);
            }
            if (sc.UserAction(nm, Restrictions.PurchaseView))
            {
                mi = new MenuItem("Закупка");
                mi.NavigateUrl = "~/Purchase.aspx";
                NavigationMenu.Items.Add(mi);
            }
            if (sc.UserAction(nm, Restrictions.MovingView))
            {
                mi = new MenuItem("Движение");
                mi.NavigateUrl = "~/StorDoc.aspx";
                NavigationMenu.Items.Add(mi);
            }
            if (sc.UserAction(nm, Restrictions.Reports))
            {
                mi = new MenuItem("Отчеты");
                mi.NavigateUrl = "~/GenDoc.aspx";
                NavigationMenu.Items.Add(mi);
            }
            if (sc.UserAction(nm, Restrictions.LibraryView) || sc.UserAction(nm, Restrictions.LibraryOrgEdit))
            {
                mi = new MenuItem("Справочники");
                mi.NavigateUrl="~/Catalog.aspx";
                NavigationMenu.Items.Add(mi);
            }
            if (sc.UserAction(nm, Restrictions.UserRolesEdit) || sc.UserAction(nm, Restrictions.Events))
            {
                mi = new MenuItem("Администрирование");
                mi.NavigateUrl="~/Administration/ManageUsers.aspx";
                if (sc.UserAction(nm, Restrictions.Events))
                    mi.NavigateUrl = "~/Administration/LogView.aspx";
                NavigationMenu.Items.Add(mi);
            }
            mi = new MenuItem("Контакты");
            mi.NavigateUrl = "~/Administration/Contacts.aspx";
            NavigationMenu.Items.Add(mi);
        }
        protected void MainLoginStatus_OnLoggingOut(object sender, LoginCancelEventArgs e)
        {
            //FormsAuthentication.SetAuthCookie("", false);
        }
    }
}
