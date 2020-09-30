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

namespace CardPerso
{
    public partial class ShowCard : System.Web.UI.Page
    {
        protected void Page_Load(object sender, EventArgs e)
        {
            if (Page.IsPostBack)
                return;
            string fname = String.Format("~/Images/{1}", ConfigurationSettings.AppSettings["ImagePath"].ToString(), Request.QueryString["im"]);
            //if (System.IO.File.Exists(fname))
            iCard.ImageUrl = fname;
        }
    }
}
