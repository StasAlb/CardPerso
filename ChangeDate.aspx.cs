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
    public partial class ChangeDate : System.Web.UI.Page
    {
        protected void Page_Load(object sender, EventArgs e)
        {
            if (IsPostBack)
                return;
            DatePickerD.Culture = System.Globalization.CultureInfo.GetCultureInfo("ru-RU");
            try
            {
                //DatePickerD.SelectedDate = Convert.ToDateTime(Request.QueryString["dt"]);
            }
            catch { }
            String dt = Request.QueryString["dt"];
            String dtNow = String.Format("{0:dd.MM.yyyy}", DateTime.Now);
            if (dt != null && dt.Length > 0 && dt.Equals(dtNow) == true)
            {
                lMessage.Text = "За какое число подтверждать документ?";
                try
                {
                    DatePickerD.SelectedDate = Convert.ToDateTime(Request.QueryString["dt"]);
                }
                catch { }
            }
            else
            {
                lMessage.Text = String.Format("Дата формирования документа ({0}) отличается от текущей. За какое число подтверждать документ?", Request.QueryString["dt"]);
                DatePickerD.SelectedDate = DateTime.Now;
            }
        }
        protected void bSave_Click(object sender, ImageClickEventArgs e)
        {
            try
            {
                DateTime.ParseExact(DatePickerD.DatePickerText, "dd.MM.yyyy", null);
                //Convert.ToDateTime(DatePickerD.DatePickerText);
            }
            catch
            {
                return;
            }
            Response.Write("<script language=javascript>window.returnValue='" + DatePickerD.SelectedDate.ToShortDateString() + "'; window.close();</script>");
        }
    }
}
