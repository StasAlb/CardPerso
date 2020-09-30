using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Web.UI;
using System.Web.UI.WebControls;
using OstCard.Data;

namespace WebApplication1.Account
{
    public partial class ChangePassword : System.Web.UI.Page
    {
        CardPerso.Administration.ServiceClass sc = new CardPerso.Administration.ServiceClass();
        
        protected void Page_Load(object sender, EventArgs e)
        {

        }

        protected void ChangeUserPassword_ChangedPassword(object sender, EventArgs e)
        {
            lock (Database.lockObjectDB)
            {
                Database.ExecuteNonQuery(String.Format("update aspnet_Users set ActivePassword=1 where UserId='{0}'", sc.UserGuid(User.Identity.Name)), null);
            }
        }
    }
}
