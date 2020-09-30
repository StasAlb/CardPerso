using System.Web;
using System;
namespace CardPerso
{
    public class DialogUtils
    {
        public static void SetScriptSubmit(System.Web.UI.Page page)
        {
            //page.ClientScript.RegisterOnSubmitStatement(page.GetType(), "OnSubmitScript", "setTimerDownLoad();ShowWait(null,{timeOut:500});");
        }
        public static void SetCookieResponse(System.Web.HttpResponse resp)
        {
            /*
            HttpCookie cookie = new HttpCookie("downLoadEnd", "+++");
            cookie.Expires = DateTime.Now.AddHours(24d);
            resp.AppendCookie(cookie);
            */
             
        }
    }
}