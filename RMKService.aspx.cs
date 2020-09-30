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
using System.IO;
using System.Collections.Generic;

namespace CardPerso
{
    public partial class RMKService : System.Web.UI.Page
    {
        protected void Page_Load(object sender, EventArgs e)
        {

            if (Application["RMKGuide"] == null)
            {
                Application["RMKGuide"] = new List<string>();
            }

            List<string> RMKGuide = (List<string>)(Application["RMKGuide"]);
            
            Response.ContentType = "text/xml";
            StreamReader reader = new StreamReader(Request.InputStream);
            String xmlData = reader.ReadToEnd();
            RMKRequestData rqd = new RMKRequestData();
            RMKResponseData rsd = new RMKResponseData();
            try
            {
                rqd.parseRequest(xmlData);
                if (rqd.operationId.Length < 1)
                {
                    rsd.operationId = Guid.NewGuid().ToString();
                    rsd.message = "Операция успешно сохранена";
                }
                else rsd.operationId = "";
            }
            catch (Exception e1)
            {
                rsd.status = false;
                rsd.operationId = "";
                rsd.message = e1.Message;
            }

            Application["RMKGuide"] = RMKGuide;

            Response.Write(rsd.createResponse());
            Response.End();

        }
    }
}
