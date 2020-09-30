using System;
using System.Data;
using System.Configuration;
using System.Linq;
using System.Web;
using System.Web.Security;
using System.Web.UI;
using System.Web.UI.HtmlControls;
using System.Web.UI.WebControls;
using System.Web.UI.WebControls.WebParts;
using System.Text;
using System.Net;
using System.IO;
using System.Collections.Generic;
using System.Diagnostics;
using System.ServiceModel;
using System.Runtime.Serialization;
using System.ServiceModel.Security;
using System.Xml;

namespace CardPerso
{
    public class RMKRequestData
    {
        public string materialValue = String.Empty;
        public int count = 0;
        public DateTime operDate = DateTime.MinValue;
        public string userSeries  = String.Empty;
        public string userNumber = String.Empty; 
        public string accountableSeries = String.Empty;
        public string accountableNumber = String.Empty;
        public string branchCode = String.Empty;
        public string operationType = String.Empty;
        public string operationId = String.Empty;
        public string packetId = String.Empty;
        public string personnelNumber = String.Empty;

        public string OperDate 
        {
            get 
            { 
                if (operDate == DateTime.MinValue) return String.Empty; 
                return operDate.ToString("yyyy-MM-dd"); 
            }
        }

        public void parseRequest(string xml)
        {
            XmlDocument xmldoc = new XmlDocument();
            try
            {
                xmldoc.LoadXml(xml);
                materialValue = xmldoc.DocumentElement.SelectSingleNode("MaterialValueID").InnerText;
                string c = xmldoc.DocumentElement.SelectSingleNode("Count").InnerText;
                if(c.Length>0) count = Convert.ToInt32(c);
                operDate = DateTime.MinValue;
                try
                {
                    operDate = DateTime.Parse(xmldoc.DocumentElement.SelectSingleNode("OperDate").InnerText);
                }
                catch {}
                userSeries = xmldoc.DocumentElement.SelectSingleNode("Creator_Series").InnerText;
                userNumber = xmldoc.DocumentElement.SelectSingleNode("Creator_Number").InnerText;
                accountableSeries = xmldoc.DocumentElement.SelectSingleNode("Accountable_Person_Series").InnerText;
                accountableNumber = xmldoc.DocumentElement.SelectSingleNode("Accountable_Person_Number").InnerText;
                branchCode = xmldoc.DocumentElement.SelectSingleNode("InnerCode").InnerText;
                operationType = xmldoc.DocumentElement.SelectSingleNode("OperationTypeID").InnerText;
                operationId = xmldoc.DocumentElement.SelectSingleNode("OperationID").InnerText;
                packetId = xmldoc.DocumentElement.SelectSingleNode("DocID").InnerText;
                personnelNumber = xmldoc.DocumentElement.SelectSingleNode("CreatorTabularNumber").InnerText;
            }
            catch (Exception e)
            {
                throw new Exception(e.Message);
            }
        }

        public string createRequest()
        {
            StringBuilder bodyPost = new StringBuilder();
            bodyPost.AppendLine("<?xml version=\"1.0\" encoding=\"UTF-8\"?>");
            bodyPost.AppendLine("<RmkData>");
            if (!String.IsNullOrEmpty((materialValue)))
                bodyPost.AppendLine(string.Format("\t<MaterialValueID>{0}</MaterialValueID>", materialValue));
            if (count > 0)
                bodyPost.AppendLine(string.Format("\t<Count>{0}</Count>", count));
            if (!String.IsNullOrEmpty(OperDate))
                bodyPost.AppendLine(string.Format("\t<OperDate>{0}</OperDate>", OperDate));
            if (!String.IsNullOrEmpty(userSeries))
            {
                userSeries = userSeries.Replace(" ", "");
                if (userSeries.Length == 4 && !userSeries.ToLower().Equals("null"))
                    userSeries = String.Format("{0} {1}", userSeries.Substring(0, 2), userSeries.Substring(2,2));
                bodyPost.AppendLine(string.Format("\t<Creator_Series>{0}</Creator_Series>", userSeries));
            }
            if (!String.IsNullOrEmpty(userNumber))
                bodyPost.AppendLine(string.Format("\t<Creator_Number>{0}</Creator_Number>", userNumber));
            if (!String.IsNullOrEmpty(accountableSeries))
            {
                accountableSeries = accountableSeries.Replace(" ", "");
                if (accountableSeries.Length == 4)
                    accountableSeries = String.Format("{0} {1}", accountableSeries.Substring(0, 2), accountableSeries.Substring(2, 2));
                bodyPost.AppendLine(string.Format("\t<Accountable_Person_Series>{0}</Accountable_Person_Series>", accountableSeries));
            }
            if (!String.IsNullOrEmpty(accountableNumber))
                bodyPost.AppendLine(string.Format("\t<Accountable_Person_Number>{0}</Accountable_Person_Number>", accountableNumber));
            if (!String.IsNullOrEmpty(branchCode))
            {
                //bodyPost.AppendLine(string.Format("\t<InnerCode>0013</InnerCode>", branchCode));
                bodyPost.AppendLine(string.Format("\t<InnerCode>{0}</InnerCode>", branchCode));
            }

            if (!String.IsNullOrEmpty(operationType))
                bodyPost.AppendLine(string.Format("\t<OperationTypeID>{0}</OperationTypeID>", operationType));
            if (!String.IsNullOrEmpty(operationId))
                bodyPost.AppendLine(string.Format("\t<OperationID>{0}</OperationID>", operationId));
            //11.03.2020 было закоментарено DocID. Раскоментарил. Потом у дениса все свалилось, заккоментарил обратно
            //if (!String.IsNullOrEmpty(packetId))
            //    bodyPost.AppendLine(string.Format("\t<DocID>{0}</DocID>", packetId));
            if (!String.IsNullOrEmpty(personnelNumber))
                bodyPost.AppendLine(string.Format("\t<CreatorTabularNumber>{0}</CreatorTabularNumber>", personnelNumber));
            bodyPost.AppendLine("</RmkData>");
            return bodyPost.ToString();
        }
    }

    public class RMKResponseData
    {
        public bool status;
        public string operationId;
        public string message;

        public RMKResponseData()
        {
            status = true;
            operationId = string.Empty;
            message = "";
        }

        public void parseResponse(string xml)
        {
            XmlDocument xmldoc = new XmlDocument();
            try
            {
                xmldoc.LoadXml(xml);
                status = Convert.ToBoolean(xmldoc.DocumentElement.SelectSingleNode("status").InnerText);
                operationId = xmldoc.DocumentElement.SelectSingleNode("OperationID").InnerText;
            }
            catch (Exception e)
            {
                throw new Exception(e.Message);
            }

            try
            {
                message = xmldoc.DocumentElement.SelectSingleNode("message").InnerText;
            }
            catch { }
            if (message.Length > 0)
                throw new Exception(message);
        }

        public string createResponse()
        {
            StringBuilder body = new StringBuilder();
            body.AppendLine("<?xml version=\"1.0\" encoding=\"UTF-8\"?>");
            body.AppendLine("<rmkresponse>");
            body.AppendLine(string.Format("\t<status>{0}</status>", status));
            body.AppendLine(string.Format("\t<OperationID>{0}</OperationID>", operationId));
            body.AppendLine(string.Format("\t<message>{0}</message>", message));
            body.AppendLine("</rmkresponse>");
            return body.ToString();
        }
    }

    public class RMKClient
    {
        private string _url;
        public bool isService = false;
        public string LogData = "";

        public RMKClient(string url)
        {
            _url = url;
        }


        public RMKResponseData RMKCardPersoData(RMKRequestData req)
        {

            return RunRequestService((req.createRequest()));
            //return RunRequest(req.createRequest());
        }

        private RMKResponseData RunRequestService(string body)
        {
            RMKResponseData res = null;
            //string str = "<?xml version =\"1.0\" encoding =\"utf-16\"?><RmkResponse xmlns:xsi =\"http://www.w3.org/2001/XMLSchema-instance\" xmlns:xsd = \"http://www.w3.org/2001/XMLSchema\"><status>true</status><OperationID>1d51cf78-e64f-468a-a2ea-f507f43cbbc3</OperationID></RmkResponse>";
            //res = new RMKResponseData();
            //WebLog.LogClass.WriteToLog("Ответ сервиса: {0}", str);
            //if (String.IsNullOrEmpty(str))
            //{
            //    res.status = false;
            //    res.operationId = "";
            //    res.message = "Не получены данные от службы";
            //}
            //else
            //{
            //    res.parseResponse(str);
            //}
            //return res;
            
            BasicHttpBinding binding = new BasicHttpBinding();
            binding.Name = "BasicHttpBinding_CaPService";
            binding.CloseTimeout = TimeSpan.FromMinutes(1);
            binding.OpenTimeout = TimeSpan.FromMinutes(1);
            binding.ReceiveTimeout = TimeSpan.FromMinutes(10);
            binding.SendTimeout = TimeSpan.FromMinutes(1);
            binding.AllowCookies = false;
            binding.BypassProxyOnLocal = false;
            binding.HostNameComparisonMode = HostNameComparisonMode.StrongWildcard;
            binding.MaxBufferSize = 65536;
            binding.MaxBufferPoolSize = 524288;
            binding.MessageEncoding = WSMessageEncoding.Text;
            binding.TextEncoding = System.Text.Encoding.UTF8;
            binding.TransferMode = TransferMode.Buffered;
            binding.UseDefaultWebProxy = true;

            binding.ReaderQuotas.MaxDepth = 32;
            binding.ReaderQuotas.MaxStringContentLength = 8192;
            binding.ReaderQuotas.MaxArrayLength = 16384;
            binding.ReaderQuotas.MaxBytesPerRead = 4096;
            binding.ReaderQuotas.MaxNameTableCharCount = 16384;
            LogData = "RMK request:\n";
            if (String.IsNullOrEmpty(_url))
            {
                if (!isService)
                    WebLog.LogClass.WriteToLog(body);
                else
                    LogData += body;
                RMKResponseData r = new RMKResponseData();
                r.status = true;
                r.message = "debug result: ok";
                r.operationId =  Guid.NewGuid().ToString();
                LogData += "\nRMK emulation response:\n";
                LogData += r.createResponse();

                return r;
            }
            Uri baseAddress = new Uri(_url);
            if (!isService)
                WebLog.LogClass.WriteToLog(body);
            else
                LogData += body;
            using (ChannelFactory<RmkTwoCaPService.IBasicHttpBinding_CaPService> ServiceProxy =
                new ChannelFactory<RmkTwoCaPService.IBasicHttpBinding_CaPService>(binding, new EndpointAddress(baseAddress)))
            {
                ServiceProxy.Open();
                RmkTwoCaPService.IBasicHttpBinding_CaPService client = ServiceProxy.CreateChannel();
                string stream = client.CreateCaPOperation(body);
                res = new RMKResponseData();
                LogData += "\nRMK response:\n";
                if (!isService)
                    WebLog.LogClass.WriteToLog("Ответ сервиса: {0}", stream);
                else
                    LogData += stream;
                if (String.IsNullOrEmpty(stream))
                {
                    res.status = false;
                    res.operationId = "";
                    res.message = "Не получены данные от службы";
                }
                else
                {
                    res.parseResponse(stream);
                }
                return res;
            }
        }
        private RMKResponseData RunRequest(string bodyPost)
        {
            HttpWebRequest req = null;
            try
            {

                byte[] body = Encoding.GetEncoding("UTF-8").GetBytes(bodyPost);
                req = (HttpWebRequest)WebRequest.Create(_url);
                req.UserAgent = "Mozilla/5.0 (Windows NT 6.1; WOW64; rv:43.0) Gecko/20100101 Firefox/43.0";
                req.Timeout = 45 * 1000;
                req.Method = "POST";
                req.ContentType = "text/xml";
                req.ProtocolVersion = HttpVersion.Version11;
                req.ContentLength = 0;
                if (body != null)
                {
                    req.ContentLength = body.Length;
                    using (var writer = req.GetRequestStream())
                    {
                        writer.Write(body, 0, body.Length);
                        writer.Flush();
                        writer.Close();
                    }
                }
                using (var resp = (HttpWebResponse)req.GetResponse())
                {
                    StreamReader streamR = new StreamReader(resp.GetResponseStream(), Encoding.GetEncoding("UTF-8"));
                    String content = streamR.ReadToEnd();
                    streamR.Close();
                    resp.Close();
                    RMKResponseData rd = new RMKResponseData();
                    if (content == null || content.Length < 1)
                    {
                        rd.status = false;
                        rd.operationId = "";
                        rd.message = "Не получены данные от службы";
                    }
                    else
                    {
                        rd.parseResponse(content);
                    }
                    return rd;
                }
            }
            catch (WebException we)
            {
                throw new Exception(we.Status.ToString() + " " + we.Message);
            }
            catch (Exception e)
            {
                RMKResponseData rd = new RMKResponseData();
                rd.status = false;
                rd.operationId = "";
                rd.message = e.Message;
                return rd;
            }
            finally
            {
                if (req != null) req.Abort();
                req = null;
            }
        }
    }
}
