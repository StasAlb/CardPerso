using System;
using System.Data.SqlClient;
using System.Collections.Generic;
using System.Configuration;
using System.Data;
using System.Linq;
using System.Web;
using System.Web.Services;
using System.Web.Mvc;
using JQueryDataTables.Models;

namespace JQueryDataTables.Contollers
{
    public class HomeController : Controller
    {
        public string AddItem()
        {
            return "<Root></Root>";
        }
        public ActionResult GenDoc()
        {
            return View();
        }


        public ActionResult ReportQueueRefresh(JQueryDataTableParamModel param)
        {
            //WebLog.LogClass.WriteToLog("RefreshTable");
            CardPerso.Administration.ServiceClass sc = new CardPerso.Administration.ServiceClass();
            var UserId = sc.UserGuid(User.Identity.Name);

            DataSet ds = new DataSet();
            using (SqlConnection conn =
                new SqlConnection(ConfigurationManager.ConnectionStrings["ConString"].ConnectionString))
            {
                conn.Open();
                using (SqlCommand comm = conn.CreateCommand())
                {
                    comm.CommandText = $"SELECT UserId, ReportType, ReportStatus, ReportDate, id, ReportParameters, (select count(*) from ReportQuery where id<r.id and ReportStatus=1) as waitcnt FROM ReportQuery r WHERE UserId = '{UserId}' ORDER BY ReportDate DESC";
                    SqlDataAdapter da = new SqlDataAdapter(comm);
                    da.Fill(ds);
                }
                conn.Close();
            }

            IList<Report> reports = new List<Report>();
            foreach (DataRow r in ds.Tables[0].Rows)
            {
                reports.Add(new Report() { Date = $"{r["ReportDate"]:dd.MM.yyyy HH:mm:ss}", Status = Convert.ToInt32(r["ReportStatus"]),
                    Type = Convert.ToInt32(r["ReportType"]), Parameters = Convert.ToString(r["ReportParameters"]),
                    id = Convert.ToInt32(r["id"]), WaitCnt = Convert.ToInt32(r["WaitCnt"])});
            }

            //var employees = DataRepository.GetEmployees();

            ////"Business logic" methog that filter employees by the employer id
            //var companyEmployees = (from e in employees
            //    where (CompanyID == null || e.CompanyID == CompanyID)
            //    select e).ToList();

            ////UI processing logic that filter company employees by name and paginates them
            //var filteredEmployees = (from e in companyEmployees
            //    where (param.sSearch == null || e.Name.ToLower().Contains(param.sSearch.ToLower()))
            //    select e).ToList();
            var result = from r in reports.Skip(param.iDisplayStart).Take(param.iDisplayLength)
                         select new[] { r.Title, r.Date, r.StatusName, r.Status.ToString(), r.id.ToString(), r.Type.ToString()};
            return Json(new
            {
                sEcho = param.sEcho,
                iTotalRecords = reports.Count,
                iTotalDisplayRecords = reports.Count,
                aaData = result
            },
                JsonRequestBehavior.AllowGet);
        }

        [HttpPost]
        public JsonResult DeleteReport(int? reportId)
        {
            OstCard.Data.Database2.ExecuteNonQuery($"delete from ReportQuery where id={reportId}", null);
            return Json(data: "Deleted", behavior: JsonRequestBehavior.AllowGet);
       }
        [HttpPost]
        public JsonResult DownloadReport(int? reportId)
        {
            object bytes = null;

            var file = OstCard.Data.Database2.ExecuteScalar($"SELECT ReportFile FROM ReportQuery WHERE id = {reportId}", ref bytes, null);

            string filename = Guid.NewGuid().ToString();
            string path = System.IO.Path.Combine(Server.MapPath("~/Attachments/Temp"), filename + ".xls");

            System.IO.File.WriteAllBytes(path, (byte[])bytes);


            //System.Web.HttpResponse response = System.Web.HttpContext.Current.Response;


            //System.Web.HttpContext.Current.Response.ClearHeaders();
            //System.Web.HttpContext.Current.Response.ClearContent();

            //System.Web.HttpContext.Current.Response.HeaderEncoding = System.Text.Encoding.Default;
            //System.Web.HttpContext.Current.Response.AddHeader("Content-Disposition", "attachment; filename=" + "report.xls");
            //System.Web.HttpContext.Current.Response.AddHeader("Content-Length", ((byte[])bytes).Length.ToString());
            //System.Web.HttpContext.Current.Response.ContentType = "application/octet-stream";
            //System.Web.HttpContext.Current.Response.Cache.SetCacheability(System.Web.HttpCacheability.NoCache);

            //System.Web.HttpContext.Current.Response.BufferOutput = true;


            //System.Web.HttpContext.Current.Response.BinaryWrite((byte[])bytes);
            //System.Web.HttpContext.Current.Response.Flush();

            //System.Web.HttpContext.Current.ApplicationInstance.CompleteRequest();
            return Json(data: filename);
        }
        [HttpGet]
        [DeleteFileAttribute]
        public ActionResult Download(string filename, int type)
        {
            //get the temp folder and file path in server
            string path = System.IO.Path.Combine(Server.MapPath("~/Attachments/Temp"), filename + ".xls");

            return File(path, "application/octet-stream", SelectDocTemplate(type));
        }
        public string SelectDocTemplate(int reportType)
        {
            switch (reportType)
            {
                case (1):
                    return "Attachment1.xls";
                case (2):
                    return "Attachment7.xls";
                case (3):
                    return "Report1.xls";
                case (4):
                    return "Report1.xls";
                case 5:
                    return "Attachment1.xls";
                case 6:
                    return "Report2.xls";
                case 7:
                    return "Report2.xls";
                case 8:
                    return "Attachment1.xls";
                case 9:
                    return "Report1.xls";
                case 10:
                    return "Report1.xls";
                case 11:
                    return "Report1.xls";
                case 12:
                    return "Attachment9.xls";
                case 13:
                    return "Attachment10.xls";
                case 14:
                    return "Report1.xls";
                case 15:
                    return "Attachment10_1.xls";
                case 16:
                    return "Attachment15.xls";
                case 17:
                    return "Attachment15_1.xls";
                case 18:
                    return "Attachment23.xls";
                case 19:
                    return "Attachment24.xls";
                case 20:
                    return "Attachment1_4502.xls";
                case 41:
                    return "Attachment41.xls";
                case 42:
                    return "AttMemoric.xls";
                case 43:
                    return "Attachment10_2.xls";
                case 44:
                    return "Attachment44.xls";
                case 45:
                    return "Attachment124_r.xls";
                default:
                        return "Report.xls";
            }
        }
    }
    public class DeleteFileAttribute : ActionFilterAttribute
    {
        public override void OnResultExecuted(ResultExecutedContext filterContext)
        {
            filterContext.HttpContext.Response.Flush();
            string filePath = (filterContext.Result as FilePathResult).FileName;
            System.IO.File.Delete(filePath);
        }
    }
}