using System;
using System.IO;
using System.Net;


namespace CardPerso
{

    //public class LogClass
    //{
    //    private static object lockLog = new object();

    //    private static string getDir()
    //    {
    //        try
    //        {
    //            //string dir = System.AppDomain.CurrentDomain.BaseDirectory + "Logs";
    //            string errorDir = "";
    //            try
    //            {
    //                errorDir = System.Configuration.ConfigurationManager.AppSettings["ErrorDir"];
    //            }
    //            catch
    //            {
    //                errorDir = @"c:\CardPerso";
    //            }
    //            //string dir = "c:\\CardPerso";
    //            Directory.CreateDirectory(errorDir);
    //            errorDir += "\\";
    //            return errorDir;
    //        }
    //        catch (Exception e)
    //        {
    //            throw new Exception("LogClass.getDir: " + e.Message);
    //        }
    //    }
    //    public static void WriteToLogErr(string str, params object[] pars)
    //    {
    //        WriteToLog("Ошибка: " + str, pars);
    //    }
    //    public static void WriteToLog(string str, params object[] pars)
    //    {
    //        lock (lockLog)
    //        {
    //            DateTime dt = DateTime.Now;
    //            try
    //            {
    //                string siteName = "", machineName = Dns.GetHostName();
    //                if (System.Web.Hosting.HostingEnvironment.ApplicationHost != null)
    //                    siteName = System.Web.Hosting.HostingEnvironment.ApplicationHost.GetSiteName();
    //                else
    //                    siteName = "CardPerso";

    //                System.IO.StreamWriter sw = new StreamWriter(string.Format("{0}{1}{2}_{3:yyMMdd}.log", getDir(), machineName, siteName, dt), true, System.Text.Encoding.GetEncoding(1251));
    //                sw.Write("{0:HH.mm.ss}:{1:000}\t", dt, dt.Millisecond);
    //                sw.WriteLine(str, pars);
    //                sw.Close();
    //            }
    //            catch
    //            {

    //            }
    //        }
    //    }
    //}
}
