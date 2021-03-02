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
using System.Xml.Linq;
using System.IO;
using Excel = Microsoft.Office.Interop.Excel;
using System.Runtime.InteropServices;
using System.Diagnostics;

namespace CardPerso
{
    public class ExcelAp
    {
        private Excel.Application eap;
        private Excel.Workbook ewb;
        private Excel.Worksheet ewsh;
        private Excel.Range erng;
        private object tms = Type.Missing;

        public bool RunApp(string path)
        {

            bool res = true;
            try
            {
                WebLog.LogClass.WriteToLog(path);
                WebLog.LogClass.WriteToLog("1");
                eap = new Excel.Application();
                WebLog.LogClass.WriteToLog("1");
                if (path == "") ewb = eap.Workbooks.Add(tms);
                else ewb = eap.Workbooks.Open(path, tms, tms, tms, tms, tms, tms, tms, tms, tms, tms, tms, tms, tms, tms);
                WebLog.LogClass.WriteToLog("1");
                WebLog.LogClass.WriteToLog("Excel opened");
            }
            catch (Exception e)
            {
                WebLog.LogClass.WriteToLog("Excel.RunApp: " + e.Message);
                res = false;
            }
            return res;
        }
        public void Show()
        {
            eap.Visible = true;
        }
        public void Close()
        {
            int handle = 0;
            try
            {
                handle = eap.Hinstance;
            }
            catch (Exception e)
            {
                WebLog.LogClass.WriteToLog("Excel.Close hinstance: " + e.Message);
            }
            try
            {
                if (erng != null) Marshal.ReleaseComObject(erng);
                erng = null;
            }
            catch (Exception e)
            {
                WebLog.LogClass.WriteToLog("Excel.Close 1: " + e.Message);
            }
            try
            {
                if (ewsh != null) Marshal.ReleaseComObject(ewsh);
                ewsh = null;
            }
            catch (Exception e)
            {
                WebLog.LogClass.WriteToLog("Excel.Close 2: " + e.Message);
            }
            try
            {
                if (ewb != null)
                {
                    ewb.Close(false, tms, tms);
                    Marshal.ReleaseComObject(ewb);
                }
                ewb = null;
                eap.Quit();
            }
            catch (Exception e)
            {
                WebLog.LogClass.WriteToLog("Excel.Close 3: " + e.Message);
            }
            try
            {
                Marshal.ReleaseComObject(eap);
                eap = null;
                if (handle > 0)
                {
                    Process[] proc = Process.GetProcessesByName("EXCEL");
                    foreach (Process p in proc)
                    {
                        if (p.MainModule.BaseAddress.ToInt32() == handle && p.MainWindowHandle.ToInt32() == 0)
                            p.Kill();
                    }
                }
            }
            catch (Exception e)
            {
                WebLog.LogClass.WriteToLog("Excel.Close 4: " + e.Message);
            }

        }
        public void SaveDoc(bool alert)
        {
            eap.DisplayAlerts = alert;
            ewb.Save();
        }
        public bool SaveAsDoc(string path, bool alert)
        {
            bool res = true;
            try
            {
                //path = "C:\\Temp\\Temp\\Attachment2.xls";
                eap.DisplayAlerts = alert;
                // если при параметре true нажать отмена, то косяк Excel
                ewb.SaveAs(path, tms, tms, tms, tms, tms, Excel.XlSaveAsAccessMode.xlNoChange, tms, tms, tms, tms, tms);
            }
            catch (Exception ex)
            {
                WebLog.LogClass.WriteToLog("Excel.Save: " + path + "\n" + ex.ToString());
                res = false;
            }
            return res;
        }

        public void SetChartData(int dataCol, int legCol, int rowStart, int rowEnd, Excel.XlChartType chartType, bool showLegend)
        {
            Excel.Worksheet wss = (Excel.Worksheet)ewb.Sheets["Graph"];
            Excel.ChartObjects cos = (Excel.ChartObjects)wss.ChartObjects(Type.Missing);
            Excel.ChartObject co = (Excel.ChartObject)cos.Item("Диаграмма 1");
            Excel.Chart ch = (Excel.Chart)co.Chart;
            ch.HasLegend = showLegend;
            ch.ChartType = chartType;
            ch.SetSourceData(ewsh.Range[ewsh.Cells[rowStart, legCol], ewsh.Cells[rowEnd, dataCol]], Type.Missing);
            ReleaseObject(ch);
            ReleaseObject(co);
            ReleaseObject(cos);
            ReleaseObject(wss);
        }

        private void ReleaseObject(object obj)
        {
            try
            {
                Marshal.ReleaseComObject(obj);
            }
            catch (Exception e)
            {
                WebLog.LogClass.WriteToLog("Excel.Release: " + e.Message);
            }

        }

        public void SetWorkSheet(int i)
        {
            ewsh = (Excel.Worksheet)ewb.Sheets[i];
            //ewsh.Activate();
        }
        public string GetCell(int row, int col)
        {
            erng = (Excel.Range)ewsh.Cells[row, col];
            return erng.Value2.ToString();
        }
        public void SetText(int row, int col, string val)
        {
            erng = (Excel.Range)ewsh.Cells[row, col];
            erng.Value2 = val;
        }
        public void AddText(int row, int col, string val)
        {
            erng = (Excel.Range)ewsh.Cells[row, col];
            erng.Value2 = erng.Value2 + val;
        }
        public void SetText(string name_def, string val)
        {
            try
            {
                ewsh.Range[name_def, tms].Value2 = val;
            }
            catch { }
        }

        public void SetText_Name(string name_def, string val)
        {
            try
            {
                ewsh.Range[name_def, tms].Value2 = val;
            }
            catch { }
        }

        public bool GetNameRC(string name_def, ref int row, ref int col)
        {
            for (int i = 0; i < eap.Names.Count; i++)
            {
                if (((Excel.Name)eap.Names.Item(i + 1, tms, tms)).Name == name_def)
                {
                    string s = ((Excel.Name)eap.Names.Item(i + 1, tms, tms)).RefersToR1C1Local.ToString();
                    int pos_c = s.IndexOf("C");
                    col = Convert.ToInt32(s.Substring(pos_c + 1, s.Length - pos_c - 1));
                    int pos_r = s.IndexOf("R");
                    row = Convert.ToInt32(s.Substring(pos_r + 1, pos_c - pos_r - 1));
                    return true;
                }
            }
            return false;
        }
        public void ShowRows(int rowStart, int rowEnd)
        {
            erng = (Excel.Range)ewsh.Range[ewsh.Cells[rowStart, 1], ewsh.Cells[rowEnd, 1]];
            erng.Rows.Hidden = false;
        }
        public void SetFormat(int row1, int col1, int row2, int col2, string fmt)
        {
            ewsh.Range[ewsh.Cells[row1, col1], ewsh.Cells[row2, col2]].NumberFormat = fmt;
        }
        public void SetFormat(int row, int col, string fmt)
        {
            erng = (Excel.Range)(ewsh.Cells[row, col]);
            erng.NumberFormat = fmt;
        }

        public void CopyRangeFormat(int source_row1, int source_col1, int source_row2, int source_col2,
            int destination_row1, int destination_col1, int destination_row2, int destination_col2)
        {
            ewsh.Range[ewsh.Cells[source_row1, source_col1], ewsh.Cells[source_row2, source_col2]].Copy();
            ewsh.Range[ewsh.Cells[destination_row1, destination_col1], ewsh.Cells[destination_row2, destination_col2]]
                .PasteSpecial(Excel.XlPasteType.xlPasteFormats);
        }

        public void MoveRange(int source_row1, int source_col1, int source_row2, int source_col2,
            int destination_row1, int destination_col1, int destination_row2, int destination_col2)
        {
            ewsh.Range[ewsh.Cells[source_row1, source_col1], ewsh.Cells[source_row2, source_col2]].Copy();
            
            ewsh.Range[ewsh.Cells[destination_row1, destination_col1], ewsh.Cells[destination_row2, destination_col2]]
                .PasteSpecial(Excel.XlPasteType.xlPasteFormats);

            ewsh.Range[ewsh.Cells[destination_row1, destination_col1], ewsh.Cells[destination_row2, destination_col2]]
                    .Value2 =
                ewsh.Range[ewsh.Cells[source_row1, source_col1], ewsh.Cells[source_row2, source_col2]].Value2;

            ewsh.Range[ewsh.Cells[source_row1, source_col1], ewsh.Cells[source_row2, source_col2]].ClearFormats();
        }
        public void SetRangeData(int row1, int col1, int row2, int col2, object[,] dt)
        {
            ewsh.Range[ewsh.Cells[row1, col1], ewsh.Cells[row2, col2]].Value2 = dt;
        }
        public void SetRangeAutoFit(int row1, int col1, int row2, int col2)
        {
            ewsh.Range[ewsh.Cells[row1, col1], ewsh.Cells[row2, col2]].Columns.AutoFit();
        }
        public void SetRangeAutoFit(string row, string col)
        {
            ewsh.Range[row, col].Columns.AutoFit();
        }

        public void SetRangeAlignment(int row1, int col1, int row2, int col2, Excel.Constants algn)
        {
            ewsh.Range[ewsh.Cells[row1, col1], ewsh.Cells[row2, col2]].Columns.HorizontalAlignment = algn;
        }

        public void SetRangeColorIndex(int row1, int col1, int row2, int col2, int colorIndex)
        {
            ewsh.Range[ewsh.Cells[row1, col1], ewsh.Cells[row2, col2]].Interior.ColorIndex = colorIndex;
        }

        public void SetRangeBold(int row1, int col1, int row2, int col2)
        {
            ewsh.Range[ewsh.Cells[row1, col1], ewsh.Cells[row2, col2]].Font.Bold = 1;
        }
        public void SetRangeBold(string row, string col)
        {
            ewsh.Range[row, col].Font.Bold = 1;
        }
        public void SetRangeBorders(int row1, int col1, int row2, int col2)
        {
            ewsh.Range[ewsh.Cells[row1, col1], ewsh.Cells[row2, col2]].Borders[Excel.XlBordersIndex.xlEdgeLeft].LineStyle = 1;
            ewsh.Range[ewsh.Cells[row1, col1], ewsh.Cells[row2, col2]].Borders[Excel.XlBordersIndex.xlEdgeRight].LineStyle = 1;
            ewsh.Range[ewsh.Cells[row1, col1], ewsh.Cells[row2, col2]].Borders[Excel.XlBordersIndex.xlEdgeTop].LineStyle = 1;
            ewsh.Range[ewsh.Cells[row1, col1], ewsh.Cells[row2, col2]].Borders[Excel.XlBordersIndex.xlEdgeBottom].LineStyle = 1;
            try
            {
                ewsh.Range[ewsh.Cells[row1, col1], ewsh.Cells[row2, col2]].Borders[Excel.XlBordersIndex.xlInsideHorizontal].LineStyle = 1;
                ewsh.Range[ewsh.Cells[row1, col1], ewsh.Cells[row2, col2]].Borders[Excel.XlBordersIndex.xlInsideVertical].LineStyle = 1;
            }
            catch
            { }
        }
        public void SetRangeBottom(int row1, int col1, int row2, int col2)
        {
            ewsh.Range[ewsh.Cells[row1, col1], ewsh.Cells[row2, col2]].Borders[Excel.XlBordersIndex.xlEdgeBottom].LineStyle = 1;
        }


        public void SetRangeXlMedium(int row1, int col1, int row2, int col2)
        {
            ewsh.Range[ewsh.Cells[row1, col1], ewsh.Cells[row2, col2]].Borders[Excel.XlBordersIndex.xlEdgeLeft].Weight = 3;
            ewsh.Range[ewsh.Cells[row1, col1], ewsh.Cells[row2, col2]].Borders[Excel.XlBordersIndex.xlEdgeRight].Weight = 3;
            ewsh.Range[ewsh.Cells[row1, col1], ewsh.Cells[row2, col2]].Borders[Excel.XlBordersIndex.xlEdgeTop].Weight = 3;
            ewsh.Range[ewsh.Cells[row1, col1], ewsh.Cells[row2, col2]].Borders[Excel.XlBordersIndex.xlEdgeBottom].Weight = 3;
            try
            {
                ewsh.Range[ewsh.Cells[row1, col1], ewsh.Cells[row2, col2]].Borders[Excel.XlBordersIndex.xlInsideHorizontal].Weight = 3;
                ewsh.Range[ewsh.Cells[row1, col1], ewsh.Cells[row2, col2]].Borders[Excel.XlBordersIndex.xlInsideVertical].Weight = 3;
            }
            catch
            { }
        }

        public void SetRangeBorders(string row, string col)
        {
            ewsh.Range[row, col].Borders[Excel.XlBordersIndex.xlEdgeLeft].LineStyle = 1;
            ewsh.Range[row, col].Borders[Excel.XlBordersIndex.xlEdgeRight].LineStyle = 1;
            ewsh.Range[row, col].Borders[Excel.XlBordersIndex.xlEdgeTop].LineStyle = 1;
            ewsh.Range[row, col].Borders[Excel.XlBordersIndex.xlEdgeBottom].LineStyle = 1;

            try
            {
                ewsh.Range[row, col].Borders[Excel.XlBordersIndex.xlInsideHorizontal].LineStyle = 1;
                ewsh.Range[row, col].Borders[Excel.XlBordersIndex.xlInsideVertical].LineStyle = 1;
            }
            catch
            { }
        }
        public void SetOrientation(bool LandScape)
        {
            if (LandScape)
                ewsh.PageSetup.Orientation = Excel.XlPageOrientation.xlLandscape;
            else
                ewsh.PageSetup.Orientation = Excel.XlPageOrientation.xlPortrait;
        }

        public void ExportGridExcel(GridView gv)
        {
            int cnt_col = 0;

            for (int i = 0; i < gv.Columns.Count; i++)
            {
                BoundField bf = gv.Columns[i] as BoundField;
                if (bf != null)
                {
                    SetText(1, i + 1, gv.Columns[i].HeaderText);
                    cnt_col++;
                }
                TemplateField tf = gv.Columns[i] as TemplateField;
                if (tf != null && gv.Columns[i].Visible)
                {
                    SetText(1, cnt_col + 1, gv.Columns[i].HeaderText);
                    cnt_col++;
                }

            }
            SetFormat(1, 1, gv.Rows.Count + 1, cnt_col, "@");
            for (int j = 0; j < gv.Rows.Count; j++)
            {
                for (int i = 0; i < gv.Columns.Count; i++)
                {
                    BoundField bf = gv.Columns[i] as BoundField;
                    if (bf != null)
                        SetText(j + 2, i + 1, gv.Rows[j].Cells[i].Text.Replace("&nbsp;", ""));
                    TemplateField tf = gv.Columns[i] as TemplateField;
                    if (tf != null)
                    {
                        string str = "";
                        try
                        {
                            str = (gv.Rows[j].Cells[i].Controls[0] as DataBoundLiteralControl).Text;
                            try
                            {
                                str = str.Substring(0, str.IndexOf('<')); //откидываем все тэги форматирования html
                            }
                            catch { }
                            str = str.Replace((char)13, (char)32).Replace((char)10, (char)32).Replace("&nbsp;", "").Trim();
                        }                    
                        catch
                        { str = ""; }
                        SetText(j + 2, i + 1, str);
                    }   
                }

            }

            SetRangeAutoFit(1, 1, gv.Rows.Count + 1, cnt_col);
            SetRangeBold(1, 1, 1, cnt_col);
            SetRangeBorders(1, 1, gv.Rows.Count + 1, cnt_col);
        }
        public void ExportGridExcel_Cards(GridView gv, DataTable dt)
        {
            int cnt_col = 0;
            //DataTable dt = (DataTable)gv.DataSource;
            for (int i = 0; i < gv.Columns.Count; i++)
            {
                BoundField bf = gv.Columns[i] as BoundField;
                if (bf != null && gv.Columns[i].Visible)
                {
                    SetText(1, cnt_col + 1, gv.Columns[i].HeaderText);
                    cnt_col++;
                }
                TemplateField tf = gv.Columns[i] as TemplateField;
                if (tf != null && tf.SortExpression == "fio")
                {
                    SetText(1, cnt_col + 1, gv.Columns[i].HeaderText);
                    cnt_col++;
                }
            }
            SetFormat(1, 1, dt.Rows.Count + 1, cnt_col, "@");
            object[,] list = new object[dt.Rows.Count, gv.Columns.Count];

            for (int j = 0; j < dt.Rows.Count; j++)
            {
                int t = 0;
                for (int i = 0; i < gv.Columns.Count; i++)
                {
                    BoundField bf = gv.Columns[i] as BoundField;
                    if (bf != null && gv.Columns[i].Visible)
                    {

                        if (bf.DataFormatString.Length == 0)
                            list[j, t] = dt.Rows[j][bf.DataField].ToString().Replace("&nbsp;", "");
                              //SetText(j + 2, t + 1, dt.Rows[j][bf.DataField].ToString().Replace("&nbsp;", ""));
                    if (bf.DataFormatString == "{0:d}" && dt.Rows[j][bf.DataField] != DBNull.Value)
                        list[j, t] = String.Format("{0:dd.MM.yyyy}", Convert.ToDateTime(dt.Rows[j][bf.DataField]));
                    //            SetText(j + 2, t + 1, String.Format("{0:dd.MM.yyyy}", Convert.ToDateTime(dt.Rows[j][bf.DataField])));
                           t++;
                    }
                    TemplateField tf = gv.Columns[i] as TemplateField;
                    if (tf != null && tf.SortExpression == "fio")
                    {
                      //  SetText(j+2, t+1, dt.Rows[j]["fio"].ToString());
                      list[j, t] = dt.Rows[j]["fio"].ToString();
                        t++;
                    }

                }
            }
            SetRangeData(2, 1, dt.Rows.Count+1, cnt_col, list);

            SetRangeAutoFit(1, 1, dt.Rows.Count + 1, cnt_col);
            SetRangeBold(1, 1, 1, cnt_col);
            SetRangeBorders(1, 1, dt.Rows.Count + 1, cnt_col);
        }
        public bool ReturnXls(System.Web.HttpResponse resp, string fileName)
        {
            System.IO.FileInfo f = new System.IO.FileInfo(fileName);
            resp.ClearHeaders();
            resp.ClearContent();

            DialogUtils.SetCookieResponse(resp);

            resp.HeaderEncoding = System.Text.Encoding.Default;
            resp.AddHeader("Content-Disposition", "attachment; filename=" + f.Name);
            resp.AddHeader("Content-Length", f.Length.ToString());
            resp.ContentType = "application/octet-stream";
            resp.Cache.SetCacheability(HttpCacheability.NoCache);
            /*
            resp.BufferOutput = false;
            resp.WriteFile(f.FullName);
            resp.Flush();
            resp.End();
            */

            resp.BufferOutput = true;
            resp.WriteFile(f.FullName);
            //resp.End();
            return true;
        }
        public bool ReturnXlsBytes(System.Web.HttpResponse resp, byte[] data, string fileName)
        {
            resp.ClearHeaders();
            resp.ClearContent();

            DialogUtils.SetCookieResponse(resp);

            resp.HeaderEncoding = System.Text.Encoding.Default;
            resp.AddHeader("Content-Disposition", "attachment; filename=" + fileName);
            resp.AddHeader("Content-Length", data.Length.ToString());
            resp.ContentType = "application/octet-stream";
            resp.Cache.SetCacheability(HttpCacheability.NoCache);
            /*
            resp.BufferOutput = false;
            resp.WriteFile(f.FullName);
            resp.Flush();
            resp.End();
            */

            resp.BufferOutput = true;
            resp.BinaryWrite(data);
            //resp.End();
            return true;
        }

        public void SetWorkSheetName(int cur, string s)
        {
            s = s.Replace(".", " ");
            if (s.Length > 30)
                s = s.Substring(0, 30);

            ewsh = (Excel.Worksheet)ewb.Sheets[cur];
            //ewsh = (Excel.Worksheet)ewb.ActiveSheet;
            try
            {
                ewsh.Name = s;
            }
            catch (Exception e)
            {
                WebLog.LogClass.WriteToLog("Excel.SetWorkSheetName: " + e.Message);
            }
        }

        public string GetWorkSheetName(int cur)
        {

            //if (cur < 1 || cur > ewb.Sheets.Count) return "";
            ewsh = (Excel.Worksheet)ewb.Sheets[cur];
            try
            {
                return ewsh.Name;
            }
            catch (Exception e)
            {
                WebLog.LogClass.WriteToLog("Excel.GetWorkSheetName: " + e.Message);
            }

            return "";
        }

        public void AddWorkSheet()
        {
            //ewsh.Copy(tms, ewb.ActiveSheet);
            ewsh.Copy(tms, ewb.Sheets[ewb.Sheets.Count]);
        }
        public void AddWorkSheet(int index)
        {
            //ewsh.Copy(tms, ewb.ActiveSheet);
            ewsh.Copy(tms, ewb.Sheets[index]);
        }

        public void DelWorkSheet(int index)
        {
            ewsh = (Excel.Worksheet)ewb.Sheets[index];
            try
            {
                ewsh.Delete();
            }
            catch (Exception e)
            {
                WebLog.LogClass.WriteToLog("Excel.DelWorkSheet: " + e.Message);
            }
        }

        public void CopyCells(int frow1, int fcol1, int frow2, int fcol2, int trow1, int tcol1, int trow2, int tcol2)
        {
            ewsh.Range[ewsh.Cells[frow1, fcol1], ewsh.Cells[frow2, fcol2]].Copy(ewsh.Range[ewsh.Cells[trow1, tcol1], ewsh.Cells[trow2, tcol2]]);

        }

        public void InsertRow(int row)
        {
            erng = (Excel.Range)ewsh.Cells[row, 1];
            erng = erng.EntireRow;
            erng.Insert(Excel.XlInsertShiftDirection.xlShiftDown, false);
        }

        public void InsertRow(int row, bool bCopy)
        {
            erng = (Excel.Range)ewsh.Cells[row, 1];
            erng = erng.EntireRow;
            erng.Insert(Excel.XlInsertShiftDirection.xlShiftDown, false);
            if (bCopy)
            {
                erng.Copy(((Excel.Range)ewsh.Cells[row, 1]).EntireRow);
            }
        }
    }
}
