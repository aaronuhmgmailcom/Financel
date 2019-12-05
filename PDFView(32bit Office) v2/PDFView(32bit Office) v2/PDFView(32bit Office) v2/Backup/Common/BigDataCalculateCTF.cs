using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
namespace ExcelAddIn4
{
    public partial class Finance_Tools
    {
        /// <summary>
        /// 
        /// </summary>
        /// <param name="lastRowNum"></param>
        /// <param name="list"></param>
        /// <param name="StartColumn"></param>
        /// <param name="LineIndicator"></param>
        /// <param name="s"></param>
        /// <param name="ws"></param>
        /// <param name="errorStr"></param>
        /// <param name="newlist"></param>
        /// <param name="start"></param>
        /// <param name="over"></param>
        private Microsoft.Office.Interop.Excel.Worksheet CalTransCTF(string lastRowNum, List<RowCreateTextFile> list, string StartColumn, string LineIndicator, string s, Microsoft.Office.Interop.Excel.Worksheet ws, ref List<RowCreateTextFile> newlist, int start, int over)
        {
            for (int j = start; j <= over; j++)
            {
                System.Windows.Forms.Application.DoEvents();
                for (int i = 0; i < list.Count; i++)
                {
                    RowCreateTextFile re;
                    try
                    {
                        try
                        {
                            var lastrow = ws.Cells.Find("*", ws.Cells[1, 1], Microsoft.Office.Interop.Excel.XlFindLookIn.xlFormulas, Microsoft.Office.Interop.Excel.XlLookAt.xlPart, Microsoft.Office.Interop.Excel.XlSearchOrder.xlByRows, Microsoft.Office.Interop.Excel.XlSearchDirection.xlPrevious, false, Type.Missing, Type.Missing);
                            string LastRowNumber = Finance_Tools.RemoveNotNumber(lastrow.Address);
                            if (LastRowNumber != Ribbon2.LastRowNumber)
                            {
                                Globals.ThisAddIn.Application.DisplayAlerts = false;
                                Ribbon2.wsRrigin.get_Range("A1", Ribbon2.LastColumnName + Ribbon2.LastRowNumber).Copy(ws.get_Range("A1"));
                                Globals.ThisAddIn.Application.DisplayAlerts = true;
                                System.Windows.Forms.Clipboard.Clear();
                            }
                        }
                        catch
                        {//Copy sheet to make ws not disappear.
                            Globals.ThisAddIn.Application.DisplayAlerts = false;
                            Ribbon2.wsRrigin.get_Range("A1", Ribbon2.LastColumnName + Ribbon2.LastRowNumber).Copy(ws.get_Range("A1"));
                            Globals.ThisAddIn.Application.DisplayAlerts = true;
                            System.Windows.Forms.Clipboard.Clear();
                        }
                        string str = list[i].ToString();
                        if (string.IsNullOrEmpty(str)) continue;
                        if (!IsStartingCellRowContainLineIndicator(StartColumn + (j + int.Parse(s)).ToString(), LineIndicator)) continue;
                        if (SessionInfo.UserInfo.CurrentRef != list[i].ReferenceNumber) continue;
                        re = new RowCreateTextFile();
                        re.SunComponent = list[i].SunComponent;
                        re.SunMethod = list[i].SunMethod;
                        re.Column1 = string.IsNullOrEmpty(list[i].Column1) ? "" : GetEntityValueInExcel(list[i].Column1, s, j, ws);
                        re.Column2 = string.IsNullOrEmpty(list[i].Column2) ? "" : GetEntityValueInExcel(list[i].Column2, s, j, ws);
                        re.Column3 = string.IsNullOrEmpty(list[i].Column3) ? "" : GetEntityValueInExcel(list[i].Column3, s, j, ws);
                        re.Column4 = string.IsNullOrEmpty(list[i].Column4) ? "" : GetEntityValueInExcel(list[i].Column4, s, j, ws);
                        re.Column5 = string.IsNullOrEmpty(list[i].Column5) ? "" : GetEntityValueInExcel(list[i].Column5, s, j, ws);
                        re.Column6 = string.IsNullOrEmpty(list[i].Column6) ? "" : GetEntityValueInExcel(list[i].Column6, s, j, ws);
                        re.Column7 = string.IsNullOrEmpty(list[i].Column7) ? "" : GetEntityValueInExcel(list[i].Column7, s, j, ws);
                        re.Column8 = string.IsNullOrEmpty(list[i].Column8) ? "" : GetEntityValueInExcel(list[i].Column8, s, j, ws);
                        re.Column9 = string.IsNullOrEmpty(list[i].Column9) ? "" : GetEntityValueInExcel(list[i].Column9, s, j, ws);
                        re.Column10 = string.IsNullOrEmpty(list[i].Column10) ? "" : GetEntityValueInExcel(list[i].Column10, s, j, ws);
                        re.Column11 = string.IsNullOrEmpty(list[i].Column11) ? "" : GetEntityValueInExcel(list[i].Column11, s, j, ws);
                        re.Column12 = string.IsNullOrEmpty(list[i].Column12) ? "" : GetEntityValueInExcel(list[i].Column12, s, j, ws);
                        re.Column13 = string.IsNullOrEmpty(list[i].Column13) ? "" : GetEntityValueInExcel(list[i].Column13, s, j, ws);
                        re.Column14 = string.IsNullOrEmpty(list[i].Column14) ? "" : GetEntityValueInExcel(list[i].Column14, s, j, ws);
                        re.Column15 = string.IsNullOrEmpty(list[i].Column15) ? "" : GetEntityValueInExcel(list[i].Column15, s, j, ws);
                        re.Column16 = string.IsNullOrEmpty(list[i].Column16) ? "" : GetEntityValueInExcel(list[i].Column16, s, j, ws);
                        re.Column17 = string.IsNullOrEmpty(list[i].Column17) ? "" : GetEntityValueInExcel(list[i].Column17, s, j, ws);
                        re.Column18 = string.IsNullOrEmpty(list[i].Column18) ? "" : GetEntityValueInExcel(list[i].Column18, s, j, ws);
                        re.Column19 = string.IsNullOrEmpty(list[i].Column19) ? "" : GetEntityValueInExcel(list[i].Column19, s, j, ws);
                        re.Column20 = string.IsNullOrEmpty(list[i].Column20) ? "" : GetEntityValueInExcel(list[i].Column20, s, j, ws);
                        re.Column21 = string.IsNullOrEmpty(list[i].Column21) ? "" : GetEntityValueInExcel(list[i].Column21, s, j, ws);
                        re.Column22 = string.IsNullOrEmpty(list[i].Column22) ? "" : GetEntityValueInExcel(list[i].Column22, s, j, ws);
                        re.Column23 = string.IsNullOrEmpty(list[i].Column23) ? "" : GetEntityValueInExcel(list[i].Column23, s, j, ws);
                        re.Column24 = string.IsNullOrEmpty(list[i].Column24) ? "" : GetEntityValueInExcel(list[i].Column24, s, j, ws);
                        re.Column25 = string.IsNullOrEmpty(list[i].Column25) ? "" : GetEntityValueInExcel(list[i].Column25, s, j, ws);
                        re.Column26 = string.IsNullOrEmpty(list[i].Column26) ? "" : GetEntityValueInExcel(list[i].Column26, s, j, ws);
                        re.Column27 = string.IsNullOrEmpty(list[i].Column27) ? "" : GetEntityValueInExcel(list[i].Column27, s, j, ws);
                        re.Column28 = string.IsNullOrEmpty(list[i].Column28) ? "" : GetEntityValueInExcel(list[i].Column28, s, j, ws);
                        re.Column29 = string.IsNullOrEmpty(list[i].Column29) ? "" : GetEntityValueInExcel(list[i].Column29, s, j, ws);
                        re.Column30 = string.IsNullOrEmpty(list[i].Column30) ? "" : GetEntityValueInExcel(list[i].Column30, s, j, ws);
                        re.Column31 = string.IsNullOrEmpty(list[i].Column31) ? "" : GetEntityValueInExcel(list[i].Column31, s, j, ws);
                        re.Column32 = string.IsNullOrEmpty(list[i].Column32) ? "" : GetEntityValueInExcel(list[i].Column32, s, j, ws);
                        re.Column33 = string.IsNullOrEmpty(list[i].Column33) ? "" : GetEntityValueInExcel(list[i].Column33, s, j, ws);
                        re.Column34 = string.IsNullOrEmpty(list[i].Column34) ? "" : GetEntityValueInExcel(list[i].Column34, s, j, ws);
                        re.Column35 = string.IsNullOrEmpty(list[i].Column35) ? "" : GetEntityValueInExcel(list[i].Column35, s, j, ws);
                        re.Column36 = string.IsNullOrEmpty(list[i].Column36) ? "" : GetEntityValueInExcel(list[i].Column36, s, j, ws);
                        re.Column37 = string.IsNullOrEmpty(list[i].Column37) ? "" : GetEntityValueInExcel(list[i].Column37, s, j, ws);
                        re.Column38 = string.IsNullOrEmpty(list[i].Column38) ? "" : GetEntityValueInExcel(list[i].Column38, s, j, ws);
                        re.Column39 = string.IsNullOrEmpty(list[i].Column39) ? "" : GetEntityValueInExcel(list[i].Column39, s, j, ws);
                        re.Column40 = string.IsNullOrEmpty(list[i].Column40) ? "" : GetEntityValueInExcel(list[i].Column40, s, j, ws);
                        re.Column41 = string.IsNullOrEmpty(list[i].Column41) ? "" : GetEntityValueInExcel(list[i].Column41, s, j, ws);
                        re.Column42 = string.IsNullOrEmpty(list[i].Column42) ? "" : GetEntityValueInExcel(list[i].Column42, s, j, ws);
                        re.Column43 = string.IsNullOrEmpty(list[i].Column43) ? "" : GetEntityValueInExcel(list[i].Column43, s, j, ws);
                        re.Column44 = string.IsNullOrEmpty(list[i].Column44) ? "" : GetEntityValueInExcel(list[i].Column44, s, j, ws);
                        re.Column45 = string.IsNullOrEmpty(list[i].Column45) ? "" : GetEntityValueInExcel(list[i].Column45, s, j, ws);
                        re.Column46 = string.IsNullOrEmpty(list[i].Column46) ? "" : GetEntityValueInExcel(list[i].Column46, s, j, ws);
                        re.Column47 = string.IsNullOrEmpty(list[i].Column47) ? "" : GetEntityValueInExcel(list[i].Column47, s, j, ws);
                        re.Column48 = string.IsNullOrEmpty(list[i].Column48) ? "" : GetEntityValueInExcel(list[i].Column48, s, j, ws);
                        re.Column49 = string.IsNullOrEmpty(list[i].Column49) ? "" : GetEntityValueInExcel(list[i].Column49, s, j, ws);
                        re.Column50 = string.IsNullOrEmpty(list[i].Column50) ? "" : GetEntityValueInExcel(list[i].Column50, s, j, ws);
                    }
                    catch (Exception ex)
                    {
                        LogErrorCTF("The data in Line " + (j + int.Parse(s)).ToString() + " has error! " + ex.Message + "\r\n");
                        continue;
                    }
                    newlist.Add(re);
                }
            }
            return ws;
        }
        private delegate void dgetLogErrorCTF(string strLogessage);
        public void LogErrorCTF(string LogText)
        {
            try
            {
                if (CreateTextFileForm.richTextBox1.InvokeRequired)
                    CreateTextFileForm.richTextBox1.BeginInvoke(new dgetLogErrorCTF(LogErrorCTF), new object[] { LogText });
                else
                    CreateTextFileForm.richTextBox1.Text += LogText;
            }
            catch
            { }
        }
        private delegate void dgetLogTextCTF(int strLogessage);
        public void LogCTF(int LogText)
        {
            try
            {
                if (CreateTextFileForm.progressBar1.InvokeRequired)
                    CreateTextFileForm.progressBar1.BeginInvoke(new dgetLogTextCTF(LogCTF), new object[] { LogText });
                else
                    CreateTextFileForm.progressBar1.Value = LogText;
            }
            catch { }
        }
        /// <summary>
        /// 
        /// </summary>
        /// <param name="lastRowNum"></param>
        /// <param name="list"></param>
        /// <param name="StartColumn"></param>
        /// <param name="LineIndicator"></param>
        /// <param name="s"></param>
        /// <param name="ws"></param>
        /// <param name="errorStr"></param>
        /// <param name="newlist"></param>
        public void BigCalculateCTF(string lastRowNum, List<RowCreateTextFile> list, string StartColumn, string LineIndicator, string s, Microsoft.Office.Interop.Excel.Worksheet ws, ref List<RowCreateTextFile> newlist)
        {
            int rowCount = int.Parse(lastRowNum) - int.Parse(s);
            if (rowCount > 500)
            {
                CreateTextFileForm.progressBar1.Visible = true;
                CreateTextFileForm.progressBar1.Maximum = rowCount;
            }
            //================================================================below 300======================================================
            List<RowCreateTextFile> newlist1 = new List<RowCreateTextFile>();
            if (rowCount <= 300)
            {
                ws = CalTransCTF(lastRowNum, list, StartColumn, LineIndicator, s, ws, ref newlist1, 0, rowCount);
                newlist.AddRange(newlist1);
            }
            else if (rowCount > 300)//================================================================above 300======================================================
            {
                ws = CalTransCTF(lastRowNum, list, StartColumn, LineIndicator, s, ws, ref newlist1, 0, 100);
                newlist.AddRange(newlist1);
                newlist1.Clear();
                GC.Collect();
                ws = CalTransCTF(lastRowNum, list, StartColumn, LineIndicator, s, ws, ref newlist1, 101, 200);
                newlist.AddRange(newlist1);
                newlist1.Clear();
                GC.Collect();
                ws = CalTransCTF(lastRowNum, list, StartColumn, LineIndicator, s, ws, ref newlist1, 201, 300);
                newlist.AddRange(newlist1);
                newlist1.Clear();
                GC.Collect();
                //===================================================================================400=============================================================
                if (rowCount <= 400)
                {
                    ws = CalTransCTF(lastRowNum, list, StartColumn, LineIndicator, s, ws, ref newlist1, 301, rowCount);
                    newlist.AddRange(newlist1);
                    return;
                }
                ws = CalTransCTF(lastRowNum, list, StartColumn, LineIndicator, s, ws, ref newlist1, 301, 400);
                newlist.AddRange(newlist1);
                newlist1.Clear();
                GC.Collect();
                //===================================================================================500=============================================================
                if (rowCount <= 500)
                {
                    ws = CalTransCTF(lastRowNum, list, StartColumn, LineIndicator, s, ws, ref newlist1, 401, rowCount);
                    newlist.AddRange(newlist1);
                    return;
                }
                ws = CalTransCTF(lastRowNum, list, StartColumn, LineIndicator, s, ws, ref newlist1, 401, 500);
                newlist.AddRange(newlist1);
                newlist1.Clear();
                LogCTF(500);
                GC.Collect();
                //===================================================================================600=============================================================
                if (rowCount <= 600)
                {
                    ws = CalTransCTF(lastRowNum, list, StartColumn, LineIndicator, s, ws, ref newlist1, 501, rowCount);
                    newlist.AddRange(newlist1);
                    return;
                }
                ws = CalTransCTF(lastRowNum, list, StartColumn, LineIndicator, s, ws, ref newlist1, 501, 600);
                newlist.AddRange(newlist1);
                newlist1.Clear();
                LogCTF(600);
                GC.Collect();
                //===================================================================================700=============================================================
                if (rowCount <= 700)
                {
                    ws = CalTransCTF(lastRowNum, list, StartColumn, LineIndicator, s, ws, ref newlist1, 601, rowCount);
                    newlist.AddRange(newlist1);
                    return;
                }
                ws = CalTransCTF(lastRowNum, list, StartColumn, LineIndicator, s, ws, ref newlist1, 601, 700);
                newlist.AddRange(newlist1);
                newlist1.Clear();
                LogCTF(700);
                GC.Collect();
                //===================================================================================800=============================================================
                if (rowCount <= 800)
                {
                    ws = CalTransCTF(lastRowNum, list, StartColumn, LineIndicator, s, ws, ref newlist1, 701, rowCount);
                    newlist.AddRange(newlist1);
                    return;
                }
                ws = CalTransCTF(lastRowNum, list, StartColumn, LineIndicator, s, ws, ref newlist1, 701, 800);
                newlist.AddRange(newlist1);
                newlist1.Clear();
                LogCTF(800);
                GC.Collect();
                //===================================================================================900=============================================================
                if (rowCount <= 900)
                {
                    ws = CalTransCTF(lastRowNum, list, StartColumn, LineIndicator, s, ws, ref newlist1, 801, rowCount);
                    newlist.AddRange(newlist1);
                    return;
                }
                ws = CalTransCTF(lastRowNum, list, StartColumn, LineIndicator, s, ws, ref newlist1, 801, 900);
                newlist.AddRange(newlist1);
                newlist1.Clear();
                LogCTF(900);
                GC.Collect();
                //===================================================================================1000=============================================================
                if (rowCount <= 1000)
                {
                    ws = CalTransCTF(lastRowNum, list, StartColumn, LineIndicator, s, ws, ref newlist1, 901, rowCount);
                    newlist.AddRange(newlist1);
                    return;
                }
                ws = CalTransCTF(lastRowNum, list, StartColumn, LineIndicator, s, ws, ref newlist1, 901, 1000);
                newlist.AddRange(newlist1);
                newlist1.Clear();
                LogCTF(1000);
                GC.Collect();
                //===================================================================================1100=============================================================
                if (rowCount <= 1100)
                {
                    ws = CalTransCTF(lastRowNum, list, StartColumn, LineIndicator, s, ws, ref newlist1, 1001, rowCount);
                    newlist.AddRange(newlist1);
                    return;
                }
                ws = CalTransCTF(lastRowNum, list, StartColumn, LineIndicator, s, ws, ref newlist1, 1001, 1100);
                newlist.AddRange(newlist1);
                newlist1.Clear();
                LogCTF(1100);
                GC.Collect();
                //===================================================================================1200=============================================================
                if (rowCount <= 1200)
                {
                    ws = CalTransCTF(lastRowNum, list, StartColumn, LineIndicator, s, ws, ref newlist1, 1101, rowCount);
                    newlist.AddRange(newlist1);
                    return;
                }
                ws = CalTransCTF(lastRowNum, list, StartColumn, LineIndicator, s, ws, ref newlist1, 1101, 1200);
                newlist.AddRange(newlist1);
                newlist1.Clear();
                LogCTF(1200);
                GC.Collect();
                //===================================================================================1300=============================================================
                if (rowCount <= 1300)
                {
                    ws = CalTransCTF(lastRowNum, list, StartColumn, LineIndicator, s, ws, ref newlist1, 1201, rowCount);
                    newlist.AddRange(newlist1);
                    return;
                }
                ws = CalTransCTF(lastRowNum, list, StartColumn, LineIndicator, s, ws, ref newlist1, 1201, 1300);
                newlist.AddRange(newlist1);
                newlist1.Clear();
                LogCTF(1300);
                GC.Collect();
                //===================================================================================1400=============================================================
                if (rowCount <= 1400)
                {
                    ws = CalTransCTF(lastRowNum, list, StartColumn, LineIndicator, s, ws, ref newlist1, 1301, rowCount);
                    newlist.AddRange(newlist1);
                    return;
                }
                ws = CalTransCTF(lastRowNum, list, StartColumn, LineIndicator, s, ws, ref newlist1, 1301, 1400);
                newlist.AddRange(newlist1);
                newlist1.Clear();
                LogCTF(1400);
                GC.Collect();
                //===================================================================================1500=============================================================
                if (rowCount <= 1500)
                {
                    ws = CalTransCTF(lastRowNum, list, StartColumn, LineIndicator, s, ws, ref newlist1, 1401, rowCount);
                    newlist.AddRange(newlist1);
                    return;
                }
                ws = CalTransCTF(lastRowNum, list, StartColumn, LineIndicator, s, ws, ref newlist1, 1401, 1500);
                newlist.AddRange(newlist1);
                newlist1.Clear();
                LogCTF(1500);
                GC.Collect();
                //===================================================================================1600=============================================================
                if (rowCount <= 1600)
                {
                    ws = CalTransCTF(lastRowNum, list, StartColumn, LineIndicator, s, ws, ref newlist1, 1501, rowCount);
                    newlist.AddRange(newlist1);
                    return;
                }
                ws = CalTransCTF(lastRowNum, list, StartColumn, LineIndicator, s, ws, ref newlist1, 1501, 1600);
                newlist.AddRange(newlist1);
                newlist1.Clear();
                LogCTF(1600);
                GC.Collect();
                //===================================================================================1700=============================================================
                if (rowCount <= 1700)
                {
                    ws = CalTransCTF(lastRowNum, list, StartColumn, LineIndicator, s, ws, ref newlist1, 1601, rowCount);
                    newlist.AddRange(newlist1);
                    return;
                }
                ws = CalTransCTF(lastRowNum, list, StartColumn, LineIndicator, s, ws, ref newlist1, 1601, 1700);
                newlist.AddRange(newlist1);
                newlist1.Clear();
                LogCTF(1700);
                GC.Collect();
                //===================================================================================1800=============================================================
                if (rowCount <= 1800)
                {
                    ws = CalTransCTF(lastRowNum, list, StartColumn, LineIndicator, s, ws, ref newlist1, 1701, rowCount);
                    newlist.AddRange(newlist1);
                    return;
                }
                ws = CalTransCTF(lastRowNum, list, StartColumn, LineIndicator, s, ws, ref newlist1, 1701, 1800);
                newlist.AddRange(newlist1);
                newlist1.Clear();
                LogCTF(1800);
                GC.Collect();
                //===================================================================================1900=============================================================
                if (rowCount <= 1900)
                {
                    ws = CalTransCTF(lastRowNum, list, StartColumn, LineIndicator, s, ws, ref newlist1, 1801, rowCount);
                    newlist.AddRange(newlist1);
                    return;
                }
                ws = CalTransCTF(lastRowNum, list, StartColumn, LineIndicator, s, ws, ref newlist1, 1801, 1900);
                newlist.AddRange(newlist1);
                newlist1.Clear();
                LogCTF(1900);
                GC.Collect();
                //===================================================================================2000=============================================================
                if (rowCount <= 2000)
                {
                    ws = CalTransCTF(lastRowNum, list, StartColumn, LineIndicator, s, ws, ref newlist1, 1901, rowCount);
                    newlist.AddRange(newlist1);
                    return;
                }
                ws = CalTransCTF(lastRowNum, list, StartColumn, LineIndicator, s, ws, ref newlist1, 1901, 2000);
                newlist.AddRange(newlist1);
                newlist1.Clear();
                LogCTF(2000);
                GC.Collect();
                //===================================================================================2100=============================================================
                if (rowCount <= 2100)
                {
                    ws = CalTransCTF(lastRowNum, list, StartColumn, LineIndicator, s, ws, ref newlist1, 2001, rowCount);
                    newlist.AddRange(newlist1);
                    return;
                }
                ws = CalTransCTF(lastRowNum, list, StartColumn, LineIndicator, s, ws, ref newlist1, 2001, 2100);
                newlist.AddRange(newlist1);
                newlist1.Clear();
                LogCTF(2100);
                GC.Collect();
                //===================================================================================2200=============================================================
                if (rowCount <= 2200)
                {
                    ws = CalTransCTF(lastRowNum, list, StartColumn, LineIndicator, s, ws, ref newlist1, 2101, rowCount);
                    newlist.AddRange(newlist1);
                    return;
                }
                ws = CalTransCTF(lastRowNum, list, StartColumn, LineIndicator, s, ws, ref newlist1, 2101, 2200);
                newlist.AddRange(newlist1);
                newlist1.Clear();
                LogCTF(2200);
                GC.Collect();
                //===================================================================================2300=============================================================
                if (rowCount <= 2300)
                {
                    ws = CalTransCTF(lastRowNum, list, StartColumn, LineIndicator, s, ws, ref newlist1, 2201, rowCount);
                    newlist.AddRange(newlist1);
                    return;
                }
                ws = CalTransCTF(lastRowNum, list, StartColumn, LineIndicator, s, ws, ref newlist1, 2201, 2300);
                newlist.AddRange(newlist1);
                newlist1.Clear();
                LogCTF(2300);
                GC.Collect();
                //===================================================================================2400=============================================================
                if (rowCount <= 2400)
                {
                    ws = CalTransCTF(lastRowNum, list, StartColumn, LineIndicator, s, ws, ref newlist1, 2301, rowCount);
                    newlist.AddRange(newlist1);
                    return;
                }
                ws = CalTransCTF(lastRowNum, list, StartColumn, LineIndicator, s, ws, ref newlist1, 2301, 2400);
                newlist.AddRange(newlist1);
                newlist1.Clear();
                LogCTF(2400);
                GC.Collect();
                //===================================================================================2500=============================================================
                if (rowCount <= 2500)
                {
                    ws = CalTransCTF(lastRowNum, list, StartColumn, LineIndicator, s, ws, ref newlist1, 2401, rowCount);
                    newlist.AddRange(newlist1);
                    return;
                }
                ws = CalTransCTF(lastRowNum, list, StartColumn, LineIndicator, s, ws, ref newlist1, 2401, 2500);
                newlist.AddRange(newlist1);
                newlist1.Clear();
                LogCTF(2500);
                GC.Collect();
                //===================================================================================2600=============================================================
                if (rowCount <= 2600)
                {
                    ws = CalTransCTF(lastRowNum, list, StartColumn, LineIndicator, s, ws, ref newlist1, 2501, rowCount);
                    newlist.AddRange(newlist1);
                    return;
                }
                ws = CalTransCTF(lastRowNum, list, StartColumn, LineIndicator, s, ws, ref newlist1, 2501, 2600);
                newlist.AddRange(newlist1);
                newlist1.Clear();
                LogCTF(2600);
                GC.Collect();
                //===================================================================================2700=============================================================
                if (rowCount <= 2700)
                {
                    ws = CalTransCTF(lastRowNum, list, StartColumn, LineIndicator, s, ws, ref newlist1, 2601, rowCount);
                    newlist.AddRange(newlist1);
                    return;
                }
                ws = CalTransCTF(lastRowNum, list, StartColumn, LineIndicator, s, ws, ref newlist1, 2601, 2700);
                newlist.AddRange(newlist1);
                newlist1.Clear();
                LogCTF(2700);
                GC.Collect();
                //===================================================================================2800=============================================================
                if (rowCount <= 2800)
                {
                    ws = CalTransCTF(lastRowNum, list, StartColumn, LineIndicator, s, ws, ref newlist1, 2701, rowCount);
                    newlist.AddRange(newlist1);
                    return;
                }
                ws = CalTransCTF(lastRowNum, list, StartColumn, LineIndicator, s, ws, ref newlist1, 2701, 2800);
                newlist.AddRange(newlist1);
                newlist1.Clear();
                LogCTF(2800);
                GC.Collect();
                //===================================================================================2900=============================================================
                if (rowCount <= 2900)
                {
                    ws = CalTransCTF(lastRowNum, list, StartColumn, LineIndicator, s, ws, ref newlist1, 2801, rowCount);
                    newlist.AddRange(newlist1);
                    return;
                }
                ws = CalTransCTF(lastRowNum, list, StartColumn, LineIndicator, s, ws, ref newlist1, 2801, 2900);
                newlist.AddRange(newlist1);
                newlist1.Clear();
                LogCTF(2900);
                GC.Collect();
                //===================================================================================3000=============================================================
                if (rowCount <= 3000)
                {
                    ws = CalTransCTF(lastRowNum, list, StartColumn, LineIndicator, s, ws, ref newlist1, 2901, rowCount);
                    newlist.AddRange(newlist1);
                    return;
                }
                ws = CalTransCTF(lastRowNum, list, StartColumn, LineIndicator, s, ws, ref newlist1, 2901, 3000);
                newlist.AddRange(newlist1);
                newlist1.Clear();
                LogCTF(3000);
                GC.Collect();
                //===================================================================================3100=============================================================
                if (rowCount <= 3100)
                {
                    ws = CalTransCTF(lastRowNum, list, StartColumn, LineIndicator, s, ws, ref newlist1, 3001, rowCount);
                    newlist.AddRange(newlist1);
                    return;
                }
                ws = CalTransCTF(lastRowNum, list, StartColumn, LineIndicator, s, ws, ref newlist1, 3001, 3100);
                newlist.AddRange(newlist1);
                newlist1.Clear();
                LogCTF(3100);
                GC.Collect();
                //===================================================================================3200=============================================================
                if (rowCount <= 3200)
                {
                    ws = CalTransCTF(lastRowNum, list, StartColumn, LineIndicator, s, ws, ref newlist1, 3101, rowCount);
                    newlist.AddRange(newlist1);
                    return;
                }
                ws = CalTransCTF(lastRowNum, list, StartColumn, LineIndicator, s, ws, ref newlist1, 3101, 3200);
                newlist.AddRange(newlist1);
                newlist1.Clear();
                LogCTF(3200);
                GC.Collect();
                //===================================================================================3300=============================================================
                if (rowCount <= 3300)
                {
                    ws = CalTransCTF(lastRowNum, list, StartColumn, LineIndicator, s, ws, ref newlist1, 3201, rowCount);
                    newlist.AddRange(newlist1);
                    return;
                }
                ws = CalTransCTF(lastRowNum, list, StartColumn, LineIndicator, s, ws, ref newlist1, 3201, 3300);
                newlist.AddRange(newlist1);
                newlist1.Clear();
                LogCTF(3300);
                GC.Collect();
                //===================================================================================3400=============================================================
                if (rowCount <= 3400)
                {
                    ws = CalTransCTF(lastRowNum, list, StartColumn, LineIndicator, s, ws, ref newlist1, 3301, rowCount);
                    newlist.AddRange(newlist1);
                    return;
                }
                ws = CalTransCTF(lastRowNum, list, StartColumn, LineIndicator, s, ws, ref newlist1, 3301, 3400);
                newlist.AddRange(newlist1);
                newlist1.Clear();
                LogCTF(3400);
                GC.Collect();
                //===================================================================================3500=============================================================
                if (rowCount <= 3500)
                {
                    ws = CalTransCTF(lastRowNum, list, StartColumn, LineIndicator, s, ws, ref newlist1, 3401, rowCount);
                    newlist.AddRange(newlist1);
                    return;
                }
                ws = CalTransCTF(lastRowNum, list, StartColumn, LineIndicator, s, ws, ref newlist1, 3401, 3500);
                newlist.AddRange(newlist1);
                newlist1.Clear();
                LogCTF(3500);
                GC.Collect();
                //===================================================================================3600=============================================================
                if (rowCount <= 3600)
                {
                    ws = CalTransCTF(lastRowNum, list, StartColumn, LineIndicator, s, ws, ref newlist1, 3501, rowCount);
                    newlist.AddRange(newlist1);
                    return;
                }
                ws = CalTransCTF(lastRowNum, list, StartColumn, LineIndicator, s, ws, ref newlist1, 3501, 3600);
                newlist.AddRange(newlist1);
                newlist1.Clear();
                LogCTF(3600);
                GC.Collect();
                //===================================================================================3700=============================================================
                if (rowCount <= 3700)
                {
                    ws = CalTransCTF(lastRowNum, list, StartColumn, LineIndicator, s, ws, ref newlist1, 3601, rowCount);
                    newlist.AddRange(newlist1);
                    return;
                }
                ws = CalTransCTF(lastRowNum, list, StartColumn, LineIndicator, s, ws, ref newlist1, 3601, 3700);
                newlist.AddRange(newlist1);
                newlist1.Clear();
                LogCTF(3700);
                GC.Collect();
                //===================================================================================3800=============================================================
                if (rowCount <= 3800)
                {
                    ws = CalTransCTF(lastRowNum, list, StartColumn, LineIndicator, s, ws, ref newlist1, 3701, rowCount);
                    newlist.AddRange(newlist1);
                    return;
                }
                ws = CalTransCTF(lastRowNum, list, StartColumn, LineIndicator, s, ws, ref newlist1, 3701, 3800);
                newlist.AddRange(newlist1);
                newlist1.Clear();
                LogCTF(3800);
                GC.Collect();
                //===================================================================================3900=============================================================
                if (rowCount <= 3900)
                {
                    ws = CalTransCTF(lastRowNum, list, StartColumn, LineIndicator, s, ws, ref newlist1, 3801, rowCount);
                    newlist.AddRange(newlist1);
                    return;
                }
                ws = CalTransCTF(lastRowNum, list, StartColumn, LineIndicator, s, ws, ref newlist1, 3801, 3900);
                newlist.AddRange(newlist1);
                newlist1.Clear();
                LogCTF(3900);
                GC.Collect();
                //===================================================================================4000=============================================================
                if (rowCount <= 4000)
                {
                    ws = CalTransCTF(lastRowNum, list, StartColumn, LineIndicator, s, ws, ref newlist1, 3901, rowCount);
                    newlist.AddRange(newlist1);
                    return;
                }
                ws = CalTransCTF(lastRowNum, list, StartColumn, LineIndicator, s, ws, ref newlist1, 3901, 4000);
                newlist.AddRange(newlist1);
                newlist1.Clear();
                LogCTF(4000);
                GC.Collect();
                //===================================================================================4100=============================================================
                if (rowCount <= 4100)
                {
                    ws = CalTransCTF(lastRowNum, list, StartColumn, LineIndicator, s, ws, ref newlist1, 4001, rowCount);
                    newlist.AddRange(newlist1);
                    return;
                }
                ws = CalTransCTF(lastRowNum, list, StartColumn, LineIndicator, s, ws, ref newlist1, 4001, 4100);
                newlist.AddRange(newlist1);
                newlist1.Clear();
                LogCTF(4100);
                GC.Collect();
                //===================================================================================4200=============================================================
                if (rowCount <= 4200)
                {
                    ws = CalTransCTF(lastRowNum, list, StartColumn, LineIndicator, s, ws, ref newlist1, 4101, rowCount);
                    newlist.AddRange(newlist1);
                    return;
                }
                ws = CalTransCTF(lastRowNum, list, StartColumn, LineIndicator, s, ws, ref newlist1, 4101, 4200);
                newlist.AddRange(newlist1);
                newlist1.Clear();
                LogCTF(4200);
                GC.Collect();
                //===================================================================================4300=============================================================
                if (rowCount <= 4300)
                {
                    ws = CalTransCTF(lastRowNum, list, StartColumn, LineIndicator, s, ws, ref newlist1, 4201, rowCount);
                    newlist.AddRange(newlist1);
                    return;
                }
                ws = CalTransCTF(lastRowNum, list, StartColumn, LineIndicator, s, ws, ref newlist1, 4201, 4300);
                newlist.AddRange(newlist1);
                newlist1.Clear();
                LogCTF(4300);
                GC.Collect();
                //===================================================================================4400=============================================================
                if (rowCount <= 4400)
                {
                    ws = CalTransCTF(lastRowNum, list, StartColumn, LineIndicator, s, ws, ref newlist1, 4301, rowCount);
                    newlist.AddRange(newlist1);
                    return;
                }
                ws = CalTransCTF(lastRowNum, list, StartColumn, LineIndicator, s, ws, ref newlist1, 4301, 4400);
                newlist.AddRange(newlist1);
                newlist1.Clear();
                LogCTF(4400);
                GC.Collect();
                //===================================================================================4500=============================================================
                if (rowCount <= 4500)
                {
                    ws = CalTransCTF(lastRowNum, list, StartColumn, LineIndicator, s, ws, ref newlist1, 4401, rowCount);
                    newlist.AddRange(newlist1);
                    return;
                }
                ws = CalTransCTF(lastRowNum, list, StartColumn, LineIndicator, s, ws, ref newlist1, 4401, 4500);
                newlist.AddRange(newlist1);
                newlist1.Clear();
                LogCTF(4500);
                GC.Collect();
                //===================================================================================4600=============================================================
                if (rowCount <= 4600)
                {
                    ws = CalTransCTF(lastRowNum, list, StartColumn, LineIndicator, s, ws, ref newlist1, 4501, rowCount);
                    newlist.AddRange(newlist1);
                    return;
                }
                ws = CalTransCTF(lastRowNum, list, StartColumn, LineIndicator, s, ws, ref newlist1, 4501, 4600);
                newlist.AddRange(newlist1);
                newlist1.Clear();
                LogCTF(4600);
                GC.Collect();
                //===================================================================================4700=============================================================
                if (rowCount <= 4700)
                {
                    ws = CalTransCTF(lastRowNum, list, StartColumn, LineIndicator, s, ws, ref newlist1, 4601, rowCount);
                    newlist.AddRange(newlist1);
                    return;
                }
                ws = CalTransCTF(lastRowNum, list, StartColumn, LineIndicator, s, ws, ref newlist1, 4601, 4700);
                newlist.AddRange(newlist1);
                newlist1.Clear();
                LogCTF(4700);
                GC.Collect();
                //===================================================================================4800=============================================================
                if (rowCount <= 4800)
                {
                    ws = CalTransCTF(lastRowNum, list, StartColumn, LineIndicator, s, ws, ref newlist1, 4701, rowCount);
                    newlist.AddRange(newlist1);
                    return;
                }
                ws = CalTransCTF(lastRowNum, list, StartColumn, LineIndicator, s, ws, ref newlist1, 4701, 4800);
                newlist.AddRange(newlist1);
                newlist1.Clear();
                LogCTF(4800);
                GC.Collect();
                //===================================================================================4900=============================================================
                if (rowCount <= 4900)
                {
                    ws = CalTransCTF(lastRowNum, list, StartColumn, LineIndicator, s, ws, ref newlist1, 4801, rowCount);
                    newlist.AddRange(newlist1);
                    return;
                }
                ws = CalTransCTF(lastRowNum, list, StartColumn, LineIndicator, s, ws, ref newlist1, 4801, 4900);
                newlist.AddRange(newlist1);
                newlist1.Clear();
                LogCTF(4900);
                GC.Collect();
                //===================================================================================5000=============================================================
                if (rowCount <= 5000)
                {
                    ws = CalTransCTF(lastRowNum, list, StartColumn, LineIndicator, s, ws, ref newlist1, 4901, rowCount);
                    newlist.AddRange(newlist1);
                    return;
                }
                ws = CalTransCTF(lastRowNum, list, StartColumn, LineIndicator, s, ws, ref newlist1, 4901, 5000);
                newlist.AddRange(newlist1);
                newlist1.Clear();
                LogCTF(5000);
                GC.Collect();
                //===================================================================================5100=============================================================
                if (rowCount <= 5100)
                {
                    ws = CalTransCTF(lastRowNum, list, StartColumn, LineIndicator, s, ws, ref newlist1, 5001, rowCount);
                    newlist.AddRange(newlist1);
                    return;
                }
                ws = CalTransCTF(lastRowNum, list, StartColumn, LineIndicator, s, ws, ref newlist1, 5001, 5100);
                newlist.AddRange(newlist1);
                newlist1.Clear();
                LogCTF(5100);
                GC.Collect();
                //===================================================================================5200=============================================================
                if (rowCount <= 5200)
                {
                    ws = CalTransCTF(lastRowNum, list, StartColumn, LineIndicator, s, ws, ref newlist1, 5101, rowCount);
                    newlist.AddRange(newlist1);
                    return;
                }
                ws = CalTransCTF(lastRowNum, list, StartColumn, LineIndicator, s, ws, ref newlist1, 5101, 5200);
                newlist.AddRange(newlist1);
                newlist1.Clear();
                LogCTF(5200);
                GC.Collect();
                //===================================================================================5300=============================================================
                if (rowCount <= 5300)
                {
                    ws = CalTransCTF(lastRowNum, list, StartColumn, LineIndicator, s, ws, ref newlist1, 5201, rowCount);
                    newlist.AddRange(newlist1);
                    return;
                }
                ws = CalTransCTF(lastRowNum, list, StartColumn, LineIndicator, s, ws, ref newlist1, 5201, 5300);
                newlist.AddRange(newlist1);
                newlist1.Clear();
                LogCTF(5300);
                GC.Collect();
                //===================================================================================5400=============================================================
                if (rowCount <= 5400)
                {
                    ws = CalTransCTF(lastRowNum, list, StartColumn, LineIndicator, s, ws, ref newlist1, 5301, rowCount);
                    newlist.AddRange(newlist1);
                    return;
                }
                ws = CalTransCTF(lastRowNum, list, StartColumn, LineIndicator, s, ws, ref newlist1, 5301, 5400);
                newlist.AddRange(newlist1);
                newlist1.Clear();
                LogCTF(5400);
                GC.Collect();
                //===================================================================================5500=============================================================
                if (rowCount <= 5500)
                {
                    ws = CalTransCTF(lastRowNum, list, StartColumn, LineIndicator, s, ws, ref newlist1, 5401, rowCount);
                    newlist.AddRange(newlist1);
                    return;
                }
                ws = CalTransCTF(lastRowNum, list, StartColumn, LineIndicator, s, ws, ref newlist1, 5401, 5500);
                newlist.AddRange(newlist1);
                newlist1.Clear();
                LogCTF(5500);
                GC.Collect();
                //===================================================================================5600=============================================================
                if (rowCount <= 5600)
                {
                    ws = CalTransCTF(lastRowNum, list, StartColumn, LineIndicator, s, ws, ref newlist1, 5501, rowCount);
                    newlist.AddRange(newlist1);
                    return;
                }
                ws = CalTransCTF(lastRowNum, list, StartColumn, LineIndicator, s, ws, ref newlist1, 5501, 5600);
                newlist.AddRange(newlist1);
                newlist1.Clear();
                LogCTF(5600);
                GC.Collect();
                //===================================================================================5700=============================================================
                if (rowCount <= 5700)
                {
                    ws = CalTransCTF(lastRowNum, list, StartColumn, LineIndicator, s, ws, ref newlist1, 5601, rowCount);
                    newlist.AddRange(newlist1);
                    return;
                }
                ws = CalTransCTF(lastRowNum, list, StartColumn, LineIndicator, s, ws, ref newlist1, 5601, 5700);
                newlist.AddRange(newlist1);
                newlist1.Clear();
                LogCTF(5700);
                GC.Collect();
                //===================================================================================5800=============================================================
                if (rowCount <= 5800)
                {
                    ws = CalTransCTF(lastRowNum, list, StartColumn, LineIndicator, s, ws, ref newlist1, 5701, rowCount);
                    newlist.AddRange(newlist1);
                    return;
                }
                ws = CalTransCTF(lastRowNum, list, StartColumn, LineIndicator, s, ws, ref newlist1, 5701, 5800);
                newlist.AddRange(newlist1);
                newlist1.Clear();
                LogCTF(5800);
                GC.Collect();
                //===================================================================================5900=============================================================
                if (rowCount <= 5900)
                {
                    ws = CalTransCTF(lastRowNum, list, StartColumn, LineIndicator, s, ws, ref newlist1, 5801, rowCount);
                    newlist.AddRange(newlist1);
                    return;
                }
                ws = CalTransCTF(lastRowNum, list, StartColumn, LineIndicator, s, ws, ref newlist1, 5801, 5900);
                newlist.AddRange(newlist1);
                newlist1.Clear();
                LogCTF(5900);
                GC.Collect();
                //===================================================================================6000=============================================================
                if (rowCount <= 6000)
                {
                    ws = CalTransCTF(lastRowNum, list, StartColumn, LineIndicator, s, ws, ref newlist1, 5901, rowCount);
                    newlist.AddRange(newlist1);
                    return;
                }
                ws = CalTransCTF(lastRowNum, list, StartColumn, LineIndicator, s, ws, ref newlist1, 5901, 6000);
                newlist.AddRange(newlist1);
                newlist1.Clear();
                LogCTF(6000);
                GC.Collect();
                //===================================================================================6100=============================================================
                if (rowCount <= 6100)
                {
                    ws = CalTransCTF(lastRowNum, list, StartColumn, LineIndicator, s, ws, ref newlist1, 6001, rowCount);
                    newlist.AddRange(newlist1);
                    return;
                }
                ws = CalTransCTF(lastRowNum, list, StartColumn, LineIndicator, s, ws, ref newlist1, 6001, 6100);
                newlist.AddRange(newlist1);
                newlist1.Clear();
                LogCTF(6100);
                GC.Collect();
                //===================================================================================6200=============================================================
                if (rowCount <= 6200)
                {
                    ws = CalTransCTF(lastRowNum, list, StartColumn, LineIndicator, s, ws, ref newlist1, 6101, rowCount);
                    newlist.AddRange(newlist1);
                    return;
                }
                ws = CalTransCTF(lastRowNum, list, StartColumn, LineIndicator, s, ws, ref newlist1, 6101, 6200);
                newlist.AddRange(newlist1);
                newlist1.Clear();
                LogCTF(6200);
                GC.Collect();
                //===================================================================================6300=============================================================
                if (rowCount <= 6300)
                {
                    ws = CalTransCTF(lastRowNum, list, StartColumn, LineIndicator, s, ws, ref newlist1, 6201, rowCount);
                    newlist.AddRange(newlist1);
                    return;
                }
                ws = CalTransCTF(lastRowNum, list, StartColumn, LineIndicator, s, ws, ref newlist1, 6201, 6300);
                newlist.AddRange(newlist1);
                newlist1.Clear();
                LogCTF(6300);
                GC.Collect();
                //===================================================================================6400=============================================================
                if (rowCount <= 6400)
                {
                    ws = CalTransCTF(lastRowNum, list, StartColumn, LineIndicator, s, ws, ref newlist1, 6301, rowCount);
                    newlist.AddRange(newlist1);
                    return;
                }
                ws = CalTransCTF(lastRowNum, list, StartColumn, LineIndicator, s, ws, ref newlist1, 6301, 6400);
                newlist.AddRange(newlist1);
                newlist1.Clear();
                LogCTF(6400);
                GC.Collect();
                //===================================================================================6500=============================================================
                if (rowCount <= 6500)
                {
                    ws = CalTransCTF(lastRowNum, list, StartColumn, LineIndicator, s, ws, ref newlist1, 6401, rowCount);
                    newlist.AddRange(newlist1);
                    return;
                }
                ws = CalTransCTF(lastRowNum, list, StartColumn, LineIndicator, s, ws, ref newlist1, 6401, 6500);
                newlist.AddRange(newlist1);
                newlist1.Clear();
                LogCTF(6500);
                GC.Collect();
                //===================================================================================6600=============================================================
                if (rowCount <= 6600)
                {
                    ws = CalTransCTF(lastRowNum, list, StartColumn, LineIndicator, s, ws, ref newlist1, 6501, rowCount);
                    newlist.AddRange(newlist1);
                    return;
                }
                ws = CalTransCTF(lastRowNum, list, StartColumn, LineIndicator, s, ws, ref newlist1, 6501, 6600);
                newlist.AddRange(newlist1);
                newlist1.Clear();
                LogCTF(6600);
                GC.Collect();
                //===================================================================================6700=============================================================
                if (rowCount <= 6700)
                {
                    ws = CalTransCTF(lastRowNum, list, StartColumn, LineIndicator, s, ws, ref newlist1, 6601, rowCount);
                    newlist.AddRange(newlist1);
                    return;
                }
                ws = CalTransCTF(lastRowNum, list, StartColumn, LineIndicator, s, ws, ref newlist1, 6601, 6700);
                newlist.AddRange(newlist1);
                newlist1.Clear();
                LogCTF(6700);
                GC.Collect();
                //===================================================================================6800=============================================================
                if (rowCount <= 6800)
                {
                    ws = CalTransCTF(lastRowNum, list, StartColumn, LineIndicator, s, ws, ref newlist1, 6701, rowCount);
                    newlist.AddRange(newlist1);
                    return;
                }
                ws = CalTransCTF(lastRowNum, list, StartColumn, LineIndicator, s, ws, ref newlist1, 6701, 6800);
                newlist.AddRange(newlist1);
                newlist1.Clear();
                LogCTF(6800);
                GC.Collect();
                //===================================================================================6900=============================================================
                if (rowCount <= 6900)
                {
                    ws = CalTransCTF(lastRowNum, list, StartColumn, LineIndicator, s, ws, ref newlist1, 6801, rowCount);
                    newlist.AddRange(newlist1);
                    return;
                }
                ws = CalTransCTF(lastRowNum, list, StartColumn, LineIndicator, s, ws, ref newlist1, 6801, 6900);
                newlist.AddRange(newlist1);
                newlist1.Clear();
                LogCTF(6900);
                GC.Collect();
                //===================================================================================7000=============================================================
                if (rowCount <= 7000)
                {
                    ws = CalTransCTF(lastRowNum, list, StartColumn, LineIndicator, s, ws, ref newlist1, 6901, rowCount);
                    newlist.AddRange(newlist1);
                    return;
                }
                ws = CalTransCTF(lastRowNum, list, StartColumn, LineIndicator, s, ws, ref newlist1, 6901, 7000);
                newlist.AddRange(newlist1);
                newlist1.Clear();
                LogCTF(7000);
                GC.Collect();
                //===================================================================================7100=============================================================
                if (rowCount <= 7100)
                {
                    ws = CalTransCTF(lastRowNum, list, StartColumn, LineIndicator, s, ws, ref newlist1, 7001, rowCount);
                    newlist.AddRange(newlist1);
                    return;
                }
                ws = CalTransCTF(lastRowNum, list, StartColumn, LineIndicator, s, ws, ref newlist1, 7001, 7100);
                newlist.AddRange(newlist1);
                newlist1.Clear();
                LogCTF(7100);
                GC.Collect();
                //===================================================================================7200=============================================================
                if (rowCount <= 7200)
                {
                    ws = CalTransCTF(lastRowNum, list, StartColumn, LineIndicator, s, ws, ref newlist1, 7101, rowCount);
                    newlist.AddRange(newlist1);
                    return;
                }
                ws = CalTransCTF(lastRowNum, list, StartColumn, LineIndicator, s, ws, ref newlist1, 7101, 7200);
                newlist.AddRange(newlist1);
                newlist1.Clear();
                LogCTF(7200);
                GC.Collect();
                //===================================================================================7300=============================================================
                if (rowCount <= 7300)
                {
                    ws = CalTransCTF(lastRowNum, list, StartColumn, LineIndicator, s, ws, ref newlist1, 7201, rowCount);
                    newlist.AddRange(newlist1);
                    return;
                }
                ws = CalTransCTF(lastRowNum, list, StartColumn, LineIndicator, s, ws, ref newlist1, 7201, 7300);
                newlist.AddRange(newlist1);
                newlist1.Clear();
                LogCTF(7300);
                GC.Collect();
                //===================================================================================7400=============================================================
                if (rowCount <= 7400)
                {
                    ws = CalTransCTF(lastRowNum, list, StartColumn, LineIndicator, s, ws, ref newlist1, 7301, rowCount);
                    newlist.AddRange(newlist1);
                    return;
                }
                ws = CalTransCTF(lastRowNum, list, StartColumn, LineIndicator, s, ws, ref newlist1, 7301, 7400);
                newlist.AddRange(newlist1);
                newlist1.Clear();
                LogCTF(7400);
                GC.Collect();
                //===================================================================================7500=============================================================
                if (rowCount <= 7500)
                {
                    ws = CalTransCTF(lastRowNum, list, StartColumn, LineIndicator, s, ws, ref newlist1, 7401, rowCount);
                    newlist.AddRange(newlist1);
                    return;
                }
                ws = CalTransCTF(lastRowNum, list, StartColumn, LineIndicator, s, ws, ref newlist1, 7401, 7500);
                newlist.AddRange(newlist1);
                newlist1.Clear();
                LogCTF(7500);
                GC.Collect();
                //===================================================================================7600=============================================================
                if (rowCount <= 7600)
                {
                    ws = CalTransCTF(lastRowNum, list, StartColumn, LineIndicator, s, ws, ref newlist1, 7501, rowCount);
                    newlist.AddRange(newlist1);
                    return;
                }
                ws = CalTransCTF(lastRowNum, list, StartColumn, LineIndicator, s, ws, ref newlist1, 7501, 7600);
                newlist.AddRange(newlist1);
                newlist1.Clear();
                LogCTF(7600);
                GC.Collect();
                //===================================================================================7700=============================================================
                if (rowCount <= 7700)
                {
                    ws = CalTransCTF(lastRowNum, list, StartColumn, LineIndicator, s, ws, ref newlist1, 7601, rowCount);
                    newlist.AddRange(newlist1);
                    return;
                }
                ws = CalTransCTF(lastRowNum, list, StartColumn, LineIndicator, s, ws, ref newlist1, 7601, 7700);
                newlist.AddRange(newlist1);
                newlist1.Clear();
                LogCTF(7700);
                GC.Collect();
                //===================================================================================7800=============================================================
                if (rowCount <= 7800)
                {
                    ws = CalTransCTF(lastRowNum, list, StartColumn, LineIndicator, s, ws, ref newlist1, 7701, rowCount);
                    newlist.AddRange(newlist1);
                    return;
                }
                ws = CalTransCTF(lastRowNum, list, StartColumn, LineIndicator, s, ws, ref newlist1, 7701, 7800);
                newlist.AddRange(newlist1);
                newlist1.Clear();
                LogCTF(7800);
                GC.Collect();
                //===================================================================================7900=============================================================
                if (rowCount <= 7900)
                {
                    ws = CalTransCTF(lastRowNum, list, StartColumn, LineIndicator, s, ws, ref newlist1, 7801, rowCount);
                    newlist.AddRange(newlist1);
                    return;
                }
                ws = CalTransCTF(lastRowNum, list, StartColumn, LineIndicator, s, ws, ref newlist1, 7801, 7900);
                newlist.AddRange(newlist1);
                newlist1.Clear();
                LogCTF(7900);
                GC.Collect();
                //===================================================================================8000=============================================================
                if (rowCount <= 8000)
                {
                    ws = CalTransCTF(lastRowNum, list, StartColumn, LineIndicator, s, ws, ref newlist1, 7901, rowCount);
                    newlist.AddRange(newlist1);
                    return;
                }
                ws = CalTransCTF(lastRowNum, list, StartColumn, LineIndicator, s, ws, ref newlist1, 7901, 8000);
                newlist.AddRange(newlist1);
                newlist1.Clear();
                LogCTF(8000);
                GC.Collect();
                //===================================================================================8100=============================================================
                if (rowCount <= 8100)
                {
                    ws = CalTransCTF(lastRowNum, list, StartColumn, LineIndicator, s, ws, ref newlist1, 8001, rowCount);
                    newlist.AddRange(newlist1);
                    return;
                }
                ws = CalTransCTF(lastRowNum, list, StartColumn, LineIndicator, s, ws, ref newlist1, 8001, 8100);
                newlist.AddRange(newlist1);
                newlist1.Clear();
                LogCTF(8100);
                GC.Collect();
                //===================================================================================8200=============================================================
                if (rowCount <= 8200)
                {
                    ws = CalTransCTF(lastRowNum, list, StartColumn, LineIndicator, s, ws, ref newlist1, 8101, rowCount);
                    newlist.AddRange(newlist1);
                    return;
                }
                ws = CalTransCTF(lastRowNum, list, StartColumn, LineIndicator, s, ws, ref newlist1, 8101, 8200);
                newlist.AddRange(newlist1);
                newlist1.Clear();
                LogCTF(8200);
                GC.Collect();
                //===================================================================================8300=============================================================
                if (rowCount <= 8300)
                {
                    ws = CalTransCTF(lastRowNum, list, StartColumn, LineIndicator, s, ws, ref newlist1, 8201, rowCount);
                    newlist.AddRange(newlist1);
                    return;
                }
                ws = CalTransCTF(lastRowNum, list, StartColumn, LineIndicator, s, ws, ref newlist1, 8201, 8300);
                newlist.AddRange(newlist1);
                newlist1.Clear();
                LogCTF(8300);
                GC.Collect();
                //===================================================================================8400=============================================================
                if (rowCount <= 8400)
                {
                    ws = CalTransCTF(lastRowNum, list, StartColumn, LineIndicator, s, ws, ref newlist1, 8301, rowCount);
                    newlist.AddRange(newlist1);
                    return;
                }
                ws = CalTransCTF(lastRowNum, list, StartColumn, LineIndicator, s, ws, ref newlist1, 8301, 8400);
                newlist.AddRange(newlist1);
                newlist1.Clear();
                LogCTF(8400);
                GC.Collect();
                //===================================================================================8500=============================================================
                if (rowCount <= 8500)
                {
                    ws = CalTransCTF(lastRowNum, list, StartColumn, LineIndicator, s, ws, ref newlist1, 8401, rowCount);
                    newlist.AddRange(newlist1);
                    return;
                }
                ws = CalTransCTF(lastRowNum, list, StartColumn, LineIndicator, s, ws, ref newlist1, 8401, 8500);
                newlist.AddRange(newlist1);
                newlist1.Clear();
                LogCTF(8500);
                GC.Collect();
                //===================================================================================8600=============================================================
                if (rowCount <= 8600)
                {
                    ws = CalTransCTF(lastRowNum, list, StartColumn, LineIndicator, s, ws, ref newlist1, 8501, rowCount);
                    newlist.AddRange(newlist1);
                    return;
                }
                ws = CalTransCTF(lastRowNum, list, StartColumn, LineIndicator, s, ws, ref newlist1, 8501, 8600);
                newlist.AddRange(newlist1);
                newlist1.Clear();
                LogCTF(8600);
                GC.Collect();
                //===================================================================================8700=============================================================
                if (rowCount <= 8700)
                {
                    ws = CalTransCTF(lastRowNum, list, StartColumn, LineIndicator, s, ws, ref newlist1, 8601, rowCount);
                    newlist.AddRange(newlist1);
                    return;
                }
                ws = CalTransCTF(lastRowNum, list, StartColumn, LineIndicator, s, ws, ref newlist1, 8601, 8700);
                newlist.AddRange(newlist1);
                newlist1.Clear();
                LogCTF(8700);
                GC.Collect();
                //===================================================================================8800=============================================================
                if (rowCount <= 8800)
                {
                    ws = CalTransCTF(lastRowNum, list, StartColumn, LineIndicator, s, ws, ref newlist1, 8701, rowCount);
                    newlist.AddRange(newlist1);
                    return;
                }
                ws = CalTransCTF(lastRowNum, list, StartColumn, LineIndicator, s, ws, ref newlist1, 8701, 8800);
                newlist.AddRange(newlist1);
                newlist1.Clear();
                LogCTF(8800);
                GC.Collect();
                //===================================================================================8900=============================================================
                if (rowCount <= 8900)
                {
                    ws = CalTransCTF(lastRowNum, list, StartColumn, LineIndicator, s, ws, ref newlist1, 8801, rowCount);
                    newlist.AddRange(newlist1);
                    return;
                }
                ws = CalTransCTF(lastRowNum, list, StartColumn, LineIndicator, s, ws, ref newlist1, 8801, 8900);
                newlist.AddRange(newlist1);
                newlist1.Clear();
                LogCTF(8900);
                GC.Collect();
                //===================================================================================9000=============================================================
                if (rowCount <= 9000)
                {
                    ws = CalTransCTF(lastRowNum, list, StartColumn, LineIndicator, s, ws, ref newlist1, 8901, rowCount);
                    newlist.AddRange(newlist1);
                    return;
                }
                ws = CalTransCTF(lastRowNum, list, StartColumn, LineIndicator, s, ws, ref newlist1, 8901, 9000);
                newlist.AddRange(newlist1);
                newlist1.Clear();
                LogCTF(9000);
                GC.Collect();
                //===================================================================================9100=============================================================
                if (rowCount <= 9100)
                {
                    ws = CalTransCTF(lastRowNum, list, StartColumn, LineIndicator, s, ws, ref newlist1, 9001, rowCount);
                    newlist.AddRange(newlist1);
                    return;
                }
                ws = CalTransCTF(lastRowNum, list, StartColumn, LineIndicator, s, ws, ref newlist1, 9001, 9100);
                newlist.AddRange(newlist1);
                newlist1.Clear();
                LogCTF(9100);
                GC.Collect();
                //===================================================================================9200=============================================================
                if (rowCount <= 9200)
                {
                    ws = CalTransCTF(lastRowNum, list, StartColumn, LineIndicator, s, ws, ref newlist1, 9101, rowCount);
                    newlist.AddRange(newlist1);
                    return;
                }
                ws = CalTransCTF(lastRowNum, list, StartColumn, LineIndicator, s, ws, ref newlist1, 9101, 9200);
                newlist.AddRange(newlist1);
                newlist1.Clear();
                LogCTF(9200);
                GC.Collect();
                //===================================================================================9300=============================================================
                if (rowCount <= 9300)
                {
                    ws = CalTransCTF(lastRowNum, list, StartColumn, LineIndicator, s, ws, ref newlist1, 9201, rowCount);
                    newlist.AddRange(newlist1);
                    return;
                }
                ws = CalTransCTF(lastRowNum, list, StartColumn, LineIndicator, s, ws, ref newlist1, 9201, 9300);
                newlist.AddRange(newlist1);
                newlist1.Clear();
                LogCTF(9300);
                GC.Collect();
                //===================================================================================9400=============================================================
                if (rowCount <= 9400)
                {
                    ws = CalTransCTF(lastRowNum, list, StartColumn, LineIndicator, s, ws, ref newlist1, 9301, rowCount);
                    newlist.AddRange(newlist1);
                    return;
                }
                ws = CalTransCTF(lastRowNum, list, StartColumn, LineIndicator, s, ws, ref newlist1, 9301, 9400);
                newlist.AddRange(newlist1);
                newlist1.Clear();
                LogCTF(9400);
                GC.Collect();
                //===================================================================================9500=============================================================
                if (rowCount <= 9500)
                {
                    ws = CalTransCTF(lastRowNum, list, StartColumn, LineIndicator, s, ws, ref newlist1, 9401, rowCount);
                    newlist.AddRange(newlist1);
                    return;
                }
                ws = CalTransCTF(lastRowNum, list, StartColumn, LineIndicator, s, ws, ref newlist1, 9401, 9500);
                newlist.AddRange(newlist1);
                newlist1.Clear();
                LogCTF(9500);
                GC.Collect();
                //===================================================================================9600=============================================================
                if (rowCount <= 9600)
                {
                    ws = CalTransCTF(lastRowNum, list, StartColumn, LineIndicator, s, ws, ref newlist1, 9501, rowCount);
                    newlist.AddRange(newlist1);
                    return;
                }
                ws = CalTransCTF(lastRowNum, list, StartColumn, LineIndicator, s, ws, ref newlist1, 9501, 9600);
                newlist.AddRange(newlist1);
                newlist1.Clear();
                LogCTF(9600);
                GC.Collect();
                //===================================================================================9700=============================================================
                if (rowCount <= 9700)
                {
                    ws = CalTransCTF(lastRowNum, list, StartColumn, LineIndicator, s, ws, ref newlist1, 9601, rowCount);
                    newlist.AddRange(newlist1);
                    return;
                }
                ws = CalTransCTF(lastRowNum, list, StartColumn, LineIndicator, s, ws, ref newlist1, 9601, 9700);
                newlist.AddRange(newlist1);
                newlist1.Clear();
                LogCTF(9700);
                GC.Collect();
                //===================================================================================9800=============================================================
                if (rowCount <= 9800)
                {
                    ws = CalTransCTF(lastRowNum, list, StartColumn, LineIndicator, s, ws, ref newlist1, 9701, rowCount);
                    newlist.AddRange(newlist1);
                    return;
                }
                ws = CalTransCTF(lastRowNum, list, StartColumn, LineIndicator, s, ws, ref newlist1, 9701, 9800);
                newlist.AddRange(newlist1);
                newlist1.Clear();
                LogCTF(9800);
                GC.Collect();
                //===================================================================================9900=============================================================
                if (rowCount <= 9900)
                {
                    ws = CalTransCTF(lastRowNum, list, StartColumn, LineIndicator, s, ws, ref newlist1, 9801, rowCount);
                    newlist.AddRange(newlist1);
                    return;
                }
                ws = CalTransCTF(lastRowNum, list, StartColumn, LineIndicator, s, ws, ref newlist1, 9801, 9900);
                newlist.AddRange(newlist1);
                newlist1.Clear();
                LogCTF(9900);
                GC.Collect();
                //===================================================================================10000=============================================================
                if (rowCount <= 10000)
                {
                    ws = CalTransCTF(lastRowNum, list, StartColumn, LineIndicator, s, ws, ref newlist1, 9901, rowCount);
                    newlist.AddRange(newlist1);
                    return;
                }
                ws = CalTransCTF(lastRowNum, list, StartColumn, LineIndicator, s, ws, ref newlist1, 9901, 10000);
                newlist.AddRange(newlist1);
                LogCTF(10000);
            }
        }
    }
}
