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
        private Microsoft.Office.Interop.Excel.Worksheet CalAMU(string lastRowNum, List<ExcelAddIn4.Common.AMUEntityForSave> list, string StartColumn, string LineIndicator, string s, Microsoft.Office.Interop.Excel.Worksheet ws, ref List<ExcelAddIn4.Common.AMUEntityForSave> newlist, int start, int over)
        {
            for (int j = start; j <= over; j++)
            {
                System.Windows.Forms.Application.DoEvents();
                for (int i = 0; i < list.Count; i++)
                {
                    ExcelAddIn4.Common.AMUEntityForSave re;
                    try
                    {
                        try
                        {
                            var lastrow = ws.Cells.Find("*", ws.Cells[1, 1], Microsoft.Office.Interop.Excel.XlFindLookIn.xlFormulas, Microsoft.Office.Interop.Excel.XlLookAt.xlPart, Microsoft.Office.Interop.Excel.XlSearchOrder.xlByRows, Microsoft.Office.Interop.Excel.XlSearchDirection.xlPrevious, false, Type.Missing, Type.Missing);
                            string LastRowNumber = Finance_Tools.RemoveNotNumber(lastrow.Address);
                            if (LastRowNumber != Ribbon2.LastRowNumber)
                            {
                                //Copy sheet to make ws not disappear.
                                Globals.ThisAddIn.Application.DisplayAlerts = false;
                                //Globals.ThisAddIn.Application.Worksheets.Add(Type.Missing, Globals.ThisAddIn.Application.ActiveWorkbook.Worksheets[/*"Sheet1"*/Globals.ThisAddIn.Application.ActiveWorkbook.Worksheets.Count], 1, Type.Missing);
                                //Microsoft.Office.Interop.Excel.Worksheet ws2 = (Microsoft.Office.Interop.Excel.Worksheet)Globals.ThisAddIn.Application.Worksheets.get_Item(Globals.ThisAddIn.Application.ActiveWorkbook.Worksheets.Count);
                                //ws2.Name = "RSTMPTemplate" + Guid.NewGuid().ToString().Replace("-", "").Remove(17);
                                Ribbon2.wsRrigin.get_Range("A1", Ribbon2.LastColumnName + Ribbon2.LastRowNumber).Copy(ws.get_Range("A1"));
                                Globals.ThisAddIn.Application.DisplayAlerts = true;
                                System.Windows.Forms.Clipboard.Clear();
                            }
                        }
                        catch
                        {
                            //Copy sheet to make ws not disappear.
                            Globals.ThisAddIn.Application.DisplayAlerts = false;
                            //Globals.ThisAddIn.Application.Worksheets.Add(Type.Missing, Globals.ThisAddIn.Application.ActiveWorkbook.Worksheets[/*"Sheet1"*/Globals.ThisAddIn.Application.ActiveWorkbook.Worksheets.Count], 1, Type.Missing);
                            //Microsoft.Office.Interop.Excel.Worksheet ws2 = (Microsoft.Office.Interop.Excel.Worksheet)Globals.ThisAddIn.Application.Worksheets.get_Item(Globals.ThisAddIn.Application.ActiveWorkbook.Worksheets.Count);
                            //ws2.Name = "RSTMPTemplate" + Guid.NewGuid().ToString().Replace("-", "").Remove(17);
                            Ribbon2.wsRrigin.get_Range("A1", Ribbon2.LastColumnName + Ribbon2.LastRowNumber).Copy(ws.get_Range("A1"));
                            Globals.ThisAddIn.Application.DisplayAlerts = true;
                            System.Windows.Forms.Clipboard.Clear();
                        }

                        string str = list[i].ToString();
                        if (string.IsNullOrEmpty(str)) continue;

                        if (!IsStartingCellRowContainLineIndicator(StartColumn + (j + int.Parse(s)).ToString(), LineIndicator)) continue;

                        re = new ExcelAddIn4.Common.AMUEntityForSave();
                        //re.SelectionCriteria = new Common3.SelectionCriteria();
                        //re.NewSettings = new Common3.NewSettings();
                        re.JournalNumber = string.IsNullOrEmpty(list[i].JournalNumber) ? "" : GetEntityValueInExcel(list[i].JournalNumber, s, j, ws);

                        re.JournalLineNumber = string.IsNullOrEmpty(list[i].JournalLineNumber) ? "" : GetEntityValueInExcel(list[i].JournalLineNumber, s, j, ws);
                        re.Ledger = string.IsNullOrEmpty(list[i].Ledger) ? "" : GetEntityValueInExcel(list[i].Ledger, s, j, ws);
                        re.AccountCode = string.IsNullOrEmpty(list[i].AccountCode) ? "" : GetEntityValueInExcel(list[i].AccountCode, s, j, ws);
                        re.AccountingPeriod = GetEntityValueInExcel(list[i].AccountingPeriod, s, j, ws).Replace("/", "");
                        if (re.AccountingPeriod != null)
                        {
                            bool result;
                            string y3 = string.Empty;
                            string m = string.Empty;
                            if (re.AccountingPeriod != "")
                            {
                                result = Finance_Tools.IsPeriodString(re.AccountingPeriod);
                                if (result)
                                {
                                    //Format(Right(2012/001,3)&Left(2012/001,4),"0000000")
                                    y3 = re.AccountingPeriod.Replace("/", "").Substring(0, 4);
                                    m = re.AccountingPeriod.Replace("/", "").Substring(4, 3);
                                    int a = 0;
                                    if ((int.TryParse(y3, out a) == false) || (int.TryParse(m, out a) == false))
                                    {
                                        throw new Exception("AccountingPeriod format is not correct.");
                                    }
                                    re.AccountingPeriod = m.PadLeft(3, '0') + y3;//;
                                }
                                else
                                {
                                    throw new Exception("AccountingPeriod format is not correct.");
                                }
                            }
                        }
                        re.TransactionDate = string.IsNullOrEmpty(list[i].TransactionDate) ? "" : GetEntityValueInExcel(list[i].TransactionDate, s, j, ws);
                        re.JournalType = string.IsNullOrEmpty(list[i].JournalType) ? "" : GetEntityValueInExcel(list[i].JournalType, s, j, ws);
                        re.TransactionReference = string.IsNullOrEmpty(list[i].TransactionReference) ? "" : GetEntityValueInExcel(list[i].TransactionReference, s, j, ws);
                        re.AllocationMarker = string.IsNullOrEmpty(list[i].AllocationMarker) ? "" : GetEntityValueInExcel(list[i].AllocationMarker, s, j, ws);
                        re.AnalysisCode1 = string.IsNullOrEmpty(list[i].AnalysisCode1) ? "" : GetEntityValueInExcel(list[i].AnalysisCode1, s, j, ws);
                        re.AnalysisCode2 = string.IsNullOrEmpty(list[i].AnalysisCode2) ? "" : GetEntityValueInExcel(list[i].AnalysisCode2, s, j, ws);
                        re.AnalysisCode3 = string.IsNullOrEmpty(list[i].AnalysisCode3) ? "" : GetEntityValueInExcel(list[i].AnalysisCode3, s, j, ws);
                        re.AnalysisCode4 = string.IsNullOrEmpty(list[i].AnalysisCode4) ? "" : GetEntityValueInExcel(list[i].AnalysisCode4, s, j, ws);
                        re.AnalysisCode5 = string.IsNullOrEmpty(list[i].AnalysisCode5) ? "" : GetEntityValueInExcel(list[i].AnalysisCode5, s, j, ws);
                        re.AnalysisCode6 = string.IsNullOrEmpty(list[i].AnalysisCode6) ? "" : GetEntityValueInExcel(list[i].AnalysisCode6, s, j, ws);
                        re.AnalysisCode7 = string.IsNullOrEmpty(list[i].AnalysisCode7) ? "" : GetEntityValueInExcel(list[i].AnalysisCode7, s, j, ws);
                        re.AnalysisCode8 = string.IsNullOrEmpty(list[i].AnalysisCode8) ? "" : GetEntityValueInExcel(list[i].AnalysisCode8, s, j, ws);
                        re.AnalysisCode9 = string.IsNullOrEmpty(list[i].AnalysisCode9) ? "" : GetEntityValueInExcel(list[i].AnalysisCode9, s, j, ws);
                        re.AnalysisCode10 = string.IsNullOrEmpty(list[i].AnalysisCode10) ? "" : GetEntityValueInExcel(list[i].AnalysisCode10, s, j, ws);
                        re.DebitCredit = string.IsNullOrEmpty(list[i].DebitCredit) ? "" : GetEntityValueInExcel(list[i].DebitCredit, s, j, ws);
                        newlist.Add(re);
                    }
                    catch (Exception ex)
                    {
                        //errorStr += "The data in Line " + (j + int.Parse(s)).ToString() + " has error! " + ex.Message + "\r\n";
                        LogErrorAMU("The data in Line " + (j + int.Parse(s)).ToString() + " has error! " + ex.Message + "\r\n");
                        continue;
                    }

                }
                //if (j % 2 == 0)
                //{

                //}
            }
            return ws;
        }
        private delegate void dgetLogErrorAMU(string strLogessage);

        public void LogErrorAMU(string LogText)
        {
            try
            {
                if (AMUPostFrm.richTextBox1.InvokeRequired)
                {
                    AMUPostFrm.richTextBox1.BeginInvoke(new dgetLogErrorAMU(LogErrorAMU), new object[] { LogText });
                }
                else
                {
                    //send info to progressBar1
                    AMUPostFrm.richTextBox1.Text += LogText;
                }
            }
            catch
            {
            }
        }

        private delegate void dgetLogTextAMU(int strLogessage);

        public void LogAMU(int LogText)
        {
            try
            {
                if (AMUPostFrm.progressBar1.InvokeRequired)
                {
                    AMUPostFrm.progressBar1.BeginInvoke(new dgetLogTextAMU(LogAMU), new object[] { LogText });
                }
                else
                {
                    //send info to progressBar1
                    AMUPostFrm.progressBar1.Value = LogText;
                }
            }
            catch
            {
            }
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
        public void BigCalculateAMU(string lastRowNum, List<ExcelAddIn4.Common.AMUEntityForSave> list, string StartColumn, string LineIndicator, string s, Microsoft.Office.Interop.Excel.Worksheet ws, ref List<ExcelAddIn4.Common.AMUEntityForSave> newlist)
        {
            //Copy sheet to make ws not disappear.
            Globals.ThisAddIn.Application.DisplayAlerts = false;
            Globals.ThisAddIn.Application.Worksheets.Add(Type.Missing, Globals.ThisAddIn.Application.ActiveWorkbook.Worksheets[/*"Sheet1"*/Globals.ThisAddIn.Application.ActiveWorkbook.Worksheets.Count], 1, Type.Missing);
            Microsoft.Office.Interop.Excel.Worksheet ws1 = (Microsoft.Office.Interop.Excel.Worksheet)Globals.ThisAddIn.Application.Worksheets.get_Item(Globals.ThisAddIn.Application.ActiveWorkbook.Worksheets.Count);
            ws1.Name = "RSTMPTemplate" + Guid.NewGuid().ToString().Replace("-", "").Remove(17);
            ws.get_Range("A1", Ribbon2.LastColumnName + Ribbon2.LastRowNumber).Copy(ws1.get_Range("A1"));
            Globals.ThisAddIn.Application.DisplayAlerts = true;
            System.Windows.Forms.Clipboard.SetText("\r\n");

            int rowCount = int.Parse(lastRowNum) - int.Parse(s);
            if (rowCount > 500)
            {
                AMUPostFrm.progressBar1.Visible = true;
                AMUPostFrm.progressBar1.Maximum = rowCount;
            }
            //================================================================below 300======================================================
            List<ExcelAddIn4.Common.AMUEntityForSave> newlist1 = new List<ExcelAddIn4.Common.AMUEntityForSave>();
            if (rowCount <= 300)
            {
                ws1 = CalAMU(lastRowNum, list, StartColumn, LineIndicator, s, ws1, ref newlist1, 0, rowCount);
                newlist.AddRange(newlist1);
            }
            else if (rowCount > 300)//================================================================above 300======================================================
            {
                ws1 = CalAMU(lastRowNum, list, StartColumn, LineIndicator, s, ws1, ref newlist1, 0, 100);
                newlist.AddRange(newlist1);
                newlist1.Clear();
                System.Threading.Thread.Sleep(2000); GC.Collect();
                ws1 = CalAMU(lastRowNum, list, StartColumn, LineIndicator, s, ws1, ref newlist1, 101, 200);
                newlist.AddRange(newlist1);
                newlist1.Clear();
                System.Threading.Thread.Sleep(2000); GC.Collect();
                ws1 = CalAMU(lastRowNum, list, StartColumn, LineIndicator, s, ws1, ref newlist1, 201, 300);
                newlist.AddRange(newlist1);
                newlist1.Clear();
                System.Threading.Thread.Sleep(2000); GC.Collect();
                //===================================================================================400=============================================================
                if (rowCount <= 400)
                {
                    ws1 = CalAMU(lastRowNum, list, StartColumn, LineIndicator, s, ws1, ref newlist1, 301, rowCount);
                    newlist.AddRange(newlist1);
                    return;
                }
                ws1 = CalAMU(lastRowNum, list, StartColumn, LineIndicator, s, ws1, ref newlist1, 301, 400);
                newlist.AddRange(newlist1);
                newlist1.Clear();
                System.Threading.Thread.Sleep(2000); GC.Collect();
                //===================================================================================500=============================================================
                if (rowCount <= 500)
                {
                    ws1 = CalAMU(lastRowNum, list, StartColumn, LineIndicator, s, ws1, ref newlist1, 401, rowCount);
                    newlist.AddRange(newlist1);
                    return;
                }
                ws1 = CalAMU(lastRowNum, list, StartColumn, LineIndicator, s, ws1, ref newlist1, 401, 500);
                newlist.AddRange(newlist1);
                newlist1.Clear();
                LogAMU(500);
                System.Threading.Thread.Sleep(2000); GC.Collect();
                //===================================================================================600=============================================================
                if (rowCount <= 600)
                {
                    ws1 = CalAMU(lastRowNum, list, StartColumn, LineIndicator, s, ws1, ref newlist1, 501, rowCount);
                    newlist.AddRange(newlist1);
                    return;
                }
                ws1 = CalAMU(lastRowNum, list, StartColumn, LineIndicator, s, ws1, ref newlist1, 501, 600);
                newlist.AddRange(newlist1);
                newlist1.Clear();
                LogAMU(600);
                System.Threading.Thread.Sleep(2000); GC.Collect();
                //===================================================================================700=============================================================
                if (rowCount <= 700)
                {
                    ws1 = CalAMU(lastRowNum, list, StartColumn, LineIndicator, s, ws1, ref newlist1, 601, rowCount);
                    newlist.AddRange(newlist1);
                    return;
                }
                ws1 = CalAMU(lastRowNum, list, StartColumn, LineIndicator, s, ws1, ref newlist1, 601, 700);
                newlist.AddRange(newlist1);
                newlist1.Clear();
                LogAMU(700);
                System.Threading.Thread.Sleep(2000); GC.Collect();
                //===================================================================================800=============================================================
                if (rowCount <= 800)
                {
                    ws1 = CalAMU(lastRowNum, list, StartColumn, LineIndicator, s, ws1, ref newlist1, 701, rowCount);
                    newlist.AddRange(newlist1);
                    return;
                }
                ws1 = CalAMU(lastRowNum, list, StartColumn, LineIndicator, s, ws1, ref newlist1, 701, 800);
                newlist.AddRange(newlist1);
                newlist1.Clear();
                LogAMU(800);
                System.Threading.Thread.Sleep(2000); GC.Collect();
                //===================================================================================900=============================================================
                if (rowCount <= 900)
                {
                    ws1 = CalAMU(lastRowNum, list, StartColumn, LineIndicator, s, ws1, ref newlist1, 801, rowCount);
                    newlist.AddRange(newlist1);
                    return;
                }
                ws1 = CalAMU(lastRowNum, list, StartColumn, LineIndicator, s, ws1, ref newlist1, 801, 900);
                newlist.AddRange(newlist1);
                newlist1.Clear();
                LogAMU(900);
                System.Threading.Thread.Sleep(2000); GC.Collect();
                //===================================================================================1000=============================================================
                if (rowCount <= 1000)
                {
                    ws1 = CalAMU(lastRowNum, list, StartColumn, LineIndicator, s, ws1, ref newlist1, 901, rowCount);
                    newlist.AddRange(newlist1);
                    return;
                }
                ws1 = CalAMU(lastRowNum, list, StartColumn, LineIndicator, s, ws1, ref newlist1, 901, 1000);
                newlist.AddRange(newlist1);
                newlist1.Clear();
                LogAMU(1000);
                System.Threading.Thread.Sleep(2000); GC.Collect();
                //===================================================================================1100=============================================================
                if (rowCount <= 1100)
                {
                    ws1 = CalAMU(lastRowNum, list, StartColumn, LineIndicator, s, ws1, ref newlist1, 1001, rowCount);
                    newlist.AddRange(newlist1);
                    return;
                }
                ws1 = CalAMU(lastRowNum, list, StartColumn, LineIndicator, s, ws1, ref newlist1, 1001, 1100);
                newlist.AddRange(newlist1);
                newlist1.Clear();
                LogAMU(1100);
                System.Threading.Thread.Sleep(2000); GC.Collect();
                //===================================================================================1200=============================================================
                if (rowCount <= 1200)
                {
                    ws1 = CalAMU(lastRowNum, list, StartColumn, LineIndicator, s, ws1, ref newlist1, 1101, rowCount);
                    newlist.AddRange(newlist1);
                    return;
                }
                ws1 = CalAMU(lastRowNum, list, StartColumn, LineIndicator, s, ws1, ref newlist1, 1101, 1200);
                newlist.AddRange(newlist1);
                newlist1.Clear();
                LogAMU(1200);
                System.Threading.Thread.Sleep(2000); GC.Collect();
                //===================================================================================1300=============================================================
                if (rowCount <= 1300)
                {
                    ws1 = CalAMU(lastRowNum, list, StartColumn, LineIndicator, s, ws1, ref newlist1, 1201, rowCount);
                    newlist.AddRange(newlist1);
                    return;
                }
                ws1 = CalAMU(lastRowNum, list, StartColumn, LineIndicator, s, ws1, ref newlist1, 1201, 1300);
                newlist.AddRange(newlist1);
                newlist1.Clear();
                LogAMU(1300);
                System.Threading.Thread.Sleep(2000); GC.Collect();
                //===================================================================================1400=============================================================
                if (rowCount <= 1400)
                {
                    ws1 = CalAMU(lastRowNum, list, StartColumn, LineIndicator, s, ws1, ref newlist1, 1301, rowCount);
                    newlist.AddRange(newlist1);
                    return;
                }
                ws1 = CalAMU(lastRowNum, list, StartColumn, LineIndicator, s, ws1, ref newlist1, 1301, 1400);
                newlist.AddRange(newlist1);
                newlist1.Clear();
                LogAMU(1400);
                System.Threading.Thread.Sleep(2000); GC.Collect();
                //===================================================================================1500=============================================================
                if (rowCount <= 1500)
                {
                    ws1 = CalAMU(lastRowNum, list, StartColumn, LineIndicator, s, ws1, ref newlist1, 1401, rowCount);
                    newlist.AddRange(newlist1);
                    return;
                }
                ws1 = CalAMU(lastRowNum, list, StartColumn, LineIndicator, s, ws1, ref newlist1, 1401, 1500);
                newlist.AddRange(newlist1);
                newlist1.Clear();
                LogAMU(1500);
                System.Threading.Thread.Sleep(2000); GC.Collect();
                //===================================================================================1600=============================================================
                if (rowCount <= 1600)
                {
                    ws1 = CalAMU(lastRowNum, list, StartColumn, LineIndicator, s, ws1, ref newlist1, 1501, rowCount);
                    newlist.AddRange(newlist1);
                    return;
                }
                ws1 = CalAMU(lastRowNum, list, StartColumn, LineIndicator, s, ws1, ref newlist1, 1501, 1600);
                newlist.AddRange(newlist1);
                newlist1.Clear();
                LogAMU(1600);
                System.Threading.Thread.Sleep(2000); GC.Collect();
                //===================================================================================1700=============================================================
                if (rowCount <= 1700)
                {
                    ws1 = CalAMU(lastRowNum, list, StartColumn, LineIndicator, s, ws1, ref newlist1, 1601, rowCount);
                    newlist.AddRange(newlist1);
                    return;
                }
                ws1 = CalAMU(lastRowNum, list, StartColumn, LineIndicator, s, ws1, ref newlist1, 1601, 1700);
                newlist.AddRange(newlist1);
                newlist1.Clear();
                LogAMU(1700);
                System.Threading.Thread.Sleep(2000); GC.Collect();
                //===================================================================================1800=============================================================
                if (rowCount <= 1800)
                {
                    ws1 = CalAMU(lastRowNum, list, StartColumn, LineIndicator, s, ws1, ref newlist1, 1701, rowCount);
                    newlist.AddRange(newlist1);
                    return;
                }
                ws1 = CalAMU(lastRowNum, list, StartColumn, LineIndicator, s, ws1, ref newlist1, 1701, 1800);
                newlist.AddRange(newlist1);
                newlist1.Clear();
                LogAMU(1800);
                System.Threading.Thread.Sleep(2000); GC.Collect();
                //===================================================================================1900=============================================================
                if (rowCount <= 1900)
                {
                    ws1 = CalAMU(lastRowNum, list, StartColumn, LineIndicator, s, ws1, ref newlist1, 1801, rowCount);
                    newlist.AddRange(newlist1);
                    return;
                }
                ws1 = CalAMU(lastRowNum, list, StartColumn, LineIndicator, s, ws1, ref newlist1, 1801, 1900);
                newlist.AddRange(newlist1);
                newlist1.Clear();
                LogAMU(1900);
                System.Threading.Thread.Sleep(2000); GC.Collect();
                //===================================================================================2000=============================================================
                if (rowCount <= 2000)
                {
                    ws1 = CalAMU(lastRowNum, list, StartColumn, LineIndicator, s, ws1, ref newlist1, 1901, rowCount);
                    newlist.AddRange(newlist1);
                    return;
                }
                ws1 = CalAMU(lastRowNum, list, StartColumn, LineIndicator, s, ws1, ref newlist1, 1901, 2000);
                newlist.AddRange(newlist1);
                newlist1.Clear();
                LogAMU(2000);
                System.Threading.Thread.Sleep(2000); GC.Collect();
                //Copy sheet to make ws not disappear.
                Globals.ThisAddIn.Application.DisplayAlerts = false;
                Globals.ThisAddIn.Application.Worksheets.Add(Type.Missing, Globals.ThisAddIn.Application.ActiveWorkbook.Worksheets[/*"Sheet1"*/Globals.ThisAddIn.Application.ActiveWorkbook.Worksheets.Count], 1, Type.Missing);
                Microsoft.Office.Interop.Excel.Worksheet ws2 = (Microsoft.Office.Interop.Excel.Worksheet)Globals.ThisAddIn.Application.Worksheets.get_Item(Globals.ThisAddIn.Application.ActiveWorkbook.Worksheets.Count);
                ws2.Name = "RSTMPTemplate" + Guid.NewGuid().ToString().Replace("-", "").Remove(17);
                ws1.get_Range("A1", Ribbon2.LastColumnName + Ribbon2.LastRowNumber).Copy(ws2.get_Range("A1"));
                Globals.ThisAddIn.Application.DisplayAlerts = true;
                System.Windows.Forms.Clipboard.SetText("\r\n");
                //===================================================================================2100=============================================================
                if (rowCount <= 2100)
                {
                    ws2 = CalAMU(lastRowNum, list, StartColumn, LineIndicator, s, ws2, ref newlist1, 2001, rowCount);
                    newlist.AddRange(newlist1);
                    return;
                }
                ws2 = CalAMU(lastRowNum, list, StartColumn, LineIndicator, s, ws2, ref newlist1, 2001, 2100);
                newlist.AddRange(newlist1);
                newlist1.Clear();
                LogAMU(2100);
                System.Threading.Thread.Sleep(2000); GC.Collect();
                //===================================================================================2200=============================================================
                if (rowCount <= 2200)
                {
                    ws2 = CalAMU(lastRowNum, list, StartColumn, LineIndicator, s, ws2, ref newlist1, 2101, rowCount);
                    newlist.AddRange(newlist1);
                    return;
                }
                ws2 = CalAMU(lastRowNum, list, StartColumn, LineIndicator, s, ws2, ref newlist1, 2101, 2200);
                newlist.AddRange(newlist1);
                newlist1.Clear();
                LogAMU(2200);
                System.Threading.Thread.Sleep(2000); GC.Collect();
                //===================================================================================2300=============================================================
                if (rowCount <= 2300)
                {
                    ws2 = CalAMU(lastRowNum, list, StartColumn, LineIndicator, s, ws2, ref newlist1, 2201, rowCount);
                    newlist.AddRange(newlist1);
                    return;
                }
                ws2 = CalAMU(lastRowNum, list, StartColumn, LineIndicator, s, ws2, ref newlist1, 2201, 2300);
                newlist.AddRange(newlist1);
                newlist1.Clear();
                LogAMU(2300);
                System.Threading.Thread.Sleep(2000); GC.Collect();
                //===================================================================================2400=============================================================
                if (rowCount <= 2400)
                {
                    ws2 = CalAMU(lastRowNum, list, StartColumn, LineIndicator, s, ws2, ref newlist1, 2301, rowCount);
                    newlist.AddRange(newlist1);
                    return;
                }
                ws2 = CalAMU(lastRowNum, list, StartColumn, LineIndicator, s, ws2, ref newlist1, 2301, 2400);
                newlist.AddRange(newlist1);
                newlist1.Clear();
                LogAMU(2400);
                System.Threading.Thread.Sleep(2000); GC.Collect();
                //===================================================================================2500=============================================================
                if (rowCount <= 2500)
                {
                    ws2 = CalAMU(lastRowNum, list, StartColumn, LineIndicator, s, ws2, ref newlist1, 2401, rowCount);
                    newlist.AddRange(newlist1);
                    return;
                }
                ws2 = CalAMU(lastRowNum, list, StartColumn, LineIndicator, s, ws2, ref newlist1, 2401, 2500);
                newlist.AddRange(newlist1);
                newlist1.Clear();
                LogAMU(2500);
                System.Threading.Thread.Sleep(2000); GC.Collect();
                //===================================================================================2600=============================================================
                if (rowCount <= 2600)
                {
                    ws2 = CalAMU(lastRowNum, list, StartColumn, LineIndicator, s, ws2, ref newlist1, 2501, rowCount);
                    newlist.AddRange(newlist1);
                    return;
                }
                ws2 = CalAMU(lastRowNum, list, StartColumn, LineIndicator, s, ws2, ref newlist1, 2501, 2600);
                newlist.AddRange(newlist1);
                newlist1.Clear();
                LogAMU(2600);
                System.Threading.Thread.Sleep(2000); GC.Collect();
                //===================================================================================2700=============================================================
                if (rowCount <= 2700)
                {
                    ws2 = CalAMU(lastRowNum, list, StartColumn, LineIndicator, s, ws2, ref newlist1, 2601, rowCount);
                    newlist.AddRange(newlist1);
                    return;
                }
                ws2 = CalAMU(lastRowNum, list, StartColumn, LineIndicator, s, ws2, ref newlist1, 2601, 2700);
                newlist.AddRange(newlist1);
                newlist1.Clear();
                LogAMU(2700);
                System.Threading.Thread.Sleep(2000); GC.Collect();
                //===================================================================================2800=============================================================
                if (rowCount <= 2800)
                {
                    ws2 = CalAMU(lastRowNum, list, StartColumn, LineIndicator, s, ws2, ref newlist1, 2701, rowCount);
                    newlist.AddRange(newlist1);
                    return;
                }
                ws2 = CalAMU(lastRowNum, list, StartColumn, LineIndicator, s, ws2, ref newlist1, 2701, 2800);
                newlist.AddRange(newlist1);
                newlist1.Clear();
                LogAMU(2800);
                System.Threading.Thread.Sleep(2000); GC.Collect();
                //===================================================================================2900=============================================================
                if (rowCount <= 2900)
                {
                    ws2 = CalAMU(lastRowNum, list, StartColumn, LineIndicator, s, ws2, ref newlist1, 2801, rowCount);
                    newlist.AddRange(newlist1);
                    return;
                }
                ws2 = CalAMU(lastRowNum, list, StartColumn, LineIndicator, s, ws2, ref newlist1, 2801, 2900);
                newlist.AddRange(newlist1);
                newlist1.Clear();
                LogAMU(2900);
                System.Threading.Thread.Sleep(2000); GC.Collect();
                //===================================================================================3000=============================================================
                if (rowCount <= 3000)
                {
                    ws2 = CalAMU(lastRowNum, list, StartColumn, LineIndicator, s, ws2, ref newlist1, 2901, rowCount);
                    newlist.AddRange(newlist1);
                    return;
                }
                ws2 = CalAMU(lastRowNum, list, StartColumn, LineIndicator, s, ws2, ref newlist1, 2901, 3000);
                newlist.AddRange(newlist1);
                newlist1.Clear();
                LogAMU(3000);
                System.Threading.Thread.Sleep(2000); GC.Collect();
                //===================================================================================3100=============================================================
                if (rowCount <= 3100)
                {
                    ws2 = CalAMU(lastRowNum, list, StartColumn, LineIndicator, s, ws2, ref newlist1, 3001, rowCount);
                    newlist.AddRange(newlist1);
                    return;
                }
                ws2 = CalAMU(lastRowNum, list, StartColumn, LineIndicator, s, ws2, ref newlist1, 3001, 3100);
                newlist.AddRange(newlist1);
                newlist1.Clear();
                LogAMU(3100);
                System.Threading.Thread.Sleep(2000); GC.Collect();
                //===================================================================================3200=============================================================
                if (rowCount <= 3200)
                {
                    ws2 = CalAMU(lastRowNum, list, StartColumn, LineIndicator, s, ws2, ref newlist1, 3101, rowCount);
                    newlist.AddRange(newlist1);
                    return;
                }
                ws2 = CalAMU(lastRowNum, list, StartColumn, LineIndicator, s, ws2, ref newlist1, 3101, 3200);
                newlist.AddRange(newlist1);
                newlist1.Clear();
                LogAMU(3200);
                System.Threading.Thread.Sleep(2000); GC.Collect();
                //===================================================================================3300=============================================================
                if (rowCount <= 3300)
                {
                    ws2 = CalAMU(lastRowNum, list, StartColumn, LineIndicator, s, ws2, ref newlist1, 3201, rowCount);
                    newlist.AddRange(newlist1);
                    return;
                }
                ws2 = CalAMU(lastRowNum, list, StartColumn, LineIndicator, s, ws2, ref newlist1, 3201, 3300);
                newlist.AddRange(newlist1);
                newlist1.Clear();
                LogAMU(3300);
                System.Threading.Thread.Sleep(2000); GC.Collect();
                //===================================================================================3400=============================================================
                if (rowCount <= 3400)
                {
                    ws2 = CalAMU(lastRowNum, list, StartColumn, LineIndicator, s, ws2, ref newlist1, 3301, rowCount);
                    newlist.AddRange(newlist1);
                    return;
                }
                ws2 = CalAMU(lastRowNum, list, StartColumn, LineIndicator, s, ws2, ref newlist1, 3301, 3400);
                newlist.AddRange(newlist1);
                newlist1.Clear();
                LogAMU(3400);
                System.Threading.Thread.Sleep(2000); GC.Collect();
                //===================================================================================3500=============================================================
                if (rowCount <= 3500)
                {
                    ws2 = CalAMU(lastRowNum, list, StartColumn, LineIndicator, s, ws2, ref newlist1, 3401, rowCount);
                    newlist.AddRange(newlist1);
                    return;
                }
                ws2 = CalAMU(lastRowNum, list, StartColumn, LineIndicator, s, ws2, ref newlist1, 3401, 3500);
                newlist.AddRange(newlist1);
                newlist1.Clear();
                LogAMU(3500);
                System.Threading.Thread.Sleep(2000); GC.Collect();
                //===================================================================================3600=============================================================
                if (rowCount <= 3600)
                {
                    ws2 = CalAMU(lastRowNum, list, StartColumn, LineIndicator, s, ws2, ref newlist1, 3501, rowCount);
                    newlist.AddRange(newlist1);
                    return;
                }
                ws2 = CalAMU(lastRowNum, list, StartColumn, LineIndicator, s, ws2, ref newlist1, 3501, 3600);
                newlist.AddRange(newlist1);
                newlist1.Clear();
                LogAMU(3600);
                System.Threading.Thread.Sleep(2000); GC.Collect();
                //===================================================================================3700=============================================================
                if (rowCount <= 3700)
                {
                    ws2 = CalAMU(lastRowNum, list, StartColumn, LineIndicator, s, ws2, ref newlist1, 3601, rowCount);
                    newlist.AddRange(newlist1);
                    return;
                }
                ws2 = CalAMU(lastRowNum, list, StartColumn, LineIndicator, s, ws2, ref newlist1, 3601, 3700);
                newlist.AddRange(newlist1);
                newlist1.Clear();
                LogAMU(3700);
                System.Threading.Thread.Sleep(2000); GC.Collect();
                //===================================================================================3800=============================================================
                if (rowCount <= 3800)
                {
                    ws2 = CalAMU(lastRowNum, list, StartColumn, LineIndicator, s, ws2, ref newlist1, 3701, rowCount);
                    newlist.AddRange(newlist1);
                    return;
                }
                ws2 = CalAMU(lastRowNum, list, StartColumn, LineIndicator, s, ws2, ref newlist1, 3701, 3800);
                newlist.AddRange(newlist1);
                newlist1.Clear();
                LogAMU(3800);
                System.Threading.Thread.Sleep(2000); GC.Collect();
                //===================================================================================3900=============================================================
                if (rowCount <= 3900)
                {
                    ws2 = CalAMU(lastRowNum, list, StartColumn, LineIndicator, s, ws2, ref newlist1, 3801, rowCount);
                    newlist.AddRange(newlist1);
                    return;
                }
                ws2 = CalAMU(lastRowNum, list, StartColumn, LineIndicator, s, ws2, ref newlist1, 3801, 3900);
                newlist.AddRange(newlist1);
                newlist1.Clear();
                LogAMU(3900);
                System.Threading.Thread.Sleep(2000); GC.Collect();
                //===================================================================================4000=============================================================
                if (rowCount <= 4000)
                {
                    ws2 = CalAMU(lastRowNum, list, StartColumn, LineIndicator, s, ws2, ref newlist1, 3901, rowCount);
                    newlist.AddRange(newlist1);
                    return;
                }
                ws2 = CalAMU(lastRowNum, list, StartColumn, LineIndicator, s, ws2, ref newlist1, 3901, 4000);
                newlist.AddRange(newlist1);
                newlist1.Clear();
                LogAMU(4000);
                System.Threading.Thread.Sleep(2000); GC.Collect();
                //Copy sheet to make ws not disappear.
                Globals.ThisAddIn.Application.DisplayAlerts = false;
                Globals.ThisAddIn.Application.Worksheets.Add(Type.Missing, Globals.ThisAddIn.Application.ActiveWorkbook.Worksheets[/*"Sheet1"*/Globals.ThisAddIn.Application.ActiveWorkbook.Worksheets.Count], 1, Type.Missing);
                Microsoft.Office.Interop.Excel.Worksheet ws3 = (Microsoft.Office.Interop.Excel.Worksheet)Globals.ThisAddIn.Application.Worksheets.get_Item(Globals.ThisAddIn.Application.ActiveWorkbook.Worksheets.Count);
                ws3.Name = "RSTMPTemplate" + Guid.NewGuid().ToString().Replace("-", "").Remove(17);
                ws2.get_Range("A1", Ribbon2.LastColumnName + Ribbon2.LastRowNumber).Copy(ws3.get_Range("A1"));
                Globals.ThisAddIn.Application.DisplayAlerts = true;
                System.Windows.Forms.Clipboard.SetText("\r\n");
                //===================================================================================4100=============================================================
                if (rowCount <= 4100)
                {
                    ws3 = CalAMU(lastRowNum, list, StartColumn, LineIndicator, s, ws3, ref newlist1, 4001, rowCount);
                    newlist.AddRange(newlist1);
                    return;
                }
                ws3 = CalAMU(lastRowNum, list, StartColumn, LineIndicator, s, ws3, ref newlist1, 4001, 4100);
                newlist.AddRange(newlist1);
                newlist1.Clear();
                LogAMU(4100);
                System.Threading.Thread.Sleep(2000); GC.Collect();
                //===================================================================================4200=============================================================
                if (rowCount <= 4200)
                {
                    ws3 = CalAMU(lastRowNum, list, StartColumn, LineIndicator, s, ws3, ref newlist1, 4101, rowCount);
                    newlist.AddRange(newlist1);
                    return;
                }
                ws3 = CalAMU(lastRowNum, list, StartColumn, LineIndicator, s, ws3, ref newlist1, 4101, 4200);
                newlist.AddRange(newlist1);
                newlist1.Clear();
                LogAMU(4200);
                System.Threading.Thread.Sleep(2000); GC.Collect();
                //===================================================================================4300=============================================================
                if (rowCount <= 4300)
                {
                    ws3 = CalAMU(lastRowNum, list, StartColumn, LineIndicator, s, ws3, ref newlist1, 4201, rowCount);
                    newlist.AddRange(newlist1);
                    return;
                }
                ws3 = CalAMU(lastRowNum, list, StartColumn, LineIndicator, s, ws3, ref newlist1, 4201, 4300);
                newlist.AddRange(newlist1);
                newlist1.Clear();
                LogAMU(4300);
                System.Threading.Thread.Sleep(2000); GC.Collect();
                //===================================================================================4400=============================================================
                if (rowCount <= 4400)
                {
                    ws3 = CalAMU(lastRowNum, list, StartColumn, LineIndicator, s, ws3, ref newlist1, 4301, rowCount);
                    newlist.AddRange(newlist1);
                    return;
                }
                ws3 = CalAMU(lastRowNum, list, StartColumn, LineIndicator, s, ws3, ref newlist1, 4301, 4400);
                newlist.AddRange(newlist1);
                newlist1.Clear();
                LogAMU(4400);
                System.Threading.Thread.Sleep(2000); GC.Collect();
                //===================================================================================4500=============================================================
                if (rowCount <= 4500)
                {
                    ws3 = CalAMU(lastRowNum, list, StartColumn, LineIndicator, s, ws3, ref newlist1, 4401, rowCount);
                    newlist.AddRange(newlist1);
                    return;
                }
                ws3 = CalAMU(lastRowNum, list, StartColumn, LineIndicator, s, ws3, ref newlist1, 4401, 4500);
                newlist.AddRange(newlist1);
                newlist1.Clear();
                LogAMU(4500);
                System.Threading.Thread.Sleep(2000); GC.Collect();
                //===================================================================================4600=============================================================
                if (rowCount <= 4600)
                {
                    ws3 = CalAMU(lastRowNum, list, StartColumn, LineIndicator, s, ws3, ref newlist1, 4501, rowCount);
                    newlist.AddRange(newlist1);
                    return;
                }
                ws3 = CalAMU(lastRowNum, list, StartColumn, LineIndicator, s, ws3, ref newlist1, 4501, 4600);
                newlist.AddRange(newlist1);
                newlist1.Clear();
                LogAMU(4600);
                System.Threading.Thread.Sleep(2000); GC.Collect();
                //===================================================================================4700=============================================================
                if (rowCount <= 4700)
                {
                    ws3 = CalAMU(lastRowNum, list, StartColumn, LineIndicator, s, ws3, ref newlist1, 4601, rowCount);
                    newlist.AddRange(newlist1);
                    return;
                }
                ws3 = CalAMU(lastRowNum, list, StartColumn, LineIndicator, s, ws3, ref newlist1, 4601, 4700);
                newlist.AddRange(newlist1);
                newlist1.Clear();
                LogAMU(4700);
                System.Threading.Thread.Sleep(2000); GC.Collect();
                //===================================================================================4800=============================================================
                if (rowCount <= 4800)
                {
                    ws3 = CalAMU(lastRowNum, list, StartColumn, LineIndicator, s, ws3, ref newlist1, 4701, rowCount);
                    newlist.AddRange(newlist1);
                    return;
                }
                ws3 = CalAMU(lastRowNum, list, StartColumn, LineIndicator, s, ws3, ref newlist1, 4701, 4800);
                newlist.AddRange(newlist1);
                newlist1.Clear();
                LogAMU(4800);
                System.Threading.Thread.Sleep(2000); GC.Collect();
                //===================================================================================4900=============================================================
                if (rowCount <= 4900)
                {
                    ws3 = CalAMU(lastRowNum, list, StartColumn, LineIndicator, s, ws3, ref newlist1, 4801, rowCount);
                    newlist.AddRange(newlist1);
                    return;
                }
                ws3 = CalAMU(lastRowNum, list, StartColumn, LineIndicator, s, ws3, ref newlist1, 4801, 4900);
                newlist.AddRange(newlist1);
                newlist1.Clear();
                LogAMU(4900);
                System.Threading.Thread.Sleep(2000); GC.Collect();
                //===================================================================================5000=============================================================
                if (rowCount <= 5000)
                {
                    ws3 = CalAMU(lastRowNum, list, StartColumn, LineIndicator, s, ws3, ref newlist1, 4901, rowCount);
                    newlist.AddRange(newlist1);
                    return;
                }
                ws3 = CalAMU(lastRowNum, list, StartColumn, LineIndicator, s, ws3, ref newlist1, 4901, 5000);
                newlist.AddRange(newlist1);
                newlist1.Clear();
                LogAMU(5000);
                System.Threading.Thread.Sleep(2000); GC.Collect();
                //===================================================================================5100=============================================================
                if (rowCount <= 5100)
                {
                    ws3 = CalAMU(lastRowNum, list, StartColumn, LineIndicator, s, ws3, ref newlist1, 5001, rowCount);
                    newlist.AddRange(newlist1);
                    return;
                }
                ws3 = CalAMU(lastRowNum, list, StartColumn, LineIndicator, s, ws3, ref newlist1, 5001, 5100);
                newlist.AddRange(newlist1);
                newlist1.Clear();
                LogAMU(5100);
                System.Threading.Thread.Sleep(2000); GC.Collect();
                //===================================================================================5200=============================================================
                if (rowCount <= 5200)
                {
                    ws3 = CalAMU(lastRowNum, list, StartColumn, LineIndicator, s, ws3, ref newlist1, 5101, rowCount);
                    newlist.AddRange(newlist1);
                    return;
                }
                ws3 = CalAMU(lastRowNum, list, StartColumn, LineIndicator, s, ws3, ref newlist1, 5101, 5200);
                newlist.AddRange(newlist1);
                newlist1.Clear();
                LogAMU(5200);
                System.Threading.Thread.Sleep(2000); GC.Collect();
                //===================================================================================5300=============================================================
                if (rowCount <= 5300)
                {
                    ws3 = CalAMU(lastRowNum, list, StartColumn, LineIndicator, s, ws3, ref newlist1, 5201, rowCount);
                    newlist.AddRange(newlist1);
                    return;
                }
                ws3 = CalAMU(lastRowNum, list, StartColumn, LineIndicator, s, ws3, ref newlist1, 5201, 5300);
                newlist.AddRange(newlist1);
                newlist1.Clear();
                LogAMU(5300);
                System.Threading.Thread.Sleep(2000); GC.Collect();
                //===================================================================================5400=============================================================
                if (rowCount <= 5400)
                {
                    ws3 = CalAMU(lastRowNum, list, StartColumn, LineIndicator, s, ws3, ref newlist1, 5301, rowCount);
                    newlist.AddRange(newlist1);
                    return;
                }
                ws3 = CalAMU(lastRowNum, list, StartColumn, LineIndicator, s, ws3, ref newlist1, 5301, 5400);
                newlist.AddRange(newlist1);
                newlist1.Clear();
                LogAMU(5400);
                System.Threading.Thread.Sleep(2000); GC.Collect();
                //===================================================================================5500=============================================================
                if (rowCount <= 5500)
                {
                    ws3 = CalAMU(lastRowNum, list, StartColumn, LineIndicator, s, ws3, ref newlist1, 5401, rowCount);
                    newlist.AddRange(newlist1);
                    return;
                }
                ws3 = CalAMU(lastRowNum, list, StartColumn, LineIndicator, s, ws3, ref newlist1, 5401, 5500);
                newlist.AddRange(newlist1);
                newlist1.Clear();
                LogAMU(5500);
                System.Threading.Thread.Sleep(2000); GC.Collect();
                //===================================================================================5600=============================================================
                if (rowCount <= 5600)
                {
                    ws3 = CalAMU(lastRowNum, list, StartColumn, LineIndicator, s, ws3, ref newlist1, 5501, rowCount);
                    newlist.AddRange(newlist1);
                    return;
                }
                ws3 = CalAMU(lastRowNum, list, StartColumn, LineIndicator, s, ws3, ref newlist1, 5501, 5600);
                newlist.AddRange(newlist1);
                newlist1.Clear();
                LogAMU(5600);
                System.Threading.Thread.Sleep(2000); GC.Collect();
                //===================================================================================5700=============================================================
                if (rowCount <= 5700)
                {
                    ws3 = CalAMU(lastRowNum, list, StartColumn, LineIndicator, s, ws3, ref newlist1, 5601, rowCount);
                    newlist.AddRange(newlist1);
                    return;
                }
                ws3 = CalAMU(lastRowNum, list, StartColumn, LineIndicator, s, ws3, ref newlist1, 5601, 5700);
                newlist.AddRange(newlist1);
                newlist1.Clear();
                LogAMU(5700);
                System.Threading.Thread.Sleep(2000); GC.Collect();
                //===================================================================================5800=============================================================
                if (rowCount <= 5800)
                {
                    ws3 = CalAMU(lastRowNum, list, StartColumn, LineIndicator, s, ws3, ref newlist1, 5701, rowCount);
                    newlist.AddRange(newlist1);
                    return;
                }
                ws3 = CalAMU(lastRowNum, list, StartColumn, LineIndicator, s, ws3, ref newlist1, 5701, 5800);
                newlist.AddRange(newlist1);
                newlist1.Clear();
                LogAMU(5800);
                System.Threading.Thread.Sleep(2000); GC.Collect();
                //===================================================================================5900=============================================================
                if (rowCount <= 5900)
                {
                    ws3 = CalAMU(lastRowNum, list, StartColumn, LineIndicator, s, ws3, ref newlist1, 5801, rowCount);
                    newlist.AddRange(newlist1);
                    return;
                }
                ws3 = CalAMU(lastRowNum, list, StartColumn, LineIndicator, s, ws3, ref newlist1, 5801, 5900);
                newlist.AddRange(newlist1);
                newlist1.Clear();
                LogAMU(5900);
                System.Threading.Thread.Sleep(2000); GC.Collect();
                //===================================================================================6000=============================================================
                if (rowCount <= 6000)
                {
                    ws3 = CalAMU(lastRowNum, list, StartColumn, LineIndicator, s, ws3, ref newlist1, 5901, rowCount);
                    newlist.AddRange(newlist1);
                    return;
                }
                ws3 = CalAMU(lastRowNum, list, StartColumn, LineIndicator, s, ws3, ref newlist1, 5901, 6000);
                newlist.AddRange(newlist1);
                newlist1.Clear();
                LogAMU(6000);
                System.Threading.Thread.Sleep(2000); GC.Collect();
                //Copy sheet to make ws not disappear.
                Globals.ThisAddIn.Application.DisplayAlerts = false;
                Globals.ThisAddIn.Application.Worksheets.Add(Type.Missing, Globals.ThisAddIn.Application.ActiveWorkbook.Worksheets[/*"Sheet1"*/Globals.ThisAddIn.Application.ActiveWorkbook.Worksheets.Count], 1, Type.Missing);
                Microsoft.Office.Interop.Excel.Worksheet ws4 = (Microsoft.Office.Interop.Excel.Worksheet)Globals.ThisAddIn.Application.Worksheets.get_Item(Globals.ThisAddIn.Application.ActiveWorkbook.Worksheets.Count);
                ws4.Name = "RSTMPTemplate" + Guid.NewGuid().ToString().Replace("-", "").Remove(17);
                ws3.get_Range("A1", Ribbon2.LastColumnName + Ribbon2.LastRowNumber).Copy(ws4.get_Range("A1"));
                Globals.ThisAddIn.Application.DisplayAlerts = true;
                System.Windows.Forms.Clipboard.SetText("\r\n");
                //===================================================================================6100=============================================================
                if (rowCount <= 6100)
                {
                    ws4 = CalAMU(lastRowNum, list, StartColumn, LineIndicator, s, ws4, ref newlist1, 6001, rowCount);
                    newlist.AddRange(newlist1);
                    return;
                }
                ws4 = CalAMU(lastRowNum, list, StartColumn, LineIndicator, s, ws4, ref newlist1, 6001, 6100);
                newlist.AddRange(newlist1);
                newlist1.Clear();
                LogAMU(6100);
                System.Threading.Thread.Sleep(2000); GC.Collect();
                //===================================================================================6200=============================================================
                if (rowCount <= 6200)
                {
                    ws4 = CalAMU(lastRowNum, list, StartColumn, LineIndicator, s, ws4, ref newlist1, 6101, rowCount);
                    newlist.AddRange(newlist1);
                    return;
                }
                ws4 = CalAMU(lastRowNum, list, StartColumn, LineIndicator, s, ws4, ref newlist1, 6101, 6200);
                newlist.AddRange(newlist1);
                newlist1.Clear();
                LogAMU(6200);
                System.Threading.Thread.Sleep(2000); GC.Collect();
                //===================================================================================6300=============================================================
                if (rowCount <= 6300)
                {
                    ws4 = CalAMU(lastRowNum, list, StartColumn, LineIndicator, s, ws4, ref newlist1, 6201, rowCount);
                    newlist.AddRange(newlist1);
                    return;
                }
                ws4 = CalAMU(lastRowNum, list, StartColumn, LineIndicator, s, ws4, ref newlist1, 6201, 6300);
                newlist.AddRange(newlist1);
                newlist1.Clear();
                LogAMU(6300);
                System.Threading.Thread.Sleep(2000); GC.Collect();
                //===================================================================================6400=============================================================
                if (rowCount <= 6400)
                {
                    ws4 = CalAMU(lastRowNum, list, StartColumn, LineIndicator, s, ws4, ref newlist1, 6301, rowCount);
                    newlist.AddRange(newlist1);
                    return;
                }
                ws4 = CalAMU(lastRowNum, list, StartColumn, LineIndicator, s, ws4, ref newlist1, 6301, 6400);
                newlist.AddRange(newlist1);
                newlist1.Clear();
                LogAMU(6400);
                System.Threading.Thread.Sleep(2000); GC.Collect();
                //===================================================================================6500=============================================================
                if (rowCount <= 6500)
                {
                    ws4 = CalAMU(lastRowNum, list, StartColumn, LineIndicator, s, ws4, ref newlist1, 6401, rowCount);
                    newlist.AddRange(newlist1);
                    return;
                }
                ws4 = CalAMU(lastRowNum, list, StartColumn, LineIndicator, s, ws4, ref newlist1, 6401, 6500);
                newlist.AddRange(newlist1);
                newlist1.Clear();
                LogAMU(6500);
                System.Threading.Thread.Sleep(2000); GC.Collect();
                //===================================================================================6600=============================================================
                if (rowCount <= 6600)
                {
                    ws4 = CalAMU(lastRowNum, list, StartColumn, LineIndicator, s, ws4, ref newlist1, 6501, rowCount);
                    newlist.AddRange(newlist1);
                    return;
                }
                ws4 = CalAMU(lastRowNum, list, StartColumn, LineIndicator, s, ws4, ref newlist1, 6501, 6600);
                newlist.AddRange(newlist1);
                newlist1.Clear();
                LogAMU(6600);
                System.Threading.Thread.Sleep(2000); GC.Collect();
                //===================================================================================6700=============================================================
                if (rowCount <= 6700)
                {
                    ws4 = CalAMU(lastRowNum, list, StartColumn, LineIndicator, s, ws4, ref newlist1, 6601, rowCount);
                    newlist.AddRange(newlist1);
                    return;
                }
                ws4 = CalAMU(lastRowNum, list, StartColumn, LineIndicator, s, ws4, ref newlist1, 6601, 6700);
                newlist.AddRange(newlist1);
                newlist1.Clear();
                LogAMU(6700);
                System.Threading.Thread.Sleep(2000); GC.Collect();
                //===================================================================================6800=============================================================
                if (rowCount <= 6800)
                {
                    ws4 = CalAMU(lastRowNum, list, StartColumn, LineIndicator, s, ws4, ref newlist1, 6701, rowCount);
                    newlist.AddRange(newlist1);
                    return;
                }
                ws4 = CalAMU(lastRowNum, list, StartColumn, LineIndicator, s, ws4, ref newlist1, 6701, 6800);
                newlist.AddRange(newlist1);
                newlist1.Clear();
                LogAMU(6800);
                System.Threading.Thread.Sleep(2000); GC.Collect();
                //===================================================================================6900=============================================================
                if (rowCount <= 6900)
                {
                    ws4 = CalAMU(lastRowNum, list, StartColumn, LineIndicator, s, ws4, ref newlist1, 6801, rowCount);
                    newlist.AddRange(newlist1);
                    return;
                }
                ws4 = CalAMU(lastRowNum, list, StartColumn, LineIndicator, s, ws4, ref newlist1, 6801, 6900);
                newlist.AddRange(newlist1);
                newlist1.Clear();
                LogAMU(6900);
                System.Threading.Thread.Sleep(2000); GC.Collect();
                //===================================================================================7000=============================================================
                if (rowCount <= 7000)
                {
                    ws4 = CalAMU(lastRowNum, list, StartColumn, LineIndicator, s, ws4, ref newlist1, 6901, rowCount);
                    newlist.AddRange(newlist1);
                    return;
                }
                ws4 = CalAMU(lastRowNum, list, StartColumn, LineIndicator, s, ws4, ref newlist1, 6901, 7000);
                newlist.AddRange(newlist1);
                newlist1.Clear();
                LogAMU(7000);
                System.Threading.Thread.Sleep(2000); GC.Collect();
                //===================================================================================7100=============================================================
                if (rowCount <= 7100)
                {
                    ws4 = CalAMU(lastRowNum, list, StartColumn, LineIndicator, s, ws4, ref newlist1, 7001, rowCount);
                    newlist.AddRange(newlist1);
                    return;
                }
                ws4 = CalAMU(lastRowNum, list, StartColumn, LineIndicator, s, ws4, ref newlist1, 7001, 7100);
                newlist.AddRange(newlist1);
                newlist1.Clear();
                LogAMU(7100);
                System.Threading.Thread.Sleep(2000); GC.Collect();
                //===================================================================================7200=============================================================
                if (rowCount <= 7200)
                {
                    ws4 = CalAMU(lastRowNum, list, StartColumn, LineIndicator, s, ws4, ref newlist1, 7101, rowCount);
                    newlist.AddRange(newlist1);
                    return;
                }
                ws4 = CalAMU(lastRowNum, list, StartColumn, LineIndicator, s, ws4, ref newlist1, 7101, 7200);
                newlist.AddRange(newlist1);
                newlist1.Clear();
                LogAMU(7200);
                System.Threading.Thread.Sleep(2000); GC.Collect();
                //===================================================================================7300=============================================================
                if (rowCount <= 7300)
                {
                    ws4 = CalAMU(lastRowNum, list, StartColumn, LineIndicator, s, ws4, ref newlist1, 7201, rowCount);
                    newlist.AddRange(newlist1);
                    return;
                }
                ws4 = CalAMU(lastRowNum, list, StartColumn, LineIndicator, s, ws4, ref newlist1, 7201, 7300);
                newlist.AddRange(newlist1);
                newlist1.Clear();
                LogAMU(7300);
                System.Threading.Thread.Sleep(2000); GC.Collect();
                //===================================================================================7400=============================================================
                if (rowCount <= 7400)
                {
                    ws4 = CalAMU(lastRowNum, list, StartColumn, LineIndicator, s, ws4, ref newlist1, 7301, rowCount);
                    newlist.AddRange(newlist1);
                    return;
                }
                ws4 = CalAMU(lastRowNum, list, StartColumn, LineIndicator, s, ws4, ref newlist1, 7301, 7400);
                newlist.AddRange(newlist1);
                newlist1.Clear();
                LogAMU(7400);
                System.Threading.Thread.Sleep(2000); GC.Collect();
                //===================================================================================7500=============================================================
                if (rowCount <= 7500)
                {
                    ws4 = CalAMU(lastRowNum, list, StartColumn, LineIndicator, s, ws4, ref newlist1, 7401, rowCount);
                    newlist.AddRange(newlist1);
                    return;
                }
                ws4 = CalAMU(lastRowNum, list, StartColumn, LineIndicator, s, ws4, ref newlist1, 7401, 7500);
                newlist.AddRange(newlist1);
                newlist1.Clear();
                LogAMU(7500);
                System.Threading.Thread.Sleep(2000); GC.Collect();
                //===================================================================================7600=============================================================
                if (rowCount <= 7600)
                {
                    ws4 = CalAMU(lastRowNum, list, StartColumn, LineIndicator, s, ws4, ref newlist1, 7501, rowCount);
                    newlist.AddRange(newlist1);
                    return;
                }
                ws4 = CalAMU(lastRowNum, list, StartColumn, LineIndicator, s, ws4, ref newlist1, 7501, 7600);
                newlist.AddRange(newlist1);
                newlist1.Clear();
                LogAMU(7600);
                System.Threading.Thread.Sleep(2000); GC.Collect();
                //===================================================================================7700=============================================================
                if (rowCount <= 7700)
                {
                    ws4 = CalAMU(lastRowNum, list, StartColumn, LineIndicator, s, ws4, ref newlist1, 7601, rowCount);
                    newlist.AddRange(newlist1);
                    return;
                }
                ws4 = CalAMU(lastRowNum, list, StartColumn, LineIndicator, s, ws4, ref newlist1, 7601, 7700);
                newlist.AddRange(newlist1);
                newlist1.Clear();
                LogAMU(7700);
                System.Threading.Thread.Sleep(2000); GC.Collect();
                //===================================================================================7800=============================================================
                if (rowCount <= 7800)
                {
                    ws4 = CalAMU(lastRowNum, list, StartColumn, LineIndicator, s, ws4, ref newlist1, 7701, rowCount);
                    newlist.AddRange(newlist1);
                    return;
                }
                ws4 = CalAMU(lastRowNum, list, StartColumn, LineIndicator, s, ws4, ref newlist1, 7701, 7800);
                newlist.AddRange(newlist1);
                newlist1.Clear();
                LogAMU(7800);
                System.Threading.Thread.Sleep(2000); GC.Collect();
                //===================================================================================7900=============================================================
                if (rowCount <= 7900)
                {
                    ws4 = CalAMU(lastRowNum, list, StartColumn, LineIndicator, s, ws4, ref newlist1, 7801, rowCount);
                    newlist.AddRange(newlist1);
                    return;
                }
                ws4 = CalAMU(lastRowNum, list, StartColumn, LineIndicator, s, ws4, ref newlist1, 7801, 7900);
                newlist.AddRange(newlist1);
                newlist1.Clear();
                LogAMU(7900);
                System.Threading.Thread.Sleep(2000); GC.Collect();
                //===================================================================================8000=============================================================
                if (rowCount <= 8000)
                {
                    ws4 = CalAMU(lastRowNum, list, StartColumn, LineIndicator, s, ws4, ref newlist1, 7901, rowCount);
                    newlist.AddRange(newlist1);
                    return;
                }
                ws4 = CalAMU(lastRowNum, list, StartColumn, LineIndicator, s, ws4, ref newlist1, 7901, 8000);
                newlist.AddRange(newlist1);
                newlist1.Clear();
                LogAMU(8000);
                System.Threading.Thread.Sleep(2000); GC.Collect();
                //===================================================================================8100=============================================================
                if (rowCount <= 8100)
                {
                    ws4 = CalAMU(lastRowNum, list, StartColumn, LineIndicator, s, ws4, ref newlist1, 8001, rowCount);
                    newlist.AddRange(newlist1);
                    return;
                }
                ws4 = CalAMU(lastRowNum, list, StartColumn, LineIndicator, s, ws4, ref newlist1, 8001, 8100);
                newlist.AddRange(newlist1);
                newlist1.Clear();
                LogAMU(8100);
                System.Threading.Thread.Sleep(2000); GC.Collect();
                //===================================================================================8200=============================================================
                if (rowCount <= 8200)
                {
                    ws4 = CalAMU(lastRowNum, list, StartColumn, LineIndicator, s, ws4, ref newlist1, 8101, rowCount);
                    newlist.AddRange(newlist1);
                    return;
                }
                ws4 = CalAMU(lastRowNum, list, StartColumn, LineIndicator, s, ws4, ref newlist1, 8101, 8200);
                newlist.AddRange(newlist1);
                newlist1.Clear();
                LogAMU(8200);
                System.Threading.Thread.Sleep(2000); GC.Collect();
                //===================================================================================8300=============================================================
                if (rowCount <= 8300)
                {
                    ws4 = CalAMU(lastRowNum, list, StartColumn, LineIndicator, s, ws4, ref newlist1, 8201, rowCount);
                    newlist.AddRange(newlist1);
                    return;
                }
                ws4 = CalAMU(lastRowNum, list, StartColumn, LineIndicator, s, ws4, ref newlist1, 8201, 8300);
                newlist.AddRange(newlist1);
                newlist1.Clear();
                LogAMU(8300);
                System.Threading.Thread.Sleep(2000); GC.Collect();
                //===================================================================================8400=============================================================
                if (rowCount <= 8400)
                {
                    ws4 = CalAMU(lastRowNum, list, StartColumn, LineIndicator, s, ws4, ref newlist1, 8301, rowCount);
                    newlist.AddRange(newlist1);
                    return;
                }
                ws4 = CalAMU(lastRowNum, list, StartColumn, LineIndicator, s, ws4, ref newlist1, 8301, 8400);
                newlist.AddRange(newlist1);
                newlist1.Clear();
                LogAMU(8400);
                System.Threading.Thread.Sleep(2000); GC.Collect();
                //===================================================================================8500=============================================================
                if (rowCount <= 8500)
                {
                    ws4 = CalAMU(lastRowNum, list, StartColumn, LineIndicator, s, ws4, ref newlist1, 8401, rowCount);
                    newlist.AddRange(newlist1);
                    return;
                }
                ws4 = CalAMU(lastRowNum, list, StartColumn, LineIndicator, s, ws4, ref newlist1, 8401, 8500);
                newlist.AddRange(newlist1);
                newlist1.Clear();
                LogAMU(8500);
                System.Threading.Thread.Sleep(2000); GC.Collect();
                //===================================================================================8600=============================================================
                if (rowCount <= 8600)
                {
                    ws4 = CalAMU(lastRowNum, list, StartColumn, LineIndicator, s, ws4, ref newlist1, 8501, rowCount);
                    newlist.AddRange(newlist1);
                    return;
                }
                ws4 = CalAMU(lastRowNum, list, StartColumn, LineIndicator, s, ws4, ref newlist1, 8501, 8600);
                newlist.AddRange(newlist1);
                newlist1.Clear();
                LogAMU(8600);
                System.Threading.Thread.Sleep(2000); GC.Collect();
                //===================================================================================8700=============================================================
                if (rowCount <= 8700)
                {
                    ws4 = CalAMU(lastRowNum, list, StartColumn, LineIndicator, s, ws4, ref newlist1, 8601, rowCount);
                    newlist.AddRange(newlist1);
                    return;
                }
                ws4 = CalAMU(lastRowNum, list, StartColumn, LineIndicator, s, ws4, ref newlist1, 8601, 8700);
                newlist.AddRange(newlist1);
                newlist1.Clear();
                LogAMU(8700);
                System.Threading.Thread.Sleep(2000); GC.Collect();
                //===================================================================================8800=============================================================
                if (rowCount <= 8800)
                {
                    ws4 = CalAMU(lastRowNum, list, StartColumn, LineIndicator, s, ws4, ref newlist1, 8701, rowCount);
                    newlist.AddRange(newlist1);
                    return;
                }
                ws4 = CalAMU(lastRowNum, list, StartColumn, LineIndicator, s, ws4, ref newlist1, 8701, 8800);
                newlist.AddRange(newlist1);
                newlist1.Clear();
                LogAMU(8800);
                System.Threading.Thread.Sleep(2000); GC.Collect();
                //===================================================================================8900=============================================================
                if (rowCount <= 8900)
                {
                    ws4 = CalAMU(lastRowNum, list, StartColumn, LineIndicator, s, ws4, ref newlist1, 8801, rowCount);
                    newlist.AddRange(newlist1);
                    return;
                }
                ws4 = CalAMU(lastRowNum, list, StartColumn, LineIndicator, s, ws4, ref newlist1, 8801, 8900);
                newlist.AddRange(newlist1);
                newlist1.Clear();
                LogAMU(8900);
                System.Threading.Thread.Sleep(2000); GC.Collect();
                //===================================================================================9000=============================================================
                if (rowCount <= 9000)
                {
                    ws4 = CalAMU(lastRowNum, list, StartColumn, LineIndicator, s, ws4, ref newlist1, 8901, rowCount);
                    newlist.AddRange(newlist1);
                    return;
                }
                ws4 = CalAMU(lastRowNum, list, StartColumn, LineIndicator, s, ws4, ref newlist1, 8901, 9000);
                newlist.AddRange(newlist1);
                newlist1.Clear();
                LogAMU(9000);
                System.Threading.Thread.Sleep(2000); GC.Collect();
                //===================================================================================9100=============================================================
                if (rowCount <= 9100)
                {
                    ws4 = CalAMU(lastRowNum, list, StartColumn, LineIndicator, s, ws4, ref newlist1, 9001, rowCount);
                    newlist.AddRange(newlist1);
                    return;
                }
                ws4 = CalAMU(lastRowNum, list, StartColumn, LineIndicator, s, ws4, ref newlist1, 9001, 9100);
                newlist.AddRange(newlist1);
                newlist1.Clear();
                LogAMU(9100);
                System.Threading.Thread.Sleep(2000); GC.Collect();
                //===================================================================================9200=============================================================
                if (rowCount <= 9200)
                {
                    ws4 = CalAMU(lastRowNum, list, StartColumn, LineIndicator, s, ws4, ref newlist1, 9101, rowCount);
                    newlist.AddRange(newlist1);
                    return;
                }
                ws4 = CalAMU(lastRowNum, list, StartColumn, LineIndicator, s, ws4, ref newlist1, 9101, 9200);
                newlist.AddRange(newlist1);
                newlist1.Clear();
                LogAMU(9200);
                System.Threading.Thread.Sleep(2000); GC.Collect();
                //===================================================================================9300=============================================================
                if (rowCount <= 9300)
                {
                    ws4 = CalAMU(lastRowNum, list, StartColumn, LineIndicator, s, ws4, ref newlist1, 9201, rowCount);
                    newlist.AddRange(newlist1);
                    return;
                }
                ws4 = CalAMU(lastRowNum, list, StartColumn, LineIndicator, s, ws4, ref newlist1, 9201, 9300);
                newlist.AddRange(newlist1);
                newlist1.Clear();
                LogAMU(9300);
                System.Threading.Thread.Sleep(2000); GC.Collect();
                //===================================================================================9400=============================================================
                if (rowCount <= 9400)
                {
                    ws4 = CalAMU(lastRowNum, list, StartColumn, LineIndicator, s, ws4, ref newlist1, 9301, rowCount);
                    newlist.AddRange(newlist1);
                    return;
                }
                ws4 = CalAMU(lastRowNum, list, StartColumn, LineIndicator, s, ws4, ref newlist1, 9301, 9400);
                newlist.AddRange(newlist1);
                newlist1.Clear();
                LogAMU(9400);
                System.Threading.Thread.Sleep(2000); GC.Collect();
                //===================================================================================9500=============================================================
                if (rowCount <= 9500)
                {
                    ws4 = CalAMU(lastRowNum, list, StartColumn, LineIndicator, s, ws4, ref newlist1, 9401, rowCount);
                    newlist.AddRange(newlist1);
                    return;
                }
                ws4 = CalAMU(lastRowNum, list, StartColumn, LineIndicator, s, ws4, ref newlist1, 9401, 9500);
                newlist.AddRange(newlist1);
                newlist1.Clear();
                LogAMU(9500);
                System.Threading.Thread.Sleep(2000); GC.Collect();
                //===================================================================================9600=============================================================
                if (rowCount <= 9600)
                {
                    ws4 = CalAMU(lastRowNum, list, StartColumn, LineIndicator, s, ws4, ref newlist1, 9501, rowCount);
                    newlist.AddRange(newlist1);
                    return;
                }
                ws4 = CalAMU(lastRowNum, list, StartColumn, LineIndicator, s, ws4, ref newlist1, 9501, 9600);
                newlist.AddRange(newlist1);
                newlist1.Clear();
                LogAMU(9600);
                System.Threading.Thread.Sleep(2000); GC.Collect();
                //===================================================================================9700=============================================================
                if (rowCount <= 9700)
                {
                    ws4 = CalAMU(lastRowNum, list, StartColumn, LineIndicator, s, ws4, ref newlist1, 9601, rowCount);
                    newlist.AddRange(newlist1);
                    return;
                }
                ws4 = CalAMU(lastRowNum, list, StartColumn, LineIndicator, s, ws4, ref newlist1, 9601, 9700);
                newlist.AddRange(newlist1);
                newlist1.Clear();
                LogAMU(9700);
                System.Threading.Thread.Sleep(2000); GC.Collect();
                //===================================================================================9800=============================================================
                if (rowCount <= 9800)
                {
                    ws4 = CalAMU(lastRowNum, list, StartColumn, LineIndicator, s, ws4, ref newlist1, 9701, rowCount);
                    newlist.AddRange(newlist1);
                    return;
                }
                ws4 = CalAMU(lastRowNum, list, StartColumn, LineIndicator, s, ws4, ref newlist1, 9701, 9800);
                newlist.AddRange(newlist1);
                newlist1.Clear();
                LogAMU(9800);
                System.Threading.Thread.Sleep(2000); GC.Collect();
                //===================================================================================9900=============================================================
                if (rowCount <= 9900)
                {
                    ws4 = CalAMU(lastRowNum, list, StartColumn, LineIndicator, s, ws4, ref newlist1, 9801, rowCount);
                    newlist.AddRange(newlist1);
                    return;
                }
                ws4 = CalAMU(lastRowNum, list, StartColumn, LineIndicator, s, ws4, ref newlist1, 9801, 9900);
                newlist.AddRange(newlist1);
                newlist1.Clear();
                LogAMU(9900);
                System.Threading.Thread.Sleep(2000); GC.Collect();
                //===================================================================================10000=============================================================
                if (rowCount <= 10000)
                {
                    ws4 = CalAMU(lastRowNum, list, StartColumn, LineIndicator, s, ws4, ref newlist1, 9901, rowCount);
                    newlist.AddRange(newlist1);
                    return;
                }
                ws4 = CalAMU(lastRowNum, list, StartColumn, LineIndicator, s, ws4, ref newlist1, 9901, 10000);
                newlist.AddRange(newlist1);
                LogAMU(10000);
            }
        }
    }
}
