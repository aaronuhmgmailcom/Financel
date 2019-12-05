using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.ComponentModel.Design;
using ciloci.FormulaEngine;

namespace ExcelAddIn4
{
    public partial class Finance_Tools
    {
        public IServiceContainer MyServices;
        public FormulaEngine engine;
        public ciloci.FormulaEngine.Formula CreateFormula(String expression)
        {
            try
            {
                FormulaEngine fe = (FormulaEngine)MyServices.GetService(typeof(FormulaEngine));
                return fe.CreateFormula(expression);
            }
            catch (ciloci.FormulaEngine.InvalidFormulaException ex)
            {
                return null;
            }
        }
        private void SetFormulaEngine(FormulaEngine engine)
        {
            FormulaEngine oldEngine = (FormulaEngine)MyServices.GetService(typeof(FormulaEngine));
            MyServices.RemoveService(typeof(FormulaEngine));
            MyServices.AddService(typeof(FormulaEngine), engine);
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
        /// <param name="start"></param>
        /// <param name="over"></param>
        private Microsoft.Office.Interop.Excel.Worksheet Cal(string lastRowNum, List<Specialist> list, string StartColumn, string LineIndicator, string s, Microsoft.Office.Interop.Excel.Worksheet ws, ref List<Specialist> newlist, int start, int over)
        {
            for (int j = start; j <= over; j++)
            {
                System.Windows.Forms.Application.DoEvents();
                for (int i = 0; i < list.Count; i++)
                {
                    Specialist re;
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
                        {
                            Globals.ThisAddIn.Application.DisplayAlerts = false;
                            Ribbon2.wsRrigin.get_Range("A1", Ribbon2.LastColumnName + Ribbon2.LastRowNumber).Copy(ws.get_Range("A1"));
                            Globals.ThisAddIn.Application.DisplayAlerts = true;
                            System.Windows.Forms.Clipboard.Clear();
                        }
                        string str = list[i].ToString();
                        if (string.IsNullOrEmpty(str)) continue;
                        if (!IsStartingCellRowContainLineIndicator(StartColumn + (j + int.Parse(s)).ToString(), LineIndicator)) continue;
                        if (SessionInfo.UserInfo.CurrentRef != list[i].Reference) continue;
                        SessionInfo.UserInfo.CurrentSaveRef = list[i].SaveReference;
                        SessionInfo.UserInfo.PopulateCellWithJnNumber = list[i].populatecellwithJN;
                        SessionInfo.UserInfo.BalanceBy = list[i].BalanceBy;
                        if (string.IsNullOrEmpty(SessionInfo.UserInfo.AllowBalTran)) SessionInfo.UserInfo.AllowBalTran = list[i].AllowBalTrans;
                        if (string.IsNullOrEmpty(SessionInfo.UserInfo.AllowPostToSuspended)) SessionInfo.UserInfo.AllowPostToSuspended = list[i].AllowPostSuspAcco;
                        if (string.IsNullOrEmpty(SessionInfo.UserInfo.PostProvisional)) SessionInfo.UserInfo.PostProvisional = list[i].PostProvisional;
                        re = new Specialist();
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
                                    y3 = re.AccountingPeriod.Replace("/", "").Substring(0, 4);
                                    m = re.AccountingPeriod.Replace("/", "").Substring(4, 3);
                                    int a = 0;
                                    if ((int.TryParse(y3, out a) == false) || (int.TryParse(m, out a) == false))
                                        throw new Exception("AccountingPeriod format is not correct.");

                                    re.AccountingPeriod = m.PadLeft(3, '0') + y3;//;
                                }
                                else
                                    throw new Exception("AccountingPeriod format is not correct.");
                            }
                        }
                        re.TransactionDate = string.IsNullOrEmpty(list[i].TransactionDate) ? "" : GetEntityValueInExcel(list[i].TransactionDate, s, j, ws);
                        re.DueDate = string.IsNullOrEmpty(list[i].DueDate) ? "" : GetEntityValueInExcel(list[i].DueDate, s, j, ws);
                        re.JournalType = string.IsNullOrEmpty(list[i].JournalType) ? "" : GetEntityValueInExcel(list[i].JournalType, s, j, ws);
                        re.JournalSource = string.IsNullOrEmpty(list[i].JournalSource) ? "" : GetEntityValueInExcel(list[i].JournalSource, s, j, ws);
                        re.TransactionReference = string.IsNullOrEmpty(list[i].TransactionReference) ? "" : GetEntityValueInExcel(list[i].TransactionReference, s, j, ws);
                        re.Description = string.IsNullOrEmpty(list[i].Description) ? "" : GetEntityValueInExcel(list[i].Description, s, j, ws);
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
                        re.GenDesc1 = string.IsNullOrEmpty(list[i].GenDesc1) ? "" : GetEntityValueInExcel(list[i].GenDesc1, s, j, ws);
                        re.GenDesc2 = string.IsNullOrEmpty(list[i].GenDesc2) ? "" : GetEntityValueInExcel(list[i].GenDesc2, s, j, ws);
                        re.GenDesc3 = string.IsNullOrEmpty(list[i].GenDesc3) ? "" : GetEntityValueInExcel(list[i].GenDesc3, s, j, ws);
                        re.GenDesc4 = string.IsNullOrEmpty(list[i].GenDesc4) ? "" : GetEntityValueInExcel(list[i].GenDesc4, s, j, ws);
                        re.GenDesc5 = string.IsNullOrEmpty(list[i].GenDesc5) ? "" : GetEntityValueInExcel(list[i].GenDesc5, s, j, ws);
                        re.GenDesc6 = string.IsNullOrEmpty(list[i].GenDesc6) ? "" : GetEntityValueInExcel(list[i].GenDesc6, s, j, ws);
                        re.GenDesc7 = string.IsNullOrEmpty(list[i].GenDesc7) ? "" : GetEntityValueInExcel(list[i].GenDesc7, s, j, ws);
                        re.GenDesc8 = string.IsNullOrEmpty(list[i].GenDesc8) ? "" : GetEntityValueInExcel(list[i].GenDesc8, s, j, ws);
                        re.GenDesc9 = string.IsNullOrEmpty(list[i].GenDesc9) ? "" : GetEntityValueInExcel(list[i].GenDesc9, s, j, ws);
                        re.GenDesc10 = string.IsNullOrEmpty(list[i].GenDesc10) ? "" : GetEntityValueInExcel(list[i].GenDesc10, s, j, ws);
                        re.GenDesc11 = string.IsNullOrEmpty(list[i].GenDesc11) ? "" : GetEntityValueInExcel(list[i].GenDesc11, s, j, ws);
                        re.GenDesc12 = string.IsNullOrEmpty(list[i].GenDesc12) ? "" : GetEntityValueInExcel(list[i].GenDesc12, s, j, ws);
                        re.GenDesc13 = string.IsNullOrEmpty(list[i].GenDesc13) ? "" : GetEntityValueInExcel(list[i].GenDesc13, s, j, ws);
                        re.GenDesc14 = string.IsNullOrEmpty(list[i].GenDesc14) ? "" : GetEntityValueInExcel(list[i].GenDesc14, s, j, ws);
                        re.GenDesc15 = string.IsNullOrEmpty(list[i].GenDesc15) ? "" : GetEntityValueInExcel(list[i].GenDesc15, s, j, ws);
                        re.GenDesc16 = string.IsNullOrEmpty(list[i].GenDesc16) ? "" : GetEntityValueInExcel(list[i].GenDesc16, s, j, ws);
                        re.GenDesc17 = string.IsNullOrEmpty(list[i].GenDesc17) ? "" : GetEntityValueInExcel(list[i].GenDesc17, s, j, ws);
                        re.GenDesc18 = string.IsNullOrEmpty(list[i].GenDesc18) ? "" : GetEntityValueInExcel(list[i].GenDesc18, s, j, ws);
                        re.GenDesc19 = string.IsNullOrEmpty(list[i].GenDesc19) ? "" : GetEntityValueInExcel(list[i].GenDesc19, s, j, ws);
                        re.GenDesc20 = string.IsNullOrEmpty(list[i].GenDesc20) ? "" : GetEntityValueInExcel(list[i].GenDesc20, s, j, ws);
                        re.GenDesc21 = string.IsNullOrEmpty(list[i].GenDesc21) ? "" : GetEntityValueInExcel(list[i].GenDesc21, s, j, ws);
                        re.GenDesc22 = string.IsNullOrEmpty(list[i].GenDesc22) ? "" : GetEntityValueInExcel(list[i].GenDesc22, s, j, ws);
                        re.GenDesc23 = string.IsNullOrEmpty(list[i].GenDesc23) ? "" : GetEntityValueInExcel(list[i].GenDesc23, s, j, ws);
                        re.GenDesc24 = string.IsNullOrEmpty(list[i].GenDesc24) ? "" : GetEntityValueInExcel(list[i].GenDesc24, s, j, ws);
                        re.GenDesc25 = string.IsNullOrEmpty(list[i].GenDesc25) ? "" : GetEntityValueInExcel(list[i].GenDesc25, s, j, ws);
                        string sTransAmount = GetEntityValueInExcel(list[i].TransactionAmount, s, j, ws);
                        re.TransactionAmount = double.Parse(string.IsNullOrEmpty(sTransAmount) ? "0" : sTransAmount).ToString("0.000");
                        re.CurrencyCode = string.IsNullOrEmpty(list[i].CurrencyCode) ? "" : GetEntityValueInExcel(list[i].CurrencyCode, s, j, ws);
                        re.DebitCredit = string.IsNullOrEmpty(list[i].DebitCredit) ? "" : GetEntityValueInExcel(list[i].DebitCredit, s, j, ws);
                        string sBaseAmount = GetEntityValueInExcel(list[i].BaseAmount, s, j, ws);
                        re.BaseAmount = double.Parse(string.IsNullOrEmpty(sBaseAmount) ? "0" : sBaseAmount).ToString("0.000");
                        string sBase2ReportingAmount = GetEntityValueInExcel(list[i].Base2ReportingAmount, s, j, ws);
                        re.Base2ReportingAmount = double.Parse(string.IsNullOrEmpty(sBase2ReportingAmount) ? "0" : sBase2ReportingAmount).ToString("0.000");
                        string sValue4Amount = GetEntityValueInExcel(list[i].Value4Amount, s, j, ws);
                        re.Value4Amount = double.Parse(string.IsNullOrEmpty(sValue4Amount) ? "0" : sValue4Amount).ToString("0.000");
                    }
                    catch (Exception ex)
                    {
                        LogError("The data in Line " + (j + int.Parse(s)).ToString() + " has error! " + ex.Message + "\r\n");
                        continue;
                    }
                    if (OutputContainer.isTransUpdFlag == false)
                    {
                        if (re.TransactionAmount != "0.000" || re.BaseAmount != "0.000" || re.Base2ReportingAmount != "0.000" || re.Value4Amount != "0.000")
                            newlist.Add(re);
                    }
                    else
                    {
                        newlist.Add(re);
                    }
                }
            }
            return ws;
        }
        /// <summary>
        /// 
        /// </summary>
        /// <param name="strLogessage"></param>
        private delegate void dgetLogError(string strLogessage);
        /// <summary>
        /// 
        /// </summary>
        /// <param name="LogText"></param>
        public void LogError(string LogText)
        {
            try
            {
                if (XMLPostFrm.richTextBox1.InvokeRequired)
                    XMLPostFrm.richTextBox1.BeginInvoke(new dgetLogError(LogError), new object[] { LogText });
                else
                    XMLPostFrm.richTextBox1.Text += LogText;
            }
            catch
            { }
        }
        /// <summary>
        /// 
        /// </summary>
        /// <param name="strLogessage"></param>
        private delegate void dgetLogText(int strLogessage);
        /// <summary>
        /// 
        /// </summary>
        /// <param name="LogText"></param>
        public void Log(int LogText)
        {
            try
            {
                if (XMLPostFrm.progressBar1.InvokeRequired)
                    XMLPostFrm.progressBar1.BeginInvoke(new dgetLogText(Log), new object[] { LogText });
                else
                    XMLPostFrm.progressBar1.Value = LogText;
            }
            catch
            { }
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
        public void BigCalculate(string lastRowNum, List<Specialist> list, string StartColumn, string LineIndicator, string s, Microsoft.Office.Interop.Excel.Worksheet ws, ref List<Specialist> newlist)
        {
            int rowCount = int.Parse(lastRowNum) - int.Parse(s) + 1;
            if (rowCount > 500)
            {
                XMLPostFrm.progressBar1.Visible = true;
                XMLPostFrm.progressBar1.Maximum = rowCount;
            }
            //================================================================below 300======================================================
            List<Specialist> newlist1 = new List<Specialist>();
            if (rowCount <= 300)
            {
                ws = Cal(lastRowNum, list, StartColumn, LineIndicator, s, ws, ref newlist1, 0, rowCount);
                newlist.AddRange(newlist1);
            }
            else if (rowCount > 300)//================================================================above 300======================================================
            {
                ws = Cal(lastRowNum, list, StartColumn, LineIndicator, s, ws, ref newlist1, 0, 100);
                newlist.AddRange(newlist1);
                newlist1.Clear();
                GC.Collect();
                ws = Cal(lastRowNum, list, StartColumn, LineIndicator, s, ws, ref newlist1, 101, 200);
                newlist.AddRange(newlist1);
                newlist1.Clear();
                GC.Collect();
                ws = Cal(lastRowNum, list, StartColumn, LineIndicator, s, ws, ref newlist1, 201, 300);
                newlist.AddRange(newlist1);
                newlist1.Clear();
                GC.Collect();
                //===================================================================================400=============================================================
                if (rowCount <= 400)
                {
                    ws = Cal(lastRowNum, list, StartColumn, LineIndicator, s, ws, ref newlist1, 301, rowCount);
                    newlist.AddRange(newlist1);
                    return;
                }
                ws = Cal(lastRowNum, list, StartColumn, LineIndicator, s, ws, ref newlist1, 301, 400);
                newlist.AddRange(newlist1);
                newlist1.Clear();
                GC.Collect();
                //===================================================================================500=============================================================
                if (rowCount <= 500)
                {
                    ws = Cal(lastRowNum, list, StartColumn, LineIndicator, s, ws, ref newlist1, 401, rowCount);
                    newlist.AddRange(newlist1);
                    return;
                }
                ws = Cal(lastRowNum, list, StartColumn, LineIndicator, s, ws, ref newlist1, 401, 500);
                newlist.AddRange(newlist1);
                newlist1.Clear();
                Log(500);
                GC.Collect();
                //===================================================================================600=============================================================
                if (rowCount <= 600)
                {
                    ws = Cal(lastRowNum, list, StartColumn, LineIndicator, s, ws, ref newlist1, 501, rowCount);
                    newlist.AddRange(newlist1);
                    return;
                }
                ws = Cal(lastRowNum, list, StartColumn, LineIndicator, s, ws, ref newlist1, 501, 600);
                newlist.AddRange(newlist1);
                newlist1.Clear();
                Log(600);
                GC.Collect();
                //===================================================================================700=============================================================
                if (rowCount <= 700)
                {
                    ws = Cal(lastRowNum, list, StartColumn, LineIndicator, s, ws, ref newlist1, 601, rowCount);
                    newlist.AddRange(newlist1);
                    return;
                }
                ws = Cal(lastRowNum, list, StartColumn, LineIndicator, s, ws, ref newlist1, 601, 700);
                newlist.AddRange(newlist1);
                newlist1.Clear();
                Log(700);
                GC.Collect();
                //===================================================================================800=============================================================
                if (rowCount <= 800)
                {
                    ws = Cal(lastRowNum, list, StartColumn, LineIndicator, s, ws, ref newlist1, 701, rowCount);
                    newlist.AddRange(newlist1);
                    return;
                }
                ws = Cal(lastRowNum, list, StartColumn, LineIndicator, s, ws, ref newlist1, 701, 800);
                newlist.AddRange(newlist1);
                newlist1.Clear();
                Log(800);
                GC.Collect();
                //===================================================================================900=============================================================
                if (rowCount <= 900)
                {
                    ws = Cal(lastRowNum, list, StartColumn, LineIndicator, s, ws, ref newlist1, 801, rowCount);
                    newlist.AddRange(newlist1);
                    return;
                }
                ws = Cal(lastRowNum, list, StartColumn, LineIndicator, s, ws, ref newlist1, 801, 900);
                newlist.AddRange(newlist1);
                newlist1.Clear();
                Log(900);
                GC.Collect();
                //===================================================================================1000=============================================================
                if (rowCount <= 1000)
                {
                    ws = Cal(lastRowNum, list, StartColumn, LineIndicator, s, ws, ref newlist1, 901, rowCount);
                    newlist.AddRange(newlist1);
                    return;
                }
                ws = Cal(lastRowNum, list, StartColumn, LineIndicator, s, ws, ref newlist1, 901, 1000);
                newlist.AddRange(newlist1);
                newlist1.Clear();
                Log(1000);
                GC.Collect();
                //===================================================================================1100=============================================================
                if (rowCount <= 1100)
                {
                    ws = Cal(lastRowNum, list, StartColumn, LineIndicator, s, ws, ref newlist1, 1001, rowCount);
                    newlist.AddRange(newlist1);
                    return;
                }
                ws = Cal(lastRowNum, list, StartColumn, LineIndicator, s, ws, ref newlist1, 1001, 1100);
                newlist.AddRange(newlist1);
                newlist1.Clear();
                Log(1100);
                GC.Collect();
                //===================================================================================1200=============================================================
                if (rowCount <= 1200)
                {
                    ws = Cal(lastRowNum, list, StartColumn, LineIndicator, s, ws, ref newlist1, 1101, rowCount);
                    newlist.AddRange(newlist1);
                    return;
                }
                ws = Cal(lastRowNum, list, StartColumn, LineIndicator, s, ws, ref newlist1, 1101, 1200);
                newlist.AddRange(newlist1);
                newlist1.Clear();
                Log(1200);
                GC.Collect();
                //===================================================================================1300=============================================================
                if (rowCount <= 1300)
                {
                    ws = Cal(lastRowNum, list, StartColumn, LineIndicator, s, ws, ref newlist1, 1201, rowCount);
                    newlist.AddRange(newlist1);
                    return;
                }
                ws = Cal(lastRowNum, list, StartColumn, LineIndicator, s, ws, ref newlist1, 1201, 1300);
                newlist.AddRange(newlist1);
                newlist1.Clear();
                Log(1300);
                GC.Collect();
                //===================================================================================1400=============================================================
                if (rowCount <= 1400)
                {
                    ws = Cal(lastRowNum, list, StartColumn, LineIndicator, s, ws, ref newlist1, 1301, rowCount);
                    newlist.AddRange(newlist1);
                    return;
                }
                ws = Cal(lastRowNum, list, StartColumn, LineIndicator, s, ws, ref newlist1, 1301, 1400);
                newlist.AddRange(newlist1);
                newlist1.Clear();
                Log(1400);
                GC.Collect();
                //===================================================================================1500=============================================================
                if (rowCount <= 1500)
                {
                    ws = Cal(lastRowNum, list, StartColumn, LineIndicator, s, ws, ref newlist1, 1401, rowCount);
                    newlist.AddRange(newlist1);
                    return;
                }
                ws = Cal(lastRowNum, list, StartColumn, LineIndicator, s, ws, ref newlist1, 1401, 1500);
                newlist.AddRange(newlist1);
                newlist1.Clear();
                Log(1500);
                GC.Collect();
                //===================================================================================1600=============================================================
                if (rowCount <= 1600)
                {
                    ws = Cal(lastRowNum, list, StartColumn, LineIndicator, s, ws, ref newlist1, 1501, rowCount);
                    newlist.AddRange(newlist1);
                    return;
                }
                ws = Cal(lastRowNum, list, StartColumn, LineIndicator, s, ws, ref newlist1, 1501, 1600);
                newlist.AddRange(newlist1);
                newlist1.Clear();
                Log(1600);
                GC.Collect();
                //===================================================================================1700=============================================================
                if (rowCount <= 1700)
                {
                    ws = Cal(lastRowNum, list, StartColumn, LineIndicator, s, ws, ref newlist1, 1601, rowCount);
                    newlist.AddRange(newlist1);
                    return;
                }
                ws = Cal(lastRowNum, list, StartColumn, LineIndicator, s, ws, ref newlist1, 1601, 1700);
                newlist.AddRange(newlist1);
                newlist1.Clear();
                Log(1700);
                GC.Collect();
                //===================================================================================1800=============================================================
                if (rowCount <= 1800)
                {
                    ws = Cal(lastRowNum, list, StartColumn, LineIndicator, s, ws, ref newlist1, 1701, rowCount);
                    newlist.AddRange(newlist1);
                    return;
                }
                ws = Cal(lastRowNum, list, StartColumn, LineIndicator, s, ws, ref newlist1, 1701, 1800);
                newlist.AddRange(newlist1);
                newlist1.Clear();
                Log(1800);
                GC.Collect();
                //===================================================================================1900=============================================================
                if (rowCount <= 1900)
                {
                    ws = Cal(lastRowNum, list, StartColumn, LineIndicator, s, ws, ref newlist1, 1801, rowCount);
                    newlist.AddRange(newlist1);
                    return;
                }
                ws = Cal(lastRowNum, list, StartColumn, LineIndicator, s, ws, ref newlist1, 1801, 1900);
                newlist.AddRange(newlist1);
                newlist1.Clear();
                Log(1900);
                GC.Collect();
                //===================================================================================2000=============================================================
                if (rowCount <= 2000)
                {
                    ws = Cal(lastRowNum, list, StartColumn, LineIndicator, s, ws, ref newlist1, 1901, rowCount);
                    newlist.AddRange(newlist1);
                    return;
                }
                ws = Cal(lastRowNum, list, StartColumn, LineIndicator, s, ws, ref newlist1, 1901, 2000);
                newlist.AddRange(newlist1);
                newlist1.Clear();
                Log(2000);
                GC.Collect();
                //===================================================================================2100=============================================================
                if (rowCount <= 2100)
                {
                    ws = Cal(lastRowNum, list, StartColumn, LineIndicator, s, ws, ref newlist1, 2001, rowCount);
                    newlist.AddRange(newlist1);
                    return;
                }
                ws = Cal(lastRowNum, list, StartColumn, LineIndicator, s, ws, ref newlist1, 2001, 2100);
                newlist.AddRange(newlist1);
                newlist1.Clear();
                Log(2100);
                GC.Collect();
                //===================================================================================2200=============================================================
                if (rowCount <= 2200)
                {
                    ws = Cal(lastRowNum, list, StartColumn, LineIndicator, s, ws, ref newlist1, 2101, rowCount);
                    newlist.AddRange(newlist1);
                    return;
                }
                ws = Cal(lastRowNum, list, StartColumn, LineIndicator, s, ws, ref newlist1, 2101, 2200);
                newlist.AddRange(newlist1);
                newlist1.Clear();
                Log(2200);
                GC.Collect();
                //===================================================================================2300=============================================================
                if (rowCount <= 2300)
                {
                    ws = Cal(lastRowNum, list, StartColumn, LineIndicator, s, ws, ref newlist1, 2201, rowCount);
                    newlist.AddRange(newlist1);
                    return;
                }
                ws = Cal(lastRowNum, list, StartColumn, LineIndicator, s, ws, ref newlist1, 2201, 2300);
                newlist.AddRange(newlist1);
                newlist1.Clear();
                Log(2300);
                GC.Collect();
                //===================================================================================2400=============================================================
                if (rowCount <= 2400)
                {
                    ws = Cal(lastRowNum, list, StartColumn, LineIndicator, s, ws, ref newlist1, 2301, rowCount);
                    newlist.AddRange(newlist1);
                    return;
                }
                ws = Cal(lastRowNum, list, StartColumn, LineIndicator, s, ws, ref newlist1, 2301, 2400);
                newlist.AddRange(newlist1);
                newlist1.Clear();
                Log(2400);
                GC.Collect();
                //===================================================================================2500=============================================================
                if (rowCount <= 2500)
                {
                    ws = Cal(lastRowNum, list, StartColumn, LineIndicator, s, ws, ref newlist1, 2401, rowCount);
                    newlist.AddRange(newlist1);
                    return;
                }
                ws = Cal(lastRowNum, list, StartColumn, LineIndicator, s, ws, ref newlist1, 2401, 2500);
                newlist.AddRange(newlist1);
                newlist1.Clear();
                Log(2500);
                GC.Collect();
                //===================================================================================2600=============================================================
                if (rowCount <= 2600)
                {
                    ws = Cal(lastRowNum, list, StartColumn, LineIndicator, s, ws, ref newlist1, 2501, rowCount);
                    newlist.AddRange(newlist1);
                    return;
                }
                ws = Cal(lastRowNum, list, StartColumn, LineIndicator, s, ws, ref newlist1, 2501, 2600);
                newlist.AddRange(newlist1);
                newlist1.Clear();
                Log(2600);
                GC.Collect();
                //===================================================================================2700=============================================================
                if (rowCount <= 2700)
                {
                    ws = Cal(lastRowNum, list, StartColumn, LineIndicator, s, ws, ref newlist1, 2601, rowCount);
                    newlist.AddRange(newlist1);
                    return;
                }
                ws = Cal(lastRowNum, list, StartColumn, LineIndicator, s, ws, ref newlist1, 2601, 2700);
                newlist.AddRange(newlist1);
                newlist1.Clear();
                Log(2700);
                GC.Collect();
                //===================================================================================2800=============================================================
                if (rowCount <= 2800)
                {
                    ws = Cal(lastRowNum, list, StartColumn, LineIndicator, s, ws, ref newlist1, 2701, rowCount);
                    newlist.AddRange(newlist1);
                    return;
                }
                ws = Cal(lastRowNum, list, StartColumn, LineIndicator, s, ws, ref newlist1, 2701, 2800);
                newlist.AddRange(newlist1);
                newlist1.Clear();
                Log(2800);
                GC.Collect();
                //===================================================================================2900=============================================================
                if (rowCount <= 2900)
                {
                    ws = Cal(lastRowNum, list, StartColumn, LineIndicator, s, ws, ref newlist1, 2801, rowCount);
                    newlist.AddRange(newlist1);
                    return;
                }
                ws = Cal(lastRowNum, list, StartColumn, LineIndicator, s, ws, ref newlist1, 2801, 2900);
                newlist.AddRange(newlist1);
                newlist1.Clear();
                Log(2900);
                GC.Collect();
                //===================================================================================3000=============================================================
                if (rowCount <= 3000)
                {
                    ws = Cal(lastRowNum, list, StartColumn, LineIndicator, s, ws, ref newlist1, 2901, rowCount);
                    newlist.AddRange(newlist1);
                    return;
                }
                ws = Cal(lastRowNum, list, StartColumn, LineIndicator, s, ws, ref newlist1, 2901, 3000);
                newlist.AddRange(newlist1);
                newlist1.Clear();
                Log(3000);
                GC.Collect();
                //===================================================================================3100=============================================================
                if (rowCount <= 3100)
                {
                    ws = Cal(lastRowNum, list, StartColumn, LineIndicator, s, ws, ref newlist1, 3001, rowCount);
                    newlist.AddRange(newlist1);
                    return;
                }
                ws = Cal(lastRowNum, list, StartColumn, LineIndicator, s, ws, ref newlist1, 3001, 3100);
                newlist.AddRange(newlist1);
                newlist1.Clear();
                Log(3100);
                GC.Collect();
                //===================================================================================3200=============================================================
                if (rowCount <= 3200)
                {
                    ws = Cal(lastRowNum, list, StartColumn, LineIndicator, s, ws, ref newlist1, 3101, rowCount);
                    newlist.AddRange(newlist1);
                    return;
                }
                ws = Cal(lastRowNum, list, StartColumn, LineIndicator, s, ws, ref newlist1, 3101, 3200);
                newlist.AddRange(newlist1);
                newlist1.Clear();
                Log(3200);
                GC.Collect();
                //===================================================================================3300=============================================================
                if (rowCount <= 3300)
                {
                    ws = Cal(lastRowNum, list, StartColumn, LineIndicator, s, ws, ref newlist1, 3201, rowCount);
                    newlist.AddRange(newlist1);
                    return;
                }
                ws = Cal(lastRowNum, list, StartColumn, LineIndicator, s, ws, ref newlist1, 3201, 3300);
                newlist.AddRange(newlist1);
                newlist1.Clear();
                Log(3300);
                GC.Collect();
                //===================================================================================3400=============================================================
                if (rowCount <= 3400)
                {
                    ws = Cal(lastRowNum, list, StartColumn, LineIndicator, s, ws, ref newlist1, 3301, rowCount);
                    newlist.AddRange(newlist1);
                    return;
                }
                ws = Cal(lastRowNum, list, StartColumn, LineIndicator, s, ws, ref newlist1, 3301, 3400);
                newlist.AddRange(newlist1);
                newlist1.Clear();
                Log(3400);
                GC.Collect();
                //===================================================================================3500=============================================================
                if (rowCount <= 3500)
                {
                    ws = Cal(lastRowNum, list, StartColumn, LineIndicator, s, ws, ref newlist1, 3401, rowCount);
                    newlist.AddRange(newlist1);
                    return;
                }
                ws = Cal(lastRowNum, list, StartColumn, LineIndicator, s, ws, ref newlist1, 3401, 3500);
                newlist.AddRange(newlist1);
                newlist1.Clear();
                Log(3500);
                GC.Collect();
                //===================================================================================3600=============================================================
                if (rowCount <= 3600)
                {
                    ws = Cal(lastRowNum, list, StartColumn, LineIndicator, s, ws, ref newlist1, 3501, rowCount);
                    newlist.AddRange(newlist1);
                    return;
                }
                ws = Cal(lastRowNum, list, StartColumn, LineIndicator, s, ws, ref newlist1, 3501, 3600);
                newlist.AddRange(newlist1);
                newlist1.Clear();
                Log(3600);
                GC.Collect();
                //===================================================================================3700=============================================================
                if (rowCount <= 3700)
                {
                    ws = Cal(lastRowNum, list, StartColumn, LineIndicator, s, ws, ref newlist1, 3601, rowCount);
                    newlist.AddRange(newlist1);
                    return;
                }
                ws = Cal(lastRowNum, list, StartColumn, LineIndicator, s, ws, ref newlist1, 3601, 3700);
                newlist.AddRange(newlist1);
                newlist1.Clear();
                Log(3700);
                GC.Collect();
                //===================================================================================3800=============================================================
                if (rowCount <= 3800)
                {
                    ws = Cal(lastRowNum, list, StartColumn, LineIndicator, s, ws, ref newlist1, 3701, rowCount);
                    newlist.AddRange(newlist1);
                    return;
                }
                ws = Cal(lastRowNum, list, StartColumn, LineIndicator, s, ws, ref newlist1, 3701, 3800);
                newlist.AddRange(newlist1);
                newlist1.Clear();
                Log(3800);
                GC.Collect();
                //===================================================================================3900=============================================================
                if (rowCount <= 3900)
                {
                    ws = Cal(lastRowNum, list, StartColumn, LineIndicator, s, ws, ref newlist1, 3801, rowCount);
                    newlist.AddRange(newlist1);
                    return;
                }
                ws = Cal(lastRowNum, list, StartColumn, LineIndicator, s, ws, ref newlist1, 3801, 3900);
                newlist.AddRange(newlist1);
                newlist1.Clear();
                Log(3900);
                GC.Collect();
                //===================================================================================4000=============================================================
                if (rowCount <= 4000)
                {
                    ws = Cal(lastRowNum, list, StartColumn, LineIndicator, s, ws, ref newlist1, 3901, rowCount);
                    newlist.AddRange(newlist1);
                    return;
                }
                ws = Cal(lastRowNum, list, StartColumn, LineIndicator, s, ws, ref newlist1, 3901, 4000);
                newlist.AddRange(newlist1);
                newlist1.Clear();
                Log(4000);
                GC.Collect();
                //===================================================================================4100=============================================================
                if (rowCount <= 4100)
                {
                    ws = Cal(lastRowNum, list, StartColumn, LineIndicator, s, ws, ref newlist1, 4001, rowCount);
                    newlist.AddRange(newlist1);
                    return;
                }
                ws = Cal(lastRowNum, list, StartColumn, LineIndicator, s, ws, ref newlist1, 4001, 4100);
                newlist.AddRange(newlist1);
                newlist1.Clear();
                Log(4100);
                GC.Collect();
                //===================================================================================4200=============================================================
                if (rowCount <= 4200)
                {
                    ws = Cal(lastRowNum, list, StartColumn, LineIndicator, s, ws, ref newlist1, 4101, rowCount);
                    newlist.AddRange(newlist1);
                    return;
                }
                ws = Cal(lastRowNum, list, StartColumn, LineIndicator, s, ws, ref newlist1, 4101, 4200);
                newlist.AddRange(newlist1);
                newlist1.Clear();
                Log(4200);
                GC.Collect();
                //===================================================================================4300=============================================================
                if (rowCount <= 4300)
                {
                    ws = Cal(lastRowNum, list, StartColumn, LineIndicator, s, ws, ref newlist1, 4201, rowCount);
                    newlist.AddRange(newlist1);
                    return;
                }
                ws = Cal(lastRowNum, list, StartColumn, LineIndicator, s, ws, ref newlist1, 4201, 4300);
                newlist.AddRange(newlist1);
                newlist1.Clear();
                Log(4300);
                GC.Collect();
                //===================================================================================4400=============================================================
                if (rowCount <= 4400)
                {
                    ws = Cal(lastRowNum, list, StartColumn, LineIndicator, s, ws, ref newlist1, 4301, rowCount);
                    newlist.AddRange(newlist1);
                    return;
                }
                ws = Cal(lastRowNum, list, StartColumn, LineIndicator, s, ws, ref newlist1, 4301, 4400);
                newlist.AddRange(newlist1);
                newlist1.Clear();
                Log(4400);
                GC.Collect();
                //===================================================================================4500=============================================================
                if (rowCount <= 4500)
                {
                    ws = Cal(lastRowNum, list, StartColumn, LineIndicator, s, ws, ref newlist1, 4401, rowCount);
                    newlist.AddRange(newlist1);
                    return;
                }
                ws = Cal(lastRowNum, list, StartColumn, LineIndicator, s, ws, ref newlist1, 4401, 4500);
                newlist.AddRange(newlist1);
                newlist1.Clear();
                Log(4500);
                GC.Collect();
                //===================================================================================4600=============================================================
                if (rowCount <= 4600)
                {
                    ws = Cal(lastRowNum, list, StartColumn, LineIndicator, s, ws, ref newlist1, 4501, rowCount);
                    newlist.AddRange(newlist1);
                    return;
                }
                ws = Cal(lastRowNum, list, StartColumn, LineIndicator, s, ws, ref newlist1, 4501, 4600);
                newlist.AddRange(newlist1);
                newlist1.Clear();
                Log(4600);
                GC.Collect();
                //===================================================================================4700=============================================================
                if (rowCount <= 4700)
                {
                    ws = Cal(lastRowNum, list, StartColumn, LineIndicator, s, ws, ref newlist1, 4601, rowCount);
                    newlist.AddRange(newlist1);
                    return;
                }
                ws = Cal(lastRowNum, list, StartColumn, LineIndicator, s, ws, ref newlist1, 4601, 4700);
                newlist.AddRange(newlist1);
                newlist1.Clear();
                Log(4700);
                GC.Collect();
                //===================================================================================4800=============================================================
                if (rowCount <= 4800)
                {
                    ws = Cal(lastRowNum, list, StartColumn, LineIndicator, s, ws, ref newlist1, 4701, rowCount);
                    newlist.AddRange(newlist1);
                    return;
                }
                ws = Cal(lastRowNum, list, StartColumn, LineIndicator, s, ws, ref newlist1, 4701, 4800);
                newlist.AddRange(newlist1);
                newlist1.Clear();
                Log(4800);
                GC.Collect();
                //===================================================================================4900=============================================================
                if (rowCount <= 4900)
                {
                    ws = Cal(lastRowNum, list, StartColumn, LineIndicator, s, ws, ref newlist1, 4801, rowCount);
                    newlist.AddRange(newlist1);
                    return;
                }
                ws = Cal(lastRowNum, list, StartColumn, LineIndicator, s, ws, ref newlist1, 4801, 4900);
                newlist.AddRange(newlist1);
                newlist1.Clear();
                Log(4900);
                GC.Collect();
                //===================================================================================5000=============================================================
                if (rowCount <= 5000)
                {
                    ws = Cal(lastRowNum, list, StartColumn, LineIndicator, s, ws, ref newlist1, 4901, rowCount);
                    newlist.AddRange(newlist1);
                    return;
                }
                ws = Cal(lastRowNum, list, StartColumn, LineIndicator, s, ws, ref newlist1, 4901, 5000);
                newlist.AddRange(newlist1);
                newlist1.Clear();
                Log(5000);
                GC.Collect();
                //===================================================================================5100=============================================================
                if (rowCount <= 5100)
                {
                    ws = Cal(lastRowNum, list, StartColumn, LineIndicator, s, ws, ref newlist1, 5001, rowCount);
                    newlist.AddRange(newlist1);
                    return;
                }
                ws = Cal(lastRowNum, list, StartColumn, LineIndicator, s, ws, ref newlist1, 5001, 5100);
                newlist.AddRange(newlist1);
                newlist1.Clear();
                Log(5100);
                GC.Collect();
                //===================================================================================5200=============================================================
                if (rowCount <= 5200)
                {
                    ws = Cal(lastRowNum, list, StartColumn, LineIndicator, s, ws, ref newlist1, 5101, rowCount);
                    newlist.AddRange(newlist1);
                    return;
                }
                ws = Cal(lastRowNum, list, StartColumn, LineIndicator, s, ws, ref newlist1, 5101, 5200);
                newlist.AddRange(newlist1);
                newlist1.Clear();
                Log(5200);
                GC.Collect();
                //===================================================================================5300=============================================================
                if (rowCount <= 5300)
                {
                    ws = Cal(lastRowNum, list, StartColumn, LineIndicator, s, ws, ref newlist1, 5201, rowCount);
                    newlist.AddRange(newlist1);
                    return;
                }
                ws = Cal(lastRowNum, list, StartColumn, LineIndicator, s, ws, ref newlist1, 5201, 5300);
                newlist.AddRange(newlist1);
                newlist1.Clear();
                Log(5300);
                GC.Collect();
                //===================================================================================5400=============================================================
                if (rowCount <= 5400)
                {
                    ws = Cal(lastRowNum, list, StartColumn, LineIndicator, s, ws, ref newlist1, 5301, rowCount);
                    newlist.AddRange(newlist1);
                    return;
                }
                ws = Cal(lastRowNum, list, StartColumn, LineIndicator, s, ws, ref newlist1, 5301, 5400);
                newlist.AddRange(newlist1);
                newlist1.Clear();
                Log(5400);
                GC.Collect();
                //===================================================================================5500=============================================================
                if (rowCount <= 5500)
                {
                    ws = Cal(lastRowNum, list, StartColumn, LineIndicator, s, ws, ref newlist1, 5401, rowCount);
                    newlist.AddRange(newlist1);
                    return;
                }
                ws = Cal(lastRowNum, list, StartColumn, LineIndicator, s, ws, ref newlist1, 5401, 5500);
                newlist.AddRange(newlist1);
                newlist1.Clear();
                Log(5500);
                GC.Collect();
                //===================================================================================5600=============================================================
                if (rowCount <= 5600)
                {
                    ws = Cal(lastRowNum, list, StartColumn, LineIndicator, s, ws, ref newlist1, 5501, rowCount);
                    newlist.AddRange(newlist1);
                    return;
                }
                ws = Cal(lastRowNum, list, StartColumn, LineIndicator, s, ws, ref newlist1, 5501, 5600);
                newlist.AddRange(newlist1);
                newlist1.Clear();
                Log(5600);
                GC.Collect();
                //===================================================================================5700=============================================================
                if (rowCount <= 5700)
                {
                    ws = Cal(lastRowNum, list, StartColumn, LineIndicator, s, ws, ref newlist1, 5601, rowCount);
                    newlist.AddRange(newlist1);
                    return;
                }
                ws = Cal(lastRowNum, list, StartColumn, LineIndicator, s, ws, ref newlist1, 5601, 5700);
                newlist.AddRange(newlist1);
                newlist1.Clear();
                Log(5700);
                GC.Collect();
                //===================================================================================5800=============================================================
                if (rowCount <= 5800)
                {
                    ws = Cal(lastRowNum, list, StartColumn, LineIndicator, s, ws, ref newlist1, 5701, rowCount);
                    newlist.AddRange(newlist1);
                    return;
                }
                ws = Cal(lastRowNum, list, StartColumn, LineIndicator, s, ws, ref newlist1, 5701, 5800);
                newlist.AddRange(newlist1);
                newlist1.Clear();
                Log(5800);
                GC.Collect();
                //===================================================================================5900=============================================================
                if (rowCount <= 5900)
                {
                    ws = Cal(lastRowNum, list, StartColumn, LineIndicator, s, ws, ref newlist1, 5801, rowCount);
                    newlist.AddRange(newlist1);
                    return;
                }
                ws = Cal(lastRowNum, list, StartColumn, LineIndicator, s, ws, ref newlist1, 5801, 5900);
                newlist.AddRange(newlist1);
                newlist1.Clear();
                Log(5900);
                GC.Collect();
                //===================================================================================6000=============================================================
                if (rowCount <= 6000)
                {
                    ws = Cal(lastRowNum, list, StartColumn, LineIndicator, s, ws, ref newlist1, 5901, rowCount);
                    newlist.AddRange(newlist1);
                    return;
                }
                ws = Cal(lastRowNum, list, StartColumn, LineIndicator, s, ws, ref newlist1, 5901, 6000);
                newlist.AddRange(newlist1);
                newlist1.Clear();
                Log(6000);
                GC.Collect();
                //===================================================================================6100=============================================================
                if (rowCount <= 6100)
                {
                    ws = Cal(lastRowNum, list, StartColumn, LineIndicator, s, ws, ref newlist1, 6001, rowCount);
                    newlist.AddRange(newlist1);
                    return;
                }
                ws = Cal(lastRowNum, list, StartColumn, LineIndicator, s, ws, ref newlist1, 6001, 6100);
                newlist.AddRange(newlist1);
                newlist1.Clear();
                Log(6100);
                GC.Collect();
                //===================================================================================6200=============================================================
                if (rowCount <= 6200)
                {
                    ws = Cal(lastRowNum, list, StartColumn, LineIndicator, s, ws, ref newlist1, 6101, rowCount);
                    newlist.AddRange(newlist1);
                    return;
                }
                ws = Cal(lastRowNum, list, StartColumn, LineIndicator, s, ws, ref newlist1, 6101, 6200);
                newlist.AddRange(newlist1);
                newlist1.Clear();
                Log(6200);
                GC.Collect();
                //===================================================================================6300=============================================================
                if (rowCount <= 6300)
                {
                    ws = Cal(lastRowNum, list, StartColumn, LineIndicator, s, ws, ref newlist1, 6201, rowCount);
                    newlist.AddRange(newlist1);
                    return;
                }
                ws = Cal(lastRowNum, list, StartColumn, LineIndicator, s, ws, ref newlist1, 6201, 6300);
                newlist.AddRange(newlist1);
                newlist1.Clear();
                Log(6300);
                GC.Collect();
                //===================================================================================6400=============================================================
                if (rowCount <= 6400)
                {
                    ws = Cal(lastRowNum, list, StartColumn, LineIndicator, s, ws, ref newlist1, 6301, rowCount);
                    newlist.AddRange(newlist1);
                    return;
                }
                ws = Cal(lastRowNum, list, StartColumn, LineIndicator, s, ws, ref newlist1, 6301, 6400);
                newlist.AddRange(newlist1);
                newlist1.Clear();
                Log(6400);
                GC.Collect();
                //===================================================================================6500=============================================================
                if (rowCount <= 6500)
                {
                    ws = Cal(lastRowNum, list, StartColumn, LineIndicator, s, ws, ref newlist1, 6401, rowCount);
                    newlist.AddRange(newlist1);
                    return;
                }
                ws = Cal(lastRowNum, list, StartColumn, LineIndicator, s, ws, ref newlist1, 6401, 6500);
                newlist.AddRange(newlist1);
                newlist1.Clear();
                Log(6500);
                GC.Collect();
                //===================================================================================6600=============================================================
                if (rowCount <= 6600)
                {
                    ws = Cal(lastRowNum, list, StartColumn, LineIndicator, s, ws, ref newlist1, 6501, rowCount);
                    newlist.AddRange(newlist1);
                    return;
                }
                ws = Cal(lastRowNum, list, StartColumn, LineIndicator, s, ws, ref newlist1, 6501, 6600);
                newlist.AddRange(newlist1);
                newlist1.Clear();
                Log(6600);
                GC.Collect();
                //===================================================================================6700=============================================================
                if (rowCount <= 6700)
                {
                    ws = Cal(lastRowNum, list, StartColumn, LineIndicator, s, ws, ref newlist1, 6601, rowCount);
                    newlist.AddRange(newlist1);
                    return;
                }
                ws = Cal(lastRowNum, list, StartColumn, LineIndicator, s, ws, ref newlist1, 6601, 6700);
                newlist.AddRange(newlist1);
                newlist1.Clear();
                Log(6700);
                GC.Collect();
                //===================================================================================6800=============================================================
                if (rowCount <= 6800)
                {
                    ws = Cal(lastRowNum, list, StartColumn, LineIndicator, s, ws, ref newlist1, 6701, rowCount);
                    newlist.AddRange(newlist1);
                    return;
                }
                ws = Cal(lastRowNum, list, StartColumn, LineIndicator, s, ws, ref newlist1, 6701, 6800);
                newlist.AddRange(newlist1);
                newlist1.Clear();
                Log(6800);
                GC.Collect();
                //===================================================================================6900=============================================================
                if (rowCount <= 6900)
                {
                    ws = Cal(lastRowNum, list, StartColumn, LineIndicator, s, ws, ref newlist1, 6801, rowCount);
                    newlist.AddRange(newlist1);
                    return;
                }
                ws = Cal(lastRowNum, list, StartColumn, LineIndicator, s, ws, ref newlist1, 6801, 6900);
                newlist.AddRange(newlist1);
                newlist1.Clear();
                Log(6900);
                GC.Collect();
                //===================================================================================7000=============================================================
                if (rowCount <= 7000)
                {
                    ws = Cal(lastRowNum, list, StartColumn, LineIndicator, s, ws, ref newlist1, 6901, rowCount);
                    newlist.AddRange(newlist1);
                    return;
                }
                ws = Cal(lastRowNum, list, StartColumn, LineIndicator, s, ws, ref newlist1, 6901, 7000);
                newlist.AddRange(newlist1);
                newlist1.Clear();
                Log(7000);
                GC.Collect();
                //===================================================================================7100=============================================================
                if (rowCount <= 7100)
                {
                    ws = Cal(lastRowNum, list, StartColumn, LineIndicator, s, ws, ref newlist1, 7001, rowCount);
                    newlist.AddRange(newlist1);
                    return;
                }
                ws = Cal(lastRowNum, list, StartColumn, LineIndicator, s, ws, ref newlist1, 7001, 7100);
                newlist.AddRange(newlist1);
                newlist1.Clear();
                Log(7100);
                GC.Collect();
                //===================================================================================7200=============================================================
                if (rowCount <= 7200)
                {
                    ws = Cal(lastRowNum, list, StartColumn, LineIndicator, s, ws, ref newlist1, 7101, rowCount);
                    newlist.AddRange(newlist1);
                    return;
                }
                ws = Cal(lastRowNum, list, StartColumn, LineIndicator, s, ws, ref newlist1, 7101, 7200);
                newlist.AddRange(newlist1);
                newlist1.Clear();
                Log(7200);
                GC.Collect();
                //===================================================================================7300=============================================================
                if (rowCount <= 7300)
                {
                    ws = Cal(lastRowNum, list, StartColumn, LineIndicator, s, ws, ref newlist1, 7201, rowCount);
                    newlist.AddRange(newlist1);
                    return;
                }
                ws = Cal(lastRowNum, list, StartColumn, LineIndicator, s, ws, ref newlist1, 7201, 7300);
                newlist.AddRange(newlist1);
                newlist1.Clear();
                Log(7300);
                GC.Collect();
                //===================================================================================7400=============================================================
                if (rowCount <= 7400)
                {
                    ws = Cal(lastRowNum, list, StartColumn, LineIndicator, s, ws, ref newlist1, 7301, rowCount);
                    newlist.AddRange(newlist1);
                    return;
                }
                ws = Cal(lastRowNum, list, StartColumn, LineIndicator, s, ws, ref newlist1, 7301, 7400);
                newlist.AddRange(newlist1);
                newlist1.Clear();
                Log(7400);
                GC.Collect();
                //===================================================================================7500=============================================================
                if (rowCount <= 7500)
                {
                    ws = Cal(lastRowNum, list, StartColumn, LineIndicator, s, ws, ref newlist1, 7401, rowCount);
                    newlist.AddRange(newlist1);
                    return;
                }
                ws = Cal(lastRowNum, list, StartColumn, LineIndicator, s, ws, ref newlist1, 7401, 7500);
                newlist.AddRange(newlist1);
                newlist1.Clear();
                Log(7500);
                GC.Collect();
                //===================================================================================7600=============================================================
                if (rowCount <= 7600)
                {
                    ws = Cal(lastRowNum, list, StartColumn, LineIndicator, s, ws, ref newlist1, 7501, rowCount);
                    newlist.AddRange(newlist1);
                    return;
                }
                ws = Cal(lastRowNum, list, StartColumn, LineIndicator, s, ws, ref newlist1, 7501, 7600);
                newlist.AddRange(newlist1);
                newlist1.Clear();
                Log(7600);
                GC.Collect();
                //===================================================================================7700=============================================================
                if (rowCount <= 7700)
                {
                    ws = Cal(lastRowNum, list, StartColumn, LineIndicator, s, ws, ref newlist1, 7601, rowCount);
                    newlist.AddRange(newlist1);
                    return;
                }
                ws = Cal(lastRowNum, list, StartColumn, LineIndicator, s, ws, ref newlist1, 7601, 7700);
                newlist.AddRange(newlist1);
                newlist1.Clear();
                Log(7700);
                GC.Collect();
                //===================================================================================7800=============================================================
                if (rowCount <= 7800)
                {
                    ws = Cal(lastRowNum, list, StartColumn, LineIndicator, s, ws, ref newlist1, 7701, rowCount);
                    newlist.AddRange(newlist1);
                    return;
                }
                ws = Cal(lastRowNum, list, StartColumn, LineIndicator, s, ws, ref newlist1, 7701, 7800);
                newlist.AddRange(newlist1);
                newlist1.Clear();
                Log(7800);
                GC.Collect();
                //===================================================================================7900=============================================================
                if (rowCount <= 7900)
                {
                    ws = Cal(lastRowNum, list, StartColumn, LineIndicator, s, ws, ref newlist1, 7801, rowCount);
                    newlist.AddRange(newlist1);
                    return;
                }
                ws = Cal(lastRowNum, list, StartColumn, LineIndicator, s, ws, ref newlist1, 7801, 7900);
                newlist.AddRange(newlist1);
                newlist1.Clear();
                Log(7900);
                GC.Collect();
                //===================================================================================8000=============================================================
                if (rowCount <= 8000)
                {
                    ws = Cal(lastRowNum, list, StartColumn, LineIndicator, s, ws, ref newlist1, 7901, rowCount);
                    newlist.AddRange(newlist1);
                    return;
                }
                ws = Cal(lastRowNum, list, StartColumn, LineIndicator, s, ws, ref newlist1, 7901, 8000);
                newlist.AddRange(newlist1);
                newlist1.Clear();
                Log(8000);
                GC.Collect();
                //===================================================================================8100=============================================================
                if (rowCount <= 8100)
                {
                    ws = Cal(lastRowNum, list, StartColumn, LineIndicator, s, ws, ref newlist1, 8001, rowCount);
                    newlist.AddRange(newlist1);
                    return;
                }
                ws = Cal(lastRowNum, list, StartColumn, LineIndicator, s, ws, ref newlist1, 8001, 8100);
                newlist.AddRange(newlist1);
                newlist1.Clear();
                Log(8100);
                GC.Collect();
                //===================================================================================8200=============================================================
                if (rowCount <= 8200)
                {
                    ws = Cal(lastRowNum, list, StartColumn, LineIndicator, s, ws, ref newlist1, 8101, rowCount);
                    newlist.AddRange(newlist1);
                    return;
                }
                ws = Cal(lastRowNum, list, StartColumn, LineIndicator, s, ws, ref newlist1, 8101, 8200);
                newlist.AddRange(newlist1);
                newlist1.Clear();
                Log(8200);
                GC.Collect();
                //===================================================================================8300=============================================================
                if (rowCount <= 8300)
                {
                    ws = Cal(lastRowNum, list, StartColumn, LineIndicator, s, ws, ref newlist1, 8201, rowCount);
                    newlist.AddRange(newlist1);
                    return;
                }
                ws = Cal(lastRowNum, list, StartColumn, LineIndicator, s, ws, ref newlist1, 8201, 8300);
                newlist.AddRange(newlist1);
                newlist1.Clear();
                Log(8300);
                GC.Collect();
                //===================================================================================8400=============================================================
                if (rowCount <= 8400)
                {
                    ws = Cal(lastRowNum, list, StartColumn, LineIndicator, s, ws, ref newlist1, 8301, rowCount);
                    newlist.AddRange(newlist1);
                    return;
                }
                ws = Cal(lastRowNum, list, StartColumn, LineIndicator, s, ws, ref newlist1, 8301, 8400);
                newlist.AddRange(newlist1);
                newlist1.Clear();
                Log(8400);
                GC.Collect();
                //===================================================================================8500=============================================================
                if (rowCount <= 8500)
                {
                    ws = Cal(lastRowNum, list, StartColumn, LineIndicator, s, ws, ref newlist1, 8401, rowCount);
                    newlist.AddRange(newlist1);
                    return;
                }
                ws = Cal(lastRowNum, list, StartColumn, LineIndicator, s, ws, ref newlist1, 8401, 8500);
                newlist.AddRange(newlist1);
                newlist1.Clear();
                Log(8500);
                GC.Collect();
                //===================================================================================8600=============================================================
                if (rowCount <= 8600)
                {
                    ws = Cal(lastRowNum, list, StartColumn, LineIndicator, s, ws, ref newlist1, 8501, rowCount);
                    newlist.AddRange(newlist1);
                    return;
                }
                ws = Cal(lastRowNum, list, StartColumn, LineIndicator, s, ws, ref newlist1, 8501, 8600);
                newlist.AddRange(newlist1);
                newlist1.Clear();
                Log(8600);
                GC.Collect();
                //===================================================================================8700=============================================================
                if (rowCount <= 8700)
                {
                    ws = Cal(lastRowNum, list, StartColumn, LineIndicator, s, ws, ref newlist1, 8601, rowCount);
                    newlist.AddRange(newlist1);
                    return;
                }
                ws = Cal(lastRowNum, list, StartColumn, LineIndicator, s, ws, ref newlist1, 8601, 8700);
                newlist.AddRange(newlist1);
                newlist1.Clear();
                Log(8700);
                GC.Collect();
                //===================================================================================8800=============================================================
                if (rowCount <= 8800)
                {
                    ws = Cal(lastRowNum, list, StartColumn, LineIndicator, s, ws, ref newlist1, 8701, rowCount);
                    newlist.AddRange(newlist1);
                    return;
                }
                ws = Cal(lastRowNum, list, StartColumn, LineIndicator, s, ws, ref newlist1, 8701, 8800);
                newlist.AddRange(newlist1);
                newlist1.Clear();
                Log(8800);
                GC.Collect();
                //===================================================================================8900=============================================================
                if (rowCount <= 8900)
                {
                    ws = Cal(lastRowNum, list, StartColumn, LineIndicator, s, ws, ref newlist1, 8801, rowCount);
                    newlist.AddRange(newlist1);
                    return;
                }
                ws = Cal(lastRowNum, list, StartColumn, LineIndicator, s, ws, ref newlist1, 8801, 8900);
                newlist.AddRange(newlist1);
                newlist1.Clear();
                Log(8900);
                GC.Collect();
                //===================================================================================9000=============================================================
                if (rowCount <= 9000)
                {
                    ws = Cal(lastRowNum, list, StartColumn, LineIndicator, s, ws, ref newlist1, 8901, rowCount);
                    newlist.AddRange(newlist1);
                    return;
                }
                ws = Cal(lastRowNum, list, StartColumn, LineIndicator, s, ws, ref newlist1, 8901, 9000);
                newlist.AddRange(newlist1);
                newlist1.Clear();
                Log(9000);
                GC.Collect();
                //===================================================================================9100=============================================================
                if (rowCount <= 9100)
                {
                    ws = Cal(lastRowNum, list, StartColumn, LineIndicator, s, ws, ref newlist1, 9001, rowCount);
                    newlist.AddRange(newlist1);
                    return;
                }
                ws = Cal(lastRowNum, list, StartColumn, LineIndicator, s, ws, ref newlist1, 9001, 9100);
                newlist.AddRange(newlist1);
                newlist1.Clear();
                Log(9100);
                GC.Collect();
                //===================================================================================9200=============================================================
                if (rowCount <= 9200)
                {
                    ws = Cal(lastRowNum, list, StartColumn, LineIndicator, s, ws, ref newlist1, 9101, rowCount);
                    newlist.AddRange(newlist1);
                    return;
                }
                ws = Cal(lastRowNum, list, StartColumn, LineIndicator, s, ws, ref newlist1, 9101, 9200);
                newlist.AddRange(newlist1);
                newlist1.Clear();
                Log(9200);
                GC.Collect();
                //===================================================================================9300=============================================================
                if (rowCount <= 9300)
                {
                    ws = Cal(lastRowNum, list, StartColumn, LineIndicator, s, ws, ref newlist1, 9201, rowCount);
                    newlist.AddRange(newlist1);
                    return;
                }
                ws = Cal(lastRowNum, list, StartColumn, LineIndicator, s, ws, ref newlist1, 9201, 9300);
                newlist.AddRange(newlist1);
                newlist1.Clear();
                Log(9300);
                GC.Collect();
                //===================================================================================9400=============================================================
                if (rowCount <= 9400)
                {
                    ws = Cal(lastRowNum, list, StartColumn, LineIndicator, s, ws, ref newlist1, 9301, rowCount);
                    newlist.AddRange(newlist1);
                    return;
                }
                ws = Cal(lastRowNum, list, StartColumn, LineIndicator, s, ws, ref newlist1, 9301, 9400);
                newlist.AddRange(newlist1);
                newlist1.Clear();
                Log(9400);
                GC.Collect();
                //===================================================================================9500=============================================================
                if (rowCount <= 9500)
                {
                    ws = Cal(lastRowNum, list, StartColumn, LineIndicator, s, ws, ref newlist1, 9401, rowCount);
                    newlist.AddRange(newlist1);
                    return;
                }
                ws = Cal(lastRowNum, list, StartColumn, LineIndicator, s, ws, ref newlist1, 9401, 9500);
                newlist.AddRange(newlist1);
                newlist1.Clear();
                Log(9500);
                GC.Collect();
                //===================================================================================9600=============================================================
                if (rowCount <= 9600)
                {
                    ws = Cal(lastRowNum, list, StartColumn, LineIndicator, s, ws, ref newlist1, 9501, rowCount);
                    newlist.AddRange(newlist1);
                    return;
                }
                ws = Cal(lastRowNum, list, StartColumn, LineIndicator, s, ws, ref newlist1, 9501, 9600);
                newlist.AddRange(newlist1);
                newlist1.Clear();
                Log(9600);
                GC.Collect();
                //===================================================================================9700=============================================================
                if (rowCount <= 9700)
                {
                    ws = Cal(lastRowNum, list, StartColumn, LineIndicator, s, ws, ref newlist1, 9601, rowCount);
                    newlist.AddRange(newlist1);
                    return;
                }
                ws = Cal(lastRowNum, list, StartColumn, LineIndicator, s, ws, ref newlist1, 9601, 9700);
                newlist.AddRange(newlist1);
                newlist1.Clear();
                Log(9700);
                GC.Collect();
                //===================================================================================9800=============================================================
                if (rowCount <= 9800)
                {
                    ws = Cal(lastRowNum, list, StartColumn, LineIndicator, s, ws, ref newlist1, 9701, rowCount);
                    newlist.AddRange(newlist1);
                    return;
                }
                ws = Cal(lastRowNum, list, StartColumn, LineIndicator, s, ws, ref newlist1, 9701, 9800);
                newlist.AddRange(newlist1);
                newlist1.Clear();
                Log(9800);
                GC.Collect();
                //===================================================================================9900=============================================================
                if (rowCount <= 9900)
                {
                    ws = Cal(lastRowNum, list, StartColumn, LineIndicator, s, ws, ref newlist1, 9801, rowCount);
                    newlist.AddRange(newlist1);
                    return;
                }
                ws = Cal(lastRowNum, list, StartColumn, LineIndicator, s, ws, ref newlist1, 9801, 9900);
                newlist.AddRange(newlist1);
                newlist1.Clear();
                Log(9900);
                GC.Collect();
                //===================================================================================10000=============================================================
                if (rowCount <= 10000)
                {
                    ws = Cal(lastRowNum, list, StartColumn, LineIndicator, s, ws, ref newlist1, 9901, rowCount);
                    newlist.AddRange(newlist1);
                    return;
                }
                ws = Cal(lastRowNum, list, StartColumn, LineIndicator, s, ws, ref newlist1, 9901, 10000);
                newlist.AddRange(newlist1);
                Log(10000);
            }
        }
    }
}
