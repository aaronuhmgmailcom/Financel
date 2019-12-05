using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace ExcelAddIn4.Common
{
    public class AMUEntityForSave
    {
        /// <summary>
        /// 
        /// </summary>
        /// <returns></returns>
        public override string ToString()
        {
            return Ledger + AccountCode + AccountingPeriod + this.AllocationMarker + this.AnalysisCode1 + this.AnalysisCode2 + this.AnalysisCode3 + AnalysisCode4 + this.AnalysisCode5 + this.AnalysisCode6 + this.AnalysisCode7 + this.AnalysisCode8 + this.AnalysisCode9 + this.DebitCredit + this.JournalType + this.TransactionReference;
        }
        public string LineIndicator
        {
            get;
            set;
        }
        /// <summary>
        /// 
        /// </summary>
        public string Ledger
        {
            get;
            set;
        }
        /// <summary>
        /// 
        /// </summary>
        public string AccountCode
        {
            get;
            set;
        }
        /// <summary>
        /// 
        /// </summary>
        public string AccountingPeriod
        {
            get;
            set;
        }
        /// <summary>
        /// 
        /// </summary>
        public string AllocationMarker
        {
            get;
            set;
        }
        /// <summary>
        /// 
        /// </summary>
        public string AnalysisCode1
        {
            get;
            set;
        }
        /// <summary>
        /// 
        /// </summary>
        public string AnalysisCode10
        {
            get;
            set;
        }
        /// <summary>
        /// 
        /// </summary>
        public string AnalysisCode2
        {
            get;
            set;
        }
        /// <summary>
        /// 
        /// </summary>
        public string AnalysisCode3
        {
            get;
            set;
        }
        /// <summary>
        /// 
        /// </summary>
        public string AnalysisCode4
        {
            get;
            set;
        }
        /// <summary>
        /// 
        /// </summary>
        public string AnalysisCode5
        {
            get;
            set;
        }
        /// <summary>
        /// 
        /// </summary>
        public string AnalysisCode6
        {
            get;
            set;
        }
        /// <summary>
        /// 
        /// </summary>
        public string AnalysisCode7
        {
            get;
            set;
        }
        /// <summary>
        /// 
        /// </summary>
        public string AnalysisCode8
        {
            get;
            set;
        }
        /// <summary>
        /// 
        /// </summary>
        public string AnalysisCode9
        {
            get;
            set;
        }
        /// <summary>
        /// 
        /// </summary>
        public string DebitCredit
        {
            get;
            set;
        }
        /// <summary>
        /// 
        /// </summary>
        public string JournalType
        {
            get;
            set;
        }
        /// <summary>
        /// 
        /// </summary>
        public string TransactionDate
        {
            get;
            set;
        }
        /// <summary>
        /// 
        /// </summary>
        public string TransactionReference
        {
            get;
            set;
        }
        /// <summary>
        /// 
        /// </summary>
        public string JournalNumber
        {
            get;
            set;
        }
        /// <summary>
        /// 
        /// </summary>
        public string JournalLineNumber
        {
            get;
            set;
        }
    }
}
