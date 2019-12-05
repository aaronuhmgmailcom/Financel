
/*  
 * Module ID：<ExcelAddIn4>   
 * Function：<AllocationMarkers>   
 * Author：Peter.uhm  (yanb@shinetechchina.com)
 * Modify date：2015.10
 * Version : 2.0.0.2
 */
using System;
using System.Collections.Generic;
using System.Text;

namespace ExcelAddIn4.Common3
{
    public class AllocationMarkers
    {
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
        public string Ledger
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
        public string AnalysisCode10
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
        public string JournalType
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
        public AccountRange AccountRange
        {
            get;
            set;
        }
        /// <summary>
        /// 
        /// </summary>
        public Actions Actions
        {
            get;
            set;
        }
        /// <summary>
        /// 
        /// </summary>
        public TransactionDateRange TransactionDateRange
        {
            get;
            set;
        }
        /// <summary>
        /// 
        /// </summary>
        public Messages Messages
        {
            get;
            set;
        }
    }
    /// <summary>
    /// 
    /// </summary>
    public class Actions
    {
        /// <summary>
        /// 
        /// </summary>
        public string AllocationMarker
        {
            get;
            set;
        }
    }
    /// <summary>
    /// 
    /// </summary>
    public class TransactionDateRange
    {
        /// <summary>
        /// 
        /// </summary>
        public string TransactionDateFrom
        {
            get;
            set;
        }
        /// <summary>
        /// 
        /// </summary>
        public string TransactionDateTo
        {
            get;
            set;
        }
    }

    public class AccountRange
    {
        /// <summary>
        /// 
        /// </summary>
        public string AccountCodeFrom
        {
            get;
            set;
        }
        /// <summary>
        /// 
        /// </summary>
        public string AccountCodeTo
        {
            get;
            set;
        }
    }
}

