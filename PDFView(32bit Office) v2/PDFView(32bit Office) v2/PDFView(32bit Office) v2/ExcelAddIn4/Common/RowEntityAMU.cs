
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
    public class AccountAllocations
    {
        /// <summary>
        /// 
        /// </summary>
        public ActionAllSettings ActionAllSettings
        {
            get;
            set;
        }
        /// <summary>
        /// 
        /// </summary>
        public FlagOptions FlagOptions
        {
            get;
            set;
        }
        /// <summary>
        /// 
        /// </summary>
        public NewSettings NewSettings
        {
            get;
            set;
        }
        /// <summary>
        /// 
        /// </summary>
        public SelectionCriteria SelectionCriteria
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

    public class SelectionCriteria
    {
        /// <summary>
        /// 
        /// </summary>
        /// <returns></returns>
        public override string ToString()
        {
            return Ledger + AccountCode + AccountingPeriodFrom + AccountingPeriodTo + this.AllocationMarker + this.TransactionAnalysis1From + this.TransactionAnalysis10From + this.TransactionAnalysis2From + this.TransactionAnalysis3From + this.TransactionAnalysis4From + this.TransactionAnalysis5From + this.TransactionAnalysis6From + this.TransactionAnalysis7From + this.TransactionAnalysis8From + this.TransactionAnalysis9From + this.DebitCredit + this.JournalTypeFrom + this.TransactionReferenceFrom;
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
        public string AccountingPeriodFrom
        {
            get;
            set;
        }
        /// <summary>
        /// 
        /// </summary>
        public string AccountingPeriodTo
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
        public string TransactionAnalysis1From
        {
            get;
            set;
        }
        /// <summary>
        /// 
        /// </summary>
        public string TransactionAnalysis1To
        {
            get;
            set;
        }
        /// <summary>
        /// 
        /// </summary>
        public string TransactionAnalysis10From
        {
            get;
            set;
        }
        /// <summary>
        /// 
        /// </summary>
        public string TransactionAnalysis10To
        {
            get;
            set;
        }
        /// <summary>
        /// 
        /// </summary>
        public string TransactionAnalysis2From
        {
            get;
            set;
        }
        /// <summary>
        /// 
        /// </summary>
        public string TransactionAnalysis2To
        {
            get;
            set;
        }
        /// <summary>
        /// 
        /// </summary>
        public string TransactionAnalysis3From
        {
            get;
            set;
        }
        /// <summary>
        /// 
        /// </summary>
        public string TransactionAnalysis3To
        {
            get;
            set;
        }
        /// <summary>
        /// 
        /// </summary>
        public string TransactionAnalysis4From
        {
            get;
            set;
        }
        /// <summary>
        /// 
        /// </summary>
        public string TransactionAnalysis4To
        {
            get;
            set;
        }
        /// <summary>
        /// 
        /// </summary>
        public string TransactionAnalysis5From
        {
            get;
            set;
        }
        /// <summary>
        /// 
        /// </summary>
        public string TransactionAnalysis5To
        {
            get;
            set;
        }
        /// <summary>
        /// 
        /// </summary>
        public string TransactionAnalysis6From
        {
            get;
            set;
        }
        /// <summary>
        /// 
        /// </summary>
        public string TransactionAnalysis6To
        {
            get;
            set;
        }
        /// <summary>
        /// 
        /// </summary>
        public string TransactionAnalysis7From
        {
            get;
            set;
        }
        /// <summary>
        /// 
        /// </summary>
        public string TransactionAnalysis7To
        {
            get;
            set;
        }
        /// <summary>
        /// 
        /// </summary>
        public string TransactionAnalysis8From
        {
            get;
            set;
        }
        /// <summary>
        /// 
        /// </summary>
        public string TransactionAnalysis8To
        {
            get;
            set;
        }
        /// <summary>
        /// 
        /// </summary>
        public string TransactionAnalysis9From
        {
            get;
            set;
        }
        /// <summary>
        /// 
        /// </summary>
        public string TransactionAnalysis9To
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
        public string JournalTypeFrom
        {
            get;
            set;
        }
        /// <summary>
        /// 
        /// </summary>
        public string JournalTypeTo
        {
            get;
            set;
        }
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
        /// <summary>
        /// 
        /// </summary>
        public string TransactionReferenceFrom
        {
            get;
            set;
        }
        /// <summary>
        /// 
        /// </summary>
        public string TransactionReferenceTo
        {
            get;
            set;
        }
        /// <summary>
        /// 
        /// </summary>
        public string JournalNumberFrom
        {
            get;
            set;
        }
        /// <summary>
        /// 
        /// </summary>
        public string JournalNumberTo
        {
            get;
            set;
        }
    }
    public class FlagOptions
    {
        /// <summary>
        /// 
        /// </summary>
        public string AllowClosedOrSuspendedAccountCode
        {
            get;
            set;
        }
        /// <summary>
        /// 
        /// </summary>
        public string OverrideAllocations
        {
            get;
            set;
        }
    }
    public class ActionAllSettings
    {
        /// <summary>
        /// 
        /// </summary>
        public string AllocationMarkerFrom
        {
            get;
            set;
        }
        /// <summary>
        /// 
        /// </summary>
        public string AllocationMarkerTo
        {
            get;
            set;
        }
    }
    public class NewSettings
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

    public class Messages
    {
        /// <summary>
        /// 
        /// </summary>
        public Message Message
        {
            get;
            set;
        }
    }
    /// <summary>
    /// 
    /// </summary>
    public class Message
    {
        /// <summary>
        /// 
        /// </summary>
        public string Exception
        {
            get;
            set;
        }
        /// <summary>
        /// 
        /// </summary>
        public string UserText
        {
            get;
            set;
        }
        /// <summary>
        /// 
        /// </summary>
        public Application Application
        {
            get;
            set;
        }
    }
    /// <summary>
    /// 
    /// </summary>
    public class Application
    {
        /// <summary>
        /// 
        /// </summary>
        public string Component
        {
            get;
            set;
        }
        /// <summary>
        /// 
        /// </summary>
        public string DataItem
        {
            get;
            set;
        }
        /// <summary>
        /// 
        /// </summary>
        public string Driver
        {
            get;
            set;
        }
        /// <summary>
        /// 
        /// </summary>
        public string Item
        {
            get;
            set;
        }
        /// <summary>
        /// 
        /// </summary>
        public string LastMethod
        {
            get;
            set;
        }
        /// <summary>
        /// 
        /// </summary>
        public string Message
        {
            get;
            set;
        }
        /// <summary>
        /// 
        /// </summary>
        public string MessageNumber
        {
            get;
            set;
        }
        /// <summary>
        /// 
        /// </summary>
        public string Method
        {
            get;
            set;
        }
        /// <summary>
        /// 
        /// </summary>
        public string Type
        {
            get;
            set;
        }
        /// <summary>
        /// 
        /// </summary>
        public string Value
        {
            get;
            set;
        }
        /// <summary>
        /// 
        /// </summary>
        public string Version
        {
            get;
            set;
        }
    }
}

//public class TransactionDateRange
//{
//    /// <summary>
//    /// 
//    /// </summary>
//    public string TransactionDateFrom
//    {
//        get;
//        set;
//    }
//    /// <summary>
//    /// 
//    /// </summary>
//    public string TransactionDateTo
//    {
//        get;
//        set;
//    }
//}

//public class AccountRange
//{
//    /// <summary>
//    /// 
//    /// </summary>
//    public string AccountCodeFrom
//    {
//        get;
//        set;
//    }
//    /// <summary>
//    /// 
//    /// </summary>
//    public string AccountCodeTo
//    {
//        get;
//        set;
//    }
//}