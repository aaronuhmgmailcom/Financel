
/*  
 * Module ID：<ExcelAddIn4>   
 * Function：<Ledger Update>   
 * Author：Peter.uhm  (yanb@shinetechchina.com)
 * Modify date：2016.04
 * Modify date：2016.09
 * Version : 2.0.0.2
 */
using System;
using System.Collections.Generic;
using System.Text;
namespace ExcelAddIn4.Common2
{
    public class LedgerUpdate
    {
        /// <summary>
        /// 
        /// </summary>
        /// <returns></returns>
        public override string ToString()
        {
            return Ledger + AccountCode + AccountingPeriod + this.AllocationMarker + this.AnalysisCode1 + this.AnalysisCode10 + this.AnalysisCode2 + this.AnalysisCode3 + this.AnalysisCode4 + this.AnalysisCode5 + this.AnalysisCode6 + this.AnalysisCode7 + this.AnalysisCode8 + this.AnalysisCode9 + this.Base2ReportingAmount + this.BaseAmount + this.CurrencyCode + this.DebitCredit + this.Description + this.JournalSource + this.JournalType + this.TransactionAmount + this.TransactionDate + this.DueDate + this.TransactionReference + this.Value4Amount;
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
        public string Base2ReportingAmount
        {
            get;
            set;
        }
        /// <summary>
        /// 
        /// </summary>
        public string BaseAmount
        {
            get;
            set;
        }
        /// <summary>
        /// 
        /// </summary>
        public string CurrencyCode
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
        public string Description
        {
            get;
            set;
        }
        /// <summary>
        /// 
        /// </summary>
        public string JournalSource
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
        public string TransactionAmount
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
        public string DueDate
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
        public string Value4Amount
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
        public Messages Messages
        {
            get;
            set;
        }
    }
    public class Actions
    {
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
        public string ControlTotal
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
        public string Description
        {
            get;
            set;
        }
        /// <summary>
        /// 
        /// </summary>
        public string GeneralDescription1
        {
            get;
            set;
        }
        /// <summary>
        /// 
        /// </summary>
        public string GeneralDescription2
        {
            get;
            set;
        }
        /// <summary>
        /// 
        /// </summary>
        public string GeneralDescription3
        {
            get;
            set;
        }
        /// <summary>
        /// 
        /// </summary>
        public string GeneralDescription4
        {
            get;
            set;
        }
        /// <summary>
        /// 
        /// </summary>
        public string GeneralDescription5
        {
            get;
            set;
        }
        /// <summary>
        /// 
        /// </summary>
        public string GeneralDescription6
        {
            get;
            set;
        }
        /// <summary>
        /// 
        /// </summary>
        public string GeneralDescription7
        {
            get;
            set;
        }
        /// <summary>
        /// 
        /// </summary>
        public string GeneralDescription8
        {
            get;
            set;
        }
        /// <summary>
        /// 
        /// </summary>
        public string GeneralDescription9
        {
            get;
            set;
        }
        /// <summary>
        /// 
        /// </summary>
        public string GeneralDescription10
        {
            get;
            set;
        }
        /// <summary>
        /// 
        /// </summary>
        public string GeneralDescription11
        {
            get;
            set;
        }
        /// <summary>
        /// 
        /// </summary>
        public string GeneralDescription12
        {
            get;
            set;
        }
        /// <summary>
        /// 
        /// </summary>
        public string GeneralDescription13
        {
            get;
            set;
        }
        /// <summary>
        /// 
        /// </summary>
        public string GeneralDescription14
        {
            get;
            set;
        }
        /// <summary>
        /// 
        /// </summary>
        public string GeneralDescription15
        {
            get;
            set;
        }
        /// <summary>
        /// 
        /// </summary>
        public string GeneralDescription16
        {
            get;
            set;
        }
        /// <summary>
        /// 
        /// </summary>
        public string GeneralDescription17
        {
            get;
            set;
        }
        /// <summary>
        /// 
        /// </summary>
        public string GeneralDescription18
        {
            get;
            set;
        }
        /// <summary>
        /// 
        /// </summary>
        public string GeneralDescription19
        {
            get;
            set;
        }
        /// <summary>
        /// 
        /// </summary>
        public string GeneralDescription20
        {
            get;
            set;
        }
        /// <summary>
        /// 
        /// </summary>
        public string GeneralDescription21
        {
            get;
            set;
        }
        /// <summary>
        /// 
        /// </summary>
        public string GeneralDescription22
        {
            get;
            set;
        }
        /// <summary>
        /// 
        /// </summary>
        public string GeneralDescription23
        {
            get;
            set;
        }
        /// <summary>
        /// 
        /// </summary>
        public string GeneralDescription24
        {
            get;
            set;
        }
        /// <summary>
        /// 
        /// </summary>
        public string GeneralDescription25
        {
            get;
            set;
        }
    }
    /// <summary>
    /// 
    /// </summary>
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
    /// <summary>
    /// 
    /// </summary>
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
