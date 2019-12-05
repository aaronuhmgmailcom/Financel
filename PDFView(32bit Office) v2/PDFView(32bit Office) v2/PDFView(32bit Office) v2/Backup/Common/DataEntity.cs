
/*  
 * Module ID：<ExcelAddIn4>   
 * Function：<SSC>   
 * Author：Peter.uhm  (yanb@shinetechchina.com)
 * Modify date：2016.04
 * Modify date：2016.09
 * Version : 2.0.0.2
 */
using System;
using System.Collections.Generic;
using System.Text;
namespace ExcelAddIn4
{
    public class SSC
    {
        public ErrorContext ErrorContext
        {
            get;
            set;
        }
        public SunSystemsContext SunSystemsContext
        {
            get;
            set;
        }
        public MethodContext MethodContext
        {
            get;
            set;
        }
        public Payload Payload
        {
            get;
            set;
        }
        public string ErrorMessages
        {
            get;
            set;
        }
    }
    public class ErrorContext
    {
        public string CompatibilityMode
        {
            get;
            set;
        }
        public string ErrorOutput
        {
            get;
            set;
        }
        public string ErrorThreshold
        {
            get;
            set;
        }
    }
    public class User
    {
        public string Name
        {
            get;
            set;
        }
    }
    public class SunSystemsContext
    {
        public string BusinessUnit
        {
            get;
            set;
        }
        public string BudgetCode
        {
            get;
            set;
        }
    }
    public class MethodContext
    {
        public LedgerPostingParameters LedgerPostingParamrters
        {
            get;
            set;
        }
    }
    public class LedgerPostingParameters
    {
        public string AllowBalTran
        {
            get;
            set;
        }
        public string AllowPostToSuspended
        {
            get;
            set;
        }
        public string DefaultPeriod
        {
            get;
            set;
        }
        public string Description
        {
            get;
            set;
        }
        public string JournalType
        {
            get;
            set;
        }
        public string LoadOnly
        {
            get;
            set;
        }
        public string PostProvisional
        {
            get;
            set;
        }
        public string PostToHold
        {
            get;
            set;
        }
        public string PostingType
        {
            get;
            set;
        }
        public string ReportErrorsOnly
        {
            get;
            set;
        }
        public string ReportingAccount
        {
            get;
            set;
        }
        public string SuppressSubstitutedMessages
        {
            get;
            set;
        }
        public string SuspenseAccount
        {
            get;
            set;
        }
        public string TransactionAmountAccount
        {
            get;
            set;
        }
    }
    public class Payload
    {
        public List<Line> Ledger
        {
            get;
            set;
        }
    }
}
