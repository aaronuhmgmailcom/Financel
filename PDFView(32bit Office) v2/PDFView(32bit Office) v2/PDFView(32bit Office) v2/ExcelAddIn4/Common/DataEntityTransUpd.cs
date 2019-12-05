
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
namespace ExcelAddIn4.Common2
{
    public class SSC
    {
        public ErrorContext ErrorContext
        {
            get;
            set;
        }
        public User User
        {
            get;
            set;
        }
        public SunSystemsContext SunSystemsContext
        {
            get;
            set;
        }
        public List<LedgerUpdate> Payload
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
}
