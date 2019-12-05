
/*  
 * Module ID：<ExcelAddIn4>   
 * Function：<SSC>   
 * Author：Peter.uhm  (yanb@shinetechchina.com)
 * Modify date：2015.10
 * Version : 2.0.0.2
 */
using System;
using System.Collections.Generic;
using System.Text;

namespace ExcelAddIn4.Common4
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
        public List<ExcelAddIn4.Common3.AllocationMarkers> Payload
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
