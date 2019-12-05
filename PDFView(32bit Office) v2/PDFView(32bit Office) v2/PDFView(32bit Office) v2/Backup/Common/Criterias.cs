/*  
 * Module ID：<ExcelAddIn4>   
 * Function：<Criterias>   
 * Author：Peter.uhm  (yanb@shinetechchina.com)
 * Modify date：2016.04
 * Modify date：2016.09
 * Version : 2.0.0.2
 */
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace ExcelAddIn4.Common
{
    public class Criterias
    {
        /// <summary>
        /// 
        /// </summary>
        public string TemplatePath
        {
            get;
            set;
        }
        /// <summary>
        /// 
        /// </summary>
        public List<Criteria> CriteriaN
        {
            get;
            set;
        }
    }
    public class Criteria
    {
        /// <summary>
        /// 
        /// </summary>
        public List<string> CriteriaName
        {
            get;
            set;
        }
        /// <summary>
        /// 
        /// </summary>
        public List<string> CriteriaValue
        {
            get;
            set;
        }
    }
}
