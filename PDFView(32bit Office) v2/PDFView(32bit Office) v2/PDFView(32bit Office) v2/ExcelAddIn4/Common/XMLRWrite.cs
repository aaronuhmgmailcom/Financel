/*  
 * Module ID：<ExcelAddIn4>   
 * Function：<XmlRWrite>   
 * Author：Peter.uhm  (yanb@shinetechchina.com)
 * Modify date：2016.04
 * Modify date：2016.09
 * Version : 2.0.0.2
 */
using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Xml;
using System.Data.SqlClient;
using System.Data;
namespace ExcelAddIn4.Common
{
    /// <summary>
    ///
    /// </summary>
    public class XmlRWrite
    {
        private string _FilePath;
        private SqlConnection conn = null;
        private SqlDataReader rdr = null;
        private SqlCommand cmd;
        private string component;
        private string method;
        private string _xml;
        private string _field;
        /// <summary>
        /// 
        /// </summary>
        public string XML
        {
            set
            {
                _xml = value;
            }
        }
        /// <summary>
        /// 
        /// </summary>
        public string Field
        {
            set
            {
                _field = value;
            }
        }
        /// <summary>
        /// 
        /// </summary>
        public XmlRWrite()
        { }
        /// <summary>
        /// 
        /// </summary>
        /// <param name="con"></param>
        /// <param name="rd"></param>
        /// <param name="cm"></param>
        /// <param name="com"></param>
        /// <param name="me"></param>
        public XmlRWrite(SqlConnection con, SqlDataReader rd, SqlCommand cm, string com, string me)
        {
            conn = con;
            rd = rdr;
            cmd = cm;
            component = com;
            method = me;
            cmd = new SqlCommand("rsTemplateCreateXMLTextProfile_Ins", conn);
            cmd.CommandType = CommandType.StoredProcedure;
        }
        /// <summary>
        /// 
        /// </summary>
        public string FilePath
        {
            set
            {
                _FilePath = value;
            }
        }
        /// <summary>
        /// 
        /// </summary>
        /// <param name="XNL"></param>
        /// <returns></returns>
        private void ReadNodeValue(XmlNode XNL)
        {
            foreach (XmlNode node1 in XNL.ChildNodes)
            {
                if (node1.ChildNodes.Count >= 1 && (node1.ChildNodes[0].Name != "#text"))
                {
                    InsertNode(node1, "");
                    ReadNodeValue(node1);
                }
                else
                {
                    InsertNode(node1, node1.InnerText);
                }
            }
        }
        /// <summary>
        /// 
        /// </summary>
        /// <param name="node"></param>
        /// <param name="value"></param>
        private void InsertNode(XmlNode node, string value)
        {
            cmd.Parameters.Clear();
            cmd.Parameters.Add(new SqlParameter("@Field", node.Name));
            cmd.Parameters.Add(new SqlParameter("@FriendlyName", node.Name));
            cmd.Parameters.Add(new SqlParameter("@Visible", false));
            cmd.Parameters.Add(new SqlParameter("@DefaultValue", value));
            cmd.Parameters.Add(new SqlParameter("@SunComponentName", component));
            cmd.Parameters.Add(new SqlParameter("@SunMethod", method));
            cmd.Parameters.Add(new SqlParameter("@Mandatory", false));
            cmd.Parameters.Add(new SqlParameter("@Separator", ""));
            cmd.Parameters.Add(new SqlParameter("@TextLength", ""));
            cmd.Parameters.Add(new SqlParameter("@trimText", "None"));
            cmd.Parameters.Add(new SqlParameter("@Prefix", ""));
            cmd.Parameters.Add(new SqlParameter("@Suffix", ""));
            cmd.Parameters.Add(new SqlParameter("@RemoveCharacters", ""));
            cmd.Parameters.Add(new SqlParameter("@TextFileName", ""));
            cmd.Parameters.Add(new SqlParameter("@Parent", node.ParentNode.Name));
            cmd.Parameters.Add(new SqlParameter("@Section", "Line"));
            rdr = cmd.ExecuteReader();
            rdr.Close();
        }
        /// <summary>
        /// 
        /// </summary>
        /// <returns></returns>
        private XmlNode XNodeList()
        {
            try
            {
                XmlDocument xmlDoc = new XmlDocument();
                xmlDoc.Load(_FilePath);
                return xmlDoc.SelectSingleNode("SSC");
            }
            catch
            {
                throw new Exception("Xml file error!");
            }
        }
        /// <summary>
        /// 
        /// </summary>
        /// <param name="s"></param>
        /// <returns></returns>
        private XmlNode XNodeList(string s)
        {
            try
            {
                XmlDocument xmlDoc = new XmlDocument();
                xmlDoc.LoadXml(_xml);
                return xmlDoc.SelectSingleNode(_field);
            }
            catch
            {
                throw new Exception("Xml error!");
            }
        }
        ///// <summary>
        ///// 
        ///// </summary>
        ///// <param name="XNL"></param>
        ///// <returns></returns>
        //private void ReadNodeValueIntoString(XmlNode XNL)
        //{
        //    foreach (XmlNode node1 in XNL.ChildNodes)
        //    {
        //        if (node1.ChildNodes.Count >= 1 && (node1.ChildNodes[0].Name != "#text"))
        //        {
        //            _sSCstring += node1.Name + ",";
        //            ReadNodeValueIntoString(node1);
        //        }
        //        else
        //        {
        //            _sSCstring += node1.Name + ",";
        //        }
        //    }
        //}
        ///// <summary>
        ///// 
        ///// </summary>
        ///// <returns></returns>
        //public void SaveNodeValueIntoStr()
        //{
        //    XmlNode XNL = XNodeList("");
        //    ReadNodeValueIntoString(XNL);
        //}
        /// <summary>
        /// 
        /// </summary>
        /// <returns></returns>
        public void SaveNodeValue()
        {
            try
            {
                XmlNode XNL = XNodeList();
                ReadNodeValue(XNL);
            }
            finally
            {
                if (conn != null)
                {
                    conn.Close();
                }
                if (rdr != null)
                {
                    rdr.Close();
                }
            }
        }
    }
}
