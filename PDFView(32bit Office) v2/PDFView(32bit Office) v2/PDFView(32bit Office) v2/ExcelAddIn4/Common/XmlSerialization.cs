using System;
using System.Collections.Generic;
using System.Text;
using System.Xml.Serialization;
using System.IO;
using System.Xml;
using System.Text.RegularExpressions;
using System.Data;

namespace ExcelAddIn4
{
    public class XmlSerialization<T> where T : class
    {
        /// <summary>
        /// 
        /// </summary>
        /// <param name="instance"></param>
        /// <returns></returns>
        public static string Serialize(T instance)
        {
            XmlSerializer serializer = new XmlSerializer(typeof(T));
            using (MemoryStream ms = new MemoryStream())
            {
                serializer.Serialize(ms, instance);
                ms.Seek(0, SeekOrigin.Begin);
                return new StreamReader(ms, Encoding.UTF8).ReadToEnd();
            }
        }
        /// <summary>
        /// 
        /// </summary>
        /// <param name="xml"></param>
        /// <returns></returns>
        public static T DeSerialize(string xml)
        {
            XmlSerializer serializer = new XmlSerializer(typeof(T));
            using (MemoryStream ms = new MemoryStream())
            {
                StreamWriter sw = new StreamWriter(ms, Encoding.UTF8);
                sw.Write(xml);
                sw.Flush();
                ms.Seek(0, SeekOrigin.Begin);
                return (T)serializer.Deserialize(ms);
            }
        }
        /// <summary>
        /// 
        /// </summary>
        /// <param name="xml"></param>
        /// <returns></returns>
        public static string AvoidXmlns(string xml)
        {
            Regex reg = new Regex("xmlns\\s*\\S*\\s*=\\s*\"[^\"]*\"", RegexOptions.IgnoreCase);
            return reg.Replace(xml, "");
        }
    }
    public class XmlOperator
    {
        public static string JoinPath(params string[] nodeName)
        {
            return string.Join(".", nodeName);
        }
        private XmlDocument objXmlDoc = new XmlDocument();
        public XmlNode RootNode
        {
            get
            {
                return objXmlDoc.DocumentElement;
            }
        }
        public XmlDocument XmlDoc
        {
            get { return objXmlDoc; }
        }
        /// <summary>
        /// Load an XmlFile by it's fullname
        /// </summary>
        /// <param name="fileName">Fullname of the xml file</param>
        public XmlOperator(string xml)
        {
            try
            {
                objXmlDoc.LoadXml(xml);
            }
            catch (System.Exception ex)
            {
                throw ex;
            }
        }
        public XmlOperator(XmlDocument xmlDoc)
        {
            objXmlDoc = xmlDoc;
        }
        public XmlOperator()
        {
            objXmlDoc = new XmlDocument();
        }
        public XmlNode CreateRootNode(string rootNode)
        {
            objXmlDoc = new XmlDocument();
            XmlNode node = objXmlDoc.CreateNode(XmlNodeType.Element, rootNode, "");
            XmlDeclaration declaration = objXmlDoc.CreateXmlDeclaration("1.0", "utf-8", "");
            objXmlDoc.AppendChild(declaration);
            objXmlDoc.AppendChild(node);
            return node;

        }
        /// <summary>
        /// Read a node from xml file, return a DataSet
        /// </summary>
        /// <param name="XmlPathNode">an xpath string used to search an XmlNode</param>
        /// <returns>DataSet</returns>
        public DataSet GetData(string XmlPathNode)
        {
            DataSet ds = new DataSet();
            StringReader read = new StringReader(objXmlDoc.SelectSingleNode(XmlPathNode).OuterXml);
            ds.ReadXml(read);
            return ds;
        }
        /// <summary>
        /// Replace the inner text of an XmlNode
        /// </summary>
        /// <param name="XmlPathNode">an xpath string used to search an XmlNode</param>
        /// <param name="Content"></param>
        public XmlNode SetInnerText(string XmlPathNode, string Content)
        {
            XmlNode node = objXmlDoc.SelectSingleNode(this.xpath(XmlPathNode));
            if (!string.IsNullOrEmpty(Content))
            {
                XmlCDataSection cdata = objXmlDoc.CreateCDataSection(Content);

                node.AppendChild(cdata);
            }
            return node;
        }
        /// <summary>
        /// 
        /// </summary>
        /// <param name="XmlPathNode"></param>
        /// <param name="Content"></param>
        /// <returns></returns>
        public XmlNode SetInnerText(XmlNode XmlPathNode, string Content)
        {
            XmlNode node = XmlPathNode;
            if (!string.IsNullOrEmpty(Content))
            {
                XmlCDataSection cdata = objXmlDoc.CreateCDataSection(Content);

                node.AppendChild(cdata);
            }
            return node;
        }
        /// <summary>
        /// Delete an XmlNode
        /// </summary>
        /// <param name="Node">an xpath string used to search an XmlNode</param>
        public XmlNode Delete(string Node)
        {
            XmlNode node = objXmlDoc.SelectSingleNode(this.xpath(Node));
            node.ParentNode.RemoveChild(node);
            return node;
        }
        /// <summary>
        /// Insert a node and one of it's child node
        /// </summary>
        /// <param name="MainNode">an xpath string used to search an parent node</param>
        /// <param name="ChildNode">Name of the child node</param>
        /// <param name="Element">Name of the child node's element</param>
        /// <param name="Content">Content of the child node's element</param>
        public XmlNode InsertNode(string MainNode, string ChildNode, string Element, string Content)
        {
            XmlNode objRootNode = objXmlDoc.SelectSingleNode(this.xpath(MainNode));
            XmlElement objChildNode = objXmlDoc.CreateElement(ChildNode);
            objRootNode.AppendChild(objChildNode);
            XmlElement objElement = objXmlDoc.CreateElement(Element);
            if (!string.IsNullOrEmpty(Content))
            {
                XmlCDataSection cdata = objXmlDoc.CreateCDataSection(Content);
                objElement.AppendChild(cdata);
            }
            objChildNode.AppendChild(objElement);
            return objChildNode;
        }
        /// <summary>
        /// 
        /// </summary>
        /// <param name="MainNode"></param>
        /// <param name="ChildNode"></param>
        /// <param name="Element"></param>
        /// <param name="Content"></param>
        /// <returns></returns>
        public XmlNode InsertNode(XmlNode MainNode, string ChildNode, string Element, string Content)
        {
            XmlNode objRootNode = MainNode;
            XmlElement objChildNode = objXmlDoc.CreateElement(ChildNode);
            objRootNode.AppendChild(objChildNode);
            XmlElement objElement = objXmlDoc.CreateElement(Element);
            if (!string.IsNullOrEmpty(Content))
            {
                XmlCDataSection cdata = objXmlDoc.CreateCDataSection(Content);
                objElement.AppendChild(cdata);
            }
            objChildNode.AppendChild(objElement);
            return objChildNode;
        }
        /// <summary>
        /// Insert a node, with an attribute
        /// </summary>
        /// <param name="MainNode">an xpath string used to search parent node</param>
        /// <param name="Element">Name of the node</param>
        /// <param name="Attrib">Name of the attribute</param>
        /// <param name="AttribContent">Value of the attribute</param>
        /// <param name="Content">Inner text of the node</param>
        public XmlNode InsertElement(string MainNode, string Element, string Attrib, string AttribContent, string Content)
        {
            XmlNode objNode = objXmlDoc.SelectSingleNode(this.xpath(MainNode));
            XmlElement objElement = objXmlDoc.CreateElement(Element);
            objElement.SetAttribute(Attrib, AttribContent);
            if (!string.IsNullOrEmpty(Content))
            {
                XmlCDataSection cdata = objXmlDoc.CreateCDataSection(Content);
                objElement.AppendChild(cdata);
            }
            objNode.AppendChild(objElement);
            return objElement;
        }
        /// <summary>
        /// 
        /// </summary>
        /// <param name="MainNode"></param>
        /// <param name="Element"></param>
        /// <param name="Attrib"></param>
        /// <param name="AttribContent"></param>
        /// <param name="Content"></param>
        /// <returns></returns>
        public XmlNode InsertElement(XmlNode MainNode, string Element, string Attrib, string AttribContent, string Content)
        {
            XmlNode objNode = MainNode;
            XmlElement objElement = objXmlDoc.CreateElement(Element);
            objElement.SetAttribute(Attrib, AttribContent);
            if (!string.IsNullOrEmpty(Content))
            {
                XmlCDataSection cdata = objXmlDoc.CreateCDataSection(Content);
                objElement.AppendChild(cdata);
            }
            objNode.AppendChild(objElement);
            return objElement;
        }
        /// <summary>
        /// Insert a node, without any attribute
        /// </summary>
        /// <param name="MainNode">an xpath string used to search parent node</param>
        /// <param name="Element">Name of the node</param>
        /// <param name="Content">Inner text of the node</param>
        public XmlNode InsertElement(string MainNode, string Element, string Content)
        {
            XmlNode objNode = objXmlDoc.SelectSingleNode(this.xpath(MainNode));
            XmlElement objElement = objXmlDoc.CreateElement(Element);
            if (!string.IsNullOrEmpty(Content))
            {
                XmlCDataSection cdata = objXmlDoc.CreateCDataSection(Content);
                objElement.AppendChild(cdata);
            }
            objNode.AppendChild(objElement);
            return objElement;
        }
        /// <summary>
        /// 
        /// </summary>
        /// <param name="MainNode"></param>
        /// <param name="Element"></param>
        /// <param name="Content"></param>
        /// <returns></returns>
        public XmlNode InsertElement(XmlNode MainNode, string Element, string Content)
        {
            XmlNode objNode = MainNode;
            XmlElement objElement = objXmlDoc.CreateElement(Element);
            if (!string.IsNullOrEmpty(Content))
            {
                XmlCDataSection cdata = objXmlDoc.CreateCDataSection(Content);
                objElement.AppendChild(cdata);
            }
            objNode.AppendChild(objElement);
            return objElement;
        }
        /// <summary>
        /// Select an XmlNode using a standard xpath string
        /// </summary>
        /// <param name="xpath">a standard xpath string</param>
        /// <returns>XmlNode</returns>
        public XmlNode SelectSingleNode(string xpath)
        {
            return objXmlDoc.SelectSingleNode(this.xpath(xpath));
        }
        /// <summary>
        /// Select an array of XmlNode using a standard xpath string
        /// </summary>
        /// <param name="xpath">a standard xpath string</param>
        /// <returns>XmlNodeList</returns>
        public XmlNodeList SelectNodes(string xpath)
        {
            return objXmlDoc.SelectNodes(this.xpath(xpath));
        }
        /// <summary>
        /// Get the value of an attribute from an XmlNode
        /// </summary>
        /// <param name="node">an XmlNode whose attribute is to be readed</param>
        /// <param name="attributeName">name of the attribute</param>
        /// <returns>value of the attribute</returns>
        public string GetNodeAttribute(XmlNode node, string attributeName)
        {
            try
            {
                return node.Attributes[attributeName].Value;
            }
            catch (System.Exception ex)
            {
                throw ex;
            }
        }
        /// <summary>
        /// Get the value of an XmlNode by it's node path
        /// </summary>
        /// <param name="nodePath">The path of the XmlNode, it should be write from the root and split every childnode by "."</param>
        /// <param name="attributeName">name of the attribute</param>
        /// <returns>value of the attribute</returns>
        public string GetNodeAttribute(string nodePath, string attributeName)
        {
            XmlNode node = this.GetSelectNode(this.xpath(nodePath));
            if (node == null)
            {
                throw new System.Exception(string.Format("Could not find the XmlNode with the path \"{0}\"", nodePath));
            }
            return this.GetNodeAttribute(node, attributeName);
        }
        /// <summary>
        /// Get the value of an attribute from a childnode defined by name
        /// </summary>
        /// <param name="node">current XmlNode</param>
        /// <param name="childNodeName">name of the childnode</param>
        /// <param name="attrName">attribute's name of the childnode</param>
        /// <returns>value of the attribute</returns>
        public string GetNodeAttribute(XmlNode node, string childNodeName, string attributeName)
        {
            string xpath = string.Format("./{0}", childNodeName);
            XmlNode child = node.SelectSingleNode(xpath);
            if (child == node)
            {
                string error = string.Format("Could not find the child node named '{0}' from current XmlNode", childNodeName);
                throw new System.Exception(error);
            }
            return this.GetNodeAttribute(child, attributeName);
        }
        /// <summary>
        /// Save the value of an attribute
        /// </summary>
        /// <param name="node">current XmlNode</param>
        /// <param name="attributeName">name of the attribute</param>
        /// <param name="value">value of the attribute</param>
        public XmlNode SetNodeAttribute(XmlNode node, string attributeName, string value)
        {
            if (node.Attributes[attributeName] == null)
            {
                XmlAttribute attr = objXmlDoc.CreateAttribute(attributeName);
                node.Attributes.Append(attr);
            }
            node.Attributes[attributeName].Value = value;
            return node;
        }
        /// <summary>
        /// Save the value of an attribute by the path of an XmlNode
        /// </summary>
        /// <param name="nodePath">The path of the XmlNode, it should be write from the root and split every childnode by "."</param>
        /// <param name="attributeName">name of the attribute</param>
        /// <param name="value">value of the attribute</param>
        public XmlNode SetNodeAttribute(string nodePath, string attributeName, string value)
        {
            XmlNode node = this.GetSelectNode(this.xpath(nodePath));
            if (node == null)
            {
                throw new System.Exception(string.Format("Could not find the XmlNode with the path \"{0}\"", nodePath));
            }
            this.SetNodeAttribute(node, attributeName, value);
            return node;
        }
        /// <summary>
        /// Save the value of an XmlNode
        /// </summary>
        /// <param name="node">current XmlNode</param>
        /// <param name="value">value of the XmlNode</param>
        public XmlNode SetNode(XmlNode node, string value)
        {
            node.Value = value;
            return node;
        }
        /// <summary>
        /// Save the value of an XmlNode by it's node path
        /// </summary>
        /// <param name="nodePath">The path of the XmlNode, it should be write from the root and split every childnode by "."</param>
        /// <param name="value">value of the XmlNode</param>
        public XmlNode SetNode(string nodePath, string value)
        {
            XmlNode node = this.GetSelectNode(this.xpath(nodePath));
            if (node == null)
            {
                throw new System.Exception(string.Format("Could not find the XmlNode with the path \"{0}\"", nodePath));
            }
            this.SetNode(node, value);
            return node;
        }
        /// <summary>
        /// Get an XmlNode using it's node path
        /// </summary>
        /// <param name="nodePath">The path of the XmlNode, it should be write from the root and split every childnode by "."</param>
        /// <returns>XmlNode</returns>
        public XmlNode GetSelectNode(string nodePath)
        {
            string xpath = this.xpath(nodePath);
            return objXmlDoc.SelectSingleNode(xpath);
        }
        /// <summary>
        /// Get an XmlNode using it's node path with condition
        /// </summary>
        /// <param name="nodePath">The path of the XmlNode, it should be write from the root and split every childnode by "."</param>
        /// <param name="attributeName">name of the attribute</param>
        /// <param name="value">value of the attribute</param>
        /// <returns>XmlNode</returns>
        public XmlNode GetSelectNode(string nodePath, string attributeName, string value)
        {
            string xpath = this.xpath(nodePath, attributeName, value);
            return objXmlDoc.SelectSingleNode(xpath);
        }
        /// <summary>
        /// Get a child node of an XmlNode by name
        /// </summary>
        /// <param name="node">current XmlNode</param>
        /// <param name="childName">name of the child node</param>
        /// <returns>a XmlNode, it will return the first one if there's more than one result</returns>
        public XmlNode GetSelectNode(XmlNode node, string childName)
        {
            string xpath = string.Format("./{0}", childName);
            return node.SelectSingleNode(xpath);
        }
        /// <summary>
        /// Get a child node of an XmlNode by name with condition
        /// </summary>
        /// <param name="node">current XmlNode</param>
        /// <param name="childName">name of the child node</param>
        /// <param name="attributeName">name of the attribute of the child node</param>
        /// <param name="value">value of the attribute</param>
        /// <returns>a XmlNode, it will return the first one if there's more than one result</returns>
        public XmlNode GetSelectNode(XmlNode node, string childName, string attributeName, string value)
        {
            string xpath = string.Format("./{0}[@{1}='{2}']", childName, attributeName, value);
            return node.SelectSingleNode(xpath);
        }
        /// <summary>
        /// Get all the child nodes of an XmlNode
        /// </summary>
        /// <param name="node">current XmlNode</param>
        /// <returns>an array contains all the child nodes</returns>
        public XmlNode[] GetChildsNode(XmlNode node)
        {
            List<XmlNode> nodeList = new List<XmlNode>();
            foreach (XmlNode n in node.ChildNodes)
            {
                nodeList.Add(n);
            }
            return nodeList.ToArray();
        }
        /// <summary>
        /// Convert the node path to a standard xpath
        /// </summary>
        /// <param name="nodePath">a string of the node path</param>
        /// <returns>a standard xpath string</returns>
        private string xpath(string nodePath)
        {
            Regex format = new Regex(@"^[^\.]+(\.[^\.]+)*$");
            if (!format.IsMatch(nodePath))
            {
                throw new FormatException(string.Format("The node path \"{0}\" does not meet the required format", nodePath));
            }
            return nodePath.Replace(".", "/").Insert(0, "/");
        }
        /// <summary>
        /// Convert the node path to a standard xpath, with condition
        /// </summary>
        /// <param name="nodePath">a string of the node path</param>
        /// <param name="attrName">attribute's name of the target node</param>
        /// <param name="value">value of the attribute</param>
        /// <returns>a standard xpath string</returns>
        private string xpath(string nodePath, string attrName, string value)
        {
            string xpath = this.xpath(nodePath);
            return string.Format("{0}[@{1}='{2}']", xpath, attrName, value);
        }
    }

    public class XmlNodeName
    {
        /// <summary>
        /// 
        /// </summary>
        public const string Message = "Message";
        /// <summary>
        /// 
        /// </summary>
        public const string Table = "Table";
        /// <summary>
        /// 
        /// </summary>
        public const string Field = "Field";
    }

    public class XmlAttributeName
    {
        public const string Name = "Name";
    }

    public class XMLCollection
    {
        public const string xmlGenFolderStructure = @"<?xml version='1.0' encoding='utf-8' ?><Nodes></Nodes>";

        public const string xml = @"<?xml version='1.0' encoding='utf-8' ?>
<Nodes>
    <Ledger FieldGroup ='Payload' SunField='Ledger' FriendlyName='Ledger' Output='True' Input='True'></Ledger>
    <AccountCode FieldGroup ='Payload' SunField='AccountCode' FriendlyName='Account' Output='True' Input='True'></AccountCode>
    <AccountingPeriod FieldGroup ='Payload' SunField='AccountingPeriod' FriendlyName='Period' Output='True' Input='True'></AccountingPeriod>
    <TransactionDate FieldGroup ='Payload' SunField='TransactionDate' FriendlyName='Trans Date' Output='True' Input='True'></TransactionDate>
    <DueDate FieldGroup ='Payload' SunField='DueDate' FriendlyName='Due Date' Output='True' Input='True'></DueDate>
    <JournalType FieldGroup ='Payload' SunField='JournalType' FriendlyName='Jrnl Type' Output='True' Input='True'></JournalType>
    <JournalSource FieldGroup ='Payload' SunField='JournalSource' FriendlyName='Jrnl Source' Output='True' Input='True'></JournalSource>
    <TransactionReference FieldGroup ='Payload' SunField='TransactionReference' FriendlyName='Trans Ref' Output='True' Input='True'></TransactionReference>
    <Description FieldGroup ='Payload' SunField='Description' FriendlyName='Description' Output='True' Input='True'></Description>
    <AllocationMarker FieldGroup ='Payload' SunField='AllocationMarker' FriendlyName='Alloctn Marker' Output='True' Input='True'></AllocationMarker>
    <AnalysisCode1 FieldGroup ='Payload' SunField='AnalysisCode1' FriendlyName='LA1' Output='True' Input='True'></AnalysisCode1>
    <AnalysisCode2 FieldGroup ='Payload' SunField='AnalysisCode2' FriendlyName='LA2' Output='True' Input='True'></AnalysisCode2>
    <AnalysisCode3 FieldGroup ='Payload' SunField='AnalysisCode3' FriendlyName='LA3' Output='True' Input='True'></AnalysisCode3>
    <AnalysisCode4 FieldGroup ='Payload' SunField='AnalysisCode4' FriendlyName='LA4' Output='True' Input='True'></AnalysisCode4>
    <AnalysisCode5 FieldGroup ='Payload' SunField='AnalysisCode5' FriendlyName='LA5' Output='True' Input='True'></AnalysisCode5>
    <AnalysisCode6 FieldGroup ='Payload' SunField='AnalysisCode6' FriendlyName='LA6' Output='True' Input='True'></AnalysisCode6>
    <AnalysisCode7 FieldGroup ='Payload' SunField='AnalysisCode7' FriendlyName='LA7' Output='True' Input='True'></AnalysisCode7>
    <AnalysisCode8 FieldGroup ='Payload' SunField='AnalysisCode8' FriendlyName='LA8' Output='True' Input='True'></AnalysisCode8>
    <AnalysisCode9 FieldGroup ='Payload' SunField='AnalysisCode9' FriendlyName='LA9' Output='True' Input='True'></AnalysisCode9>
    <AnalysisCode10 FieldGroup ='Payload' SunField='AnalysisCode10' FriendlyName='LA10' Output='True' Input='True'></AnalysisCode10>
    <TransactionAmount FieldGroup ='Payload' SunField='TransactionAmount' FriendlyName='Trans Amount' Output='True' Input='True'></TransactionAmount>
    <CurrencyCode FieldGroup ='Payload' SunField='CurrencyCode' FriendlyName='Currency' Output='True' Input='True'></CurrencyCode>
    <BaseAmount FieldGroup ='Payload' SunField='BaseAmount' FriendlyName='Base Amount' Output='True' Input='True'></BaseAmount>
    <Base2ReportingAmount FieldGroup ='Payload' SunField='Base2ReportingAmount' FriendlyName='2nd Base' Output='True' Input='True'></Base2ReportingAmount>
    <Value4Amount FieldGroup ='Payload' SunField='Value4Amount' FriendlyName='4th Amount' Output='True' Input='True'></Value4Amount>
</Nodes>";

        public const string xmlGen = @"<?xml version='1.0' encoding='utf-8' ?>
<Nodes>
    <GeneralDescription1 FieldGroup ='Payload' SunField='GeneralDescription1' FriendlyName='GeneralDescription1' Output='False' XML_Query=''>
  </GeneralDescription1>
  <GeneralDescription2 FieldGroup ='Payload' SunField='GeneralDescription2' FriendlyName='GeneralDescription2' Output='False' XML_Query=''>
  </GeneralDescription2>
  <GeneralDescription3 FieldGroup ='Payload' SunField='GeneralDescription3' FriendlyName='GeneralDescription3' Output='False' XML_Query=''>
  </GeneralDescription3>
  <GeneralDescription4 FieldGroup ='Payload' SunField='GeneralDescription4' FriendlyName='GeneralDescription4' Output='False' XML_Query=''>
  </GeneralDescription4>
  <GeneralDescription5 FieldGroup ='Payload' SunField='GeneralDescription5' FriendlyName='GeneralDescription5' Output='False' XML_Query=''>
  </GeneralDescription5>
  <GeneralDescription6 FieldGroup ='Payload' SunField='GeneralDescription6' FriendlyName='GeneralDescription6' Output='False' XML_Query=''>
  </GeneralDescription6>
  <GeneralDescription7 FieldGroup ='Payload' SunField='GeneralDescription7' FriendlyName='GeneralDescription7' Output='False' XML_Query=''>
  </GeneralDescription7>
  <GeneralDescription8 FieldGroup ='Payload' SunField='GeneralDescription8' FriendlyName='GeneralDescription8' Output='False' XML_Query=''>
  </GeneralDescription8>
  <GeneralDescription9 FieldGroup ='Payload' SunField='GeneralDescription9' FriendlyName='GeneralDescription9' Output='False' XML_Query=''>
  </GeneralDescription9>
  <GeneralDescription10 FieldGroup ='Payload' SunField='GeneralDescription10' FriendlyName='GeneralDescription10' Output='False' XML_Query=''>
  </GeneralDescription10>
  <GeneralDescription11 FieldGroup ='Payload' SunField='GeneralDescription11' FriendlyName='GeneralDescription11' Output='False' XML_Query=''>
  </GeneralDescription11>
  <GeneralDescription12 FieldGroup ='Payload' SunField='GeneralDescription12' FriendlyName='GeneralDescription12' Output='False' XML_Query=''>
  </GeneralDescription12>
  <GeneralDescription13 FieldGroup ='Payload' SunField='GeneralDescription13' FriendlyName='GeneralDescription13' Output='False' XML_Query=''>
  </GeneralDescription13>
  <GeneralDescription14 FieldGroup ='Payload' SunField='GeneralDescription14' FriendlyName='GeneralDescription14' Output='False' XML_Query=''>
  </GeneralDescription14>
  <GeneralDescription15 FieldGroup ='Payload' SunField='GeneralDescription15' FriendlyName='GeneralDescription15' Output='False' XML_Query=''>
  </GeneralDescription15>
  <GeneralDescription16 FieldGroup ='Payload' SunField='GeneralDescription16' FriendlyName='GeneralDescription16' Output='False' XML_Query=''>
  </GeneralDescription16>
  <GeneralDescription17 FieldGroup ='Payload' SunField='GeneralDescription17' FriendlyName='GeneralDescription17' Output='False' XML_Query=''>
  </GeneralDescription17>
  <GeneralDescription18 FieldGroup ='Payload' SunField='GeneralDescription18' FriendlyName='GeneralDescription18' Output='False' XML_Query=''>
  </GeneralDescription18>
  <GeneralDescription19 FieldGroup ='Payload' SunField='GeneralDescription19' FriendlyName='GeneralDescription19' Output='False' XML_Query=''>
  </GeneralDescription19>
  <GeneralDescription20 FieldGroup ='Payload' SunField='GeneralDescription20' FriendlyName='GeneralDescription20' Output='False' XML_Query=''>
  </GeneralDescription20>
  <GeneralDescription21 FieldGroup ='Payload' SunField='GeneralDescription21' FriendlyName='GeneralDescription21' Output='False' XML_Query=''>
  </GeneralDescription21>
  <GeneralDescription22 FieldGroup ='Payload' SunField='GeneralDescription22' FriendlyName='GeneralDescription22' Output='False' XML_Query=''>
  </GeneralDescription22>
  <GeneralDescription23 FieldGroup ='Payload' SunField='GeneralDescription23' FriendlyName='GeneralDescription23' Output='False' XML_Query=''>
  </GeneralDescription23>
  <GeneralDescription24 FieldGroup ='Payload' SunField='GeneralDescription24' FriendlyName='GeneralDescription24' Output='False' XML_Query=''>
  </GeneralDescription24>
  <GeneralDescription25 FieldGroup ='Payload' SunField='GeneralDescription25' FriendlyName='GeneralDescription25' Output='False' XML_Query=''>
  </GeneralDescription25>
</Nodes>";
    }
}
