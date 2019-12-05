using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.IO;
using System.Windows.Forms;
using System.Diagnostics;
using System.Xml;
namespace ExcelAddIn4
{
    public class FolderFileXMLHelper
    {
        /// <summary>
        /// 
        /// </summary>
        internal DirectoryInfo di
        {
            get { return new DirectoryInfo(this.RootPath); }
        }
        /// <summary>
        /// 
        /// </summary>
        public string filesString
        {
            get;
            set;
        }
        /// <summary>
        /// 
        /// </summary>
        public string RootPath
        {
            get;
            set;
        }
        /// <summary>
        /// 
        /// </summary>
        /// <param name="path"></param>
        public FolderFileXMLHelper(string path)
        {
            RootPath = path;
            filesString = "";
        }
        /// <summary>
        /// 
        /// </summary>
        /// <returns></returns>
        public string GetFilesCount()
        {
            if (!string.IsNullOrEmpty(RootPath))
            {
                DirectoryInfo[] fldrs = di.GetDirectories("*.*");
                filesString = "";
                for (int i = fldrs.Length; i > 0; i--)
                {
                    DirectoryInfo d = fldrs[i - 1];
                    if (Finance_Tools.IsAddedToGallery(d))
                    {
                        try
                        {
                            FileInfo[] myfile = d.GetFiles();
                            for (int j = myfile.Length; j > 0; j--)
                            {
                                FileInfo f = myfile[j - 1];
                                if (f.Extension == ".xls" || f.Extension == ".xlsx" || f.Extension == ".xlsm" || f.Extension == ".pdf")
                                {
                                    string fileName = Path.GetFileNameWithoutExtension(f.Name);
                                    if (!fileName.Contains("~"))
                                        filesString += fileName;
                                }
                            }
                        }
                        catch
                        { }
                    }
                }
            }
            return filesString;
        }
        /// <summary>
        /// Initialize File Structure when initia folder path or file number changes
        /// </summary>
        public void InitializeFileXMLStructure()
        {
            if (!string.IsNullOrEmpty(RootPath))
            {
                string[] sArray = Finance_Tools.FileIds.Split(new char[3] { ',', ',', ',' });
                DirectoryInfo[] fldrs = di.GetDirectories("*.*");
                XmlDocument xdoc = new XmlDocument();
                xdoc.Load(Finance_Tools.FolderFileXMLPath + Finance_Tools.simpleID + ".xml");
                XmlNode Nodes = xdoc.SelectSingleNode("Nodes");
                Nodes.RemoveAll();
                filesString = "";
                for (int i = fldrs.Length; i > 0; i--)
                {
                    DirectoryInfo d = fldrs[i - 1];
                    if (Finance_Tools.IsAddedToGallery(d))
                    {
                        try
                        {//add folder nodes
                            XmlElement xmlElement = xdoc.CreateElement("Folder");//add attributes
                            xmlElement.SetAttribute("Name", d.Name);//The node is added to the specified node
                            XmlNode xml = Nodes.PrependChild(xmlElement);
                            FileInfo[] myfile = d.GetFiles();
                            for (int j = myfile.Length; j > 0; j--)
                            {
                                FileInfo f = myfile[j - 1];
                                if (f.Extension == ".xls" || f.Extension == ".xlsx" || f.Extension == ".xlsm" || f.Extension == ".pdf")
                                {
                                    string fileName = Path.GetFileNameWithoutExtension(f.Name);
                                    //using ShellLib project(framework2.0) for current project(framework3.5),because framework3.5 cannot using SHELL32(X64) directly.
                                    string fileDesc = string.Empty;
                                    try
                                    {
                                        ShellLib.ShellLibClass sl = new ShellLib.ShellLibClass();
                                        fileDesc = sl.getFileDetail(d.FullName, f.Name);
                                    }
                                    catch { }//add folder nodes
                                    XmlElement xmlElementSub = xdoc.CreateElement("File");//add attributes
                                    xmlElementSub.SetAttribute("Name", fileName);
                                    xmlElementSub.SetAttribute("Description", fileDesc);
                                    for (int x = 0; x < sArray.Length; x++)
                                    {
                                        if (!string.IsNullOrEmpty(sArray[x]) && sArray[x].ToLower().Contains(f.FullName.ToLower()))
                                        {
                                            xmlElementSub.SetAttribute("ID", sArray[x].Substring(sArray[x].LastIndexOf("-") + 1));
                                        }
                                    }//The node is added to the specified node
                                    xml.PrependChild(xmlElementSub);
                                    if (!fileName.Contains("~"))
                                        filesString += fileName;
                                }
                            }
                        }
                        catch { }
                    }
                }
                xdoc.Save(Finance_Tools.FolderFileXMLPath + Finance_Tools.simpleID + ".xml");
                Finance_Tools.TemplateCount = filesString;
            }
        }
    }
}
