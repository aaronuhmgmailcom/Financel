using System;
using System.Collections.Generic;
using System.Text;
using Shell32;

namespace ShellLib
{
    public class ShellLibClass
    {
        /// <summary>
        /// Shell32 for x86 using to show file details and base on Framework2.0 ,distinct to the project ExcelAddIn4(using Framework 3.5) 
        /// </summary>
        /// <param name="fullname"></param>
        /// <param name="filename"></param>
        /// <returns></returns>
        public string getFileDetail(string fullname, string filename)
        {
            string returnValue = string.Empty;
            //The file path to get a property  
            string filePath = fullname + "\\" + filename;
            //Initialize the Shell interface
            Shell32.Shell shell = new Shell32.ShellClass();
            //Gets the file where the parent directory objects  
            Folder folder = shell.NameSpace(filePath.Substring(0, filePath.LastIndexOf("\\")));
            //Gets the file corresponding to the FolderItem object
            FolderItem item = folder.ParseName(filePath.Substring(filePath.LastIndexOf("\\") + 1));
            //Key relation dictionary stored attribute names and values
            //Dictionary<string, string> Properties = new Dictionary<string, string>();  
            int i = 0;
            while (true)
            {
                //Gets the attribute name  
                string key = folder.GetDetailsOf(null, i);
                if (string.IsNullOrEmpty(key))
                {
                    //When no properties desirable, EXIT cycle
                    break;
                }
                //Getting the property value
                string value = folder.GetDetailsOf(item, i);
                if (key == "Comments")
                {
                    returnValue = value;
                    break;
                }
                //Save attribute  
                //Properties.Add(key, value);  
                i++;
            }

            return returnValue;

        }
    }
}
