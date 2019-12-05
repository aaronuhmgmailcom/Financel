using System;
using System.Collections;
using System.Collections.Generic;
using System.ComponentModel;
using System.Configuration.Install;
using System.Linq;
using System.Configuration;
using System.IO;
using System.Windows.Forms;
using System.Diagnostics;
using System.Security.AccessControl;

namespace ExcelAddIn4
{
    [RunInstaller(true)]
    public partial class InstallHelper : System.Configuration.Install.Installer
    {
        public InstallHelper()
        {
            InitializeComponent();
        }
        public override void Install(System.Collections.IDictionary stateSaver)
        {
            base.Install(stateSaver);

            try
            {
                //var map = new ExeConfigurationFileMap();
                ////Get app.config path
                //map.ExeConfigFilename = Context.Parameters["assemblypath"] + ".config";
                ////Get Config and AppSettings
                //var config = ConfigurationManager.OpenMappedExeConfiguration(map, ConfigurationUserLevel.None);
                //var appSettings = config.AppSettings;
                ////Get input value from setup project
                //var configValueList = new List<string>() { Context.Parameters["Config1"], 
                //                                           Context.Parameters["Config2"], 
                //                                           Context.Parameters["Config3"] };
                ////assign input value to appSettings
                //for (int i = 1; i <= 3; i++)
                //{
                //    appSettings.Settings["Sample.Config" + i].Value = configValueList[i - 1];
                //}
                ////save app.config
                //config.Save();
                //System.Environment.SpecialFolder.ApplicationData

                string s = Context.Parameters["Config1"];
                string s2 = Context.Parameters["Config2"];
                string s3 = Context.Parameters["Config3"];

                string path = Context.Parameters["assemblypath"].ToString().Substring(0, Context.Parameters["assemblypath"].ToString().LastIndexOf("\\"));
                //var map = new ExeConfigurationFileMap();

                //MessageBox.Show(path + "\\Application Files\\ExcelAddIn21_1_0_0_63\\" + Path.GetFileName(Context.Parameters["assemblypath"]) +".config");
                ////Get app.config path
                //map.ExeConfigFilename = path + "\\Application Files\\ExcelAddIn21_1_0_0_63\\" + Path.GetFileName(Context.Parameters["assemblypath"]) +".config.deploy" ;

                ////Get Config and AppSettings
                //var config = ConfigurationManager.OpenMappedExeConfiguration(map, ConfigurationUserLevel.None);
                //var appSettings = config.AppSettings;

                //appSettings.Settings["IntermediateConfig"].Value = path + "\\RSDataConfig\\Server.txt";
                ////save app.config
                //config.Save();

                ////System.Configuration.Configuration config = ConfigurationManager.OpenExeConfiguration(ConfigurationUserLevel.None);
                ////config.AppSettings.Settings["IntermediateConfig"].Value = path + "\\RSDataConfig\\Server.txt";
                ////config.Save(ConfigurationSaveMode.Modified);
                ////ConfigurationManager.RefreshSection("appSettings");

                string s4 = Context.Parameters["Config4"];



                DirectoryInfo mypath = new DirectoryInfo("C:\\ProgramData\\RSDataV2\\RSDataConfig");
                if (mypath.Exists)
                {
                }
                else
                {
                    mypath.Create();
                }
                if (File.Exists("C:\\ProgramData\\RSDataV2\\RSDataConfig\\Server.txt"))
                {
                    File.Delete("C:\\ProgramData\\RSDataV2\\RSDataConfig\\Server.txt");
                    //File.Delete(path + "\\RSDataConfig\\Server.txt");
                }

                FileStream aFile = new FileStream("C:\\ProgramData\\RSDataV2\\RSDataConfig\\Server.txt", FileMode.OpenOrCreate, FileAccess.Write, FileShare.ReadWrite);
                StreamWriter sw = new StreamWriter(aFile);
                sw.WriteLine("{0}", DEncrypt.Encrypt(Context.Parameters["Config1"]));
                sw.WriteLine("{0}", DEncrypt.Encrypt(Context.Parameters["Config2"]));
                sw.WriteLine("{0}", DEncrypt.Encrypt(Context.Parameters["Config3"]));
                sw.Close();

                Finance_Tools.AddDirectorySecurity("C:\\ProgramData\\RSDataV2\\RSDataConfig\\Server.txt", @"Everyone", System.Security.AccessControl.FileSystemRights.FullControl, AccessControlType.Allow);

                if (!File.Exists("C:\\ProgramData\\RSDataV2\\RSDataConfig\\DataFieldsSetting.xml"))
                {
                    FileStream aFile2 = new FileStream("C:\\ProgramData\\RSDataV2\\RSDataConfig\\DataFieldsSetting.xml", FileMode.OpenOrCreate, FileAccess.Write, FileShare.ReadWrite);
                    StreamWriter sw2 = new StreamWriter(aFile2);
                    sw2.Write("{0}", XMLCollection.xml);
                    sw2.Close();

                    Finance_Tools.AddDirectorySecurity("C:\\ProgramData\\RSDataV2\\RSDataConfig\\DataFieldsSetting.xml", @"Everyone", System.Security.AccessControl.FileSystemRights.FullControl, AccessControlType.Allow);
                }
                if (!File.Exists("C:\\ProgramData\\RSDataV2\\RSDataConfig\\GenDescFieldsSetting.xml"))
                {
                    FileStream aFile3 = new FileStream("C:\\ProgramData\\RSDataV2\\RSDataConfig\\GenDescFieldsSetting.xml", FileMode.OpenOrCreate, FileAccess.Write, FileShare.ReadWrite);
                    StreamWriter sw3 = new StreamWriter(aFile3);
                    sw3.Write("{0}", XMLCollection.xmlGen);
                    sw3.Close();

                    Finance_Tools.AddDirectorySecurity("C:\\ProgramData\\RSDataV2\\RSDataConfig\\GenDescFieldsSetting.xml", @"Everyone", System.Security.AccessControl.FileSystemRights.FullControl, AccessControlType.Allow);
                }
                //string installFile = Context.Parameters["assemblypath"].ToString().Replace("dll", "vsto");
                //Process.Start(installFile);


                //============================================================setup3============================================
                //string s = Context.Parameters["Config1"];
                //string path = Context.Parameters["assemblypath"].ToString().Substring(0, Context.Parameters["assemblypath"].ToString().LastIndexOf("\\"));

                //DirectoryInfo mypath = new DirectoryInfo("C:\\ProgramData\\RSData\\RSDataConfig");
                //if (mypath.Exists)
                //{
                //}
                //else
                //{
                //    mypath.Create();
                //}
                //if (File.Exists("C:\\ProgramData\\RSData\\RSDataConfig\\Server.txt"))
                //{
                //    File.Delete("C:\\ProgramData\\RSData\\RSDataConfig\\Server.txt");
                //    //File.Delete(path + "\\RSDataConfig\\Server.txt");
                //}

                //FileStream aFile = new FileStream("C:\\ProgramData\\RSData\\RSDataConfig\\Server.txt", FileMode.OpenOrCreate, FileAccess.Write, FileShare.ReadWrite);
                //StreamWriter sw = new StreamWriter(aFile);
                //sw.Write("{0}", Context.Parameters["Config1"]);
                //sw.Close();

                //Finance_Tools.AddDirectorySecurity("C:\\ProgramData\\RSData\\RSDataConfig\\Server.txt", @"Everyone", System.Security.AccessControl.FileSystemRights.FullControl, AccessControlType.Allow);
            }
            catch (Exception e)
            {
                string s = e.Message;
            }
        }

    }
}
