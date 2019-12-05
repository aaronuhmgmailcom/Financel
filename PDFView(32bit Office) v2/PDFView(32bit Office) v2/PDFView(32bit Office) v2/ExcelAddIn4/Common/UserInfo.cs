
/*  
 * Module ID：<ExcelAddIn4>   
 * Function：<UserInfo>   
 * Author：Peter.uhm  (yanb@shinetechchina.com)
 * Modify date：2016.04
 * Version : 2.0.0.2
 */
using System;

namespace ExcelAddIn4
{
    /// <summary>
    /// 
    /// </summary>
    public class UserInfo
    {
        #region Field
        private string Id;
        private string formUserID;
        private string formUserPassword;
        private string windowsUserID;
        private string machineName;
        private int? loginType;
        private string containerpath;
        private string file_ftid;
        private string filePath;
        private int? invNumber;
        private string fileName;
        private string cachePath;
        private string folderName;
        private string sunUserIP;
        private string sunUserID;
        private string sunUserPass;
        private string balanceBy;
        private string useSequenceNumbering;
        private string sequencePrefix;
        private string postToField;
        private string populateCell;
        private string populateCellWithJnNumber;
        private string criteria1 = "";
        private string criteria2 = "";
        private string criteria3 = "";
        private string criteria4 = "";
        private string criteria5 = "";
        private string cellReference1 = "";
        private string cellReference2 = "";
        private string cellReference3 = "";
        private string cellReference4 = "";
        private string cellReference5 = "";
        private bool opentransuponSave = false;
        private bool useCriteria;
        private string sunJournalNumber = "";
        private string allocationMakerType = "";
        private string currentRef = "";
        private string currentSaveRef = "";
        private string comName = "";
        private string methodName = "";
        private string textfilename = "";
        private string globalError = "";
        private string allowBalTran = "";
        private string allowPostToSuspended = "";
        private string postProvisional = "";
        private ExcelAddIn4.Common.SessionFileDictionary dictionary;
        #endregion

        #region Property

        #region ID
        /// <summary>
        /// user's identity
        /// </summary>
        public string ID
        {
            get { return Id; }
            set { Id = value; }
        }
        #endregion

        #region formUserID
        /// <summary>
        /// 
        /// </summary>
        public string FormUserID
        {
            get { return formUserID; }
            set { formUserID = value; }
        }
        #endregion

        #region formUserPassword
        /// <summary>
        /// user's formUserPassword
        /// </summary>
        public string FormUserPassword
        {
            get { return formUserPassword; }
            set { formUserPassword = value; }
        }
        #endregion

        #region windowsUserID
        /// <summary>
        /// user's windowsUserID
        /// </summary>
        public string WindowsUserID
        {
            get { return windowsUserID; }
            set { windowsUserID = value; }
        }
        #endregion

        #region machineName
        /// <summary>
        /// user's machineName
        /// </summary>
        public string MachineName
        {
            get { return machineName; }
            set { machineName = value; }
        }
        #endregion

        #region loginType
        /// <summary>
        /// user's loginType
        /// </summary>
        public int? LoginType
        {
            get { return loginType; }
            set { loginType = value; }
        }
        #endregion

        #region Containerpath
        /// <summary>
        /// user's Containerpath
        /// </summary>
        public string Containerpath
        {
            get { return containerpath; }
            set { containerpath = value; }
        }
        #endregion

        /// <summary>
        /// 
        /// </summary>
        public string File_ftid
        {
            get { return file_ftid; }
            set { file_ftid = value; }
        }

        #region filePath
        /// <summary>
        /// user's filePath
        /// </summary>
        public string FilePath
        {
            get { return filePath; }
            set { filePath = value; }
        }
        #endregion

        #region cachePath
        /// <summary>
        /// user's cachePath
        /// </summary>
        public string CachePath
        {
            get { return cachePath; }
            set { cachePath = value; }
        }
        #endregion

        #region invNumber
        /// <summary>
        /// user's invNumber
        /// </summary>
        public int? InvNumber
        {
            get { return invNumber; }
            set { invNumber = value; }
        }
        #endregion

        #region fileName
        /// <summary>
        /// user's fileName
        /// </summary>
        public string FileName
        {
            get { return fileName; }
            set { fileName = value; }
        }
        #endregion

        #region folderName
        /// <summary>
        /// user's folderName
        /// </summary>
        public string FolderName
        {
            get { return folderName; }
            set { folderName = value; }
        }
        #endregion

        #region sunUserIP
        /// <summary>
        /// user's sunUserIP
        /// </summary>
        public string SunUserIP
        {
            get { return sunUserIP; }
            set { sunUserIP = value; }
        }
        #endregion
        #region sunUserID
        /// <summary>
        /// user's sunUserID
        /// </summary>
        public string SunUserID
        {
            get { return sunUserID; }
            set { sunUserID = value; }
        }
        #endregion

        #region sunUserPass
        /// <summary>
        /// user's sunUserPass
        /// </summary>
        public string SunUserPass
        {
            get { return sunUserPass; }
            set { sunUserPass = value; }
        }
        #endregion

        #region balanceBy
        /// <summary>
        /// user's balanceBy
        /// </summary>
        public string BalanceBy
        {
            get { return balanceBy; }
            set { balanceBy = value; }
        }
        #endregion

        #region useSequenceNumbering
        /// <summary>
        /// user's useSequenceNumbering
        /// </summary>
        public string UseSequenceNumbering
        {
            get { return useSequenceNumbering; }
            set { useSequenceNumbering = value; }
        }
        #endregion

        #region sequencePrefix
        /// <summary>
        /// user's sequencePrefix
        /// </summary>
        public string SequencePrefix
        {
            get { return sequencePrefix; }
            set { sequencePrefix = value; }
        }
        #endregion

        #region populateCell
        /// <summary>
        /// user's populateCell
        /// </summary>
        public string PopulateCell
        {
            get { return populateCell; }
            set { populateCell = value; }
        }
        #endregion

        #region populateCellWithJnNumber
        /// <summary>
        /// user's populateCellWithJnNumber
        /// </summary>
        public string PopulateCellWithJnNumber
        {
            get { return populateCellWithJnNumber; }
            set { populateCellWithJnNumber = value; }
        }
        #endregion
        #region criteria1
        /// <summary>
        /// user's Criteria1
        /// </summary>
        public string Criteria1
        {
            get { return criteria1; }
            set { criteria1 = value; }
        }
        #endregion

        #region criteria2
        /// <summary>
        /// user's Criteria2
        /// </summary>
        public string Criteria2
        {
            get { return criteria2; }
            set { criteria2 = value; }
        }
        #endregion

        #region criteria3
        /// <summary>
        /// user's Criteria3
        /// </summary>
        public string Criteria3
        {
            get { return criteria3; }
            set { criteria3 = value; }
        }
        #endregion

        #region criteria4
        /// <summary>
        /// user's Criteria4
        /// </summary>
        public string Criteria4
        {
            get { return criteria4; }
            set { criteria4 = value; }
        }
        #endregion

        #region criteria5
        /// <summary>
        /// user's Criteria5
        /// </summary>
        public string Criteria5
        {
            get { return criteria5; }
            set { criteria5 = value; }
        }
        #endregion

        #region cellReference1
        /// <summary>
        /// user's CellReference1
        /// </summary>
        public string CellReference1
        {
            get { return cellReference1; }
            set { cellReference1 = value; }
        }
        #endregion

        #region cellReference2
        /// <summary>
        /// user's CellReference2
        /// </summary>
        public string CellReference2
        {
            get { return cellReference2; }
            set { cellReference2 = value; }
        }
        #endregion

        #region cellReference3
        /// <summary>
        /// user's CellReference3
        /// </summary>
        public string CellReference3
        {
            get { return cellReference3; }
            set { cellReference3 = value; }
        }
        #endregion

        #region cellReference4
        /// <summary>
        /// user's CellReference4
        /// </summary>
        public string CellReference4
        {
            get { return cellReference4; }
            set { cellReference4 = value; }
        }
        #endregion

        #region cellReference5
        /// <summary>
        /// user's CellReference5
        /// </summary>
        public string CellReference5
        {
            get { return cellReference5; }
            set { cellReference5 = value; }
        }
        #endregion

        #region opentransuponSave
        /// <summary>
        /// user's opentransuponSave
        /// </summary>
        public bool OpentransuponSave
        {
            get { return opentransuponSave; }
            set { opentransuponSave = value; }
        }
        #endregion

        #region useCriteria
        /// <summary>
        /// user's useCriteria
        /// </summary>
        public bool UseCriteria
        {
            get { return useCriteria; }
            set { useCriteria = value; }
        }
        #endregion

        #region sunJournalNumber
        /// <summary>
        /// user's sunJournalNumber
        /// </summary>
        public string SunJournalNumber
        {
            get { return sunJournalNumber; }
            set { sunJournalNumber = value; }
        }
        #endregion

        #region currentRef
        /// <summary>
        /// user's currentRef
        /// </summary>
        public string CurrentRef
        {
            get { return currentRef; }
            set { currentRef = value; }
        }
        #endregion

        #region currentSaveRef
        /// <summary>
        /// user's currentSaveRef
        /// </summary>
        public string CurrentSaveRef
        {
            get { return currentSaveRef; }
            set { currentSaveRef = value; }
        }
        #endregion

        #region comName
        /// <summary>
        /// user's comName
        /// </summary>
        public string ComName
        {
            get { return comName; }
            set { comName = value; }
        }
        #endregion

        #region methodName
        /// <summary>
        /// user's methodName
        /// </summary>
        public string MethodName
        {
            get { return methodName; }
            set { methodName = value; }
        }
        #endregion

        #region textfilename
        /// <summary>
        /// user's textfilename
        /// </summary>
        public string Textfilename
        {
            get { return textfilename; }
            set { textfilename = value; }
        }
        #endregion

        #region globalError
        /// <summary>
        /// user's globalError
        /// </summary>
        public string GlobalError
        {
            get { return globalError; }
            set { globalError = value; }
        }
        #endregion
        #region allowBalTran
        /// <summary>
        /// user's allowBalTran
        /// </summary>
        public string AllowBalTran
        {
            get { return allowBalTran; }
            set { allowBalTran = value; }
        }
        #endregion

        #region allowPostToSuspended
        /// <summary>
        /// user's allowPostToSuspended
        /// </summary>
        public string AllowPostToSuspended
        {
            get { return allowPostToSuspended; }
            set { allowPostToSuspended = value; }
        }
        #endregion
        #region postProvisional
        /// <summary>
        /// user's postProvisional 
        /// </summary>
        public string PostProvisional
        {
            get { return postProvisional; }
            set { postProvisional = value; }
        }
        #endregion
        #region dictionary
        /// <summary>
        /// user's dictionary
        /// </summary>
        public ExcelAddIn4.Common.SessionFileDictionary Dictionary
        {
            get { return dictionary; }
            set { dictionary = value; }
        }
        #endregion

        #endregion
    }
}