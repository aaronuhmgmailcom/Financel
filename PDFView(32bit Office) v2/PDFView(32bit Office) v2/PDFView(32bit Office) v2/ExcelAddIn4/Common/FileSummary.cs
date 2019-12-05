using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace ExcelAddIn4.Common
{
    class FileSummary
    {
        public static void SetProperty(string filename,string msg , SummaryPropId summaryType)
        {
            IPropertySetStorage propSetStorage = null;
            Guid IID_PropertySetStorage = new Guid("0000013A-0000-0000-C000-000000000046");
            uint hresult = ole32.StgOpenStorageEx(filename,(int)(STGM.SHARE_EXCLUSIVE | STGM.READWRITE),(int)STGFMT.FILE,0,(IntPtr)0,(IntPtr)0,ref IID_PropertySetStorage,ref propSetStorage);
            Guid fmtid_SummaryProperties = new Guid("F29F85E0-4FF9-1068-AB91-08002B27B3D9");
            IPropertyStorage propStorage = null;

            hresult = propSetStorage.Create(ref fmtid_SummaryProperties,(IntPtr)0,(int)PROPSETFLAG.DEFAULT,(int)(STGM.CREATE | STGM.READWRITE |STGM.SHARE_EXCLUSIVE),ref propStorage);

            PropSpec propertySpecification = new PropSpec();
            propertySpecification.ulKind = 1;
            propertySpecification.Name_Or_ID = new IntPtr((int)summaryType);

            PropVariant propertyValue = new PropVariant();
            propertyValue.FromObject(msg);
         
            propStorage.WriteMultiple(1, ref propertySpecification, ref propertyValue, 2);
            hresult = propStorage.Commit((int)STGC.DEFAULT);
            System.Runtime.InteropServices.Marshal.ReleaseComObject(propSetStorage);
            propSetStorage = null;
            GC.Collect();
        }
    }
}
