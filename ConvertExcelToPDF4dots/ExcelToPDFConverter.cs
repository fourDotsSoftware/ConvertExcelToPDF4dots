using System;
using System.Collections.Generic;
using System.Text;
using System.Drawing;
using System.Reflection;

namespace ConvertExcelToPDF4dots
{
    public class ExcelToPDFConverter
    {        
        public string err = "";

        public bool ConvertToPDF(string filepath,string outfilepath)
        {
            err = "";
            
            object oDocuments = null;
            object doc = null;

            try
            {
                OfficeHelper.CreateExcelApplication();

                oDocuments = OfficeHelper.ExcelApp.GetType().InvokeMember("Workbooks", BindingFlags.InvokeMethod | BindingFlags.GetProperty, null, OfficeHelper.ExcelApp, null);

                doc = oDocuments.GetType().InvokeMember("Open", BindingFlags.InvokeMethod | BindingFlags.GetProperty, null, oDocuments, new object[] { filepath });

                /*
                System.Threading.Thread.Sleep(100);

                OfficeHelper.PPApp.GetType().InvokeMember("Activate", BindingFlags.IgnoreReturn | BindingFlags.Public |
                BindingFlags.Static | BindingFlags.InvokeMethod, null, OfficeHelper.PPApp, null);
                */

                System.Threading.Thread.Sleep(200);

                /*
                string fp=System.IO.Path.Combine(
                    System.IO.Path.GetDirectoryName(filepath),
                    System.IO.Path.GetFileNameWithoutExtension(filepath)+".pdf"
                    );                
                */

                doc.GetType().InvokeMember("ExportAsFixedFormat", BindingFlags.InvokeMethod, null, doc, new object[] { 0, outfilepath });

                oDocuments = null;
                doc = null;

                GC.Collect();
                GC.WaitForPendingFinalizers();

                return true;
            }
            catch (Exception ex)
            {
                err += TranslateHelper.Translate("Error could not Convert Excel to PDF") + " : " + filepath + "\r\n" + ex.Message;
                return false;
            }

            return true;
        }                
    }
}