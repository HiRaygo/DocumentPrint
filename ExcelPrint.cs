using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Globalization;
using System.Runtime.InteropServices;
using System.Reflection;


namespace DocumentPrint
{
    internal class ExcelPrint : IDisposable
    {
    	private Type ExcelType;
        private object ExcelApp;
        private object oBook;
        private object oDocs;
        private object[] oParams;
        private object missing = System.Reflection.Missing.Value;
        private byte ExcelResult = 0;
        private string version = null;
        
        /// <summary>
        /// Constructor
        /// </summary>
        public ExcelPrint()
        {
            try
            {
                ExcelType = Type.GetTypeFromProgID("Excel.Application");
                ExcelApp = Activator.CreateInstance(ExcelType);
                ExcelResult = 0;
            }
            catch
            {
                ExcelResult = 3;
            }
        }
        
        /// <summary>
        /// Dispose
        /// </summary>
        public void Dispose()
        {
            Close();
            Quit();

            ExcelApp = null;
            oBook = null;
            oDocs = null;
        }
        
		/// <summary>
		/// Field: Version of Excel App
		/// </summary>
        public string Version
        {
            get
            {
                if (version == null && ExcelType != null)
                {
                    version = ExcelApp.GetType().InvokeMember("Version", BindingFlags.GetProperty, null, ExcelApp, null).ToString();
                }
                return version;
            }
        }
		
        /// <summary>
        /// Field: Result of print
        /// </summary>
        public byte PrintResult
        {
        	get{
        		return ExcelResult;
        	}
        }
 		
		/// <summary>
        /// Open Excel App
        /// </summary>
        /// <param name="strFileName">Name of file to be printed</param>
        public void Open(string strFileName)
        {
            if (ExcelType != null)
            {
            	try{
            		oParams = new object[1];
	                oParams[0] = false;
	                ExcelApp.GetType().InvokeMember("Visible", BindingFlags.SetProperty, null, ExcelApp, oParams);
	                
	                oDocs = ExcelApp.GetType().InvokeMember("Workbooks", BindingFlags.GetProperty, null, ExcelApp, null, CultureInfo.InvariantCulture);
	                oParams = new object[3];
	                oParams[0] = strFileName;
	                oParams[1] = missing;
	                oParams[2] = true;	
	                oBook = oDocs.GetType().InvokeMember("Open", BindingFlags.InvokeMethod, null, oDocs, oParams);
            	}
            	catch{
            		ExcelResult = 12;
            	}
            }
        }

        /// <summary>
        /// Close Excel App
        /// </summary>
        public void Close()
        {
            if (ExcelType != null)
            {
            	try{
	                GC.Collect();
	                GC.WaitForPendingFinalizers();
	                Marshal.FinalReleaseComObject(oDocs);
	
	                oBook.GetType().InvokeMember("Close", BindingFlags.InvokeMethod, null, oBook, null);
	                Marshal.FinalReleaseComObject(oBook);
            	}
            	catch{
            		ExcelResult = 13;
            	}
            }
        }

        /// <summary>
        /// Print the file
        /// </summary>
        /// <param name="printerName">Name of Printer</param>
        public void Print(string printerName)
        {
        	if ((ExcelType != null)&&(ExcelResult ==0))
            {
                oParams = new object[8];
                oParams[0] = missing;
                oParams[1] = missing;
                oParams[2] = missing;
                oParams[3] = missing;
                oParams[4] = printerName == null ? missing : printerName;
                oParams[5] = missing;
                oParams[6] = missing;
                oParams[7] = missing;
                try{
                	oBook.GetType().InvokeMember("PrintOut", BindingFlags.InvokeMethod, null, oBook, oParams);
                }
                catch{
                	ExcelResult = 14;
                }
            }
        }

        /// <summary>
        /// Quit the Excel App
        /// </summary>
        public void Quit()
        {
            if (ExcelType != null)
            {
            	try{
                	ExcelApp.GetType().InvokeMember("Quit", BindingFlags.InvokeMethod, null, ExcelApp, null);
                	Marshal.FinalReleaseComObject(ExcelApp);
            	}catch{}
            }
            oParams = null;
            missing = null;
            ExcelType = null;
        }
    }
}
