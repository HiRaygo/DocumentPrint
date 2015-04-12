using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Globalization;
using System.Runtime.InteropServices;
using System.Reflection;


namespace DocumentPrint
{
    internal class WordPrint : IDisposable
    {
    	private Type WordType;
        private object WordApp;
        private object oDocs;
        private object oDoc;
        private object[] oParams;
        private object missing = System.Reflection.Missing.Value;
        private byte WordResult = 0;
        private string version = null;
        
        /// <summary>
        /// Field: Version of Word App
        /// </summary>
        public string Version
        {
            get
            {
                if (version == null && WordType != null)
                {
                    version = WordApp.GetType().InvokeMember("Version", BindingFlags.GetProperty, null, WordApp, null).ToString();
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
        		return WordResult;
        	}
        }
        
        /// <summary>
        /// Constructor
        /// </summary>
        public WordPrint()
        {
        	WordResult = 0;
            try
            {
                WordType = Type.GetTypeFromProgID("Word.Application");
                WordApp = Activator.CreateInstance(WordType);                
            }
            catch
            {
                WordResult = 3;
            }
        }

        /// <summary>
        /// Dispose
        /// </summary>
        public void Dispose()
        {
            Close();
            Quit();

            WordApp = null;
            oDocs = null;
            oDoc = null;
        }
        
        /// <summary>
        /// Open the document
        /// </summary>
        /// <param name="strFileName">Name of file to be printed</param>
        public void Open(string strFileName)
        {
            if (WordType != null)
            {
            	try{
            		//Hide App
            		oParams = new object[1];
	                oParams[0] = false;
	                WordApp.GetType().InvokeMember("Visible", BindingFlags.SetProperty, null, WordApp, oParams);
	                //WordApp.GetType().InvokeMember("DisplayAlerts", BindingFlags.SetProperty, null, WordApp, oParams);
	                //Get Document	                
	                oDocs = WordApp.GetType().InvokeMember("Documents", BindingFlags.GetProperty, null, WordApp, null);
	                //Open File
	                oParams = new object[3];
	                oParams[0] = strFileName;
	                oParams[1] = missing;
	                oParams[2] = true;	
	                oDoc = oDocs.GetType().InvokeMember("Open", BindingFlags.InvokeMethod, null, oDocs, oParams);
            	}
            	catch{
            		WordResult = 12;
            	}
            }
        }

        /// <summary>
        /// Close the document
        /// </summary>
        public void Close()
        {
            if (WordType != null)
            {
            	try{
	                GC.Collect();
	                GC.WaitForPendingFinalizers();
	                //Close Document
	                oParams = new object[3];
	                oParams[0] = 0;
	                oParams[1] = 0;
	                oParams[2] = false;
	                oDoc.GetType().InvokeMember("Close", BindingFlags.InvokeMethod, null, oDoc, oParams);
	                Marshal.FinalReleaseComObject(oDoc);
	                //Close Documents
	                Marshal.FinalReleaseComObject(oDocs);
            	}
            	catch{
            		WordResult = 13;
            	}
            }
        }

        /// <summary>
        /// Print the file
        /// </summary>
        /// <param name="printerName">Name of Printer</param>
        public void Print(string printerName)
        {
        	if ((WordType != null)&&(WordResult ==0))
            {
        		oParams = new object[1];
		        oParams[0] = printerName;
        		try{	        		
		            WordApp.GetType().InvokeMember("ActivePrinter", BindingFlags.SetProperty, null, WordApp, oParams);
        		}
        		catch{
        			WordResult = 7;
        			return;
        		}
	                
                oParams = new object[1];
                oParams[0] = false;		//Background
//                oParams[1] = false;		//Append
//                oParams[2] = missing;	//Range
//                oParams[3] = missing;	//OutputFileName
//                oParams[4] = missing;	//From
//                oParams[5] = missing;	//To
//                oParams[6] = missing;	//Item
//                oParams[7] = 1;			//Copies
//                oParams[8] = missing;	//Pages
//                oParams[9] = missing;	//PageType
//                oParams[10] = false;	//PrintToFile
//                oParams[11] = missing;	//Collate
//                oParams[12] = missing;	//FileName
//                oParams[13] = missing;	//ActivePrinterMacGX
//                oParams[14] = missing;	//ManualDuplexPrint
//                oParams[15] = missing;	//PrintZoomColumn
//                oParams[16] = missing;	//PrintZoomRow
//                oParams[17] = missing;	//PrintZoomPaperWidth
//                oParams[18] = missing;	//PrintZoomPaperHeight
                try{
                	oDoc.GetType().InvokeMember("PrintOut", BindingFlags.InvokeMethod, null, oDoc, oParams);
                }
                catch{
                	WordResult = 14;
                }
            }
        }

        /// <summary>
        /// Quit the Word App
        /// </summary>
        public void Quit()
        {
            if (WordType != null)
            {
            	try{
                	WordApp.GetType().InvokeMember("Quit", BindingFlags.InvokeMethod, null, WordApp, null);
                	Marshal.FinalReleaseComObject(WordApp);
            	}catch{}
            }
            oParams = null;
            missing = null;
            WordType = null;
        }
    
    }
}
