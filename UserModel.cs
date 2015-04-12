using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Drawing;
using System.Drawing.Printing;
using System.IO;
using System.Linq;
using System.Management;
using System.Runtime.InteropServices;
using System.Text;

using Microsoft.Win32;

namespace DocumentPrint
{
    /// <summary>
    /// Returns the handled result for printing action.
    /// </summary>
    public enum DocPrintResult : byte
    {
        /// <summary>
        /// Not operated yet.
        /// </summary>
        NOT_SENT_TO_PRINTER = 0,
        /// <summary>
        /// File is sent to printer successfully.
        /// </summary>
        SENT_SUCCESSFULLY = 1,
        /// <summary>
        /// Unhandled error occured.
        /// </summary>
        UNSPECIFIED_ERROR = 2,
        /// <summary>
        /// Any application not installed on running machine.
        /// </summary>
        APPLICATION_NOT_FOUND = 3,
        /// <summary>
        /// Newer version of excel/word required.
        /// </summary>
        INVALID_APP_VERSION = 4,
        /// <summary>
        /// Tried to unsupported file format.
        /// </summary>
        UNSUPPORTED_FILE_FORMAT = 6,
        /// <summary>
        /// The target printer not found on running machine.
        /// </summary>
        PRINTER_NOT_FOUND = 7,
        /// <summary>
        /// Specified file not accessable on running machine.
        /// </summary>
        FILE_NOT_FOUND = 8,
        /// <summary>
        /// Not found any supported PDF file viewer on running machine.
        /// </summary>
        PDF_VIEWER_NOT_FOUND = 10,
        /// <summary>
        /// Not found any supported Text file viewer on running machine.
        /// </summary>
        TEXT_VIEWER_NOT_FOUND = 11,
        /// <summary>
        /// Can Not Open File
        /// </summary>
        FILE_CANNOT_OPEN = 12,
        /// <summary>
        /// Can Not Close File
        /// </summary>
        FILE_CANNOT_CLOSE = 13,
        /// <summary>
        /// Can Not Close File
        /// </summary>
        PRINTING_ERROR = 14
    }

    /// <summary>
    /// A class for the printing jobs.
    /// </summary>
    public class UserModel : IDisposable
    {
        private List<LocalPrinter> printerList;
        private const int ERROR_INSUFFICIENT_BUFFER = 122;
		private DocPrintResult printResult = DocPrintResult.NOT_SENT_TO_PRINTER;
        /// <summary>
        /// Raises after sending each file to printer.
        /// </summary>
        public event EventHandler OnAfterSendPrinter = null;

        #region Constructors
        
        /// <summary>
        /// Initialize and read settings from Windows Registery.
        /// </summary>
        public UserModel()
        {
            printerList = new List<LocalPrinter>();
            try
            {
                ReadPrintersFromRegistery();
            }
            catch (Exception)
            {

                this.printResult = DocPrintResult.UNSPECIFIED_ERROR;
            }
        }

        #endregion

        #region Private Properties
        
        private LocalPrinter Printer
        {
            get
            {
                return printerList.SingleOrDefault(p => p.DeviceID == (this.PrinterName != null ? this.PrinterName : this.DefaultPrinter));
            }
        }

        #endregion

        #region Public Properties
        /// <summary>
        /// User based UniqueID.
        /// </summary>
        public string UniqueID { get; set; }

        /// <summary>
        /// Physical file path of printing document.
        /// </summary>
        public string FileName { get; set; }

        /// <summary>
        /// Set the target printer name.
        /// </summary>
        public string PrinterName { get; set; }        
        
        /// <summary>
        /// Action result for Print method.
        /// </summary>
        public DocPrintResult PrintResult
        {
            get { return printResult; }
        }

        #endregion

        #region Private Methods

        [DllImport("winspool.drv", CharSet = CharSet.Auto, SetLastError = true)]
        private static extern bool SetDefaultPrinter(string Name);

        [DllImport("winspool.drv", CharSet = CharSet.Auto, SetLastError = true)]
        private static extern bool GetDefaultPrinter(StringBuilder pszBuffer, ref int pcchBuffer);

        private string DefaultPrinter
        {
            get
            {
                int pcchBuffer = 0;
                if (!GetDefaultPrinter(null, ref pcchBuffer))
                {
                    int lastWin32Error = Marshal.GetLastWin32Error();
                    if (lastWin32Error == ERROR_INSUFFICIENT_BUFFER)
                    {
                        StringBuilder pszBuffer = new StringBuilder(pcchBuffer);
                        if (GetDefaultPrinter(pszBuffer, ref pcchBuffer))
                        {
                            return pszBuffer.ToString();
                        }
                    }
                }
                return null;
            }
        }

        private void ReadPrintersFromRegistery()
        {
            RegistryKey TAWKAY = RegistryKey.OpenRemoteBaseKey(Microsoft.Win32.RegistryHive.CurrentUser, "").OpenSubKey(@"Software\Microsoft\Windows NT\CurrentVersion\Devices");
            foreach (string regKey in TAWKAY.GetValueNames())
            {
                LocalPrinter printer = new LocalPrinter(regKey);
                string port = TAWKAY.GetValue(regKey).ToString();
                port = port.Substring(port.IndexOf(',') + 1);
                printer.DevicePortAddress = port;
                printerList.Add(printer);
            }
            TAWKAY.Close();
        }        

        private void RaiseOnAfterSendPrinter()
        {
            if (OnAfterSendPrinter != null)
                OnAfterSendPrinter(this, EventArgs.Empty);
        }
        #endregion

        #region Public Methods

        /// <summary>
        /// Log WMI informations for local printers.
        /// </summary>
        /// <param name="logFolderPath">Physical directory path of logging informations.</param>
        public void LogPrinterWMIInformations(string logFolderPath)
        {
            ManagementObjectSearcher searchPrinters = new ManagementObjectSearcher("SELECT * FROM Win32_Printer");
            foreach (ManagementObject printer in searchPrinters.Get())
            {
                using (StreamWriter writer = new StreamWriter(logFolderPath + @"\" + Convert.ToString(printer.Properties["DriverName"].Value + ".txt")))
                {
                    foreach (PropertyData data in printer.Properties)
                    {
                        writer.WriteLine(data.Name + " = " + Convert.ToString(data.Value));
                    }
                }
            }
            searchPrinters.Dispose();
        }

        /// <summary>
        /// Logs all printer informations to given path.
        /// </summary>
        /// <param name="filePath">Physical file path to logging.</param>
        public void SavePrinterList(string filePath)
        {
            using (StreamWriter writer = new StreamWriter(filePath))
            {
                foreach (LocalPrinter printer in printerList)
                {
                    writer.WriteLine(printer.ToString());
                }
            }
        }

        /// <summary>
        /// Print given file from default printers.
        /// </summary>
        /// <param name="fileName">Physical file path of printing document.</param>
        public void Print(string fileName)
        {
            this.FileName = fileName;
            this.PrinterName = null;
            this.UniqueID = null;

            Print();
        }

        /// <summary>
        /// Print given file from specific printers.
        /// </summary>
        /// <param name="fileName">Physical file path of printing document.</param>
        /// <param name="printerName">Set the target printer name.</param>
        public void Print(string fileName, string printerName)
        {
            this.FileName = fileName;
            this.PrinterName = printerName;
            this.UniqueID = null;

            Print();
        }

        /// <summary>
        /// Print given file from specific printers with UniqueID.
        /// </summary>
        /// <param name="fileName">Physical file path of printing document.</param>
        /// <param name="printerName">Set the target printer name(set null for using default printer).</param>
        /// <param name="uniqueID">User based UniqueID, useful for cathing file info from OnAfterSendPrinter event.</param>
        public void Print(string fileName, string printerName, string uniqueID)
        {
            this.FileName = fileName;
            this.PrinterName = printerName;
            this.UniqueID = uniqueID;

            Print();
        }

        /// <summary>
        /// Prints a single document, FileName property is required.
        /// </summary>
        public void Print()
        {
            printResult = DocPrintResult.NOT_SENT_TO_PRINTER;
            FileInfo file = new FileInfo(this.FileName);
            if (!file.Exists)
            {
                printResult = DocPrintResult.FILE_NOT_FOUND;
                return;
            }

            if (this.Printer == null)
            {
                printResult = DocPrintResult.PRINTER_NOT_FOUND;
                return;
            }

            try
            {
                string extension = file.Extension.ToLower();
                switch (extension)
                {
                    case ".xls":
                    case ".xlsx":
                        {
                			printResult = PrintExcelDoc(file);
                			break;
                        }
                	case ".doc":
                    case ".docx":
                        {
                            printResult = PrintWordDoc(file);
                			break;
                        }
                    case ".txt":
                	case ".rtf":
                		{
                			printResult = PrintTextFile(file);
                			break;
                		}
                    case ".pdf":
                        {
                            printResult = PrintPdfFile(file);
                            break;
                        }
                    case ".jpg":
                    case ".jpe":
                    case ".jpeg":
                    case ".gif":
                    case ".png":
                    case ".bmp":
                    case ".tif":
                    case ".tiff":
                        {
                            printResult = PrintImage(file);
                            break;
                        }
                    default:
                        printResult = DocPrintResult.UNSUPPORTED_FILE_FORMAT;
                        break;
                }
            }
            catch(Exception ex)
            {
                Console.WriteLine(ex.Message);
                printResult = DocPrintResult.UNSPECIFIED_ERROR;
            }

            RaiseOnAfterSendPrinter();
        }
        
		/// <summary>
		/// Print Excel Document
		/// </summary>
        private DocPrintResult PrintExcelDoc(FileInfo file)
        {
        	//CultureInfo oldCI = Thread.CurrentThread.CurrentCulture;
            //Thread.CurrentThread.CurrentCulture = new CultureInfo("en-US");
            using (ExcelPrint objExcel = new ExcelPrint())
            {
                if (objExcel.PrintResult == 3){
                    return DocPrintResult.APPLICATION_NOT_FOUND;
                }
            	//Start to print
                if (objExcel.Version == "12.0"){
                    objExcel.Open(file.FullName);
                    objExcel.Print(this.Printer.DeviceID);                                        
                }
                else{
                    return DocPrintResult.INVALID_APP_VERSION;
                }
            	//Get print result
            	if (objExcel.PrintResult == 0){
            		return DocPrintResult.SENT_SUCCESSFULLY;                                		
            	}
            	else{
            		return (DocPrintResult)objExcel.PrintResult;
            	}                              	
            }
            //Thread.CurrentThread.CurrentCulture = oldCI;
        }
        
        /// <summary>
		/// Print Word Document
		/// </summary>
        private DocPrintResult PrintWordDoc(FileInfo file)
        {
            using (WordPrint objWord = new WordPrint())
            {
                if (objWord.PrintResult == 3){
                    return DocPrintResult.APPLICATION_NOT_FOUND;
                }
            	//Start to print
                if (objWord.Version == "12.0"){
                    objWord.Open(file.FullName);
                    objWord.Print(this.Printer.DeviceID);                                        
                }
                else{
                    return DocPrintResult.INVALID_APP_VERSION;
                }
            	//Get print result
            	if (objWord.PrintResult == 0){
            		return DocPrintResult.SENT_SUCCESSFULLY;                                		
            	}
            	else{
            		return (DocPrintResult)objWord.PrintResult;
            	}                              	
            }
        }
        
        /// <summary>
		/// Print Image
		/// </summary>
        private DocPrintResult PrintImage(FileInfo file)
        {
        	using (PrintDocument printer = new PrintDocument())
            {
                printer.PrintPage += delegate(object sender, PrintPageEventArgs e)
                {
                    Graphics g = e.Graphics;
                    g.DrawImage(Image.FromFile(file.FullName), printer.DefaultPageSettings.Margins.Left, printer.DefaultPageSettings.Margins.Top);
                };

                printer.PrinterSettings.PrinterName = this.Printer.DeviceID;
                printer.Print();
            }
            return DocPrintResult.SENT_SUCCESSFULLY;
        }
        
        /// <summary>
		/// Print PDF File
		/// </summary>
        private DocPrintResult PrintPdfFile(FileInfo file)
        {
        	using (PdfPrint objPdf = new PdfPrint(file.FullName))
            {
            	//Start to print
                objPdf.Print(this.Printer.DeviceID);                                       
            	//Get print result
            	if (objPdf.PrintResult == 0){
            		return DocPrintResult.SENT_SUCCESSFULLY;                                		
            	}
            	else{
            		return (DocPrintResult)objPdf.PrintResult;
            	}                              	
            }
        }
        
        /// <summary>
        /// Print Text File
        /// </summary>
        private DocPrintResult PrintTextFile(FileInfo file)
        {
        	Process pr= null;
        	DocPrintResult ret;
        	try  
			{  
        		pr = new Process();
				//Hide window
				pr.StartInfo.CreateNoWindow = true;
				pr.StartInfo.WindowStyle = ProcessWindowStyle.Hidden;
				//Auto Diagnostics Mode
				pr.StartInfo.UseShellExecute = true;
				//Set operation
				pr.StartInfo.Verb = "Print";
				//file to be print
				pr.StartInfo.FileName = file.FullName;
				//Start
				pr.Start();
				//Wait to finish
				pr.WaitForExit(10000);
				//Print Success
				ret = DocPrintResult.SENT_SUCCESSFULLY;
			}  
			catch  
			{  
				ret = DocPrintResult.PRINTING_ERROR;
			}
			finally
			{
				pr = null;
			}
			return ret;
        }
        
        /// <summary>
        /// Release the used resources.
        /// </summary>
        public void Dispose()
        {
            printerList = null;
            OnAfterSendPrinter = null;
        }

        #endregion
    }
}
