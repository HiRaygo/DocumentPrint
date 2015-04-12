using System;
using System.Diagnostics;
using System.Drawing;
using System.Drawing.Imaging;
using System.Drawing.Printing;
using System.Drawing.Drawing2D;
using PDFRender = O2S.Components.PDFRender4NET;

namespace DocumentPrint
{
	
	internal class PdfPrint: IDisposable
	{
		private byte PdfResult;
		private PDFRender.PDFFile pdfFile;
		private string FileName;
		private PrintDocument printer;
		/// <summary>
        /// Field: Result of print
        /// </summary>
        public byte PrintResult
        {
        	get{return PdfResult; }
        }
        
        /// <summary>
        /// Constructor
        /// </summary>
		public PdfPrint(string filename)
		{
			this.FileName = filename;
			PdfResult = 0;
			try{
				pdfFile = PDFRender.PDFFile.Open(FileName);
				printer = new PrintDocument();
			}
			catch{
				PdfResult = 3;
			}
		}
		
		/// <summary>
        /// Dispose
        /// </summary>
        public void Dispose()
        {
            pdfFile.Dispose();
            printer.Dispose();
        }
        
        /// <summary>
        /// Print
        /// </summary>
        /// <param name="printerName">Printer Name</param>
		public void PrintX(string printername)
		{
			if(PdfResult == 0 && pdfFile != null)
			{
				try  
				{
					//打印机设置
					PrinterSettings settings = new PrinterSettings();
					settings.PrinterName = printername;
					settings.PrintToFile = false;
					settings.Copies = 1;
					settings.PrintRange = PrintRange.AllPages;
					//设置纸张大小（可以不设置取，取默认设置）
		            PaperSize ps = new PaperSize("Your Paper Name",595,842);
		            ps.RawKind = 9; //如果是自定义纸张，就要大于118
		            
		            //打印设置
		            PDFRender.Printing.PDFPrintSettings pdfPrintSettings = new PDFRender.Printing.PDFPrintSettings(settings);
		            pdfPrintSettings.AutoRotate = false;			//自动旋转
		            pdfPrintSettings.BitmapPrintResolution = 560;	//图片打印精度		            
		            pdfPrintSettings.PaperSize = ps;				//纸张尺寸
		            pdfPrintSettings.PageScaling = PDFRender.Printing.PageScaling.FitToPrinterMarginsProportional;
					
		            pdfFile.Print(pdfPrintSettings);
                      
				}  
				catch(Exception ex)
				{  
					Console.Write(ex.Message);
					PdfResult = 14;
				}
			}
		}
		
        /// <summary>
        /// Print
        /// </summary>
        /// <param name="printerName">Printer Name</param>
		public void Print(string printername)
		{
			if(PdfResult == 0)
			{
				try  
				{
					PrinterSettings settings = new PrinterSettings();
					settings.PrinterName = printername;
					settings.PrintToFile = false;
					//设置纸张大小（可以不设置取，取默认设置）
		            //PaperSize ps = new PaperSize("Your Paper Name",595,842);
		            //ps.RawKind = 9; //如果是自定义纸张，就要大于118，
					printer.PrinterSettings = settings;
					//图片拉伸
					ImageAttributes ImgAtt = new ImageAttributes();
					ImgAtt.SetWrapMode(WrapMode.TileFlipXY);
					//分页转换
					int endPageNum = pdfFile.PageCount;
					Bitmap pageImage = null;
		            for (int i = 0; i < endPageNum; i++)
		            {
		                pageImage = pdfFile.GetPageImage(i, 280);
		                //pageImage.Save(i.ToString() + ".png",ImageFormat.Png);
		                printer.PrintPage += delegate(object sender, PrintPageEventArgs e)
		                {
		                    Graphics g = e.Graphics;
		                    g.InterpolationMode = InterpolationMode.HighQualityBicubic;
		                    //指定所绘制图像的位置和大小。将图像进行缩放以适合该矩形
		                    //var dstRect = new Rectangle(20, 20, pageImage.Width/4 , pageImage.Height/4);
		                    Point[] points = {new Point(20, 20), new Point(pageImage.Width/4, 20), new Point(20, pageImage.Height/4)};
		                    Point[] pointsA4 = {new Point(20, 20), new Point(575, 20), new Point(20, 822)};
		                    //指定 image 对象中要绘制的部分
		                    var imgRect = new Rectangle(0, 0, pageImage.Width , pageImage.Height);
		                    g.DrawImage(pageImage, pointsA4, imgRect, GraphicsUnit.Pixel, ImgAtt);
		                };
		                printer.Print();
		                pageImage.Dispose();
		            }	            
				}  
				catch(Exception ex)
				{  
					Console.Write(ex.Message);
					PdfResult = 14;
				}
			}
		}
			
	}
}
