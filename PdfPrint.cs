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
        /// <param name="filename">File Name</param>
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
		                pageImage = pdfFile.
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
		
		/// <summary>
        /// 合并图片
        /// </summary>
        /// <param name="maps"></param>
        /// <returns></returns>
        private static Bitmap MergerImg(params Bitmap[] maps)
        {
            int i = maps.Length;            
            if (i == 0)
                throw new Exception("图片数不能够为0");
            else if (i == 1)
                return maps[0];

            //创建要显示的图片对象,根据参数的个数设置宽度
            Bitmap backgroudImg = new Bitmap(maps[0].Width, i * maps[0].Height);
            Graphics g = Graphics.FromImage(backgroudImg);
            //清除画布,背景设置为白色
            g.Clear(System.Drawing.Color.White);
            for (int j = 0; j < i; j++)
            {
                g.DrawImage(maps[j], 0, j * maps[j].Height, maps[j].Width, maps[j].Height);
            }
            g.Dispose();
            return backgroudImg;
        }
        
        /// <summary>
        /// 将PDF文档转换为图片的方法
        /// </summary>
        /// <param name="pdfInputPath">PDF文件路径</param>
        /// <param name="imageOutputPath">图片输出完整路径(包括文件名)</param>
        /// <param name="startPageNum">从PDF文档的第几页开始转换</param>
        /// <param name="endPageNum">从PDF文档的第几页开始停止转换</param>
        /// <param name="imageFormat">设置所需图片格式</param>
        /// <param name="definition">设置图片的清晰度，数字越大越清晰</param>
        private static void ConvertPdf2Image(string pdfInputPath, string imageOutputPath,
             int startPageNum, int endPageNum, ImageFormat imageFormat, int definition)
        {
            
            PDFRender.PDFFile pdfFile = PDFRender.PDFFile.Open(pdfInputPath);
            
            if (startPageNum <= 0)
            {
                startPageNum = 1;
            }

            if (endPageNum > pdfFile.PageCount)
            {
                endPageNum = pdfFile.PageCount;
            }

            if (startPageNum > endPageNum)
            {
                int tempPageNum = startPageNum;
                startPageNum = endPageNum;
                endPageNum = startPageNum;
            }

            var bitMap = new Bitmap[endPageNum];

            for (int i = startPageNum; i <= endPageNum; i++)
            {
                Bitmap pageImage = pdfFile.GetPageImage(i - 1, 56 * definition);
                Bitmap newPageImage = new Bitmap(pageImage.Width/4 , pageImage.Height/4);

                Graphics g = Graphics.FromImage(newPageImage);
                g.InterpolationMode = InterpolationMode.HighQualityBicubic;
　　　　　　　　　　　//重新画图的时候Y轴减去130，高度也减去130  这样水印就看不到了
                g.DrawImage(pageImage, new Rectangle(0, 0, pageImage.Width/4 , pageImage.Height/4),
                    new Rectangle(0, 130, pageImage.Width, pageImage.Height-130), GraphicsUnit.Pixel);

                bitMap[i - 1] = newPageImage;
　　　　　　　　　g.Dispose();
            }

            //合并图片
            var mergerImg = MergerImg(bitMap);
            //保存图片
            mergerImg.Save(imageOutputPath, imageFormat);
            pdfFile.Dispose();
        }
		
	}
}
