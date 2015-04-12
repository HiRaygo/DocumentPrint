/*
 * 由SharpDevelop创建。
 * 用户： XiaoSanya
 * 日期: 2015/4/11
 * 时间: 19:06
 */
using System;
using System.Drawing;
using System.IO;
using System.Threading;
using System.Windows.Forms;

namespace DocumentPrint
{
	/// <summary>
	/// Description of Main.
	/// </summary>
	public partial class MainForm : Form
	{
		private string foldername;
		private bool isDelete;
		private bool isErrorStop;
		private bool StopPrint;
		private bool Printting;
		private int interval;
		delegate void ShowInfoCallBack(string msg);
		private ShowInfoCallBack ShowCB;
		
		public MainForm()
		{
			//
			// The InitializeComponent() call is required for Windows Forms designer support.
			//
			InitializeComponent();
			
			//
			// TODO: Add constructor code after the InitializeComponent() call.
			//
			foldername = null;
			isDelete = false;
			isErrorStop = false;
			StopPrint = false;
			Printting= false;
			interval = 30;
			ShowCB = new ShowInfoCallBack(ShowInfo);
		}
		
		void Button1Click(object sender, EventArgs e)
		{
			FolderBrowserDialog fbd = new FolderBrowserDialog();
			if(fbd.ShowDialog() == DialogResult.OK)
			{
				foldername = fbd.SelectedPath;
				this.labelFolder.Text = foldername;
			}
			else
			{
				this.labelFolder.Text = "未选择文件夹";
			}
		}
		
		private void PrintWork()
		{
			Printting = true;
			int count = 0;
			bool firsttime = true;
			while(!StopPrint)
			{				
				//Delay
				try{
					interval = int.Parse(this.textBoxInterval.Text);
				}
				catch{
					interval = 30;
				}
				if((interval >1000) || (interval <10))
				{
					interval = 30;
				}
				Thread.Sleep(1000);
				count++;
				if((count >= interval)||firsttime)
				{
					try{
						PrintFolder(foldername);
					}
					catch(Exception ex){
						this.Invoke(ShowCB,new object[1]{ex.Message});
					}
					count = 0;
					firsttime = false;
				}				
			}
			Printting = false;
		}
		
		/// <summary>
		/// Print files in a folder
		/// </summary>
		/// <param name="folder">folder</param>
		private void PrintFolder(string folder)
		{
			DirectoryInfo dirinfo = new DirectoryInfo(folder);
			if(!dirinfo.Exists)
			{
				return;
			}
			using (UserModel objPrint = new UserModel())
			{
				foreach (FileInfo fileinfo in dirinfo.GetFiles())
	            {
					bool printed = fileinfo.Name.StartsWith("Printed_");
					if(printed)
					{
						object sj = fileinfo.Name + "\r\nPrinted.";
						this.Invoke(ShowCB, new object[]{sj});
						continue;
					}
		        	objPrint.Print(fileinfo.FullName);
		        	string msg = fileinfo.Name;
		        	object[] obj = new object[1];
		        	obj[0] = msg;
		        	this.Invoke(ShowCB, obj);
		        	if(objPrint.PrintResult == DocPrintResult.SENT_SUCCESSFULLY)
		        	{
		        		obj[0] = "Send to printer success.";
		        		this.Invoke(ShowCB,obj);
		            	if(isDelete)
		            	{
		            		fileinfo.Delete();
		            		obj[0] = "Delete file finished.";
		            		this.Invoke(ShowCB, obj);
		            	}
		            	string newname = fileinfo.DirectoryName + "\\" + "Printed_" + fileinfo.Name;
		            	fileinfo.MoveTo(newname);
		        	}
		        	else
		        	{
		        		obj[0] = "Error: " + objPrint.PrintResult.ToString();
		        		this.Invoke(ShowCB, obj);
		        	}
		        	
		        	if(isErrorStop)
		        	{
		        		StopPrint = true;
		        		return;
		        	}
	            }
			}
		}		
		
		
		private void PrintFile(object ob)
		{
			PrintInfo pi = (PrintInfo)ob;
			FileInfo fileinfo = pi.fileinfo;
			UserModel objPrint = pi.pdfPrint;
			bool printed = fileinfo.Name.StartsWith("Printed_");
			if(printed)
			{
				object sj = fileinfo.Name + "\r\nPrinted.";
				this.Invoke(ShowCB, new object[]{sj});
				return;
			}
        	objPrint.Print(fileinfo.FullName);
        	string msg = fileinfo.Name;
        	object[] obj = new object[1];
        	obj[0] = msg;
        	this.Invoke(ShowCB, obj);
        	if(objPrint.PrintResult == DocPrintResult.SENT_SUCCESSFULLY)
        	{
        		obj[0] = "Send to printer success.";
        		this.Invoke(ShowCB,obj);
            	if(isDelete)
            	{
            		fileinfo.Delete();
            		obj[0] = "Delete file finished.";
            		this.Invoke(ShowCB, obj);
            	}
            	string newname = fileinfo.DirectoryName + "\\" + "Printed_" + fileinfo.Name;
            	fileinfo.MoveTo(newname);
        	}
        	else
        	{
        		obj[0] = "Error: " + objPrint.PrintResult.ToString();
        		this.Invoke(ShowCB, obj);
        	}
        	
        	if(isErrorStop)
        	{
        		StopPrint = true;
        		return;
        	}
		}
		
		private void ShowInfo(object msg)
		{
			string info = (string)msg;
			this.textBoxInfo.AppendText(info);
			this.textBoxInfo.AppendText("\r\n");
		}
		
		void ButtonStartClick(object sender, EventArgs e)
		{
			Thread tt = null;
			if(! Printting)
			{
				StopPrint = false;
				tt = new Thread(new ThreadStart(PrintWork));
				tt.IsBackground = true;
				tt.Start();	
				this.textBoxInfo.AppendText("Print work start.\r\n");
				this.buttonStart.Text = "停止打印";
			}
			else
			{
				StopPrint = true;
				Thread.Sleep(1000);
				if((tt!= null) &&(tt.IsAlive))
					tt.Abort();
				this.textBoxInfo.AppendText("Print work stop.\r\n");
				this.buttonStart.Text = "开始打印";
			}
		}
	}
	
	internal class PrintInfo
	{
		public FileInfo fileinfo;
		public UserModel pdfPrint;
	}
}
