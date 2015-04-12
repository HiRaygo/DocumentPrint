/*
 * 由SharpDevelop创建。
 * 用户： XiaoSanya
 * 日期: 2015/4/11
 * 时间: 19:06
 */
namespace DocumentPrint
{
	partial class MainForm
	{
		/// <summary>
		/// Designer variable used to keep track of non-visual components.
		/// </summary>
		private System.ComponentModel.IContainer components = null;
		
		/// <summary>
		/// Disposes resources used by the form.
		/// </summary>
		/// <param name="disposing">true if managed resources should be disposed; otherwise, false.</param>
		protected override void Dispose(bool disposing)
		{
			if (disposing) {
				if (components != null) {
					components.Dispose();
				}
			}
			base.Dispose(disposing);
		}
		
		/// <summary>
		/// This method is required for Windows Forms designer support.
		/// Do not change the method contents inside the source code editor. The Forms designer might
		/// not be able to load this method if it was changed manually.
		/// </summary>
		private void InitializeComponent()
		{
			this.groupBox1 = new System.Windows.Forms.GroupBox();
			this.buttonStart = new System.Windows.Forms.Button();
			this.label2 = new System.Windows.Forms.Label();
			this.textBoxInterval = new System.Windows.Forms.TextBox();
			this.checkBoxError = new System.Windows.Forms.CheckBox();
			this.checkBoxDelete = new System.Windows.Forms.CheckBox();
			this.buttonSelect = new System.Windows.Forms.Button();
			this.labelFolder = new System.Windows.Forms.Label();
			this.label1 = new System.Windows.Forms.Label();
			this.textBoxInfo = new System.Windows.Forms.TextBox();
			this.groupBox1.SuspendLayout();
			this.SuspendLayout();
			// 
			// groupBox1
			// 
			this.groupBox1.Controls.Add(this.buttonStart);
			this.groupBox1.Controls.Add(this.label2);
			this.groupBox1.Controls.Add(this.textBoxInterval);
			this.groupBox1.Controls.Add(this.checkBoxError);
			this.groupBox1.Controls.Add(this.checkBoxDelete);
			this.groupBox1.Controls.Add(this.buttonSelect);
			this.groupBox1.Controls.Add(this.labelFolder);
			this.groupBox1.Controls.Add(this.label1);
			this.groupBox1.Location = new System.Drawing.Point(12, 12);
			this.groupBox1.Name = "groupBox1";
			this.groupBox1.Size = new System.Drawing.Size(510, 87);
			this.groupBox1.TabIndex = 1;
			this.groupBox1.TabStop = false;
			// 
			// buttonStart
			// 
			this.buttonStart.Location = new System.Drawing.Point(403, 52);
			this.buttonStart.Name = "buttonStart";
			this.buttonStart.Size = new System.Drawing.Size(101, 23);
			this.buttonStart.TabIndex = 7;
			this.buttonStart.Text = "开始打印";
			this.buttonStart.UseVisualStyleBackColor = true;
			this.buttonStart.Click += new System.EventHandler(this.ButtonStartClick);
			// 
			// label2
			// 
			this.label2.Location = new System.Drawing.Point(114, 52);
			this.label2.Name = "label2";
			this.label2.Size = new System.Drawing.Size(30, 20);
			this.label2.TabIndex = 6;
			this.label2.Text = "秒";
			this.label2.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// textBoxInterval
			// 
			this.textBoxInterval.Location = new System.Drawing.Point(66, 51);
			this.textBoxInterval.Name = "textBoxInterval";
			this.textBoxInterval.Size = new System.Drawing.Size(40, 21);
			this.textBoxInterval.TabIndex = 5;
			this.textBoxInterval.Text = "30";
			// 
			// checkBoxError
			// 
			this.checkBoxError.Location = new System.Drawing.Point(293, 51);
			this.checkBoxError.Name = "checkBoxError";
			this.checkBoxError.Size = new System.Drawing.Size(104, 24);
			this.checkBoxError.TabIndex = 4;
			this.checkBoxError.Text = "出错后停止";
			this.checkBoxError.UseVisualStyleBackColor = true;
			// 
			// checkBoxDelete
			// 
			this.checkBoxDelete.Location = new System.Drawing.Point(153, 52);
			this.checkBoxDelete.Name = "checkBoxDelete";
			this.checkBoxDelete.Size = new System.Drawing.Size(116, 24);
			this.checkBoxDelete.TabIndex = 3;
			this.checkBoxDelete.Text = "打印后删除文件";
			this.checkBoxDelete.UseVisualStyleBackColor = true;
			// 
			// buttonSelect
			// 
			this.buttonSelect.Location = new System.Drawing.Point(403, 17);
			this.buttonSelect.Name = "buttonSelect";
			this.buttonSelect.Size = new System.Drawing.Size(101, 23);
			this.buttonSelect.TabIndex = 2;
			this.buttonSelect.Text = "选择文件夹";
			this.buttonSelect.UseVisualStyleBackColor = true;
			this.buttonSelect.Click += new System.EventHandler(this.Button1Click);
			// 
			// labelFolder
			// 
			this.labelFolder.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle;
			this.labelFolder.Location = new System.Drawing.Point(6, 19);
			this.labelFolder.Name = "labelFolder";
			this.labelFolder.Size = new System.Drawing.Size(391, 20);
			this.labelFolder.TabIndex = 1;
			this.labelFolder.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// label1
			// 
			this.label1.Location = new System.Drawing.Point(6, 52);
			this.label1.Name = "label1";
			this.label1.Size = new System.Drawing.Size(65, 20);
			this.label1.TabIndex = 0;
			this.label1.Text = "扫描周期：";
			this.label1.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// textBoxInfo
			// 
			this.textBoxInfo.Location = new System.Drawing.Point(12, 107);
			this.textBoxInfo.Multiline = true;
			this.textBoxInfo.Name = "textBoxInfo";
			this.textBoxInfo.ReadOnly = true;
			this.textBoxInfo.Size = new System.Drawing.Size(510, 215);
			this.textBoxInfo.TabIndex = 2;
			// 
			// MainForm
			// 
			this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 12F);
			this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
			this.ClientSize = new System.Drawing.Size(534, 332);
			this.Controls.Add(this.textBoxInfo);
			this.Controls.Add(this.groupBox1);
			this.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedSingle;
			this.MaximizeBox = false;
			this.Name = "MainForm";
			this.Text = "Main";
			this.groupBox1.ResumeLayout(false);
			this.groupBox1.PerformLayout();
			this.ResumeLayout(false);
			this.PerformLayout();
		}
		private System.Windows.Forms.Button buttonStart;
		private System.Windows.Forms.Label label2;
		private System.Windows.Forms.CheckBox checkBoxError;
		private System.Windows.Forms.TextBox textBoxInterval;
		private System.Windows.Forms.CheckBox checkBoxDelete;
		private System.Windows.Forms.Button buttonSelect;
		private System.Windows.Forms.TextBox textBoxInfo;
		private System.Windows.Forms.Label labelFolder;
		private System.Windows.Forms.GroupBox groupBox1;
		private System.Windows.Forms.Label label1;
	}
}
