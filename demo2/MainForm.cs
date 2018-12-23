/*
 * 用户：Jason
 * 日期: 2018/12/22
 * 时间: 23:12
 * 
 */
using System;
using System.IO;
using System.Collections.Generic;
using System.Drawing;
using System.Diagnostics;
using System.Windows.Forms;
using System.Text.RegularExpressions;

namespace demo2
{
	public partial class MainForm : Form
	{
		public MainForm()
		{
			//
			// The InitializeComponent() call is required for Windows Forms designer support.
			//
			InitializeComponent();
			
			//
			// TODO: Add constructor code after the InitializeComponent() call.
			//
		}
		
		bool CheckExcelRowColValid()
        {
            try 
            {
                int row = Convert.ToInt32(this.textBoxExcelRow.Text);
                if (0 >= row) 
                {
                    throw new Exception("row of excel must a positive integer");;
                }
            } 
            catch (Exception) 
            {
                return false;
            }
                    
            string pattern=@"^[A-Za-z]+$";
            Regex  regex = new Regex(pattern);
            return regex.IsMatch(this.textBoxExcelCol.Text);
        }
		
		void ButtonExcelWriteClick(object sender, EventArgs e)
		{
	        if (!CheckExcelRowColValid()) 
            {
                MessageBox.Show("please input excel row(1,2,3...) and col(A,B,C...)");
                return;
            }
			
	        string cell = this.textBoxExcelCol.Text + this.textBoxExcelRow.Text;
	        string filename = Path.GetFileName(this.textBoxExcelPath.Text);
	        
	        ExcelUIA excel = new ExcelUIA(filename);
	        if (excel.SetValue(cell, this.textBoxExcelValue.Text)) 
	        {
	        	MessageBox.Show(cell + "设置成功");
	        }
	        
	        excel.Close();
	        excel = null;
		}

		void ButtonExcelReadClick(object sender, EventArgs e)
		{
	        if (!CheckExcelRowColValid()) 
            {
                MessageBox.Show("please input excel row(1,2,3...) and col(A,B,C...)");
                return;
            }
	        
	        string cell = this.textBoxExcelCol.Text + this.textBoxExcelRow.Text;
	        string filename = Path.GetFileName(this.textBoxExcelPath.Text);
	        
	        ExcelUIA excel = new ExcelUIA(filename);
	        string value = "not found";
	        bool ret = excel.GetValue(cell, ref value);
	        this.textBoxExcelValue.Text = value;
	        
	        if (ret)
	        {
	        	MessageBox.Show(cell + "读取成功: " + value);
	        }
	        
	        excel.Close();
	        excel = null;
		}
		
		void LabelExcelSelectClick(object sender, EventArgs e)
		{
		    string excelPath = "";
            OpenFileDialog op = new OpenFileDialog();
            op.Filter = "excel Files|*.xls;*.xlsx";
            if (op.ShowDialog() == DialogResult.OK)
            {
                excelPath = op.FileName;
            }
            else
            {
                excelPath = "";
            }
            
            this.textBoxExcelPath.Text = excelPath;
		}
		
		void ButtonOpenExcelClick(object sender, EventArgs e)
		{
	        try 
	        {
                System.Diagnostics.Process.Start("excel");
            } 
	        catch (Exception) 
	        {
                MessageBox.Show("Excel not found, start failed!", 
	        	                "start Excel failed", 
	        	                MessageBoxButtons.OK, 
	        	                MessageBoxIcon.Error);
            }
		}
		
		void ButtonOpenOutlookClick(object sender, EventArgs e)
		{
		    try 
	        {
                System.Diagnostics.Process.Start("Outlook");
            } 
	        catch (Exception)
	        {
                MessageBox.Show("Outlook not found, start failed!", 
	        	                "start Outlook failed", 
	        	                MessageBoxButtons.OK, 
	        	                MessageBoxIcon.Error);
            }
		}

		void ButtonMailSearchClick(object sender, EventArgs e)
		{
			if (string.IsNullOrEmpty(this.textBoxMailSubject.Text))
			{
				return;
			}
			
			OutlookUIA outlook = new OutlookUIA();
			this.labelMailInfo.Text = outlook.Search(this.textBoxMailSubject.Text);
		}
	}
}
