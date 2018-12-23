/*
 * 用户：Jason
 * 日期: 2018/12/23
 * 时间: 1:07
 */
using System;
using System.Text;
using System.Windows;
using System.Diagnostics;
using System.Windows.Forms;
using System.Windows.Automation;
using System.Text.RegularExpressions;
using System.Threading;

namespace demo2
{
	/// <summary>
	/// Description of OutlookUIA.
	/// </summary>
	public class OutlookUIA
	{
		AutomationElement app_ = null;
		string outlookSearchFrameName_ = "搜索查询";
		string outlookSearchResultStatus_ = "此文件夹中的项目数";

		public OutlookUIA()
		{
		}
		
		bool FindApp()
		{
			// 我的outlook没有激活, 会弹窗, 要先关闭弹窗
			if (app_ == null)
			{
				// TODO -- kill outlook
				Process process = Process.Start("Outlook");
				UIA.WaitProcess(process.Id);
            	app_ = AutomationElement.FromHandle(process.MainWindowHandle);
			}
			
			/*
            StringBuilder sb2 = new StringBuilder();
			AutomationElementCollection ac2 = app_.FindAll(TreeScope.Descendants, 
               	new PropertyCondition(AutomationElement.ClassNameProperty, "NetUITextbox"));
			for (int i = 0; i < ac2.Count; ++i)
			{
				AutomationElement wd = ac2[i];
				sb2.Append("name: " + wd.Current.Name + ", classname: " + wd.Current.ClassName + "\n");
			}			
			MessageBox.Show(sb2.ToString());
			
			if (app_ != null)
			{
				MessageBox.Show(app_.Current.Name + " " + app_.Current.ClassName);
			}
			*/
			
			return app_ == null ? false : true;
		}
		
		AutomationElement FindElement(string elementName)
		{
			if (!FindApp())
			{
				MessageBox.Show("found app failed\n");
				return null;
			}
            
			AutomationElement searchElement = app_.FindFirst(TreeScope.Descendants,
				new PropertyCondition(AutomationElement.NameProperty, elementName));
			
			return searchElement;
		}
		
		
		AutomationElement FindSearchResultElement()
		{
			if (!FindApp())
			{
				MessageBox.Show("found app failed\n");
				return null;
			}
            
			AutomationElement searchElement = app_.FindFirst(TreeScope.Descendants,
				new PropertyCondition(AutomationElement.HelpTextProperty,
			                          outlookSearchResultStatus_));
			
			return searchElement;
		}
		
		int GetInBoxMailNumber()
		{
			AutomationElement result = FindSearchResultElement();
			if (result == null)
			{
				MessageBox.Show("GetInBoxMailNumber failed\n");
				return 0;
			}
			
			try
			{
				string mailNum = System.Text.RegularExpressions.Regex.Replace(
						result.Current.Name, @"[^0-9]+", "");
				return int.Parse(mailNum);
			}
			catch (Exception)
			{
				MessageBox.Show("GetInBoxMailNumber convert string to int failed\n");
				return 0;
			}
		}
		
		public string Search(string info)
		{
			AutomationElement search = FindElement(outlookSearchFrameName_);
			if (search == null)
			{
				MessageBox.Show("found failed\n");
				return "没有找到邮件";
			}
			
			int mailAll = GetInBoxMailNumber();

			UIA.SetValue(search, info);
			
			// 等待输入完成
			Thread.Sleep(2000);

			int mailSearchAll = GetInBoxMailNumber();
			
			return string.Format("找到{0}/{1}封邮件", mailSearchAll, mailAll);
		}
	}
}
