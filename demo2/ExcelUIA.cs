/*
 * 用户：Jason
 * 日期: 2018/12/21
 * 时间: 22:41
 */
using System;
using System.Text;
using System.Windows;
using System.Windows.Forms;
using System.Windows.Automation;

namespace demo2
{
	class ExcelUIA
	{
		AutomationElement app_ = null;
		string excelFileName_ = null;
		
		public ExcelUIA(string excelFilename)
		{
			excelFileName_ = excelFilename;
			// 需要手动打开excel，demo没有实现判断文件自动打开
			//excelFileName_ = Path.GetFileName(this.textBoxExcelPath.Text);
		}
		
		bool FindApp()
		{			
			/* 本机的Name格式(filename - Excel), 如下所示:
				demo.xlsx - Excel
				新建 Microsoft Excel 工作表.xlsx - Excel
			 */
			if (app_ == null)
			{
           		PropertyCondition nameProperty = new PropertyCondition(
					AutomationElement.NameProperty, 
					excelFileName_ + " - Excel");
            	//PropertyCondition typeProperty = new PropertyCondition(AutomationElement.ControlTypeProperty, ControlType.Window);
            	PropertyCondition typeProperty = new PropertyCondition(
            		AutomationElement.ClassNameProperty, "XLMAIN");
            	AndCondition andCondition = new AndCondition(nameProperty, typeProperty);
            	app_ = AutomationElement.RootElement.FindFirst(TreeScope.Children, andCondition);
			}
			
			return app_ == null ? false : true;
		}
		
		AutomationElement FindExcelCell(string cell)
		{
			if (!FindApp())
			{
				return null;
			}
            
			AutomationElement cellElement = app_.FindFirst(TreeScope.Descendants,
				new PropertyCondition(AutomationElement.NameProperty, cell));
			
			return cellElement;
		}
		
		private GridItemPattern GetGridItemPattern(
		    AutomationElement targetControl)
		{
		    GridItemPattern gridItemPattern = null;
		
		    try
		    {
		        gridItemPattern =
		            targetControl.GetCurrentPattern(
		            GridItemPattern.Pattern)
		            as GridItemPattern;
		    }
		    // Object doesn't support the 
		    // GridPattern control pattern
		    catch (InvalidOperationException)
		    {
		        return null;
		    }
		
		    return gridItemPattern;
		}
		
		public GridPattern GetGridPattern(AutomationElement element)
		{
		    object currentPattern;
		    if (!element.TryGetCurrentPattern(GridPattern.Pattern,
		                                      out currentPattern))
		    {
		        throw new Exception(
		    		string.Format(
		    			"Element with AutomationId '{0}' and Name '{1}' "
		    			+ "does not support the GridPattern.",
		            element.Current.AutomationId, element.Current.Name));
		    }
		    
		    return currentPattern as GridPattern;
		}
		
		public bool SetValue(string cell, string value)
		{
			AutomationElement cellElement = FindExcelCell(cell);
			if (cellElement == null)
			{
				MessageBox.Show("取窗口失败, demo下excel需要手动打开文件测试");
				return false;
			}
            
			/*
			MessageBox.Show("set " + cell + " " + value);
			MessageBox.Show(cellElement.Current.Name);		
            StringBuilder sb = new StringBuilder();
            AutomationPattern[] patterns = cellElement.GetSupportedPatterns();
			AutomationElementCollection ac = AutomationElement.RootElement.FindAll(TreeScope.Children, Condition.TrueCondition);
			foreach (AutomationPattern pattern in patterns)
			{
				sb.Append("ProgrammaticName: " + pattern.ProgrammaticName 
				           + ", classname: " + "PatternName: " 
				           + Automation.PatternName(pattern) + "\n");
			}			
			MessageBox.Show(sb.ToString());
			*/

			ValuePattern valuePattern
				= (ValuePattern)cellElement.GetCurrentPattern(ValuePatternIdentifiers.Pattern);

			// 设置的元素值可读但不可见
			valuePattern.SetValue(value);
			
			return true;
		}
		
		public bool GetValue(string cell, ref string value)
		{		
			AutomationElement cellElement = FindExcelCell(cell);
			if (cellElement == null)
			{
				MessageBox.Show("取窗口失败, demo下excel需要手动打开文件测试");
				return false;
			}
			
			ValuePattern valuePattern
				= (ValuePattern)cellElement.GetCurrentPattern(ValuePatternIdentifiers.Pattern);
			
			value = valuePattern.Current.Value.ToString();
			
			return true;
		}
		
        public void Close()
        {
        }
	}
}
