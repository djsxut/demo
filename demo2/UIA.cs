/*
 * 用户：Jason
 * 日期: 2018/12/23
 * 时间: 1:02
 */
using System;
using System.Windows.Forms;
using System.Windows.Automation;
using System.Threading;

namespace demo2
{
	/// <summary>
	/// Description of UIA.
	/// </summary>
	public class UIA
	{
		public UIA()
		{
		}
		
		public static void SetValue(AutomationElement element, string value)
		{
			ValuePattern valuePattern = (ValuePattern)element.GetCurrentPattern(ValuePattern.Pattern);
			valuePattern.SetValue(value);
		}
		
		public static void GetValue(AutomationElement element, ref string value)
		{
			TextPattern valuePattern = (TextPattern)element.GetCurrentPattern(TextPattern.Pattern);
			value = valuePattern.DocumentRange.GetText(-1);
		}
		
		public static void Invoke(AutomationElement element)
		{
            //获取按钮的invokepattern接口
            var invokeptn = element.GetCurrentPattern(InvokePattern.Pattern) as InvokePattern;
            //调用invoke接口，完成按钮点击
            invokeptn.Invoke();
		}
		
		public static void WaitProcess(int pid)
		{
			// 简单实现, 等待5s
			Thread.Sleep(5000);
		}
	}
}
