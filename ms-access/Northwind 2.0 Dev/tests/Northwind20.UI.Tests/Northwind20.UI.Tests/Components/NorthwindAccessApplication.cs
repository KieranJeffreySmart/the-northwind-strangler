using Microsoft.Office.Interop.Access;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Automation;

namespace Northwind20.UI.Tests.Components
{
    public class NorthwindAccessApplication
    {
        Process appProcess;

        AutomationElement appElement;

        public static NorthwindAccessApplication NewApplication(string path = @"../../../../../..\Database1.accdb")
        {
            Process p = new Process();
            p.StartInfo.FileName = "explorer";
            p.StartInfo.Arguments = "\"" + Path.GetFullPath(path) + "\"";
            p.Start();
            Thread.Sleep(5000);
            return new NorthwindAccessApplication(p);
        }

        public NorthwindAccessApplication(Process process)
        {
            appProcess = process;

            Condition propCondition = new PropertyCondition(
                AutomationElement.NameProperty, "Northwind Traders 2.0 Developer Edition", PropertyConditionFlags.IgnoreCase);

            AutomationElement rootElement = AutomationElement.RootElement;
            appElement = rootElement.FindFirst(TreeScope.Children, propCondition);
        }

        public bool LoginFormIsOpen { get { return appElementsByNameExists("Login"); } }

        private bool appElementsByNameExists(string name)
        {
            Condition propCondition = new PropertyCondition(
                AutomationElement.NameProperty, name, PropertyConditionFlags.IgnoreCase);

            return appElement.FindFirst(TreeScope.Children, propCondition) != null;
        }

        private AutomationElement appElementByName(string name)
        {
            Condition propCondition = new PropertyCondition(
                AutomationElement.NameProperty, name, PropertyConditionFlags.IgnoreCase);

            return appElement.FindFirst(TreeScope.Children, propCondition);
        }
        private IEnumerable<AutomationElement> childElementsByName(AutomationElement parent, string name)
        {
            Condition propCondition = new PropertyCondition(
                AutomationElement.NameProperty, name, PropertyConditionFlags.IgnoreCase);

            return parent.FindAll(TreeScope.Children, propCondition).Cast<AutomationElement>();
        }

        internal void Login(string userName)
        {
            var loginWindow = appElementByName("Login");

            if (userName != null)
            {

            }

            var loginButton = childElementsByName(loginWindow, "Login").FirstOrDefault(e => e.Current.ControlType == ControlType.Button);
            invokeClick(loginButton);
        }

        private void invokeClick(AutomationElement? element)
        {
            element.TryGetCurrentPattern(InvokePattern.Pattern, out var objPattern);
            InvokePattern invPattern = objPattern as InvokePattern;
            if (invPattern != null)
            {
                invPattern.Invoke();
            }
        }

        private WindowPattern getWindowPattern(AutomationElement targetControl)
        {
            WindowPattern windowPattern = null;

            try
            {
                windowPattern = targetControl.GetCurrentPattern(WindowPattern.Pattern) as WindowPattern;
            }
            catch (InvalidOperationException)
            {
                // object doesn't support the WindowPattern control pattern
                return null;
            }
            // Make sure the element is usable.
            if (false == windowPattern.WaitForInputIdle(10000))
            {
                // Object not responding in a timely manner
                return null;
            }
            return windowPattern;
        }

        public void CloseAppWindow()
        {
            var appWindowPattern = getWindowPattern(appElement);
            appWindowPattern.Close();
            appProcess.Kill();
            var msa = Process.GetProcessesByName("MSACCESS");
            foreach (Process msAccess in msa)
            {
                    msAccess.Kill();
            }            
        }
    }
}
