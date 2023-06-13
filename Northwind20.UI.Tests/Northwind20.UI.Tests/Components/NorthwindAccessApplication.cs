using Microsoft.Office.Interop.Access;
using Newtonsoft.Json.Linq;
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

        public static NorthwindAccessApplication NewApplication(string path = @"../../../TestData/Database1.accdb")
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

        public bool LoginFormIsOpen { get { return appElementExistsByName("Login"); } }

        public bool DatabaseLoginFormIsOpen { get { return appElementExistsByName("SQL Server Login"); } }

        public bool OrdersListIsOpen { 
            get {

                var randomPane = getAppElementByName("ODocTabs").FindFirst(TreeScope.Children, Condition.TrueCondition);
                var firstDocTabs = getChildElementsByName(randomPane, "Document Tabs").First();
                var secondDocTabs = getChildElementsByName(firstDocTabs, "Document Tabs").First();
                var isOpenWithFocus = childElementExistsByName(secondDocTabs, @"'Orders' Tab");
                var isOpen = isOpenWithFocus || childElementExistsByName(secondDocTabs, @"Orders Tab");

                return isOpen;
            } 
        }

        public bool ApplicationDialogIsOpen { get { return appElementExistsByName("Northwind Traders 2.0 Developer Edition"); } }

        internal void UserLogin(string? userName = null)
        {
            var loginWindow = getAppElementByName("Login");

            if (userName != null)
            {
                var userNameCombo = getChildElementsByName(loginWindow, "Detail Section, Select Employee: collapsed").FirstOrDefault(e => e.Current.ControlType == ControlType.ComboBox);
                changeTextValue(userNameCombo, userName);
            }

            var loginButton = getChildElementsByName(loginWindow, "Login").FirstOrDefault(e => e.Current.ControlType == ControlType.Button);
            invokeClick(loginButton);
        }

        internal void DatabaseLogin(string? userName = null, string? password = null)
        {
            var loginWindow = getAppElementByName("SQL Server Login");

            if (userName != null)
            {
                var userNameEdit = getChildElementsByName(loginWindow, "Login ID:").FirstOrDefault(e => e.Current.ControlType == ControlType.Edit);
                changeTextValue(userNameEdit, userName);
            }

            if (password != null)
            {
                var passwordEdit = getChildElementsByName(loginWindow, "Password:").FirstOrDefault(e => e.Current.ControlType == ControlType.Edit);
                changeTextValue(passwordEdit, password);
            }

            var loginButton = getChildElementsByName(loginWindow, "OK").FirstOrDefault(e => e.Current.ControlType == ControlType.Button);
            invokeClick(loginButton);
        }

        private bool appElementExistsByName(string name)
        {
            Condition propCondition = new PropertyCondition(
                AutomationElement.NameProperty, name, PropertyConditionFlags.IgnoreCase);

            return appElement.FindFirst(TreeScope.Children, propCondition) != null;
        }

        private bool childElementExistsByName(AutomationElement parent, string name)
        {
            Condition propCondition = new PropertyCondition(
                AutomationElement.NameProperty, name, PropertyConditionFlags.IgnoreCase);

            return parent.FindFirst(TreeScope.Children, propCondition) != null;
        }

        private AutomationElement getAppElementByName(string name)
        {
            Condition propCondition = new PropertyCondition(
                AutomationElement.NameProperty, name, PropertyConditionFlags.IgnoreCase);

            return appElement.FindFirst(TreeScope.Children, propCondition);
        }
        private IEnumerable<AutomationElement> getChildElementsByName(AutomationElement parent, string name)
        {
            Condition propCondition = new PropertyCondition(
                AutomationElement.NameProperty, name, PropertyConditionFlags.IgnoreCase);

            return parent.FindAll(TreeScope.Children, propCondition).Cast<AutomationElement>();
        }

        private IEnumerable<AutomationElement> getChildElementsByControlType(AutomationElement parent, ControlType controlType)
        {
            Condition propCondition = new PropertyCondition(
                AutomationElement.ControlTypeProperty, controlType);

            return parent.FindAll(TreeScope.Children, propCondition).Cast<AutomationElement>();
        }


        private void changeTextValue(AutomationElement? element, string value)
        {
            // A series of basic checks prior to attempting an insertion.
            //
            // Check #1: Is control enabled?
            // An alternative to testing for static or read-only controls
            // is to filter using
            // PropertyCondition(AutomationElement.IsEnabledProperty, true)
            // and exclude all read-only text controls from the collection.
            if (!element.Current.IsEnabled)
            {
                throw new InvalidOperationException(
                    "The control with an AutomationID of "
                    + element.Current.AutomationId.ToString()
                    + " is not enabled.\n\n");
            }

            // Check #2: Are there styles that prohibit us
            //           from sending text to this control?
            if (!(element.Current.IsKeyboardFocusable || element.Current.ControlType == ControlType.ComboBox))
            {
                throw new InvalidOperationException(
                    "The control with an AutomationID of "
                    + element.Current.AutomationId.ToString()
                    + "is read-only.\n\n");
            }

            // Once you have an instance of an AutomationElement,
            // check if it supports the ValuePattern pattern.
            object valuePattern = null;

            // Control does not support the ValuePattern pattern
            // so use keyboard input to insert content.
            //
            //
            if (!element.TryGetCurrentPattern(
                ValuePattern.Pattern, out valuePattern))
            {
                throw new InvalidOperationException(
                    "The control with an AutomationID of "
                    + element.Current.AutomationId.ToString()
                    + "does not support value pattern.\n\n");
            }
            // Control supports the ValuePattern pattern so we can
            // use the SetValue method to insert content.
            else
            {
                // Set focus for input functionality and begin.
                element.SetFocus();

                ((ValuePattern)valuePattern).SetValue(value);
            }
        }

        private bool elementIsKeyboardFocusableCombobox(AutomationElement? element)
        {
            if (element.Current.ControlType != ControlType.ComboBox) return false;
            return getChildElementsByControlType(element, ControlType.Edit).FirstOrDefault()?.Current.IsKeyboardFocusable ?? false;
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
