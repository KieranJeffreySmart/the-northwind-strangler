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
using System.Xml.Linq;

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

        public bool LoginFormIsOpen { get { return appElement.ChildElementExistsByName("Login"); } }

        public bool DatabaseLoginFormIsOpen { get { return appElement.ChildElementExistsByName("SQL Server Login"); } }

        public bool ApplicationDialogIsOpen { get { return appElement.ChildElementExistsByName("Northwind Traders 2.0 Developer Edition"); } }

        public bool InvoiceReportIsOpen { get; internal set; }

        public bool OrdersListIsOpen {
            get {
                var workspace = GetWorkspace();
                var documentsTabBar = GetDocumentsTabBar();
                if (workspace == null || documentsTabBar == null)
                    return false;

                var tabIsDisplayed = documentsTabBar.ChildElementExistsByName(@"'Orders' Tab") || documentsTabBar.ChildElementExistsByName(@"Orders Tab");
                var ordersListIsDisplayed = workspace.ChildElementExistsByName("Orders");

                return tabIsDisplayed && ordersListIsDisplayed;
            }
        }

        public void UserLogin(string? userName = null)
        {
            var loginWindow = appElement.GetChildElementsByName("Login").First();

            if (userName != null)
            {
                var userNameCombo = loginWindow.GetChildElementsByName("Detail Section, Select Employee: collapsed").First(e => e.Current.ControlType == ControlType.ComboBox);
                userNameCombo.ChangeTextValue(userName);
            }

            var loginButton = loginWindow.GetChildElementsByName("Login").First(e => e.Current.ControlType == ControlType.Button);
            loginButton.InvokeClick();
        }

        public void DatabaseLogin(string? userName = null, string? password = null)
        {
            var loginWindow = appElement.GetChildElementsByName("SQL Server Login").First();

            if (userName != null)
            {
                var userNameEdit = loginWindow.GetChildElementsByName("Login ID:").First(e => e.Current.ControlType == ControlType.Edit);
                userNameEdit.ChangeTextValue(userName);
            }

            if (password != null)
            {
                var passwordEdit = loginWindow.GetChildElementsByName("Password:").First(e => e.Current.ControlType == ControlType.Edit);
                passwordEdit.ChangeTextValue(password);
            }

            var loginButton = loginWindow.GetChildElementsByName("OK").First(e => e.Current.ControlType == ControlType.Button);
            loginButton.InvokeClick();
        }

        public void CloseAppWindow()
        {
            var appWindowPattern = appElement.GetWindowPattern();
            appWindowPattern.Close();
            appProcess.Kill();
            var msa = Process.GetProcessesByName("MSACCESS");
            foreach (Process msAccess in msa)
            {
                msAccess.Kill();
            }
        }

        public OrderDetailsForm StartNewOrder()
        {
            var workspace = GetWorkspace();
            if (workspace == null || !workspace.ChildElementExistsByName("Orders"))
                throw new Exception("Invalid Command StartNewOrder. Error: Orders List Not Open");

            AutomationElement ordersList = workspace.GetChildElementsByName("Orders").First();
            var addOrderButton = ordersList.GetChildElementsByName("Add Orders").First();
            addOrderButton.InvokeClick();
            Thread.Sleep(1000);
            var orderForm = workspace.GetChildElementsByName("Order").First();
            return new OrderDetailsForm(orderForm);
        }

        private AutomationElement? GetWorkspace()
        {
            return appElement.GetChildElementsByName("").FirstOrDefault();
        }

        private AutomationElement? GetDocumentsTabBar()
        {
            var odocs = appElement.GetChildElementsByName("ODocTabs").FirstOrDefault();
            var docsTab = odocs?.GetFirstDescendentElementByNameAndType("Document Tabs", ControlType.Tab);

            return docsTab;
        }
    }
}