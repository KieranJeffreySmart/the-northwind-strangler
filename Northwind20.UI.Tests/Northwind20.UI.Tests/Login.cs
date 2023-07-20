using System.Windows.Automation;
using System.Linq;
using System.Windows.Automation.Peers;
using System;
using Northwind20.UI.Tests.Components;

namespace Northwind20.UI.Tests
{
    public class Login
    {
        [Fact]
        public void DatabaseLogin_DefaultClient()
        {
            NorthwindAccessApplication app = NorthwindAccessApplication.NewApplication();
            try
            {
                Assert.True(app.DatabaseLoginFormIsOpen);

                app.DatabaseLogin("NorthwindClient", "N0rthw1nd2");
                Thread.Sleep(500);
                Assert.False(app.DatabaseLoginFormIsOpen);
                Assert.True(app.LoginFormIsOpen);
            }
            finally
            {
                app.CloseAppWindow();
            }
        }

        [Fact]
        public void Login_DefaultUser()
        {
            NorthwindAccessApplication app = NorthwindAccessApplication.NewApplication();
            try
            {
                app.DatabaseLogin("NorthwindClient", "N0rthw1nd2");
                Thread.Sleep(500);
                Assert.True(app.LoginFormIsOpen);

                app.UserLogin();
                Thread.Sleep(500);
                Assert.False(app.LoginFormIsOpen);
                Assert.True(app.OrdersListIsOpen);
            }
            finally
            {
                app.CloseAppWindow();
            }
        }

        [Fact]
        public void Login_WithValidUser()
        {
            NorthwindAccessApplication app = NorthwindAccessApplication.NewApplication();
            try
            {
                app.DatabaseLogin("NorthwindClient", "N0rthw1nd2");
                Thread.Sleep(500);
                Assert.True(app.LoginFormIsOpen);
                
                app.UserLogin("Andrew Cencini");
                Thread.Sleep(500);
                Assert.False(app.LoginFormIsOpen);
                Assert.True(app.OrdersListIsOpen);
            }
            finally
            {
                app.CloseAppWindow();
            }
        }

        [Fact]
        public void Login_WithInvalidUser()
        {
            NorthwindAccessApplication app = NorthwindAccessApplication.NewApplication();
            try
            {
                app.DatabaseLogin("NorthwindClient", "N0rthw1nd2");
                Thread.Sleep(500);
                Assert.True(app.LoginFormIsOpen);

                app.UserLogin("Non User");
                Thread.Sleep(500);
                Assert.True(app.LoginFormIsOpen);
                Assert.True(app.ApplicationDialogIsOpen);
            }
            finally
            {
                app.CloseAppWindow();
            }
        }
    }
}