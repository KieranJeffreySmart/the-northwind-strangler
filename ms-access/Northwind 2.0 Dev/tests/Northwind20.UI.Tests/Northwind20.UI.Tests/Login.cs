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
        public void Login_DefaultUser()
        {
            NorthwindAccessApplication app = NorthwindAccessApplication.NewApplication();
            Assert.True(app.LoginFormIsOpen);

            app.Login();
            Assert.False(app.LoginFormIsOpen);

            app.CloseAppWindow();
        }

        [Fact]
        public void Login_WithValidUser()
        {
            NorthwindAccessApplication app = NorthwindAccessApplication.NewApplication();
            Assert.True(app.LoginFormIsOpen);

            app.Login("Andrew Cencini");
            Assert.False(app.LoginFormIsOpen);

            app.CloseAppWindow();
        }

        [Fact]
        public void Login_WithInvalidUser()
        {
            NorthwindAccessApplication app = NorthwindAccessApplication.NewApplication();
            Assert.True(app.LoginFormIsOpen);

            app.Login();
            Assert.False(app.LoginFormIsOpen);

            app.CloseAppWindow();
        }
    }
}