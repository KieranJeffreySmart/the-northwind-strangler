using Northwind20.UI.Tests.Components;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Northwind20.UI.Tests
{
    public class NewOrders
    {

        [Fact]
        public void PlaceAnOrderAndInvoice()
        {

            // Given a customer has placed an order for 100 units of beer
            var customer = "Adatum Corporation";
            var shippingFee = 10;
            var productCategory = "Beverages";
            var product = "Beer";
            var quantity = 100;

            // When I Create an invoice
            NorthwindAccessApplication app = NorthwindAccessApplication.NewApplication();
            try
            {
                Login(app);
                var orderDetailForm = app.StartNewOrder();
                Assert.NotNull(orderDetailForm);
                orderDetailForm.SetCustomer(customer);
                orderDetailForm.SetShippingFee(shippingFee);
                orderDetailForm.AddItem(productCategory, product, quantity);
                orderDetailForm.CreateInvoice();
                Thread.Sleep(10000);
                Assert.True(app.InvoiceReportIsOpen);
            }
            finally
            {
                app.CloseAppWindow();
            }
        }

        private void Login(NorthwindAccessApplication app)
        {
            app.DatabaseLogin("NorthwindClient", "N0rthw1nd2");
            Thread.Sleep(500);
            Assert.True(app.LoginFormIsOpen);
            app.UserLogin();
            // Then a report should appear, ready for printing
            Thread.Sleep(500);
            Assert.False(app.LoginFormIsOpen);
            Assert.True(app.OrdersListIsOpen);
        }
    }
}
