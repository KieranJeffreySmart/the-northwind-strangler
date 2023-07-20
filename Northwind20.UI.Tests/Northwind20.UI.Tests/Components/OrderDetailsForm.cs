using System.Windows.Automation;

namespace Northwind20.UI.Tests.Components
{
    public class OrderDetailsForm
    {
        private AutomationElement orderForm;

        public OrderDetailsForm(AutomationElement orderForm)
        {
            this.orderForm = orderForm;
        }

        public void AddItem(string productCategory, string product, int quantity)
        {
            var orderItems = orderForm.GetChildElementsByName("sfrmOrderLineItems").First();
            orderItems.GetChildElementsByName("Detail Section, cboProductCategories collapsed").First().ChangeTextValue(productCategory);
            orderItems.GetChildElementsByName("ProductName").First().ChangeTextValue(product);
            orderItems.GetChildElementsByName("Quantity").First().ChangeTextValue(quantity.ToString());
        }

        public void CreateInvoice()
        {
            orderForm.GetChildElementsByName("1: Create Invoice1: Create Invoice").First().InvokeClick();
        }

        public void SetCustomer(string customer)
        {
            orderForm.GetChildElementsByName("Customer collapsed").First().ChangeTextValue(customer);
        }

        public void SetShippingFee(int shippingFee)
        {            
            orderForm.GetChildElementsByName("Detail Section, 1: Shipping Fee").First().ChangeTextValue($"${shippingFee}.00");
        }
    }
}