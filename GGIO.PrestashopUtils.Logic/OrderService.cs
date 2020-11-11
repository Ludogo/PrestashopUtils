using System.Threading.Tasks;
using Bukimedia.PrestaSharp.Factories;

namespace GGIO.PrestashopUtils.Logic
{
    public class OrderService
    {
        private readonly OrderFactory orderFactory;
        private readonly OrderInvoiceFactory orderInvoiceFactory;

        public OrderService(OrderFactory orderFactory, OrderInvoiceFactory orderInvoiceFactory)
        {
            this.orderFactory = orderFactory;
            this.orderInvoiceFactory = orderInvoiceFactory;
        }
        public async Task GetOrder(long orderId)
        {
            var res = await orderFactory.GetAsync(orderId);
            var invoice = await orderInvoiceFactory.GetAsync(res.invoice_number);

        }
    }
}