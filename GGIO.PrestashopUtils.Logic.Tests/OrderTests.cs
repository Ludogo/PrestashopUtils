using System;
using System.IO;
using System.Threading.Tasks;
using Bukimedia.PrestaSharp.Factories;
using Microsoft.VisualStudio.TestTools.UnitTesting;

namespace GGIO.PrestashopUtils.Logic.Tests
{
    [TestClass]
    public class OrderTests
    {
        [TestMethod]
        public async Task ProductImportTest()
        {
            string url = "https://poubelle.boutique.coop/api";
            string account = "CLEQ2RIEPIN53WP4GSMBNFU1RRTRD8AK";
            var orderFactory = new OrderFactory(url,account, null,null);
            var invoiceFactory = new OrderInvoiceFactory(url, account, null,null);

            var service = new OrderService(orderFactory, invoiceFactory);

            await service.GetOrder(9);
        }
    }
}
