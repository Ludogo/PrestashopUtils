using System;
using System.IO;
using System.Threading.Tasks;
using Bukimedia.PrestaSharp.Factories;
using Microsoft.VisualStudio.TestTools.UnitTesting;

namespace GGIO.PrestashopUtils.Logic.Tests
{
    [TestClass]
    public class ImportTests
    {
        [TestMethod]
        public async Task ProductImportTest()
        {
            string url = "https://poubelle.boutique.coop/api";
            string account = "CLEQ2RIEPIN53WP4GSMBNFU1RRTRD8AK";
            var productFactory = new Bukimedia.PrestaSharp.Factories.ProductFactory(url, account, string.Empty);

            var service = new ProductImportService(productFactory, null);

            await service.ImportAsync(File.Open("/home/ludo/Dev/PrestashopUtils/Test Import Presta M20S.xlsx", FileMode.Open), 5);
        }
    }
}
