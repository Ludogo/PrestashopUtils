using System.Linq;
using System.Threading.Tasks;
using Bukimedia.PrestaSharp.Entities;
using Bukimedia.PrestaSharp.Factories;
using ClosedXML;
using ClosedXML.Excel;

namespace GGIO.PrestashopUtils.Logic
{
    public class ProductImportService
    {
        public ProductFactory ProductClient { get; }
        public CategoryFactory CategoryFactory { get; }
        public ProductImportService(ProductFactory productClient, CategoryFactory categoryFactory)
        {
            this.CategoryFactory = categoryFactory;
            this.ProductClient = productClient;

        }

        public async Task ImportAsync(System.IO.Stream file, int shopId)
        {
            using (var workbook = new XLWorkbook(file))
            {
                var categories = await CategoryFactory.GetAllAsync();

                foreach (var row in workbook.Worksheets.First().Rows())
                {
                    int index = 1;
                    var category = categories.FirstOrDefault(x => x.name.Any(j => j.Value == row.Cell(index++).GetValue<string>()) && x.id_shop_default == shopId);
                    var product = new product();

                    product.id_shop_default = shopId;
                    product.id_category_default = category.id;
                    product.description.Add(new Bukimedia.PrestaSharp.Entities.AuxEntities.language { id = 1, Value = row.Cell(index++).GetValue<string>() });


                    product.associations.categories.Add(new Bukimedia.PrestaSharp.Entities.AuxEntities.category { id = category.id.Value });
                    //to do : traiter multi valeur de caracteristiques
                    product.associations.product_features.Add(new Bukimedia.PrestaSharp.Entities.AuxEntities.product_feature { id_feature_value = row.Cell(index++).GetValue<int>() });
                    product.id_manufacturer = row.Cell(index++).GetValue<int>();
                    product.reference = row.Cell(index++).GetValue<string>();
                    product.id_shop_default = row.Cell(index++).GetValue<int>();
                    //to do : traiter le sotck product.quan
                    index++;
                    product.price = row.Cell(index++).GetValue<decimal>();
                    index++;
                    product.id_tax_rules_group = row.Cell(index++).GetValue<int>();
                    product.location = row.Cell(index++).GetValue<string>();
                    


                }
            }
        }
    }
}