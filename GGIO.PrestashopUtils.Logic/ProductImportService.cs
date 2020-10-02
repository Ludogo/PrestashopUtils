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
        public StockAvailableFactory StockClient { get; }

        public ProductImportService(ProductFactory productClient, StockAvailableFactory stockClient)
        {
            this.StockClient = stockClient;
            this.ProductClient = productClient;
        }

        public ProductImportService()
        {
        }

        public async Task ImportAsync(System.IO.Stream file, int shopId)
        {
            using (var workbook = new XLWorkbook(file))
            {

                foreach (var row in workbook.Worksheets.First().Rows().Skip(1))
                {
                    if (string.IsNullOrWhiteSpace(row.Cell(1).GetString()))
                    {
                        return;
                    }

                    int index = 1;
                    var product = new product();
                    product.id_shop_default = shopId;
                    product.active = 1;
                    product.state = 1;

                    //A
                    product.AddName(new Bukimedia.PrestaSharp.Entities.AuxEntities.language { id = 1, Value = row.Cell(index++).GetValue<string>() });
                    //B
                    index++;
                    //C
                    product.description.Add(new Bukimedia.PrestaSharp.Entities.AuxEntities.language { id = 1, Value = row.Cell(index++).GetValue<string>() });
                    //D //E
                    product.associations.product_features.Add(new Bukimedia.PrestaSharp.Entities.AuxEntities.product_feature { id = row.Cell(index++).GetValue<int>(), id_feature_value = row.Cell(index++).GetValue<int>() });
                    //to do : traiter multi valeur de caracteristiques


                    //F
                    product.id_manufacturer = row.Cell(index++).GetValue<int>();
                    //G
                    product.reference = row.Cell(index++).GetValue<string>();
                    //H
                    var quantity = row.Cell(index++).GetValue<int>();
                    product.minimal_quantity = quantity;

                    //I
                    product.price = row.Cell(index++).GetValue<decimal>();
                    //J
                    index++;
                    //K
                    product.id_tax_rules_group = row.Cell(index++).GetValue<int>();

                    //L
                    int categoryId = row.Cell(index++).GetValue<int>();


                    //M
                    index++;
                    //N
                    index++;
                    //O
                    index++;
                    //P
                    product.location = row.Cell(index++).GetValue<string>();
                    //Q
                    index++;
                    //R
                    product.width = row.Cell(index++).GetValue<decimal>();

                    //S
                    product.height = row.Cell(index++).GetValue<decimal>();

                    //T
                    product.depth = row.Cell(index++).GetValue<decimal>();

                    //U
                    product.weight = row.Cell(index++).GetValue<decimal>();

                    //V
                    index++;

                    //W
                    index++;

                    //X
                    product.wholesale_price = row.Cell(index++).GetValue<decimal>();

                    //Y
                    index++;

                    //Z
                    product.AddLinkRewrite(new Bukimedia.PrestaSharp.Entities.AuxEntities.language { id = 1, Value = row.Cell(index++).GetValue<string>() });

                    //AA
                    product.meta_description.Add(new Bukimedia.PrestaSharp.Entities.AuxEntities.language { id = 1, Value = row.Cell(index++).GetValue<string>() });

                    //AB
                    index++;

                    //AC
                    index++;

                    //AD
                    product.condition = row.Cell(index++).GetValue<string>();

                    //AE
                    product.show_condition = row.Cell(index++).GetValue<int>();

                    //AF
                    product.ean13 = row.Cell(index++).GetValue<string>();

                    var newProduct = await ProductClient.AddAsync(product);
                    newProduct = await ProductClient.GetAsync(newProduct.id.Value);

                    newProduct.id_category_default = categoryId;
                    if (newProduct.associations.categories.Any())
                    {
                        newProduct.associations.categories.Clear();
                    }
                    newProduct.associations.categories.Add(new Bukimedia.PrestaSharp.Entities.AuxEntities.category(categoryId));
                    newProduct.position_in_category = 0;
                    newProduct.id_shop_default = shopId;

                    await ProductClient.UpdateAsync(newProduct);

                    var productStockAvailable = newProduct.associations.stock_availables.FirstOrDefault();
                    if (productStockAvailable != null)
                    {
                        var stockAvailable = await StockClient.GetAsync(productStockAvailable.id);
                        stockAvailable.id_shop = shopId;
                        stockAvailable.location = product.location;
                        stockAvailable.quantity = quantity;

                        await StockClient.UpdateAsync(stockAvailable);
                    }

                    


                }
            }
        }
    }
}