using System;
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
                    try
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
                        product.description.Add(new Bukimedia.PrestaSharp.Entities.AuxEntities.language { id = 1, Value = row.Cell(index++).GetValue<string>() });
                        //C //D
                        if (TryGetValue<int>(row, ref index, out var featureId))
                        {
                            //get value id 
                            var valuesIds = row.Cell(index++).GetValue<string>().Split(';').Select(x => int.Parse(x));
                            foreach (var valueId in valuesIds)
                            {
                                product.associations.product_features.Add(new Bukimedia.PrestaSharp.Entities.AuxEntities.product_feature { id = featureId, id_feature_value = valueId });
                            }
                        }
                        //E
                        if (TryGetValue<int>(row, ref index, out var manufaturerId))
                        {
                            product.id_manufacturer = manufaturerId;
                        }
                        //F
                        product.reference = row.Cell(index++).GetValue<string>();

                        //G  
                        if (TryGetValue<int>(row, ref index, out var quantity))
                        {
                            product.minimal_quantity = quantity;
                        }

                        //H
                        if (TryGetValue<decimal>(row, ref index, out var price))
                        {
                            product.price = price;
                        }

                        //I
                        index++;
                        //J
                        if (TryGetValue<int>(row, ref index, out var taxRule))
                        {
                            product.id_tax_rules_group = taxRule;
                        }

                        //K
                        int categoryId = row.Cell(index++).GetValue<int>();

                        //L 
                        index ++;
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
                        if (TryGetValue<decimal>(row, ref index, out var width))
                        {
                            product.width = width;
                        }

                        //S
                        if (TryGetValue<decimal>(row, ref index, out var heigth))
                        {
                            product.height = heigth;
                        }

                        //T
                        if (TryGetValue<decimal>(row, ref index, out var depth))
                        {
                            product.depth = depth;
                        }

                        //U
                        if (TryGetValue<decimal>(row, ref index, out var weight))
                        {
                            product.weight = weight;
                        }

                        //V
                        index++;

                        //W
                        index++;

                        //X
                        if (TryGetValue<decimal>(row, ref index, out var wholeSale))
                        {
                            product.wholesale_price = wholeSale;
                        }

                        //Y
                        product.meta_title.Add(new Bukimedia.PrestaSharp.Entities.AuxEntities.language { id = 1, Value = row.Cell(index++).GetValue<string>() });

                        //Z  
                        var metaDescription = row.Cell(index++).GetValue<string>();
                        product.AddLinkRewrite(new Bukimedia.PrestaSharp.Entities.AuxEntities.language { id = 1, Value = metaDescription });
                        product.meta_keywords.Add(new Bukimedia.PrestaSharp.Entities.AuxEntities.language { id = 1, Value = metaDescription });

                        //AA
                        product.meta_description.Add(new Bukimedia.PrestaSharp.Entities.AuxEntities.language { id = 1, Value = row.Cell(index++).GetValue<string>() });

                        //AB
                        index++;

                        //AC
                        index++;

                        //AD
                        product.condition = row.Cell(index++).GetValue<string>();

                        //AE
                        if (TryGetValue<int>(row, ref index, out var showCondition))
                        {
                            product.show_condition = showCondition;
                        }

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
                    catch (Exception )
                    {
                        
                    }
                }
            }
        }

        private bool TryGetValue<T>(IXLRow row, ref int index, out T value)
        {
            bool ret = false;
            value = default(T);
            if (!string.IsNullOrEmpty(row.Cell(index).GetString()))
            {
                ret = row.Cell(index++).TryGetValue<T>(out value);
            }

            return ret;
        }
    }
}