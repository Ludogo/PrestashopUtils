using System;
using System.Threading.Tasks;
using McMaster.Extensions.CommandLineUtils;
using Microsoft.Extensions.DependencyInjection;
using Microsoft.Extensions.Logging;
using Microsoft.Extensions.Logging.Console;
using GGIO.PrestashopUtils.Abstraction;
using GGIO.PrestashopUtils.Logic;
using System.IO;

namespace GG_IO.PrestashopUtils
{
    class Program
    {
        static int Main(string[] args)
        {
            var app = new CommandLineApplication
            {
                Name = "PrestashopUtils",
                Description = "Offre un panel d'outils pour intégrer des données à Prestashop via API"
            };

            app.HelpOption(inherited: true);
            app.Command("product", configCmd =>
            {
                configCmd.OnExecute(() =>
                {
                    Console.WriteLine("Specify a subcommand");
                    configCmd.ShowHelp();
                    return 1;
                });

                configCmd.Command("import", importProduct =>
                {
                    importProduct.Description = "Importe des produits à partir d'Excel (xslx)";
                    var urlOpt = importProduct.Option("-u | --url", "Prestashop API's url", CommandOptionType.SingleValue);
                    var keyOpt = importProduct.Option("-k |--key","Prestashop API's KEY",CommandOptionType.SingleValue);
                    var languageOpt = importProduct.Option("-l | --languageId",
                                                            "Identifiant de la langue à utiliser pour les insertions",
                                                            CommandOptionType.SingleValue).IsRequired();
                    var ShopIdOpt = importProduct.Option("-s | --shopId",
                                                         "Identifiant de la boutique dans laquelle les produits doivent être importés",
                                                         CommandOptionType.SingleValue).IsRequired();
                    var file = importProduct.Argument("file", "Chemin du fichier Excel source").IsRequired().Accepts(v => v.ExistingFile("Le fichier spécifié n'existe pas"));

                    importProduct.OnExecuteAsync(async (token) =>
                    {
                        Console.WriteLine($"Importation des produits du fichier {file.Value} sur {urlOpt.Value()} avec la clé {keyOpt.Value()} avec le language {languageOpt.Value()} dans la boutique {ShopIdOpt.Value()}");

                        int shopId = int.Parse(ShopIdOpt.Value());
                        var services = BuildServices(urlOpt.Value(),keyOpt.Value(), int.Parse(ShopIdOpt.Value()));

                        var importService = services.GetRequiredService<IProductImportService>();
                        await importService.ImportAsync(File.Open("/home/ludo/Dev/PrestashopUtils/Test Import Presta M20S.xlsx", FileMode.Open),shopId, int.Parse(languageOpt.Value()));
                        
                    });
                });


            });

            app.OnExecute(() =>
            {
                Console.WriteLine("Specify a subcommand");
                app.ShowHelp();
                return 1;
            });


            return app.Execute(args);
        }

        public static ServiceProvider BuildServices(string prestaUrl , string prestaKey, int shopId)
        {
            var serviceColl = new ServiceCollection();

            serviceColl.AddLogging( conf => 
            {
                conf.AddConsole();
            });

            serviceColl.AddScoped<IProductImportService, ProductImportService>();
            serviceColl.AddScoped<Bukimedia.PrestaSharp.Factories.ProductFactory>((service ) => new Bukimedia.PrestaSharp.Factories.ProductFactory(prestaUrl, prestaKey, string.Empty,shopId) );
            serviceColl.AddScoped<Bukimedia.PrestaSharp.Factories.StockAvailableFactory>((service) => new Bukimedia.PrestaSharp.Factories.StockAvailableFactory(prestaUrl, prestaKey, string.Empty,shopId));
                      


            return serviceColl.BuildServiceProvider();
        }
    }
}
