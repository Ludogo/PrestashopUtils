using System.IO;
using System.Threading.Tasks;

namespace GGIO.PrestashopUtils.Abstraction
{
    public interface IProductImportService
    {
        Task ImportAsync(Stream file, int shopId, int languageId);
    }

}