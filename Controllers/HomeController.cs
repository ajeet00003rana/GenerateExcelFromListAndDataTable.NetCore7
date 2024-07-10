using GenerateExcelFromList.NetCore7.Common;
using GenerateExcelFromList.NetCore7.Models;
using Microsoft.AspNetCore.Mvc;
using System.Data;
using System.Diagnostics;

namespace GenerateExcelFromList.NetCore7.Controllers
{
    public class HomeController : Controller
    {
        private readonly ILogger<HomeController> _logger;
        private readonly ExcelGenerator _excelGenerator;

        public HomeController(ILogger<HomeController> logger)
        {
            _logger = logger;
            _excelGenerator = new ExcelGenerator();
        }

        public IActionResult Index()
        {
            return View();
        }

        public IActionResult Privacy()
        {
            return View();
        }

        [ResponseCache(Duration = 0, Location = ResponseCacheLocation.None, NoStore = true)]
        public IActionResult Error()
        {
            return View(new ErrorViewModel { RequestId = Activity.Current?.Id ?? HttpContext.TraceIdentifier });
        }

        public IActionResult DownloadExcel()
        {
            List<Product> items = new List<Product> {
                new Product { Id = 1, SKu = "Sku1", Amount = 10.05M, Code = "Code1", Cost = 11.00M, ProductAddedOn = DateTime.Now},
                new Product { Id = 2, SKu = "Sku2", Amount = 20.05M, Code = "Code2", Cost = 12.00M, ProductAddedOn = DateTime.Now},
                new Product { Id = 3, SKu = "Sku3", Amount = 30.05M, Code = "Code3", Cost = 13.00M, ProductAddedOn = DateTime.Now},
                new Product { Id = 4, SKu = "Sku4", Amount = 40.05M, Code = "Code4", Cost = 14.05M, ProductAddedOn = DateTime.Now},
                new Product { Id = 5, SKu = "Sku5", Amount = 50.05M, Code = "Code5", Cost = 15.101M, ProductAddedOn = DateTime.Now}
            };
            byte[] excelBytes = _excelGenerator.GenerateExcel(items);

            string fileName = $"Products_{DateTime.Now.ToString("yyyyMMddHHmmss")}.xlsx";
            string contentType = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet";

            return File(excelBytes, contentType, fileName);
        }

        public IActionResult DownloadDTExcel()
        {
            DataTableHelper dataTableHelper = new DataTableHelper();
            MemoryStreamHelper memoryStreamHelper = new MemoryStreamHelper();

            // Create sample DataTable
            DataTable dataTable = dataTableHelper.CreateSampleDataTable();

            var excelBytes = memoryStreamHelper.ConvertDataTableToMemoryStream(dataTable);

            string fileName = $"Products_{DateTime.Now.ToString("yyyyMMddHHmmss")}.csv";
            string contentType = "text/csv"; 

            return File(excelBytes, contentType, fileName);
        }
    }
}