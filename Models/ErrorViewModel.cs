using System.Data;

namespace GenerateExcelFromList.NetCore7.Models
{
    public class ErrorViewModel
    {
        public string? RequestId { get; set; }

        public bool ShowRequestId => !string.IsNullOrEmpty(RequestId);
    }

    public class Product
    {
        public int Id { get; set; }
        public string Code { get; set; }
        public string SKu { get; set; }
        public DateTime ProductAddedOn { get; set; }
        public decimal Amount { get; set; }
        public decimal Cost { get; set; }
    }

    public class DataTableHelper
    {
        public DataTable CreateSampleDataTable()
        {
            // Create a DataTable
            DataTable dataTable = new DataTable("Products");

            // Add columns
            dataTable.Columns.Add("Id", typeof(int));
            dataTable.Columns.Add("SKU", typeof(string));
            dataTable.Columns.Add("Amount", typeof(int));
            dataTable.Columns.Add("Code", typeof(string));
            dataTable.Columns.Add("Cost", typeof(double));
            dataTable.Columns.Add("ProductAddedOn", typeof(DateTime));

            // Add rows
            dataTable.Rows.Add(1, "Sku1", 10, "Code1", 11.0, DateTime.Now);
            dataTable.Rows.Add(2, "Sku2", 20, "Code2", 12.0, DateTime.Now);
            dataTable.Rows.Add(3, "Sku3", 30, "Code3", 13.0, DateTime.Now);
            dataTable.Rows.Add(4, "Sku4", 40, "Code4", 14.0, DateTime.Now);
            dataTable.Rows.Add(5, "Sku5", 50, "Code5", 15.0, DateTime.Now);

            return dataTable;
        }
    }

}