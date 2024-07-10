using ClosedXML.Excel;
using System.Data;

namespace GenerateExcelFromList.NetCore7.Common
{
    public class ExcelGenerator
    {
        public byte[] GenerateExcel<T>(IEnumerable<T> items)
        {
            using (var workbook = new XLWorkbook())
            {
                var worksheet = workbook.Worksheets.Add("Products");

                // Add headers based on Product properties
                var properties = typeof(T).GetProperties();
                for (int i = 0; i < properties.Length; i++)
                {
                    var cell = worksheet.Cell(1, i + 1);
                    cell.Value = properties[i].Name;

                    // Apply formatting to header cell
                    cell.Style.Font.Bold = true;
                    cell.Style.Fill.BackgroundColor = XLColor.LightGray;
                    cell.Style.Border.BottomBorder = XLBorderStyleValues.Thin;
                    cell.Style.Border.BottomBorderColor = XLColor.Black;
                    cell.Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
                }

                // Add data rows
                int row = 2; // Start from row 2 after headers
                foreach (var item in items)
                {
                    int col = 1;
                    foreach (var prop in properties)
                    {
                        var value = prop.GetValue(item);
                        if (value != null)
                        {
                            if (prop.PropertyType == typeof(DateTime))
                            {
                                worksheet.Cell(row, col).Value = (DateTime)value;
                            }
                            else if (prop.PropertyType == typeof(int))
                            {
                                worksheet.Cell(row, col).Value = (int)value;
                            }
                            else if (prop.PropertyType == typeof(double))
                            {
                                worksheet.Cell(row, col).Value = (double)value;
                            }
                            else if (prop.PropertyType == typeof(decimal))
                            {
                                worksheet.Cell(row, col).Value = (decimal)value;
                            }
                            else
                            {
                                worksheet.Cell(row, col).Value = value.ToString();
                            }
                        }
                        col++;
                    }
                    row++;
                }

                // Auto-fit all columns except the last one (if there are any columns)
                if (properties.Length > 0)
                {
                    worksheet.Columns(1, properties.Length - 1).AdjustToContents();
                }

                // Auto-fit the last column explicitly
                if (properties.Length > 0)
                {
                    worksheet.Column(properties.Length).AdjustToContents();
                }

                using (var stream = new MemoryStream())
                {
                    workbook.SaveAs(stream);
                    return stream.ToArray();
                }
            }
        }
    }
    public class MemoryStreamHelper
    {
        public byte[] ConvertDataTableToMemoryStream(DataTable dataTable)
        {
            using (var stream = new MemoryStream())
            {
                using (var writer = new StreamWriter(stream))
                {
                    WriteDataTable(dataTable, writer);
                }
                return stream.ToArray();
            }
        }

        private void WriteDataTable(DataTable table, TextWriter writer)
        {
            // Write columns header
            foreach (DataColumn column in table.Columns)
            {
                writer.Write(column.ColumnName);
                if (column.Ordinal < table.Columns.Count - 1)
                    writer.Write(",");
            }
            writer.WriteLine();

            // Write row values
            foreach (DataRow row in table.Rows)
            {
                for (int i = 0; i < table.Columns.Count; i++)
                {
                    if (!Convert.IsDBNull(row[i]))
                    {
                        string value = row[i].ToString();
                        if (value.Contains(","))
                            value = String.Format("\"{0}\"", value);

                        writer.Write(value);
                    }
                    if (i < table.Columns.Count - 1)
                        writer.Write(",");
                }
                writer.WriteLine();
            }
            writer.Flush();
        }
    }

}
