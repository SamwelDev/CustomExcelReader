using ClosedXML.Excel;
using DocumentFormat.OpenXml.Spreadsheet;
using System;
using System.Collections.Generic;
using System.Reflection;
using System.Text;

namespace CustomExcelReader.Resources.Resources.Helpers
{
    public static class ReaderHelper
    {
        public static byte[] ExportToExcel<T>(List<T> data, string sheetName = "Sheet01")
        {
            using var workbook = new XLWorkbook();
            var _ws = workbook.Worksheets.Add(sheetName);
            var props = typeof(T).GetProperties(BindingFlags.Public | BindingFlags.Instance);
            for (int i = 0; i < props.Length; i++)
            {
                _ws.Cell(1, i + 1).Value = props[i].Name;
                _ws.Cell(1, i + 1).Style.Font.Bold = true;
            }
            for (int r = 0; r < data.Count; r++)
            {
                for (int c = 0; c < props.Length; c++)
                {
                    var value = props[c].GetValue(data[r]);
                    _ws.Cell(r + 2, c + 1).Value = value?.ToString() ?? "";
                }
            }
            _ws.Columns().AdjustToContents();
            using var stream = new MemoryStream();
            workbook.SaveAs(stream);
            return stream.ToArray();
        }

        public static List<T> ImportFromExcel<T>(Stream excelStream) where T : new()
        {
            var result = new List<T>();
            var props = typeof(T).GetProperties();
            using var workbook = new XLWorkbook(excelStream);
            var sheet = workbook.Worksheet(1);
            var rows = sheet.RowsUsed().Skip(1); // skip header
            foreach (var row in rows)
            {
                var instance = new T();

                for (int c = 0; c < props.Length; c++)
                {
                    var cell = row.Cell(c + 1);
                    var value = cell.Value;

                    var prop = props[c];
                    if (!prop.CanWrite) continue;

                    try
                    {
                        object? converted = ConvertValue(cell, prop.PropertyType);
                        prop.SetValue(instance, converted);
                    }
                    catch
                    {
                       
                    }
                }

                result.Add(instance);
            }

            return result;
        }
        private static object? ConvertValue(IXLCell cell, Type type)
        {
            if (cell.IsEmpty())
                return null;

            if (type == typeof(string))
                return cell.GetString();

            if (type == typeof(bool) || type == typeof(bool?))
                return cell.GetBoolean();

            if (type == typeof(int) || type == typeof(int?))
                return (int)cell.GetDouble();

            if (type == typeof(long) || type == typeof(long?))
                return (long)cell.GetDouble();

            if (type == typeof(double) || type == typeof(double?))
                return cell.GetDouble();

            if (type == typeof(DateTime) || type == typeof(DateTime?))
                return cell.GetDateTime();

            return Convert.ChangeType(cell.GetString(), type);
        }
    }
}
