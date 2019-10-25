using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Reflection;
using Microsoft.Office.Interop.Excel;
using System.Drawing;

namespace OfficeLibrary
{
    public class ExcelUtilities
    {
        public string ExportPath { get; }

        public ExcelUtilities(string path = "")
        {
            ExportPath = path == string.Empty
                ? Path.Combine(AppDomain.CurrentDomain.BaseDirectory, $"{DateTime.Now:yyyyMMdd-HHmmss}.xlsx")
                : path;
        }


        public void ListToExcel<T>(List<T> list, bool isShowHeader = true, bool isVisible = false, int startRow = 1,
            int startColumn = 2, string indexName = "No")
        {
            if (startRow < 1 || startColumn < 2)
                throw new ArgumentException("Argument row start or column start is not valid.");
            if (list == null || list.Count == 0) throw new ArgumentException("Argument list is null or not valid.");


            // Init Export Excel app,workbook, worksheet
            var appExcel = new ApplicationClass {Visible = isVisible};
            var workbook = appExcel.Workbooks.Add(Missing.Value);
            var sheet1 = workbook.Sheets[1] as _Worksheet;
            var properties = list.First().GetType().GetProperties();
            var headers = properties.Select(o => o.Name).ToList();
            // Header Range
            var rangeHeader = sheet1?.Range[sheet1.Cells[startRow, startColumn],
                sheet1.Cells[startRow, startColumn + headers.Count - 1]];

            // Fill header
            var currentRow = 0;
            if (isShowHeader)
            {
                for (var i = 0; i < headers.Count; i++)
                {
                    if (sheet1 != null) sheet1.Cells[startRow, startColumn + i] = headers[i];
                }

                currentRow++;
            }

            // Data Range            
            var rangeData = sheet1?.Range[sheet1.Cells[startRow + currentRow, startColumn],
                sheet1.Cells[startRow + currentRow + list.Count - 1, startColumn + headers.Count - 1]];

            // Fill Data
            foreach (var item in list)
            {
                for (var i = 0; i < headers.Count; i++)
                {
                    if (sheet1?.Cells[currentRow + startRow, i + startColumn] is Range cell)
                        cell.Value2 = GetPropValue(item, headers[i]);
                }

                currentRow++;
            }

            Range rangeTotal;
            if (isShowHeader && rangeHeader != null && rangeData != null)
            {
                var styleHeader = workbook.Styles.Add("MyExport Style");
                styleHeader.Font.Bold = true;

                rangeHeader.Style = styleHeader;
                rangeTotal = sheet1.Range[rangeHeader, rangeData];
            }
            else
            {
                rangeTotal = rangeData;
            }

            if (rangeTotal != null) rangeTotal.Borders.LineStyle = XlLineStyle.xlContinuous;
            workbook.SaveAs(ExportPath, XlFileFormat.xlWorkbookDefault);
            appExcel.Quit();
        }

        private static dynamic GetPropValue(object src, string propName)
        {
            try
            {
                return src.GetType().GetProperty(propName)?.GetValue(src, null);
            }
            catch
            {
                return null;
            }
        }
    }
}