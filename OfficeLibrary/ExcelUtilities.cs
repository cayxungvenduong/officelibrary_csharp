using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Reflection;
using OfficeOpenXml;
using System.Drawing;
using System.Runtime.InteropServices;
using OfficeOpenXml.Style;

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


        public void ListToExcel<T>(IEnumerable<T> listExport, bool isShowHeader = true,
            int startRow = 1,
            int startColumn = 1)
        {
            if (startRow < 1 || startColumn < 1)
                throw new ArgumentException("Argument row start or column start is not valid.");
            if (listExport == null) throw new ArgumentException("Argument list is null or not valid.");
            var list = listExport.ToList();

            using (var app = new ExcelPackage(new FileInfo(ExportPath)))
            {
                app.Workbook.Worksheets.Add("Sheet1");
                var sheet1 = app.Workbook.Worksheets[1];

                sheet1.Cells[startRow, startColumn].LoadFromCollection(list, isShowHeader);
                var properties = typeof(T).GetProperties();
                var headers = properties.Select(o => o.Name).ToList();
                if (isShowHeader)
                {
                    
                    var headerRange = sheet1.Cells[startRow, startColumn, startRow, startColumn + headers.Count - 1];

                    headerRange.Style.Font.Bold = true;
                    headerRange.Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                    headerRange.Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                    
                }

                var totalRange = sheet1.Cells[startRow, startColumn, startRow + list.Count,
                    startColumn + headers.Count - 1];
                
                totalRange.Style.Border.Top.Style=ExcelBorderStyle.Thin;
                totalRange.Style.Border.Right.Style=ExcelBorderStyle.Thin;
                totalRange.Style.Border.Bottom.Style=ExcelBorderStyle.Thin;
                totalRange.Style.Border.Left.Style=ExcelBorderStyle.Thin;
                
                totalRange.AutoFitColumns();

                app.Save();
                app.Workbook.Worksheets.Dispose();
                app.Workbook.Dispose();
                app.Dispose();
            }

//
//            var properties = typeof(T).GetProperties();
//            var headers = properties.Select(o => o.Name).ToList();

            //if (sheet1 != null)
            //{
            //    // Header Range
//            var rangeHeader = sheet1.Range[sheet1.Cells[startRow, startColumn],
//                sheet1.Cells[startRow, startColumn + headers.Count - 1]];

            //    // Fill header
            //    var currentRow = 0;
            //    if (isShowHeader)
            //    {
            //        for (var i = 0; i < headers.Count; i++)
            //        {
            //            sheet1.Cells[startRow, startColumn + i] = headers[i];
            //        }

            //        currentRow++;
            //    }

            //    // Data Range            
            //    var rangeData = sheet1.Range[sheet1.Cells[startRow + currentRow, startColumn],
            //        sheet1.Cells[startRow + currentRow + list.Count - 1, startColumn + headers.Count - 1]];

            //    // Fill Data
            //    foreach (var item in list)
            //    {
            //        for (var i = 0; i < headers.Count; i++)
            //        {
            //            if (sheet1.Cells[currentRow + startRow, i + startColumn] is Range cell)
            //                cell.Value2 = GetPropValue(item, headers[i]);
            //        }

            //        currentRow++;
            //    }

            //    Range rangeTotal;
            //    if (isShowHeader && rangeHeader != null && rangeData != null)
            //    {
            //        var freezeRow = (Range)sheet1.Rows[startRow + 1];
            //        freezeRow.Select();

            //        workbook.Application.ActiveWindow.FreezePanes = true;
            //        var styleHeader = workbook.Styles.Add("MyExport Style");
            //        styleHeader.Font.Bold = true;

            //        rangeHeader.Style = styleHeader;
            //        rangeTotal = sheet1.Range[rangeHeader, rangeData];
            //    }
            //    else
            //    {
            //        rangeTotal = rangeData;
            //    }

            //    if (rangeTotal != null)
            //    {

            //        rangeTotal.Columns.AutoFit();
            //        rangeTotal.Borders.LineStyle = XlLineStyle.xlContinuous;
            //    }
            //}

            //workbook.SaveAs(ExportPath, XlFileFormat.xlWorkbookDefault);
            //appExcel.Application.Workbooks.Close();
            //Marshal.FinalReleaseComObject(appExcel.Application.Workbooks);
            //appExcel.Quit();
            //Marshal.FinalReleaseComObject(appExcel);
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