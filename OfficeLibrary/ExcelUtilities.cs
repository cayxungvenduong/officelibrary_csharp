using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Reflection;
using Microsoft.Office.Interop.Excel;

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

     
        public void ListToExcel<T>(List<T> list,bool isShowHeader=true, bool isVisible = false, int startRow = 1, int startColumn = 1)
        {
            // Init Export Excel app,workbook, worksheet
            var appExcel = new ApplicationClass {Visible = isVisible};
            var workbook = appExcel.Workbooks.Add(Missing.Value);
            var sheet1 = workbook.Sheets[1] as _Worksheet;
            var properties = list.First().GetType().GetProperties();
            var headers = properties.Select(o => o.Name).ToList();
            
            // Fill header
            if (isShowHeader)
            {
                for (var i = 0; i < headers.Count; i++)
                {
                    if (sheet1 != null) sheet1.Cells[startRow, startColumn + i] = headers[i];
                }
            }
        }
    }
}