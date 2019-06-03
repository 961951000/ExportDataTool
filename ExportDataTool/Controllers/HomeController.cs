using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.Threading.Tasks;
using Microsoft.AspNetCore.Mvc;
using ExportDataTool.Models;
using System.Data;
using System.Data.SqlClient;
using Dapper;
using System.IO;
using OfficeOpenXml;
using System.Net.Mime;

namespace ExportDataTool.Controllers
{
    public class HomeController : Controller
    {
        public IActionResult Index()
        {
            return View();
        }

        public IActionResult Privacy()
        {
            return View();
        }

        public IActionResult ExportData(ExportModel model)
        {
            if (!ModelState.IsValid)
            {
                return View("Index");
            }

            var table = new DataTable();
            using (var connection = new SqlConnection(model.ConnectionString.Trim()))
            {
                var reader = connection.ExecuteReader(model.SQL.Trim());
                table.Load(reader);

                var fileStream = TableToStream(table);

                return this.File(fileStream, MediaTypeNames.Application.Octet, $"{Guid.NewGuid().ToString()}.xlsx");
            }
        }

        [ResponseCache(Duration = 0, Location = ResponseCacheLocation.None, NoStore = true)]
        public IActionResult Error()
        {
            return View(new ErrorViewModel { RequestId = Activity.Current?.Id ?? HttpContext.TraceIdentifier });
        }

        #region Private methods

        public static MemoryStream TableToStream(DataTable data, string worksheetName = "Sheet1")
        {
            using (var package = new ExcelPackage())
            {
                ExcelWorksheet worksheet = package.Workbook.Worksheets.Add(worksheetName);
                worksheet.Cells["A1"].LoadFromDataTable(data, true);
                var stream = new MemoryStream(package.GetAsByteArray());

                return stream;
            }
        }

        #endregion
    }
}
