using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Web;
using System.Web.Mvc;

namespace PocExportToExcel.Controllers
{
    public class HomeController : Controller
    {
        public ActionResult Index()
        {
            return View();
        }

        public ActionResult ExportToExcel()
        {
            var data = new[]{
                new{ Name="Ram", Email="ram@techbrij.com", Phone="111-222-3333" },
                new{ Name="Shyam", Email="shyam@techbrij.com", Phone="159-222-1596" },
                new{ Name="Mohan", Email="mohan@techbrij.com", Phone="456-222-4569" },
                new{ Name="Sohan", Email="sohan@techbrij.com", Phone="789-456-3333" },
                new{ Name="Karan", Email="karan@techbrij.com", Phone="111-222-1234" },
                new{ Name="Brij", Email="brij@techbrij.com", Phone="111-222-3333" }
            };

            using (var excelPackage = new ExcelPackage())
            {
                var workSheet = excelPackage.Workbook.Worksheets.Add("Sheet1");
                workSheet.Cells[1, 1].LoadFromCollection(data, true);
                var memoryStream = new MemoryStream();
                {
                    excelPackage.SaveAs(memoryStream);
                    memoryStream.Position = 0;

                    return File(memoryStream, "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", "contacts.xlsx");
                }
            }
        }
    }
}