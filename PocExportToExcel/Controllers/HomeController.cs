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
            var data = new List<ContactData>();

            for (int i = 0; i < 10000; i++)
            {
                data.Add(new ContactData("Ram", "ram@techbrij.com", "111-222-3333"));
                data.Add(new ContactData("Shyam", "shyam@techbrij.com", "159-222-1596"));
                data.Add(new ContactData("Mohan", "mohan@techbrij.com", "456-222-4569"));
                data.Add(new ContactData("Sohan", "sohan@techbrij.com", "789-456-3333"));
                data.Add(new ContactData("Karan", "karan@techbrij.com", "111-222-1234"));
                data.Add(new ContactData("Brij", "brij@techbrij.com", "111-222-3333"));
            }

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

    public class ContactData
    {
        public string Name { get; set; }
        public string Email { get; set; }
        public string Phone { get; set; }

        public ContactData() { }
        public ContactData(string name, string email, string phone)
        {
            Name = name;
            Email = email;
            Phone = phone;
        }
    }
}