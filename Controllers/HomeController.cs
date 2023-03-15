using ClosedXML.Excel;
using Microsoft.AspNetCore.Mvc;
using BlackCoderExportToExcelWithCompanyLogo.Models;
using System.Diagnostics;

namespace BlackCoderExportToExcelWithCompanyLogo.Controllers
{
        public class HomeController : Controller
        {
            public IActionResult Index()
            {
                var data = new List<Person>
            {
                new Person { CopyRights = "BlackCoder", Bid = 30 },
                new Person { CopyRights = "BlackCoder", Bid = 25 },
                new Person { CopyRights = "BlackCoder", Bid = 40 }
            };
                return View(data);
            }

            public IActionResult ExportToExcel()
            {
                var data = new List<Person>
            {
                new Person { CopyRights = "BlackCoder", Bid = 30 },
                new Person { CopyRights = "BlackCoder", Bid = 25 },
                new Person { CopyRights = "BlackCoder", Bid = 40 }
            };

                var wb = new XLWorkbook();
                var ws = wb.Worksheets.Add("Sheet1");

                var picture = ws.AddPicture("wwwroot/images/TrickSologic.png")
                    .MoveTo(ws.Cell("B2"))
                    .Scale(0.5);

                for (int i = 0; i < data.Count; i++)
                {
                    ws.Cell(i + 2, 1).Value = data[i].CopyRights;
                    ws.Cell(i + 2, 2).Value = data[i].Bid;
                }

                var stream = new MemoryStream();
                wb.SaveAs(stream);
                var content = stream.ToArray();

                return File(content, "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", "data.xlsx");
            }
        }

        public class Person
        {
            public string? CopyRights { get; set; }
            public int Bid { get; set; }
        }
    }
