using DemoExcel.Models;
using Microsoft.AspNetCore.Mvc;
using OfficeOpenXml;
using OfficeOpenXml.Style;
using System.Diagnostics;
using System.Drawing;

namespace DemoExcel.Controllers
{
    public class HomeController : Controller
    {
        private readonly ILogger<HomeController> _logger;
        private readonly IWebHostEnvironment _hostingEnvironment;
        public HomeController(ILogger<HomeController> logger, IWebHostEnvironment hostingEnvironment)
        {
            _logger = logger;
            _hostingEnvironment = hostingEnvironment;
        }

        public IActionResult Index()
        {
            var lst = new List<Student>()
            {
                new Student()
                {
                    Code = "1141060075",
                    Name = "Đỗ Như Nam"
                },
                new Student()
                {
                    Code = "3213141234",
                    Name = "ABC"
                }
            };
            // If you are a commercial business and have
            // purchased commercial licenses use the static property
            // LicenseContext of the ExcelPackage class:
            ExcelPackage.LicenseContext = LicenseContext.Commercial;

            // If you use EPPlus in a noncommercial context
            // according to the Polyform Noncommercial license:
            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
            using (ExcelPackage p = new ExcelPackage())
            {
                p.Workbook.Properties.Author = "1";
                p.Workbook.Properties.Title = "2";
                p.Workbook.Worksheets.Add("Sheet 1");

                ExcelWorksheet ws = p.Workbook.Worksheets[0];
                ws.Cells.AutoFitColumns();
                ws.Name = "T1";
                ws.Cells.Style.Font.Size = 11;
                ws.Cells.Style.Font.Name = "Calibri";

                string[] arrCol = new string[] { "Mã sinh viên", "Họ tên" };
                var countColHeader = arrCol.Count();

                ws.Cells[1, 1].Value = "Thống kê thông tin sinh viên";
                ws.Cells[1, 1, 1, countColHeader].Merge = true;
                ws.Cells[1, 1, 1, countColHeader].Style.Font.Bold = true;
                ws.Cells[1, 1, 1, countColHeader].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                ws.Cells[1, 1, 1, countColHeader].AutoFitColumns();
                int colIndex = 1;
                int rowIndex = 2;
                //tạo các header từ column header đã tạo từ bên trên
                foreach (var item in arrCol)
                {
                    var cell = ws.Cells[rowIndex, colIndex];
                    cell.AutoFitColumns();
                    //set màu thành gray
                    var fill = cell.Style.Fill;
                    fill.PatternType = ExcelFillStyle.Solid;
                    fill.BackgroundColor.SetColor(System.Drawing.Color.Yellow);
                    
                    var border = cell.Style.Border;
                    border.Bottom.Style =
                    border.Top.Style =
                    border.Left.Style =
                    border.Right.Style = ExcelBorderStyle.Thin;
                    cell.Value = item;
                    colIndex++;
                }

                foreach (var item in lst)
                {
                    colIndex = 1;
                    rowIndex++;
                    ws.Cells[rowIndex, colIndex++].Value = item.Code;
                    ws.Cells[rowIndex, colIndex++].Value = item.Name;
                }

                Byte[] bin = p.GetAsByteArray();
                string path = Path.Combine(_hostingEnvironment.WebRootPath, "test2.xlsx");

                System.IO.File.WriteAllBytes(path, bin);

            }
            return View();
        }

        public IActionResult Privacy()
        {
            return View();
        }

        [ResponseCache(Duration = 0, Location = ResponseCacheLocation.None, NoStore = true)]
        public IActionResult Error()
        {
            return View(new ErrorViewModel { RequestId = Activity.Current?.Id ?? HttpContext.TraceIdentifier });
        }
    }

    public class Student
    {
        public string Code { get; set;}
        public string Name { get; set; }
    }
}