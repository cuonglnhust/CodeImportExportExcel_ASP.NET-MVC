using ExcelTest.Models.Dao;
using ExcelTest.Models.EF;
using OfficeOpenXml;
using OfficeOpenXml.Table;
using System;
using System.Collections.Generic;
using System.Configuration;
using System.Data;
using System.Data.SqlClient;
using System.IO;
using System.Linq;
using System.Web;
using System.Web.Mvc;

namespace ExcelTest.Controllers
{
    public class HomeController : Controller
    {

        public ActionResult Index()
        {
            return View();
        }

        public List<Product> CreateItem()
        {
            var NewListItem = new List<Product>();
            var model = new ProductDao().ListALL();
            foreach (var item in model)
            {
                var a = new Product()
                {
                    ID = item.ID,
                    Code = item.Code,
                    CreateDate = item.CreateDate,
                    Status = item.Status,
                    Detail = item.Detail,
                    Price = item.Price,
                };
                NewListItem.Add(a);
            }
            return NewListItem;
        }


        private Stream CreateExcelFile(Stream stream = null)
        {

            var list = CreateItem();
            using (var excelPackge = new ExcelPackage(stream ?? new MemoryStream()))
            {

                //Tao thong tin cho file excel
                excelPackge.Workbook.Properties.Author = "Cuongln";
                excelPackge.Workbook.Properties.Title = "test EPPlus";
                excelPackge.Workbook.Properties.Comments = "this is my excel";
                // Add sheet vua tao vao file excel
                excelPackge.Workbook.Worksheets.Add("Luu Nhan Cuong");
                // lay sheet vua tao ra de thao tac
                var workSheet = excelPackge.Workbook.Worksheets[1];
                //Do data vao Excel File
                workSheet.Cells[1, 1].LoadFromCollection(list, true, TableStyles.Dark9);
                excelPackge.Save();
                return excelPackge.Stream;
            }
        }

        [HttpGet]
        public ActionResult Export()
        {
            // Goi lai ham de tao excel file
            var stream = CreateExcelFile();
            // Tao buffer memory stream de hung file excel
            var buffer = stream as MemoryStream;
            //Tao Content type 
            Response.ContentType = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet";
            //// Dòng này rất quan trọng, vì chạy trên firefox hay IE thì dòng này sẽ hiện Save As dialog cho người dùng chọn thư mục để lưu
            // File name của Excel này là ExcelDemoCuonln
            Response.AddHeader("Content-Disposition", "attachment;filename = ExcelDemoCuonglnHust.xlsx");
            //Luu file excel cua chung ta nhuw 1 mang byte de tra ve response
            Response.BinaryWrite(buffer.ToArray());
            //Gui tat ca output bytes ve phia client
            Response.Flush();
            Response.End();
            // Redirect ve trang inde
            return RedirectToAction("Index");
        }


        private DataTable ReadFileExcel(string path, string sheetName)
        {
            //Khoiwr taoj Data table
            DataTable dt = new DataTable();
            //Load file excel vaf cacs setting ban dau
            using (ExcelPackage package = new ExcelPackage(new FileInfo(path)))
            {
                if (package.Workbook.Worksheets.Count < 1)
                {
                    //log ko sheet nao ton tai trong file excel
                    return null;
                }
                // Khởi Lấy Sheet đầu tiện trong file Excel để truy vấn, truyền vào name của Sheet để lấy ra sheet cần, nếu name = null thì lấy sheet đầu tiên
                ExcelWorksheet worksheet = package.Workbook.Worksheets.FirstOrDefault(x => x.Name == sheetName) ?? package.Workbook.Worksheets.FirstOrDefault();
                //Doc tat ca header
                foreach (var firstRowCell in worksheet.Cells[1, 1, 1, worksheet.Dimension.End.Column])
                {
                    dt.Columns.Add(firstRowCell.Text);

                }
                //Docj tat ca data tu row 2
                for (var rowNumber = 2; rowNumber <= worksheet.Dimension.End.Row; rowNumber++)
                {
                    //lay 1 row ra de truy van 
                    var row = worksheet.Cells[rowNumber, 1, rowNumber, worksheet.Dimension.End.Column];
                    //tao 1 row trong data table 
                    var newRow = dt.NewRow();
                    foreach (var cell in row)
                    {
                        newRow[cell.Start.Column - 1] = cell.Text;
                    }
                    dt.Rows.Add(newRow);
                }
                return dt;
            }

           
        }
        
       
    
        [HttpGet]
        public ActionResult ReadFileExcel()
        {
            var data = ReadFileExcel(@"C:\Users\Luu Nhan Cuong\Downloads\ExcelDemoCuonglnHust.xlsx", "Luu Nhan Cuong");
            return View(data);
        }
    }
}