using ExcelTest.Models.EF;
using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Web.Mvc;
using System.Web.Services.Description;

namespace ExcelTest.Controllers
{
    public class ImportExcelController : Controller
    {
        // GET: ImportExcel
        public ActionResult Index()
        {
            return View();
        }
        public ActionResult Upload(FormCollection formCollection)
        {
            var productList = new List<Product>();
            if (Request != null)
            {
                HttpPostedFileBase file = Request.Files["UploadedFile"];
                if ((file != null) && (file.ContentLength > 0) && !string.IsNullOrEmpty(file.FileName))
                {
                    string fileName = file.FileName;
                    string fileContentType = file.ContentType;
                    byte[] fileBytes = new byte[file.ContentLength];
                    var data = file.InputStream.Read(fileBytes, 0, Convert.ToInt32(file.ContentLength));
                    using (var package = new ExcelPackage(file.InputStream))
                    {
                        var currentSheet = package.Workbook.Worksheets;
                        var workSheet = currentSheet.First(); 
                        var noOfCol = workSheet.Dimension.End.Column;
                        var noOfRow = workSheet.Dimension.End.Row;
                        for (int rowIterator = 2; rowIterator <= noOfRow; rowIterator++)
                        {
                            var product = new Product();
                            product.ID = Convert.ToInt32(workSheet.Cells[rowIterator, 1].Value);
                            product.Code = workSheet.Cells[rowIterator, 2].Value.ToString();
                            product.MetaTitle = workSheet.Cells[rowIterator, 3].ToString();
                            product.Name = workSheet.Cells[rowIterator, 4].ToString();
                            product.Price = Convert.ToInt32(workSheet.Cells[rowIterator, 8].Value);
                            productList.Add(product);
                        }
                    }
                }
            }
            using (EntityModel excelImportDBEntities = new EntityModel())
            {
                foreach (var item in productList)
                {
                    excelImportDBEntities.Products.Add(item);
                }
                excelImportDBEntities.SaveChanges();
                
            }
            
            return View("Index");
        }
        
    }
}