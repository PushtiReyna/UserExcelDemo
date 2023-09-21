using Microsoft.AspNetCore.Mvc;
using Microsoft.Data.SqlClient;
using OfficeOpenXml;
using UserExcel.Models;
using Microsoft.EntityFrameworkCore;
using Newtonsoft.Json;
using iTextSharp.text.pdf;
using iTextSharp.text;
using NuGet.Protocol;
using Microsoft.AspNetCore.Hosting;
using Microsoft.Win32;
using Microsoft.AspNetCore.Http.HttpResults;
using Microsoft.EntityFrameworkCore.Metadata.Internal;
using System.Numerics;
//using iTextSharp.tool.xml;


namespace UserExcel.Controllers
{
    public class UserController : Controller
    {
        private readonly UserDbContext _db;
        private readonly IConfiguration _configuration;
        private readonly IWebHostEnvironment _webHostEnvironment;
        public UserController(UserDbContext db, IConfiguration configuration, IWebHostEnvironment webHostEnvironment)
        {
            _db = db;
            _configuration = configuration;
            _webHostEnvironment = webHostEnvironment;
        }
        [HttpGet]
        public IActionResult ImportExcelFile()
        {
            return View();
        }

        public IActionResult ImportExcelFile(IFormFile formFile)
        {
            if (formFile != null)
            {
                var usersList = new List<UserMst>();
                using (var stream = new MemoryStream())
                {
                    formFile.CopyTo(stream);

                    ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
                    using (var package = new ExcelPackage(stream))
                    {
                        ExcelWorksheet worksheet = package.Workbook.Worksheets[0];
                        var rowcount = worksheet.Dimension.Rows;
                        for (int rowIterator = 2; rowIterator <= rowcount; rowIterator++)
                        {
                            var user = new UserMst();
                            user.FirstName = worksheet.Cells[rowIterator, 1].Value.ToString();
                            user.MiddleName = worksheet.Cells[rowIterator, 2].Value.ToString();
                            user.LastName = worksheet.Cells[rowIterator, 3].Value.ToString();
                            user.UserName = worksheet.Cells[rowIterator, 4].Value.ToString();
                            user.Password = worksheet.Cells[rowIterator, 5].Value.ToString();
                            user.Address = worksheet.Cells[rowIterator, 6].Value.ToString();
                            user.Pincode = worksheet.Cells[rowIterator, 7].Value.ToString();
                            user.Mobile1 = worksheet.Cells[rowIterator, 8].Value.ToString();
                            user.Mobile2 = worksheet.Cells[rowIterator, 9].Value.ToString();
                            user.Email = worksheet.Cells[rowIterator, 10].Value.ToString();
                            user.CompanyName = worksheet.Cells[rowIterator, 11].Value.ToString();
                            user.CreateDate = DateTime.Now;
                            user.IsActive = true;
                            usersList.Add(user);
                        }
                    }
                }

                foreach (var item in usersList)
                {
                    //_db.Entry(item).State = EntityState.Modified;
                    _db.UserMsts.Add(item);
                }
                _db.SaveChanges();

                ViewBag.message = "file uploaded";
                return View();
            }
            else
            {
                ViewBag.message = "file not uploaded.";
                return View();
            }
        }

        [HttpGet]
        public string GetDataTable()
        {
            var userList = _db.UserMsts.ToList();
            var result = JsonConvert.SerializeObject( new { data = userList });
            return result;
        }

        [HttpGet]
        public IActionResult GetPDF()
        {
            return View();
        }
        //create pdf of the view.
        // [HttpPost]
        //public IActionResult GetPDF(string ExportData)
        //{
        //    using (MemoryStream stream = new System.IO.MemoryStream())
        //    {
        //        StringReader reader = new StringReader(ExportData);
        //        Document PdfFile = new Document(PageSize.A4);
        //        PdfWriter writer = PdfWriter.GetInstance(PdfFile, stream);
        //        PdfFile.Open();
        //        XMLWorkerHelper.GetInstance().ParseXHtml(writer, PdfFile, reader);
        //        PdfFile.Close();
        //        return File(stream.ToArray(), "application/pdf", "ExportData.pdf");
        //    }
        //}

        [HttpPost]
        public ActionResult PDFDownload()
        {
            
            Document pdfDoc = new Document(PageSize.A4, 25, 25, 25, 15);
            
            if (System.IO.File.Exists("D:\\Example.pdf"))
            {
                System.IO.File.Delete("D:\\Example.pdf");
            }
            FileStream FS = new FileStream("D:\\Example.pdf", FileMode.Create);
           
            PdfWriter pdfWriter = PdfWriter.GetInstance(pdfDoc, FS);

            pdfDoc.Open();

            Paragraph header = new Paragraph("Registration Form", new Font(Font.FontFamily.HELVETICA,16,Font.BOLD));
            header.SpacingBefore = 10f;
            header.SpacingAfter = 10f;
            header.IndentationLeft = 200f;
            pdfDoc.Add(header);

            Paragraph para = new Paragraph();
            para.Add("Date: 4th September 2023\r\n");
            para.SpacingAfter = 20f;
            para.IndentationLeft = 350f;
            para.IndentationRight = 20f;
            pdfDoc.Add(para);

            para = new Paragraph();
            para.Add("Create registration form which contain following fields.");
            para.SpacingAfter = 10f;
            para.IndentationLeft = 50f;
            pdfDoc.Add(para);

            List list = new List(List.ORDERED);
           
            list.Add(new ListItem("First Name : Drashti"));
            
            list.Add(new ListItem("Last Name  : Patel"));
            list.Add(new ListItem("DOB          : 01/01/2001"));
            list.Add(new ListItem("Gender       : Female"));
            list.Add(new ListItem("Email        : drashti@gmail.com"));
            list.Add(new ListItem("Phone        : 9898083705"));
            list.Add(new ListItem("Username     : Drashti12"));
            list.Add(new ListItem("Password     : 12@Patel"));
            list.Add(new ListItem("Department   : CSE"));
            list.Add(new ListItem("Is Active    : Active"));
            list.IndentationLeft = 50f;
            list.IndentationRight = 100f;
            pdfDoc.Add(list);          

            para = new Paragraph();
            para.Add("Create database using required field -MS SQL server.");
            para.SpacingBefore = 10f;
            para.SpacingAfter = 10f;
            para.IndentationLeft = 50f;
            pdfDoc.Add(para);

            para = new Paragraph();
            para.Add("Create page using proper validation and design.");
            para.SpacingAfter = 10f;
            para.IndentationLeft = 50f;
            pdfDoc.Add(para);

            para = new Paragraph();
            para.Add("Complete task before EOD");
            para.SpacingAfter = 10f;
            para.IndentationLeft = 50f;
            pdfDoc.Add(para);

            pdfWriter.CloseStream = false;
            pdfDoc.Close();

            FS.Close();
            TempData["message"] = "PDF download in D Drive";
            return RedirectToAction("GetPDF");
        }
    }

}


