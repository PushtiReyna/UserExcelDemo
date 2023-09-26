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
using static System.Runtime.InteropServices.JavaScript.JSType;
using System.Data;
using OfficeOpenXml.Style;
using Microsoft.EntityFrameworkCore.Metadata;
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

        public IActionResult Report()
        {
            return View();
        }

        public IActionResult ReportDownload()
        {
            Document pdfDoc = new Document(PageSize.A3, 25, 25, 25, 15);

            //if (System.IO.File.Exists("D:\\CHW.pdf"))
            //{
            //    System.IO.File.Delete("D:\\CHW.pdf");
            //}
            FileStream FS = new FileStream("D:\\Summary Report.pdf", FileMode.Create);

            PdfWriter pdfWriter = PdfWriter.GetInstance(pdfDoc, FS);

            pdfDoc.Open();

            #region Title

            //Paragraph header = new Paragraph("Click Heal Weal", new Font(Font.FontFamily.HELVETICA, 19, Font.NORMAL));
            //header.SpacingBefore = 10f;
            //header.IndentationLeft = 50f;
            //pdfDoc.Add(header);

            #endregion

            #region 1st table

            // CRAETE 1ST TABLE
            PdfPTable table = new PdfPTable(3);
            table.WidthPercentage = 100;
            table.HorizontalAlignment = 20;
            table.SpacingBefore = 20f;
            table.SpacingAfter = 30f;
             
            PdfPCell Title = new PdfPCell(new Phrase("Summary Report", new Font(Font.FontFamily.HELVETICA, 22, Font.BOLD)));
            Title.HorizontalAlignment = Element.ALIGN_CENTER;
            Title.VerticalAlignment = Element.ALIGN_MIDDLE;
            Title.Colspan = 3;
            Title.FixedHeight = 40f;
            Title.BorderColor = BaseColor.BLACK;
            table.AddCell(Title);
            pdfDoc.Add(table);

            #endregion

            #region 2nd table

            //CREATE 2ND TABLE
            
            //  Table - Patient Information
            PdfPTable SecondTable = new PdfPTable(3);
            SecondTable.TotalWidth = 260f;
            SecondTable.LockedWidth = true;
            

            PdfPCell cell = new PdfPCell();
            
            Paragraph Header = new Paragraph("Patient Information", new Font(Font.FontFamily.HELVETICA, 12, Font.BOLD));
            cell.AddElement(Header);
            cell.Colspan = 3;
            cell.HorizontalAlignment = 1;
            cell.BorderColor = BaseColor.BLACK;
            cell.FixedHeight = 30f;
            cell.PaddingLeft = 6f;
            //Border
            cell.Border = 0;
            cell.BorderWidthLeft = 1f;
            cell.BorderWidthTop = 1f;
            cell.BorderWidthRight = 1f;
            SecondTable.AddCell(cell);

            //create sub header
            cell = new PdfPCell();
            Paragraph subHeader = new Paragraph("First Name", new Font(Font.FontFamily.HELVETICA, 11, Font.BOLD));
            cell.FixedHeight = 30f;
            SecondTable.DefaultCell.Border = 0;
            cell.AddElement(subHeader);
            cell.PaddingLeft = 6f;
            cell.Border = 0;
            cell.BorderWidthLeft = 1f;
            SecondTable.AddCell(cell);

            cell = new PdfPCell();
            subHeader = new Paragraph("Last Name", new Font(Font.FontFamily.HELVETICA, 11, Font.BOLD));
            cell.FixedHeight = 30f;
            cell.AddElement(subHeader);
            cell.PaddingLeft = 6f;
            //Border
            cell.Border = 0;
            SecondTable.AddCell(cell);

            cell = new PdfPCell();
            subHeader = new Paragraph("Dignosis", new Font(Font.FontFamily.HELVETICA, 12, Font.BOLD));
            cell.FixedHeight = 30f;
            cell.PaddingLeft = 10f;
            cell.Border = 0;
            cell.BorderWidthRight = 1f;
            cell.AddElement(subHeader);
            SecondTable.AddCell(cell);

            cell = new PdfPCell();
            Paragraph Data = new Paragraph("jhon");
            cell.FixedHeight = 30f;
            cell.Border = 0;
            cell.BorderWidthLeft = 1f;
            cell.BorderWidthBottom = 1f;
            cell.PaddingLeft = 16f;
            cell.AddElement(Data);
            SecondTable.AddCell(cell);

            cell = new PdfPCell();
            Data = new Paragraph("deo");
            cell.FixedHeight = 30f;
            cell.Border = 0;
            cell.BorderWidthBottom = 1f;
            cell.PaddingLeft = 16f;
            cell.AddElement(Data);
            SecondTable.AddCell(cell);

            cell = new PdfPCell();
            Data = new Paragraph("hypertension");
            cell.FixedHeight = 30f;
            cell.Border = 0;
            cell.BorderWidthRight = 1f;
            cell.BorderWidthBottom = 1f;
            cell.PaddingLeft = 10f;
            cell.AddElement(Data);
            SecondTable.AddCell(cell);

            SecondTable.WriteSelectedRows(0, -1, pdfDoc.Left, pdfDoc.Top - 100, pdfWriter.DirectContent);


            //  Table - Physician Information
            SecondTable = new PdfPTable(1);
            SecondTable.TotalWidth = 260f;
            SecondTable.LockedWidth = true;

            cell = new PdfPCell();
            Header = new Paragraph("Physician Information", new Font(Font.FontFamily.HELVETICA, 12, Font.BOLD));
            cell.AddElement(Header);
            cell.BorderColor = BaseColor.BLACK;
            cell.HorizontalAlignment = 0;
            cell.FixedHeight = 30f;
            //Border
            cell.Border = 0;
            cell.BorderWidthLeft = 1f;
            cell.BorderWidthTop = 1f;
            cell.BorderWidthRight = 1f;
            cell.PaddingLeft = 6f;
            SecondTable.AddCell(cell);

            cell = new PdfPCell();
            subHeader = new Paragraph("Name", new Font(Font.FontFamily.HELVETICA, 12, Font.BOLD));
            cell.AddElement(subHeader);
            cell.FixedHeight = 30f;
            //Border
            cell.Border = 0;
            cell.BorderWidthLeft = 1f;
            cell.BorderWidthRight = 1f;
            cell.PaddingLeft = 6f;
            SecondTable.AddCell(cell);

            cell = new PdfPCell();
            Data = new Paragraph("Jhon");
            cell.FixedHeight = 30f;
            cell.Border = 0;
            cell.BorderWidthLeft = 1f;
            cell.BorderWidthRight = 1f;
            cell.BorderWidthBottom = 1f;
            cell.PaddingLeft = 10f;
            cell.AddElement(Data);
            SecondTable.AddCell(cell);

            //SecondTable.AddCell("Jhon");
            SecondTable.WriteSelectedRows(0, -1, pdfDoc.Left + 266, pdfDoc.Top - 100, pdfWriter.DirectContent);


            //  Table - Appointment Information
            SecondTable = new PdfPTable(3);
            SecondTable.TotalWidth = 260f;
            SecondTable.LockedWidth = true;

            cell = new PdfPCell();
            Header = new Paragraph("Appointment Information", new Font(Font.FontFamily.HELVETICA, 12, Font.BOLD));
            cell.AddElement(Header);
            cell.Colspan = 3;
            cell.BorderColor = BaseColor.BLACK;
            cell.HorizontalAlignment = 2;
            cell.PaddingLeft = 6f;
            cell.FixedHeight = 30f;
            //Border
            cell.Border = 0;
            cell.BorderWidthLeft = 1f;
            cell.BorderWidthTop = 1f;
            cell.BorderWidthRight = 1f;
            SecondTable.AddCell(cell);

            cell = new PdfPCell();
            subHeader = new Paragraph("Date", new Font(Font.FontFamily.HELVETICA, 12, Font.BOLD));
            cell.AddElement(subHeader);
            cell.FixedHeight = 30f;
            cell.Border = 0;
            cell.BorderWidthLeft = 1f;
            cell.PaddingLeft = 16f;
            SecondTable.AddCell(cell);

            cell = new PdfPCell();
            subHeader = new Paragraph("Time", new Font(Font.FontFamily.HELVETICA, 12, Font.BOLD));
            cell.AddElement(subHeader);
            cell.PaddingLeft = 16f;
            cell.FixedHeight = 30f;
            cell.Border = 0;
            SecondTable.AddCell(cell);

            cell = new PdfPCell();
            subHeader = new Paragraph("Loction", new Font(Font.FontFamily.HELVETICA, 12, Font.BOLD));
            cell.FixedHeight = 30f;
            cell.PaddingLeft = 16f;
            cell.Border = 0;
            cell.BorderWidthRight = 1f;
            cell.AddElement(subHeader);
            SecondTable.AddCell(cell);

     
            cell = new PdfPCell();
            Data = new Paragraph("28 / 07 / 2022");
            cell.FixedHeight = 30f;
            cell.Border = 0;
            cell.BorderWidthLeft = 1f;
            cell.BorderWidthBottom = 1f;
            cell.PaddingLeft = 6f;
            cell.AddElement(Data);
            SecondTable.AddCell(cell);


            cell = new PdfPCell();
            Data = new Paragraph("01:00");
            cell.FixedHeight = 30f;
            cell.Border = 0;
            cell.BorderWidthBottom = 1f;
            cell.PaddingLeft = 16f;
            cell.AddElement(Data);
            SecondTable.AddCell(cell);

            cell = new PdfPCell();
            Data = new Paragraph("Vadodara");
            cell.AddElement(Data);
            cell.FixedHeight = 30f;
            cell.Border = 0;
            cell.BorderWidthRight = 1f;
            cell.BorderWidthBottom = 1f;
            cell.PaddingLeft = 16f;
            SecondTable.AddCell(cell);
            
            SecondTable.WriteSelectedRows(0, -1, pdfDoc.Left + 532, pdfDoc.Top - 100, pdfWriter.DirectContent);
           // pdfDoc.Add(SecondTable);

            #endregion

            #region 3rd Table


            //PdfPTable UserTable = new PdfPTable(3);
            //UserTable.WidthPercentage = 100;
            //UserTable.HorizontalAlignment = 20;
            //UserTable.SpacingBefore = 150f;
            //UserTable.SpacingAfter = 30f;
           

            //cell = new PdfPCell();
            //Paragraph UserHeader = new Paragraph("First Name", new Font(Font.FontFamily.HELVETICA, 12, Font.BOLD));
            //cell.AddElement(UserHeader);
            //UserTable.AddCell(cell);

            //cell = new PdfPCell();
            //UserHeader = new Paragraph("Middle Name", new Font(Font.FontFamily.HELVETICA, 12, Font.BOLD));
            //cell.AddElement(UserHeader);
            //UserTable.AddCell(cell);

            //cell = new PdfPCell();
            //UserHeader = new Paragraph("Last Name", new Font(Font.FontFamily.HELVETICA, 12, Font.BOLD));
            //cell.AddElement(UserHeader);
            //UserTable.AddCell(cell);
            

            string connect = "Server=ARCHE-ITD440\\SQLEXPRESS;Database=UserDB;Trusted_Connection=True;TrustServerCertificate=True;";

            using (SqlConnection conn = new SqlConnection(connect))

            {
                string query = "SELECT FirstName, MiddleName, LastName FROM UserMst";
                SqlCommand cmd = new SqlCommand(query, conn);
                conn.Open();
                using (SqlDataReader rdr = cmd.ExecuteReader())
                {
                    while (rdr.Read())
                    {
                        PdfPTable UserTable = new PdfPTable(3);
                        UserTable.TotalWidth = 260f;
                        UserTable.LockedWidth = true;                     
                        UserTable.SpacingBefore = 90f;
                          
                        cell = new PdfPCell();
                        Header = new Paragraph("User Detail", new Font(Font.FontFamily.HELVETICA, 12, Font.BOLD));
                        cell.AddElement(Header);
                        cell.Colspan = 3;
                        cell.BorderColor = BaseColor.BLACK;
                        cell.HorizontalAlignment = 0;
                        cell.PaddingLeft = 6f;
                        cell.FixedHeight = 30f;
                        //Border
                        cell.Border = 0;
                        cell.BorderWidthLeft = 1f;
                        cell.BorderWidthTop = 1f;
                        cell.BorderWidthRight = 1f;
                        UserTable.AddCell(cell);

                        cell = new PdfPCell();
                        subHeader = new Paragraph("First Name", new Font(Font.FontFamily.HELVETICA, 12, Font.BOLD));
                        cell.AddElement(subHeader);
                        cell.PaddingLeft = 6f;
                        cell.FixedHeight = 30f;
                        //Border
                        cell.Border = 0;
                        cell.BorderWidthLeft = 1f;
                       
                        UserTable.AddCell(cell);

                        cell = new PdfPCell();
                        subHeader = new Paragraph("Middle Name", new Font(Font.FontFamily.HELVETICA, 12, Font.BOLD));
                        cell.AddElement(subHeader);
                        cell.PaddingLeft = 6f;
                        cell.FixedHeight = 30f;
                        //Border
                        cell.Border = 0;
                        UserTable.AddCell(cell);

                        cell = new PdfPCell();
                        subHeader = new Paragraph("Last Name", new Font(Font.FontFamily.HELVETICA, 12, Font.BOLD));
                        cell.AddElement(subHeader);
                        cell.PaddingLeft = 6f;
                        cell.FixedHeight = 30f;
                        //Border
                        cell.Border = 0;
                        cell.BorderWidthRight = 1f;
                        UserTable.AddCell(cell);

                        cell = new PdfPCell();
                        subHeader = new Paragraph(rdr[0].ToString());
                        cell.AddElement(subHeader);
                        cell.PaddingLeft = 16f;
                        cell.FixedHeight = 30f;
                        //Border
                        cell.Border = 0;
                        cell.BorderWidthLeft = 1f;
                        cell.BorderWidthBottom = 1f;
                        UserTable.AddCell(cell);

                        cell = new PdfPCell();
                        subHeader = new Paragraph(rdr[1].ToString());
                        cell.AddElement(subHeader);
                        cell.PaddingLeft = 16f;
                        cell.FixedHeight = 30f;
                        //Border
                        cell.Border = 0;
                        cell.BorderWidthBottom = 1f;
                        UserTable.AddCell(cell);

                        cell = new PdfPCell();
                        subHeader = new Paragraph(rdr[2].ToString());
                        cell.AddElement(subHeader);
                        cell.PaddingLeft = 16f;
                        cell.FixedHeight = 30f;
                        //Border
                        cell.Border = 0;
                        cell.BorderWidthRight = 1f;
                        cell.BorderWidthBottom = 1f;
                        UserTable.AddCell(cell);


                        // UserTable.AddCell(rdr[0].ToString());
                        // UserTable.AddCell(rdr[1].ToString());
                        // UserTable.AddCell(rdr[2].ToString());

                       // UserTable.WriteSelectedRows(0, -1, pdfDoc.Left , pdfDoc.Top - 200, pdfWriter.DirectContent);
                        pdfDoc.Add(UserTable);
                    } 
                }
            }

            #endregion

            pdfWriter.CloseStream = false;
            pdfDoc.Close();

            FS.Close();
            TempData["message"] = "PDF download in D Drive";

            return RedirectToAction("Report");
        }
    }

}


