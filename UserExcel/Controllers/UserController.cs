using Microsoft.AspNetCore.Mvc;
using Microsoft.Data.SqlClient;
using OfficeOpenXml;
using UserExcel.Models;
using Microsoft.EntityFrameworkCore;
using Newtonsoft.Json;

namespace UserExcel.Controllers
{
    public class UserController : Controller
    {
        private readonly UserDbContext _db;
        private readonly IConfiguration _configuration;

        public UserController(UserDbContext db, IConfiguration configuration)
        {
            _db = db;
            _configuration = configuration;
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

        public IActionResult GetPDF()
        {
            return View();
        }

    }

}
