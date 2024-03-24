using CuaHangThoiTrang.DAO;
using CuaHangThoiTrang.Models;
using PagedList;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.InteropServices;
using System.Web;
using System.Web.Mvc;
using Excel = Microsoft.Office.Interop.Excel;
using System.IO;
using System.Data.SqlClient;
using System.Configuration;
using System.Data.OleDb;
using System.Data;
using OfficeOpenXml;

namespace CuaHangThoiTrang.Areas.Admin.Controllers
{
    public class SanphamController : Controller
    {
        private CHThoiTrangDbContext _context = new CHThoiTrangDbContext();
        public string ProcessUpload(HttpPostedFileBase file)
        {
            if (file == null)
            {
                return "";
            }
            file.SaveAs(Server.MapPath("~/Content/img/" + file.FileName));
            return "/Content/img/" + file.FileName;
        }
        public ActionResult Index(int? page)
        {                     
            if (page == null)
                page = 1;
            var all_sanpham = _context.SANPHAMs.OrderBy(m => m.maSP);
            int pageSize = 5;
            int pageNum = page ?? 1;
            return View(all_sanpham.ToPagedList(pageNum, pageSize));


        }

        [HttpPost]
        public ActionResult Index(HttpPostedFileBase file, int? page)
        {
            string filename = Guid.NewGuid() + Path.GetExtension(file.FileName);
            string filepath = "/excelfolder/" + filename;
            file.SaveAs(Path.Combine(Server.MapPath("/excelfolder"), filename));
            InsertExceldata(filepath, filename);
            if (page == null)
                page = 1;
            var all_sanpham = _context.SANPHAMs.OrderBy(m => m.maSP);
            int pageSize = 5;
            int pageNum = page ?? 1;
            return View(all_sanpham.ToPagedList(pageNum, pageSize));

        }
        public ActionResult Details(int? id)
        {
            if (id == null)
                return new HttpStatusCodeResult(System.Net.HttpStatusCode.BadRequest);
            SANPHAM sanpham = _context.SANPHAMs.Find(id);
            if (sanpham == null)
                return HttpNotFound();
            return View(sanpham);

        }
        public ActionResult Create()
        {
            return View();
        }
        [HttpPost]
        [ValidateAntiForgeryToken]
        public ActionResult Create([Bind(Include = "maSP,maTH,maDM,tenSP,gia,hinh,hinh1,hinh2,soLuong,thongTin")] SANPHAM sanpham)
        {
            if (ModelState.IsValid)
            {
                _context.SANPHAMs.Add(sanpham);
                _context.SaveChanges();
                return RedirectToAction("Index");
            }

            return View(sanpham);
        }
        public ActionResult Edit(int? id)
        {
            if (id == null)
                return new HttpStatusCodeResult(System.Net.HttpStatusCode.BadRequest);
            SANPHAM sanpham = _context.SANPHAMs.Find(id);
            if (sanpham == null)
                return HttpNotFound();
            return View(sanpham);
        }

        [HttpPost]
        [ValidateAntiForgeryToken]
        public ActionResult Edit([Bind(Include = "maSP,maTH,maDM,tenSP,gia,hinh,hinh1,hinh2,soLuong,thongTin")] SANPHAM sanpham)
        {
            if (ModelState.IsValid)
            {
                _context.Entry(sanpham).State = System.Data.Entity.EntityState.Modified;
                _context.SaveChanges();
                return RedirectToAction("Index");
            }
            return View(sanpham);
        }
        public ActionResult Delete(int? id)
        {
            if (id == null)
                return new HttpStatusCodeResult(System.Net.HttpStatusCode.BadRequest);
            SANPHAM sanpham = _context.SANPHAMs.Find(id);
            if (sanpham == null)
                return HttpNotFound();
            return View(sanpham);
        }

        [HttpPost]
        [ValidateAntiForgeryToken]
        public ActionResult Delete(int id)
        {
            SANPHAM sanpham = _context.SANPHAMs.Find(id);
            _context.SANPHAMs.Remove(sanpham);
            _context.SaveChanges();
            return RedirectToAction("Index");
        }
        public ActionResult SearchSP(string searchString)
        {
            var sp = (from ss in _context.SANPHAMs select  ss).Where(p=> p.tenSP.Contains(searchString));
            return View(sp.ToList());
        }


        public ActionResult Export()
        {
            try
            {
                Excel.Application application = new Excel.Application();
                Excel.Workbook workbook = application.Workbooks.Add(System.Reflection.Missing.Value);
                Excel.Worksheet worksheet = workbook.ActiveSheet;
                SanPhamDAO spd = new SanPhamDAO();
                //worksheet.Cells[1, 1] = "Mã sản phẩm";
                worksheet.Cells[1, 1] = "Mã thương hiệu";
                worksheet.Cells[1, 2] = "Mã danh mục";
                worksheet.Cells[1, 3] = "Tên sản phẩm";
                worksheet.Cells[1, 4] = "Giá";
                worksheet.Cells[1, 5] = "Hình 1";
                worksheet.Cells[1, 6] = "Hình 2";
                worksheet.Cells[1, 7] = "Hình 3";
                worksheet.Cells[1, 8] = "Số lượng";
                worksheet.Cells[1, 9] = "Thông tin";
                int row = 2;
                foreach (SANPHAM sp in spd.ListSP())
                {
                    //worksheet.Cells[row, 1] = sp.maSP;
                    worksheet.Cells[row, 1] = sp.maTH;
                    worksheet.Cells[row, 2] = sp.maDM;
                    worksheet.Cells[row, 3] = sp.tenSP;
                    worksheet.Cells[row, 4] = sp.gia;
                    worksheet.Cells[row, 5] = sp.hinh;
                    worksheet.Cells[row, 6] = sp.hinh1;
                    worksheet.Cells[row, 7] = sp.hinh2;
                    worksheet.Cells[row, 8] = sp.soLuong;
                    worksheet.Cells[row, 9] = sp.thongTin;
                    row++;
                }

                worksheet.get_Range("A1", "I1").EntireColumn.AutoFit();
                var range_heading = worksheet.get_Range("A1", "I1");
                range_heading.Font.Bold = true;
                range_heading.Font.Color = System.Drawing.Color.Red;
                range_heading.Font.Size = 13;

                var range_currency = worksheet.get_Range("D2", "D100");
                range_currency.NumberFormat = "###,###,###";

                          
                workbook.SaveAs("F:\\DACN\\sanpham.xls");
                workbook.Close();
                Marshal.ReleaseComObject(workbook);

                application.Quit();
                Marshal.FinalReleaseComObject(application);
                ViewBag.Result = "Done";
            }
            catch(Exception ex)
            {
                ViewBag.Result = ex.Message;
            }
            return View("Success");
        }

        SqlConnection con = new SqlConnection(ConfigurationManager.ConnectionStrings["CHThoiTrangDbContext"].ConnectionString);
        OleDbConnection Econ;

        private void ExcelConn(string filepath)
        {
            string constr = string.Format(@"Provider=Microsoft.ACE.OLEDB.12.0;Data Source={0};Extended Properties= ""Excel 12.0 Xml;HDR=YES;""", filepath);
            Econ = new OleDbConnection(constr);
        }

        private void InsertExceldata(string fileepath, string filename)
        {
            string fullpath = Server.MapPath("/excelfolder/") + filename;
            ExcelConn(fullpath);
            string query = string.Format("Select * from [{0}]", "Sheet1$");
            OleDbCommand Ecom = new OleDbCommand(query, Econ);
            Econ.Open();

            DataSet ds = new DataSet();
            OleDbDataAdapter oda = new OleDbDataAdapter(query, Econ);
            Econ.Close();
            oda.Fill(ds);

            DataTable dt = ds.Tables[0];

            

            SqlBulkCopy objbulk = new SqlBulkCopy(con);
            objbulk.DestinationTableName = "SANPHAM";

            objbulk.ColumnMappings.Add("Mã thương hiệu", "maTH");
            objbulk.ColumnMappings.Add("Mã danh mục", "maDM");
            objbulk.ColumnMappings.Add("Tên sản phẩm", "tenSP");
            objbulk.ColumnMappings.Add("Giá", "gia");
            objbulk.ColumnMappings.Add("Hình 1", "hinh");
            objbulk.ColumnMappings.Add("Hình 2", "hinh1");
            objbulk.ColumnMappings.Add("Hình 3", "hinh2");
            objbulk.ColumnMappings.Add("Số lượng", "soLuong");
            objbulk.ColumnMappings.Add("Thông tin", "thongTin");
            con.Open();

            objbulk.WriteToServer(dt);
            con.Close();
        }


        
    }



}
