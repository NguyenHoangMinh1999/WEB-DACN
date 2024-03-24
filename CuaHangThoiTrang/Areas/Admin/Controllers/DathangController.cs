using CuaHangThoiTrang.DAO;
using CuaHangThoiTrang.Models;
using MLVshop.DAO;
using System;
using System.Collections.Generic;
using System.Data.Entity;
using System.Linq;
using System.Net;
using System.Net.Mail;
using System.Runtime.InteropServices;
using System.Web;
using System.Web.Mvc;
using Excel = Microsoft.Office.Interop.Excel;

namespace CuaHangThoiTrang.Areas.Admin.Controllers
{
    public class DathangController : Controller
    {
        CHThoiTrangDbContext db = new CHThoiTrangDbContext();
        // GET: Admin/Dathang
        public ActionResult Index()
        {
            var donhangs = db.DATHANGs.Include(d => d.NGUOIDUNG);
            return View(donhangs.ToList());
        }
        public ActionResult Details(int? id)
        {
            if (id == null)
            {
                return new HttpStatusCodeResult(HttpStatusCode.BadRequest);
            }
            DATHANG dathang = db.DATHANGs.Find(id);
            if (dathang == null)
            {
                return HttpNotFound();
            }
            return View(dathang);
        }
 
        public ActionResult Xacnhan(int? id)
        {
            if (id == null)
            {
                return new HttpStatusCodeResult(HttpStatusCode.BadRequest);
            }
            DATHANG dathang = db.DATHANGs.Find(id);

            if (dathang == null)
            {
                return HttpNotFound();
            }
            dathang.trangThai = 1;
            db.SaveChanges();

            var chitietdonhang = db.CHITIETDONHANGs.Where(x => x.maDH == id).ToList();
            foreach (var ctdh in chitietdonhang)
            {
                // update so luong
                SANPHAM sp = db.SANPHAMs.FirstOrDefault(x => x.maSP == ctdh.maSP);
                sp.soLuong = sp.soLuong - ctdh.soluong;
                db.SaveChanges();
            }
            return RedirectToAction("Index");
        }

        public ActionResult Create()
        {
            ViewBag.MaNguoidung = new SelectList(db.NGUOIDUNGs, "MaND");
            return View();
        }

        [HttpPost]
        [ValidateAntiForgeryToken]
        public ActionResult Create([Bind(Include = "MaDH,ngayDat,trangThai,maPT,maND,diachi")] DATHANG dathang)
        {
            if (ModelState.IsValid)
            {
                db.DATHANGs.Add(dathang);
                db.SaveChanges();
                return RedirectToAction("Index");
            }

            ViewBag.MaNguoidung = new SelectList(db.NGUOIDUNGs, "maND", "hotenKH", dathang.maND);
            return View(dathang);
        }

        [HttpGet] 
        public ActionResult Edit(int? id)
        {
            if (id == null)
            {
                return new HttpStatusCodeResult(HttpStatusCode.BadRequest);
            }
            DATHANG dathang = db.DATHANGs.Find(id);
            if (dathang == null)
            {
                return HttpNotFound();
            }
            ViewBag.MaNguoidung = new SelectList(db.NGUOIDUNGs, "maND", "hotenKH", dathang.maND);
            return View(dathang);
        }
        [HttpPost]
        [ValidateAntiForgeryToken]
        public ActionResult Edit([Bind(Include = "Madon,Ngaydat,Tinhtrang,Thanhtoan,MaNguoidung,Diachinhanhang")] DATHANG dathang)
        {
            if (ModelState.IsValid)
            {
                db.Entry(dathang).State = EntityState.Modified;
                db.SaveChanges();
                return RedirectToAction("Index");
            }
            ViewBag.MaNguoidung = new SelectList(db.NGUOIDUNGs, "maND", "hotenKH", dathang.maND);
            return View(dathang);
        }
        [HttpGet]
        public ActionResult Delete(int? id)
        {
            if (id == null)
            {
                return new HttpStatusCodeResult(HttpStatusCode.BadRequest);
            }
            DATHANG donhang = db.DATHANGs.Find(id);
            if (donhang == null)
            {
                return HttpNotFound();
            }
            return View(donhang);
        }
        [HttpPost, ActionName("Delete")]
        [ValidateAntiForgeryToken]
        public ActionResult DeleteConfirmed(int id)
        {
            DATHANG dathang = db.DATHANGs.Find(id);
            db.DATHANGs.Remove(dathang);
            db.SaveChanges();
            return RedirectToAction("Index");
        }

        public ActionResult Export()
        {
            try
            {
                Excel.Application application = new Excel.Application();
                Excel.Workbook workbook = application.Workbooks.Add(System.Reflection.Missing.Value);
                Excel.Worksheet worksheet = workbook.ActiveSheet;
                DonHangDAO dh = new DonHangDAO();
                worksheet.Cells[1, 1] = "Họ tên khách hàng";
                worksheet.Cells[1, 2] = "Ngày đặt";
                worksheet.Cells[1, 3] = "Tình trạng đơn hàng";
                worksheet.Cells[1, 4] = "Tổng giá trị";
                worksheet.Cells[1, 5] = "Phương thức thanh toán";
                worksheet.Cells[1, 6] = "Địa chỉ nhận hàng";
                
                int row = 2;
                foreach (DATHANG ddh in dh.ListDH())
                {
                    worksheet.Cells[row, 1] = ddh.hotenKH;
                    worksheet.Cells[row, 2] = ddh.ngayDat;
                    if(ddh.trangThai == 1)
                        worksheet.Cells[row, 3] = "Đã xác nhận";
                    else
                        worksheet.Cells[row, 3] = "Chưa xác nhận";

                    worksheet.Cells[row, 4] = ddh.tongGiaTri;                    
                    worksheet.Cells[row, 5] = ddh.PHUONGTHUCTHANHTOAN.tenPT;
                    worksheet.Cells[row, 6] = ddh.diaChi;                  
                    row++;
                }

                worksheet.get_Range("A1", "F1").EntireColumn.AutoFit();
                var range_heading = worksheet.get_Range("A1", "F1");
                range_heading.Font.Bold = true;
                range_heading.Font.Color = System.Drawing.Color.Red;
                range_heading.Font.Size = 13;

                var range_currency = worksheet.get_Range("D2", "D100");
                range_currency.NumberFormat = "###,###,###";


                workbook.SaveAs("F:\\DACN\\dondathang.xls");
                workbook.Close();
                Marshal.ReleaseComObject(workbook);

                application.Quit();
                Marshal.FinalReleaseComObject(application);
                ViewBag.Result = "Done";
            }
            catch (Exception ex)
            {
                ViewBag.Result = ex.Message;
            }
            return View("Success");
        }



        

    }
}