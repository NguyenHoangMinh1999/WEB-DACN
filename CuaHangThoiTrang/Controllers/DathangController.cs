using CuaHangThoiTrang.Models;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Net;
using System.Web;
using System.Web.Mvc;

namespace CuaHangThoiTrang.Controllers
{
    public class DathangController : Controller
    {
        // GET: Dathang
        CHThoiTrangDbContext db = new CHThoiTrangDbContext();
        public ActionResult Index()
        {
            //Kiểm tra đang đăng nhập
            if (Session["UserId"] == null || Session["UserId"].ToString() == "")
            {
                return RedirectToAction("Login", "Nguoidung");
            }
            NGUOIDUNG kh = (NGUOIDUNG)Session["UserId"];
            int amaND = kh.maND;
            var dathangs = db.DATHANGs.Where(d => d.maND == amaND).ToList();
            return View(dathangs.ToList());
        }
        //Hiển thị chi tiết đơn hàng
        public ActionResult Details(int? id)
        {
            if (id == null)
            {
                return new HttpStatusCodeResult(HttpStatusCode.BadRequest);
            }
            DATHANG dathang = db.DATHANGs.Find(id);
            var chitiet = db.CHITIETDONHANGs.Where(d => d.maDH == id).ToList();
            if (dathang == null)
            {
                return HttpNotFound();
            }
            return View(chitiet);
        }
    }
}