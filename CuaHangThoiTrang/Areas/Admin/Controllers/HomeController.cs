using CuaHangThoiTrang.Models;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Security.Cryptography;
using System.Web;
using System.Web.Mvc;

namespace CuaHangThoiTrang.Areas.Admin.Controllers
{
    public class HomeController : Controller
    {
        // GET: Admin/Home
        CHThoiTrangDbContext db =new CHThoiTrangDbContext();
        public ActionResult Index()
        {
            if (Session["UserId"] == null)
            {
                return HttpNotFound();
            }
            else
            {
                var user = (NGUOIDUNG)Session["UserId"];
                if(user.maRole != 2)
                    return HttpNotFound();
                DateTime now = DateTime.Now;
                ViewBag.datamonth = db.DATHANGs.OrderBy(m => m.maDH).Where(t => t.ngayDat.Month == now.Month).Sum(m => m.tongGiaTri);
                ViewBag.dathang = db.DATHANGs.Sum(m => m.tongGiaTri);
                ViewBag.request = db.DATHANGs.OrderBy(m => m.maDH).Count(d => d.trangThai == 0);
                return View();  
            }
        }
    }
}