using CuaHangThoiTrang.Models;
using PagedList;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Web.Mvc;
using System.Web.UI;

namespace CuaHangThoiTrang.Controllers
{
    public class SanPhamController : Controller
    {
        // GET: SanPham
        CHThoiTrangDbContext db = new CHThoiTrangDbContext();
        public ActionResult Index(int? page)
        {
            if (page == null)
                page = 1;
            var all_sanpham = db.SANPHAMs.OrderBy(m => m.maSP);
            int pageSize = 20;
            int pageNum = page ?? 1;
            return View(all_sanpham.ToPagedList(pageNum, pageSize));
        }
    }
}