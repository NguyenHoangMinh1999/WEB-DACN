using CuaHangThoiTrang.DAO;
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
    public class HomeController : Controller
    {
        CHThoiTrangDbContext db = new CHThoiTrangDbContext();
        public static SanPhamDAO sp = new SanPhamDAO();
        IQueryable<SANPHAM> listSP = sp.ListSP();
        public ActionResult Index()
        {
            return View();
        }

        public ActionResult About()
        {
            return View();
        }

        public ActionResult Contact()
        {
            return View();
        }

        //sản phẩm
        public ActionResult SPTheoMuc(int? madm, int? page)
        {
            DANHMUC dm = db.DANHMUCs.FirstOrDefault(s => s.maDM == madm);
            ViewBag.Danhmuc = dm;
            if (page == null)
                page = 1;
            var all_sanpham = db.SANPHAMs.OrderBy(m => m.maSP).Where(n => n.maDM == madm);
            int pageSize = 20;
            int pageNum = page ?? 1;
            return View(all_sanpham.ToPagedList(pageNum, pageSize));
        }

        public ActionResult SPTheoThuongHieu(int? math, int? page)
        {
            THUONGHIEU th = db.THUONGHIEUx.FirstOrDefault(s => s.maTH == math);
            ViewBag.Thuonghieu = th;
            if (page == null)
                page = 1;
            var all_sanpham = db.SANPHAMs.OrderBy(m => m.maSP).Where(n => n.maTH == math);
            int pageSize = 20;
            int pageNum = page ?? 1;
            return View(all_sanpham.ToPagedList(pageNum, pageSize));
        }

        public ActionResult timkiem(string tensp)
        {
            SanPhamDAO sp = new SanPhamDAO();
            ViewData["TimKiem"] = sp.listspTimkiem(tensp);
            return View();
        }

        public ActionResult ChitietSP(int id)
        {
            SANPHAM sp = new SANPHAM();
            foreach (var item in listSP)
            {
                if(item.maSP == id)
                    sp = item;
            }
            ViewBag.Chitiet = sp;
            return View();
        }
    }
}