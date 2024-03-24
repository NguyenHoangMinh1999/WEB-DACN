using CuaHangThoiTrang.Models;
using PagedList;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Web.Mvc;
using System.Web.Security;

namespace CuaHangThoiTrang.Areas.Admin.Controllers
{
    public class NguoidungController : Controller
    {
        private CHThoiTrangDbContext _context = new CHThoiTrangDbContext();
        public ActionResult Index(int? page)
        {
            if (page == null)
                page = 1;
            var all_nguoidung = _context.NGUOIDUNGs.OrderBy(m => m.maND);
            int pageSize = 5;
            int pageNum = page ?? 1;
            return View(all_nguoidung.ToPagedList(pageNum, pageSize));
        }
        public ActionResult Details(int? id)
        {
            if (id == null)
                return new HttpStatusCodeResult(System.Net.HttpStatusCode.BadRequest);
            NGUOIDUNG nguoidung = _context.NGUOIDUNGs.Find(id);
            if (nguoidung == null)
                return HttpNotFound();
            return View(nguoidung);
        }
        public ActionResult Create()
        {
            return View();
        }
        [HttpPost]
        [ValidateAntiForgeryToken]
        public ActionResult Create([Bind(Include = "maND,maRole,hoten,gioitinh,email,sodienthoai,diachi,ngaysinh,password")] NGUOIDUNG nguoidung)
        {
            if (ModelState.IsValid)
            {
                _context.NGUOIDUNGs.Add(nguoidung);
                _context.SaveChanges();
                return RedirectToAction("Index");
            }

            return View(nguoidung);
        }
        public ActionResult Edit(int? id)
        {
            if (id == null)
                return new HttpStatusCodeResult(System.Net.HttpStatusCode.BadRequest);
            NGUOIDUNG nguoidung = _context.NGUOIDUNGs.Find(id);
            if (nguoidung == null)
                return HttpNotFound();
            return View(nguoidung);
        }

        [HttpPost]
        [ValidateAntiForgeryToken]
        public ActionResult Edit([Bind(Include = "maND,maRole,hoten,gioitinh,email,sodienthoai,diachi,ngaysinh,password")] NGUOIDUNG nguoidung)
        {
            if (ModelState.IsValid)
            {
                _context.Entry(nguoidung).State = System.Data.Entity.EntityState.Modified;
                _context.SaveChanges();
                return RedirectToAction("Index");
            }
            return View(nguoidung);
        }
        public ActionResult Delete(int? id)
        {
            if (id == null)
                return new HttpStatusCodeResult(System.Net.HttpStatusCode.BadRequest);
            NGUOIDUNG nguoidung = _context.NGUOIDUNGs.Find(id);
            if (nguoidung == null)
                return HttpNotFound();
            return View(nguoidung);
        }

        [HttpPost]
        [ValidateAntiForgeryToken]
        public ActionResult Delete(int id)
        {
            NGUOIDUNG nguoidung = _context.NGUOIDUNGs.Find(id);
            _context.NGUOIDUNGs.Remove(nguoidung);
            _context.SaveChanges();
            return RedirectToAction("Index");
        }
    }
}