using CuaHangThoiTrang.Models;
using PagedList;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Web.Mvc;

namespace CuaHangThoiTrang.Areas.Admin.Controllers
{
    public class ThuonghieuController : Controller
    {
        private CHThoiTrangDbContext _context = new CHThoiTrangDbContext();
        public ActionResult Index(int ? page)
        {
            if(page == null)
                page = 1;
            var all_thuonghieu = _context.THUONGHIEUx.OrderBy(m => m.maTH);
            int pageSize = 5;
            int pageNum = page ?? 1;
            return View(all_thuonghieu.ToPagedList(pageNum,pageSize));
        }
        public ActionResult Details(int? id)
        {
            if (id == null)
                return new HttpStatusCodeResult(System.Net.HttpStatusCode.BadRequest);
            THUONGHIEU thuonghieu = _context.THUONGHIEUx.Find(id);
            if (thuonghieu == null)
                return HttpNotFound();
            return View(thuonghieu);
        }
        public ActionResult Create()
        {
            return View();
        }
        [HttpPost]
        [ValidateAntiForgeryToken]
        public ActionResult Create([Bind(Include = "maTH,tenTH,mota,logo")] THUONGHIEU thuonghieu)
        {
            if (ModelState.IsValid)
            {
                _context.THUONGHIEUx.Add(thuonghieu);
                _context.SaveChanges();
                return RedirectToAction("Index");
            }

            return View(thuonghieu);
        }
        public ActionResult Edit(int? id)
        {
            if (id == null)
                return new HttpStatusCodeResult(System.Net.HttpStatusCode.BadRequest);
            THUONGHIEU thuonghieu = _context.THUONGHIEUx.Find(id);
            if (thuonghieu == null)
                return HttpNotFound();
            return View(thuonghieu);
        }

        [HttpPost]
        [ValidateAntiForgeryToken]
        public ActionResult Edit([Bind(Include = "maTH,tenTH,mota,logo")] THUONGHIEU thuonghieu)
        {
            if (ModelState.IsValid)
            {
                _context.Entry(thuonghieu).State = System.Data.Entity.EntityState.Modified;
                _context.SaveChanges();
                return RedirectToAction("Index");
            }
            return View(thuonghieu);
        }
        public ActionResult Delete(int? id)
        {
            if (id == null)
                return new HttpStatusCodeResult(System.Net.HttpStatusCode.BadRequest);
            THUONGHIEU thuonghieu = _context.THUONGHIEUx.Find(id);
            if (thuonghieu == null)
                return HttpNotFound();
            return View(thuonghieu);
        }

        [HttpPost]
        [ValidateAntiForgeryToken]
        public ActionResult Delete(int id)
        {
            THUONGHIEU thuonghieu = _context.THUONGHIEUx.Find(id);
            _context.THUONGHIEUx.Remove(thuonghieu);
            _context.SaveChanges();
            return RedirectToAction("Index");
        }
    }
}