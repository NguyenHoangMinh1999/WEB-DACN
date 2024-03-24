using CuaHangThoiTrang.Models;
using PagedList;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Web.Mvc;

namespace CuaHangThoiTrang.Areas.Admin.Controllers
{
    public class SizeController : Controller
    {
        private CHThoiTrangDbContext _context = new CHThoiTrangDbContext();
        public ActionResult Index(int? page)
        {
            if (page == null)
                page = 1;
            var all_size = _context.SIZEs.OrderBy(m => m.maSize);
            int pageSize = 5;
            int pageNum = page ?? 1;
            return View(all_size.ToPagedList(pageNum, pageSize));
        }
        public ActionResult Details(int? id)
        {
            if (id == null)
                return new HttpStatusCodeResult(System.Net.HttpStatusCode.BadRequest);
            SIZE size = _context.SIZEs.Find(id);
            if (size == null)
                return HttpNotFound();
            return View(size);
            
        }
        public ActionResult Create()
        {
            return View();
        }
        [HttpPost]
        [ValidateAntiForgeryToken]
        public ActionResult Create([Bind(Include = "maSize,tenS,mota")] SIZE size)
        {
            if (ModelState.IsValid)
            {
                _context.SIZEs.Add(size);
                _context.SaveChanges();
                return RedirectToAction("Index");
            }

            return View(size);
        }
        public ActionResult Edit(int? id)
        {
            if (id == null)
                return new HttpStatusCodeResult(System.Net.HttpStatusCode.BadRequest);
            SIZE size = _context.SIZEs.Find(id);
            if (size == null)
                return HttpNotFound();
            return View(size);
        }
        [HttpPost]
        [ValidateAntiForgeryToken]
        public ActionResult Edit([Bind(Include = "maSize,tenS,mota")] SIZE size)
        {
            if (ModelState.IsValid)
            {
                _context.Entry(size).State = System.Data.Entity.EntityState.Modified;
                _context.SaveChanges();
                return RedirectToAction("Index");
            }
            return View(size);
        }
        public ActionResult Delete(int? id)
        {
            if (id == null)
                return new HttpStatusCodeResult(System.Net.HttpStatusCode.BadRequest);
            SIZE size = _context.SIZEs.Find(id);
            if (size == null)
                return HttpNotFound();
            return View(size);
        }

        [HttpPost]
        [ValidateAntiForgeryToken]
        public ActionResult Delete(int id)
        {
            SIZE size = _context.SIZEs.Find(id);
            _context.SIZEs.Remove(size);
            _context.SaveChanges();
            return RedirectToAction("Index");
        }
    }
}