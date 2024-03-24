using CuaHangThoiTrang.Models;
using PagedList;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Web.Mvc;

namespace CuaHangThoiTrang.Areas.Admin.Controllers
{
    public class MauController : Controller
    {
        private CHThoiTrangDbContext _context = new CHThoiTrangDbContext();
        public ActionResult Index(int? page)
        {
            if (page == null)
                page = 1;
            var all_mau = _context.MAUs.OrderBy(m => m.maMau);
            int pageSize = 5;
            int pageNum = page ?? 1;
            return View(all_mau.ToPagedList(pageNum, pageSize));
        }
        public ActionResult Details(int? id)
        {
            if (id == null)
                return new HttpStatusCodeResult(System.Net.HttpStatusCode.BadRequest);
            MAU mau = _context.MAUs.Find(id);
            if (mau == null)
                return HttpNotFound();
            return View(mau);

        }
        public ActionResult Create()
        {
            return View();
        }
        [HttpPost]
        [ValidateAntiForgeryToken]
        public ActionResult Create([Bind(Include = "maMau,maSP,tenMau")] MAU mau)
        {
            if (ModelState.IsValid)
            {
                _context.MAUs.Add(mau);
                _context.SaveChanges();
                return RedirectToAction("Index");
            }

            return View(mau);
        }
        public ActionResult Edit(int? id)
        {
            if (id == null)
                return new HttpStatusCodeResult(System.Net.HttpStatusCode.BadRequest);
            MAU mau = _context.MAUs.Find(id);
            if (mau == null)
                return HttpNotFound();
            return View(mau);
        }

        [HttpPost]
        [ValidateAntiForgeryToken]
        public ActionResult Edit([Bind(Include = "maMau,maSP,tenMau")] MAU mau)
        {
            if (ModelState.IsValid)
            {
                _context.Entry(mau).State = System.Data.Entity.EntityState.Modified;
                _context.SaveChanges();
                return RedirectToAction("Index");
            }
            return View(mau);
        }
        public ActionResult Delete(int? id)
        {
            if (id == null)
                return new HttpStatusCodeResult(System.Net.HttpStatusCode.BadRequest);
            MAU mau = _context.MAUs.Find(id);
            if (mau == null)
                return HttpNotFound();
            return View(mau);
        }

        [HttpPost]
        [ValidateAntiForgeryToken]
        public ActionResult Delete(int id)
        {
            MAU mau = _context.MAUs.Find(id);
            _context.MAUs.Remove(mau);
            _context.SaveChanges();
            return RedirectToAction("Index");
        }
    }
}