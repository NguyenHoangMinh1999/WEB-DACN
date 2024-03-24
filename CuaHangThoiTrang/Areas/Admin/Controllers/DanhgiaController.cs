using CuaHangThoiTrang.Models;
using PagedList;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Web.Mvc;

namespace CuaHangThoiTrang.Areas.Admin.Controllers
{
    public class DanhgiaController : Controller
    {
        private CHThoiTrangDbContext _context = new CHThoiTrangDbContext();

        public ActionResult Index(int? page)
        {
            if (page == null)
                page = 1;
            var all_danhgia = _context.DANHGIAs.OrderBy(m => m.maDG);
            int pageSize = 5;
            int pageNum = page ?? 1;
            return View(all_danhgia.ToPagedList(pageNum, pageSize));
        }

        public ActionResult Details(int? id)
        {
            if (id == null)
                return new HttpStatusCodeResult(System.Net.HttpStatusCode.BadRequest);
            DANHGIA danhgia = _context.DANHGIAs.Find(id);
            if (danhgia == null)
                return HttpNotFound();
            return View(danhgia);
        }

        public ActionResult Create()
        {
            return View();
        }
        [HttpPost]
        [ValidateAntiForgeryToken]
        public ActionResult Create([Bind(Include = "maDG,maSP,noidungDG,sosao")] DANHGIA danhgia)
        {
            if (ModelState.IsValid)
            {
                _context.DANHGIAs.Add(danhgia);
                _context.SaveChanges();
                return RedirectToAction("Index");
            }

            return View(danhgia);
        }

        public ActionResult Edit(int? id)
        {
            if (id == null)
                return new HttpStatusCodeResult(System.Net.HttpStatusCode.BadRequest);
            DANHGIA danhgia = _context.DANHGIAs.Find(id);
            if (danhgia == null)
                return HttpNotFound();
            return View(danhgia);
        }

        [HttpPost]
        [ValidateAntiForgeryToken]
        public ActionResult Edit([Bind(Include = "maDG,maSP,noidungDG,sosao")] DANHGIA danhgia)
        {
            if (ModelState.IsValid)
            {
                _context.Entry(danhgia).State = System.Data.Entity.EntityState.Modified;
                _context.SaveChanges();
                return RedirectToAction("Index");
            }
            return View(danhgia);
        }

        public ActionResult Delete(int? id)
        {
            if (id == null)
                return new HttpStatusCodeResult(System.Net.HttpStatusCode.BadRequest);
            DANHGIA danhgia = _context.DANHGIAs.Find(id);
            if (danhgia == null)
                return HttpNotFound();
            return View(danhgia);
        }

        [HttpPost]
        [ValidateAntiForgeryToken]
        public ActionResult Delete(int id)
        {
            DANHGIA danhgia = _context.DANHGIAs.Find(id);
            _context.DANHGIAs.Remove(danhgia);
            _context.SaveChanges();
            return RedirectToAction("Index");
        }
    }
}