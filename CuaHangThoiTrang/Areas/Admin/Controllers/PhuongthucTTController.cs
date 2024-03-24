using CuaHangThoiTrang.Models;
using PagedList;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Web.Mvc;

namespace CuaHangThoiTrang.Areas.Admin.Controllers
{
    public class PhuongthucTTController : Controller
    {
        // GET: Admin/PhuongthucTT
        private CHThoiTrangDbContext _context = new CHThoiTrangDbContext();
        public ActionResult Index(int? page)
        {
            if (page == null)
                page = 1;
            var all_phuongthuc = _context.PHUONGTHUCTHANHTOANs.OrderBy(m => m.maPT);
            int pageSize = 5;
            int pageNum = page ?? 1;
            return View(all_phuongthuc.ToPagedList(pageNum, pageSize));
        }
        public ActionResult Details(int? id)
        {
            if (id == null)
                return new HttpStatusCodeResult(System.Net.HttpStatusCode.BadRequest);
            PHUONGTHUCTHANHTOAN phuongthuc = _context.PHUONGTHUCTHANHTOANs.Find(id);
            if (phuongthuc == null)
                return HttpNotFound();
            return View(phuongthuc);
        }
        public ActionResult Create()
        {
            return View();
        }
        [HttpPost]
        [ValidateAntiForgeryToken]
        public ActionResult Create([Bind(Include = "maPT,tenPT")] PHUONGTHUCTHANHTOAN phuongthuc)
        {
            if (ModelState.IsValid)
            {
                _context.PHUONGTHUCTHANHTOANs.Add(phuongthuc);
                _context.SaveChanges();
                return RedirectToAction("Index");
            }

            return View(phuongthuc);
        }
        public ActionResult Edit(int? id)
        {
            if (id == null)
                return new HttpStatusCodeResult(System.Net.HttpStatusCode.BadRequest);
            PHUONGTHUCTHANHTOAN phuongthuc = _context.PHUONGTHUCTHANHTOANs.Find(id);
            if (phuongthuc == null)
                return HttpNotFound();
            return View(phuongthuc);
        }

        [HttpPost]
        [ValidateAntiForgeryToken]
        public ActionResult Edit([Bind(Include = "maPT,tenPT")] PHUONGTHUCTHANHTOAN phuongthuc)
        {
            if (ModelState.IsValid)
            {
                _context.Entry(phuongthuc).State = System.Data.Entity.EntityState.Modified;
                _context.SaveChanges();
                return RedirectToAction("Index");
            }
            return View(phuongthuc);
        }
        public ActionResult Delete(int? id)
        {
            if (id == null)
                return new HttpStatusCodeResult(System.Net.HttpStatusCode.BadRequest);
            PHUONGTHUCTHANHTOAN phuongthuc = _context.PHUONGTHUCTHANHTOANs.Find(id);
            if (phuongthuc == null)
                return HttpNotFound();
            return View(phuongthuc);
        }

        [HttpPost]
        [ValidateAntiForgeryToken]
        public ActionResult Delete(int id)
        {
            PHUONGTHUCTHANHTOAN phuongthuc = _context.PHUONGTHUCTHANHTOANs.Find(id);
            _context.PHUONGTHUCTHANHTOANs.Remove(phuongthuc);
            _context.SaveChanges();
            return RedirectToAction("Index");
        }
    }
}