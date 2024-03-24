using CuaHangThoiTrang.Models;
using PagedList;
using System;
using System.Collections.Generic;
using System.Drawing.Printing;
using System.Linq;
using System.Web;
using System.Web.Mvc;
using System.Web.UI;

namespace CuaHangThoiTrang.Areas.Admin.Controllers
{
    public class DanhmucController : Controller
    {
        // GET: Admin/Danhmuc
        private CHThoiTrangDbContext _context = new CHThoiTrangDbContext();
        public ActionResult Index(int ? page)
        {
            if (page == null)
                page = 1;
            var all_danhmuc = _context.DANHMUCs.OrderBy(m => m.maDM);
            int pageSize = 5;
            int pageNum = page ?? 1;
            return View(all_danhmuc.ToPagedList(pageNum, pageSize));
        }
        public ActionResult Details(int? id)
        {
            if(id == null)
                return new HttpStatusCodeResult(System.Net.HttpStatusCode.BadRequest);
            DANHMUC danhmuc = _context.DANHMUCs.Find(id);
            if(danhmuc == null)
                return HttpNotFound();
            return View(danhmuc);
        }
        public ActionResult Create()
        {
            return View();
        }
        [HttpPost]
        [ValidateAntiForgeryToken]
        public ActionResult Create([Bind(Include = "maDM,tenDM")] DANHMUC danhmuc)
        {
            if(ModelState.IsValid)
            {
                _context.DANHMUCs.Add(danhmuc);
                _context.SaveChanges();
                return RedirectToAction("Index");
            }

            return View(danhmuc);
        }
        public ActionResult Edit(int? id)
        {
            if (id == null)
                return new HttpStatusCodeResult(System.Net.HttpStatusCode.BadRequest);
            DANHMUC danhmuc = _context.DANHMUCs.Find(id);
            if (danhmuc == null)
                return HttpNotFound();
            return View(danhmuc);
        }

        [HttpPost]
        [ValidateAntiForgeryToken]
        public ActionResult Edit([Bind(Include = "maDM,tenDM")] DANHMUC danhmuc)
        {
            if (ModelState.IsValid)
            {
                _context.Entry(danhmuc).State = System.Data.Entity.EntityState.Modified;
                _context.SaveChanges();
                return RedirectToAction("Index");
            }
            return View(danhmuc);
        }
        public ActionResult Delete(int? id) 
        {
            if (id == null)
                return new HttpStatusCodeResult(System.Net.HttpStatusCode.BadRequest);
            DANHMUC danhmuc = _context.DANHMUCs.Find(id);
            if(danhmuc == null)
                return HttpNotFound();
            return View(danhmuc);
        }

        [HttpPost]
        [ValidateAntiForgeryToken]
        public ActionResult Delete(int id)
        {
            DANHMUC danhmuc = _context.DANHMUCs.Find(id);
            _context.DANHMUCs.Remove(danhmuc);
            _context.SaveChanges();
            return RedirectToAction("Index");
        }
    }
}