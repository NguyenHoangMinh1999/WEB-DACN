using CuaHangThoiTrang.Models;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Net;
using System.Net.Mail;
using System.Security.Cryptography;
using System.Text;
using System.Text.RegularExpressions;
using System.Web;
using System.Web.Helpers;
using System.Web.Mvc;
using System.Web.Security;

namespace CuaHangThoiTrang.Controllers
{
    public class NguoidungController : Controller
    {
        // GET: Nguoidung
        CHThoiTrangDbContext db = new CHThoiTrangDbContext();
        public static bool ValidateVNPhoneNumber(string phoneNumber)
        {
            phoneNumber = phoneNumber.Replace("+84", "0");
            Regex regex = new
            Regex(@"^(0)(86|96|97|98|32|33|34|35|36|37|38|39|91|94|83|84|85|81|82|90|93|70|79|77|76|78|92|56|58|99|59|55|87)\d{7}$");
            return regex.IsMatch(phoneNumber);
        }
        public bool ValidateEmail(string email)
        {
            Regex regex = new Regex(@"^[a-zA-Z0-9._%+-]+@[a-zA-Z0-9.-]+\.[a-zA-Z]{2,}$");
            return regex.IsMatch(email);
        }

        //đăng ký
        public ActionResult Register()
        {
            return View();
        }

        // ĐĂNG KÝ PHƯƠNG THỨC POST
        [HttpPost]
        public ActionResult Register(NGUOIDUNG nguoidung)
        {
            if (ModelState.IsValid)
            {
                nguoidung.maRole = 1;
                var check = db.NGUOIDUNGs.FirstOrDefault(s => s.email == nguoidung.email);
                if (check == null)
                {
                    nguoidung.password = md5(nguoidung.password);
                    db.Configuration.ValidateOnSaveEnabled = false;
                    db.NGUOIDUNGs.Add(nguoidung);
                    db.SaveChanges();
                    return RedirectToAction("Login");
                }
                else
                {
                    ViewBag.error = "Email already exists";
                    return View();
                }
            }
            return View();
        }
        [HttpGet]
        public ActionResult Login()
        {
            return View();
        }


        [HttpPost]
        public ActionResult Login(LoginModel model)
        {
            if (ModelState.IsValid)
            {
                var f_password = md5(model.password);
                var user = db.NGUOIDUNGs.FirstOrDefault(u => u.email == model.userMail && u.password == f_password);
                if (user != null)
                {
                    Session["UserId"] = user;
                    if (user.maRole == 2)
                    {
                        return RedirectToAction("Index", "Admin/Home");
                    }
                    else
                    {
                        return RedirectToAction("Index", "Home");
                    }

                }
                else
                {
                    ViewBag.Fail = "Tài khoản hoặc mật khẩu không chính xác.";
                }
            }
            return View("Login");

        }

        public ActionResult Logout()
        {
            Session["GioHang"] = null;
            Session["UserId"] = null;
            return RedirectToAction("Index", "Home");

        }
        public ActionResult Profile(int? id)
        {
            if (id == null)
            {
                return new HttpStatusCodeResult(HttpStatusCode.BadRequest);
            }
            NGUOIDUNG nguoiDung = db.NGUOIDUNGs.Find(id);
            if (nguoiDung == null)
            {
                return HttpNotFound();
            }
            return View(nguoiDung);
        }
        public static byte[] encryptData(string data)
        {
            System.Security.Cryptography.MD5CryptoServiceProvider md5Hasher = new System.Security.Cryptography.MD5CryptoServiceProvider();
            byte[] hashedBytes;
            System.Text.UTF8Encoding encoder = new System.Text.UTF8Encoding();
            hashedBytes = md5Hasher.ComputeHash(encoder.GetBytes(data));
            return hashedBytes;
        }
        public static string md5(string data)
        {
            MD5 md5 = new MD5CryptoServiceProvider();
            byte[] fromData = Encoding.UTF8.GetBytes(data);
            byte[] targetData = md5.ComputeHash(fromData);
            string byte2String = null;

            for (int i = 0; i < targetData.Length; i++)
            {
                byte2String += targetData[i].ToString("x2");

            }
            return byte2String;
        }
    }
}