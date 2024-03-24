using CuaHangThoiTrang.Models;
using CuaHangThoiTrang.Orther;
using CuaHangThoiTrang.Orther.Momo;
using Newtonsoft.Json.Linq;
using System;
using System.Collections.Generic;
using System.Configuration;
using System.Linq;
using System.Web;
using System.Web.Helpers;
using System.Web.Mvc;
using System.IO;
using System.Net.Mail;
using System.Net;



namespace CuaHangThoiTrang.Controllers
{
    public class GioHangController : Controller
    {
        // GET: GioHang
        CHThoiTrangDbContext db = new CHThoiTrangDbContext();

        //Lấy giỏ hàng 
        public List<GIOHANG> LayGioHang()
        {
            List<GIOHANG> lstGioHang = Session["GioHang"] as List<GIOHANG>;
            if (lstGioHang == null)
            {
                //Nếu giỏ hàng chưa tồn tại thì mình tiến hành khởi tao list giỏ hàng (sessionGioHang)
                lstGioHang = new List<GIOHANG>();
                Session["GioHang"] = lstGioHang;
            }
            return lstGioHang;
        }
        //Thêm giỏ hàng
        public ActionResult ThemGioHang(int iMasp, string strURL)
        {
            if (Session["UserId"] == null || Session["UserId"].ToString() == "")
            {
                return RedirectToAction("Login", "Nguoidung");
            }
            //Kiểm tra giỏ hàng
            if (Session["GioHang"] == null)
            {
                RedirectToAction("Index", "Home");
            }
            SANPHAM sp = db.SANPHAMs.SingleOrDefault(n => n.maSP == iMasp);
            if (sp == null)
            {
                Response.StatusCode = 404;
                return null;
            }
            //Lấy ra session giỏ hàng
            List<GIOHANG> lstGioHang = LayGioHang();
            //Kiểm tra sp này đã tồn tại trong session[giohang] chưa
            GIOHANG sanpham = lstGioHang.Find(n => n.imaSP == iMasp);
            if (sanpham == null)
            {
                sanpham = new GIOHANG(iMasp);
                //Add sản phẩm mới thêm vào list
                lstGioHang.Add(sanpham);
                return Redirect(strURL);
            }
            else
            {
                sanpham.isoluong++;
                return Redirect(strURL);
            }
        }
        public ActionResult CapNhatGioHang(int id, FormCollection collection)
        {
            var sp = db.SANPHAMs.FirstOrDefault(m => m.maSP == id); // lấy sách dựa trên mã
            var txtsl = Int32.Parse(collection["txtSoLg"]);// số lượng tử form nhập vào
            List<GIOHANG> lstGiohang = LayGioHang();
            GIOHANG sanpham = lstGiohang.SingleOrDefault(n => n.imaSP == id);
            if (sanpham != null)
            {
                if (txtsl > 0) //kiểm tra số lượng lớn hơn 0
                {
                    if (sp.soLuong < txtsl) // nếu số lượng tồn nhỏ hơn số lượng nhập vào
                    {
                        TempData["Error"] = "Số lượng tồn không đủ";//báo lỗi
                    }
                    else
                    {
                        sanpham.isoluong = txtsl;//gán số lượng cho sanpham
                    }
                }
                else if (txtsl == 0) // nếu số lượng nhập vào là 0
                {
                    XoaGioHang(id);// xóa sản phẩm đó ra khỏi giỏ hàng

                }
                else
                {
                    TempData["Error"] = "số lượng không hợp lệ";// lỗi số lượng âm
                }

            }
            return RedirectToAction("GioHang");
        }
        //Xóa giỏ hàng
        public ActionResult XoaGioHang(int iMaSP)
        {
            //Kiểm tra masp
            SANPHAM sp = db.SANPHAMs.SingleOrDefault(n => n.maSP == iMaSP);
            //Nếu get sai masp thì sẽ trả về trang lỗi 404
            if (sp == null)
            {
                Response.StatusCode = 404;
                return null;
            }
            //Lấy giỏ hàng ra từ session
            List<GIOHANG> lstGioHang = LayGioHang();
            GIOHANG sanpham = lstGioHang.SingleOrDefault(n => n.imaSP == iMaSP);
            //Nếu mà tồn tại thì chúng ta cho sửa số lượng
            if (sanpham != null)
            {
                lstGioHang.RemoveAll(n => n.imaSP == iMaSP);

            }
            if (lstGioHang.Count == 0)
            {
                return RedirectToAction("Index", "Home");
            }
            return RedirectToAction("GioHang");
        }
        //Xây dựng trang giỏ hàng
        public ActionResult GioHang()
        {
            if (Session["GioHang"] == null)
            {
                return RedirectToAction("Index", "Home");
            }
            List<GIOHANG> lstGioHang = LayGioHang();
            return View(lstGioHang);
        }
        //Tính tổng số lượng và tổng tiền
        //Tính tổng số lượng
        private int TongSoLuong()
        {
            int iTongSoLuong = 0;
            List<GIOHANG> lstGioHang = Session["GioHang"] as List<GIOHANG>;
            if (lstGioHang != null)
            {
                iTongSoLuong = lstGioHang.Sum(n => n.isoluong);
            }
            return iTongSoLuong;
        }

        //Tính tổng thành tiền
        private double TongTien()
        {
            double dTongTien = 0;
            List<GIOHANG> lstGioHang = Session["GioHang"] as List<GIOHANG>;
            if (lstGioHang != null)
            {
                dTongTien = lstGioHang.Sum(n => n.dthanhtien);
            }
            return dTongTien;
        }
        //tạo partial giỏ hàng
        public ActionResult GioHangPartial()
        {
            if (TongSoLuong() == 0)
            {
                return PartialView();
            }
            ViewBag.TongSoLuong = TongSoLuong();
            ViewBag.TongTien = TongTien();
            return PartialView();
        }
        //Xây dựng 1 view cho người dùng chỉnh sửa giỏ hàng
        public ActionResult SuaGioHang()
        {
            if (Session["GioHang"] == null)
            {
                return RedirectToAction("Index", "Home");
            }
            List<GIOHANG> lstGioHang = LayGioHang();
            return View(lstGioHang);

        }

        #region 
        //Xây dựng chức năng đặt hàng
        [HttpPost]
        public ActionResult DatHang(FormCollection donhangForm)
        {
            ViewBag.Phuongthuc = db.PHUONGTHUCTHANHTOANs.ToList();
            //Kiểm tra đăng đăng nhập
            if (Session["UserId"] == null || Session["UserId"].ToString() == "")
            {
                return RedirectToAction("Login", "Nguoidung");
            }
            //Kiểm tra giỏ hàng
            if (Session["GioHang"] == null)
            {
                RedirectToAction("Index", "Home");
            }
            Console.WriteLine(donhangForm);
            string diachinhanhang = donhangForm["diachi"].ToString();
            Session["diachinhanhang"] = diachinhanhang;
            string thanhtoan = donhangForm["MaTT"].ToString();
            int ptthanhtoan = Int32.Parse(thanhtoan);
            if (ptthanhtoan == 2)
            {
                return RedirectToAction("Momo", "GioHang");
            } 
            //Thêm đơn hàng
            DATHANG ddh = new DATHANG();
            NGUOIDUNG kh = (NGUOIDUNG)Session["UserId"];
            List<GIOHANG> gh = LayGioHang();
            ddh.maND = kh.maND;
            ddh.hotenKH = kh.hoten;
            ddh.email = kh.email;
            ddh.ngayDat = DateTime.Now;
            ddh.maPT = ptthanhtoan;
            ddh.diaChi = diachinhanhang;
            decimal tongtien = 0;
            foreach (var item in gh)
            {
                decimal thanhtien = item.isoluong * (decimal)item.dgia;
                tongtien += thanhtien;
            }
            ddh.tongGiaTri = (decimal)tongtien;
            ddh.trangThai = 0;
            db.DATHANGs.Add(ddh);
            db.SaveChanges();
            //Thêm chi tiết đơn hàng
            foreach (var item in gh)
            {
                CHITIETDONHANG ctDH = new CHITIETDONHANG();
                decimal thanhtien = item.isoluong * (decimal)item.dgia;
                ctDH.maDH = ddh.maDH;
                ctDH.maSP = item.imaSP;
                ctDH.soluong = item.isoluong;
                ctDH.gia = (decimal)item.dgia;
                ctDH.tonggiatri = thanhtien;
                db.CHITIETDONHANGs.Add(ctDH);
            }
            db.SaveChanges();
            //SendOrderConfirmationEmail();
            Session["GioHang"] = null;
            return RedirectToAction("Index", "Dathang");
           
        }
        #endregion

        public ActionResult ThanhToanDonHang()
        {
            ViewBag.MaTT = new SelectList(db.PHUONGTHUCTHANHTOANs, "maPT", "tenPT",1);
            

            //Kiểm tra đăng đăng nhập
            if (Session["UserId"] == null || Session["UserId"].ToString() == "")
            {
                return RedirectToAction("Login", "Nguoidung");
            }
            //Kiểm tra giỏ hàng
            if (Session["GioHang"] == null)
            {
                RedirectToAction("Index", "Home");
            }

            //Thêm đơn hàng
            DATHANG ddh = new DATHANG();
            NGUOIDUNG kh = (NGUOIDUNG)Session["UserId"];
            ViewBag.MaNguoiDung = new SelectList(db.NGUOIDUNGs, "maND", "hoten");
            List<GIOHANG> gh = LayGioHang();
            decimal tongtien = 0;
            foreach (var item in gh)
            {
                decimal thanhtien = item.isoluong * (decimal)item.dgia;
                tongtien += thanhtien;
            }
            ddh.maND = kh.maND;
            ddh.ngayDat = DateTime.Now;
            CHITIETDONHANG ctDH = new CHITIETDONHANG();
            ViewBag.tongtien = tongtien;

            

            return View(ddh);
            /*db.SaveChanges*/
        }
        public ActionResult Momo()
        {
            List<GIOHANG> gh = LayGioHang();
            decimal tongtien = 0;
            foreach (var item in gh)
            {
                decimal thanhtien = item.isoluong * (decimal)item.dgia;
                tongtien += thanhtien;
            }
            //request params need to request to MoMo system
            string endpoint = "https://test-payment.momo.vn/gw_payment/transactionProcessor";
            string partnerCode = "MOMOOJOI20210710";
            string accessKey = "iPXneGmrJH0G8FOP";
            string serectkey = "sFcbSGRSJjwGxwhhcEktCHWYUuTuPNDB";
            string orderInfo = "Thanh toán đơn hàng";
            string returnUrl = "https://localhost:44391/GioHang/Thanhtoanmomo";
            string notifyurl = "https://4c8d-2001-ee0-5045-50-58c1-b2ec-3123-740d.ap.ngrok.io/GioHang/LuuThanhToan";

            string amount = tongtien.ToString();
            string orderid = DateTime.Now.Ticks.ToString();
            string requestId = DateTime.Now.Ticks.ToString();
            string extraData = "";

            //Before sign HMAC SHA256 signature
            string rawHash = "partnerCode=" +
                partnerCode + "&accessKey=" +
                accessKey + "&requestId=" +
                requestId + "&amount=" +
                amount + "&orderId=" +
                orderid + "&orderInfo=" +
                orderInfo + "&returnUrl=" +
                returnUrl + "&notifyUrl=" +
                notifyurl + "&extraData=" +
                extraData;

            MoMoSecurity crypto = new MoMoSecurity();
            //sign signature SHA256
            string signature = crypto.signSHA256(rawHash, serectkey);

            //build body json request
            JObject message = new JObject
            {
                { "partnerCode", partnerCode },
                { "accessKey", accessKey },
                { "requestId", requestId },
                { "amount", amount },
                { "orderId", orderid },
                { "orderInfo", orderInfo },
                { "returnUrl", returnUrl },
                { "notifyUrl", notifyurl },
                { "extraData", extraData },
                { "requestType", "captureMoMoWallet" },
                { "signature", signature }
            };
            string responseFromMomo = PaymentRequest.sendPaymentRequest(endpoint, message.ToString());

            JObject jmessage = JObject.Parse(responseFromMomo);

            return Redirect(jmessage.GetValue("payUrl").ToString());
        }
        public ActionResult Thanhtoanmomo(Result result)
        {
            DATHANG ddh = new DATHANG();
            NGUOIDUNG kh = (NGUOIDUNG)Session["UserId"];
            List<GIOHANG> gh = LayGioHang();
            ddh.maND = kh.maND;
            ddh.hotenKH = kh.hoten;
            ddh.email = kh.email;
            ddh.ngayDat = DateTime.Now;
            ddh.maPT = 2;
            ddh.diaChi = Session["diachinhanhang"].ToString(); ;         
            decimal tongtien = 0;
            foreach (var item in gh)
            {
                decimal thanhtien = item.isoluong * (decimal)item.dgia;
                tongtien += thanhtien;
            }
            ddh.tongGiaTri = tongtien;
            db.DATHANGs.Add(ddh);
            db.SaveChanges();
            //Thêm chi tiết đơn hàng
            foreach (var item in gh)
            {
                CHITIETDONHANG ctDH = new CHITIETDONHANG();
                decimal thanhtien = item.isoluong * (decimal)item.dgia;
                ctDH.maDH = ddh.maDH;
                ctDH.maSP = item.imaSP;
                ctDH.soluong = item.isoluong;
                ctDH.gia = (decimal)item.dgia;
                ctDH.tonggiatri = (decimal)thanhtien;
                db.CHITIETDONHANGs.Add(ctDH);
            }
            db.SaveChanges();
            return View();
        }

        [HttpPost]
        public void Luuthanhtoan()
        {
            //cập nhật dữ liệu vào db
        }
        //Send Email
        //public ActionResult SendMail()
        //{
        //    return View();
        //}
        //[HttpPost]
        //public ActionResult SendEmail(string email)
        //{
        //    string subject = "Xác Nhận Thanh Toán";
        //    string body = "Cảm ơn quý khách đã đặt hàng";
        //    WebMail.Send(email, subject, body, null, null, null, true, null, null, null, null, null);
        //    ViewBag.msg = "Email đã được gửi";
        //    return View();
        //}

        //private void SendOrderConfirmationEmail()
        //{
        //    // Khởi tạo đối tượng SmtpClient và cấu hình từ web.config
        //    SmtpClient smtpClient = new SmtpClient();
        //    smtpClient.UseDefaultCredentials = false;
        //    smtpClient.Host = System.Configuration.ConfigurationManager.AppSettings["smtpHost"];
        //    smtpClient.Port = int.Parse(System.Configuration.ConfigurationManager.AppSettings["smtpPort"]);
        //    smtpClient.EnableSsl = bool.Parse(System.Configuration.ConfigurationManager.AppSettings["smtpEnableSsl"]);
        //    smtpClient.Credentials = new NetworkCredential(
        //        System.Configuration.ConfigurationManager.AppSettings["smtpUserName"],
        //        System.Configuration.ConfigurationManager.AppSettings["smtpPassword"]
        //    );

        //    // Tạo đối tượng MailMessage để tạo email
        //    MailMessage mailMessage = new MailMessage();
        //    mailMessage.From = new MailAddress(System.Configuration.ConfigurationManager.AppSettings["smtpUserName"]);
        //    mailMessage.To.Add("nguyenbagialam@gmail.com");
        //    mailMessage.Subject = "Subject of the Email";
        //    mailMessage.Body = "Body of the Email";

        //    // Gửi email
        //    smtpClient.Send(mailMessage);
        //}
    }
}