using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;

namespace CuaHangThoiTrang.Models
{
    public class GIOHANG
    {
        CHThoiTrangDbContext db = new CHThoiTrangDbContext();
        public int imaSP { get; set; }
        public string stenSP { get; set; }
        public string shinh { get; set; }
        public double dgia { get; set; }
        public int isoluong { get; set; }
        public double dthanhtien 
        { 
            get { return isoluong * dgia; }  
        }
        //tạo giỏ hàng
        public GIOHANG(int id)
        {
            imaSP = id;
            SANPHAM sp = db.SANPHAMs.Single(n => n.maSP == imaSP);
            stenSP = sp.tenSP;
            shinh = sp.hinh;
            dgia = sp.gia;
            isoluong = 1;
        }
    }
}