using CuaHangThoiTrang.Models;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;

namespace CuaHangThoiTrang.DAO
{
    public class ThuongHieuDAO
    {
        CHThoiTrangDbContext context;

        public ThuongHieuDAO()
        {
            context = new CHThoiTrangDbContext();
        }

        public List<THUONGHIEU> ListThuongHieu()
        {
            return context.THUONGHIEUx.ToList();
        }
    }
}