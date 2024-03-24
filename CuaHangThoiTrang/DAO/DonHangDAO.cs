using CuaHangThoiTrang.Models;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;

namespace MLVshop.DAO
{
    public class DonHangDAO
    {
        CHThoiTrangDbContext context;
        public DonHangDAO()
        {
            context = new CHThoiTrangDbContext();
        }

        public IQueryable<DATHANG> ListDH()
        {
            var res = (from dh in context.DATHANGs select dh);
            return res;
        }
    }
}