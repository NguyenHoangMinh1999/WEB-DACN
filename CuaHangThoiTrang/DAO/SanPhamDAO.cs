using CuaHangThoiTrang.Models;
using PagedList;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;

namespace CuaHangThoiTrang.DAO
{
    public class SanPhamDAO
    {
        CHThoiTrangDbContext context;

        public SanPhamDAO()
        {
            context = new CHThoiTrangDbContext();
        }

        public IQueryable<SANPHAM> ListSP()
        {
            var res = (from sp in context.SANPHAMs select sp);
            return res;
        }

        public List<SANPHAM> listspTimkiem(string tensp)
        {
            string search = "select * from SanPham where ten like '%" + tensp + "%'";
            var rs = context.SANPHAMs.SqlQuery(search).ToList();
            return rs;
        }

        
    }
}