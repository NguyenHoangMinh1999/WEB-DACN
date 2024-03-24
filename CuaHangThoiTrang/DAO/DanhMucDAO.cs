using CuaHangThoiTrang.Models;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;

namespace CuaHangThoiTrang.DAO
{
    public class DanhMucDAO
    {
        CHThoiTrangDbContext context;
        
        public DanhMucDAO()
        {
            context = new CHThoiTrangDbContext();
        }

        public List<DANHMUC> ListDanhMuc()
        {
            return context.DANHMUCs.ToList();
        }
    }
}