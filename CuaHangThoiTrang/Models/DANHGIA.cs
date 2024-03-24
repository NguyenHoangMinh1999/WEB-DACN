namespace CuaHangThoiTrang.Models
{
    using System;
    using System.Collections.Generic;
    using System.ComponentModel.DataAnnotations;
    using System.ComponentModel.DataAnnotations.Schema;
    using System.Data.Entity.Spatial;

    [Table("DANHGIA")]
    public partial class DANHGIA
    {
        [Key]
        public int maDG { get; set; }

        public int maSP { get; set; }

        [StringLength(200)]
        public string noidungDG { get; set; }

        public int sosao { get; set; }

        public virtual SANPHAM SANPHAM { get; set; }
    }
}
