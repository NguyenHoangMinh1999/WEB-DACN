namespace CuaHangThoiTrang.Models
{
    using System;
    using System.Collections.Generic;
    using System.ComponentModel.DataAnnotations;
    using System.ComponentModel.DataAnnotations.Schema;
    using System.Data.Entity.Spatial;

    [Table("MAU")]
    public partial class MAU
    {
        [Key]
        public int maMau { get; set; }

        public int maSP { get; set; }

        [Required]
        [StringLength(10)]
        public string tenMau { get; set; }

        public virtual SANPHAM SANPHAM { get; set; }
    }
}
