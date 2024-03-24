namespace CuaHangThoiTrang.Models
{
    using System;
    using System.Collections.Generic;
    using System.ComponentModel.DataAnnotations;
    using System.ComponentModel.DataAnnotations.Schema;
    using System.Data.Entity.Spatial;

    [Table("CHITIETDONHANG")]
    public partial class CHITIETDONHANG
    {
        [Key]
        [Column(Order = 0)]
        [DatabaseGenerated(DatabaseGeneratedOption.None)]
        public int maDH { get; set; }

        [Key]
        [Column(Order = 1)]
        [DatabaseGenerated(DatabaseGeneratedOption.None)]
        public int maSP { get; set; }

        public int soluong { get; set; }

        [Column(TypeName = "money")]
        public decimal gia { get; set; }

        [Column(TypeName = "money")]
        public decimal? tonggiatri { get; set; }

        public virtual DATHANG DATHANG { get; set; }

        public virtual SANPHAM SANPHAM { get; set; }
    }
}
