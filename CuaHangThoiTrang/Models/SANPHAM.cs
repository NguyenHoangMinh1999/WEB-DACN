namespace CuaHangThoiTrang.Models
{
    using System;
    using System.Collections.Generic;
    using System.ComponentModel.DataAnnotations;
    using System.ComponentModel.DataAnnotations.Schema;
    using System.Data.Entity.Spatial;

    [Table("SANPHAM")]
    public partial class SANPHAM
    {
        [System.Diagnostics.CodeAnalysis.SuppressMessage("Microsoft.Usage", "CA2214:DoNotCallOverridableMethodsInConstructors")]
        public SANPHAM()
        {
            CHITIETDONHANGs = new HashSet<CHITIETDONHANG>();
            CHITIETSIZEs = new HashSet<CHITIETSIZE>();
            DANHGIAs = new HashSet<DANHGIA>();
            MAUs = new HashSet<MAU>();
        }

        [Key]
        public int maSP { get; set; }

        public int maTH { get; set; }

        public int maDM { get; set; }

        [Required]
        [StringLength(50)]
        public string tenSP { get; set; }

        public int gia { get; set; }

        [Required]
        [StringLength(30)]
        public string hinh { get; set; }

        [StringLength(30)]
        public string hinh1 { get; set; }

        [StringLength(30)]
        public string hinh2 { get; set; }

        public int soLuong { get; set; }

        [StringLength(1000)]
        public string thongTin { get; set; }

        [System.Diagnostics.CodeAnalysis.SuppressMessage("Microsoft.Usage", "CA2227:CollectionPropertiesShouldBeReadOnly")]
        public virtual ICollection<CHITIETDONHANG> CHITIETDONHANGs { get; set; }

        [System.Diagnostics.CodeAnalysis.SuppressMessage("Microsoft.Usage", "CA2227:CollectionPropertiesShouldBeReadOnly")]
        public virtual ICollection<CHITIETSIZE> CHITIETSIZEs { get; set; }

        [System.Diagnostics.CodeAnalysis.SuppressMessage("Microsoft.Usage", "CA2227:CollectionPropertiesShouldBeReadOnly")]
        public virtual ICollection<DANHGIA> DANHGIAs { get; set; }

        public virtual DANHMUC DANHMUC { get; set; }

        [System.Diagnostics.CodeAnalysis.SuppressMessage("Microsoft.Usage", "CA2227:CollectionPropertiesShouldBeReadOnly")]
        public virtual ICollection<MAU> MAUs { get; set; }

        public virtual THUONGHIEU THUONGHIEU { get; set; }
    }
}
