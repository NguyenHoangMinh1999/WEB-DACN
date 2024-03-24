namespace CuaHangThoiTrang.Models
{
    using System;
    using System.Collections.Generic;
    using System.ComponentModel.DataAnnotations;
    using System.ComponentModel.DataAnnotations.Schema;
    using System.Data.Entity.Spatial;

    [Table("DATHANG")]
    public partial class DATHANG
    {
        [System.Diagnostics.CodeAnalysis.SuppressMessage("Microsoft.Usage", "CA2214:DoNotCallOverridableMethodsInConstructors")]
        public DATHANG()
        {
            CHITIETDONHANGs = new HashSet<CHITIETDONHANG>();
        }

        [Key]
        public int maDH { get; set; }

        public int maPT { get; set; }

        public int maND { get; set; }

        [Required]
        [StringLength(30)]
        public string hotenKH { get; set; }

        [Required]
        [StringLength(30)]
        public string email { get; set; }

        public DateTime ngayDat { get; set; }

        [Column(TypeName = "money")]
        public decimal tongGiaTri { get; set; }

        public int trangThai { get; set; }

        [Required]
        [StringLength(200)]
        public string diaChi { get; set; }

        [StringLength(200)]
        public string Ghichu { get; set; }

        [System.Diagnostics.CodeAnalysis.SuppressMessage("Microsoft.Usage", "CA2227:CollectionPropertiesShouldBeReadOnly")]
        public virtual ICollection<CHITIETDONHANG> CHITIETDONHANGs { get; set; }

        public virtual NGUOIDUNG NGUOIDUNG { get; set; }

        public virtual PHUONGTHUCTHANHTOAN PHUONGTHUCTHANHTOAN { get; set; }
    }
}
