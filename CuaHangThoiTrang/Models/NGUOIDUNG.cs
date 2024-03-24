using System;
using System.Collections.Generic;
using System.ComponentModel.DataAnnotations;
using System.ComponentModel.DataAnnotations.Schema;
using System.Data.Entity.Spatial;

namespace CuaHangThoiTrang.Models
{

    [Table("NGUOIDUNG")]
    public partial class NGUOIDUNG
    {
        [System.Diagnostics.CodeAnalysis.SuppressMessage("Microsoft.Usage", "CA2214:DoNotCallOverridableMethodsInConstructors")]
        public NGUOIDUNG()
        {
            DATHANGs = new HashSet<DATHANG>();
        }

        [Key]
        public int maND { get; set; }

        public int maRole { get; set; }

        [Required]
        [StringLength(50)]
        public string hoten { get; set; }

        public bool gioitinh { get; set; }

        [Required]
        [StringLength(50)]
        public string email { get; set; }

        [Required]
        [StringLength(11)]
        public string sodienthoai { get; set; }

        [StringLength(100)]
        public string diachi { get; set; }

        public DateTime? ngaysinh { get; set; }

        [Required]
        [StringLength(50)]
        public string password { get; set; }

        [Required]
        [StringLength(50)]
        public string confirmpassword { get; set; }

        [System.Diagnostics.CodeAnalysis.SuppressMessage("Microsoft.Usage", "CA2227:CollectionPropertiesShouldBeReadOnly")]
        public virtual ICollection<DATHANG> DATHANGs { get; set; }

        public virtual ROLE ROLE { get; set; }
    }
}