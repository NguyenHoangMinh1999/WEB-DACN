namespace CuaHangThoiTrang.Models
{
    using System;
    using System.Collections.Generic;
    using System.ComponentModel.DataAnnotations;
    using System.ComponentModel.DataAnnotations.Schema;
    using System.Data.Entity.Spatial;

    [Table("SIZE")]
    public partial class SIZE
    {
        [System.Diagnostics.CodeAnalysis.SuppressMessage("Microsoft.Usage", "CA2214:DoNotCallOverridableMethodsInConstructors")]
        public SIZE()
        {
            CHITIETSIZEs = new HashSet<CHITIETSIZE>();
        }

        [Key]
        public int maSize { get; set; }


        [Required]
        [StringLength(6)]
        public string tenS { get; set; }

        public int? mota { get; set; }

        [System.Diagnostics.CodeAnalysis.SuppressMessage("Microsoft.Usage", "CA2227:CollectionPropertiesShouldBeReadOnly")]
        public virtual ICollection<CHITIETSIZE> CHITIETSIZEs { get; set; }
    }
}
