using System;
using System.ComponentModel.DataAnnotations.Schema;
using System.Data.Entity;
using System.Linq;

namespace CuaHangThoiTrang.Models
{
    public partial class CHThoiTrangDbContext : DbContext
    {
        public CHThoiTrangDbContext()
            : base("name=CHThoiTrangDbContext")
        {
        }

        public virtual DbSet<CHITIETDONHANG> CHITIETDONHANGs { get; set; }
        public virtual DbSet<CHITIETSIZE> CHITIETSIZEs { get; set; }
        public virtual DbSet<DANHGIA> DANHGIAs { get; set; }
        public virtual DbSet<DANHMUC> DANHMUCs { get; set; }
        public virtual DbSet<DATHANG> DATHANGs { get; set; }
        public virtual DbSet<MAU> MAUs { get; set; }
        public virtual DbSet<NGUOIDUNG> NGUOIDUNGs { get; set; }
        public virtual DbSet<PHUONGTHUCTHANHTOAN> PHUONGTHUCTHANHTOANs { get; set; }
        public virtual DbSet<ROLE> ROLES { get; set; }
        public virtual DbSet<SANPHAM> SANPHAMs { get; set; }
        public virtual DbSet<SIZE> SIZEs { get; set; }
        public virtual DbSet<sysdiagram> sysdiagrams { get; set; }
        public virtual DbSet<THUONGHIEU> THUONGHIEUx { get; set; }

        protected override void OnModelCreating(DbModelBuilder modelBuilder)
        {
            modelBuilder.Entity<CHITIETDONHANG>()
                .Property(e => e.gia)
                .HasPrecision(19, 4);

            modelBuilder.Entity<CHITIETDONHANG>()
                .Property(e => e.tonggiatri)
                .HasPrecision(19, 4);

            modelBuilder.Entity<DANHMUC>()
                .HasMany(e => e.SANPHAMs)
                .WithRequired(e => e.DANHMUC)
                .WillCascadeOnDelete(false);

            modelBuilder.Entity<DATHANG>()
                .Property(e => e.tongGiaTri)
                .HasPrecision(19, 4);

            modelBuilder.Entity<DATHANG>()
                .HasMany(e => e.CHITIETDONHANGs)
                .WithRequired(e => e.DATHANG)
                .WillCascadeOnDelete(false);

            modelBuilder.Entity<NGUOIDUNG>()
                .Property(e => e.sodienthoai)
                .IsFixedLength()
                .IsUnicode(false);

            modelBuilder.Entity<NGUOIDUNG>()
                .Property(e => e.password)
                .IsUnicode(false);

            modelBuilder.Entity<NGUOIDUNG>()
                .Property(e => e.confirmpassword)
                .IsUnicode(false);

            modelBuilder.Entity<NGUOIDUNG>()
                .HasMany(e => e.DATHANGs)
                .WithRequired(e => e.NGUOIDUNG)
                .WillCascadeOnDelete(false);

            modelBuilder.Entity<PHUONGTHUCTHANHTOAN>()
                .HasMany(e => e.DATHANGs)
                .WithRequired(e => e.PHUONGTHUCTHANHTOAN)
                .WillCascadeOnDelete(false);

            modelBuilder.Entity<ROLE>()
                .Property(e => e.tenRole)
                .IsUnicode(false);

            modelBuilder.Entity<ROLE>()
                .HasMany(e => e.NGUOIDUNGs)
                .WithRequired(e => e.ROLE)
                .HasForeignKey(e => e.maRole)
                .WillCascadeOnDelete(false);;

            modelBuilder.Entity<SANPHAM>()
                .HasMany(e => e.CHITIETDONHANGs)
                .WithRequired(e => e.SANPHAM)
                .WillCascadeOnDelete(false);

            modelBuilder.Entity<SANPHAM>()
                .HasMany(e => e.CHITIETSIZEs)
                .WithRequired(e => e.SANPHAM)
                .WillCascadeOnDelete(false);

            modelBuilder.Entity<SANPHAM>()
                .HasMany(e => e.DANHGIAs)
                .WithRequired(e => e.SANPHAM)
                .WillCascadeOnDelete(false);

            modelBuilder.Entity<SANPHAM>()
                .HasMany(e => e.MAUs)
                .WithRequired(e => e.SANPHAM)
                .WillCascadeOnDelete(false);

            modelBuilder.Entity<SIZE>()
                .Property(e => e.tenS)
                .IsFixedLength()
                .IsUnicode(false);

            modelBuilder.Entity<SIZE>()
                .HasMany(e => e.CHITIETSIZEs)
                .WithRequired(e => e.SIZE)
                .WillCascadeOnDelete(false);

            modelBuilder.Entity<THUONGHIEU>()
                .HasMany(e => e.SANPHAMs)
                .WithRequired(e => e.THUONGHIEU)
                .WillCascadeOnDelete(false);
        }
    }
}
