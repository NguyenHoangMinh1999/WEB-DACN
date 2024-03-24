using System;
using System.Collections.Generic;
using System.ComponentModel.DataAnnotations;
using System.Linq;
using System.Web;

namespace CuaHangThoiTrang.Models
{
    public class LoginModel
    {
        [Key]
        [Display(Name="Email")]
        [Required(ErrorMessage = "Please enter your email")]
        public string userMail { get; set; }
        [Display(Name = "Mật khẩu")]
        [Required(ErrorMessage = "Please enter your password")]
        public string password { get; set; }
        public bool RemmemberMe { get; set; }
    }
}