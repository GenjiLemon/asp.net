using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.ComponentModel;
using System.Web.Mvc;
using System.ComponentModel.DataAnnotations;

namespace MaskShoppingCart.Models
{

    public class MaskOrder
    {
        [ScaffoldColumn(false)]
        public int Id { get; set; }
        [Required(ErrorMessage = "请输入要购买的数量")]
        [Range(1, 1000,
        ErrorMessage = "Price must be between 1 and 1000")]
        public int Quantity { get; set; }
        [Required(ErrorMessage = "请输入您的地址")]
        [StringLength(1024)]
        public string Address { get; set; }
        [Required(ErrorMessage = "请输入您的姓名")]
        public string Buyer { get; set; }
        [Required(ErrorMessage = "请输入您的手机号")]
        public string PhoneNum { get; set; }
       
 
    }
}