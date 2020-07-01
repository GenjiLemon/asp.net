using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Web.Mvc;

namespace GuestBook.Models
{
    [Bind(Exclude = "Name")]
    public class Guestbook
    {
        
        public int Id { get; set; }
        public string Name { get; set; }
        public string Email { get; set; }
        public string Content { get; set; }
    }
}