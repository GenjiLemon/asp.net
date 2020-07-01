using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Web.Mvc;
using MVC.Models;
namespace MVC.Controllers
{
    public class HomeController : Controller
    {
        // GET: Home
        public void Index()
        {

          
        }
     
        [HttpGet]
        public ActionResult Form()
        {
            return View();
        }
        [HttpPost]
        public void GetForm(Student stu)
        {

            Response.Write(stu.Email + "<br>" + stu.Name);

        }
        public ActionResult Show()
        {
            Student stu = new Student()
            {
                Email = "fay.com",
                Name = "fay"
            };
           
            return View(stu);

        }
        [HttpPost]
        [ValidateAntiForgeryToken]
        public ActionResult Form(Student stu)
        {
            if (ModelState.IsValid)
            {
                ModelState.AddModelError("", "对");
                return View();
            }
            else
            {
                ModelState.AddModelError("", "错啦");
                return View();
            }
        }
        
       
    }
}