using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Web.Mvc;

namespace EFDemo.Controllers
{
    public class MVCController : Controller
    {
        // GET: MVC
        public ActionResult Index()
        {
            return View();
        }

        // GET: MVC/Details/5
        public ActionResult Details(int id)
        {
            return View();
        }

        // GET: MVC/Create
        public ActionResult Create()
        {
            return View();
        }

        // POST: MVC/Create
        [HttpPost]
        public ActionResult Create(FormCollection collection)
        {
            try
            {
                // TODO: Add insert logic here

                return RedirectToAction("Index");
            }
            catch
            {
                return View();
            }
        }

        // GET: MVC/Edit/5
        public ActionResult Edit(int id)
        {
            return View();
        }

        // POST: MVC/Edit/5
        [HttpPost]
        public ActionResult Edit(int id, FormCollection collection)
        {
            try
            {
                // TODO: Add update logic here

                return RedirectToAction("Index");
            }
            catch
            {
                return View();
            }
        }

        // GET: MVC/Delete/5
        public ActionResult Delete(int id)
        {
            return View();
        }

        // POST: MVC/Delete/5
        [HttpPost]
        public ActionResult Delete(int id, FormCollection collection)
        {
            try
            {
                // TODO: Add delete logic here

                return RedirectToAction("Index");
            }
            catch
            {
                return View();
            }
        }
    }
}
