using System;
using System.Collections.Generic;
using System.Data;
using System.Data.Entity;
using System.Linq;
using System.Net;
using System.Web;
using System.Web.Mvc;
using MaskShoppingCart.Models;

namespace MaskShoppingCart.Controllers
{
    public class MaskOrdersController : Controller
    {
        private MaskOrderContext db = new MaskOrderContext();

        // GET: MaskOrders
        public ActionResult Index()
        {
            return View(db.MaskOrders.ToList());
        }

        // GET: MaskOrders/Details/5
        public ActionResult Details(int? id)
        {
            if (id == null)
            {
                return new HttpStatusCodeResult(HttpStatusCode.BadRequest);
            }
            MaskOrder maskOrder = db.MaskOrders.Find(id);
            if (maskOrder == null)
            {
                return HttpNotFound();
            }
            return View(maskOrder);
        }

        // GET: MaskOrders/Create
        public ActionResult Create()
        {
            return View();
        }

        // POST: MaskOrders/Create
        [HttpPost]
        [ValidateAntiForgeryToken]
        public ActionResult Create(MaskOrder maskOrder)
        {
            if (ModelState.IsValid)
            {
                db.MaskOrders.Add(maskOrder);
                db.SaveChanges();
                return RedirectToAction("Index");
            }

            return View(maskOrder);
        }

        // GET: MaskOrders/Edit/5
        public ActionResult Edit(int? id)
        {
            if (id == null)
            {
                return new HttpStatusCodeResult(HttpStatusCode.BadRequest);
            }
            MaskOrder maskOrder = db.MaskOrders.Find(id);
            if (maskOrder == null)
            {
                return HttpNotFound();
            }
            return View(maskOrder);
        }

        // POST: MaskOrders/Edit/5
        [HttpPost]
        [ValidateAntiForgeryToken]
        public ActionResult Edit(MaskOrder maskOrder)
        {
            if (ModelState.IsValid)
            {
                db.Entry(maskOrder).State = EntityState.Modified;
                db.SaveChanges();
                return RedirectToAction("Index");
            }
            return View(maskOrder);
        }

        // GET: MaskOrders/Delete/5
        public ActionResult Delete(int? id)
        {
            if (id == null)
            {
                return new HttpStatusCodeResult(HttpStatusCode.BadRequest);
            }
            MaskOrder maskOrder = db.MaskOrders.Find(id);
            if (maskOrder == null)
            {
                return HttpNotFound();
            }
            return View(maskOrder);
        }

        // POST: MaskOrders/Delete/5
        [HttpPost, ActionName("Delete")]
        [ValidateAntiForgeryToken]
        public ActionResult DeleteConfirmed(int id)
        {
            MaskOrder maskOrder = db.MaskOrders.Find(id);
            db.MaskOrders.Remove(maskOrder);
            db.SaveChanges();
            return RedirectToAction("Index");
        }

        protected override void Dispose(bool disposing)
        {
            if (disposing)
            {
                db.Dispose();
            }
            base.Dispose(disposing);
        }
    }
}
