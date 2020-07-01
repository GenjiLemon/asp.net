using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Web.Mvc;
using MvcMusicStore.Models;
namespace MvcMusicStore.Controllers
{
    public class StoreController : Controller
    {
        MusicStoreEntities storeDB = new MusicStoreEntities();
        // GET: Store
        public ActionResult  Index()
        {
            var genres = storeDB.Genres.ToList();
            
            return this.View(genres);
        }
        public ActionResult Browse(string genre)
        {
            var genremodel = storeDB.Genres.Include("Albums").Single(p => p.Name == genre);

            return View(genremodel);
        }
        public ActionResult Details(int id)
        {
            var album = storeDB.Albums.Find(id);
            return View(album);
        }

    }
}