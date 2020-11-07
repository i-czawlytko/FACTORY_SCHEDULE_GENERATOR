using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.Threading.Tasks;
using Microsoft.AspNetCore.Mvc;
using Microsoft.Extensions.Logging;
using MetallFactory.Models;

namespace MetallFactory.Controllers
{
    public class HomeController : Controller
    {
        private readonly ILogger<HomeController> _logger;
        private IRepository repository;
        private ScheduleGenerator scheduleGenerator;

        public HomeController(ILogger<HomeController> logger, IRepository repo, ScheduleGenerator _scheduleGenerator)
        {
            _logger = logger;
            repository = repo;
            repo.Load();
            scheduleGenerator = _scheduleGenerator;
        }

        public IActionResult Index()
        {
            return View(repository.Machines);
        }

        public IActionResult ExportToExcel()
        {
            scheduleGenerator.ExportToXlxs();
            return RedirectToAction("Index");
        }

        public IActionResult Schedule()
        {
            scheduleGenerator.Generate();
            return View(scheduleGenerator.GetSchedule());
        }
        public JsonResult GetChart()
        {

            var groups = from p in repository.Parties
                      join mat in repository.Materials on p.MaterialId equals mat.Id
                      group repository.Parties by mat.Name into g
                      select new {Name = g.Key, Count = g.Count() };
            var Names = groups.Select(x => x.Name);
            var Quantity = groups.Select(x => x.Count);


            return Json(new {
                names = Names,
                quantity = Quantity});
        }

        [ResponseCache(Duration = 0, Location = ResponseCacheLocation.None, NoStore = true)]
        public IActionResult Error()
        {
            return View(new ErrorViewModel { RequestId = Activity.Current?.Id ?? HttpContext.TraceIdentifier });
        }
    }
}
