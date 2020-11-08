using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.Threading.Tasks;
using Microsoft.AspNetCore.Mvc;
using Microsoft.Extensions.Logging;
using MetallFactory.Models;
using MetallFactory.ViewModels;

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
                scheduleGenerator = _scheduleGenerator;
        }

        public IActionResult Index()
        {
            try
            {               
                repository.Load();
                return View(new MainViewModel
                {
                    Machines = repository.Machines,
                    Errors = repository.CheckOut()
                });
            }
            catch (ExcelDataException e)
            {
                TempData["message"] = "Неверные данные в xlsx-файлах // " + e.Message;
                return View("Error");
            }
            catch (Exception e)
            {
                TempData["message"] = e.Message;
                return View("Error");
            }

        }

        public IActionResult ExportToExcel()
        {
            scheduleGenerator.ExportToXlxs();
            return RedirectToAction("Index");
        }

        public IActionResult Schedule()
        {
            try
            {
                scheduleGenerator.Generate();
                return View(scheduleGenerator.GetSchedule());
            }
            catch (Exception e)
            {
                TempData["message"] = e.Message;
                return View("Error");
            }

        }
        public JsonResult GetChart()
        {
            repository.Load();
            var groups = from p in repository.Parties
                         join mat in repository.Materials on p.MaterialId equals mat.Id
                         group repository.Parties by mat.Name into g
                         select new { Name = g.Key, Count = g.Count() };
            var Names = groups.Select(x => x.Name);
            var Quantity = groups.Select(x => x.Count);


            return Json(new
            {
                names = Names,
                quantity = Quantity
            });
        }

        public IActionResult Competitors()
        {
            repository.Load();
            
            return View(repository.StructuredTimes);
        }
    }
}
