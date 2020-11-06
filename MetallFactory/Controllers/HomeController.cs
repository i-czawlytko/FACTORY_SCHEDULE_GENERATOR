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
            scheduleGenerator = _scheduleGenerator;
        }

        public IActionResult Index()
        {
            //repository.Load();
            return View(scheduleGenerator.Generate());
        }

        public IActionResult Schedule()
        {
            //repository.Load();
            return View(scheduleGenerator.Generate());
        }
        public IActionResult Privacy()
        {
            return View();
        }

        [ResponseCache(Duration = 0, Location = ResponseCacheLocation.None, NoStore = true)]
        public IActionResult Error()
        {
            return View(new ErrorViewModel { RequestId = Activity.Current?.Id ?? HttpContext.TraceIdentifier });
        }
    }
}
