using MetallFactory.Models;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;

namespace MetallFactory.ViewModels
{
    public class MainViewModel
    {
        public IEnumerable<Machine> Machines { get; set; }
        public IEnumerable<string> Errors { get; set; }
    }
}
