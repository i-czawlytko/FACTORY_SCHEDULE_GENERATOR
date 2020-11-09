using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;

namespace MetallFactory.Models
{
    public interface IRepository
    {
        public List<Material> Materials { get; set; }
        public List<Machine> Machines { get; set; }
        public List<Party> Parties { get; set; }
        public List<TimeInfo> Times { get; set; }
        public List<TIStructured> StructuredTimes { get; set; }
        public Dictionary<int, Dictionary<int, int>> Competitors { get; set; }
        public void Load();
        public List<string> CheckOut();

        public List<List<TIStructured>> AllCombinations { get; set; }
    }
}
