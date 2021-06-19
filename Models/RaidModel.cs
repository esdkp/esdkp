using ES_DKP_Utils.Models;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ES_DKP_Utils
{
    public class RaidModel
    {
        public DateTime Date { get; set; } = new DateTime();
        public string Name { get; set; }
        public double AttendanceMultiplier { get; set; } = 1.0;
        public List<string> Raiders { get; set; } = new List<string>();

        public List<LootModel> Loots { get; set; } = new List<LootModel>();
    }
}
