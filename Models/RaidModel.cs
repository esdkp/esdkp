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
        public DateTime Date { get; set; }
        public string Name { get; set; }
        public List<string> Raiders { get; set; }

        /// <summary>
        /// List of Tuples representing awarded loots in Loot, Raider, Price order.
        /// </summary>
        public List<LootModel> Loots { get; set; }

        public RaidModel()
        {
            Date = new DateTime();
            Name = "";
            Raiders = new List<string>();
            Loots = new List<LootModel>();
        }
    }
}
