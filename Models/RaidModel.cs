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

        /// <summary>
        /// Amount of attendance this raid is worth
        /// </summary>
        public double Attendance { get; set; } = 1.0;
        
        /// <summary>
        /// List of raiders who have attended the raid
        /// </summary>
        public List<string> Raiders { get; set; } = new List<string>();
        
        /// <summary>
        /// List of awarded loots for a given raid
        /// </summary>
        public List<LootModel> Loots { get; set; } = new List<LootModel>();
    }
}
