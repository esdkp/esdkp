using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ES_DKP_Utils.Models
{
    public class LootModel
    {
        /// <summary>
        /// Name of Loot Awarded
        /// </summary>
        public string Name { get; set; }

        /// <summary>
        /// Name of Raider loot was awarded to
        /// </summary>
        public string Raider { get; set; }

        /// <summary>
        /// Amount of Points loot is worth
        /// </summary>
        public double Points { get; set; }

        public LootModel(string name, string raider, double points)
        {
            this.Name = name;
            this.Raider = raider;
            this.Points = points;
        }
    }
}
