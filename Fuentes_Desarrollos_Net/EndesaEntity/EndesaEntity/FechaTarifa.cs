using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace EndesaEntity
{
    public class FechaTarifa
    {
        public DateTime f { get; set; }
        public int[] pt { get; set; }

        public FechaTarifa()
        {
            pt = new int[26];
        }
    }
}
