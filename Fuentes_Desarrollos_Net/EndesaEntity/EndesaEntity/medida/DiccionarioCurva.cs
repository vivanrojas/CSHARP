using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace EndesaEntity.medida
{
    public class DiccionarioCurva
    {
        public Dictionary<DateTime,Curva> dic { get; set; }

        public DiccionarioCurva()
        {
            dic = new Dictionary<DateTime, Curva>();
        }
    }
}
