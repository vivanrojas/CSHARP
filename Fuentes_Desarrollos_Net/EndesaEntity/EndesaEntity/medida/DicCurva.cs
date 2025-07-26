using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace EndesaEntity.medida
{
    public class DicCurva
    {
        public Dictionary<DateTime, EndesaEntity.medida.CurvaDeCarga> dic { get; set; }
        public DicCurva()
        {
            dic = new Dictionary<DateTime, EndesaEntity.medida.CurvaDeCarga>();
        }
    }
}
