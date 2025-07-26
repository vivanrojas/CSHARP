using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace EndesaEntity.medida
{
    public class P1_Dic
    {


        public Dictionary<DateTime, List<P1>> dic { get; set; }

        public P1_Dic()
        {
            dic = new Dictionary<DateTime, List<P1>>();
        }
    }



}
