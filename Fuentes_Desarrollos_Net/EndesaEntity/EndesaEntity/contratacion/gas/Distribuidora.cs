﻿using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace EndesaEntity.contratacion.gas
{
    public class Distribuidora
    {
        public string distribuidora { get; set; }
        public DateTime fecha_desde { get; set; }
        public DateTime fecha_hasta { get; set; }
        public string mail { get; set; }
        public string tramitacion { get; set; }

    }
}
