﻿using System;
using System.Collections.Generic;
using System.Data.SqlTypes;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Xml.Serialization;

namespace EndesaEntity.cnmc.V21_2019_12_17
{
    [XmlRoot(ElementName = "MensajeBajaSuspension", Namespace = "http://localhost/elegibilidad", IsNullable = true)]
    public class TipoMensajeB101
    {
        public Cabecera Cabecera { get; set; }
        public EndesaEntity.cnmc.V30_2022_21_01.BajaSuspension BajaSuspension { get; set; }

        public TipoMensajeB101()
        {
            Cabecera = new Cabecera();
            BajaSuspension = new V30_2022_21_01.BajaSuspension();
        }
    }
}
