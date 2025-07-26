using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace EndesaBusiness.facturacion
{
    public class CuadroMandoFacturacion
    {

        EndesaBusiness.gas.PuntosActivosGas puntosActivosGas;
        public CuadroMandoFacturacion()
        {
            DateTime fd = new DateTime(2020, 08, 01);
            DateTime fh = new DateTime(2020, 08, 31);
            puntosActivosGas = new gas.PuntosActivosGas(fd, fh, "POR");
        }


    }
}
