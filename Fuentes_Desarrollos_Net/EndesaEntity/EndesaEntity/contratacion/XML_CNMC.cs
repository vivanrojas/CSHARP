using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace EndesaEntity.contratacion
{
    public class XML_CNMC
    {
        public string codigoREEEmpresaEmisora { get; set; }
        public string codigoREEEmpresaDestino { get; set; }
        public string codigoDelProceso { get; set; }
        public string codigoDePaso { get; set; }
        public string codigoDeSolicitud { get; set; }
        public string secuencialDeSolicitud { get; set; }
        public DateTime fechaSolicitud { get; set; }
        public string cups { get; set; }
        public int cnae { get; set; }
        public double potenciaExtension { get; set; }
        public double potenciaDeAcceso { get; set; }
        public double potenciaInstAT { get; set; }
        public string indicativoDeInterrumpibilidad { get; set; }
        public string pais { get; set; }
        public string provincia { get; set; }
        public string desc_provincia { get; set; }
        public string municipio { get; set; }
        public string poblacion { get; set; }
        public string descripcionPoblacion { get; set; }
        public string tipoVia { get; set; }
        public string codPostal { get; set; }
        public string calle { get; set; }
        public string numeroFinca { get; set; }
        public string aclaradorFinca { get; set; }
        public string tipoIdentificador { get; set; }
        public string identificador { get; set; }
        public string tipoPersona { get; set; }
        public string razonSocial { get; set; }
        public string prefijoPais { get; set; }
        public string numero { get; set; }
        public string telefono { get; set; }
        public string correoElectronico { get; set; }
        public string indicadorTipoDireccion { get; set; }
        public string indicativoDeDireccionExterna { get; set; }
        public string linea1DeLaDireccionExterna { get; set; }
        public string linea2DeLaDireccionExterna { get; set; }
        public string linea3DeLaDireccionExterna { get; set; }
        public string linea4DeLaDireccionExterna { get; set; }
        public string idioma { get; set; }
        public DateTime fechaActivacion { get; set; } // FechaActivacion
        public string codContrato { get; set; }
        public string tipoAutoconsumo { get; set; }
        public string tipoContratoATR { get; set; }
        public string tarifaATR { get; set; }
        public string periodicidadFacturacion { get; set; }
        public string tipodeTelegestion { get; set; }
        public double[] potenciaPeriodo { get; set; }
        public string marcaMedidaConPerdidas { get; set; }
        public int tensionDelSuministro { get; set; }
        public bool encontrado_registro { get; set; }
        public string fichero { get; set; }
        public string paisCliente { get; set; }
        public string provinciaCliente { get; set; }
        public string municipioCliente { get; set; }
        public string poblacionCliente { get; set; }
        public string descripcionPoblacionCliente { get; set; }
        public string codPostalCliente { get; set; }
        public string calleCliente { get; set; }
        public string numeroFincaCliente { get; set; }
        public string pisoCliente { get; set; }
        public string tipoViaCliente { get; set; }
        public string indActivacion { get; set; }
        public DateTime fechaPrevistaAccion { get; set; }
        public string tipoActivacionPrevista { get; set; }
        public DateTime fechaActivacionPrevista { get; set; }


        // Datos relativos a PS_AT
        public bool existe_en_psat { get; set; }


        public XML_CNMC()
        {
            encontrado_registro = false;
            potenciaPeriodo = new double[7];
            existe_en_psat = false;
        }
    }
}
