using System;
using System.Collections.Generic;
using System.Diagnostics.Contracts;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace EndesaEntity.contratacion.xxi
{
    public class XML_Datos
    {
        public string codigoREEEmpresaEmisora { get; set; }
        public string codigoREEEmpresaDestino { get; set; }
        public string codigoDelProceso { get; set; }
        public string codigoDePaso { get; set; }
        public string codigoDeSolicitud { get; set; }
        public string secuencialDeSolicitud { get; set; }
        public DateTime fechaSolicitud { get; set; }
        public string cups { get; set; }
        public string motivoTraspaso { get; set; }
        public DateTime fechaPrevistaAccion { get; set; }
        public DateTime fechaFinalizacion { get; set; }
        public string cnae { get; set; }
        public string indEsencial { get; set; }        
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
        public string numeroFinca {get; set;}
        public string tipoAclaradorFinca { get; set; }
        public string piso { get; set; }
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
        public int vasTrafo { get; set; }
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

        public string modoControlPotencia { get; set; }
        public string contacto_PersonaDeContacto { get; set; }
        public string contacto_PrefijoPais { get; set; }
        public string contacto_Numero { get; set; }
        public string contacto_CorreoElectronico { get; set; }
        //public string cliente_TipoIdentificador { get; set; }
        //public string cliente_Identificador { get; set; }
        //public string cliente_TipoPersona { get; set; }
        //public string cliente_NombreDePila { get; set; }
        //public string cliente_PrimerApellido { get; set; }
        //public string cliente_SegundoApellido { get; set; }
        //public string cliente_PrefijoPais { get; set; }
        //public string cliente_Numero { get; set; }
        //public string cliente_CorreoElectronico { get; set; }
        //public string cliente_IndicadorTipoDireccion { get; set; }
        //public string cliente_Pais { get; set; }
        //public string cliente_Provincia { get; set; }
        //public string cliente_Municipio { get; set; }
        //public string cliente_Poblacion { get; set; }
        //public string cliente_CodPostal { get; set; }
        //public string cliente_TipoVia { get; set; }
        //public string cliente_Calle { get; set; }
        //public string cliente_NumeroFinca { get; set; }
        public string direccionPS_Pais { get; set; }
        public string direccionPS_Provincia { get; set; }
        public string direccionPS_Municipio { get; set; }
        public string direccionPS_Poblacion { get; set; }
        public string direccionPS_CodPostal { get; set; }
        public string direccionPS_Calle { get; set; }
        public string direccionPS_NumeroFinca { get; set; }
        public string direccionPS_AclaradorFinca { get; set; }
        public string direccionPS_TipoVia { get; set; }
        public string motivo { get; set; }
        public bool rechazar { get; set; }
        public bool existe { get; set; }





        public List<XML_PuntoDeMedida> lista_PM { get; set; }

        // Datos relativos a PS_AT
        public bool existe_en_psat { get; set; }

        


        public XML_Datos()
        {
            codigoREEEmpresaEmisora = "";
            codigoREEEmpresaDestino = "";
            codigoDelProceso = "";
            codigoDePaso = "";
            codigoDeSolicitud = "";
            secuencialDeSolicitud = "";
  
             cups = "";
             motivoTraspaso = "";
         
         
           indEsencial = "";
             
          indicativoDeInterrumpibilidad = "N";
             pais = "";
             provincia = "";
             desc_provincia = "";
             municipio = "";
             poblacion = "";
             descripcionPoblacion = "";
             tipoVia = "";
             codPostal = "";
             calle = "";
             numeroFinca = "";
             tipoAclaradorFinca = "";
             piso = "";
             aclaradorFinca = "";
             tipoIdentificador = "";
             identificador = "";
             tipoPersona = "";
             razonSocial = "";
             prefijoPais = "";
             numero = "";
             telefono = "";
             correoElectronico = "";
             indicadorTipoDireccion = "F";
             indicativoDeDireccionExterna = "S";
             linea1DeLaDireccionExterna = "";
             linea2DeLaDireccionExterna = "";
             linea3DeLaDireccionExterna = "";
             linea4DeLaDireccionExterna = "";
             idioma = "ES";
            
             codContrato = "";
             tipoAutoconsumo = "";
             tipoContratoATR = "";
             tarifaATR = "";
             periodicidadFacturacion = "";
             tipodeTelegestion = "";
            
          marcaMedidaConPerdidas = "";
            
            
          fichero = "";
             paisCliente = "";
             provinciaCliente = "";
             municipioCliente = "";
             poblacionCliente = "";
             descripcionPoblacionCliente = "";
             codPostalCliente = "";
             calleCliente = "";
             numeroFincaCliente = "";
             pisoCliente = "";
             tipoViaCliente = "";

             modoControlPotencia = "";
             contacto_PersonaDeContacto = "";
             contacto_PrefijoPais = "";
             contacto_Numero = "";
             contacto_CorreoElectronico = "";
             //cliente_TipoIdentificador = "";
             //cliente_Identificador = "";
             //cliente_TipoPersona = "";
             //cliente_NombreDePila = "";
             //cliente_PrimerApellido = "";
             //cliente_SegundoApellido = "";
             //cliente_PrefijoPais = "";
             //cliente_Numero = "";
             //cliente_CorreoElectronico = "";
             //cliente_IndicadorTipoDireccion = "";
             //cliente_Pais = "";
             //cliente_Provincia = "";
             //cliente_Municipio = "";
             //cliente_Poblacion = "";
             //cliente_CodPostal = "";
             //cliente_TipoVia = "";
             //cliente_Calle = "";
             //cliente_NumeroFinca = "";
            direccionPS_Pais = "";
              direccionPS_Provincia = "";
             direccionPS_Municipio = "";
             direccionPS_Poblacion = "";
             direccionPS_CodPostal = "";
             direccionPS_Calle = "";
             direccionPS_NumeroFinca = "";

            lista_PM = new List<XML_PuntoDeMedida>();
            encontrado_registro = false;
            potenciaPeriodo = new double[7];
            existe_en_psat = false;
            existe = false;
        }
                
    }
}
