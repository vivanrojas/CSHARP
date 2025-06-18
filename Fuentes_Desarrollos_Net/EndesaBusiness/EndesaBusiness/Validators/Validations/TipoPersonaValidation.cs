using Aspose.Cells;
using Google.Protobuf.WellKnownTypes;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using EndesaEntity.extrasistemas;
using EndesaBusiness.cnmc;

namespace EndesaBusiness.Validators.Validations
{
    class TipoPersonaValidation :ICellValidator
    {

        public CellValidationResult Validate(string cellValue, Global global, CNMC diccionary, int row, int colum)
        {
            if (!string.IsNullOrEmpty(cellValue))
            {
                global.tipo_persona = cellValue.Substring(0, 1);

                //if (global.tipo_persona == "J")
                //{
                //    global.razon_social = razonSocial?.Substring(0, Math.Min(razonSocial.Length, 50));
                //}
                //else
                //{
                //    global.nombre_de_pila = nombre?.Substring(0, Math.Min(nombre.Length, 50));
                //    global.primer_apellido = apellido?.Substring(0, Math.Min(apellido.Length, 50));
                //}

                return CellValidationResult.WithSuccess();
            }
            else
            {
                return CellValidationResult.WithError("Hoja: " + global.hoja + " Fila: " + row + " --> TipoPersona: el campo está vacío o nulo.");
            }
        }



    }
}
