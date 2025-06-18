using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using EndesaBusiness.cnmc;
using EndesaEntity.extrasistemas;

namespace EndesaBusiness.Validators.Validations
{
    class TiposolicitudadministrativaValidation : ICellValidator
    {
        public CellValidationResult Validate(string cellValue, Global global, CNMC dictionary, int row, int column)
        {
            // Tipo_solicitud_administrativa - Obligatorio X(1) sólo si tipo_modificacion = 'A' o 'S'
            if (global.tipo_modificacion == "A" || global.tipo_modificacion == "S")
            {
                if (cellValue != null && cellValue.Length >= 1)
                {
                    global.tipo_solicitud_administrativa = cellValue.Substring(0, 1);
                    return CellValidationResult.WithSuccess();
                }
                else
                {
                    return CellValidationResult.WithError(
                        "Hoja: " + global.hoja +
                        " Fila: " + row +
                        " --> Tipo_solicitud_administrativa: el campo es requerido para modificaciones de tipo 'A' o 'S' y es nulo o tiene una longitud inválida X(1)."
                    );
                }
            }
            else
            {
                // No es obligatorio para otros tipos de modificación
                return CellValidationResult.WithSuccess();
            }
        }

    }
}
