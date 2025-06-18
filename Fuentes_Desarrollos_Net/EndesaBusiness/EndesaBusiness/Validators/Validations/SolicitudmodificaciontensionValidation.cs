using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using EndesaBusiness.cnmc;
using EndesaEntity.extrasistemas;

namespace EndesaBusiness.Validators.Validations
{
    class SolicitudmodificaciontensionValidation :ICellValidator
    {
        public CellValidationResult Validate(string cellValue, Global global, CNMC diccionary, int row, int column)
        {
            // solicitud_modificacion_tension – Obligatorio X(1)

            if (!string.IsNullOrWhiteSpace(cellValue) && cellValue.Length >= 1)
            {
                global.solicitud_modificacion_tension = cellValue.Substring(0, 1);
                return CellValidationResult.WithSuccess();
            }
            else
            {
                return CellValidationResult.WithError(
                    $"Hoja: {global.hoja} Fila: {row} --> solicitud_modificacion_tension: el campo está vacío o no tiene al menos 1 carácter."
                );
            }
        }

    }
}
