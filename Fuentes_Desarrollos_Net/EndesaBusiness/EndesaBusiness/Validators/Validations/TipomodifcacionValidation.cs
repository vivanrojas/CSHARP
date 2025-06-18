using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using EndesaBusiness.cnmc;
using EndesaEntity.extrasistemas;

namespace EndesaBusiness.Validators.Validations
{
    class TipomodifcacionValidation : ICellValidator
    {
        public CellValidationResult Validate(string cellValue, Global global, CNMC dictionary, int row, int column)
        {
            // Tipo_modificacion - Obligatorio X(1)

            if (cellValue != null && cellValue.Length >= 1)
            {
                // Solo el primer carácter
                global.tipo_modificacion = cellValue.Substring(0, 1);
                return CellValidationResult.WithSuccess();
            }
            else
            {
                // Error si es nulo o longitud insuficiente
                return CellValidationResult.WithError(
                    "Hoja: " + global.hoja +
                    " Fila: " + row +
                    " --> Tipo_modificacion: el campo tipo de modificación es nulo o no tiene una longitud válida X(1)."
                );
            }
        }

    }
}
