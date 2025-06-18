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
    class Cups22Validation : ICellValidator
    {
        public CellValidationResult Validate(string cellValue, Global global, CNMC diccionary, int row, int column)
        {
            if (cellValue != null && cellValue.Length >= 22)
            {
                // Tomamos sólo los primeros 22 caracteres
                global.cups22 = cellValue.Substring(0, 22);
                return CellValidationResult.WithSuccess();
            }
            else
            {
                // Mensaje de error si es nulo o de longitud insuficiente
                return CellValidationResult.WithError(
                    "Hoja: " + global.hoja +
                    " Fila: " + row +
                    " --> CUPS: el campo CUPS es nulo o no tiene una longitud válida X(22)."
                );
            }


        }
    }
}

