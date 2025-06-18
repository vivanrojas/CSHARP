using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using EndesaBusiness.cnmc;
using EndesaEntity.extrasistemas;

namespace EndesaBusiness.Validators.Validations
{
    class ColectivoValidation : ICellValidator
    {
        public CellValidationResult Validate(string cellValue, Global global,CNMC diccionary , int row, int column)
        {
            if (cellValue != null && (Convert.ToString(cellValue) == "S" || Convert.ToString(cellValue) == "N"))

            {
                global.colectivo = Convert.ToString(cellValue);
                return CellValidationResult.WithSuccess();
            }
            else
            {
                // Mensaje de error si es nulo o de longitud insuficiente
                return CellValidationResult.WithError(
                    "Hoja: " + global.hoja +
                    " Fila: " + row +
                    " --> colectivo: el campo colectivo es nulo o no tiene un valor válido"
                );

            }
        }
    }
}
