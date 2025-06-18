using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using EndesaBusiness.cnmc;
using EndesaEntity.extrasistemas;

namespace EndesaBusiness.Validators.Validations
{
    class TipoCupsValidation : ICellValidator
    {
        public CellValidationResult Validate(string cellValue, Global global, CNMC diccionary, int row, int column)
        {
            if (!string.IsNullOrWhiteSpace(cellValue) && cellValue.Length >= 2)
            {
                // Tomar las primeras dos letras en mayúsculas
                string celdatipocups = cellValue.Substring(0, 2).ToUpper();
                global.tipocups = celdatipocups;

                return CellValidationResult.WithSuccess();
            }
            else
            {
                return CellValidationResult.WithError(
                    $"Hoja: {global.hoja} --> Fila: {row} tipocups: el campo tipocups es nulo o no tiene una longitud válida (mínimo 2 caracteres)."
                );
            }
        }

    }
}
