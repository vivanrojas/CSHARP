using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using EndesaBusiness.cnmc;
using EndesaEntity.extrasistemas;

namespace EndesaBusiness.Validators.Validations
{
    class TipoSubseccionValidation : ICellValidator
    {
        public CellValidationResult Validate(string cellValue, Global global, CNMC diccionary, int row, int column)
        {
            if (!string.IsNullOrWhiteSpace(cellValue) && cellValue.Length >= 2)
            {
                // Se obtienen las primeras dos letras en mayúsculas
                string tiposubseccionValor = cellValue.Substring(0, 2).ToUpper();
                global.tipo_subseccion = tiposubseccionValor;

                return CellValidationResult.WithSuccess();
            }
            else
            {
                return CellValidationResult.WithError(
                    $"Hoja: {global.hoja} --> Fila: {row} tiposubseccion: el campo tiposubseccion es nulo o no tiene una longitud válida (mínimo 2 caracteres)."
                );
            }
        }
    }
}
