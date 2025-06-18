using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using EndesaBusiness.cnmc;
using EndesaEntity.extrasistemas;

namespace EndesaBusiness.Validators.Validations
{
    class PaisValidation : ICellValidator
    {
        public CellValidationResult Validate(string cellValue, Global global, CNMC diccionary, int row, int column)
        {
            // país - Obligatorio

            if (!string.IsNullOrWhiteSpace(cellValue))
            {
                global.pais = cellValue.Trim();
                return CellValidationResult.WithSuccess();
            }
            else
            {
                return CellValidationResult.WithError(
                    $"Hoja: {global.hoja} Fila: {row} --> CUPS: {global.cups22} El campo 'país' es obligatorio y no está informado."
                );
            }
        }

    }
}
