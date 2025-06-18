using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using EndesaBusiness.cnmc;
using EndesaEntity.extrasistemas;

namespace EndesaBusiness.Validators.Validations
{
    class PrimerApellidoValidation : ICellValidator
    {
        public CellValidationResult Validate(string cellValue, Global global, CNMC diccionary, int row, int column)
        {
            // primer_apellido - Obligatorio X(50)

            if (!string.IsNullOrWhiteSpace(cellValue))
            {
                global.primer_apellido = cellValue.Substring(0, Math.Min(cellValue.Length, 50));
                return CellValidationResult.WithSuccess();
            }
            else
            {
                return CellValidationResult.WithError(
                    $"Hoja: {global.hoja} Fila: {row} --> CUPS: {global.cups22} primer_apellido: el campo es obligatorio y no puede estar vacío."
                );
            }
        }

    }
}
