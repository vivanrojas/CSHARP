using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using EndesaBusiness.cnmc;
using EndesaEntity.extrasistemas;

namespace EndesaBusiness.Validators.Validations
{
    class PotInstaladaGenValidation : ICellValidator

    {
        public CellValidationResult Validate(string cellValue, Global global, CNMC diccionary, int row, int column)
        {
            if (!string.IsNullOrWhiteSpace(cellValue))
            {
                string trimmedValue = cellValue.Trim();

                if (trimmedValue.Length <= 14 && long.TryParse(trimmedValue, out long potInstaladaGenValor))
                {
                    // Asignar el valor directamente
                    global.potinstaladagen = (int)potInstaladaGenValor;
                    return CellValidationResult.WithSuccess();
                }
            }

            return CellValidationResult.WithError(
                $"Hoja: {global.hoja} --> Fila: {row} potInstaladaGen: el campo potInstaladaGen es nulo, no es numérico o excede las 14 posiciones."
            );
        }
    }
}
