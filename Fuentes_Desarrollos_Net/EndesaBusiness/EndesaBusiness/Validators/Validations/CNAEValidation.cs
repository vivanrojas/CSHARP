using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using EndesaBusiness.cnmc;
using EndesaEntity.extrasistemas;

namespace EndesaBusiness.Validators.Validations
{
    class CNAEValidation :ICellValidator
    {
        public CellValidationResult Validate(string cellValue, Global global, CNMC diccionary, int row, int column)
        {
            // CNAE - Obligatorio X(4)

            if (cellValue != null && cellValue.Length >= 4)
            {
                global.cnae = cellValue.Substring(0, 4);
                return CellValidationResult.WithSuccess();
            }
            else
            {
                return CellValidationResult.WithError(
                    $"Hoja: {global.hoja} Fila: {row} --> CUPS: {global.cups22}  CNAE: el campo CNAE es nulo o no tiene una longitud válida X(4)."
                );
            }
        }




    }
}
