using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Aspose.Cells;
using EndesaBusiness.cnmc;
using EndesaEntity.extrasistemas;

namespace EndesaBusiness.Validators.Validations
{
    class CauValidation : ICellValidator
    {
        public CellValidationResult Validate(string cellValue, Global global, CNMC diccionary, int row, int column)
        {
            if (global.ssaa == "S")
            {
                if (!string.IsNullOrWhiteSpace(cellValue) && (cellValue == "S" || cellValue == "N"))
                {
                    global.unicocontrato = cellValue;
                    return CellValidationResult.WithSuccess();
                }
                else
                {
                    return CellValidationResult.WithError(
                        $"Hoja: {global.hoja} --> Fila: {row} unicocontrato: el campo es obligatorio y debe contener 'S' o 'N' cuando SSAA es 'S'."
                    );
                }
            }
            else
            {
                global.unicocontrato = cellValue; // se asigna aunque esté vacío o nulo
                return CellValidationResult.WithSuccess();
            }
        }
        
    }
}
