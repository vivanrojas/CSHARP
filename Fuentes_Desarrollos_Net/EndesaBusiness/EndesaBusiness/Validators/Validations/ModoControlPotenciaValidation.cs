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
    class ModoControlPotenciaValidation : ICellValidator
    {
        public CellValidationResult Validate(string cellValue, Global global, CNMC diccionary,  int row,
                                   int column)
        {
            // modo_control_potencia – Obligatorio X(1)
            if (cellValue != null)
            {
                var cellValuem = cellValue.ToString().Trim();
                if (cellValuem.Length >= 1)
                {
                    global.modo_control_potencia = cellValuem.Substring(0, 1).ToUpper();
                    return CellValidationResult.WithSuccess();
                }
            }

            return CellValidationResult.WithError(
                $"Hoja: {global.hoja} Fila: {row} --> CUPS: {global.cups22} modo_control_potencia: el campo está vacío o no tiene la longitud mínima X(1)."
            );
        }


    }
}
