using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using EndesaBusiness.cnmc;
using EndesaEntity.extrasistemas;

namespace EndesaBusiness.Validators.Validations
{
    class RazonSocialValidation : ICellValidator
    {
        public CellValidationResult Validate(string cellValue, Global global, CNMC diccionary, int row, int column)
       
        {
            // modo_control_potencia – Obligatorio X(1)

            if (global.tipo_persona == "J")
            {
              global.razon_social = Convert.ToString(cellValue);
                      // global.modo_control_potencia = cellValue.Substring(0, 1).ToUpper();
                return CellValidationResult.WithSuccess();
            }
            else
            {
                return CellValidationResult.WithError(
                    $"Hoja: {global.hoja} Fila: {row} --> CUPS: {global.cups22} modo_control_potencia: el campo está vacío o no tiene una longitud válida X(1)."
                );
            }
        }



    }
}
