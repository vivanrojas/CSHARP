using EndesaBusiness.cnmc;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using EndesaEntity.extrasistemas;

namespace EndesaBusiness.Validators.Validations
{
    class TarifaValidation : ICellValidator 
    {
        public CellValidationResult Validate(string cellValue, Global global, CNMC diccionary, int row, int column)
        {
            // tarifa – Obligatorio X(3), valor en cnmc.dic_tarifa_atr
            if (cellValue != null)
            {
                var cellValueT = cellValue.ToString().Trim();
                if (cellValueT.Length >= 3)
                {
                    var tarifaStr = cellValueT.Substring(0, 3).ToUpper();
                    // Atención: la clave es la descripción (ContainsValue)
                    if (diccionary.dic_tarifa_atr.ContainsValue(tarifaStr))
                    {
                        global.tarifa = tarifaStr;
                        return CellValidationResult.WithSuccess();
                    }
                    else
                    {
                        return CellValidationResult.WithError(
                            $"Hoja: {global.hoja} Fila: {row} --> tarifa atr: el campo tarifa atr no contiene un valor válido."
                        );
                    }
                }
            }

            // Nulo o longitud insuficiente
            return CellValidationResult.WithError(
                $"Hoja: {global.hoja} Fila: {row} --> tarifa atr: el campo tarifa atr es nulo o no tiene una longitud válida X(3)."
            );
        }

    }
}
