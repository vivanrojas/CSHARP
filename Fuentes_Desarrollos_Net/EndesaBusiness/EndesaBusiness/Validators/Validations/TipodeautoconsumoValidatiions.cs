using EndesaBusiness.cnmc;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using EndesaEntity.extrasistemas;

namespace EndesaBusiness.Validators.Validations
{
    class TipodeautoconsumoValidatiions : ICellValidator
    {
        public CellValidationResult Validate(
                           string cellValue, Global global, CNMC dictionary,int row, int column)
        {
            // tipo_autoconsumo – Obligatorio X(2), valor en cnmc.dic_autoconsumo

            if (!string.IsNullOrWhiteSpace(cellValue) && cellValue.Length >= 2)
            {
                var tipoAutoconsumo = cellValue.Substring(0, 2).ToUpper();
                if (dictionary.dic_autoconsumo.ContainsValue(tipoAutoconsumo))
                {
                    global.tipo_autoconsumo = tipoAutoconsumo;
                    return CellValidationResult.WithSuccess();
                }
                else
                {
                    return CellValidationResult.WithError(
                        $"Hoja: {global.hoja} Fila: {row} --> tipo_autoconsumo: el campo no contiene un valor válido."
                    );
                }
            }
            else
            {
                return CellValidationResult.WithError(
                    $"Hoja: {global.hoja} Fila: {row} --> tipo_autoconsumo: el campo es nulo o no tiene una longitud válida (mínimo 2 caracteres)."
                );
            }
        }




    }
}
