using EndesaBusiness.cnmc;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using EndesaEntity.extrasistemas;

namespace EndesaBusiness.Validators.Validations
{
    class TipocontratoValidation   : ICellValidator
    {
        public CellValidationResult Validate(
               string cellValue,
               Global global,
               CNMC diccionary,
               int row,
               int column)
        {
            // tipo_contrato_atr – Obligatorio X(2), valor en cnmc.dic_contrato_atr

            if (!string.IsNullOrWhiteSpace(cellValue) && cellValue.Length >= 2)
            {
                var tipoContrato = cellValue.Substring(0, 2).ToUpper();
                if (diccionary.dic_contrato_atr.ContainsValue(tipoContrato))
                {
                    global.tipo_contrato_atr = tipoContrato;
                    return CellValidationResult.WithSuccess();
                }
                else
                {
                    return CellValidationResult.WithError(
                        $"Hoja: {global.hoja} Fila: {row} --> tipo_contrato_atr: el campo tipo_contrato_atr no contiene un valor válido."
                    );
                }
            }
            else
            {
                return CellValidationResult.WithError(
                    $"Hoja: {global.hoja} Fila: {row} --> tipo_contrato_atr: el campo tipo_contrato_atr es nulo o no tiene una longitud válida X(2)."
                );
            }
        }



    }
}
