using EndesaBusiness.cnmc;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using EndesaEntity.extrasistemas;

namespace EndesaBusiness.Validators.Validations
{
    class ContratacionIncondicionalPSValidation :ICellValidator
    {
        public CellValidationResult Validate(
               string cellValue, Global global, CNMC dictionary, int row,  int column)
        {
            // ContratacionIncondicionalPS – Obligatorio X(1), valor en cnmc.dic_indicativo_sino

            if (!string.IsNullOrWhiteSpace(cellValue) && cellValue.Length >= 1)
            {
                var valor = cellValue.Substring(0, 1).ToUpper();
                if (dictionary.dic_indicativo_sino.ContainsValue(valor))
                {
                    global.contratacion_incondicional_ps = valor;
                    return CellValidationResult.WithSuccess();
                }
                else
                {
                    return CellValidationResult.WithError(
                        $"Hoja: {global.hoja} Fila: {row} --> ContratacionIncondicionalPS: el campo ContratacionIncondicionalPS no contiene un valor válido."
                    );
                }
            }
            else
            {
                return CellValidationResult.WithError(
                    $"Hoja: {global.hoja} Fila: {row} --> ContratacionIncondicionalPS: el campo ContratacionIncondicionalPS es nulo o no tiene una longitud válida X(1)."
                );
            }
        }


    }
}
