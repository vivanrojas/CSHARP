using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using EndesaBusiness.cnmc;
using EndesaEntity.extrasistemas;

namespace EndesaBusiness.Validators.Validations
{
    class TipodeviaValidation : ICellValidator
    {
        public CellValidationResult Validate(string cellValue, Global global, CNMC diccionary, int row, int column)
        {
            // Tipo de vía - Obligatorio, 2 caracteres

            if (!string.IsNullOrWhiteSpace(cellValue) && cellValue.Length >= 2)
            {
                string valortipoVia = cellValue.Substring(0, 2).ToUpper();

                if (diccionary.dic_via.ContainsValue(valortipoVia))
                {
                    global.tipo_via = valortipoVia;
                    return CellValidationResult.WithSuccess();
                }
                else
                {
                    return CellValidationResult.WithError(
                        $"Hoja: {global.hoja} Fila: {row} --> tipo via: el campo tipo via no contiene un valor válido."
                    );
                }
            }
            else
            {
                return CellValidationResult.WithError(
                    $"Hoja: {global.hoja} Fila: {row} --> tipo via: el campo tipo via es nulo o no tiene una longitud válida (2 caracteres)."
                );
            }
        }

    }
}
