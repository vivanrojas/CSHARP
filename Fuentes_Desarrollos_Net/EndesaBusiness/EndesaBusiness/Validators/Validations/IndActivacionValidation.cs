using EndesaBusiness.cnmc;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using EndesaEntity.extrasistemas;

namespace EndesaBusiness.Validators.Validations
{
    class IndActivacionValidation : ICellValidator
    {
        public CellValidationResult Validate( string cellValue, Global global, CNMC diccionary,
          int row,
          int column)
        {
            // IndActivacion - Obligatorio X(1) y con valor en cnmc.dic_indicativo_activacion

            if (cellValue != null && cellValue.Length >= 1)
            {
                // Solo el primer carácter en mayúscula
                var valor = cellValue.Substring(0, 1).ToUpper();

                if (diccionary.dic_indicativo_activacion.ContainsValue(valor))
                {
                    global.ind_activacion = valor;
                    return CellValidationResult.WithSuccess();
                }
                else
                {
                    return CellValidationResult.WithError(
                        $"Hoja: {global.hoja} Fila: {row} --> CUPS: {global.cups22} IndActivacion: el campo IndActivacion no contiene un valor válido."
                    );
                }
            }
            else
            {
                return CellValidationResult.WithError(
                    $"Hoja: {global.hoja} Fila: {row} --> CUPS: {global.cups22} IndActivacion: el campo IndActivacion es nulo o no tiene una longitud válida X(1)."
                );
            }
        }



    }
}
