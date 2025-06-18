using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using EndesaBusiness.cnmc;
using EndesaEntity.extrasistemas;

namespace EndesaBusiness.Validators.Validations
{
    class NIdentificadorValidation : ICellValidator
    {
        public CellValidationResult Validate(string cellValue, Global global, CNMC diccionary, int row, int column)
        {
            // n_identificador – Obligatorio, máximo 14 caracteres
            if (!string.IsNullOrWhiteSpace(cellValue))
            {
                var valor = cellValue.Trim();
                if (valor.Length <= 14)
                {
                    global.n_identificador = valor.ToUpperInvariant();
                    return CellValidationResult.WithSuccess();
                }
                else
                {
                    return CellValidationResult.WithError(
                        $"Hoja: {global.hoja} Fila: {row} --> n_identificador: " +
                        $"el valor excede la longitud máxima permitida (14 caracteres)."
                    );
                }
            }
            else
            {
                return CellValidationResult.WithError(
                    $"Hoja: {global.hoja} Fila: {row} --> n_identificador: el campo es nulo o está vacío."
                );
            }
        }


    }
}
