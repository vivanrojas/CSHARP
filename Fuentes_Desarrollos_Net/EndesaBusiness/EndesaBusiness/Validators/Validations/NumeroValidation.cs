using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using EndesaBusiness.cnmc;
using EndesaEntity.extrasistemas;

namespace EndesaBusiness.Validators.Validations
{
    class NumeroValidation : ICellValidator
    {

        public CellValidationResult Validate(string cellValue, Global global, CNMC diccionary, int row, int column)
        {
            const int MaxLength = 5;

            if (cellValue == null)
            {
                return CellValidationResult.WithError(
                $"Hoja: {global.hoja} Fila: {row} --> Número: el campo está nulo."
                );
            }

            string numero = cellValue.Trim();

            if (numero.Length == 0)
            {
                return CellValidationResult.WithError(
                $"Hoja: {global.hoja} Fila: {row} --> Número: el campo está vacío."
                );
            }

            if (numero.Length > MaxLength)
            {
                global.numero = numero.Substring(0, MaxLength);
                return CellValidationResult.WithError(
                $"Hoja: {global.hoja} Fila: {row} --> Número: el texto excede los {MaxLength} caracteres y fue truncado."
                );
            }

            global.numero = numero;
            return CellValidationResult.WithSuccess();
        }

    }
}
