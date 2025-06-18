using EndesaBusiness.cnmc;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using EndesaEntity.extrasistemas;

namespace EndesaBusiness.Validators.Validations
{
    class SsaaValidation : ICellValidator
    {
        public CellValidationResult Validate(string cellValue, Global global, CNMC diccionary, int row, int column)
        {
            string valor = cellValue?.Trim().ToUpper();

            if (valor == "S" || valor == "N")
            {
                global.ssaa = valor;
                return CellValidationResult.WithSuccess();
            }
            else if (global.tipocups == "01")
            {
                return CellValidationResult.WithError(
                    $"Hoja: {global.hoja} --> Fila: {row} ssaa: el campo ssaa es obligatorio y debe contener 'S' o 'N' cuando TipoCups es '01'."
                );
            }
            else
            {
                return CellValidationResult.WithError(
                    $"Hoja: {global.hoja} --> Fila: {row} ssaa: el campo ssaa es nulo o no tiene un valor válido."
                );
            }
        }

    }
}
