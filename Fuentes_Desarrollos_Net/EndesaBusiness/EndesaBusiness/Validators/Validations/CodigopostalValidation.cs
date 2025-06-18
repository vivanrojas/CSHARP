using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using EndesaBusiness.cnmc;
using EndesaEntity.extrasistemas;

namespace EndesaBusiness.Validators.Validations
{
    class CodigopostalValidation : ICellValidator
    {
        public CellValidationResult Validate(string cellValue, Global global, CNMC diccionary, int row, int column)
        {
            // Código Postal - Obligatorio, debe ser numérico y tener exactamente 5 dígitos

            if (!string.IsNullOrWhiteSpace(cellValue))
            {
                string valorCP = cellValue.Trim();

                if (int.TryParse(valorCP, out int codigoPostal))
                {
                    string codigoPostalStr = codigoPostal.ToString().PadLeft(5, '0');

                    if (codigoPostalStr.Length == 5)
                    {
                        global.codigo_postal = codigoPostalStr;
                        return CellValidationResult.WithSuccess();
                    }
                    else
                    {
                        return CellValidationResult.WithError(
                            $"Hoja: {global.hoja} Fila: {row} --> Código Postal: debe contener exactamente 5 dígitos numéricos."
                        );
                    }
                }
                else
                {
                    return CellValidationResult.WithError(
                        $"Hoja: {global.hoja} Fila: {row} --> Código Postal: el valor ingresado no es numérico."
                    );
                }
            }
            else
            {
                return CellValidationResult.WithError(
                    $"Hoja: {global.hoja} Fila: {row} --> Código Postal: el campo está vacío."
                );
            }
        }

    }
}
