using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using EndesaBusiness.cnmc;
using EndesaEntity.extrasistemas;

namespace EndesaBusiness.Validators.Validations
{
    class NombreViaValidation : ICellValidator
    {
        public CellValidationResult Validate(string cellValue, Global global, CNMC diccionary, int row, int column)
        {
            // Nombre de Vía - Obligatorio, máximo 30 caracteres

            if (!string.IsNullOrWhiteSpace(cellValue))
            {
                string nombreVia = cellValue.Trim();

                if (nombreVia.Length > 30)
                {
                    global.nombre_via = nombreVia.Substring(0, 30); // Truncar a 30 caracteres
                    return CellValidationResult.WithError(
                        $"Hoja: {global.hoja} Fila: {row} --> Nombre de Vía: el texto excede los 30 caracteres y fue truncado."
                    );
                }
                else
                {
                    global.nombre_via = nombreVia;
                    return CellValidationResult.WithSuccess();
                }
            }
            else
            {
                return CellValidationResult.WithError(
                    $"Hoja: {global.hoja} Fila: {row} --> Nombre de Vía: el campo está vacío o es nulo."
                );
            }
        }

    }
}
