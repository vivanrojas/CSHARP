using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using EndesaBusiness.cnmc;
using EndesaEntity.extrasistemas;

namespace EndesaBusiness.Validators.Validations
{
    class TipoInstalacionValidation : ICellValidator
    {
        public CellValidationResult Validate(string cellValue, Global global, CNMC diccionary, int row, int column)
        {
            string tipoAutoconsumo = global.tipo_autoconsumo;

            if (!string.IsNullOrWhiteSpace(cellValue) && cellValue.Length >= 2)
            {
                string tipoInstalacionValor = cellValue.Substring(0, 2).ToUpper();

                if (tipoAutoconsumo == "11")
                {
                    // Solo se permite "01" o "02" si tipo_autoconsumo es "11"
                    if (tipoInstalacionValor == "01" || tipoInstalacionValor == "02")
                    {
                        global.tipoinstalacion = tipoInstalacionValor;
                        return CellValidationResult.WithSuccess();
                    }
                    else
                    {
                        return CellValidationResult.WithError(
                            $"Hoja: {global.hoja} --> Fila: {row} TipoInstalacion: cuando TipoAutoconsumo es '11', solo puede ser '01' o '02'. Se recibió '{tipoInstalacionValor}'."
                        );
                    }
                }
                else
                {
                    global.tipoinstalacion = tipoInstalacionValor;
                    return CellValidationResult.WithSuccess();
                }
            }
            else
            {
                if (tipoAutoconsumo == "11")
                {
                    return CellValidationResult.WithError(
                        $"Hoja: {global.hoja} --> Fila: {row} TipoInstalacion: cuando TipoAutoconsumo es '11', el campo TipoInstalacion es obligatorio ('01' o '02')."
                                             );
                }
                else
                {
                    global.tipoinstalacion = null;
                    return CellValidationResult.WithSuccess(); // Campo vacío permitido si no es "11"
                }
            }
        }

    }
}

