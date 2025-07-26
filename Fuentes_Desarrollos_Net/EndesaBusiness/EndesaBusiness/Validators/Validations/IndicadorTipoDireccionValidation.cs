using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using EndesaBusiness.cnmc;
using EndesaEntity.extrasistemas;

namespace EndesaBusiness.Validators.Validations
{
    class IndicadorTipoDireccionValidation : ICellValidator
    {

        public CellValidationResult Validate(
                             string cellValue, Global global, CNMC diccionary ,int row, int column)
        {
            // IndicadorTipoDireccion – Obligatorio, longitud mínima 1 carácter
            if (!string.IsNullOrWhiteSpace(cellValue))
            {
                global.indicador_tipo_direccion = cellValue.Substring(0, 1).ToUpperInvariant();
                return CellValidationResult.WithSuccess();
            }
            else
            {
                return CellValidationResult.WithError(
                    $"Hoja: {global.hoja} --> CUPS: {global.cups22} El campo 'IndicadorTipoDireccion' no tiene un valor válido."
                );
            }
        }
    }

}