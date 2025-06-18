using EndesaBusiness.cnmc;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using EndesaEntity.extrasistemas;

namespace EndesaBusiness.Validators.Validations
{
    class TensionsolicitadaValidation :ICellValidator
    {
        public CellValidationResult Validate(string cellValue, Global global, CNMC diccionary, int row, int column)
        {
            // Tension_solicitada – Obligatorio X(2) solo si solicitud_modificacion_tension = 'S'
            if (global.solicitud_modificacion_tension == "S")
            {
                if (!string.IsNullOrWhiteSpace(cellValue) && cellValue.Length >= 2)
                {
                    var valorTension = cellValue.Substring(0, 2);
                    if (diccionary.dic_tensiones.ContainsValue(valorTension))
                    {
                        global.tension_solicitada = valorTension;
                        return CellValidationResult.WithSuccess();
                    }
                    else
                    {
                        return CellValidationResult.WithError(
                            $"Hoja: {global.hoja} Fila: {row} --> Tension_solicitada: el valor '{valorTension}' no se encuentra en el diccionario de tensiones válido."
                        );
                    }
                }
                else
                {
                    return CellValidationResult.WithError(
                        $"Hoja: {global.hoja} Fila: {row} --> Tension_solicitada: campo requerido porque solicitud_modificacion_tension es 'S', pero es nulo o no tiene una longitud válida X(2)."
                    );
                }
            }
            // No aplica si no se solicitó modificación de tensión
            return CellValidationResult.WithSuccess();
        }
    }
}
