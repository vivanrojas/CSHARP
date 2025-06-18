using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using EndesaBusiness.cnmc;
using EndesaEntity.extrasistemas;

namespace EndesaBusiness.Validators.Validations
{
    class FechaactivacionValidation : ICellValidator
     {
        public CellValidationResult Validate(string cellValue, Global global, CNMC diccionary, int row, int column)
        {
            // FechaPrevistaAccion (fecha_activacion) – Obligatorio solo si IndActivacion = 'F', formato X(10) “AAAA-MM-DD”
            if (global.ind_activacion == "F")
            {
                if (!string.IsNullOrWhiteSpace(cellValue))
                {
                    var fechaStr = cellValue.Trim();
                    if (DateTime.TryParse(fechaStr, out DateTime fechaParsed))
                    {
                        // Verifica formato exacto “yyyy-MM-dd”
                        if (fechaParsed.ToString("yyyy-MM-dd") == fechaStr)
                        {
                            global.fecha_activacion = fechaParsed;
                            return CellValidationResult.WithSuccess();
                        }
                        else
                        {
                            return CellValidationResult.WithError(
                                $"Hoja: {global.hoja} Fila: {row} --> CUPS: {global.cups22} FechaPrevistaAccion: el campo fecha_activación no tiene el formato AAAA-MM-DD."
                            );
                        }
                    }
                    else
                    {
                        return CellValidationResult.WithError(
                            $"Hoja: {global.hoja} Fila: {row} --> CUPS: {global.cups22} FechaPrevistaAccion: el campo fecha_activación no es una fecha válida."
                        );
                    }
                }
                else
                {
                    return CellValidationResult.WithError(
                        $"Hoja: {global.hoja} Fila: {row} --> CUPS: {global.cups22} FechaPrevistaAccion: el campo fecha_activación está vacío."
                    );
                }
            }
            // No es obligatorio si IndActivacion != 'F'
            return CellValidationResult.WithSuccess();
        }


    }
}
