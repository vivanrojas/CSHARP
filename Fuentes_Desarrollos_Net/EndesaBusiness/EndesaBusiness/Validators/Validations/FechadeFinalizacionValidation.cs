using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using EndesaBusiness.cnmc;
using EndesaEntity.extrasistemas;

namespace EndesaBusiness.Validators.Validations
{
     //internal 
    class FechadeFinalizacionValidation :ICellValidator
    {
        public CellValidationResult Validate( string cellValue, Global global, CNMC diccionary, int row, int column)
        {
            // Fecha Finalización – Obligatorio X(10) “AAAA-MM-DD” solo si tipo_contrato_atr está en { "02", "03", "09" }
            var contratosRequierenFecha = new[] { "02", "03", "09" };

            if (contratosRequierenFecha.Contains(global.tipo_contrato_atr))
            {
                // Se espera un valor no nulo
                if (!string.IsNullOrWhiteSpace(cellValue))
                {
                    var fechaStr = cellValue.Trim();
                    // Parse exacto “yyyy-MM-dd”
                    if (DateTime.TryParseExact(
                            fechaStr,
                            "yyyy-MM-dd",
                            CultureInfo.InvariantCulture,
                            DateTimeStyles.None,
                            out DateTime fechaFinalizacion))
                    {
                        global.fecha_finalizacion = fechaFinalizacion;
                        return CellValidationResult.WithSuccess();
                    }
                    else
                    {
                        return CellValidationResult.WithError(
                            $"Hoja: {global.hoja} Fila: {row} --> CUPS: {global.cups22} FechaFinalización: '{fechaStr}' no es válida (debe ser AAAA-MM-DD)."
                        );
                    }
                }
                else
                {
                    return CellValidationResult.WithError(
                        $"Hoja: {global.hoja} Fila: {row} --> CUPS: {global.cups22} FechaFinalización: es obligatoria para tipo_contrato_atr '{global.tipo_contrato_atr}'."
                    );
                }
            }

            // No aplica si el contrato no requiere fecha
            return CellValidationResult.WithSuccess();
        }


    }
}
