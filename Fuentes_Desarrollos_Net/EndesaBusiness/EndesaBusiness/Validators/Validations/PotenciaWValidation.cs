using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using EndesaBusiness.cnmc;
using EndesaEntity.extrasistemas;
using FluentFTP.Helpers;

namespace EndesaBusiness.Validators.Validations
{
    class PotenciaWValidation : ICellValidator
    {
        public CellValidationResult Validate(string cellValue, Global global, CNMC diccionary, int row, int column)
        {
            int potenciasInformadas = 0;
            bool esTarifa20TD = global.tarifa == "2.0TD";
            int[] potencias = new int[6];

            for (int i = 0; i < 6; i++)
            {
                var potenciaRaw = 10;
                //var potenciaRaw = diccionary.GetCellValue(row, column + i); // Se asume un método auxiliar para obtener celdas relacionadas
                string potenciaStr = Convert.ToString(potenciaRaw);

                if (!string.IsNullOrWhiteSpace(potenciaStr) && int.TryParse(potenciaStr, out int potencia))
                {
                    if (potencia >= 0)
                    {
                        potencias[i] = potencia;
                        potenciasInformadas++;
                    }
                    else
                    {
                        return CellValidationResult.WithError(
                            $"Hoja: {global.hoja} Fila: {row} --> Potencia {i + 1}: el valor debe ser un número positivo."
                        );
                    }
                }
                else
                {
                    return CellValidationResult.WithError(
                        $"Hoja: {global.hoja} Fila: {row} --> Potencia {i + 1}: el valor no es un número válido o está en blanco."
                    );
                }
            }

            if (!esTarifa20TD)
            {
                for (int i = 1; i < 6; i++)
                {
                    if (potencias[i] < potencias[i - 1])
                    {
                        return CellValidationResult.WithError(
                            $"Hoja: {global.hoja} Fila: {row} --> Las potencias deben ser crecientes. Potencia {i + 1} ({potencias[i]}) es menor que Potencia {i} ({potencias[i - 1]})."
                        );
                    }
                }
            }

            global.potencias = potencias;
            return CellValidationResult.WithSuccess();
        }
    }
}




