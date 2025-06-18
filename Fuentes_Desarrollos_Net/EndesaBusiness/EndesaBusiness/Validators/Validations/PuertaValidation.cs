using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using EndesaBusiness.cnmc;
using EndesaEntity.extrasistemas;

namespace EndesaBusiness.Validators.Validations
{
    public class PuertaValidation : ICellValidator
    {
        public CellValidationResult Validate(string cellValue, Global global, CNMC diccionario, int row, int column)
        {
            if (string.IsNullOrWhiteSpace(cellValue))
            {
                global.puerta = "";
                return CellValidationResult.WithSuccess();
            }

            string valorCeldaPuerta = cellValue.Trim().ToUpper();
            string[] partesPuerta = valorCeldaPuerta.Split('-');
            string codigoPuerta = partesPuerta[0].Trim();

            // Verificar si es un número de hasta 3 dígitos
            if (codigoPuerta.Length <= 3 && int.TryParse(codigoPuerta, out int numeroPuerta))
            {
                global.puerta = numeroPuerta.ToString("D3"); // Formato 001, 002, etc.
                return CellValidationResult.WithSuccess();
            }

            // Verificar abreviatura
            string codigoPuertaAbreviado = codigoPuerta.Length >= 2 ? codigoPuerta.Substring(0, 2) : codigoPuerta;

            if (diccionario.dic_puerta.ContainsValue(codigoPuertaAbreviado))
            {
                global.puerta = codigoPuertaAbreviado;
                return CellValidationResult.WithSuccess();
            }

            return CellValidationResult.WithError(
                $"Hoja: {global.hoja} Fila: {row} --> El valor del campo 'puerta' ({codigoPuerta}) no es válido."
            );
        }
    }
}
