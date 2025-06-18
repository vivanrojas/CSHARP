using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using EndesaBusiness.cnmc;
using EndesaEntity.extrasistemas;

namespace EndesaBusiness.Validators.Validations
{
    class PisoValidation : ICellValidator
    {
        public CellValidationResult Validate(string cellValue, Global global, CNMC diccionary, int row, int column)
        {
            if (string.IsNullOrWhiteSpace(cellValue))
            {
                global.piso = "";
                return CellValidationResult.WithSuccess();
            }

            string valorCeldaPiso = cellValue.Trim().ToUpper();
            string[] partesPiso = valorCeldaPiso.Split('-');
            string codigoPiso = partesPiso[0].Trim();

            // Intentar convertir a número de hasta 3 dígitos
            if (codigoPiso.Length <= 3 && int.TryParse(codigoPiso, out int numeroPiso))
            {
                global.piso = numeroPiso.ToString("D3");  // Formato 001, 002, etc.
                return CellValidationResult.WithSuccess();
            }

            // Verificar si el valor abreviado está en el diccionario
            string codigoPisoAbreviado = codigoPiso.Length >= 2 ? codigoPiso.Substring(0, 2) : codigoPiso;

            if (diccionary.dic_piso.ContainsValue(codigoPisoAbreviado))
            {
                global.piso = codigoPisoAbreviado;
                return CellValidationResult.WithSuccess();
            }

            return CellValidationResult.WithError(
                $"Hoja: {global.hoja} Fila: {row} --> El valor del campo 'piso' ({codigoPiso}) no es válido."
            );
        }
    }
}
