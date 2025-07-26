using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using EndesaBusiness.cnmc;
using EndesaEntity.extrasistemas;

namespace EndesaBusiness.Validators.Validations
{
    class PersonadeContactoValidation  : ICellValidator
    {
        public CellValidationResult Validate(string cellValue, Global global, CNMC diccionary, int row, int column)
        {
            if (!string.IsNullOrWhiteSpace(cellValue))
            {
                // Cortar a máximo 150 caracteres
                var contacto = cellValue.Substring(0, Math.Min(cellValue.Length, 150));
                global.persona_contacto = contacto;
                global.contacto = contacto;
                return CellValidationResult.WithSuccess();
            }
            else
            {
                return CellValidationResult.WithError(
                    $"Hoja: {global.hoja} Fila: {row} --> CUPS: {global.cups22} El campo 'persona de contacto' es obligatorio y está vacío."
                );
            }
        }



    }
}
