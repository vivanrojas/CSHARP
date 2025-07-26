using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using EndesaBusiness.cnmc;
using EndesaEntity.extrasistemas;

namespace EndesaBusiness.Validators.Validations
{
    class TlfcontactoValidation : ICellValidator
    {

        public CellValidationResult Validate(string cellValue, Global global, CNMC dictionary, int row, int column)
        {
            if (string.IsNullOrWhiteSpace(cellValue))
            {
                return CellValidationResult.WithError(
                $"Hoja: {global.hoja} --> CUPS: {global.cups22} El campo teléfono no tiene un valor válido."
                );
            }

            // Eliminar todos los caracteres que no sean dígitos
            string telefono = new string(cellValue.Where(char.IsDigit).ToArray());

            if (telefono.Length >= 6 && telefono.Length <= 12)
            {
                global.tlf_contracto = telefono;
                return CellValidationResult.WithSuccess();
            }
            else
            {
                return CellValidationResult.WithError(
                $"Hoja: {global.hoja} --> CUPS: {global.cups22} El campo teléfono debe tener entre 6 y 12 dígitos."
                );
            }
        }

    }
}
