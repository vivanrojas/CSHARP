using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using EndesaBusiness.cnmc;
using EndesaEntity.extrasistemas;

namespace EndesaBusiness.Validators.Validations
{
    class TelefonoValidation : ICellValidator
    {
        public CellValidationResult Validate(string cellValue, Global global, CNMC diccionario, int row, int column)
        {
            if (string.IsNullOrWhiteSpace(cellValue))
            {
                return CellValidationResult.WithError(
                    $"Hoja: {global.hoja} Fila: {row} --> Teléfono: el campo está vacío."
                );
            }

            string valorTelefono = cellValue.Trim();

            // Verificar si contiene más de un teléfono
            if (valorTelefono.Contains(","))
            {
                return CellValidationResult.WithError(
                    $"Hoja: {global.hoja} Fila: {row} --> Teléfono: el campo contiene más de un número de teléfono."
                );
            }

            // Verificar formato: solo dígitos, entre 6 y 12 caracteres
            if (!System.Text.RegularExpressions.Regex.IsMatch(valorTelefono, @"^\d{6,12}$"))
            {
                return CellValidationResult.WithError(
                    $"Hoja: {global.hoja} Fila: {row} --> Teléfono: el número '{valorTelefono}' no es válido. Debe tener entre 6 y 12 dígitos."
                );
            }

            // Asignación si válido
            global.telefono = valorTelefono;
            global.tlf_contacto = valorTelefono;

            return CellValidationResult.WithSuccess();
        }
    }
}
