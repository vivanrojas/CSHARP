using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using EndesaBusiness.cnmc;
using EndesaEntity.extrasistemas;

namespace EndesaBusiness.Validators.Validations
{
    class NombreDePilaValidation : ICellValidator
    {

        public CellValidationResult Validate(string cellValue, Global global, CNMC diccionary, int row, int column)

        {
            if (!string.IsNullOrEmpty(cellValue))
            {
                global.nombre_de_pila = cellValue.Length > 50 ? cellValue.Substring(0, 50) : cellValue;
                return CellValidationResult.WithSuccess();
            }
            else
            {
                return CellValidationResult.WithError("Hoja: " + global.hoja + " Fila: " + row + " --> Nombre: el campo nombre de pila es nulo o vacío.");
            }
        }




    }
}
