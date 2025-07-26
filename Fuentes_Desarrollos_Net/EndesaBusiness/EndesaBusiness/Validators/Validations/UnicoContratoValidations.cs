using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using EndesaBusiness.cnmc;
using EndesaEntity.extrasistemas;



namespace EndesaBusiness.Validators.Validations
{
    class UnicoContratoValidations : ICellValidator
    {

        public CellValidationResult Validate(string cellValue, Global global, CNMC diccionary, int row, int column)
        {
            if (cellValue != null && cellValue.Length >= 26)
            {
                //global.cau = Convert.ToString(cellValue);
                global.cau = cellValue;
                return CellValidationResult.WithSuccess();
            }
            else
            {
                // Mensaje de error si es nulo o de longitud insuficiente
                return CellValidationResult.WithError(
                    "Hoja: " + global.hoja +
                    " Fila: " + row +
                    " --> cau: el campo cau es nulo o no tiene una longitud válida."
                );
            }

        }
    }
}
