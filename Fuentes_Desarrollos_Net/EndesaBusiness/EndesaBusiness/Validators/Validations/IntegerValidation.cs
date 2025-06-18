using EndesaBusiness.cnmc;
using EndesaEntity.extrasistemas;

namespace EndesaBusiness.Validators.Validators.Validations
{
    public class IntegerValidation : ICellValidator
    {
        public CellValidationResult Validate(string cellValue, Global global, CNMC diccionary, int row, int column)
        {
            var isValid = int.TryParse(cellValue, out _);

            return isValid 
                ? CellValidationResult.WithSuccess() 
                : CellValidationResult.WithError("Error: cell value is not a valid integer");
        }
      
    }
}