using EndesaBusiness.cnmc;
using EndesaEntity.extrasistemas;

namespace EndesaBusiness.Validators.Validations
{
    public class StringNotNullValidation : ICellValidator
    {
        public CellValidationResult Validate(string cellValue, Global global, CNMC diccionary, int row, int column)
        {
            var isValid = !string.IsNullOrEmpty(cellValue);

            return isValid 
                ? CellValidationResult.WithSuccess() 
                : CellValidationResult.WithError("Error: cell value is null or empty");
        }
    }
}