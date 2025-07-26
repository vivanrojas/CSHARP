namespace EndesaBusiness.Validators
{
    public class CellValidationResult
    {
        public readonly bool IsValid;
        public readonly string ErrorMessage;

        private CellValidationResult(bool isValid, string errorMessage = null)
        {
            IsValid = isValid;
            ErrorMessage = errorMessage;
        }

        public static CellValidationResult WithSuccess()
        {
            return new CellValidationResult(true);
        }

        public static CellValidationResult WithError(string errorMessage)
        {
            return new CellValidationResult(false, errorMessage);
        }
    }
}