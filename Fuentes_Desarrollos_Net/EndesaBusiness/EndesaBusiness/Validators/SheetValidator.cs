using System;
using System.Collections.Generic;
using System.Runtime.InteropServices;
using System.Text;
using Aspose.Cells;
using EndesaBusiness.cnmc;
using EndesaEntity.extrasistemas;
//using OfficeOpenXml;

namespace EndesaBusiness.Validators
{
    public class SheetValidator
    {
      // private interface ICellValidator;
        private readonly IList<ICellValidator> _cellValidators = new List<ICellValidator>();
        private readonly int _firstRow;
        private readonly Worksheet _worksheet;

        private SheetValidator(int firstRow, Worksheet worksheet)
        {
            _firstRow = firstRow;
            _worksheet = worksheet;
        }

        public SheetValidationResult Validate(Global global, CNMC diccionary)  // se añadio cnmc
        {
            if (_cellValidators.Count == 0) throw new InvalidOperationException("no validate found");
            var errors = new StringBuilder(_cellValidators.Count);
            var cellCont = 0;
            foreach (var cellValidator in _cellValidators)
            {
                var cellValue = _worksheet.Cells[_firstRow, cellCont].Value;
                var cellResult = cellValidator.Validate(cellValue.ToString(), global,diccionary, _firstRow, cellCont); 
                cellCont++;

                if (cellResult.IsValid) errors.Append(cellResult.ErrorMessage);
            }

            return new SheetValidationResult()
            {
                IsValid = errors.Length == 0,
                ErrorMessage = errors.ToString()
            };
        }

        public SheetValidator WithValidator(ICellValidator validator)
        {
            
            if (validator == null)
                throw new ArgumentNullException(nameof(validator));
            _cellValidators.Add(validator);

            return this;
        }

        public static SheetValidator Create(int firstLine, Worksheet worksheet)
        {
            if (firstLine < 0)
                throw new ArgumentOutOfRangeException(nameof(firstLine));
                      
            if (worksheet == null)
                throw new ArgumentNullException(nameof(worksheet));

            return new SheetValidator(firstLine, worksheet);
        }

        
    }
}