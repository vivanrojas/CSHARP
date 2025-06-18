using System;
using System.Collections.Generic;
using System.Runtime.InteropServices;
using System.Text;
using EndesaBusiness.cnmc;

//using Aspose.Cells;
using EndesaEntity.extrasistemas;
using OfficeOpenXml;

namespace EndesaBusiness.Validators
{
    public class OpenOfficeSheetValidator
    {
        private readonly IList<ICellValidator> _cellValidators = new List<ICellValidator>();
        private readonly int _firstRow;
        private readonly ExcelWorksheet _worksheet;

        private OpenOfficeSheetValidator(int firstRow, ExcelWorksheet worksheet)
        {
            _firstRow = firstRow;
            _worksheet = worksheet;
        }
        
        public SheetValidationResult Validate(Global global, CNMC dictionary)
        {
            if (_cellValidators.Count == 0) throw new InvalidOperationException("no validate found");
            var errors = new StringBuilder(_cellValidators.Count);
            var cellCont = 0;
            foreach (var cellValidator in _cellValidators)
            {
                var cellValue = _worksheet.Cells[_firstRow, cellCont].Value;
                var cellResult = cellValidator.Validate(cellValue.ToString(), global, dictionary, _firstRow, cellCont); 
                cellCont++;

                if (cellResult.IsValid) errors.Append(cellResult.ErrorMessage);
            }

            return new SheetValidationResult()
            {
                IsValid = errors.Length == 0,
                ErrorMessage = errors.ToString()
            };
        }

        public OpenOfficeSheetValidator WithValidator(ICellValidator validator)
        {
            
            if (validator == null)
                throw new ArgumentNullException(nameof(validator));
            _cellValidators.Add(validator);

            return this;
        }

        public static OpenOfficeSheetValidator Create(int firstLine, ExcelWorksheet worksheet)
        {
            if (firstLine < 0)
                throw new ArgumentOutOfRangeException(nameof(firstLine));
                      
            if (worksheet == null)
                throw new ArgumentNullException(nameof(worksheet));

            return new OpenOfficeSheetValidator(firstLine, worksheet);
        }
    }
}