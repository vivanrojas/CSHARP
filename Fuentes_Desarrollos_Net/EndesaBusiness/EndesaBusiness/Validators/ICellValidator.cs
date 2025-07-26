using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using EndesaBusiness.cnmc;
using EndesaEntity.extrasistemas;
using EndesaBusiness.Validators;


namespace EndesaBusiness.Validators
{
    public interface ICellValidator
    {

        CellValidationResult Validate(string cellValue, Global global, CNMC diccionary, int row, int column);


    }
    //public class StringNotNullValidation : ICellValidator
    //{
    //    public CellValidationResult Validate(string cellValue, Global global, CNMC diccionary, int row, int column)
    //    {
    //        if (string.IsNullOrWhiteSpace(cellValue))
    //        {
    //            return CellValidationResult.WithError(
    //                $"Hoja: {global.hoja} Fila: {row} --> El valor no puede ser nulo o vacío."
    //            );
    //        }

    //        return CellValidationResult.WithSuccess();
    //    }
    //}
}