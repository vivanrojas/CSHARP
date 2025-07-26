using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using EndesaBusiness.cnmc;
using EndesaEntity.extrasistemas;

namespace EndesaBusiness.Validators.Validations
{
    class ContratIncondBSValidation : ICellValidator
    {

        public CellValidationResult Validate(string cellValue, Global global, CNMC diccionary, int row, int column)
        {
            // ContratacionIncondicionalBS – Valor fijo X(1) = "N"
            global.contratacion_incondicional_bs = "N";
            return CellValidationResult.WithSuccess();
        }



    }
}
