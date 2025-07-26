using EndesaBusiness.cnmc;
using EndesaEntity.extrasistemas;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace EndesaBusiness.Validators.Validations
{
    class DistribuidoraValidation : ICellValidator
    {
        public CellValidationResult Validate(string cellValue, Global global, CNMC diccionary, int row, int column)
        
        {
            if (cellValue != null && cellValue.Length >= 4)
            {
                if (int.TryParse(cellValue.Substring(0, 4), out _))
                {

                    global.distribuidora = (cellValue).Substring(0, 4);
                    return CellValidationResult.WithSuccess();
                }
                else
                {
                    return CellValidationResult.WithError("Hoja: " + global.hoja + " Fila: " + row +
                        " --> " + " Distribuidora: el campo código empresa destino (distribuidora) no tiene un formato válido.");
                }   
            }
            else
            {
                return CellValidationResult.WithError("Hoja: " + global.hoja + " Fila: " + row +
                    " --> " + " Distribuidora: el campo código empresa destino (distribuidora) es nulo o no tiene una longitud válida X(4).");
                             
            }

        }

    }

    
}
