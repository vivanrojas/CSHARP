using Aspose.Cells;
using EndesaBusiness.cnmc;
using EndesaEntity.extrasistemas;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace EndesaBusiness.Validators.Validations
{

    class EmpresaEmisoraValidation : ICellValidator
    {

        public CellValidationResult Validate(string cellValue, Global global, CNMC diccionary, int row, int column)
        {
            
            if (cellValue != null && cellValue.Length >= 4)
            {
                if (int.TryParse(cellValue.Substring(0, 4), out _))
                {
                    global.empresa_emisora = (cellValue).Substring(0, 4);
                    return CellValidationResult.WithSuccess();
                }
                else
                {
                   return CellValidationResult.WithError( "Hoja: " + global.hoja + " Fila: " + row + " --> " + " Empresa_emisora: el campo código empresa emisora no tiene un formato válido.");

                   // registro_valido = false;
                }
            }
            else
            {
                return CellValidationResult.WithError("Hoja: " + global.hoja + " Fila: " + row + " --> " + " Empresa_emisora: el campo código empresa emisora es nulo o no tiene una longitud válida X(4).");
                //lista_log.Add(System.Environment.NewLine);
                //registro_valido = false;
            }

        }
    }
}

