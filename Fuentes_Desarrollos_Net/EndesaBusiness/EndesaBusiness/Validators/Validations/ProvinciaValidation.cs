using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using EndesaBusiness.cnmc;
using EndesaEntity.extrasistemas;

namespace EndesaBusiness.Validators.Validations
{
    class ProvinciaValidation : ICellValidator
    {
        public CellValidationResult Validate(string cellValue, Global global, CNMC dictionary, int row, int column)
        {
            // Provincia - Obligatorio X(2)

            if (!string.IsNullOrWhiteSpace(cellValue) && cellValue.Trim().Length >= 2)
            {
                string valorProvincia = cellValue.Trim().Substring(0, 2).ToUpper();

                //var tipoAutoconsumo = cellValue.Substring(0, 2).ToUpper();
                //if (cnmc.dic_autoconsumo.ContainsValue(tipoAutoconsumo))

                    //if (dicProvincias.ContainsValue(valorProvincia))
                if (dictionary.dic_provincias.ContainsValue(valorProvincia))
                {
                    global.provincia = valorProvincia;
                    return CellValidationResult.WithSuccess();
                }
                else
                {
                    return CellValidationResult.WithError(
                        $"Hoja: {global.hoja} Fila: {row} --> provincia: el campo no contiene un valor válido."
                    );
                }
            }
            else
            {
                return CellValidationResult.WithError(
                    $"Hoja: {global.hoja} Fila: {row} --> provincia: el campo es nulo o no tiene una longitud válida (X2)."
                );
            }
        }


    }
}
