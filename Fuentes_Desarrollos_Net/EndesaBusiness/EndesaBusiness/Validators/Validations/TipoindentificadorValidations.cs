using EndesaBusiness.cnmc;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using EndesaEntity.extrasistemas;

namespace EndesaBusiness.Validators.Validations
{
    class TipoindentificadorValidations  : ICellValidator 
    
    {
        public CellValidationResult Validate(string cellValue, Global global, CNMC dictionary, int row, int column)
        {
            // tipo_identificador – Obligatorio, mapeo desde dic_identificador
            if (cellValue != null)
            {
                // Normalizar entrada
                var descripcion = cellValue.ToString().Trim().ToUpperInvariant();
                if (!string.IsNullOrEmpty(descripcion))
                {
                    // Normalizar claves del diccionario (sin espacios, mayúsculas)
                    var dicNorm = dictionary.dic_identificador
                        .ToDictionary(
                            kv => kv.Key.Trim().ToUpperInvariant(),
                            kv => kv.Value
                        );

                    // Intentar coincidencia exacta
                    dicNorm.TryGetValue(descripcion, out var valor);

                    // Si no se encontró, intentar coincidencias parciales
                    if (string.IsNullOrEmpty(valor))
                    {
                        foreach (var kvp in dicNorm)
                        {
                            var clavesDic = kvp.Key
                                .Split(new[] { ',', '-', ';' }, StringSplitOptions.RemoveEmptyEntries)
                                .Select(s => s.Trim().ToUpperInvariant());
                            var clavesEnt = descripcion
                                .Split(new[] { ',', '-', ';' }, StringSplitOptions.RemoveEmptyEntries)
                                .Select(s => s.Trim().ToUpperInvariant());

                            if (clavesDic.Intersect(clavesEnt).Any())
                            {
                                valor = kvp.Value;
                                break;
                            }
                        }
                    }

                    // Validar y asignar los dos últimos caracteres
                    if (!string.IsNullOrEmpty(valor))
                    {
                        if (valor.Length >= 2)
                        {
                            global.tipo_identificador = valor
                                .Substring(valor.Length - 2)
                                .ToUpperInvariant();
                            return CellValidationResult.WithSuccess();
                        }
                        else
                        {
                            return CellValidationResult.WithError(
                                $"Hoja: {global.hoja} Fila: {row} --> TipoIdentificador: " +
                                $"el valor asociado a '{descripcion}' en el diccionario es demasiado corto ('{valor}')."
                            );
                        }
                    }
                    else
                    {
                        return CellValidationResult.WithError(
                            $"Hoja: {global.hoja} Fila: {row} --> TipoIdentificador: " +
                            $"la descripción '{descripcion}' no se encontró en el diccionario."
                        );
                    }
                }
                else
                {
                    return CellValidationResult.WithError(
                        $"Hoja: {global.hoja} Fila: {row} --> TipoIdentificador: el valor en la celda está vacío o contiene solo espacios."
                    );
                }
            }
            else
            {
                return CellValidationResult.WithError(
                    $"Hoja: {global.hoja} Fila: {row} --> TipoIdentificador: el campo es nulo."
                );
            }
        }




    }
}
