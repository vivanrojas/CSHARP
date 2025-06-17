using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Xml.Serialization;
using System.Runtime.Serialization;
using System.ServiceModel;
using System.Collections;

namespace EndesaEntity.cnmc.V21_2019_12_17
{ 
    [XmlRoot(ElementName = "PotenciasContratadas")]
    public class PotenciasContratadas   
    {

        
        [XmlElement(ElementName = "Potencia")] public List<Potencia> Potencia { get; set; }
        
        

        public PotenciasContratadas()
        {            
            
            Potencia = new List<Potencia>();  

        }



    }
}
