using EndesaEntity.cnmc.gas.V25_2019_12_17;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Xml.Serialization;

namespace EndesaEntity.cnmc.gas
{
    [XmlRoot(ElementName = "a1550")]
    public class A1550
    {
        [XmlElement(ElementName = "reqcode")] public string reqcode { get; set; }
        [XmlElement(ElementName = "responsedate")] public DateTime responsedate { get; set; }
        [XmlElement(ElementName = "responsehour")] public DateTime responsehour { get; set; }
        [XmlElement(ElementName = "cups")] public string cups { get; set; }
        [XmlElement(ElementName = "atrcode")] public string atrcode { get; set; }
        [XmlElement(ElementName = "transfereffectivedate")] public DateTime transfereffectivedate { get; set; }
        [XmlElement(ElementName = "enservicio")] public string enservicio { get; set; }
        [XmlElement(ElementName = "tolltype")] public string tolltype { get; set; }
        [XmlElement(ElementName = "telemetering")] public string telemetering { get; set; }
        [XmlElement(ElementName = "finalclientyearlyconsumption")] public int finalclientyearlyconsumption { get; set; }
        [XmlElement(ElementName = "gasusetype")] public string gasusetype { get; set; }        
        [XmlElement(ElementName = "netsituation")] public string netsituation { get; set; }
        [XmlElement(ElementName = "result")] public string result { get; set; }
        [XmlElement(ElementName = "resultdesc")] public string resultdesc { get; set; }
        [XmlElement(ElementName = "nationality")] public string nationality { get; set; }
        [XmlElement(ElementName = "documenttype")] public string documenttype { get; set; }
        [XmlElement(ElementName = "documentnum")] public string documentnum { get; set; }
        [XmlElement(ElementName = "titulartype")] public string titulartype { get; set; }
        [XmlElement(ElementName = "firstname")] public string firstname { get; set; }
        [XmlElement(ElementName = "familyname1")] public string familyname1 { get; set; }
        [XmlElement(ElementName = "familyname2")] public string familyname2 { get; set; }        
        [XmlElement(ElementName = "telephone1")] public string telephone1 { get; set; }
        [XmlElement(ElementName = "email")] public string email { get; set; }
        public AddressPS AddressPS { get; set; }
        [XmlElement(ElementName = "canonircperiodicity")] public string canonircperiodicity { get; set; }        
        [XmlElement(ElementName = "lastinspectionsdate")] public DateTime lastinspectionsdate { get; set; }
        [XmlElement(ElementName = "lastinspectionsresult")] public string lastinspectionsresult { get; set; }
        [XmlElement(ElementName = "StatusPS")] public string StatusPS { get; set; }
        [XmlElement(ElementName = "readingtype")] public string readingtype { get; set; }
        [XmlElement(ElementName = "counterlist")] public CounterList counterlist { get; set; }

        public A1550()
        {
            AddressPS = new AddressPS();
            counterlist = new CounterList();
        }
    }
}
