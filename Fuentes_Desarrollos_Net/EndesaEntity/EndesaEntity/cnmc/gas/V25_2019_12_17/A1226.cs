using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Xml.Serialization;

namespace EndesaEntity.cnmc.gas.V25_2019_12_17
{
    [XmlRoot(ElementName = "a1226")]
    public class A1226
    {
        [XmlElement(ElementName = "reqdate")] public DateTime reqdate { get; set; }
        [XmlElement(ElementName = "reqhour")] public DateTime reqhour { get; set; }
        [XmlElement(ElementName = "atrcode")] public string atrcode { get; set; }
        [XmlElement(ElementName = "cups")] public string cups { get; set; }
        [XmlElement(ElementName = "nationality")] public string nationality { get; set; }
        [XmlElement(ElementName = "documenttype")] public string documenttype { get; set; }
        [XmlElement(ElementName = "documentnum")] public string documentnum { get; set; }
        [XmlElement(ElementName = "firstname")] public string firstname { get; set; }
        [XmlElement(ElementName = "familyname1")] public string familyname1 { get; set; }
        [XmlElement(ElementName = "familyname2")] public string familyname2 { get; set; }
        [XmlElement(ElementName = "telephone")] public string telephone { get; set; }
        [XmlElement(ElementName = "fax")] public string fax { get; set; }
        [XmlElement(ElementName = "newcustomer")] public string newcustomer { get; set; }
        [XmlElement(ElementName = "email")] public string email { get; set; }
        [XmlElement(ElementName = "streettype")] public string streettype { get; set; }
        [XmlElement(ElementName = "street")] public string street { get; set; }
        [XmlElement(ElementName = "streetnumber")] public string streetnumber { get; set; }
        [XmlElement(ElementName = "portal")] public string portal { get; set; }
        [XmlElement(ElementName = "staircase")] public string staircase { get; set; }
        [XmlElement(ElementName = "floor")] public string floor { get; set; }
        [XmlElement(ElementName = "door")] public string door { get; set; }
        [XmlElement(ElementName = "province")] public string province { get; set; }
        [XmlElement(ElementName = "city")] public string city { get; set; }
        [XmlElement(ElementName = "zipcode")] public string zipcode { get; set; }
        [XmlElement(ElementName = "tolltype")] public string tolltype { get; set; }
        [XmlElement(ElementName = "qdgranted")] public double qdgranted { get; set; }
        [XmlElement(ElementName = "qhgranted")] public double qhgranted { get; set; }
        [XmlElement(ElementName = "singlenomination")] public string singlenomination { get; set; }
        [XmlElement(ElementName = "transfereffectivedate")] public DateTime transfereffectivedate { get; set; }
        [XmlElement(ElementName = "finalclientyearlyconsumption")] public int finalclientyearlyconsumption { get; set; }
        [XmlElement(ElementName = "netsituation")] public string netsituation { get; set; }
        [XmlElement(ElementName = "outgoingpressuregranted")] public double outgoingpressuregranted { get; set; }
        [XmlElement(ElementName = "lastinspectionsdate")] public DateTime lastinspectionsdate { get; set; }
        [XmlElement(ElementName = "lastinspectionsresult")] public string lastinspectionsresult { get; set; }
        [XmlElement(ElementName = "readingtype")] public string readingtype { get; set; }
        [XmlElement(ElementName = "rentingamount")] public double rentingamount { get; set; }
        [XmlElement(ElementName = "rentingperiodicity")] public string rentingperiodicity { get; set; }
        [XmlElement(ElementName = "canonircamount")] public double canonircamount { get; set; }
        [XmlElement(ElementName = "canonircperiodicity")] public string canonircperiodicity { get; set; }
        [XmlElement(ElementName = "canonircforlife")] public string canonircforlife { get; set; }
        [XmlElement(ElementName = "canonircdate")] public DateTime canonircdate { get; set; }
        [XmlElement(ElementName = "canonircmonth")] public string canonircmonth { get; set; }
        [XmlElement(ElementName = "othersamount")] public double othersamount { get; set; }
        [XmlElement(ElementName = "othersperiodicity")] public string othersperiodicity { get; set; }
        [XmlElement(ElementName = "readingperiodicitycode")] public string readingperiodicitycode { get; set; }
        [XmlElement(ElementName = "transporter")] public string transporter { get; set; }
        [XmlElement(ElementName = "transnet")] public string transnet { get; set; }
        [XmlElement(ElementName = "gasusetype")] public string gasusetype { get; set; }
        [XmlElement(ElementName = "caecode")] public string caecode { get; set; }
        [XmlElement(ElementName = "communicationreason")] public string communicationreason { get; set; }
        [XmlElement(ElementName = "titulartype")] public string titulartype { get; set; }
        [XmlElement(ElementName = "regularaddress")] public string regularaddress { get; set; }
        [XmlElement(ElementName = "counterlist")]  public CounterList counterlist { get; set; }

        public A1226()
        {
            counterlist = new CounterList();
        }
       

    }
        
}
