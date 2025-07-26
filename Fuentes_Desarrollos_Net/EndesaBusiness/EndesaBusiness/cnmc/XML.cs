using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Xml;

namespace EndesaBusiness.cnmc
{
    public class XML
    {
        public XML()
        {

        }

        public void CreaXML_A1_43(FileInfo file, EndesaEntity.cnmc.XML_A1_43 xml_a143)
        {

            string url = "http://localhost/sctd/A143";

            XmlTextWriter writer = new XmlTextWriter(file.FullName, Encoding.UTF8);


            writer.WriteStartDocument();
            writer.WriteStartElement("sctdapplication", url);
            {
                writer.WriteStartElement("heading");
                {
                    writer.WriteElementString("dispatchingcode", "GML");
                    writer.WriteElementString("dispatchingcompany", xml_a143.dispatchingcompany);
                    writer.WriteElementString("destinycompany", xml_a143.destinycompany);                    
                    writer.WriteElementString("communicationsdate", DateTime.Now.ToString("yyyy-MM-dd"));
                    writer.WriteElementString("communicationshour", DateTime.Now.ToString("HH:mm:ss"));
                    writer.WriteElementString("processcode", "43");
                    writer.WriteElementString("messagetype", "A1");
                }
                writer.WriteEndElement(); // heading

                writer.WriteStartElement("a143");
                {
                   
                    writer.WriteElementString("comreferencenum", xml_a143.comreferencenum);
                    writer.WriteElementString("reqdate", DateTime.Now.ToString("yyyy-MM-dd"));
                    writer.WriteElementString("reqhour", DateTime.Now.ToString("HH:mm:ss"));
                    writer.WriteElementString("titulartype", "J"); 
                    writer.WriteElementString("nationality", "ES");
                    writer.WriteElementString("documenttype", "01");
                    writer.WriteElementString("documentnum", xml_a143.documentnum);
                    writer.WriteElementString("cups", xml_a143.cups);
                    writer.WriteElementString("modeffectdate", "04");
                    writer.WriteElementString("reqtransferdate", xml_a143.productstartdate.ToString("yyyy-MM-dd"));
                    writer.WriteStartElement("productlist");
                    {
                        writer.WriteStartElement("product");
                        {
                            //writer.WriteElementString("reqtype", "03");
                            xml_a143.reqtype = xml_a143.reqtype + 1;
                            writer.WriteElementString("reqtype", xml_a143.reqtype.ToString().PadLeft(2,'0'));
                            if(xml_a143.reqtype == 1 || xml_a143.reqtype == 2)
                                writer.WriteElementString("productcode", xml_a143.productcode);

                            writer.WriteElementString("producttype", xml_a143.producttype);
                            writer.WriteElementString("producttolltype", xml_a143.producttolltype); // Peaje

                            if (xml_a143.producttype == "06" && xml_a143.productqi == 0)                            
                                writer.WriteElementString("productqd", xml_a143.productqd.ToString());                            
                            else if(xml_a143.producttype == "06" && xml_a143.productqd == 0)
                                writer.WriteElementString("productqi", xml_a143.productqi.ToString());
                            else
                                writer.WriteElementString("productqd", xml_a143.productqd.ToString());

                            if(xml_a143.productqd > 0)
                                writer.WriteElementString("productqa", (xml_a143.productqd * 330).ToString());
                            else
                                writer.WriteElementString("productqa", (xml_a143.productqi * 330).ToString());

                            if (xml_a143.producttype == "06")
                            {
                                writer.WriteElementString("productstarthour", xml_a143.startHour.ToString("HH:mm:ss"));
                            }
                                
                        }
                        writer.WriteEndElement(); // product
                    }
                    writer.WriteEndElement(); // productlist
                    
                }
                writer.WriteEndElement(); // a104


            }


            writer.WriteEndElement(); //sctdapplication
            writer.WriteEndDocument();

            writer.Close();


        }

        public void CreaXML_A1_05(FileInfo file, EndesaEntity.cnmc.XML_A1_43 xml_a143)
        {
            string url = "http://localhost/sctd/A105";

            XmlTextWriter writer = new XmlTextWriter(file.FullName, Encoding.UTF8);


            writer.WriteStartDocument();
            writer.WriteStartElement("sctdapplication", url);            
            {
                writer.WriteStartElement("heading");
                writer.WriteElementString("dispatchingcode", "GML");
                writer.WriteElementString("dispatchingcompany", xml_a143.dispatchingcompany);
                writer.WriteElementString("destinycompany", xml_a143.destinycompany);
                writer.WriteElementString("communicationsdate", DateTime.Now.ToString("yyyy-MM-dd"));
                writer.WriteElementString("communicationshour", DateTime.Now.ToString("HH:mm:ss"));
                writer.WriteElementString("processcode", "05");
                writer.WriteElementString("messagetype", "A1");
                writer.WriteEndElement(); // heading
            }

            writer.WriteStartElement("a105");
            {
                writer.WriteElementString("comreferencenum", xml_a143.comreferencenum);
                writer.WriteElementString("reqdate", DateTime.Now.ToString("yyyy-MM-dd"));
                writer.WriteElementString("reqhour", DateTime.Now.ToString("HH:mm:ss"));                
                writer.WriteElementString("nationality", "ES");
                writer.WriteElementString("documenttype", "01");
                writer.WriteElementString("documentnum", xml_a143.documentnum);
                writer.WriteElementString("cups", xml_a143.cups);
                writer.WriteElementString("modeffectdate", "04");
                writer.WriteElementString("reqtransferdate", "2021-10-01");                
                writer.WriteElementString("updatereason", "25");
                writer.WriteElementString("newtolltype", xml_a143.newtolltype);


            }

            

            writer.WriteEndElement(); //sctdapplication
            writer.WriteEndDocument();

            writer.Close();
        }

    }
}
