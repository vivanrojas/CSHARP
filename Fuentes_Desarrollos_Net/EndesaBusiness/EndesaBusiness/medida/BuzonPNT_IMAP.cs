using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Net;


namespace EndesaBusiness.medida
{
    public class BuzonPNT_IMAP
    {

        utilidades.Param param;
        utilidades.ParamUser paramUser;
        EndesaEntity.ReglaCorreo regla;

        logs.Log ficheroLog = new logs.Log(Environment.CurrentDirectory, "logs", "MED_pnts");

        //Dictionary<string, EndesaEntity.Pnt  meterpoint.MeterPoint> mpListCUPS20 = 
        //    new Dictionary<string, meterpoint.MeterPoint>(); Dictionary<string, meterpoint.MeterPoint> mpListCUPS20 = new Dictionary<string, meterpoint.MeterPoint>();
        public BuzonPNT_IMAP()
        {
            utilidades.Credenciales credenciales = new utilidades.Credenciales("RSIOPEGMA001");
            param = new utilidades.Param("pnt_param", servidores.MySQLDB.Esquemas.MED);
                
            paramUser = new utilidades.ParamUser("pnt_param_user", "RSIOPEGMA001", servidores.MySQLDB.Esquemas.MED);
            regla = new EndesaEntity.ReglaCorreo();
        }

        public void RecorreInbox()
        {
            using (var ic = new AE.Net.Mail.ImapClient("outlook.office365.com", "rsiope.gma@enel.com", "@@Peti1004@@", AE.Net.Mail.AuthMethods.Login, 993, true))
            {
                ic.SelectMailbox("pnts@enel.com");
                //MailMessage[] mm = ic.GetMessages(0, 10);
            }
        }
    }
}
