using EndesaBusiness.servidores;
using Org.BouncyCastle.Crypto.Agreement.Srp;
using System;
using System.Collections.Generic;
using System.Data.Odbc;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace EndesaBusiness.contratacion.Redshift
{
    public class Rechazos_ESP_D4
    {
        utilidades.Param p;
        utilidades.Seguimiento_Procesos ss_pp;
        public Rechazos_ESP_D4()
        {
           
        }


        public void EjecutaExtraccion()
        {
            servidores.RedShiftServer db;
            OdbcCommand command;
            OdbcDataReader r;
            

            db = new servidores.RedShiftServer(RedShiftServer.Entornos.PROD);
            command = new OdbcCommand(Consulta(), db.con);
            r = command.ExecuteReader();
            while (r.Read())
            {

            }
            db.CloseConnection();
        }


        private string Consulta()
        {
            string strSql = "";

            strSql = "select"
                + "atr.de_empr,atr.cd_cups,atr.id_solicitud_atr,atr.req_name,"
                + " atr.de_tipo_sol,atr.de_sub_tipo_sol,atr.fh_rechazo :: date,"
                + " atr.cd_tp_cli,atr.cd_nif_cif_cli,atr.tx_nom_cli,atr.tx_apell_cli,"
                + " atr.cd_linea_negocio,atr.nm_pot_ctatada_1,atr.nm_pot_ctatada_2,"
                + " atr.nm_pot_ctatada_3,atr.nm_pot_ctatada_4,atr.nm_pot_ctatada_5,"
                + " atr.nm_pot_ctatada_6,"
                + " case when de_tipo_sol='D1(Cambios ATR desde distribuidor)' then request.RD4_REQ_RejectionReason__c else"
                + " cd_mot_rechazo_1 end as cd_mot_rechazo_1,case when de_tipo_sol='D1(Cambios ATR desde distribuidor)' then"
                + " rejection.label else de_mot_rechazo_1 end as de_mot_rechazo_1,case when de_tipo_sol='D1(Cambios ATR desde distribuidor)' then"
                + " request.RD4_REQ_RejectionComment__c else de_comentario_mot_rechazo1 end as de_comentario_mot_rechazo1,"
                + " atr.cd_mot_rechazo_2,atr.de_mot_rechazo_2,atr.de_comentario_mot_rechazo2,atr.cd_mot_rechazo_3,"
                + " atr.de_mot_rechazo_3,atr.de_comentario_mot_rechazo3,atr.cd_mot_rechazo_4,atr.de_mot_rechazo_4,"
                + " atr.de_comentario_mot_rechazo4,atr.cd_mot_rechazo_5,atr.de_mot_rechazo_5,atr.de_comentario_mot_rechazo5,"
                + " atr.cd_mot_rechazo_6,atr.de_mot_rechazo_6,atr.de_comentario_mot_rechazo6,atr.cd_mot_rechazo_7,"
                + " atr.de_mot_rechazo_7,atr.de_comentario_mot_rechazo7,atr.cd_mot_rechazo_8,atr.de_mot_rechazo_8,"
                + " atr.de_comentario_mot_rechazo8,atr.cd_mot_rechazo_9,atr.de_mot_rechazo_9,atr.de_comentario_mot_rechazo9,"
                + " atr.cd_mot_rechazo_10,atr.de_mot_rechazo_10,atr.de_comentario_mot_rechazo10,"
                + " atr.fh_env_sistema :: date,atr.atr_gestionado_cliente,atr.de_territory,atr.de_canal_entrada,"
                + " atr.de_tipo_rechazo,atr.lg_rechazo_tras_aceptacion,atr.lg_rechazo_manual,atr.name_empr_distdora,"
                + " atr.de_tension_contratar,atr.de_tarifa_contratar"
                + " from ed_owner.t_ed_h_sol_atr atrleft join ods_owner.ods_rstt_rd4_request__c requeston atr.id=request.id "
                + " left join (select value, label from ods_owner.ods_rstt_picklistvalueinfo"
                + " where durableid like 'RD4_Request__c.RD4_REQ_RejectionReason__c%' group by value, label) rejection on "
                + " request.RD4_REQ_RejectionReason__c=rejection.valuewhereatr.cd_pais<>'PORTUGAL' and"
                + " atr.cd_tp_tension='SE' and atr.cd_linea_negocio='Electricidad' and rtrim(ltrim(cd_estado_sol)) in ('11','12')"
                + " and cd_sub_tipo_sol<>'MI'and to_char(fh_rechazo,'MMYYYY')=to_char(dateadd(mm,-1,sysdate),'MMYYYY')";

            return strSql;
        }

    }
}
