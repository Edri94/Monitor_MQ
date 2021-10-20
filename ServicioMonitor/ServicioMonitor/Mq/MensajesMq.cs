using ConexionBDSQL;
using ServicioMonitor.Models;
using System;
using System.Data;

namespace ServicioMonitor.Mq
{
    public class MensajesMq : MqSeries
    {
        public string gsAccesoActual;

        public string gsPswdDB;
        public string gsUserDB;
        public string gsNameDB;
        public string gsCataDB;
        public string gsDSNDB;
        public string gsSrvr;

        public string strQuery;

        public ConexionBD cnnConexion;

        /// <summary>
        /// Inserta fila en BITACORA_ERRORES_MENSAJES_PU
        /// </summary>
        /// <param name="bitacora"></param>
        /// <returns></returns>
        public string psInsertarSQL(Bitacora_Errores_Mensajes_Pu bitacora)
        {
            strQuery = "Insert into BITACORA_ERRORES_MENSAJES_PU ";
            strQuery += "(fecha_hora, error_numero, error_descripcion, aplicacion) ";
            strQuery += $"Values ('{bitacora.fecha_hora}', {bitacora.error_numero}, '{bitacora.error_descripcion}', '{bitacora.aplicacion}')";

            return $"Se afectaron {ejecutarInsert(strQuery)} fila(s)";
        }

        /// <summary>
        /// Ejecutar un insert con un query dado
        /// </summary>
        /// <param name="query">query select</param>
        /// <returns></returns>
        public int ejecutarInsert(string query)
        {
            try
            {
                cnnConexion.ActiveConnection = true;
                cnnConexion.ParametersContains = false;
                cnnConexion.CommandType = CommandType.Text;
                cnnConexion.ActiveConnection = true;

                int afectados = cnnConexion.ExecuteNonQuery(query);

                return afectados;
            }
            catch (Exception ex)
            {
                return -1;
            }
        }
    }
}
