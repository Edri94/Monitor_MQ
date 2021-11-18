using ConexionBDSQL;
using ServicioMonitor.Helpers;
using ServicioMonitor.Models;
using System;
using System.Data;
using System.Data.SqlClient;


namespace ServicioMonitor.Data
{
    public class FuncionesBd
    {
        public string gsPswdDB;
        public string gsUserDB;
        public string gsNameDB;
        public string gsCataDB;
        public string gsDSNDB;
        public string gsSrvr;

        public string strQuery;

        public ConexionBD cnnConexion;

        private Encriptacion encriptacion;
        private Funcion funcion;

        public FuncionesBd()
        {
            encriptacion = new Encriptacion();
            funcion = new Funcion();
        }

        public bool ConectDB()
        {
            bool ConectDB = false;
            string section = "conexion";
            try
            {
                string a = funcion.getValueAppConfig("DBCata", section);
                this.gsCataDB = encriptacion.Decrypt(a);
                this.gsDSNDB = encriptacion.Decrypt(funcion.getValueAppConfig("DBDSN", section));
                this.gsSrvr = encriptacion.Decrypt(funcion.getValueAppConfig("DBSrvr", section));
                this.gsUserDB = encriptacion.Decrypt(funcion.getValueAppConfig("DBUser", section));
                this.gsPswdDB = encriptacion.Decrypt(funcion.getValueAppConfig("DBPswd", section));
                this.gsNameDB = encriptacion.Decrypt(funcion.getValueAppConfig("DBName", section));

                string conn_str = $"Data source ={this.gsSrvr}; uid ={this.gsUserDB}; PWD ={this.gsPswdDB}; initial catalog = {this.gsNameDB}";
                this.cnnConexion = new ConexionBD(conn_str);

                ConectDB = true;

                return ConectDB;

            }
            catch (Exception ex)
            {
                funcion.Escribe(ex, "Error");
            }
            return ConectDB;
        }


        /// <summary>
        /// Realizar la inserción del registro de error en la base de datos del Ticket durante el procesamiento del mensaje de la QUEUE
        /// </summary>
        /// <param name="pnNumeroError">Número de error establecido por la aplicación</param>
        /// <param name="psDescripcion"> Descripción del error</param>
        /// <param name="psAplicacion">Aplicación en la que se presentó el error.</param>
        /// <param name="psFuncion">Función en la que se presentó el error.</param>
        public void psInsertarSQL(int pnNumeroError, string psDescripcion, string psAplicacion, string psFuncion)
        {
            string strQuery;
            int resultado;
            try
            {
                DateTime dateTime = DateTime.Now;

                strQuery = "Insert into BITACORA_ERRORES_MENSAJES_PU ";
                strQuery += "(fecha_hora, error_numero, error_descripcion, aplicacion) ";
                strQuery += "Values ('" + dateTime.ToString("yyyy-MM-dd HHMM") + "', " + pnNumeroError + ", '" + psDescripcion + "', '" + psAplicacion + "')";

                resultado = InsertaQuery(strQuery);
                return;
            }
            catch (Exception ex)
            {
                funcion.Escribe("Se presentó un error en la función: psInsertaSQL; " + ex.Message + " - " + ex.Data, "Mensaje");
                return;
            }
        }

        /// <summary>
        /// Inserta fila en BITACORA_ERRORES_MENSAJES_PU
        /// </summary>
        /// <param name="bitacora"></param>
        /// <returns></returns>
        public string psInsertarSQL(Bitacora_Errores_Mensajes_Pu bitacora)
        {
            strQuery = "Insert into BITACORA_ERRORES_MENSAJES_PU ";
            strQuery += "(fecha_hora, error_numero, error_descripcion, aplicacion) ";
            strQuery += $"Values ('{bitacora.fecha_hora.ToString("yyyy-MM-dd hh:mm:ss")}', {bitacora.error_numero}, '{bitacora.error_descripcion}', '{bitacora.aplicacion}')";

            return $"Se afectaron {ejecutarInsert(strQuery)} fila(s)";
        }

        public int InsertaQuery(string lsQuery)
        {
            int resultado;
            try
            {
                resultado = Insertar(lsQuery);
                return resultado;
            }
            catch (Exception ex)
            {
                funcion.Escribe("Se presentó un error durante la conexión a la base de datos: ", "Mensaje");
                return 0;
            }

        }

        public int Insertar(string sqlquery)
        {
            SqlCommand comando;
            int i = 0;
            try
            {
                //Connection.cnnConexion.Open(); //  [SQL]
                //comando = new SqlCommand(sqlquery, Connection.cnnConexion);
                //comando.CommandTimeout = 2480;
                //i = Math.Abs(comando.ExecuteNonQuery());
                if (i == 0)
                {
                    return 0;
                }
                else
                {
                    return i;
                }

            }
            catch (Exception ex)
            {
                funcion.Escribe("Error al ejecutar el query de inserción: " + sqlquery + " - " + ex.Message + " - " + ex.Data, "Error");
                return 0;
            }

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
                funcion.Escribe(ex);
                return -1;
            }
        }

        /// <summary>
        /// Obtiene un datatable con la informacion del query
        /// </summary>
        /// <param name="query"></param>
        /// <returns></returns>
        public DataTable ConsultaMQQUEUEFunc(string query)
        {
            try
            {
                SqlDataReader dr = this.ejecutarConsulta(query);

                DataTable dt = new DataTable();

                dt.Clear();

                dt.Columns.Add("id_funcionario");
                dt.Columns.Add("centro_regional");
                dt.Columns.Add("numero_funcionario");
                dt.Columns.Add("producto");
                dt.Columns.Add("subproducto");
                dt.Columns.Add("fecha_alta");
                dt.Columns.Add("fecha_baja");
                dt.Columns.Add("fecha_ultimo_mant");
                dt.Columns.Add("tipo_peticion");
                dt.Columns.Add("status_envio");
                dt.Columns.Add("columna_11");
                dt.Columns.Add("columna_12");
                dt.Columns.Add("id_transaccion");
                dt.Columns.Add("tipo");


                if (dr != null)
                {
                    while (dr.Read())
                    {
                        DataRow _row = dt.NewRow();

                        _row["id_funcionario"] = dr.GetInt32(0);
                        _row["centro_regional"] = dr.GetString(1);
                        _row["numero_funcionario"] = dr.GetString(2);
                        _row["producto"] = dr.GetString(3);
                        _row["subproducto"] = dr.GetString(4);
                        _row["fecha_alta"] = dr.GetString(5);
                        _row["fecha_baja"] = dr.GetString(6);
                        _row["fecha_ultimo_mant"] = dr.GetString(7);
                        _row["tipo_peticion"] = dr.GetString(8);
                        _row["status_envio"] = dr.GetByte(9);
                        _row["columna_11"] = dr.GetString(10);
                        _row["columna_12"] = dr.GetString(11);
                        _row["id_transaccion"] = dr.GetInt32(12);
                        _row["tipo"] = dr.GetString(13);

                        dt.Rows.Add(_row);
                    }
                    dr.Close();
                    return dt;
                }
            }
            catch (SqlException ex)
            {
                funcion.Escribe(ex, "Error");
            }
            catch (Exception ex)
            {
                funcion.Escribe(ex, "Error");
            }
            return null;
        }


        /// <summary>
        /// Ejecutar una consulta con un query dado
        /// </summary>
        /// <param name="query">query select</param>
        /// <returns></returns>
        private SqlDataReader ejecutarConsulta(string query)
        {
            try
            {
                cnnConexion.ActiveConnection = true;
                cnnConexion.ParametersContains = false;
                cnnConexion.CommandType = CommandType.Text;
                cnnConexion.ActiveConnection = true;

                SqlDataReader sqlRecord = cnnConexion.ExecuteDataReader(query);

                return sqlRecord;
            }
            catch (Exception ex)
            {
                return null;
                throw;
            }
        }


        public DataTable ConsultaMQQUEUEAuto(string query)
        {
            try
            {
                SqlDataReader dr = this.ejecutarConsulta(query);

                DataTable dt = new DataTable();
                dt.Clear();

                dt.Columns.Add("operacion");
                dt.Columns.Add("oficina");
                dt.Columns.Add("numero_funcionario");
                dt.Columns.Add("id_transaccion");
                dt.Columns.Add("codigo_operacion");
                dt.Columns.Add("cuenta");
                dt.Columns.Add("divisa");
                dt.Columns.Add("importe");
                dt.Columns.Add("fecha_operacion");
                dt.Columns.Add("folio_autorizacion");
                dt.Columns.Add("status_envio");
                dt.Columns.Add("fecha");
                dt.Columns.Add("hora");


                if (dr != null)
                {
                    while (dr.Read())
                    {
                        DataRow _row = dt.NewRow();

                        _row["operacion"] = dr.GetInt32(0);
                        _row["oficina"] = dr.GetString(1);
                        _row["numero_funcionario"] = dr.GetString(2);
                        _row["id_transaccion"] = dr.GetString(3);
                        _row["codigo_operacion"] = dr.GetString(4);
                        _row["cuenta"] = dr.GetString(5);
                        _row["divisa"] = dr.GetString(6);
                        _row["importe"] = dr.GetString(7);
                        _row["fecha_operacion"] = dr.GetString(8);
                        _row["folio_autorizacion"] = dr.GetString(9);
                        _row["status_envio"] = dr.GetByte(10);
                        _row["fecha"] = dr.GetString(11);
                        _row["hora"] = dr.GetString(12);

                        dt.Rows.Add(_row);
                    }
                    return dt;
                }
            }
            catch (SqlException ex)
            {
                funcion.Escribe(ex, "Error");
            }
            catch (Exception ex)
            {
                funcion.Escribe(ex, "Error");
            }
            return null;
        }

        public void Desconectar()
        {
            //mQ.cnnConexion.Close();
            //Set rssRegistro = Nothing
            this.cnnConexion = null;
        }

    }
}
