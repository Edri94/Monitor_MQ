using ConexionBDSQL;
using ServicioMonitor.Data;
using ServicioMonitor.Helpers;
using ServicioMonitor.Models;
using ServicioMonitor.Mq;
using System;
using System.Data;
using System.Data.SqlClient;

namespace ServicioMonitor.Processes
{
    public class Tkt
    {
        private string Archivo;
        private string ArchivoIni;
        private string Ls_Archivo;
        private string lsCommandLine;

        public string Ms_MQMngL;         //'MQManager de Lectura
        public string Ms_MQLeer;         //'MQQueue   de Lectura
        public string Ms_MQMngE;         //'MQManager de Escritura
        public string Ms_MQEscr;         //'MQQueue   de Escritura
        public string Gs_MsgRes;         //'Respuesta al mensaje procesado
        private string Ms_BanRetorno;     //'BANDERA QUE NOS INDICA SI SE VA A REGRESAR UNA RESPUESTA
        private Boolean mbFuncionBloque;  //
        public string Bandera;            //'Indicador de encripción de constantes de conexión a SQL Server (1=encriptado; 0=no encriptado)

        //Variables para BD
        public static string CadenaConexion;
        public static ConexionBD sqlConexionUtil;

        // Variables para el control del log
        private string strlogFileName;
        private string strlogFilePath;
        private bool Mb_GrabaLog;

        public Autorizacion laAutoriz;


        // Variables para el registro de los valores del header IH
        private string strFuncionHost; // Valor que indica el programa que invocara el CICSBRIDGE
        private string strHeaderTagIni; // Bandera que indica el comienzo del Header
        private string strIDProtocol; // Identificador  del protocolo (PS9)
        private string strLogical; // Terminal Lógico Asigna Arquitectura ASTA
        private string strAccount; // Terminal Contable (CR Contable)
        private string strUser; // Usuario. Debe ser diferente de espacios
        private string strSeqNumber; // Número de Secuencia (indicador de paginación)
        private string strTXCode; // Función específica Asigna Arquitectura Central
        private string strUserOption; // Tecla de función (no aplica)
        private string strCommit; // Indicador de commit: Permite realizar commit
        private string strMsgType; // Tipo de mensaje: Nuevo requerimiento
        private string strProcessType; // Tipo de proceso: on line
        private string strChannel; // Canal Asigna Arquitectura Central
        private string strPreFormat; // Indicador de preformateo: Arquitectura no deberá de preformatear los datos
        private string strLenguage; // Idioma: Español
        private string strHeaderTagEnd; // Bandera que indica el final del header

        // Variables para el registro de los valores del header ME
        private string strMETAGINI; // Bandera que indica el comienzo del mensaje
                                    // Private strMsgColecMax          As String 'Longitud del layout  del colector
        private string strMsgTypeCole; // Tipo de mensaje: Copy
                                       // Private strMaxMsgCole           As String 'Máximo X(30641)
        private string strMETAGEND; // Bandera que indica el fin del mensaje

        // Variables para el registro de los valores Default
        private string strFechaBaja; // fecha_baja
        private string strColectorMaxLeng; // Maxima longitud del COLECTOR
        private string strMsgMaxLeng; // Maxima longitud del del bloque ME
        private string strPS9MaxLeng; // Maxima longitud del formato PS9
        private string strReplyToMQ; // MQueue de respuesta para HOST
        private string strFuncionSQL; // Funcion a ejecutar al recibir la respuesta
        private string strRndLogTerm; // Indica que el atributo Logical Terminal es random

        // Variables para el manejo de los parametros de la base de datos
        // Public gsSeccRegWdw             As String

        // VARIABLES NUVAS PARA EL ENVIO DE MENSAJE
        private string sPersistencia;
        private string sExpirar;

        private string Gs_MQManager;       // MQManager de Escritura
        private string Gs_MQQueueEscritura;       // MQQueue de Escritura
        private string gsEjecutable;       // Ejecutable a realizar


        //MqSeriesConfig mqSeriesConfig;
        //EscribeArchivoLOGConfig escribeArchivoLOGConfig;
        //ConexionConfig conexionConfig;

        TktMq mqSeries;
        Funcion_Tkt funcion;
        Encriptacion encriptacion;
        TktBd bd;
        public bool Mb_Detalles;

        public Tkt()
        {
            mqSeries = new TktMq();
            funcion = new Funcion_Tkt();
            encriptacion = new Encriptacion();
            bd = new TktBd();
        }


        public bool fValidaLayout(string psFuncion, int piLenMensaje, string psMensaje)
        {
            //            string ls_msg = "";
            string sParam;
            int iLnParam = 0;

            try
            {
                funcion.Escribe("=================================", "Mensaje");
                funcion.Escribe("Entra a la funcion fValidaLayout.", "Mensaje");
                funcion.Escribe("=================================", "Mensaje");
                funcion.Escribe("declara variables y abre conexion", "Mensaje");
                DataSet dsRegistro = new DataSet();
                SqlParameter pFuncion = sqlConexionUtil.CreateParameters("@psFuncion", SqlDbType.NVarChar, ParameterDirection.Input, psFuncion);
                SqlCommand command = new SqlCommand();
                SqlConnection sql_conn = new SqlConnection(sqlConexionUtil.ConnectionString);
                sql_conn.Open();
                funcion.Escribe("sql_conn.Open()  correcto.", "Mensaje");

                string lsQuery = "SELECT SUM(B.longitud)";
                //lsQuery += "FROM " + gsNameDB + "..TIPO_TRANSACCION_PIU A, " + gsNameDB + "..ENTRADA_PIU B ";
                lsQuery += "FROM TICKET..TIPO_TRANSACCION_PIU A, TICKET..ENTRADA_PIU B ";
                lsQuery += "WHERE A.tipo_transaccion = B.tipo_transaccion ";
                lsQuery += "AND A.funcion = '" + psFuncion.Trim() + "' ";

                funcion.Escribe("Query a ejecutar" + lsQuery, "Mensaje");
                funcion.Escribe("=================================", "Mensaje");
                funcion.Escribe("declara parametros para ejecución del query", "Mensaje");

                command.CommandType = CommandType.Text;
                command.CommandText = lsQuery;
                command.Parameters.Add(pFuncion);
                command.Connection = sql_conn;
                SqlDataAdapter dataadapter = new SqlDataAdapter(command);
                dataadapter.Fill(dsRegistro);

                funcion.Escribe("Ejecuta query, se asigna en dsRegistro.", "Mensaje");
                funcion.Escribe("=================================", "Mensaje");
                funcion.Escribe("declara datatable y le asigna el contenido de dsRegistro", "Mensaje");

                DataTable Tabla = new DataTable();
                Tabla = dsRegistro.Tables[0];

                funcion.Escribe("=================================", "Mensaje");
                funcion.Escribe("cierra la conexion", "Mensaje");
                sql_conn.Close();

                if ((dsRegistro.Tables[0].Rows[0] == null) || (dsRegistro.Tables[0].Rows[0].ToString() == "") || (Convert.ToInt32(dsRegistro.Tables[0].Rows[0]) == 0))
                {
                    funcion.Escribe("Error no es posible extraer la longitud de los paramtros del Store Pocedure, funcion fValidaLayout.", "Error");
                }
                else
                {
                    iLnParam = Convert.ToInt32(dsRegistro.Tables[0].Rows[0]);
                }
                funcion.Escribe("=================================", "Mensaje");
                funcion.Escribe("limpia el objeto dsRegistro y libera memoria", "Mensaje");

                dsRegistro.Clear();
                dsRegistro = null;
                sParam = psMensaje.Substring(1, psMensaje.Length);

                funcion.Escribe("=================================", "Mensaje");
                funcion.Escribe("sParam: " + sParam, "Mensaje");

                if (iLnParam > sParam.Length)
                {
                    return false;
                }

                return true;
            }
            catch (Exception ex)
            {
                funcion.Escribe("Error al validar el layout, funcion fValidaLayout: '" + ex.Message + "'", "Error");
                throw new Exception(ex.Message);
            }
            //return true;
        }


        // ***************************************        
        // ***************************************
        // ***************************************

        public void ProcesarMensajes(string strRutaIni, string strParametros)
        {
            //string strFuncion;
            //int Li_TotReg;
            //int intLongBia;
            bool OpNoCompletada = false;
            string lsCommandLine;
            string[] Parametros = new string[4];        //' Arreglo para almacenar los parametros via línea de comando
            string Ls_MsgVal = "";       // Mensaje con el resultado de la validación

            //float LnDiferencia;       // Minutos transcurridos desde el último intento de acceso


            try
            {
                mbFuncionBloque = false;

                lsCommandLine = strParametros.Trim();

                if (lsCommandLine.Equals("") == false)
                {
                    Parametros = lsCommandLine.Split('-');
                    Ms_MQMngL = Parametros[0].Trim();
                    Ms_MQMngE = Parametros[0].Trim();
                    Ms_MQLeer = Parametros[1].Trim();
                    Ms_MQEscr = Parametros[2].Trim();
                    Ms_BanRetorno = Parametros[3].Trim();
                }
                else
                {
                    ObtenerInfoMQ();
                }

                funcion.Escribe("Comienza la función MAIN de la aplicación MensajesMQ: " + DateTime.Now.ToString() + " Tipo Función: '" + gsEjecutable + "'", "Mensaje");

                if (!ValidaInfoMQ(Ls_MsgVal))
                {
                    funcion.Escribe("Se presentó un error en la información de validación de configuración de la sección de MQSeries.", "Error");
                    funcion.Escribe("Termina el acceso a la aplicación TKTMQ.", "Error");
                    funcion.Escribe("", "Error");
                    return;       //Sale de la Función Principal ProcesarMensajes
                }

                //if (mqSeries.MQConectar(Ms_MQMngL, mqSeries.mqsManager))
                if (mqSeries.ConectarMQ(Ms_MQMngL))
                {
                    mqSeries.blnConectado = true;
                }
                else
                {
                    funcion.Escribe("Fallo conexión al servidor de MQ-Series del QManager ", "Mensaje");
                    return;       //Sale de la Función Principal ProcesarMensajes
                }

                long lngMQOpen = (long)MqSeries.MQOPEN.MQOO_INQUIRE;
                funcion.Escribe(" ---> llamando a funcion: mqSeries.MQRevisaQueue", "Mensaje");

                //if (mqSeries.MQRevisaQueue(mqSeries.mqsManager, Ms_MQLeer, (MqSeries.MQOPEN)lngMQOpen) == 0)
                if (mqSeries.RevisaQueueMq(Ms_MQLeer, (MqSeries.MQOPEN)lngMQOpen) == 0)
                {
                    funcion.Escribe(" ---> llamando a funcion: mqSeries.MQDesconectar", "Mensaje");
                    //mqSeries.MQDesconectar(ref mqSeries.mqsManager, ref mqSeries.mqsEscribir, ref mqSeries.mqsLectura);
                    mqSeries.DesconectarMQ();
                    funcion.Escribe("Termina TKTMQ Recepción de Solicitudes. Al revisar la QUEUE no encontramos mensajes ", "Mensaje");
                    funcion.Escribe("", "Mensaje");
                    lngMQOpen = 0;
                    return;       //Sale de la Función Principal ProcesarMensajes
                }

                funcion.Escribe(" ---> llamando a funcion: ConectDB", "Mensaje");
                if (!bd.ConectDB())
                {
                    funcion.Escribe("Error al conectar con la BD ", "Mensaje");
                    funcion.Escribe("", "Mensaje");
                    return;       //Sale de la Función Principal ProcesarMensajes
                }

                //Recupera el mensaje de la queue
                string strReturn = "";
                //if (mqSeries.MQRecibir(mqSeries.mqsManager, Ms_MQLeer, mqSeries.mqsLectura, mqSeries.mqsMsglectura, (MqSeries.MQOPEN)lngMQOpen, ref strReturn))
                funcion.Escribe(" ---> llamando a funcion: mqSeries.MQRecibir", "Mensaje");
                if (mqSeries.RecibirMq(Ms_MQLeer))
                {
                    funcion.Escribe(" ---> llamando a funcion: mqSeries.MQDesconectar", "Mensaje");
                    //mqSeries.MQDesconectar(ref mqSeries.mqsManager, ref mqSeries.mqsEscribir, ref mqSeries.mqsLectura);
                    mqSeries.DesconectarMQ();
                    if (Mb_Detalles)
                    {
                        funcion.Escribe("Mensaje recibido correctamente", "Mensaje");
                    }
                }
                else
                {
                    mqSeries.DesconectarMQ();
                    if (Mb_Detalles)
                    {

                        bd.psInsertarSQL(new Bitacora_Errores_Mensajes_Pu { error_numero = 8, error_descripcion = "No se realizo la recepción del mensaje", aplicacion = "TKT" });
                        funcion.Escribe("No se realizo la recepción del mensaje", "Error");
                    }
                    funcion.Escribe("", "Mensaje");
                    mqSeries.DesconectarMQ();
                    return;       //Sale de la Función Principal ProcesarMensajes
                }

                funcion.Escribe("Mensaje recuperado de la Queue: " + strReturn, "Mensaje");
                Gs_MsgRes = "";

                funcion.Escribe("", "Mensaje");
                funcion.Escribe("Llamando a la función: mqSeries.ProcesMessage", "Mensaje");
                //Procesa el mensaje                 //if (!mqSeries.ProcesMessage(ref mqSeries.mqsMsglectura))
                if (!ProcesMessage(strReturn))
                {
                    string respuesta = strReturn;
                    OpNoCompletada = true;
                }

                funcion.Escribe("1 Valida si el Mensaje fue completado:", "Mensaje");
                if (!OpNoCompletada)
                {
                    funcion.Escribe("el Mensaje SI fue completado:", "Mensaje");
                    if (Ms_BanRetorno == "0" || mbFuncionBloque)
                    {
                        funcion.Escribe("Termina el acceso a la aplicación MensajesMQ. Función SQL: " + strFuncionSQL, "Mensaje");
                        funcion.Escribe("", "Mensaje");
                        mqSeries.DesconectarMQ();
                        return;       //Sale de la Función Principal ProcesarMensajes
                    }

                    //if (mqSeries.MQConectar(Ms_MQMngL, mqSeries.mqsManager))
                    if (mqSeries.ConectarMQ(Ms_MQMngL))
                    {
                        mqSeries.blnConectado = true;
                        if (mqSeries.MQEnviar(Ms_MQEscr, Gs_MsgRes))
                        {
                            funcion.Escribe("", "Mensaje");
                            mqSeries.DesconectarMQ();
                        }
                        else
                        {
                            mqSeries.DesconectarMQ();
                            bd.psInsertarSQL(new Bitacora_Errores_Mensajes_Pu { error_numero = 5, error_descripcion = "No se pudo enviar el mensaje", aplicacion = "TKT" });
                            OpNoCompletada = true;
                        }
                    }
                    else
                    {
                        bd.psInsertarSQL(new Bitacora_Errores_Mensajes_Pu { error_numero = 3, error_descripcion = "Error al abrir la MQ Queue de respuesta", aplicacion = "TKT" });
                        OpNoCompletada = true;
                    }

                    funcion.Escribe("2 Valida si el Mensaje fue completado:", "Mensaje");
                    if (!OpNoCompletada)
                    {
                        funcion.Escribe("2 el Mensaje NO fue completado:", "Mensaje");
                        mqSeries.DesconectarMQ();
                        //se desconecta de la BD
                        Desconectar();
                        funcion.Escribe("Mensaje Regresado correctamente. Termina TKTMQ Recepción de Solicitudes.", "Mensaje");
                        funcion.Escribe("", "Mensaje");
                        return;       //Sale de la Función Principal ProcesarMensajes
                    }
                    else
                    {
                        funcion.Escribe("2 el Mensaje NO fue completado:", "Mensaje");
                        if (Gs_MsgRes == "ENCABEZADO")
                        {
                            funcion.Escribe("Termina TKTMQ Recepción de Solicitudes.", "Mensaje");
                            funcion.Escribe("", "Mensaje");
                            return;   //Sale de la Función Principal ProcesarMensajes
                        }
                        if (!Mb_Detalles)
                        {
                            funcion.Escribe("El procesamiento del mensaje no pudo ser completado correctamente, para ver mas detalles active Mb_Detalles = 1 en el archivo appConfig.", "Mensaje");
                            funcion.Escribe("", "Mensaje");
                        }
                        return;     //Sale de la Función Principal ProcesarMensajes
                    }

                }
                else
                {
                    funcion.Escribe("el Mensaje NO fue completado:", "Mensaje");
                    if (Gs_MsgRes == "ENCABEZADO")
                    {
                        funcion.Escribe("Termina TKTMQ Recepción de Solicitudes.", "Mensaje");
                        funcion.Escribe("", "Mensaje");
                        return;       //Sale de la Función Principal ProcesarMensajes
                    }
                    if (!Mb_Detalles)
                    {
                        funcion.Escribe("El procesamiento del mensaje no pudo ser completado correctamente, para ver mas detalles active Mb_Detalles = 1 en el archivo appConfig.", "Mensaje");
                        funcion.Escribe("", "Mensaje");
                    }
                    return;       //Sale de la Función Principal ProcesarMensajes
                }
            }
            catch (Exception ex)
            {
                funcion.Escribe("Termina el acceso a la aplicación TKTMQ porque se presentó un error en la función MAIN. Error. " + ex.Message + " - " + ex.Data, "Error");
                funcion.Escribe("", "Error");
                //se desconecta de la BD
                Desconectar();
            }

        }


        private void ObtenerInfoMQ()
        {
            string section = "mqSeries";
            Gs_MQManager = funcion.getValueAppConfig("MQManager", section);
            Gs_MQQueueEscritura = funcion.getValueAppConfig("MQEscritura", section);
            gsEjecutable = funcion.getValueAppConfig("FGEjecutable", section);

        }

        private void ConfiguraFileLog()
        {
            string section = "escribeArchivoLOG";

            strlogFileName = funcion.getValueAppConfig("logFileName", section);
            strlogFilePath = funcion.getValueAppConfig("logFilePath", section);
            Mb_GrabaLog = Boolean.Parse(funcion.getValueAppConfig("Estatus", section));
            Mb_Detalles = Boolean.Parse(funcion.getValueAppConfig("Detalles", section));
        }





        private bool ValidaInfoMQ(string ps_MsgVal)
        {
            string ls_msg = "";

            if (Ms_MQMngL.Trim() == "")
            {
                ls_msg = ls_msg + "";
            }
            if (Ms_MQMngE.Trim() == "")
            {
                ls_msg = ls_msg + "";
            }
            if (Ms_MQLeer.Trim() == "")
            {
                ls_msg = ls_msg + "";
            }
            if (Ms_MQEscr.Trim() == "")
            {
                ls_msg = ls_msg + "";
            }

            if (ls_msg == "")
            {
                return true;
            }
            ps_MsgVal = ls_msg;
            return false;
        }

        public bool ProcesMessage(string MsgLeido)
        {
            string ls_mensaje;
            string lsQuery;
            string lsDatos;
            string lsNewMsg;
            string Ls_Servicio;
            string psBodyMsg;
            string psLongitud;
            string psTipo;
            string sBodyMsg;
            string sAux;
            //string lsMsgErr;
            bool sMensajeBloque;
            //double Ld_CodigoExecNTHOST;
            //string EjecutableMSG;
            //variable contador de emensaje
            int iMensaje;
            bool lbYaProcesado;
            int ipiLenMensaje;
            bool bNEXTOC;

            //ProcesMessage = False
            sMensajeBloque = false;
            //se maneja como variable privada y en falso
            mbFuncionBloque = false;
            psBodyMsg = "";
            psLongitud = "";
            psTipo = "";
            lsNewMsg = "";
            ipiLenMensaje = 0;
            bNEXTOC = false;
            lsQuery = "";
            funcion.Escribe("======================================", "Mensaje");
            funcion.Escribe("Entro a la función: ProcesMessage", "Mensaje");
            //Obtenemos el mensaje contenido en el objeto Mensaje
            //ls_mensaje = objMsgLeido.MessageData;
            ls_mensaje = MsgLeido;
            funcion.Escribe("Mensaje recibido: " + ls_mensaje, "Mensaje");
            //'Esta línea revisa si el mensaje que solo es el encabezado y no contaban con cuerpo. El procesamiento marcaba un error que era escrito en el archivo .log
            // InStr(1, ls_mensaje, "<OC>", vbTextCompare) = 0
            //busca la cadena "<OC>" en ls_mensaje.
            //asigna la posición en la que se encuentra a iInicia, si no existe retorna -1
            //int iIniciaOC = ls_mensaje.IndexOf("<OC>") + "<OC>".Length;
            int iIniciaOC = ls_mensaje.IndexOf("<OC>");

            //if ( string.Compare(ls_mensaje, "<OC>") == 0 )
            if (iIniciaOC == -1)
            {
                Gs_MsgRes = "ENCABEZADO";
                return false;
            }
            //iIniciaOC += "</OC>".Length;
            int iFinOC = ls_mensaje.IndexOf("</OC>") + "</OC>".Length;

            funcion.Escribe("--> iIniciaOC: " + iIniciaOC, "Mensaje");
            funcion.Escribe("--> iFinOC: " + iFinOC, "Mensaje");
            //sBodyMsg = ls_mensaje.Substring(string.Compare(ls_mensaje, "<OC>"), ls_mensaje.Length);
            sBodyMsg = ls_mensaje.Substring(iIniciaOC, iFinOC - iIniciaOC);
            //sBodyMsg = Mid(ls_mensaje, InStr(1, ls_mensaje, "<OC>", vbTextCompare), Len(ls_mensaje))
            iMensaje = 0;
            funcion.Escribe("sBodyMsg:" + sBodyMsg, "Mensaje");
            funcion.Escribe("Entra al DO para procesar contenido del sBodyMsg", "Mensaje");
            do
            {
                lbYaProcesado = false;
                //sAux = sBodyMsg.Substring(1, string.Compare(sBodyMsg, "<OC>") + 4);
                sAux = sBodyMsg;
                funcion.Escribe("sAux: " + sAux, "Mensaje");

                funcion.Escribe("Llamando a la función: psMsgAStructPS9", "Mensaje");
                //'Descompone el mensaje en una estructura PS9
                psMsgAStructPS9(sAux, psBodyMsg, ref psLongitud, psTipo);

                funcion.Escribe("SE REGRESO psLongitud: " + psLongitud, "Mensaje");

                //laAutoriz.FuncionSQL = psBodyMsg.Substring(1, 8);
                funcion.Escribe("psLongitud.Substring(0, 8): " + psLongitud.Substring(0, 8), "Mensaje");
                laAutoriz.FuncionSQL = psLongitud.Substring(0, 8);
                funcion.Escribe("laAutoriz.FuncionSQL: " + laAutoriz.FuncionSQL, "Mensaje");

                funcion.Escribe("--> Llamando a la función: 1 Fl_ObtieneServicio", "Mensaje");
                Ls_Servicio = Fl_ObtieneServicio(laAutoriz.FuncionSQL, "SinProcesar", "EntraLog0", 1);

                funcion.Escribe("<-- Saliendo de la función: 1 Fl_ObtieneServicio", "Mensaje");

                funcion.Escribe("1 Ls_Servicio recuperado: " + Ls_Servicio, "Mensaje");

                funcion.Escribe("Validando longitud de Ls_Servicio:" + Ls_Servicio.Length, "Mensaje");
                if (Ls_Servicio.Length != 0)
                {
                    funcion.Escribe("Entro al IF Validando longitud de Ls_Servicio.", "Mensaje");
                    lsDatos = laAutoriz.DatosTemp;
                    funcion.Escribe("lsDatos: " + lsDatos, "Mensaje");

                    funcion.Escribe("---> Llamando a la función: ReArmaPS9", "Mensaje");
                    ReArmaPS9(lsDatos, ls_mensaje, psBodyMsg, ref lsNewMsg, psTipo);
                    funcion.Escribe("<--- Saliendo de la función: ReArmaPS9", "Mensaje");

                    Gs_MsgRes = lsNewMsg;
                    funcion.Escribe("lsNewMsg: " + lsNewMsg, "Mensaje");
                    funcion.Escribe("Gs_MsgRes: " + Gs_MsgRes, "Mensaje");
                    funcion.Escribe("La funcion pertenece al bloque de pruebas (no procesar): " + laAutoriz.FuncionSQL, "Mensaje");
                    return false;   //sale de la función como FALSO
                }

                funcion.Escribe("--> Llamando a la función: 2 Fl_ObtieneServicio", "Mensaje");
                Ls_Servicio = Fl_ObtieneServicio(laAutoriz.FuncionSQL, "Servicios", "Servicio3", 2);
                funcion.Escribe("<-- Saliendo de la función: 2 Fl_ObtieneServicio", "Mensaje");

                funcion.Escribe("2 Ls_Servicio: " + Ls_Servicio, "Mensaje");
                if (Ls_Servicio.Length == 0)
                {
                    Ls_Servicio = Fl_ObtieneServicio(laAutoriz.FuncionSQL, "FuncionBloques", "Servicio0", 2);
                    funcion.Escribe("2a Ls_Servicio: " + Ls_Servicio, "Mensaje");
                    if (Ls_Servicio.Length != 0)
                    {
                        sMensajeBloque = true;
                        laAutoriz.Espera = psBodyMsg.Substring(9, 1);
                        laAutoriz.FechaProce = psBodyMsg.Substring(10, 10);
                        laAutoriz.HoraProce = psBodyMsg.Substring(21, 8);
                        laAutoriz.DatosTemp = psBodyMsg;
                    }
                    else
                    {
                        funcion.Escribe("La funcion " + laAutoriz.FuncionSQL + " no existe.", "Mensaje");
                    }
                }
                else
                {
                    sMensajeBloque = false;
                    laAutoriz.FechaProce = psBodyMsg.Substring(9, 8);
                    laAutoriz.HoraProce = psBodyMsg.Substring(17, 8);
                    laAutoriz.DatosTemp = psBodyMsg;
                }

                //if (!fValidaLayout(laAutoriz.FuncionSQL, IIf(sMensajeBloque = True, Len(psBodyMsg) - 8, (psLongitud)), psBodyMsg))
                if (sMensajeBloque)
                {
                    ipiLenMensaje = psBodyMsg.Length - 8;
                }
                else
                {
                    ipiLenMensaje = Convert.ToInt32(psLongitud);
                    //ipiLenMensaje = 0;
                }
                funcion.Escribe("Valida layout: tktMQ.fValidaLayout", "Mensaje");
                if (!fValidaLayout(laAutoriz.FuncionSQL, ipiLenMensaje, psBodyMsg))
                {
                    //int Resp = tktMQ.psInsertaSQL(4, "La longitud de los parametros en el mansaje, no es compatible con la longuitud de los parametros para generar la ejecución del sp", "TKT", "ProcesMessage");
                    bd.psInsertarSQL(new Bitacora_Errores_Mensajes_Pu { error_numero = 4, error_descripcion = "La longitud de los parametros en el mansaje, no es compatible con la longuitud de los parametros para generar la ejecución del sp", aplicacion = "TKT" });
                    bNEXTOC = true;
                }

                if (!bNEXTOC)   //Si entra, es porque la variable esta en falso.
                {
                    if (sMensajeBloque)
                    {
                        ipiLenMensaje = psBodyMsg.Length - 8;
                    }
                    else
                    {
                        ipiLenMensaje = Convert.ToInt32(psLongitud);
                        //ipiLenMensaje = 0;
                    }

                    if (!BuscaSPyLongPar(laAutoriz.FuncionSQL, ipiLenMensaje, psBodyMsg, ref lsQuery))
                    {
                        ReArmaPS9("TKT1010 FUNCION NO EXISTENTE", ls_mensaje, psBodyMsg, ref lsNewMsg, psTipo);
                        bNEXTOC = true;
                    }

                    if (!bNEXTOC)   //Si entra, es porque la variable esta en falso.
                    {
                        lsDatos = "";
                        if (!EjecSPFuncion(lsQuery, ref lsDatos))
                        {
                            ReArmaPS9(lsDatos, ls_mensaje, psBodyMsg, ref lsNewMsg, psTipo);
                            lbYaProcesado = true;
                            bNEXTOC = true;
                        }
                        if (!bNEXTOC)   //Si entra, es porque la variable esta en falso.
                        {
                            if (Mb_Detalles)
                            {
                                if (lsDatos != null || lsDatos != "")
                                {
                                    funcion.Escribe("Resultado de la ejecución del SP: " + lsDatos, "Mensaje");
                                }
                                else
                                {
                                    funcion.Escribe("El SP no regresa datos.", "Mensaje");
                                }
                            }
                        }
                    }

                }

                sBodyMsg = sBodyMsg.Substring((sAux.Length + 1), (sBodyMsg.Length));
                iMensaje += 1;

            } while (sBodyMsg != "");


            if (sMensajeBloque)
            {
                funcion.Escribe("Total de mensajes <OC> obtenidos :" + iMensaje.ToString(), "Mensaje");
            }

            if (!lbYaProcesado)
            {
                lsDatos = "";
                //'Rearma el PS9 con los datos de la respuesta (Proceso Completo)
                ReArmaPS9(lsDatos, ls_mensaje, psBodyMsg, ref lsNewMsg, psTipo);
            }

            if (sMensajeBloque)
            {
                if (laAutoriz.Espera == "1")
                {
                    //TktMQ mQp = new TktMQ();
                    Bitacora objBitacoras = new Bitacora();
                    try
                    {
                        //string strParams = mQp.Ms_MQMngE + "-" + mQp.Ms_MQEscr + "-" + "1" + "-" + laAutoriz.FuncionSQL;
                        //objBitacoras.ProcesarBitacora(gstrRutaIni, strParams);
                        //objBitacoras = null;
                    }
                    catch (Exception ex)
                    {
                        funcion.Escribe("Error al ejecutar funcion ProcesMessage - objBitacoras.ProcesarBitacora : " + ex.Message + " - " + ex.Data, "Mensaje");
                        return false;
                    }
                }
            }

            Gs_MsgRes = lsNewMsg;
            return true;
        }

        private void psMsgAStructPS9(string psMsg, string psBodyMsg, ref string psLongitud, string psTipo)
        {
            int iIni;
            int iFin;
            int iExisteHE;
            int iExisteOC;
            string sEncabezado;

            funcion.Escribe("Entro a funcion: psMsgAStructPS9", "Mensaje");

            iExisteHE = psMsg.IndexOf("<ME>");
            iExisteOC = psMsg.IndexOf("<OC>");
            funcion.Escribe("--> iExisteHE:" + iExisteHE, "Mensaje");
            funcion.Escribe("--> iExisteOC:" + iExisteOC, "Mensaje");

            if (iExisteHE >= 0)
            {
                psTipo = "<ME>";
                funcion.Escribe("psTipo: " + psTipo, "Mensaje");
                iIni = psMsg.IndexOf("<ME>");
                iFin = psMsg.IndexOf("</ME>") + "</ME>".Length;
                funcion.Escribe("--> iIni:" + iIni, "Mensaje");
                funcion.Escribe("--> iFin:" + iFin, "Mensaje");
                sEncabezado = psMsg.Substring(iIni, 4);
                funcion.Escribe("--> sEncabezado:" + sEncabezado, "Mensaje");
                psLongitud = psMsg.Substring(iIni + sEncabezado.Length, 5);
                funcion.Escribe("--> psLongitud:" + psLongitud, "Mensaje");
                psBodyMsg = psMsg.Substring((iIni + sEncabezado.Length + psLongitud.Length), (iFin - (iIni + sEncabezado.Length + psLongitud.Length)));
                funcion.Escribe("--> psBodyMsg:" + psBodyMsg, "Mensaje");
                psLongitud = psLongitud.Substring(1, 4);
                funcion.Escribe("--> 2 psLongitud:" + psLongitud, "Mensaje");
            }
            else if (iExisteOC >= 0)
            {
                psTipo = "<OC>";
                funcion.Escribe("psTipo: " + psTipo, "Mensaje");
                iIni = psMsg.IndexOf("<OC>");
                iFin = psMsg.IndexOf("</OC>") + "</OC>".Length;
                funcion.Escribe("--> iIni:" + iIni, "Mensaje");
                funcion.Escribe("--> iFin:" + iFin, "Mensaje");
                sEncabezado = psMsg.Substring(iIni, 4);
                funcion.Escribe("--> sEncabezado:" + sEncabezado, "Mensaje");
                psLongitud = psMsg.Substring((iIni + sEncabezado.Length + 5 + 5), (psMsg.Length - 19));
                funcion.Escribe("--> psLongitud:" + psLongitud, "Mensaje");
                psBodyMsg = psMsg.Substring((iIni + sEncabezado.Length + 10), (iFin - (iIni + sEncabezado.Length + 10)));
                funcion.Escribe("--> psBodyMsg:" + psBodyMsg, "Mensaje");
            }
            funcion.Escribe(" ", "Mensaje");

        }

        public string Fl_ObtieneServicio(string Ls_Funcion, string psBloque, string psLlave, int arreglo)
        {
            int Ln_Ind;
            string Ls_Servicio;
            string sFlObtieneServicio;
            string sValorDo;
            string[] tContenidos;
            long i;

            i = 0;
            sFlObtieneServicio = "";
            string section = psBloque;
            string llave = psLlave;

            funcion.Escribe("Entro a la función: Fl_ObtieneServicio", "Mensaje");
            funcion.Escribe("--> arreglo: " + arreglo, "Mensaje");
            funcion.Escribe("--> section: " + section, "Mensaje");
            funcion.Escribe("--> llave: " + llave, "Mensaje");

            string sValor = funcion.getValueAppConfig(llave, section);
            int iValor = Convert.ToInt32(sValor);
            funcion.Escribe("--> sValor: " + sValor, "Mensaje");

            funcion.Escribe("", "Mensaje");
            if (sValor != null)
            {
                funcion.Escribe("--> Entro al IF", "Mensaje");
                do
                {
                    funcion.Escribe("--> llave" + llave + ": " + llave, "Mensaje");
                    i += 1;
                    llave = llave + i.ToString();
                    funcion.Escribe("--> llave: " + llave, "Mensaje");
                    sValorDo = funcion.getValueAppConfig(llave, section);
                    funcion.Escribe("--> sValorDo: " + sValorDo, "Mensaje");
                    tContenidos = sValorDo.Split(',');
                    if (arreglo == 1)
                    {
                        //Ln_Ind = string.Compare(tContenidos[0], ",");
                        Ls_Servicio = tContenidos[0];
                        sFlObtieneServicio = Ls_Servicio;
                        funcion.Escribe("--> sFlObtieneServicio: " + sFlObtieneServicio, "Mensaje");
                        return sFlObtieneServicio;
                    }
                    if (arreglo == 2)
                    {
                        Ln_Ind = string.Compare(tContenidos[1], Ls_Funcion);
                        funcion.Escribe("--> Ln_Ind: " + Ln_Ind, "Mensaje");
                        if (Ln_Ind > 0)
                        {
                            Ln_Ind = string.Compare(tContenidos[1], ",");
                            funcion.Escribe("---> Ln_Ind: " + Ln_Ind, "Mensaje");
                            Ls_Servicio = tContenidos[1];
                            sFlObtieneServicio = Ls_Servicio;
                            funcion.Escribe("--> sFlObtieneServicio: " + sFlObtieneServicio, "Mensaje");
                            return sFlObtieneServicio;
                        }
                    }

                } while (i < iValor);
            }
            return sFlObtieneServicio;
        }


        public string ReArmaPS9(string psDatos, string psMensaje, string psColector, ref string psNuevoMsg, string psTipo)
        {
            int LnLongColector;
            string Header;
            string sTMsg;
            string sLMsg;
            string sColector;
            //string sCero;
            string sFlujo;

            try
            {
                funcion.Escribe("", "Mensaje");
                funcion.Escribe("Entrando a la función: ReArmaPS9", "Mensaje");
                funcion.Escribe("psTipo: " + psTipo, "Mensaje");
                //'Determinamos la longitud del colector
                LnLongColector = psDatos.Length;
                sColector = "";
                if (psTipo == "<ME>")
                {
                    Header = psMensaje.Substring(1, string.Compare(psMensaje, psTipo) - 1) + "<HE>";
                    sLMsg = LnLongColector.ToString("0000");
                    sTMsg = psMensaje.Substring(string.Compare(psMensaje, psTipo) + 8, 1);
                    if (psDatos.Length == 0)
                    {
                        sColector = psDatos.PadRight(psDatos.Length - LnLongColector) + "</HE>";
                    }
                    else
                    {
                        sColector = "</HE>";
                    }
                    psNuevoMsg = Header + sLMsg + sTMsg + sColector;
                }
                else if (psTipo == "<OC>")
                {
                    Header = psMensaje.Substring(1, 4);
                    sFlujo = psMensaje.Substring(5, 5);
                    sLMsg = LnLongColector.ToString("0000");
                    sTMsg = psMensaje.Substring(14, 1);
                    if (psDatos.Length == 0)
                    {
                        sColector = psDatos.PadRight(psDatos.Length - LnLongColector) + "</OH>";
                    }
                    else
                    {
                        sColector = "</OH>";
                    }
                    psNuevoMsg = Header + sFlujo + sLMsg + sTMsg + sColector;
                }
                return psNuevoMsg;
            }
            catch (Exception ex)
            {
                funcion.Escribe("ocurrio un error en la funcion: ReArmaPS9, " + ex.Message + " - " + ex.Data, "Error");
                return psNuevoMsg;
            }

        }


        public bool BuscaSPyLongPar(string psFuncion, int pnLongPar, string psBia, ref string psQuery)
        {
            string lsQuery;
            string lsParam;
            int lnCont;
            string Ls_Servicio;
            string sSPejecutar;

            lsQuery = "";

            try
            {
                Ls_Servicio = Fl_ObtieneServicio(psFuncion.Trim(), "servicios", "Servicio3", 1);
                if (Ls_Servicio.Length == 0)
                {
                    Ls_Servicio = Fl_ObtieneServicio(psFuncion.Trim(), "FuncionBloques", "Servicio0", 2);
                    if (Ls_Servicio.Length == 0)
                    {
                        funcion.Escribe("La funcion: " + psFuncion + " NO existe.", "Mensaje");
                    }
                }
                else
                {
                    funcion.Escribe("Funcion a ejecutar: " + psFuncion, "Mensaje");
                }
            }
            catch (Exception ex)
            {
                funcion.Escribe("Error en la Funcion BuscaSPyLongPar al buscar el servicio Fl_ObtieneServicio: " + ex.Message + " - " + ex.Data, "Error");
                return false;
            }

            try
            {
                lsQuery = "SELECT A.stored_procedure, B.longitud, B.orden_campo ";
                lsQuery += "FROM TICKET..TIPO_TRANSACCION_PIU A, TICKET..ENTRADA_PIU B ";
                lsQuery += "WHERE A.tipo_transaccion = B.tipo_transaccion ";
                lsQuery += "AND A.funcion = '" + psFuncion.Trim() + "' ";
                lsQuery += "ORDER BY B.orden_campo";
                //EJECUTA QUERY

                var dt = new DataTable();
                dt.Columns.Add("stored_procedure", typeof(string));
                dt.Columns.Add("longitud", typeof(int));
                dt.Columns.Add("orden_campo", typeof(int));
                dt.Clear();

                //dt = ConsultaQuery(lsQuery);  //[SQL]

                //asigna valor a variable
                sSPejecutar = dt.Rows[0]["stored_procedure"].ToString();
                // string sSPejecutar = rssRegistro(0).Value
                lnCont = 0;
                lsQuery = "EXEC TICKET.." + sSPejecutar + " ";
                lsParam = psBia.Substring(1, pnLongPar);

                foreach (DataRow drow in dt.Rows)
                {
                    if (lnCont > 0)
                    {
                        lsQuery += ", ";
                    }
                    lsQuery = lsQuery + "'" + DaParamBia(Convert.ToInt32(drow["longitud"]), ref lsParam) + "'";
                    lnCont += 1;
                }
            }
            catch (Exception ex)
            {
                funcion.Escribe("Error en la Funcion BuscaSPyLongPar, al ejecutar el query: " + lsQuery + " - Error:" + ex.Message + " - " + ex.Data, "Error");
                return false;
            }

            //rssRegistro.close;

            if (Mb_Detalles)
            {
                funcion.Escribe("Stored Procedure a ser ejecutado: " + lsQuery, "Mensaje");
            }
            psQuery = lsQuery;
            return true;
        }

        private bool EjecSPFuncion(string psQuery, ref string psDatos)
        {
            string lsDatos;
            int lnCont;
            DataTable dtSPquery;
            try
            {
                //rssRegistro.Open psQuery;
                //var dtSPquery = new DataTable();
                dtSPquery = ejecutaSPquery(psQuery);

                //dtSPquery.Columns.Add("stored_procedure", typeof(string));
                //dtSPquery.Columns.Add("longitud", typeof(int));
                //dtSPquery.Columns.Add("orden_campo", typeof(int));
                //dtSPquery.Clear();

                if (dtSPquery.Rows.Count > 0)
                {
                    lsDatos = "";
                    lnCont = 0;
                    foreach (DataRow drow in dtSPquery.Rows)
                    {

                        lsDatos += drow[0].ToString() + drow[1].ToString() + drow[2].ToString();
                        lnCont += 1;
                    }
                    psDatos = lsDatos;
                    return true;
                }
                else
                {
                    funcion.Escribe("No obtuvo información EjecSPFuncion - ejecutaSPquery.", "Mensaje");
                    return false;
                }
                //ejecuta query

            }
            catch (Exception ex)
            {
                funcion.Escribe("Error al ejecutar funcion EjecSPFuncion: " + ex.Message + " - " + ex.Data, "Mensaje");
                return false;
            }

        }

        private DataTable ejecutaSPquery(string querySP)
        {
            DataTable dt = new DataTable();
            return dt;
            //return exeSPquery(querySP, "execSPquery");  //[SQL]
        }

        private string DaParamBia(int pnLong, ref string psParam)
        {
            string sDaParamBia;
            try
            {
                sDaParamBia = "";                                  //'Inicia variables
                sDaParamBia = psParam.Substring(1, pnLong);        //'Obtiene el parámetro
                psParam = psParam.Substring(1, pnLong + 1);          //'Prepara la cadena para el siguiente parámetro
            }
            catch (Exception ex)
            {
                funcion.Escribe("Error en la Funcion DaParamBia - Error:" + ex.Message + " - " + ex.Data, "Error");
                return "";
            }
            return sDaParamBia;
        }

        public void Desconectar()
        {
            //Connection.cnnConexion.Close();
            ////rssRegistro = Nothing
            //Connection.cnnConexion = null;

        }

    }
}
