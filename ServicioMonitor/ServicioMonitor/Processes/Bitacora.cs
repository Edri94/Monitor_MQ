using ServicioMonitor.Helpers;
using ServicioMonitor.Mq;
using System;
namespace ServicioMonitor.Processes
{
    public class Bitacora
    {
        private string ArchivoIni;
        private string lsCommandLine;

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
        private string strMsgTypeCole; // Tipo de mensaje: Copy
        private string strMETAGEND; // Bandera que indica el fin del mensaje

        // Variables para el registro de los valores Default
        string strColectorMaxLeng; // Maxima longitud del COLECTOR
        string strMsgMaxLeng; // Maxima longitud del del bloque ME
        string strPS9MaxLeng; // Maxima longitud del formato PS9
        string strReplyToMQ; // MQueue de respuesta para HOST
        string strFuncionSQL; // Funcion a ejecutar al recibir la respuesta
        string strRndLogTerm; // Indica que el atributo Logical Terminal es 

        private string Gs_MQManager;   // MQManager de Escritura
        private string Gs_MQQueueEscritura;   // MQQueue de Escritura
        public DateTime gsAccesoActual;   // Fecha/Hora actual del sistema. La tomamos del servidor NT

        BitacorasMq mqSeries;
        Funcion_Bitacora funcion;

        public Bitacora()
        {
            mqSeries = new BitacorasMq();
            funcion = new Funcion_Bitacora();
        }

        public void ProcesarBitacora(string ruta, string strParametros)
        {
            string[] parametros;
            string ls_MsgVal = "";

            try
            {
                //ArchivoIni = strRutaIni + @"\Bitacoras.ini";
                ConfiguraFileLog("escribeArchivoLOG");
                gsAccesoActual = DateTime.Now;

                lsCommandLine = strParametros.Trim();

                if (lsCommandLine.Equals("") == false)
                {
                    parametros = lsCommandLine.Split('-');

                    Gs_MQManager = parametros[0].Trim();
                    Gs_MQQueueEscritura = parametros[1].Trim();
                    strFuncionSQL = parametros[3];
                }
                else
                {
                    ObtenerInfoMq();
                }


                ConfiguraHeader_IH_ME();

                if (!ValidaInfoMQ(ls_MsgVal))
                {
                    funcion.Escribe("Se presentó un error en la función ValidaInfoMQ invocada desde el MAIN: " + ls_MsgVal + ". Función SQL: " + strFuncionSQL, "Mensaje");
                    return;
                }

                ProcesoBDtoMQQUEUE();

                funcion.Escribe("Termina proceso bitácoras. Función SQL: " + strFuncionSQL, "Mensaje");


            }
            catch (Exception Err)
            {
                //mqSeries.MQDesconectar(mqSeries.queueManager, mqSeries.queue);
                funcion.Escribe(Err, "Error");
            }
        }

        private void ConfiguraFileLog(string section)
        {
            string strlogFileName = funcion.getValueAppConfig("logFileName", section);
            string strlogFilePath = funcion.getValueAppConfig("logFilePath", section);

            bool Mb_GrabaLog = true;
        }

        private void ObtenerInfoMq()
        {
            string section = "mqSeries";
            Gs_MQManager = funcion.getValueAppConfig("MQManager", section);
            Gs_MQQueueEscritura = funcion.getValueAppConfig("MQEscritura", section);
            strFuncionSQL = funcion.getValueAppConfig("FGBitacora", section);
        }

        private void ConfiguraHeader_IH_ME()
        {
            string section = "headerih";
            strFuncionHost = funcion.getValueAppConfig("PRIMERVALOR", section);
            strHeaderTagIni = $"<{ funcion.getValueAppConfig("IHTAGINI", section)}>";
            strIDProtocol = funcion.getValueAppConfig("IDPROTOCOL", section);
            strLogical = funcion.getValueAppConfig("FGBitacora", section);
            strAccount = funcion.getValueAppConfig("ACCOUNT", section);
            strUser = funcion.getValueAppConfig("User", section);
            strSeqNumber = funcion.getValueAppConfig("SEQNUMBER", section);
            strTXCode = funcion.getValueAppConfig("TXCODE", section);
            strUserOption = funcion.getValueAppConfig("USEROPT", section);
            strCommit = funcion.getValueAppConfig("Commit", section);
            strMsgType = funcion.getValueAppConfig("MSGTYPE", section);
            strProcessType = funcion.getValueAppConfig("PROCESSTYPE", section);
            strChannel = funcion.getValueAppConfig("CHANNEL", section);
            strPreFormat = funcion.getValueAppConfig("PREFORMATIND", section);
            strLenguage = funcion.getValueAppConfig("LANGUAGE", section);
            strHeaderTagEnd = $"</{ funcion.getValueAppConfig("IHTAGEND", section)}>";

            section = "headerme";

            strMETAGINI = $"<{ funcion.getValueAppConfig("METAGINI", section)}>";
            strMsgTypeCole = funcion.getValueAppConfig("TIPOMSG", section);
            strMETAGEND = $"</{ funcion.getValueAppConfig("METAGEND", section)}>";

            section = "defaultValues";

            strColectorMaxLeng = funcion.getValueAppConfig("COLMAXLENG", section);
            strMsgMaxLeng = funcion.getValueAppConfig("MSGMAXLENG", section);
            strPS9MaxLeng = funcion.getValueAppConfig("PS9MAXLENG", section);
            strReplyToMQ = funcion.getValueAppConfig("ReplyToQueue", section);
            strRndLogTerm = funcion.getValueAppConfig("RandomLogTerm", section);
        }

        private bool ValidaInfoMQ(string ps_MsgVal)
        {
            bool validaInfoMQ = false;
            string ls_msg = "";

            if (Gs_MQManager.Trim() == "")
            {
                ls_msg = ls_msg + (ls_msg.Length > 0 ? ((char)13).ToString() : "") + "Falta MQ Manager envio.";
            }
            if (Gs_MQQueueEscritura.Trim() == "")
            {
                ls_msg = ls_msg + (ls_msg.Length > 0 ? ((char)13).ToString() : "") + "Falta MQ Queue envio.";
            }
            if (ls_msg.Trim() == "")
            {
                validaInfoMQ = true;
            }

            ps_MsgVal = ls_msg;

            return validaInfoMQ;

        }

        private void ProcesoBDtoMQQUEUE()
        {
            string Ls_MensajeMQ;
            string Ls_MsgColector;

            string sFechaEnvio;
            string sEnvioConse;
            string sMensajeEnvio = "";


            try
            {
                funcion.Escribe("Inicia envío de mensajes a Host: " + gsAccesoActual + " Función SQL: " + strFuncionSQL, "Mensaje");

                if (mqSeries.ConectarMQ(Gs_MQManager))
                {
                    //SetParameterAppSettings("valorTk14", "TKCONSECUTIVO", "algo");
                    mqSeries.blnConectado = true;
                }
                else
                {
                    funcion.Escribe("Fallo conexión MQ-Manager ", "Mensaje");
                    return;
                }

                sFechaEnvio = funcion.Left(DateTime.Now.ToString("yyyymmddhhmmss") + funcion.Space(26), 26);
                sEnvioConse = funcion.Left(funcion.getValueAppConfig("valorTk14", "TKCONSECUTIVO") + funcion.Space(1), 1);

                Ls_MsgColector = (strFuncionSQL + funcion.Space(8)).Substring(0, 8) + sFechaEnvio;

                if (Ls_MsgColector.Length > 0)
                {
                    Ls_MensajeMQ = ASTA_ENTRADA(Ls_MsgColector);
                    if (Ls_MensajeMQ != "")
                    {
                        funcion.Escribe("Mensaje Enviado: " + Ls_MensajeMQ, "Mensaje");

                        if (mqSeries.EnviarMensajeMQ(Gs_MQQueueEscritura))
                        {
                            sMensajeEnvio = (sEnvioConse + 1).ToString();

                            if (Int32.Parse(sMensajeEnvio) > 9)
                            {
                                sMensajeEnvio = ((char)(1)).ToString();
                            }
                            //SetParameterAppSettings("valorTk14", "TKCONSECUTIVO", sMensajeEnvio);
                        }
                        else
                        {
                            funcion.Escribe("Se ha presentado un error al escribir la solicitud en la MQ QUEUE:", "Mensaje");
                        }
                    }
                    else
                    {
                        funcion.Escribe("Se ha presentado un error durante el armado del formato PS9 funcion ASTA_ENTRADA.Colector: " + Ls_MsgColector, "Mensaje");
                    }
                }
                else
                {
                    funcion.Escribe("Se ha presentado un error al armar el Layout TKT14. No existe longitud en el Colector", "Mensaje");
                }

                mqSeries.DesconectarMQ();

                funcion.Escribe("Envio de solicitures TKT -> Host Terminado", "Mensaje");
                funcion.Escribe("Solicitudes enviadas a MQ: " + sMensajeEnvio, "Mensaje");
            }
            catch (Exception ex)
            {
                funcion.Escribe(ex, "Error");
            }

        }

        private string ASTA_ENTRADA(string strMsgColector)
        {
            string ASTA_ENTRADA;

            string ls_TempColectorMsg;
            string ls_BloqueME;
            int ln_longCOLECTOR;
            int ln_AccTerminal;

            try
            {
                ls_TempColectorMsg = strMsgColector;
                funcion.Escribe("Entre a  funcion: ASTA_ENTRADA");
                funcion.Escribe("ls_TempColectorMsg: " + ls_TempColectorMsg);
                funcion.Escribe("strColectorMaxLeng: " + strColectorMaxLeng);

                if (ls_TempColectorMsg.Length > Int32.Parse(strColectorMaxLeng))
                {
                    funcion.Escribe("La longitud del colector supera el maximo permitido", "Mensaje");
                    //GoTo ErrorASTA
                }

                ls_BloqueME = funcion.Left((strMETAGINI + funcion.Space(4)).Trim(), 4);
                ls_BloqueME = ls_BloqueME + funcion.Right("0000" + (ls_TempColectorMsg.Length.ToString()), 4);
                ls_BloqueME = ls_BloqueME + funcion.Left((strMsgTypeCole.Trim() + funcion.Space(1)), 1);
                ls_BloqueME = ls_BloqueME + ls_TempColectorMsg;
                ls_BloqueME = ls_BloqueME + funcion.Left(strMETAGEND + funcion.Space(5), 5);

                if (ls_BloqueME.Length > Int32.Parse(strMsgMaxLeng))
                {
                    funcion.Escribe("La longitud del Bloque ME supera el maximo permitido", "Mensaje");
                    //GoTo ErrorASTA
                }

                ASTA_ENTRADA = funcion.Left(strFuncionHost.Trim() + funcion.Space(8), 8);
                ASTA_ENTRADA = ASTA_ENTRADA + funcion.Left(strHeaderTagIni.Trim() + funcion.Space(4), 4);
                ASTA_ENTRADA = ASTA_ENTRADA + funcion.Left(strIDProtocol.Trim() + funcion.Space(2), 2);

                if (strRndLogTerm.Trim().Equals("1"))
                {
                    ln_AccTerminal = 0;
                    do
                    {
                        //new  Random(DateTime.Now.Second);
                        var Rnd = new Random(DateTime.Now.Second * 1000);
                        ln_AccTerminal = Rnd.Next();

                    } while (ln_AccTerminal > 0 && ln_AccTerminal < 2000);

                    ASTA_ENTRADA = ASTA_ENTRADA + funcion.Left((ln_AccTerminal.ToString("D4") + funcion.Space(8)), 8);
                }
                else
                {
                    ASTA_ENTRADA = ASTA_ENTRADA + funcion.Left(strAccount.Trim() + funcion.Space(8), 8);
                }

                ASTA_ENTRADA = ASTA_ENTRADA + funcion.Left(strAccount.Trim() + funcion.Space(8), 8);
                ASTA_ENTRADA = ASTA_ENTRADA + funcion.Left(strUser.Trim() + funcion.Space(8), 8);
                ASTA_ENTRADA = ASTA_ENTRADA + funcion.Left(strSeqNumber.Trim() + funcion.Space(8), 8);
                ASTA_ENTRADA = ASTA_ENTRADA + funcion.Left(strTXCode.Trim() + funcion.Space(8), 8);
                ASTA_ENTRADA = ASTA_ENTRADA + funcion.Left(strUserOption.Trim() + funcion.Space(2), 2);

                ln_longCOLECTOR = 65 + ls_BloqueME.Length;

                if (ln_longCOLECTOR > Int32.Parse(strPS9MaxLeng.Trim()))
                {
                    funcion.Escribe("La longitud del Layout PS9 supera el maximo permitido", "Mensaje");
                    //GoTo ErrorASTA
                }

                ASTA_ENTRADA = ASTA_ENTRADA + funcion.Right("00000" + ln_longCOLECTOR.ToString(), 5);
                ASTA_ENTRADA = ASTA_ENTRADA + funcion.Left(strCommit.Trim() + funcion.Space(1), 1);
                ASTA_ENTRADA = ASTA_ENTRADA + funcion.Left(strMsgType.Trim() + funcion.Space(1), 1);
                ASTA_ENTRADA = ASTA_ENTRADA + funcion.Left(strProcessType.Trim() + funcion.Space(1), 1);
                ASTA_ENTRADA = ASTA_ENTRADA + funcion.Left(strChannel.Trim() + funcion.Space(2), 2);
                ASTA_ENTRADA = ASTA_ENTRADA + funcion.Left(strPreFormat.Trim() + funcion.Space(1), 1);
                ASTA_ENTRADA = ASTA_ENTRADA + funcion.Left(strLenguage.Trim() + funcion.Space(1), 1);
                ASTA_ENTRADA = ASTA_ENTRADA + funcion.Left(strHeaderTagEnd.Trim() + funcion.Space(5), 5);
                ASTA_ENTRADA = ASTA_ENTRADA + ls_BloqueME;


                return ASTA_ENTRADA;
            }
            catch (Exception ex)
            {
                funcion.Escribe(ex, "Error");
                return ex.Message;
            }
        }

    }
}
