using IBM.WMQ;
using ServicioMonitor.Helpers;
using ServicioMonitor.Mq;
using ServicioMonitor.Processes;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.ServiceProcess;
using static ServicioMonitor.Mq.MqSeries;

namespace MonitorMQTKT
{
    public partial class MonitorMQTKT : ServiceBase
    {
        bool blBandera = false;
        bool blIniMonitorMQTKT = false; //SGGG 20-01-2022 - Se Genera bandera para imprimir en bitacora inicio del servicio
        bool blIniMensajesMQ = false;   //SGGG 21-01-2022 - Se Genera bandera para imprimir en bitacora inicio del servicio    
        bool blIniTKTMQ = false;        //SGGG 21-01-2022 - Se Genera bandera para imprimir en bitacora inicio del servicio
        bool blIniBitaora = false;      //SGGG 21-01-2022 - Se Genera bandera para imprimir en bitacora inicio del servicio



        private bool ActivoProcFuncAuto;        // Variable para determinar si se desea ejecutar el proceso del Monitoreo
        private bool ModoMonitor;               // Variable para determinar el modo de operacion del monitor


        // ***** Para realizar el monitoreo de bitacoras
        private int miTotalMonitor;


        private double MensajesMQ;

        private Funcion_Monitor funcion;
        private MqSeries mqSeries ;
        private MqMonitorTicket monitorTicket;


        public MonitorMQTKT()
        {
            funcion = new Funcion_Monitor();
            mqSeries = new MqSeries();
            monitorTicket = new MqMonitorTicket();
            InitializeComponent();
         }

        protected override void OnStart(string[] args)
        {
            // TODO: agregar código aquí para iniciar el servicio.
            blIniMonitorMQTKT = true;   //SGGG 20-01-2022 - Se inicializa bandera al iniciarse el servicio MonitorMQTKT
            blIniMensajesMQ = true;     //SGGG 21-01-2022 - Se inicializa bandera al iniciarse el servicio MEnsajesMQ
            blIniTKTMQ = true;          //SGGG 21-01-2022 - Se inicializa bandera al iniciarse el servicio TKTMQ
            blIniBitaora = true;        //SGGG 21-01-2022 - Se inicializa bandera al iniciarse el servicio Bitacora

            tmrMonitorMQTKT.Start();
        }

        protected override void OnStop()
        {
            // TODO: agregar código aquí para realizar cualquier anulación necesaria para detener el servicio.
            tmrMonitorMQTKT.Stop();
        }
     


        private void tmrMonitorMQTKT_Elapsed(object sender, System.Timers.ElapsedEventArgs e)
        {

            if (blBandera) return;
            try
            {
                blBandera = true;
                EventLog.WriteEntry("Ejecuta Monitor MonitorMQPU", EventLogEntryType.Information);
                                

                FrmMonitor_Load();


                monitorTicket.dblCiclosBitacoras += 10;
                monitorTicket.dblCiclosTKTMQ += 10;
                monitorTicket.dblCiclosFuncionarios += 10;
                monitorTicket.dblCiclosAutorizaciones += 10;

           

                if (monitorTicket.intgMonitor == 1)
                {
                    if (monitorTicket.strFormatoTiempoBitacoras != "S")
                    {
                        if (monitorTicket.dblCiclosBitacoras >= (monitorTicket.intTiempoBitacoras * 60))
                        { 
                            TmrBitacora();
                            monitorTicket.dblCiclosBitacoras = 0;
                        }
                    }
                    else
                    {
                        if (monitorTicket.dblCiclosBitacoras >= monitorTicket.intTiempoBitacoras)
                        {

                            TmrBitacora();
                            monitorTicket.dblCiclosBitacoras = 0;
                        }
                    }
                }


                if (monitorTicket.strFormatoTiempoTKTMQ != "S")
                {
                    if (monitorTicket.dblCiclosTKTMQ >= (monitorTicket.intTiempoTKTMQ * 60))
                    {
                        TmrTKTMQ();
                        monitorTicket.dblCiclosTKTMQ = 0;
                    }
                }
                else
                {
                    if (monitorTicket.dblCiclosTKTMQ >= monitorTicket.intTiempoTKTMQ)
                    {
                        TmrTKTMQ();
                        monitorTicket.dblCiclosTKTMQ = 0;
                    }
                }

                if (monitorTicket.strFormatoTiempoFuncionarios != "S")
                {
                    if (monitorTicket.dblCiclosFuncionarios >= (monitorTicket.intTiempoFuncionarios * 60))
                    {
                        ActivarEnvioFuncAuto("F");
                        monitorTicket.dblCiclosFuncionarios = 0;
                    }
                }
                else
                {
                    if (monitorTicket.dblCiclosFuncionarios >= monitorTicket.intTiempoFuncionarios)
                    {
                        ActivarEnvioFuncAuto("F");
                        monitorTicket.dblCiclosFuncionarios = 0;
                    }
                }


                if (monitorTicket.strFormatoTiempoAutorizaciones != "S")
                {
                    if (monitorTicket.dblCiclosAutorizaciones >= (monitorTicket.intTiempoAutorizaciones * 60))
                    {
                        ActivarEnvioFuncAuto("A");
                        monitorTicket.dblCiclosAutorizaciones = 0;
                    }
                }
                else
                {
                    if (monitorTicket.dblCiclosAutorizaciones >= monitorTicket.intTiempoAutorizaciones)
                    {
                        ActivarEnvioFuncAuto("A");
                        monitorTicket.dblCiclosAutorizaciones = 0;
                    }
                }








            }

            catch (Exception ex)
            {
                EventLog.WriteEntry(ex.Message, EventLogEntryType.Error);
            }

            blBandera = false;
        }




       

        //public FrmMonitor()
        //{
        //    //Inicializando variables-----------
        //    funcion = new Funcion();
        //    //mqSeries = new MqSeries(); //Solo se podran usar mnetodos que hereden de esta clase
        //    monitorTicket = new MqMonitorTicket();
        //    //----------------------------------
        //    InitializeComponent();
        //}


        private void GuardarLog()
        {
            if (blIniMonitorMQTKT ==  true) //SGGG 20-01-2022 - Se agrega validación para que sólo imprima al iniciarse el servicio
            {
                funcion.Escribe("El siguiente reporte se genera a partir del botón 'Guardar el Registro de Operaciones' o cuando ha cambiado el dia de monitoreo o cuando se ha pulsado el botón 'Salir' del Monitor.");
                funcion.Escribe("---------  Reporte del estado de los procesos  ---------");
                funcion.Escribe("*********  Registro de operaciones procesadas  *********");
                funcion.Escribe("   Respuestas (HOST->NT) registradas proceso de Monitoreo");
                //'Escribe "       > Duración del CICLO[seg] : " & IntRecepResMonitor
                funcion.Escribe("   Solicitudes (NT->HOST) registradas proceso de Funcionarios y Autorizaciones");
                //'Escribe( "       > Duración del CICLO[min] : " & IntEnvioMsgMonitor
                funcion.Escribe("---------  Fin del reporte del estado de los procesos  ---------");
                funcion.Escribe("");
                
            }
        }

        private bool ResetMonitor()
        {
            bool reset_monitor = false;

            try
            {
                funcion.Escribe("Respaldo del estado del monitor " + funcion.ObtenFechaFormato(1));

                GuardarLog();

                if (monitorTicket.QUEUE != null)
                {
                    if (monitorTicket.QUEUE.IsOpen)
                    {
                        monitorTicket.CerrarColaMQ();
                    }
                }

                monitorTicket.DesconectarMQ();

                if (monitorTicket.ConectarMQ(monitorTicket.strMQManager))
                {
                    monitorTicket.blnConectado = true;
                }
                else
                {
                    funcion.Escribe("Falla en Monitor < Error al conectarse con la MQ > : " + funcion.ObtenFechaFormato(1));
                    funcion.Escribe("Detalles : " + monitorTicket.QUEUE.ReasonCode + " " + monitorTicket.QUEUE.ReasonName);
                }


            }
            catch (Exception error)
            {
                funcion.Escribe("Falla en Monitor < Error al conectarse con la MQ > : " + funcion.ObtenFechaFormato(1));
                funcion.Escribe("Detalles : " + monitorTicket.QMGR.ReasonCode + " " + monitorTicket.QMGR.ReasonName);
                funcion.Escribe(error);
            }

            return reset_monitor;
        }

        private void FrmMonitor_Load()
        {        
            if (monitorTicket.Inicia())
            {
               


                ModoMonitor = (monitorTicket.intgModoMonitor == 1) ? true : false;
                ActivoProcFuncAuto = (monitorTicket.intgActv_FuncAuto == 1) ? true : false;

                funcion.Escribe("Aplicación Monitor iniciado : " + funcion.ObtenFechaFormato(1));

                CargaInfMonitoreo();

                Iniciar();
            }
            else
            {
                funcion.Escribe("No se puede continuar con la carga. Archivo Ini no existe.");
            }

        }


        private void CargaInfMonitoreo()
        {
            miTotalMonitor = Int32.Parse(funcion.getValueAppConfig("PMONITOREOS"));
        }

        private void FrmMonitor_FormClosing()
        {
            Detener();

            if (monitorTicket.QUEUE != null)
            {
                if (monitorTicket.QUEUE.IsOpen) monitorTicket.CerrarColaMQ();
            }

            monitorTicket.DesconectarMQ();

            monitorTicket.QUEUE = null;
            monitorTicket.QMGR = null;
        }

        private void Detener()
        {
            try
            {
                if (monitorTicket.QUEUE != null)
                {
                    if (monitorTicket.QUEUE.IsOpen) monitorTicket.CerrarColaMQ();
                }

                if (monitorTicket.DesconectarMQ())
                {
                    monitorTicket.blnConectado = false;
                }
            }
            catch (MQException ex)
            {
                funcion.Escribe(ex);
                funcion.Escribe("" + monitorTicket.QUEUE.ReasonCode + " " + monitorTicket.QUEUE.ReasonName);
                funcion.Escribe("Falla en Monitor < Falla en el cierre de MQ-Series > : " + funcion.ObtenFechaFormato(1));
            }
            catch (Exception ex)
            {
                funcion.Escribe(ex);
            }
        }

        private void Iniciar()
        {
            try
            {
                if (monitorTicket.ConectarMQ(monitorTicket.strMQManager))//cambiar
                {
                    monitorTicket.blnConectado = true;
                }
                else
                {
                    funcion.Escribe("Falla en Monitor < Falla en el cierre de MQ-Series > : " + funcion.ObtenFechaFormato(1));
                    return;
                    //funcion.Escribe("" + monitorTicket.QUEUE.ReasonCode + " " + monitorTicket.QUEUE.ReasonName);
                }

                if (ModoMonitor == true)
                {
                    funcion.Escribe("Monitor iniciado en modo de monitoreo: " + funcion.ObtenFechaFormato(1));

                    if (!ActivoProcFuncAuto)
                    {
                        funcion.Escribe("El procesos de Funcionarios-Autorizaciones se encuentra en estado inactivo");
                    }
                    else
                    {
                        funcion.Escribe("El procesos de Funcionarios-Autorizaciones se encuentra en estado activo");
                    }
                }
                else
                {
                    funcion.Escribe("Monitor iniciado en modo de procesamiento: " + funcion.ObtenFechaFormato(1));
                }

                tmrRestar.Enabled = true;
                tmrRestar.Interval = monitorTicket.intgtmrRestar * 1000;

                TimeSpan Diff_dates = Convert.ToDateTime(monitorTicket.FechaRestar).Subtract(monitorTicket.date); //opcion 1
                int dias_diferiencia = (monitorTicket.date - Convert.ToDateTime(monitorTicket.FechaRestar)).Days; //opcion 2
                
                if (dias_diferiencia != 0)
                {
                    monitorTicket.FechaRestar = monitorTicket.date.AddDays(1).ToString();
                }

                
                if (!funcion.UpdateAppSettings("RestarMonitor", monitorTicket.FechaRestar))
                {
                    funcion.Escribe("Iniciar() No se encontro el archivo");
                    //this.Close();
                }
                else
                {
                    funcion.Escribe("Iniciar() Se actulizo [FechaRestar] en el archivo App.Settings " + monitorTicket.FechaRestar);
                }

            }
            catch (Exception ex)
            {
                funcion.Escribe("Error en la conexion con el servidor MQ: " + monitorTicket.QUEUE.ReasonCode + " " + monitorTicket.QUEUE.ReasonName);
                funcion.Escribe(ex);
            }
        }
       

      


        private void TmrTKTMQ()
        {
            double ln_MsgEncontrado;
            ln_MsgEncontrado = RevisaMQ(monitorTicket.strMQManager, monitorTicket.strMQQMonitorLectura, monitorTicket.strMQQMonitorEscritura, "0");
        }

        private void ActivarEnvioFuncAuto(string psMonitor)
        {
            Mensaje mensaje; ;

            string LsProceso = "";
            string sMensaje;

            switch (psMonitor)
            {
                case "A":
                    LsProceso = "PROCESO2";
                    break;
                case "F":
                    LsProceso = "PROCESO1";
                    break;
                default:
                    break;
            }


            if (ActivoProcFuncAuto)
            {
                string a = funcion.getValueAppConfig(LsProceso);
                string b = funcion.getValueAppConfig(LsProceso);

                var prueba = funcion.InStr(3, b, ",");

                sMensaje = funcion.Mid(a, 3, funcion.InStr(3, b, ",") - 3);


                funcion.Escribe("EJECUTANDO fValidaEjecucion(): " + sMensaje);  //[PRUEBAS]

                if (fValidaEjecucion(sMensaje))
                {
                    funcion.Escribe("La operación: " + sMensaje + " SI esta habilitada para este día " + funcion.ObtenFechaFormato(1));  //[PRUEBAS]
                    mensaje = new Mensaje();
                    mensaje.ProcesarMensajes(ref blIniMensajesMQ, monitorTicket.strMQManager + "-" + monitorTicket.strMQQMonitorEscritura + "-" + psMonitor); //SGGG - 21-01-2022 - Sea grega nuevo parametro para la bancwera de inicio de servicio
                    mensaje = null;
                }
                else
                {
                    funcion.Escribe("La operación: " + sMensaje + " NO esta habilitada para este día " + funcion.ObtenFechaFormato(1));
                }
            }
        }


        private double RevisaMQ(string MQManager, string MQQLectura, string MQQEscritura, string psOtros)
        {
            long lngErr;
            int j;
            //long lngMQOpen;
            string lsExeParam;
            double RevisaMQ;

            try
            {
                //lngMQOpen = (long)MQOPEN.MQOO_INQUIRE;

                if(monitorTicket.blnConectado == true)
                {
                    if (monitorTicket.AbrirColaMQ(MQQLectura, MqMonitorTicket.MQOPEN.MQOO_INQUIRE)) //cambiar
                    {
                        j = 1;
                        lngErr = monitorTicket.QUEUE.ReasonCode;

                        MensajesMQ = monitorTicket.QUEUE.CurrentDepth;

                        if (MensajesMQ < 0)
                        {
                            RevisaMQ = 0;
                            return RevisaMQ;
                        }

                        monitorTicket.CerrarColaMQ();


                        if (!ModoMonitor)
                        {
                            if (MensajesMQ > 0)
                            {
                                do
                                {
                                    if (psOtros.CompareTo("") != 0)
                                    {
                                        lsExeParam = MQManager + "-" + MQQLectura + "-" + MQQEscritura + "-" + psOtros;
                                    }
                                    else
                                    {
                                        lsExeParam = MQManager + "-" + MQQLectura + "-" + MQQEscritura + "-0";
                                    }

                                    Tkt tkt = new Tkt();
                                    tkt.ProcesarMensajes(lsExeParam, ref blIniTKTMQ); //[CAMBIAR POR APP.PATH]  //SGGG - 21-01-2022 - Sea grega nuevo parametro para la bancwera de inicio de servicio

                                    j++;

                                } while (j <= MensajesMQ);
                            }
                        }
                    }
                    else
                    {
                        funcion.Escribe("Error en la conexion con el MQ Manager : " + monitorTicket.QMGR.ReasonCode + " " + monitorTicket.QMGR.ReasonName);
                        funcion.Escribe("Ejecutamos el reinicio del monitor por problemas en la comunicacion con MQManager " + monitorTicket.strMQManager);
                        ReConectar();
                    }

                    RevisaMQ = MensajesMQ;
                }
                else
                {
                    funcion.Escribe("No esta conectada la MQ", "Errro");
                    RevisaMQ = -1;
                }

                return RevisaMQ;

            }
            catch (Exception ex)
            {
                funcion.Escribe(ex);
                funcion.Escribe("Monitor error al ejecutar el proceso de TKTMQ el " + funcion.ObtenFechaFormato(1));
                RevisaMQ = -1;
                return RevisaMQ;
            }
        }

        private void TmrBitacora()
        {
            string Ejecutable;
            List<string> Parametro = new List<string>();
            int intlBitacoras;
            string sValor;
            string[] vntBitacora;

            for (intlBitacoras = 1; intlBitacoras < miTotalMonitor; intlBitacoras++)
            {
                sValor = funcion.getValueAppConfig("PARAMETRO" + intlBitacoras);
                vntBitacora = sValor.Split(',');


                if (Int32.Parse(vntBitacora[0]) == 1)
                {
                    funcion.Escribe("EJECUTANDO fValidaEjecucion(): " + vntBitacora[1]);  //[PRUEBAS]

                    if (fValidaEjecucion(vntBitacora[1]))
                    {
                        funcion.Escribe("La operación: " + vntBitacora[1] + " SI esta habilitada para este día " + funcion.ObtenFechaFormato(1));  //[PRUEBAS]

                        Parametro = funcion.getValueAppConfig(vntBitacora[1]).Split(',').ToList();
                        Ejecutable = Parametro[0];

                        if (Ejecutable == "M")
                        {
                            Mensaje mensajes_MQ = new Mensaje();

                            mensajes_MQ.ProcesarMensajes(ref blIniMensajesMQ,monitorTicket.strMQManager + "-" + monitorTicket.strMQQMonitorEscritura + "-1-" + Parametro[1]);  //SGGG - 21-01-2022 - Sea grega nuevo parametro para la bancwera de inicio de servicio

                        }
                        else
                        {
                            Bitacora bitacoras_MQ = new Bitacora();

                            bitacoras_MQ.ProcesarBitacora(monitorTicket.strMQManager + "-" + monitorTicket.strMQQMonitorEscritura + "-1-" + Parametro[1],ref blIniTKTMQ);  //SGGG - 21-01-2022 - Sea grega nuevo parametro para la bancwera de inicio de servicio

                        }
                    }
                    else
                    {
                        funcion.Escribe("La operación: " + vntBitacora[1] + " NO esta habilitada para este día " + funcion.ObtenFechaFormato(1));  //[PRUEBAS]
                    }
                }
            }
        }

        private bool fValidaEjecucion(string psBitacora)
        {
            string iTotalProcesos;
            int iRow;
            string sValor;
            string[] sParametros;
            int intCuenta;
            bool fValidaEjecucion = false;
            try
            {
                iTotalProcesos = funcion.getValueAppConfig("PROCESOS"); 

                funcion.Escribe("iTotalProcesos:" + iTotalProcesos); //[PRUEBAS]

                for (iRow = 1; iRow <= Int32.Parse(iTotalProcesos); iRow++)
                {
                    funcion.Escribe("iRow:" + iRow); //[PRUEBAS]
                    sValor = funcion.getValueAppConfig("PROCESO" + iRow);

                    sParametros = sValor.Split(',');

                    funcion.Escribe("sParametros[9]:" + sParametros[9]); //[PRUEBAS]
                    for (intCuenta = 0; intCuenta <= Int32.Parse(sParametros[9]); intCuenta++)
                    {
                        funcion.Escribe("intCuenta:" + intCuenta); //[PRUEBAS]
                        
                        funcion.Escribe("si (sParametros[0]:" + sParametros[0] + " es igual a 1) y ( sParametros[1]:" + sParametros[1] + " es igual a psBitacora:" + psBitacora + ")"); //[PRUEBAS]                      
                        if (Int32.Parse(sParametros[0]) == 1 && sParametros[1] == psBitacora)
                        {
                            funcion.Escribe("si (sParametros[intCuenta + 2]:" + sParametros[intCuenta + 2] + " es igual a SI"); //[PRUEBAS]
                            if (sParametros[intCuenta + 2] == "Si")
                            {
                                funcion.Escribe("si ((int)DateTime.Now.DayOfWeek: " + (int)DateTime.Now.DayOfWeek + " es igual a intCuenta + 1:" + (intCuenta + 1)); //[PRUEBAS]
                                if ((int)DateTime.Now.DayOfWeek == intCuenta + 1)
                                {
                                    fValidaEjecucion = true;
                                    intCuenta = Int32.Parse(sParametros[9]);
                                    iRow = Int32.Parse(iTotalProcesos);
                                    break;
                                }
                            }
                        }
                        else
                        {
                            fValidaEjecucion = false;
                            intCuenta = Int32.Parse(sParametros[9]);
                        }
                    }
                }
                return fValidaEjecucion;
            }
            catch (Exception ex)
            {
                funcion.Escribe(ex);
                return fValidaEjecucion;
            }
        }


        private void ReConectar()
        {
            try
            {
                if (monitorTicket.QUEUE != null)
                {
                    if (monitorTicket.QUEUE.IsOpen)
                    {
                        monitorTicket.CerrarColaMQ();
                    }

                    monitorTicket.DesconectarMQ();

                    if (monitorTicket.ConectarMQ(monitorTicket.strMQManager))
                    {
                        monitorTicket.blnConectado = true;
                    }
                    else
                    {
                        funcion.Escribe("Falla en Monitor < Error al reconectarse con la MQ > : " + funcion.ObtenFechaFormato(1));
                        funcion.Escribe("Detalles : " + monitorTicket.QUEUE.ReasonCode + " " + monitorTicket.QUEUE.ReasonName);
                    }
                }
            }
            catch (Exception ex)
            {
                funcion.Escribe(ex);
            }
        }

        private void btnTest_Click(object sender, EventArgs e)
        {
            double ln_MsgEncontrados;

            RevisaMQ(monitorTicket.strMQManager, monitorTicket.strMQQMonitorLectura, monitorTicket.strMQQMonitorEscritura, "0");
            ln_MsgEncontrados = monitorTicket.dblRevisaMQ;

        }

        private void tmrRestar_Elapsed(object sender, System.Timers.ElapsedEventArgs e)
        {
            tmrRestar.Enabled = false;

            if (monitorTicket.date > Convert.ToDateTime(monitorTicket.FechaRestar))
            {
                ResetMonitor();

                monitorTicket.FechaRestar = monitorTicket.date.ToString();

                funcion.UpdateAppSettings("RestarMonitor", monitorTicket.FechaRestar);

                if (blIniMonitorMQTKT == true) //SGGG 20-01-2022 - Se agrega validación para que sólo imprima al iniciarse el servicio
                {
                    funcion.Escribe("Aplicación Monitor iniciado : " + monitorTicket.currentDate, "Mensaje");
                    funcion.Escribe("Monitor iniciado en modo de procesamiento: " + monitorTicket.currentDate, "Mensaje");
                    blIniMonitorMQTKT = false; //SGGG 20-01-2022 - Se apaga la bandera para que no imprima en cada timer
                }
            }

            tmrRestar.Enabled = true;
        }

    }
}
