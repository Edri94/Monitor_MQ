using ServicioMonitor.Helpers;
using System;
using System.Security.Cryptography;


namespace ServicioMonitor.Mq
{
    public class MqMonitorTicket : MqSeries
    {
        private static readonly TripleDESCryptoServiceProvider DES = new TripleDESCryptoServiceProvider();
        private static readonly MD5CryptoServiceProvider MD5 = new MD5CryptoServiceProvider();

        public string ArchivoLog;
        public string dateFormat = "yyyyMMdd";
        public string dateFormatHora = "yyyy-MM-dd HHmmss";

        public bool bIniciarmod;
        public bool bMQConectar;
        public bool bMQAbrirCola;
        public bool bfValidaEjecucion;
        public bool bReConectar;
        public double dblRevisaMQ;

        //' Variables para controlar los periodos de monitoreo
        public string FechaRestar;                      //' Valor que determina la fecha y hora para ejecutar un restar del monitor
        public string StrLogPath;                       //' Ruta para almacenar el log del monitor
        public string StrlogName;                       //' Nombre y extención del archivo log
        public string StrlogActivo;                     //' Nombre y extención del archivo log



        public double dblCiclosBitacoras;
        public double dblCiclosTKTMQ;
        public double dblCiclosFuncionarios;
        public double dblCiclosAutorizaciones;

        public string strMQManager;
        public string strMQQMonitorEscritura;
        public string strMQQMonitorLectura;

        public int intgModoMonitor;
        public int intgActv_FuncAuto;
        public int intgMonitor;

        public string inttmrRestar;
        public string inttmrMonitor;
        public string inttmrBitacora;

        public int intgtmrRestar;
        public int intgtmrMonitor;
        public int intgtmrBitacora;


        //private string[] Temporal = TiempoBitacoras.Split(',');
        public string strFormatoTiempoBitacoras;
        public string strFormatoTiempoTKTMQ;
        public string strFormatoTiempoFuncionarios;
        public string strFormatoTiempoAutorizaciones;

        public string TiempoBitacoras;
        public string TiempoMensajes;
        public string Tiempofuncionarios;
        public string TiempoAutorizacion;

        public int intTiempoBitacoras;
        public int intTiempoTKTMQ;
        public int intTiempoFuncionarios;
        public int intTiempoAutorizaciones;

        public static string LogPath;
        public static string LogFile;
        public static string LogActivo;

        public string PMONITOREOS;
        public string PARAMETRO1;
        public string PARAMETRO2;
        public string PARAMETRO3;
        public string PARAMETRO4;

        public string PROCESOS;
        public string PROCESO1;
        public string PROCESO2;
        public string PROCESO3;
        public string PROCESO4;
        public string PROCESO5;
        public string PROCESO6;

        public string Autorizaciones_Operatoria;
        public string Log_Apertura_de_Cuentas;
        public string LOG_Ordenes_de_Pago;
        public string LOG_Operaciones_CED;
        public string LOG_Operaciones_TDD;

        public string Ambiente;
        public string Prefijo;
        public string fileFuncionarios;
        public string fileSucursales;
        public string fileUsuarios;

        public string dirBitacora;
        public string Bitacora;


        public DateTime date;
        public DateTime currentDate;
        public new Funcion_Monitor funcion;

        public MqMonitorTicket()
        {
            funcion = new Funcion_Monitor();

            date = DateTime.Now;
            currentDate = DateTime.Now;

            strMQManager = funcion.getValueAppConfig("MQManager", "");
            strMQQMonitorEscritura = funcion.getValueAppConfig("MQEnvioMsgMonitor", "");
            strMQQMonitorLectura = funcion.getValueAppConfig("MQRecepResMonitor", "");

            intgModoMonitor = Convert.ToInt32(funcion.getValueAppConfig("intModoMonitor", ""));
            intgActv_FuncAuto = Convert.ToInt32(funcion.getValueAppConfig("intgActv_FuncAuto", ""));
            intgMonitor = Convert.ToInt32(funcion.getValueAppConfig("intMonitor", ""));

            inttmrRestar = funcion.getValueAppConfig("inttmrRestar", "");
            inttmrMonitor = funcion.getValueAppConfig("inttmrMonitor", "");
            inttmrBitacora = funcion.getValueAppConfig("inttmrBitacora", "");

            TiempoBitacoras = funcion.getValueAppConfig("TiempoBitacoras", "");
            TiempoMensajes = funcion.getValueAppConfig("TiempoMensajes", "");
            Tiempofuncionarios = funcion.getValueAppConfig("Tiempofuncionarios", "");
            TiempoAutorizacion = funcion.getValueAppConfig("TiempoAutorizacion", "");
       

            String CadenaRsultante, strPathServicio;
            strPathServicio = System.Reflection.Assembly.GetEntryAssembly().Location;
            CadenaRsultante = strPathServicio.Substring(strPathServicio.LastIndexOf("\\", strPathServicio.Length) + 1);

            LogPath = strPathServicio.Substring(0, strPathServicio.Length - CadenaRsultante.Length) + "\\Logs\\";

            //LogPath = funcion.getValueAppConfig("escribeArchivoLOG.LogPath", "");
            LogFile = funcion.getValueAppConfig("escribeArchivoLOG.LogFile", "");
            LogActivo = funcion.getValueAppConfig("LogActivo", "");

            PMONITOREOS = funcion.getValueAppConfig("PMONITOREOS", "");
            PARAMETRO1 = funcion.getValueAppConfig("PARAMETRO1", "");
            PARAMETRO2 = funcion.getValueAppConfig("PARAMETRO2", "");
            PARAMETRO3 = funcion.getValueAppConfig("PARAMETRO3", "");
            PARAMETRO4 = funcion.getValueAppConfig("PARAMETRO4", "");

            PROCESOS = funcion.getValueAppConfig("PROCESOS", "");
            PROCESO1 = funcion.getValueAppConfig("PROCESO1", "");
            PROCESO2 = funcion.getValueAppConfig("PROCESO2", "");
            PROCESO3 = funcion.getValueAppConfig("PROCESO3", "");
            PROCESO4 = funcion.getValueAppConfig("PROCESO4", "");
            PROCESO5 = funcion.getValueAppConfig("PROCESO5", "");
            PROCESO6 = funcion.getValueAppConfig("PROCESO6", "");

            Autorizaciones_Operatoria = funcion.getValueAppConfig("Autorizaciones Operatoria", "");
            Log_Apertura_de_Cuentas = funcion.getValueAppConfig("Log Apertura de Cuentas", "");
            LOG_Ordenes_de_Pago = funcion.getValueAppConfig("LOG Ordenes de Pago", "");
            LOG_Operaciones_CED = funcion.getValueAppConfig("LOG Operaciones CED", "");
            LOG_Operaciones_TDD = funcion.getValueAppConfig("LOG Operaciones TDD", "");

            Ambiente = funcion.getValueAppConfig("Ambiente", "TEST");
            Prefijo = funcion.getValueAppConfig("Prefijo", "P_");
            fileFuncionarios = funcion.getValueAppConfig("fileFuncionarios", "PFuncionarios.txt");
            fileSucursales = funcion.getValueAppConfig("fileSucursales", "PSucursales.txt");
            fileUsuarios = funcion.getValueAppConfig("fileUsuarios", "PUsuarios.txt");

            dirBitacora = funcion.getValueAppConfig("RutaBitacora", "");
            Bitacora = funcion.getValueAppConfig("Bitacora", "");

            FechaRestar = funcion.getValueAppConfig("RestarMonitor", "");
        }
        public bool Inicia()
        {
            //falta validación de entradas
            bool inicia;

            funcion.Escribe("", "entro a Inicia");

            //ModFunciones.date = DateTime.Now;
            //currentDate = ModFunciones.date.ToString(dateFormat);

            strMQManager = funcion.getValueAppConfig("MQManager", "");
            strMQQMonitorEscritura = funcion.getValueAppConfig("MQEnvioMsgMonitor", "");
            strMQQMonitorLectura = funcion.getValueAppConfig("MQRecepResMonitor", "");

            intgModoMonitor = Convert.ToInt32(funcion.getValueAppConfig("intModoMonitor", ""));
            intgActv_FuncAuto = Convert.ToInt32(funcion.getValueAppConfig("intgActv_FuncAuto", ""));
            intgMonitor = Convert.ToInt32(funcion.getValueAppConfig("intMonitor", ""));

            intgtmrRestar = Convert.ToInt32(inttmrRestar);
            intgtmrMonitor = Convert.ToInt32(inttmrMonitor);
            intgtmrBitacora = Convert.ToInt32(inttmrBitacora);

            TiempoBitacoras = funcion.getValueAppConfig("TiempoBitacoras", "");

            string[] vs = TiempoBitacoras.Split(',');
            string[] Temporal = vs;
            strFormatoTiempoBitacoras = Temporal[0];
            intTiempoBitacoras = Convert.ToInt32(Temporal[1]);

            TiempoMensajes = funcion.getValueAppConfig("TiempoMensajes", "");
            vs = TiempoMensajes.Split(',');
            Temporal = vs;
            strFormatoTiempoTKTMQ = Temporal[0];
            intTiempoTKTMQ = Convert.ToInt32(Temporal[1]);

            Tiempofuncionarios = funcion.getValueAppConfig("Tiempofuncionarios", "");
            vs = Tiempofuncionarios.Split(',');
            Temporal = vs;
            strFormatoTiempoFuncionarios = Temporal[0];
            intTiempoFuncionarios = Convert.ToInt32(Temporal[1]);

            TiempoAutorizacion = funcion.getValueAppConfig("TiempoAutorizacion", "");
            vs = TiempoAutorizacion.Split(',');
            Temporal = vs;
            strFormatoTiempoAutorizaciones = Temporal[0];
            intTiempoAutorizaciones = Convert.ToInt32(Temporal[1]);

            //FechaRestar = RestarMonitor;
            StrlogName = LogFile;
            StrLogPath = LogPath;
            StrlogActivo = LogActivo;

            if (String.IsNullOrEmpty(strMQManager))
            {
                inicia = false;
            }
            else
            {
                inicia = true;
            }

            return inicia;
        }
    }
}
