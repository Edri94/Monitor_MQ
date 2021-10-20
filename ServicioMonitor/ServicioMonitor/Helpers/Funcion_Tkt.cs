
using System;
using System.Diagnostics;
using System.IO;
using System.Reflection;

namespace ServicioMonitor.Helpers
{
    public class Funcion_Tkt : Funcion
    {

        /// <summary>
        /// escribe en el log
        /// </summary>
        /// <param name="vData"></param>
        public override void Escribe(string vData, string tipo = "Mensaje")
        {
            StackTrace trace = new StackTrace(StackTrace.METHODS_TO_SKIP + 2);
            StackFrame frame = trace.GetFrame(0);
            MethodBase caller = frame.GetMethod();

            string clase = "Tkt";
            string funcion = caller.Name;

            string seccion = "escribeArchivoLOG";
            string nombre_archivo = DateTime.Now.ToString("ddMMyyyy") + "-" + getValueAppConfig("logFileName", seccion);
            nombre_archivo = nombre_archivo.Replace("@clase", clase);

            if (true)
            {
                using (StreamWriter outputFile = new StreamWriter(Path.Combine(getValueAppConfig("logFilePath", seccion), nombre_archivo), append: true))
                {
                    vData = $"[{DateTime.Now.ToString("dd-MM-yyyy hh:mm:ss")}]  {tipo} desde {funcion}:  {vData}";
                    Console.WriteLine(vData);
                    outputFile.WriteLine(vData);
                }

            }
        }

        /// <summary>
        /// escribe en el log
        /// </summary>
        /// <param name="vData"></param>
        public override void Escribe(Exception ex, string tipo = "Error")
        {
            StackTrace trace = new StackTrace(StackTrace.METHODS_TO_SKIP + 2);
            StackFrame frame = trace.GetFrame(0);
            MethodBase caller = frame.GetMethod();

            string clase = "Tkt";
            string funcion = caller.Name;

            string vData;
            string seccion = "escribeArchivoLOG";
            string nombre_archivo = DateTime.Now.ToString("ddMMyyyy") + "-" + getValueAppConfig("logFileName", seccion);
            nombre_archivo = nombre_archivo.Replace("@clase", clase);

            if (true)
            {
                using (StreamWriter outputFile = new StreamWriter(Path.Combine(getValueAppConfig("logFilePath", seccion), nombre_archivo), append: true))
                {
                    vData = $"[{DateTime.Now.ToString("dd-MM-yyyy hh:mm:ss")}] {(char)13}" +
                        $"*{tipo} desde {funcion}:  {ex.Message} {(char)13}" +
                        $"*InnerException: {ex.InnerException} {(char)13}" +
                        $"*Source: {ex.Source}  {(char)13}" +
                        $"*Data: {ex.Data}  {(char)13}" +
                        $"*HelpLink: {ex.HelpLink}  {(char)13}" +
                        $"*StackTrace: {ex.StackTrace}  {(char)13}" +
                        $"*HResult: {ex.HResult}  {(char)13}" +
                        $"*TargetSite: {ex.TargetSite}  {(char)13}";
                    Console.Write(vData);
                    outputFile.WriteLine(vData);
                }

            }
        }
    }
}
