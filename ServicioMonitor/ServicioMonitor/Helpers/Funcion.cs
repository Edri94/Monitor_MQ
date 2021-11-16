using System;
using System.Configuration;
using System.Diagnostics;
using System.IO;
using System.Reflection;


namespace ServicioMonitor.Helpers
{
    public class Funcion
    {
        public string strLogPath;
        public Funcion()
        {
            String CadenaRsultante, strPathServicio;
            strPathServicio = System.Reflection.Assembly.GetEntryAssembly().Location;
            CadenaRsultante = strPathServicio.Substring(strPathServicio.LastIndexOf("\\", strPathServicio.Length) + 1);

            strLogPath = strPathServicio.Substring(0, strPathServicio.Length - CadenaRsultante.Length) + "\\Logs\\";
        }
        
        /// <summary>
        /// escribe en el log
        /// </summary>
        /// <param name="vData"></param>
        public virtual void Escribe(string vData, string tipo = "Mensaje")
        {
            StackTrace trace = new StackTrace(StackTrace.METHODS_TO_SKIP + 2);
            StackFrame frame = trace.GetFrame(0);
            MethodBase caller = frame.GetMethod();

            string clase = caller.ReflectedType.Name;
            string funcion = caller.Name;

            string seccion = "escribeArchivoLOG";       
            string nombre_archivo = DateTime.Now.ToString("ddMMyyyy") + "-" + getValueAppConfig("logFileName", seccion);
            nombre_archivo = nombre_archivo.Replace("@clase", clase);

            if (true)
            {
                using (StreamWriter outputFile = new StreamWriter(Path.Combine(getValueAppConfig(strLogPath, seccion), nombre_archivo), append: true))
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
        public virtual void Escribe(Exception ex, string tipo = "Error")
        {
            StackTrace trace = new StackTrace(StackTrace.METHODS_TO_SKIP + 2);
            StackFrame frame = trace.GetFrame(0);
            MethodBase caller = frame.GetMethod();

            string clase = caller.ReflectedType.Name;
            string funcion = caller.Name;

            string vData;
            string seccion = "escribeArchivoLOG";         
            string nombre_archivo = DateTime.Now.ToString("ddMMyyyy") + "-" + getValueAppConfig("logFileName", seccion);
            nombre_archivo = nombre_archivo.Replace("@clase", clase);

            if (true)
            {
                using (StreamWriter outputFile = new StreamWriter(Path.Combine(getValueAppConfig(strLogPath, seccion), nombre_archivo), append: true))
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

        /// <summary>
        /// Obtiene valor del parametro dado desde el app.config
        /// </summary>
        /// <param name="section">Seccion donde buscara</param>
        /// <param name="value">Valor que buscas</param>
        /// <returns></returns>
        public string getValueAppConfig(string key, string section = "")
        {
            if (section.Length >= 1)
            {
                return ConfigurationManager.AppSettings[$"{section}.{key}"];
            }
            else
            {
                return ConfigurationManager.AppSettings[$"{key}"];
            }

        }

        /// <summary>
        /// Obtiene fecha dependiendo el formato escogido
        /// </summary>
        /// <param name="formato"></param>
        /// <returns></returns>
        public string ObtenFechaFormato(int formato)
        {
            try
            {
                switch (formato)
                {
                    case 1:
                        return new DateTime().ToString("dd/MM/yyyy");
                    default:
                        return new DateTime().ToString("dd/MM/yyyy");
                }
            }
            catch (Exception error)
            {
                Escribe(error);
                return null;
            }
        }


        /// <summary>
        /// Escribe en el App.config en la seccion y key dada en parametros
        /// </summary>
        /// <param name="section">seccion en appsettings</param>
        /// <param name="key">key en appsetitngs</param>
        /// <param name="value">valor nuevo</param>
        public bool SetParameterAppSettings(string key, string value, string section = "")
        {
            //string nombre_appconfig = "MonitorMQTKT.exe.config";
            string nombre_appconfig = "App.config";

            bool bandera_archivo_existe = false;
            try
            {
                string appPath = System.IO.Path.GetDirectoryName(System.Reflection.Assembly.GetExecutingAssembly().Location);
                Escribe("Split aplicado a:   " + appPath);
                string[] appPath_arr = appPath.Split('\\');

                appPath = "";
                for (int i = 0; i < (appPath_arr.Length); i++)
                {
                    appPath = (i > 0) ? appPath + "\\" + appPath_arr[i] : appPath + appPath_arr[i];
                    string busqueda = $"{appPath}\\{nombre_appconfig}";
                    Escribe("Buscando:    " + busqueda);
                    bandera_archivo_existe = File.Exists(busqueda);
                    if (bandera_archivo_existe) break;
                }
                if (bandera_archivo_existe)
                {
                    appPath = appPath.Substring(1, appPath.Length - 1);
                    string configFile = System.IO.Path.Combine(appPath, nombre_appconfig);
                    ExeConfigurationFileMap configFileMap = new ExeConfigurationFileMap();
                    configFileMap.ExeConfigFilename = configFile;
                    System.Configuration.Configuration config = ConfigurationManager.OpenMappedExeConfiguration(configFileMap, ConfigurationUserLevel.None);
                    if (section.Length > 0)
                    {
                        config.AppSettings.Settings[$"{section}.{key}"].Value = value;
                    }
                    else
                    {
                        config.AppSettings.Settings[key].Value = value;
                    }
                    config.Save();
                    return true;
                }
                else
                {
                    Escribe("No se encontro el archivo", "Error");
                    return false;
                }

            }
            catch (Exception ex)
            {
                Escribe(ex, "Error");
                return false;
            }

        }

        /// <summary>
        /// Devuelve cadena igual a el numero de espacios dados en el parametro
        /// </summary>
        /// <param name="veces">numero de espacios</param>
        /// <returns></returns>
        public string Space(int veces)
        {
            return new String(' ', veces);
        }


        /// <summary>
        /// Devuelve una variante ( cadena ) que contiene un número especificado de caracteres del lado izquierdo de una cadena.
        /// </summary>
        /// <param name="cadena">Cadena</param>
        /// <param name="posiciones">Posiciones a tomar</param>
        /// <returns></returns>
        public string Left(string cadena, int posiciones)
        {
            return cadena.Substring(0, posiciones);
        }

        /// <summary>
        /// Devuelve una variante ( cadena ) que contiene un número específico de caracteres de una cadena.
        /// </summary>
        /// <param name="cadena"></param>
        /// <param name="start"></param>
        /// <param name="length"></param>
        /// <returns></returns>
        public string Mid(string cadena, int start, int length)
        {
            start--;
            return cadena.Substring(start, length);
        }
        /// <summary>
        /// Devuelve una variante ( cadena ) que contiene un número específico de caracteres del lado derecho de una cadena.
        /// </summary>
        /// <param name="cadena">Cadena</param>
        /// <param name="posiciones">Posiciones a tomar</param>
        /// <returns></returns>
        public string Right(string cadena, int posiciones)
        {
            return cadena.Substring((cadena.Length - posiciones), posiciones);
        }


        /// <summary>
        /// Se Actualiza el dato de configuracion con el app.config, solo para Alertas
        /// </summary>
        /// <param name="Key"></param>
        /// <param name="value"></param>
        /// <returns></returns>
        public bool UpdateAppSettings(string Key, string value)
        {
            try
            {
                //string s = ConfigurationManager.AppSettings[Key];
                var configFile = ConfigurationManager.OpenExeConfiguration(ConfigurationUserLevel.None);
                var settings = configFile.AppSettings.Settings;

                if (settings[Key] == null)
                {
                    return false;
                }
                else
                {
                    settings[Key].Value = value;
                }
                configFile.Save(ConfigurationSaveMode.Modified);
                ConfigurationManager.RefreshSection(configFile.AppSettings.SectionInformation.Name);
                return true;
            }
            catch (ConfigurationErrorsException)
            {
                Console.WriteLine("Error writing app settings");
                return false;
            }
        }

        /// <summary>
        /// Devuelve una variante ( larga ) que especifica la posición de la primera aparición de una cadena dentro de otra.
        /// </summary>
        /// <param name="inicio"></param>
        /// <param name="cadena1"></param>
        /// <param name="cadena2"></param>
        /// <returns></returns>
        public int InStr(int inicio, string cadena1, string cadena2)
        {
            try
            {
                int contador;
                string caracter;

                for (contador = 0; contador < cadena1.Length; contador++)
                {
                    caracter = cadena1.Substring(contador, 1);
                    if (caracter == cadena2 & contador >= inicio)
                    {
                        break;
                    }
                }

                return contador + 1;
            }
            catch (Exception)
            {
                return -1;
            }
        }
    }
}
