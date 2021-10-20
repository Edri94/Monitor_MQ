Attribute VB_Name = "modFunciones"
Option Explicit

' Funciones para leer y escribir archivos INI
Public Declare Function GetPrivateProfileString Lib "kernel32" Alias "GetPrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpDefault As String, ByVal lpReturnedString As String, ByVal nSize As Long, ByVal lpFileName As String) As Long
Public Declare Function WritePrivateProfileString Lib "kernel32" Alias "WritePrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpString As Any, ByVal lpFileName As String) As Long

' Variable para el archivo .ini
Public strArchivoIni As String
Dim ArchivoLog    As String

'Enumeración para las opciones de abrir la cola
Enum MQOPEN
    MQOO_INPUT_AS_Q_DEF = &H1
    MQOO_INPUT_SHARED = &H2
    MQOO_INPUT_EXCLUSIVE = &H4
    MQOO_BROWSE = &H8
    MQOO_OUTPUT = &H10
    MQOO_INQUIRE = &H20
    MQOO_SET = &H40
    MQOO_BIND_ON_OPEN = &H4000
    MQOO_BIND_NOT_FIXED = &H8000
    MQOO_BIND_AS_Q_DEF = &H0
    MQOO_SAVE_ALL_CONTEXT = &H80
    MQOO_PASS_IDENTITY_CONTEXT = &H100
    MQOO_PASS_ALL_CONTEXT = &H200
    MQOO_SET_IDENTITY_CONTEXT = &H400
    MQOO_SET_ALL_CONTEXT = &H800
    MQOO_ALTERNATE_USER_AUTHORITY = &H1000
    MQOO_FAIL_IF_QUIESCING = &H2000
End Enum


' Declaraciones de los objetos para MQSeries
' Referencia: IBM MQSeries Automation Classes for ActiveX
Public mqsConexion          As New MQSession    ' Objeto Session para conexión con el servidor MQSeries
Public mqsManager           As MQQueueManager   ' Objeto QueueManager para accesar al maestro de colas
Public MQQMonitorLectura       As MQQueue          ' Objeto Queue para Leer la MQ Queue Respuestas de Funcionarios


' Variables de parametro para MQSeries
Public strMQManager            As String   ' Nombre del MQ Manager
Public strMQQMonitorLectura    As String   ' Nombre de la MQ Queue para Leer las respuestas de una solicitud de Funcionarios
Public strMQQMonitorEscritura  As String   ' Nombre de la MQ Queue para Escribir una solicitud de Funcionarios


' Variables para controlar los periodos de monitoreo
'Public StrTipoLapso         As String       ' Valor para controlar el tipo de lapso que se usara para el monitor (Funcionarios)
'Public IntRecepResMonitor   As Integer      ' Valor para disparar el monitor HOST >> NT (Funcionarios)
'Public IntEnvioMsgMonitor   As Integer      ' Valor para disparar el monitor NT >> HOST (Funcionarios)
Public FechaRestar          As String       ' Valor que determina la fecha y hora para ejecutar un restar del monitor

' Variables para controlar los periodos de monitoreo
Public strLogPath           As String       ' Ruta para almacenar el log del monitor
Public strlogName           As String       ' Nombre y extención del archivo log

' Variable para validar la conexión
Public blnConectado         As Boolean

'///////////////////////////////////////////////////////////////////////////////////
'///////////////////////////////////////////////////////////////////////////////////

'Variables nuevas para convertirlo a Servicio
Public intgModoMonitor      As Integer
Public intgActv_FuncAuto    As Integer
Public intgMonitor          As Integer

Public intgtmrRestar        As Integer
Public intgtmrMonitor       As Integer
Public intgtmrBitacora      As Integer

Public strFormatoTiempoBitacoras       As String
Public strFormatoTiempoTKTMQ           As String
Public strFormatoTiempoFuncionarios    As String
Public strFormatoTiempoAutorizaciones  As String

Public intTiempoBitacoras       As Integer
Public intTiempoTKTMQ           As Integer
Public intTiempoFuncionarios    As Integer
Public intTiempoAutorizaciones  As Integer

Public dblCiclosBitacoras      As Double
Public dblCiclosTKTMQ          As Double
Public dblCiclosFuncionarios   As Double
Public dblCiclosAutorizaciones As Double

'///////////////////////////////////////////////////////////////////////////////////
'///////////////////////////////////////////////////////////////////////////////////

Public Function Inicia() As Boolean
    
    Dim Temporal() As String
            
'Escribe "Inicio de Inicia"
    'Verifica si existe el archivo
'Escribe "Valida archivo ini"
    
    If ExisteArchivo(strArchivoIni) Then
    
'Escribe "INI: Dentro del if"
'Escribe "INI: Antes de obtener los valores del ini"
       
       strMQManager = GetProfile(strArchivoIni, "MQSeries", "MQManager")
'Escribe "INI: Obtiene el valor de strMQManager:" & strMQManager
       
       strMQQMonitorEscritura = GetProfile(strArchivoIni, "MQSeries", "MQEnvioMsgMonitor")
'Escribe "INI: Obtiene el valor de strMQQMonitorEscritura:" & strMQQMonitorEscritura
       
       strMQQMonitorLectura = GetProfile(strArchivoIni, "MQSeries", "MQRecepResMonitor")
'Escribe "INI: Obtiene el valor de strMQQMonitorLectura:" & strMQQMonitorLectura
       
              
       'Si existe obtiene los parametros para configurar los disparos de los procesos
       'StrTipoLapso = GetProfile(strArchivoIni, "INTERTIEMPO", "TipoLapso")
'Escribe "INI: Obtiene el valor de StrTipoLapso:" & StrTipoLapso
       
       'IntRecepResMonitor = Val(GetProfile(strArchivoIni, "INTERTIEMPO", "LapRespMonitor"))
'Escribe "INI: Obtiene el valor de IntRecepResMonitor:" & IntRecepResMonitor
       
       'IntEnvioMsgMonitor = Val(GetProfile(strArchivoIni, "INTERTIEMPO", "LapEnvioMsgMonitor"))
'Escribe "INI: Obtiene el valor de IntEnvioMsgMonitor:" & IntEnvioMsgMonitor
                   
       'Servicio
        intgModoMonitor = GetProfile(strArchivoIni, "SERVICIO", "intModoMonitor")
'Escribe "INI: Obtiene el valor de intgModoMonitor:" & intgModoMonitor
        
        intgActv_FuncAuto = GetProfile(strArchivoIni, "SERVICIO", "intActv_FuncAuto")
'Escribe "INI: Obtiene el valor de intgActv_FuncAuto:" & intgActv_FuncAuto
        
        intgMonitor = GetProfile(strArchivoIni, "SERVICIO", "intMonitor")
'Escribe "INI: Obtiene el valor de intgMonitor:" & intgMonitor
        
        intgtmrRestar = GetProfile(strArchivoIni, "SERVICIO", "inttmrRestar")
'Escribe "INI: Obtiene el valor de intgtmrRestar:" & intgtmrRestar
        
        intgtmrMonitor = GetProfile(strArchivoIni, "SERVICIO", "inttmrMonitor")
'Escribe "INI: Obtiene el valor de intgtmrMonitor:" & intgtmrMonitor
        
        intgtmrBitacora = GetProfile(strArchivoIni, "SERVICIO", "inttmrBitacora")
'Escribe "INI: Obtiene el valor de intgtmrBitacora:" & intgtmrBitacora
               
        Temporal = Split(GetProfile(strArchivoIni, "PARAMETROSTIEMPO", "TiempoBitacoras"), ",")
        
        strFormatoTiempoBitacoras = Temporal(0)
'Escribe "INI: Obtiene el valor de strFormatoTiempoBitacoras:" & strFormatoTiempoBitacoras
        
        intTiempoBitacoras = Temporal(1)
'Escribe "INI: Obtiene el valor de intTiempoBitacoras:" & intTiempoBitacoras
        
        Temporal = Split(GetProfile(strArchivoIni, "PARAMETROSTIEMPO", "TiempoMensajes"), ",")

        
        strFormatoTiempoTKTMQ = Temporal(0)
'Escribe "INI: Obtiene el valor de strFormatoTiempoTKTMQ:" & strFormatoTiempoTKTMQ
        
        intTiempoTKTMQ = Temporal(1)
'Escribe "INI: Obtiene el valor de intTiempoTKTMQ:" & intTiempoTKTMQ
        
        Temporal = Split(GetProfile(strArchivoIni, "PARAMETROSTIEMPO", "TiempoFuncionarios"), ",")

        
        strFormatoTiempoFuncionarios = Temporal(0)
'Escribe "INI: Obtiene el valor de strFormatoTiempoFuncionarios:" & strFormatoTiempoFuncionarios
        
        intTiempoFuncionarios = Temporal(1)
'Escribe "INI: Obtiene el valor de intTiempoFuncionarios:" & intTiempoFuncionarios
                
        Temporal = Split(GetProfile(strArchivoIni, "PARAMETROSTIEMPO", "TiempoAutorizaciones"), ",")

        
        strFormatoTiempoAutorizaciones = Temporal(0)
'Escribe "INI: Obtiene el valor de strFormatoTiempoAutorizaciones :" & strFormatoTiempoAutorizaciones
        
        intTiempoAutorizaciones = Temporal(1)
'Escribe "INI: Obtiene el valor de intTiempoAutorizaciones:" & intTiempoAutorizaciones
        
        FechaRestar = GetProfile(strArchivoIni, "INTERTIEMPO", "RestarMonitor")
'Escribe "INI: Obtiene el valor de FechaRestar:" & FechaRestar
        
        
        strlogName = GetProfile(strArchivoIni, "LOGDATA", "LogFile")
'Escribe "INI: Obtiene el valor de strlogName:" & strlogName
        
       'Si existe obtiene los parametros para configurar el log del monitor
        strLogPath = GetProfile(strArchivoIni, "LOGDATA", "LogPath")
'Escribe "INI: Obtiene el valor de strLogPath:" & strLogPath
               
        
        Inicia = True
    Else
        Inicia = False
    End If
    
End Function


Public Function MQConectar(ByRef objMQConexion As MQSession, ByVal strMQManager As String, ByRef objMQManager As MQQueueManager) As Boolean
    On Error GoTo ErrMQConectar
    
    ' Se crear una sesión con el servidor de MQSeries
    ' y se accesa al Queue Manager.
    Set objMQManager = objMQConexion.AccessQueueManager(strMQManager)
    
    MQConectar = True
    Exit Function
    
ErrMQConectar:
    MQConectar = False
End Function

Public Function MQDesconectar(ByRef objMQManager As MQQueueManager) As Boolean
    On Error GoTo ErrMQDesconectar
    
    ' Verificamos si existe el objeto
    If Not Nothing Is objMQManager Then
        ' Si esta conectado
        If objMQManager.IsConnected Then
            ' Se desconecta
            objMQManager.Disconnect
            MQDesconectar = True
        End If
    End If
    
    Exit Function
    
ErrMQDesconectar:
    MQDesconectar = False
End Function

Public Function MQAbrirCola(ByRef objMQManager As MQQueueManager, ByVal strMQCola As String, ByRef objMQCola As MQQueue, ByVal lngOpciones As MQOPEN) As Boolean
    On Error GoTo ErrMQAbrirCola
    
    ' Se accesa la cola ya sea para leer o escribir
    Set objMQCola = objMQManager.AccessQueue(strMQCola, lngOpciones, mqsManager.Name, "AMQ.*")
    
    MQAbrirCola = True
    Exit Function
    
ErrMQAbrirCola:
    MQAbrirCola = False
End Function

Public Function MQCerrarCola(ByRef objMQCola As MQQueue) As Boolean
    On Error GoTo ErrMQCerrarCola
    
    ' Vaerificamos si existe el objeto
    If Not Nothing Is objMQCola Then
        ' Si esta abierta la cola
        If objMQCola.IsOpen Then
            ' Se cierra
            objMQCola.Close
            MQCerrarCola = True
        End If
    End If
    
    Exit Function
    
ErrMQCerrarCola:
    MQCerrarCola = False
End Function

'Función para verificar si existe un archivo
Public Function ExisteArchivo(strArchivo As String) As Boolean
       On Error GoTo NoExiste
       
       Dim lngArch As Long
              
       lngArch = GetAttr(strArchivo)
       ExisteArchivo = True
       Exit Function
NoExiste:
       ExisteArchivo = False
End Function

'Función para obtener un valor de un archivo ini
Public Function GetProfile(lpFileName As String, lpAppName As String, lpKeyName As String, Optional lpString As String = "") As String
       Dim retval As String, worked As Long
        
       retval = String$(255, 0)
        
       worked = GetPrivateProfileString(lpAppName, lpKeyName, lpString, retval, Len(retval), lpFileName)
       If worked = 0 Then
          GetProfile = lpString
       Else
          GetProfile = Left(retval, worked)
       End If
End Function

'Función para guardar un valor de un archivo ini
Public Sub SaveProfile(lpFileName As String, lpAppName As String, lpKeyName As String, lpString As String)
       Dim Ret As Long
       Ret = WritePrivateProfileString(lpAppName, lpKeyName, lpString, lpFileName)
End Sub


Public Function ObtenFechaFormato(ByVal strFormato As String, _
                                    Optional ByVal strFecha As String, _
                                    Optional ByVal strMeridiano As String, _
                                    Optional ByVal strTiempo As String) As String

    Dim strDia          As String, strMes       As String, strAnio      As String
    Dim strHora         As String, strMinuto    As String, strSegundo   As String
    Dim strMeridian     As String, Fecha        As String, strTemporal  As String
    Dim Tiempo          As String
    
    If strFecha = "" Then
        Fecha = Date & " " & Time
        'Tiempo = Time
    Else
        If Len(strFecha) < 11 Then
            Tiempo = " " & Time
        Else
            Tiempo = ""
        End If
        
        Fecha = strFecha & Tiempo
    End If
        
    strDia = DatePart("d", Fecha)
    strDia = IIf(Len(strDia) = 1, "0" & strDia, strDia)
    strMes = DatePart("m", Fecha)
    strMes = IIf(Len(strMes) = 1, "0" & strMes, strMes)
    strAnio = DatePart("yyyy", Fecha)
    
    strHora = String(2 - Len(DatePart("h", Fecha)), "0") & DatePart("h", Fecha)
    strMinuto = String(2 - Len(DatePart("s", Fecha)), "0") & DatePart("s", Fecha)
    strSegundo = String(2 - Len(DatePart("n", Fecha)), "0") & DatePart("n", Fecha)

    
    'If Len(Time) > 8 Then
        'If Len(strFecha) > 10 Then
            strMeridian = Mid(Fecha, InStr(1, Fecha, " ") + 10)
        'Else
        '    strMeridian = Mid(Time, InStr(1, Time, " ") + 1)
        'End If
    'End If
    
    Select Case strFormato
        Case "1" 'yyyy/MM/dd hh:mm:ss
            strTemporal = strAnio & "/" & strMes & "/" & strDia & " " & _
                          strHora & ":" & strMinuto & ":" & strSegundo
        Case "2"
            strTemporal = strDia & "-" & strMes & "-" & strAnio & " " & _
                          strHora & ":" & strMinuto & ":" & strSegundo
        Case "3"
            strTemporal = strAnio & strMes & strDia
    End Select
    
    If strMeridiano <> "" And strMeridian <> "" Then
        strTemporal = strTemporal & " " & strMeridian
    End If
    
    ObtenFechaFormato = strTemporal
    
End Function

Public Function Escribe(ByVal vData As String)

    If strLogPath = "" Then
        strLogPath = App.Path & "\"
    End If
    If strlogName = "" Then
        strlogName = "Carga.log"
    End If
    
    ArchivoLog = strLogPath & ObtenFechaFormato(3) & "-" & strlogName
    
    Open ArchivoLog For Append As #1
    Print #1, vData
    Close #1
End Function

