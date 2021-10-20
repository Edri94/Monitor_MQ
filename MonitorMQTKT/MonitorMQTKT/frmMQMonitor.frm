VERSION 5.00
Begin VB.Form Monitor 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   0  'None
   Caption         =   "TKTMQMonitor"
   ClientHeight    =   90
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   90
   ControlBox      =   0   'False
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmMQMonitor.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   90
   ScaleWidth      =   90
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Visible         =   0   'False
   Begin VB.Timer tmrMonitorMQTKT 
      Interval        =   10000
      Left            =   735
      Top             =   600
   End
   Begin VB.Timer tmrRestar 
      Enabled         =   0   'False
      Left            =   -15
      Top             =   -60
   End
End
Attribute VB_Name = "Monitor"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim MensajesMQ              As Double
Dim ErrorEjecucion          As Boolean

Dim ActivoProcFuncAuto       As Boolean      'Variable para determinar si se desea ejecutar el proceso del Monitoreo
Dim ModoMonitor             As Boolean      'Variable para determinar el modo de operacion del monitor


Dim DetallesShow            As Boolean      ' Variable para determinar el status de los detalles

'***** Para realizar el monitoreo de bitacoras
Dim miTipoMonitoreo As Integer
Dim miTotalMonitor As Integer

Dim ItemProceso As Integer
Dim ItemSeleccion As Integer

Dim bCambioManual As Boolean

Private Sub GuardarLog()
    Escribe "El siguiente reporte se genera a partir del botón 'Guardar el Registro de Operaciones' o cuando ha cambiado el dia de monitoreo o cuando se ha pulsado el botón 'Salir' del Monitor."
    Escribe "---------  Reporte del estado de los procesos  ---------"
    Escribe "*********  Registro de operaciones procesadas  *********"
    Escribe "   Respuestas (HOST->NT) registradas proceso de Monitoreo"
    'Escribe "       > Duración del CICLO[seg] : " & IntRecepResMonitor
    Escribe "   Solicitudes (NT->HOST) registradas proceso de Funcionarios y Autorizaciones"
    'Escribe "       > Duración del CICLO[min] : " & IntEnvioMsgMonitor
    Escribe "---------  Fin del reporte del estado de los procesos  ---------"
    Escribe ""
End Sub

Private Function ResetMonitor() As Boolean
    
On Error GoTo ErrResetMonitor
    
    ' Guardamos el estado del monitor
        
    Escribe "Respaldo del estado del monitor " & ObtenFechaFormato(1)
    
    GuardarLog
        
    'Desconecta del MQ Manager
    ' Objeto Queue para Leer la MQ Queue Respuestas de FUNCIONARIOS/AUTORIZACIONES
    If Not Nothing Is MQQMonitorLectura Then
        If MQQMonitorLectura.IsOpen Then Call MQCerrarCola(MQQMonitorLectura)
    End If
        
    ' Desconectando del MQManager
    Call MQDesconectar(mqsManager)
    
    ' Conectando con el MQManager
    If MQConectar(mqsConexion, strMQManager, mqsManager) Then
        blnConectado = True
    Else
        Escribe ""

        Escribe "Falla en Monitor < Error al conectarse con la MQ > : " & ObtenFechaFormato(1)
        Escribe "Detalles : " & mqsConexion.ReasonCode & " " & mqsConexion.ReasonName
        GoTo ErrResetMonitor
    End If
    
    Exit Function
    
ErrResetMonitor:
    Escribe "Error en la ejecución del reinicio de la aplicación: " & Err.Description
End Function

Private Sub Form_Load()

    strArchivoIni = App.Path & "\MonitorMQTKT.ini"

    If modFunciones.Inicia Then
        
        bCambioManual = False
        
        ' Establece el valor de la variable de control para determinar el modo de operación del monitor
        
        'Monitoreo sin ejecución
'Escribe "Antes del If intgModoMonitor = 1 Then"
        If intgModoMonitor = 1 Then
            ModoMonitor = True
        Else
            ModoMonitor = False
        End If
'Escribe "Despues del If intgModoMonitor = 1 Then"
        ' Establece el valor de la variable de control para determinar si se desea ejecutar el proceso funcionarios/autorizaciones
        
'Escribe "Antes del  If intgActv_FuncAuto = 1 Then"
        If intgActv_FuncAuto = 1 Then
            ActivoProcFuncAuto = True
        Else
            ActivoProcFuncAuto = False
        End If
'Escribe "Despues del  If intgActv_FuncAuto = 1 Then"
          
        Escribe ""
                
        Escribe "Aplicación Monitor iniciado : " & ObtenFechaFormato(1)
            
        'OBTIENE LA INFORMACION PARA REALIZAR EL MONITOREO DE BITACORAS
'Escribe "Antes de funcion CargaInfMonitoreo"
        CargaInfMonitoreo
'Escribe "DEspues de función CargaInfMonitoreo"
        
'Escribe "Antes de funcion Iniciar"
        Iniciar
'Escribe "Despues de funcion Iniciar"
    Else
        Escribe "No se puede continuar con la carga. Archivo Ini no existe."
        End
    End If
End Sub

Private Sub CargaInfMonitoreo()
    miTotalMonitor = GetProfile(strArchivoIni, "MONITOREO", "PMONITOREOS")
End Sub


Private Sub Form_Unload(Cancel As Integer)
    
    Detener
    
    ' Objeto Queue para Leer la MQ Queue Respuestas del Monitor
    If Not Nothing Is MQQMonitorLectura Then
        If MQQMonitorLectura.IsOpen Then Call MQCerrarCola(MQQMonitorLectura)
    End If
    
    Call MQDesconectar(mqsManager)
    
    ' Liberamos espacio en memoria
    Set mqsConexion = Nothing
    Set mqsManager = Nothing
    Set MQQMonitorLectura = Nothing
    
    'Salir
End Sub

Private Sub Detener()
On Error GoTo ErrorDetener
    
    'Cierra cola leer
    If Not Nothing Is MQQMonitorLectura Then
        If MQQMonitorLectura.IsOpen Then Call MQCerrarCola(MQQMonitorLectura)
    End If

    'Cierra la conexión
    If MQDesconectar(mqsManager) Then
        blnConectado = False
    End If
    
       
  Exit Sub
  
ErrorDetener:
    Escribe mqsConexion.ReasonCode & " " & mqsConexion.ReasonName
    Escribe ""
    
    Escribe "Falla en Monitor < Falla en el cierre de MQ-Series > : " & ObtenFechaFormato(1)
    Escribe "Detalles : " & mqsConexion.ReasonCode & " " & mqsConexion.ReasonName
End Sub

Private Sub Iniciar()
Dim iRow As Integer

On Error GoTo FallaInicio
    
    ' Logica para conectarse al Servidor MQSeries
    If MQConectar(mqsConexion, strMQManager, mqsManager) Then
      blnConectado = True
    Else
      Escribe ""
      Escribe "Falla en Monitor < Error al conectarse con la MQ > : " & ObtenFechaFormato(1)
      Escribe "Detalles : " & mqsConexion.ReasonCode & " " & mqsConexion.ReasonName
      ErrorEjecucion = True
      GoTo FallaInicio
    End If
      
    Escribe ""
      
    If ModoMonitor Then
          Escribe "Monitor iniciado en modo de monitoreo: " & ObtenFechaFormato(1)
          
          If Not ActivoProcFuncAuto Then        ' Ejecución Funcionarios/Autorizaciones no activada
              Escribe "El procesos de Funcionarios-Autorizaciones se encuentra en estado inactivo"
          Else                                  ' Ejecución Funcionarios/Autorizaciones activa
              Escribe "El procesos de Funcionarios-Autorizaciones se encuentra en estado activo"
          End If
    Else
          Escribe "Monitor iniciado en modo de procesamiento: " & ObtenFechaFormato(1)
    End If
    
    
    ' Activa el timer que controla la ejecución del restar
    tmrRestar.Enabled = True
    tmrRestar.Interval = intgtmrRestar * 1000
        
    ' Revisa que no haya diferencia en días menor
    If DateDiff("d", Date, CDate(FechaRestar)) <> 1 Then
       FechaRestar = ObtenFechaFormato(1, DateAdd("d", 1, Date))
    End If
        
    ' Almacena el valor del proximo restart del monitor en el archivo de ini
    SaveProfile strArchivoIni, "INTERTIEMPO", "RestarMonitor", FechaRestar
  Exit Sub
  
FallaInicio:
    
    Escribe "Error en la conexion con el servidor MQ: " & mqsConexion.ReasonCode & " " & mqsConexion.ReasonName
End Sub

Private Sub tmrMonitorMQTKT_Timer()
                        
    dblCiclosBitacoras = dblCiclosBitacoras + 10
    dblCiclosTKTMQ = dblCiclosTKTMQ + 10
    dblCiclosFuncionarios = dblCiclosFuncionarios + 10
    dblCiclosAutorizaciones = dblCiclosAutorizaciones + 10
    
    If intgMonitor = 1 Then
        If strFormatoTiempoBitacoras <> "S" Then
            If dblCiclosBitacoras >= (intTiempoBitacoras * 60) Then
                tmrBitacora
                dblCiclosBitacoras = 0
            End If
        Else
            If dblCiclosBitacoras >= intTiempoBitacoras Then
                tmrBitacora
                dblCiclosBitacoras = 0
            End If
        End If
    End If
    
    If strFormatoTiempoTKTMQ <> "S" Then
        If dblCiclosTKTMQ >= (intTiempoTKTMQ * 60) Then
'Escribe "Antes de entrar a la funcion tmrTKTMQ en Minutos"
            tmrTKTMQ
'Escribe "Despues de entrar a la funcion tmrTKTMQ en Minutos"
            dblCiclosTKTMQ = 0
        End If
    Else
        If dblCiclosTKTMQ >= intTiempoTKTMQ Then
'Escribe "Antes de entrar a la funcion tmrTKTMQ en Segundos"
            tmrTKTMQ
'Escribe "Despues de entrar a la funcion tmrTKTMQ en Segundos"
            dblCiclosTKTMQ = 0
        End If
    End If
    
    If strFormatoTiempoFuncionarios <> "S" Then
        If dblCiclosFuncionarios >= (intTiempoFuncionarios * 60) Then
            Call ActivarEnvioFuncAuto("F")
            dblCiclosFuncionarios = 0
        End If
    Else
        If dblCiclosFuncionarios >= intTiempoFuncionarios Then
            Call ActivarEnvioFuncAuto("F")
            dblCiclosFuncionarios = 0
        End If
    End If
    
    If strFormatoTiempoAutorizaciones <> "S" Then
        If dblCiclosAutorizaciones >= (intTiempoAutorizaciones * 60) Then
            Call ActivarEnvioFuncAuto("A")
            dblCiclosAutorizaciones = 0
        End If
    Else
        If dblCiclosAutorizaciones >= intTiempoAutorizaciones Then
            Call ActivarEnvioFuncAuto("A")
            dblCiclosAutorizaciones = 0
        End If
    End If
End Sub

Private Sub tmrRestar_Timer()

    tmrRestar.Enabled = False
    
    If Date > CDate(FechaRestar) Then
    
        ' Ejecuta un restar del monitor
        ResetMonitor
        
        ' Establece la fecha del siguiente restar
        FechaRestar = Date
        
        'Guarda el valor en log
        SaveProfile strArchivoIni, "INTERTIEMPO", "RestarMonitor", FechaRestar
        
        Escribe "Aplicación Monitor iniciado : " & ObtenFechaFormato("1")
        Escribe "Monitor iniciado en modo de procesamiento: " & ObtenFechaFormato("1")
        
    End If
    
    tmrRestar.Enabled = True

End Sub

Private Sub tmrTKTMQ()
    'Escribe "Entro al sub tmrTKTMQ"
    Dim ln_MsgEncontrados As Double
    'Escribe "Antes del llamado de la funcion RevisaMQ"
    ln_MsgEncontrados = RevisaMQ(MQQMonitorLectura, strMQManager, strMQQMonitorLectura, strMQQMonitorEscritura, "0", ActivoProcFuncAuto)
    'Escribe "Despues del llamado de la funcion RevisaMQ"
            
End Sub

Private Sub ActivarEnvioFuncAuto(psMonitor As String)

    Dim MensajesMQ As Object
    
    Dim Ld_CodigoExecNTHOST     As Double
    Dim EjecutableMSG           As String
    Dim LsProceso               As String
    
    Dim sMensaje As String
    
    Select Case psMonitor
        Case "A"
            LsProceso = "PROCESOS2"
        Case "F"
            LsProceso = "PROCESOS1"
    End Select
    
    ' Armado de los parametros que se incluiran para el ejecutable
   
    If ActivoProcFuncAuto Then

        sMensaje = Mid(GetProfile(strArchivoIni, "CONFIGURACION", LsProceso), 3, InStr(3, GetProfile(strArchivoIni, "CONFIGURACION", LsProceso), ",") - 3)
        If fValidaEjecucion(sMensaje) Then
            Set MensajesMQ = CreateObject("MensajesMQ.cMensajes")
            MensajesMQ.ProcesarMensajes App.Path, strMQManager & "-" & strMQQMonitorEscritura & "-" & psMonitor
            Set MensajesMQ = Nothing
        Else
            Escribe "La operación: " & sMensaje & " no esta habilitada para este día " & ObtenFechaFormato(1)
        End If
    End If
End Sub

Private Function RevisaMQ(ByRef MQParaLectura As MQQueue, _
                          ByRef MQManager As String, _
                          ByRef MQQLectura As String, _
                          ByRef MQQEscritura As String, _
                          ByRef psOtros As String, _
                          ByRef StatusProceso As Boolean) As Double
                          
    Dim lngErr      As Long
    Dim j           As Integer
    Dim lngMQOpen   As Long
    Dim lsExeParam  As String
  
On Error GoTo ErrorRevisaMQ

'Escribe "Entro a la funcion RevisaMQ"

    'Abre cola leer
'Escribe "Asigna valor a la variable lngMQOpen"
    lngMQOpen = MQOO_INQUIRE         ' Permite leer las propiedades de la QUEUE
'Escribe "Asigna valor a la variable lngMQOpen: " & lngMQOpen
    
'Escribe "Antes de validar si se puede o no abrir la Queue, a la Queue MQQLectura:" & MQQLectura
    
    If MQAbrirCola(mqsManager, MQQLectura, MQParaLectura, lngMQOpen) Then
'Escribe "Si pudo abrirla correctamente"
        j = 1
'Escribe "Antes de obtener el status del queue manager"
        lngErr = mqsConexion.ReasonCode
'Escribe "Obtiene el status del queue manager a traves de mqsConexion.ReasonCode= " & lngErr

'Escribe "Antes de obtener el numero de mensajes"
        ' Obtención del numero de mensajes en la QUEUE
        MensajesMQ = MQParaLectura.CurrentDepth
'Escribe "Despues de obtener el numero de mensajes con MQParaLectura.CurrentDepth: " & MensajesMQ

'Escribe "Si hay mensajes se continua"
        ' si el número de mensajes es mayor a cero continuamos
        If MensajesMQ < 0 Then
            RevisaMQ = 0
            Exit Function
        End If
        
'Escribe "Se continuo y se cierra nuevamente la queue"
        'Cierra cola leer
        Call MQCerrarCola(MQParaLectura)
'Escribe "Cerro la queue"

        If Not ModoMonitor Then        ' Monitor en modo de ejecución
                
                'Creamos una instancia de TKTMQ por cada mensaje leido
                If MensajesMQ > 0 Then
'Escribe "Antes del ciclo while"
                    Do While j <= MensajesMQ
                        ' Todo ejecutable es acompañado por parametros
                        ' (MQQLectura, MQQEscritura, otros datos)

                        If StrComp(psOtros, "", vbTextCompare) <> 0 Then
                            lsExeParam = MQManager & "-" & MQQLectura & "-" & MQQEscritura & "-" & psOtros
                        Else
                            lsExeParam = MQManager & "-" & MQQLectura & "-" & MQQEscritura & "-0"
                        End If
'Escribe "Antes de la Declaración de la variable para crear el objeto de TKTMQ"
                        Dim TKTMQ As Object
'Escribe "Despues de la Declaración de la variable para crear el objeto de TKTMQ"

'Escribe "Antes de instanciar la variable TKTMQ"
                        Set TKTMQ = CreateObject("TKTMQ.cProcesaMSG")
'Escribe "Despues de instanciar la variable TKTMQ"

'Escribe "Antes de llamar a la clase ProcesarMensaje de la dll, Enviandole los siguientes valores"
'Escribe "App.Path:" & App.Path & " #### lsExeParam:" & lsExeParam
                        TKTMQ.ProcesarMensaje App.Path, lsExeParam
'Escribe "Despues de llamar a la clase ProcesarMensaje de la dll"

'Escribe "Antes de destruir el objeto dll"
                        Set TKTMQ = Nothing
'Escribe "Despues de destruir el objeto dll"
                        j = j + 1
                        
                    Loop
'Escribe "Despues del ciclo while"
                End If
        End If
    Else
        Escribe "Error en la conexion con el MQ Manager : " & CStr(mqsConexion.ReasonCode) & " " & mqsConexion.ReasonName
        Escribe "Ejecutamos el reinicio del monitor por problemas en la comunicacion con MQManager " & strMQManager
        ReConectar
        Exit Function
    End If
    
    RevisaMQ = MensajesMQ
    
    Exit Function
    
ErrorRevisaMQ:
    Escribe ""
    Escribe "Monitor error al ejecutar el proceso de TKTMQ el " & ObtenFechaFormato(1)
    RevisaMQ = -1
End Function

Private Sub tmrBitacora()
Dim Ld_CodigoExecNTHOST     As Double
Dim EjecutableMSG           As String
Dim icont As Integer
Dim Ejecutable As String
Dim Parametro() As String
Dim intlBitacoras As Integer
Dim sValor As String
Dim vntBitacora() As String

    For intlBitacoras = 1 To miTotalMonitor
        sValor = GetProfile(strArchivoIni, "MONITOREO", "Parametro" & intlBitacoras)
        vntBitacora = Split(sValor, ",")
        If vntBitacora(0) = 1 Then
            If fValidaEjecucion(vntBitacora(1)) Then
                Parametro = Split(GetProfile(strArchivoIni, "CAPTION_VS_EXE", vntBitacora(1)), ",")
                Ejecutable = Parametro(0)
                
                If Ejecutable = "M" Then
                    Dim MensajesMQ As Object
                    Set MensajesMQ = CreateObject("MensajesMQ.cMensajes")
                    MensajesMQ.ProcesarMensajes App.Path, strMQManager & "-" & strMQQMonitorEscritura & "-" & "1" & "-" & Parametro(1)
                    Set MensajesMQ = Nothing
                Else
                    Dim Bitacoras As Object
                    Set Bitacoras = CreateObject("Bitacoras.cBitacoras")
                    Bitacoras.ProcesarBitacora App.Path, strMQManager & "-" & strMQQMonitorEscritura & "-" & "1" & "-" & Parametro(1)
                    Set Bitacoras = Nothing
                End If
                
            Else
                Escribe "La operación: " & vntBitacora(1) & " no esta habilitada para este día " & ObtenFechaFormato(1)
            End If
        End If
    Next
End Sub

Private Function fValidaEjecucion(ByVal psBitacora As String) As Boolean
Dim iTotalProcesos As String
Dim iRow As Integer
Dim sValor As String
Dim sParametros() As String
Dim intCuenta As Integer
On Error GoTo error

    fValidaEjecucion = False
    
    iTotalProcesos = GetProfile(strArchivoIni, "CONFIGURACION", "PPROCESOS")
    For iRow = 1 To iTotalProcesos
        sValor = GetProfile(strArchivoIni, "CONFIGURACION", "PROCESOS" & iRow)
        sParametros = Split(sValor, ",")
        For intCuenta = 0 To sParametros(9)
            If sParametros(0) = 1 And sParametros(1) = psBitacora Then
                If sParametros(intCuenta + 2) = "Si" Then
                        If Weekday(Date, vbMonday) = intCuenta + 1 Then
                            fValidaEjecucion = True
                            intCuenta = sParametros(9)
                            iRow = iTotalProcesos
                        End If
                End If
            Else
                 fValidaEjecucion = False
                 intCuenta = sParametros(9)
            End If
        Next
    Next

Exit Function

error:
    Escribe Err.Description
    fValidaEjecucion = False
    
End Function

Public Sub ReConectar()
   
On Error GoTo ErrReConectar
    
    'Desconecta del MQ Manager
    ' Objeto Queue para Leer la MQ Queue Respuestas de FUNCIONARIOS/AUTORIZACIONES
    If Not Nothing Is MQQMonitorLectura Then
        If MQQMonitorLectura.IsOpen Then Call MQCerrarCola(MQQMonitorLectura)
    End If
        
    ' Desconectando del MQManager
    Call MQDesconectar(mqsManager)
    
    ' Conectando con el MQManager
    If MQConectar(mqsConexion, strMQManager, mqsManager) Then
        blnConectado = True
    Else
        Escribe ""
        Escribe "Falla en Monitor < Error al reconectarse con la MQ > : " & ObtenFechaFormato(1)
        Escribe "Detalles : " & mqsConexion.ReasonCode & " " & mqsConexion.ReasonName
        GoTo ErrReConectar
    End If
    
    Exit Sub
    
ErrReConectar:
    Escribe "Error en la ejecución del reconexión del Monitor. " & Err.Description
End Sub
