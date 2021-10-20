Attribute VB_Name = "modMQ"
Option Explicit

Public ParamEncripcion  As MNICript.clsEncripta 'Objeto de Encripción
Public cnnConexion As Connection
Public rssRegistro As Recordset
Public Mb_Detalles     As Boolean
Public gstrRutaIni      As String
Public gsPswdDB         As String
Public gsUserDB         As String
Public gsNameDB         As String
Public gsCataDB         As String
Public gsDSNDB          As String

Public strQuery         As String               'Cadena para almacenar el Query a ejecutarse en la base de datos

' Variables para configuración del archivo log
Public lsLogPath       As String
Public lsLogName       As String

Private Archivo         As String

Public Mb_GrabaLog     As Boolean

'Public gsAccesoActual   As String               'Fecha/Hora actual del sistema. La tomamos del servidor NT y no de SQL porque precisamente el
                                                'requerimiento de la fecha/hora fue porque a veces se pierde la conexión a SQL

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

'Enumeración para el tipo de acción
Enum TipoAccion
    eMQConectar = 0
    eMQDesconectar = 1
    eMQAbrirCola = 2
    eMQCerrarCola = 3
    eMQLeerCola = 4
    eMQEscribirCola = 5
    eMQOtro = 6
End Enum

' Variable para validar la conexión
Public blnConectado As Boolean

'******************************************************************************************
'Variables y objectos publicos para conectarse al MQSeries

' Declaraciones de los objetos para MQSeries
' Referencia: IBM MQSeries Automation Classes for ActiveX
Public mqsConexion       As New MQSession       ' Objeto Session para conexión con el servidor MQSeries
Public mqsManager        As MQQueueManager  ' Objeto QueueManager para accesar al maestro de colas
Public mqsEscribir       As MQQueue             ' Objeto Queue para escribir
Public mqsLeer           As MQQueue             ' Objeto Queue para leer
Public mqsMsgEscribir    As MQMessage           ' Objeto Message para escribir
Public mqsMsgLeer        As MQMessage           ' Objeto Message para leer
Public mqsMsgReporte     As MQMessage           ' Objeto Message para reportes

Public gsError As String                        'Cadena que contiene un mensaje de error


Public Function MQConectar(ByRef objMQConexion As MQSession, _
                           ByVal strMQManager As String, _
                           ByRef objMQManager As MQQueueManager) As Boolean
                           
    On Error GoTo ErrMQConectar
    
    ' Se crear una sesión con el servidor de MQSeries
    ' y se accesa al Queue Manager.
    Set objMQManager = objMQConexion.AccessQueueManager(strMQManager)
    
    MQConectar = True

    Exit Function
    
ErrMQConectar:
    MQConectar = False

End Function

Public Function MQDesconectar(ByRef objMQManager As MQQueueManager, _
                              ByRef objMQEscribir As MQQueue, _
                              ByRef objMQLeer As MQQueue) As Boolean
                              
    On Error GoTo ErrMQDesconectar
    
    'Cierra cola escribir
    If Not Nothing Is objMQEscribir Then
        If objMQEscribir.IsOpen Then Call MQCerrarCola(objMQEscribir): Set objMQEscribir = Nothing
    End If
    
    'Cierra cola leer
    If Not Nothing Is objMQLeer Then
        If objMQLeer.IsOpen Then Call MQCerrarCola(objMQLeer): Set objMQLeer = Nothing
    End If
    
    ' Verificamos si existe el objeto
    If Not Nothing Is objMQManager Then
        ' Si esta conectado
        If objMQManager.IsConnected Then
            ' Se desconecta
            objMQManager.Disconnect
            Set objMQManager = Nothing
            MQDesconectar = True
        Else
            Set objMQManager = Nothing
        End If
    End If
    
    Set mqsConexion = Nothing
    

    Exit Function
    
ErrMQDesconectar:
    MQDesconectar = False
    psInsertaSQL 2, "Error en la desconexión a MQ", "TKT", "MQDesconectar"

End Function

Public Function MQAbrirCola(ByRef objMQManager As MQQueueManager, _
                            ByVal strMQCola As String, _
                            ByRef objMQCola As MQQueue, _
                            ByVal lngOpciones As MQOPEN) As Boolean
                            
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
    psInsertaSQL 7, "Error durante el cierre de la cola de MQSeries", "TKT", "MQCerrarCola"

End Function

Public Function MQRecibir(ByVal objMQConexion As MQSession, _
                          ByVal objMQManager As MQQueueManager, _
                          ByVal strMQCola As String, _
                          ByRef objMQCola As MQQueue, _
                          ByRef objMQMensaje As MQMessage) As Boolean
    
    On Error GoTo ErrMQRecibir

    Dim objMQOpciones   As MQGetMessageOptions ' Objeto PutMessageOptions para las opciones de lectura
    Dim strBuffer       As String
    Dim strMensaje      As String
    Dim lngErr          As Long
    Dim intLineas       As Integer
    Dim li_Indice       As Integer

    MQRecibir = True
    
    gsError = Space(0)
       
    ' Se accesan a la opciones de lectura por default
    Set objMQOpciones = objMQConexion.AccessGetMessageOptions()

    ' Con esta opción se leen los mensajes y se borran de la cola
    objMQOpciones.Options = MQGMO_NO_WAIT + MQGMO_COMPLETE_MSG
    strMensaje = ""
    intLineas = 0
    
    If MQAbrirCola(objMQManager, strMQCola, objMQCola, MQOO_INPUT_AS_Q_DEF) Then
        'Se accesa al mensaje.
        Set objMQMensaje = objMQConexion.AccessMessage()
        
        'Obtiene el mensaje de la cola con sus opciones
        objMQCola.Get objMQMensaje, objMQOpciones
        
        If Len(objMQMensaje.MessageData) = 0 Then
            Call MQCerrarCola(objMQCola)

            MQRecibir = False
            Exit Function
        End If
        
        If Mb_Detalles Then
            Escribe "Mb_Detalles ha sido activado. Longitud de mensaje recuperado: " & objMQMensaje.MessageLength
        End If
        Escribe "Contenido del mensaje recuperado: " & objMQMensaje.MessageData
        Call MQCerrarCola(objMQCola)
    Else
        psInsertaSQL 6, "Se presentó un fallo en la conexion MQ-Queue " & strMQCola & ": " & strMQCola & ": " & mqsConexion.ReasonCode & " " & mqsConexion.ReasonName & ". Fallo abrir MQ-Queue: " & strMQCola, "TKT", "MQAbrirCola"

        MQRecibir = False
        Exit Function
    End If
    

    Exit Function
    
ErrMQRecibir:
'    If Len(objMQMensaje.MessageData) = 0 Then Resume Next

    psInsertaSQL 8, "Se presentó un fallo en la conexion MQ-Queue " & strMQCola & ": " & strMQCola & ": " & mqsConexion.ReasonCode & " " & mqsConexion.ReasonName, "TKT", "MQRecibir"
    MQRecibir = False
End Function

Public Function MQEnviar(ByVal objMQConexion As MQSession, _
                         ByVal objMQManager As MQQueueManager, _
                         ByVal strMQCola As String, _
                         ByRef objMQCola As MQQueue, _
                         ByRef objMQMensaje As MQMessage, _
                         ByVal strMsgRespuesta As String) As Boolean

    On Error GoTo ErrMQEnviar
    
    Dim objMQOpciones   As MQPutMessageOptions ' Objeto PutMessageOptions para las opciones de escritura

    
    ' Se accesan a la opciones de escritura por default
    Set objMQOpciones = objMQConexion.AccessPutMessageOptions()
    objMQOpciones.Options = objMQOpciones.Options Or MQPMO_NO_SYNCPOINT
    Set objMQMensaje = objMQConexion.AccessMessage
        
        If MQAbrirCola(objMQManager, strMQCola, objMQCola, MQOO_OUTPUT) Then
            objMQCola.Put objMQMensaje, objMQOpciones
            objMQMensaje.ClearMessage
            objMQMensaje.Format = "MQSTR   "
            objMQMensaje.MessageType = MQMT_DATAGRAM
            objMQMensaje.WriteLong Len(strMsgRespuesta)
            objMQMensaje.MessageData = strMsgRespuesta
            objMQCola.Put objMQMensaje, objMQOpciones
            ' Cierra cola escribir
            Call MQCerrarCola(objMQCola)
            MQEnviar = True
        End If

    Exit Function
    
ErrMQEnviar:
    MQEnviar = False

End Function


Public Function VerificarMQManager(ByRef objMQConexion As MQSession, _
                           ByVal strMQManager As String, _
                           ByRef objMQManager As MQQueueManager) As String

    On Error GoTo ErrMQConectar
    

    
    ' Se crear una sesión con el servidor de MQSeries
    ' y se accesa al Queue Manager.
    Set objMQManager = objMQConexion.AccessQueueManager(strMQManager)
    
    VerificarMQManager = ""

    objMQManager.Disconnect
    Exit Function
    
ErrMQConectar:
   VerificarMQManager = objMQConexion.ReasonCode & vbCrLf & objMQConexion.ReasonName

End Function

Public Function VerificarMQQueue(ByRef objMQManager As MQQueueManager, _
                            ByVal strMQCola As String, _
                            ByRef objMQCola As MQQueue, _
                            ByVal lngOpciones As MQOPEN) As String
   
   On Error GoTo ErrMQAbrirCola
   

   
   ' Se accesa la cola ya sea para leer o escribir
   Set objMQCola = objMQManager.AccessQueue(strMQCola, lngOpciones, mqsManager.Name, "AMQ.*")
    
   objMQCola.Close
   VerificarMQQueue = ""

   Exit Function
    
ErrMQAbrirCola:
   VerificarMQQueue = objMQManager.ReasonCode & vbCrLf & objMQManager.ReasonName

End Function

Public Sub psInsertaSQL(pnNumeroError As Integer, psDescripcion As String, _
                 psAplicacion As String, psFuncion As String)
'*******************************************************************************************************************
'Procedimiento: psInsertaSQL
'Objetivo:      Realizar la inserción del registro de error en la base de datos del Ticket durante el procesamiento
'               del mensaje de la QUEUE
'Autor:         EDS-AGO
'Entradas:      psFechaHora. Fecha-Hora de la presentación del error.
'               pnNumeroError. Número de error establecido por la aplicación
'               psDescripcion. Descripción del error
'               psAplicacion. Aplicación en la que se presentó el error.
'               psFuncion. Función en la que se presentó el error.
'Fecha:         22/05/2006
'*******************************************************************************************************************

    'Arma el query de inserción
    strQuery = "Insert into " & gsNameDB & "..BITACORA_ERRORES_MENSAJES_PU "
    strQuery = strQuery & "(fecha_hora, error_numero, error_descripcion, aplicacion) "
    strQuery = strQuery & "Values ('" & Now & "', " & pnNumeroError & ", '" & psDescripcion & "', '" & psAplicacion & "')"
    rssRegistro.Open strQuery

End Sub

Public Function MQRevisar(ByVal objMQConexion As MQSession, _
                          ByVal objMQManager As MQQueueManager, _
                          ByVal strMQCola As String, _
                          ByRef objMQCola As MQQueue, _
                          ByRef objMQMensaje As MQMessage) As Integer
    
    On Error GoTo ErrMQRevisar

    MQRevisar = False
    
    MQRevisar = 0
              
    If MQAbrirCola(objMQManager, strMQCola, objMQCola, MQOO_INQUIRE) Then
        MQRevisar = objMQCola.CurrentDepth
    Else
        psInsertaSQL 6, "Se presentó un fallo en la conexion MQ-Queue " & strMQCola & ": " & strMQCola & ": " & mqsConexion.ReasonCode & " " & mqsConexion.ReasonName & ". Fallo abrir MQ-Queue: " & strMQCola, "TKT", "MQAbrirCola"
    End If
        
    Call MQCerrarCola(objMQCola)
            
    Exit Function
    
ErrMQRevisar:
    psInsertaSQL 8, "Se presentó un fallo en la conexion MQ-Queue " & strMQCola & ": " & strMQCola & ": " & mqsConexion.ReasonCode & " " & mqsConexion.ReasonName, "TKT", "MQRevisar"
End Function



Public Sub Escribe(vData As String)
    ' Configuración del archivo log
    Archivo = lsLogPath + Format(Now(), "yyyyMMdd") + "-" + lsLogName
    If Mb_GrabaLog = True Then
        Open Archivo For Append As #1
        Print #1, vData
        Close #1
    End If
    
End Sub


