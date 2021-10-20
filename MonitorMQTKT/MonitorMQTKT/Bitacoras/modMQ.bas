Attribute VB_Name = "modMQ"
Option Explicit

Declare Function GetPrivateProfileString Lib "kernel32" Alias "GetPrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpDefault As String, ByVal lpReturnedString As String, ByVal nSize As Long, ByVal lpFileName As String) As Long

Public Declare Function WritePrivateProfileString Lib "kernel32" Alias "WritePrivateProfileStringA" _
(ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpString As Any, ByVal lpFileName As String) As Long

Private Declare Function GetPrivateProfileSection Lib "kernel32" Alias "GetPrivateProfileSectionA" _
(ByVal lpAppName As String, ByVal lpReturnedString As String, ByVal nSize As Long, ByVal lpFileName As String) As Long

Public strlogFilePath  As String
Public strlogFileName  As String
Public Mb_GrabaLog     As Boolean

Private Archivo         As String


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
Public mqSession        As New mqSession           ' Objeto Session para conexión con el servidor MQSeries
Public mqManager        As New MQQueueManager      ' -Objeto QueueManager para accesar al maestro de colas
Public mqsEscribir      As New MQQueue             ' -Objeto Queue para escribir
Public mqsLectura       As New MQQueue             ' -Objeto Queue para lectura
Public mqsMsgEscribir   As New MQMessage           ' -Objeto Message para escribir
Public mqsMsglectura    As New MQMessage           ' -Objeto Message para lectura

Public Function MQConectar(ByRef objMQConexion As mqSession, _
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
                              ByRef objMQEscribir As MQQueue) As Boolean
                              
    On Error GoTo ErrMQDesconectar
    

    
    'Cierra cola escribir
    If Not Nothing Is objMQEscribir Then
        If objMQEscribir.IsOpen Then Call MQCerrarCola(objMQEscribir): Set objMQEscribir = Nothing
    End If
    
    ' Verificamos si existe el objeto
    If Not Nothing Is objMQManager Then
        ' Si esta conectado
        If objMQManager.IsConnected Then
            ' Se desconecta
            objMQManager.Disconnect
            Set objMQManager = Nothing
            MQDesconectar = True
        End If
    End If
    
    Set mqSession = Nothing
    

    Exit Function
    
ErrMQDesconectar:
    MQDesconectar = False
    Escribe "Error en la desconexión a MQ"

End Function

Public Function MQAbrirCola(ByRef objMQManager As MQQueueManager, _
                            ByVal strMQCola As String, _
                            ByRef objMQCola As MQQueue, _
                            ByVal lngOpciones As MQOPEN) As Boolean
                            
    On Error GoTo ErrMQAbrirCola
    

    
    ' Se accesa la cola ya sea para leer o escribir
    Set objMQCola = objMQManager.AccessQueue(strMQCola, lngOpciones, mqManager.Name, "AMQ.*")
    
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

Public Function VerificarMQQueue(ByRef objMQManager As MQQueueManager, _
                            ByVal strMQCola As String, _
                            ByRef objMQCola As MQQueue, _
                            ByVal lngOpciones As MQOPEN) As String

   On Error GoTo ErrMQAbrirCola



   ' Se accesa la cola ya sea para leer o escribir
   Set objMQCola = objMQManager.AccessQueue(strMQCola, lngOpciones, mqManager.Name, "AMQ.*")

   objMQCola.Close
   VerificarMQQueue = ""

   Exit Function

ErrMQAbrirCola:
   VerificarMQQueue = objMQManager.ReasonCode & vbCrLf & objMQManager.ReasonName

End Function

Public Function MQEnviarMsg(ByVal objMQConexion As mqSession, _
                         ByVal objMQManager As MQQueueManager, _
                         ByVal strMQCola As String, _
                         ByRef objMQCola As MQQueue, _
                         ByRef objMQMensaje As MQMessage, _
                         ByRef ls_mensaje As String, _
                         ByVal Ls_ReplayMQQueue As String, _
                         Optional ByVal strMensajeID As String = "") As Boolean
    On Error GoTo ErrMQEnviar

    Dim mqsMQOpciones   As MQPutMessageOptions
    Dim strMensaje      As String



    ' Se accesan a la opciones de escritura por default
    Set mqsMQOpciones = objMQConexion.AccessPutMessageOptions
    mqsMQOpciones.Options = mqsMQOpciones.Options Or MQPMO_NO_SYNCPOINT
    Set objMQMensaje = objMQConexion.AccessMessage
           
    If MQAbrirCola(objMQManager, strMQCola, objMQCola, MQOO_OUTPUT) Then
        strMensaje = ls_mensaje
        objMQMensaje.ClearMessage
        objMQMensaje.Format = "MQSTR   "
        objMQMensaje.MessageType = MQMT_DATAGRAM
        objMQMensaje.WriteLong Len(Trim(strMensaje))
        objMQMensaje.MessageData = Trim(strMensaje)
        objMQMensaje.ReplyToQueueName = Ls_ReplayMQQueue
        objMQCola.Put objMQMensaje, mqsMQOpciones
        ' Cierra cola escribir
        Call MQCerrarCola(objMQCola)
        MQEnviarMsg = True
    End If
    

    Exit Function

ErrMQEnviar:
    Escribe ("Error durante en el envio del mensaje objMQMensaje: " & objMQMensaje.ReasonCode & " : " & objMQMensaje.ReasonName)
    Escribe ("Error durante en el envio del mensaje objMQCola: " & objMQCola.ReasonCode & " : " & objMQCola.ReasonName)
    MQEnviarMsg = False

End Function



Public Sub Escribe(vData As String)
    Archivo = strlogFilePath & Format(Now(), "yyyyMMdd") & "-" & strlogFileName
    If Mb_GrabaLog Then
        Open Archivo For Append As #1
        Print #1, vData
        Close #1
    End If
End Sub

