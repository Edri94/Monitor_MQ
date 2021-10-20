VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cBitacoras"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit


'declaracion de variables

Private ArchivoIni      As String
Private lsCommandLine   As String

 'Variables para el registro de los valores del header IH
Private strFuncionHost  As String 'Valor que indica el programa que invocara el CICSBRIDGE
Private strHeaderTagIni As String 'Bandera que indica el comienzo del Header
Private strIDProtocol   As String 'Identificador  del protocolo (PS9)
Private strLogical      As String 'Terminal Lógico Asigna Arquitectura ASTA
Private strAccount      As String 'Terminal Contable (CR Contable)
Private strUser         As String 'Usuario. Debe ser diferente de espacios
Private strSeqNumber    As String 'Número de Secuencia (indicador de paginación)
Private strTXCode       As String 'Función específica Asigna Arquitectura Central
Private strUserOption   As String 'Tecla de función (no aplica)
Private strCommit       As String 'Indicador de commit: Permite realizar commit
Private strMsgType      As String 'Tipo de mensaje: Nuevo requerimiento
Private strProcessType  As String 'Tipo de proceso: on line
Private strChannel      As String 'Canal Asigna Arquitectura Central
Private strPreFormat    As String 'Indicador de preformateo: Arquitectura no deberá de preformatear los datos
Private strLenguage     As String 'Idioma: Español
Private strHeaderTagEnd As String 'Bandera que indica el final del header

' Variables para el registro de los valores del header ME
Private strMETAGINI     As String 'Bandera que indica el comienzo del mensaje
'Private strMsgColecMax  As String 'Longitud del layout  del colector
Private strMsgTypeCole  As String 'Tipo de mensaje: Copy
'Private strMaxMsgCole   As String 'Máximo X(30641)
Private strMETAGEND     As String 'Bandera que indica el fin del mensaje

' Variables para el registro de los valores Default
'Dim strFechaBaja        As String 'fecha_baja
Dim strColectorMaxLeng  As String 'Maxima longitud del COLECTOR
Dim strMsgMaxLeng       As String 'Maxima longitud del del bloque ME
Dim strPS9MaxLeng       As String 'Maxima longitud del formato PS9
Dim strReplyToMQ        As String 'MQueue de respuesta para HOST
Dim strFuncionSQL       As String 'Funcion a ejecutar al recibir la respuesta
Dim strRndLogTerm       As String 'Indica que el atributo Logical Terminal es random

Private Gs_MQManager            As String   'MQManager de Escritura
Private Gs_MQQueueEscritura     As String   'MQQueue de Escritura
Public gsAccesoActual           As String   'Fecha/Hora actual del sistema. La tomamos del servidor NT y no de SQL porque precisamente el
                                            'requerimiento de la fecha/hora fue porque a veces se pierde la conexión a SQL

Public Function ProcesarBitacora(ByVal strRutaIni As String, _
                                Optional ByVal strParametros As String)
                                
    Dim Parametros()        As String       ' Arreglo para almacenar los parametros via línea de comando
    Dim Ls_MsgVal           As String       ' Mensaje con el resultado de la validación
    
On Error GoTo Form_Load_Err

    ArchivoIni = strRutaIni & "\Bitacoras.ini"
    Call ConfiguraFileLog("EscribeArchivoLOG", ArchivoIni)
    
    'Esta es la fecha y hora actual
    gsAccesoActual = Now
    
    lsCommandLine = Trim(strParametros)
   
    If StrComp(lsCommandLine, "", vbTextCompare) <> 0 Then
        Erase Parametros()
        Parametros = Split(lsCommandLine, "-")
        Gs_MQManager = Trim(Parametros(0))
        Gs_MQQueueEscritura = Trim(Parametros(1))
        strFuncionSQL = Parametros(3)
    Else
        Call ObtenerInfoMQ("MQ Series", ArchivoIni)
    End If

    Call ConfiguraHeader_IH_ME(ArchivoIni)
  
    If Not ValidaInfoMQ(Ls_MsgVal) Then
        Escribe "Se presentó un error en la función ValidaInfoMQ invocada desde el MAIN: " & Ls_MsgVal & ". Función SQL: " & strFuncionSQL
        Exit Function
    End If
        
    ProcesoBDtoMQQUEUE
    
    Escribe "Termina proceso bitácoras. Función SQL: " & strFuncionSQL
    
Exit Function

Form_Load_Err:
    Call MQDesconectar(mqManager, mqsEscribir)
    Escribe "Termina el acceso a la aplicación Bitácoras porque se presentó un error en la función MAIN. Función SQL: " & strFuncionSQL & ". Error. " & Err.Number & "-" & Err.Description
End Function


Private Function ProcesoBDtoMQQUEUE()
    Dim Ls_MensajeMQ        As String       ' Cadena con el mensaje armado con los registros de la base de datos
    Dim Ls_MsgColector      As String       ' Cadena para almecenar el COLECTOR
    
    Dim sFechaEnvio As String
    Dim sEnvioConse As String
    Dim sMensajeEnvio As String
    
On Error GoTo ErrorProcesoBDMQ
    
    Escribe ""
    Escribe "Inicia envío de mensajes a Host: " & gsAccesoActual & " Función SQL: " & strFuncionSQL
    
    If MQConectar(mqSession, Gs_MQManager, mqManager) Then
        blnConectado = True
    Else
        Escribe "Fallo conexión MQ-Manager " & Gs_MQManager & ": " & mqSession.ReasonCode & " - " & mqSession.ReasonName
        Exit Function
    End If
                       
    sFechaEnvio = Left(Format(Now, "yyyymmddhhnnss") + Space(26), 26)
    sEnvioConse = Left(ObtenParametroIni("VALORTK14", "TKCONSECUTIVO", "", ArchivoIni) + Space(1), 1)
   
    Ls_MsgColector = Left(strFuncionSQL + Space(8), 8) + sFechaEnvio
                     '
    If Len(Ls_MsgColector) > 0 Then
        Ls_MensajeMQ = ASTA_ENTRADA(Ls_MsgColector)
        If Ls_MensajeMQ <> "" Then
            Escribe "Mensaje Enviado: " & Ls_MensajeMQ
            If MQEnviarMsg(mqSession, mqManager, Gs_MQQueueEscritura, mqsEscribir, mqsMsgEscribir, Ls_MensajeMQ, strReplyToMQ) Then
                'REESCRIBE EN .INI
                sMensajeEnvio = CStr(sEnvioConse + 1)
                If Val(sMensajeEnvio) > 9 Then
                    sMensajeEnvio = CStr(1)
                End If
                
                SaveProfile ArchivoIni, "VALORTK14", "TKCONSECUTIVO", sMensajeEnvio
            Else
                Escribe "Se ha presentado un error al escribir la solicitud en la MQ QUEUE: " & Gs_MQQueueEscritura & ". Mensaje: " & Ls_MensajeMQ
            End If
        Else
            Escribe "Se ha presentado un error durante el armado del formato PS9 funcion ASTA_ENTRADA. Colector: " & Ls_MsgColector
        End If
    Else
        Escribe "Se ha presentado un error al armar el Layout TKT14. No existe longitud en el Colector"
    End If
        
    Call MQDesconectar(mqManager, mqsEscribir)
    
    Escribe "Envio de solicitures TKT -> Host Terminado"
    Escribe "Solicitudes enviadas a MQ: " & sMensajeEnvio
    Exit Function

ErrorProcesoBDMQ:
    Escribe "Se presentó un error durante la ejecución de la función ProcesoBDtoMQQUEUE"
End Function

Private Function ValidaInfoMQ(ByRef ps_MsgVal As String) As Boolean
    Dim ls_msg As String
    
    If Trim(Gs_MQManager) = "" Then ls_msg = ls_msg & IIf(Len(ls_msg) > 0, Chr(13), "") & "Falta MQ Manager envio."
    If Trim(Gs_MQQueueEscritura) = "" Then ls_msg = ls_msg & IIf(Len(ls_msg) > 0, Chr(13), "") & "Falta MQ Queue envio."
    If Trim(ls_msg) = "" Then ValidaInfoMQ = True
    ps_MsgVal = ls_msg
End Function

Private Function ObtenParametroIni(Ls_Grupo As String, Ls_Variable As String, Ls_Default As String, Ls_AppPath As String) As String
    Dim Ls_Buffer As String * 80
    Dim Li_Long As Integer
    Li_Long = GetPrivateProfileString(ByVal Ls_Grupo, ByVal Ls_Variable, ByVal Ls_Default, ByVal Ls_Buffer, Len(Ls_Buffer), Ls_AppPath)
    Li_Long = InStr(Ls_Buffer, Chr(0))
    ObtenParametroIni = Trim(Left(Ls_Buffer, Li_Long - 1))
End Function

Private Sub ObtenerInfoMQ(Ls_Tit As String, Ls_Path As String)
    Gs_MQManager = ObtenParametroIni(Ls_Tit, "MQManager", "", Ls_Path)
    Gs_MQQueueEscritura = ObtenParametroIni(Ls_Tit, "MQEscritura", "", Ls_Path)
    strFuncionSQL = ObtenParametroIni(Ls_Tit, "FGBitacora", "", Ls_Path)
End Sub

Private Sub ConfiguraFileLog(Ls_Tit As String, Ls_Path As String)
    strlogFileName = ObtenParametroIni(Ls_Tit, "logFileName", "", Ls_Path)
    strlogFilePath = ObtenParametroIni(Ls_Tit, "logFilePath", "", Ls_Path)
    Mb_GrabaLog = True
End Sub

Private Sub ConfiguraHeader_IH_ME(Ls_Path As String)

    strFuncionHost = ObtenParametroIni("HEADERIH", "PRIMERVALOR", "", Ls_Path)
    strHeaderTagIni = ObtenParametroIni("HEADERIH", "IHTAGINI", "", Ls_Path)
    strIDProtocol = ObtenParametroIni("HEADERIH", "IDPROTOCOL", "", Ls_Path)
    strLogical = ObtenParametroIni("HEADERIH", "Logical", "", Ls_Path)
    strAccount = ObtenParametroIni("HEADERIH", "ACCOUNT", "", Ls_Path)
    strUser = ObtenParametroIni("HEADERIH", "User", "", Ls_Path)
    strSeqNumber = ObtenParametroIni("HEADERIH", "SEQNUMBER", "", Ls_Path)
    strTXCode = ObtenParametroIni("HEADERIH", "TXCODE", "", Ls_Path)
    strUserOption = ObtenParametroIni("HEADERIH", "USEROPT", "", Ls_Path)
    strCommit = ObtenParametroIni("HEADERIH", "Commit", "", Ls_Path)
    strMsgType = ObtenParametroIni("HEADERIH", "MSGTYPE", "", Ls_Path)
    strProcessType = ObtenParametroIni("HEADERIH", "PROCESSTYPE", "", Ls_Path)
    strChannel = ObtenParametroIni("HEADERIH", "CHANNEL", "", Ls_Path)
    strPreFormat = ObtenParametroIni("HEADERIH", "PREFORMATIND", "", Ls_Path)
    strLenguage = ObtenParametroIni("HEADERIH", "LANGUAGE", "", Ls_Path)
    strHeaderTagEnd = ObtenParametroIni("HEADERIH", "IHTAGEND", "", Ls_Path)

    strMETAGINI = ObtenParametroIni("HEADERME", "METAGINI", "", Ls_Path)
    'strMsgColecMax = ObtenParametroIni("HEADERME", "COLECTORLENGTH", "", Ls_Path)
    strMsgTypeCole = ObtenParametroIni("HEADERME", "TIPOMSG", "", Ls_Path)
    'strMaxMsgCole = ObtenParametroIni("HEADERME", "MSGMAXLENG", "", Ls_Path)
    strMETAGEND = ObtenParametroIni("HEADERME", "METAGEND", "", Ls_Path)

    'strFechaBaja = ObtenParametroIni("DEFAULTVALUES", "FECHABAJA", "", Ls_Path)
    strColectorMaxLeng = ObtenParametroIni("DEFAULTVALUES", "COLMAXLENG", "", Ls_Path)
    strMsgMaxLeng = ObtenParametroIni("DEFAULTVALUES", "MSGMAXLENG", "", Ls_Path)
    strPS9MaxLeng = ObtenParametroIni("DEFAULTVALUES", "PS9MAXLENG", "", Ls_Path)
    strReplyToMQ = ObtenParametroIni("DEFAULTVALUES", "ReplyToQueue", "", Ls_Path)
    strRndLogTerm = ObtenParametroIni("DEFAULTVALUES", "RandomLogTerm", "", Ls_Path)
End Sub

Private Function ASTA_ENTRADA(strMsgColector As String) As String
    Dim ls_TempColectorMsg      As String
    Dim ls_BloqueME             As String
    Dim ln_longCOLECTOR         As Integer
    Dim ln_AccTerminal          As Integer
    
    On Error GoTo ErrorASTA
    
    ls_TempColectorMsg = strMsgColector
    If Len(ls_TempColectorMsg) > CInt(Trim(strColectorMaxLeng)) Then
        Escribe ("La longitud del colector supera el maximo permitido")
        GoTo ErrorASTA
    End If
    
    ls_BloqueME = Left(Trim(strMETAGINI) + Space(4), 4)
    ls_BloqueME = ls_BloqueME + Right("0000" + CStr(Len(ls_TempColectorMsg)), 4)
    ls_BloqueME = ls_BloqueME + Left(Trim(strMsgTypeCole) + Space(1), 1)
    ls_BloqueME = ls_BloqueME + ls_TempColectorMsg
    ls_BloqueME = ls_BloqueME + Left(Trim(strMETAGEND) + Space(5), 5)
    
    If Len(ls_BloqueME) > CInt(Trim(strMsgMaxLeng)) Then
        Escribe ("La longitud del Bloque ME supera el maximo permitido")
        GoTo ErrorASTA
    End If

    ASTA_ENTRADA = Left(Trim(strFuncionHost) + Space(8), 8)
    ASTA_ENTRADA = ASTA_ENTRADA + Left(Trim(strHeaderTagIni) + Space(4), 4)
    ASTA_ENTRADA = ASTA_ENTRADA + Left(Trim(strIDProtocol) + Space(2), 2)
    If StrComp(Trim(strRndLogTerm), "1", vbTextCompare) = 0 Then
        ln_AccTerminal = 0
        Do While Not (ln_AccTerminal > 0 And ln_AccTerminal < 2000)
            Call Randomize(DateTime.DatePart("s", DateTime.Now))
            ln_AccTerminal = Rnd(DateTime.DatePart("s", DateTime.Now)) * 1000
        Loop
        ASTA_ENTRADA = ASTA_ENTRADA + Left(Format(CStr(ln_AccTerminal), "0000") + Space(8), 8)
    Else
        ASTA_ENTRADA = ASTA_ENTRADA + Left(Trim(strLogical) + Space(8), 8)
    End If
    
    ASTA_ENTRADA = ASTA_ENTRADA + Left(Trim(strAccount) + Space(8), 8)
    ASTA_ENTRADA = ASTA_ENTRADA + Left(Trim(strUser) + Space(8), 8)
    ASTA_ENTRADA = ASTA_ENTRADA + Left(Trim(strSeqNumber) + Space(8), 8)
    ASTA_ENTRADA = ASTA_ENTRADA + Left(Trim(strTXCode) + Space(8), 8)
    ASTA_ENTRADA = ASTA_ENTRADA + Left(Trim(strUserOption) + Space(2), 2)
    
    ln_longCOLECTOR = 65 + Len(ls_BloqueME)
    
    If ln_longCOLECTOR > CInt(Trim(strPS9MaxLeng)) Then
        Escribe ("La longitud del Layout PS9 supera el maximo permitido")
        GoTo ErrorASTA
    End If
    
    ASTA_ENTRADA = ASTA_ENTRADA + Right("00000" + CStr(ln_longCOLECTOR), 5)
    ASTA_ENTRADA = ASTA_ENTRADA + Left(Trim(strCommit) + Space(1), 1)
    ASTA_ENTRADA = ASTA_ENTRADA + Left(Trim(strMsgType) + Space(1), 1)
    ASTA_ENTRADA = ASTA_ENTRADA + Left(Trim(strProcessType) + Space(1), 1)
    ASTA_ENTRADA = ASTA_ENTRADA + Left(Trim(strChannel) + Space(2), 2)
    ASTA_ENTRADA = ASTA_ENTRADA + Left(Trim(strPreFormat) + Space(1), 1)
    ASTA_ENTRADA = ASTA_ENTRADA + Left(Trim(strLenguage) + Space(1), 1)
    ASTA_ENTRADA = ASTA_ENTRADA + Left(Trim(strHeaderTagEnd) + Space(5), 5)
    ASTA_ENTRADA = ASTA_ENTRADA + ls_BloqueME
    
Exit Function

ErrorASTA:
    ASTA_ENTRADA = ""
    Escribe "Error al armar el mensaje en la función ASTA_ENTRADA"
End Function

Private Sub SaveProfile(lpFileName As String, lpAppName As String, lpKeyName As String, lpString As String)
       Dim Ret As Long
       Ret = WritePrivateProfileString(lpAppName, lpKeyName, lpString, lpFileName)
End Sub



