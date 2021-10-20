VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cMensajes"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

'declaracion de variables
Private Archivo                 As String
Private ArchivoIni              As String
Private Ls_Archivo              As String
Private lsCommandLine           As String

' Variables para el control del log
Private strlogFileName          As String
Private strlogFilePath          As String
Private Mb_GrabaLog             As Boolean

' Variables para el registro de los valores del header IH
Private strFuncionHost          As String 'Valor que indica el programa que invocara el CICSBRIDGE
Private strHeaderTagIni         As String 'Bandera que indica el comienzo del Header
Private strIDProtocol           As String 'Identificador  del protocolo (PS9)
Private strLogical              As String 'Terminal Lógico Asigna Arquitectura ASTA
Private strAccount              As String 'Terminal Contable (CR Contable)
Private strUser                 As String 'Usuario. Debe ser diferente de espacios
Private strSeqNumber            As String 'Número de Secuencia (indicador de paginación)
Private strTXCode               As String 'Función específica Asigna Arquitectura Central
Private strUserOption           As String 'Tecla de función (no aplica)
Private strCommit               As String 'Indicador de commit: Permite realizar commit
Private strMsgType              As String 'Tipo de mensaje: Nuevo requerimiento
Private strProcessType          As String 'Tipo de proceso: on line
Private strChannel              As String 'Canal Asigna Arquitectura Central
Private strPreFormat            As String 'Indicador de preformateo: Arquitectura no deberá de preformatear los datos
Private strLenguage             As String 'Idioma: Español
Private strHeaderTagEnd         As String 'Bandera que indica el final del header

' Variables para el registro de los valores del header ME
Private strMETAGINI             As String 'Bandera que indica el comienzo del mensaje
'Private strMsgColecMax          As String 'Longitud del layout  del colector
Private strMsgTypeCole          As String 'Tipo de mensaje: Copy
'Private strMaxMsgCole           As String 'Máximo X(30641)
Private strMETAGEND             As String 'Bandera que indica el fin del mensaje

' Variables para el registro de los valores Default
Dim strFechaBaja                As String 'fecha_baja
Dim strColectorMaxLeng          As String 'Maxima longitud del COLECTOR
Dim strMsgMaxLeng               As String 'Maxima longitud del del bloque ME
Dim strPS9MaxLeng               As String 'Maxima longitud del formato PS9
Dim strReplyToMQ                As String 'MQueue de respuesta para HOST
Dim strFuncionSQL               As String 'Funcion a ejecutar al recibir la respuesta
Dim strRndLogTerm               As String 'Indica que el atributo Logical Terminal es random

' Variables para el manejo de los parametros de la base de datos
'Public gsSeccRegWdw             As String

'VARIABLES NUVAS PARA EL ENVIO DE MENSAJE
Private sPersistencia           As String
Private sExpirar                As String

Private Gs_MQManager            As String       'MQManager de Escritura
Private Gs_MQQueueEscritura     As String       'MQQueue de Escritura
Private gsEjecutable            As String       'Ejecutable a realizar

Public Bandera                  As String           'Indicador de encripción de constantes de conexión a SQL Server (1=encriptado; 0=no encriptado)

Public Function ProcesarMensajes(ByVal strRutaIni As String, _
                                Optional ByVal strParametros As String)
    Dim Parametros()        As String       ' Arreglo para almacenar los parametros via línea de comando
    Dim Ls_MsgVal           As String       ' Mensaje con el resultado de la validación
    Dim LnDiferencia        As Single       ' Minutos transcurridos desde el último intento de acceso
    
    ArchivoIni = strRutaIni & "\MensajesMQ.ini"
    gstrRutaIni = ArchivoIni
    
    ' Recupera parametro via línea de comandos
    lsCommandLine = Trim(strParametros)
    
    If StrComp(lsCommandLine, "", vbTextCompare) <> 0 Then
        Erase Parametros()
        Parametros = Split(lsCommandLine, "-")
        Gs_MQManager = Trim(Parametros(0))
        Gs_MQQueueEscritura = Trim(Parametros(1))
        gsEjecutable = Trim(Parametros(2))
    Else
        ' Recupera la información del archivo ini dedicada a la configuración de la MQ
        Call ObtenerInfoMQ("MQ Series", ArchivoIni)
    End If
        
    ' Recupera la información del archivo ini dedicada a la configuración del archivo log de la aplicación
    Call ConfiguraFileLog("EscribeArchivoLOG", ArchivoIni)
              
    'Configura el Header de MQ. Se toma directamente del archivo INI
    Call ConfiguraHeader_IH_ME(ArchivoIni)
    
    'Obtiene la sección del registro de windows donde se encuentra la información de conexión
    'gsSeccRegWdw = ObtenParametroIni("CONEXIONSQL", "SECCREGWDW", "", ArchivoIni)
        
    'Verifica si hay conexión al servidor SQL
    If ConectDB = False Then Exit Function
                
    'Solo hasta este punto la conexión ha sido establecida
    Escribe "Comienza la función MAIN de la aplicación MensajesMQ: " & Now & " Tipo Función: '" & gsEjecutable & "'"
    gsAccesoActual = Date & " " & Time
    'Verifica la existencia de la información de la Queue
    If Not ValidaInfoMQ(Ls_MsgVal) Then
        psInsertaSQL gsAccesoActual, 1, Ls_MsgVal, "MSG", "ValidaInfoMQ"
        Escribe "Termina el acceso a la aplicación MensajesMQ. Cheque la bitácora de errores en SQL. Tipo Función: '" & gsEjecutable & "'"
        Escribe ""
        Desconectar
        Exit Function
    End If
                
    'Operación Armado de MSG para ser depositados en la MQQUEUE
    Select Case gsEjecutable
        Case "F"
            ProcesoBDtoMQQUEUEFunc
        Case "A"
            ProcesoBDtoMQQUEUEAuto
    End Select
        
    ' Desconexión con la base de datos
    Desconectar
    
    ' Escribe la hora de terminación de la aplicación
    Escribe "Termina el acceso a la aplicación MensajesMQ. Función SQL: " & strFuncionSQL
    Escribe ""
    
End Function

Public Sub Escribe(vData As String)
    Archivo = strlogFilePath & Format(Now(), "yyyyMMdd") & "-" & strlogFileName
    If Mb_GrabaLog Then
        Open Archivo For Append As #1
        Print #1, vData
        Close #1
    End If
End Sub

Private Function ProcesoBDtoMQQUEUEFunc()
    Dim Ls_MensajeMQ        As String       ' Cadena con el mensaje armado con los registros de la base de datos
    Dim Ls_MsgColector      As String       ' Cadena para almecenar el COLECTOR
    Dim Ls_HeaderMsg        As String       ' Cadena para almacenar el HEADER del mensaje
    Dim NumeroMsgEnviados   As Integer      ' Contador para almacenar el número de mensajes procesados
    Dim las_Funcionarios()  As String       ' Arreglo para ingresar todos los registros que han sido enviados correctamente
    ' Para el armado de la solicitud
    Dim ls_IDFuncionario    As String
    Dim ls_CentroRegional   As String       '1  centro_regional
    Dim ls_NumRegistro      As String       '2  numero_registro
    Dim ls_Producto         As String       '3  producto
    Dim ls_SubProducto      As String       '4  subproducto
    Dim ls_FechaAlta        As String       '5  fecha_alta
    Dim ls_TipoPeticion     As String       '8  tipo_peticion
    Dim ls_IdTransaccion    As String       '12 id_transaccion
    Dim ls_Tipo             As String       '13 tipo
    Dim ls_Fecha            As String
    Dim ls_Hora             As String
    

On Error GoTo ErrorProcesoBDMQ
        
    Escribe "Inicia el envío de mensajes a Host: " & gsAccesoActual & " Función: " & strFuncionSQL
    NumeroMsgEnviados = 0
    
   'Logica para recuperar los n mensajes de la tabla temporal en db.funcionario
   'Logica para procesar cada registro y convertirlo en un mensaje
    strQuery = "SELECT" & vbCrLf
    strQuery = strQuery + "id_funcionario," & vbCrLf                       '0  id_funcionario
    strQuery = strQuery + "centro_regional," & vbCrLf                      '1  centro_regional
    strQuery = strQuery + "numero_funcionario," & vbCrLf                   '2  numero_
    strQuery = strQuery + "producto," & vbCrLf                             '3  producto
    strQuery = strQuery + "subproducto," & vbCrLf                          '4  subproducto
    strQuery = strQuery + "CONVERT(char(11), fecha_alta, 105) + CONVERT(char(5), fecha_alta, 108)," & vbCrLf                           '5  fecha_alta
    strQuery = strQuery + "CONVERT(char(11), fecha_baja, 105) + CONVERT(char(5), fecha_baja, 108)," & vbCrLf                           '6  fecha_baja
    strQuery = strQuery + "CONVERT(char(11), fecha_ultimo_mant, 105) + CONVERT(char(6), fecha_ultimo_mant, 108)," & vbCrLf                    '7  fecha_ultimo_mant
    strQuery = strQuery + "tipo_peticion," & vbCrLf                        '8  tipo_peticion
    strQuery = strQuery + "status_envio," & vbCrLf                          '9  status_envio
    strQuery = strQuery + "CONVERT(char(8),getdate(),112)," & vbCrLf        '10
    strQuery = strQuery + "CONVERT(char(5),getdate(),108)," & vbCrLf        '11
    strQuery = strQuery + "id_transaccion," & vbCrLf                        '12  id transaccion en TKT
    strQuery = strQuery + "tipo " & vbCrLf                                  '13  Tipo  A-Alta, B-Baja, M-Mantenimiento
    strQuery = strQuery + "FROM" & vbCrLf
    strQuery = strQuery + gsNameDB & "..TMP_FUNCIONARIOS_PU" & vbCrLf
    strQuery = strQuery + "WHERE status_envio = 0"
        
    rssRegistro.Open strQuery
            
    If Not rssRegistro.EOF Then
        
        If MQConectar(mqSession, Gs_MQManager, mqManager) Then
            blnConectado = True
        Else
            'Registra en el log de SQL de la aplicación que fue imposible conectarse con el MQManager
            psInsertaSQL gsAccesoActual, 3, "ProcesoBDtoMQQUEUEFunc. Fallo conexión MQ-Manager " & Gs_MQManager & ": " & mqSession.ReasonCode & " - " & mqSession.ReasonName, "MSG", "MQConectar"
            Exit Function
        End If
        
        Do While Not rssRegistro.EOF
            ' Almacenado de variables
            ls_IDFuncionario = Left(Trim(CStr(Format(rssRegistro(0).Value, "0000000"))) + "       ", 7)
            ls_CentroRegional = Left(Trim(CStr(rssRegistro(1).Value)) + "    ", 4)
            ls_NumRegistro = Left(Trim(CStr(rssRegistro(2).Value)) + "        ", 8)
            ls_Producto = Left(Trim(CStr(rssRegistro(3).Value)) + "  ", 2)
            ls_SubProducto = Right("0000000000" + Trim(CStr(rssRegistro(4).Value)), 10)
            If rssRegistro(5).Value <> "" Then
                ls_FechaAlta = CStr(rssRegistro(5).Value)
                ls_FechaAlta = Mid(ls_FechaAlta, 1, 10)
            End If
            
            ls_TipoPeticion = Left(Trim(CStr(rssRegistro(8).Value)) + "0", 1)
            ls_Fecha = Left(Trim(CStr(rssRegistro(10).Value)) + Space$(8), 8)
            ls_Hora = Left(Replace(Trim(CStr(rssRegistro(11).Value)), ":", "") + Space$(4), 4)
            ls_IdTransaccion = Left(Trim(CStr(Format(rssRegistro(12).Value, "0000000000"))) + "       ", 10)
            ls_Tipo = Left(Trim(CStr(rssRegistro(13).Value)) + " ", 1)
            ' Logica para armado de una seccion del colector
            ' Se incluye la funcion a ser ejecutada por SQL cuando el msg
            ' con la respuesta sea recibido
            
            Ls_MsgColector = Left(Trim(strFuncionSQL) + "        ", 8)
            Ls_MsgColector = Ls_MsgColector + ls_Fecha + ls_Hora
            Ls_MsgColector = Ls_MsgColector + ls_TipoPeticion + ls_CentroRegional
            Ls_MsgColector = Ls_MsgColector + ls_NumRegistro + ls_Producto
            Ls_MsgColector = Ls_MsgColector + ls_SubProducto + ls_FechaAlta
            Ls_MsgColector = Ls_MsgColector + strFechaBaja + ls_IDFuncionario
            Ls_MsgColector = Ls_MsgColector + ls_IdTransaccion + ls_Tipo
            Ls_MsgColector = Ls_MsgColector + Space(43)
            
            
            If Len(Ls_MsgColector) > 0 Then
                ' Logica Armado ASTA ENTRADA
                Ls_MensajeMQ = ASTA_ENTRADA(Ls_MsgColector, " Funcionario: " & ls_IDFuncionario)
                ' Logica para el envio a la MQQUEUE
                If Ls_MensajeMQ <> "" Then
                    Escribe "Mensaje Enviado: " & Ls_MensajeMQ
                    If MQEnviarMsg(mqSession, mqManager, Gs_MQQueueEscritura, mqsEscribir, mqsMsgEscribir, Ls_MensajeMQ, strReplyToMQ, sPersistencia, sExpirar) Then
                        ReDim Preserve las_Funcionarios(NumeroMsgEnviados)
                        las_Funcionarios(NumeroMsgEnviados) = ls_IDFuncionario
                        NumeroMsgEnviados = NumeroMsgEnviados + 1
                    Else
                        'Registra en el log de SQL de la aplicación que el MQEnviarMsg tuvo error en el armado
                        psInsertaSQL gsAccesoActual, 5, "ProcesoBDtoMQQUEUEFunc. Error al escribir la solicitud en la MQ QUEUE: " & Gs_MQQueueEscritura & ". Error con el Funcionario: " & ls_IDFuncionario, "MSG", "MQEnviarMsg"
                    End If
                Else
                    'Registra en el log de SQL de la aplicación que el ASTA_ENTRADA tuvo error en el armado
                    psInsertaSQL gsAccesoActual, 4, "ProcesoBDtoMQQUEUEFunc. Error durante el armado del formato PS9 funcion ASTA_ENTRADA. Error con el Funcionario: " & ls_IDFuncionario, "MSG", "ASTA_ENTRADA"
                End If
            Else
                Escribe "Error al armar el Layout Alta-Mantenimiento-Baja de Funcionarios TKT-CED. Error con el Funcionario : " & ls_IDFuncionario
            End If
            rssRegistro.MoveNext
            
        Loop
        rssRegistro.Close
    Else
        Escribe "No existen registros en la consulta de los datos de tabla TMP_FUNCIONARIOS_PU. ProcesoBDtoMQQUEUEFunc"
    End If
   
    ' Cierre de la conexion con el MQQUEUE
    Call MQDesconectar(mqManager, mqsEscribir)
    
    If NumeroMsgEnviados > 0 Then
        If Not ActualizaRegistrosFunc(las_Funcionarios) Then
            Escribe "Existieron errores al actualizar la tabla TMP_FUNCIONARIOS_PU"
        End If
    End If
    
    Escribe "Envio de solicitures TKT -> Host Terminado. ProcesoBDtoMQQUEUEFunc"
    Escribe "Solicitudes enviadas a MQ: " & NumeroMsgEnviados
    Exit Function

ErrorProcesoBDMQ:
    Escribe "Se presentó un error durante la ejecución de la función ProcesoBDtoMQQUEUEFunc. Vea log y tabla TMP_FUNCIONARIOS_PU" & Err.Source & " %% " & Err.Description
End Function

Private Function ProcesoBDtoMQQUEUEAuto()
    Dim Ls_MensajeMQ         As String       ' Cadena con el mensaje armado con los registros de la base de datos
    Dim Ls_MsgColector       As String       ' Cadena para almecenar el COLECTOR
    Dim Ls_HeaderMsg         As String       ' Cadena para almacenar el HEADER del mensaje
    Dim strQuery             As String       ' Cadena para almacenar el Query a ejecutarse en la base de datos
    Dim NumeroMsgEnviados    As Integer      ' Contador para almacenar el número de mensajes procesados
    Dim las_Autorizaciones() As String    ' Arreglo para ingresar todos los registros que han sido enviados correctamente
    ' Para el armado de la solicitud
    Dim ls_Operacion         As String    '1  operacion
    Dim ls_Oficina           As String    '2  oficina
    Dim ls_NumeroFunc        As String    '3  codusu
    Dim ls_Transaccion       As String    '4  transaccion
    Dim ls_CodigoOperacion   As String    '5  tipo-oper
    Dim ls_Cuenta            As String    '6  cuenta-ced
    Dim ls_Divisa            As String    '7  divisa
    Dim ls_Importe           As String    '8  importe
    Dim ls_Fecha_Ope         As String    '9  Fecha (operacion)
    Dim ls_Folio_Ope         As String    '10 Folio
    Dim ls_Status_Envio      As String    '11 Status
    Dim ls_Fecha            As String
    Dim ls_Hora             As String

On Error GoTo ErrorProcesoBDMQ
    
    Escribe "Inicia el envío de mensajes a Host: " & gsAccesoActual & " Función: " & strFuncionSQL
    NumeroMsgEnviados = 0

   'Logica para recuperar los n mensajes de la tabla temporal en db.funcionario
   'Logica para procesar cada registro y convertirlo en un mensaje
    strQuery = "SELECT" & vbCrLf
    strQuery = strQuery + "operacion," & vbCrLf                         '0  operacion
    strQuery = strQuery + "oficina," & vbCrLf                           '1  oficina
    strQuery = strQuery + "numero_funcionario," & vbCrLf                '2  codusu
    strQuery = strQuery + "id_transaccion," & vbCrLf                    '3  transaccion
    strQuery = strQuery + "codigo_operacion," & vbCrLf                  '4  tipo-oper
    strQuery = strQuery + "cuenta," & vbCrLf                            '5  cuenta-ced
    strQuery = strQuery + "divisa," & vbCrLf                            '6  divisa
    strQuery = strQuery + "importe," & vbCrLf                           '7  importe
    strQuery = strQuery + "fecha_operacion," & vbCrLf                   '8  fecha
    strQuery = strQuery + "folio_autorizacion," & vbCrLf                '9  folio
    strQuery = strQuery + "status_envio," & vbCrLf                      '10 status
    strQuery = strQuery + "CONVERT(char(8),getdate(),112)," & vbCrLf    '11 fecha
    strQuery = strQuery + "CONVERT(char(5),getdate(),108)" & vbCrLf     '12 hora
    strQuery = strQuery + "FROM " & vbCrLf
    strQuery = strQuery + gsNameDB & "..TMP_AUTORIZACIONES_PU" & vbCrLf
    strQuery = strQuery + "WHERE status_envio = 0"
    rssRegistro.Open strQuery
    
    If Not rssRegistro.EOF Then
        If MQConectar(mqSession, Gs_MQManager, mqManager) Then
            blnConectado = True
        Else
            'Registra en el log de SQL de la aplicación que fue imposible conectarse con el MQManager
            psInsertaSQL gsAccesoActual, 3, "ProcesoBDtoMQQUEUEAuto. Fallo conexión MQ-Manager " & Gs_MQManager & ": " & mqSession.ReasonCode & " - " & mqSession.ReasonName, "MSG", "MQConectar"
            Exit Function
        End If
        
        Do While Not rssRegistro.EOF 'dbError
            ' Almacenado de variables
            ls_Operacion = Trim(Format(rssRegistro(0).Value, "0000000"))
            ls_Oficina = Trim(Format(rssRegistro(1).Value, "0000", 4))
            ls_NumeroFunc = Trim(CStr(rssRegistro(2).Value)) + Space(8 - Len(Trim(CStr(rssRegistro(2).Value))))
            ls_Transaccion = Trim(Format(rssRegistro(3).Value, "0000"))
            ls_CodigoOperacion = Trim(CStr(rssRegistro(4).Value + Space(3)))
            ls_Cuenta = Trim(CStr(rssRegistro(5).Value + Space(10)))
            ls_Divisa = Trim(CStr(rssRegistro(6).Value + Space(3)))
            ls_Importe = rssRegistro(7).Value
            ls_Fecha_Ope = rssRegistro(8).Value
            ls_Folio_Ope = Format(rssRegistro(9).Value, "000000000000")
            ls_Status_Envio = Trim(Format(rssRegistro(10).Value, "0"))
            
            ls_Fecha = Left(Trim(CStr(rssRegistro(11).Value)) + Space$(8), 8)
            ls_Hora = Left(Replace(Trim(CStr(rssRegistro(12).Value)), ":", "") + Space$(4), 4)
            
            ' Logica para armado de una seccion del colector
            ' Se incluye la funcion a ser ejecutada por SQL cuando el msg
            ' con la respuesta sea recibido
            
            Ls_MsgColector = Left(Trim(strFuncionSQL) + "        ", 8)
            Ls_MsgColector = Ls_MsgColector + ls_Fecha + ls_Hora
            Ls_MsgColector = Ls_MsgColector + ls_Operacion + ls_Oficina
            Ls_MsgColector = Ls_MsgColector + ls_NumeroFunc + ls_Transaccion
            Ls_MsgColector = Ls_MsgColector + ls_CodigoOperacion + ls_Cuenta
            Ls_MsgColector = Ls_MsgColector + ls_Divisa + ls_Importe
            Ls_MsgColector = Ls_MsgColector + ls_Fecha_Ope + ls_Folio_Ope
            
            If Len(Ls_MsgColector) > 0 Then
                ' Logica Armado ASTA ENTRADA
                Ls_MensajeMQ = ASTA_ENTRADA(Ls_MsgColector, " Operación: " & ls_Operacion)
                ' Logica para el envio a la MQQUEUE
                If Ls_MensajeMQ <> "" Then
                    Escribe "Mensaje Enviado: " & Ls_MensajeMQ

                    If MQEnviarMsg(mqSession, mqManager, Gs_MQQueueEscritura, mqsEscribir, mqsMsgEscribir, Ls_MensajeMQ, strReplyToMQ, sPersistencia, sExpirar) Then
                        ReDim Preserve las_Autorizaciones(NumeroMsgEnviados)
                        las_Autorizaciones(NumeroMsgEnviados) = ls_Operacion
                        NumeroMsgEnviados = NumeroMsgEnviados + 1
                    Else
                        'Registra en el log de SQL de la aplicación que el MQEnviarMsg tuvo error en el armado
                        psInsertaSQL gsAccesoActual, 5, "ProcesoBDtoMQQUEUEAuto. Error al escribir la solicitud en la MQ QUEUE: " & Gs_MQQueueEscritura & ". Error con la Operación: " & ls_Operacion, "MSG", "MQEnviarMsg"
                    End If
                Else
                    'Registra en el log de SQL de la aplicación que el ASTA_ENTRADA tuvo error en el armado
                    psInsertaSQL gsAccesoActual, 4, "ProcesoBDtoMQQUEUEAuto. Error durante el armado del formato PS9 funcion ASTA_ENTRADA. Error con la Operacion: " & ls_Operacion, "MSG", "ASTA_ENTRADA"
                End If
            Else
                Escribe "Error al armar el Layout Actualización del Autorizaciones TKT-CED. Error con la Operación : " & ls_Operacion
            End If
            rssRegistro.MoveNext

        Loop
        rssRegistro.Close
   Else
        Escribe "Cero registros en la consulta de los datos, tabla TMP_AUTORIZACIONES_PU. ProcesoBDtoMQQUEUEAuto"
   End If
   
    ' Cierre de la conexion con el MQQUEUE
    Call MQDesconectar(mqManager, mqsEscribir)
    
    If NumeroMsgEnviados > 0 Then
        If Not ActualizaRegistrosAuto(las_Autorizaciones) Then
            Escribe "Existieron errores al actualizar la tabla TMP_AUTORIZACIONES_PU"
        End If
    End If
    
    Escribe "Envio de solicitures TKT -> Host Terminado. ProcesoBDtoMQQUEUEAuto"
    Escribe "Solicitudes enviadas a MQ: " & NumeroMsgEnviados
    Exit Function

ErrorProcesoBDMQ:
    Escribe "Se presentó un error durante la ejecución de la función ProcesoBDtoMQQUEUEAuto. Vea log y tabla TMP_AUTORIZACIONES_PU"
End Function

'*************************************************************************************
' Valida el contenido de las variables de conexion a el sistema MQ
'*************************************************************************************
Private Function ValidaInfoMQ(ByRef ps_MsgVal As String) As Boolean
    Dim ls_msg As String
    
    If Trim(Gs_MQManager) = "" Then ls_msg = ls_msg & "Falta MQ Manager envio."
    If Trim(Gs_MQQueueEscritura) = "" Then ls_msg = ls_msg & "Falta MQ Queue envio."
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
    ' Asigna el MQManager que se encuentra en el archivo INI
    Gs_MQManager = ObtenParametroIni(Ls_Tit, "MQManager", "", Ls_Path)
    ' Asigna el MQQUEUE de envios que se encuentra en el archivo INI
    Gs_MQQueueEscritura = ObtenParametroIni(Ls_Tit, "MQEscritura", "", Ls_Path)
    gsEjecutable = ObtenParametroIni(Ls_Tit, "FGEjecutable", "", Ls_Path)
End Sub

Private Sub ConfiguraFileLog(Ls_Tit As String, Ls_Path As String)
    ' Asigna el nombre del archivo usado para almacenar el log de esta aplicación
    strlogFileName = ObtenParametroIni(Ls_Tit, "logFileName", "", Ls_Path)
    ' Asigna la ruta para almacenar el archivo del log de esta aplicación
    strlogFilePath = ObtenParametroIni(Ls_Tit, "logFilePath", "", Ls_Path)
    ' Asigna el valor de la bandera para escritura del log indicada en el INI
    Mb_GrabaLog = ObtenParametroIni(Ls_Tit, "Estatus", "", Ls_Path)
End Sub

Private Sub ConfiguraHeader_IH_ME(Ls_Path As String)
'Valor inicial MQ Message
    strFuncionHost = ObtenParametroIni("HEADERIH", "PRIMERVALOR", "", Ls_Path)
    strHeaderTagIni = ObtenParametroIni("HEADERIH", "IHTAGINI", "", Ls_Path)
    strIDProtocol = ObtenParametroIni("HEADERIH", "IDPROTOCOL", "", Ls_Path)
    strLogical = ObtenParametroIni("HEADERIH", "Logical", "", Ls_Path)
    strAccount = ObtenParametroIni("HEADERIH", "ACCOUNT", "", Ls_Path)
    strUser = ObtenParametroIni("HEADERIH", "User", "", Ls_Path)
    strSeqNumber = ObtenParametroIni("HEADERIH", "SEQNUMBER", "", Ls_Path)
    strTXCode = ObtenParametroIni("HEADERIH", "TXCODE" & gsEjecutable, "", Ls_Path)
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
    
    strFechaBaja = ObtenParametroIni("DEFAULTVALUES", "FECHABAJA", "", Ls_Path)
    strColectorMaxLeng = ObtenParametroIni("DEFAULTVALUES", "COLMAXLENG", "", Ls_Path)
    strMsgMaxLeng = ObtenParametroIni("DEFAULTVALUES", "MSGMAXLENG", "", Ls_Path)
    strPS9MaxLeng = ObtenParametroIni("DEFAULTVALUES", "PS9MAXLENG", "", Ls_Path)
    strReplyToMQ = ObtenParametroIni("DEFAULTVALUES", "ReplyToQueue", "", Ls_Path)
    Select Case gsEjecutable
        Case "F"
            strFuncionSQL = ObtenParametroIni("DEFAULTVALUES", "FuncionSQLF", "", Ls_Path)
        Case "A"
            strFuncionSQL = ObtenParametroIni("DEFAULTVALUES", "FuncionSQLA", "", Ls_Path)
    End Select
    strRndLogTerm = ObtenParametroIni("DEFAULTVALUES", "RandomLogTerm", "", Ls_Path)
    sPersistencia = ObtenParametroIni("DEFAULTVALUES", "PPERSISTENCE", "", Ls_Path)
    sExpirar = ObtenParametroIni("DEFAULTVALUES", "PEXPIRY", "", Ls_Path)
    
End Sub

Private Function ASTA_ENTRADA(strMsgColector As String, psTipo As String) As String
    Dim ls_TempColectorMsg      As String
    Dim ls_BloqueME             As String
    Dim ln_longCOLECTOR         As Integer
    Dim ln_AccTerminal          As Integer
    
    On Error GoTo ErrorASTA
    
    ' Armado del layout del COLECTOR
        'Datos de entrada (layout colector) Aquí vendrán los datos propios del layout de comunicación de extranjero.
        ls_TempColectorMsg = strMsgColector
            
    If Len(ls_TempColectorMsg) > CInt(Trim(strColectorMaxLeng)) Then  ' Se determina si el COLECTOR no excede el número de bytes maximo
        Escribe "La longitud del colector supera el maximo permitido"
        GoTo ErrorASTA
    End If
    
    ' HEADER DE MENSAJE (ME) Armado
    'X(04)   Bandera que indica el comienzo del mensaje
    ls_BloqueME = Left(Trim(strMETAGINI) + "    ", 4)
    '9(04)   Longitud del layout  del colector
    ls_BloqueME = ls_BloqueME + Right("0000" + CStr(Len(ls_TempColectorMsg)), 4)
    'X(01)   Tipo de mensaje: Copy
    ls_BloqueME = ls_BloqueME + Left(Trim(strMsgTypeCole) + " ", 1)
    ' Layout del COLECTOR
    ls_BloqueME = ls_BloqueME + ls_TempColectorMsg
    'X(05)   Bandera que indica el fin del mensaje
    ls_BloqueME = ls_BloqueME + Left(Trim(strMETAGEND) + "     ", 5)
    
    If Len(ls_BloqueME) > CInt(Trim(strMsgMaxLeng)) Then
        Escribe "La longitud del Bloque ME supera el maximo permitido"
        GoTo ErrorASTA
    End If

    
    'Para el uso de MQ-SERIES y CICSBRIDGE se requiere anteponer
    'al HEADER DE ENTRADA(IH) un valor que indique el programa
    'que invocara el CICSBRIDGE
    'X(08)  Indica el programa que invocara el CICSBRIDGE
    ASTA_ENTRADA = Left(Trim(strFuncionHost) + "        ", 8)
    
    'Header IH Armado
    'X(04)  Bandera que indica el comienzo del Header
    ASTA_ENTRADA = ASTA_ENTRADA + Left(Trim(strHeaderTagIni) + "    ", 4)
    'X(02)  Identificador  del protocolo (PS9)
    ASTA_ENTRADA = ASTA_ENTRADA + Left(Trim(strIDProtocol) + "  ", 2)
    'X(08)  Terminal Lógico Asigna Arquitectura ASTA
    If StrComp(Trim(strRndLogTerm), "1", vbTextCompare) = 0 Then
        ' Obtenemos un valor aleatorio en el rango[1-2000]
        ln_AccTerminal = 0
        Do While Not (ln_AccTerminal > 0 And ln_AccTerminal < 2000)
            Call Randomize(DateTime.DatePart("s", DateTime.Now))
            ln_AccTerminal = Rnd(DateTime.DatePart("s", DateTime.Now)) * 1000
        Loop
        'modificacion por formato en layout: ogj bbva bancomer 14 feb 2005
        'ASTA_ENTRADA = ASTA_ENTRADA + Left(CStr(ln_AccTerminal) + "        ", 8)
        ASTA_ENTRADA = ASTA_ENTRADA + Left(Format(CStr(ln_AccTerminal), "0000") + "        ", 8)
    Else
        ASTA_ENTRADA = ASTA_ENTRADA + Left(Trim(strLogical) + "        ", 8)
    End If
    'X(08)  Terminal Contable (CR Contable)
    ASTA_ENTRADA = ASTA_ENTRADA + Left(Trim(strAccount) + "        ", 8)
    'X(08)  Usuario. Debe ser diferente de espacions
    ASTA_ENTRADA = ASTA_ENTRADA + Left(Trim(strUser) + "        ", 8)
    'X(08)  Número de Secuencia (indicador de paginación)
    ASTA_ENTRADA = ASTA_ENTRADA + Left(Trim(strSeqNumber) + "        ", 8)
    'X(08)  Función específica Asigna Arquitectura Central
    ASTA_ENTRADA = ASTA_ENTRADA + Left(Trim(strTXCode) + "        ", 8)
    'X(02)  Tecla de función (no aplica)
    ASTA_ENTRADA = ASTA_ENTRADA + Left(Trim(strUserOption) + "  ", 2)
    '9(05)  Longitud total del mensaje = longitud del header de entrada +  Longitud de mensaje  (65+36+longitud del layout colector)
    ln_longCOLECTOR = 65 + Len(ls_BloqueME)
    If ln_longCOLECTOR > CInt(Trim(strPS9MaxLeng)) Then
        Escribe "La longitud del Layout PS9 supera el maximo permitido"
        GoTo ErrorASTA
    End If
    ASTA_ENTRADA = ASTA_ENTRADA + Right("00000" + CStr(ln_longCOLECTOR), 5)
    'X(01)  Indicador de commit: Permite realizar  commit
    ASTA_ENTRADA = ASTA_ENTRADA + Left(Trim(strCommit) + " ", 1)
    'X(01)  Tipo de mensaje: Nuevo requerimiento
    ASTA_ENTRADA = ASTA_ENTRADA + Left(Trim(strMsgType) + " ", 1)
    'X(01)   Tipo de proceso: on line
    ASTA_ENTRADA = ASTA_ENTRADA + Left(Trim(strProcessType) + " ", 1)
    'X(02)   Canal Asigna Arquitectura Central
    ASTA_ENTRADA = ASTA_ENTRADA + Left(Trim(strChannel) + "  ", 2)
    'X(01)   Indicador de preformateo: Arquitectura no deberá de preformatear los datos
    ASTA_ENTRADA = ASTA_ENTRADA + Left(Trim(strPreFormat) + " ", 1)
    'X(01)   Idioma: Español
    ASTA_ENTRADA = ASTA_ENTRADA + Left(Trim(strLenguage) + " ", 1)
    'X(05)   Bandera que indica el final del header
    ASTA_ENTRADA = ASTA_ENTRADA + Left(Trim(strHeaderTagEnd) + "     ", 5)
    ' Incorporacion del bloque ME al bloque IH
    ASTA_ENTRADA = ASTA_ENTRADA + ls_BloqueME
    
Exit Function

ErrorASTA:
    ASTA_ENTRADA = ""
    Escribe ("Error al armar el mensaje para " & psTipo)
End Function

Private Function FormatFecha(fecha As Date) As String
    FormatFecha = Right("00" + CStr(DateTime.DatePart("d", fecha)), 2)
    FormatFecha = FormatFecha + "-" + Right("00" + CStr(DateTime.DatePart("m", fecha)), 2)
    FormatFecha = FormatFecha + "-" + CStr(DateTime.DatePart("yyyy", fecha))
End Function

Private Function ActualizaRegistrosFunc(IDFuncionario() As String) As Boolean
    Dim strQueryUpDate As String
    Dim ln_indice As Integer
    
    ActualizaRegistrosFunc = False
    
On Error GoTo ErrorActRegistrosFunc
    
    ActualizaRegistrosFunc = True
    
    For ln_indice = 0 To UBound(IDFuncionario())
        strQueryUpDate = "UPDATE " & gsNameDB & "..TMP_FUNCIONARIOS_PU" & vbCrLf
        strQueryUpDate = strQueryUpDate + "SET  status_envio = 1" & vbCrLf
        strQueryUpDate = strQueryUpDate + "--  ,fecha_ultimo_mant = GETDATE()," & vbCrLf
        strQueryUpDate = strQueryUpDate + "WHERE status_envio = 0" & vbCrLf
        strQueryUpDate = strQueryUpDate + "AND id_funcionario = " & IDFuncionario(ln_indice)
        rssRegistro.Open strQueryUpDate
    Next ln_indice
    
    ActualizaRegistrosFunc = True
    Exit Function
    
ErrorActRegistrosFunc:
    Escribe "Error al realizar la actualización en la tabla TMP_FUNCIONARIOS_PU. Función ActualizaRegistrosFunc"
End Function

Private Function ActualizaRegistrosAuto(IDAutorizacion() As String) As Boolean
    Dim strQueryUpDate As String
    Dim ln_indice As Integer
    
    ActualizaRegistrosAuto = False
    
On Error GoTo ErrorActRegistrosAuto
    
    ActualizaRegistrosAuto = True
    
    For ln_indice = 0 To UBound(IDAutorizacion())
        strQueryUpDate = "UPDATE " & gsNameDB & "..TMP_AUTORIZACIONES_PU " & vbCrLf
        strQueryUpDate = strQueryUpDate + "SET  status_envio = 1 " & vbCrLf
        strQueryUpDate = strQueryUpDate + "WHERE status_envio = 0 " & vbCrLf
        strQueryUpDate = strQueryUpDate + "AND operacion = " & IDAutorizacion(ln_indice)
        rssRegistro.Open strQueryUpDate
    Next ln_indice
    
    ActualizaRegistrosAuto = True
    Exit Function
    
ErrorActRegistrosAuto:
    Escribe "Error al realizar la actualización en la tabla TMP_AUTORIZACION_PU. Función ActualizaRegistrosAuto"
End Function

Public Sub Desencriptar()
'***********************************************************************************************************************
'Objetivo:      Los datos encriptados de conexión a SQL Server son desencriptados en este procedimiento
'Autor:         EDS-BAGO
'Fecha:         19/Abr/2006
'***********************************************************************************************************************
    Dim ParamEncripcion  As MNICript.clsEncripta
    Set ParamEncripcion = New MNICript.clsEncripta

    gsNameDB = ParamEncripcion.VerificaClaves(2, gsNameDB)
    
    'Nombre de la base de datos de CATALOGOS
    gsCataDB = ParamEncripcion.VerificaClaves(2, gsCataDB)
    
    'Nombre del usuario de conexión a la BD
    gsUserDB = ParamEncripcion.VerificaClaves(2, gsUserDB)
    
    'Password de conexion
    gsPswdDB = ParamEncripcion.VerificaClaves(2, gsPswdDB)
    
    'Nombre del DSN
    gsDSNDB = ParamEncripcion.VerificaClaves(2, gsDSNDB)

    Set ParamEncripcion = Nothing

End Sub

'Función para guardar un valor de un archivo ini
Public Sub SaveProfile(lpFileName As String, lpAppName As String, lpKeyName As String, lpString As String)
    Dim Ret As Long
    Ret = WritePrivateProfileString(lpAppName, lpKeyName, lpString, lpFileName)
End Sub

Private Function ConectDB() As Boolean
'*******************************************************************************************************************
'Procedimiento: ConectDB
'Objetivo:      Generar y obtener los datos previos para realizar la conexión a la base de datos del Ticket
'               Los datos los obtiene del registro de windows. Verifica la cantidad de intentos de conexión que va a
'               realizar la aplicación en el periodo estipulado
'Autor:         EDS-AGO
'Fecha:         22/05/2006
'*******************************************************************************************************************

On Error GoTo Error
            
    gsCataDB = ObtenParametroIni("Conexion", "DBCata", "", gstrRutaIni)
    gsDSNDB = ObtenParametroIni("Conexion", "DBDSN", "", gstrRutaIni)
    gsUserDB = ObtenParametroIni("Conexion", "DBUser", "", gstrRutaIni)
    gsPswdDB = ObtenParametroIni("Conexion", "DBPswd", "", gstrRutaIni)
    gsNameDB = ObtenParametroIni("Conexion", "DBName", "", gstrRutaIni)
    
    Desencriptar
    
    Set cnnConexion = New Connection
    Set rssRegistro = New Recordset
    
    cnnConexion.ConnectionString = "DSN=" & gsDSNDB & _
                                    ";DB=" & gsCataDB & _
                                    ";UID=" & gsUserDB & _
                                    ";PWD=" & gsPswdDB
        
    cnnConexion.Open
    
    rssRegistro.ActiveConnection = cnnConexion
    
    ConectDB = True
    
    On Error GoTo 0
    
    Exit Function
        
Error:
    Escribe "Se presentó un error durante la conexión a la base de datos: " + Err.Number + "-" & Err.Description
    Escribe ""
    ConectDB = False
End Function

Private Function Desconectar()

    cnnConexion.Close
    
    Set rssRegistro = Nothing
    Set cnnConexion = Nothing
    
End Function

