VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cProcesaMSG"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private Declare Function GetPrivateProfileString Lib "kernel32" Alias "GetPrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpDefault As String, ByVal lpReturnedString As String, ByVal nSize As Long, ByVal lpFileName As String) As Long
Private Declare Function WritePrivateProfileString Lib "kernel32" Alias "WritePrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpString As Any, ByVal lpFileName As String) As Long
Private Declare Function GetPrivateProfileSection Lib "kernel32" Alias "GetPrivateProfileSectionA" _
(ByVal lpAppName As String, ByVal lpReturnedString As String, ByVal nSize As Long, ByVal lpFileName As String) As Long

Private ArchivoIni      As String
Private Ls_Servicio     As String
Private sBuffer         As String
'Public gsSeccRegWdw     As String
Private laAutoriz       As Autorizacion   'Estructura para almacenar el layout autorizacion
Private Ms_MQMngL       As String         'MQManager de Lectura
Private Ms_MQLeer       As String         'MQQueue   de Lectura
Private Ms_MQMngE       As String         'MQManager de Escritura
Private Ms_MQEscr       As String         'MQQueue   de Escritura
Private Gs_MsgRes       As String         'Respuesta al mensaje procesado
Private Ms_BanRetorno   As String         'BANDERA QUE NOS INDICA SI SE VA A REGRESAR UNA RESPUESTA
Private mbFuncionBloque As Boolean
Public Bandera          As String         'Indicador de encripción de constantes de conexión a SQL Server (1=encriptado; 0=no encriptado)


'------------------------------------------------------------------------------------------------
'Definicion de la estructura para los datos del proceso Autorizaciones
'------------------------------------------------------------------------------------------------
Private Type Autorizacion
    FuncionSQL  As String * 8       '   Función a buscar en la base de datos
    FechaProce  As String           '   Fecha del proceso
    HoraProce   As String           '   Hora del proceso
    DatosTemp   As String           '   Mensaje de prueba
    Espera      As String           '   Indica si existe mensjae en espera
End Type

Private Function fValidaLeyout(ByVal psFuncion As String, ByVal piLenMensaje As Integer, psMensaje) As Boolean
    Dim sParam As String
    Dim iLenParam As Integer
    Dim lsQuery As String

On Error GoTo Error
    
    fValidaLeyout = True
    
    lsQuery = "SELECT SUM(B.longitud)" & vbCrLf
    lsQuery = lsQuery + "FROM " & gsNameDB & "..TIPO_TRANSACCION_PIU A, " & gsNameDB & "..ENTRADA_PIU B " & vbCrLf
    lsQuery = lsQuery + "WHERE A.tipo_transaccion = B.tipo_transaccion " & vbCrLf
    lsQuery = lsQuery + "AND A.funcion = '" + Trim$(psFuncion) + "' " & vbCrLf
    
    rssRegistro.Open lsQuery
    
    If rssRegistro.EOF Then
        Escribe "Error no fue posible extraer la longitud de los parametros del Stored Procedure: " & Err.Description
        fValidaLeyout = False
        Exit Function
    Else
        iLenParam = CInt(rssRegistro(0).Value)
    End If
    
    rssRegistro.Close
    
    sParam = Mid(psMensaje, 1, Len(psMensaje))
    
    If iLenParam > Len(sParam) Then
        fValidaLeyout = False
        Exit Function
    End If
    
    Exit Function
        
Error:
    Escribe "Error al verificar los datos del layout contra los parametros del SP: " & Err.Description
    fValidaLeyout = False
End Function

Public Function ProcesarMensaje(ByVal strRutaIni As String, _
                                Optional ByVal strParametros As String)
                                
    Dim strFuncion      As String
    Dim Li_TotReg       As Integer
    Dim intLongBia      As Integer
    Dim Ls_MsgVal       As String
    Dim lsCommandLine   As String
    Dim Parametros()    As String       ' Arreglo para almacenar los parametros via línea de comando
    

On Error GoTo Form_Load_Err

    ArchivoIni = strRutaIni & "\TKTMQ.ini"
'Escribe "ArchivoIni:" & ArchivoIni
    
    gstrRutaIni = strRutaIni
    
    ' Recupera la información del archivo ini dedicada a
    'la configuración del archivo log de la aplicación
'Escribe "Antes de entrar a la funcion Escribe ConfiguraFileLog"
    ConfiguraFileLog "EscribeArchivoLOG", ArchivoIni
'Escribe "Despues de entrar a la funcion Escribe ConfiguraFileLog"

    'Obtiene la sección del registro de windows donde se encuentra la información de conexión
'    gsSeccRegWdw = ObtenParametroIni("CONEXIONSQL", "SECCREGWDW", "", ArchivoIni)
    
    'se inicializa la variable booleana
    mbFuncionBloque = False
    
'Escribe "Antes de obtener los parametros"
    ' Recupera parametros
    lsCommandLine = Trim(strParametros)
'Escribe "Despues de obtener los parametros: " & lsCommandLine
    
'Escribe "If para validar si hay o no valores en los parametros"
    If StrComp(lsCommandLine, "", vbTextCompare) <> 0 Then
'Escribe "Dentro del If, Si hay valores en parametros"
        Erase Parametros()
        Parametros = Split(lsCommandLine, "-")
        
        Ms_MQMngL = Trim(Parametros(0))
'Escribe "Split del primer valor Ms_MQMngL: " & Ms_MQMngL

        Ms_MQMngE = Trim(Parametros(0))
'Escribe "Split del segundo valor Ms_MQMngE: " & Ms_MQMngE

        Ms_MQLeer = Trim(Parametros(1))
'Escribe "Split del tercer valor Ms_MQLeer: " & Ms_MQLeer

        Ms_MQEscr = Trim(Parametros(2))
'Escribe "Split del cuarto valor Ms_MQEscr: " & Ms_MQEscr

        Ms_BanRetorno = Trim(Parametros(3))
'Escribe "Split del quinto valor Ms_BanRetorno: " & Ms_BanRetorno

    Else
'Escribe "En caso de que no existan los parametros los toma del Ini"
        ' Recupera la información del archivo ini dedicada a la configuración de la MQ
        ObtenerInfoMQ "MQ Series", ArchivoIni
    End If
'Escribe "saliendo del IF"
    'Solo hasta este punto la conexión ha sido establecida

    
'Escribe "Llama a la funcion ValidaInfoMQ y le envia Ls_MsgVal: " & Ls_MsgVal
    'Verifica la existencia de la información de la Queue
    If Not ValidaInfoMQ(Ls_MsgVal) Then
        
        Escribe "Se presentó un error en la información de validación de configuración de la sección de MQSeries."
    
        Escribe "Termina el acceso a la aplicación TKTMQ."
    
        Escribe ""
    
        Exit Function
    End If
    
'Escribe "Conecta a MQ manager:" & Ms_MQMngL

    If MQConectar(mqsConexion, Ms_MQMngL, mqsManager) Then
'Escribe "Se conecto a MQ"
        blnConectado = True
    Else
        Escribe "Se presentó un error en la conexión al servidor de MQSeries."
        Escribe "Termina el acceso a la aplicación TKTMQ."
        Escribe ""
        Exit Function
    End If
    
'Escribe "Llama a la funcion MQRevisar"
    ' Revisamos si existen datos en la QUEUE de lectura
    If MQRevisar(mqsConexion, mqsManager, Ms_MQLeer, mqsLeer, mqsMsgLeer) = 0 Then
'Escribe "Ok, en la funcion MQRevisar"
'Escribe "Llama a la funcion MQDesconectar"
        MQDesconectar mqsManager, mqsEscribir, mqsLeer
'Escribe "Despues de llamar a la funcion MQDesconectar"
        Escribe "Termina TKTMQ Recepción de Solicitudes. Al revisar la QUEUE no encontramos mensajes " '+ gsAccesoActual
        Escribe ""
        Exit Function
    End If
        
'Escribe "Antes de conectarse a SQL"
    'Verifica si hay conexión al servidor SQL
    If ConectDB = False Then Exit Function
'Escribe "Despues de conectarse a SQL"

'Escribe "llamado a la funcio MQRecibir, con los valores Ms_MQLeer:" & Ms_MQLeer
    ' Recuperación del mensaje
    If MQRecibir(mqsConexion, mqsManager, Ms_MQLeer, mqsLeer, mqsMsgLeer) Then
'Escribe "todo, ok. llamado de funcion MQDesconectar"
        MQDesconectar mqsManager, mqsEscribir, mqsLeer
'Escribe "Termino de funcion"
'Escribe "valor de variable Mb_Detalles:" & Mb_Detalles
        If Mb_Detalles Then Escribe "Mensaje recibido correctamente"
    Else
        MQDesconectar mqsManager, mqsEscribir, mqsLeer
        If Mb_Detalles Then psInsertaSQL 8, "No se realizo la recepción del mensaje", "TKT", "MQRecibir"
        Escribe ""
        Desconectar
        Exit Function
    End If
      
    Gs_MsgRes = "" ' Borrado del mensaje de respuesta
        
'Escribe "Llamado a la funcion ProcesMessage"
    ' Procesamiento del mensaje
    If Not ProcesMessage(mqsMsgLeer) Then
        GoTo OpNoCompletada
    End If
'Escribe "Continua con la validacion If Ms_BanRetorno = 0 Or mbFuncionBloque = True Then"
    ' Envio del mensaje cuando fue procesado correctamente
    If Ms_BanRetorno = 0 Or mbFuncionBloque = True Then
        Escribe "Mensaje procesado correctamente. Termina TKTMQ Recepción de Solicitudes. Función SQL: " & laAutoriz.FuncionSQL
        Escribe ""
        Desconectar
        Exit Function
    End If
                
    If MQConectar(mqsConexion, Ms_MQMngE, mqsManager) Then
        blnConectado = True
        If MQEnviar(mqsConexion, mqsManager, Ms_MQEscr, mqsEscribir, mqsMsgEscribir, Gs_MsgRes) Then
            Call MQDesconectar(mqsManager, mqsEscribir, mqsLeer)
        Else
            Call MQDesconectar(mqsManager, mqsEscribir, mqsLeer)
            psInsertaSQL 5, "No se pudo enviar el mensaje", "TKT", "MQEnviar"
            GoTo OpNoCompletada
        End If
    Else
        psInsertaSQL 3, "Error al abrir la MQ Queue de respuesta", "TKT", "MQConectar"
        GoTo OpNoCompletada
    End If
       
   ' Desconexión con el MQ Manager
    Call MQDesconectar(mqsManager, mqsEscribir, mqsLeer)
    ' Desconexión con la base de datos
    Desconectar
      
    'Termina la recepción de solicitudes
    Escribe "Mensaje Regresado correctamente. Termina TKTMQ Recepción de Solicitudes."
    Escribe ""
      
    Exit Function

OpNoCompletada:
    Desconectar
    'AGO-EDS. Condición generada porque se estaba escribiendo el encabezado en el archivo log como error cuando no existía
    'cuerpo
    If Gs_MsgRes = "ENCABEZADO" Then
        Escribe "Termina TKTMQ Recepción de Solicitudes."
        Escribe ""
        Exit Function
    End If
    
    If Mb_Detalles = 0 Then
        Escribe "El procesamiento del mensaje no pudo ser completado correctamente, para ver mas detalles active Mb_Detalles = 1 en el archivo .Ini."
        Escribe ""
    End If
    
    Exit Function
    
Form_Load_Err:
    Escribe "Termina el acceso a la aplicación TKTMQ porque se presentó un error en la función MAIN. Error. " & CStr(Err.Number) & "-" & Err.Description
    Escribe ""
    Desconectar
End Function

Private Sub ObtenerDatosConexion(Ls_Tit As String, Ls_Path As String)
    'Obtiene datos de configuración del LOG
   lsLogPath = ObtenParametroIni("EscribeArchivoLOG", "LogPath", "", Ls_Path)
   lsLogName = ObtenParametroIni("EscribeArchivoLOG", "LogFile", "", Ls_Path)
   
   Mb_GrabaLog = ObtenParametroIni("EscribeArchivoLOG", "Estatus", "", Ls_Path)
   Mb_Detalles = ObtenParametroIni("EscribeArchivoLOG", "DETALLES", "", Ls_Path)
End Sub



Private Function ConectDB() As Boolean
'*******************************************************************************************************************
'Procedimiento: ConectDB
'Objetivo:      Generar y obtener los datos previos para realizar la conexión a la base de datos del Ticket
'               Los datos los obtiene del registro de windows. Verifica la cantidad de intentos de conexión que va a
'               realizar la aplicación en el periodo estipulado
'Autor:         EDS-DGI
'Fecha:         19/06/2006
'*******************************************************************************************************************

On Error GoTo Error
    
'Escribe "Obtengo valores del INI"
    gsCataDB = ObtenParametroIni("Conexion", "DBCata", "", gstrRutaIni & "\TKTMQ.ini")
'Escribe "Obtengo gsCataDB:" & gsCataDB

    gsDSNDB = ObtenParametroIni("Conexion", "DBDSN", "", gstrRutaIni & "\TKTMQ.ini")
'Escribe "Obtengo gsDSNDB:" & gsDSNDB

    gsUserDB = ObtenParametroIni("Conexion", "DBUser", "", gstrRutaIni & "\TKTMQ.ini")
'Escribe "Obtengo gsUserDB:" & gsUserDB

    gsPswdDB = ObtenParametroIni("Conexion", "DBPswd", "", gstrRutaIni & "\TKTMQ.ini")
'Escribe "Obtengo gsPswdDB:" & gsPswdDB

    gsNameDB = ObtenParametroIni("Conexion", "DBName", "", gstrRutaIni & "\TKTMQ.ini")
'Escribe "Obtengo gsNameDB:" & gsNameDB
        
    Set ParamEncripcion = New MNICript.clsEncripta
    
    
    Desencriptar
    
    
    Set ParamEncripcion = Nothing
    
    Set cnnConexion = New Connection
    Set rssRegistro = New Recordset
    
'Escribe "ANTES DE CONECTARME " & "DSN=" & gsDSNDB & _
                                    ";DB=" & gsCataDB & _
                                    ";UID=" & gsUserDB & _
                                    ";PWD=" & gsPswdDB
                                    
    cnnConexion.ConnectionString = "DSN=" & gsDSNDB & _
                                    ";DB=" & gsCataDB & _
                                    ";UID=" & gsUserDB & _
                                    ";PWD=" & gsPswdDB


    cnnConexion.Open

'Escribe "Conectado"

    rssRegistro.ActiveConnection = cnnConexion

    ConectDB = True
'Escribe "Termino de conectarme"

    On Error GoTo 0
    
    Exit Function
        
Error:
    Escribe "Se presentó un error durante la conexión a la base de datos: " + Err.Number + "-" & Err.Description
    Escribe ""
    ConectDB = False
End Function

Public Function ProcesMessage(ByRef objMsgLeido As MQMessage) As Boolean
    Dim ls_mensaje  As String
    Dim lsQuery     As String
    Dim lsDatos     As String
    Dim lsNewMsg    As String
    Dim Ls_Servicio As String
    Dim psBodyMsg As String
    Dim psLongitud As String
    Dim psTipo As String
    
    Dim sBodyMsg As String
    Dim sAux As String
    
    Dim lsMsgErr As String
    Dim sMensajeBloque As Boolean
    Dim Ld_CodigoExecNTHOST     As Double
    Dim EjecutableMSG           As String
    
    'variable contador de emensaje
    Dim iMensaje    As Integer
    Dim lbYaProcesado As Boolean
    
    
On Error GoTo ProcesMsgError
    ProcesMessage = False
    sMensajeBloque = False
    mbFuncionBloque = False
    
    'Obtenemos el mensaje contenido en el objeto Mensaje
    ls_mensaje = objMsgLeido.MessageData
    'AGO-EDS
    'Esta línea condicionante fué creada porque se estaban procesando mensajes que solo tenían el encabezado y no contaban
    'con cuerpo. El procesamiento marcaba un error que era escrito en el archivo .log
    If InStr(1, ls_mensaje, "<OC>", vbTextCompare) = 0 Then
        Gs_MsgRes = "ENCABEZADO"
        Exit Function
    End If
    
    sBodyMsg = Mid(ls_mensaje, InStr(1, ls_mensaje, "<OC>", vbTextCompare), Len(ls_mensaje))
    
    iMensaje = 0
    Do While sBodyMsg <> ""
        lbYaProcesado = False
        sAux = Mid(sBodyMsg, 1, InStr(1, sBodyMsg, "</OC>", vbTextCompare) + 4)
        
        'Descompone el mensaje en una estructura PS9
        psMsgAStructPS9 sAux, psBodyMsg, psLongitud, psTipo
        laAutoriz.FuncionSQL = Left$(psBodyMsg, 8)
        
        Ls_Servicio = Fl_ObtieneServicio(laAutoriz.FuncionSQL, "SinProcesar")
        If Len(Ls_Servicio) <> 0 Then
            lsDatos = laAutoriz.DatosTemp
            ReArmaPS9 lsDatos, ls_mensaje, psBodyMsg, lsNewMsg, psTipo
            Gs_MsgRes = lsNewMsg
            Escribe "La funcion pertenece al bloque de pruebas (no procesar): " + laAutoriz.FuncionSQL
            Exit Function
        End If
        
        Ls_Servicio = Fl_ObtieneServicio(Trim$(laAutoriz.FuncionSQL), "Servicios")
        If Len(Ls_Servicio) = 0 Then
            Ls_Servicio = Fl_ObtieneServicio(Trim$(laAutoriz.FuncionSQL), "FUNCIONBLOQUES")
            If Len(Ls_Servicio) <> 0 Then
                mbFuncionBloque = True
                sMensajeBloque = True
                laAutoriz.Espera = Mid$(psBodyMsg, 9, 1)
                laAutoriz.FechaProce = Mid$(psBodyMsg, 10, 10)
                laAutoriz.HoraProce = Mid$(psBodyMsg, 21, 8)
                laAutoriz.DatosTemp = psBodyMsg
            Else
                Escribe "La función " + laAutoriz.FuncionSQL + " no existe"
            End If
        Else
            sMensajeBloque = False
            laAutoriz.FechaProce = Mid$(psBodyMsg, 9, 8)
            laAutoriz.HoraProce = Mid$(psBodyMsg, 17, 8)
            laAutoriz.DatosTemp = psBodyMsg
        End If
        
        If Not fValidaLeyout(laAutoriz.FuncionSQL, IIf(sMensajeBloque = True, Len(psBodyMsg) - 8, (psLongitud)), psBodyMsg) Then
            psInsertaSQL 4, "La longitud de los parametros en el mansaje, no es compatible con la longuitud de los parametros para generar la ejecución del sp", "TKT", "ProcesMessage"
            GoTo NEXTOC
        End If
        
        'Genera el Query funcional de la aplicación para el SP
        If Not BuscaSPyLongPar(laAutoriz.FuncionSQL, IIf(sMensajeBloque = True, Len(psBodyMsg) - 8, (psLongitud)), psBodyMsg, lsQuery) Then
            ReArmaPS9 "TKT1010 FUNCION NO EXISTENTE", ls_mensaje, psBodyMsg, lsNewMsg, psTipo
            GoTo NEXTOC
        End If
        
        'Ejecuta el Stored procedure y regresa los datos, resultado de la operación
        lsDatos = ""
        If Not EjecSPFuncion(lsQuery, lsDatos) Then
            ReArmaPS9 lsDatos, ls_mensaje, psBodyMsg, lsNewMsg, psTipo
            lbYaProcesado = True
            GoTo NEXTOC
        End If
        If Mb_Detalles Then
            If lsDatos <> Space(0) Then
                Escribe "Resultado de la ejecución del SP: " + lsDatos
            Else
                Escribe "El SP no regresa datos"
            End If
        End If

'Etiqueta para siguiente bloque
NEXTOC:
        'Verifica si existen mas cadenas OC
        sBodyMsg = Mid(sBodyMsg, Len(sAux) + 1, Len(sBodyMsg))
        
        iMensaje = iMensaje + 1
    Loop
    
    If sMensajeBloque = True Then
        Escribe "Total de mensajes <OC> obtenidos : " & iMensaje
    End If
    
    If Not lbYaProcesado Then
        'Rearma el PS9 con los datos de la respuesta (Proceso Completo)
        ReArmaPS9 lsDatos, ls_mensaje, psBodyMsg, lsNewMsg, psTipo
    End If
    
    If sMensajeBloque = True Then
        If laAutoriz.Espera = 1 Then
        
            Dim objBitacoras As Object
            Set objBitacoras = CreateObject("Bitacoras.cBitacoras")
            objBitacoras.ProcesarBitacora gstrRutaIni, Ms_MQMngE + "-" + Ms_MQEscr + "-" + "1" + "-" + laAutoriz.FuncionSQL
            Set objBitacoras = Nothing
        
        End If
    End If
    
    'Configuración del mensaje a enviar a HOST
    Gs_MsgRes = lsNewMsg
    ProcesMessage = True
    Exit Function
    
ProcesMsgError:
    psInsertaSQL 4, "Error al procesar el mensaje " + laAutoriz.FuncionSQL + " (ProcesMessage): " + Err.Description, "TKT", "ProcesMessage"
End Function

Private Sub ReArmaPS9(ByVal psDatos As String, ByVal psMensaje As String, ByVal psColector As String, ByRef psNuevoMsg As String, ByVal psTipo As String)
    Dim LnLongColector  As Integer
    Dim Header As String
    Dim sTMsg As String
    Dim sLMsg As String
    Dim sColector As String
    Dim sCero As String
    Dim sFlujo As String
    
    'Determinamos la longitud del colector
    LnLongColector = Len(psDatos)
    
    Select Case psTipo
        Case "<ME>"
            
            Header = Mid(psMensaje, 1, InStr(1, psMensaje, psTipo) - 1) + "<HE>"
            sLMsg = Format(CStr(LnLongColector), "0000")
            sTMsg = Mid(psMensaje, InStr(1, psMensaje, psTipo) + 8, 1)
            sColector = Left(psDatos + Space(LnLongColector), LnLongColector) + "</HE>"
            psNuevoMsg = Header + sLMsg + sTMsg + sColector
                    
        Case "<OC>"
            
            Header = Mid(psMensaje, 1, 4)
            sFlujo = Mid(psMensaje, 5, 5)
            sLMsg = Format(CStr(LnLongColector), "0000")
            sTMsg = Mid(psMensaje, 14, 1)
            sColector = Left(psDatos + Space(LnLongColector), LnLongColector) + "</OH>"
            psNuevoMsg = Header + sFlujo + sLMsg + sTMsg + sColector
            
    End Select
    
    Escribe ("Respuesta generada : " & psNuevoMsg)
    
End Sub

Private Function psMsgAStructPS9(ByVal psMsg As String, ByRef psBodyMsg As String, ByRef psLongitud As String, ByRef psTipo As String)
Dim iIni As Integer
Dim iFin As Integer
Dim iExisteHE As Integer
Dim iExisteOC As Integer
Dim sEncabezado As String

    iExisteHE = InStr(1, psMsg, "<ME>")
    iExisteOC = InStr(1, psMsg, "<OC>")
    
    If iExisteHE > 0 Then
        psTipo = "<ME>"
    
        'MENSAJE DE ORIGEN NT --> HOST
        iIni = InStr(1, psMsg, "<ME>")
        iFin = InStr(1, psMsg, "</ME>")
        
        sEncabezado = Mid(psMsg, iIni, 4)
        psLongitud = Mid(psMsg, iIni + Len(sEncabezado), 5)
        psBodyMsg = Mid(psMsg, (iIni + Len(sEncabezado) + Len(psLongitud)), (iFin) - (iIni + Len(sEncabezado) + Len(psLongitud)))
        psLongitud = Mid(psLongitud, 1, 4)
            
    ElseIf iExisteOC > 0 Then
        psTipo = "<OC>"

        'MENSAJE DE ORIGEN HOST --> NT
        iIni = InStr(1, psMsg, "<OC>")
        iFin = InStr(1, psMsg, "</OC>")
        
        sEncabezado = Mid(psMsg, iIni, 4)
        'psLongitud = Mid(psMsg, iIni + Len(sEncabezado) + 5, 4)
        psLongitud = Len(Mid(psMsg, iIni + Len(sEncabezado) + 5 + 5, Len(psMsg) - 19))
        psBodyMsg = Mid(psMsg, (iIni + Len(sEncabezado) + 10), (iFin) - (iIni + Len(sEncabezado) + 10))
    End If
End Function

Private Function BuscaSPyLongPar(ByVal psFuncion As String, ByVal pnLongPar As Integer, ByVal psBia As String, ByRef psQuery As String) As Boolean
    Dim lsQuery     As String
    Dim lsParam     As String
    Dim lnCont      As Integer
    
    On Error GoTo BuscaSPyLongPar_Err:
    'Inicializa la función a falso
    BuscaSPyLongPar = False
    
    'Obtiene el Nombre del Servicio a partir del Nombre de la Funcion
    Ls_Servicio = Fl_ObtieneServicio(Trim$(psFuncion), "Servicios")
    'Si no exite el Servicio lo guarda para ponerlo en el Log.
    If Len(Ls_Servicio) = 0 Then
        Ls_Servicio = Fl_ObtieneServicio(Trim$(psFuncion), "FUNCIONBLOQUES")
        If Len(Ls_Servicio) = 0 Then
            Escribe "La función: " + psFuncion + " no existe."
            Exit Function
        End If
    Else
        Escribe "Funcion a ejecutar: " + psFuncion
    End If
            
    'Prepara el query
    lsQuery = "SELECT" & vbCrLf
    lsQuery = lsQuery + "A.stored_procedure," & vbCrLf
    lsQuery = lsQuery + "B.longitud, " & vbCrLf
    lsQuery = lsQuery + "B.orden_campo" & vbCrLf
    lsQuery = lsQuery + "FROM" & vbCrLf
    lsQuery = lsQuery + gsNameDB & "..TIPO_TRANSACCION_PIU A, " & vbCrLf
    lsQuery = lsQuery + gsNameDB & "..ENTRADA_PIU B " & vbCrLf
    lsQuery = lsQuery + "WHERE" & vbCrLf
    lsQuery = lsQuery + "A.tipo_transaccion = B.tipo_transaccion " & vbCrLf
    lsQuery = lsQuery + "AND A.funcion = '" + Trim$(psFuncion) + "' " & vbCrLf
    lsQuery = lsQuery + "ORDER BY orden_campo"

    'Ejecuta query
    rssRegistro.Open lsQuery
    
    'No hubo error
    If rssRegistro.EOF Then
        'Obtiene error
        Escribe "Error al ejecutar query: " + lsQuery
        Exit Function
    Else
        'Inicializa datos de operación query
    
        lnCont = 0
        lsQuery = "EXEC " & gsNameDB & ".." & RemplazaCaracter(QuitaApostrofe(Trim$(rssRegistro(0).Value))) + " "
        lsParam = Mid(psBia, 1, pnLongPar)
        'Arma cadena de parámetros
        While Not rssRegistro.EOF
            If lnCont > 0 Then
                lsQuery = lsQuery + ", "
            End If
            lsQuery = lsQuery + "'" + RemplazaCaracter(QuitaApostrofe(DaParamBia(Val(rssRegistro(1).Value), lsParam))) + "'"
            
            rssRegistro.MoveNext
            
            lnCont = lnCont + 1
        Wend
    End If
    
    'Termina el query
    rssRegistro.Close
    
    If Mb_Detalles = True Then
        If Mb_Detalles Then Escribe "Stored Procedure a ser ejecutado: " + lsQuery
    End If
        
    'Regresa verdadera la función
    BuscaSPyLongPar = True
    psQuery = lsQuery
    Exit Function

BuscaSPyLongPar_Err:
    Escribe "Error en el armado del SP (BuscaSPyLongPar). Error: " & CStr(Err.Number) + ", Descripción: " + Err.Description + ", Query generado: " + lsQuery
    BuscaSPyLongPar = False
End Function

Private Function ValidaInfoMQ(ByRef ps_MsgVal As String) As Boolean
    Dim ls_Msg As String
    If Trim(Ms_MQMngL) = "" Then ls_Msg = "Falta MQ Manager Origen."
    If Trim(Ms_MQMngE) = "" Then ls_Msg = ls_Msg & IIf(Len(ls_Msg) > 0, Chr(13), "") & "Falta MQ Manager Destino."
    If Trim(Ms_MQLeer) = "" Then ls_Msg = ls_Msg & IIf(Len(ls_Msg) > 0, Chr(13), "") & "Falta MQ Queue Origen."
    If Trim(Ms_MQEscr) = "" Then ls_Msg = ls_Msg & IIf(Len(ls_Msg) > 0, Chr(13), "") & "Falta MQ Queue Destino."
    If Trim(ls_Msg) = "" Then ValidaInfoMQ = True
    ps_MsgVal = ls_Msg
End Function

Private Function ObtenParametroIni(Ls_Grupo As String, Ls_Variable As String, Ls_Default As String, Ls_AppPath As String) As String
   Dim Ls_Buffer As String * 80
   Dim Li_Long As Integer
   Li_Long = GetPrivateProfileString(ByVal Ls_Grupo, ByVal Ls_Variable, ByVal Ls_Default, ByVal Ls_Buffer, Len(Ls_Buffer), Ls_AppPath)
   Li_Long = InStr(Ls_Buffer, Chr(0))
   ObtenParametroIni = Trim(Left(Ls_Buffer, Li_Long - 1))
   Exit Function
End Function

Private Sub ObtenerInfoMQ(Ls_Tit As String, Ls_Path As String)
'Escribe "Entra a la funcion ObtenerInfoMQ, con los siguientes valores:"
'Escribe "Ls_Tit:" & Ls_Tit & " Ls_Path:" & Ls_Path

   Ms_MQMngL = ObtenParametroIni(Ls_Tit, "MQManagerL", "", Ls_Path)
'Escribe "En caso de que no existan los parametros los toma del Ini"

   Ms_MQMngE = ObtenParametroIni(Ls_Tit, "MQManagerE", "", Ls_Path)
'Escribe "En caso de que no existan los parametros los toma del Ini"

   Ms_MQLeer = ObtenParametroIni(Ls_Tit, "MQLectura", "", Ls_Path)
'Escribe "En caso de que no existan los parametros los toma del Ini"

   Ms_MQEscr = ObtenParametroIni(Ls_Tit, "MQEscritu", "", Ls_Path)
'Escribe "En caso de que no existan los parametros los toma del Ini"

   Ms_BanRetorno = ObtenParametroIni(Ls_Tit, "MQBanRetorno", "", Ls_Path)
'Escribe "En caso de que no existan los parametros los toma del Ini"

'Escribe "Saliendo de la funcion ObtenerInfoMQ"

End Sub

Private Sub ConfiguraFileLog(Ls_Tit As String, Ls_Path As String)
    'Obtiene datos de configuración del LOG
'Escribe "Dentro de la funcion ConfiguraFileLog"
'Escribe "antes de leer lsLogPath"
   lsLogPath = ObtenParametroIni("EscribeArchivoLOG", "LogPath", "", Ls_Path)
'Escribe "despues de leer lsLogPath:" & lsLogPath

'Escribe "antes de leer lsLogName"
   lsLogName = ObtenParametroIni("EscribeArchivoLOG", "LogFile", "", Ls_Path)
'Escribe "despues de leer lsLogName:" & lsLogName

'Escribe "antes de leer Mb_GrabaLog"
   Mb_GrabaLog = ObtenParametroIni("EscribeArchivoLOG", "Estatus", "", Ls_Path)
'Escribe "despues de leer Mb_GrabaLog:" & Mb_GrabaLog

'Escribe "antes de leer Mb_Detalles"
   Mb_Detalles = ObtenParametroIni("EscribeArchivoLOG", "DETALLES", "", Ls_Path)
'Escribe "despues de leer Mb_Detalles:" & Mb_Detalles
'Escribe "Saliendo de la funcion ConfiguraFileLog"
End Sub

'******************************************************************************************************
'******************************************************************************************************
'                                  clase SrvInterface
'******************************************************************************************************
'******************************************************************************************************
Private Function Fl_ObtieneServicio(Ls_Funcion As String, ByVal psBloque As String) As String

    Dim Ls_Linea    As String
    Dim Ln_Ind      As Byte
    Dim Ls_Servicio As String
    Dim tContenidos As Collection
    Dim i           As Long
    
    Fl_ObtieneServicio = ""
    Set tContenidos = IniGetSection(ArchivoIni, psBloque)
    If Not tContenidos Is Nothing Then
        For i = 1 To tContenidos.Count
            Ln_Ind = InStr(1, tContenidos(i), Trim(Ls_Funcion))
            If Ln_Ind > 0 Then
                Ln_Ind = InStr(1, tContenidos(i), ",")
                Ls_Servicio = Mid(tContenidos(i), 1, Ln_Ind - 1)
                Fl_ObtieneServicio = Ls_Servicio
                Exit For
            End If
        Next i
    Else
    End If
End Function

Private Function QuitaApostrofe(lsline As String) As String
    
    Dim CadenaSinApostrofe  As String
    Dim sOtroApostrofe      As String
    Dim lscad1              As String
    Dim lscad2              As String
    Dim r                   As String
    Dim lnposicion          As Integer
    Dim i                   As Integer
    Dim n1                  As Integer

    Do While InStr(1, lsline, "'") <> 0
        n1 = Len(lsline)
        lnposicion = InStr(1, lsline, "'")
        lnposicion = lnposicion - 1
        lscad1 = Mid(lsline, 1, lnposicion)
        lscad2 = Mid(lsline, lnposicion + 2)
        CadenaSinApostrofe = lscad1 & "`" & lscad2
        lsline = CadenaSinApostrofe
    Loop
    
    Do While InStr(1, lsline, """") <> 0
        n1 = Len(lsline)
        lnposicion = InStr(1, lsline, """")
        lnposicion = lnposicion - 1
        lscad1 = Mid(lsline, 1, lnposicion)
        lscad2 = Mid(lsline, lnposicion + 2)
        CadenaSinApostrofe = lscad1 & "`" & lscad2
        lsline = CadenaSinApostrofe
    Loop
    
    QuitaApostrofe = lsline
End Function

Private Function RemplazaCaracter(lsline As String) As String

    Dim CadenaSinSimbolos   As String
    Dim lnposicion          As Integer
    Dim n1                  As Integer
    Dim lscad1              As String
    Dim lscad2              As String
    
    Do While InStr(1, lsline, "¥") <> 0
        n1 = Len(lsline)
        lnposicion = InStr(1, lsline, "¥")
        lnposicion = lnposicion - 1
        lscad1 = Mid(lsline, 1, lnposicion)
        lscad2 = Mid(lsline, lnposicion + 2)
        CadenaSinSimbolos = lscad1 & "Ñ" & lscad2
        lsline = CadenaSinSimbolos
    Loop
    RemplazaCaracter = lsline
End Function

Private Function DaParamBia(ByVal pnLong As Integer, ByRef psParam As String) As String
   
    On Error GoTo DaParamBia_Err:
    DaParamBia = ""                             'Inicia variables
    DaParamBia = Mid(psParam, 1, pnLong)        'Obtiene el parámetro
    psParam = Mid(psParam, pnLong + 1)          'Prepara la cadena para el siguiente parámetro
    
DaParamBia_Err:
    Exit Function
End Function

Private Function IniGetSection(ByVal lpFileName As String, ByVal lpAppName As String) As Variant
    
    Dim tContenidos As Collection
    Dim nSize       As Long
    Dim i           As Long
    Dim j           As Long
    Dim sTmp        As String
    Dim sClave      As String
    Dim sValor      As String
    
    ' El tamaño máximo para Windows 95
    sBuffer = String$(32767, Chr$(0))
    
    nSize = GetPrivateProfileSection(lpAppName, sBuffer, Len(sBuffer), lpFileName)
        
    If nSize Then
        Set tContenidos = New Collection
        
        ' Cortar la cadena al número de caracteres devueltos
        sBuffer = Left$(sBuffer, nSize)
        ' Quitar los vbNullChar extras del final
        i = InStr(sBuffer, vbNullChar & vbNullChar)
        If i Then
            sBuffer = Left$(sBuffer, i - 1)
        End If
        
        ' Cada una de las entradas estará separada por un Chr$(0)
        Do
            i = InStr(sBuffer, Chr$(0))
            If i Then
                sTmp = LTrim$(Left$(sBuffer, i - 1))
                If Len(sTmp) Then
                    ' Comprobar si tiene el signo igual
                    j = InStr(sTmp, "=")
                    If j Then
                        sClave = Left$(sTmp, j - 1)
                        sValor = LTrim$(Mid$(sTmp, j + 1))
                        ' Asignar la clave y el valor
                        tContenidos.Add sValor
                    End If
                End If
                sBuffer = Mid$(sBuffer, i + 1)
            End If
        Loop While i
        ' Por si aún queda algo...
        If Len(sBuffer) Then
            j = InStr(sBuffer, "=")
            If j Then
                sClave = Left$(sBuffer, j - 1)
                sValor = LTrim$(Mid$(sBuffer, j + 1))
                tContenidos.Add sClave
                tContenidos.Add sValor
            End If
        End If
    End If
    Set IniGetSection = tContenidos
End Function

Private Function EjecSPFuncion(ByVal psQuery As String, ByRef psDatos As String) As Boolean
   
    Dim lsDatos     As String
    Dim lnCont      As Integer
    
    On Error GoTo EjecSPFuncion_Err:
    'Inicializa la función a falso
    EjecSPFuncion = False
    
    rssRegistro.Open psQuery
    'No hubo error
    If rssRegistro.State = 1 Then
        If rssRegistro.EOF Then
            'Obtiene error
            Escribe "Error al ejecutar query: " + psQuery
            Exit Function
        Else
            'Inicializa datos de operación query
            lsDatos = ""
            
            'Prepara cadena de datos
            While Not rssRegistro.EOF  'dbError
                lnCont = 0
                    lsDatos = rssRegistro(0).Value & rssRegistro(1).Value & rssRegistro(2).Value
                rssRegistro.MoveNext
            Wend
        End If
    End If
    If rssRegistro.State = 1 Then
        'Termina el query
        rssRegistro.Close
    End If
    
    
    'Regresa verdadera la función y valores finales
    EjecSPFuncion = True
    psDatos = lsDatos
    
    Exit Function
    
EjecSPFuncion_Err:
    Escribe "Error al ejecutar funcion EjecSPFuncion: " & CStr(Err.Number) + " " + Err.Description + " " + psQuery
    rssRegistro.Close
    EjecSPFuncion = False
End Function

Public Sub Desencriptar()
'***********************************************************************************************************************
'Objetivo:      Los datos encriptados de conexión a SQL Server son desencriptados en este procedimiento
'Autor:         EDS-BAGO
'Fecha:         19/Abr/2006
'***********************************************************************************************************************
Dim LsAuxiliar As String        'Cadena auxiliar para trabajar con las variables de conexión

    'Nombre de la base de datos de TICKET
    gsNameDB = ParamEncripcion.VerificaClaves(2, gsNameDB)
    
    'Nombre de la base de datos de CATALOGOS
    gsCataDB = ParamEncripcion.VerificaClaves(2, gsCataDB)
    
    'Nombre del usuario de conexión a la BD
    gsUserDB = ParamEncripcion.VerificaClaves(2, gsUserDB)
    
    'Password de conexion
    gsPswdDB = ParamEncripcion.VerificaClaves(2, gsPswdDB)
    
    'Nombre del DSN
    gsDSNDB = ParamEncripcion.VerificaClaves(2, gsDSNDB)
        
End Sub

'Función para guardar un valor de un archivo ini
Public Sub SaveProfile(lpFileName As String, lpAppName As String, lpKeyName As String, lpString As String)
    Dim Ret As Long
    Ret = WritePrivateProfileString(lpAppName, lpKeyName, lpString, lpFileName)
End Sub


Private Function Desconectar()
    
    cnnConexion.Close
    
    Set rssRegistro = Nothing
    Set cnnConexion = Nothing
    
End Function
