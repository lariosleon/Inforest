Attribute VB_Name = "modAdministracion"
Public hk As New License
 
 Sub Main()
 Dim existe As Boolean
  Dim AnoMesSv As String
  Dim directorioSistema  As String
  Dim archI As String
  Dim verificaAdministracionCentralizada As String
    
    ' dic 2010 un exe varias bd
    ultimoConectado = True
    moduloUso = "Administracion"
    ' dic 2010
    
    verificaAdministracionCentralizada = Trim(Trim(LeerIni(App.Path + "\INFOREST.INI", "AdministracionCentralizada", "CENTRALIZADA", "OFF")))
    If verificaAdministracionCentralizada = "ON" Then ' si es "ON" significa que jala informacion del servidor central
        lCentral = True  '  flag de adm. centralizada
        sServidorCentral = Trim(LeerIni(App.Path + "\INFOREST.INI", "AdministracionCentralizada", "SERVIDOR", "0.0.0.0")) ' leer servidor central
        bdInforestCentral = Trim(LeerIni(App.Path + "\INFOREST.INI", "AdministracionCentralizada", "BASEDATO", "INFOREST")) '  leer base de datos en servidor central, a consultar para actualizar datos
    End If
    
    '==========================================================
    lVersionEducativa = False
    
    lHARDkey = False
 
    Screen.MousePointer = vbHourglass
    ' Configuracion ini
    sRuta = Trim(LeerIni(App.Path + "\INFOREST.INI", "Conexion", "SERVIDOR", "."))
    sMDB = Trim(LeerIni(App.Path + "\INFOREST.INI", "Conexion", "BASEDATO", "INFOREST"))
    sCaja = Trim(LeerIni(App.Path + "\INFOREST.INI", "Configuracion", "CAJA", "001"))
    sSalon = Trim(LeerIni(App.Path + "\INFOREST.INI", "Configuracion", "SALON", "01"))
    sEmpresa = Trim(LeerIni(App.Path + "\INFOREST.INI", "Configuracion", "EMPRESA", "000"))
    
    Dim RsParametro As Recordset
    Dim RsCaja As Recordset
        
    Set Lib = New Libreria16.Applications
    sUserName = "infhotel"
    sUserPassword = "4gust1n-fl0r14n"
    
    
    'auditoria
    nCorrelativoAcceso = 0
    tModuloSeg = "13" 'XX= CODIGO DEL MODULO DE LA TABLA DE MMODULO DE SEGURIDAD
    Set CnSeg = New Connection
      
    CnSeg.Provider = "SQLOLEDB"
    CnSeg.CursorLocation = adUseServer
    CnSeg.ConnectionString = "User ID=" & sUserName & _
                          ";password=" & sUserPassword & _
                          ";Data Source=" & sRuta & _
                          ";Initial Catalog=INFSEGURIDAD"
    CnSeg.CommandTimeout = 250
    
    'auditoria
    
    Set Cn = New Connection
    Cn.Provider = "SQLOLEDB"
    Cn.CursorLocation = adUseServer
    Cn.ConnectionString = "User ID=" & sUserName & _
                          ";password=" & sUserPassword & _
                          ";Data Source=" & sRuta & _
                          ";Initial Catalog=" & sMDB
    Cn.CommandTimeout = 300
    Cn.Open
    '===============================================0
    
    'dic 2010 un exe varias bd
    Set cnDefault = Cn
    'dic 2010 un exe varias bd
           
    '====================================================
    If lVersionEducativa = True Then
     directorioSistema = ObtenerDirectorioSO
     directorioSistema = directorioSistema & "\Infhotel.csg"
     existe = FileExists(directorioSistema)
     If existe = True Then
        AnoMesSv = Calcular("select dbo.returnAnoMes(getdate()) as codigo", Cn)
        archI = obtieneAnoMes(directorioSistema)
        If Val(AnoMesSv) > Val(archI) Then
            MsgBox "Error Fatal: Su Licencia Académica a Caducado. Comuníquese con Personal De INFHOTEL", vbCritical, sMensaje
            End
        End If
     Else
        MsgBox "Error Fatal: Comuníquese con Personal De INFHOTEL", vbCritical, sMensaje
        End
     End If
    End If
    
    '===================================================
    'extranjeroBolivia
    '=================================================
    pais = ObtienePais
    '=================================================
    
    '------VERIFICACION DE VERSION Actualización automática
        
    Dim sVersion As String
    Dim sVersionExe As String
    Dim RsVersion As Recordset
    
    sVersion = ""
    sVersionExe = App.Major & "." & App.Minor & "." & App.Revision
    sVersion = Calcular("SELECT tVersion As Codigo FROM tParametro", Cn)
    
    If sVersion <> sVersionExe Then
        MsgBox "Existe Una Nueva Version Disponible", vbInformation, sMensaje
    
        'CREA LA CARPETA PARA BACKUP DE EXES
        Dim Backup As String
        Backup = App.Path + "\ExesHistoricos"
    
        'CREA LA CARPETA Y VALIDA QUE NO EXISTA
        If Dir(Backup, vbDirectory) = "" Then
           MkDir (Backup)
        End If
        
        Shell App.Path & "\Actualizador.exe" & " " & App.EXEName + "1", vbNormalFocus
        End
    End If
       
    '------
'   If App.PrevInstance Then
'    MsgBox "Ya se esta ejecutando el Aplicativo!", vbInformation, "Atención!!!"
'    End
'   End If

    ' ValidarLlave
    Isql = "select * from TPARAMETRO"
    Set RsParametro = Lib.OpenRecordset(Isql, Cn)
    lAlmacen = IIf(IsNull(RsParametro!lAlmacen), False, RsParametro!lAlmacen)
    Set CnAlmacen = New Connection
    Set CnAlmacenRemoto = New Connection
    Set cnAlmacenDefault = New Connection
    If lAlmacen Then
       ' Configuracion ini
       sAlmacenRuta = Trim(LeerIni(App.Path + "\ALMACEN.INI", "Conexion", "SERVIDOR", "."))
       sAlmacenMDB = Trim(LeerIni(App.Path + "\ALMACEN.INI", "Conexion", "BASEDATO", "ALMACEN"))
       sLocal = Trim(LeerIni(App.Path + "\ALMACEN.INI", "Configuracion", "LOCAL", "001"))

      
       CnAlmacen.Provider = "SQLOLEDB"
       CnAlmacen.CursorLocation = adUseServer
       CnAlmacen.ConnectionString = "User ID=" & sUserName & _
                                   ";password=" & sUserPassword & _
                                   ";Data Source=" & sAlmacenRuta & _
                                   ";Initial Catalog=" & sAlmacenMDB
       CnAlmacen.Open
       
              nLongitudAlmacen = Calcular("select ISNULL(AVG(LEN(TCODIGOLOCAL)),0) as codigo from TLOCALIDADES", CnAlmacen)
       
        If Len(sLocal) <> nLongitudAlmacen Then
        MsgBox "Error en la configuración del Local en el archivo ALMACEN.INI"
        Exit Sub
        End If
      
       
       
       Set cnAlmacenDefault = CnAlmacen
       localConectado = Calcular("select isnull(tresumido,'0') as codigo from vlocalidades where ip='" & sRuta & "' and bdinf='" & sMDB & "'", CnAlmacen)

       'almacenremoto
       verificaAlmacenRemoto = Trim(LeerIni(App.Path + "\ALMACEN.INI", "AlmacenRemoto", "REMOTO", "OFF"))
       If verificaAlmacenRemoto = "ON" Then
             lAlmacenRemoto = True
            sRutaAlmacenRemoto = Trim(LeerIni(App.Path + "\ALMACEN.INI", "AlmacenRemoto", "SERVIDOR", "LOCAL"))
            sMDBAlmacenRemoto = Trim(LeerIni(App.Path + "\ALMACEN.INI", "AlmacenRemoto", "BASEDATO", "ALMACEN"))
            
       End If
       
     End If

    Call ValidaHARDkey 'OO
    'frmFlash.Label5.Caption = "Módulo de Administración"
    'frmFlash.Show vbModal
    
    'Variables Públicas
    sModulo = "ADMINISTRACION"
    sImpuesto1 = IIf(IsNull(RsParametro!tImpuesto1), "", RsParametro!tImpuesto1)
    sImpuesto2 = IIf(IsNull(RsParametro!tImpuesto2), "", RsParametro!tImpuesto2)
    sImpuesto3 = IIf(IsNull(RsParametro!tImpuesto3), "", RsParametro!tImpuesto3)
    nPorcentaje1 = IIf(IsNull(RsParametro!IMPUESTO1), 0, RsParametro!IMPUESTO1)
    nPorcentaje2 = IIf(IsNull(RsParametro!IMPUESTO2), 0, RsParametro!IMPUESTO2)
    nPorcentaje3 = IIf(IsNull(RsParametro!IMPUESTO3), 0, RsParametro!IMPUESTO3)
    nTiempo = IIf(IsNull(RsParametro!nTiempo), 0, RsParametro!nTiempo)
    nChkTiempo = IIf(IsNull(RsParametro!nChkTiempo), 0, RsParametro!nChkTiempo)
    nDELIVERY = IIf(IsNull(RsParametro!nDELIVERY), 0, RsParametro!nDELIVERY)
    nLlevar = IIf(IsNull(RsParametro!nLlevar), 0, RsParametro!nLlevar)
    sRazonSocial = IIf(IsNull(RsParametro!tRazonSocial), "", RsParametro!tRazonSocial)
    sRazonComercial = IIf(IsNull(RsParametro!tRazonComercial), "", RsParametro!tRazonComercial)
    sDireccion = IIf(IsNull(RsParametro!tDireccion), "", RsParametro!tDireccion)
    sDireccion2 = IIf(IsNull(RsParametro!tDireccion2), "", RsParametro!tDireccion2)

    sTelefono = IIf(IsNull(RsParametro!tTelefono), "", RsParametro!tTelefono)
    sFax = IIf(IsNull(RsParametro!tFax), "", RsParametro!tFax)
    
    sWeb = IIf(IsNull(RsParametro!tWebPage), "", RsParametro!tWebPage)
    sMail = IIf(IsNull(RsParametro!temail), "", RsParametro!temail)
    sRUC = IIf(IsNull(RsParametro!tIdentificacionTributaria), "", RsParametro!tIdentificacionTributaria)
    sMonN = IIf(IsNull(RsParametro!tMonN), "", RsParametro!tMonN)
    sMonedaN = IIf(IsNull(RsParametro!tMonedaN), "", RsParametro!tMonedaN)
    sMonE = IIf(IsNull(RsParametro!tMonE), "", RsParametro!tMonE)
    sMonedaE = IIf(IsNull(RsParametro!tMonedaE), "", RsParametro!tMonedaE)
    sPAdmin = UCase(Desencapsula(IIf(IsNull(RsParametro!tpassword), "", RsParametro!tpassword)))
    lBotonTrans = IIf(IsNull(RsParametro!lBotonTrans), False, RsParametro!lBotonTrans)
    lLongitud = IIf(IsNull(RsParametro!lLongitud), False, RsParametro!lLongitud)
    nLongitud = IIf(IsNull(RsParametro!nLongitud), 11, RsParametro!nLongitud)
    lInfhotel = IIf(IsNull(RsParametro!lInfhotel), False, RsParametro!lInfhotel)
    nDecimal = IIf(IsNull(RsParametro!nDecimal), 11, RsParametro!nDecimal)
    
    
    'agenteretencion
   tTextoAgenteRetencion = IIf(IsNull(RsParametro!tAgenteRetencion), "", RsParametro!tAgenteRetencion)
   
    
    
    'KDS
    lKDS = IIf(IsNull(RsParametro!lKDS), False, RsParametro!lKDS)
    sOrderInfo = IIf(IsNull(RsParametro!tOrderInfo), "", RsParametro!tOrderInfo)
        'HUELLA
     lHuellaDigitalPersona = IIf(IsNull(RsParametro!lHUELLADIGITAL), False, RsParametro!lHUELLADIGITAL)
      lHuellaSecugen = IIf(IsNull(RsParametro!lHuellaSecugen), False, RsParametro!lHuellaSecugen)
    'KDS
    'insumoscritico
    lPrinter = IIf(IsNull(RsParametro!lPrinter), False, RsParametro!lPrinter)
    
    
    'descargo al cierre turno
    lActivaConsultaDescargo = IIf(IsNull(RsParametro!lActivaConsultaDescargo), False, RsParametro!lActivaConsultaDescargo)
    
    
    
    ' UN EXE VARIAS BD
    multiLocal = IIf(IsNull(RsParametro!lmultilocal), False, RsParametro!lmultilocal)
        
    'club
    lClub = IIf(IsNull(RsParametro!lClub), False, RsParametro!lClub)
     
    'Facturacion Electronica
    lFacturacionE = IIf(IsNull(RsParametro!lFacturacionE), False, RsParametro!lFacturacionE)
    
    'FE paperlees
    lFEpape = IIf(IsNull(RsParametro!lFEpape), False, RsParametro!lFEpape)
    
    'Canal de Venta
    sBoton1 = Calcular("select tDetallado as Codigo from TCANALVENTA where tCodigoCanalVenta='01'", Cn)
    sBoton2 = Calcular("select tDetallado as Codigo from TCANALVENTA where tCodigoCanalVenta='02'", Cn)
    sBoton3 = Calcular("select tDetallado as Codigo from TCANALVENTA where tCodigoCanalVenta='03'", Cn)
    sBoton4 = Calcular("select tDetallado as Codigo from TCANALVENTA where tCodigoCanalVenta='04'", Cn)
    sBoton5 = Calcular("select tDetallado as Codigo from TCANALVENTA where tCodigoCanalVenta='05'", Cn)
    
    '--- SAP----------
    lSAP = IIf(IsNull(RsParametro!lInteSAP), 0, IIf(RsParametro!lInteSAP = True, 1, 0))

    If lSAP Then
        sServidorSAp = IIf(IsNull(RsParametro!tservidorSAP), "", RsParametro!tservidorSAP)
        sBdSAP = IIf(IsNull(RsParametro!tBDSAP), "", RsParametro!tBDSAP)
        sCodSap = IIf(IsNull(RsParametro!tCodAlmcSAP), "", RsParametro!tCodAlmcSAP)
    Else
        sServidorSAp = "" 'IIf(IsNull(RsParametro!tservidorSAP), "", RsParametro!tservidorSAP)
        sBdSAP = ""  ' IIf(IsNull(RsParametro!tBDSAP), "", RsParametro!tBDSAP)
        sCodSap = "" 'IIf(IsNull(RsParametro!tCodAlmcSAP), "", RsParametro!tCodAlmcSAP)
    End If
   '------------------
   
    lFESpring = IIf(IsNull(RsParametro!lFESpring), 0, IIf(RsParametro!lFESpring = True, 1, 0))
    lFECarbajal = IIf(IsNull(RsParametro!lFECarbajal), 0, IIf(RsParametro!lFECarbajal = True, 1, 0))
     
     obtieneVencimientoConexiones
     
     '''''''''''''''''''
    If localConectado = "0" Or localConectado = "" Then
            localConectado = sRazonComercial
    End If
                  
    'Proceso de Caja
    Set RsCaja = Lib.OpenRecordset("select * from TCAJA", Cn)
    If RsCaja.RecordCount <> 0 Then
       RsCaja.MoveFirst
       RsCaja.Find ("tCaja='" & sCaja & "'")
       
       If RsCaja.EOF Then
          MsgBox "Error Faltal: No existe Caja Configurada", vbCritical, sMensaje
          End
       Else
         sPreCuenta = IIf(IsNull(RsCaja!tPrecuenta), "001", RsCaja!tPrecuenta)
       End If
    Else
       MsgBox "Error Faltal: No existen Cajas", vbCritical, sMensaje
       End
    End If
    
    Dim sFecha As String
    sFecha = Calcular("SELECT MAX(fFecha) AS Codigo From dbo.TTIPOCAMBIO", Cn)
    If sFecha = "0" Then
       nTC = 1
    Else
       nTC = Calcular("select nVenta as Codigo from TTIPOCAMBIO where fFecha='" & Format(sFecha, "yyyy/mm/dd") & "'", Cn)
    End If
    Set RsParametro = Nothing
    
    wInicio = False
    frmAcceso.Caption = "Inforest Módulo de Administración " & "v." & App.Major & "." & App.Minor & "." & App.Revision
    frmAcceso.Show vbModal
    If wEnter = True Then
        Set RsCaja = Lib.OpenRecordset("select * from TCAJA Where tCaja = '" + sCaja + "'", Cn)
        If RsCaja.RecordCount <> 0 Then
            lMCPV = IIf(IsNull(RsCaja!lMCPV), False, RsCaja!lMCPV)
        End If
       mdiAdministracion.Show
    End If
    Exit Sub
    
InforestIni:
    If err.Number = "-2147467259" Then
       MsgBox "Inforest, SQL Server no Encontrado, Falla de Conectividad", vbCritical, vbOKOnly
    Else
       MsgBox err.Description & ":" & err.Number
    End If
    Exit Sub
AlmacenIni:
    If err.Number = "-2147467259" Then
       MsgBox "Almacen, SQL Server no Encontrado, Falla de Conectividad", vbCritical + vbOKOnly
       Exit Sub
    Else
       MsgBox err.Description & ":" & err.Number
    End If
    Exit Sub
End Sub

Sub AsignaMenu(objNombre As Object)
    Dim Control As Control
    For Each Control In mdiAdministracion.Controls
        If Control.Name = objNombre Then
            Control.Enabled = True
            Exit Sub
        End If
    Next Control
End Sub
Public Sub EliminaTemporal()
   On Error Resume Next
   Screen.MousePointer = vbHourglass
  ' mdiAdministracion.StatusBar.Panels(1).Text = "Eleminando Temporales..."
   Dim RsTemp As Recordset
   Isql = "SELECT NAME AS Nombre FROM SYSOBJECTS WHERE TYPE='U' AND NAME LIKE 'TMP%'"
   Set RsTemp = Lib.OpenRecordset(Isql, Cn)
   
   If RsTemp.RecordCount > 0 Then
      Do While Not RsTemp.EOF
         Cn.Execute "Drop Table " & RsTemp!nombre
         RsTemp.MoveNext
      Loop
   End If
   Set RsTemp = Nothing
   Cn.Execute "DBCC SHRINKDATABASE (" & sMDB & ", 0)"
   'mdiAdministracion.StatusBar.Panels(1).Text = "Sistema Listo"
   Screen.MousePointer = vbDefault
End Sub

Private Sub ValidaHARDkey()
    On Error GoTo ErrHARDkey
    clave1 = "67332"
    clave2 = "5877"
    '************Validacion Hard Key***************************
    If lHARDkey Then
        Dim verif2 As Boolean
        Dim str As String
        Dim verif As String
            
        hk.SetClaves clave1, clave2
        verif = hk.IniciaConexion(Aplicacion.Administracion)
    
        If verif <> "" Then
            MsgBox verif, vbCritical, "Aviso"
            End
        End If
    
        str = hk.ValidaFechaVencimiento
    
        If str = "5" Then
            MsgBox "No se encontro la llave", vbCritical, "Aviso"
            End
        End If
    
        If Len(str) > 5 Then
            MsgBox str, vbExclamation, "Aviso"
        End If
    
        If str = "4" Then
            MsgBox "Su licencia ha caducado.", vbCritical, "Aviso"
            verif2 = hk.FinalizarConexion(InfhotelHK.Administracion)
            End
        End If
    End If
    '*********************************************************
ErrHARDkey:
End Sub

Private Sub FinalizaConexionHK()
    Dim result As Boolean
    result = hk.FinalizarConexion(Aplicacion.Administracion)
End Sub
