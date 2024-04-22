Attribute VB_Name = "modConsulta"
Public hk As New License
 
 Sub Main()
 Dim existe As Boolean
 Dim AnoMesSv As String
 Dim directorioSistema   As String
 Dim archI As String
    lVersionEducativa = False
    
    lHARDkey = False
    
    moduloUso = "Consulta"

    ultimoConectado = True
    
    ' Configuracion ini
    sRuta = Trim(LeerIni(App.Path + "\INFOREST.INI", "Conexion", "SERVIDOR", "."))
    sMDB = Trim(LeerIni(App.Path + "\INFOREST.INI", "Conexion", "BASEDATO", "INFOREST"))
    sCaja = Trim(LeerIni(App.Path + "\INFOREST.INI", "Configuracion", "CAJA", "001"))
    sSalon = Trim(LeerIni(App.Path + "\INFOREST.INI", "Configuracion", "SALON", "01"))
    sEmpresa = Trim(LeerIni(App.Path + "\INFOREST.INI", "Configuracion", "EMPRESA", "000"))
    
    
    'FACTURACION_E_PERU
    sRutaFE = Trim(LeerIni(App.Path + "\FACTURACION.INI", "Conexion", "SERVIDOR", "."))
    sMDBFE = Trim(LeerIni(App.Path + "\FACTURACION.INI", "Conexion", "BASEDATO", "BDEFACT"))
    '-----------------
   
    
    Dim RsParametro As Recordset
    Dim RsTc As Recordset
    Dim RsCaja As Recordset
    Dim RsBusca As Recordset
    Dim RsFactura As Recordset
    
    Set Lib = New Libreria16.Applications
    sUserName = "infhotel"
    sUserPassword = "4gust1n-fl0r14n"
    
    'auditoria
    nCorrelativoAcceso = 0
    tModuloSeg = "14" 'XX= CODIGO DEL MODULO DE LA TABLA DE MMODULO DE SEGURIDAD
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
            MsgBox "Error Fatal: Su Licencia Acad�mica a Caducado. Comun�quese con Personal De INFHOTEL", vbCritical, sMensaje
            End
        End If
     Else
        MsgBox "Error Fatal: Comun�quese con Personal De INFHOTEL", vbCritical, sMensaje
        End
     End If
    End If
    
    
    '------VERIFICACION DE VERSION Actualizaci�n autom�tica
        
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

    
    
    
    '=================================================
    pais = ObtienePais
    '=================================================
    '===================================================

    'Configuraci�n
   'Provider=SQLOLEDB.1;Persist Security Info=False;User ID=sa;Initial Catalog=Inforest;Data Source=LUIS
   'Provider=Microsoft.Jet.OLEDB.4.0;Data Source=C:\Inforest\Data\inforest.mdb;Persist Security Info=False
   
   Isql = "select * from TPARAMETRO"
   Set RsParametro = Lib.OpenRecordset(Isql, Cn)
   sModulo = "CONSULTA"
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
   sPie = IIf(IsNull(RsParametro!tPie), "", RsParametro!tPie)
   sPiePreCuenta = IIf(IsNull(RsParametro!tPiePreCuenta), "", RsParametro!tPiePreCuenta)
   sRazonComercial = IIf(IsNull(RsParametro!tRazonComercial), "", RsParametro!tRazonComercial)
   sDireccion = IIf(IsNull(RsParametro!tDireccion), "", RsParametro!tDireccion)
   sDireccion2 = IIf(IsNull(RsParametro!tDireccion2), "", RsParametro!tDireccion2)
   sTelefono = IIf(IsNull(RsParametro!tTelefono), "", RsParametro!tTelefono)
   sFax = IIf(IsNull(RsParametro!tFax), "", RsParametro!tFax)
   sWeb = IIf(IsNull(RsParametro!tWebPage), "", RsParametro!tWebPage)
   sMail = IIf(IsNull(RsParametro!tEmail), "", RsParametro!tEmail)
   nCabecera = IIf(IsNull(RsParametro!nCabecera), 0, RsParametro!nCabecera)
   nDetalle = IIf(IsNull(RsParametro!nDetalle), 0, RsParametro!nDetalle)
   sRUC = IIf(IsNull(RsParametro!tIdentificacionTributaria), "", RsParametro!tIdentificacionTributaria)
   sMonN = IIf(IsNull(RsParametro!tMonN), "", RsParametro!tMonN)
   sMonedaN = IIf(IsNull(RsParametro!tMonedaN), "", RsParametro!tMonedaN)
   sMonE = IIf(IsNull(RsParametro!tMonE), "", RsParametro!tMonE)
   sMonedaE = IIf(IsNull(RsParametro!tMonedaE), "", RsParametro!tMonedaE)
   sPAdmin = UCase(Desencapsula(IIf(IsNull(RsParametro!tpassword), "", RsParametro!tpassword)))
   sElimina = IIf(IsNull(RsParametro!tElimina), "", RsParametro!tElimina)
   lInfhotel = IIf(IsNull(RsParametro!lInfhotel), False, RsParametro!lInfhotel)
   sClub = IIf(IsNull(RsParametro!tClub), "", RsParametro!tClub)
   lCierre = IIf(IsNull(RsParametro!lCierre), False, RsParametro!lCierre)
   
       'HUELLA
     lHuellaDigitalPersona = IIf(IsNull(RsParametro!lHUELLADIGITAL), False, RsParametro!lHUELLADIGITAL)
      lHuellaSecugen = IIf(IsNull(RsParametro!lHuellaSecugen), False, RsParametro!lHuellaSecugen)
    'KDS
   'FACTURACION ELECTRONICA
   lFacturacionE = IIf(IsNull(RsParametro!lFacturacionE), False, RsParametro!lFacturacionE)
   tCodigoFE = IIf(IsNull(RsParametro!tCodigoFE), "000", RsParametro!tCodigoFE)
   tPieDocumento1 = IIf(IsNull(RsParametro!tPieDocumento1), " ", RsParametro!tPieDocumento1)
   lAmbienteProduccion = IIf(IsNull(RsParametro!lAmbienteFE), False, RsParametro!lAmbienteFE)
   
   'agenteretencion
   tTextoAgenteRetencion = IIf(IsNull(RsParametro!tAgenteRetencion), "", RsParametro!tAgenteRetencion)
   
   lImpresionCodigoBarras = IIf(IsNull(RsParametro!lImprimeCodigoBarras), False, RsParametro!lImprimeCodigoBarras)
   
   lNcOfisis = IIf(IsNull(RsParametro!lNcOfisis), False, RsParametro!lNcOfisis)
   
   'TCANALVENTA
   sBoton1 = Calcular("select tDetallado as Codigo from TCANALVENTA where tCodigoCanalVenta='01'", Cn)
   sBoton2 = Calcular("select tDetallado as Codigo from TCANALVENTA where tCodigoCanalVenta='02'", Cn)
   sBoton3 = Calcular("select tDetallado as Codigo from TCANALVENTA where tCodigoCanalVenta='03'", Cn)
   sBoton4 = Calcular("select tDetallado as Codigo from TCANALVENTA where tCodigoCanalVenta='04'", Cn)
   sBoton5 = Calcular("select tDetallado as Codigo from TCANALVENTA where tCodigoCanalVenta='05'", Cn)
   nDias = IIf(IsNull(RsParametro!nDias), 2, RsParametro!nDias)
   nPunto = IIf(IsNull(RsParametro!nPunto), 1, RsParametro!nPunto)
   lAlmacen = IIf(IsNull(RsParametro!lAlmacen), False, RsParametro!lAlmacen)
   nLongitud = IIf(IsNull(RsParametro!nLongitud), 11, RsParametro!nLongitud)
   
   'club
   lClub = IIf(IsNull(RsParametro!lClub), False, RsParametro!lClub)
   obtieneVencimientoConexiones
   
   'fe
    lQRFE = IIf(IsNull(RsParametro!lCodigoQrFE), 0, IIf(RsParametro!lCodigoQrFE = True, 1, 0))
    RutaImgFE = IIf(IsNull(RsParametro!tRutaFE), "", RsParametro!tRutaFE)
   
   'fepaperllees
   lFEpape = IIf(IsNull(RsParametro!lFEpape), 0, IIf(RsParametro!lFEpape = True, 1, 0))
   lDesactivaNCFP = IIf(IsNull(RsParametro!lDesactivaNCFP), 0, IIf(RsParametro!lDesactivaNCFP = True, 1, 0))
   
   'FACTURACION_E_PERU
   On Error GoTo FacturacionIni
   
   If lFacturacionE And lFEOfisis = False And lFEpape = False Then
            Set CnFE = New Connection
            CnFE.Provider = "SQLOLEDB"
            CnFE.CursorLocation = adUseServer
            CnFE.ConnectionString = "User ID=" & sUserName & _
            ";password=" & sUserPassword & _
            ";Data Source=" & sRutaFE & _
            ";Initial Catalog=" & sMDBFE
            CnFE.CommandTimeout = 250
            CnFE.Open
    End If
   
FacturacionIni:
    If err.Number = "-2147467259" Then
       MsgBox "Facturacion Electronica, SQL Server no Encontrado, Falla de Conectividad", vbCritical + vbOKOnly
       Exit Sub
    End If
    '-------------------------------
    
   
   
   
   
   'KDS Para reporte KDS
   lKDS = IIf(IsNull(RsParametro!lKDS), False, RsParametro!lKDS)
   If (lKDS) And sModulo <> "CONSULTA" Then
        sBump = IIf(IsNull(RsParametro!tBump), "", RsParametro!tBump)
        KDS_ListarBumpNotification (sBump)
   End If
      
   Set CnAlmacen = New Connection
   Set CnAlmacenRemoto = New Connection
   Set cnAlmacenDefault = New Connection
   
   If lAlmacen Then
      ' Configuracion ini
      sAlmacenRuta = Trim(LeerIni(App.Path + "\ALMACEN.INI", "Conexion", "SERVIDOR", "."))
      sAlmacenMDB = Trim(LeerIni(App.Path + "\ALMACEN.INI", "Conexion", "BASEDATO", "ALMACEN"))
      sLocal = Trim(LeerIni(App.Path + "\ALMACEN.INI", "Configuracion", "LOCAL", "01"))
  
      CnAlmacen.Provider = "SQLOLEDB"
      CnAlmacen.CursorLocation = adUseServer
      CnAlmacen.ConnectionString = "User ID=" & sUserName & _
                                   ";password=" & sUserPassword & _
                                   ";Data Source=" & sAlmacenRuta & _
                                   ";Initial Catalog=" & sAlmacenMDB
      CnAlmacen.Open
      Set cnAlmacenDefault = CnAlmacen
        localConectado = Calcular("select isnull(tresumido,'0') as codigo from vlocalidades where ip='" & sRuta & "' and bdinf='" & sMDB & "'", CnAlmacen)
      
'      '''''''''''''''''''''''''' almacen remoto
'      Dim verificaAlmacenRemoto As String
'      verificaAdministracionCentralizada = Trim(LeerIni(App.Path + "\ALMACEN.INI", "AlmacenRemoto", "REMOTO", "OFF"))
'      If verificaAdministracionCentralizada = "ON" Then
'            lAlmacenRemoto = True
'            sRutaAlmacenRemoto = Trim(LeerIni(App.Path + "\ALMACEN.INI", "AlmacenRemoto", "SERVIDOR", "LOCAL"))
'            sMDBAlmacenRemoto = Trim(LeerIni(App.Path + "\ALMACEN.INI", "AlmacenRemoto", "BASEDATO", "ALMACEN"))
'      End If
      ''''''''''''''''''''''''''
   End If
    ' UN EXE VARIAS BD
    multiLocal = IIf(IsNull(RsParametro!lmultilocal), False, RsParametro!lmultilocal)
    'dic 2010 un exe varias bd

    If localConectado = "0" Or localConectado = "" Then
            localConectado = sRazonComercial
    End If
    
    Call ValidaHARDkey 'OO
    'frmFlash.Label5.Caption = "M�dulo de Consultas"
    'frmFlash.Show vbModal
                               
                             
   'Proceso de Caja
   Set RsCaja = Lib.OpenRecordset("select * from TCAJA", Cn)
   If RsCaja.RecordCount <> 0 Then
      RsCaja.MoveFirst
      RsCaja.Find ("tCaja='" & sCaja & "'")
      If RsCaja.EOF Then
         MsgBox "Error Fatal: No existe Caja Configurada", vbCritical, sMensaje
         End
      Else
         sPreCuenta = IIf(IsNull(RsCaja!tPrecuenta), "001", RsCaja!tPrecuenta)
         
         'descripcion alternativa
         lDescripcionAlternativa = IIf(IsNull(RsCaja!lActivaImpDscAlternativa), False, RsCaja!lActivaImpDscAlternativa)
        lImprimeImagCabPrecuenta = IIf(IsNull(RsCaja!lImprimeImagCabPrecuenta), False, RsCaja!lImprimeImagCabPrecuenta)
        lImprimeImagPiePrecuenta = IIf(IsNull(RsCaja!lImprimeImagPiePrecuenta), False, RsCaja!lImprimeImagPiePrecuenta)

      End If
   Else
      MsgBox "Error Fatal: No existen Cajas", vbCritical, sMensaje
      End
   End If
   
   Set RsTc = Nothing
   Set RsCaja = Nothing
   Set RsParametro = Nothing
   Set RsBusca = Nothing
   Set RsFactura = Nothing
   wInicio = False
   frmAcceso.Caption = "Inforest m�dulo de Consultas v." & App.Major & "." & App.Minor & "." & App.Revision
   frmAcceso.Show vbModal
   If wEnter = True Then
      mdiConsulta.Show
   End If
   Exit Sub
InforestIni:
    If err.Number = "-2147467259" Then
       MsgBox "Inforest, SQL Server no Encontrado, Falla de Conectividad", vbCritical + vbOKOnly
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

Private Sub ValidaHARDkey()
    On Error GoTo ErrHARDkey
    clave1 = "67332"
    clave2 = "5877"
    
    If lHARDkey Then
        '************Validacion Hard Key***************************
        Dim verif2 As Boolean
        Dim str As String
        Dim verif As String
            
        hk.SetClaves clave1, clave2
        verif = hk.IniciaConexion(Aplicacion.Consultas)
            
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
            verif2 = hk.FinalizarConexion(Aplicacion.Consultas)
            End
        End If
        '*********************************************************
    End If
ErrHARDkey:
End Sub

Private Sub FinalizaConexionHK()
    Dim result As Boolean
    result = hk.FinalizarConexion(Aplicacion.Consultas)
End Sub
