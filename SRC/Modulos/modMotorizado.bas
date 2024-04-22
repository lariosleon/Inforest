Attribute VB_Name = "modMotorizado"
 Sub Main()
    ' Configuracion ini
    sRuta = Trim(LeerIni(App.Path + "\INFOREST.INI", "Conexion", "SERVIDOR", "."))
    sMDB = Trim(LeerIni(App.Path + "\INFOREST.INI", "Conexion", "BASEDATO", "INFOREST"))
    sCaja = Trim(LeerIni(App.Path + "\INFOREST.INI", "Configuracion", "CAJA", "001"))
    sSalon = Trim(LeerIni(App.Path + "\INFOREST.INI", "Configuracion", "SALON", "01"))
    sEmpresa = Trim(LeerIni(App.Path + "\INFOREST.INI", "Configuracion", "EMPRESA", "000"))
    sRutaCD = Trim(LeerIni(App.Path + "\INFOREST.INI", "CentralDelivery", "SERVIDOR", "."))
    sMDBCD = Trim(LeerIni(App.Path + "\INFOREST.INI", "CentralDelivery", "BASEDATO", "INFOREST"))
        
    Dim RsParametro As Recordset
    
    Set Lib = New Libreria16.Applications
    sUserName = "Infhotel"
    sUserPassword = "4gust1n-fl0r14n"
    
    
    Set Cn = New Connection
    Cn.Provider = "SQLOLEDB"
    Cn.CursorLocation = adUseServer
    Cn.ConnectionString = "User ID=" & sUserName & _
                          ";password=" & sUserPassword & _
                          ";Data Source=" & sRuta & _
                          ";Initial Catalog=" & sMDB
   Cn.CommandTimeout = 300
   Cn.Open
   CD = Calcular("select lCD as Codigo from TCAJA where tCaja='" & sCaja & "'", Cn)
 
    Isql = "select * from TPARAMETRO"
   Set RsParametro = Lib.OpenRecordset(Isql, Cn)
      sRUC = IIf(IsNull(RsParametro!tIdentificacionTributaria), "", RsParametro!tIdentificacionTributaria)
     nLongitud = IIf(IsNull(RsParametro!nLongitud), 11, RsParametro!nLongitud)
     
     
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
     
   frmLlegadaSalida.Show

End Sub
