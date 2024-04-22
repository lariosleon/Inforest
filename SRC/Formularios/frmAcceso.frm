VERSION 5.00
Begin VB.Form frmAcceso 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   0  'None
   Caption         =   "Acceso al Sistema"
   ClientHeight    =   9000
   ClientLeft      =   3435
   ClientTop       =   2820
   ClientWidth     =   12060
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   9.75
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmAcceso.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   9000
   ScaleWidth      =   12060
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame3 
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   0  'None
      Height          =   9015
      Left            =   11880
      TabIndex        =   15
      Top             =   0
      Width           =   255
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   0  'None
      Height          =   255
      Left            =   4680
      TabIndex        =   14
      Top             =   0
      Width           =   7335
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   0  'None
      Height          =   255
      Left            =   4680
      TabIndex        =   13
      Top             =   8760
      Width           =   7335
   End
   Begin VB.CommandButton cmdOpcion 
      Caption         =   "Aceptar"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   675
      Index           =   0
      Left            =   12120
      Picture         =   "frmAcceso.frx":0442
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   8760
      Width           =   1275
   End
   Begin VB.TextBox txtPassword 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   465
      IMEMode         =   3  'DISABLE
      Left            =   7200
      MaxLength       =   20
      PasswordChar    =   "*"
      TabIndex        =   1
      Top             =   5000
      Width           =   2940
   End
   Begin VB.TextBox txtUsuario 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   480
      Left            =   7200
      TabIndex        =   0
      Top             =   3900
      Width           =   2940
   End
   Begin VB.CommandButton cmdOpcion 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   800
      Index           =   4
      Left            =   5880
      Picture         =   "frmAcceso.frx":0544
      Style           =   1  'Graphical
      TabIndex        =   10
      Top             =   9960
      Width           =   1080
   End
   Begin VB.CommandButton cmdOpcion 
      Caption         =   "PassWord"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   555
      Index           =   3
      Left            =   8640
      TabIndex        =   9
      Top             =   9600
      Width           =   1275
   End
   Begin VB.CommandButton cmdOpcion 
      Caption         =   "Usuario"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   555
      Index           =   2
      Left            =   7320
      TabIndex        =   8
      Top             =   9720
      Width           =   1275
   End
   Begin VB.CommandButton cmdOpcion 
      Caption         =   "Cancelar"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   675
      Index           =   1
      Left            =   9840
      Picture         =   "frmAcceso.frx":1286
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   10560
      Width           =   1275
   End
   Begin VB.TextBox txtCaja1 
      BackColor       =   &H00E0E0E0&
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   10080
      TabIndex        =   2
      Top             =   9720
      Width           =   1905
   End
   Begin VB.Image imgOpcion 
      Height          =   375
      Index           =   2
      Left            =   6000
      Top             =   3960
      Width           =   975
   End
   Begin VB.Image ImagePais 
      Height          =   800
      Left            =   11020
      Stretch         =   -1  'True
      Top             =   280
      Width           =   800
   End
   Begin VB.Image imgNewOpcion 
      Height          =   495
      Index           =   0
      Left            =   7200
      Top             =   5760
      Width           =   1935
   End
   Begin VB.Image Image1 
      Height          =   855
      Left            =   10920
      Top             =   7920
      Width           =   855
   End
   Begin VB.Label txtCaja 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Label1"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   0
      TabIndex        =   12
      Top             =   8040
      Width           =   4575
   End
   Begin VB.Image imgOpcion 
      Height          =   375
      Index           =   3
      Left            =   6000
      Top             =   5040
      Width           =   975
   End
   Begin VB.Image imgOpcion 
      Height          =   855
      Index           =   1
      Left            =   0
      Top             =   0
      Width           =   975
   End
   Begin VB.Image imgOpcion 
      Height          =   1335
      Index           =   4
      Left            =   7680
      Top             =   2040
      Width           =   975
   End
   Begin VB.Label txtBD 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Label1"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   0
      TabIndex        =   11
      Top             =   8520
      Width           =   4575
   End
   Begin VB.Label Label 
      AutoSize        =   -1  'True
      Caption         =   "Caja :"
      Height          =   240
      Index           =   0
      Left            =   2400
      TabIndex        =   5
      Top             =   9360
      Width           =   510
   End
   Begin VB.Label Label 
      AutoSize        =   -1  'True
      Caption         =   "PassWord :"
      Height          =   240
      Index           =   2
      Left            =   2160
      TabIndex        =   4
      Top             =   9840
      Width           =   1050
   End
   Begin VB.Label Label 
      AutoSize        =   -1  'True
      Caption         =   "Usuario :"
      Height          =   240
      Index           =   1
      Left            =   2280
      TabIndex        =   3
      Top             =   9600
      Width           =   795
   End
   Begin VB.Image Image 
      Height          =   1185
      Left            =   3960
      Picture         =   "frmAcceso.frx":1388
      Stretch         =   -1  'True
      Top             =   9360
      Width           =   1170
   End
   Begin VB.Image imgInforest2 
      Height          =   9000
      Left            =   0
      Picture         =   "frmAcceso.frx":17CA
      Top             =   9120
      Visible         =   0   'False
      Width           =   12000
   End
   Begin VB.Image ImgAdministracion2 
      Height          =   9000
      Left            =   1200
      Picture         =   "frmAcceso.frx":187BD
      Top             =   9000
      Visible         =   0   'False
      Width           =   12000
   End
   Begin VB.Image imgConsulta2 
      Height          =   9000
      Left            =   12120
      Picture         =   "frmAcceso.frx":2DCBB
      Top             =   7680
      Visible         =   0   'False
      Width           =   12000
   End
   Begin VB.Image imgInforest 
      Height          =   9000
      Left            =   0
      Picture         =   "frmAcceso.frx":4265B
      Top             =   0
      Visible         =   0   'False
      Width           =   12000
   End
   Begin VB.Image imgAdministracion 
      Height          =   9000
      Left            =   0
      Picture         =   "frmAcceso.frx":59134
      Top             =   0
      Visible         =   0   'False
      Width           =   12000
   End
   Begin VB.Image imgconsulta 
      Height          =   9000
      Left            =   0
      Picture         =   "frmAcceso.frx":743F8
      Top             =   0
      Visible         =   0   'False
      Width           =   12000
   End
End
Attribute VB_Name = "frmAcceso"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim RsUsuario As Recordset
Dim RsAcceso As Recordset
Dim i As Integer
Dim validaCierreHK As Boolean
Dim fso As Object
'HUELLA
Dim lModulo As String

Private Sub cmdOpcion_Click(Index As Integer)
    Select Case Index
           Case Is = 0 ' Aceptar
                If lHARDkey Then
                    ValidacionEntrarConLicencia
                End If
                 
                'Chequea Datos
                
                wEnter = False
                
                If txtUsuario.Text = "" Then MsgBox "Ingrese su usuario", vbExclamation, sMensaje: txtUsuario.SetFocus: Exit Sub
                If txtPassword.Text = "" Then MsgBox "Ingrese su password", vbExclamation, sMensaje: txtPassword.SetFocus: Exit Sub
                RsUsuario.MoveFirst
                RsUsuario.Find ("tResumido = '" & Trim(txtUsuario.Text) & "' ")
                   
                If RsUsuario.EOF Then
                    i = i + 1
                    MsgBox "Usuario No Encontrado", vbCritical, sMensaje
                    txtPassword.Text = ""
                    txtUsuario.SetFocus
                Else
                    If Desencapsula(RsUsuario!tpassword) = UCase(txtPassword.Text) Or (Desencapsula(RsUsuario!tBandaMagnetica) = UCase(Extrae(txtPassword.Text)) And RsUsuario!tBandaMagnetica <> "") Then
                        sPassword = UCase(txtPassword.Text)
                        sUsuario = UCase(txtUsuario.Text)
                        xUsuario = Mid(RsUsuario!tCodigoUsuario, 3, 3)
                        tcodigoUsuarioA = RsUsuario!tCodigoUsuario 'para controler

                        wEnter = True
                       
                        Open App.Path & "\USUARIO.INI" For Output As #1
                        Print #1, IIf(Mid(sUsuario, 1, 1) = "*", Mid(sUsuario, 2, 15), sUsuario)
                        Close #1
                        If lHARDkey Then
                            validaCierreHK = False
                        End If
             
                        'audirotia
                        
                        registroAccesoAuditoria "I", sUsuario
                        If nCorrelativoAcceso = -1 Then
                        End
                        End If
                        'auditoria
                        
                        Unload Me
                    Else
                        i = i + 1
                        MsgBox "Password Erroneo", vbCritical, sMensaje
                        txtPassword.Text = ""
                        txtPassword.SetFocus
                    End If
                End If
                If i = 4 Then End
           
           Case Is = 1 ' Cancelar
                If lHARDkey Then
                    '----------Verifica Llave HK----------------------------------
                    If hk.ValidaLlave Then
                        'MsgBox "Fallo la validacion de la llave", vbCritical, "Aviso"
                        Dim result As Boolean
                        Select Case sModulo
                            Case "INFOREST"
                                result = hk.FinalizarConexion(Aplicacion.PuntoVenta) 'InfhotelHK.PuntoVenta)
                            Case "ADMINISTRACION"
                                result = hk.FinalizarConexion(Aplicacion.Administracion) 'InfhotelHK.Administracion)
                            Case "CONSULTA"
                                result = hk.FinalizarConexion(Aplicacion.Consultas) 'InfhotelHK.Consultas)
                            Case Else
                        End Select
                        End
                    End If
                '--------------------------------------------------------------
                End If
                End
                
           Case Is = 2 ' Usuario
                frmKeyBoard.txtResultado.Text = txtUsuario.Text
                frmKeyBoard.Show vbModal
                If wEnter Then
                   txtUsuario.Text = sDescrip
                End If
                wEnter = False
                
           Case Is = 3 ' Password
                frmPassword.cmdOpcion.Visible = False
                frmPassword.Show vbModal
                If wEnter Then
                   txtPassword.Text = sDescrip
                End If
                wEnter = False
                
           Case Is = 4 'HUELLA
                wenterHuellaSup = False
                lUsuarioHuella = True
                frmVerificacionHuellaSup.Opcion lModulo
                frmVerificacionHuellaSup.Show vbModal
                If wenterHuellaSup Then
                    wEnter = True
                    lUsuarioHuella = False
                    sUsuario = sVar1
                    Unload Me
                End If
                
                
    End Select
End Sub

Private Sub Form_Activate()
   If sUsuario = "" Then
      txtUsuario.SetFocus
   Else
      txtPassword.SetFocus
   End If
   lUsuarioHuella = False
   
'    'Actualización automática
'    Dim sVersion As String
'    Dim sVersionExe As String
'    Dim RsVersion As Recordset
'
'    sVersion = ""
'    sVersionExe = App.Major & "." & App.Minor & "." & App.Revision
'    sVersion = Calcular("SELECT tVersion As Codigo FROM tParametro", Cn)
'
'    If sVersion <> sVersionExe Then
'        MsgBox "Existe Una Nueva Version Disponible", vbInformation, sMensaje
'
'        'CREA LA CARPETA PARA BACKUP DE EXES
'        Dim Backup As String
'        Backup = App.Path + "\ExesHistoricos"
'
'        'CREA LA CARPETA Y VALIDA QUE NO EXISTA
'        If Dir(Backup, vbDirectory) = "" Then
'           MkDir (Backup)
'        End If
'
'        Shell App.Path & "\Actualizador.exe" & " " & App.EXEName + "1", vbNormalFocus
'        End
'    End If
   
End Sub

Private Sub Form_Load()
  On Error GoTo err:
  
  '************** del flash
'  If lAlmacen Then
'      Tiempo.Interval = 1
'   Else
'      Tiempo.Interval = 1500
'   End If
    If lAlmacen Then
      Actualiza
   End If
   If sModulo = "INTEGRACION" Then
      Integra
   End If
   
   If sModulo = "ADMINISTRACION" Then
    Me.imgAdministracion.Visible = True
   End If
   If sModulo = "INFOREST" Then
    Me.imgInforest.Visible = True
   End If
   If sModulo = "CONSULTA" Then
    Me.imgconsulta.Visible = True
   End If
   'TIPO CAMBIO
   If pais = "002" Then
      RTipoCambio
   End If
  '*************************
  Set fso = CreateObject("Scripting.FileSystemObject")
      Select Case pais
        Case Is = "001"
            If fso.FileExists(App.Path & "\bmps\Paises\001.jpg") Then
               ImagePais.Picture = LoadPicture(App.Path & "\bmps\Paises\001.jpg")
            End If
        Case Is = "002"
            If fso.FileExists(App.Path & "\bmps\Paises\002.jpg") Then
               ImagePais.Picture = LoadPicture(App.Path & "\bmps\Paises\002.jpg")
            End If
        Case Else
            If fso.FileExists(App.Path & "\bmps\Paises\000.jpg") Then
               ImagePais.Picture = LoadPicture(App.Path & "\bmps\Paises\000.jpg")
            End If
    End Select

    Set fso = Nothing
  
  If lHARDkey Then
        validaCierreHK = True
  End If
  Open App.Path & "\USUARIO.INI" For Input As #1   ' Abre el archivo para recibir los datos.
  Do While Not EOF(1)                              ' Repite el bucle hasta el final del archivo.
     Input #1, sUsuario                            ' Lee el carácter en dos variables
  Loop
  Close #1
  txtUsuario.Text = sUsuario

  AccesoInicio
  Screen.MousePointer = vbDefault
Exit Sub

err:
   txtUsuario.Text = ""
   AccesoInicio
   Screen.MousePointer = vbDefault
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    If lHARDkey Then
        If validaCierreHK Then
            Dim Verifica As Boolean
            Select Case sModulo
                Case "INFOREST"
                    Verifica = hk.FinalizarConexion(Aplicacion.PuntoVenta)
                Case "ADMINISTRACION"
                    Verifica = hk.FinalizarConexion(Aplicacion.Administracion)
                Case "CONSULTA"
                    Verifica = hk.FinalizarConexion(Aplicacion.Consultas)
                Case Else
            End Select
        End If
    End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set RsUsuario = Nothing
    Set frmAcceso = Nothing
End Sub

Private Sub Image1_Click()
    frmAbout.Show vbModal
End Sub

Private Sub imgNewOpcion_Click(Index As Integer)
    MsgBox "Proceso en desarrollo!! Aun no Culminado!", vbInformation, sMensaje
End Sub

Private Sub imgOpcion_Click(Index As Integer)
 Select Case Index
           Case Is = 0 ' Aceptar
                If lHARDkey Then
                    ValidacionEntrarConLicencia
                End If
                 
                'Chequea Datos
                
                wEnter = False
                
                If txtUsuario.Text = "" Then MsgBox "Ingrese su usuario", vbExclamation, sMensaje: txtUsuario.SetFocus: Exit Sub
                If txtPassword.Text = "" Then MsgBox "Ingrese su password", vbExclamation, sMensaje: txtPassword.SetFocus: Exit Sub
                RsUsuario.MoveFirst
                RsUsuario.Find ("tResumido = '" & Trim(txtUsuario.Text) & "' ")
                   
                If RsUsuario.EOF Then
                    i = i + 1
                    MsgBox "Usuario No Encontrado", vbCritical, sMensaje
                    txtPassword.Text = ""
                    txtUsuario.SetFocus
                Else
                    If Desencapsula(RsUsuario!tpassword) = UCase(txtPassword.Text) Or (Desencapsula(RsUsuario!tBandaMagnetica) = UCase(Extrae(txtPassword.Text)) And RsUsuario!tBandaMagnetica <> "") Then
                        sPassword = UCase(txtPassword.Text)
                        sUsuario = UCase(txtUsuario.Text)
                        xUsuario = Mid(RsUsuario!tCodigoUsuario, 3, 3)
                        tcodigoUsuarioA = RsUsuario!tCodigoUsuario 'para controler

                        wEnter = True
                       
                        Open App.Path & "\USUARIO.INI" For Output As #1
                        Print #1, IIf(Mid(sUsuario, 1, 1) = "*", Mid(sUsuario, 2, 15), sUsuario)
                        Close #1
                        If lHARDkey Then
                            validaCierreHK = False
                        End If
             
                        'audirotia
                        
                        registroAccesoAuditoria "I", sUsuario
                        If nCorrelativoAcceso = -1 Then
                        End
                        End If
                        'auditoria
                        
                        Unload Me
                    Else
                        i = i + 1
                        MsgBox "Password Erroneo", vbCritical, sMensaje
                        txtPassword.Text = ""
                        txtPassword.SetFocus
                    End If
                End If
                If i = 4 Then End
           
           Case Is = 1 ' Cancelar
                If lHARDkey Then
                    '----------Verifica Llave HK----------------------------------
                    If hk.ValidaLlave Then
                        'MsgBox "Fallo la validacion de la llave", vbCritical, "Aviso"
                        Dim result As Boolean
                        Select Case sModulo
                            Case "INFOREST"
                                result = hk.FinalizarConexion(Aplicacion.PuntoVenta) 'InfhotelHK.PuntoVenta)
                            Case "ADMINISTRACION"
                                result = hk.FinalizarConexion(Aplicacion.Administracion) 'InfhotelHK.Administracion)
                            Case "CONSULTA"
                                result = hk.FinalizarConexion(Aplicacion.Consultas) 'InfhotelHK.Consultas)
                            Case Else
                        End Select
                        End
                    End If
                '--------------------------------------------------------------
                End If
                End
                
           Case Is = 2 ' Usuario
                frmKeyBoard.txtResultado.Text = txtUsuario.Text
                frmKeyBoard.Show vbModal
                If wEnter Then
                   txtUsuario.Text = sDescrip
                End If
                wEnter = False
                
           Case Is = 3 ' Password
                frmPassword.cmdOpcion.Visible = False
                frmPassword.Show vbModal
                If wEnter Then
                   'txtPassword.SetFocus
                   txtPassword.Text = sDescrip
                   'cmdOpcion_Click (0)
                End If
                wEnter = False
                
           Case Is = 4 'HUELLA
                wenterHuellaSup = False
                lUsuarioHuella = True
                frmVerificacionHuellaSup.Opcion lModulo
                frmVerificacionHuellaSup.Show vbModal
                If wenterHuellaSup Then
                    wEnter = True
                    lUsuarioHuella = False
                    sUsuario = sVar1
                    Unload Me
                End If
                
                
    End Select
End Sub

Private Sub txtPassword_GotFocus()
   If Trim(txtPassword.Text) <> "" Then
      imgOpcion_Click (0)
   End If
End Sub

Private Sub txtPassword_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      imgOpcion_Click (0)
   End If
End Sub

Private Sub txtUsuario_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      imgOpcion_Click (0)
   End If
End Sub

Private Sub txtUsuario_LostFocus()
   txtUsuario.Text = UCase(txtUsuario)
End Sub

Public Sub AccesoInicio()
    wEnter = False
    txtCaja.Caption = "CAJA  " & sCaja
    txtBD.Caption = UCase(sRuta) & " : " & UCase(sMDB)
    i = 1
    
    'HUELLA
    pTipo = "M"
    
    If sModulo = "INFOREST" Then
      ' Set RsUsuario = Lib.OpenRecordset("select * from vGrupoUSUARIO where lActivo=1 and ActivoGrupo=1 and lModulo01=1", Cn)
       lModulo = "01"
    ElseIf sModulo = "ADMINISTRACION" Then
       'Set RsUsuario = Lib.OpenRecordset("select * from vGrupoUSUARIO where lActivo=1 and ActivoGrupo=1 and lModulo02=1", Cn)
       lModulo = "02"
    Else
      ' Set RsUsuario = Lib.OpenRecordset("select * from vGrupoUSUARIO where lActivo=1 and ActivoGrupo=1 and lModulo03=1", Cn)
       lModulo = "03"
    End If
    Set RsUsuario = Lib.OpenRecordset("usp_Inforest_ObtieneUsuarios '" & sModulo & "'", Cn)
    If RsUsuario.RecordCount = 0 Then
       MsgBox "No existen Usuarios..!", vbCritical, sMensaje
       End
    End If
End Sub

Private Sub ValidacionEntrarConLicencia()
    If lHARDkey Then
        '----------Verifica Llave HK----------------------------------
        Dim verif As Boolean
        verif = hk.VerificaConexion
                        
        If verif = False Then
            Dim str As String
            str = ""
            Select Case sModulo
                Case "INFOREST"
                    str = hk.IniciaConexion(Aplicacion.PuntoVenta)
                Case "ADMINISTRACION"
                    str = hk.IniciaConexion(Aplicacion.Administracion)
                Case "CONSULTA"
                    str = hk.IniciaConexion(Aplicacion.Consultas)
                Case Else
                    
            End Select
            If str <> "" Then
                MsgBox str, vbCritical, "Aviso"
                End
            End If
        End If
        '--------------------------------------------------------------
    End If
End Sub
Public Sub Actualiza()
   Dim RsTemp As Recordset
   Screen.MousePointer = vbHourglass
   CnAlmacen.Execute "sp_ActualizaReceta"
   
    Cn.Execute "usp_Inforest_InicializaCostos"
    
    Dim oComandox As clsComando
    Set oComandox = New clsComando
    If Not oComandox.CreateCmdSp("usp_Inforest_ActualizaCostos", Cn) Then
       Set oComandox = Nothing
       Exit Sub
    End If
    oComandox.CreateParameter "@tNombreInforest", adVarChar, adParamInput, 50, sMDB
    oComandox.CreateParameter "@tNombreAlmacen", adVarChar, adParamInput, 50, sAlmacenMDB
    oComandox.CreateParameter "@tLocal", adVarChar, adParamInput, 5, sLocal
    If Not oComandox.GetParamOK Then
       Set oComandox = Nothing
       Exit Sub
    End If
    If Not oComandox.ExecSP Then
    Set oComandox = Nothing
    Exit Sub
    End If

'Actualiza los precios de Venta de Transferencia a almacen

If Not oComandox.CreateCmdSp("Usp_ActualizarPreciosTransferenciaAlmacen", Cn) Then
   Set oComandox = Nothing
   Exit Sub
End If
oComandox.CreateParameter "@SubGrupo", adVarChar, adParamInput, 50, ""
oComandox.CreateParameter "@BaseDatoAlmacen", adVarChar, adParamInput, 50, sAlmacenMDB
oComandox.CreateParameter "@tipooper", adInteger, adParamInput, 5, 2
If Not oComandox.GetParamOK Then
   Set oComandox = Nothing
   Exit Sub
End If
If Not oComandox.ExecSP Then
    Set oComandox = Nothing
    Exit Sub
End If

   
   Screen.MousePointer = vbDefault
   Exit Sub
End Sub
Public Sub Integra()
    'frmServidores.cargaModo False
    'frmServidores.llenaGrid
End Sub

Public Sub RTipoCambio()
   'TIPO DE CAMBIO
   Dim rsTipoCambio As Recordset
   
   Isql = "select * from TTIPOCAMBIO WHERE CONVERT(NVARCHAR,fFecha,103)= '" & FechaServidorTipoCambio() & "' "
   Set rsTipoCambio = Lib.OpenRecordset(Isql, Cn)
     
   If rsTipoCambio.RecordCount = 0 Then
                     
                  Dim oComando As clsComando
                  Set oComando = New clsComando
                  If Not oComando.CreateCmdSp("spIns_TipoCambio", Cn) Then
                     Set oComando = Nothing
                     Exit Sub
                  End If

                  oComando.CreateParameter "@nTc", adDouble, adParamInput, 0, 1
                  oComando.CreateParameter "@tUSUARIO", adVarChar, adParamInput, 15, ""
                  oComando.CreateParameter "@nTco", adDouble, adParamInput, 0, 0
                  If Not oComando.GetParamOK Then
                     Set oComando = Nothing
                     Exit Sub
                  End If
                  If Not oComando.ExecSP Then
                     Set oComando = Nothing
                     Exit Sub
                  End If
   End If
End Sub

