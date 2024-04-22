VERSION 5.00
Object = "{C1A8AF28-1257-101B-8FB0-0020AF039CA3}#1.1#0"; "mci32.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "mscomctl.ocx"
Begin VB.Form frmMozoUsuario 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   0  'None
   Caption         =   "Acceso "
   ClientHeight    =   9090
   ClientLeft      =   0
   ClientTop       =   -105
   ClientWidth     =   12120
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   8.25
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ForeColor       =   &H00C0C000&
   Icon            =   "frmMozoUsuario.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "frmMozoUsuario.frx":000C
   ScaleHeight     =   9090
   ScaleWidth      =   12120
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox Text2 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      Height          =   375
      Left            =   840
      TabIndex        =   16
      Top             =   6600
      Width           =   10575
   End
   Begin VB.PictureBox imgLogoPic 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   3900
      Left            =   5400
      ScaleHeight     =   3870
      ScaleWidth      =   6495
      TabIndex        =   15
      Top             =   90
      Visible         =   0   'False
      Width           =   6520
   End
   Begin VB.CommandButton cmdOpcion 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Aceptar"
      Height          =   800
      Index           =   0
      Left            =   2640
      Picture         =   "frmMozoUsuario.frx":182E9
      Style           =   1  'Graphical
      TabIndex        =   13
      Top             =   11760
      Width           =   1170
   End
   Begin VB.TextBox txtPassword 
      Alignment       =   2  'Center
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      IMEMode         =   3  'DISABLE
      Left            =   4200
      MaxLength       =   15
      PasswordChar    =   "*"
      TabIndex        =   12
      Top             =   4350
      Width           =   4680
   End
   Begin VB.CommandButton cmdOpcion 
      BackColor       =   &H00C0C0C0&
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
      Height          =   800
      Index           =   2
      Left            =   360
      Style           =   1  'Graphical
      TabIndex        =   11
      Top             =   12000
      Width           =   1170
   End
   Begin VB.CommandButton cmdOpcion 
      BackColor       =   &H00C0C0C0&
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
      Index           =   3
      Left            =   720
      Picture         =   "frmMozoUsuario.frx":183EB
      Style           =   1  'Graphical
      TabIndex        =   10
      Top             =   11640
      Width           =   840
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00C0C0C0&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4770
      Left            =   15600
      TabIndex        =   3
      Top             =   360
      Width           =   11760
   End
   Begin VB.Frame Frame3 
      BackColor       =   &H00C0C0C0&
      Caption         =   " Mensajes "
      Height          =   1845
      Left            =   8640
      TabIndex        =   6
      Top             =   11640
      Width           =   7770
   End
   Begin VB.Timer Timer_girar 
      Enabled         =   0   'False
      Interval        =   2000
      Left            =   16830
      Top             =   9765
   End
   Begin VB.Timer Timer_LlenaRecordSet 
      Interval        =   1000
      Left            =   16350
      Top             =   9765
   End
   Begin VB.CommandButton Command12 
      BackColor       =   &H00C0C0C0&
      DragIcon        =   "frmMozoUsuario.frx":1912D
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   15480
      Picture         =   "frmMozoUsuario.frx":1956F
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   8160
      Width           =   765
   End
   Begin VB.CommandButton cmdOpcion 
      BackColor       =   &H00C0C0C0&
      Height          =   750
      Index           =   1
      Left            =   15960
      MaskColor       =   &H000000C0&
      Picture         =   "frmMozoUsuario.frx":199B1
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   7080
      Width           =   765
   End
   Begin VB.Timer Timer1 
      Interval        =   1000
      Left            =   15870
      Top             =   9765
   End
   Begin VB.Timer timSalida 
      Interval        =   3000
      Left            =   15390
      Top             =   9765
   End
   Begin MCI.MMControl mmControl 
      Height          =   375
      Left            =   17280
      TabIndex        =   0
      Top             =   9780
      Visible         =   0   'False
      Width           =   480
      _ExtentX        =   847
      _ExtentY        =   661
      _Version        =   393216
      PrevVisible     =   0   'False
      NextVisible     =   0   'False
      PlayVisible     =   0   'False
      PauseVisible    =   0   'False
      BackVisible     =   0   'False
      StepVisible     =   0   'False
      StopVisible     =   0   'False
      RecordVisible   =   0   'False
      EjectVisible    =   0   'False
      DeviceType      =   ""
      FileName        =   ""
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00C0C0C0&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2130
      Left            =   12600
      TabIndex        =   4
      Top             =   11880
      Width           =   7605
   End
   Begin VB.CommandButton cmdConsultaSaldo 
      BackColor       =   &H00C0C0C0&
      DragIcon        =   "frmMozoUsuario.frx":19AA3
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   15360
      Picture         =   "frmMozoUsuario.frx":19EE5
      Style           =   1  'Graphical
      TabIndex        =   9
      Top             =   10320
      Visible         =   0   'False
      Width           =   765
   End
   Begin MSComctlLib.ListView ListView1 
      Height          =   1605
      Left            =   1140
      TabIndex        =   14
      Top             =   6600
      Width           =   9720
      _ExtentX        =   17145
      _ExtentY        =   2831
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      FlatScrollBar   =   -1  'True
      _Version        =   393217
      ForeColor       =   8388608
      BackColor       =   16777215
      Appearance      =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial Black"
         Size            =   9
         Charset         =   0
         Weight          =   900
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      NumItems        =   0
   End
   Begin VB.Image imgNewProceso 
      Height          =   495
      Index           =   2
      Left            =   5040
      Top             =   4920
      Width           =   1935
   End
   Begin VB.Image imgNewProceso 
      Height          =   735
      Index           =   1
      Left            =   480
      Top             =   10080
      Width           =   2175
   End
   Begin VB.Image imgNewProceso 
      Height          =   1695
      Index           =   0
      Left            =   840
      Top             =   4200
      Width           =   1335
   End
   Begin VB.Image Command1 
      Height          =   1095
      Left            =   11160
      Top             =   7920
      Width           =   855
   End
   Begin VB.Image imgOpcion 
      Height          =   495
      Index           =   2
      Left            =   2880
      Top             =   4320
      Width           =   975
   End
   Begin VB.Image imgOpcion 
      Height          =   1575
      Index           =   3
      Left            =   9960
      Top             =   4320
      Width           =   1095
   End
   Begin VB.Image ImgLogo 
      Height          =   3900
      Left            =   5400
      Stretch         =   -1  'True
      Top             =   90
      Width           =   6520
   End
   Begin VB.Image imgOpcion 
      Height          =   855
      Index           =   1
      Left            =   0
      Top             =   1560
      Width           =   855
   End
   Begin VB.Image ImagePais 
      Height          =   480
      Left            =   11160
      Top             =   165
      Width           =   630
   End
   Begin VB.Label text1 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   36
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   855
      Left            =   4320
      TabIndex        =   8
      Top             =   5400
      Width           =   3375
   End
   Begin VB.Image ImgLogo2 
      Height          =   1305
      Left            =   1920
      Stretch         =   -1  'True
      Top             =   11400
      Width           =   3000
   End
   Begin VB.Image Image2 
      Height          =   1005
      Left            =   5040
      Picture         =   "frmMozoUsuario.frx":1A627
      Stretch         =   -1  'True
      Top             =   11520
      Width           =   3030
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      Caption         =   "www.infhotel.com.pe"
      BeginProperty Font 
         Name            =   "Arial Black"
         Size            =   9.75
         Charset         =   0
         Weight          =   900
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00BF8801&
      Height          =   255
      Left            =   5040
      TabIndex        =   7
      Top             =   12480
      Width           =   3030
   End
   Begin VB.Label txtFecha 
      Alignment       =   2  'Center
      BackColor       =   &H00C0C0C0&
      BackStyle       =   0  'Transparent
      Caption         =   "Fecha"
      BeginProperty Font 
         Name            =   "Arial Black"
         Size            =   15
         Charset         =   0
         Weight          =   900
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00E0E0E0&
      Height          =   465
      Left            =   0
      TabIndex        =   1
      Top             =   120
      Width           =   5355
   End
   Begin VB.Image Image1 
      Height          =   9000
      Left            =   0
      Picture         =   "frmMozoUsuario.frx":ABD53
      Top             =   0
      Width           =   12000
   End
End
Attribute VB_Name = "frmMozoUsuario"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim RsMozo As Recordset
Dim rsMensajeCocina As Recordset
Dim fso As Object
Dim sql_mensaje As String
Dim nroElementos As Integer
   Dim RsTc As New ADODB.Recordset
Dim i As Integer
Private Const PI = 3.14159265
Dim inicio As Boolean

Private Sub ValidacionEntrarConLicencia()
    If lHARDkey Then
        '----------Verifica Llave HK----------------------------------
        Dim verif As Boolean
        verif = hk.VerificaConexion
                        
        If verif = False Then
            Dim str As String
            str = hk.IniciaConexion(InfhotelHK.Adicion)
            If str <> "" Then
                MsgBox str, vbCritical, "Aviso"
                End
            End If
        End If
        '--------------------------------------------------------------
    End If
End Sub

Private Sub cmdConsultaSaldo_Click()
    frmConsultaSaldo.Show vbModal
End Sub

Private Sub cmdOpcion_Click(Index As Integer)
    Select Case Index
           Case Is = 0 ' Aceptar
                'TIPO CAMBIO
                  If lMCPV = False Then
                If pais <> "002" Then
                        Dim RsTc As New ADODB.Recordset
                        Set RsTc = Lib.OpenRecordset("SELECT * From TTIPOCAMBIO WHERE (fFecha = {fn CURDATE() })", Cn)
                
                        If RsTc.EOF Then
                         nTC = 0
                        Else
                         nTC = IIf(IsNull(RsTc!nVenta), 0, IIf(IsNull(RsTc!nVenta), 0, RsTc!nVenta))
                        End If
                
        '                Set rstc = Nothing
        '                Set RsCaja = Nothing
        '                Set RsParametro = Nothing
                        wInicio = False
                        If nTC = 0 Then
                         MsgBox "Error: No se ha ingresado el Tipo de Cambio", vbCritical, sMensaje
                         txtPassword.Text = ""
                         txtPassword.SetFocus
                         Exit Sub
                        End If
                End If
                 End If
                
                If lHARDkey Then
                    ValidacionEntrarConLicencia
                End If
                If txtPassword.Text = "" Then
                   MsgBox "Ingrese su password", vbExclamation, sMensaje
                   txtPassword.SetFocus
                   Exit Sub
                Else
                    RsMozo.MoveFirst
                    Do While Not RsMozo.EOF
                       If Desencapsula(IIf(IsNull(RsMozo!tValor), "", RsMozo!tValor)) = UCase(txtPassword.Text) Or (IIf(IsNull(RsMozo!tBandaMagnetica), "", RsMozo!tBandaMagnetica) = Encapsula(UCase(Extrae(txtPassword.Text))) And Encapsula(UCase(Extrae(txtPassword.Text))) <> "") Then
                             sVar1 = RsMozo!tResumido
                             If RsMozo!nTamano = 1 Then
                                Dim sCambio As String
                                MsgBox "Por motivo de seguridad, deberá cambiar su password", vbCritical, sMensaje
                                
                                frmPassword.Caption = "Ingrese su nueva clave"
                                frmPassword.Show vbModal
                                If Not wEnter Then
                                   Exit Sub
                                End If
                                sCambio = sDescrip
                                frmPassword.Caption = "Confirme su nueva clave"
                                frmPassword.Show vbModal
                                If Not wEnter Then
                                   Exit Sub
                                End If
                                If sCambio <> sDescrip Or sDescrip = "" Then
                                   MsgBox "Confirmación erronea, no se realizó el cambio", vbCritical, sMensaje
                                   Exit Sub
                                End If
                                RsMozo!nTamano = 0
                                RsMozo!tValor = Encapsula(sDescrip)
                             End If
                          sPassword = UCase(txtPassword.Text)
                          sMozo = RsMozo!codigo
                          lSomelier = IIf(IsNull(RsMozo!nValor), 0, RsMozo!nValor)
                          txtPassword.Text = ""
                          If lMCPV Then
                             If sModulo = "INFOREST" Then
                                sUsuario = sVar1
                                Unload Me
                             Else
                             
                                'audirotia
                                
                                registroAccesoAuditoria "I", sVar1
                                
                                'auditoria


                                Me.Timer_LlenaRecordSet.Enabled = False
                                Me.Timer_LlenaRecordSet.Interval = 0
                                frmCargoMozo.Show vbModal
                                Me.Timer_LlenaRecordSet.Enabled = True
                                Me.Timer_LlenaRecordSet.Interval = 1000
                                
                                
                             End If
                          Else
                          
                                'audirotia
                                
                                registroAccesoAuditoria "I", sVar1
                                
                                'auditoria
                                
                                
                                Me.Timer_LlenaRecordSet.Enabled = False
                                Me.Timer_LlenaRecordSet.Interval = 0
                                frmCargoMozo.Show vbModal
                                Me.Timer_LlenaRecordSet.Enabled = True
                                Me.Timer_LlenaRecordSet.Interval = 1000
                             
                             
                          End If
                          timSalida.Enabled = False
                          Exit Sub
                       End If
                       RsMozo.MoveNext
                    Loop
                    MsgBox "Usuario no Encontrado", vbCritical, sMensaje
                    txtPassword.Text = ""
                    txtPassword.SetFocus
                   Exit Sub
                End If
                                
           Case Is = 2
                frmPassword.cmdOpcion.Visible = False
                frmPassword.Show vbModal
                If wEnter Then
                   txtPassword.Text = sDescrip
                End If
                
           'HUELLA
           Case Is = 3
                   If pais <> "002" Then
                      
                                Set RsTc = Lib.OpenRecordset("SELECT * From TTIPOCAMBIO WHERE (fFecha = {fn CURDATE() })", Cn)
                        
                                If RsTc.EOF Then
                                 nTC = 0
                                Else
                                 nTC = IIf(IsNull(RsTc!nVenta), 0, IIf(IsNull(RsTc!nVenta), 0, RsTc!nVenta))
                                End If
                        
                '                Set rstc = Nothing
                '                Set RsCaja = Nothing
                '                Set RsParametro = Nothing
                                wInicio = False
                                If nTC = 0 Then
                                 MsgBox "Error: No se ha ingresado el Tipo de Cambio", vbCritical, sMensaje
                                 Exit Sub
                                End If
                        End If
                wEnterHuella = False
                frmVerificacionHuella.Show vbModal
                If wEnterHuella Then
                                Me.Timer_LlenaRecordSet.Enabled = False
                                Me.Timer_LlenaRecordSet.Interval = 0
                                frmCargoMozo.Show vbModal
                                Me.Timer_LlenaRecordSet.Enabled = True
                                Me.Timer_LlenaRecordSet.Interval = 1000
                End If
                
                
           Case Is = 1
                If lHARDkey Then
                    '----------Verifica Llave HK----------------------------------
                    If hk.ValidaLlave Then
                        'MsgBox "Fallo la validacion de la llave", vbCritical, "Aviso"
                        Dim result As Boolean
                        If sModulo = "INFOREST" Then
                            result = hk.FinalizarConexion(Aplicacion.PuntoVenta)
                        End If
                        If sModulo = "ADICION" Then
                            result = hk.FinalizarConexion(Aplicacion.Adicion)
                        End If
                        End
                    End If
                    '--------------------------------------------------------------
                End If
                End

    End Select
    txtPassword.SelStart = Len(txtPassword.Text)
    txtPassword.SetFocus
End Sub

'Private Declare Function GetCursorPos Lib "user32" (lpPoint As POINTAPI) As Long
  Private Sub Command1_Click()
  frmAbout.Show vbModal

End Sub


Private Sub Form_Activate()
On Error GoTo fin
    txtPassword.SetFocus
fin:
End Sub

Private Sub Form_Load()
   Centrar Me
   inicio = True
   timSalida.Enabled = False
   timSalida.Interval = nSalir
   text1.Caption = Format(Time, "HH:mm:ss")

   
   'Frame3.Visible = False
   If lSiab Then
      cmdConsultaSaldo.Visible = True
      cmdOpcion(1).Height = 830
   End If
   txtPassword.Text = ""
   If lMCPV Then
      Isql = "select tCodigoUsuario as Codigo, tResumido, tPassword as tValor, tBandaMagnetica, nTamano=0, nValor=0 from vGrupousuario Where lActivo = 1 And lModulo01 = 1"
   Else
      Isql = "select * from vMozo where lActivo = 1 Order by nBoton"
   End If
   Set RsMozo = Lib.OpenRecordset(Isql, Cn)
   txtFecha.Caption = Format(FechaServidor(), "long date")

    'On Error Resume Next
    Set fso = CreateObject("Scripting.FileSystemObject")
    If fso.FileExists(App.Path & "\bmps\Cliente.avi") Then
       ImgLogo.Visible = False
       imgLogoPic.Visible = True
       With mmControl
            .FileName = App.Path & "\bmps\Cliente.avi"
            .Command = "open"
            .hWndDisplay = imgLogoPic.hwnd
            .DeviceType = "AVIVideo"
            .From = 0
            .Notify = True
            .Command = "play"
       End With
    Else
       If fso.FileExists(App.Path & "\bmps\Cliente.jpg") Then
          ImgLogo.Picture = LoadPicture(App.Path & "\bmps\Cliente.jpg")
       End If
    End If
    If fso.FileExists(App.Path & "\bmps\logo.jpg") Then
       ImgLogo2.Picture = LoadPicture(App.Path & "\bmps\Logo.jpg")
    End If
    
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
    
'    If fso.FileExists(App.Path & "\bmps\Pais.jpg") Then
'       ImagePais.Picture = LoadPicture(App.Path & "\bmps\Pais.jpg")
'    End If
    
    
    Set fso = Nothing
    'txtPassword.SetFocus
End Sub




Private Sub imgNewProceso_Click(Index As Integer)
    MsgBox "Proceso en Desarrollo, Aun no culminado!!!", vbInformation, sMensaje
End Sub

Private Sub imgOpcion_Click(Index As Integer)
    Select Case Index
           Case Is = 0 ' Aceptar
                'TIPO CAMBIO
                  If lMCPV = False Then
                If pais <> "002" Then
                        Dim RsTc As New ADODB.Recordset
                        Set RsTc = Lib.OpenRecordset("SELECT * From TTIPOCAMBIO WHERE (fFecha = {fn CURDATE() })", Cn)
                
                        If RsTc.EOF Then
                         nTC = 0
                        Else
                         nTC = IIf(IsNull(RsTc!nVenta), 0, IIf(IsNull(RsTc!nVenta), 0, RsTc!nVenta))
                        End If
                
        '                Set rstc = Nothing
        '                Set RsCaja = Nothing
        '                Set RsParametro = Nothing
                        wInicio = False
                        If nTC = 0 Then
                         MsgBox "Error: No se ha ingresado el Tipo de Cambio", vbCritical, sMensaje
                         txtPassword.Text = ""
                         txtPassword.SetFocus
                         Exit Sub
                        End If
                End If
                 End If
                
                If lHARDkey Then
                    ValidacionEntrarConLicencia
                End If
                If txtPassword.Text = "" Then
                   MsgBox "Ingrese su password", vbExclamation, sMensaje
                   txtPassword.SetFocus
                   Exit Sub
                Else
                    RsMozo.MoveFirst
                    Do While Not RsMozo.EOF
                       If Desencapsula(IIf(IsNull(RsMozo!tValor), "", RsMozo!tValor)) = UCase(txtPassword.Text) Or (IIf(IsNull(RsMozo!tBandaMagnetica), "", RsMozo!tBandaMagnetica) = Encapsula(UCase(Extrae(txtPassword.Text))) And Encapsula(UCase(Extrae(txtPassword.Text))) <> "") Then
                             sVar1 = RsMozo!tResumido
                             If RsMozo!nTamano = 1 Then
                                Dim sCambio As String
                                MsgBox "Por motivo de seguridad, deberá cambiar su password", vbCritical, sMensaje
                                
                                frmPassword.Caption = "Ingrese su nueva clave"
                                frmPassword.Show vbModal
                                If Not wEnter Then
                                   Exit Sub
                                End If
                                sCambio = sDescrip
                                frmPassword.Caption = "Confirme su nueva clave"
                                frmPassword.Show vbModal
                                If Not wEnter Then
                                   Exit Sub
                                End If
                                If sCambio <> sDescrip Or sDescrip = "" Then
                                   MsgBox "Confirmación erronea, no se realizó el cambio", vbCritical, sMensaje
                                   Exit Sub
                                End If
                                RsMozo!nTamano = 0
                                RsMozo!tValor = Encapsula(sDescrip)
                             End If
                          sPassword = UCase(txtPassword.Text)
                          sMozo = RsMozo!codigo
                          lSomelier = IIf(IsNull(RsMozo!nValor), 0, RsMozo!nValor)
                          txtPassword.Text = ""
                          If lMCPV Then
                             If sModulo = "INFOREST" Then
                                sUsuario = sVar1
                                Unload Me
                             Else
                             
                                'audirotia
                                
                                registroAccesoAuditoria "I", sVar1
                                
                                'auditoria


                                Me.Timer_LlenaRecordSet.Enabled = False
                                Me.Timer_LlenaRecordSet.Interval = 0
                                frmCargoMozo.Show vbModal
                                Me.Timer_LlenaRecordSet.Enabled = True
                                Me.Timer_LlenaRecordSet.Interval = 1000
                                
                                
                             End If
                          Else
                          
                                'audirotia
                                
                                registroAccesoAuditoria "I", sVar1
                                
                                'auditoria
                                
                                
                                Me.Timer_LlenaRecordSet.Enabled = False
                                Me.Timer_LlenaRecordSet.Interval = 0
                                frmCargoMozo.Show vbModal
                                Me.Timer_LlenaRecordSet.Enabled = True
                                Me.Timer_LlenaRecordSet.Interval = 1000
                             
                             
                          End If
                          timSalida.Enabled = False
                          Exit Sub
                       End If
                       RsMozo.MoveNext
                    Loop
                    MsgBox "Usuario no Encontrado", vbCritical, sMensaje
                    txtPassword.Text = ""
                    txtPassword.SetFocus
                   Exit Sub
                End If
                                
           Case Is = 2
                frmPassword.cmdOpcion.Visible = False
                frmPassword.Show vbModal
                If wEnter Then
                   txtPassword.Text = sDescrip
                End If
                
           'HUELLA
           Case Is = 3
                   If pais <> "002" Then
                      
                                Set RsTc = Lib.OpenRecordset("SELECT * From TTIPOCAMBIO WHERE (fFecha = {fn CURDATE() })", Cn)
                        
                                If RsTc.EOF Then
                                 nTC = 0
                                Else
                                 nTC = IIf(IsNull(RsTc!nVenta), 0, IIf(IsNull(RsTc!nVenta), 0, RsTc!nVenta))
                                End If
                        
                '                Set rstc = Nothing
                '                Set RsCaja = Nothing
                '                Set RsParametro = Nothing
                                wInicio = False
                                If nTC = 0 Then
                                 MsgBox "Error: No se ha ingresado el Tipo de Cambio", vbCritical, sMensaje
                                 Exit Sub
                                End If
                        End If
                wEnterHuella = False
                frmVerificacionHuella.Show vbModal
                If wEnterHuella Then
                                Me.Timer_LlenaRecordSet.Enabled = False
                                Me.Timer_LlenaRecordSet.Interval = 0
                                frmCargoMozo.Show vbModal
                                Me.Timer_LlenaRecordSet.Enabled = True
                                Me.Timer_LlenaRecordSet.Interval = 1000
                End If
                
                
           Case Is = 1
                If lHARDkey Then
                    '----------Verifica Llave HK----------------------------------
                    If hk.ValidaLlave Then
                        'MsgBox "Fallo la validacion de la llave", vbCritical, "Aviso"
                        Dim result As Boolean
                        If sModulo = "INFOREST" Then
                            result = hk.FinalizarConexion(Aplicacion.PuntoVenta)
                        End If
                        If sModulo = "ADICION" Then
                            result = hk.FinalizarConexion(Aplicacion.Adicion)
                        End If
                        End
                    End If
                    '--------------------------------------------------------------
                End If
                End

    End Select
    txtPassword.SelStart = Len(txtPassword.Text)
    txtPassword.SetFocus
End Sub

Private Sub mmControl_Done(NotifyCode As Integer)
     With mmControl
       .From = 0
       .Command = "play"
    End With
End Sub

Private Sub Timer_girar_Timer()
   Generar_Movimiento
End Sub

Private Sub Timer_LlenaRecordSet_Timer()
    ListView1.ListItems.Clear
     Llenar_RecordSet
End Sub

Private Sub txtPassword_GotFocus()
    If Trim(txtPassword.Text) <> "" Then
      cmdOpcion_Click (0)
   End If
End Sub

Private Sub txtPassword_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      cmdOpcion_Click (0)
   End If
End Sub

Private Sub timSalida_Timer()
   Do While Screen.ActiveForm.Name <> "frmMozoUsuario"
      Unload Screen.ActiveForm
   Loop
End Sub

Public Sub ReseteaTimer()
   frmMozoUsuario.timSalida.Enabled = False
   frmMozoUsuario.timSalida.Enabled = True
End Sub

Private Sub Reloj()
   Static last_time As Date

   Dim Cx As Single
   Dim cy As Single
   Dim num As Single
   Dim radius As Single
   Dim theta As Single

     If last_time = Now Then Exit Sub

     last_time = Now
     Picture1.Cls
     Picture1.ForeColor = vbBlue
     Cx = Picture1.ScaleWidth / 2
     cy = Picture1.ScaleHeight / 2

     ' Horas
     num = 5 * (DatePart("h", last_time) + DatePart("n", last_time) / _
                                 60 + DatePart("s", last_time) / 3600)
     theta = MinutesToRadians(num)
     radius = Picture1.ScaleWidth * 0.24
     Picture1.ForeColor = &H26050F
     Picture1.DrawWidth = 7
     Picture1.Line (Cx, cy)-Step(radius * Cos(theta), -radius * Sin(theta))

     ' Los Minutos
     num = DatePart("n", last_time)
     theta = MinutesToRadians(num)
     radius = Picture1.ScaleWidth * 0.37
     Picture1.ForeColor = &H26050F
     Picture1.DrawWidth = 6
     Picture1.Line (Cx, cy)-Step(radius * Cos(theta), -radius * Sin(theta))

     ' Los segundos
     num = DatePart("s", last_time)
     theta = MinutesToRadians(num)
     radius = Picture1.ScaleWidth * 0.34
     Picture1.ForeColor = &H1E32CD
     Picture1.DrawWidth = 4
     Picture1.Line (Cx, cy)-Step(radius * Cos(theta), -radius * Sin(theta))
    'Call RetornarMensajes
    
 End Sub
 Private Function MinutesToRadians(ByVal num As Single) As Single
     MinutesToRadians = (15 - num) * 2 * PI / 60
 End Function
Private Sub Timer1_Timer()
    text1.Caption = Format(Time, "HH:mm:ss")
    'Reloj
        
         
End Sub

Private Sub Llenar_RecordSet()
    Dim X As Integer
    Dim Item As ListItem
    Isql = "usp_listadoMensajes"
    Set rsMensajeCocina = Lib.OpenRecordset(Isql, Cn)
    If rsMensajeCocina.EOF Or rsMensajeCocina.BOF Then
       'Frame3.Visible = False
       'ListView1.Visible = False
       Timer_girar.Enabled = False
    Else
            With ListView1
                .View = lvwReport
                .ListItems.Clear
                .ColumnHeaders.Clear
            End With
            nroElementos = rsMensajeCocina.RecordCount
            Timer_LlenaRecordSet.Interval = (nroElementos + 1) * 1000
            Me.MousePointer = vbHourglass
            ListView1.ColumnHeaders.Add , , "", 7000
            
            rsMensajeCocina.MoveFirst
            ListView1.ListItems.Clear
            While Not rsMensajeCocina.EOF
                Set Item = ListView1.ListItems.Add(, , rsMensajeCocina.Fields(0))
                rsMensajeCocina.MoveNext
            Wend
            If rsMensajeCocina.EOF Then
                ListView1.ListItems.Add , , "-"
            End If
            Me.MousePointer = vbDefault
            'Frame3.Visible = True
            ListView1.Visible = True
             Timer_girar.Enabled = True
             Timer_girar.Interval = 2000
    End If
End Sub
Private Sub Generar_Movimiento()
    Dim temp As String
    Dim X As Integer
    Dim numero_lista As Integer
    numero_lista = ListView1.ListItems.Count
    temp = ListView1.ListItems.Item(1)
    For X = 1 To numero_lista - 1
        ListView1.ListItems.Item(X) = ListView1.ListItems.Item(X + 1)
        If X = numero_lista - 1 Then
            ListView1.ListItems.Item(numero_lista) = temp
        End If
    Next X
End Sub


