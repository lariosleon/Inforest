VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmInicio 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Inicio de Turno"
   ClientHeight    =   4440
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7065
   Icon            =   "frmInicio.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4440
   ScaleWidth      =   7065
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdOpcion 
      Caption         =   "Tipo de Cambio Oficial"
      Height          =   555
      Index           =   5
      Left            =   2880
      TabIndex        =   27
      Top             =   1260
      Width           =   1275
   End
   Begin VB.Frame fraMontos 
      Caption         =   " Montos "
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1860
      Left            =   90
      TabIndex        =   10
      Top             =   1920
      Width           =   6900
      Begin VB.TextBox txtAbonoE 
         Alignment       =   1  'Right Justify
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00808000&
         Height          =   285
         Left            =   2475
         TabIndex        =   14
         Top             =   1350
         Width           =   1365
      End
      Begin VB.TextBox txtAbonoN 
         Alignment       =   1  'Right Justify
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00808000&
         Height          =   285
         Left            =   2475
         TabIndex        =   13
         Top             =   765
         Width           =   1365
      End
      Begin VB.CommandButton cmdOpcion 
         Caption         =   "Abono MN"
         Height          =   555
         Index           =   3
         Left            =   210
         TabIndex        =   12
         Top             =   555
         Width           =   1275
      End
      Begin VB.CommandButton cmdOpcion 
         Caption         =   "Abono ME"
         Height          =   555
         Index           =   4
         Left            =   210
         TabIndex        =   11
         Top             =   1155
         Width           =   1275
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Saldo Anterior"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   4005
         TabIndex        =   23
         Top             =   360
         Width           =   1215
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "Monto Inicial"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   5505
         TabIndex        =   22
         Top             =   360
         Width           =   1110
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         Caption         =   "Abono Actual"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   2580
         TabIndex        =   21
         Top             =   360
         Width           =   1155
      End
      Begin VB.Label txtAnteriorN 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00808000&
         Height          =   285
         Left            =   3930
         TabIndex        =   20
         Top             =   765
         Width           =   1365
      End
      Begin VB.Label txtAnteriorE 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00808000&
         Height          =   285
         Left            =   3930
         TabIndex        =   19
         Top             =   1350
         Width           =   1365
      End
      Begin VB.Label txtSaldoN 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   285
         Left            =   5385
         TabIndex        =   18
         Top             =   765
         Width           =   1365
      End
      Begin VB.Label txtSaldoE 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   285
         Left            =   5385
         TabIndex        =   17
         Top             =   1350
         Width           =   1365
      End
      Begin VB.Label txtME 
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   1605
         TabIndex        =   16
         Top             =   1350
         Width           =   690
      End
      Begin VB.Label txtMN 
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   1605
         TabIndex        =   15
         Top             =   765
         Width           =   690
      End
   End
   Begin VB.CommandButton cmdOpcion 
      Caption         =   "Tipo de Cambio"
      Height          =   555
      Index           =   2
      Left            =   2880
      TabIndex        =   9
      Top             =   495
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
      Height          =   555
      Index           =   1
      Left            =   4380
      Picture         =   "frmInicio.frx":0442
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   3810
      Width           =   1275
   End
   Begin VB.CommandButton cmdOpcion 
      Caption         =   "Apertura"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   555
      Index           =   0
      Left            =   5730
      Picture         =   "frmInicio.frx":0544
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   3810
      Width           =   1275
   End
   Begin MSComCtl2.Animation aniVideo 
      Height          =   540
      Left            =   120
      TabIndex        =   24
      Top             =   3840
      Visible         =   0   'False
      Width           =   735
      _ExtentX        =   1296
      _ExtentY        =   953
      _Version        =   393216
      FullWidth       =   49
      FullHeight      =   36
   End
   Begin VB.Label Label9 
      AutoSize        =   -1  'True
      Caption         =   "Oficial :"
      Height          =   195
      Left            =   780
      TabIndex        =   28
      Top             =   1560
      Width           =   525
   End
   Begin VB.Label Label8 
      AutoSize        =   -1  'True
      Caption         =   "Tipo de Cambio"
      Height          =   195
      Left            =   120
      TabIndex        =   26
      Top             =   1305
      Width           =   1110
   End
   Begin VB.Label txtTCO 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   1395
      TabIndex        =   6
      Top             =   1260
      Width           =   1365
   End
   Begin VB.Label lblProceso 
      BackStyle       =   0  'Transparent
      Caption         =   "Actualizando datos. Este proceso puede tomar algunos minutos."
      ForeColor       =   &H00404040&
      Height          =   585
      Left            =   900
      TabIndex        =   25
      Top             =   3780
      Visible         =   0   'False
      Width           =   3225
   End
   Begin VB.Label txtTC 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   1395
      TabIndex        =   5
      Top             =   765
      Width           =   1365
   End
   Begin VB.Label TxtFecha 
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   1395
      TabIndex        =   4
      Top             =   405
      Width           =   1365
   End
   Begin VB.Label txtUsuario 
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   1395
      TabIndex        =   3
      Top             =   45
      Width           =   1365
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "Fecha :"
      Height          =   195
      Left            =   750
      TabIndex        =   2
      Top             =   450
      Width           =   540
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Tipo de Cambio :"
      Height          =   195
      Left            =   90
      TabIndex        =   1
      Top             =   810
      Width           =   1200
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Usuario :"
      Height          =   195
      Left            =   660
      TabIndex        =   0
      Top             =   90
      Width           =   630
   End
End
Attribute VB_Name = "frmInicio"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim RsTc As Recordset
Dim RsTurno As Recordset
Dim wAgrega As Boolean
Dim nAnteriorN As Double
Dim nAnteriorE As Double
Dim nAbonoN As Double
Dim nAbonoE As Double
Dim nMontoSN As Double
Dim nMontoSE As Double
 

Private Sub cmdOpcion_Click(Index As Integer)
   Dim nCorrela As String
   Dim xMensaje As String
   Dim wPasa As Boolean
   
   Select Case Index
          Case Is = 0 ' Aperturar
               If sMonE <> "" And sMonN <> sMonE And nTC = 0 And nTCO = 0 Then
                  MsgBox "Tipo de cambio no ingresado", vbExclamation, sMensaje
                  Exit Sub
               End If
               '-----------------
               If RsTurno.RecordCount = 0 Then

                  If MsgBox("Seguro de Aperturar el Turno?", vbQuestion + vbOKCancel, sMensaje) = vbCancel Then
                     Exit Sub
                  End If
               Else
                  If RsTurno!lCierre = True Then
                     If MsgBox("Seguro de Aperturar el Turno?", vbQuestion + vbOKCancel, sMensaje) = vbCancel Then
                        Exit Sub
                     End If
                  Else
                     If MsgBox("Seguro de Re Aperturar el Turno?", vbQuestion + vbOKCancel, sMensaje) = vbCancel Then
                        Exit Sub
                     End If
                  End If
               End If
               
               If wAgrega Then
                  Dim oComando As clsComando
                  Set oComando = New clsComando
                  If Not oComando.CreateCmdSp("spIns_TipoCambio", Cn) Then
                     Set oComando = Nothing
                     Exit Sub
                  End If
                  '---CESAR - tipocambio
                  oComando.CreateParameter "@nTc", adDouble, adParamInput, 0, nTC
                  oComando.CreateParameter "@tUSUARIO", adVarChar, adParamInput, 15, sUsuario
                  oComando.CreateParameter "@nTco", adDouble, adParamInput, 0, nTCO
                  If Not oComando.GetParamOK Then
                  '------------
                     Set oComando = Nothing
                     Exit Sub
                  End If
                  If Not oComando.ExecSP Then
                     Set oComando = Nothing
                     Exit Sub
                  End If
               End If
                              
               If RsTurno.RecordCount = 0 Then
                  wPasa = True
               Else
                  wPasa = RsTurno!lCierre
               End If
                  
               If wPasa Then
                  'Obtiene el Correlativo
                  nCorrela = Calcular("select max(tTurno) as Codigo from MTURNO where substring(tTurno,1,2)= substring(ltrim(str(year(getdate()))),3,2)", Cn)
                  If IsNull(nCorrela) Or Mid(nCorrela, 1, 2) <> Mid(Trim(str(Year(FechaServidor()))), 3, 2) Then
                     nCorrela = Mid(Trim(str(Year(FechaServidor()))), 3, 2) & "00000001"
                  Else
                     nCorrela = Mid(Trim(str(Year(FechaServidor()))), 3, 2) & Lib.Correlativo(Mid(nCorrela, 3, 8), 8)
                  End If
                  
                  Isql = "insert into MTURNO( " & _
                         "tTurno, tCaja, tSalon, fInicial, tUsuario, lCierre, nMontoIN, nMontoIE) " & _
                         "values ('" & nCorrela & "', " & _
                                 "'" & sCaja & "', " & _
                                 "'" & sSalon & "', " & _
                                 "getdate() , " & _
                                 "'" & sUsuario & "', " & _
                                        0 & ", " & _
                                 nAbonoN & ", " & _
                                 nAbonoE & ")"
                  Cn.Execute Isql
                  sTurno = nCorrela
                  If lAlmacenRemoto = True Then
                     actualizaDatosSistemaAlmacen
                  End If
                  
               Else
                  sTurno = RsTurno!tTurno
                   
                  Isql = "update MTURNO set " & _
                         "tUsuario ='" & sUsuario & "', " & _
                         "nMontoIN = " & nAbonoN & ", " & _
                         "nMontoIE = " & nAbonoE & " " & _
                         "where tTurno ='" & sTurno & "'"
                   Cn.Execute Isql
               End If
               ActivaInicio (True)
               wInicio = True
               Unload Me
          
          Case Is = 1 ' Cancelar
               Unload Me

          Case Is = 2 ' Tipo de Cambio
              sTipo = "TC"
               frmNumPad.Show vbModal
               If wEnter Then
                  txtTC.Caption = Format(sDescrip, "###,###,##0.000")
                  nTC = Val(sDescrip)
               End If
          Case Is = 3 ' Abono MN
               sTipo = ""
               frmNumPad.Show vbModal
               nAbonoN = IIf(wEnter = True, sDescrip, nAbonoN)
               txtAbonoN.Text = Format(nAbonoN, "###,###,###,##0.00")
               txtSaldoN.Caption = Format(nAbonoN + nAnteriorN, "###,###,###,##0.00")
          
          Case Is = 4 ' Abono ME
               sTipo = ""
               frmNumPad.Show vbModal
               nAbonoE = IIf(wEnter = True, sDescrip, nAbonoE)
               txtAbonoE.Text = Format(nAbonoE, "###,###,###,##0.00")
               txtSaldoE.Caption = Format(nAbonoE + nAnteriorE, "###,###,###,##0.00")
               
         '---CESAR tipo cambio Oficial
         Case Is = 5
                sTipo = "TC"
               frmNumPad.Show vbModal
               If wEnter Then
                  txtTCO.Caption = Format(sDescrip, "###,###,##0.000")
                  nTCO = Val(sDescrip)
               End If
          '---------------
   End Select
End Sub

Private Sub Form_Activate()
   If CDate(Format(txtFecha.Caption, "short date")) > FechaServidor() Then
      MsgBox "Error : Ha querido ingresar un turno con fecha anterior", vbCritical, sMensaje
      Unload Me
   End If
End Sub

Private Sub Form_Load()
  
   Centrar Me
   On Error Resume Next
   aniVideo.Open App.Path & "\bmps\FileMove.avi"
   If lMCPV Then
      Isql = "select * from MTURNO where tUsuario ='" & sUsuario & "' order by tTurno"
   Else
      Isql = "select * from MTURNO where tCaja ='" & sCaja & "' order by tTurno"
   End If
   Set RsTurno = Lib.OpenRecordset(Isql, Cn)
   
   If RsTurno.RecordCount = 0 Then
      nAbonoN = 0
      nAbonoE = 0
      nAnteriorN = 0
      nAnteriorE = 0
      nMontoSN = 0
      nMontoSE = 0
      
      txtAbonoN.Text = Format(nAbonoN, "###,###,##0.00")
      txtAbonoE.Text = Format(nAbonoE, "###,###,##0.00")
      txtAnteriorN.Caption = Format(nAnteriorN, "###,###,##0.00")
      txtAnteriorE.Caption = Format(nAnteriorE, "###,###,##0.00")
      txtSaldoN.Caption = Format(nMontoSN, "###,###,##0.00")
      txtSaldoE.Caption = Format(nMontoSE, "###,###,##0.00")
            
      txtFecha.Caption = FechaServidor()
      Me.Caption = "Apertura de Turno"
   Else
      RsTurno.MoveLast
                  
      If Not RsTurno!lCierre = True Then
         txtFecha.Caption = IIf(IsNull(RsTurno!finicial), Now, RsTurno!finicial)
         Me.Caption = "Re Apertura de Turno"
         nAbonoN = IIf(IsNull(RsTurno!nMontoIN), 0, RsTurno!nMontoIN)
         nAbonoE = IIf(IsNull(RsTurno!nMontoIE), 0, RsTurno!nMontoIE)
         nAnteriorN = 0
         nAnteriorE = 0
         nMontoSN = nAbonoN + nAnteriorN
         nMontoSE = nAbonoE + nAnteriorE
                           
         txtAbonoN.Text = Format(nAbonoN, "###,###,##0.00")
         txtAbonoE.Text = Format(nAbonoE, "###,###,##0.00")
         txtAnteriorN.Caption = Format(nAnteriorN, "###,###,##0.00")
         txtAnteriorE.Caption = Format(nAnteriorE, "###,###,##0.00")
         txtSaldoN.Caption = Format(nMontoSN, "###,###,##0.00")
         txtSaldoE.Caption = Format(nMontoSE, "###,###,##0.00")
      Else
         Me.Caption = "Apertura de Turno"
         txtFecha.Caption = FechaServidor()
         
         nAbonoN = 0
         nAbonoE = 0
         nAnteriorN = 0
         nAnteriorE = 0
         
         nMontoSN = nAbonoN + nAnteriorN
         nMontoSE = nAbonoE + nAnteriorE
         
         txtAbonoN.Text = Format(nAbonoN, "###,###,##0.00")
         txtAbonoE.Text = Format(nAbonoE, "###,###,##0.00")
         txtAnteriorN.Caption = Format(nAnteriorN, "###,###,##0.00")
         txtAnteriorE.Caption = Format(nAnteriorE, "###,###,##0.00")
         txtSaldoN.Caption = Format(nMontoSN, "###,###,##0.00")
         txtSaldoE.Caption = Format(nMontoSE, "###,###,##0.00")
      End If
   End If
   
   If sMonE <> "" And sMonN <> sMonE Then
      Set RsTc = Lib.OpenRecordset("SELECT * From TTIPOCAMBIO WHERE (fFecha = {fn CURDATE() })", Cn)
           
      If RsTc.EOF Then
         nTC = 0
         nTCO = 0
         wAgrega = True
      Else
         nTC = IIf(IsNull(RsTc!nVenta), 0, IIf(IsNull(RsTc!nVenta), 0, RsTc!nVenta))
         'CESAR-----para mostrar el tipo cambio si ya esta registrado
         nTCO = IIf(IsNull(RsTc!nOficial), 0, IIf(IsNull(RsTc!nOficial), 0, RsTc!nOficial))
         wAgrega = False
         
         If nTC = 0 Then: wAgrega = True
         If nTCO = 0 Then: wAgrega = True
         
      End If
      If nTC = 0 And Not lInfhotel Then
         cmdOpcion(2).Visible = True
      Else
         cmdOpcion(2).Visible = False
      End If
      
      If nTCO = 0 And Not lInfhotel Then
         '---CESAR
         cmdOpcion(5).Visible = True
      Else
         '---CESAR
         cmdOpcion(5).Visible = False
      End If
      
      txtTC.Caption = Format(nTC, "#,###,##0.000")
      '---CESAR tipo cambio
      txtTCO.Caption = Format(nTCO, "#,###,##0.000")
      txtMN.Caption = sMonN
      txtME.Caption = sMonE
   Else
      cmdOpcion(2).Visible = False
      cmdOpcion(4).Visible = False
      '---CESAR tipo cambio
      cmdOpcion(5).Visible = False
      txtMN.Caption = sMonN
      txtME.Visible = False
      txtAbonoE.Visible = False
      txtAbonoE.Visible = False
      txtSaldoE.Visible = False
      txtAnteriorE.Visible = False
      txtTC.Caption = "0.000"
      nTC = 1
      '---CESAR tipo cambio
      txtTCO.Caption = "0.000"
      nTCO = 1
      
   End If
   txtUsuario.Caption = sUsuario
   
    'TIPO CAMBIO
    If pais = "002" Then
        Label2.Visible = False
        Label8.Visible = False
        Label9.Visible = False
        txtTC.Visible = False
        txtTCO.Visible = False
        cmdOpcion(2).Visible = False
        cmdOpcion(5).Visible = False
    Else
        Label2.Visible = True
        Label8.Visible = True
        Label9.Visible = True
        txtTC.Visible = True
        txtTCO.Visible = True
        cmdOpcion(2).Visible = True
        cmdOpcion(5).Visible = True
    End If
 
End Sub

Private Sub Form_Unload(Cancel As Integer)
   Set RsTc = Nothing
   Set RsTurno = Nothing
   Set frmInicio = Nothing
End Sub

Public Sub actualizaDatosSistemaAlmacen()
    Screen.MousePointer = vbHourglass
    aniVideo.Visible = True
    aniVideo.AutoPlay = True
    lblProceso.Visible = True
    If VerificaConexionAlmacenRemoto = True Then
        CnAlmacenRemoto.Execute "sp_ActualizaReceta"
        CargaTablasAlmacenRemoto
    End If
    aniVideo.AutoPlay = False
    aniVideo.Visible = False
    lblProceso.Visible = False
    Screen.MousePointer = vbDefault
End Sub

Function VerificaConexionAlmacenRemoto() As Boolean
    On Error GoTo err:
    If sRutaAlmacenRemoto <> "" And sMDBAlmacenRemoto <> "" Then
        Set CnAlmacenRemoto = New ADODB.Connection
        CnAlmacenRemoto.Provider = "SQLOLEDB"
        CnAlmacenRemoto.CursorLocation = adUseServer
        CnAlmacenRemoto.ConnectionString = "User ID=" & sUserName & _
                ";password=" & sUserPassword & _
                ";Data Source=" & sRutaAlmacenRemoto & _
                ";Initial Catalog=" & sMDBAlmacenRemoto
        CnAlmacenRemoto.CommandTimeout = 0
        CnAlmacenRemoto.Open
            If CnAlmacenRemoto.State Then
                VerificaConexionAlmacenRemoto = True
            Else
                VerificaConexionAlmacenRemoto = False
            End If
    End If
    Exit Function
err:
    VerificaConexionAlmacenRemoto = False
End Function

 
