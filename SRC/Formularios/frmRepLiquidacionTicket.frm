VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "msdatlst.ocx"
Begin VB.Form frmRepLiquidacionTicket 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Liquidaci�n de Cajero por Ticketera"
   ClientHeight    =   8280
   ClientLeft      =   1350
   ClientTop       =   885
   ClientWidth     =   7110
   Icon            =   "frmRepLiquidacionTicket.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8280
   ScaleWidth      =   7110
   ShowInTaskbar   =   0   'False
   Begin VB.Frame Frame1 
      Caption         =   " Opciones "
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   7575
      Left            =   0
      TabIndex        =   23
      Top             =   0
      Width           =   7065
      Begin VB.Frame Frame13 
         Caption         =   " Opciones de Resumenes "
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   5040
         Left            =   135
         TabIndex        =   31
         Top             =   2430
         Width           =   4815
         Begin VB.CheckBox chkPedidos 
            Caption         =   " Resumen Pedidos cancelados en otro rango"
            Height          =   240
            Left            =   1170
            TabIndex        =   20
            Top             =   4590
            Width           =   3525
         End
         Begin VB.CheckBox chkCortesia 
            Caption         =   " Resumen Documentos por Cortesia"
            Height          =   240
            Left            =   1170
            TabIndex        =   7
            Top             =   658
            Width           =   3390
         End
         Begin VB.CheckBox chkNoCortesia 
            Caption         =   "Visualizar"
            CausesValidation=   0   'False
            Height          =   240
            Left            =   135
            TabIndex        =   32
            Top             =   675
            Width           =   1005
         End
         Begin VB.CheckBox chkPagoCortes�a 
            Caption         =   " Resumen Pagos con Cortes�a"
            Height          =   285
            Left            =   1170
            TabIndex        =   19
            Top             =   4234
            Width           =   3525
         End
         Begin VB.CheckBox chkGenerado 
            Caption         =   " Resumen Documentos Generados"
            Height          =   240
            Left            =   1170
            TabIndex        =   6
            Top             =   360
            Width           =   3615
         End
         Begin VB.CheckBox chkAnulado 
            Caption         =   " Resumen Documentos Anulados"
            Height          =   240
            Left            =   1170
            TabIndex        =   9
            Top             =   1254
            Width           =   3615
         End
         Begin VB.CheckBox chkCobrar 
            Caption         =   " Resumen Documentos por Cobrar"
            Height          =   240
            Left            =   1170
            TabIndex        =   8
            Top             =   956
            Width           =   3615
         End
         Begin VB.CheckBox chkIngreso 
            Caption         =   " Resumen Recibos de Ingreso"
            Height          =   240
            Left            =   1170
            TabIndex        =   11
            Top             =   1850
            Width           =   3615
         End
         Begin VB.CheckBox chkAnticipo 
            Caption         =   " Resumen Recibos por Anticipo"
            Height          =   240
            Left            =   1170
            TabIndex        =   12
            Top             =   2148
            Width           =   3615
         End
         Begin VB.CheckBox chkEgreso 
            Caption         =   " Resumen Recibos de Egreso"
            Height          =   240
            Left            =   1170
            TabIndex        =   13
            Top             =   2446
            Width           =   3570
         End
         Begin VB.CheckBox chkCuenta 
            Caption         =   " Resumen Cuentas Corrientes"
            Height          =   240
            Left            =   1170
            TabIndex        =   10
            Top             =   1552
            Width           =   3615
         End
         Begin VB.CheckBox chkEfectivo 
            Caption         =   " Resumen Pagos con Efectivo"
            Height          =   240
            Left            =   1170
            TabIndex        =   14
            Top             =   2744
            Width           =   3570
         End
         Begin VB.CheckBox chkTarjeta 
            Caption         =   " Resumen Pagos con Tarjeta de Cr�dito"
            Height          =   240
            Left            =   1170
            TabIndex        =   15
            Top             =   3042
            Width           =   3570
         End
         Begin VB.CheckBox chkCheque 
            Caption         =   " Resumen Pagos con Cheque / Dep�sito"
            Height          =   240
            Left            =   1170
            TabIndex        =   16
            Top             =   3340
            Width           =   3570
         End
         Begin VB.CheckBox chkPunto 
            Caption         =   " Resumen Pagos con Puntos"
            Height          =   240
            Left            =   1170
            TabIndex        =   18
            Top             =   3936
            Width           =   3570
         End
         Begin VB.CheckBox chkOtro 
            Caption         =   " Resumen Otros Tipos de Pago"
            Height          =   240
            Left            =   1170
            TabIndex        =   17
            Top             =   3638
            Width           =   3570
         End
      End
      Begin VB.CheckBox chkTurno 
         Caption         =   "Todos los Turnos"
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
         Left            =   4995
         TabIndex        =   26
         Top             =   300
         Width           =   1875
      End
      Begin VB.CheckBox chkUsuario 
         Caption         =   "Todos los Usuarios"
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
         Left            =   4995
         TabIndex        =   25
         Top             =   693
         Value           =   1  'Checked
         Width           =   1995
      End
      Begin VB.CommandButton cmdBusca 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   0
         Left            =   3270
         Picture         =   "frmRepLiquidacionTicket.frx":000C
         Style           =   1  'Graphical
         TabIndex        =   0
         Top             =   255
         Width           =   765
      End
      Begin VB.TextBox txtTurno 
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
         Height          =   315
         Left            =   1350
         Locked          =   -1  'True
         TabIndex        =   24
         Top             =   285
         Width           =   1815
      End
      Begin MSComCtl2.DTPicker dtpHorIni 
         Height          =   375
         Left            =   3270
         TabIndex        =   3
         Top             =   1071
         Width           =   1635
         _ExtentX        =   2884
         _ExtentY        =   661
         _Version        =   393216
         Enabled         =   0   'False
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         CustomFormat    =   "HH:mm"
         Format          =   88670211
         UpDown          =   -1  'True
         CurrentDate     =   37539
      End
      Begin MSComCtl2.DTPicker dtpHorFin 
         Height          =   375
         Left            =   3270
         TabIndex        =   5
         Top             =   1524
         Width           =   1635
         _ExtentX        =   2884
         _ExtentY        =   661
         _Version        =   393216
         Enabled         =   0   'False
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         CustomFormat    =   "HH:mm"
         Format          =   88670211
         UpDown          =   -1  'True
         CurrentDate     =   37541.9993055556
      End
      Begin MSComCtl2.DTPicker dtpFecFin 
         Height          =   375
         Left            =   1350
         TabIndex        =   4
         Top             =   1524
         Width           =   1815
         _ExtentX        =   3201
         _ExtentY        =   661
         _Version        =   393216
         Enabled         =   0   'False
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         CustomFormat    =   "dd/MM/yyyy hh:mm"
         Format          =   88670209
         CurrentDate     =   37541.9993055556
      End
      Begin MSComCtl2.DTPicker dtpFecIni 
         Height          =   375
         Left            =   1350
         TabIndex        =   2
         Top             =   1071
         Width           =   1815
         _ExtentX        =   3201
         _ExtentY        =   661
         _Version        =   393216
         Enabled         =   0   'False
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         CustomFormat    =   "dd/MM/yyyy hh:mm"
         Format          =   88670209
         CurrentDate     =   37539.2083333333
      End
      Begin MSDataListLib.DataCombo cboUsuario 
         Height          =   315
         Left            =   1350
         TabIndex        =   1
         Top             =   678
         Width           =   1815
         _ExtentX        =   3201
         _ExtentY        =   556
         _Version        =   393216
         MatchEntry      =   -1  'True
         Style           =   2
         BackColor       =   16777215
         ListField       =   ""
         BoundColumn     =   ""
         Text            =   ""
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin MSDataListLib.DataCombo cboImpresora 
         Height          =   360
         Left            =   1350
         TabIndex        =   33
         Top             =   1980
         Width           =   2790
         _ExtentX        =   4921
         _ExtentY        =   635
         _Version        =   393216
         MatchEntry      =   -1  'True
         Style           =   2
         BackColor       =   16777215
         ListField       =   ""
         BoundColumn     =   ""
         Text            =   ""
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin VB.Label Label 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Impresora :"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   195
         Index           =   7
         Left            =   360
         TabIndex        =   34
         Top             =   2063
         Width           =   960
      End
      Begin VB.Label Label 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Fecha Final :"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   195
         Index           =   1
         Left            =   195
         TabIndex        =   30
         Top             =   1614
         Width           =   1125
      End
      Begin VB.Label Label 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Fecha Inicial :"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   195
         Index           =   0
         Left            =   90
         TabIndex        =   29
         Top             =   1161
         Width           =   1230
      End
      Begin VB.Label Label 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Turno :"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   195
         Index           =   2
         Left            =   690
         TabIndex        =   28
         Top             =   345
         Width           =   630
      End
      Begin VB.Label Label 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Usuario :"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   195
         Index           =   3
         Left            =   540
         TabIndex        =   27
         Top             =   738
         Width           =   780
      End
   End
   Begin VB.CommandButton cmdOpcion 
      Caption         =   "Salir"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Index           =   0
      Left            =   5595
      Picture         =   "frmRepLiquidacionTicket.frx":010E
      Style           =   1  'Graphical
      TabIndex        =   22
      Top             =   7620
      Width           =   1455
   End
   Begin VB.CommandButton cmdOpcion 
      Caption         =   "Ticketera"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Index           =   1
      Left            =   4050
      Picture         =   "frmRepLiquidacionTicket.frx":0200
      Style           =   1  'Graphical
      TabIndex        =   21
      Top             =   7620
      Width           =   1455
   End
End
Attribute VB_Name = "frmRepLiquidacionTicket"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim rsReporte As Recordset
Dim RsSumas As Recordset
Dim RsUsuario As Recordset
Dim RsImpresora As Recordset
Dim RsTc As Recordset
Dim RsTurno As Recordset

Dim sNombre   As String
Dim sSuma     As String
Dim sCriterio As String
Dim sCriterioNC As String
Dim sBase     As String
Dim xBase     As String
Dim yBase     As String
Dim sTitulo1  As String
Dim sTitulo2  As String
Dim sTitulo3  As String
Dim sTitulo4  As String

Dim sFechaFin As String
Dim nFondoN As Double
Dim nFondoE As Double
Dim i As Integer
Dim nCambio As Double
Dim xFecha As Date
Dim xParte1 As String
Dim xParte2 As String
Dim xParte3 As String
Dim xParte4 As String
Dim nTotal As Double
Dim nVenta1 As Double
Dim nVenta2 As Double
Dim sLinea As String

Dim nDocumento As Double
Dim nEfectivoN As Double
Dim nEfectivoE As Double
Dim xEfectivoE As Double
Dim nIngresoN As Double
Dim nIngresoE As Double
Dim xIngresoE As Double
Dim nIngresoAN As Double
Dim nIngresoAE As Double
Dim xIngresoAE As Double
Dim nEgresoN As Double
Dim nEgresoE As Double
Dim xEgresoE As Double
Dim nCortesia As Double

' notas de credito
Dim nNotaCreditoN As Double
Dim nNotaCreditoE As Double
Dim xNotaCreditoE As Double
  
Dim tOtro(20) As String
Dim nOtro(20) As Double
Dim xOtro(20) As Double
Dim nDiferencia(20) As Double

Dim nTotalDiferencia As Double
Dim tTarjeta(8) As String
Dim nTarjeta(8) As Double
Dim nPropina(8) As Double
Dim nRecibo(8) As Double
Dim nTarjetaT As Double
Dim nPropinaT As Double
Dim nReciboT As Double

Dim nChequeN As Double
Dim nChequeE As Double
Dim xChequeE As Double
Dim nOtroN As Double
Dim nOtroE As Double
Dim xOtroE As Double
Dim nPunto As Double

Dim nTotalCobrado As Double
Dim nTotalCobradoN As Double
Dim nTotalCobradoE As Double
Dim xTotalCobradoE As Double

Dim nTotalEfectivo As Double
Dim nTotalEfectivoN As Double
Dim nTotalEfectivoE As Double
Dim xTotalEfectivoE As Double
Dim nFisEfectivo As Double
Dim nFisEfectivoN As Double

Dim nAdulto As Double
Dim nNeto As Double
Dim nImpuesto1 As Double
Dim nImpuesto2 As Double
Dim nImpuesto3 As Double
Dim nTotalDoc As Double
Dim nOtroDoc As Double

Dim nTotalCierre As Double
Dim nCobrar As Double
Dim nNoCobrado As Double
Dim Doc(50) As String
Dim Min(50) As String
Dim Max(50) As String
Dim Emi(50) As Integer
Dim Anu(50) As Integer
Dim nDolar As Double
Dim xDolar As Double

Dim nValorVenta(6) As Double
Dim nImp1(6) As Double
Dim nImp2(6) As Double
Dim nImp3(6) As Double
Dim nVenta(6) As Double
Dim nTotalPromedio(6) As Double
Dim sQuiebre(100) As String
Dim nCantidad As Integer
Dim nTotalCorrelativo As Integer

Dim sTurnoConsulta As String


Sub LlenaCombos()
    With cboUsuario
         Isql = "Select * from TUSUARIO where lActivo = 1 order by tCodigoUsuario"
         Set RsUsuario = Lib.OpenRecordset(Isql, Cn)
         Set .RowSource = RsUsuario
         .DataField = "tResumido"
         .ListField = "tResumido"
         .BoundColumn = "tCodigoUsuario"
    End With

    With cboImpresora
         Isql = "Select * from TIMPRESORA Where tCaja = '" & sCaja & "'"
         Set RsImpresora = Lib.OpenRecordset(Isql, Cn)
         Set .RowSource = RsImpresora
             .DataField = "tDescripcion"
             .ListField = "tDescripcion"
             .BoundColumn = "tImpresora"
    End With
End Sub


Private Sub chkNoCortesia_Click()
   If chkNoCortesia.value Then
      chkCortesia.Enabled = True
   Else
      chkCortesia.Enabled = False
   End If
End Sub

Private Sub cmdBusca_Click(Index As Integer)
      Isql = "Select tTurno as Codigo, tCaja, convert(nVarChar,fInicial, 120) as Descripcion, convert(nvarchar,fFinal,120) as fFinal from MTURNO order by tTurno DESC"
      frmBusca.nPredeterm = 2
      Call ConfGrilla(4, frmBusca.grdGrilla, "Turno", 2, "Codigo", 1400, 0, 0, "", _
                                             "Caja", 2, "tCaja", 1100, 0, 0, "", _
                                             "Fecha Inicial", 2, "Descripcion", 2400, 0, 0, "", _
                                             "Fecha Final", 2, "fFinal", 2400, 0, 0, "")
      frmBusca.Show vbModal
      If Not wEnter Then
         Exit Sub
      End If
      sTurnoConsulta = scodigo
      txtTurno.Text = scodigo
End Sub

Private Sub chkTurno_Click()
   If chkTurno.value = 1 Then
      cmdBusca(0).Enabled = False
      txtTurno.Text = ""
      sTurnoConsulta = ""
      txtTurno.Text = ""
      dtpFecIni.Enabled = True
      dtpFecFin.Enabled = True
      dtpHorIni.Enabled = True
      dtpHorFin.Enabled = True
   Else
      cmdBusca(0).Enabled = True
      dtpFecIni.Enabled = False
      dtpFecFin.Enabled = False
      dtpHorIni.Enabled = False
      dtpHorFin.Enabled = False
   End If
End Sub

Private Sub chkUsuario_Click()
   If chkUsuario.value = 1 Then
      cboUsuario.Enabled = False
      cboUsuario.Text = ""
   Else
      cboUsuario.Enabled = True
   End If
End Sub

Private Sub cmdOpcion_Click(Index As Integer)
   If Index = 0 Then
      Unload Me
      Exit Sub
   End If
      
   sTitulo1 = ""
   sTitulo2 = ""
   sTitulo3 = ""
   sTitulo4 = ""
    sTurnoConsulta = Trim(txtTurno.Text)
    
   If cboImpresora.Text = "" Then
      MsgBox "Escoja una impresora", vbCritical, sMensaje
      Exit Sub
   End If
      
   'Fecha del Turno
   If chkTurno.value = 0 Then
      xFecha = Calcular("select fInicial as codigo from MTURNO where tTurno='" & sTurnoConsulta & "'", Cn)
   Else
      xFecha = dtpFecIni.value
   End If
      
   sCriterio = "isnull(tCaja,'0')<>'0'"
   sCriterioNC = "isnull(DBO.MNOTACREDITO.tCaja,'0')<>'0'"
   sBase = "isnull(MDOCUMENTO.tCaja,'0')<>'0'"
   xBase = "isnull(DPAGODOCUMENTO.tDocumento,'0')<>'0'"
   yBase = "isnull(MPEDIDO.tCaja,'0')<>'0'"
      
      
   If chkTurno.value = 0 Then
      If sTurnoConsulta = "" Then
         MsgBox "No se ha asignado el Turno", vbExclamation, sMensaje
         Exit Sub
      End If
      
      Set RsTurno = Lib.OpenRecordset("select * from MTURNO where tTurno='" & sTurnoConsulta & "'", Cn)
      If Not RsTurno.EOF Then
         nFondoN = IIf(IsNull(RsTurno!nMontoIN), 0, RsTurno!nMontoIN)
         nFondoE = IIf(IsNull(RsTurno!nMontoIE), 0, RsTurno!nMontoIE)
         sFechaFin = IIf(IsNull(RsTurno!fFinal), "", RsTurno!fFinal)
      Else
         nFondoN = 0
         nFondoE = 0
         sFechaFin = ""
      End If
      
      sTitulo1 = "Turno  : " & sTurnoConsulta
      sTitulo2 = "Rango  : Del " & Format(xFecha, "") & " Hr"
      sTitulo3 = "         Al " & Format(sFechaFin) & " Hr"
            
      sCriterio = sCriterio & " And (tTurno ='" & sTurnoConsulta & "') "
      sCriterioNC = sCriterioNC & " And (DBO.MNOTACREDITO.tTurno ='" & sTurnoConsulta & "') "
      sBase = sBase & " And MDOCUMENTO.tTurno ='" & sTurnoConsulta & "'"
      xBase = xBase & " And DPAGODOCUMENTO.tTurno ='" & sTurnoConsulta & "'"
      yBase = yBase & " And MPEDIDO.tTurno ='" & sTurnoConsulta & "'"
   Else
      If dtpFecIni.value + dtpHorIni.value > dtpFecFin.value + dtpHorFin.value Then
          MsgBox "Error en Rango de Fechas", vbCritical, sMensaje
          Exit Sub
          dtpFecFin.SetFocus
      End If
      nFondoN = 0
      nFondoE = 0
      sFechaFin = dtpFecFin.value
      
      sTitulo1 = "Turno   : Todos los Turnos"
      sTitulo2 = "Rango   : Del " & dtpFecIni.value & " " & Format(dtpHorIni.value, "HH:nn") & " Hrs "
      sTitulo3 = "          Al " & dtpFecFin.value & " " & Format(dtpHorFin.value, "HH:nn") & " Hrs "
      sCriterio = sCriterio & " And fRegistro >= '" & Format(dtpFecIni.value, "yyyy/MM/dd") & " " & Format(dtpHorIni.value, "hh:mm") & "' And fRegistro <= '" & Format(dtpFecFin.value, "yyyy/MM/dd") & " " & Format(dtpHorFin.value, "HH:nn") & "'"
      sCriterioNC = sCriterioNC & " And DBO.MNOTACREDITO.fRegistro >= '" & Format(dtpFecIni.value, "yyyy/MM/dd") & " " & Format(dtpHorIni.value, "hh:mm") & "' And DBO.MNOTACREDITO.fRegistro <= '" & Format(dtpFecFin.value, "yyyy/MM/dd") & " " & Format(dtpHorFin.value, "HH:nn") & "'"
      sBase = sBase & " And ((MDOCUMENTO.fRegistro >= '" & Format(dtpFecIni.value, "yyyy/MM/dd") & " " & Format(dtpHorIni.value, "HH:nn") & "' And MDOCUMENTO.fRegistro <= '" & Format(dtpFecFin.value, "yyyy/MM/dd") & " " & Format(dtpHorFin.value, "HH:nn") & "')) "
      xBase = xBase & " And ((DPAGODOCUMENTO.fRegistro >= '" & Format(dtpFecIni.value, "yyyy/MM/dd") & " " & Format(dtpHorIni.value, "HH:nn") & "' And DPAGODOCUMENTO.fRegistro <= '" & Format(dtpFecFin.value, "yyyy/MM/dd") & " " & Format(dtpHorFin.value, "HH:nn") & "')) "
      yBase = yBase & " And MPEDIDO.tTurno ='" & sTurnoConsulta & "'"
   End If
   
   If chkUsuario.value = 0 Then
      If cboUsuario.Text = "" Then
         MsgBox "No se ha asignado el Usuario", vbExclamation, sMensaje
         Exit Sub
      End If
      sCriterio = sCriterio & " And tUsuario ='" & cboUsuario.Text & "'"
      sCriterioNC = sCriterioNC & " And DBO.MNOTACREDITO.tUsuario ='" & cboUsuario.Text & "'"
      sBase = sBase & " And MDOCUMENTO.tUsuario ='" & cboUsuario.Text & "'"
      xBase = xBase & " And DPAGODOCUMENTO.tUsuario ='" & cboUsuario.Text & "'"
      yBase = yBase & " And MPEDIDO.tUsuario ='" & cboUsuario.Text & "'"
      sTitulo4 = "Usuario: " & cboUsuario.Text
   Else
      sTitulo4 = "Usuario: Todos los Usuarios"
   End If

   Genera
   
   If rsReporte.EOF = True Then
      Screen.MousePointer = vbDefault
      MsgBox "No hay Datos para Mostrar", vbExclamation, "Mensaje del Sistema"
      Exit Sub
   End If
   
   Set RsSumas = Lib.OpenRecordset("select Distinct tGrupo, tSubGrupo From " & sNombre & " order by tgrupo,tsubgrupo", Cn)
   nCantidad = RsSumas.RecordCount
   If nCantidad > 0 Then
      RsSumas.MoveFirst
      For i = 0 To nCantidad - 1
          sQuiebre(i) = RsSumas!tgrupo + RsSumas!tSubGrupo
          RsSumas.MoveNext
      Next i
   End If
   
'    sQuiebre(0) = "00"   'Emitidos
'    sQuiebre(1) = "01"   'Efectivo
'    sQuiebre(2) = "02"   'Tarjeta
'    sQuiebre(3) = "03"   'Cheques
'    sQuiebre(4) = "04"   'Otros
'    sQuiebre(5) = "05"   'Puntos
'    sQuiebre(6) = "06"   'Cuentas por Cobrar
'    sQuiebre(7) = "07"   'Recibos Ingreso
'    sQuiebre(8) = "08"   'Recibos Anticipos
'    sQuiebre(9) = "09"   'Egreso
'   sQuiebre(10) = "10"   'Cortesias No va
'   sQuiebre(11) = "11"   'Cuentas Corrientes
'   sQuiebre(12) = "12"   'Anulados
        
   'Configura la impresora la impresion Font
   Imprimir (cboImpresora.BoundText)
   Printer.FontName = sFont
   Printer.FontBold = False
   
   'Cabecera
   ImprimeXCentro "Liquidacion de Cajero", 40
   ImprimeXCentro sRazonSocial, 40
   Printer.Print ""
   
   Printer.Print sTitulo1
   If lMCPV Then
      Printer.Print "Usuario: " & sUsuario
   Else
      Printer.Print sTitulo2
      Printer.Print sTitulo3
      Printer.Print sTitulo4
   End If
   Printer.Print ""
   Printer.Print "Fecha  : " & Date & " " & Time
   Printer.Print String(40, "-")
         
   'Detalle
   rsReporte.MoveFirst
   
   nTotal = 0
   nVenta1 = 0
   nVenta2 = 0
   
   For i = 0 To nCantidad - 1
       rsReporte.Filter = "tGrupo = '" & Mid(sQuiebre(i), 1, 2) & "' and tSubGrupo='" & Mid(sQuiebre(i), 3, 3) & "'"
       If Not rsReporte.EOF Then
          Cabecera rsReporte!tgrupo, rsReporte!tSubGrupo
          rsReporte.MoveFirst
          Do While Not rsReporte.EOF
             nTotal = nTotal + 1
             nVenta1 = nVenta1 + IIf(IsNull(rsReporte!nVenta1), 0, rsReporte!nVenta1)
             nVenta2 = nVenta2 + IIf(IsNull(rsReporte!nVenta2), 0, rsReporte!nVenta2)

             Detalle rsReporte!tgrupo, rsReporte!tSubGrupo
             rsReporte.MoveNext
          Loop

          rsReporte.MoveLast
          Final rsReporte!tgrupo, rsReporte!tSubGrupo

          nTotal = 0
          nVenta1 = 0
          nVenta2 = 0
       End If
   Next i
   
   Resumen
   Printer.EndDoc
   Screen.MousePointer = vbDefault
   
   If sModulo = "INFOREST" Then
      Unload Me
   End If
End Sub

Private Sub dtpFecfin_LostFocus()
   If dtpFecIni.value > dtpFecFin.value Then
      MsgBox "Error en Rango de Fechas", vbCritical, sMensaje
      dtpFecFin.SetFocus
   End If
End Sub

Private Sub Form_Load()
   sNombre = dbTemporal(sCaja, 13, "tGrupo", "nVarChar(2)", _
                                   "Grupo", "nVarChar(50)", _
                                   "tSubGrupo", "nVarChar(3)", _
                                   "SubGrupo", "nVarChar(50)", _
                                   "tDocumento", "nVarChar(20)", _
                                   "tUsuario", "nVarChar(15)", _
                                   "fFecha", "nVarChar(35)", _
                                   "fPago", "nVarChar(35)", _
                                   "nTC", "Float", _
                                   "nVenta", "Float", _
                                   "nVenta1", "Float", _
                                   "nVenta2", "Float", _
                                   "tObservacion", "nVarChar(250)")

   Centrar Me
   LlenaCombos
   cboUsuario.Enabled = False
   dtpFecIni.value = Date
   dtpFecFin.value = Date
   dtpHorIni.value = "12:00 am"
   dtpHorFin.value = Time
   sSuma = dbTemporal(sCaja, 8, "tDocumento", "nVarChar(20)", _
                                 "nNeto", "float", _
                                 "nImpuesto1", "float", _
                                 "nImpuesto2", "float", _
                                 "nImpuesto3", "float", _
                                 "nVenta", "float", _
                                 "tTipoPedido", "nVarChar(2)", "RV", "NVARCHAR(1)")

   cboImpresora.BoundText = sPreCuenta
End Sub

Private Sub Form_Unload(Cancel As Integer)
   Cn.Execute "Drop Table " & sNombre
   Cn.Execute "Drop Table " & sSuma

   Set RsSumas = Nothing
   Set RsTc = Nothing
   Set rsReporte = Nothing
   Set RsUsuario = Nothing
   Set RsTurno = Nothing
   Set frmRepLiquidacionTicket = Nothing
End Sub

Public Sub Genera()
   Screen.MousePointer = vbHourglass
   
   'Turnos
   If chkTurno.value = 0 Then
      Set RsTurno = Lib.OpenRecordset("select * from MTURNO where tTurno='" & sTurnoConsulta & "'", Cn)
      If Not RsTurno.EOF Then
         nFondoN = IIf(IsNull(RsTurno!nMontoIN), 0, RsTurno!nMontoIN)
         nFondoE = IIf(IsNull(RsTurno!nMontoIE), 0, RsTurno!nMontoIE)
         sFechaFin = IIf(IsNull(RsTurno!fFinal), "", RsTurno!fFinal)
       Else
         nFondoN = 0
         nFondoE = 0
         sFechaFin = ""
      End If
   Else
      nFondoN = 0
      nFondoE = 0
   End If
   
   'Tipo de Cambio del turno
   nCambio = Calcular("select nVenta as Codigo from TTIPOCAMBIO where fFecha='" & Format(xFecha, "yyyy/mm/dd") & "'", Cn)
   
   'Blanqueo la tabla
   Cn.Execute "delete from " & sNombre
      
   'Inserta los documentos generados
   If chkGenerado.value Then
      Isql = "Insert Into " & sNombre & " (tGrupo, Grupo, tSubGrupo, SubGrupo, tDocumento, tUsuario, fFecha, fPago, nVenta, nVenta1, nVenta2, tObservacion ) " & _
             "SELECT '00' AS tGrupo, 'Documentos Generados' AS Grupo, '00' AS tSubGrupo, '' AS SubGrupo, '' AS tDocumento, '' AS tUsuario, '' AS fFecha, '' AS fPago, SUM(nVenta) AS nVenta, SUM(CASE WHEN tTipoDocumento = '00' THEN 0 WHEN tEstadoDocumento = '04' THEN 0 ELSE nventa END) As nVenta1, 0 As nVenta2, Count(nVenta) As tObservacion " & _
             "From dbo.MDOCUMENTO " & _
             "WHERE " & sBase & IIf(chkNoCortesia.value, "", " and tTipoDocumento <> '00'")
      Cn.Execute Isql
      Cn.Execute "delete from " & sNombre & " where tGrupo='00' and nVenta=0"
      
        If lNcOfisis Then
             Isql = "Insert Into " & sNombre & " (tGrupo, Grupo, tSubGrupo, SubGrupo, tDocumento, tUsuario, fFecha, fPago,  nVenta, nVenta1,nVenta2, tObservacion ) " & _
                 " SELECT '00' AS tGrupo, 'Documentos Generados' AS Grupo, '00' AS tSubGrupo, '' AS SubGrupo, '' AS tDocumento, '' AS tUsuario,  '' AS fFecha, '' AS fPago ,sum((dbo.MNOTACREDITO.nVenta*(-1))) As nVenta,sum((dbo.MNOTACREDITO.nVenta*(-1))) as nVenta1,0 as nVenta2, count(nventa) AS tObservacion " & _
                 " from dbo.MNOTACREDITO " & _
                 " WHERE dbo.MNOTACREDITO.tEstadoDocumento in ('05','02') and  " & sCriterio
                 '" WHERE dbo.MNOTACREDITO.tEstadoDocumento='05' and  " & sCriterio

             Cn.Execute Isql
        End If

   Else
      Isql = "Insert Into " & sNombre & " (tGrupo, Grupo, tSubGrupo, SubGrupo, tDocumento, tUsuario, fFecha, fPago, nVenta, nVenta1, tObservacion ) " & _
             "SELECT Distinct '00' AS tGrupo, 'Documentos Generados' AS Grupo, '00' AS tSubGrupo, '' AS SubGrupo, dbo.MDOCUMENTO.tDocumento, dbo.MDOCUMENTO.tUsuarioAutoriza AS tUsuario, " & _
             "CONVERT(nvarchar(8), dbo.MDOCUMENTO.fRegistro, 3) + ' ' + CONVERT(nvarchar(5), dbo.MDOCUMENTO.fRegistro, 8) AS fRegistro, CONVERT(nvarchar(8), dbo.MDOCUMENTO.fPago, 3) + ' ' + CONVERT(nvarchar(5), dbo.MDOCUMENTO.fPago, 8) AS fPago, " & _
             "dbo.MDOCUMENTO.nVenta AS nVenta, CASE WHEN tTipoDocumento = '00' THEN 0 WHEN tEstadoDocumento = '04' THEN 0 ELSE nventa END As nVenta1,CASE WHEN MDOCUMENTO.TESTADODOCUMENTO <>'04' THEN  case when isnull(tclientepago,' ')<>' ' then 'POR COBRAR ' ELSE dbo.vEstadoDocumento.Descripcion +  CASE WHEN isnull(dbo.dPagoDocumento.tTurno, '')='' then ''else ' - ' END + ISNULL(dbo.DPAGODOCUMENTO.tTurno, '') +   ' ' + ISNULL(dbo.vCortesia.Descripcion, '') END ELSE   case when isnull(tclientepago,' ')<>' ' then  'POR COBRAR '  ELSE dbo.vEstadoDocumento.Descripcion +  CASE WHEN isnull(dbo.dPagoDocumento.tTurno,'')='' then '' else ' - ' END + ISNULL(dbo.DPAGODOCUMENTO.tTurno,  '' ) + ' ' + ISNULL(dbo.vCortesia.Descripcion , '') END END  AS tObservacion " & _
             "FROM dbo.MDOCUMENTO LEFT OUTER JOIN dbo.vCortesia ON dbo.MDOCUMENTO.tCortesia = dbo.vCortesia.Codigo LEFT OUTER JOIN dbo.vEstadoDocumento ON dbo.MDOCUMENTO.tEstadoDocumento = dbo.vEstadoDocumento.Codigo LEFT OUTER JOIN dbo.DPAGODOCUMENTO ON dbo.MDOCUMENTO.tDocumento = dbo.DPAGODOCUMENTO.tDocumento " & _
             "WHERE " & sBase & IIf(chkNoCortesia.value, "", " and tTipoDocumento <> '00'")
      Cn.Execute Isql
      
      If lNcOfisis Then
              Isql = "Insert Into " & sNombre & " (tGrupo, Grupo, tSubGrupo, SubGrupo, tDocumento, tUsuario, fFecha, fPago,  nVenta, nVenta1, tObservacion ) " & _
               " SELECT Distinct '00' AS tGrupo, 'Documentos Generados' AS Grupo, '00' AS tSubGrupo, '' AS SubGrupo, dbo.MNOTACREDITO.tNotaCredito, dbo.MNOTACREDITO.tUsuario AS tUsuario,  " & _
               " CONVERT(nvarchar(8), dbo.MNOTACREDITO.fRegistro, 3) + ' ' + CONVERT(nvarchar(5), dbo.MNOTACREDITO.fRegistro, 8) AS fRegistro, " & _
               " CONVERT(nvarchar(8), dbo.MNOTACREDITO.fRegistro, 3) + ' ' + CONVERT(nvarchar(5), dbo.MNOTACREDITO.fRegistro, 8) AS fPago," & _
               "  (dbo.MNOTACREDITO.nVenta*(-1)) As nVenta,(dbo.MNOTACREDITO.nVenta*(-1)) as nVenta1, 'NC. Gen. : ' + dbo.MNOTACREDITO.tTurno AS tObservacion " & _
               " from dbo.MNOTACREDITO " & _
               " WHERE dbo.MNOTACREDITO.tEstadoDocumento in ('05','02') and  " & sCriterio

          Cn.Execute Isql
      End If

   End If
      
   'Inserta Efectivo
   If chkEfectivo.value Then
      Isql = "Insert Into " & sNombre & " (tGrupo, Grupo, tSubGrupo, SubGrupo, tDocumento, tUsuario, fFecha, fPago, nVenta, nVenta1, nVenta2, tObservacion, nTC ) " & _
             "SELECT '01' AS tGrupo, 'Efectivo' AS Grupo, '00' AS tSubGrupo, '' AS SubGrupo, '' AS tDocumento, '' AS tUsuario, '' AS fFecha, '' AS fPago, " & _
             "SUM(nVenta) AS nVenta, " & _
             "SUM(CASE WHEN dbo.DPAGODOCUMENTO.tMoneda = '01' THEN dbo.DPAGODOCUMENTO.nMonto ELSE 0 END) As nVenta1, " & _
             "SUM(CASE WHEN dbo.DPAGODOCUMENTO.tMoneda = '02' THEN dbo.DPAGODOCUMENTO.nMonto ELSE 0 END) As nVenta2, " & _
             "Count(nVenta) As tObservacion, AVG(nTipoCambio) " & _
             "FROM dbo.MDOCUMENTO LEFT OUTER JOIN dbo.DPAGODOCUMENTO ON dbo.MDOCUMENTO.tDocumento = dbo.DPAGODOCUMENTO.tDocumento " & _
             "WHERE dbo.DPAGODOCUMENTO.tTipoPago='01' and " & xBase
      Cn.Execute Isql
      Cn.Execute "delete from " & sNombre & " where tGrupo='01' and nVenta=0"

   Else
      Isql = "Insert Into " & sNombre & " (tGrupo, Grupo, tSubGrupo, SubGrupo, tDocumento, tUsuario, fFecha, fPago, nTC, nVenta, nVenta1, nVenta2, tObservacion ) " & _
             "SELECT '01' AS tGrupo, 'Efectivo' AS Grupo, '00' AS tSubGrupo, '' AS SubGrupo, dbo.MDOCUMENTO.tDocumento, " & _
             "dbo.DPAGODOCUMENTO.tUsuario AS tUsuario, CONVERT(nvarchar(8), dbo.MDOCUMENTO.fRegistro, 3) + ' ' + CONVERT(nvarchar(5), dbo.MDOCUMENTO.fRegistro, 8) AS fRegistro, " & _
             "CONVERT(nvarchar(8), dbo.DPAGODOCUMENTO.fRegistro, 3) + ' ' + CONVERT(nvarchar(5), dbo.DPAGODOCUMENTO.fRegistro, 8) AS fPago, dbo.DPAGODOCUMENTO.nTipoCambio, " & _
             "dbo.MDOCUMENTO.nVenta As nVenta, " & _
             "CASE WHEN dbo.DPAGODOCUMENTO.tMoneda = '01' THEN dbo.DPAGODOCUMENTO.nMonto ELSE 0 END As nVenta1, " & _
             "CASE WHEN dbo.DPAGODOCUMENTO.tMoneda = '02' THEN dbo.DPAGODOCUMENTO.nMonto ELSE 0 END As nVenta2, " & _
             "'Doc. Gen. : ' + dbo.MDOCUMENTO.tTurno AS tObservacion " & _
             "FROM dbo.MDOCUMENTO LEFT OUTER JOIN dbo.DPAGODOCUMENTO ON dbo.MDOCUMENTO.tDocumento = dbo.DPAGODOCUMENTO.tDocumento " & _
             "WHERE dbo.DPAGODOCUMENTO.tTipoPago='01' and " & xBase
      Cn.Execute Isql
    
   End If
   
   'Inserta Tarjeta Credito
   If chkTarjeta.value Then
      Isql = "Insert Into " & sNombre & " (tGrupo, Grupo, tSubGrupo, SubGrupo, tDocumento, tUsuario, fFecha, fPago, nVenta, nVenta1, nVenta2, tObservacion ) " & _
             "SELECT '02' AS tGrupo, 'Tarjeta Cr�dito' AS Grupo, dbo.DPAGODOCUMENTO.tTarjeta AS tSubGrupo, dbo.TTARJETACREDITO.tResumido AS SubGrupo, '' AS tDocumento, '' AS tUsuario, '' AS fFecha, '' AS fPago, " & _
             "SUM(nVenta) AS nVenta, " & _
             "SUM(dbo.DPAGODOCUMENTO.nMonto) as nVenta1, " & _
             "SUM(dbo.DPAGODOCUMENTO.nPropina) As nVenta2, " & _
             "Count(nVenta) as tObservacion " & _
             "FROM dbo.TTARJETACREDITO RIGHT OUTER JOIN dbo.DPAGODOCUMENTO ON dbo.TTARJETACREDITO.tCodigoTarjeta = dbo.DPAGODOCUMENTO.tTarjeta RIGHT OUTER JOIN dbo.MDOCUMENTO ON dbo.DPAGODOCUMENTO.tDocumento = dbo.MDOCUMENTO.tDocumento " & _
             "WHERE dbo.DPAGODOCUMENTO.tTipoPago='02' and " & xBase & _
             "GROUP BY dbo.DPAGODOCUMENTO.tTarjeta, dbo.TTARJETACREDITO.tResumido"
      Cn.Execute Isql
      Cn.Execute "delete from " & sNombre & " where tGrupo='02' and isnull(nVenta,0)=0"
   Else
      Isql = "Insert Into " & sNombre & " (tGrupo, Grupo, tSubGrupo, SubGrupo, tDocumento, tUsuario, fFecha, fPago, nTC, nVenta, nVenta1, nVenta2, tObservacion ) " & _
             "SELECT '02' AS tGrupo, 'Tarjeta Cr�dito' AS Grupo, dbo.DPAGODOCUMENTO.tTarjeta AS tSubGrupo, dbo.TTARJETACREDITO.tResumido AS SubGrupo, dbo.MDOCUMENTO.tDocumento, " & _
             "dbo.DPAGODOCUMENTO.tUsuario AS tUsuario, CONVERT(nvarchar(8), dbo.MDOCUMENTO.fRegistro, 3) + ' ' + CONVERT(nvarchar(5), dbo.MDOCUMENTO.fRegistro, 8) AS fRegistro, " & _
             "CONVERT(nvarchar(8), dbo.DPAGODOCUMENTO.fRegistro, 3) + ' ' + CONVERT(nvarchar(5), dbo.DPAGODOCUMENTO.fRegistro, 8) AS fPago, dbo.DPAGODOCUMENTO.nTipoCambio, " & _
             "dbo.MDOCUMENTO.nVenta As nVenta, " & _
             "dbo.DPAGODOCUMENTO.nMonto as nVenta1, " & _
             "dbo.DPAGODOCUMENTO.nPropina As nVenta2, " & _
             "'Doc. Gen. : ' + dbo.MDOCUMENTO.tTurno + ' Nro. ' + dbo.DPAGODOCUMENTO.tNumero AS tObservacion " & _
             "FROM dbo.TTARJETACREDITO RIGHT OUTER JOIN dbo.DPAGODOCUMENTO ON dbo.TTARJETACREDITO.tCodigoTarjeta = dbo.DPAGODOCUMENTO.tTarjeta RIGHT OUTER JOIN dbo.MDOCUMENTO ON dbo.DPAGODOCUMENTO.tDocumento = dbo.MDOCUMENTO.tDocumento " & _
             "WHERE dbo.DPAGODOCUMENTO.tTipoPago='02' and " & xBase
      Cn.Execute Isql
   End If
   
   'Inserta Cheque
   If chkCheque.value Then
      Isql = "Insert Into " & sNombre & " (tGrupo, Grupo, tSubGrupo, SubGrupo, tDocumento, tUsuario, fFecha, fPago, nVenta, nVenta1, nVenta2, tObservacion, nTC ) " & _
             "SELECT '03' AS tGrupo, 'Cheque / Deposito' AS Grupo, '00' AS tSubGrupo, '' AS SubGrupo, '' AS tDocumento, '' AS tUsuario, '' AS fFecha, '' AS fPago, " & _
             "SUM(nVenta) AS nVenta, " & _
             "SUM(CASE WHEN dbo.DPAGODOCUMENTO.tMoneda = '01' THEN dbo.DPAGODOCUMENTO.nMonto ELSE 0 END) As nVenta1, " & _
             "SUM(CASE WHEN dbo.DPAGODOCUMENTO.tMoneda = '02' THEN dbo.DPAGODOCUMENTO.nMonto ELSE 0 END) As nVenta2, " & _
             "Count(nVenta) As tObservacion, AVG(nTipoCambio) " & _
             "FROM dbo.MDOCUMENTO LEFT OUTER JOIN dbo.DPAGODOCUMENTO ON dbo.MDOCUMENTO.tDocumento = dbo.DPAGODOCUMENTO.tDocumento " & _
             "WHERE dbo.DPAGODOCUMENTO.tTipoPago='03' and " & xBase
      Cn.Execute Isql
      Cn.Execute "delete from " & sNombre & " where tGrupo='03' and isnull(nVenta,0)=0"
   Else
      Isql = "Insert Into " & sNombre & " (tGrupo, Grupo, tSubGrupo, SubGrupo, tDocumento, tUsuario, fFecha, fPago, nTC, nVenta, nVenta1, nVenta2, tObservacion ) " & _
             "SELECT '03' AS tGrupo, 'Cheque / Deposito' AS Grupo, '00' AS tSubGrupo, '' AS SubGrupo, dbo.MDOCUMENTO.tDocumento, " & _
             "dbo.DPAGODOCUMENTO.tUsuario AS tUsuario, CONVERT(nvarchar(8), dbo.MDOCUMENTO.fRegistro, 3) + ' ' + CONVERT(nvarchar(5), dbo.MDOCUMENTO.fRegistro, 8) AS fRegistro, " & _
             "CONVERT(nvarchar(8), dbo.DPAGODOCUMENTO.fRegistro, 3) + ' ' + CONVERT(nvarchar(5), dbo.DPAGODOCUMENTO.fRegistro, 8) AS fPago, dbo.DPAGODOCUMENTO.nTipoCambio, " & _
             "dbo.MDOCUMENTO.nVenta As nVenta, " & _
             "CASE WHEN dbo.DPAGODOCUMENTO.tMoneda = '01' THEN dbo.DPAGODOCUMENTO.nMonto ELSE 0 END As nVenta1, " & _
             "CASE WHEN dbo.DPAGODOCUMENTO.tMoneda = '02' THEN dbo.DPAGODOCUMENTO.nMonto ELSE 0 END As nVenta2, " & _
             "'Doc. Gen. : ' + dbo.MDOCUMENTO.tTurno + ' Nro. ' + dbo.DPAGODOCUMENTO.tNumero AS tObservacion " & _
             "FROM dbo.MDOCUMENTO LEFT OUTER JOIN dbo.DPAGODOCUMENTO ON dbo.MDOCUMENTO.tDocumento = dbo.DPAGODOCUMENTO.tDocumento " & _
             "WHERE dbo.DPAGODOCUMENTO.tTipoPago='03' and " & xBase
      Cn.Execute Isql
   End If
      
   'Inserta Pagos Varios
   If chkOtro.value Then
      Isql = "Insert Into " & sNombre & " (tGrupo, Grupo, tSubGrupo, SubGrupo, tDocumento, tUsuario, fFecha, fPago, nVenta, nVenta1, nVenta2, tObservacion, nTC ) " & _
             "SELECT '04' AS tGrupo, 'Otros Pagos' AS Grupo, dbo.DPAGODOCUMENTO.tOtroTipoPago AS tSubGrupo, vTipoCancelacion.Descripcion AS SubGrupo, '' AS tDocumento, '' AS tUsuario, '' AS fFecha, '' AS fPago, " & _
             "SUM(nVenta) AS nVenta, " & _
             "SUM(CASE WHEN dbo.DPAGODOCUMENTO.tMoneda = '01' THEN dbo.DPAGODOCUMENTO.nMonto ELSE 0 END) As nVenta1, " & _
             "SUM(CASE WHEN dbo.DPAGODOCUMENTO.tMoneda = '02' THEN dbo.DPAGODOCUMENTO.nMonto ELSE 0 END) As nVenta2, " & _
             "Count(nVenta) As tObservacion, AVG(nTipoCambio) " & _
             "FROM dbo.vTipoCancelacion RIGHT OUTER JOIN dbo.DPAGODOCUMENTO ON dbo.vTipoCancelacion.Codigo = dbo.DPAGODOCUMENTO.tOtroTipoPago RIGHT OUTER JOIN dbo.MDOCUMENTO ON dbo.DPAGODOCUMENTO.tDocumento = dbo.MDOCUMENTO.tDocumento " & _
             "WHERE dbo.DPAGODOCUMENTO.tTipoPago='04' and " & xBase & _
             "GROUP BY dbo.DPAGODOCUMENTO.tOtroTipoPago, dbo.vTipoCancelacion.Descripcion"
      Cn.Execute Isql
      Cn.Execute "delete from " & sNombre & " where tGrupo='04' and isnull(nVenta,0)=0"
   Else
      Isql = "Insert Into " & sNombre & " (tGrupo, Grupo, tSubGrupo, SubGrupo, tDocumento, tUsuario, fFecha, fPago, nTC, nVenta, nVenta1, nVenta2, tObservacion ) " & _
             "SELECT '04' AS tGrupo, 'Otros Pagos' AS Grupo, dbo.DPAGODOCUMENTO.tOtroTipoPago AS tSubGrupo, vTipoCancelacion.Descripcion AS SubGrupo, dbo.MDOCUMENTO.tDocumento, " & _
             "dbo.DPAGODOCUMENTO.tUsuario AS tUsuario, CONVERT(nvarchar(8), dbo.MDOCUMENTO.fRegistro, 3) + ' ' + CONVERT(nvarchar(5), dbo.MDOCUMENTO.fRegistro, 8) AS fRegistro, " & _
             "CONVERT(nvarchar(8), dbo.DPAGODOCUMENTO.fRegistro, 3) + ' ' + CONVERT(nvarchar(5), dbo.DPAGODOCUMENTO.fRegistro, 8) AS fPago, dbo.DPAGODOCUMENTO.nTipoCambio, " & _
             "dbo.MDOCUMENTO.nVenta As nVenta, " & _
             "CASE WHEN dbo.DPAGODOCUMENTO.tMoneda = '01' THEN dbo.DPAGODOCUMENTO.nMonto ELSE 0 END As nVenta1, " & _
             "CASE WHEN dbo.DPAGODOCUMENTO.tMoneda = '02' THEN dbo.DPAGODOCUMENTO.nMonto ELSE 0 END As nVenta2, " & _
             "'Doc. Gen. : ' + dbo.MDOCUMENTO.tTurno + ' ' + CASE WHEN isnull(dbo.dPagoDocumento.tNumero,'')='' then '' else 'Doc: ' END + ltrim(dbo.DPAGODOCUMENTO.tNumero) + ' ' + lTrim(dbo.DPAGODOCUMENTO.tBanco) AS tObservacion  " & _
             "FROM dbo.vTipoCancelacion RIGHT OUTER JOIN dbo.DPAGODOCUMENTO ON dbo.vTipoCancelacion.Codigo = dbo.DPAGODOCUMENTO.tOtroTipoPago RIGHT OUTER JOIN dbo.MDOCUMENTO ON dbo.DPAGODOCUMENTO.tDocumento = dbo.MDOCUMENTO.tDocumento " & _
             "WHERE dbo.DPAGODOCUMENTO.tTipoPago='04' and " & xBase
      Cn.Execute Isql
   End If

   'Inserta Puntos
   If chkPunto.value Then
      Isql = "Insert Into " & sNombre & " (tGrupo, Grupo, tSubGrupo, SubGrupo, tDocumento, tUsuario, fFecha, fPago, nVenta, nVenta1, nVenta2, tObservacion ) " & _
             "SELECT '05' AS tGrupo, '" & sClub & "' as Grupo, '00' AS tSubGrupo, '' AS SubGrupo, '' AS tDocumento, '' AS tUsuario, '' AS fFecha, '' AS fPago, " & _
             "SUM(nVenta) AS nVenta, " & _
             "SUM(CASE WHEN dbo.DPAGODOCUMENTO.tMoneda = '01' THEN dbo.DPAGODOCUMENTO.nMonto ELSE 0 END) As nVenta1, " & _
             "0 As nVenta2, " & _
             "'' As tObservacion " & _
             "FROM dbo.vDelivery RIGHT OUTER JOIN dbo.DPAGODOCUMENTO ON dbo.vDelivery.Codigo = dbo.DPAGODOCUMENTO.tNumero RIGHT OUTER JOIN dbo.MDOCUMENTO ON dbo.DPAGODOCUMENTO.tDocumento = dbo.MDOCUMENTO.tDocumento " & _
             "WHERE dbo.DPAGODOCUMENTO.tTipoPago='05' and " & xBase
      Cn.Execute Isql
      Cn.Execute "delete from " & sNombre & " where tGrupo='05' and isnull(nVenta,0)=0"
   Else
      Isql = "Insert Into " & sNombre & " (tGrupo, Grupo, tSubGrupo, SubGrupo, tDocumento, tUsuario, fFecha, fPago, nTC, nVenta, nVenta1, nVenta2, tObservacion ) " & _
             "SELECT '05' AS tGrupo, '" & sClub & "' AS Grupo, '00' AS tSubGrupo, '' AS SubGrupo, dbo.MDOCUMENTO.tDocumento, " & _
             "dbo.DPAGODOCUMENTO.tUsuario AS tUsuario, CONVERT(nvarchar(8), dbo.MDOCUMENTO.fRegistro, 3) + ' ' + CONVERT(nvarchar(5), dbo.MDOCUMENTO.fRegistro, 8) AS fRegistro, " & _
             "CONVERT(nvarchar(8), dbo.DPAGODOCUMENTO.fRegistro, 3) + ' ' + CONVERT(nvarchar(5), dbo.DPAGODOCUMENTO.fRegistro, 8) AS fPago, dbo.DPAGODOCUMENTO.nTipoCambio, " & _
             "dbo.MDOCUMENTO.nVenta As nVenta, " & _
             "CASE WHEN dbo.DPAGODOCUMENTO.tMoneda = '01' THEN dbo.DPAGODOCUMENTO.nMonto ELSE 0 END As nVenta1, " & _
             "0 As nVenta2, " & _
             "'Doc. Gen. : ' + dbo.MDOCUMENTO.tTurno + ' Cliente: ' + ltrim(dbo.vDelivery.Cliente) AS tObservacion " & _
             "FROM dbo.vDelivery RIGHT OUTER JOIN dbo.DPAGODOCUMENTO ON dbo.vDelivery.Codigo = dbo.DPAGODOCUMENTO.tNumero RIGHT OUTER JOIN dbo.MDOCUMENTO ON dbo.DPAGODOCUMENTO.tDocumento = dbo.MDOCUMENTO.tDocumento " & _
             "WHERE dbo.DPAGODOCUMENTO.tTipoPago='05' and " & xBase
      Cn.Execute Isql
   End If
   
   'Inserta Cancelacion con Cortes�a
   If chkPagoCortes�a.value Then
'      Isql = " Insert Into " & sNombre & " (tGrupo, Grupo, tSubGrupo, SubGrupo, tDocumento, tUsuario, fFecha, fPago, nVenta, nVenta1, nVenta2, tObservacion ) " & _
'             " SELECT '06' AS tGrupo, 'Cancelacion Cortesias' as Grupo, '00' AS tSubGrupo, '' AS SubGrupo, mdocumento.tDocumento, tUsuario, " & _
'             " convert(nvarchar(8), MDOCUMENTO.fRegistro,3) + ' ' + convert(nvarchar(5), MDOCUMENTO.fRegistro,8) as fFecha, fPago, nVenta, nVenta as nVenta1, 0 as nVenta2, (vCortesia.descripcion) as tObservacion  " & _
'             " From dbo.mDOCUMENTO " & _
'             " left join vcortesia on mdocumento.tcortesia = vcortesia.Codigo " & _
'             " WHERE ltrim(rtrim(isnull(dbo.mDOCUMENTO.tCortesia,'')))<>'' and mdocumento.ttipodocumento<>'00' and " & sBase & _
'             " and testadodocumento='02'"
      Isql = " Insert Into " & sNombre & " (tGrupo, Grupo, tSubGrupo, SubGrupo, tDocumento, tUsuario, fFecha, fPago, nVenta, nVenta1, nVenta2, tObservacion ) " & _
             " SELECT '06' AS tGrupo, 'Cancelacion Cortesias' as Grupo, '00' AS tSubGrupo, '' AS SubGrupo, MAX(mdocumento.tDocumento) as tDocumento, MAX(tUsuario) as tUsuario, " & _
             " MAX(convert(nvarchar(8), MDOCUMENTO.fRegistro,3) + ' ' + convert(nvarchar(5), MDOCUMENTO.fRegistro,8)) as fFecha, MAX(fPago) as fPago, SUM(nVenta) as nVenta, SUM(nVenta) as nVenta1, 0 as nVenta2, Count(nVenta) as tObservacion  " & _
             " From dbo.mDOCUMENTO " & _
             " left join vcortesia on mdocumento.tcortesia = vcortesia.Codigo " & _
             " WHERE ltrim(rtrim(isnull(dbo.mDOCUMENTO.tCortesia,'')))<>'' and mdocumento.ttipodocumento<>'00' and " & sBase & _
             " and testadodocumento='02' GROUP BY dbo.MDOCUMENTO.nAbono"
      Cn.Execute Isql
      Cn.Execute "delete from " & sNombre & " where tGrupo='06' and isnull(nVenta,0)=0"
   Else
      Isql = " Insert Into " & sNombre & " (tGrupo, Grupo, tSubGrupo, SubGrupo, tDocumento, tUsuario, fFecha, fPago, nVenta, nVenta1, nVenta2, tObservacion ) " & _
             " SELECT '06' AS tGrupo, 'Cancelacion Cortesias' as Grupo, '00' AS tSubGrupo, '' AS SubGrupo, mdocumento.tDocumento, tUsuario, " & _
             " convert(nvarchar(8), MDOCUMENTO.fRegistro,3) + ' ' + convert(nvarchar(5), MDOCUMENTO.fRegistro,8) as fFecha, fPago, SUM(nVenta) AS nVenta, Sum(nVenta) as nVenta1, 0 as nVenta2, max(vCortesia.descripcion) as tObservacion " & _
             " From dbo.mDOCUMENTO " & _
             " left join vcortesia on mdocumento.tcortesia = vcortesia.Codigo " & _
             " WHERE ltrim(rtrim(isnull(dbo.mDOCUMENTO.tCortesia,'')))<>'' and mdocumento.ttipodocumento<>'00' and " & sBase & _
             " and testadodocumento='02' group by tdocumento, fRegistro, fpago, tusuario"
      Cn.Execute Isql
   End If
    
   'Inserta los CXC
   If chkCobrar.value Then
      Isql = "insert into " & sNombre & " (tGrupo, Grupo, tSubGrupo, SubGrupo, tDocumento, tUsuario, fFecha, nVenta, nVenta1, tObservacion ) " & _
             "SELECT '07' AS tGrupo, 'Cuentas por Cobrar' AS Grupo, '00' AS tSubGrupo, '' AS SubGrupo, '' AS tDocumento, '' AS tUsuario, '' AS fFecha, " & _
             "SUM(MDOCUMENTO.nVenta) AS nVenta, SUM(MDOCUMENTO.nVenta) AS nVenta1, '' AS tObservacion " & _
             "FROM MDOCUMENTO LEFT JOIN vCOMPANIA ON MDOCUMENTO.tClientePago = vCompania.Codigo  " & _
             "WHERE     isnull(MDOCUMENTO.tClientePago,'') <> '' and " & sBase
             
             '"WHERE  MDOCUMENTO.tEstadoDocumento = '03' and isnull(MDOCUMENTO.tClientePago,'') <> '' and " & sBase
      Cn.Execute Isql
      Cn.Execute "delete from " & sNombre & " where tGrupo='11' and isnull(nVenta,0)=0"
   Else
      Isql = "insert into " & sNombre & " (tGrupo, Grupo, tSubGrupo, SubGrupo, tDocumento, tUsuario, fFecha, nVenta, nVenta1, tObservacion ) " & _
             "SELECT '07' as tGrupo, 'Cuentas por Cobrar' AS Grupo, '00' AS tSubGrupo, '' AS SubGrupo, MDOCUMENTO.tDocumento, MDOCUMENTO.tUsuarioAutoriza, convert(nvarchar(8), MDOCUMENTO.fRegistro,3) + ' ' + convert(nvarchar(5), MDOCUMENTO.fRegistro,8) AS fRegistro, " & _
             "MDOCUMENTO.nVenta, MDOCUMENTO.nVenta as nVenta1, vCOMPANIA.Descripcion AS tObservacion " & _
             "FROM MDOCUMENTO LEFT JOIN vCOMPANIA ON MDOCUMENTO.tClientePago = vCompania.Codigo  " & _
             "WHERE  isnull(MDOCUMENTO.tClientePago,'') <> '' and " & sBase
             '"WHERE  MDOCUMENTO.tEstadoDocumento = '03' and isnull(MDOCUMENTO.tClientePago,'') <> '' and " & sBase
      Cn.Execute Isql
   End If

   'Borra las cuentas corrientes canceladas en el Rango
   Cn.Execute "delete from " & sNombre & " where tGrupo='07' and tDocumento in (select distinct tdocumento from " & sNombre & " where (tGrupo='01' or tGrupo='02' or tGrupo='03' or tGrupo='04' or tGrupo='05' or tGrupo='06'))"
   
   'Recibos Ingreso en el Turno
   If chkIngreso.value Then
      Isql = "Insert Into " & sNombre & " (tGrupo, Grupo, tSubGrupo, SubGrupo, tDocumento, tUsuario, fFecha, fPago, nVenta, nVenta1, nVenta2, tObservacion ) " & _
             "SELECT '08' AS tGrupo, 'Recibo de Ingreso' as Grupo, tTipoPago AS tSubGrupo, vTipoPago.Descripcion AS SubGrupo, '' AS tDocumento, '' AS tUsuario, '' AS fFecha, '' AS fPago, " & _
             "SUM(dbo.vIngreso.nMonto) As nVenta, " & _
             "SUM(CASE WHEN dbo.vIngreso.tMoneda = '01' THEN dbo.vIngreso.nMonto ELSE 0 END) As nVenta1, " & _
             "SUM(CASE WHEN dbo.vIngreso.tMoneda = '02' THEN dbo.vIngreso.nMonto ELSE 0 END) As nVenta2, " & _
             "Count(dbo.vIngreso.nMonto) As tObservacion " & _
             "FROM dbo.vIngreso LEFT OUTER JOIN dbo.vTipoPago ON dbo.vIngreso.tTipoPago = dbo.vTipoPago.Codigo " & _
             "WHERE dbo.vIngreso.lAnticipo=0 and dbo.vIngreso.tEstadoDocumento='01' and " & sCriterio & _
             " Group by tTipoPago, vTipoPago.Descripcion"
      Cn.Execute Isql
      Cn.Execute "delete from " & sNombre & " where tGrupo='06' and isnull(nVenta,0)=0"
   Else
      Isql = "Insert Into " & sNombre & " (tGrupo, Grupo, tSubGrupo, SubGrupo, tDocumento, tUsuario, fFecha, fPago, nTC, nVenta, nVenta1, nVenta2, tObservacion ) " & _
             "SELECT '08' AS tGrupo, 'Recibos de Ingreso' AS Grupo, tTipoPago AS tSubGrupo, vTipoPago.Descripcion AS SubGrupo, dbo.vIngreso.tRecibo, " & _
             "dbo.vIngreso.tUsuario AS tUsuario, CONVERT(nvarchar(8), dbo.vIngreso.fRegistro, 3) + ' ' + CONVERT(nvarchar(5), dbo.vIngreso.fRegistro, 8) AS fRegistro, " & _
             "CONVERT(nvarchar(8), dbo.vIngreso.fRegistro, 3) + ' ' + CONVERT(nvarchar(5), dbo.vIngreso.fRegistro, 8) AS fPago, dbo.vIngreso.nTipoCambio, " & _
             "dbo.vIngreso.nMonto As nVenta, " & _
             "CASE WHEN dbo.vIngreso.tMoneda = '01' THEN dbo.vIngreso.nMonto ELSE 0 END As nVenta1, " & _
             "CASE WHEN dbo.vIngreso.tMoneda = '02' THEN dbo.vIngreso.nMonto ELSE 0 END As nVenta2, " & _
             "'Doc. Gen. : ' + dbo.vIngreso.tTurno + ' Obs: ' + ltrim(dbo.vIngreso.tDescripcion) AS tObservacion " & _
             "FROM dbo.vIngreso LEFT OUTER JOIN dbo.vTipoPago ON dbo.vIngreso.tTipoPago = dbo.vTipoPago.Codigo " & _
             "WHERE dbo.vIngreso.lAnticipo=0 and dbo.vIngreso.tEstadoDocumento='01' and " & sCriterio
      Cn.Execute Isql
   End If
   
   'Recibos Ingreso Anticipos en el Turno
   If chkAnticipo.value Then
      Isql = "Insert Into " & sNombre & " (tGrupo, Grupo, tSubGrupo, SubGrupo, tDocumento, tUsuario, fFecha, fPago, nVenta, nVenta1, nVenta2, tObservacion ) " & _
             "SELECT '09' AS tGrupo, 'Recibos de Anticipos' as Grupo, tTipoPago AS tSubGrupo, vTipoPago.Descripcion AS SubGrupo, '' AS tDocumento, '' AS tUsuario, '' AS fFecha, '' AS fPago, " & _
             "SUM(dbo.vIngreso.nMonto) As nVenta, " & _
             "SUM(CASE WHEN dbo.vIngreso.tMoneda = '01' THEN dbo.vIngreso.nMonto ELSE 0 END) As nVenta1, " & _
             "SUM(CASE WHEN dbo.vIngreso.tMoneda = '02' THEN dbo.vIngreso.nMonto ELSE 0 END) As nVenta2, " & _
             "Count(dbo.vIngreso.nMonto) As tObservacion " & _
             "FROM dbo.vIngreso LEFT OUTER JOIN dbo.vTipoPago ON dbo.vIngreso.tTipoPago = dbo.vTipoPago.Codigo " & _
             "WHERE dbo.vIngreso.lAnticipo=1 and dbo.vIngreso.tEstadoDocumento='01' and " & sCriterio & _
             " Group by tTipoPago, vTipoPago.Descripcion"
      Cn.Execute Isql
      Cn.Execute "delete from " & sNombre & " where tGrupo='07' and isnull(nVenta,0)=0"
   Else
      Isql = "Insert Into " & sNombre & " (tGrupo, Grupo, tSubGrupo, SubGrupo, tDocumento, tUsuario, fFecha, fPago, nTC, nVenta, nVenta1, nVenta2, tObservacion ) " & _
             "SELECT '09' AS tGrupo, 'Recibos de Anticipos' AS Grupo, tTipoPago AS tSubGrupo, vTipoPago.Descripcion AS SubGrupo, dbo.vIngreso.tRecibo, " & _
             "dbo.vIngreso.tUsuario AS tUsuario, CONVERT(nvarchar(8), dbo.vIngreso.fRegistro, 3) + ' ' + CONVERT(nvarchar(5), dbo.vIngreso.fRegistro, 8) AS fRegistro, " & _
             "CONVERT(nvarchar(8), dbo.vIngreso.fRegistro, 3) + ' ' + CONVERT(nvarchar(5), dbo.vIngreso.fRegistro, 8) AS fPago, dbo.vIngreso.nTipoCambio, " & _
             "dbo.vIngreso.nMonto As nVenta, " & _
             "CASE WHEN dbo.vIngreso.tMoneda = '01' THEN dbo.vIngreso.nMonto ELSE 0 END As nVenta1, " & _
             "CASE WHEN dbo.vIngreso.tMoneda = '02' THEN dbo.vIngreso.nMonto ELSE 0 END As nVenta2, " & _
             "'Doc. Gen. : ' + dbo.vIngreso.tTurno + ' Obs: ' + ltrim(dbo.vIngreso.tDescripcion) AS tObservacion " & _
             "FROM dbo.vIngreso LEFT OUTER JOIN dbo.vTipoPago ON dbo.vIngreso.tTipoPago = dbo.vTipoPago.Codigo " & _
             "WHERE dbo.vIngreso.lAnticipo=1 and dbo.vIngreso.tEstadoDocumento='01' and " & sCriterio
      Cn.Execute Isql
   End If
   
   'Recibos Egreso
   If chkEgreso.value Then
      Isql = "Insert Into " & sNombre & " (tGrupo, Grupo, tSubGrupo, SubGrupo, tDocumento, tUsuario, fFecha, fPago, nVenta, nVenta1, nVenta2, tObservacion ) " & _
             "SELECT '10' AS tGrupo, 'Recibos de Egreso' as Grupo, '00' AS tSubGrupo, '' AS SubGrupo, '' AS tDocumento, '' AS tUsuario, '' AS fFecha, '' AS fPago, " & _
             "SUM(dbo.vEgreso.nMonto) As nVenta, " & _
             "SUM(CASE WHEN dbo.vEgreso.tMoneda = '01' THEN dbo.vEgreso.nMonto ELSE 0 END) As nVenta1, " & _
             "SUM(CASE WHEN dbo.vEgreso.tMoneda = '02' THEN dbo.vEgreso.nMonto ELSE 0 END) As nVenta2, " & _
             "Count(dbo.vEgreso.nMonto) As tObservacion " & _
             "FROM dbo.vEgreso " & _
             "WHERE dbo.vEgreso.tEstadoDocumento='01' and " & sCriterio
      Cn.Execute Isql
      Cn.Execute "delete from " & sNombre & " where tGrupo='08' and isnull(nVenta,0)=0"
   Else
      Isql = "Insert Into " & sNombre & " (tGrupo, Grupo, tSubGrupo, SubGrupo, tDocumento, tUsuario, fFecha, fPago, nTC, nVenta, nVenta1, nVenta2, tObservacion ) " & _
             "SELECT '10' AS tGrupo, 'Recibos de Egreso' AS Grupo, substring(tTipoEgreso,2,2) AS tSubGrupo, vegreso.tipoegreso AS SubGrupo, dbo.vEgreso.tRecibo, " & _
             "dbo.vEgreso.tUsuario AS tUsuario, CONVERT(nvarchar(8), dbo.vEgreso.fRegistro, 3) + ' ' + CONVERT(nvarchar(5), dbo.vEgreso.fRegistro, 8) AS fRegistro, " & _
             "CONVERT(nvarchar(8), dbo.vEgreso.fRegistro, 3) + ' ' + CONVERT(nvarchar(5), dbo.vEgreso.fRegistro, 8) AS fPago, dbo.vEgreso.nTipoCambio, " & _
             "dbo.vEgreso.nMonto As nVenta, " & _
             "CASE WHEN dbo.vEgreso.tMoneda = '01' THEN dbo.vEgreso.nMonto ELSE 0 END As nVenta1, " & _
             "CASE WHEN dbo.vEgreso.tMoneda = '02' THEN dbo.vEgreso.nMonto ELSE 0 END As nVenta2, " & _
             "'Doc. Gen. : ' + dbo.vEgreso.tTurno + ' Obs: ' + ltrim(dbo.vEgreso.tDescripcion) AS tObservacion " & _
             "FROM dbo.vEgreso " & _
             "WHERE dbo.vEgreso.tEstadoDocumento='01' and " & sCriterio
      Cn.Execute Isql
   End If
   
   'Cortesias
   If chkNoCortesia.value Then
      If chkCortesia.value Then
         Isql = "Insert Into " & sNombre & " (tGrupo, Grupo, tSubGrupo, SubGrupo, tDocumento, tUsuario, fFecha, fPago, nVenta, nVenta1, tObservacion ) " & _
                "SELECT '11' AS tGrupo, 'Cortesias' AS Grupo, '00' AS tSubGrupo, '' AS SubGrupo, '' AS tDocumento, '' AS tUsuario, '' AS fFecha, '' AS fPago, " & _
                "SUM(dbo.MDOCUMENTO.nVenta) AS nVenta, SUM(dbo.MDOCUMENTO.nVenta) AS nVenta1, " & _
                "'' as tObservacion " & _
                "From dbo.MDOCUMENTO " & _
                "WHERE dbo.MDOCUMENTO.tTipoDocumento ='00' and " & sBase
         Cn.Execute Isql
         Cn.Execute "delete from " & sNombre & " where tGrupo='09' and isnull(nVenta,0)=0"
      Else
         Isql = "Insert Into " & sNombre & " (tGrupo, Grupo, tSubGrupo, SubGrupo, tDocumento, tUsuario, fFecha, fPago, nVenta, nVenta1, tObservacion ) " & _
                "SELECT '11' AS tGrupo, 'Cortesias' AS Grupo, '00' AS tSubGrupo, dbo.vCortesia.Descripcion AS SubGrupo, dbo.MDOCUMENTO.tDocumento, dbo.MDOCUMENTO.tUsuarioAutoriza AS tUsuario, " & _
                "CONVERT(nvarchar(8), dbo.MDOCUMENTO.fRegistro, 3) + ' ' + CONVERT(nvarchar(5), dbo.MDOCUMENTO.fRegistro, 8) AS fRegistro, CONVERT(nvarchar(8), dbo.MDOCUMENTO.fPago, 3) + ' ' + CONVERT(nvarchar(5), dbo.MDOCUMENTO.fPago, 8) AS fPago, " & _
                "dbo.MDOCUMENTO.nVenta AS nVenta, dbo.MDOCUMENTO.nVenta AS nVenta1, " & _
                "dbo.MDOCUMENTO.tTurno + ' ' + ISNULL(dbo.vCortesia.Descripcion, N'') AS tObservacion " & _
                "FROM dbo.MDOCUMENTO LEFT OUTER JOIN dbo.vCortesia ON dbo.MDOCUMENTO.tCortesia = dbo.vCortesia.Codigo LEFT OUTER JOIN dbo.vEstadoDocumento ON dbo.MDOCUMENTO.tEstadoDocumento = dbo.vEstadoDocumento.Codigo " & _
                "WHERE dbo.MDOCUMENTO.tTipoDocumento ='00' and " & sBase
         Cn.Execute Isql
      End If
   End If
      
   'Cuentas Corrientes
   If chkCuenta.value Then
         Isql = "Insert Into " & sNombre & " (tGrupo, Grupo, tSubGrupo, SubGrupo, tDocumento, tUsuario, fFecha, nVenta, nVenta1, tObservacion ) " & _
                "SELECT '12' AS tGrupo, 'Cuenta Corriente' AS Grupo, '00' AS tSubGrupo, '' AS SubGrupo, '' AS tDocumento, '' AS tUsuario, '' AS fFecha, " & _
                "SUM(DPEDIDO.nVenta) AS nVenta, SUM(DPEDIDO.nVenta) AS nVenta1, '' AS tObservacion " & _
                "FROM dbo.MPEDIDO LEFT OUTER JOIN DPEDIDO ON dbo.MPEDIDO.tCodigoPedido = dbo.DPEDIDO.tCodigoPedido " & _
                "WHERE " & yBase & " AND (dbo.MPEDIDO.tEstadoPedido<> N'03') AND (dbo.DPEDIDO.tEstadoItem <> N'A') and isnull(mpedido.tclientectacte,'')<>'' "
      Cn.Execute Isql
      Cn.Execute "delete from " & sNombre & " where tGrupo='10' and isnull(nVenta,0)=0"
   Else
      Isql = "Insert Into " & sNombre & " (tGrupo, Grupo, tSubGrupo, SubGrupo, tDocumento,tUsuario,fFecha,nVenta,nVenta1,tObservacion ) " & _
             "SELECT '12' AS tGrupo, 'Cuenta Corriente' AS Grupo, '00' AS tSubGrupo,' ' AS SubGrupo,MPEDIDO.tCodigoPedido AS tDocumento, " & _
             "MAX(MPEDIDO.tUsuario) AS tUsuario,MAX(CONVERT(nvarchar(8),MPEDIDO.fRegistro, 3) + ' ' + CONVERT(nvarchar(5), " & _
             "MPEDIDO.fRegistro, 8)) AS fRegistro, SUM(DPEDIDO.nVenta) AS nVenta1,SUM(DPEDIDO.nVenta) AS nVenta," & _
             "(vCompania.tNomSoc) AS tObservacion " & _
             "FROM dbo.MPEDIDO LEFT OUTER JOIN " & _
             "DPEDIDO ON dbo.MPEDIDO.tCodigoPedido = dbo.DPEDIDO.tCodigoPedido RIGHT OUTER JOIN " & _
             "vCompania ON dbo.MPEDIDO.tClienteCtaCte = dbo.vCompania.Codigo " & _
             "WHERE " & yBase & " AND (dbo.MPEDIDO.tEstadoPedido <> N'03') AND (dbo.DPEDIDO.tEstadoItem <> N'A') AND ISNULL(MPEDIDO.TCLIENTECTACTE,'')<>'' " & _
             "GROUP BY dbo.vCompania.tNomSoc,dbo.MPEDIDO.tCodigoPedido , dbo.MPEDIDO.fFecha, dbo.MPEDIDO.tClienteCtaCte"
      Cn.Execute Isql
   End If
       
   'Inserta los Anulados
   If chkAnulado.value Then
      Isql = "insert into " & sNombre & " (tGrupo, Grupo, tSubGrupo, SubGrupo, tDocumento, tUsuario, fFecha, nVenta, tObservacion ) " & _
             "SELECT '13' AS tGrupo, 'Documentos Anulados' AS Grupo, '00' AS tSubGrupo, '' AS SubGrupo, '' AS tDocumento, '' AS tUsuario, '' AS fFecha, " & _
             "SUM(MDOCUMENTO.nVenta) as nVenta, '' as tOBservacion " & _
             "FROM MDOCUMENTO " & _
             "WHERE tEstadoDocumento = '04' and " & sBase
      Cn.Execute Isql
      Cn.Execute "delete from " & sNombre & " where tGrupo='12' and isnull(nVenta,0)=0"
   Else
      Isql = "insert into " & sNombre & " (tGrupo, Grupo, tSubGrupo, SubGrupo, tDocumento, tUsuario, fFecha, nVenta, tObservacion ) " & _
             "SELECT '13' as tGrupo, 'Documentos Anulados' AS Grupo, '00' AS tSubGrupo, '' AS SubGrupo, MDOCUMENTO.tDocumento, MDOCUMENTO.tUsuario, " & _
             "convert(nvarchar(8), MDOCUMENTO.fRegistro,3) + ' ' + convert(nvarchar(5), MDOCUMENTO.fRegistro,8) AS fRegistro, " & _
             "MDOCUMENTO.nVenta, MDOCUMENTO.tObservacion " & _
             "FROM MDOCUMENTO " & _
             "WHERE tEstadoDocumento = '04' and " & sBase
      Cn.Execute Isql
   End If
      
   'Inserta Pedidos facturados en otro rango
   If chkPedidos.value Then
      Isql = "Insert Into " & sNombre & " (tGrupo, Grupo, tSubGrupo, SubGrupo, tDocumento, tUsuario, fFecha, nVenta, nVenta1,tObservacion ) " & _
             "SELECT '14' AS tGrupo, 'Pedidos Facturados en otro Rango' AS Grupo, '00' AS tSubGrupo,' ' AS SubGrupo, '' AS tDocumento, " & _
             "'', '', sum(T1.nVenta) as nVenta, sum(T1.nVenta) as nVenta2, '' " & _
             "FROM (SELECT DISTINCT  dbo.MPEDIDO.tCodigoPedido, dbo.MPEDIDO.fFecha, dbo.DPEDIDO.nVenta, dbo.MPEDIDO.tUsuario, dbo.DPEDIDO.tDocumento, dbo.DPEDIDO.tItem " & _
             "FROM dbo.DPEDIDO INNER JOIN dbo.MPEDIDO ON dbo.DPEDIDO.tCodigoPedido = dbo.MPEDIDO.tCodigoPedido RIGHT OUTER JOIN dbo.MDOCUMENTO INNER JOIN dbo.DDOCUMENTO ON dbo.MDOCUMENTO.tDocumento = dbo.DDOCUMENTO.tDocumento ON dbo.MPEDIDO.tCodigoPedido = dbo.DDOCUMENTO.tCodigoPedido " & _
             "WHERE (" & sBase & " AND not (" & yBase & ")) ) T1 "
      Cn.Execute Isql
      Cn.Execute "delete from " & sNombre & " where tGrupo='15' and isnull(nVenta,0)=0"
   Else
      Isql = "Insert Into " & sNombre & " (tGrupo, Grupo, tSubGrupo, SubGrupo, tDocumento, tUsuario, fFecha, nVenta, nVenta1,tObservacion ) " & _
             "SELECT '14' AS tGrupo, 'Pedidos Facturados en otro Rango' AS Grupo, '00' AS tSubGrupo,' ' AS SubGrupo, T1.tCodigoPedido AS tDocumento, " & _
             "T1.tUsuario, T1.fFecha, sum(T1.nVenta) as nVenta, sum(T1.nVenta) as nVenta2, T1.tDocumento " & _
             "FROM (SELECT DISTINCT  dbo.MPEDIDO.tCodigoPedido, dbo.MPEDIDO.fFecha, dbo.DPEDIDO.nVenta, dbo.MPEDIDO.tUsuario, dbo.DPEDIDO.tDocumento, dbo.DPEDIDO.tItem " & _
             "FROM dbo.DPEDIDO INNER JOIN dbo.MPEDIDO ON dbo.DPEDIDO.tCodigoPedido = dbo.MPEDIDO.tCodigoPedido RIGHT OUTER JOIN dbo.MDOCUMENTO INNER JOIN dbo.DDOCUMENTO ON dbo.MDOCUMENTO.tDocumento = dbo.DDOCUMENTO.tDocumento ON dbo.MPEDIDO.tCodigoPedido = dbo.DDOCUMENTO.tCodigoPedido " & _
             "WHERE (" & sBase & " AND not (" & yBase & ")) ) T1 " & _
             "Group by T1.tcodigopedido, T1.tUsuario, T1.fFecha, T1.tDocumento"
      Cn.Execute Isql
   End If
      
   'Inserta Correlativo de Documentos
   Isql = "insert into " & sNombre & " (tGrupo, Grupo, tSubGrupo, SubGrupo, tDocumento, fFecha, nVenta1, nVenta2, tObservacion ) " & _
          "SELECT '15' as tGrupo, 'Correlativo Documentos' AS Grupo, '00' AS tSubGrupo, '' AS SubGrupo, min(tDocumento) as Minimo, max(tDocumento) as Maximo, " & _
          "sum(case when tEstadoDocumento<>'04' then 1 else 0 End) as Emitido, sum(case when tEstadoDocumento='04' then 1 else 0 End) As Anulado, " & _
          "vTipoDocumento.Descripcion as Documento " & _
          "from MDOCUMENTO LEFT JOIN vTipoDocumento ON MDOCUMENTO.tTipoDocumento = vTipoDocumento.Codigo " & _
          "where tTipoDocumento <>'00' and " & sCriterio & " " & _
          "Group by tTipoDocumento, vTipoDocumento.Descripcion, substring(tDocumento,2,5) "
   Cn.Execute Isql
   
    Isql = "Insert Into " & sNombre & " (tGrupo, Grupo, tSubGrupo, SubGrupo, tDocumento, tUsuario, fFecha, fPago, nVenta, nVenta1, tObservacion ) " & _
           " SELECT '16' AS tGrupo, 'Documentos No Fiscales' AS Grupo, '00' AS tSubGrupo, '' AS SubGrupo, dbo.MDOCUMENTO.tDocumento, dbo.MDOCUMENTO.tUsuario AS tUsuario, " & _
           " CONVERT(nvarchar(8), dbo.MDOCUMENTO.fRegistro, 3) + ' '+ CONVERT(nvarchar(5), dbo.MDOCUMENTO.fRegistro, 8) AS fRegistro, CONVERT(nvarchar(8), dbo.MDOCUMENTO.fPago, 3) + ' '+ CONVERT(nvarchar(5), dbo.MDOCUMENTO.fPago, 8) AS fPago," & _
           " dbo.MDOCUMENTO.nVenta AS nVenta, dbo.MDOCUMENTO.nVenta AS nVenta1, CASE WHEN dbo.MDOCUMENTO.tOBSERVACION IS NULL THEN ' ' ELSE dbo.MDOCUMENTO.tOBSERVACION END    AS tObservacion " & _
           " FROM dbo.MDOCUMENTO LEFT OUTER JOIN dbo.vCortesia ON dbo.MDOCUMENTO.tCortesia = dbo.vCortesia.Codigo LEFT OUTER JOIN dbo.vEstadoDocumento ON dbo.MDOCUMENTO.tEstadoDocumento = dbo.vEstadoDocumento.Codigo " & _
           " WHERE  mdocumento.ttipodocumento in (select codigo from vtipodocumento where registroventa=0) and dbo.mdocumento.testadodocumento<>'04' and " & sBase
    
    Cn.Execute Isql
   
   '---- NOtas de Credito
        If lNcOfisis Then
            Isql = "Insert Into " & sNombre & " (tGrupo, Grupo, tSubGrupo, SubGrupo, tDocumento, tUsuario, fFecha, fPago, nTC, nVenta, nVenta1, nVenta2, tObservacion ) " & _
                    " select  '17' AS tGrupo, 'Notas de Credito' AS Grupo, '00' AS tSubGrupo, '' AS SubGrupo, dbo.MNOTACREDITO.tNotaCredito, dbo.DPAGODOCUMENTO.tUsuario AS tUsuario, CONVERT(nvarchar(8), " & _
                    " dbo.MNOTACREDITO.fRegistro, 3) + ' ' + CONVERT(nvarchar(5), dbo.MNOTACREDITO.fRegistro, 8) AS fRegistro, CONVERT(nvarchar(8), dbo.MNOTACREDITO.fRegistro, 3) + ' ' + CONVERT(nvarchar(5), " & _
                    " dbo.DPAGODOCUMENTO.fRegistro, 8) AS fPago, dbo.DPAGODOCUMENTO.nTipoCambio, (dbo.MNOTACREDITO.nVenta*(-1)) As nVenta," & _
                    " (CASE WHEN dbo.DPAGODOCUMENTO.tMoneda = '01'  THEN dbo.DPAGODOCUMENTO.nMonto ELSE 0 END *(-1))As nVenta1, " & _
                    " (CASE WHEN dbo.DPAGODOCUMENTO.tMoneda = '02' THEN dbo.DPAGODOCUMENTO.nMonto ELSE 0 END*(-1)) As nVenta2," & _
                    " 'Doc. Gen. : ' + dbo.MNOTACREDITO.tTurno AS tObservacion  from MNOTACREDITO left outer join MDOCUMENTO on MNOTACREDITO.tDocumento=MDOCUMENTO.tDocumento " & _
                    " left outer join DPAGODOCUMENTO on MNOTACREDITO.tDocumento=DPAGODOCUMENTO.tDocumento where dbo.MNOTACREDITO.tEstadoDocumento in ('05','02') and " & sCriterioNC

            Cn.Execute Isql

        End If
    
   'Select del Reporte
   Isql = "SELECT * FROM [" & sNombre & "] order by tGrupo, tSubGrupo, SubGrupo, tDocumento "
  ' Cn.Execute "select * into prueba1 from  [" & sNombre & "] order by tGrupo, tSubGrupo, SubGrupo, tDocumento "
   Set rsReporte = Lib.OpenRecordset(Isql, Cn)
           
   'Calcula Totales
   'Total de Documentos
   nDocumento = Calcular("Select Sum(nVenta1) as Codigo From " & sNombre & " Where tGrupo='00'", Cn)
   Set RsSumas = Lib.OpenRecordset("select tGrupo, sum(nVenta1) as nVenta1, sum(nVenta2) as nVenta2, sum(nVenta2*isnull(nTC,0)) as nVenta3 From  " & sNombre & " Group by tGrupo", Cn)
      
      
   'Efectivo
   RsSumas.Filter = "tGrupo='01'"
   If RsSumas.RecordCount > 0 Then
      nEfectivoN = IIf(IsNull(RsSumas!nVenta1), 0, RsSumas!nVenta1)
      nEfectivoE = IIf(IsNull(RsSumas!nVenta2), 0, RsSumas!nVenta2)
      xEfectivoE = IIf(IsNull(RsSumas!nVenta3), 0, RsSumas!nVenta3)
   Else
      nEfectivoN = 0
      nEfectivoE = 0
      xEfectivoE = 0
   End If
                 
   'Cheques
   RsSumas.Filter = "tGrupo='03'"
   If RsSumas.RecordCount > 0 Then
      nChequeN = IIf(IsNull(RsSumas!nVenta1), 0, RsSumas!nVenta1)
      nChequeE = IIf(IsNull(RsSumas!nVenta2), 0, RsSumas!nVenta2)
      xChequeE = IIf(IsNull(RsSumas!nVenta3), 0, RsSumas!nVenta3)
   Else
      nChequeN = 0
      nChequeE = 0
      xChequeE = 0
   End If
   
   'Otros
   RsSumas.Filter = "tGrupo='04'"
   If RsSumas.RecordCount > 0 Then
      nOtroN = IIf(IsNull(RsSumas!nVenta1), 0, RsSumas!nVenta1)
      nOtroE = IIf(IsNull(RsSumas!nVenta2), 0, RsSumas!nVenta2)
      xOtroE = IIf(IsNull(RsSumas!nVenta3), 0, RsSumas!nVenta3)
   Else
      nOtroN = 0
      nOtroE = 0
      xOtroE = 0
   End If
   
   'Puntos
   RsSumas.Filter = "tGrupo='05'"
   If RsSumas.RecordCount > 0 Then
      nPunto = IIf(IsNull(RsSumas!nVenta1), 0, RsSumas!nVenta1)
   Else
      nPunto = 0
   End If
   
   'Cancelacion Cortesias
   RsSumas.Filter = "tGrupo='06'"
   If RsSumas.RecordCount > 0 Then
      nCortesia = IIf(IsNull(RsSumas!nVenta1), 0, RsSumas!nVenta1)
   Else
      nCortesia = 0
   End If
          
   'Cuentas por Cobrar
   RsSumas.Filter = "tGrupo='07'"
   If RsSumas.RecordCount > 0 Then
      nCobrar = IIf(IsNull(RsSumas!nVenta1), 0, RsSumas!nVenta1)
   Else
      nCobrar = 0
   End If
   
   'Recibos de egreso
   RsSumas.Filter = "tGrupo='10'"
   If RsSumas.RecordCount > 0 Then
      nEgresoN = IIf(IsNull(RsSumas!nVenta1), 0, RsSumas!nVenta1)
      nEgresoE = IIf(IsNull(RsSumas!nVenta2), 0, RsSumas!nVenta2)
      xEgresoE = IIf(IsNull(RsSumas!nVenta3), 0, RsSumas!nVenta3)
   Else
      nEgresoN = 0
      nEgresoE = 0
      xEgresoE = 0
   End If
   
   If lNcOfisis Then
    'Notas de Credito
        RsSumas.Filter = "tGrupo='17'"
        If RsSumas.RecordCount > 0 Then
           nNotaCreditoN = IIf(IsNull(RsSumas!nVenta1), 0, RsSumas!nVenta1)
           nNotaCreditoE = IIf(IsNull(RsSumas!nVenta2), 0, RsSumas!nVenta2)
           xNotaCreditoE = IIf(IsNull(RsSumas!nVenta3), 0, RsSumas!nVenta3)
        Else
           nNotaCreditoN = 0
           nNotaCreditoE = 0
           xNotaCreditoE = 0
        End If
    Else
           nNotaCreditoN = 0
           nNotaCreditoE = 0
           xNotaCreditoE = 0
   End If
   'Recibos de ingreso Efectivo
   Set RsSumas = Lib.OpenRecordset("select tGrupo, sum(nVenta1) as nVenta1, sum(nVenta2) as nVenta2, sum(nVenta2*isnull(nTC,0)) as nVenta3 From  " & sNombre & " where tGrupo='08' and tSubGrupo='01' Group by tGrupo", Cn)
   If RsSumas.RecordCount > 0 Then
      nIngresoN = IIf(IsNull(RsSumas!nVenta1), 0, RsSumas!nVenta1)
      nIngresoE = IIf(IsNull(RsSumas!nVenta2), 0, RsSumas!nVenta2)
      xIngresoE = IIf(IsNull(RsSumas!nVenta3), 0, RsSumas!nVenta3)
   Else
      nIngresoN = 0
      nIngresoE = 0
      xIngresoE = 0
   End If
   

   
   'Tarjetas
   Set RsSumas = Lib.OpenRecordset("Select tCodigoTarjeta, tDetallado from TTARJETACREDITO where lActivo=1", Cn)
   For i = 1 To 8
       RsSumas.Filter = "tCodigoTarjeta='0" & LTrim(str(i)) & "'"
       If RsSumas.RecordCount > 0 Then
          tTarjeta(i) = IIf(IsNull(RsSumas!tDetallado), 0, RsSumas!tDetallado)
       Else
          tTarjeta(i) = "no definido"
       End If
   Next i
   
   Set RsSumas = Lib.OpenRecordset("select tSubGrupo, SubGrupo, sum(nVenta1) as nVenta1, sum(nVenta2) as nVenta2 From  " & sNombre & " where tGrupo='02' Group by tSubGrupo, SubGrupo", Cn)
   nTarjetaT = 0
   nPropinaT = 0
   nReciboT = 0
   For i = 1 To 8
       RsSumas.Filter = "tSubGrupo='0" & LTrim(str(i)) & "'"
       If RsSumas.RecordCount > 0 Then
          nTarjeta(i) = IIf(IsNull(RsSumas!nVenta1), 0, RsSumas!nVenta1)
          nPropina(i) = IIf(IsNull(RsSumas!nVenta2), 0, RsSumas!nVenta2)
       Else
          nTarjeta(i) = 0
          nPropina(i) = 0
       End If
       nTarjetaT = nTarjetaT + nTarjeta(i)
       nPropinaT = nPropinaT + nPropina(i)
   Next i
   
   'Recibos de ingreso Anticipos Efectivo
   Set RsSumas = Lib.OpenRecordset("select tGrupo, sum(nVenta1) as nVenta1, sum(nVenta2) as nVenta2, sum(nVenta2*isnull(nTC,0)) as nVenta3 From  " & sNombre & " where tGrupo='09' and tSubGrupo='01' Group by tGrupo", Cn)
   If RsSumas.RecordCount > 0 Then
      nIngresoAN = IIf(IsNull(RsSumas!nVenta1), 0, RsSumas!nVenta1)
      nIngresoAE = IIf(IsNull(RsSumas!nVenta2), 0, RsSumas!nVenta2)
      xIngresoAE = IIf(IsNull(RsSumas!nVenta3), 0, RsSumas!nVenta3)
   Else
      nIngresoAN = 0
      nIngresoAE = 0
      xIngresoAE = 0
   End If
   
   'Ingresos por Tarjeta
   Set RsSumas = Lib.OpenRecordset("select tTarjeta, sum(nMonto) as nVenta1 From vIngreso where tTipoPago='02' and " & sCriterio & " Group by tTarjeta", Cn)
   For i = 1 To 8
       RsSumas.Filter = "tTarjeta='0" & LTrim(str(i)) & "'"
       If RsSumas.RecordCount > 0 Then
          nRecibo(i) = IIf(IsNull(RsSumas!nVenta1), 0, RsSumas!nVenta1)
       Else
          nRecibo(i) = 0
       End If
       nReciboT = nReciboT + nRecibo(i)
   Next i
  
   'Documentos cobrados de otro turno
   Isql = "select sum(t1.nventa) as codigo From (select distinct dpagodocumento.tdocumento, nventa " & _
          "FROM dbo.MDOCUMENTO INNER JOIN dbo.DPAGODOCUMENTO ON dbo.MDOCUMENTO.tDocumento = dbo.DPAGODOCUMENTO.tDocumento " & _
          "where " & xBase & " and dpagodocumento.tdocumento not in (select tdocumento from mdocumento where tTipoDocumento<>'00' and testadodocumento='02' and " & sBase & ")) T1"
   nOtroDoc = Calcular(Isql, Cn)
     
   'Dolares Cobrados
   Isql = "select sum(nDolar) as Suma, sum(nDolar * nTipoCambio) as xSuma from dpagodocumento where tTipopago='01' and tMoneda='02' and " & xBase
   Set RsSumas = Lib.OpenRecordset(Isql, Cn)
   If RsSumas.RecordCount > 0 Then
      nDolar = IIf(IsNull(RsSumas!Suma), 0, RsSumas!Suma)
      xDolar = IIf(IsNull(RsSumas!xSuma), 0, RsSumas!xSuma)
   Else
      nDolar = 0
      xDolar = 0
   End If
   nNoCobrado = Calcular("select sum(nVenta) as codigo from MDOCUMENTO where tEstadoDocumento='01' and " & sCriterio, Cn)

   'Totales de los documentos emitidos
   Isql = "select sum(nNeto) as nNeto, sum(nPrecioImpuesto1) as nImpuesto1, sum(nPrecioImpuesto2) as nImpuesto2, sum(nPrecioImpuesto3) as nImpuesto3, sum(nVenta) as nVenta " & _
          "From mdocumento where tEstadoDocumento<>'04' and tTipoDocumento <>'00' and " & sCriterio

   Set RsSumas = Lib.OpenRecordset(Isql, Cn)
   If RsSumas.RecordCount > 0 Then
      nNeto = IIf(IsNull(RsSumas!nNeto), 0, RsSumas!nNeto)
      nImpuesto1 = IIf(IsNull(RsSumas!nImpuesto1), 0, RsSumas!nImpuesto1)
      nImpuesto2 = IIf(IsNull(RsSumas!nImpuesto2), 0, RsSumas!nImpuesto2)
      nImpuesto3 = IIf(IsNull(RsSumas!nImpuesto3), 0, RsSumas!nImpuesto3)
      nTotalDoc = IIf(IsNull(RsSumas!nVenta), 0, RsSumas!nVenta)
   Else
      nNeto = 0
      nImpuesto1 = 0
      nImpuesto2 = 0
      nImpuesto3 = 0
      nTotalDoc = 0
   End If
                
   'Adultos y Ni�os
   Set RsSumas = Lib.OpenRecordset("select sum(nAdulto) as Adulto from MPEDIDO where MPEDIDO.tTipoPedido = '01' and MPEDIDO.tEstadoPedido <> '03' and " & sCriterio, Cn)
   If RsSumas.RecordCount > 0 Then
      nAdulto = IIf(IsNull(RsSumas!Adulto), 0, RsSumas!Adulto)
   Else
      nAdulto = 0
   End If
   
   'Totales por Tipo Pedido
   Cn.Execute "delete from " & sSuma
'   Isql = "insert " & sSuma & " (tDocumento, nNeto, nImpuesto1, nImpuesto2, nImpuesto3, nVenta, tTipoPedido,RV) " & _
'              "SELECT DISTINCT  dbo.MDOCUMENTO.tDocumento, dbo.MDOCUMENTO.nNeto, dbo.MDOCUMENTO.nPrecioImpuesto1, dbo.MDOCUMENTO.nPrecioImpuesto2, dbo.MDOCUMENTO.nPrecioImpuesto3, dbo.MDOCUMENTO.nVenta, dbo.MPEDIDO.tTipoPedido,CASE WHEN ISNULL(dbo.vTipoDocumento.RegistroVenta, 0) = 1 THEN 'R' ELSE 'N' END AS RVENTA " & _
'              "FROM  dbo.vTipoDocumento INNER JOIN dbo.MDOCUMENTO ON dbo.vTipoDocumento.Codigo = dbo.MDOCUMENTO.tTipoDocumento LEFT OUTER JOIN dbo.MPEDIDO INNER JOIN dbo.DDOCUMENTO ON dbo.MPEDIDO.tCodigoPedido = dbo.DDOCUMENTO.tCodigoPedido ON dbo.MDOCUMENTO.tDocumento = dbo.DDOCUMENTO.tDocumento " & _
'              "where tEstadoDocumento<>'04' and tTipoDocumento<>'00' and " & sBase
              
    If lNcOfisis Then
        Cn.Execute "insert " & sSuma & " (tDocumento, nNeto, nImpuesto1, nImpuesto2, nImpuesto3, nVenta, tTipoPedido,RV) " & _
                   "SELECT DISTINCT  dbo.MDOCUMENTO.tDocumento, dbo.MDOCUMENTO.nNeto, dbo.MDOCUMENTO.nPrecioImpuesto1, dbo.MDOCUMENTO.nPrecioImpuesto2, dbo.MDOCUMENTO.nPrecioImpuesto3, dbo.MDOCUMENTO.nVenta, dbo.MPEDIDO.tTipoPedido,CASE WHEN ISNULL(dbo.vTipoDocumento.RegistroVenta, 0) = 1 THEN 'S' ELSE 'N' END AS RVENTA " & _
                   "FROM  dbo.vTipoDocumento INNER JOIN dbo.MDOCUMENTO ON dbo.vTipoDocumento.Codigo = dbo.MDOCUMENTO.tTipoDocumento LEFT OUTER JOIN dbo.MPEDIDO INNER JOIN dbo.DDOCUMENTO ON dbo.MPEDIDO.tCodigoPedido = dbo.DDOCUMENTO.tCodigoPedido ON dbo.MDOCUMENTO.tDocumento = dbo.DDOCUMENTO.tDocumento " & _
                   "where tEstadoDocumento<>'04' and tTipoDocumento<>'00' and " & sBase

'
        Cn.Execute "insert " & sSuma & " (tDocumento, nNeto, nImpuesto1, nImpuesto2, nImpuesto3, nVenta, tTipoPedido,RV) " & _
                   "SELECT DISTINCT  dbo.MNOTACREDITO.tNotaCredito, (dbo.MNOTACREDITO.nNeto*(-1)), (dbo.MNOTACREDITO.nImpuesto1*(-1)), (dbo.MNOTACREDITO.nImpuesto2*(-1)), (dbo.MNOTACREDITO.nImpuesto3*(-1)), (dbo.MNOTACREDITO.nVenta*(-1)), dbo.MPEDIDO.tTipoPedido,CASE WHEN ISNULL(dbo.vTipoDocumento.RegistroVenta, 0) = 1 THEN 'S' ELSE 'N' END AS RVENTA " & _
                   "FROM  dbo.vTipoDocumento INNER JOIN dbo.MNOTACREDITO ON dbo.vTipoDocumento.Codigo = ( SELECT DBO.MDOCUMENTO.tTipoDocumento FROM MDOCUMENTO WHERE DBO.MDOCUMENTO.TDOCUMENTO = dbo.MNOTACREDITO.tDocumento) LEFT OUTER JOIN dbo.MPEDIDO INNER JOIN dbo.DDOCUMENTO ON dbo.MPEDIDO.tCodigoPedido = dbo.DDOCUMENTO.tCodigoPedido ON dbo.MNOTACREDITO.tDocumento = dbo.DDOCUMENTO.tDocumento " & _
                   "where tEstadoDocumento<>'04' and tTipoDocumento<>'00'  and   dbo.MNOTACREDITO.tEstadoDocumento in ('05','02') and " & sCriterioNC

    Else
        Cn.Execute "insert " & sSuma & " (tDocumento, nNeto, nImpuesto1, nImpuesto2, nImpuesto3, nVenta, tTipoPedido,RV) " & _
                   "SELECT DISTINCT  dbo.MDOCUMENTO.tDocumento, dbo.MDOCUMENTO.nNeto, dbo.MDOCUMENTO.nPrecioImpuesto1, dbo.MDOCUMENTO.nPrecioImpuesto2, dbo.MDOCUMENTO.nPrecioImpuesto3, dbo.MDOCUMENTO.nVenta, dbo.MPEDIDO.tTipoPedido,CASE WHEN ISNULL(dbo.vTipoDocumento.RegistroVenta, 0) = 1 THEN 'S' ELSE 'N' END AS RVENTA " & _
                   "FROM  dbo.vTipoDocumento INNER JOIN dbo.MDOCUMENTO ON dbo.vTipoDocumento.Codigo = dbo.MDOCUMENTO.tTipoDocumento LEFT OUTER JOIN dbo.MPEDIDO INNER JOIN dbo.DDOCUMENTO ON dbo.MPEDIDO.tCodigoPedido = dbo.DDOCUMENTO.tCodigoPedido ON dbo.MDOCUMENTO.tDocumento = dbo.DDOCUMENTO.tDocumento " & _
                   "where tEstadoDocumento<>'04' and tTipoDocumento<>'00' and " & sBase
    End If
'    Isql = "SELECT DISTINCT dbo.MDOCUMENTO.tDocumento, dbo.MDOCUMENTO.nNeto, dbo.MDOCUMENTO.nPrecioImpuesto1, dbo.MDOCUMENTO.nPrecioImpuesto2, dbo.MDOCUMENTO.nPrecioImpuesto3, dbo.MDOCUMENTO.nVenta, tTipoPedido " & _
              "FROM dbo.MPEDIDO INNER JOIN dbo.DDOCUMENTO ON dbo.MPEDIDO.tCodigoPedido = dbo.DDOCUMENTO.tCodigoPedido RIGHT OUTER JOIN dbo.MDOCUMENTO ON dbo.DDOCUMENTO.tDocumento = dbo.MDOCUMENTO.tDocumento " & _
              "where tEstadoDocumento<>'04' and tTipoDocumento<>'00' and " & sBase

   Isql = "select sum(nNeto) as nNeto, sum(nImpuesto1) as nImpuesto1, sum(nImpuesto2) as nImpuesto2, sum(nImpuesto3) as nImpuesto3, sum(nVenta) as nVenta from " & sSuma & " where tTipoPedido='01' AND RV='S'"
   Set RsSumas = Lib.OpenRecordset(Isql, Cn)
   If Not RsSumas.EOF Then
      nValorVenta(1) = IIf(IsNull(RsSumas!nNeto), 0, RsSumas!nNeto)
      nImp1(1) = IIf(IsNull(RsSumas!nImpuesto1), 0, RsSumas!nImpuesto1)
      nImp2(1) = IIf(IsNull(RsSumas!nImpuesto2), 0, RsSumas!nImpuesto2)
      nImp3(1) = IIf(IsNull(RsSumas!nImpuesto3), 0, RsSumas!nImpuesto3)
      nVenta(1) = IIf(IsNull(RsSumas!nVenta), 0, RsSumas!nVenta)
      nTotalPromedio(1) = Calcular("select count(tDocumento) as Codigo from " & sSuma & " where tTipoPedido='01' and rv='S'", Cn)
   End If
      
   Isql = "select sum(nNeto) as nNeto, sum(nImpuesto1) as nImpuesto1, sum(nImpuesto2) as nImpuesto2, sum(nImpuesto3) as nImpuesto3, sum(nVenta) as nVenta from " & sSuma & " where tTipoPedido='02' AND RV='S'"
   Set RsSumas = Lib.OpenRecordset(Isql, Cn)
   If Not RsSumas.EOF Then
      nValorVenta(2) = IIf(IsNull(RsSumas!nNeto), 0, RsSumas!nNeto)
      nImp1(2) = IIf(IsNull(RsSumas!nImpuesto1), 0, RsSumas!nImpuesto1)
      nImp2(2) = IIf(IsNull(RsSumas!nImpuesto2), 0, RsSumas!nImpuesto2)
      nImp3(2) = IIf(IsNull(RsSumas!nImpuesto3), 0, RsSumas!nImpuesto3)
      nVenta(2) = IIf(IsNull(RsSumas!nVenta), 0, RsSumas!nVenta)
      nTotalPromedio(2) = Calcular("select count(tDocumento) as Codigo from " & sSuma & " where tTipoPedido='02' AND RV='S'", Cn)
   End If
   
   Isql = "select sum(nNeto) as nNeto, sum(nImpuesto1) as nImpuesto1, sum(nImpuesto2) as nImpuesto2, sum(nImpuesto3) as nImpuesto3, sum(nVenta) as nVenta from " & sSuma & " where tTipoPedido='03' AND RV='S'"
   Set RsSumas = Lib.OpenRecordset(Isql, Cn)
   If Not RsSumas.EOF Then
      nValorVenta(3) = IIf(IsNull(RsSumas!nNeto), 0, RsSumas!nNeto)
      nImp1(3) = IIf(IsNull(RsSumas!nImpuesto1), 0, RsSumas!nImpuesto1)
      nImp2(3) = IIf(IsNull(RsSumas!nImpuesto2), 0, RsSumas!nImpuesto2)
      nImp3(3) = IIf(IsNull(RsSumas!nImpuesto3), 0, RsSumas!nImpuesto3)
      nVenta(3) = IIf(IsNull(RsSumas!nVenta), 0, RsSumas!nVenta)
      nTotalPromedio(3) = Calcular("select count(tDocumento) as Codigo from " & sSuma & " where tTipoPedido='03' AND RV='S'", Cn)
   End If
                                                                                                                    
   Isql = "select sum(nNeto) as nNeto, sum(nImpuesto1) as nImpuesto1, sum(nImpuesto2) as nImpuesto2, sum(nImpuesto3) as nImpuesto3, sum(nVenta) as nVenta from " & sSuma & " where tTipoPedido='04' AND RV='S'"
   Set RsSumas = Lib.OpenRecordset(Isql, Cn)
   If Not RsSumas.EOF Then
      nValorVenta(4) = IIf(IsNull(RsSumas!nNeto), 0, RsSumas!nNeto)
      nImp1(4) = IIf(IsNull(RsSumas!nImpuesto1), 0, RsSumas!nImpuesto1)
      nImp2(4) = IIf(IsNull(RsSumas!nImpuesto2), 0, RsSumas!nImpuesto2)
      nImp3(4) = IIf(IsNull(RsSumas!nImpuesto3), 0, RsSumas!nImpuesto3)
      nVenta(4) = IIf(IsNull(RsSumas!nVenta), 0, RsSumas!nVenta)
      nTotalPromedio(4) = Calcular("select count(tDocumento) as Codigo from " & sSuma & " where tTipoPedido='04'  AND RV='S'", Cn)
   End If
                                                                                                                    
   Isql = "select sum(nNeto) as nNeto, sum(nImpuesto1) as nImpuesto1, sum(nImpuesto2) as nImpuesto2, sum(nImpuesto3) as nImpuesto3, sum(nVenta) as nVenta from " & sSuma & " where tTipoPedido='05' AND RV='S'"
   Set RsSumas = Lib.OpenRecordset(Isql, Cn)
   If Not RsSumas.EOF Then
      nValorVenta(5) = IIf(IsNull(RsSumas!nNeto), 0, RsSumas!nNeto)
      nImp1(5) = IIf(IsNull(RsSumas!nImpuesto1), 0, RsSumas!nImpuesto1)
      nImp2(5) = IIf(IsNull(RsSumas!nImpuesto2), 0, RsSumas!nImpuesto2)
      nImp3(5) = IIf(IsNull(RsSumas!nImpuesto3), 0, RsSumas!nImpuesto3)
      nVenta(5) = IIf(IsNull(RsSumas!nVenta), 0, RsSumas!nVenta)
      nTotalPromedio(5) = Calcular("select count(tDocumento) as Codigo from " & sSuma & " where tTipoPedido='05' AND RV='S'", Cn)
   End If
                                                                                                                    
   Isql = "select sum(nNeto) as nNeto, sum(nImpuesto1) as nImpuesto1, sum(nImpuesto2) as nImpuesto2, sum(nImpuesto3) as nImpuesto3, sum(nVenta) as nVenta from " & sSuma & " where RV='N'"
   Set RsSumas = Lib.OpenRecordset(Isql, Cn)
   If Not RsSumas.EOF Then
      nValorVenta(6) = IIf(IsNull(RsSumas!nNeto), 0, RsSumas!nNeto)
      nImp1(6) = IIf(IsNull(RsSumas!nImpuesto1), 0, RsSumas!nImpuesto1)
      nImp2(6) = IIf(IsNull(RsSumas!nImpuesto2), 0, RsSumas!nImpuesto2)
      nImp3(6) = IIf(IsNull(RsSumas!nImpuesto3), 0, RsSumas!nImpuesto3)
      nVenta(6) = IIf(IsNull(RsSumas!nVenta), 0, RsSumas!nVenta)
      nTotalPromedio(6) = Calcular("select count(tDocumento) as Codigo from " & sSuma & " where RV='N'", Cn)
   End If
   'Otros Tipos de Pago
   Isql = "select * from vTipoCancelacion where lActivo=1"
   Set RsSumas = Lib.OpenRecordset(Isql, Cn)
   If RsSumas.RecordCount > 0 Then
      RsSumas.MoveFirst
      For i = 1 To 20
          If RsSumas.EOF Then
             tOtro(i) = "no definido"
             nOtro(i) = 0
          Else
             tOtro(i) = RsSumas!Descripcion
             nOtro(i) = Calcular("select sum(nVenta1) as Codigo from " & sNombre & " where tGrupo='04' and tSubGrupo='" & RsSumas!codigo & "'", Cn)
             xOtro(i) = Calcular("select sum(nVenta2) as Codigo from " & sNombre & " where tGrupo='04' and tSubGrupo='" & RsSumas!codigo & "'", Cn)
             nDiferencia(i) = Calcular("select sum(case when round(nVenta1,2) > round(nVenta,2) then nVenta1 - nVenta else 0 end) as Codigo from " & sNombre & " where tGrupo='04' and tSubGrupo='" & RsSumas!codigo & "'", Cn)
             RsSumas.MoveNext
          End If
      Next i
      nTotalDiferencia = Calcular("select sum(case when round(nVenta1,2) > round(nVenta,2) then nVenta1 - nVenta else 0 end) as Codigo from " & sNombre & " where tGrupo='04'", Cn)
   End If
                                                                                                                                                                                    
                                                                                                                                                                                    
    If lNcOfisis Then
        Isql = "SELECT vTipoDocumento.Descripcion as Documento, min(tDocumento) as Minimo, max(tDocumento) as Maximo, " & _
               "sum(case when tEstadoDocumento<>'04' then 1 else 0 End) as Emitido, sum(case when tEstadoDocumento='04' then 1 else 0 End) As Anulado " & _
               "from MDOCUMENTO LEFT JOIN vTipoDocumento ON MDOCUMENTO.tTipoDocumento = vTipoDocumento.Codigo " & _
               "where tTipoDocumento <>'00' and " & sCriterio & " " & _
               "Group by tTipoDocumento, vTipoDocumento.Descripcion, substring(tDocumento,2,5) " & _
                " union " & _
                "SELECT vTipoDocumento.Descripcion as Documento, min(tNotaCredito) as Minimo, max(tNotaCredito) as Maximo, " & _
               "sum(case when tEstadoDocumento<>'04' then 1 else 0 End) as Emitido, sum(case when tEstadoDocumento='04' then 1 else 0 End) As Anulado " & _
               "from MNOTACREDITO LEFT JOIN vTipoDocumento ON MNOTACREDITO.tTipoDocumento = vTipoDocumento.Codigo " & _
               "where tTipoDocumento <>'00' and dbo.MNOTACREDITO.tEstadoDocumento in ('05','02') and " & sCriterio & " " & _
               "Group by tTipoDocumento, vTipoDocumento.Descripcion, substring(tDocumento,2,5) "
        Set RsSumas = Lib.OpenRecordset(Isql, Cn)
   Else
        Isql = "SELECT vTipoDocumento.Descripcion as Documento, min(tDocumento) as Minimo, max(tDocumento) as Maximo, " & _
                "sum(case when tEstadoDocumento<>'04' then 1 else 0 End) as Emitido, sum(case when tEstadoDocumento='04' then 1 else 0 End) As Anulado " & _
                "from MDOCUMENTO LEFT JOIN vTipoDocumento ON MDOCUMENTO.tTipoDocumento = vTipoDocumento.Codigo " & _
                "where tTipoDocumento <>'00' and " & sCriterio & " " & _
                "Group by tTipoDocumento, vTipoDocumento.Descripcion, substring(tDocumento,2,5) "
         Set RsSumas = Lib.OpenRecordset(Isql, Cn)
   
    End If

   nTotalCorrelativo = RsSumas.RecordCount
   'Correlativo de Documentos
   If RsSumas.RecordCount > 0 Then
      RsSumas.MoveFirst
      For i = 1 To nTotalCorrelativo
          Doc(i) = IIf(IsNull(RsSumas!Documento), "", RsSumas!Documento)
          Min(i) = IIf(IsNull(RsSumas!Minimo), "", RsSumas!Minimo)
          Max(i) = IIf(IsNull(RsSumas!Maximo), "", RsSumas!Maximo)
          Emi(i) = IIf(IsNull(RsSumas!Emitido), 0, RsSumas!Emitido)
          Anu(i) = IIf(IsNull(RsSumas!Anulado), 0, RsSumas!Anulado)
          RsSumas.MoveNext
      Next i
   End If
                          
   'Calcula Totales
   nTotalCobradoN = nEfectivoN + nTarjetaT + nOtroN + nChequeN + nPunto + nCortesia + nNotaCreditoN
   nTotalCobradoE = nEfectivoE + nOtroE + nChequeE + nNotaCreditoE
   xTotalCobradoE = xEfectivoE + xOtroE + xChequeE + xNotaCreditoE
   
   nTotalCobrado = nTotalCobradoN + xTotalCobradoE
   nTotalEfectivoN = nEfectivoN + nIngresoN + nIngresoAN - nEgresoN
   nTotalEfectivoE = nEfectivoE + nIngresoE + nIngresoAE - nEgresoE
   xTotalEfectivoE = xEfectivoE + xIngresoE + xIngresoAE - xEgresoE
   
   nTotalEfectivo = nTotalEfectivoN + xTotalEfectivoE
   nFisEfectivoN = nTotalEfectivoN - (xDolar - xEgresoE - xTotalEfectivoE)
   nFisEfectivo = nFisEfectivoN + xDolar - xEgresoE
   
End Sub

Public Sub Cabecera(Grupo As String, SubGrupo As String)
    Dim nContador As Integer
    nContador = 1
        
    Select Case Grupo
           Case Is = "00" 'Documentos Generado
                Printer.Print "Documentos Generados"
                'Printer.Print " "
                If chkGenerado.value = False Then
                   Printer.Print "Documentos             Monto Usuario Est"
                Else
                   Printer.Print "Documentos         Monto MN"
                End If
                Printer.Print "----------------------------------------"
                
           Case Is = "01"  'Efectivos
                Printer.Print "Pagos con Efectivo"
                'Printer.Print " "
                If chkEfectivo.value = False Then
                   Printer.Print "Documentos          Monto MN   Monto ME"
                Else
                   Printer.Print "Documentos         Monto MN     Monto ME"
                End If
                Printer.Print "----------------------------------------"
                
           Case Is = "02" 'TARJETA
                Printer.Print "Pagos con Tarjeta " & rsReporte!SubGrupo
                'Printer.Print " "
                If chkTarjeta.value = False Then
                   Printer.Print "Documentos         Monto MN Propina Num."
                Else
                   Printer.Print "Documentos         Monto MN      Propina"
                End If
                Printer.Print "----------------------------------------"
                
           Case Is = "03" 'Cheques
                Printer.Print " "
                Printer.Print "Pagos con Cheque/Deposito"
                'Printer.Print " "
                If chkCheque.value = False Then
                   Printer.Print "Documentos        Monto MN Monto ME Num."
                Else
                   Printer.Print "Documentos         Monto MN     Monto ME"
                End If
                 Printer.Print "----------------------------------------"
                 
           Case Is = "04" 'Otros
                Printer.Print "Otros Pagos con " & rsReporte!SubGrupo
                'Printer.Print " "
                If chkCheque.value = False Then
                   Printer.Print "Documentos          Monto MN  Monto ME"
                Else
                   Printer.Print "Documentos         Monto MN     Monto ME"
                End If
                Printer.Print "----------------------------------------"
                   
           Case Is = "05" 'Puntos
                Printer.Print "Pagos con Puntos"
                'Printer.Print " "
                If chkPunto.value = False Then
                   Printer.Print "Documentos          Monto MN Cliente"
                Else
                   Printer.Print "Documentos            Monto MN"
                End If
                Printer.Print "----------------------------------------"
                
          Case Is = "06" ' Transferencia Gratuita
                Printer.Print "Transferencias Gratuitas" 'Cancelaci�n de Cortes�as
                'Printer.Print " "
                If chkPagoCortes�a.value = False Then
                   Printer.Print "Documentos          Monto MN Cliente"
                Else
                   Printer.Print "Documentos         Monto MN"
                End If
                Printer.Print "----------------------------------------"
                
                
          Case Is = "07" ' Cuentas por cobrar
                Printer.Print "Cuentas por Cobrar"
                'Printer.Print " "
                If chkCobrar.value = False Then
                   Printer.Print "Documentos           Monto MN Cliente"
                Else
                   Printer.Print "Documentos           Monto MN"
                End If
                Printer.Print "----------------------------------------"
           
           Case Is = "08" 'Recibos Ingreso
                Printer.Print "Recibos de Ingreso con " & rsReporte!SubGrupo
                'Printer.Print " "
                If chkIngreso.value = False Then
                   Printer.Print "Documentos   Monto MN Monto ME Observacion"
                Else
                   Printer.Print "Documentos           Monto MN     Monto ME"
                End If
                Printer.Print "----------------------------------------"
                
           Case Is = "09" 'Recibos Anticipos
                Printer.Print "Recibos de Anticipo con " & rsReporte!SubGrupo
                'Printer.Print " "
                If chkAnticipo.value = False Then
                   Printer.Print "Documentos   Monto MN Monto ME Observac."
                Else
                   Printer.Print "Documentos         Monto MN     Monto ME"
                End If
                Printer.Print "----------------------------------------"
           
           Case Is = "10" 'Recibos de Egreso
                Printer.Print "Recibos de Egreso " & rsReporte!SubGrupo
                'Printer.Print " "
                If chkEgreso.value = False Then
                   Printer.Print "Documentos   Monto MN Monto ME Observac."
                Else
                   Printer.Print "Documentos         Monto MN     Monto ME"
                End If
                Printer.Print "----------------------------------------"
                
           Case Is = "11" 'Cortesias
                Printer.Print "Cortesias"
                If chkCortesia.value = False Then
                   Printer.Print "Documentos            Monto Descripcion"
                   Printer.Print "---------------------------------------"
                Else
                   Printer.Print "Documentos         Monto MN"
                   Printer.Print "---------------------------------------"
                End If
                                                                          
           Case Is = "12" ' Cuentas corrientes
                Printer.Print "Cuentas Corrientes"
                If chkCuenta.value = False Then
                   Printer.Print "Documentos          Monto MN"
                Else
                   Printer.Print "Documentos          Monto MN"
                End If
                Printer.Print "----------------------------------------"
                                                                          
                
           Case Is = "13" ' Anulados
                Printer.Print "Documentos Anulados"
                If chkAnulado.value = False Then
                   Printer.Print "Documentos        Observacion"
                Else
                   Printer.Print "Documentos"
                End If
                Printer.Print "----------------------------------------"
                
           Case Is = "14" ' Facturado en otro Rango
                Printer.Print "Facturados de otro Rango"
                If chkPedidos.value = False Then
                   Printer.Print "Pedidos           Monto Documento"
                Else
                   Printer.Print "Pedidos            Monto MN"
                End If
                Printer.Print "--------------------------------------"
                
          Case Is = "16" ' Consumos facturados
                Printer.Print "Consumos Facturados"
                Printer.Print "Documentos         Monto MN    Monto ME"
                Printer.Print "----------------------------------------"
                
        Case Is = "17" ' Notas de creditos
            If lNcOfisis Then
                Printer.Print "Notas de Credito"
                Printer.Print "Documentos         Monto MN    Monto ME"
                Printer.Print "----------------------------------------"
            End If
            
    End Select
End Sub

Public Sub Detalle(Grupo As String, SubGrupo As String)
    xParte1 = ""
    xParte2 = ""
    xParte3 = ""
    xParte4 = ""
    sLinea = ""
    
    Select Case Grupo
           Case Is = "00" 'Documentos Generado
                If chkGenerado.value = False Then
                   xParte1 = Left(rsReporte!tDocumento & String(17, " "), 17)
                   xParte2 = Right(String(10, " ") & Format(rsReporte!nVenta1, "###,##0.00"), 10)
                   xParte3 = Left(rsReporte!tUsuario & String(7, " "), 7)
                   xParte4 = Left(rsReporte!tObservacion & String(3, " "), 3)
                Else
                   xParte1 = Right(String(10, " ") & Format(rsReporte!tObservacion, "#,##0.00"), 10) & "    "
                   xParte2 = Right(String(12, " ") & Format(rsReporte!nVenta1, "###,##0.00"), 12)
                   xParte3 = ""
                   xParte4 = ""
                End If
                
           Case Is = "01"  'Efectivos
                If chkEfectivo.value = False Then
                   xParte1 = Left(rsReporte!tDocumento & String(17, " "), 17)
                   xParte2 = Right(String(10, " ") & Format(rsReporte!nVenta1, "###,##0.00"), 10)
                   xParte3 = Right(String(10, " ") & Format(rsReporte!nVenta2, "###,##0.00"), 10)
                   xParte4 = ""
                Else
                   xParte1 = Right(String(10, " ") & Format(rsReporte!tObservacion, "#,##0.00"), 10) & "    "
                   xParte2 = Right(String(12, " ") & Format(rsReporte!nVenta1, "###,##0.00"), 12)
                   xParte3 = Right(String(12, " ") & Format(rsReporte!nVenta2, "###,##0.00"), 12)
                   xParte4 = ""
                End If
                
           Case Is = "02" 'TARJETA
                If chkTarjeta.value = False Then
                   xParte1 = Left(rsReporte!tDocumento & String(17, " "), 17)
                   xParte2 = Right(String(8, " ") & Format(rsReporte!nVenta1, "#,##0.00"), 8)
                   xParte3 = Right(String(8, " ") & Format(rsReporte!nVenta2, "#,##0.00"), 8)
                   xParte4 = Left(Mid(rsReporte!tObservacion, 29) & String(4, " "), 4)
                Else
                   xParte1 = Right(String(10, " ") & Format(rsReporte!tObservacion, "#,##0.00"), 10) & "    "
                   xParte2 = Right(String(12, " ") & Format(rsReporte!nVenta1, "###,##0.00"), 12)
                   xParte3 = Right(String(12, " ") & Format(rsReporte!nVenta2, "###,##0.00"), 12)
                   xParte4 = ""
                End If
           
           Case Is = "03" 'Cheques"
                If chkCheque.value = False Then
                   xParte1 = Left(rsReporte!tDocumento & String(17, " "), 17)
                   xParte2 = Right(String(8, " ") & Format(rsReporte!nVenta1, "#,##0.00"), 8)
                   xParte3 = Right(String(8, " ") & Format(rsReporte!nVenta2, "#,##0.00"), 8)
                   xParte4 = Left(Mid(rsReporte!tObservacion, 29) & String(4, " "), 4)
                Else
                   xParte1 = Right(String(10, " ") & Format(rsReporte!tObservacion, "#,##0.00"), 10) & "    "
                   xParte2 = Right(String(12, " ") & Format(rsReporte!nVenta1, "###,##0.00"), 12)
                   xParte3 = Right(String(12, " ") & Format(rsReporte!nVenta2, "###,##0.00"), 12)
                   xParte4 = ""
                End If
                                     
           Case Is = "04" 'Otros
                If chkOtro.value = False Then
                   xParte1 = Left(rsReporte!tDocumento & String(18, " "), 18)
                   xParte2 = Right(String(9, " ") & Format(rsReporte!nVenta1, "##,##0.00"), 9)
                   xParte3 = Right(String(9, " ") & Format(rsReporte!nVenta2, "##,##0.00"), 9)
                   xParte4 = ""
                Else
                   xParte1 = Right(String(10, " ") & Format(rsReporte!tObservacion, "#,##0.00"), 10) & "    "
                   xParte2 = Right(String(12, " ") & Format(rsReporte!nVenta1, "###,##0.00"), 12)
                   xParte3 = Right(String(12, " ") & Format(rsReporte!nVenta2, "###,##0.00"), 12)
                   xParte4 = ""
                End If
                                     
           Case Is = "05" 'Puntos
                If chkPunto.value = False Then
                   xParte1 = Left(rsReporte!tDocumento & String(18, " "), 18)
                   xParte2 = Right(String(9, " ") & Format(rsReporte!nVenta1, "##,##0.00"), 9)
                   xParte3 = Left(Mid(rsReporte!tObservacion, 33) & String(12, " "), 12)
                   xParte4 = ""
                Else
                   xParte1 = Right(String(10, " ") & Format(rsReporte!tObservacion, "#,##0.00"), 10) & "    "
                   xParte2 = Right(String(12, " ") & Format(rsReporte!nVenta1, "###,##0.00"), 12)
                   xParte3 = ""
                   xParte4 = ""
                End If
                                     
           Case Is = "06" 'Cancelacion de Cortesias
                If chkPagoCortes�a.value = False Then
                   xParte1 = Left(rsReporte!tDocumento & String(18, " "), 18)
                   xParte2 = Right(String(9, " ") & Format(rsReporte!nVenta1, "##,##0.00"), 9)
                   xParte3 = Left(Mid(rsReporte!tObservacion, 33) & String(12, " "), 12)
                   xParte4 = ""
                Else
                   xParte1 = Right(String(10, " ") & Format(rsReporte!tObservacion, "#,##0.00"), 10) & "    "
                   xParte2 = Right(String(12, " ") & Format(rsReporte!nVenta1, "###,##0.00"), 12)
                   xParte3 = ""
                   xParte4 = ""
                End If
                                     
           Case Is = "07" ' Cuentas por cobrar
                If chkCobrar.value = False Then
                   xParte1 = Left(rsReporte!tDocumento & String(18, " "), 18)
                   xParte2 = Right(String(9, " ") & Format(rsReporte!nVenta1, "##,##0.00"), 9)
                   xParte3 = Left(rsReporte!tObservacion & String(12, " "), 12)
                   xParte4 = ""
                Else
                   xParte1 = Right(String(10, " ") & Format(rsReporte!tObservacion, "#,##0.00"), 10) & "    "
                   xParte2 = Right(String(12, " ") & Format(rsReporte!nVenta1, "###,##0.00"), 12)
                   xParte3 = ""
                   xParte4 = ""
                End If
                                     
           Case Is = "08" 'Recibos Ingreso
                If chkIngreso.value = False Then
                   xParte1 = Left(rsReporte!tDocumento & String(10, " "), 10)
                   xParte2 = Right(String(8, " ") & Format(rsReporte!nVenta1, "#,##0.00"), 8)
                   xParte3 = Right(String(8, " ") & Format(rsReporte!nVenta2, "#,##0.00"), 8)
                   xParte4 = Left(Mid(rsReporte!tObservacion, 29) & String(10, " "), 10)
                 Else
                   xParte1 = Right(String(10, " ") & Format(rsReporte!tObservacion, "#,##0.00"), 10) & "    "
                   xParte2 = Right(String(12, " ") & Format(rsReporte!nVenta1, "###,##0.00"), 12)
                   xParte3 = Right(String(12, " ") & Format(rsReporte!nVenta2, "###,##0.00"), 12)
                   xParte4 = ""
                End If
                
           Case Is = "09" 'Recibos Anticipos
                If chkAnticipo.value = False Then
                   xParte1 = Left(rsReporte!tDocumento & String(10, " "), 10)
                   xParte2 = Right(String(9, " ") & Format(rsReporte!nVenta1, "#,##0.00"), 9)
                   xParte3 = Right(String(8, " ") & Format(rsReporte!nVenta2, "#,##0.00"), 8)
                   xParte4 = Left(Mid(rsReporte!tObservacion, 29) & String(9, " "), 9)
                 Else
                   xParte1 = Right(String(10, " ") & Format(rsReporte!tObservacion, "#,##0.00"), 10) & "    "
                   xParte2 = Right(String(12, " ") & Format(rsReporte!nVenta1, "###,##0.00"), 12)
                   xParte3 = Right(String(12, " ") & Format(rsReporte!nVenta2, "###,##0.00"), 12)
                   xParte4 = ""
                End If
                                     
           Case Is = "10" 'Recibos de Egreso
                If chkEgreso.value = False Then
                   xParte1 = Left(rsReporte!tDocumento & String(10, " "), 10)
                   xParte2 = Right(String(9, " ") & Format(rsReporte!nVenta1, "#,##0.00"), 9)
                   xParte3 = Right(String(8, " ") & Format(rsReporte!nVenta2, "#,##0.00"), 8)
                   xParte4 = Left(Mid(rsReporte!tObservacion, 29) & String(9, " "), 9)
                Else
                   xParte1 = Right(String(10, " ") & Format(rsReporte!tObservacion, "#,##0.00"), 10) & "    "
                   xParte2 = Right(String(12, " ") & Format(rsReporte!nVenta1, "###,##0.00"), 12)
                   xParte3 = Right(String(12, " ") & Format(rsReporte!nVenta2, "###,##0.00"), 12)
                   xParte4 = ""
                End If
                                                                          
           Case Is = "11" 'Cortesias
                If chkCortesia.value = False Then
                   xParte1 = Left(rsReporte!tDocumento & String(16, " "), 16)
                   xParte2 = Right(String(10, " ") & Format(rsReporte!nVenta1, "##,##0.00"), 10)
                   xParte3 = Mid(rsReporte!tObservacion, 12, 13)
                   xParte4 = ""
                Else
                   xParte1 = Right(String(10, " ") & Format(rsReporte!tObservacion, "#,##0.00"), 10) & "    "
                   xParte2 = Right(String(12, " ") & Format(rsReporte!nVenta1, "###,##0.00"), 12)
                   xParte3 = ""
                   xParte4 = ""
                End If
                                                                          
           Case Is = "12" ' Cuentas corrientes
                If chkCuenta.value = False Then
                   xParte1 = Left(rsReporte!tDocumento & String(16, " "), 16)
                   xParte2 = Right(String(10, " ") & Format(rsReporte!nVenta1, "###,##0.00"), 10)
                   xParte3 = Left(rsReporte!tObservacion & String(10, " "), 10)
                   xParte4 = ""
                Else
                   xParte1 = Right(String(10, " ") & Format(rsReporte!tObservacion, "#,##0.00"), 10) & "    "
                   xParte2 = Right(String(12, " ") & Format(rsReporte!nVenta1, "###,##0.00"), 12)
                   xParte3 = ""
                   xParte4 = ""
                End If
                                                                                                                                                    
           Case Is = "13" ' Anulados
                If chkAnulado.value = False Then
                   xParte1 = Left(rsReporte!tDocumento & String(17, " "), 17)
                   xParte2 = Left(rsReporte!tObservacion & String(20, " "), 20)
                   xParte3 = ""
                   xParte4 = ""
                Else
                   xParte1 = Right(String(10, " ") & Format(rsReporte!tObservacion, "#,##0.00"), 10) & "    "
                   xParte2 = ""
                   xParte3 = ""
                   xParte4 = ""
                End If
                
           Case Is = "14" 'Documentos Facturados de otro rango
                If chkPedidos.value = False Then
                   xParte1 = Left(rsReporte!tDocumento & String(15, " "), 15)
                   xParte2 = Right(String(10, " ") & Format(rsReporte!nVenta1, "###,##0.00"), 10)
                   xParte3 = Left(rsReporte!tObservacion & String(11, " "), 11)
                   xParte4 = ""
                Else
                   xParte1 = Right(String(10, " ") & Format(rsReporte!tObservacion, "#,##0.00"), 10) & "    "
                   xParte2 = Right(String(12, " ") & Format(rsReporte!nVenta1, "###,##0.00"), 12)
                   xParte3 = ""
                   xParte4 = ""
                End If
          Case Is = "16" ' consumos facturados
                   xParte1 = Left(rsReporte!tDocumento & String(15, " "), 15)
                   xParte2 = Right(String(10, " ") & Format(rsReporte!nVenta1, "###,##0.00"), 10)
                   xParte3 = Left(rsReporte!tUsuario & String(7, " "), 7)
                   xParte4 = Left(rsReporte!tObservacion & String(3, " "), 3)
                   
'
          Case Is = "17" ' Notas de CRedito
                If lNcOfisis Then
                   xParte1 = Left(rsReporte!tDocumento & String(17, " "), 17)
                   xParte2 = Right(String(10, " ") & Format(rsReporte!nVenta1, "###,##0.00"), 10)
                   xParte3 = Right(String(10, " ") & Format(rsReporte!nVenta2, "###,##0.00"), 10)
                   xParte4 = ""
'                   xParte1 = Left(rsReporte!tDocumento & String(15, " "), 15)
'                   xParte2 = Right(String(10, " ") & Format(rsReporte!nVenta1, "###,##0.00"), 10)
'                   xParte3 = Left(rsReporte!tUsuario & String(7, " "), 7)
'                   xParte4 = Left(rsReporte!tObservacion & String(3, " "), 3)
                End If
    End Select
    If Grupo <> "15" Then
        sLinea = xParte1 & " " & xParte2 & " " & xParte3 & " " & xParte4
        If sLinea <> "" Then
            Printer.Print sLinea
        End If
    End If
End Sub

Public Sub Final(Grupo As String, SubGrupo As String)
    Select Case Grupo
           Case Is = "00" 'Documentos Generado
                If chkGenerado.value = False Then
                   Printer.Print String(40, "-")
                   Printer.Print "          Total : " & Right(String(10, " ") & Format(nVenta1, "###,##0.00"), 10)
                   Printer.Print "     Documentos : " & Right(String(10, " ") & Format(nTotal, "###,##0.00"), 10)
                End If
                
           Case Is = "01"  'Efectivos
                If chkEfectivo.value = False Then
                   Printer.Print String(40, "-")
                   Printer.Print "         Total : " & Right(String(11, " ") & Format(nVenta1, "###,##0.00"), 11) & Right(String(11, " ") & Format(nVenta2, "###,##0.00"), 11)
                   Printer.Print "    Documentos : " & Right(String(11, " ") & Format(nTotal, "###,##0.00"), 11)
                End If
                
           Case Is = "02" 'TARJETA
                If chkTarjeta.value = False Then
                   Printer.Print String(40, "-")
                   Printer.Print "         Total : " & Right(String(9, " ") & Format(nVenta1, "##,##0.00"), 9) & Right(String(9, " ") & Format(nVenta2, "##,##0.00"), 9)
                   Printer.Print "    Documentos : " & Right(String(9, " ") & Format(nTotal, "##,##0.00"), 9)
                End If
                
           Case Is = "03" 'Cheques
                Printer.Print " "
                If chkCheque.value = False Then
                   Printer.Print String(40, "-")
                   Printer.Print "         Total : " & Right(String(9, " ") & Format(nVenta1, "##,##0.00"), 9) & Right(String(9, " ") & Format(nVenta2, "##,##0.00"), 9)
                   Printer.Print "    Documentos : " & Right(String(9, " ") & Format(nTotal, "##,##0.00"), 9)
                End If
                
           Case Is = "04" 'Otros
                If chkOtro.value = False Then
                   Printer.Print String(40, "-")
                   Printer.Print "           Total : " & Right(String(9, " ") & Format(nVenta1, "##,##0.00"), 9) & Right(String(10, " ") & Format(nVenta2, "###,##0.00"), 10)
                   Printer.Print "      Documentos : " & Right(String(9, " ") & Format(nTotal, "##,##0.00"), 9)
                End If
                
           Case Is = "05" 'Puntos
                If chkPunto.value = False Then
                   Printer.Print String(40, "-")
                   Printer.Print "           Total : " & Right(String(9, " ") & Format(nVenta1, "##,##0.00"), 9)
                   Printer.Print "      Documentos : " & Right(String(9, " ") & Format(nTotal, "##,##0.00"), 9)
                End If

           Case Is = "06" 'Cancelacion de Cortesias
                If chkPagoCortes�a.value = False Then
                   Printer.Print String(40, "-")
                   Printer.Print "           Total : " & Right(String(13, " ") & Format(nVenta1, "##,##0.00"), 9)
                   Printer.Print "      Documentos : " & Right(String(13, " ") & Format(nTotal, "##,##0.00"), 9)
                End If
                                                                                                                                                                                      
           Case Is = "07" ' Cuentas por cobrar
                If chkCobrar.value = False Then
                   Printer.Print String(40, "-")
                   Printer.Print "           Total : " & Right(String(10, " ") & Format(nVenta1, "###,##0.00"), 10)
                   Printer.Print "      Documentos : " & Right(String(10, " ") & Format(nTotal, "###,##0.00"), 10)
                End If
                                     
           Case Is = "08" 'Recibos Ingreso
                If chkIngreso.value = False Then
                   Printer.Print String(40, "-")
                   Printer.Print "     Total:" & Right(String(13, " ") & Format(nVenta1, "#,##0.00"), 8) & Right(String(11, " ") & Format(nVenta2, "#,##0.00"), 8)
                   Printer.Print "Documentos:" & Right(String(13, " ") & Format(nTotal, "#,##0.00"), 8)
                End If
                
           Case Is = "09" 'Recibos Anticipos
                If chkAnticipo.value = False Then
                   Printer.Print String(40, "-")
                   Printer.Print "     Total:" & Right(String(7, " ") & Format(nVenta1, "##,##0.00"), 9) & Right(String(12, " ") & Format(nVenta2, "#,##0.00"), 8)
                   Printer.Print "Documentos:" & Right(String(9, " ") & Format(nTotal, "#,##0.00"), 9)
                End If
                
           Case Is = "10" 'Recibos de Egreso
                If chkEgreso.value = False Then
                   Printer.Print String(40, "-")
                   Printer.Print "     Total:" & Right(String(7, " ") & Format(nVenta1, "##,##0.00"), 9) & Right(String(12, " ") & Format(nVenta2, "#,##0.00"), 8)
                   Printer.Print "Documentos:" & Right(String(9, " ") & Format(nTotal, "#,##0.00"), 9)
                End If
                
           Case Is = "11" 'Cortesias
                If chkCortesia.value = False Then
                   Printer.Print String(40, "-")
                   Printer.Print "           Total : " & Right(String(10, " ") & Format(nVenta1, "###,##0.00"), 10)
                   Printer.Print "      Documentos : " & Right(String(10, " ") & Format(nTotal, "###,##0.00"), 10)
                End If
                
           Case Is = "12" ' Cuentas corrientes
                If chkCuenta.value = False Then
                   Printer.Print String(40, "-")
                   Printer.Print "          Total : " & Right(String(10, " ") & Format(nVenta1, "###,##0.00"), 10)
                   Printer.Print "     Documentos : " & Right(String(10, " ") & Format(nTotal, "###,##0.00"), 10)
                End If
                                
           Case Is = "13" ' Anulados
                If chkAnulado.value = False Then
                   Printer.Print String(40, "-")
                   Printer.Print "        Total : " & Right(String(10, " ") & Format(nVenta1, "###,##0.00"), 10)
                   Printer.Print "   Documentos : " & Right(String(10, " ") & Format(nTotal, "###,##0.00"), 10)
                End If
                
           Case Is = "14" ' Facturados en otro rango
                If chkPedidos.value = False Then
                   Printer.Print String(40, "-")
                   Printer.Print "           Total : " & Right(String(10, " ") & Format(nVenta1, "###,##0.00"), 10)
                   Printer.Print "      Documentos : " & Right(String(10, " ") & Format(nTotal, "###,##0.00"), 10)
                End If
                
           Case Is = "16" ' Facturados en otro rango
                   Printer.Print String(40, "-")
                   Printer.Print "           Total : " & Right(String(10, " ") & Format(nVenta1, "###,##0.00"), 10)
                   Printer.Print "      Documentos : " & Right(String(10, " ") & Format(nTotal, "###,##0.00"), 10)
                   
         
            Case Is = "17" ' Notas de Credito
                    If lNcOfisis Then
                        Printer.Print String(40, "-")
                        Printer.Print "           Total : " & Right(String(10, " ") & Format(nVenta1, "###,##0.00"), 10)
                        Printer.Print "      Documentos : " & Right(String(10, " ") & Format(nTotal, "###,##0.00"), 10)
                    End If
    End Select
    Printer.Print " "
End Sub
    
Public Sub Resumen()
   'Imprime en el Reporte
   'Titulos
   Printer.Print " "
   Printer.Print String(40, "-")
   ImprimeXCentro "RESUMEN", 40
   Printer.Print "T/Cambio : " & Format(nCambio, "#,##0.000")
   Printer.Print
   
   'Produccion
   Printer.FontBold = True
   Printer.Print "PRODUCCION     Total     Mon MN   Mon ME"
   Printer.Print String(40, "-")
   Printer.FontBold = False
   
   'Efectivo
   sLinea = "Efectivo :" & Right(String(10, " ") & Format(nEfectivoN + xEfectivoE, "###,##0.00"), 10) & " " & Right(String(10, " ") & Format(nEfectivoN, "###,###,##0.00"), 10) & Right(String(9, " ") & Format(nEfectivoE, "##,##0.00"), 9) & " "
   Printer.Print sLinea
   'Tarjeta
   sLinea = "Tar.Cred.:" & Right(String(10, " ") & Format(nTarjetaT, "###,##0.00"), 10) & " " & Right(String(10, " ") & Format(nTarjetaT, "###,##0.00"), 10)
   Printer.Print sLinea
   'Cheque
   sLinea = "Che/Depos:" & Right(String(10, " ") & Format(nChequeN + xChequeE, "###,##0.00"), 10) & " " & Right(String(12, " ") & Format(nChequeN, "###,##0.00"), 10) & Right(String(9, " ") & Format(nChequeE, "##,##0.00"), 9)
   Printer.Print sLinea
   'Notas de Credito
   If lNcOfisis Then
    sLinea = "Nota Cred:" & Right(String(10, " ") & Format(nNotaCreditoN + xNotaCreditoE, "###,##0.00"), 10) & " " & Right(String(12, " ") & Format(nNotaCreditoN, "###,##0.00"), 10) & Right(String(9, " ") & Format(xNotaCreditoE, "##,##0.00"), 9)
    Printer.Print sLinea
   End If
   'Otros
   sLinea = "Otros    :" & Right(String(10, " ") & Format(nOtroN + xOtroE, "###,##0.00"), 10) & " " & Right(String(12, " ") & Format(nOtroN, "###,##0.00"), 10) & Right(String(9, " ") & Format(nOtroE, "##,##0.00"), 9)
   Printer.Print sLinea
   'Puntos
   sLinea = "Puntos   :" & Right(String(10, " ") & Format(nPunto, "###,##0.00"), 10) & " " & Right(String(10, " ") & Format(nPunto, "###,##0.00"), 10)
   Printer.Print sLinea
   'Cancelacion cortesias
   sLinea = "Transf. Gratuitas:" & Right(String(10, " ") & Format(nCortesia, "###,##0.00"), 10) & " " & Right(String(10, " ") & Format(nCortesia, "###,##0.00"), 10)
   Printer.Print sLinea
         
   'Cobrado
   Printer.Print " "
   sLinea = "Cobrados :" & Right(String(10, " ") & Format(nTotalCobradoN + xTotalCobradoE, "###,##0.00"), 10) & " " & Right(String(10, " ") & Format(nTotalCobradoN, "###,##0.00"), 10) & Right(String(9, " ") & Format(nTotalCobradoE, "##,##0.00"), 9) & " "
   Printer.Print sLinea
   Printer.Print " "
   Printer.Print "X Cobrar :" & Right(String(10, " ") & Format(nCobrar, "##,###,##0.00"), 10)
   Printer.Print " "
   
   'Diferencia
   Printer.Print "Total Doc:" & Right(String(10, " ") & Format(nDocumento, "##,###,##0.00"), 10)
   Printer.Print " "
   Printer.Print "Cob.Otro :" & Right(String(10, " ") & Format(nOtroDoc, "##,###,##0.00"), 10)
   Printer.Print " "
   'Printer.Print "No Cobrad:" & Right(String(10, " ") & Format(nDocumento - (nTotalCobradoN + xTotalCobradoE + nCobrar), "###,##0.00"), 10)
   Printer.Print "No Cobrad:" & Right(String(10, " ") & Format((nNoCobrado + (nDocumento - nTotalCobradoN - xTotalCobradoE - nNoCobrado - nCobrar + nOtroDoc)), "###,##0.00"), 10)
        
   'Documentos por Tipo de Pedido
   Printer.FontBold = True
   Printer.Print " "
   Printer.Print "TOTAL DOCUMENTOS"
   Printer.Print String(40, "-")
   Printer.FontBold = False
   
   Dim netoRv, Impuesto1Rv, Impuesto2Rv, Impuesto3Rv, totalVRv As Double
   Dim netoCF, Impuesto1CF, Impuesto2CF, Impuesto3CF, totalVCF As Double
   
   netoRv = 0: Impuesto1Rv = 0: Impuesto2Rv = 0: Impuesto3Rv = 0: totalVRv = 0
   netoCF = 0: Impuesto1CF = 0: Impuesto2CF = 0: Impuesto3CF = 0: totalVCF = 0
   
   Printer.Print ""
   Printer.Print Right(String(20, " ") & IIf(sBoton1 <> "", sBoton1, "N/D"), 20) & "     Docum  Promedio"
   Printer.Print "Neto     :" & Right(String(9, " ") & Format(nValorVenta(1), "###,##0.00"), 10)
   netoRv = nValorVenta(1)
   If sImpuesto1 <> "" Then
      Printer.Print Mid(sImpuesto1 & String(9, " "), 1, 9) & ":" & Right(String(10, " ") & Format(nImp1(1), "###,##0.00"), 10)
      Impuesto1Rv = nImp1(1)
   End If
   If sImpuesto2 <> "" Then
      Printer.Print Mid(sImpuesto2 & String(9, " "), 1, 9) & ":" & Right(String(10, " ") & Format(nImp2(1), "###,##0.00"), 10)
      Impuesto2Rv = nImp2(1)
   End If
   If sImpuesto3 <> "" Then
      Printer.Print Mid(sImpuesto3 & String(9, " "), 1, 9) & ":" & Right(String(10, " ") & Format(nImp3(1), "###,##0.00"), 10)
      Impuesto3Rv = nImp3(1)
   End If
   
   Printer.Print "Total    :" & Right(String(10, " ") & Format(nVenta(1), "###,##0.00"), 10) & "  " & Right(String(8, " ") & Format(nTotalPromedio(1), "###,##0"), 8) & Right(String(10, " ") & Format(nValorVenta(1) / IIf(nTotalPromedio(1) = 0, 1, nTotalPromedio(1)), "###,##0.00"), 10) & " "
   totalVRv = nVenta(1)
   If sBoton2 <> "" Then
      Printer.Print ""
      Printer.Print Right(String(20, " ") & IIf(sBoton2 <> "", sBoton2, "N/D"), 20) & "     Docum  Promedio"
      Printer.Print "Neto     :" & Right(String(9, " ") & Format(nValorVenta(2), "###,##0.00"), 10)
      netoRv = netoRv + nValorVenta(2)
      If sImpuesto1 <> "" Then
         Printer.Print Mid(sImpuesto1 & String(9, " "), 1, 9) & ":" & Right(String(10, " ") & Format(nImp1(2), "###,##0.00"), 10)
         Impuesto1Rv = Impuesto1Rv + nImp1(2)
      End If
      If sImpuesto2 <> "" Then
         Printer.Print Mid(sImpuesto2 & String(9, " "), 1, 9) & ":" & Right(String(10, " ") & Format(nImp2(2), "###,##0.00"), 10)
         Impuesto2Rv = Impuesto2Rv + nImp2(2)
      End If
      If sImpuesto3 <> "" Then
         Printer.Print Mid(sImpuesto3 & String(9, " "), 1, 9) & ":" & Right(String(10, " ") & Format(nImp3(2), "###,##0.00"), 10)
         Impuesto3Rv = Impuesto3Rv + nImp3(2)
      End If
      Printer.Print "Total    :" & Right(String(10, " ") & Format(nVenta(2), "###,##0.00"), 10) & "  " & Right(String(8, " ") & Format(nTotalPromedio(2), "###,##0"), 8) & Right(String(10, " ") & Format(nValorVenta(2) / IIf(nTotalPromedio(2) = 0, 1, nTotalPromedio(2)), "###,##0.00"), 10) & " "
      totalVRv = totalVRv + nVenta(2)
   End If
      
   If sBoton3 <> "" Then
      Printer.Print ""
      Printer.Print Right(String(20, " ") & IIf(sBoton3 <> "", sBoton3, "N/D"), 20) & "     Docum  Promedio"
      Printer.Print "Neto     :" & Right(String(9, " ") & Format(nValorVenta(3), "###,##0.00"), 10)
      netoRv = netoRv + nValorVenta(3)
      If sImpuesto1 <> "" Then
         Printer.Print Mid(sImpuesto1 & String(9, " "), 1, 9) & ":" & Right(String(10, " ") & Format(nImp1(3), "###,##0.00"), 10)
         Impuesto1Rv = Impuesto1Rv + nImp1(3)
      End If
      If sImpuesto2 <> "" Then
         Printer.Print Mid(sImpuesto2 & String(9, " "), 1, 9) & ":" & Right(String(10, " ") & Format(nImp2(3), "###,##0.00"), 10)
         Impuesto2Rv = Impuesto2Rv + nImp2(3)
      End If
      If sImpuesto3 <> "" Then
         Printer.Print Mid(sImpuesto3 & String(9, " "), 1, 9) & ":" & Right(String(10, " ") & Format(nImp3(3), "###,##0.00"), 10)
         Impuesto3Rv = Impuesto3Rv + nImp3(3)
      End If
      Printer.Print "Total    :" & Right(String(10, " ") & Format(nVenta(3), "###,##0.00"), 10) & "  " & Right(String(8, " ") & Format(nTotalPromedio(3), "###,##0"), 8) & Right(String(10, " ") & Format(nValorVenta(3) / IIf(nTotalPromedio(3) = 0, 1, nTotalPromedio(3)), "###,##0.00"), 10) & " "
      totalVRv = totalVRv + nVenta(3)
   End If
      
   If sBoton4 <> "" Then
      Printer.Print ""
      Printer.Print Right(String(20, " ") & IIf(sBoton4 <> "", sBoton4, "N/D"), 20) & "     Docum  Promedio"
      Printer.Print "Neto     :" & Right(String(9, " ") & Format(nValorVenta(4), "###,##0.00"), 10)
      netoRv = netoRv + nValorVenta(4)
      If sImpuesto1 <> "" Then
         Printer.Print Mid(sImpuesto1 & String(9, " "), 1, 9) & ":" & Right(String(10, " ") & Format(nImp1(4), "###,##0.00"), 10)
         Impuesto1Rv = Impuesto1Rv + nImp1(4)
      End If
      If sImpuesto2 <> "" Then
         Printer.Print Mid(sImpuesto2 & String(9, " "), 1, 9) & ":" & Right(String(10, " ") & Format(nImp2(4), "###,##0.00"), 10)
         Impuesto2Rv = Impuesto2Rv + nImp2(4)
      End If
      If sImpuesto3 <> "" Then
         Printer.Print Mid(sImpuesto3 & String(9, " "), 1, 9) & ":" & Right(String(10, " ") & Format(nImp3(4), "###,##0.00"), 10)
         Impuesto3Rv = Impuesto3Rv + nImp3(4)
      End If
      Printer.Print "Total    :" & Right(String(10, " ") & Format(nVenta(4), "###,##0.00"), 10) & "  " & Right(String(8, " ") & Format(nTotalPromedio(4), "###,##0"), 8) & Right(String(10, " ") & Format(nValorVenta(4) / IIf(nTotalPromedio(4) = 0, 1, nTotalPromedio(4)), "###,##0.00"), 10) & " "
      totalVRv = totalVRv + nVenta(4)
   End If
      
   If sBoton5 <> "" Then
      Printer.Print ""
      Printer.Print Right(String(20, " ") & IIf(sBoton5 <> "", sBoton5, "N/D"), 20) & "     Docum  Promedio"
      Printer.Print "Neto     :" & Right(String(9, " ") & Format(nValorVenta(5), "###,##0.00"), 10)
      netoRv = netoRv + nValorVenta(5)
      If sImpuesto1 <> "" Then
         Printer.Print Mid(sImpuesto1 & String(9, " "), 1, 9) & ":" & Right(String(10, " ") & Format(nImp1(5), "###,##0.00"), 10)
         Impuesto1Rv = Impuesto1Rv + nImp1(5)
      End If
      If sImpuesto2 <> "" Then
         Printer.Print Mid(sImpuesto2 & String(9, " "), 1, 9) & ":" & Right(String(10, " ") & Format(nImp2(5), "###,##0.00"), 10)
         Impuesto2Rv = Impuesto2Rv + nImp2(5)
      End If
      If sImpuesto3 <> "" Then
         Printer.Print Mid(sImpuesto3 & String(9, " "), 1, 9) & ":" & Right(String(10, " ") & Format(nImp3(5), "###,##0.00"), 10)
         Impuesto3Rv = Impuesto3Rv + nImp3(5)
      End If
      Printer.Print "Total    :" & Right(String(10, " ") & Format(nVenta(5), "###,##0.00"), 10) & "  " & Right(String(8, " ") & Format(nTotalPromedio(5), "###,##0"), 8) & Right(String(10, " ") & Format(nValorVenta(5) / IIf(nTotalPromedio(5) = 0, 1, nTotalPromedio(5)), "###,##0.00"), 10) & " "
      totalVRv = totalVRv + nVenta(5)
   End If
   
      Printer.Print ""
      Printer.Print Right(String(20, " ") & ("Cons Facturado"), 20) & "     Docum  Promedio"
      Printer.Print "Neto     :" & Right(String(9, " ") & Format(nValorVenta(6), "###,##0.00"), 10)
      netoCF = nValorVenta(6)
         
      If sImpuesto1 <> "" Then
         Printer.Print Mid(sImpuesto1 & String(9, " "), 1, 9) & ":" & Right(String(10, " ") & Format(nImp1(6), "###,##0.00"), 10)
         Impuesto1CF = nImp1(6)
      End If
      If sImpuesto2 <> "" Then
         Printer.Print Mid(sImpuesto2 & String(9, " "), 1, 9) & ":" & Right(String(10, " ") & Format(nImp2(6), "###,##0.00"), 10)
         Impuesto2CF = nImp2(6)
      End If
      If sImpuesto3 <> "" Then
         Printer.Print Mid(sImpuesto3 & String(9, " "), 1, 9) & ":" & Right(String(10, " ") & Format(nImp3(6), "###,##0.00"), 10)
         Impuesto3CF = nImp3(6)
      End If
      Printer.Print "Total    :" & Right(String(10, " ") & Format(nVenta(6), "###,##0.00"), 10) & "  " & Right(String(8, " ") & Format(nTotalPromedio(6), "###,##0"), 8) & Right(String(10, " ") & Format(nValorVenta(6) / IIf(nTotalPromedio(6) = 0, 1, nTotalPromedio(6)), "###,##0.00"), 10) & " "
      totalVCF = nVenta(6)
  
  Printer.Print " "
  Printer.Print " "
  Printer.Print "Total Reg Venta"
  Printer.Print String(40, "-")
  Printer.Print "Neto     :" & Right(String(9, " ") & Format(netoRv, "###,##0.00"), 10)
      If sImpuesto1 <> "" Then
         Printer.Print Mid(sImpuesto1 & String(9, " "), 1, 9) & ":" & Right(String(10, " ") & Format(Impuesto1Rv, "###,##0.00"), 10)
         Impuesto1CF = nImp1(6)
      End If
      If sImpuesto2 <> "" Then
         Printer.Print Mid(sImpuesto2 & String(9, " "), 1, 9) & ":" & Right(String(10, " ") & Format(Impuesto2Rv, "###,##0.00"), 10)
         Impuesto2CF = nImp2(6)
      End If
      If sImpuesto3 <> "" Then
         Printer.Print Mid(sImpuesto3 & String(9, " "), 1, 9) & ":" & Right(String(10, " ") & Format(Impuesto3Rv, "###,##0.00"), 10)
         Impuesto3CF = nImp3(6)
      End If
  Printer.Print "Total     :" & Right(String(9, " ") & Format(totalVRv, "###,##0.00"), 10)
  Printer.Print " "
  Printer.Print "Total Cons Facturado"
   Printer.Print String(40, "-")
   Printer.Print "Neto     :" & Right(String(9, " ") & Format(netoCF, "###,##0.00"), 10)
      If sImpuesto1 <> "" Then
         Printer.Print Mid(sImpuesto1 & String(9, " "), 1, 9) & ":" & Right(String(10, " ") & Format(Impuesto1CF, "###,##0.00"), 10)
         Impuesto1CF = nImp1(6)
      End If
      If sImpuesto2 <> "" Then
         Printer.Print Mid(sImpuesto2 & String(9, " "), 1, 9) & ":" & Right(String(10, " ") & Format(Impuesto2CF, "###,##0.00"), 10)
         Impuesto2CF = nImp2(6)
      End If
      If sImpuesto3 <> "" Then
         Printer.Print Mid(sImpuesto3 & String(9, " "), 1, 9) & ":" & Right(String(10, " ") & Format(Impuesto3CF, "###,##0.00"), 10)
         Impuesto3CF = nImp3(6)
      End If
   Printer.Print "Total     :" & Right(String(9, " ") & Format(totalVCF, "###,##0.00"), 10)
   
   Printer.Print " "
   'Tarjetas
   Printer.FontBold = True
   Printer.Print " "
   Printer.Print "TARJ.CREDITO   Monto    Propina  Recibos"
   Printer.Print String(40, "-")
   Printer.FontBold = False
      
   If tTarjeta(1) <> "no definido" Then
      Printer.Print Mid(tTarjeta(1) & String(10, " "), 1, 10) & Right(String(10, " ") & Format(nTarjeta(1), "###,##0.00"), 10) & " " & Right(String(10, " ") & Format(nPropina(1), "###,##0.00"), 10) & Right(String(9, " ") & Format(nRecibo(1), "##,##0.00"), 9)
   End If
   
   If tTarjeta(2) <> "no definido" Then
      Printer.Print Mid(tTarjeta(2) & String(10, " "), 1, 10) & Right(String(10, " ") & Format(nTarjeta(2), "###,##0.00"), 10) & " " & Right(String(10, " ") & Format(nPropina(2), "###,##0.00"), 10) & Right(String(9, " ") & Format(nRecibo(2), "##,##0.00"), 9)
   End If
   
   If tTarjeta(3) <> "no definido" Then
      Printer.Print Mid(tTarjeta(3) & String(10, " "), 1, 10) & Right(String(10, " ") & Format(nTarjeta(3), "###,##0.00"), 10) & " " & Right(String(10, " ") & Format(nPropina(3), "###,##0.00"), 10) & Right(String(9, " ") & Format(nRecibo(3), "##,##0.00"), 9)
   End If
   
   If tTarjeta(4) <> "no definido" Then
      Printer.Print Mid(tTarjeta(4) & String(10, " "), 1, 10) & Right(String(10, " ") & Format(nTarjeta(4), "###,##0.00"), 10) & " " & Right(String(10, " ") & Format(nPropina(4), "###,##0.00"), 10) & Right(String(9, " ") & Format(nRecibo(4), "##,##0.00"), 9)
   End If
   
   If tTarjeta(5) <> "no definido" Then
      Printer.Print Mid(tTarjeta(5) & String(10, " "), 1, 10) & Right(String(10, " ") & Format(nTarjeta(5), "###,##0.00"), 10) & " " & Right(String(10, " ") & Format(nPropina(5), "###,##0.00"), 10) & Right(String(9, " ") & Format(nRecibo(5), "##,##0.00"), 9)
   End If
   
   If tTarjeta(6) <> "no definido" Then
      Printer.Print Mid(tTarjeta(6) & String(10, " "), 1, 10) & Right(String(10, " ") & Format(nTarjeta(6), "###,##0.00"), 10) & " " & Right(String(10, " ") & Format(nPropina(6), "###,##0.00"), 10) & Right(String(9, " ") & Format(nRecibo(6), "##,##0.00"), 9)
   End If
   
   If tTarjeta(7) <> "no definido" Then
      Printer.Print Mid(tTarjeta(7) & String(10, " "), 1, 10) & Right(String(10, " ") & Format(nTarjeta(7), "###,##0.00"), 10) & " " & Right(String(10, " ") & Format(nPropina(7), "###,##0.00"), 10) & Right(String(9, " ") & Format(nRecibo(7), "##,##0.00"), 9)
   End If
   
   If tTarjeta(8) <> "no definido" Then
      Printer.Print Mid(tTarjeta(8) & String(10, " "), 1, 10) & Right(String(10, " ") & Format(nTarjeta(8), "###,##0.00"), 10) & " " & Right(String(10, " ") & Format(nPropina(8), "###,##0.00"), 10) & Right(String(9, " ") & Format(nRecibo(8), "##,##0.00"), 9)
   End If
   
   Printer.Print " "
   Printer.Print "Total    :" & Right(String(10, " ") & Format(nTarjetaT, "###,##0.00"), 10) & " " & Right(String(10, " ") & Format(nPropinaT, "###,##0.00"), 10) & Right(String(9, " ") & Format(nReciboT, "##,##0.00"), 9)
   Printer.Print String(40, "-")
      
   'Otro tipo Pago
   Printer.FontBold = True
   Printer.Print " "
   Printer.Print "OTRO TIPO PAGO Mon MN   Mon ME   Vuelto"
   Printer.Print String(40, "-")
   Printer.FontBold = False
      
   If tOtro(1) <> "no definido" Then
      Printer.Print Mid(tOtro(1) & String(10, " "), 1, 10) & Right(String(10, " ") & Format(nOtro(1), "###,##0.00"), 10) & Right(String(10, " ") & Format(xOtro(1), "###,##0.00"), 10) & Right(String(10, " ") & Format(nDiferencia(1), "###,##0.00"), 10)
   End If
   
   If tOtro(2) <> "no definido" Then
      Printer.Print Mid(tOtro(2) & String(10, " "), 1, 10) & Right(String(10, " ") & Format(nOtro(2), "###,##0.00"), 10) & Right(String(10, " ") & Format(xOtro(2), "###,##0.00"), 10) & Right(String(10, " ") & Format(nDiferencia(2), "###,##0.00"), 10)
   End If
   
   If tOtro(3) <> "no definido" Then
      Printer.Print Mid(tOtro(3) & String(10, " "), 1, 10) & Right(String(10, " ") & Format(nOtro(3), "###,##0.00"), 10) & Right(String(10, " ") & Format(xOtro(3), "###,##0.00"), 10) & Right(String(10, " ") & Format(nDiferencia(3), "###,##0.00"), 10)
   End If
   
   If tOtro(4) <> "no definido" Then
      Printer.Print Mid(tOtro(4) & String(10, " "), 1, 10) & Right(String(10, " ") & Format(nOtro(4), "###,##0.00"), 10) & Right(String(10, " ") & Format(xOtro(4), "###,##0.00"), 10) & Right(String(10, " ") & Format(nDiferencia(4), "###,##0.00"), 10)
   End If
   
   If tOtro(5) <> "no definido" Then
      Printer.Print Mid(tOtro(5) & String(10, " "), 1, 10) & Right(String(10, " ") & Format(nOtro(5), "###,##0.00"), 10) & Right(String(10, " ") & Format(xOtro(5), "###,##0.00"), 10) & Right(String(10, " ") & Format(nDiferencia(5), "###,##0.00"), 10)
   End If
   
   If tOtro(6) <> "no definido" Then
      Printer.Print Mid(tOtro(6) & String(10, " "), 1, 10) & Right(String(10, " ") & Format(nOtro(6), "###,##0.00"), 10) & Right(String(10, " ") & Format(xOtro(6), "###,##0.00"), 10) & Right(String(10, " ") & Format(nDiferencia(6), "###,##0.00"), 10)
   End If
   
   If tOtro(7) <> "no definido" Then
      Printer.Print Mid(tOtro(7) & String(10, " "), 1, 10) & Right(String(10, " ") & Format(nOtro(7), "###,##0.00"), 10) & Right(String(10, " ") & Format(xOtro(7), "###,##0.00"), 10) & Right(String(10, " ") & Format(nDiferencia(7), "###,##0.00"), 10)
   End If
   
   If tOtro(8) <> "no definido" Then
      Printer.Print Mid(tOtro(8) & String(10, " "), 1, 10) & Right(String(10, " ") & Format(nOtro(8), "###,##0.00"), 10) & Right(String(10, " ") & Format(xOtro(8), "###,##0.00"), 10) & Right(String(10, " ") & Format(nDiferencia(8), "###,##0.00"), 10)
   End If
   
   If tOtro(9) <> "no definido" Then
      Printer.Print Mid(tOtro(9) & String(10, " "), 1, 10) & Right(String(10, " ") & Format(nOtro(9), "###,##0.00"), 10) & Right(String(10, " ") & Format(xOtro(9), "###,##0.00"), 10) & Right(String(10, " ") & Format(nDiferencia(9), "###,##0.00"), 10)
   End If
   
   If tOtro(10) <> "no definido" Then
      Printer.Print Mid(tOtro(10) & String(10, " "), 1, 10) & Right(String(10, " ") & Format(nOtro(10), "###,##0.00"), 10) & Right(String(10, " ") & Format(xOtro(10), "###,##0.00"), 10) & Right(String(10, " ") & Format(nDiferencia(10), "###,##0.00"), 10)
   End If
   
   If tOtro(11) <> "no definido" Then
      Printer.Print Mid(tOtro(11) & String(10, " "), 1, 10) & Right(String(10, " ") & Format(nOtro(11), "###,##0.00"), 10) & Right(String(10, " ") & Format(xOtro(11), "###,##0.00"), 10) & Right(String(10, " ") & Format(nDiferencia(11), "###,##0.00"), 10)
   End If
   
   If tOtro(12) <> "no definido" Then
      Printer.Print Mid(tOtro(12) & String(10, " "), 1, 10) & Right(String(10, " ") & Format(nOtro(12), "###,##0.00"), 10) & Right(String(10, " ") & Format(xOtro(12), "###,##0.00"), 10) & Right(String(10, " ") & Format(nDiferencia(12), "###,##0.00"), 10)
   End If
   
   If tOtro(13) <> "no definido" Then
      Printer.Print Mid(tOtro(13) & String(10, " "), 1, 10) & Right(String(10, " ") & Format(nOtro(13), "###,##0.00"), 10) & Right(String(10, " ") & Format(xOtro(13), "###,##0.00"), 10) & Right(String(10, " ") & Format(nDiferencia(13), "###,##0.00"), 10)
   End If
   
   If tOtro(14) <> "no definido" Then
      Printer.Print Mid(tOtro(14) & String(10, " "), 1, 10) & Right(String(10, " ") & Format(nOtro(14), "###,##0.00"), 10) & Right(String(10, " ") & Format(xOtro(14), "###,##0.00"), 10) & Right(String(10, " ") & Format(nDiferencia(14), "###,##0.00"), 10)
   End If
   
   If tOtro(15) <> "no definido" Then
      Printer.Print Mid(tOtro(15) & String(10, " "), 1, 10) & Right(String(10, " ") & Format(nOtro(15), "###,##0.00"), 10) & Right(String(10, " ") & Format(xOtro(15), "###,##0.00"), 10) & Right(String(10, " ") & Format(nDiferencia(15), "###,##0.00"), 10)
   End If
   
   If tOtro(16) <> "no definido" Then
      Printer.Print Mid(tOtro(16) & String(10, " "), 1, 10) & Right(String(10, " ") & Format(nOtro(16), "###,##0.00"), 10) & Right(String(10, " ") & Format(xOtro(16), "###,##0.00"), 10) & Right(String(10, " ") & Format(nDiferencia(16), "###,##0.00"), 10)
   End If
   
   If tOtro(17) <> "no definido" Then
      Printer.Print Mid(tOtro(17) & String(10, " "), 1, 10) & Right(String(10, " ") & Format(nOtro(17), "###,##0.00"), 10) & Right(String(10, " ") & Format(xOtro(17), "###,##0.00"), 10) & Right(String(10, " ") & Format(nDiferencia(17), "###,##0.00"), 10)
   End If
   
   If tOtro(18) <> "no definido" Then
      Printer.Print Mid(tOtro(18) & String(10, " "), 1, 10) & Right(String(10, " ") & Format(nOtro(18), "###,##0.00"), 10) & Right(String(10, " ") & Format(xOtro(18), "###,##0.00"), 10) & Right(String(10, " ") & Format(nDiferencia(18), "###,##0.00"), 10)
   End If
   
   If tOtro(19) <> "no definido" Then
      Printer.Print Mid(tOtro(19) & String(10, " "), 1, 10) & Right(String(10, " ") & Format(nOtro(19), "###,##0.00"), 10) & Right(String(10, " ") & Format(xOtro(19), "###,##0.00"), 10) & Right(String(10, " ") & Format(nDiferencia(19), "###,##0.00"), 10)
   End If
   
   If tOtro(20) <> "no definido" Then
      Printer.Print Mid(tOtro(20) & String(10, " "), 1, 10) & Right(String(10, " ") & Format(nOtro(20), "###,##0.00"), 10) & Right(String(10, " ") & Format(xOtro(20), "###,##0.00"), 10) & Right(String(10, " ") & Format(nDiferencia(20), "###,##0.00"), 10)
   End If
   
   Printer.Print " "
   Printer.Print "Total    :" & Right(String(10, " ") & Format(nOtroN, "###,##0.00"), 10) & Right(String(10, " ") & Format(nOtroE, "###,##0.00"), 10) & Right(String(10, " ") & Format(nTotalDiferencia, "###,##0.00"), 10)
   Printer.Print String(40, "-")
      
   'Recibos
   Printer.FontBold = True
   Printer.Print " "
   Printer.Print "RECIBOS        Total     Mon MN   Mon ME"
   Printer.Print String(40, "-")
   Printer.FontBold = False
      
   sLinea = "Ingresos :" & Right(String(10, " ") & Format(nIngresoN + xIngresoE, "###,##0.00"), 10) & " " & Right(String(10, " ") & Format(nIngresoN, "###,##0.00"), 10) & Right(String(9, " ") & Format(nIngresoE, "##,##0.00"), 9)
   Printer.Print sLinea
   
   sLinea = "Anticipos:" & Right(String(10, " ") & Format(nIngresoAN + xIngresoAE, "###,##0.00"), 10) & " " & Right(String(10, " ") & Format(nIngresoAN, "###,##0.00"), 10) & Right(String(9, " ") & Format(nIngresoAE, "##,##0.00"), 9)
   Printer.Print sLinea
   
   sLinea = "Egresos  :" & Right(String(10, " ") & Format(nEgresoN + xEgresoE, "###,##0.00"), 10) & " " & Right(String(10, " ") & Format(nEgresoN, "###,##0.00"), 10) & Right(String(12, " ") & Format(nEgresoE, "##,##0.00"), 9)
   Printer.Print sLinea
      
   'Cierre
   Printer.FontBold = True
   Printer.Print " "
   Printer.Print "RESULTADOS     Total     Mon MN   Mon ME"
   Printer.Print String(40, "-")
   Printer.FontBold = False
      
   sLinea = "Cobr+Prop:" & Right(String(10, " ") & Format(nTotalCobrado + nPropinaT, "###,##0.00"), 10)
   Printer.Print sLinea
   sLinea = "Fond.Caja:" & Right(String(10, " ") & Format(nFondoN + (nFondoE * nCambio), "###,##0.00"), 10) & " " & Right(String(10, " ") & Format(nFondoN, "###,##0.00"), 10) & Right(String(9, " ") & Format(nFondoE, "##,##0.00"), 9)
   Printer.Print sLinea
   Printer.Print " "
   sLinea = "Efec.Prod:" & Right(String(10, " ") & Format(nTotalEfectivo, "###,##0.00"), 10) & " " & Right(String(10, " ") & Format(nTotalEfectivoN, "###,##0.00"), 10) & Right(String(9, " ") & Format(nTotalEfectivoE, "##,##0.00"), 9)
   Printer.Print sLinea
   sLinea = "Efec.Fis.:" & Right(String(10, " ") & Format(nFisEfectivo, "###,##0.00"), 10) & " " & Right(String(10, " ") & Format(nFisEfectivoN, "###,##0.00"), 10) & Right(String(9, " ") & Format(nDolar - nEgresoE, "##,##0.00"), 9)
   Printer.Print sLinea
                                    
   'Correlativos
   Printer.FontBold = True
   Printer.Print " "
   Printer.Print "CORRELATIVOS DE DOCUMENTOS"
   Printer.Print String(40, "-")
   Printer.FontBold = False
                        
   For i = 1 To nTotalCorrelativo
       If Doc(i) <> "" Then
          Printer.Print Doc(i)
          Printer.Print "Del     : " & Min(i)
          Printer.Print "Al      : " & Max(i)
          Printer.Print "Emitido : " & Format(Emi(i), "##,##0")
          Printer.Print "Anulado : " & Format(Anu(i), "##,##0")
       End If
       Printer.Print " "
   Next i
      
   Printer.Print " "
   Printer.Print String(40, "-")
   Printer.Print "Imp.Caja: " & sCaja
   Printer.Print " "
End Sub

