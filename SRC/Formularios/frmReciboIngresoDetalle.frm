VERSION 5.00
Begin VB.Form frmReciboIngresoDetalle 
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   5925
   ClientLeft      =   2520
   ClientTop       =   2640
   ClientWidth     =   9900
   Icon            =   "frmReciboIngresoDetalle.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5925
   ScaleWidth      =   9900
   Begin VB.Frame fraDetalle 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   5235
      Left            =   1920
      TabIndex        =   13
      Top             =   0
      Width           =   7995
      Begin VB.Frame frmDocumento 
         BackColor       =   &H008080FF&
         Caption         =   "Documento"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   1095
         Left            =   1200
         TabIndex        =   50
         Top             =   3960
         Visible         =   0   'False
         Width           =   2895
         Begin VB.CommandButton cmdTipoDocumento 
            Height          =   735
            Index           =   1
            Left            =   120
            TabIndex        =   52
            Top             =   240
            Visible         =   0   'False
            Width           =   1215
         End
         Begin VB.CommandButton cmdTipoDocumento 
            Height          =   735
            Index           =   2
            Left            =   1440
            TabIndex        =   51
            Top             =   240
            Visible         =   0   'False
            Width           =   1335
         End
      End
      Begin VB.CommandButton cmdBUscar 
         Caption         =   "Buscar Pedido"
         Height          =   615
         Left            =   5640
         TabIndex        =   49
         Top             =   4320
         Width           =   975
      End
      Begin VB.TextBox txtPedido 
         BackColor       =   &H00E0E0E0&
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
         ForeColor       =   &H000000C0&
         Height          =   285
         Left            =   4200
         Locked          =   -1  'True
         TabIndex        =   48
         TabStop         =   0   'False
         Top             =   4560
         Width           =   1335
      End
      Begin VB.CommandButton cmdOpcion 
         Caption         =   "Si"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Index           =   6
         Left            =   1200
         Style           =   1  'Graphical
         TabIndex        =   44
         Top             =   2340
         Width           =   1170
      End
      Begin VB.TextBox txtReferencia 
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
         Height          =   285
         Left            =   1200
         MaxLength       =   30
         TabIndex        =   41
         Top             =   1080
         Width           =   1575
      End
      Begin VB.Frame Frame1 
         Caption         =   " Monto "
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800080&
         Height          =   2205
         Left            =   2940
         TabIndex        =   29
         Top             =   600
         Width           =   4935
         Begin VB.CommandButton cmdCancelacion 
            Caption         =   "Cheque"
            Height          =   585
            Index           =   11
            Left            =   2460
            TabIndex        =   43
            Top             =   270
            Width           =   1185
         End
         Begin VB.CommandButton cmdReferencia 
            Caption         =   "Referencia"
            Height          =   585
            Left            =   3645
            TabIndex        =   42
            Top             =   270
            Width           =   1185
         End
         Begin VB.CommandButton cmdCancelacion 
            Caption         =   "Efectivo"
            Height          =   585
            Index           =   9
            Left            =   90
            TabIndex        =   39
            Top             =   270
            Width           =   1185
         End
         Begin VB.CommandButton cmdCancelacion 
            Caption         =   "Efectivo"
            Height          =   585
            Index           =   10
            Left            =   1275
            TabIndex        =   38
            Top             =   270
            Width           =   1185
         End
         Begin VB.CommandButton cmdCancelacion 
            Height          =   585
            Index           =   1
            Left            =   90
            TabIndex        =   37
            Top             =   900
            Width           =   1185
         End
         Begin VB.CommandButton cmdCancelacion 
            Height          =   585
            Index           =   2
            Left            =   1275
            TabIndex        =   36
            Top             =   900
            Width           =   1185
         End
         Begin VB.CommandButton cmdCancelacion 
            Height          =   585
            Index           =   3
            Left            =   2460
            TabIndex        =   35
            Top             =   900
            Width           =   1185
         End
         Begin VB.CommandButton cmdCancelacion 
            Height          =   585
            Index           =   4
            Left            =   3645
            TabIndex        =   34
            Top             =   900
            Width           =   1185
         End
         Begin VB.CommandButton cmdCancelacion 
            Height          =   585
            Index           =   5
            Left            =   90
            TabIndex        =   33
            Top             =   1530
            Width           =   1185
         End
         Begin VB.CommandButton cmdCancelacion 
            Height          =   585
            Index           =   6
            Left            =   1275
            TabIndex        =   32
            Top             =   1530
            Width           =   1185
         End
         Begin VB.CommandButton cmdCancelacion 
            Height          =   585
            Index           =   7
            Left            =   2460
            TabIndex        =   31
            Top             =   1530
            Width           =   1185
         End
         Begin VB.CommandButton cmdCancelacion 
            Height          =   585
            Index           =   8
            Left            =   3645
            TabIndex        =   30
            Top             =   1530
            Width           =   1185
         End
      End
      Begin VB.TextBox txtDescripcion 
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
         Height          =   285
         Left            =   1200
         MaxLength       =   100
         TabIndex        =   21
         Text            =   " "
         Top             =   3000
         Width           =   5340
      End
      Begin VB.TextBox txtMonto 
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
         Height          =   285
         Left            =   1200
         MaxLength       =   15
         TabIndex        =   20
         Text            =   " 0.00"
         Top             =   720
         Width           =   1575
      End
      Begin VB.TextBox txtRecibo 
         BackColor       =   &H00E0E0E0&
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
         ForeColor       =   &H000000C0&
         Height          =   285
         Left            =   1200
         Locked          =   -1  'True
         TabIndex        =   19
         TabStop         =   0   'False
         Top             =   240
         Width           =   1575
      End
      Begin VB.TextBox txtTurno 
         BackColor       =   &H00E0E0E0&
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
         ForeColor       =   &H000000C0&
         Height          =   285
         Left            =   1200
         Locked          =   -1  'True
         TabIndex        =   18
         TabStop         =   0   'False
         Top             =   4560
         Width           =   1575
      End
      Begin VB.TextBox txtFecha 
         BackColor       =   &H00E0E0E0&
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
         ForeColor       =   &H000000C0&
         Height          =   285
         Left            =   3600
         Locked          =   -1  'True
         TabIndex        =   17
         TabStop         =   0   'False
         Top             =   240
         Width           =   2655
      End
      Begin VB.TextBox txtAutorizacion 
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
         Height          =   285
         Left            =   1200
         MaxLength       =   50
         TabIndex        =   16
         Text            =   " "
         Top             =   3600
         Width           =   5340
      End
      Begin VB.CommandButton cmdOpcion 
         Caption         =   "Descripci�n"
         Height          =   615
         Index           =   4
         Left            =   6720
         Picture         =   "frmReciboIngresoDetalle.frx":030A
         Style           =   1  'Graphical
         TabIndex        =   15
         Top             =   2880
         Width           =   1170
      End
      Begin VB.CommandButton cmdOpcion 
         Caption         =   "Autorizaci�n"
         Height          =   615
         Index           =   5
         Left            =   6720
         Picture         =   "frmReciboIngresoDetalle.frx":040C
         Style           =   1  'Graphical
         TabIndex        =   14
         Top             =   3600
         Width           =   1170
      End
      Begin VB.Image imagepIE 
         Height          =   375
         Left            =   6840
         Top             =   4800
         Visible         =   0   'False
         Width           =   375
      End
      Begin VB.Image imageCab 
         Height          =   375
         Left            =   6840
         Top             =   4320
         Visible         =   0   'False
         Width           =   375
      End
      Begin VB.Image imageHash 
         Height          =   735
         Left            =   7320
         Top             =   4440
         Visible         =   0   'False
         Width           =   615
      End
      Begin VB.Label Label 
         AutoSize        =   -1  'True
         Caption         =   "Pedido :"
         Height          =   195
         Index           =   8
         Left            =   3480
         TabIndex        =   47
         Top             =   4560
         Width           =   585
      End
      Begin VB.Label Label 
         AutoSize        =   -1  'True
         Caption         =   "Anticipo :"
         Height          =   195
         Index           =   7
         Left            =   420
         TabIndex        =   45
         Top             =   2640
         Width           =   660
      End
      Begin VB.Label Label 
         AutoSize        =   -1  'True
         Caption         =   "Referencia :"
         Height          =   195
         Index           =   3
         Left            =   210
         TabIndex        =   40
         Top             =   1080
         Width           =   870
      End
      Begin VB.Label lblEstado 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Atendido"
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
         Left            =   6360
         TabIndex        =   28
         Top             =   240
         Width           =   1515
      End
      Begin VB.Label Label 
         AutoSize        =   -1  'True
         Caption         =   "Fecha :"
         Height          =   195
         Index           =   1
         Left            =   2880
         TabIndex        =   27
         Top             =   240
         Width           =   540
      End
      Begin VB.Label Label 
         AutoSize        =   -1  'True
         Caption         =   "Descripci�n :"
         Height          =   195
         Index           =   4
         Left            =   150
         TabIndex        =   26
         Top             =   3060
         Width           =   930
      End
      Begin VB.Label Label 
         AutoSize        =   -1  'True
         Caption         =   "Autorizaci�n :"
         Height          =   195
         Index           =   5
         Left            =   120
         TabIndex        =   25
         Top             =   3600
         Width           =   960
      End
      Begin VB.Label Label 
         AutoSize        =   -1  'True
         Caption         =   "Monto :"
         Height          =   195
         Index           =   2
         Left            =   540
         TabIndex        =   24
         Top             =   720
         Width           =   540
      End
      Begin VB.Label Label 
         AutoSize        =   -1  'True
         Caption         =   "Recibo :"
         Height          =   195
         Index           =   0
         Left            =   480
         TabIndex        =   23
         Top             =   240
         Width           =   600
      End
      Begin VB.Label Label 
         AutoSize        =   -1  'True
         Caption         =   "Turno :"
         Height          =   195
         Index           =   6
         Left            =   570
         TabIndex        =   22
         Top             =   4560
         Width           =   510
      End
   End
   Begin VB.PictureBox Picture 
      Align           =   2  'Align Bottom
      Height          =   750
      Left            =   0
      ScaleHeight     =   690
      ScaleWidth      =   9840
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   5175
      Width           =   9900
      Begin VB.CommandButton cmdOpcion 
         Caption         =   "Agregar"
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
         Left            =   5100
         Picture         =   "frmReciboIngresoDetalle.frx":050E
         Style           =   1  'Graphical
         TabIndex        =   9
         Top             =   60
         Width           =   1170
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
         Index           =   3
         Left            =   8610
         Picture         =   "frmReciboIngresoDetalle.frx":0A40
         Style           =   1  'Graphical
         TabIndex        =   11
         Top             =   60
         Width           =   1170
      End
      Begin VB.CommandButton cmdOpcion 
         Caption         =   "Eliminar"
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
         Index           =   2
         Left            =   7440
         Picture         =   "frmReciboIngresoDetalle.frx":0B32
         Style           =   1  'Graphical
         TabIndex        =   10
         Top             =   60
         Width           =   1170
      End
      Begin VB.CommandButton cmdOpcion 
         Caption         =   "Grabar"
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
         Left            =   6270
         Picture         =   "frmReciboIngresoDetalle.frx":0C34
         Style           =   1  'Graphical
         TabIndex        =   8
         Top             =   60
         Width           =   1170
      End
      Begin VB.PictureBox PicNavegacion 
         BackColor       =   &H80000004&
         Height          =   615
         Left            =   30
         ScaleHeight     =   555
         ScaleWidth      =   4920
         TabIndex        =   1
         Top             =   60
         Width           =   4980
         Begin VB.CommandButton cmdNavegar 
            Height          =   555
            Index           =   1
            Left            =   480
            Picture         =   "frmReciboIngresoDetalle.frx":1166
            Style           =   1  'Graphical
            TabIndex        =   7
            TabStop         =   0   'False
            Top             =   0
            Width           =   480
         End
         Begin VB.CommandButton cmdNavegar 
            Height          =   555
            Index           =   2
            Left            =   960
            Picture         =   "frmReciboIngresoDetalle.frx":16A8
            Style           =   1  'Graphical
            TabIndex        =   6
            TabStop         =   0   'False
            Top             =   0
            Width           =   480
         End
         Begin VB.CommandButton cmdNavegar 
            Height          =   555
            Index           =   0
            Left            =   0
            Picture         =   "frmReciboIngresoDetalle.frx":1BEA
            Style           =   1  'Graphical
            TabIndex        =   5
            TabStop         =   0   'False
            Top             =   0
            Width           =   480
         End
         Begin VB.CommandButton cmdNavegar 
            Height          =   555
            Index           =   5
            Left            =   4440
            Picture         =   "frmReciboIngresoDetalle.frx":212C
            Style           =   1  'Graphical
            TabIndex        =   4
            TabStop         =   0   'False
            Top             =   0
            Width           =   480
         End
         Begin VB.CommandButton cmdNavegar 
            Height          =   555
            Index           =   4
            Left            =   3960
            Picture         =   "frmReciboIngresoDetalle.frx":266E
            Style           =   1  'Graphical
            TabIndex        =   3
            TabStop         =   0   'False
            Top             =   0
            Width           =   480
         End
         Begin VB.CommandButton cmdNavegar 
            Height          =   555
            Index           =   3
            Left            =   3480
            Picture         =   "frmReciboIngresoDetalle.frx":2BB0
            Style           =   1  'Graphical
            TabIndex        =   2
            TabStop         =   0   'False
            Top             =   0
            Width           =   480
         End
         Begin VB.Label cmdTexto 
            Alignment       =   2  'Center
            Caption         =   "Registro"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   1470
            TabIndex        =   12
            Top             =   150
            Width           =   1845
         End
      End
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00FFFFFF&
      FillStyle       =   0  'Solid
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   1212
      Left            =   0
      ScaleHeight     =   1185
      ScaleWidth      =   3825
      TabIndex        =   46
      Top             =   5175
      Width           =   3852
   End
   Begin VB.Image Image 
      Height          =   5115
      Left            =   60
      Picture         =   "frmReciboIngresoDetalle.frx":30F2
      Stretch         =   -1  'True
      Top             =   60
      Width           =   1785
   End
End
Attribute VB_Name = "frmReciboIngresoDetalle"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim i As Integer
Dim sEstado As String
Dim nMonto As Double
Dim RsTarjeta As Recordset
Dim sTipoPago As String
Dim sTarjeta As String
Dim sMoneda As String
Dim RsTipoDocumento As Recordset
Dim RsIngreso As Recordset
Public TipoOper As String
Dim oComando As clsComando

Dim sMontoTotalAnticipo As Double

Sub Asignar()
    With frmReciboIngreso.RsCabecera
        'Cuadro de Texto
        For i = 1 To 11
            cmdCancelacion(i).FontBold = False
        Next i
        
        txtRecibo = IIf(IsNull(!tRecibo), "", !tRecibo)
        txtFecha = IIf(IsNull(!fFecha), "", !fFecha)
        nMonto = IIf(IsNull(!nMonto), 0, !nMonto)
        txtMonto.Text = Format(nMonto, "###,###,##0.00")
        txtDescripcion = IIf(IsNull(!tDescripcion), "", !tDescripcion)
        txtReferencia = IIf(IsNull(!tReferencia), "", !tReferencia)
        txtAutorizacion = IIf(IsNull(!tAutoriza), "", !tAutoriza)
        txtTurno = IIf(IsNull(!tTurno), "", !tTurno)
        txtPedido.Text = IIf(IsNull(!Pedido), "", !Pedido)
        sTipoPago = IIf(IsNull(!tTipoPago), "", !tTipoPago)
        sTarjeta = IIf(IsNull(!tTarjeta), "", !tTarjeta)
        sMoneda = IIf(IsNull(!tMoneda), "", !tMoneda)
        cmdOpcion(6).Caption = IIf(IIf(IsNull(!lAnticipo), 0, !lAnticipo), "Si", "No")
                
        Select Case sTipoPago
               Case Is = "01"
                    If sMoneda = "01" Then
                       i = 9
                    Else
                       i = 10
                    End If
               Case Is = "02"
                    i = Val(sTarjeta)
               Case Is = "03"
                    i = 11
        End Select
        
        cmdCancelacion(i).FontBold = True
                        
        lblEstado.ForeColor = IIf(IsNull(!tEstadoDocumento), vbBlack, IIf(!tEstadoDocumento = "04", vbRed, vbBlack))
        sEstado = IIf(IsNull(!tEstadoDocumento), "01", !tEstadoDocumento)
        lblEstado.Caption = IIf(IsNull(!tEstadoDocumento), "", IIf(!tEstadoDocumento = "04", "ANULADO", IIf(!tEstadoDocumento = "02", "PAGADO", "EMITIDO")))
        
        'If Periodo(!fFecha) Or sEstado = "04" Or sEstado = "02" Or sTurno <> !tTurno Then
        If Periodo(!fFecha) Or sEstado = "04" Or sEstado = "02" Then
           Bloquear True
        Else
           Bloquear False
           cmdOpcion(2).Enabled = True
        End If
         If sEstado = "01" Then
            cmdOpcion(1).Enabled = False
        End If
    End With
    cmdTexto.Caption = "Registro " & frmReciboIngreso.RsCabecera.AbsolutePosition & " de " & frmReciboIngreso.RsCabecera.RecordCount
End Sub

Private Sub cmdBuscar_Click()
On Error GoTo fin

               Isql = "select tCodigoPedido as Codigo, isnull((select isnull(tNombre,'') + ' ' + isnull(tApellido,'') from TDELIVERY where tCodigoDelivery=MPEDIDO.tClienteDelivery),'') as Descripcion,isnull( (select tResumido from TMESA where tCodigoMesa= mpedido.tMesa),'Sin Mesa') as Descripcion2  from MPEDIDO where tEstadoPedido='01' and tCaja='" & sCaja & "' and tTurno='" & sTurno & "'"
               
               Call ConfGrilla(3, frmBusquedaRapida.grdGrilla, "Codigo", 2, "Codigo", 1300, 2, 0, "", _
                                                               "Descripcion", 2, "Descripcion", 3400, 0, 0, "dd/mm/yyyy", _
                                                               "Mesa", 2, "Descripcion2", 1200, 0, 0, "")
               frmBusquedaRapida.nPredeterm = 1
               frmBusquedaRapida.Show vbModal
               If wEnter Then
               txtPedido.Text = sCodigo

               Else
                  Exit Sub
               End If


Exit Sub
fin:

End Sub

Private Sub cmdCancelacion_Click(Index As Integer)

   frmNumPad.Show vbModal
   nMonto = IIf(wEnter, Val(sDescrip), nMonto)
   txtMonto.Text = Format(nMonto, "###,##0.00")

   Select Case Index
          Case Is = 1 ' Tarjeta 1
               sTipoPago = "02"
               sTarjeta = "01"
               sMoneda = "01"
               
          Case Is = 2 ' Tarjeta 2
               sTipoPago = "02"
               sTarjeta = "02"
               sMoneda = "01"
          
          Case Is = 3 ' Tarjeta 3
               sTipoPago = "02"
               sTarjeta = "03"
               sMoneda = "01"
               
          Case Is = 4 ' Tarjeta 4
               sTipoPago = "02"
               sTarjeta = "04"
               sMoneda = "01"
               
          Case Is = 5 ' Tarjeta 5
               sTipoPago = "02"
               sTarjeta = "05"
               sMoneda = "01"
               
          Case Is = 6 ' Tarjeta 6
               sTipoPago = "02"
               sTarjeta = "06"
               sMoneda = "01"
               
          Case Is = 7 ' Tarjeta 7
               sTipoPago = "02"
               sTarjeta = "07"
               sMoneda = "01"
               
          Case Is = 8 ' Tarjeta 8
               sTipoPago = "02"
               sTarjeta = "08"
               sMoneda = "01"
               
          Case Is = 9 ' Efectivo MN
               sTipoPago = "01"
               sTarjeta = ""
               sMoneda = "01"
               
          Case Is = 10 ' Efectivo ME
               sTipoPago = "01"
               sTarjeta = ""
               sMoneda = "02"
               
          Case Is = 11 ' Cheque
               sTipoPago = "03"
               sTarjeta = ""
               sMoneda = "01"
               
   End Select
End Sub

Private Sub cmdNavegar_Click(Index As Integer)
    Select Case Index
           Case Is = 0 'Primero
                MoverPuntero Primero, frmReciboIngreso.grdGrilla
           Case Is = 1 'PgUp
                MoverPuntero pgup, frmReciboIngreso.grdGrilla
           Case Is = 2 'Previo
                MoverPuntero previo, frmReciboIngreso.grdGrilla
           Case Is = 3 'Siguiente
                MoverPuntero siguiente, frmReciboIngreso.grdGrilla
           Case Is = 4 'PgDn
                MoverPuntero pgdn, frmReciboIngreso.grdGrilla
           Case Is = 5 'Ultimo
                MoverPuntero Ultimo, frmReciboIngreso.grdGrilla
    End Select
   Asignar

End Sub

Private Sub cmdOpcion_Click(Index As Integer)
   Select Case Index
          Case Is = 0 'Agregar
               Sw = True
               ActivarBotones (False)
               Blanquear Me
               
               'Cambia el Nombre del Primer Text
               txtTurno.Text = sTurno
               txtFecha.Text = FechaServidor()
               lblEstado.Caption = ""
               sMoneda = "01"
               sTipoPago = "01"
               sTarjeta = ""
               cmdOpcion(6).Caption = "Si"
               For i = 1 To 11
                   cmdCancelacion(i).FontBold = False
               Next i
               Bloquear False
               cmdBuscar.Enabled = True
          Case Is = 1 'Grabar
               'Chequea Datos
               Dim nCorrela As String
               Dim nPos As Integer
               If txtMonto.Text = "" Then MsgBox "Ingrese el Monto del Recibo", vbExclamation, sMensaje: Exit Sub
               If txtDescripcion.Text = "" Then MsgBox "Ingrese la Descripci�n del Recibo", vbExclamation, sMensaje: Exit Sub
                       
               If Sw Then
                  'Obtiene el Numero de Orden
                  nCorrela = Calcular("select max(tRecibo)  as Codigo from MINGRESO where substring(tRecibo,1,2)= substring(ltrim(str(year(getdate()))),3,2)", Cn)
                  If IsNull(nCorrela) Or Mid(nCorrela, 1, 2) <> Mid(Trim(str(Year(FechaServidor()))), 3, 2) Then
                      txtRecibo.Text = Mid(Trim(str(Year(FechaServidor()))), 3, 2) & "00000001"
                  Else
                      txtRecibo.Text = Mid(Trim(str(Year(FechaServidor()))), 3, 2) & Lib.Correlativo(Mid(nCorrela, 3, 8), 8)
                  End If
                  
                  Sw = False
                    Call DrawBarcode(txtRecibo.Text, Picture1)
                  'Cambiar el SQL
                  Isql = "insert into MINGRESO( " & _
                         "tRecibo, tCaja, tTurno, fFecha, tMoneda, nTipoCambio, nMonto, tTarjeta, tTipoPago, tReferencia, tDescripcion, tAutoriza, lAnticipo, tEstadoDocumento, tUsuario, fRegistro,fDiaContable,tpedido) " & _
                         "values ('" & txtRecibo.Text & "', " & _
                                " '" & sCaja & "', " & _
                                " '" & txtTurno.Text & "', " & _
                                "{fn curdate()}, " & _
                                "'" & sMoneda & "', " & _
                                nTC & ", " & _
                                nMonto & ", " & _
                                "'" & sTarjeta & "', " & _
                                "'" & sTipoPago & "', " & _
                                "'" & txtReferencia.Text & "', " & _
                                " '" & txtDescripcion.Text & "', " & _
                                " '" & txtAutorizacion.Text & "', " & _
                                IIf(cmdOpcion(6).Caption = "Si", 1, 0) & ", " & _
                                "'01', " & _
                                "'" & sUsuario & "'," & _
                                "getdate() ,'" & Format(obtieneDiaContable, "yyyyMMdd") & "', '" & Trim(txtPedido.Text) & "')"

                       Cn.Execute Isql
                       frmReciboIngreso.RsCabecera.Requery
                       frmReciboIngreso.RsCabecera.Find "tRecibo ='" & txtRecibo.Text & "'"
                       MsgBox "Registro Guardado", vbInformation, sMensaje
                       ActivarBotones (True)
                       cmdOpcion(1).Enabled = False
                       cmdTexto.Caption = "Registro " & IIf(frmReciboIngreso.RsCabecera.RecordCount = 0, 0, frmReciboIngreso.RsCabecera.AbsolutePosition) & " de " & frmReciboIngreso.RsCabecera.RecordCount

                       Isql = "select tRecibo, tTurno, fFecha, tMoneda, nMonto, tDescripcion, tAutoriza, tEstadoDocumento, tUsuario, fRegistro, estadodocumento as Estado, Moneda, treferencia, ttipopago,ttarjeta,lanticipo, str(nMonto,10,2) as xMonto, convert(varchar, fFecha, 103) as xFecha from vingreso " & _
                              "where tRecibo = '" & txtRecibo.Text & "'"

                       Set RsIngreso = Lib.OpenRecordset(Isql, Cn)
                         If lCodigoReciboIngreso Then
                                ImprimeReciboIngreso RsIngreso, Picture1
                         Else
                                ImprimeReciboIngreso RsIngreso, Picture1
                         End If
                       Unload Me
               Else
                  'Cambiar el SQL
                  Isql = "update MINGRESO set " & _
                         "tMoneda ='" & sMoneda & "', " & _
                         "nMonto =" & nMonto & ", " & _
                         "tTarjeta ='" & sTarjeta & "', " & _
                         "tTipoPago ='" & sTipoPago & "', " & _
                         "tReferencia ='" & txtReferencia.Text & "', " & _
                         "tDescripcion ='" & txtDescripcion.Text & "', " & _
                         "lAnticipo =" & IIf(cmdOpcion(6).Caption = "Si", 1, 0) & ", " & _
                         "tAutoriza ='" & txtAutorizacion.Text & "', lReplica=1 " & _
                         " where tRecibo = '" & txtRecibo & "'"

                   Cn.Execute Isql
                   'Cambiar el Nombre del Formulario
                   nPos = frmReciboIngreso.RsCabecera.AbsolutePosition
                   frmReciboIngreso.RsCabecera.Requery
                   frmReciboIngreso.RsCabecera.AbsolutePosition = nPos
                   Asignar
                   MsgBox "Registro Modificado", vbInformation, sMensaje
               End If
               'ImprimeReciboIngreso frmReciboIngreso.RsCabecera
          
          Case Is = 2 'Eliminar
               If frmReciboIngreso.RsCabecera.RecordCount = 0 Then
                  Exit Sub
               End If
               
               If frmReciboIngreso.RsCabecera!tTurno = sTurno Then
                  'Password
                  If Supervisor("07") = False Then
                     MsgBox "Clave no permitida", vbExclamation, sMensaje
                     Exit Sub
                  End If
               Else
                  'Password
                  If Supervisor("06") = False Then
                     MsgBox "Clave no permitida", vbExclamation, sMensaje
                     Exit Sub
                  End If
               End If
               
               'Cambia el MsgBox
               If MsgBox("Seguro de Eliminar el Recibo Nro." & txtRecibo & "?", vbQuestion + vbOKCancel, sMensaje) = vbCancel Then
                  Exit Sub
               End If
               
               'Cambia el Delete
               Cn.Execute "update MINGRESO set tEstadoDocumento = '04' where tRecibo = '" & txtRecibo & "'"
               nPos = frmReciboIngreso.RsCabecera.AbsolutePosition
               frmReciboIngreso.RsCabecera.Requery
               frmReciboIngreso.RsCabecera.AbsolutePosition = nPos
               Asignar
                              
          Case Is = 3 'Salir
               Unload Me
          
          Case Is = 4 'KB Descripcion
               frmKeyBoard.txtResultado.Text = txtDescripcion.Text
               frmKeyBoard.Show vbModal
               txtDescripcion.Text = IIf(wEnter, sDescrip, txtDescripcion.Text)
               
          Case Is = 5 'KB Autorizacion
               frmKeyBoard.txtResultado.Text = txtAutorizacion.Text
               frmKeyBoard.Show vbModal
               txtAutorizacion.Text = IIf(wEnter, sDescrip, txtAutorizacion.Text)
                                   
          Case Is = 6 'Si / No
               cmdOpcion(6).Caption = IIf(cmdOpcion(6).Caption = "Si", "No", "Si")
                                   
   End Select
End Sub

Private Sub cmdReferencia_Click()
   sTipo = "Numero"
   frmNumPad.Show vbModal
   txtReferencia.Text = IIf(wEnter, sDescrip, txtReferencia.Text)
End Sub
Function InsertarPedidoAnticipo() As String
On Error GoTo fin
    Dim CodProd As String
    Dim RsProducto As Recordset
    Dim nPVenta As Double
    Dim nPBase As Double
    Dim nImpuesto1 As Double
    Dim nImpuesto2 As Double
    Dim nImpuesto3 As Double
    Dim nRecargo As Double
    Dim nDescuento As Double
    Dim nOficial As Double
    Dim nCantidad As Double
    Dim Acumulado As Double
    '**********
    Dim nValor As Double
    Dim lImp1 As Boolean
    Dim lImp2 As Boolean
    Dim lImp3 As Boolean
    Dim Oficial As Double

    '**********
    
    CodProd = Calcular("select top 1 isnull(TCodAnticipo,'') as codigo from tparametro", Cn)
    
    If CodProd = "" Then
        MsgBox "No hay producto configurado como anticipo en parametros generales!!!"
        Exit Function
    End If
    '****** RECARGA PRODUCTOS Y PRODUCTO ANTICIPO DE PARAMETRO GENERAL ******
    Set RsProducto = Lib.OpenRecordset("usp_Inforest_ObtieneProductos", Cn)
    RsProducto.Find "Codigo = '" & Trim(CodProd) & "'"
    
    If RsProducto.BOF Or RsProducto.EOF Then
        Exit Function
    End If
    
    
    'Cn.BeginTrans
    '****** GENERA EL PROCESO DE INSERCCION DE PEDIDO EN EL SISTEMA *****
    sPedido = ""
    Set oComando = New clsComando
    If Not oComando.CreateCmdSp("spIns_MPEDIDO", Cn) Then
       Set oComando = Nothing
       GoTo fin
    End If
    oComando.CreateParameter "@tCliente", adVarChar, adParamInput, 7, sCliente
    oComando.CreateParameter "@tTipoPedido", adVarChar, adParamInput, 2, "01"
    oComando.CreateParameter "@lPrioridad", adBoolean, adParamInput, 1, 0
    oComando.CreateParameter "@tTipoAtencion", adVarChar, adParamInput, 2, "01"
    oComando.CreateParameter "@tMesa", adVarChar, adParamInput, 3, ""
    oComando.CreateParameter "@tMozo", adVarChar, adParamInput, 4, ""
    oComando.CreateParameter "@tMotorizado", adVarChar, adParamInput, 4, ""
    oComando.CreateParameter "@tCaja", adVarChar, adParamInput, 3, sCaja
    oComando.CreateParameter "@tSalon", adVarChar, adParamInput, 2, sSalon
    oComando.CreateParameter "@tTurno", adVarChar, adParamInput, 10, sTurno
    oComando.CreateParameter "@tObservacion", adVarChar, adParamInput, 250, txtDescripcion.Text
    oComando.CreateParameter "@nTiempo", adInteger, adParamInput, 10, 0
    oComando.CreateParameter "@tUsuario", adVarChar, adParamInput, 15, Right(sUsuario, 15)
    oComando.CreateParameter "@nAdulto", adInteger, adParamInput, 10, 0
    oComando.CreateParameter "@nNino", adInteger, adParamInput, 10, 0
    oComando.CreateParameter "@nMesa", adInteger, adParamInput, 10, 0
    oComando.CreateParameter "@tPuntoVenta", adVarChar, adParamInput, 2, ""
    oComando.CreateParameter "@tHabitacion", adVarChar, adParamInput, 6, ""
    oComando.CreateParameter "@tReserva", adVarChar, adParamInput, 6, ""
    oComando.CreateParameter "@tPasajero", adVarChar, adParamInput, 50, ""
    oComando.CreateParameter "@tCompania", adVarChar, adParamInput, 5, ""
    oComando.CreateParameter "@tContacto", adVarChar, adParamInput, 4, ""
    oComando.CreateParameter "@nDescuento", adDouble, adParamInput, 10, 0
    oComando.CreateParameter "@tDescuento", adVarChar, adParamInput, 3, ""
    oComando.CreateParameter "@tObservacionDescuento", adVarChar, adParamInput, 250, ""
    oComando.CreateParameter "@tAutorizaDescuento", adVarChar, adParamInput, 15, ""
    oComando.CreateParameter "@nTiempoDelivery", adInteger, adParamInput, 10, 0
    oComando.CreateParameter "@tTienda", adVarChar, adParamInput, 3, ""
    oComando.CreateParameter "@fDiaContable", adDate, adParamInput, 10, obtieneDiaContable
    oComando.CreateParameter "@fProgramacion", adDate, adParamInput, 20, Format(FechaServidor, "dd/MM/yyyy HH:mm")
    oComando.CreateParameter "@tCodigoInvitado", adVarChar, adParamInput, 10, ""
    oComando.CreateParameter "@tCodigopariente", adVarChar, adParamInput, 7, ""
    oComando.CreateParameter "@tEntregarA", adVarChar, adParamInput, 20, ""
    oComando.CreateParameter "@nTiempoAntesEnvio", adInteger, adParamInput, 10, 0
    oComando.CreateParameter "@nMontoMaximo", adVarChar, adParamInput, 250, 0
    oComando.CreateParameter "@tPedido", adVarChar, adParamOutput, 10, sPedido

   oComando.CreateParameter "@codigoOrigenVentas", adVarChar, adParamInput, 2, ""
                      
                      
    If Not oComando.GetParamOK Then
       Set oComando = Nothing
       GoTo fin
    End If
    If Not oComando.ExecSP Then
       Set oComando = Nothing
       GoTo fin
    Else
       sPedido = oComando.GetParameterValue("@tPedido")
    End If
                                                                             
    If sPedido <> "" Then
        nRecargo = 0
        nDescuento = 0
        nValor = 0
        nValor = nValor + IIf(RsProducto!lImpuesto1, nPorcentaje1, 0)
        nValor = nValor + IIf(RsProducto!lImpuesto2, nPorcentaje2, 0)
        nValor = nValor + IIf(RsProducto!lImpuesto3, nPorcentaje3, 0)
        
        lImp1 = RsProducto!lImpuesto1
        lImp2 = RsProducto!lImpuesto2
        lImp3 = RsProducto!lImpuesto3
        
       nPVenta = sMontoTotalAnticipo
       nOficial = sMontoTotalAnticipo
       xCantidad = 1
       
       Oficial = IIf(IsNull(RsProducto!nprecioVenta), 0, RsProducto!nprecioVenta)
       nOficial = IIf(RsProducto!tMoneda = "02", nOficial * nTC, nOficial)
    
       Select Case pais 'ok
           Case "001" 'Bolivia
                   nValor = (nValor / 100)
                   nImpuesto1 = IIf(lImp1, nPVenta * nPorcentaje1 / 100, 0)
                   nImpuesto2 = IIf(lImp2, nPVenta * nPorcentaje2 / 100, 0)
                   nImpuesto3 = IIf(lImp3, nPVenta * nPorcentaje3 / 100, 0)
                   nPBase = nPVenta - nImpuesto1 - nImpuesto2 - nImpuesto3
           
           Case Else 'Peru, Ecuador
                   nValor = 1 + (nValor / 100)
                   nImpuesto1 = IIf(lImp1, nPVenta / nValor * nPorcentaje1 / 100, 0)
                   nImpuesto2 = IIf(lImp2, nPVenta / nValor * nPorcentaje2 / 100, 0)
                   nImpuesto3 = IIf(lImp3, nPVenta / nValor * nPorcentaje3 / 100, 0)
                   nPBase = nPVenta - nImpuesto1 - nImpuesto2 - nImpuesto3
       End Select
    
        Dim nInsumo As Double
        Dim nGasto As Double
        Dim nMObra As Double
    
'    If sTipoPedido = "01" Then
       nInsumo = IIf(IsNull(RsProducto!nInsumo), 0, RsProducto!nInsumo)
       nGasto = IIf(IsNull(RsProducto!nGasto), 0, RsProducto!nGasto)
       nMObra = IIf(IsNull(RsProducto!nManoObra), 0, RsProducto!nManoObra)
'    ElseIf sTipoPedido = "02" Then
'       nInsumo = IIf(IsNull(RsProducto!nInsumo2), 0, RsProducto!nInsumo2)
'       nGasto = IIf(IsNull(RsProducto!nGasto2), 0, RsProducto!nGasto2)
'       nMObra = IIf(IsNull(RsProducto!nManoObra2), 0, RsProducto!nManoObra2)
'    ElseIf sTipoPedido = "03" Then
'       nInsumo = IIf(IsNull(RsProducto!nInsumo3), 0, RsProducto!nInsumo3)
'       nGasto = IIf(IsNull(RsProducto!nGasto3), 0, RsProducto!nGasto3)
'       nMObra = IIf(IsNull(RsProducto!nManoObra3), 0, RsProducto!nManoObra3)
'    ElseIf sTipoPedido = "04" Then
'       nInsumo = IIf(IsNull(RsProducto!nInsumo4), 0, RsProducto!nInsumo4)
'       nGasto = IIf(IsNull(RsProducto!nGasto4), 0, RsProducto!nGasto4)
'       nMObra = IIf(IsNull(RsProducto!nManoObra4), 0, RsProducto!nManoObra4)
'    Else
'       nInsumo = IIf(IsNull(RsProducto!nInsumo5), 0, RsProducto!nInsumo5)
'       nGasto = IIf(IsNull(RsProducto!nGasto5), 0, RsProducto!nGasto5)
'       nMObra = IIf(IsNull(RsProducto!nManoObra5), 0, RsProducto!nManoObra5)
'    End If
'
        Isql = "insert into DPEDIDO " & _
           "(tCodigoPedido, tTipoPedido, tItem, tCodigoProducto, tCodigoGrupo, tCodigoSubGrupo, " & _
           "nPrecioNeto, nRecargo, nDescuento, nPrecioOficial, nPrecioImpuesto1, nPrecioImpuesto2, nPrecioImpuesto3, nPrecioVenta, " & _
           "nCantidad, nVenta, nImpuesto1, nImpuesto2, nImpuesto3, " & _
           "lImprime, tArea, lImprimeArea, lCombinacion, nCombinacion, tEstadoItem, tComanda, fRegistro, tMozoD, tUsuarioD, nInsumo, nGasto, nManoObra, nOrden, tUnidadNegocio, tOferta, tsubalmacen,fdiacontable,tcajad ) " & _
           "Values(   '" & sPedido & "', " _
                   & "'01', '001', '" & CodProd & "', " _
                   & "'" & IIf(IsNull(RsProducto!tgrupo), "", RsProducto!tgrupo) & "', " _
                   & "'" & IIf(IsNull(RsProducto!tSubGrupo), "", RsProducto!tSubGrupo) & "', " _
                   & nPBase & ", " & nRecargo & ", " & nDescuento & ", " & nOficial & ", " _
                   & nImpuesto1 & ", " & nImpuesto2 & ", " & nImpuesto3 & ", " _
                   & nPVenta & ", " & xCantidad & ", " & nPVenta * xCantidad & ", " _
                   & nImpuesto1 * xCantidad & ", " & nImpuesto2 * xCantidad & ", " & nImpuesto3 * xCantidad & ", " _
                   & "0, '" & RsProducto!tArea & "', " _
                   & IIf(RsProducto!lImprimeArea, -1, 0) & "," _
                   & IIf(RsProducto!lCombinacion, -1, 0) & "," _
                   & IIf(IsNull(RsProducto!nCombinacion), 0, RsProducto!nCombinacion) & "," _
                   & "'N', '" & sComanda & "', getDate(), '0000', '" & sUsuario & "', " _
                   & nInsumo & ", " & nGasto & ", " & nMObra & ", " _
                   & "1, '" & sUnidadNegocio & "','','','" & Format(obtieneDiaContable, "yyyyMMdd") & "', '" & sCaja & "')"
        Cn.Execute Isql
    Else

    End If
    
'Cn.CommitTrans
InsertarPedidoAnticipo = sPedido
Exit Function

fin:
'Cn.RollbackTrans
InsertarPedidoAnticipo = ""
MsgBox "ERROR: " & error
End Function
Private Sub cmdTipoDocumento_Click(Index As Integer)
On Error GoTo fin
    Dim xUltimoCorrelativo As String
    nTotalPR = txtMonto.Text
    Dim imgcodigobarra As Image
    lDocumEquivaPrecuenta = False
    
    Dim RscadenaCodigoHash As Recordset
    
    Dim fso1 As Object
    
    Dim sSerie As String
    Dim sCorrela As String
    Dim sPrefijo As String
    Dim RsSuma As Recordset
    Dim sTipoDocumento As String
    Dim sImp As String
    Dim wConsumo As Boolean
    Dim i As Integer
    Dim j As Integer
    Dim X As Integer
    Dim nRespuesta As Integer
    Dim nTotal As Double
    Dim sClienteFactura As String
    
    'sUsuarioAutoriza = sUsuario
    
    'Ecuador
    Dim sAutorizacion As String
    Dim fInicio As Date
    Dim fCaducidad As Date
   
    wConsumo = False

    sVar1 = sTipoPedido
    'txtBarra.SetFocus

    '--------------------------------------------------------------------------------
    Dim nMontoPedidoFacturarMInimo As Long
    Dim nMontoPedidoFacturar As Long
    Dim TimpresionDolaresDelivery As Boolean
    Dim sTipoDocum As String
    Dim lValidaEmail As Boolean
    Dim sEmail As String
    Dim lImprimeAlternativa  As Boolean
    Dim RsImpresion As Recordset
    Dim rstFuente As Recordset
    Dim numeroSerieImpresora As String
    Dim codigoImpresora As String
    Dim cadenaCodigoHash As String
    Dim sDetalleConsumo As String
    lImprimeAlternativa = False
    sDetalleConsumo = ""
        
    
    
    
    
    sMontoTotalAnticipo = CDbl(txtMonto.Text)
    ' verifica
    If sMontoTotalAnticipo = 0 Then
        MsgBox "Error: No se puede Facturar un Anticipo con Monto '0', Favor de Verificar!!!", vbInformation, sMensaje
        Exit Sub
    End If
    sMontoTotalAnticipo = IIf(sMoneda = "02", sMontoTotalAnticipo * nTC, sMontoTotalAnticipo)
    'Chequea Consistencia
    RsTipoDocumento.Requery
    RsTipoDocumento.MoveFirst
    RsTipoDocumento.Find ("Descripcion='" & cmdTipoDocumento(Index).Caption & "'")
    If RsTipoDocumento.EOF Then
       MsgBox "Error: Configure los Documentos", vbCritical, sMensaje
       Exit Sub
    Else
       xlTipoDocumento = Calcular("Select lValidaRuc As Codigo From TTIPODOCUMENTO Where tCodigoTipoDocumento = '" & RsTipoDocumento!TTipoEmision & "'", Cn)
    End If
    
    Dim DsOLD As String
    DsOLD = Calcular("select top 1 ttipodocumento as codigo from mdocumento where treserva='" & frmReservaDetalle.txtCodigo.Text & "'", Cn)
    
    If (RsTipoDocumento!TTipoEmision <> DsOLD And DsOLD <> "" And DsOLD <> "0") Then
        MsgBox "El documento seleccionado no Coincide con los Documentos generados a la Reserva, favor de Verificar!!!!", vbInformation, sMensaje
        Exit Sub
    End If
    
 'validacionMontoMinimo
    nMontoPedidoFacturarMInimo = 0
    nMontoPedidoFacturarMInimo = Calcular("select isnull(nMontoMinimo,0) as codigo   from vTipoDocumento where Codigo='" & RsTipoDocumento!TTipoEmision & "'", Cn)
    If nMontoPedidoFacturarMInimo > 0 Then
        If nMontoPedidoFacturarMInimo >= CDbl(txtMonto.Text) Then
            MsgBox "El Monto a Facturar no llega al Minimo Permitido al Tipo de Documento"
            Exit Sub
        End If
    End If

    'validacionMontoMaximo
    nMontoPedidoFacturar = 0
    nMontoPedidoFacturar = Calcular("select isnull(nMontoMaximo,0) as codigo   from vTipoDocumento where Codigo='" & RsTipoDocumento!TTipoEmision & "'", Cn)
    If nMontoPedidoFacturar > 0 Then
        If nMontoPedidoFacturar <= CDbl(txtMonto.Text) Then
            MsgBox "El Monto a Facturar supera al M�ximo Permitido al Tipo de Documento"
            Exit Sub
        End If
    End If


    TimpresionDolaresDelivery = False
    '-------- impresion en dolares si esta activo el check en el cliente delivery.
    If Calcular("select isnull(lEmisionMonedaExtranjera,0) as codigo from tdelivery where tcodigodelivery='" & sCliente & "'", Cn) Then
        If MsgBox("�Desea Imprimir en " & sMonedaE & "?", vbQuestion + vbYesNo + vbDefaultButton2) = vbYes Then
            TimpresionDolaresDelivery = True
        Else
            TimpresionDolaresDelivery = False
        End If
    Else
        TimpresionDolaresDelivery = False
    End If
    
    'TIPO EMISION
    If Not wConsumo And RsTipoDocumento!tFormulario <> "01" Then
    '------------- no es permitido emitir en otro formato
    Else
       '---------------------------------------------------------------------------
       '  ************************* EMISION POR TICKETERA *************************
       '---------------------------------------------------------------------------
         
       'Calcula el total de la cabecera
       nTotal = sMontoTotalAnticipo 'Calcular("select sum(nVenta) as Codigo from DPEDIDO where (isnull(tFacturado,'0') = '0' or len(ltrim(tFacturado)) = 0) and tCodigoPedido ='" & sPedido & "'", Cn)
                                   
       If RsTipoDocumento!Cliente And (RsTipoDocumento!Monto <= nTotal Or RsTipoDocumento!Monto = 0) Then
                'Consistencia Factura
                frmBusquedaRapida.nPredeterm = 1
                Isql = "exec usp_Inforest_ObtieneClientesFactura '" & sCliente & "','" & RsTipoDocumento!TTipoEmision & "'"
            
                If Calcular("select COUNT(*) as codigo from TDELIVERYCLIENTE Where tCodigoDelivery='" & sCliente & "'", Cn) = 0 Then
                   sTemp = ""
                   frmBusquedaRapida.nPredeterm = 1
                End If
                
                Select Case pais 'ok
                    Case "001" 'Bolivia
                    
                        Call ConfGrilla(3, frmBusquedaRapida.grdGrilla, "C�digo", 2, "Codigo", 1200, 2, 0, "", _
                                                                         "Ident", 2, "tIdentidad", 1800, 2, 0, "", _
                                                                         "Cliente", 2, "Descripcion", 5200, 0, 0, "")
                    Case Else 'Peru, Ecuador
                        If lClub Then
                            Call ConfGrilla(4, frmBusquedaRapida.grdGrilla, "C�digo", 2, "Codigo", 1000, 2, 0, "", _
                                                                          "Ident", 2, "tIdentidad", 1600, 2, 0, "", _
                                                                          "Cliente", 2, "Descripcion", 4500, 0, 0, "", _
                                                                          "Enlace", 2, "tEnlace", 1100, 0, 0, "")
                        Else
                            Call ConfGrilla(3, frmBusquedaRapida.grdGrilla, "C�digo", 2, "Codigo", 1100, 2, 0, "", _
                                                                          "Ident", 2, "tIdentidad", 1600, 2, 0, "", _
                                                                          "Cliente", 2, "Descripcion", 5500, 0, 0, "")
                        End If
                End Select
                
                '------VALIDA CORREO----------
                sTipoDocum = RsTipoDocumento!TTipoEmision
                
                frmBusquedaRapida.cmdOpcion(1).Enabled = True
                frmBusquedaRapida.cmdOpcion(2).Enabled = True
                frmBusquedaRapida.cmdOpcion(3).Enabled = True
                frmBusquedaRapida.Show vbModal
                
                If wEnter = True And sCodigo <> "" Then
                      sTemp = sCliente
                      sClienteFactura = sCodigo
            
                      If pais = "000" Then
                      
                          'IMPRIME DNI
                          Dim RsTc1 As ADODB.Recordset
                          Set RsTc1 = New ADODB.Recordset
                          Set RsTc1 = Lib.OpenRecordset("usp_Inforest_ValidaClienteSel '" & sTipoDocum & "','" & sClienteFactura & "'", Cn)
                          If Not (RsTc1.EOF Or RsTc1.BOF) Then
                              RsTc1.MoveFirst
                              If RsTc1.Fields(0) <> "ok" Then
                                  MsgBox "Error: El tipo de Identidad del Cliente no Corresponde al Tipo de Documento", vbCritical, sMensaje
                                  Exit Sub
                              End If
                          End If
                      
                      End If
                  
                      lValidaEmail = Calcular("Select lValidaEmail As codigo From vTipoDocumento where Codigo='" & sTipoDocum & "'", Cn)
                      
                      If lValidaEmail = True Then
                         sEmail = Calcular("Select ISNULL(tcorreo,'') As codigo From vCLIENTE where Codigo ='" & sCodigo & "' ", Cn)
                      
                         If sEmail = "" Then
                            MsgBox "El cliente no tiene Email registrado", vbCritical, sMensaje
                            Exit Sub
                         End If
                      End If
                        
                    If Calcular("Select lValidaUbigeo As codigo From vTipoDocumento where Codigo='" & sTipoDocum & "'", Cn) = True Then
                        Dim TempUbigeo As String
                        Dim TempUrbaniza As String
                        TempUbigeo = Calcular("Select ISNULL(CodigoUbigeo,'') As codigo From vCLIENTE where Codigo ='" & sCodigo & "' ", Cn)
                        TempUrbaniza = Calcular("Select ISNULL(Urbanizacion,'') As codigo From vCLIENTE where Codigo ='" & sCodigo & "' ", Cn)
                        If Trim(TempUbigeo) = "" Or Trim(TempUrbaniza) = "" Then
                            MsgBox "El cliente no tiene Ubigeo � Urbanizacion registrado, Favor de verificar!!!", vbCritical, sMensaje
                            Exit Sub
                        End If
                    End If

                      Cn.Execute "usp_Inforest_ValidaDeliveryCliente '" & sCliente & "','" & sClienteFactura & "'"
                Else
                   Exit Sub
                End If

        Else
           sClienteFactura = ""
        End If
        
        'Cn.BeginTrans
        '****** INSERCCION DE PEDIDO DE ANTICIPO SEGUN CONFIGURACION ****************
        If InsertarPedidoAnticipo() = "" Then
            Exit Sub
        End If
        '***************************************************************************
        
        'Genera y Actualiza los Numero de Documento
        sSerie = RsTipoDocumento!tSerie
        sCorrela = Lib.Correlativo(RsTipoDocumento!tUltimoNumero, 9)
        sPrefijo = RsTipoDocumento!prefijo
        sTipoDocumento = RsTipoDocumento!TTipoEmision
        sImp = RsTipoDocumento!timpresora
        sDocumento = sPrefijo & sSerie & sCorrela
        sResumen = RsTipoDocumento!lResumen
              
        Set RsSuma = Lib.OpenRecordset("select sum(nVenta) as nVenta from DPEDIDO where (isnull(tFacturado,'0') = '0' or len(ltrim(tFacturado)) = 0 OR ISNULL(tFacturado,'')='') and tCodigoPedido ='" & sPedido & "' group by tCodigoPedido ", Cn)
        nCargo = Round(RsSuma!nVenta, 2)

        Select Case pais
            Case "001" 'Bolivia
'                    tAutorizacion = obtieneAutorizacionDosificacion(sCaja, "1")
'                    tDosificacion = obtieneAutorizacionDosificacion(sCaja, "2")
'                    If tAutorizacion <> "" And tDosificacion <> "" Then
'                        tcodigoControl = devuelveCodigoControl(sCaja, sCorrela, tAutorizacion, tDosificacion, sClienteFactura, nTotal)
'                        If tcodigoControl = "" Then: MsgBox "Error al generar C�digo de Control", vbCritical, sMensaje: Exit Sub
'                    Else
'                           MsgBox "Error al obtener N�mero de Autorizaci�n o Dosificaci�n. Verifique.", vbCritical, sMensaje
'                           Exit Sub
'                    End If
                    
            Case "002" 'Ecuador
                'tAutorizacion = RsTipoDocumento!tNumeroAutorizacion
            Case Else 'Peru, Ecuador
                'tcodigoControl = ""
                'tDosificacion = ""
        End Select
        
        ''ELDCQ (INSERTA DOCUMENTO)
        If Not INSERTADOCUMENTO(sPedido, sDocumento, sTipoDocumento, IIf(sClienteFactura = "", "", sClienteFactura), "01", sCaja, sTurno, sSalon, sUsuario, sUsuario, Format(obtieneDiaContable, "yyyy/MM/dd"), "", "", IIf(TimpresionDolaresDelivery, 1, 0), "", "", "", Format(IIf(IsNull(RsTipoDocumento!fInicio), "01/01/1900", RsTipoDocumento!fInicio), "yyyy/mm/dd"), Format(IIf(IsNull(RsTipoDocumento!fCaducidad), "01/01/1900", RsTipoDocumento!fCaducidad), "yyyy/mm/dd"), tContribuyenteEspecial, 1) Then
            GoTo fin
        End If

        Cn.Execute "Update TTIPODOCUMENTOIMPRESORA Set tUltimoNumero = '" & sCorrela & "' where tTipoEmision ='" & sTipoDocumento & "' and tCaja ='" & sCaja & "'"
        wEnter = True
        
        Cn.Execute "Update MDOCUMENTO Set TRESERVA='" & frmReservaDetalle.txtCodigo.Text & "' where TDOCUMENTO='" & sDocumento & "'"
        xTipo = ""

        Isql = "insert into DPAGODOCUMENTO " & _
               "( tDocumento, tCorrelativo, tTurno, tTipoPago, tMoneda, tReferencia, nTipoCambio," & _
               " nMonto, npropina, tTarjeta, tNumero, tFechaVencimiento, fRegistro, tUsuario,fDiaContable,tNumeroATarjeta, ndolar ) " & _
               "Values(    '" & sDocumento & "',1 ,'" & sTurno & "', '" & sTipoPago & "', '" & sMoneda & "'," _
                        & "'', " _
                        & nTC & ", " & CDbl(txtMonto.Text) & ", 0, " _
                        & "'" & sTarjeta & "','" & Me.txtReferencia.Text & "','', getdate()," _
                        & "'" & sUsuario & "','" & Format(obtieneDiaContable, "yyyyMMdd") & "', ''," & IIf(sMoneda = "02", CDbl(txtMonto.Text), 0) & ")"
    
        Cn.Execute Isql
        Cn.Execute "Update MPEDIDO set fLlegada=getdate(),fEntrega=getdate(), tEstadoPedido = '02', lReplica = 1 where tCodigoPedido ='" & sPedido & "'"
        Cn.Execute "Update DPEDIDO set tFacturado = 'P', tDocumento = '" & sDocumento & "',tCortesia='' where tCodigoPedido ='" & sPedido & "' and (isnull(tFacturado,'0')='0' or len(ltrim(tFacturado))=0)"
        
                       
        Isql = "Update MDOCUMENTO set " & _
                  "tEstadoDocumento = '02', " & _
                  "nAbono = " & sMontoTotalAnticipo & ", " & _
                  "fPago = getdate()," & _
                  "nVuelto = 0 " & _
                  ",lreplica=1, TCORTESIA=''  where tDocumento ='" & sDocumento & "'"
        Cn.Execute Isql
        
        'PARA NO FISCALES
        Cn.Execute "UPDATE DPEDIDO SET lregistroventa=(select case when registroventa=0 then 0 else 1 end from vtipodocumento where codigo='" & sTipoDocumento & "') where tCodigoPedido ='" & sPedido & "' and  tDocumento ='" & sDocumento & "'"
        
        '-----------------------------------------------------------------------------------------------------
        If pais = "000" And lFEpape And IIf(RsTipoDocumento!lFacturacionElectronica = True, 1, 0) Then
             If Not FacturarTCPIP(2, sDocumento, 0) Then
                 'Cn.Execute "Update TMESA set tEstadoMesa = '02' where tCodigoMesa ='" & sMesa & "'"
                 Cn.Execute "delete MPEDIDO  where tCodigoPedido ='" & sPedido & "'"
                 Cn.Execute "delete DPEDIDO  where tCodigoPedido ='" & sPedido & "'"
                 Cn.Execute "Delete MDOCUMENTO Where tDocumento= '" & sDocumento & "'"
                 Cn.Execute "Delete DDOCUMENTO Where tDocumento= '" & sDocumento & "'"
                 
                 Cn.Execute "Delete DPAGODOCUMENTO Where tDocumento= '" & sDocumento & "'"
                 xUltimoCorrelativo = Calcular("select MAX(tDocumento) as codigo from MDOCUMENTO where tcaja='" & sCaja & "' and tTipoDocumento='" & sTipoDocumento & "'", Cn)
                 xUltimoCorrelativo = Right(xUltimoCorrelativo, 9)
                 Cn.Execute "Update TTIPODOCUMENTOIMPRESORA Set tUltimoNumero = '" & xUltimoCorrelativo & "' where tTipoEmision ='" & sTipoDocumento & "' and tCaja ='" & sCaja & "'"

                GoTo fin
             End If
        End If
        '-------------------------------------------------------------------------------------------------------
        
        '============= nov 2010
        Dim imprimeDolaDocumentos As String
        imprimeDolaDocumentos = Calcular("select isnull(lequivadolares,0) as codigo from vtipodocumentoimpresora where tcaja='" & sCaja & "' and ttipoemision='" & sTipoDocumento & "' ", Cn)
        If imprimeDolaDocumentos = "Verdadero" Then
            lDocumEquivaPrecuenta = True
        Else
            lDocumEquivaPrecuenta = False
        End If
        
            If wConsumo = False And lDescripcionAlternativa = True Then
             If validaImpresionAlternativa(sDocumento) = False Then
                    If MsgBox("Desea imprimir descripci�n Alternativa? ", vbQuestion + vbYesNo + vbDefaultButton2, sMensaje) = vbYes Then
                          lImprimeAlternativa = True
                    End If
              End If
            End If
        '-------------------------------  SE INTRDUJO LAS CONSULTAS A  UN STORE PROCEDURE -------------------------
        ' ELDCQ 15/11/2017
        'Imprime Documentos
         If lImprimeAlternativa = False Then
            If lDocumentoAgrupado Then
                Isql = "EXEC usp_Inforest_Impresion '" & sDocumento & "',1"
                'FACTURACION_E_PERU
                IsqlFact = "EXEC usp_Inforest_Impresion '" & sDocumento & "',2"
            Else
                Isql = "EXEC usp_Inforest_Impresion '" & sDocumento & "',3"
                'FACTURACION_E_PERU
                IsqlFact = "EXEC usp_Inforest_Impresion '" & sDocumento & "',4"
            End If
        Else
            If lDocumentoAgrupado Then
                Isql = "EXEC usp_Inforest_Impresion '" & sDocumento & "',5"
                'FACTURACION_E_PERU
                IsqlFact = "EXEC usp_Inforest_Impresion '" & sDocumento & "',6"
            Else
                Isql = "EXEC usp_Inforest_Impresion '" & sDocumento & "',7"
               'FACTURACION_E_PERU
                IsqlFact = "EXEC usp_Inforest_Impresion '" & sDocumento & "',8"
            End If
        End If
        
        '---------------------------------------
        Set RsImpresion = Lib.OpenRecordset(Isql, Cn)
        Set rstFuente = New ADODB.Recordset
        imageCab.Picture = Nothing
        'imagepIE.Picture = Nothing
'        Set rstFuente = Lib.OpenRecordset("select iImagenCabDoc AS foto, iImagenPieDoc as fotoPie  from tcaja where tcaja='" & sCaja & "'", Cn)
'        imageCab.DataField = "foto"
'        Set imageCab.DataSource = rstFuente
'        imagepIE.DataField = "fotoPie"
'        Set imagepIE.DataSource = rstFuente

       If RsImpresion.RecordCount = 0 Then
          LimpiaRs
          MsgBox "No existen Datos a Imprimir", vbExclamation, sMensaje
       Else
        'SUNAT
        numeroSerieImpresora = obtieneNumeroSerieImpresora(sCaja, sImp)
        codigoImpresora = sImp
        'SUNAT
        Cn.Execute " update mdocumento set timpresora='" & codigoImpresora & "', tSerieImpresora='" & numeroSerieImpresora & "' where tdocumento ='" & sDocumento & "' "

               'FACTURACION_E_PERU
               If pais = "000" Then
                   If lFacturacionE Then
                   
                       If lFEOfisis Then 'OFISIS
                    
                       ElseIf lFESpring Then
                                                               
                       ElseIf lFECarbajal Then
                           
                       ElseIf lFEpape Then
                            If IIf(RsTipoDocumento!lFacturacionElectronica = True, 1, 0) Then
                                If tCodigoFE = "000" Then
                                     If lQRFE Then
                                         Set imageHash.Picture = LoadPicture(CrearImagenQR(PapeTermico))
                                     Else
                                         If lImpresionCodigoBarras Then
                                             'Set imageHash.Picture = LoadPicture(ImagenQR(sDocumento))
                                         Else
                                             cadenaCodigoHash = PapeMatricial
                                         End If
                                     End If
                                 End If
                            End If
                       ElseIf lFEBiz Then
                           If IIf(RsTipoDocumento!lFacturacionElectronica = True, 1, 0) Then
                                If Not INSERTA_FE_INFOREST(sDocumento, 1, DateTime.Date) Then '----CABECERA
                                    Cn.Execute "Update TMESA set tEstadoMesa = '02' where tCodigoMesa ='" & sMesa & "'"
                                    Cn.Execute "Delete MDOCUMENTO Where tDocumento= '" & sDocumento & "'"
                                    Cn.Execute "Delete DDOCUMENTO Where tDocumento= '" & sDocumento & "'"
                                    Cn.Execute "Delete DPAGODOCUMENTO Where tDocumento= '" & sDocumento & "'"
                                    xUltimoCorrelativo = Calcular("select MAX(tDocumento) as codigo from MDOCUMENTO where tcaja='" & sCaja & "' and tTipoDocumento='" & sTipoDocumento & "'", Cn)
                                    xUltimoCorrelativo = Right(xUltimoCorrelativo, 9)
                                    Cn.Execute "Update TTIPODOCUMENTOIMPRESORA Set tUltimoNumero = '" & xUltimoCorrelativo & "' where tTipoEmision ='" & sTipoDocumento & "' and tCaja ='" & sCaja & "'"
                                    Cn.Execute "Update DPEDIDO set tFacturado = NULL, tDocumento = NULL where tCodigoPedido ='" & sPedido & "' "
                                    Cn.Execute "Update MPEDIDO set tEstadoPedido = '01', lReplica = 1 where tCodigoPedido = '" & sPedido & "'"
                                    Cn.Execute "UPDATE DPEDIDO SET lregistroventa = NULL where tCodigoPedido ='" & sPedido & "' and  tDocumento ='" & sDocumento & "'"
                                    'MsgBox "Se Produjo un Problema al enviar el Documento" & vbNewLine & "Favor de Refrescar los valores e intentarlo nuevamente!!!", vbInformation, sMensaje
                                    GoTo fin
                                 End If
                                 Sleep 1000
                                 'VALIDAR RESPUESTA DE CODIGO HASH Y CODIGO DE BARRA
                                 If tCodigoFE = "000" Then
                                     If lQRFE Then
                                         Set imageHash.Picture = LoadPicture(QRHASH_FE_INFOREST(3, sDocumento, 0))
                                     Else
                                         If lImpresionCodigoBarras Then
                                             Set imageHash.Picture = LoadPicture(QRHASH_FE_INFOREST(1, sDocumento, 0))
                                         Else
                                             cadenaCodigoHash = QRHASH_FE_INFOREST(2, sDocumento, 0)
                                         End If
                                     End If
                                 End If
                            End If
                            
                       Else 'INFOFACT

                       End If
                   End If
            End If
            '---------------------------------------
            'Configura la Impresora
            Imprimir (sImp)
            Printer.FontName = sFont
            Printer.FontBold = False
            
            If wConsumo Then
               If RsTipoDocumento!tFormulario = "01" Then
               
                  If RsTipoDocumento!Cliente And RsTipoDocumento!Monto = 0 Then
                      'FACTURACION ELECTRONICA
                      If IIf(RsTipoDocumento!lFacturacionElectronica = True, 1, 0) And tCodigoFE <> "999" Then
                           ImprimeFacturaConsumoElectronico RsImpresion, sDetalleConsumo, imageHash, sTipoDocumento, imageCab, imagepIE, cadenaCodigoHash, TimpresionDolaresDelivery
                      Else
                           ImprimeFacturaConsumoT RsImpresion, sDetalleConsumo, sTipoDocumento, imageCab, imagepIE, TimpresionDolaresDelivery
                      End If
                     
                      NFactura = sCorrela
                      'lblFactura.Caption = NFactura
                  Else
                      'FACTURACION ELECTRONICA
                      If IIf(RsTipoDocumento!lFacturacionElectronica = True, 1, 0) And tCodigoFE <> "999" Then
                          ImprimeBoletaConsumoElectronico RsImpresion, sDetalleConsumo, imageHash, sTipoDocumento, imageCab, imagepIE, cadenaCodigoHash, TimpresionDolaresDelivery
                      Else
                          ImprimeBoletaConsumoT RsImpresion, sDetalleConsumo, sTipoDocumento, imageCab, imagepIE, TimpresionDolaresDelivery
                      End If
                  End If
                  
               Else
               
'                  If lFacturacionE And RsTipoDocumento!lFacturacionElectronica And lFEOfisis = False And lFESpring = False And lFEpape = False And lFECarbajal = False Then
'                        If Generar_Imagen(CnFE, "select imagen from IMAGENCODIGOBARRA where nro_efact='" & fDocumento & "'", "imagen", "\fact.bmp") = True Then
'                            ImprimeFormatoAConsumo
'                            Kill App.Path & "\fact.bmp"
'                        Else
'                            ImprimeFormatoAConsumo
'                        End If
'                  ElseIf lFacturacionE And lFECarbajal Then
'                        ImprimeFormatoAConsumo
'
'                        Set fso1 = CreateObject("Scripting.FileSystemObject")
'                        If fso1.FileExists(App.Path & "\fact.bmp") Then
'                            Kill App.Path & "\fact.bmp"
'                        End If
'                    ElseIf lFacturacionE = False Then
'                        If RsTipoDocumento!Cliente And RsTipoDocumento!Monto = 0 Then
'                           ImprimeFacturaConsumoN RsImpresion, sDetalleConsumo, sEmpresa
'                           NFactura = sCorrela
'                           lblFactura.Caption = NFactura
'                        Else
'                           ImprimeBoletaConsumoN RsImpresion, sDetalleConsumo, sEmpresa
'                        End If
'                  End If
                  
               End If
            Else
               If RsTipoDocumento!tFormulario = "01" Then
                  If RsTipoDocumento!Cliente And RsTipoDocumento!Monto = 0 Then
            
                          'FACTURACION ELECTRONICA
                          If IIf(RsTipoDocumento!lFacturacionElectronica = True, 1, 0) And tCodigoFE <> "999" Then
                             ImprimeFacturaElectronica RsImpresion, imageHash, sTipoDocumento, imageCab, imagepIE, cadenaCodigoHash, TimpresionDolaresDelivery
                          Else
                             ImprimeFacturaT RsImpresion, sTipoDocumento, imageCab, imagepIE, TimpresionDolaresDelivery
                          End If
                     
                        NFactura = sCorrela
                        'lblFactura.Caption = NFactura
                          
                  ElseIf RsTipoDocumento!TTipoEmision = "00" Then
                          If MsgBox("Deseas imprimir el Voucher", vbQuestion + vbYesNo, sMensaje) = vbYes Then
                             ImprimeCortesia RsImpresion, "00", imageCab, imagepIE
                          End If
                  Else
            
                          'FACTURACION ELECTRONICA
                          If IIf(RsTipoDocumento!lFacturacionElectronica = True, 1, 0) And tCodigoFE <> "999" Then
                             ImprimeBoletaElectronica RsImpresion, imageHash, sTipoDocumento, imageCab, imagepIE, cadenaCodigoHash, TimpresionDolaresDelivery
                          Else
                             ImprimeBoletaT RsImpresion, sTipoDocumento, imageCab, imagepIE, TimpresionDolaresDelivery
                          End If
                        
                  End If
                  
               Else
                  If RsTipoDocumento!Cliente And RsTipoDocumento!Monto = 0 Then
                     ImprimeFacturaN RsImpresion, sEmpresa, sTipoDocumento
                     NFactura = sCorrela
                     'lblFactura.Caption = NFactura
                  ElseIf RsTipoDocumento!TTipoEmision = "00" Then
                     If MsgBox("Deseas imprimir el Voucher", vbQuestion + vbYesNo, sMensaje) = vbYes Then
                        ImprimeCortesia RsImpresion, "00", imageCab, imagepIE
                     End If
                  Else
                     ImprimeBoletaN RsImpresion, sEmpresa, sTipoDocumento
                  End If
                  
               End If
               
            End If
       End If
       
       LimpiaRs

    End If

    If sCliente <> "" Then
       Cn.Execute "Update TDELIVERY set tCodigoCliente = '" & sClienteFactura & "' where tCodigoDelivery ='" & sTemp & "'"
    End If
    sTemp = ""
    
    '-----------------------
     If pais = "000" And lFEpape And IIf(RsTipoDocumento!lFacturacionElectronica = True, 1, 0) Then
          If Not FacturarTCPIP(3, sDocumento, 0) Then
             MsgBox ("La confirmacion ha fallado favor de contactarse con paperlees"), vbInformation, sMensaje
          End If
     End If
     '------------------------
    
    
    Screen.MousePointer = vbDefault
    
    MsgBox "Anticipo Creado Correctamente!!", vbInformation, sMensaje
    Unload Me
    Exit Sub
fin:

    Screen.MousePointer = vbDefault
    Call Log_Inforest("PUNTO VENTA", "EMISION DE DOCUMENTO PUNTO DE VENTA", sPedido, "", sDocumento, Estado & " " & error, "", "FALLA AL GENERAR DOCUMENTO EN PUNTO DE VENTA ", sUsuario)
    MsgBox "Error: " & error & vbNewLine & "- Proceso: Emision de Documentos Punto Venta", vbCritical, sMensaje
    
End Sub

Private Sub Form_Load()
   Screen.MousePointer = vbHourglass
   Centrar Me
   
   
   If lCodigoReciboIngreso Then
        ''codigo barra
        Picture1.ScaleMode = 3
        Picture1.Height = Picture1.Height * (1.4 * 40 / Picture1.ScaleHeight)
        Picture1.FontSize = 8
   End If
   'Ingrese el SubTitulo
   Me.Caption = " Mantenimiento de Recibos de Ingreso "
   fraDetalle.Caption = Me.Caption
   
   'Tarjetas de Cr�dito
   Isql = "select * from TTARJETACREDITO where nBoton>0 and lActivo = 1 Order by nBoton"
   Set RsTarjeta = Lib.OpenRecordset(Isql, Cn)
    
   Call AsignaTarjeta(8, RsTarjeta, cmdCancelacion())
   
   nMonto = 0
   cmdCancelacion(9).Caption = cmdCancelacion(9).Caption & " " & sMonN

   If sMonE = "" Then
      cmdCancelacion(10).Visible = False
   Else
      cmdCancelacion(10).Caption = cmdCancelacion(10).Caption & " " & sMonE
   End If
        
   If Sw = True Then
      ActivarBotones (False)
      Blanquear Me
      txtTurno.Text = sTurno
      txtFecha.Text = FechaServidor()
      txtMonto.Text = "0.00"
      lblEstado.Caption = ""
      sMoneda = "01"
      sTipoPago = "01"
      sTarjeta = ""
      cmdOpcion(6).Caption = "Si"
      cmdBuscar.Enabled = True
   Else
   cmdBuscar.Enabled = False
      Asignar
      ActivarBotones (True)
   End If
   If TipoOper = "Anticipo" Then
    LlenarAnticipo
   Else
    cmdTexto.Caption = "Registro " & IIf(frmReciboIngreso.RsCabecera.RecordCount = 0, 0, frmReciboIngreso.RsCabecera.AbsolutePosition) & " de " & frmReciboIngreso.RsCabecera.RecordCount
    Screen.MousePointer = vbDefault
   End If
End Sub
Private Sub LlenarAnticipo()
    Me.fraDetalle.Caption = ""
    Me.Caption = "Anticipo"
    Me.cmdOpcion(6).Visible = False
    Me.Label(6).Visible = False
    Me.Label(7).Visible = False
    Me.Label(0).Visible = False
    Me.Label(8).Visible = False
    Me.cmdBuscar.Visible = False
    Me.txtPedido.Visible = False
    Me.txtRecibo.Visible = False
    Me.txtTurno.Visible = False
    
    Set RsTipoDocumento = Lib.OpenRecordset("select top 2  * from vTipoDocumentoImpresora where tCaja ='" & sCaja & "' and Transporte=0 And lNotaCredito = 0 And lActivo = 1 and tTipoEmision<>'00' order by tTipoEmision", Cn)
    AsignaComando 2, RsTipoDocumento, cmdTipoDocumento()
    
    cmdTipoDocumento(1).Visible = True
    cmdTipoDocumento(2).Visible = True
    frmDocumento.Visible = True
    cmdOpcion(0).Visible = False
    cmdOpcion(1).Visible = False
    cmdOpcion(2).Visible = False
End Sub
Private Sub Form_Unload(Cancel As Integer)
    'Cambia el Nombre del Formulario
    Set frmReciboIngresoDetalle = Nothing
End Sub

Sub ActivarBotones(ByVal Activa As Boolean)
    cmdNavegar(0).Enabled = Activa
    cmdNavegar(1).Enabled = Activa
    cmdNavegar(2).Enabled = Activa
    cmdNavegar(3).Enabled = Activa
    cmdNavegar(4).Enabled = Activa
    cmdNavegar(5).Enabled = Activa
End Sub

Public Sub AsignaTarjeta(nBoton As Integer, RsAsigna As Recordset, cmdBoton As Object)
   Screen.MousePointer = vbHourglass
   Dim i As Integer
   With RsAsigna
        If .RecordCount > 0 Then
           For i = 1 To nBoton
               .MoveFirst
               .Find "nboton = " & Trim(str(i))
               If Not .EOF Then
                  cmdBoton(i).Visible = True
                  cmdBoton(i).Caption = !tResumido
                Else
                    cmdBoton(i).Visible = False
                End If
           Next i
        Else
           For i = 1 To nBoton
               cmdBoton(i).Visible = False
           Next i
        End If
  End With
  Screen.MousePointer = vbDefault
End Sub

Public Sub Bloquear(Activo As Boolean)
   cmdOpcion(1).Enabled = Not Activo
   cmdOpcion(2).Enabled = Not Activo
   cmdOpcion(4).Enabled = Not Activo
   cmdOpcion(5).Enabled = Not Activo
   For i = 1 To 11
       cmdCancelacion(i).Enabled = Not Activo
   Next i
   cmdReferencia.Enabled = Not Activo
End Sub
'diaContable
Public Function obtieneDiaContable() As Date
   Dim oComando As New clsComando
   Dim DiaContable As Date
   Dim rst1 As New ADODB.Recordset
   Set oComando = New clsComando
                  If Not oComando.CreateCmdSp("usp_GenObtieneDiaContable", Cn) Then
                     Set oComando = Nothing
                     Exit Function
                  End If
                  
                  oComando.CreateParameter "@lDiaContable", adBoolean, adParamInput, 1, lDiaContable
                  oComando.CreateParameter "@sHoraCierre", adVarChar, adParamInput, 5, tHoraCierreDiaContable
                  oComando.CreateParameter "@tUsuario", adVarChar, adParamInput, 15, sUsuario
                 oComando.CreateParameter "@fDiaContable", adDBDate, adParamOutput, 10, DiaContable
                If Not oComando.GetParamOK Then
                   Set oComando = Nothing
                   Exit Function
                End If
                    Set rst1 = oComando.GetSP()
                obtieneDiaContable = oComando.GetParameterValue("@fDiaContable")
.End Function

'diaContable


