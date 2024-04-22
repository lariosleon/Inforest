VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "mscomctl.ocx"
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "mswinsck.ocx"
Begin VB.MDIForm mdiPuntoVenta 
   BackColor       =   &H8000000C&
   Caption         =   "Punto de Venta"
   ClientHeight    =   9960
   ClientLeft      =   -2700
   ClientTop       =   -1290
   ClientWidth     =   13560
   Icon            =   "mdiPuntoVenta.frx":0000
   LinkTopic       =   "MDIForm1"
   Begin VB.PictureBox picStretch 
      Align           =   1  'Align Top
      BorderStyle     =   0  'None
      Height          =   8955
      Left            =   0
      Picture         =   "mdiPuntoVenta.frx":57E2
      ScaleHeight     =   8955
      ScaleWidth      =   13560
      TabIndex        =   21
      Top             =   705
      Visible         =   0   'False
      Width           =   13560
   End
   Begin VB.Timer Timer 
      Left            =   135
      Top             =   810
   End
   Begin MSWinsockLib.Winsock Winsock1 
      Left            =   630
      Top             =   810
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin VB.PictureBox xPicture 
      Align           =   1  'Align Top
      Height          =   705
      Left            =   0
      ScaleHeight     =   645
      ScaleWidth      =   13500
      TabIndex        =   1
      Top             =   0
      Width           =   13560
      Begin VB.CommandButton cmdConsultaSaldo 
         Caption         =   "Consultar Saldos"
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   6.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   645
         Left            =   10290
         Picture         =   "mdiPuntoVenta.frx":12D9C
         Style           =   1  'Graphical
         TabIndex        =   20
         Top             =   675
         Visible         =   0   'False
         Width           =   1245
      End
      Begin VB.CommandButton cmdOpcionMensaje 
         Caption         =   "Mensajes"
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   6.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   645
         Left            =   9000
         Picture         =   "mdiPuntoVenta.frx":13126
         Style           =   1  'Graphical
         TabIndex        =   19
         Top             =   675
         Width           =   1245
      End
      Begin VB.CommandButton cmdAmpliar 
         Caption         =   "Ampliar"
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   6.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   645
         Left            =   10290
         Picture         =   "mdiPuntoVenta.frx":136B0
         Style           =   1  'Graphical
         TabIndex        =   18
         Top             =   0
         Width           =   1245
      End
      Begin VB.CommandButton cmdOpcion10 
         Caption         =   "Ctas Corrient"
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   6.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   645
         Left            =   2580
         Picture         =   "mdiPuntoVenta.frx":13802
         Style           =   1  'Graphical
         TabIndex        =   17
         Top             =   675
         Width           =   1245
      End
      Begin VB.CommandButton cmdOpcion3 
         Caption         =   "&Cierre"
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   6.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   645
         Left            =   2580
         Picture         =   "mdiPuntoVenta.frx":13D8C
         Style           =   1  'Graphical
         TabIndex        =   16
         Top             =   0
         Width           =   1245
      End
      Begin VB.CommandButton cmdOpcion12 
         Caption         =   "Carta de Productos"
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   6.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   645
         Left            =   5160
         Picture         =   "mdiPuntoVenta.frx":13E8E
         Style           =   1  'Graphical
         TabIndex        =   15
         Top             =   675
         Width           =   1245
      End
      Begin VB.CommandButton cmdOpcion13 
         Caption         =   "Delivery en Transito"
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   6.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   645
         Left            =   6450
         Picture         =   "mdiPuntoVenta.frx":14418
         Style           =   1  'Graphical
         TabIndex        =   14
         Top             =   675
         Width           =   1245
      End
      Begin VB.CommandButton cmdOpcion9 
         Caption         =   "Documentos"
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   6.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   645
         Left            =   1290
         Picture         =   "mdiPuntoVenta.frx":149A2
         Style           =   1  'Graphical
         TabIndex        =   13
         Top             =   675
         Width           =   1245
      End
      Begin VB.CommandButton cmdOpcion15 
         Caption         =   "Activar Pin Pad"
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   6.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   645
         Left            =   10290
         Picture         =   "mdiPuntoVenta.frx":14F2C
         Style           =   1  'Graphical
         TabIndex        =   12
         Top             =   675
         Visible         =   0   'False
         Width           =   1245
      End
      Begin VB.CommandButton cmdOpcion8 
         Caption         =   "Pedidos"
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   6.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   645
         Left            =   0
         Picture         =   "mdiPuntoVenta.frx":152B6
         Style           =   1  'Graphical
         TabIndex        =   11
         Top             =   675
         Width           =   1245
      End
      Begin VB.CommandButton cmdOpcion14 
         Caption         =   "Delivery Entregados"
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   6.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   645
         Left            =   7740
         Picture         =   "mdiPuntoVenta.frx":15400
         Style           =   1  'Graphical
         TabIndex        =   10
         Top             =   675
         Width           =   1245
      End
      Begin VB.CommandButton cmdOpcion6 
         Caption         =   "Recibos de Ingresos"
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   6.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   645
         Left            =   6450
         Picture         =   "mdiPuntoVenta.frx":1598A
         Style           =   1  'Graphical
         TabIndex        =   9
         Top             =   0
         Width           =   1245
      End
      Begin VB.CommandButton cmdOpcion11 
         Caption         =   "Ctas x Cobrar"
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   6.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   645
         Left            =   3870
         Picture         =   "mdiPuntoVenta.frx":15A8C
         Style           =   1  'Graphical
         TabIndex        =   8
         Top             =   675
         Width           =   1245
      End
      Begin VB.CommandButton cmdOpcion16 
         Caption         =   "Salir"
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   6.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   645
         Left            =   9000
         Picture         =   "mdiPuntoVenta.frx":16016
         Style           =   1  'Graphical
         TabIndex        =   7
         Top             =   0
         Width           =   1245
      End
      Begin VB.CommandButton cmdOpcion4 
         Caption         =   "&Mesas"
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   6.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   645
         Left            =   3870
         Picture         =   "mdiPuntoVenta.frx":16108
         Style           =   1  'Graphical
         TabIndex        =   6
         Top             =   0
         Width           =   1245
      End
      Begin VB.CommandButton cmdOpcion2 
         Caption         =   "&Punto Venta"
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   6.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   645
         Left            =   1320
         Picture         =   "mdiPuntoVenta.frx":16202
         Style           =   1  'Graphical
         TabIndex        =   5
         Top             =   0
         Width           =   1245
      End
      Begin VB.CommandButton cmdOpcion1 
         Caption         =   "&Apertura"
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   6.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   645
         Left            =   0
         Picture         =   "mdiPuntoVenta.frx":16304
         Style           =   1  'Graphical
         TabIndex        =   4
         Top             =   0
         Width           =   1245
      End
      Begin VB.CommandButton cmdOpcion5 
         Caption         =   "Recibos de Egresos"
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   6.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   645
         Left            =   5160
         Picture         =   "mdiPuntoVenta.frx":16406
         Style           =   1  'Graphical
         TabIndex        =   3
         Top             =   0
         Width           =   1245
      End
      Begin VB.CommandButton cmdOpcion7 
         Caption         =   "Reservas"
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   6.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   645
         Left            =   7740
         Picture         =   "mdiPuntoVenta.frx":16508
         Style           =   1  'Graphical
         TabIndex        =   2
         Top             =   0
         Width           =   1245
      End
      Begin VB.Image ImagePais 
         Height          =   360
         Left            =   12840
         Top             =   80
         Width           =   645
      End
   End
   Begin MSComctlLib.StatusBar StatusBar 
      Align           =   2  'Align Bottom
      Height          =   300
      Left            =   0
      TabIndex        =   0
      Top             =   9660
      Width           =   13560
      _ExtentX        =   23918
      _ExtentY        =   529
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   8
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            AutoSize        =   1
            Bevel           =   2
            Object.Width           =   4983
            MinWidth        =   3528
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   2
            Object.Width           =   3889
            MinWidth        =   3881
         EndProperty
         BeginProperty Panel4 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   3881
            MinWidth        =   3881
         EndProperty
         BeginProperty Panel5 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   1
            Alignment       =   1
            AutoSize        =   2
            Enabled         =   0   'False
            Object.Width           =   1773
            MinWidth        =   1764
            TextSave        =   "MAYÚS"
         EndProperty
         BeginProperty Panel6 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   2
            Alignment       =   1
            AutoSize        =   2
            Object.Width           =   1773
            MinWidth        =   1764
            TextSave        =   "NÚM"
         EndProperty
         BeginProperty Panel7 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   6
            Alignment       =   1
            AutoSize        =   2
            Object.Width           =   2196
            MinWidth        =   2205
            Text            =   "Now"
            TextSave        =   "27/02/2019"
         EndProperty
         BeginProperty Panel8 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   5
            Alignment       =   1
            AutoSize        =   2
            Object.Width           =   2196
            MinWidth        =   2205
            TextSave        =   "02:43 p.m."
         EndProperty
      EndProperty
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
   Begin VB.Menu mnuMovimiento 
      Caption         =   "M&ovimientos"
      Begin VB.Menu mnuInicio 
         Caption         =   "&Apertura de Turno"
      End
      Begin VB.Menu mnuVenta 
         Caption         =   "&Punto de Venta"
      End
      Begin VB.Menu mnuPinPad 
         Caption         =   "Activar PinPad (No Financiera)"
      End
      Begin VB.Menu mnuCierre 
         Caption         =   "Ci&erre de Turno"
      End
      Begin VB.Menu linea8 
         Caption         =   "-"
      End
      Begin VB.Menu mnuCorrelativo 
         Caption         =   "Correlativo de Documentos"
      End
      Begin VB.Menu linea7 
         Caption         =   "-"
      End
      Begin VB.Menu mnuMesa 
         Caption         =   "&Mesas"
      End
      Begin VB.Menu mnuInsumoCritico 
         Caption         =   "&Insumos/Platos de Stock Crítico"
      End
      Begin VB.Menu mnuImportacion 
         Caption         =   "&Importación de Requerimientos"
      End
   End
   Begin VB.Menu mnuCuentas 
      Caption         =   "&Correlativos"
      Begin VB.Menu mnuCorrelativoPedido 
         Caption         =   "Correlativo de Pedidos"
      End
      Begin VB.Menu mnuCorrelativoDocumento 
         Caption         =   "Correlativo de Documentos"
      End
      Begin VB.Menu mnuGuiaTransporte 
         Caption         =   "Correlativo de Guias de Transporte"
      End
      Begin VB.Menu mnuCtaCte 
         Caption         =   "Cuentas Corrientes"
      End
      Begin VB.Menu mnuRecibo 
         Caption         =   "Recibos de Egreso"
      End
      Begin VB.Menu mnuReciboIngreso 
         Caption         =   "Recibos de Ingreso"
      End
      Begin VB.Menu mnuNotaCredito 
         Caption         =   "Notas de Crédito"
      End
      Begin VB.Menu mnuReserva 
         Caption         =   "Reservas"
      End
      Begin VB.Menu mnuDocumentoElectronico 
         Caption         =   "Documentos Electrónicos"
      End
      Begin VB.Menu mnuCorrelativoCentralPedidos 
         Caption         =   "Pedidos Central Producción"
      End
      Begin VB.Menu linea6 
         Caption         =   "-"
      End
      Begin VB.Menu mnuCuentaCobrar 
         Caption         =   "Cuenta por Cobrar"
      End
   End
   Begin VB.Menu mnuConexion 
      Caption         =   "C&onexión"
      Begin VB.Menu mnuCambiaLocal 
         Caption         =   "Cambiar de Local"
         Shortcut        =   {F8}
      End
   End
   Begin VB.Menu mnuAyuda 
      Caption         =   "A&yuda"
      Begin VB.Menu mnuAcerca 
         Caption         =   "Acerca de ..."
      End
   End
   Begin VB.Menu mnuSalir 
      Caption         =   "&Salir"
   End
End
Attribute VB_Name = "mdiPuntoVenta"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Const IMAGESIZE = 0.566893424036281

Private Sub cmdAmpliar_Click()
    If cmdAmpliar.Caption = "Esconder" Then
       xPicture.Height = 700
       cmdAmpliar.Caption = "Ampliar"
    Else
       xPicture.Height = 1380
       cmdAmpliar.Caption = "Esconder"
    End If
End Sub

Private Sub cmdConsultaSaldo_Click()
   frmConsultaSaldo.Show vbModal
End Sub

Private Sub cmdOpcion1_Click()
    If lDiaContable = False Then 'manual
        frmDiaContable.obtieneModoIngreso "Apertura"
        frmDiaContable.Show vbModal
        StatusBar.Panels.Item(1).Text = "Día Contable : " & obtieneDiaContable
    End If




    If lDiaContable = False Then ' manual
                    If lDiaContableAperturado = True Then
                            If lMultiCajero = True Then
                                If validaInicioCajaRapida = False Then
                                        frmInicio.Show vbModal
                                End If
                            Else
                                frmInicio.Show vbModal
                                
                            End If
                    End If
            Else ' automatico
            
            If lMultiCajero = True Then
                    If validaInicioCajaRapida = False Then
                                        frmInicio.Show vbModal
                                End If
                Else
                frmInicio.Show vbModal
            End If
    End If
    
    
    
    
    StatusBar.Panels.Item(4).Text = IIf(wInicio, "Turno : " & sTurno, "Turno : No Iniciado")
    
    
    cmdOpcion1.Enabled = IIf(wInicio, False, True)
    mnuInicio.Enabled = IIf(wInicio, False, True)
End Sub

Public Function validaInicioCajaRapida() As Boolean
    validaInicioCajaRapida = True ' ingresadirecto
    Dim rsTipoCambio As New ADODB.Recordset
    
    'valida tipo cambio
    Set rsTipoCambio = Lib.OpenRecordset("SELECT * From TTIPOCAMBIO WHERE (fFecha = {fn CURDATE() })", Cn)
    If rsTipoCambio.EOF Or rsTipoCambio.BOF Then
    
        validaInicioCajaRapida = False ' pide tipocambio
        Exit Function
    Else
        nTC = IIf(IsNull(rsTipoCambio!nVenta), 0, IIf(IsNull(rsTipoCambio!nVenta), 0, rsTipoCambio!nVenta))
         If nTC = 0 Then
                validaInicioCajaRapida = False ' pide tipocambio
                Exit Function
        End If
         
         
    End If
    'valida existencia de turno
    Dim rstTurno As New ADODB.Recordset
    Dim nAbonoN As Double
    Dim nAbonoE As Double
    If lMCPV Then
      Isql = "select * from MTURNO where tUsuario ='" & sUsuario & "' order by tTurno"
   Else
      Isql = "select * from MTURNO where tCaja ='" & sCaja & "' order by tTurno"
   End If
   Set rstTurno = Lib.OpenRecordset(Isql, Cn)
   If rstTurno.RecordCount = 0 Then
        validaInicioCajaRapida = False ' pide turno
        Exit Function
    Else
      rstTurno.MoveLast
                  
      If Not rstTurno!lCierre = True Then
                sTurno = rstTurno!tTurno
                nAbonoN = IIf(IsNull(rstTurno!nMontoIN), 0, rstTurno!nMontoIN)
                nAbonoE = IIf(IsNull(rstTurno!nMontoIE), 0, rstTurno!nMontoIE)
                
         
                  Isql = "update MTURNO set " & _
                         "tUsuario ='" & sUsuario & "', " & _
                         "nMontoIN = " & nAbonoN & ", " & _
                         "nMontoIE = " & nAbonoE & " " & _
                         "where tTurno ='" & sTurno & "'"
                   Cn.Execute Isql
                   ActivaInicio (True)
                    wInicio = True
                    
                    validaInicioCajaRapida = True ' ingresadirecto
       Else
            validaInicioCajaRapida = False ' pide turno
                Exit Function
       End If
    End If
End Function



Private Sub cmdOpcion10_Click()
    frmCtaCte.Show
End Sub

Private Sub cmdOpcion11_Click()
    frmCuentaCobrar.Show
End Sub

Private Sub cmdOpcion12_Click()
    frmPrecios.Show
End Sub

Private Sub cmdOpcion13_Click()
    frmPedidoDelivery.Show
End Sub

Private Sub cmdOpcion14_Click()
    frmPedidoDeliveryNo.Show
End Sub

Private Sub cmdOpcion15_Click()
    Call mnuPinPad_Click
End Sub

Private Sub cmdOpcion16_Click()
    'OO---------------------------------------
    If lMCPV Then
        'Metodo para cerrar todo e Inicializar()
        frmMozoUsuario.Show vbModal
        InicializaMCPV
    Else
        Salir
    End If
    '---------------------------------------
End Sub

Private Sub cmdOpcion2_Click()

'Oscar Ortega
    sTipo = "TC"
    If nTC = 0 Then
        frmNumPad.Show vbModal
        If wEnter Then
            nTC = Val(sDescrip)
            
            Dim oComando As clsComando
            Set oComando = New clsComando
            If Not oComando.CreateCmdSp("spIns_TipoCambio", Cn) Then
                Set oComando = Nothing
                Exit Sub
            End If
            
            oComando.CreateParameter "@nTc", adDouble, adParamInput, 0, nTC
            oComando.CreateParameter "@tUSUARIO", adVarChar, adParamInput, 15, sUsuario
            oComando.CreateParameter "@nTCO", adDouble, adParamInput, 0, nTCO

            If Not oComando.GetParamOK Then
                Set oComando = Nothing
                Exit Sub
            End If
            If Not oComando.ExecSP Then
                Set oComando = Nothing
                Exit Sub
            End If
            
        End If
    End If

   Screen.MousePointer = vbHourglass
   If lCajaRapida Or lMultiCajero Then
      frmCajaRapida.Show vbModal
      If Not wEnter Then
         frmVenta.Show vbModal
      End If
   Else
      frmVenta.Show vbModal
      
      If lMCPV Then
         frmMozoUsuario.Show vbModal
         InicializaMCPV
      End If
  End If
End Sub

Private Sub cmdOpcion3_Click()
    frmLiquidacionDetalle.Show vbModal

    If wInicio = False Then
    ActivaInicio (False)
    End If
End Sub

Private Sub cmdOpcion4_Click()
    Screen.MousePointer = vbHourglass
    sTipo = "V"
    frmMesaConsulta.Show
End Sub

Private Sub cmdOpcion5_Click()
    frmReciboEgreso.Show
End Sub

Private Sub cmdOpcion6_Click()
    frmReciboIngreso.Show
End Sub

Private Sub cmdOpcion7_Click()
    frmReserva.Show
End Sub

Private Sub cmdOpcion8_Click()
    frmPedidoCorrelativo.Show
End Sub

Private Sub cmdOpcion9_Click()
    frmDocumentoCorrelativo.Show
End Sub

'CGMiranda-------------------------------------------------
Private Sub cmdOpcionMensaje_Click()
    frmMensajeCocina.Show
End Sub
'Fin CGMiranda---------------------------------------------


Private Sub MDIForm_Resize()
On Error Resume Next
    Dim ImageWidth As Single
    Dim ImageHeight As Single
    picStretch.Visible = False
    picStretch.AutoRedraw = True
    picStretch.Height = Me.Height
    picStretch.Width = Me.Width
    ImageWidth = picStretch.Picture.Width * IMAGESIZE
    ImageHeight = (picStretch.Picture.Height * IMAGESIZE) + 3000
    picStretch.PaintPicture picStretch.Picture, 0, 0, Me.Width, Me.Height, 0, 0, ImageWidth, ImageHeight
    Set Me.Picture = picStretch.Image
End Sub

Private Sub MDIForm_Load()
    Centrar Me
    Me.Caption = "Punto de Venta : Local " & localConectado
    'OO--------------------------------
    If lMCPV Then
        Me.Visible = False
        frmMozoUsuario.Show vbModal
        Me.Visible = True
    End If
    '----------------------------------
    ActivaInicio (False)
    If nPuerto > 0 Then
       Visor String(Int((20 - Len(tMensaje1)) / 2), " ") & tMensaje1, String(Int((20 - Len(tMensaje2)) / 2), " ") & tMensaje2, nPuerto, "N"
    End If
    
    If Not lAlmacen Then
       mnuImportacion.Enabled = False
      ' mnuCorrelativoCentralPedidos.Enabled = False
    End If
    
    
    If Not lFacturacionE Then
            mnuDocumentoElectronico.Enabled = False
    Else
       If lFEOfisis Then
            mnuDocumentoElectronico.Enabled = False
       End If
    End If
    
    
    
'    If pais = "000" Then
'        If lFacturacionE Then
'            If lFESpring Then
'                mnuNotaCredito.Enabled = False
'            End If
'        End If
'    End If
    
    
    If lMCPV Then
        InicializaMCPV
    End If
    StatusBar.Panels.Item(1).Text = "Día Contable : No Iniciado"
    StatusBar.Panels.Item(2).Text = "Caja : " & sCaja
    StatusBar.Panels.Item(3).Text = "Usuario : " & IIf(Mid(sUsuario, 1, 1) = "*", Mid(sUsuario, 2, 15), sUsuario)
    StatusBar.Panels.Item(4).Text = IIf(wInicio, "Turno :" & sTurno, "Turno : No Iniciado")
    
    If lSiab Then
       cmdConsultaSaldo.Visible = True
    End If
       
    If lVisaNet Then
       cmdOpcion15.Visible = True
       Open App.Path & "\DLL3500.INI" For Input As #1
       Do While Not EOF(1)
          Dim xStr As String
          Input #1, sLinea
          
          Select Case Mid(sLinea, 1, 4)
                 Case "HOST"
                      xStr = Trim(Mid(sLinea, InStr(1, sLinea, "=") + 1))
                      IpPinPad = Mid(xStr, 2, Len(xStr) - 2)
                 
                 Case "PORT"
                       xStr = Trim(Mid(sLinea, InStr(1, sLinea, "=") + 1))
                       xStr = Mid(xStr, 2, Len(xStr) - 2)
                       IpPort = Val(xStr)
          
                 Case "APPL"
                       xStr = Trim(Mid(sLinea, InStr(1, sLinea, "=") + 1))
                       xStr = Mid(xStr, 2, Len(xStr) - 2)
                       nTimeOut = Val(xStr)
                       nTimeOut = nTimeOut * 2
          End Select
       Loop
       Close #1

       If Not ValidaIP(IpPinPad, IpPort) Then
'          MsgBox "Error de conexión", vbCritical, "VisaNet"
          lVisaNet = False
'          mnuPinPad.Visible = False
'          Exit Sub
       End If
    
       nRet = fiOpenPort(App.Path & "\DLL3500.INI")
       If nRet = RET_NOK Then
          MsgBox "Error de Puerto", vbCritical, "VisaNet"
          lVisaNet = False
          mnuPinPad.Visible = False
          Exit Sub
       End If
    Else
       mnuPinPad.Visible = False
    End If
    If multiLocal = False Then
    mnuConexion.Visible = False
    End If
   
    Call Accesos(Me, "02", sUsuario)
    On Error Resume Next
    Me.Picture = LoadPicture(App.Path & "\bmps\Inforest.EMF")

 
    'IMAGEN PAIS
    Dim fso As Object
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
    
'    If fso.FileExists(App.Path & "\bmps\Pais.jpg") Then
'       ImagePais.Picture = LoadPicture(App.Path & "\bmps\Pais.jpg")
'    End If
  
    Set fso = Nothing
    
End Sub

Public Sub InicializaMCPV()
    Dim RsTurno As Recordset
    Isql = "select * from MTURNO where tUsuario ='" & sUsuario & "' and lcierre=0 order by tTurno"
    Set RsTurno = Lib.OpenRecordset(Isql, Cn)
    If RsTurno.RecordCount > 0 Then
        nTC = Calcular("SELECT nVenta as Codigo From TTIPOCAMBIO WHERE (fFecha = {fn CURDATE() })", Cn)
        If nTC = 0 Then
            wInicio = True
            ActivaInicio (False)
            cmdOpcion1.Enabled = True
            sTurno = RsTurno!tTurno
        Else
            wInicio = True
            ActivaInicio (True)
            cmdOpcion1.Enabled = False
            sTurno = RsTurno!tTurno
        End If
    Else
         wInicio = False
         ActivaInicio (False)
         cmdOpcion1.Enabled = True
         sTurno = ""
    End If
    StatusBar.Panels.Item(2).Text = "Caja : " & sCaja
    StatusBar.Panels.Item(3).Text = "Usuario : " & IIf(Mid(sUsuario, 1, 1) = "*", Mid(sUsuario, 2, 15), sUsuario)
    StatusBar.Panels.Item(4).Text = IIf(wInicio, "Turno :" & sTurno, "Turno : No Iniciado")
        
    Call Accesos(Me, "02", sUsuario)
End Sub

Public Sub Salir()
   If lMCPV Then
      If nPuerto > 0 Then
         Visor "", "", nPuerto, "N"
      End If
      
      If lVisaNet Then
         fiClosePort
      End If
      Unload Me
   Else
'      sino = MsgBox("Deseas Salir del Sistema", vbDefaultButton1 + vbYesNo + vbQuestion, sMensaje)
'      If sino = vbYes Then
         If nPuerto > 0 Then
            Visor "", "", nPuerto, "N"
         End If
           
         If lVisaNet Then
            fiClosePort
         End If
         If lHARDkey Then
            '--------Libera Licencia de la Llave------------------------
            Dim Verifica As Boolean
            Verifica = hk.FinalizarConexion(Aplicacion.PuntoVenta)
            '-----------------------------------------------------------
         End If
         
                
         'auditoria
        
         registroAccesoAuditoria "S", sUsuario
         'auditoria
         End
      'End If
  End If
End Sub

Private Sub MDIForm_Unload(Cancel As Integer)
    If lHARDkey Then
        '--------Libera Licencia de la Llave------------------------
        Dim Verifica As Boolean
        Verifica = hk.FinalizarConexion(Aplicacion.PuntoVenta)
        '-----------------------------------------------------------
    End If
           
 'auditoria

 registroAccesoAuditoria "S", sUsuario
 'auditoria
 
  '  End
End Sub

Private Sub mnuAcerca_Click()
   frmAbout.Show vbModal
End Sub

Private Sub mnuAnulado_Click()
   frmAnulado.Show
End Sub

Private Sub mnuCancelado_Click()
   frmCancelado.Show
End Sub

Private Sub mnuCambiaLocal_Click()
        Dim verificaForms As Boolean
        verificaForms = VerificaFormAbiertos
        If verificaForms <> True Then
            frmServidorEnlace.Show
        Else
            MsgBox "Debe cerrar todos los formularios", vbInformation, sMensaje
        End If
End Sub

Private Sub mnuCierre_Click()
   cmdOpcion3_Click
End Sub

Private Sub mnuClienteDeuda_Click()
   frmRepClienteDeuda.Show
End Sub

Private Sub mnuCorrelativo_Click()
   frmFactura.Show vbModal
End Sub

Private Sub mnuCorrelativoCentralPedidos_Click()
  frmCentralPedidos.Show
End Sub

Private Sub mnuCorrelativoDocumento_Click()
   frmDocumentoCorrelativo.Show
End Sub

Private Sub mnuCorrelativoPedido_Click()
   frmPedidoCorrelativo.Show
End Sub

Private Sub mnuCtaCte_Click()
   frmCtaCte.Show
End Sub

Private Sub mnuCuentaCorriente_Click()
   frmCuentaCobrar.Show
End Sub

Private Sub mnuCuentaCobrar_Click()
   frmCuentaCobrar.Show
End Sub

Private Sub mnuDocumentoElectronico_Click()
   frmDocumentoElectronicoCorrelativo.Show
End Sub

Private Sub mnuGuiaTransporte_Click()
   frmGuiaTransporte.Show
End Sub

Private Sub mnuImportacion_Click()
   frmImportacionRequerimientos.Show
End Sub

Private Sub mnuInicio_Click()
   cmdOpcion1_Click
End Sub

Private Sub mnuInsumoCritico_Click()
    frmInsumo.Show
End Sub

Private Sub mnuMesa_Click()
   Screen.MousePointer = vbHourglass
   sTipo = "V"
   frmMesaConsulta.Show vbModal
End Sub

Private Sub mnuPuntoVenta_Click()
   cmdOpcion2_Click
End Sub

Private Sub mnuRanking_Click()
   frmRepRanking.Show
End Sub

Private Sub mnuNotaCredito_Click()
   frmNotaCredito.Show
End Sub

Private Sub mnuPinPad_Click()
   Dim nRet As Integer
   Dim sOperacion As String
   Dim sRetorno As String * 512
   Dim lLoop As Boolean
   Dim nContador As Integer
      
   sOperacion = OP_NO_FINANCIERA & "A" & "0000000000.00" & Chr$(FS) & _
                                   "B" & "000000000000" & Chr$(FS) & _
                                   "C" & "0" & Chr$(FS) & _
                                   "D" & sEmpresa & Chr$(FS) & _
                                   "E" & sCaja

   nRet = fiStartOperation(sOperacion, 2, sRetorno)
      
   If nRet = RET_OK Or nRet = RET_RUNNING Then
      
      If Not Imprimir(sPreCuenta) Then
         Exit Sub
      End If
      Printer.FontName = sFont
      Printer.FontBold = False
      lLoop = True
      nContador = 0

      Do
        sRetorno = ""
        nRet = fiGetStatus(sRetorno, 512)
        MensajePinPad sRetorno
        
        Mensaje "PinPad Listo. Esperando...", "PinPad", 500
        nContador = nContador + 1
        If nContador >= nTimeOut Then
           If MsgBox("Tiempo de espera agotado, deseas mas tiempo?", vbExclamation + vbOKCancel, "VisaNet") = vbOK Then
              lLoop = True
              nContador = nTimeOut / 2
           Else
              lLoop = False
           End If
        End If
        
        If nRet <> "0" Then
           nContador = 0
        End If
      Loop While (Mid$(sRetorno, 5, 2) <> "C1") And lLoop
      Printer.EndDoc
   Else
      MsgBox "Error de conectividad", vbCritical, sMensaje
   End If
              
End Sub

Private Sub mnuRecibo_Click()
   cmdOpcion5_Click
End Sub

Private Sub mnuRegistroVenta_Click()
   frmRepRegistroVenta.Show
End Sub

Private Sub mnuReporteVenta_Click()
   frmRepPaloteo.Show
End Sub

Private Sub mnuReciboIngreso_Click()
   cmdOpcion6_Click
End Sub

Private Sub mnuReserva_Click()
   frmReserva.Show
End Sub

Private Sub mnuSalir_Click()
   Salir
End Sub

Private Sub mnuVenta_Click()
    Screen.MousePointer = vbHourglass
    frmVenta.Show vbModal
End Sub

Private Sub mnuVentaMensual_Click()
   frmRepVentaMensual.Show
End Sub

'UN EXE VARIAS BD
Public Sub reinicia()
    Unload Me
    mdiPuntoVenta.Show
End Sub
Private Sub Timer_Timer()
    If multiLocal = True Then
        If ultimoConectado = False Then
             frmServidorEnlace.Show
            Timer.Interval = 0
        End If
    End If
End Sub
 
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
End Function

