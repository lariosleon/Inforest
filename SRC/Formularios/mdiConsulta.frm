VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "mscomctl.ocx"
Begin VB.MDIForm mdiConsulta 
   BackColor       =   &H8000000C&
   Caption         =   "Sistema Integral para Restaurantes"
   ClientHeight    =   9945
   ClientLeft      =   -2700
   ClientTop       =   1155
   ClientWidth     =   13605
   Icon            =   "mdiConsulta.frx":0000
   LinkTopic       =   "MDIForm1"
   LockControls    =   -1  'True
   Begin VB.PictureBox picStretch 
      Align           =   1  'Align Top
      BorderStyle     =   0  'None
      Height          =   9000
      Left            =   0
      Picture         =   "mdiConsulta.frx":57E2
      ScaleHeight     =   9000
      ScaleWidth      =   13605
      TabIndex        =   12
      Top             =   615
      Visible         =   0   'False
      Width           =   13605
   End
   Begin VB.PictureBox Picture 
      Align           =   1  'Align Top
      Height          =   615
      Left            =   0
      ScaleHeight     =   555
      ScaleWidth      =   13545
      TabIndex        =   0
      Top             =   0
      Width           =   13605
      Begin VB.CommandButton cmdOpcion6 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   525
         Left            =   3705
         Picture         =   "mdiConsulta.frx":1154D
         Style           =   1  'Graphical
         TabIndex        =   10
         Top             =   0
         Width           =   570
      End
      Begin VB.CommandButton cmdOpcion9 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   525
         Left            =   0
         Picture         =   "mdiConsulta.frx":1198F
         Style           =   1  'Graphical
         TabIndex        =   9
         Top             =   0
         Width           =   570
      End
      Begin VB.CommandButton cmdOpcion1 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   525
         Left            =   585
         Picture         =   "mdiConsulta.frx":11C99
         Style           =   1  'Graphical
         TabIndex        =   7
         Top             =   0
         Width           =   570
      End
      Begin VB.CommandButton cmdOpcion2 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   525
         Left            =   1290
         Picture         =   "mdiConsulta.frx":13963
         Style           =   1  'Graphical
         TabIndex        =   6
         Top             =   0
         Width           =   570
      End
      Begin VB.CommandButton cmdOpcion3 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   525
         Left            =   1860
         Picture         =   "mdiConsulta.frx":13DA5
         Style           =   1  'Graphical
         TabIndex        =   5
         Top             =   0
         Width           =   570
      End
      Begin VB.CommandButton cmdOpcion4 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   525
         Left            =   2565
         Picture         =   "mdiConsulta.frx":141E7
         Style           =   1  'Graphical
         TabIndex        =   4
         Top             =   0
         Width           =   570
      End
      Begin VB.CommandButton cmdOpcion5 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   525
         Left            =   3135
         Picture         =   "mdiConsulta.frx":15969
         Style           =   1  'Graphical
         TabIndex        =   3
         Top             =   0
         Width           =   570
      End
      Begin VB.CommandButton cmdOpcion8 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   525
         Left            =   4980
         Picture         =   "mdiConsulta.frx":15C73
         Style           =   1  'Graphical
         TabIndex        =   2
         Top             =   0
         Width           =   570
      End
      Begin VB.CommandButton cmdOpcion7 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   525
         Left            =   4275
         Picture         =   "mdiConsulta.frx":15D65
         Style           =   1  'Graphical
         TabIndex        =   1
         Top             =   0
         Width           =   570
      End
      Begin VB.Image ImagePais 
         Height          =   360
         Left            =   12840
         Top             =   40
         Width           =   645
      End
   End
   Begin MSComctlLib.ImageList ImageList 
      Left            =   90
      Top             =   6150
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   32
      ImageHeight     =   32
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   10
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mdiConsulta.frx":1606F
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mdiConsulta.frx":16173
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mdiConsulta.frx":1648F
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mdiConsulta.frx":168E3
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mdiConsulta.frx":16D37
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mdiConsulta.frx":184CB
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mdiConsulta.frx":187E7
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mdiConsulta.frx":18B03
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mdiConsulta.frx":1B097
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mdiConsulta.frx":1B0F5
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComDlg.CommonDialog dlgPrinter 
      Left            =   720
      Top             =   6150
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.PictureBox Picture1 
      Align           =   1  'Align Top
      Height          =   0
      Left            =   0
      ScaleHeight     =   0
      ScaleWidth      =   13605
      TabIndex        =   8
      Top             =   9615
      Width           =   13605
   End
   Begin MSComctlLib.StatusBar StatusBar 
      Align           =   2  'Align Bottom
      Height          =   300
      Left            =   0
      TabIndex        =   11
      Top             =   9645
      Width           =   13605
      _ExtentX        =   23998
      _ExtentY        =   529
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   7
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Bevel           =   2
            Object.Width           =   8978
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
            Style           =   1
            Alignment       =   1
            AutoSize        =   2
            Enabled         =   0   'False
            Object.Width           =   1773
            MinWidth        =   1764
            TextSave        =   "MAY�S"
         EndProperty
         BeginProperty Panel5 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   2
            Alignment       =   1
            AutoSize        =   2
            Object.Width           =   1773
            MinWidth        =   1764
            TextSave        =   "N�M"
         EndProperty
         BeginProperty Panel6 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   6
            Alignment       =   1
            AutoSize        =   2
            Object.Width           =   2196
            MinWidth        =   2205
            Text            =   "Now"
            TextSave        =   "26/03/2019"
         EndProperty
         BeginProperty Panel7 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   5
            Alignment       =   1
            AutoSize        =   2
            Object.Width           =   2196
            MinWidth        =   2205
            TextSave        =   "06:30 a.m."
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
   Begin VB.Menu mnuCuentas 
      Caption         =   "&Correlativos"
      Begin VB.Menu mnuCorrelativoPedido 
         Caption         =   "Pedidos"
      End
      Begin VB.Menu mnuCorrelativoDocumento 
         Caption         =   "Documentos"
      End
      Begin VB.Menu mnuCtaCte 
         Caption         =   "Cuentas Corrientes"
      End
      Begin VB.Menu mnuCuentaCobrar 
         Caption         =   "Cuentas por Cobrar"
      End
      Begin VB.Menu mnunotacredito 
         Caption         =   "Notas de Cr�dito"
      End
      Begin VB.Menu mnuReserva 
         Caption         =   "Reservas"
      End
      Begin VB.Menu Linea3 
         Caption         =   "-"
         Index           =   0
      End
      Begin VB.Menu mnuRecibo 
         Caption         =   "Recibos de Egresos"
      End
      Begin VB.Menu mnuIngreso 
         Caption         =   "Recibos de Ingreso"
      End
      Begin VB.Menu Linea1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuTurno 
         Caption         =   "Turnos"
      End
   End
   Begin VB.Menu mnuReporte 
      Caption         =   "&Reportes"
      Begin VB.Menu mnuControl 
         Caption         =   "De Control"
         Begin VB.Menu mnuReporNotacredito 
            Caption         =   "Notas de Credito Trazabilidad"
            Visible         =   0   'False
         End
         Begin VB.Menu mnuLiquidacion 
            Caption         =   "Liquidaci�n de Cajero"
         End
         Begin VB.Menu mnuLiquidacion2 
            Caption         =   "Liquidaci�n de Cajero Formato 2"
         End
         Begin VB.Menu Linea8 
            Caption         =   "-"
         End
         Begin VB.Menu mnuPaloteo 
            Caption         =   "Paloteo de Producci�n"
         End
         Begin VB.Menu mnuPropiedades 
            Caption         =   "Paloteo de Propiedades"
         End
         Begin VB.Menu mnuPaloteoInsumo 
            Caption         =   "Paloteo de Insumos"
         End
         Begin VB.Menu mnuRepEquivalencias 
            Caption         =   "Paloteo de Equivalencias"
            Enabled         =   0   'False
            Visible         =   0   'False
         End
         Begin VB.Menu mnuPaloteoOfertas 
            Caption         =   "Paloteo de Ofertas"
         End
         Begin VB.Menu mnuPaloteoProductoMes 
            Caption         =   "Paloteo de Productos por Meses"
         End
         Begin VB.Menu Linea10 
            Caption         =   "-"
         End
         Begin VB.Menu mnuComanda 
            Caption         =   "Comandas"
         End
         Begin VB.Menu mnuDocumentos 
            Caption         =   "Documentos"
         End
         Begin VB.Menu Linea6 
            Caption         =   "-"
         End
         Begin VB.Menu mnuCorrelativo 
            Caption         =   "Pedidos"
         End
         Begin VB.Menu mnuProductosNoEnlazados 
            Caption         =   "Productos No Enlazados"
         End
         Begin VB.Menu mnuPropina 
            Caption         =   "Propinas"
         End
         Begin VB.Menu mnuCortesia 
            Caption         =   "Cortesias"
         End
         Begin VB.Menu mnuDescuento 
            Caption         =   "Descuentos"
         End
         Begin VB.Menu mnuRepCtaCte 
            Caption         =   "Cuentas Corrientes"
         End
         Begin VB.Menu mnuClienteDeuda 
            Caption         =   "Cuentas por Cobrar"
         End
         Begin VB.Menu mnuContacto 
            Caption         =   "Contactos"
         End
         Begin VB.Menu mnuMensajesUsuarios 
            Caption         =   "Mensajes Usuarios"
         End
         Begin VB.Menu Linea11 
            Caption         =   "-"
         End
         Begin VB.Menu mnuEnviosAutorizados 
            Caption         =   "Control de Envios Autorizados"
         End
         Begin VB.Menu mnuAnulado 
            Caption         =   "Control de Transacciones"
         End
         Begin VB.Menu mnuDiferencias 
            Caption         =   "Diferencias entre Paloteo y Liquidacion"
         End
         Begin VB.Menu mnuDescargo 
            Caption         =   "Descargo de Ventas"
         End
      End
      Begin VB.Menu mnuContables 
         Caption         =   "Contables"
         Begin VB.Menu mnuRegistroVenta 
            Caption         =   "Registro de Ventas"
         End
         Begin VB.Menu mnuPrincipal 
            Caption         =   "Principales Clientes"
         End
         Begin VB.Menu mnuCobranza 
            Caption         =   "Cobranzas (fecha documento)"
         End
      End
      Begin VB.Menu mnuEstadistico 
         Caption         =   "Estad�sticos"
         Begin VB.Menu mnuRanking 
            Caption         =   "Ranking de Producci�n"
         End
         Begin VB.Menu mnuRepEntregasReg 
            Caption         =   "Registros (Central de Pedidos)"
         End
         Begin VB.Menu mnuRepEntregas 
            Caption         =   "Entregas (Central de Pedidos)"
         End
         Begin VB.Menu linea5 
            Caption         =   "-"
         End
         Begin VB.Menu mnuMozo 
            Caption         =   "Anal�tico de Productos por Mesero"
         End
         Begin VB.Menu mnuMotorizado 
            Caption         =   "Analitico de Productos por Motorizados"
         End
         Begin VB.Menu mnuFrecuente 
            Caption         =   "Anal�tico de Clientes Frecuentes"
         End
         Begin VB.Menu mnuEstPropina 
            Caption         =   "Producci�n por Mesero"
            Visible         =   0   'False
         End
         Begin VB.Menu mnuPlanillaMotorizados 
            Caption         =   "Planilla de Motorizados"
         End
         Begin VB.Menu linea7 
            Caption         =   "-"
         End
         Begin VB.Menu mnuTiempoSalon 
            Caption         =   "Tiempos en Salon"
         End
         Begin VB.Menu mnuTiempoDelivery 
            Caption         =   "Tiempos Delivery"
         End
         Begin VB.Menu mnuTiempoKDS 
            Caption         =   "Tiempos KDS"
         End
         Begin VB.Menu mnuTiempoChefControl 
            Caption         =   "Tiempos Chef Control"
         End
         Begin VB.Menu mnudiferenciaDelivery 
            Caption         =   "Diferencias de Tiempos Delivery"
         End
         Begin VB.Menu liena1 
            Caption         =   "-"
         End
         Begin VB.Menu mnuRotacion 
            Caption         =   "Rotaci�n de Mesas"
         End
         Begin VB.Menu mnuOcupabilidad 
            Caption         =   "Ocupabilidad de Mesas"
         End
         Begin VB.Menu Linea9 
            Caption         =   "-"
            Index           =   0
         End
         Begin VB.Menu mnuPaloteoComparativo 
            Caption         =   "Paloteo Comparativo"
         End
         Begin VB.Menu mnuResultadoOperativo 
            Caption         =   "Resultados Operativos de Ventas"
         End
         Begin VB.Menu mnuRecargoalconsumo 
            Caption         =   "Distribuci�n de retenciones por forma de Pago"
         End
      End
      Begin VB.Menu mnuAnalitico 
         Caption         =   "Gerencial"
         Begin VB.Menu mnuVenta 
            Caption         =   "Venta Anual por Meses"
         End
         Begin VB.Menu mnuVentasFechas 
            Caption         =   "Venta Mensual por Fechas"
         End
         Begin VB.Menu mnuVentaComparada 
            Caption         =   "Venta Comparativa Anual"
         End
         Begin VB.Menu mnuVentaComparadaMensual 
            Caption         =   "Venta Comparativa Mensual"
         End
         Begin VB.Menu mnuVentaTurno 
            Caption         =   "Venta por Turnos"
         End
         Begin VB.Menu L9 
            Caption         =   "-"
         End
         Begin VB.Menu mnuCobranzaFecha 
            Caption         =   "Cobranza Mensual por Fechas"
         End
      End
   End
   Begin VB.Menu mnuTicketera 
      Caption         =   "&Ticketera"
      Begin VB.Menu mnuLiquidacionTicket 
         Caption         =   "Liquidaci�n de Cajero"
      End
      Begin VB.Menu mnuPaloteoTicket 
         Caption         =   "Paloteo"
      End
   End
   Begin VB.Menu mnuConexion 
      Caption         =   "&Conexi�n"
      Begin VB.Menu mnuCambiaLocal 
         Caption         =   "Cambiar de Local"
         Shortcut        =   {F8}
      End
   End
   Begin VB.Menu mnuAyuda 
      Caption         =   "&Ayuda"
      Begin VB.Menu mnuAcerca 
         Caption         =   "Acerca de ..."
      End
   End
   Begin VB.Menu mnuPaperlees 
      Caption         =   "&FE-Paperlees"
      Begin VB.Menu mnuFormato 
         Caption         =   "Formato de Intercambio � Cuadratura"
      End
      Begin VB.Menu mnuBajaFE 
         Caption         =   "Informe de Bajas"
      End
   End
   Begin VB.Menu mnugenerarsunat 
      Caption         =   "&Generar TXT SUNAT"
   End
   Begin VB.Menu mnuSalir 
      Caption         =   "&Salir"
   End
End
Attribute VB_Name = "mdiConsulta"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Const IMAGESIZE = 0.566893424036281
Private Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" ( _
                    ByVal hwnd As Long, _
                    ByVal lpOperation As String, _
                    ByVal lpFile As String, _
                    ByVal lpParameters As String, _
                    ByVal lpDirectory As String, _
                    ByVal nShowCmd As Long) As Long
Private Sub cmdOpcion1_Click()
    Screen.MousePointer = vbHourglass
    sTipo = "V"
    frmMesaConsulta.Show vbModal
End Sub

Private Sub cmdOpcion2_Click()
    frmPedidoCorrelativo.Show
End Sub

Private Sub cmdOpcion3_Click()
    frmDocumentoCorrelativo.Show
End Sub

Private Sub cmdOpcion4_Click()
    frmRepLiquidacion.Show vbModal
End Sub

Private Sub cmdOpcion5_Click()
    frmRepRegistroVenta.Show vbModal
End Sub

Private Sub cmdOpcion6_Click()
    frmRepPaloteo.Show vbModal
End Sub

Private Sub cmdOpcion7_Click()
    frmRepPropina.Show vbModal
End Sub

Private Sub cmdOpcion8_Click()
    Salir
End Sub

Private Sub cmdOpcion9_Click()
    On Error Resume Next
    dlgPrinter.ShowPrinter
End Sub
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
    StatusBar.Panels.Item(1).Text = "Local: " & localConectado
    StatusBar.Panels.Item(2).Text = "Caja : " & sCaja
    StatusBar.Panels.Item(3).Text = "Usuario : " & sUsuario
    Centrar Me
    
    If lFEpape Then
        mnuPaperlees.Visible = True
        mnugenerarsunat.Visible = True 'JCDPFH 170718
    Else
        mnuPaperlees.Visible = False
        mnugenerarsunat.Visible = False 'JCDPFH 170718
    End If
    
    If Not lAlmacen Then
       mnuPaloteoInsumo.Enabled = False
       mnuResultadoOperativo.Enabled = False
    End If
           
    If Not lInfhotel Then
       mnuContacto.Visible = False
    End If
    If multiLocal = False Then
        mnuConexion.Visible = False
      '  mnuCambiaLocal.Visible = False
    End If
    Call Accesos(Me, "04", sUsuario)
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

Public Sub Salir()
  ' sino = MsgBox("Deseas Salir del Sistema", vbDefaultButton1 + vbYesNo + vbQuestion, sMensaje)
  ' If sino = vbYes Then
    If lHARDkey Then
        '--------Libera Licencia de la Llave------------------------
        Dim Verifica As Boolean
        Verifica = hk.FinalizarConexion(Aplicacion.Consultas)
        '-----------------------------------------------------------
    End If
    'auditoria
    
    registroAccesoAuditoria "S", sUsuario
    
    
    'auditoria
    End
   'End If
End Sub

Private Sub MDIForm_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    If lHARDkey Then
        '--------Libera Licencia de la Llave------------------------
        Dim Verifica As Boolean
        Verifica = hk.FinalizarConexion(Aplicacion.Consultas)
        '-----------------------------------------------------------
    End If
        'auditoria
    
    registroAccesoAuditoria "S", sUsuario
    
    
    'auditoria

End Sub

Private Sub MDIForm_Unload(Cancel As Integer)
    'auditoria
    
    registroAccesoAuditoria "S", sUsuario
    
    
    'auditoria

End Sub

Private Sub mnuAcerca_Click()
   frmAbout.Show vbModal
End Sub

Private Sub mnuAnulado_Click()
    '10092018 ReporteExtraido
    
    Dim Ruta As String
    Dim nombre As String
    
    Dim RepFoRepTrans As String
    
    RepFoRepTrans = "" & sRuta & "|" & sMDB & "|" & sUserName & "|" & sUserPassword & "|"
    
    On Error GoTo Ir
        Ruta = App.Path & "\Reportes Extraidos\RepFoRepTrans.exe"
        nombre = Dir(Ruta)
        If Len(nombre) > 0 Then
            ShellExecute Me.hwnd, "open", Ruta, RepFoRepTrans, "C:\", SW_SHOWNORMAL
        Else
            GoTo Ir
        End If
    Exit Sub
Ir:
    'frmRepRegistroVenta.Show
   frmRepAnulado.Show vbModal
End Sub

Private Sub mnuCancelado_Click()
   frmCancelado.Show vbModal
End Sub

Private Sub mnuCliente_Click()
   frmRepClienteDelivery.Show vbModal
End Sub

Private Sub mnuBajaFE_Click()
    frmPapeCuadratura.TipoProceso = 2
    frmPapeCuadratura.Show vbModal
End Sub

Private Sub mnuCambiaLocal_Click()
 frmServidorEnlace.Show
End Sub

Private Sub mnuClienteDeuda_Click()
   frmRepClienteDeuda.Show vbModal
End Sub

Private Sub mnuCobranzaFecha_Click()
   frmRepCobranzaFecha.Show vbModal
End Sub


Private Sub mnuContacto_Click()
   frmRepContacto.Show vbModal
End Sub

Private Sub mnuCorrelativo_Click()
    frmRepPedido.Show vbModal
End Sub

Private Sub mnuCombinaMozo_Click()
   FrmRepMozosCombo.Show vbModal
End Sub

Private Sub mnuCobranza_Click()
   frmRepCancelacion.Show vbModal
End Sub

Private Sub mnuComanda_Click()
   frmRepComanda.Show vbModal
End Sub

Private Sub mnuCorrelativoDocumento_Click()
   frmDocumentoCorrelativo.Show
End Sub

Private Sub mnuCorrelativoPedido_Click()
   frmPedidoCorrelativo.Show
End Sub

Private Sub mnucortesia_Click()
    frmRepCortesia.Show vbModal
End Sub

Private Sub mnuCortesiaMozo_Click()
  FrmRepMozosCortesia.Show vbModal
End Sub

Private Sub mnuCtaCte_Click()
    frmCtaCte.Show
End Sub

Private Sub mnuCuentaCobrar_Click()
   frmCuentaCobrar.Show
End Sub

Private Sub mnuDescargo_Click()
    frmRepDescargoVentas.Show vbModal
End Sub

Private Sub mnuDescuento_Click()
   frmRepDescuento.Show vbModal
End Sub

Private Sub mnudiferenciaDelivery_Click()
   frmRepDiferenciaTiempoDelivery.Show vbModal
End Sub

Private Sub mnuDiferencias_Click()
   frmRepDiferencia.Show vbModal
End Sub

Private Sub mnuDocumentos_Click()
    frmRepDocumentosAnticipos.Show 1
End Sub

Private Sub mnuEnviosAutorizados_Click()
    frmRepControlEnviosAutorizados.Show vbModal
End Sub

Private Sub mnuEstPropina_Click()
   frmRepProduccionMozo.Show vbModal
End Sub

Private Sub mnuFormato_Click()
    frmPapeCuadratura.TipoProceso = 1
    frmPapeCuadratura.Show vbModal
End Sub

Private Sub mnuFrecuente_Click()
   frmRepClieFrecuentes.Show vbModal
End Sub

Private Sub mnugenerarsunat_Click()
    frmGenerarSunat.Show vbModal
End Sub

Private Sub mnuIngreso_Click()
   frmReciboIngreso.Show
End Sub

Private Sub mnuLiquidacion_Click()
   frmRepLiquidacion.Show vbModal
End Sub

Private Sub mnuLiquidacion2_Click()
    Dim Ruta As String
    Dim nombre As String
    
    On Error GoTo Ir
        Ruta = App.Path & "\Reportes Extraidos\RptLiquidacion\Reportes.exe"
        nombre = Dir(Ruta)
        If Len(nombre) > 0 Then
            ShellExecute Me.hwnd, "open", Ruta, RepExtParameters, "C:\", SW_SHOWNORMAL
        Else
            GoTo Ir
        End If
    Exit Sub
Ir:

End Sub

Private Sub mnuLiquidacionTicket_Click()
   frmRepLiquidacionTicket.Show vbModal
End Sub

Private Sub mnuMensajesUsuarios_Click()
  frmMensajeUsuario.Show vbModal
End Sub

Private Sub mnuMotorizado_Click()
  frmRepAnaliticoMotorizado.Show vbModal
End Sub

Private Sub mnuMozo_Click()
   frmRepAnaliticoMozo.Show vbModal
End Sub

Private Sub mnuNotaCredito_Click()
  frmNotaCredito.Show
End Sub

Private Sub mnuOcupabilidad_Click()
   frmRepOcupabilidad.Show vbModal
End Sub

Private Sub mnuPaloteo_Click()
  frmRepPaloteo.Show vbModal
End Sub

Private Sub mnuPaloteoCombo_Click()
   frmRepPaloteoCombo.Show vbModal
End Sub

Private Sub mnuPaloteoCortesia_Click()
   frmRepPaloteoCortesia.Show vbModal
End Sub

Private Sub mnuPaloteoVenta_Click()
   frmRepPaloteoVenta.Show vbModal
End Sub

Private Sub mnuProduccionMozo_Click()
  frmRepMozoProduccion.Show vbModal
End Sub

Private Sub mnuPedido_Click()
   frmRepPedido.Show vbModal
End Sub

Private Sub mnuPaloteoComparativo_Click()
   frmRepPaloteoComparativo.Show vbModal
End Sub

Private Sub mnuPaloteoInsumo_Click()
   frmRepInsumoVentas.Show vbModal
End Sub

Private Sub mnuPaloteoOfertas_Click()
   frmRepPaloteoOfertas.Show vbModal
End Sub

Private Sub mnuPaloteoProductoMes_Click()
    frmRepProductoMes.Show vbModal
End Sub

Private Sub mnuPaloteoTicket_Click()
   frmRepPaloteoTicket.Show vbModal
End Sub

Private Sub mnuPlanillaMotorizados_Click()
frmRepPlanillaMovilidadMotorizado.Show vbModal

End Sub

Private Sub mnuPrincipal_Click()
   frmRepPrincipal.Show vbModal
End Sub

Private Sub mnuProductosNoEnlazados_Click()
    frmRepProductosNoEnlazados.Show vbModal
End Sub

Private Sub mnuPropiedades_Click()
  frmRepPaloteoPropiedades.Show vbModal
End Sub

Private Sub mnuPropina_Click()
   frmRepPropina.Show vbModal
End Sub

Private Sub mnuRankingCombinacion_Click()
   frmRepRankingCombo.Show vbModal
End Sub

Private Sub mnuRankingCortesia_Click()
   frmRepRankingCortesia.Show vbModal
End Sub

Private Sub mnuRankingProduccion_Click()
   frmRepRankingProduccion.Show vbModal
End Sub

Private Sub mnuRankingVenta_Click()
   frmRepRankingVenta.Show vbModal
End Sub

Private Sub mnuRanking_Click()
   frmRepRanking.Show vbModal
End Sub

Private Sub mnuRecargoalconsumo_Click()
    frmRepRecargoConsumo.Show vbModal
End Sub

Private Sub mnuRecibo_Click()
   frmReciboEgreso.Show
End Sub

Private Sub mnuRegistroVenta_Click()
   'frmRepRegistroVenta.Show vbModal
    Dim Ruta As String
    Dim nombre As String
    
    Dim RepExtParameters As String
    
    RepExtParameters = "" & sRuta & "|" & sMDB & "|" & sUserName & "|" & sUserPassword & "|"
    
    On Error GoTo Ir
        Ruta = App.Path & "\Reportes Extraidos\RepFoRegVen.exe"
        nombre = Dir(Ruta)
        If Len(nombre) > 0 Then
            ShellExecute Me.hwnd, "open", Ruta, RepExtParameters, "C:\", SW_SHOWNORMAL
        Else
            GoTo Ir
        End If
    Exit Sub
Ir:
    frmRepRegistroVenta.Show
End Sub

Private Sub mnuRepCtaCte_Click()
   frmRepCtaCte.Show vbModal
End Sub

Private Sub mnuRepEntregas_Click()
   frmRepEntrega.Show vbModal
End Sub

Private Sub mnuRepEntregasReg_Click()
   frmRepEntregaRegistro.Show vbModal
End Sub

Private Sub mnuRepEquivalencias_Click()
   frmRepPaloteoSubProd.Show vbModal
End Sub

Private Sub mnuReporNotacredito_Click()
frmReportNotaCredito.Show vbModal
End Sub

Private Sub mnuReserva_Click()
   frmReserva.Show
End Sub

Private Sub mnuResultadoOperativo_Click()
   frmRepResultadoOperativo.Show vbModal
End Sub

Private Sub mnuRotacion_Click()
   frmRepRotacionMesa.Show
End Sub

Private Sub mnuSalir_Click()
   Salir
End Sub

Private Sub mnuTiempoChefControl_Click()
    frmRepTiempoChefControl.Show vbModal
End Sub

Private Sub mnuTiempoDelivery_Click()
   frmRepTiempoDelivery.Show vbModal
End Sub

Private Sub mnuTiempoKDS_Click()
   frmRepTiempoKDS.Show vbModal
End Sub

Private Sub mnuTiempoSalon_Click()
   frmRepTiempoSalon.Show vbModal
End Sub

Private Sub mnuTurno_Click()
   frmLiquidacion.Show
End Sub

Private Sub mnuVenta_Click()
   frmRepVentaAcumulada.Show vbModal
End Sub

Private Sub mnuVentaMozo_Click()
  FrmRepMozosVentas.Show vbModal
End Sub

Private Sub mnuVentaComparada_Click()
  frmRepVentaCompAnual.Show vbModal
End Sub

Private Sub mnuVentaComparadaMensual_Click()
  frmRepVentaCompMensual.Show vbModal
End Sub

Private Sub mnuVentasFechas_Click()
  frmRepVentaFecha.Show vbModal
End Sub

Private Sub mnuVentaTurno_Click()
   frmRepVentasTurno.Show vbModal
End Sub


Public Sub reinicia()
    Unload Me
    mdiConsulta.Show
End Sub

