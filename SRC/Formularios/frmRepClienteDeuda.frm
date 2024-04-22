VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Begin VB.Form frmRepClienteDeuda 
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Deuda de Clientes"
   ClientHeight    =   3420
   ClientLeft      =   5100
   ClientTop       =   6285
   ClientWidth     =   8265
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   9.75
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmRepClienteDeuda.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3420
   ScaleWidth      =   8265
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton cmdOpcion 
      Caption         =   "Exportar"
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
      Left            =   4132
      Picture         =   "frmRepClienteDeuda.frx":000C
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   2760
      Width           =   1455
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
      Index           =   2
      Left            =   5587
      Picture         =   "frmRepClienteDeuda.frx":082E
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   2760
      Width           =   1455
   End
   Begin VB.CommandButton cmdOpcion 
      Caption         =   "Impresora"
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
      Left            =   2677
      Picture         =   "frmRepClienteDeuda.frx":0920
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   2760
      Width           =   1455
   End
   Begin VB.CommandButton cmdOpcion 
      Caption         =   "Emite"
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
      Left            =   1222
      Picture         =   "frmRepClienteDeuda.frx":0E52
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   2760
      Width           =   1455
   End
   Begin MSComDlg.CommonDialog cmdSave 
      Left            =   6960
      Top             =   240
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
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
      Height          =   2685
      Left            =   0
      TabIndex        =   17
      Top             =   0
      Width           =   8205
      Begin VB.OptionButton optOpcion 
         Caption         =   "Listado Histórico"
         Height          =   240
         Index           =   2
         Left            =   5475
         TabIndex        =   24
         Top             =   2400
         Width           =   2295
      End
      Begin VB.CheckBox chkEstado 
         Caption         =   "Todos los Estados"
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
         Left            =   5730
         TabIndex        =   14
         Top             =   2040
         Value           =   1  'Checked
         Width           =   1965
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
         Left            =   4845
         Picture         =   "frmRepClienteDeuda.frx":1384
         Style           =   1  'Graphical
         TabIndex        =   9
         Top             =   1170
         Width           =   765
      End
      Begin VB.TextBox txtCliente 
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
         Left            =   2010
         Locked          =   -1  'True
         TabIndex        =   8
         Top             =   1200
         Width           =   2775
      End
      Begin VB.OptionButton optOpcion 
         Caption         =   "Detallado"
         Height          =   240
         Index           =   0
         Left            =   2010
         TabIndex        =   15
         Top             =   2400
         Value           =   -1  'True
         Width           =   1545
      End
      Begin VB.OptionButton optOpcion 
         Caption         =   "Resumido"
         Height          =   240
         Index           =   1
         Left            =   3900
         TabIndex        =   16
         Top             =   2400
         Width           =   1545
      End
      Begin VB.CheckBox chkTipoDocumento 
         Caption         =   "Todos los Documentos"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   5730
         TabIndex        =   12
         Top             =   1665
         Value           =   1  'Checked
         Width           =   2325
      End
      Begin VB.CheckBox chkCliente 
         Caption         =   "Todos los Clientes"
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
         Left            =   5730
         TabIndex        =   10
         Top             =   1215
         Value           =   1  'Checked
         Width           =   1905
      End
      Begin MSComCtl2.DTPicker dtpFecFin 
         Height          =   330
         Left            =   2010
         TabIndex        =   6
         Top             =   765
         Width           =   1620
         _ExtentX        =   2858
         _ExtentY        =   582
         _Version        =   393216
         Format          =   80019457
         CurrentDate     =   37539
      End
      Begin MSComCtl2.DTPicker dtpFecIni 
         Height          =   330
         Left            =   2010
         TabIndex        =   4
         Top             =   330
         Width           =   1620
         _ExtentX        =   2858
         _ExtentY        =   582
         _Version        =   393216
         Format          =   80019457
         CurrentDate     =   37539
      End
      Begin MSDataListLib.DataCombo cboTipoDocumento 
         Height          =   315
         Left            =   2010
         TabIndex        =   11
         Top             =   1620
         Width           =   3615
         _ExtentX        =   6376
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
      Begin MSDataListLib.DataCombo cboEstado 
         Height          =   315
         Left            =   2010
         TabIndex        =   13
         Top             =   2040
         Width           =   3615
         _ExtentX        =   6376
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
      Begin MSComCtl2.DTPicker dtpHoraIni 
         Height          =   330
         Left            =   3660
         TabIndex        =   5
         Top             =   330
         Width           =   1110
         _ExtentX        =   1958
         _ExtentY        =   582
         _Version        =   393216
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
         Format          =   80019459
         UpDown          =   -1  'True
         CurrentDate     =   37539
      End
      Begin MSComCtl2.DTPicker dtpHoraFin 
         Height          =   330
         Left            =   3660
         TabIndex        =   7
         Top             =   765
         Width           =   1110
         _ExtentX        =   1958
         _ExtentY        =   582
         _Version        =   393216
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
         Format          =   80019459
         UpDown          =   -1  'True
         CurrentDate     =   37541.9993055556
      End
      Begin VB.Label Label 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Estado Documento :"
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
         Left            =   180
         TabIndex        =   23
         Top             =   2100
         Width           =   1740
      End
      Begin VB.Label Label 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Tipo Documento :"
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
         Index           =   5
         Left            =   390
         TabIndex        =   22
         Top             =   1680
         Width           =   1530
      End
      Begin VB.Label Label 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Cliente :"
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
         Index           =   10
         Left            =   1200
         TabIndex        =   21
         Top             =   1260
         Width           =   720
      End
      Begin VB.Label Label 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Tipo Reporte :"
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
         Left            =   675
         TabIndex        =   20
         Top             =   2430
         Width           =   1245
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Fecha Final :"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   210
         Left            =   930
         TabIndex        =   19
         Top             =   825
         Width           =   990
      End
      Begin VB.Label Label21 
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
         Left            =   690
         TabIndex        =   18
         Top             =   405
         Width           =   1230
      End
   End
End
Attribute VB_Name = "frmRepClienteDeuda"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim Detallado As New dsrClienteDeudaD
Dim Resumido  As New dsrClienteDeudaR
Dim Listado As New dsrClienteDeudaHistorico
Dim RsCliente       As Recordset
Dim RsTipoDocumento As Recordset
Dim RsEstado        As Recordset
Dim RsPrinter       As Recordset
Dim sReporte        As String
Dim sCriterio       As String
Dim fInicio         As Date
Dim fFinal          As Date

Dim clienteSel    As String
Dim estadoSel     As String
Dim tipoSel       As String


Private Sub cmdBusca_Click()
   Dim xCriterio As String
   Isql = "Select Codigo as Codigo, Identidad, Descripcion as Descripcion from vCOMPANIA order by descripcion"
   frmBusca.cboCriterio.Enabled = True
   frmBusca.nPredeterm = 2
   Call ConfGrilla(3, frmBusca.grdGrilla, "Codigo", 2, "Codigo", 1200, 0, 0, "", _
                                          "Identificador", 2, "Identidad", 1500, 0, 0, "", _
                                          "Razón Comercial", 2, "Descripcion", 4500, 0, 0, "")
   frmBusca.Show vbModal
   If Not wEnter Then
      Exit Sub
   End If
   sCliente = sCodigo
   txtCliente.Text = sDescrip
End Sub

Private Sub chkCliente_Click()
   If chkCliente.value = 1 Then
      sCliente = ""
      txtCliente.Text = ""
      cmdBusca.Enabled = False
   Else
      cmdBusca.Enabled = True
   End If
End Sub

Private Sub chkEstado_Click()
   If chkEstado.value = 1 Then
      cboEstado.Enabled = False
      cboEstado.Text = ""
   Else
      cboEstado.Enabled = True
   End If
End Sub

Private Sub chkTipoDocumento_Click()
   If chkTipoDocumento.value = 1 Then
      cboTipoDocumento.Enabled = False
      cboTipoDocumento.Text = ""
   Else
      cboTipoDocumento.Enabled = True
   End If
End Sub

Private Sub cmdOpcion_Click(Index As Integer)
   If Index = 2 Then
      Unload Me
      Exit Sub
   End If
   
   sCriterio = ""
   clienteSel = ""
   estadoSel = ""
   tipoSel = ""
   'Cliente
   If chkCliente.value = 0 Then
      If sCliente = "" Then
         MsgBox "Debe escoger el Cliente", vbCritical, sMensaje
         Exit Sub
      End If
      clienteSel = sCliente
   End If
   'estado
    If chkEstado.value = 0 Then
      If cboEstado.Text = "" Then
         MsgBox "Debe escoger un Estado del Documento", vbCritical, sMensaje
         Exit Sub
      End If
      estadoSel = cboEstado.BoundText
   End If
   'Documento
   If chkTipoDocumento.value = 0 Then
      If cboTipoDocumento.Text = "" Then
         MsgBox "Debe escoger un Tipo de Documento", vbCritical, sMensaje
         Exit Sub
      End If
      tipoSel = cboTipoDocumento.BoundText
   End If
        
  
    fInicio = Format(dtpFecIni.value, "yyyy/mm/dd") & " " & Format(dtpHoraIni.value, "HH:mm")
    fFinal = Format(dtpFecFin.value, "yyyy/mm/dd") & " " & Format(dtpHoraFin.value, "HH:mm")

    
   If fInicio > fFinal Then
      MsgBox "Error en Rango de Fechas", vbCritical, sMensaje
      dtpFecFin.SetFocus
      Exit Sub
   End If
        

   
   Select Case Index
          Case Is = 0 ' Preview
              Genera
              If RsPrinter.EOF = True Then
                  Screen.MousePointer = vbDefault
                  MsgBox "No hay Datos para Mostrar", vbExclamation, "Mensaje del Sistema"
                  Exit Sub
              End If
              If optOpcion(2).value = True Then
                  Listado.SelectPrinter Printer.DriverName, Printer.DeviceName, Printer.Port
                  Listado.PaperOrientation = crPortrait
              Else
                        If optOpcion(0).value = True Then
                           Detallado.SelectPrinter Printer.DriverName, Printer.DeviceName, Printer.Port
                           Detallado.PaperOrientation = crPortrait
                        Else
                           Resumido.SelectPrinter Printer.DriverName, Printer.DeviceName, Printer.Port
                           Resumido.PaperOrientation = crPortrait
                        End If
              End If
                 frmEmite.CRViewer.ViewReport
                 frmEmite.Show vbModal
          
          Case Is = 1 ' Imprimir
               Genera
               Screen.MousePointer = vbDefault
               If RsPrinter.EOF = True Then
                  MsgBox "No hay Datos para Mostrar", vbExclamation, "Mensaje del Sistema"
                  Exit Sub
               End If
               If optOpcion(2).value = True Then
                  Listado.SelectPrinter Printer.DriverName, Printer.DeviceName, Printer.Port
                  Listado.PaperOrientation = crPortrait
                  Listado.PrintOut
               
               Else
               
               If optOpcion(0).value = True Then
                  Detallado.SelectPrinter Printer.DriverName, Printer.DeviceName, Printer.Port
                  Detallado.PaperOrientation = crPortrait
                  Detallado.PrintOut
               Else
                  Resumido.SelectPrinter Printer.DriverName, Printer.DeviceName, Printer.Port
                  Resumido.PaperOrientation = crPortrait
                  Resumido.PrintOut
               End If
               End If
          Case Is = 2 ' Salir
               Unload Me
               
          Case Is = 3 ' Exportar
               Genera
               Screen.MousePointer = vbDefault
               If RsPrinter.EOF = True Then
                  MsgBox "No hay Datos para Mostrar", vbExclamation, "Mensaje del Sistema"
                  Exit Sub
               End If
               If optOpcion(2).value = True Then
                  Listado.SelectPrinter Printer.DriverName, Printer.DeviceName, Printer.Port
                  Listado.PaperOrientation = crPortrait
                  Listado.ExportOptions.FormatType = 21
                  Listado.ExportOptions.DestinationType = 1
                  
                  cmdSave.Filter = "Libro de Microsoft Excel|*.xls"
                  cmdSave.ShowSave
                  If cmdSave.FileName = "" Then
                     Exit Sub
                  End If
                  Listado.ExportOptions.DiskFileName = cmdSave.FileName
                  Listado.Export False
               Else
               If optOpcion(0).value = True Then
                  Detallado.SelectPrinter Printer.DriverName, Printer.DeviceName, Printer.Port
                  Detallado.PaperOrientation = crPortrait
                  Detallado.ExportOptions.FormatType = 21
                  Detallado.ExportOptions.DestinationType = 1
                  
                  cmdSave.Filter = "Libro de Microsoft Excel|*.xls"
                  cmdSave.ShowSave
                  If cmdSave.FileName = "" Then
                     Exit Sub
                  End If
                  Detallado.ExportOptions.DiskFileName = cmdSave.FileName
                  Detallado.Export False
                Else
                  Resumido.SelectPrinter Printer.DriverName, Printer.DeviceName, Printer.Port
                  Resumido.PaperOrientation = crPortrait
                  Resumido.ExportOptions.FormatType = 21
                  Resumido.ExportOptions.DestinationType = 1
                  cmdSave.Filter = "Libro de Microsoft Excel|*.xls"
                  cmdSave.ShowSave
                  If cmdSave.FileName = "" Then
                     Exit Sub
                  End If
                  Resumido.ExportOptions.DiskFileName = cmdSave.FileName
                  Resumido.Export False
                End If
               End If
   End Select
End Sub

Private Sub dtpFecfin_LostFocus()
   If dtpFecIni.value > dtpFecFin.value Then
      MsgBox "Error en Rango de Fechas", vbCritical, sMensaje
      dtpFecFin.SetFocus
   End If
End Sub

Private Sub Form_Load()
    Centrar Me
    LlenaCombos
    dtpFecIni.value = Date
    dtpFecFin.value = Date + 1

    cmdBusca.Enabled = False
    cboTipoDocumento.Enabled = False
    cboEstado.Enabled = False
End Sub

Sub LlenaCombos()
   With cboTipoDocumento
      'Compania
      Isql = "Select * from vTipoDocumento where Codigo<>'00' Order By Descripcion"
      Set RsTipoDocumento = Lib.OpenRecordset(Isql, Cn)
      Set .RowSource = RsTipoDocumento
      .DataField = "Descripcion"
      .ListField = "Descripcion"
      .BoundColumn = "Codigo"
   End With
   
   With cboEstado
      'Estado
      Isql = "Select * from vEstadoDocumento where lActivo=1"
      Set RsEstado = Lib.OpenRecordset(Isql, Cn)
      Set .RowSource = RsEstado
      .DataField = "Descripcion"
      .ListField = "Descripcion"
      .BoundColumn = "Codigo"
   End With
   
End Sub

Public Sub Genera()
   Dim oComando As clsComando
   Screen.MousePointer = vbHourglass
   Set oComando = New clsComando
    If Not oComando.CreateCmdSp("spRep_CuentasCobrar", Cn) Then
       Set oComando = Nothing
       Exit Sub
    End If
    
    oComando.CreateParameter "@tipoListado", adVarChar, adParamInput, 1, IIf(Me.optOpcion(2).value = True, "1", "0")
    oComando.CreateParameter "@tipo", adVarChar, adParamInput, 10, IIf(Me.optOpcion(0).value = True, "DETALLADO", "RESUMIDO")
    oComando.CreateParameter "@cliente", adVarChar, adParamInput, 10, clienteSel
    oComando.CreateParameter "@tipoDoc", adVarChar, adParamInput, 5, tipoSel
    oComando.CreateParameter "@estadodoc", adVarChar, adParamInput, 5, estadoSel
    oComando.CreateParameter "@fInicio", adDBDate, adParamInput, 10, fInicio
    oComando.CreateParameter "@fFinal", adDBDate, adParamInput, 10, fFinal
    If Not oComando.GetParamOK Then
       Set oComando = Nothing
       Exit Sub
    End If
    Set RsPrinter = oComando.GetSP()
   If optOpcion(2).value = True Then
      Listado.DiscardSavedData
      Listado.Database.SetDataSource RsPrinter
      Listado.ReportTitle = "Del " & dtpFecIni.value & " Al " & dtpFecFin.value
      Listado.Text28.SetText sRazonSocial
      Listado.Text3.SetText localConectado
      frmEmite.CRViewer.DisplayGroupTree = False
      frmEmite.CRViewer.ReportSource = Listado
   Else
   If optOpcion(0).value = True Then
      Detallado.DiscardSavedData
      Detallado.Database.SetDataSource RsPrinter
      Detallado.ReportTitle = "Del " & dtpFecIni.value & " Al " & dtpFecFin.value
      Detallado.Text28.SetText sRazonSocial
      Detallado.Text3.SetText localConectado
      frmEmite.CRViewer.DisplayGroupTree = False
      frmEmite.CRViewer.ReportSource = Detallado
   Else
      Resumido.DiscardSavedData
      Resumido.Database.SetDataSource RsPrinter
      Resumido.ReportTitle = "Del " & dtpFecIni.value & " Al " & dtpFecFin.value
      Resumido.Text2.SetText localConectado
      Resumido.Text28.SetText sRazonSocial
      frmEmite.CRViewer.DisplayGroupTree = False
      frmEmite.CRViewer.ReportSource = Resumido
   End If
   End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set RsPrinter = Nothing
    Set RsTipoDocumento = Nothing
    Set RsCliente = Nothing
    Set frmRepClienteDeuda = Nothing
    
    If sReporte <> "" Then
      Cn.Execute "Drop Table " & sReporte
   End If
End Sub

Private Sub Label2_Click()

End Sub

