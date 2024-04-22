VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Begin VB.Form frmRepInsumoVentas 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Paloteo de Insumos por Ventas"
   ClientHeight    =   5190
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7740
   ForeColor       =   &H8000000C&
   Icon            =   "frmRepInsumoVentas.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5190
   ScaleWidth      =   7740
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
      Left            =   4770
      Picture         =   "frmRepInsumoVentas.frx":000C
      Style           =   1  'Graphical
      TabIndex        =   23
      Top             =   4515
      Width           =   1455
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
      Height          =   4485
      Left            =   45
      TabIndex        =   26
      Top             =   0
      Width           =   7635
      Begin VB.Frame Frame4 
         Caption         =   "Detallado"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1665
         Left            =   120
         TabIndex        =   35
         Top             =   2700
         Width           =   2565
         Begin VB.OptionButton optInsumo 
            Caption         =   "Todos los Insumos"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   300
            Index           =   0
            Left            =   120
            TabIndex        =   12
            Top             =   405
            Value           =   -1  'True
            Width           =   2235
         End
         Begin VB.OptionButton optInsumo 
            Caption         =   "Insumos de Control Diario"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   525
            Index           =   1
            Left            =   120
            TabIndex        =   13
            Top             =   810
            Width           =   2415
         End
      End
      Begin VB.CheckBox chkArea 
         Caption         =   "Todas las Areas"
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
         Left            =   5040
         TabIndex        =   7
         Top             =   1485
         Value           =   1  'Checked
         Width           =   1905
      End
      Begin VB.Frame Frame3 
         Caption         =   " Origen de Datos "
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1455
         Left            =   5040
         TabIndex        =   33
         Top             =   2700
         Width           =   2490
         Begin VB.CheckBox chkPCombo 
            Caption         =   "Propiedades de los Combos"
            Height          =   195
            Left            =   135
            TabIndex        =   20
            Top             =   1125
            Value           =   1  'Checked
            Width           =   2265
         End
         Begin VB.CheckBox chkPropiedad 
            Caption         =   "Propiedades de los Platos"
            Height          =   195
            Left            =   120
            TabIndex        =   19
            Top             =   855
            Value           =   1  'Checked
            Width           =   2130
         End
         Begin VB.CheckBox chkCombo 
            Caption         =   "Combos"
            Height          =   195
            Left            =   135
            TabIndex        =   18
            Top             =   585
            Value           =   1  'Checked
            Width           =   1725
         End
         Begin VB.CheckBox chkPlato 
            Caption         =   "Platos de Venta"
            Height          =   195
            Left            =   135
            TabIndex        =   17
            Top             =   315
            Value           =   1  'Checked
            Width           =   1770
         End
      End
      Begin VB.Frame Frame2 
         Caption         =   " Tipo de Reporte "
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1665
         Left            =   2760
         TabIndex        =   32
         Top             =   2700
         Width           =   2205
         Begin VB.OptionButton optOpcion 
            Caption         =   "Consolidado"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   300
            Index           =   2
            Left            =   360
            TabIndex        =   16
            Top             =   1080
            Width           =   1500
         End
         Begin VB.OptionButton optOpcion 
            Caption         =   "Detallado"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   300
            Index           =   0
            Left            =   360
            TabIndex        =   14
            Top             =   360
            Width           =   1575
         End
         Begin VB.OptionButton optOpcion 
            Caption         =   "Resumido"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   300
            Index           =   1
            Left            =   360
            TabIndex        =   15
            Top             =   720
            Value           =   -1  'True
            Width           =   1500
         End
      End
      Begin VB.CheckBox chkGrupo 
         Caption         =   "Todas las Familias"
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
         Left            =   5040
         TabIndex        =   1
         Top             =   270
         Value           =   1  'Checked
         Width           =   2115
      End
      Begin VB.CheckBox chkSubGrupo 
         Caption         =   "Todos las Sub Familias"
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
         Left            =   5040
         TabIndex        =   3
         Top             =   675
         Value           =   1  'Checked
         Width           =   2340
      End
      Begin VB.TextBox txtInsumo 
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
         Left            =   1680
         Locked          =   -1  'True
         TabIndex        =   25
         Top             =   1080
         Width           =   2370
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
         Left            =   4140
         Picture         =   "frmRepInsumoVentas.frx":082E
         Style           =   1  'Graphical
         TabIndex        =   4
         Top             =   1035
         Width           =   765
      End
      Begin VB.CheckBox chkProducto 
         Caption         =   "Todos los Insumos"
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
         Left            =   5040
         TabIndex        =   5
         Top             =   1080
         Value           =   1  'Checked
         Width           =   2115
      End
      Begin MSComCtl2.DTPicker dtpFecFin 
         Height          =   345
         Left            =   1680
         TabIndex        =   10
         Top             =   2295
         Width           =   1905
         _ExtentX        =   3360
         _ExtentY        =   609
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
         Format          =   128450561
         CurrentDate     =   37541.9993055556
      End
      Begin MSComCtl2.DTPicker dtpFecIni 
         Height          =   345
         Left            =   1680
         TabIndex        =   8
         Top             =   1890
         Width           =   1905
         _ExtentX        =   3360
         _ExtentY        =   609
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
         Format          =   128450561
         CurrentDate     =   37539.2083333333
      End
      Begin MSComCtl2.DTPicker dtpHorIni 
         Height          =   375
         Left            =   3675
         TabIndex        =   9
         Top             =   1890
         Width           =   1230
         _ExtentX        =   2170
         _ExtentY        =   661
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
         Format          =   128450563
         UpDown          =   -1  'True
         CurrentDate     =   37539
      End
      Begin MSComCtl2.DTPicker dtpHorFin 
         Height          =   375
         Left            =   3675
         TabIndex        =   11
         Top             =   2295
         Width           =   1230
         _ExtentX        =   2170
         _ExtentY        =   661
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
         Format          =   128450563
         UpDown          =   -1  'True
         CurrentDate     =   37541.9993055556
      End
      Begin MSDataListLib.DataCombo cboGrupo 
         Height          =   315
         Left            =   1680
         TabIndex        =   0
         Top             =   270
         Width           =   3225
         _ExtentX        =   5689
         _ExtentY        =   556
         _Version        =   393216
         Enabled         =   0   'False
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
      Begin MSDataListLib.DataCombo cboSubGrupo 
         Height          =   315
         Left            =   1680
         TabIndex        =   2
         Top             =   675
         Width           =   3225
         _ExtentX        =   5689
         _ExtentY        =   556
         _Version        =   393216
         Enabled         =   0   'False
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
      Begin MSDataListLib.DataCombo cboArea 
         Height          =   315
         Left            =   1680
         TabIndex        =   6
         Top             =   1485
         Width           =   3135
         _ExtentX        =   5530
         _ExtentY        =   556
         _Version        =   393216
         Enabled         =   0   'False
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
      Begin VB.Label Label 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Area Producción :"
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
         Left            =   120
         TabIndex        =   34
         Top             =   1485
         Width           =   1545
      End
      Begin VB.Label Label 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Familia :"
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
         Index           =   4
         Left            =   120
         TabIndex        =   31
         Top             =   270
         Width           =   1545
      End
      Begin VB.Label Label 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Sub Familia :"
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
         Index           =   6
         Left            =   120
         TabIndex        =   30
         Top             =   675
         Width           =   1545
      End
      Begin VB.Label Label 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Insumos :"
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
         Left            =   120
         TabIndex        =   29
         Top             =   1080
         Width           =   1545
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Right Justify
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
         Left            =   120
         TabIndex        =   28
         Top             =   2295
         Width           =   1545
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
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
         Left            =   120
         TabIndex        =   27
         Top             =   1890
         Width           =   1545
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
      Index           =   2
      Left            =   6225
      Picture         =   "frmRepInsumoVentas.frx":0930
      Style           =   1  'Graphical
      TabIndex        =   24
      Top             =   4515
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
      Left            =   3315
      Picture         =   "frmRepInsumoVentas.frx":0A22
      Style           =   1  'Graphical
      TabIndex        =   22
      Top             =   4515
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
      Left            =   1845
      Picture         =   "frmRepInsumoVentas.frx":0F54
      Style           =   1  'Graphical
      TabIndex        =   21
      Top             =   4515
      Width           =   1455
   End
   Begin MSComDlg.CommonDialog cmdSave 
      Left            =   638
      Top             =   4575
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
End
Attribute VB_Name = "frmRepInsumoVentas"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim sNombre     As String
Dim rsReporte   As Recordset
Dim RsArea      As Recordset
Dim RsProducto  As Recordset

Dim Detallado   As New dsrInsumosD
Dim Resumido    As New dsrInsumosR
Dim ControlDiario    As New dsrInsumosCD

Dim sCriterio   As String
Dim sInsumo     As String
Dim sFiltro     As String
Dim sTitulo     As String
Dim sPrecio     As String
Dim sTexto      As String

Dim fInicio As Date
Dim fFinal As Date
Dim familia As String
Dim subFamilia As String
Dim Area As String

Dim RsGrupo As Recordset
Dim RsSubGrupo As Recordset

Sub LlenaCombos()
    With cboGrupo
         Isql = "Select * from vFamilia order by Descripcion"
         Set RsGrupo = Lib.OpenRecordset(Isql, CnAlmacen)
         Set .RowSource = RsGrupo
             .DataField = "Descripcion"
             .ListField = "Descripcion"
             .BoundColumn = "Codigo"
    End With
           
    With cboSubGrupo
         Isql = "Select * from vSubFamilia order by Descripcion"
         Set RsSubGrupo = Lib.OpenRecordset(Isql, CnAlmacen)
         Set .RowSource = RsSubGrupo
             .DataField = "Descripcion"
             .ListField = "Descripcion"
             .BoundColumn = "Codigo"
    End With
    
    With cboArea
         Isql = "Select * from vArea"
         Set RsArea = Lib.OpenRecordset(Isql, CnAlmacen)
         Set .RowSource = RsArea
             .DataField = "Descripcion"
             .ListField = "Descripcion"
             .BoundColumn = "Codigo"
    End With
End Sub

Private Sub cboGrupo_Click(Area As Integer)
    cboSubGrupo.BoundText = ""
    With cboSubGrupo
         Isql = "Select * from vSubFamilia " & IIf(cboGrupo.BoundText = "", "", "where substring(Codigo,1,2) = '" & cboGrupo.BoundText & "'") & " order by Descripcion "
         Set RsSubGrupo = Lib.OpenRecordset(Isql, CnAlmacen)
         Set .RowSource = RsSubGrupo
    End With
End Sub

Private Sub chkArea_Click()
    If chkArea.value = 1 Then
      cboArea.Enabled = False
      cboArea.Text = ""
   Else
      cboArea.Enabled = True
   End If
End Sub

Private Sub ChkGrupo_Click()
   If chkGrupo.value = 1 Then
      cboGrupo.Enabled = False
      cboGrupo.Text = ""
   Else
      cboGrupo.Enabled = True
   End If
End Sub

Private Sub chkSubGrupo_Click()
   If chkSubGrupo.value = 1 Then
      cboSubGrupo.Enabled = False
      cboSubGrupo.Text = ""
   Else
      cboSubGrupo.Enabled = True
   End If
End Sub
Private Sub chkproducto_Click()
   If chkProducto.value = 1 Then
      sInsumo = ""
      txtInsumo.Text = ""
      cmdBusca(0).Enabled = False
   Else
      cmdBusca(0).Enabled = True
   End If
End Sub

Private Sub cmdOpcion_Click(Index As Integer)
   If Index = 2 Then
      Unload Me
      Exit Sub
   End If
   Set Detallado = New dsrInsumosD
   Set Resumido = New dsrInsumosR
   Set ControlDiario = New dsrInsumosCD

   familia = ""
   subFamilia = ""

   Area = ""
   sCriterio = ""
   sTitulo = ""
   sFiltro = ""
   
   If chkGrupo.value = 0 Then
      If cboGrupo.Text = "" Then
         MsgBox "Debe escoger la Familia", vbCritical, sMensaje
         Exit Sub
      End If
        familia = cboGrupo.Text
   End If
    
   If chkSubGrupo.value = 0 Then
      If cboSubGrupo.Text = "" Then
         MsgBox "Debe escoger la sub familia", vbCritical, sMensaje
         Exit Sub
      End If
        subFamilia = cboSubGrupo.Text
   End If
   If chkProducto.value = 0 Then
      If sInsumo = "" Then
         MsgBox "Debe escoger el producto", vbCritical, sMensaje
         Exit Sub
      End If
   End If
   
   If chkArea.value = 0 Then
    If Me.cboArea.Text = "" Then
        MsgBox "Debe escoger el area", vbCritical, sMensaje
        Exit Sub
    End If
    'CMiranda------------------------------------------------------------------------------
    Area = Me.cboArea.BoundText
    'Fin CMiranda--------------------------------------------------------------------------
   End If
   Select Case Index
          Case Is = 0 ' Preview
                If optOpcion(2).value = True Then
                    Genera2
                Else
                    Genera
                End If
               If rsReporte.EOF = True Then
                  MsgBox "No hay Datos para Mostrar", vbExclamation, "Mensaje del Sistema"
                  Screen.MousePointer = vbDefault
                  Exit Sub
               End If

               If optOpcion(0).value = True Then
                  frmEmite.CRViewer.DisplayGroupTree = True
                  Detallado.SelectPrinter Printer.DriverName, Printer.DeviceName, Printer.Port
                  Detallado.PaperOrientation = crPortrait
               Else
                    If optOpcion(2).value = True Then
                        frmEmite.CRViewer.DisplayGroupTree = True
                        ControlDiario.SelectPrinter Printer.DriverName, Printer.DeviceName, Printer.Port
                        ControlDiario.PaperOrientation = crPortrait
                    Else
                        frmEmite.CRViewer.DisplayGroupTree = True
                        Resumido.SelectPrinter Printer.DriverName, Printer.DeviceName, Printer.Port
                        Resumido.PaperOrientation = crPortrait
                    End If
               End If

               
               frmEmite.CRViewer.ViewReport
               frmEmite.Show vbModal
          
          Case Is = 1 ' Imprimir
                If optOpcion(2).value = True Then
                    Genera2
                Else
                    Genera
                End If
               Screen.MousePointer = vbDefault
               If rsReporte.EOF = True Then
                  MsgBox "No hay Datos para Mostrar", vbExclamation, "Mensaje del Sistema"
                  Exit Sub
               End If
               If optOpcion(0).value = True Then
                  Detallado.SelectPrinter Printer.DriverName, Printer.DeviceName, Printer.Port
                  Detallado.PaperOrientation = crPortrait
                  Detallado.PrintOut
               Else
                    If optOpcion(2).value = True Then
                        ControlDiario.SelectPrinter Printer.DriverName, Printer.DeviceName, Printer.Port
                        ControlDiario.PaperOrientation = crPortrait
                        ControlDiario.PrintOut
                    Else
                        Resumido.SelectPrinter Printer.DriverName, Printer.DeviceName, Printer.Port
                        Resumido.PaperOrientation = crPortrait
                        Resumido.PrintOut
                    End If
               End If
          
          Case Is = 2 ' Salir
               Unload Me
          
          Case Is = 3 ' Exportar
                If optOpcion(2).value = True Then
                    Genera2
                Else
                    Genera
                End If
               Screen.MousePointer = vbDefault
               If rsReporte.EOF = True Then
                  MsgBox "No hay Datos para Mostrar", vbExclamation, "Mensaje del Sistema"
                  Exit Sub
               End If
               If optOpcion(0).value = True Then
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
                    If optOpcion(2).value = True Then
                        ControlDiario.ExportOptions.FormatType = 21
                        ControlDiario.ExportOptions.DestinationType = 1
                        cmdSave.Filter = "Libro de Microsoft Excel|*.xls"
                        cmdSave.ShowSave
                        If cmdSave.FileName = "" Then
                           Exit Sub
                        End If
                        ControlDiario.ExportOptions.DiskFileName = cmdSave.FileName
                        ControlDiario.Export False
                    Else
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

Private Sub cmdBusca_Click(Index As Integer)
   Dim xCriterio As String
   Select Case Index
          Case Is = 0
               xCriterio = "lActivo = 1 "
               Isql = "select tCodigoProducto as Codigo, tDetallado as Descripcion from vProducto where " & xCriterio & " order by Descripcion"
               frmBuscaAlmacen.nPredeterm = 1
               Call ConfGrilla(2, frmBuscaAlmacen.grdGrilla, "Codigo", 2, "Codigo", 2300, 2, 0, "", _
                                                      "Insumo", 2, "Descripcion", 5000, 0, 0, "")
               frmBuscaAlmacen.Show vbModal
               If Not wEnter Then
                  Exit Sub
               End If
               sInsumo = sCodigo
               txtInsumo.Text = sDescrip
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
    dtpFecFin.value = Date
    cmdBusca(0).Enabled = False
End Sub

Private Sub Form_Unload(Cancel As Integer)
   Set rsReporte = Nothing
End Sub

Public Sub Genera()
   Screen.MousePointer = vbHourglass
   Dim oComando As clsComando
    Set oComando = New clsComando
    If Not oComando.CreateCmdSp("spRep_PaloteoInsumo", Cn) Then
        Set oComando = Nothing
        Exit Sub
    End If
        fInicio = Format(dtpFecIni.value, "yyyy/MM/dd") & " " & Format(dtpHorIni.value, "HH:mm")
        fFinal = Format(dtpFecFin.value, "yyyy/MM/dd") & " " & Format(dtpHorFin.value, "HH:mm")
        oComando.CreateParameter "@dbAlmacen", adVarChar, adParamInput, 35, sAlmacenMDB
        oComando.CreateParameter "@Slocal", adVarChar, adParamInput, 50, sLocal
        oComando.CreateParameter "@Familia", adVarChar, adParamInput, 100, cboGrupo.Text
        oComando.CreateParameter "@SubFamilia", adVarChar, adParamInput, 100, cboSubGrupo.Text
        oComando.CreateParameter "@tCodigoProducto", adVarChar, adParamInput, 15, sInsumo
        'CMiranda------------------------------------------------------------------------
        oComando.CreateParameter "@Area", adVarChar, adParamInput, 100, cboArea.BoundText
        'Fin CMiranda--------------------------------------------------------------------
        oComando.CreateParameter "@flagPlato", adBoolean, adParamInput, 1, Me.chkPlato.value
        oComando.CreateParameter "@flagCombo", adBoolean, adParamInput, 1, Me.chkCombo.value
        oComando.CreateParameter "@flagPropiedad", adBoolean, adParamInput, 1, Me.chkPropiedad.value
        oComando.CreateParameter "@flagPCombo", adBoolean, adParamInput, 1, Me.chkPCombo.value
        oComando.CreateParameter "@flagTipo", adBoolean, adParamInput, 1, optOpcion(2).value
        oComando.CreateParameter "@flagTipoR", adBoolean, adParamInput, 1, optOpcion(1).value
        oComando.CreateParameter "@fInicio", adDBDate, adParamInput, 10, fInicio
        oComando.CreateParameter "@fFinal", adDBDate, adParamInput, 10, fFinal
        oComando.CreateParameter "@Insumo", adVarChar, adParamInput, 1, IIf(Me.optInsumo(0).value, "T", "D")
   If Not oComando.GetParamOK Then
      Set oComando = Nothing
      Exit Sub
   End If
   Set rsReporte = oComando.GetSP()
   If optOpcion(0).value = True Then
      Detallado.DiscardSavedData
      Detallado.Text25.SetText localConectado
      Detallado.Database.SetDataSource rsReporte
      Detallado.Text10.SetText "Insumos Por Venta Detallado"
      Detallado.Text14.SetText sRazonSocial
      If optInsumo(0).value Then
        Detallado.ReportTitle = "Todos los Insumos del " & Format(dtpFecIni.value, "dd/mmm/yyyy") & " " & Format(dtpHorIni.value, "HH:mm") & " Al " & Format(dtpFecFin.value, "dd/mmm/yyyy") & " " & Format(dtpHorFin.value, "HH:mm")
      Else
        Detallado.ReportTitle = "Insumos de Control Diario del " & Format(dtpFecIni.value, "dd/mmm/yyyy") & " " & Format(dtpHorIni.value, "HH:mm") & " Al " & Format(dtpFecFin.value, "dd/mmm/yyyy") & " " & Format(dtpHorFin.value, "HH:mm")
      End If
      frmEmite.CRViewer.ReportSource = Detallado
   Else
      Resumido.DiscardSavedData
      Resumido.Database.SetDataSource rsReporte
      Resumido.Text10.SetText "Insumos Por Venta Resumido"
      Resumido.Text14.SetText sRazonSocial
      Resumido.Text25.SetText localConectado
      If optInsumo(0).value Then
        Resumido.ReportTitle = "Todos los Insumos del " & Format(dtpFecIni.value, "dd/mmm/yyyy") & " " & Format(dtpHorIni.value, "HH:mm") & " Al " & Format(dtpFecFin.value, "dd/mmm/yyyy") & " " & Format(dtpHorFin.value, "HH:mm")
      Else
        Resumido.ReportTitle = "Insumos de Control Diario del " & Format(dtpFecIni.value, "dd/mmm/yyyy") & " " & Format(dtpHorIni.value, "HH:mm") & " Al " & Format(dtpFecFin.value, "dd/mmm/yyyy") & " " & Format(dtpHorFin.value, "HH:mm")
      End If
      frmEmite.CRViewer.ReportSource = Resumido
   End If
End Sub

Public Sub Genera2()
   Screen.MousePointer = vbHourglass
   Dim oComando As clsComando
    Set oComando = New clsComando
    
    
    If Not oComando.CreateCmdSp("spRep_PaloteoInsumo", Cn) Then
        Set oComando = Nothing
        Exit Sub
    End If
    
        fInicio = Format(dtpFecIni.value, "yyyy/MM/dd") & " " & Format(dtpHorIni.value, "HH:mm")
        fFinal = Format(dtpFecFin.value, "yyyy/MM/dd") & " " & Format(dtpHorFin.value, "HH:mm")
        oComando.CreateParameter "@dbAlmacen", adVarChar, adParamInput, 35, sAlmacenMDB
        oComando.CreateParameter "@Slocal", adVarChar, adParamInput, 50, sLocal
        oComando.CreateParameter "@Familia", adVarChar, adParamInput, 100, familia
        oComando.CreateParameter "@SubFamilia", adVarChar, adParamInput, 100, subFamilia
        oComando.CreateParameter "@tCodigoProducto", adVarChar, adParamInput, 15, sInsumo
        'CMiranda----------------------------------------------------------------------------
        oComando.CreateParameter "@Area", adVarChar, adParamInput, 100, Area
        'Fin CMiranda------------------------------------------------------------------------
        oComando.CreateParameter "@flagPlato", adBoolean, adParamInput, 1, Me.chkPlato.value
        oComando.CreateParameter "@flagCombo", adBoolean, adParamInput, 1, Me.chkCombo.value
        oComando.CreateParameter "@flagPropiedad", adBoolean, adParamInput, 1, Me.chkPropiedad.value
        oComando.CreateParameter "@flagPCombo", adBoolean, adParamInput, 1, Me.chkPCombo.value
        oComando.CreateParameter "@flagTipo", adBoolean, adParamInput, 1, optOpcion(2).value
        oComando.CreateParameter "@flagTipoR", adBoolean, adParamInput, 1, optOpcion(1).value
        oComando.CreateParameter "@fInicio", adDBDate, adParamInput, 10, fInicio
        oComando.CreateParameter "@fFinal", adDBDate, adParamInput, 10, fFinal
        oComando.CreateParameter "@Insumo", adVarChar, adParamInput, 1, IIf(Me.optInsumo(0).value, "T", "D")
        
   If Not oComando.GetParamOK Then
      Set oComando = Nothing
      Exit Sub
   End If
   Set rsReporte = oComando.GetSP()
   ControlDiario.DiscardSavedData
   ControlDiario.Database.SetDataSource rsReporte
   ControlDiario.Text10.SetText "Insumos Por Venta Consolidado"
   ControlDiario.Text14.SetText sRazonSocial
   ControlDiario.Text25.SetText localConectado
   If optInsumo(0).value = True Then
    ControlDiario.ReportTitle = "Todos los Insumos del " & Format(dtpFecIni.value, "dd/mmm/yyyy") & " " & Format(dtpHorIni.value, "HH:mm") & " Al " & Format(dtpFecFin.value, "dd/mmm/yyyy") & " " & Format(dtpHorFin.value, "HH:mm")
   Else
    ControlDiario.ReportTitle = "Insumos de Control Diario del " & Format(dtpFecIni.value, "dd/mmm/yyyy") & " " & Format(dtpHorIni.value, "HH:mm") & " Al " & Format(dtpFecFin.value, "dd/mmm/yyyy") & " " & Format(dtpHorFin.value, "HH:mm")
   End If
   frmEmite.CRViewer.ReportSource = ControlDiario
End Sub


