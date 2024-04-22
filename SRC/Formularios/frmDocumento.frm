VERSION 5.00
Object = "{562E3E04-2C31-4ECE-83F4-4017EEE51D40}#8.0#0"; "todg8.ocx"
Begin VB.Form frmDocumento 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Generaci�n de Documentos"
   ClientHeight    =   8280
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   11880
   Icon            =   "frmDocumento.frx":0000
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8280
   ScaleWidth      =   11880
   WindowState     =   2  'Maximized
   Begin VB.CommandButton cmdNotasCredito 
      Caption         =   "Notas de Credito"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   585
      Index           =   13
      Left            =   3810
      TabIndex        =   42
      Top             =   7695
      Width           =   1260
   End
   Begin TrueOleDBGrid80.TDBGrid grdDocumento 
      Height          =   3150
      Left            =   30
      TabIndex        =   27
      Top             =   4470
      Width           =   9680
      _ExtentX        =   17066
      _ExtentY        =   5556
      _LayoutType     =   4
      _RowHeight      =   23
      _WasPersistedAsPixels=   0
      Columns(0)._VlistStyle=   0
      Columns(0)._MaxComboItems=   5
      Columns(0).DataField=   ""
      Columns(0).NumberFormat=   "True/False"
      Columns(0)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns.Count   =   1
      Splits(0)._UserFlags=   0
      Splits(0).MarqueeStyle=   3
      Splits(0).RecordSelectorWidth=   503
      Splits(0)._SavedRecordSelectors=   -1  'True
      Splits(0).ScrollBars=   2
      Splits(0).AllowColSelect=   0   'False
      Splits(0).FetchRowStyle=   -1  'True
      Splits(0).DividerStyle=   2
      Splits(0).DividerColor=   32768
      Splits(0).SpringMode=   0   'False
      Splits(0)._PropDict=   "_ColumnProps,515,0;_UserFlags,518,3"
      Splits(0)._ColumnProps(0)=   "Columns.Count=1"
      Splits(0)._ColumnProps(1)=   "Column(0).Width=2725"
      Splits(0)._ColumnProps(2)=   "Column(0).DividerColor=0"
      Splits(0)._ColumnProps(3)=   "Column(0)._WidthInPix=2646"
      Splits(0)._ColumnProps(4)=   "Column(0)._EditAlways=0"
      Splits(0)._ColumnProps(5)=   "Column(0).AllowSizing=0"
      Splits(0)._ColumnProps(6)=   "Column(0)._ColStyle=20"
      Splits(0)._ColumnProps(7)=   "Column(0).Order=1"
      Splits.Count    =   1
      PrintInfos(0)._StateFlags=   0
      PrintInfos(0).Name=   "piInternal 0"
      PrintInfos(0).PageHeaderFont=   "Size=6.75,Charset=0,Weight=400,Underline=0,Italic=0,Strikethrough=0,Name=Small Fonts"
      PrintInfos(0).PageFooterFont=   "Size=6.75,Charset=0,Weight=400,Underline=0,Italic=0,Strikethrough=0,Name=Small Fonts"
      PrintInfos(0).PageHeaderHeight=   0
      PrintInfos(0).PageFooterHeight=   0
      PrintInfos.Count=   1
      AllowUpdate     =   0   'False
      DefColWidth     =   0
      HeadLines       =   1
      FootLines       =   1
      Caption         =   "Documentos por Cobrar"
      MultipleLines   =   0
      CellTipsWidth   =   0
      DeadAreaBackColor=   12632256
      RowDividerColor =   12632256
      RowSubDividerColor=   12632256
      DirectionAfterEnter=   1
      DirectionAfterTab=   1
      MaxRows         =   250000
      ViewColumnCaptionWidth=   0
      ViewColumnWidth =   0
      _PropDict       =   "_ExtentX,2003,3;_ExtentY,2004,3;_LayoutType,512,2;_RowHeight,16,3;_StyleDefs,513,0;_WasPersistedAsPixels,516,2"
      _StyleDefs(0)   =   "_StyleRoot:id=0,.parent=-1,.alignment=3,.valignment=0,.bgcolor=&H80000005&"
      _StyleDefs(1)   =   ":id=0,.fgcolor=&H80000008&,.wraptext=0,.locked=0,.transparentBmp=0"
      _StyleDefs(2)   =   ":id=0,.fgpicPosition=0,.bgpicMode=0,.appearance=0,.borderSize=0,.ellipsis=0"
      _StyleDefs(3)   =   ":id=0,.borderColor=&H80000005&,.borderType=0,.bold=0,.fontsize=825,.italic=0"
      _StyleDefs(4)   =   ":id=0,.underline=0,.strikethrough=0,.charset=0"
      _StyleDefs(5)   =   ":id=0,.fontname=MS Sans Serif"
      _StyleDefs(6)   =   "Style:id=1,.parent=0,.namedParent=33,.bgcolor=&H80000018&,.locked=0"
      _StyleDefs(7)   =   "CaptionStyle:id=4,.parent=2,.alignment=2,.bgcolor=&H8000000A&,.fgcolor=&H0&"
      _StyleDefs(8)   =   ":id=4,.bold=-1,.fontsize=825,.italic=0,.underline=0,.strikethrough=0,.charset=0"
      _StyleDefs(9)   =   ":id=4,.fontname=Arial"
      _StyleDefs(10)  =   "HeadingStyle:id=2,.parent=1,.bgcolor=&H80000000&,.borderSize=1,.bold=-1"
      _StyleDefs(11)  =   ":id=2,.fontsize=825,.italic=0,.underline=0,.strikethrough=0,.charset=0"
      _StyleDefs(12)  =   ":id=2,.fontname=Arial"
      _StyleDefs(13)  =   "FooterStyle:id=3,.parent=1"
      _StyleDefs(14)  =   "InactiveStyle:id=5,.parent=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
      _StyleDefs(15)  =   "SelectedStyle:id=6,.parent=1,.bgcolor=&HE7FAB6&"
      _StyleDefs(16)  =   "EditorStyle:id=7,.parent=1"
      _StyleDefs(17)  =   "HighlightRowStyle:id=8,.parent=1,.bgcolor=&H808000&"
      _StyleDefs(18)  =   "EvenRowStyle:id=9,.parent=1"
      _StyleDefs(19)  =   "OddRowStyle:id=10,.parent=1"
      _StyleDefs(20)  =   "RecordSelectorStyle:id=11,.parent=2,.namedParent=41"
      _StyleDefs(21)  =   "FilterBarStyle:id=12,.parent=1"
      _StyleDefs(22)  =   "Splits(0).Style:id=13,.parent=1"
      _StyleDefs(23)  =   "Splits(0).CaptionStyle:id=22,.parent=4"
      _StyleDefs(24)  =   "Splits(0).HeadingStyle:id=14,.parent=2"
      _StyleDefs(25)  =   "Splits(0).FooterStyle:id=15,.parent=3"
      _StyleDefs(26)  =   "Splits(0).InactiveStyle:id=16,.parent=5"
      _StyleDefs(27)  =   "Splits(0).SelectedStyle:id=18,.parent=6"
      _StyleDefs(28)  =   "Splits(0).EditorStyle:id=17,.parent=7"
      _StyleDefs(29)  =   "Splits(0).HighlightRowStyle:id=19,.parent=8,.namedParent=38"
      _StyleDefs(30)  =   "Splits(0).EvenRowStyle:id=20,.parent=9"
      _StyleDefs(31)  =   "Splits(0).OddRowStyle:id=21,.parent=10"
      _StyleDefs(32)  =   "Splits(0).RecordSelectorStyle:id=23,.parent=11"
      _StyleDefs(33)  =   "Splits(0).FilterBarStyle:id=24,.parent=12"
      _StyleDefs(34)  =   "Splits(0).Columns(0).Style:id=32,.parent=13"
      _StyleDefs(35)  =   "Splits(0).Columns(0).HeadingStyle:id=29,.parent=14"
      _StyleDefs(36)  =   "Splits(0).Columns(0).FooterStyle:id=30,.parent=15"
      _StyleDefs(37)  =   "Splits(0).Columns(0).EditorStyle:id=31,.parent=17"
      _StyleDefs(38)  =   "Named:id=33:Normal"
      _StyleDefs(39)  =   ":id=33,.parent=0,.valignment=2,.bgcolor=&H80000018&,.locked=-1,.appearance=0"
      _StyleDefs(40)  =   ":id=33,.borderSize=1,.borderColor=&H80000005&,.borderType=0,.bold=0"
      _StyleDefs(41)  =   ":id=33,.fontsize=675,.italic=0,.underline=0,.strikethrough=0,.charset=0"
      _StyleDefs(42)  =   ":id=33,.fontname=Small Fonts"
      _StyleDefs(43)  =   "Named:id=34:Heading"
      _StyleDefs(44)  =   ":id=34,.parent=33,.alignment=2,.valignment=2,.bgcolor=&HC0C0C0&"
      _StyleDefs(45)  =   ":id=34,.fgcolor=&H80000012&,.wraptext=-1,.locked=0,.borderSize=1,.bold=-1"
      _StyleDefs(46)  =   ":id=34,.fontsize=975,.italic=0,.underline=0,.strikethrough=0,.charset=0"
      _StyleDefs(47)  =   ":id=34,.fontname=Arial"
      _StyleDefs(48)  =   "Named:id=35:Footing"
      _StyleDefs(49)  =   ":id=35,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
      _StyleDefs(50)  =   "Named:id=36:Selected"
      _StyleDefs(51)  =   ":id=36,.parent=33,.bgcolor=&HE7FAB6&,.fgcolor=&H0&,.borderColor=&H808000&"
      _StyleDefs(52)  =   ":id=36,.bold=-1,.fontsize=900,.italic=0,.underline=0,.strikethrough=0,.charset=0"
      _StyleDefs(53)  =   ":id=36,.fontname=Arial"
      _StyleDefs(54)  =   "Named:id=37:Caption"
      _StyleDefs(55)  =   ":id=37,.parent=34,.alignment=2"
      _StyleDefs(56)  =   "Named:id=38:HighlightRow"
      _StyleDefs(57)  =   ":id=38,.parent=33,.bgcolor=&HE7FAB6&,.fgcolor=&H80000012&,.bold=-1,.fontsize=675"
      _StyleDefs(58)  =   ":id=38,.italic=0,.underline=0,.strikethrough=0,.charset=0"
      _StyleDefs(59)  =   ":id=38,.fontname=Small Fonts"
      _StyleDefs(60)  =   "Named:id=39:EvenRow"
      _StyleDefs(61)  =   ":id=39,.parent=33,.bgcolor=&HFFFF00&"
      _StyleDefs(62)  =   "Named:id=40:OddRow"
      _StyleDefs(63)  =   ":id=40,.parent=33"
      _StyleDefs(64)  =   "Named:id=41:RecordSelector"
      _StyleDefs(65)  =   ":id=41,.parent=34"
      _StyleDefs(66)  =   "Named:id=42:FilterBar"
      _StyleDefs(67)  =   ":id=42,.parent=33"
   End
   Begin VB.CheckBox ChkDetalle 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   585
      Left            =   10125
      Picture         =   "frmDocumento.frx":0442
      Style           =   1  'Graphical
      TabIndex        =   41
      Top             =   7695
      Width           =   585
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
      Height          =   585
      Index           =   0
      Left            =   10710
      Picture         =   "frmDocumento.frx":09CC
      Style           =   1  'Graphical
      TabIndex        =   40
      Top             =   7695
      Width           =   1170
   End
   Begin VB.CommandButton cmdOpcion 
      Caption         =   "Pagos del Documento"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   585
      Index           =   6
      Left            =   8865
      TabIndex        =   39
      Top             =   7695
      Width           =   1265
   End
   Begin VB.CommandButton cmdOpcion 
      Caption         =   "Re Impresi�n del Documento"
      Height          =   585
      Index           =   7
      Left            =   6330
      TabIndex        =   38
      Top             =   7695
      Width           =   1265
   End
   Begin VB.CommandButton cmdOpcion 
      Caption         =   "Cuentas por Cobrar"
      Height          =   585
      Index           =   8
      Left            =   5070
      TabIndex        =   37
      Top             =   7695
      Width           =   1265
   End
   Begin VB.CommandButton cmdOpcion 
      Caption         =   "Anulaci�n Documento"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   585
      Index           =   9
      Left            =   3810
      TabIndex        =   36
      Top             =   7695
      Width           =   1265
   End
   Begin VB.CommandButton cmdOpcion 
      Caption         =   "Forma Pago Delivery"
      Height          =   585
      Index           =   1
      Left            =   30
      TabIndex        =   35
      Top             =   7695
      Width           =   1265
   End
   Begin VB.CommandButton cmdOpcion 
      Caption         =   "PreCuenta Delivery"
      Height          =   585
      Index           =   10
      Left            =   1290
      Style           =   1  'Graphical
      TabIndex        =   34
      Top             =   7695
      Width           =   1265
   End
   Begin VB.CommandButton cmdOpcion 
      Caption         =   "Anulaci�n de Pagos"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   585
      Index           =   11
      Left            =   7590
      TabIndex        =   33
      Top             =   7695
      Width           =   1270
   End
   Begin VB.CommandButton cmdOpcion 
      Caption         =   "Liquidacion Delivery"
      Height          =   585
      Index           =   12
      Left            =   2550
      Style           =   1  'Graphical
      TabIndex        =   32
      Top             =   7695
      Width           =   1265
   End
   Begin VB.Frame Frame1 
      Caption         =   " Dividir en "
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
      Height          =   615
      Left            =   9750
      TabIndex        =   21
      Top             =   6990
      Width           =   2070
      Begin VB.Label txtTotal 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "28,500.00"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   300
         Left            =   810
         TabIndex        =   24
         Top             =   210
         Width           =   1170
      End
      Begin VB.Label txtDivision 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "1"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   300
         Left            =   90
         TabIndex        =   23
         Top             =   210
         Width           =   450
      End
      Begin VB.Label Label2 
         Caption         =   "x"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   630
         TabIndex        =   22
         Top             =   270
         Width           =   105
      End
   End
   Begin VB.CommandButton cmdNavegar 
      Height          =   600
      Index           =   11
      Left            =   9750
      Picture         =   "frmDocumento.frx":0ABE
      Style           =   1  'Graphical
      TabIndex        =   20
      Top             =   4500
      Width           =   735
   End
   Begin VB.CommandButton cmdOpcion 
      Caption         =   "Generar Documento"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   600
      Index           =   2
      Left            =   10530
      TabIndex        =   19
      Top             =   4500
      Width           =   1305
   End
   Begin VB.CommandButton cmdNavegar 
      Height          =   600
      Index           =   8
      Left            =   9750
      Picture         =   "frmDocumento.frx":1388
      Style           =   1  'Graphical
      TabIndex        =   18
      Top             =   6300
      Width           =   735
   End
   Begin VB.CommandButton cmdNavegar 
      Height          =   600
      Index           =   9
      Left            =   9750
      Picture         =   "frmDocumento.frx":1C52
      Style           =   1  'Graphical
      TabIndex        =   17
      Top             =   5700
      Width           =   735
   End
   Begin VB.CommandButton cmdNavegar 
      Height          =   600
      Index           =   10
      Left            =   9750
      Picture         =   "frmDocumento.frx":251C
      Style           =   1  'Graphical
      TabIndex        =   16
      Top             =   5100
      Width           =   735
   End
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
      ForeColor       =   &H00808000&
      Height          =   4455
      Left            =   30
      TabIndex        =   3
      Top             =   -15
      Width           =   11805
      Begin VB.CommandButton cmdNavegar 
         Height          =   555
         Index           =   0
         Left            =   10470
         Picture         =   "frmDocumento.frx":2DE6
         Style           =   1  'Graphical
         TabIndex        =   15
         Top             =   3795
         Width           =   1215
      End
      Begin VB.CommandButton cmdNavegar 
         Height          =   555
         Index           =   1
         Left            =   9180
         Picture         =   "frmDocumento.frx":36B0
         Style           =   1  'Graphical
         TabIndex        =   14
         Top             =   3795
         Width           =   1215
      End
      Begin VB.CommandButton cmdNavegar 
         Height          =   555
         Index           =   2
         Left            =   7890
         Picture         =   "frmDocumento.frx":3F7A
         Style           =   1  'Graphical
         TabIndex        =   13
         Top             =   3795
         Width           =   1215
      End
      Begin VB.CommandButton cmdNavegar 
         Height          =   555
         Index           =   3
         Left            =   6600
         Picture         =   "frmDocumento.frx":4844
         Style           =   1  'Graphical
         TabIndex        =   12
         Top             =   3795
         Width           =   1215
      End
      Begin VB.CommandButton cmdNavegar 
         Height          =   555
         Index           =   4
         Left            =   3960
         Picture         =   "frmDocumento.frx":510E
         Style           =   1  'Graphical
         TabIndex        =   11
         Top             =   3795
         Width           =   1215
      End
      Begin VB.CommandButton cmdNavegar 
         Height          =   555
         Index           =   5
         Left            =   2670
         Picture         =   "frmDocumento.frx":59D8
         Style           =   1  'Graphical
         TabIndex        =   10
         Top             =   3795
         Width           =   1215
      End
      Begin VB.CommandButton cmdNavegar 
         Height          =   555
         Index           =   6
         Left            =   1380
         Picture         =   "frmDocumento.frx":62A2
         Style           =   1  'Graphical
         TabIndex        =   9
         Top             =   3795
         Width           =   1215
      End
      Begin VB.CommandButton cmdNavegar 
         Height          =   555
         Index           =   7
         Left            =   90
         Picture         =   "frmDocumento.frx":6B6C
         Style           =   1  'Graphical
         TabIndex        =   8
         Top             =   3795
         Width           =   1215
      End
      Begin VB.CommandButton cmdMovimiento 
         Caption         =   "Pasar Todos"
         Height          =   645
         Index           =   0
         Left            =   5250
         Picture         =   "frmDocumento.frx":7436
         Style           =   1  'Graphical
         TabIndex        =   7
         Top             =   420
         Width           =   1275
      End
      Begin VB.CommandButton cmdMovimiento 
         Caption         =   "Pasar Uno"
         Height          =   645
         Index           =   2
         Left            =   5250
         Picture         =   "frmDocumento.frx":79C0
         Style           =   1  'Graphical
         TabIndex        =   6
         Top             =   1125
         Width           =   1275
      End
      Begin VB.CommandButton cmdMovimiento 
         Caption         =   "Quitar Todos"
         Height          =   645
         Index           =   1
         Left            =   5250
         Picture         =   "frmDocumento.frx":7F4A
         Style           =   1  'Graphical
         TabIndex        =   5
         Top             =   2730
         Width           =   1275
      End
      Begin VB.CommandButton cmdMovimiento 
         Caption         =   "Quitar Uno"
         Height          =   645
         Index           =   3
         Left            =   5250
         Picture         =   "frmDocumento.frx":828C
         Style           =   1  'Graphical
         TabIndex        =   4
         Top             =   2025
         Width           =   1275
      End
      Begin TrueOleDBGrid80.TDBGrid grdDetalle2 
         Height          =   3225
         Left            =   90
         TabIndex        =   25
         Top             =   150
         Width           =   5085
         _ExtentX        =   8969
         _ExtentY        =   5689
         _LayoutType     =   4
         _RowHeight      =   23
         _WasPersistedAsPixels=   0
         Columns(0)._VlistStyle=   0
         Columns(0)._MaxComboItems=   5
         Columns(0).DataField=   ""
         Columns(0).NumberFormat=   "True/False"
         Columns(0)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns.Count   =   1
         Splits(0)._UserFlags=   0
         Splits(0).MarqueeStyle=   3
         Splits(0).RecordSelectorWidth=   503
         Splits(0)._SavedRecordSelectors=   -1  'True
         Splits(0).ScrollBars=   0
         Splits(0).AllowColSelect=   0   'False
         Splits(0).FetchRowStyle=   -1  'True
         Splits(0).DividerStyle=   2
         Splits(0).DividerColor=   32768
         Splits(0).SpringMode=   0   'False
         Splits(0)._PropDict=   "_ColumnProps,515,0;_UserFlags,518,3"
         Splits(0)._ColumnProps(0)=   "Columns.Count=1"
         Splits(0)._ColumnProps(1)=   "Column(0).Width=2725"
         Splits(0)._ColumnProps(2)=   "Column(0).DividerColor=0"
         Splits(0)._ColumnProps(3)=   "Column(0)._WidthInPix=2646"
         Splits(0)._ColumnProps(4)=   "Column(0)._EditAlways=0"
         Splits(0)._ColumnProps(5)=   "Column(0).AllowSizing=0"
         Splits(0)._ColumnProps(6)=   "Column(0)._ColStyle=20"
         Splits(0)._ColumnProps(7)=   "Column(0).Order=1"
         Splits.Count    =   1
         PrintInfos(0)._StateFlags=   0
         PrintInfos(0).Name=   "piInternal 0"
         PrintInfos(0).PageHeaderFont=   "Size=6.75,Charset=0,Weight=400,Underline=0,Italic=0,Strikethrough=0,Name=Small Fonts"
         PrintInfos(0).PageFooterFont=   "Size=6.75,Charset=0,Weight=400,Underline=0,Italic=0,Strikethrough=0,Name=Small Fonts"
         PrintInfos(0).PageHeaderHeight=   0
         PrintInfos(0).PageFooterHeight=   0
         PrintInfos.Count=   1
         AllowUpdate     =   0   'False
         DefColWidth     =   0
         HeadLines       =   1
         FootLines       =   1
         Caption         =   "Productos a Facturar"
         MultipleLines   =   0
         CellTipsWidth   =   0
         DeadAreaBackColor=   12632256
         RowDividerColor =   12632256
         RowSubDividerColor=   12632256
         DirectionAfterEnter=   1
         DirectionAfterTab=   1
         MaxRows         =   250000
         ViewColumnCaptionWidth=   0
         ViewColumnWidth =   0
         _PropDict       =   "_ExtentX,2003,3;_ExtentY,2004,3;_LayoutType,512,2;_RowHeight,16,3;_StyleDefs,513,0;_WasPersistedAsPixels,516,2"
         _StyleDefs(0)   =   "_StyleRoot:id=0,.parent=-1,.alignment=3,.valignment=0,.bgcolor=&H80000005&"
         _StyleDefs(1)   =   ":id=0,.fgcolor=&H80000008&,.wraptext=0,.locked=0,.transparentBmp=0"
         _StyleDefs(2)   =   ":id=0,.fgpicPosition=0,.bgpicMode=0,.appearance=0,.borderSize=0,.ellipsis=0"
         _StyleDefs(3)   =   ":id=0,.borderColor=&H80000005&,.borderType=0,.bold=0,.fontsize=825,.italic=0"
         _StyleDefs(4)   =   ":id=0,.underline=0,.strikethrough=0,.charset=0"
         _StyleDefs(5)   =   ":id=0,.fontname=MS Sans Serif"
         _StyleDefs(6)   =   "Style:id=1,.parent=0,.namedParent=33,.bgcolor=&H80000018&,.locked=0"
         _StyleDefs(7)   =   "CaptionStyle:id=4,.parent=2,.alignment=2,.bgcolor=&H8000000A&,.fgcolor=&H0&"
         _StyleDefs(8)   =   ":id=4,.bold=-1,.fontsize=825,.italic=0,.underline=0,.strikethrough=0,.charset=0"
         _StyleDefs(9)   =   ":id=4,.fontname=Arial"
         _StyleDefs(10)  =   "HeadingStyle:id=2,.parent=1,.bgcolor=&H80000000&,.borderSize=1,.bold=-1"
         _StyleDefs(11)  =   ":id=2,.fontsize=825,.italic=0,.underline=0,.strikethrough=0,.charset=0"
         _StyleDefs(12)  =   ":id=2,.fontname=Arial"
         _StyleDefs(13)  =   "FooterStyle:id=3,.parent=1"
         _StyleDefs(14)  =   "InactiveStyle:id=5,.parent=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
         _StyleDefs(15)  =   "SelectedStyle:id=6,.parent=1,.bgcolor=&HE7FAB6&"
         _StyleDefs(16)  =   "EditorStyle:id=7,.parent=1"
         _StyleDefs(17)  =   "HighlightRowStyle:id=8,.parent=1,.bgcolor=&H808000&"
         _StyleDefs(18)  =   "EvenRowStyle:id=9,.parent=1"
         _StyleDefs(19)  =   "OddRowStyle:id=10,.parent=1"
         _StyleDefs(20)  =   "RecordSelectorStyle:id=11,.parent=2,.namedParent=41"
         _StyleDefs(21)  =   "FilterBarStyle:id=12,.parent=1"
         _StyleDefs(22)  =   "Splits(0).Style:id=13,.parent=1"
         _StyleDefs(23)  =   "Splits(0).CaptionStyle:id=22,.parent=4"
         _StyleDefs(24)  =   "Splits(0).HeadingStyle:id=14,.parent=2"
         _StyleDefs(25)  =   "Splits(0).FooterStyle:id=15,.parent=3"
         _StyleDefs(26)  =   "Splits(0).InactiveStyle:id=16,.parent=5"
         _StyleDefs(27)  =   "Splits(0).SelectedStyle:id=18,.parent=6"
         _StyleDefs(28)  =   "Splits(0).EditorStyle:id=17,.parent=7"
         _StyleDefs(29)  =   "Splits(0).HighlightRowStyle:id=19,.parent=8,.namedParent=38"
         _StyleDefs(30)  =   "Splits(0).EvenRowStyle:id=20,.parent=9"
         _StyleDefs(31)  =   "Splits(0).OddRowStyle:id=21,.parent=10"
         _StyleDefs(32)  =   "Splits(0).RecordSelectorStyle:id=23,.parent=11"
         _StyleDefs(33)  =   "Splits(0).FilterBarStyle:id=24,.parent=12"
         _StyleDefs(34)  =   "Splits(0).Columns(0).Style:id=32,.parent=13"
         _StyleDefs(35)  =   "Splits(0).Columns(0).HeadingStyle:id=29,.parent=14"
         _StyleDefs(36)  =   "Splits(0).Columns(0).FooterStyle:id=30,.parent=15"
         _StyleDefs(37)  =   "Splits(0).Columns(0).EditorStyle:id=31,.parent=17"
         _StyleDefs(38)  =   "Named:id=33:Normal"
         _StyleDefs(39)  =   ":id=33,.parent=0,.valignment=2,.bgcolor=&H80000018&,.locked=-1,.appearance=0"
         _StyleDefs(40)  =   ":id=33,.borderSize=1,.borderColor=&H80000005&,.borderType=0,.bold=0"
         _StyleDefs(41)  =   ":id=33,.fontsize=675,.italic=0,.underline=0,.strikethrough=0,.charset=0"
         _StyleDefs(42)  =   ":id=33,.fontname=Small Fonts"
         _StyleDefs(43)  =   "Named:id=34:Heading"
         _StyleDefs(44)  =   ":id=34,.parent=33,.alignment=2,.valignment=2,.bgcolor=&HC0C0C0&"
         _StyleDefs(45)  =   ":id=34,.fgcolor=&H80000012&,.wraptext=-1,.locked=0,.borderSize=1,.bold=-1"
         _StyleDefs(46)  =   ":id=34,.fontsize=975,.italic=0,.underline=0,.strikethrough=0,.charset=0"
         _StyleDefs(47)  =   ":id=34,.fontname=Arial"
         _StyleDefs(48)  =   "Named:id=35:Footing"
         _StyleDefs(49)  =   ":id=35,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
         _StyleDefs(50)  =   "Named:id=36:Selected"
         _StyleDefs(51)  =   ":id=36,.parent=33,.bgcolor=&HE7FAB6&,.fgcolor=&H0&,.borderColor=&H808000&"
         _StyleDefs(52)  =   ":id=36,.bold=-1,.fontsize=900,.italic=0,.underline=0,.strikethrough=0,.charset=0"
         _StyleDefs(53)  =   ":id=36,.fontname=Arial"
         _StyleDefs(54)  =   "Named:id=37:Caption"
         _StyleDefs(55)  =   ":id=37,.parent=34,.alignment=2"
         _StyleDefs(56)  =   "Named:id=38:HighlightRow"
         _StyleDefs(57)  =   ":id=38,.parent=33,.bgcolor=&HE7FAB6&,.fgcolor=&H80000012&,.bold=-1,.fontsize=675"
         _StyleDefs(58)  =   ":id=38,.italic=0,.underline=0,.strikethrough=0,.charset=0"
         _StyleDefs(59)  =   ":id=38,.fontname=Small Fonts"
         _StyleDefs(60)  =   "Named:id=39:EvenRow"
         _StyleDefs(61)  =   ":id=39,.parent=33,.bgcolor=&HFFFF00&"
         _StyleDefs(62)  =   "Named:id=40:OddRow"
         _StyleDefs(63)  =   ":id=40,.parent=33"
         _StyleDefs(64)  =   "Named:id=41:RecordSelector"
         _StyleDefs(65)  =   ":id=41,.parent=34"
         _StyleDefs(66)  =   "Named:id=42:FilterBar"
         _StyleDefs(67)  =   ":id=42,.parent=33"
      End
      Begin TrueOleDBGrid80.TDBGrid grdDetalle1 
         Height          =   3225
         Left            =   6600
         TabIndex        =   26
         Top             =   150
         Width           =   5085
         _ExtentX        =   8969
         _ExtentY        =   5689
         _LayoutType     =   4
         _RowHeight      =   23
         _WasPersistedAsPixels=   0
         Columns(0)._VlistStyle=   0
         Columns(0)._MaxComboItems=   5
         Columns(0).DataField=   ""
         Columns(0).NumberFormat=   "True/False"
         Columns(0)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns.Count   =   1
         Splits(0)._UserFlags=   0
         Splits(0).MarqueeStyle=   3
         Splits(0).RecordSelectorWidth=   503
         Splits(0)._SavedRecordSelectors=   -1  'True
         Splits(0).ScrollBars=   0
         Splits(0).AllowColSelect=   0   'False
         Splits(0).FetchRowStyle=   -1  'True
         Splits(0).DividerStyle=   2
         Splits(0).DividerColor=   32768
         Splits(0).SpringMode=   0   'False
         Splits(0)._PropDict=   "_ColumnProps,515,0;_UserFlags,518,3"
         Splits(0)._ColumnProps(0)=   "Columns.Count=1"
         Splits(0)._ColumnProps(1)=   "Column(0).Width=2725"
         Splits(0)._ColumnProps(2)=   "Column(0).DividerColor=0"
         Splits(0)._ColumnProps(3)=   "Column(0)._WidthInPix=2646"
         Splits(0)._ColumnProps(4)=   "Column(0)._EditAlways=0"
         Splits(0)._ColumnProps(5)=   "Column(0).AllowSizing=0"
         Splits(0)._ColumnProps(6)=   "Column(0)._ColStyle=20"
         Splits(0)._ColumnProps(7)=   "Column(0).Order=1"
         Splits.Count    =   1
         PrintInfos(0)._StateFlags=   0
         PrintInfos(0).Name=   "piInternal 0"
         PrintInfos(0).PageHeaderFont=   "Size=6.75,Charset=0,Weight=400,Underline=0,Italic=0,Strikethrough=0,Name=Small Fonts"
         PrintInfos(0).PageFooterFont=   "Size=6.75,Charset=0,Weight=400,Underline=0,Italic=0,Strikethrough=0,Name=Small Fonts"
         PrintInfos(0).PageHeaderHeight=   0
         PrintInfos(0).PageFooterHeight=   0
         PrintInfos.Count=   1
         AllowUpdate     =   0   'False
         DefColWidth     =   0
         HeadLines       =   1
         FootLines       =   1
         Caption         =   "Productos por Facturar"
         MultipleLines   =   0
         CellTipsWidth   =   0
         DeadAreaBackColor=   12632256
         RowDividerColor =   12632256
         RowSubDividerColor=   12632256
         DirectionAfterEnter=   1
         DirectionAfterTab=   1
         MaxRows         =   250000
         ViewColumnCaptionWidth=   0
         ViewColumnWidth =   0
         _PropDict       =   "_ExtentX,2003,3;_ExtentY,2004,3;_LayoutType,512,2;_RowHeight,16,3;_StyleDefs,513,0;_WasPersistedAsPixels,516,2"
         _StyleDefs(0)   =   "_StyleRoot:id=0,.parent=-1,.alignment=3,.valignment=0,.bgcolor=&H80000005&"
         _StyleDefs(1)   =   ":id=0,.fgcolor=&H80000008&,.wraptext=0,.locked=0,.transparentBmp=0"
         _StyleDefs(2)   =   ":id=0,.fgpicPosition=0,.bgpicMode=0,.appearance=0,.borderSize=0,.ellipsis=0"
         _StyleDefs(3)   =   ":id=0,.borderColor=&H80000005&,.borderType=0,.bold=0,.fontsize=825,.italic=0"
         _StyleDefs(4)   =   ":id=0,.underline=0,.strikethrough=0,.charset=0"
         _StyleDefs(5)   =   ":id=0,.fontname=MS Sans Serif"
         _StyleDefs(6)   =   "Style:id=1,.parent=0,.namedParent=33,.bgcolor=&H80000018&,.locked=0"
         _StyleDefs(7)   =   "CaptionStyle:id=4,.parent=2,.alignment=2,.bgcolor=&H8000000A&,.fgcolor=&H0&"
         _StyleDefs(8)   =   ":id=4,.bold=-1,.fontsize=825,.italic=0,.underline=0,.strikethrough=0,.charset=0"
         _StyleDefs(9)   =   ":id=4,.fontname=Arial"
         _StyleDefs(10)  =   "HeadingStyle:id=2,.parent=1,.bgcolor=&H80000000&,.borderSize=1,.bold=-1"
         _StyleDefs(11)  =   ":id=2,.fontsize=825,.italic=0,.underline=0,.strikethrough=0,.charset=0"
         _StyleDefs(12)  =   ":id=2,.fontname=Arial"
         _StyleDefs(13)  =   "FooterStyle:id=3,.parent=1"
         _StyleDefs(14)  =   "InactiveStyle:id=5,.parent=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
         _StyleDefs(15)  =   "SelectedStyle:id=6,.parent=1,.bgcolor=&HE7FAB6&"
         _StyleDefs(16)  =   "EditorStyle:id=7,.parent=1"
         _StyleDefs(17)  =   "HighlightRowStyle:id=8,.parent=1,.bgcolor=&H808000&"
         _StyleDefs(18)  =   "EvenRowStyle:id=9,.parent=1"
         _StyleDefs(19)  =   "OddRowStyle:id=10,.parent=1"
         _StyleDefs(20)  =   "RecordSelectorStyle:id=11,.parent=2,.namedParent=41"
         _StyleDefs(21)  =   "FilterBarStyle:id=12,.parent=1"
         _StyleDefs(22)  =   "Splits(0).Style:id=13,.parent=1"
         _StyleDefs(23)  =   "Splits(0).CaptionStyle:id=22,.parent=4"
         _StyleDefs(24)  =   "Splits(0).HeadingStyle:id=14,.parent=2"
         _StyleDefs(25)  =   "Splits(0).FooterStyle:id=15,.parent=3"
         _StyleDefs(26)  =   "Splits(0).InactiveStyle:id=16,.parent=5"
         _StyleDefs(27)  =   "Splits(0).SelectedStyle:id=18,.parent=6"
         _StyleDefs(28)  =   "Splits(0).EditorStyle:id=17,.parent=7"
         _StyleDefs(29)  =   "Splits(0).HighlightRowStyle:id=19,.parent=8,.namedParent=38"
         _StyleDefs(30)  =   "Splits(0).EvenRowStyle:id=20,.parent=9"
         _StyleDefs(31)  =   "Splits(0).OddRowStyle:id=21,.parent=10"
         _StyleDefs(32)  =   "Splits(0).RecordSelectorStyle:id=23,.parent=11"
         _StyleDefs(33)  =   "Splits(0).FilterBarStyle:id=24,.parent=12"
         _StyleDefs(34)  =   "Splits(0).Columns(0).Style:id=32,.parent=13"
         _StyleDefs(35)  =   "Splits(0).Columns(0).HeadingStyle:id=29,.parent=14"
         _StyleDefs(36)  =   "Splits(0).Columns(0).FooterStyle:id=30,.parent=15"
         _StyleDefs(37)  =   "Splits(0).Columns(0).EditorStyle:id=31,.parent=17"
         _StyleDefs(38)  =   "Named:id=33:Normal"
         _StyleDefs(39)  =   ":id=33,.parent=0,.valignment=2,.bgcolor=&H80000018&,.locked=-1,.appearance=0"
         _StyleDefs(40)  =   ":id=33,.borderSize=1,.borderColor=&H80000005&,.borderType=0,.bold=0"
         _StyleDefs(41)  =   ":id=33,.fontsize=675,.italic=0,.underline=0,.strikethrough=0,.charset=0"
         _StyleDefs(42)  =   ":id=33,.fontname=Small Fonts"
         _StyleDefs(43)  =   "Named:id=34:Heading"
         _StyleDefs(44)  =   ":id=34,.parent=33,.alignment=2,.valignment=2,.bgcolor=&HC0C0C0&"
         _StyleDefs(45)  =   ":id=34,.fgcolor=&H80000012&,.wraptext=-1,.locked=0,.borderSize=1,.bold=-1"
         _StyleDefs(46)  =   ":id=34,.fontsize=975,.italic=0,.underline=0,.strikethrough=0,.charset=0"
         _StyleDefs(47)  =   ":id=34,.fontname=Arial"
         _StyleDefs(48)  =   "Named:id=35:Footing"
         _StyleDefs(49)  =   ":id=35,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
         _StyleDefs(50)  =   "Named:id=36:Selected"
         _StyleDefs(51)  =   ":id=36,.parent=33,.bgcolor=&HE7FAB6&,.fgcolor=&H0&,.borderColor=&H808000&"
         _StyleDefs(52)  =   ":id=36,.bold=-1,.fontsize=900,.italic=0,.underline=0,.strikethrough=0,.charset=0"
         _StyleDefs(53)  =   ":id=36,.fontname=Arial"
         _StyleDefs(54)  =   "Named:id=37:Caption"
         _StyleDefs(55)  =   ":id=37,.parent=34,.alignment=2"
         _StyleDefs(56)  =   "Named:id=38:HighlightRow"
         _StyleDefs(57)  =   ":id=38,.parent=33,.bgcolor=&HE7FAB6&,.fgcolor=&H80000012&,.bold=-1,.fontsize=675"
         _StyleDefs(58)  =   ":id=38,.italic=0,.underline=0,.strikethrough=0,.charset=0"
         _StyleDefs(59)  =   ":id=38,.fontname=Small Fonts"
         _StyleDefs(60)  =   "Named:id=39:EvenRow"
         _StyleDefs(61)  =   ":id=39,.parent=33,.bgcolor=&HFFFF00&"
         _StyleDefs(62)  =   "Named:id=40:OddRow"
         _StyleDefs(63)  =   ":id=40,.parent=33"
         _StyleDefs(64)  =   "Named:id=41:RecordSelector"
         _StyleDefs(65)  =   ":id=41,.parent=34"
         _StyleDefs(66)  =   "Named:id=42:FilterBar"
         _StyleDefs(67)  =   ":id=42,.parent=33"
      End
      Begin VB.Label txtTotal1 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   6.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   240
         Left            =   10455
         TabIndex        =   31
         Top             =   3465
         Width           =   1230
      End
      Begin VB.Label Label 
         AutoSize        =   -1  'True
         Caption         =   "Total a Pagar :"
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   165
         Index           =   3
         Left            =   9495
         TabIndex        =   30
         Top             =   3495
         Width           =   900
      End
      Begin VB.Label txtTotal2 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   6.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   240
         Left            =   3945
         TabIndex        =   29
         Top             =   3465
         Width           =   1230
      End
      Begin VB.Label Label 
         AutoSize        =   -1  'True
         Caption         =   "Total a Pagar :"
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   165
         Index           =   4
         Left            =   2970
         TabIndex        =   28
         Top             =   3480
         Width           =   900
      End
   End
   Begin VB.CommandButton cmdOpcion 
      Caption         =   "Agregar Pedido"
      Height          =   600
      Index           =   3
      Left            =   10530
      TabIndex        =   2
      Top             =   5100
      Visible         =   0   'False
      Width           =   1305
   End
   Begin VB.CommandButton cmdOpcion 
      Caption         =   "Divisi�n"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   600
      Index           =   4
      Left            =   10530
      Picture         =   "frmDocumento.frx":8816
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   5700
      Width           =   1305
   End
   Begin VB.CommandButton cmdOpcion 
      Caption         =   "Divisi�n"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   600
      Index           =   5
      Left            =   10530
      Picture         =   "frmDocumento.frx":8918
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   6300
      Width           =   1305
   End
   Begin VB.Image imageHash 
      Height          =   615
      Left            =   11880
      Top             =   3720
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.Image imageCab 
      Height          =   375
      Left            =   360
      Top             =   4200
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.Image imagepIE 
      Height          =   135
      Left            =   0
      Top             =   0
      Visible         =   0   'False
      Width           =   135
   End
End
Attribute VB_Name = "frmDocumento"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim RsDetalle1 As Recordset
Dim RsDetalle2 As Recordset
Dim RsDocumento As Recordset
Dim rst1 As Recordset
Dim RsT2 As Recordset
Dim RsT3 As Recordset
Dim RsImpresion As Recordset
Dim RsEmision As Recordset
Dim RsProductoPropiedad As Recordset
Dim RsComboPropiedad As Recordset
' cambio en anulacion por nota de credito
Public RsTparametro As Recordset
'-------------------------------
Dim Area As String
Dim xPedido As String
Dim nTotal As Double
Dim nRespuesta As Integer
Dim lImprimeAlternativa As Boolean
Dim rstFuente As ADODB.Recordset

Dim nTotalPorCobrarCliente As Double
Dim nLineaPorCobrarCliente As Double

'anulacion por nota de credito
'Dim anroDocumento As String
Dim nEmision As Integer

'FACTURACION_E_PERU
Dim RsImpDocumentoE As New Recordset
Dim RsCodigoHash As New ADODB.Recordset
Dim fDocumento As String
Dim lcodigoHash As Boolean
Dim xMontoTexto As String
Dim iImagenCab As Boolean
Dim xImpresionFE As String
Dim xImpresioDE As String
Dim cadenaCodigoHash As String
Dim TimpresionDolaresDelivery As Boolean

               
Private Sub cmdMovimiento_Click(Index As Integer)
    Screen.MousePointer = vbHourglass
    
    Select Case Index
        Case Is = 0  ' Pasar Todas
             Cn.Execute "update " & sTemporal & " set Estado ='X'"
        
        Case Is = 1  ' Quitar Todas
             Cn.Execute "update " & sTemporal & " set Estado =' '"
        
        Case Is = 2  ' Pasar una
             If RsDetalle1.RecordCount <> 0 Then
               Cn.Execute "update " & sTemporal & " set Estado ='X' where tItem = '" & RsDetalle1!tItem & "' and tCodigoPedido = '" & RsDetalle1!tCodigoPedido & "'"
             End If
        
        Case Is = 3  ' Quitar una
             If RsDetalle2.RecordCount <> 0 Then
                Cn.Execute "update " & sTemporal & " set Estado =' ' where tItem = '" & RsDetalle2!tItem & "' and tCodigoPedido = '" & RsDetalle2!tCodigoPedido & "'"
             End If
    End Select
    
    RsDetalle1.Requery
    RsDetalle2.Requery
    CalculaMonto
    Screen.MousePointer = vbDefault
End Sub

Private Sub cmdNavegar_Click(Index As Integer)
    Select Case Index
           Case Is = 0  ' Ultimo Detalle1
                MoverPuntero Ultimo, grdDetalle1
           Case Is = 1  ' Siguiente Detalle1
                MoverPuntero siguiente, grdDetalle1
           Case Is = 2  ' Previo Detalle1
                MoverPuntero previo, grdDetalle1
           Case Is = 3  ' Primero Detalle1
                MoverPuntero Primero, grdDetalle1
           Case Is = 4  ' Ultimo Detalle2
                MoverPuntero Ultimo, grdDetalle2
           Case Is = 5  ' Siguiente Detalle2
                MoverPuntero siguiente, grdDetalle2
           Case Is = 6  ' Previo Detalle2
                MoverPuntero previo, grdDetalle2
           Case Is = 7  ' Primero Detalle2
                MoverPuntero Primero, grdDetalle2
           Case Is = 8 ' Ultimo Documento
                MoverPuntero Ultimo, grdDocumento
           Case Is = 9 ' Siguiente Documento
                MoverPuntero siguiente, grdDocumento
           Case Is = 10  ' Previo Documento
                MoverPuntero previo, grdDocumento
           Case Is = 11  ' Primero Documento
                MoverPuntero Primero, grdDocumento
    End Select
End Sub

Private Sub cmdNotasCredito_Click(Index As Integer)
    Select Case Index
           Case Is = 13  'Nuevo
                Sw = True
                
                If Supervisor("27") = False Then
                      MsgBox "Clave no permitida", vbExclamation, sMensaje
                      Exit Sub
               End If
                  modProcedimiento.aNotaCredito ("documento")
                  frmNotaCreditoDetalle.Show vbModal
             
                
                End Select

            
End Sub
Private Sub cmdOpcion_Click(Index As Integer)
    Dim sUsuarioAutoriza As String
    Dim fso1 As Object
    
    Select Case Index
          Case Is = 0 ' Salir
               Unload Me
               
          Case Is = 1 ' Forma de Pago Delivery
               If RsDocumento.RecordCount = 0 Then
                  Exit Sub
               End If
               
               If RsDocumento!tEstadoDocumento <> "02" Then
                  xTipo = ""
                  nCargo = Round(RsDocumento!nVenta, 2)
                  sDocumento = grdDocumento.Columns(0).value
                  frmPrePago.Show vbModal
               End If
          
          Case Is = 2 ' Generar Documento
               If RsDetalle2.RecordCount <> 0 Then
                  Screen.MousePointer = vbHourglass
                  sTemp = Val(txtDivision.Caption)
                  frmGeneracion.Show vbModal
                  If wEnter Then
                     RsDetalle1.Requery
                     RsDetalle2.Requery
                     RsDocumento.Requery
                     RsDocumento.Requery
                     If RsDocumento.RecordCount <> 0 Then
                        If RsDocumento!tTipoPedido = "02" Then
                           cmdOpcion(1).Enabled = True
                           cmdOpcion(10).Enabled = True
                        Else
                           cmdOpcion(1).Enabled = False
                           cmdOpcion(10).Enabled = False
                        End If
                     End If
                  
                  End If
               Else
                  MsgBox "No existe items a facturar", vbExclamation, sMensaje
               End If
               
               If lPagoAntesImpresion Then
                    Unload Me
               End If
               
          Case Is = 3 ' Agregar Pedido
               sTemp = ""
               Isql = "select *, 'Importar' as Caso from vPedidoGrilla " & _
                      "where Codigo not in (select tCodigoPedido from " & sTemporal & " group by tCodigoPedido) and tCaja='" & sCaja & "' and tEstadoPedido='01' and tTipoPedido<>'04' "

               Call ConfGrilla(7, frmBusquedaRapida.grdGrilla, "Funci�n", 2, "Caso", 1000, 0, 0, "", _
                                                               "Caja", 2, "tCaja", 600, 2, 0, "", _
                                                               "Mesa", 2, "Mesa", 1200, 0, 0, "", _
                                                               "Observacion", 2, "tObservacion", 2000, 0, 0, "", _
                                                               "Pedido", 2, "Descripcion", 1100, 0, 0, "", _
                                                               "Monto", 2, "Suma", 1100, 1, 0, "###,##0.00", _
                                                               "Usuario", 2, "tUsuario", 1200, 0, 0, "")
               frmBusquedaRapida.nPredeterm = 1
               frmBusquedaRapida.Show vbModal
            
               If wEnter Then
                  sPedido = sCodigo
                  'Cambio Aqui
                  Cn.Execute "Insert into " & sTemporal & " select tCodigoPedido, tItem, tTipoPedido, tCodigoProducto, tCodigoGrupo, tCodigoSubGrupo,   tMoneda, nPrecioNeto,nPrecioImpuesto1, nPrecioImpuesto2,  nPrecioImpuesto3,  nPrecioVenta,    nRecargo,  nDescuento,  nPrecioOficial, nCantidad,nImpuesto1, nImpuesto2,  nImpuesto3,   nVenta, tObservacion, tCortesia, lImprime, tEstadoItem, tArea, lCombinacion,  nCombinacion, lImprimeArea,   tFacturado,  tDocumento, lTransferido,  tComanda, fRegistro, tMozoD,tUsuarioD, nInsumo, nGasto, nManoObra,  nOrden,lCorte,tPosicion, fEnvio,nEnvio, tUnidadNegocio,tOferta, tAutorizaOferta,  tsubalmacen, ' ' as Estado From DPEDIDO where tCodigoPedido='" & sPedido & "' and (isnull(tFacturado,'0')='0' or len(ltrim(tFacturado)) = 0) and tEstadoItem='N'"
                  RsDetalle1.Requery
                  CalculaMonto
               End If
   
          Case Is = 4 ' Suma
               txtDivision.Caption = Format(Val(txtDivision.Caption) + 1, "###")
               txtTotal.Caption = Format(Round(nTotal / Val(txtDivision.Caption), 2), "###,###,###,##0.00")
            
          Case Is = 5 ' Resta
               txtDivision.Caption = Format(IIf(Val(txtDivision.Caption) = 1, "1", Val(txtDivision.Caption) - 1), "###")
               txtTotal.Caption = Format(nTotal / Val(txtDivision.Caption), "###,###,###,##0.00")
               
          Case Is = 6  ' Pagos
               If RsDocumento.RecordCount = 0 Then
                  Exit Sub
               End If
               If RsDocumento!tTipoDocumento <> "00" Then
               If RsDocumento!tEstadoDocumento <> "02" Then
                  xTipo = ""
                  nCargo = Round(RsDocumento!nVenta, 2)
                  '0j0000000000000000000000000
                  'por los centimos en tarjetas de credito
               '   nCargo = Format(grdDocumento.Columns(5).value, "#,###,##0.00")

                  sDocumento = grdDocumento.Columns(0).value
                  If RsDocumento!tTipoPedido = "02" Then
                     sVar1 = "02"
                  Else
                     sVar1 = "01"
                  End If
                  sFormulario = ""
                  xTipo = ""
                  frmPago.Show vbModal
                  If wEnter Then
                     RsDocumento.Requery
                     If RsDocumento.RecordCount <> 0 Then
                        If RsDocumento!tTipoPedido = "02" Then
                           cmdOpcion(1).Enabled = True
                           cmdOpcion(10).Enabled = True
                        Else
                           cmdOpcion(1).Enabled = False
                           cmdOpcion(10).Enabled = False
                        End If
                     End If
                  End If
               End If
               Else
                MsgBox "Para Tipo de Documento no se registra Pagos"
               End If
                         
          Case Is = 7  ' Re Emite Documento
               lReimpresion = True
               lImprimeAlternativa = False
               If RsDocumento.RecordCount = 0 Then
                  Exit Sub
               End If
               
               RsEmision.MoveFirst
               
               sDocumento = grdDocumento.Columns(0).value
               RsEmision.Find ("Prefijo='" & Mid(sDocumento, 1, 1) & "'")
               If RsEmision.EOF Then
                  MsgBox "Error: Configuraci�n en Caja", vbCritical, sMensaje
                  Exit Sub
               End If
                                                            
               If Supervisor("12") = False Then
                  MsgBox "Clave no permitida", vbExclamation, sMensaje
                  Exit Sub
               End If
               
               Select Case pais
               Case "001" 'Bolivia
                     nRespuesta = MsgBox("Desea imprimir una copia del comprobante?", vbQuestion + vbYesNoCancel + vbDefaultButton2, sMensaje)
                     If nRespuesta = vbYes Then
                         estadoReimpresion = True
                     ElseIf nRespuesta = vbCancel Then
                         Exit Sub
                     End If
                Case "002"
                     estadoReimpresion = True
                Case Else 'Peru
                     estadoReimpresion = False
                End Select
                
                'Imprime Documentos
                  Dim wConsumo As Boolean
                  wConsumo = False
                  
                  nRespuesta = MsgBox("Por Consumo? ", vbQuestion + vbYesNoCancel + vbDefaultButton2, sMensaje)
                  If nRespuesta = vbYes Then
                      'frmKeyBoard.txtResultado = "POR CONSUMO"
                      frmKeyBoard.txtResultado = tTextoConsumo
                      frmKeyBoard.Show vbModal
                      If sDescrip = "" Or Not wEnter Then
                         MsgBox "Error: La descripcion no puede ser en blanco", vbCritical, sMensaje
                         Exit Sub
                      End If
                      wConsumo = True
                  ElseIf nRespuesta = vbCancel Then
                      Exit Sub
                  End If
                
                  If wConsumo = False And lDescripcionAlternativa = True Then
                     If validaImpresionAlternativa(grdDocumento.Columns(0).value) = False Then
                         If MsgBox("Desea imprimir descripci�n Alternativa? ", vbQuestion + vbYesNo + vbDefaultButton2, sMensaje) = vbYes Then
                            lImprimeAlternativa = True
                         End If
                     End If
                  End If
                  
                    TimpresionDolaresDelivery = False
                    '------ impresion en dolares para check de cliente delivery
                    If Calcular("select isnull(lEmisionMonedaExtranjera,0) as codigo from TDELIVERY where tCodigoDelivery=(select tClienteDelivery from MPEDIDO where tCodigoPedido=(select top 1 tCodigoPedido from DPEDIDO where tDocumento='" & sDocumento & "'))", Cn) Then
                        If MsgBox("�Desea Imprimir en " & sMonedaE & "?", vbQuestion + vbYesNo + vbDefaultButton2) = vbYes Then
                            TimpresionDolaresDelivery = True
                        Else
                            TimpresionDolaresDelivery = False
                        End If
                    Else
                        TimpresionDolaresDelivery = False
                    End If
                '-------------------------------  SE INTRDUJO LAS CONSULTAS A  UN STORE PROCEDURE -------------------------
                ' ELDCQ 15/11/2017
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
               '---------------------------------------------------------------------------------------------------------------
               'FACTURACION_E_PERU
               Set RsImpDocumentoE = Lib.OpenRecordset(IsqlFact, Cn)
               '---------------------------------------
               Set RsImpresion = Lib.OpenRecordset(Isql, Cn)
       
                  'impresion imagen
                  Set rstFuente = New ADODB.Recordset
                  imageCab.Picture = Nothing
                  imagepIE.Picture = Nothing
                  Set rstFuente = Lib.OpenRecordset("select iImagenCabDoc AS foto, iImagenPieDoc as fotoPie  from tcaja where tcaja='" & sCaja & "'", Cn)
                  imageCab.DataField = "foto"
                  Set imageCab.DataSource = rstFuente
                  imagepIE.DataField = "fotoPie"
                  Set imagepIE.DataSource = rstFuente
                
                
                  'FACTURACION ELECTRONICA
                  Dim xlFacturacionE As Boolean
                  Dim xlTipoDoc As String
                  
                  xlTipoDoc = Calcular("Select tTipoDocumento As Codigo From MDOCUMENTO Where tDocumento = '" & sDocumento & "'", Cn)
                  RsEmision.Find "tTipoEmision='" & xlTipoDoc & "'"
                  xlFacturacionE = IIf(RsEmision!lFacturacionElectronica = True, 1, 0)
                  xImpresionFE = Calcular(" SELECT tImpresionFE as codigo FROM vTipodocumento WHERE Codigo=(SELECT TTIPODOCUMENTO FROM MDOCUMENTO WHERE tDocumento='" & sDocumento & "')", Cn)
                  xImpresioDE = Mid(sDocumento, 1, 1)
                  '------------------------------------------
                  
                  'FACTURACION_E_PERU
                  If pais = "000" Then
                     If lFacturacionE Then
                            If lFEOfisis Then 'OFISIS
                                 Dim lDocumentoElectronicoOfisis As Boolean
                                 lDocumentoElectronicoOfisis = Calcular("select isnull(tdi.lDocumentoElectronicoOfisis,0) as codigo from TTIPODOCUMENTOIMPRESORA tdi inner join MDOCUMENTO m on tdi.tTipoEmision = m.tTipoDocumento and tdi.tCaja = m.tCaja  where m.tDocumento= '" & sDocumento & "'", Cn)
                                 If lDocumentoElectronicoOfisis Then
                                     fDocumento = Mid(sDocumento, 1, 1) + Mid(sDocumento, 4, 3) + "-" + CStr(CLng(Mid(sDocumento, 8, 8)))
                                     If lImpresionCodigoBarras Then
                                         imageHash.DataField = "foto"
                                         Set RsCodigoHash = Lib.OpenRecordset("USP_FactObtenerCodigoBarraOfisis '" & fDocumento & "','" & Mid(sDocumento, 1, 1) & "','' ", Cn)
                                         Set imageHash.DataSource = RsCodigoHash
                                         
                                    ElseIf lQRFE Then
                                        Set imageHash.Picture = LoadPicture(ImagenQR_Ofisis(fDocumento, sDocumento))
                                     Else
                                         Dim RscadenaCodigoHash As Recordset
                                         Set RscadenaCodigoHash = Lib.OpenRecordset("USP_FactConsultaHash '" & fDocumento & "','0' ", Cn)
                                         If RscadenaCodigoHash.RecordCount > 0 Then
                                              cadenaCodigoHash = RscadenaCodigoHash!codigo
                                         End If

                                     End If
                                 End If
                                 
                            ElseIf lFESpring Then
                                If xlFacturacionE Then
                                    If tCodigoFE = "000" Then
                                         If lQRFE Then
                                             Set imageHash.Picture = LoadPicture(ImagenFeSpring(lQRFE, sDocumento))
                                         Else
                                             If lImpresionCodigoBarras Then
                                                 'Set imageHash.Picture = LoadPicture(ImagenQR(sDocumento))
                                             Else
                                                 'cadenaCodigoHash = Calcular("select isnull(tRespFacturacion,'') as codigo from mdocumento where tdocumento='" & sDocumento & "'", Cn)
                                                 cadenaCodigoHash = ImagenFeSpring(lQRFE, sDocumento)
                                             End If
                                         End If
                                     End If
                                End If
                                
                            ElseIf lFECarbajal Then
                                If xlFacturacionE Then
                                    If tCodigoFE = "000" Then
                                         If lQRFE Then
                                             Set imageHash.Picture = LoadPicture(ImagenFeCarvajal(3, sDocumento, 0))
                                         Else
                                             If lImpresionCodigoBarras Then
                                                 Set imageHash.Picture = LoadPicture(ImagenFeCarvajal(1, sDocumento, 0))
                                             Else
                                                 cadenaCodigoHash = ImagenFeCarvajal(2, sDocumento, 0)
                                             End If
                                         End If
                                     End If
                                End If
                            ElseIf lFEpape Then
                                If xlFacturacionE Then
                                    If tCodigoFE = "000" Then
                                         If lQRFE Then
                                             Set imageHash.Picture = LoadPicture(CrearImagenQR(TCPQR(Calcular("select isnull(tImprTermica,'') as codigo from mdocumento where tdocumento='" & sDocumento & "'", Cn))))
                                         Else
                                             If lImpresionCodigoBarras Then
                                                 'Set imageHash.Picture = LoadPicture(ImagenQR(sDocumento))
                                             Else
                                                 cadenaCodigoHash = Calcular("select isnull(tRespFacturacion,'') as codigo from mdocumento where tdocumento='" & sDocumento & "'", Cn)
                                             End If
                                         End If
                                     End If
                                End If
                            Else 'INFOFACT

                                 If xlFacturacionE Then 'DOC ELECTRONICO INFOFACT
                                    fDocumento = Mid(sDocumento, 1, 1) + Mid(sDocumento, 4, 3) + Mid(sDocumento, 8, 8)
                                    If lQRFE Then
                                        Set imageHash.Picture = LoadPicture(ImagenQR(sDocumento))
                                    Else
                                        If lImpresionCodigoBarras Then
                                            Set imageHash.Picture = LoadPicture(lValidaCodBarra(lImpresionCodigoBarras, sDocumento))
                                        Else
                                            cadenaCodigoHash = lValidaCodBarra(lImpresionCodigoBarras, sDocumento)
                                        End If
                                    End If

                                 End If
                                 
                            End If

                      End If
                  End If
                  
                  '------------------------------------------
        
            
               If RsImpresion.RecordCount = 0 Then
                  LimpiaRs
                  MsgBox "No existen Datos a Imprimir", vbExclamation, sMensaje
               Else
                  'Configura la Impresora
                  Imprimir (RsEmision!timpresora)
                  sResumen = RsEmision!lResumen
                  Printer.FontName = sFont
                  Printer.FontBold = False
                                                                               
                  'JESUS 06/03/2018
                  lPagocortesiaAI = IIf(RsImpresion!tCortesia = "", False, True)
        
                  If wConsumo Then
                        If RsEmision!tFormulario = "01" Then
                           If RsEmision!Cliente And RsEmision!Monto = 0 Then
                              
                                'FACTURACION ELECTRONICA
                                If xlFacturacionE And tCodigoFE <> "999" Then
                                     ImprimeFacturaConsumoElectronico RsImpresion, sDescrip, imageHash, Calcular("select ttipodocumento as codigo from mdocumento where tdocumento='" & sDocumento & "'", Cn), imageCab, imagepIE, cadenaCodigoHash, TimpresionDolaresDelivery
                                Else
                                
                                     ImprimeFacturaConsumoT RsImpresion, sDescrip, Calcular("select ttipodocumento as codigo from mdocumento where tdocumento='" & sDocumento & "'", Cn), imageCab, imagepIE, TimpresionDolaresDelivery
                                End If
                           
                           Else
                           
                                'FACTURACION ELECTRONICA
                                If xlFacturacionE And tCodigoFE <> "999" Then
                                     ImprimeBoletaConsumoElectronico RsImpresion, sDescrip, imageHash, Calcular("select ttipodocumento as codigo from mdocumento where tdocumento='" & sDocumento & "'", Cn), imageCab, imagepIE, cadenaCodigoHash, TimpresionDolaresDelivery
                                Else
                                     ImprimeBoletaConsumoT RsImpresion, sDescrip, Calcular("select ttipodocumento as codigo from mdocumento where tdocumento='" & sDocumento & "'", Cn), imageCab, imagepIE, TimpresionDolaresDelivery
                                End If
                           
                           End If
                        Else
                           If lFacturacionE And xlFacturacionE And lFEOfisis = False And lFESpring = False And lFEpape = False And lFECarbajal = False Then
                           
                                    If Generar_Imagen(CnFE, "select imagen from IMAGENCODIGOBARRA where nro_efact='" & fDocumento & "'", "imagen", "\fact.bmp") = True Then
                                        ImprimeFormatoAConsumo
                                        Kill App.Path & "\fact.bmp"
                                    Else
                                        ImprimeFormatoAConsumo
                                    End If
                            ElseIf lFacturacionE And lFECarbajal Then
                                ImprimeFormatoAConsumo
                                Set fso1 = CreateObject("Scripting.FileSystemObject")
                                If fso1.FileExists(App.Path & "\fact.bmp") Then
                                    Kill App.Path & "\fact.bmp"
                                End If
                           ElseIf lFacturacionE = False Then
                                    If RsEmision!Cliente And RsEmision!Monto = 0 Then
                                       ImprimeFacturaConsumoN RsImpresion, sDescrip, sEmpresa
                                    Else
                                       ImprimeBoletaConsumoN RsImpresion, sDescrip, sEmpresa
                                    End If
                           End If
                           
                        End If
                  Else
                        If RsEmision!tFormulario = "01" Then
                           If RsEmision!Cliente And RsEmision!Monto = 0 Then
                              
                                'FACTURACION ELECTRONICA
                                If xlFacturacionE And tCodigoFE <> "999" Then
                                   ImprimeFacturaElectronica RsImpresion, imageHash, Calcular("select ttipodocumento as codigo from mdocumento where tdocumento='" & sDocumento & "'", Cn), imageCab, imagepIE, cadenaCodigoHash, TimpresionDolaresDelivery
                                Else
                                   ImprimeFacturaT RsImpresion, Calcular("select ttipodocumento as codigo from mdocumento where tdocumento='" & sDocumento & "'", Cn), imageCab, imagepIE, TimpresionDolaresDelivery
                                End If
                           
                           ElseIf RsEmision!TTipoEmision = "00" Then
                           
                                If MsgBox("Deseas imprimir el Voucher", vbQuestion + vbYesNo, sMensaje) = vbYes Then
                                   ImprimeCortesia RsImpresion, Calcular("select ttipodocumento as codigo from mdocumento where tdocumento='" & sDocumento & "'", Cn), imageCab, imagepIE
                                End If
                           Else
                           
                                'FACTURACION ELECTRONICA
                                If xlFacturacionE And tCodigoFE <> "999" Then
                                   ImprimeBoletaElectronica RsImpresion, imageHash, Calcular("select ttipodocumento as codigo from mdocumento where tdocumento='" & sDocumento & "'", Cn), imageCab, imagepIE, cadenaCodigoHash, TimpresionDolaresDelivery
                                Else
                                   ImprimeBoletaT RsImpresion, Calcular("select ttipodocumento as codigo from mdocumento where tdocumento='" & sDocumento & "'", Cn), imageCab, imagepIE, TimpresionDolaresDelivery
                                End If
                           
                           End If
                        Else
                        
                                If lFacturacionE And xlFacturacionE And lFEOfisis = False And lFESpring = False And lFEpape = False And lFECarbajal = False Then
                                
                                        If Generar_Imagen(CnFE, "select imagen from IMAGENCODIGOBARRA where nro_efact='" & fDocumento & "'", "imagen", "\fact.bmp") = True Then
                                            ImprimeFormatoA
                                            Kill App.Path & "\fact.bmp"
                                        Else
                                            ImprimeFormatoA
                                        End If
                                ElseIf lFacturacionE And lFECarbajal Then
                                        ImprimeFormatoA
                                        Set fso1 = CreateObject("Scripting.FileSystemObject")
                                        If fso1.FileExists(App.Path & "\fact.bmp") Then
                                            Kill App.Path & "\fact.bmp"
                                        End If
                                 ElseIf lFacturacionE = False Then
                                
                                        If RsEmision!Cliente And RsEmision!Monto = 0 Then
                                           ImprimeFacturaN RsImpresion, sEmpresa, xlTipoDoc
                                        ElseIf RsEmision!TTipoEmision = "00" Then
                                           If MsgBox("Deseas imprimir el Voucher", vbQuestion + vbYesNo, sMensaje) = vbYes Then
                                              ImprimeCortesia RsImpresion, Calcular("select ttipodocumento as codigo from mdocumento where tdocumento='" & sDocumento & "'", Cn), imageCab, imagepIE
                                           End If
                                        Else
                                           ImprimeBoletaN RsImpresion, sEmpresa, xlTipoDoc
                                        End If
                                    
                                 End If
                                
                        End If
                        
                  End If

               End If
               lReimpresion = True
               LimpiaRs
               Screen.MousePointer = vbDefault
          
       Case Is = 8 ' Cuentas por Cobrar
               Dim antesvar As String
               If RsDocumento.RecordCount <> 0 Then
                  sDocumento = RsDocumento!tDocumento
               Else
                  Exit Sub
               End If
               antesvar = sVar1
               If lPasswordPorCobrar = True Then
                    If Supervisor("08") = False Then
                       MsgBox "Clave no permitida", vbExclamation, sMensaje
                       Exit Sub
                    End If
               Else
                    sVar1 = sUsuario
               End If
               
               sTemp = ""
               Isql = "select * from vCompania where lActivo = 1 order by Descripcion "
               Call ConfGrilla(5, frmBusquedaRapida.grdGrilla, "C�digo", 2, "Codigo", 900, 2, 0, "", _
                                                               "Identidad", 2, "Identidad", 1700, 0, 0, "", _
                                                               "Cliente", 2, "Descripcion", 3400, 0, 0, "", _
                                                               "Linea", 2, "nLineaPorCobrar", 1100, 1, 0, "##,##0.00", _
                                                               "Consumo", 2, "nConsumoPorCobrar", 1100, 1, 0, "##,##0.00")
               frmBusquedaRapida.nPredeterm = 1
               frmBusquedaRapida.Show vbModal
               
               If Not wEnter Then   'And RsGrilla.RecordCount <> 0 Then
                  Exit Sub
               End If
               
               'validadcuentacobrar'
               nTotalPorCobrarCliente = 0
               nLineaPorCobrarCliente = 0
               nLineaPorCobrarCliente = Calcular("select isnull(nlineaporcobrar,0) as codigo from vdelivery where codigo='" & sCodigo & "'", Cn)
               If nLineaPorCobrarCliente > 0 Then
                   nTotalPorCobrarCliente = Calcular("select isnull(nconsumoporcobrar,0) as codigo from vdelivery where codigo='" & sCodigo & "'", Cn)
                   Dim nVentaDocActual As Double
                   nVentaDocActual = Calcular("select isnull(nventa,0) as codigo from mdocumento where tdocumento='" & sDocumento & "' ", Cn)
                   If nTotalPorCobrarCliente + nVentaDocActual > nLineaPorCobrarCliente Then
                            MsgBox "Con esta asignaci�n el cliente: " & sDescrip & " supera su linea aprobada para Cuentas por Cobrar " & vbCrLf & " Linea asignada :  " & nLineaPorCobrarCliente & ". Utilizado :  " & nTotalPorCobrarCliente, vbInformation
                            Exit Sub
                    End If
               End If
               
               
               'validadcuentacobrar'
               If MsgBox("Esta seguro de Enviar el documento Nro: " & Format(sDocumento, "@-@@@@@-@@@@@@@@@") & _
                  Chr(13) & "a Cuentas por Cobrar por el Cliente " & sDescrip & "?", vbQuestion + vbOKCancel, sMensaje) = vbCancel Then
                  Exit Sub
               End If

               Cn.Execute "Update MDOCUMENTO set tEstadoDocumento = '03', tClientePago = '" & sCodigo & "', tUsuarioAutoriza='" & sVar1 & "',lreplica=1  where tDocumento ='" & sDocumento & "'"
               
               'validadcuentacobrar'
               Cn.Execute "update tdelivery set nConsumoPorCobrar= isnull(nconsumoporcobrar,0) + " & nVentaDocActual & " where tcodigodelivery='" & sCodigo & "' "
               'validadcuentacobrar'
               
               RsDocumento.Requery
               
             
               'Liberacion
               Dim i As Integer
               Set rst1 = Lib.OpenRecordset("select tCodigoPedido from DDOCUMENTO where tDocumento ='" & sDocumento & "'", Cn)
               Cn.Execute "Update DPEDIDO set tFacturado ='P' where tDocumento ='" & sDocumento & "'"
               If rst1.RecordCount <> 0 Then
                  rst1.MoveFirst
                  For i = 1 To rst1.RecordCount
                      Set RsT2 = Lib.OpenRecordset("select tCodigoPedido from DPEDIDO where tCodigoPedido ='" & rst1!tCodigoPedido & "' and (isnull(tFacturado,'0')='0' or tfacturado='F')", Cn)
                      If RsT2.RecordCount = 0 Then
                         Cn.Execute "Update MPEDIDO set tEstadoPedido ='02'  where tCodigoPedido ='" & rst1!tCodigoPedido & "'"
                         Cn.Execute "Update TMESA set tEstadoMesa = '04' where tCodigoMesa in (select tMesa from MPEDIDO where tCodigoPedido='" & rst1!tCodigoPedido & "') "
                      End If
                      rst1.MoveNext
                  Next i
               End If
               
                                '---- mesa247 --- avisa a mesa que fue enviada a cuentas por cobrar
            If MESA247 Then
                Cn.Execute "Update mpedido set lestadomesapedido = 3 where tcodigopedido ='" & rst1!tCodigoPedido & "' and idpedidomesa<>'' "
            End If
                
               sVar1 = antesvar
                
          Case Is = 9 ' Anulaci�n de Documentos
               If RsDocumento.RecordCount <> 0 Then
                  sDocumento = RsDocumento!tDocumento
                  xPedido = IIf(IsNull(RsDocumento!tCodigoPedido), "", RsDocumento!tCodigoPedido)
               Else
                  Exit Sub
               End If
                              
               If Calcular("select count(tdocumento) as Codigo from DPAGOTARJETA where tdocumento='" & sDocumento & "'", Cn) > 0 Then
                  MsgBox "Documento con Pago TC Pinpad, anular primero el pago", vbExclamation, sMensaje
                  Exit Sub
               End If
               
                If lFEBiz Then
                    Dim DiaDoc As Long
                    DiaDoc = Calcular("SELECT DATEDIFF(DAY,fRegistro,GETDATE()) as codigo FROM MDOCUMENTO where tdocumento='" & sDocumento & "'", Cn)
                    If DiaDoc > 7 Then
                        MsgBox "No se puede eliminar el documento,Tiempo maximo de eliminacion superada: 7 Dias, Emita una Nota de Credito", vbInformation, sMensaje
                        Exit Sub
                    End If
                End If
                
               sUsuarioAutoriza = sUsuario
               If Supervisor("05") = False Then
                  MsgBox "Clave no permitida", vbExclamation, sMensaje
                  Exit Sub
               End If
                sUsuarioAutoriza = sVar1
               
               If MsgBox("Esta seguro de Eliminar el documento Nro: " & Format(sDocumento, "@-@@@@@-@@@@@@@@@") & "?", vbQuestion + vbOKCancel, sMensaje) = vbCancel Then
                  Exit Sub
               End If
                
               wEnter = False
               frmKeyBoard.Caption = "Motivo Anulaci�n"

               frmKeyBoard.Show vbModal
               If Not wEnter Or sDescrip = "" Then
                  MsgBox "Debe ingresar el Motivo", vbExclamation, sMensaje
                  Exit Sub
               End If
               
               ' facturacion electronica
                Dim oComandoBaja As clsComando
                Set oComandoBaja = New clsComando
                '
               
               'FACTURACION_E_PERU
               If pais = "000" Then
                    If lFacturacionE Then
                    
                           If lFEOfisis Then 'OFISIS
                                 Dim lDocElecOfisis As Boolean
                                 lDocElecOfisis = Calcular("select isnull(tdi.lDocumentoElectronicoOfisis,0) as codigo from TTIPODOCUMENTOIMPRESORA tdi inner join MDOCUMENTO m on tdi.tTipoEmision = m.tTipoDocumento and tdi.tCaja = m.tCaja  where m.tDocumento= '" & sDocumento & "'", Cn)
                                 
                                 If lDocElecOfisis Then 'DOC ELECTRONICO OFISIS
                                        Dim xCDROfisis As String
                                        Dim RsDocumentoOfisis As Recordset
                                        Dim xContOfisis As Integer
                                        
                                        fDocumento = Mid(sDocumento, 1, 1) + Mid(sDocumento, 4, 3) + "-" + CStr(CLng(Mid(sDocumento, 8, 8)))
                                        Isql = "Select * From dbo.TCFACT_ELEC where NU_DOCU='" & fDocumento & "'"
                                        Set RsDocumentoOfisis = Lib.OpenRecordset(Isql, CnFE)
                                        
                                        If RsDocumentoOfisis.RecordCount > 0 Then
                                            CnFE.Execute "Update TCFACT_ELEC set CO_ESTA_DOCU = 'ANU' Where NU_DOCU = '" & fDocumento & "' and TI_DOCU <> 'D'"
                                        End If
                                        
                                        Set RsT3 = Lib.OpenRecordset("select distinct tCodigoPedido from DPEDIDO where tDocumento = '" & sDocumento & "'", Cn)
                                        Cn.Execute "Update MDOCUMENTO set tEstadoDocumento = '04', tUsuarioAnulado = '" & sUsuarioAutoriza & "', fRegistroAnulado = getdate() , tObservacion ='" & sDescrip & "', lreplica=1  where tDocumento ='" & sDocumento & "'"
                                        Cn.Execute "Update DPEDIDO set tDocumento = '', tFacturado = '' where tDocumento ='" & sDocumento & "' "
                                 Else
                                        Set RsT3 = Lib.OpenRecordset("select distinct tCodigoPedido from DPEDIDO where tDocumento = '" & sDocumento & "'", Cn)
                                        Cn.Execute "Update MDOCUMENTO set tEstadoDocumento = '04', tUsuarioAnulado = '" & sUsuarioAutoriza & "', fRegistroAnulado = getdate() , tObservacion ='" & sDescrip & "', lreplica=1  where tDocumento ='" & sDocumento & "'"
                                        Cn.Execute "Update DPEDIDO set tDocumento = '', tFacturado = '' where tDocumento ='" & sDocumento & "' "
                                 End If
                                                              
                           ElseIf lFEpape Then ' PAPERLEES
                                Dim DayEliminar, DayDoc As Long
                                DayEliminar = CLng(Trim(LeerIni(App.Path + "\INFOREST.INI", "TCPIP", "DAYANULACION", "10000")))
                                DayDoc = Calcular("SELECT DATEDIFF(DAY,fRegistro,GETDATE()) as codigo FROM MDOCUMENTO where tdocumento='" & sDocumento & "'", Cn)
                                If DayDoc > DayEliminar Then
                                    MsgBox "No se puede eliminar el documento,Tiempo maximo de eliminacion superada: " & DayEliminar & " Dias, Emita una Nota de Credito", vbInformation, sMensaje
                                    Exit Sub
                                End If
                                
                                Set RsT3 = Lib.OpenRecordset("select distinct tCodigoPedido from DPEDIDO where tDocumento = '" & sDocumento & "'", Cn)
                                Cn.Execute "Update MDOCUMENTO set tEstadoDocumento = '04', tUsuarioAnulado = '" & sUsuarioAutoriza & "', fRegistroAnulado = getdate() , tObservacion ='" & sDescrip & "', lreplica=1  where tDocumento ='" & sDocumento & "'"
                                Cn.Execute "Update DPEDIDO set tDocumento = '', tFacturado = '' where tDocumento ='" & sDocumento & "' "
                           
                           ElseIf lFESpring Then ' SPRING
                           
                                Set RsT3 = Lib.OpenRecordset("select distinct tCodigoPedido from DPEDIDO where tDocumento = '" & sDocumento & "'", Cn)
                                Cn.Execute "Update MDOCUMENTO set tEstadoDocumento = '04', tUsuarioAnulado = '" & sUsuarioAutoriza & "', fRegistroAnulado = getdate() , tObservacion ='" & sDescrip & "', lreplica=1  where tDocumento ='" & sDocumento & "'"
                                Cn.Execute "Update DPEDIDO set tDocumento = '', tFacturado = '' where tDocumento ='" & sDocumento & "' "
                                
                           ElseIf lFECarbajal Then
                           
                                Set RsT3 = Lib.OpenRecordset("select distinct tCodigoPedido from DPEDIDO where tDocumento = '" & sDocumento & "'", Cn)
                                Cn.Execute "Update MDOCUMENTO set tObservacion ='" & sDescrip & "' where tDocumento ='" & sDocumento & "'"
                                Dim lDocElec As Boolean
                                Dim sImporteLetra As String
                                lDocElec = Calcular("select isnull(tdi.lFacturacionElectronica,0) as codigo from TTIPODOCUMENTOIMPRESORA tdi inner join MDOCUMENTO m on tdi.tTipoEmision = m.tTipoDocumento and tdi.tCaja = m.tCaja  where m.tDocumento= '" & sDocumento & "'", Cn)
                                If lDocElec Then 'DOC ELECTRONICO INFOFACT
                                    sImporteLetra = NumeroCadena(str(Calcular("select isnull(nVenta,0) as Codigo from mDocumento where tDocumento='" & sDocumento & "'", Cn))) + " " + sMonedaN
                                    If Not INSERTAFE_CARVAJAL(sDocumento, sImporteLetra, 0, 1) Then '----CABECERA
                                        Cn.Execute "Update MDOCUMENTO set tObservacion ='' where tDocumento ='" & sDocumento & "'"
                                        Exit Sub
                                    End If
                                End If
                                Cn.Execute "Update MDOCUMENTO set tEstadoDocumento = '04', tUsuarioAnulado = '" & sUsuarioAutoriza & "', fRegistroAnulado = getdate() , tObservacion ='" & sDescrip & "', lreplica=1  where tDocumento ='" & sDocumento & "'"
                                Cn.Execute "Update DPEDIDO set tDocumento = '', tFacturado = '' where tDocumento ='" & sDocumento & "' "
                           ElseIf lFEBiz Then
                                    Set RsT3 = Lib.OpenRecordset("select distinct tCodigoPedido from DPEDIDO where tDocumento = '" & sDocumento & "'", Cn)
                                    Cn.Execute "Update MDOCUMENTO set tEstadoDocumento = '04', tUsuarioAnulado = '" & sUsuarioAutoriza & "', fRegistroAnulado = getdate() , tObservacion ='" & sDescrip & "', lreplica=1  where tDocumento ='" & sDocumento & "'"
                                    Cn.Execute "Update DPEDIDO set tDocumento = '', tFacturado = '' where tDocumento ='" & sDocumento & "' "
                           ElseIf lFEGesa Then
                                
                               If Not oComandoBaja.CreateCmdSp("USP_FactDocumentoBaja", Cn) Then
                                     Set oComandoBaja = Nothing
                                     'Exit Sub
                                End If
                                oComandoBaja.CreateParameter "@CodigoDocumento", adVarChar, adParamInput, 20, sDocumento
                
                                If Not oComandoBaja.GetParamOK Then
                                     Set oComandoBaja = Nothing
                                     'Exit Sub
                                End If
                                If Not oComandoBaja.ExecSP Then
                                     Set oComandoBaja = Nothing
                                     'Exit Sub
                                End If
                                
                                ' ANULACION
                                Set RsT3 = Lib.OpenRecordset("select distinct tCodigoPedido from DPEDIDO where tDocumento = '" & sDocumento & "'", Cn)
                                Cn.Execute "Update MDOCUMENTO set tEstadoDocumento = '04', tUsuarioAnulado = '" & sUsuarioAutoriza & "', fRegistroAnulado = getdate() , tObservacion ='" & sDescrip & "', lreplica=1  where tDocumento ='" & sDocumento & "'"
                                Cn.Execute "Update DPEDIDO set tDocumento = '', tFacturado = '' where tDocumento ='" & sDocumento & "' "
                           
                           Else 'INFOFACT
                           
                                 Dim lDocElecInfofact As Boolean
                                 lDocElecInfofact = Calcular("select isnull(tdi.lFacturacionElectronica,0) as codigo from TTIPODOCUMENTOIMPRESORA tdi inner join MDOCUMENTO m on tdi.tTipoEmision = m.tTipoDocumento and tdi.tCaja = m.tCaja  where m.tDocumento= '" & sDocumento & "'", Cn)
                                 
                                 If lDocElecInfofact Then 'DOC ELECTRONICO INFOFACT
                                        Dim xCDR As String
                                        Dim RsDocumentoVenta As Recordset
                                        Dim xCont As Integer
                                        
                                        fDocumento = Mid(sDocumento, 1, 1) + Mid(sDocumento, 4, 3) + Mid(sDocumento, 8, 8)
                                        Isql = "Select * From dbo.DOCUMENTOVENTA where nro_efact='" & fDocumento & "'"
                                        Set RsDocumentoVenta = Lib.OpenRecordset(Isql, CnFE)
                                        
                                        If RsDocumentoVenta.RecordCount > 0 Then
                                            
                                                 xCDR = IIf(IsNull(RsDocumentoVenta!cdr), "", RsDocumentoVenta!cdr)
                                            
                                                 'Dim oComandoBaja As clsComando
                                                 'Set oComandoBaja = New clsComando
                                                 
                                                 If Mid(sDocumento, 1, 1) = "F" Then
                                                        'ENVIO DOCUMENTO DE BAJA
'                                                        If xCDR = "" Then
'                                                                MsgBox "El Documento no esta declarado", vbExclamation, sMensaje
'                                                                Exit Sub
'                                                        Else
                                                                'If xCDR = "0" Or xCDR > "3999" Or DateDiff("d", RsDocumentoVenta!fRegistro, Now) < 8 Then
                                                                 
                                                                    'ENVIO DOCUMENTO DE BAJA
                                                                    If Not oComandoBaja.CreateCmdSp("USP_FactDocumentoBaja", Cn) Then
                                                                         Set oComandoBaja = Nothing
                                                                         Exit Sub
                                                                    End If
                                                                    oComandoBaja.CreateParameter "@CodigoDocumento", adVarChar, adParamInput, 20, sDocumento
                                                    
                                                                    If Not oComandoBaja.GetParamOK Then
                                                                         Set oComandoBaja = Nothing
                                                                         Exit Sub
                                                                    End If
                                                                    If Not oComandoBaja.ExecSP Then
                                                                         Set oComandoBaja = Nothing
                                                                         Exit Sub
                                                                    End If
                                                                    
                                                                    'ANULACION
                                                                    Set RsT3 = Lib.OpenRecordset("select distinct tCodigoPedido from DPEDIDO where tDocumento = '" & sDocumento & "'", Cn)
                                                                    Cn.Execute "Update MDOCUMENTO set tEstadoDocumento = '04', tUsuarioAnulado = '" & sUsuarioAutoriza & "', fRegistroAnulado = getdate() , tObservacion ='" & sDescrip & "', lreplica=1  where tDocumento ='" & sDocumento & "'"
                                                                    Cn.Execute "Update DPEDIDO set tDocumento = '', tFacturado = '' where tDocumento ='" & sDocumento & "' "
                                                                'Else
                                                                '     MsgBox "Documento no puede ser Anulado", vbExclamation, sMensaje
                                                                '     Exit Sub
                                                                'End If
                                                        'End If
                                                        '----------------------
                                                 Else
                                                        'If DateDiff("d", RsDocumentoVenta!fRegistro, Now) < 8 Then
                            
                                                             'ENVIO DOCUMENTO DE BAJA
                                                             If Not oComandoBaja.CreateCmdSp("USP_FactDocumentoBaja", Cn) Then
                                                                  Set oComandoBaja = Nothing
                                                                  Exit Sub
                                                             End If
                                                             oComandoBaja.CreateParameter "@CodigoDocumento", adVarChar, adParamInput, 20, sDocumento
                                             
                                                             If Not oComandoBaja.GetParamOK Then
                                                                  Set oComandoBaja = Nothing
                                                                  Exit Sub
                                                             End If
                                                             If Not oComandoBaja.ExecSP Then
                                                                  Set oComandoBaja = Nothing
                                                                  Exit Sub
                                                             End If
                                                             
                                                             ' ANULACION
                                                             Set RsT3 = Lib.OpenRecordset("select distinct tCodigoPedido from DPEDIDO where tDocumento = '" & sDocumento & "'", Cn)
                                                             Cn.Execute "Update MDOCUMENTO set tEstadoDocumento = '04', tUsuarioAnulado = '" & sUsuarioAutoriza & "', fRegistroAnulado = getdate() , tObservacion ='" & sDescrip & "', lreplica=1  where tDocumento ='" & sDocumento & "'"
                                                             Cn.Execute "Update DPEDIDO set tDocumento = '', tFacturado = '' where tDocumento ='" & sDocumento & "' "
                                                        'Else
                                                        '     MsgBox "Documento no puede ser Anulado, se supero el limite de dias desde su emisi�n", vbExclamation, sMensaje
                                                        '     Exit Sub
                                                        'End If
                                                        '----------------------
                                                 End If
                                                  
                                        Else
                                            Set RsT3 = Lib.OpenRecordset("select distinct tCodigoPedido from DPEDIDO where tDocumento = '" & sDocumento & "'", Cn)
                                            Cn.Execute "Update MDOCUMENTO set tEstadoDocumento = '04', tUsuarioAnulado = '" & sUsuarioAutoriza & "', fRegistroAnulado = getdate() , tObservacion ='" & sDescrip & "', lreplica=1  where tDocumento ='" & sDocumento & "'"
                                            Cn.Execute "Update DPEDIDO set tDocumento = '', tFacturado = '' where tDocumento ='" & sDocumento & "' "
                                        End If
                                    
 
                                 Else   'NO ELECTRONICO
                                        Set RsT3 = Lib.OpenRecordset("select distinct tCodigoPedido from DPEDIDO where tDocumento = '" & sDocumento & "'", Cn)
                                        Cn.Execute "Update MDOCUMENTO set tEstadoDocumento = '04', tUsuarioAnulado = '" & sUsuarioAutoriza & "', fRegistroAnulado = getdate() , tObservacion ='" & sDescrip & "', lreplica=1  where tDocumento ='" & sDocumento & "'"
                                        Cn.Execute "Update DPEDIDO set tDocumento = '', tFacturado = '' where tDocumento ='" & sDocumento & "' "
                                 End If
                                 
                           
                           End If
                            
                     Else
                            Set RsT3 = Lib.OpenRecordset("select distinct tCodigoPedido from DPEDIDO where tDocumento = '" & sDocumento & "'", Cn)
                            Cn.Execute "Update MDOCUMENTO set tEstadoDocumento = '04', tUsuarioAnulado = '" & sUsuarioAutoriza & "', fRegistroAnulado = getdate() , tObservacion ='" & sDescrip & "', lreplica=1  where tDocumento ='" & sDocumento & "'"
                            Cn.Execute "Update DPEDIDO set tDocumento = '', tFacturado = '' where tDocumento ='" & sDocumento & "' "
                     End If
                     
               Else
                    Set RsT3 = Lib.OpenRecordset("select distinct tCodigoPedido from DPEDIDO where tDocumento = '" & sDocumento & "'", Cn)
                    Cn.Execute "Update MDOCUMENTO set tEstadoDocumento = '04', tUsuarioAnulado = '" & sUsuarioAutoriza & "', fRegistroAnulado = getdate() , tObservacion ='" & sDescrip & "', lreplica=1  where tDocumento ='" & sDocumento & "'"
                    Cn.Execute "Update DPEDIDO set tDocumento = '', tFacturado = '' where tDocumento ='" & sDocumento & "' "
               End If
               '-----------------

            If Calcular("select count(*) as codigo from mguiatransporte where tdocumento='" & sDocumento & "'", Cn) > 0 Then
                Cn.Execute "Update mguiatransporte set tEstadoGuia = '01',tdocumento='' where tdocumento ='" & sDocumento & "'"
                If RsT3.RecordCount <> 0 Then
                  RsT3.MoveFirst
                  Do While Not RsT3.EOF
                    Cn.Execute "Update MPEDIDO set tEstadoPedido = '04', lReplica=1  where tCodigoPedido ='" & RsT3!tCodigoPedido & "'"
                  Loop
                End If
            Else
               Dim xEM As String
               Dim xCM As String
               If RsT3.RecordCount <> 0 Then
                  RsT3.MoveFirst
                  Do While Not RsT3.EOF
                     
                     xCM = Calcular("select tMesa as Codigo from MPEDIDO where tCodigoPedido='" & RsT3!tCodigoPedido & "'", Cn)
                     xEM = Calcular("select tEstadoMesa as Codigo from TMESA where tCodigoMesa='" & xCM & "'", Cn)
                     
                     If xEM = "01" Or xEM = "04" Or xEM = "05" Then
                        Cn.Execute "Update MPEDIDO set tEstadoPedido = '01', lReplica=1  where tCodigoPedido ='" & RsT3!tCodigoPedido & "'"
                        Cn.Execute "update TMESA set tEstadoMesa='02' where tCodigoMesa='" & xCM & "'"
                     Else
                        Cn.Execute "Update MPEDIDO set tEstadoPedido = '01', tmesa = '', lReplica=1  where tCodigoPedido ='" & RsT3!tCodigoPedido & "'"
                     End If
                     RsT3.MoveNext
                  Loop
               End If
             End If
             
            If lInfhotel Then
               Dim sComandaInfhotel As String
               Dim sPuntoVenta As String
               sComandaInfhotel = Calcular("select tComanda as Codigo From MPEDIDO where tCodigoPedido='" & xPedido & "'", Cn)
               sPuntoVenta = Calcular("select tPuntoVenta as Codigo From MPEDIDO where tCodigoPedido='" & xPedido & "'", Cn)
               CnInfhotel.Execute "update MCOMANDA set TESTADO='04', tDocumento='' where tComanda ='" & sComandaInfhotel & "' and tPuntoVenta='" & sPuntoVenta & "'"
               CnInfhotel.Execute "delete from DCOMANDA where tcomanda='" & sComandaInfhotel & "' and tcodigoitem='100000' and tPuntoVenta='" & sPuntoVenta & "'"
               CnInfhotel.Execute "delete from WMCOMANDA where tComanda ='" & sComandaInfhotel & "' and tPuntoVenta='" & sPuntoVenta & "'"
               CnInfhotel.Execute "delete from WDCOMANDA where tComanda ='" & sComandaInfhotel & "' and tPuntoVenta='" & sPuntoVenta & "'"
            End If
            RsDocumento.Requery
            
          Case Is = 10  ' Precuenta Delivery
            If RsDocumento.RecordCount = 0 Then
               Exit Sub
            End If
            sPedido = grdDocumento.Columns(4).value
            Dim nLP As Integer
            nLP = Calcular("select nPrecuenta as Codigo FROM MPEDIDO where tCodigoPedido='" & sPedido & "'", Cn)
            If nLimitePrecuenta <> 0 And nLimitePrecuenta <= nLP Then
               If Supervisor("14") = False Then
                  MsgBox "Clave no permitida", vbExclamation, sMensaje
                  Exit Sub
               End If
            End If
            
            Screen.MousePointer = vbHourglass
                        
            Cn.Execute "update MPEDIDO set nPrecuenta = isnull(nPrecuenta,0) + 1  where tCodigoPedido='" & sPedido & "'"
            'Propiedades
            Dim xSql As String
            If lAlmacen Then
               Dim RsOp As Recordset
               Set RsOp = Lib.OpenRecordset("select Codigo, Descripcion from vOperador where lStockMenos=1", Cn)
               If RsOp.RecordCount > 0 Then
                  xSql = "select tCodigoPropiedad as Codigo, TPROPIEDAD.tDetallado as Descripcion, tProducto, TOPERADOR.tOperador as tOperador, TOPERADOR.tDetallado as Operador, nPrecio, tEnlace, " & _
                         "nInsumo, nGasto, nManoObra " & _
                         "FROM dbo.TPROPIEDAD LEFT OUTER JOIN dbo.TOPERADOR ON dbo.TPROPIEDAD.tOperador = dbo.TOPERADOR.tOperador " & _
                         "Where TPROPIEDAD.lActivo = 1 And IsNull(TOPERADOR.lStockMenos, 0) <> 1 union " & _
                         "select '9999' as Codigo, tDetallado as Descripcion, tCodigoPlato as tProducto, '" & RsOp!codigo & "' as tOperador, '" & RsOp!Descripcion & "' as Operador, 0, " & sAlmacenMDB & ".dbo.DRECETAVENTA.tCodigoProducto as tEnlace, nCantidad * nPrecio as nInsumo, 0, 0 " & _
                         "FROM " & sAlmacenMDB & ".dbo.DRECETAVENTA INNER JOIN " & sAlmacenMDB & ".dbo.MRECETAVENTA ON " & sAlmacenMDB & ".dbo.DRECETAVENTA.tLocal = " & sAlmacenMDB & ".dbo.MRECETAVENTA.tLocal AND " & sAlmacenMDB & ".dbo.DRECETAVENTA.tRecetaVenta = " & sAlmacenMDB & ".dbo.MRECETAVENTA.tRecetaVenta INNER JOIN " & sAlmacenMDB & ".dbo.TPRODUCTO ON " & sAlmacenMDB & ".dbo.DRECETAVENTA.tCodigoProducto = " & sAlmacenMDB & ".dbo.TPRODUCTO.tCodigoProducto " & _
                         "Where lNoDescargo = 1"
               Else
                  xSql = "select tCodigoPropiedad as Codigo, TPROPIEDAD.tDetallado as Descripcion, tProducto, TPROPIEDAD.tOperador, nPrecio, tEnlace, " & _
                         "nInsumo, nGasto, nManoObra, toperador.tDetallado AS Operador " & _
                         "FROM dbo.TPROPIEDAD LEFT OUTER JOIN dbo.TOPERADOR ON dbo.TPROPIEDAD.tOperador = dbo.TOPERADOR.tOperador " & _
                         "Where TPROPIEDAD.lActivo = 1 And IsNull(TOPERADOR.lStockMenos, 0) <> 1"
               End If
            Else
                  xSql = "select tCodigoPropiedad as Codigo, TPROPIEDAD.tDetallado as Descripcion, tProducto, TPROPIEDAD.tOperador, nPrecio, tEnlace, " & _
                         "nInsumo, nGasto, nManoObra, toperador.tDetallado AS Operador " & _
                         "FROM dbo.TPROPIEDAD LEFT OUTER JOIN dbo.TOPERADOR ON dbo.TPROPIEDAD.tOperador = dbo.TOPERADOR.tOperador " & _
                         "Where TPROPIEDAD.lActivo = 1 And IsNull(TOPERADOR.lStockMenos, 0) <> 1"
            End If
            
            'Producto Propiedad
            Isql = "SELECT dbo.TPRODUCTOPROPIEDAD.tCodigoPedido, dbo.TPRODUCTOPROPIEDAD.tItem, dbo.TPRODUCTOPROPIEDAD.tCodigoPropiedad, dbo.TPRODUCTOPROPIEDAD.tProducto, t1.Descripcion AS Descripcion, Operador " & _
                   "FROM dbo.TPRODUCTOPROPIEDAD INNER JOIN (" & xSql & ") T1 ON dbo.TPRODUCTOPROPIEDAD.tCodigoPropiedad = T1.Codigo and dbo.TPRODUCTOPROPIEDAD.tProducto = T1.tProducto and dbo.TPRODUCTOPROPIEDAD.tenlace= t1.tenlace " & _
                   "where tCodigoPedido='" & sPedido & "'"
            Set RsProductoPropiedad = Lib.OpenRecordset(Isql, Cn)
               
            'Combo Propiedad
            Isql = "SELECT dbo.TCOMBOPROPIEDAD.tCodigoPedido, dbo.TCOMBOPROPIEDAD.tItem, dbo.TCOMBOPROPIEDAD.tItemCombo, T1.Descripcion, T1.Operador " & _
                   "FROM dbo.TCOMBOPROPIEDAD INNER JOIN (" & xSql & ") T1 ON dbo.TCOMBOPROPIEDAD.tCodigoPropiedad = T1.Codigo AND dbo.TCOMBOPROPIEDAD.tProducto = T1.tProducto AND dbo.TCOMBOPROPIEDAD.tEnlace = T1.tEnlace " & _
                   "where tCodigoPedido='" & sPedido & "'"
            Set RsComboPropiedad = Lib.OpenRecordset(Isql, Cn)
            
            
            Isql = "select * from vPreCuentaDelivery " & _
                   "WHERE Codigo='" & sPedido & "' ORDER BY tItem, tItemCombo"

            Set RsImpresion = Lib.OpenRecordset(Isql, Cn)
            If RsImpresion.RecordCount = 0 Then
               LimpiaRs
               MsgBox "No existen Datos a Imprimir", vbExclamation, sMensaje
            Else
               ImprimeDelivery RsImpresion, RsProductoPropiedad, RsComboPropiedad
            End If
            LimpiaRs
               
          Case Is = 11  ' Anulaci�n de Pago
               If Supervisor("05") = False Then
                  MsgBox "Clave no permitida", vbExclamation, sMensaje
                  Exit Sub
               End If
               
               Isql = "SELECT MAX(dbo.DDOCUMENTO.tCodigoPedido) AS Descripcion, dbo.MDOCUMENTO.tDocumento AS Codigo, MAX(dbo.MDOCUMENTO.fRegistro) AS fFecha, dbo.MDOCUMENTO.tUsuario, dbo.TCLIENTE.tEmpresa AS Cliente, dbo.MDOCUMENTO.nVenta, MAX(dbo.MPEDIDO.tObservacion) AS tObservacion, dbo.TMESA.tDetallado AS Mesa " & _
                      "FROM dbo.MPEDIDO LEFT OUTER JOIN dbo.TMESA ON dbo.MPEDIDO.tMesa = dbo.TMESA.tCodigoMesa RIGHT OUTER JOIN dbo.DDOCUMENTO ON dbo.MPEDIDO.tCodigoPedido = dbo.DDOCUMENTO.tCodigoPedido RIGHT OUTER JOIN dbo.MDOCUMENTO LEFT OUTER JOIN dbo.TCLIENTE ON dbo.MDOCUMENTO.tCodigoCliente = dbo.TCLIENTE.tCodigoCliente ON dbo.DDOCUMENTO.tDocumento = dbo.MDOCUMENTO.tDocumento " & _
                      "where dbo.MDOCUMENTO.tTurno='" & sTurno & "' and tEstadoDocumento = '02' and tTipodocumento<>'00' " & _
                      "GROUP BY dbo.MDOCUMENTO.tDocumento, dbo.TCLIENTE.tEmpresa, dbo.MDOCUMENTO.nVenta, dbo.MDOCUMENTO.tUsuario, dbo.TMESA.tDetallado ORDER BY dbo.MDOCUMENTO.tDocumento "
                                             
               Call ConfGrilla(6, frmBusquedaRapida.grdGrilla, "Documento", 2, "Codigo", 1500, 0, 0, "", _
                                                               "Fec.Emis", 2, "fFecha", 1000, 0, 0, "", _
                                                               "Monto", 2, "nVenta", 1000, 1, 0, "###,###,##0.00", _
                                                               "Mesa", 2, "Mesa", 1000, 0, 0, "", _
                                                               "Observacion", 2, "tObservacion", 2000, 0, 0, "", _
                                                               "Cliente", 2, "Cliente", 1700, 0, 0, "")
               frmBusquedaRapida.nPredeterm = 0
               frmBusquedaRapida.Show vbModal
               
                If lFEBiz Then
                    Dim DiaDoc2 As Long
                    DiaDoc2 = Calcular("SELECT DATEDIFF(DAY,fRegistro,GETDATE()) as codigo FROM MDOCUMENTO where tdocumento='" & sCodigo & "'", Cn)
                    If DiaDoc2 > 7 Then
                        If MsgBox("El documento tiene mas 7 Dias , no se podra eliminar, Desea continuar con la Anulacion del Pago?", vbYesNo, sMensaje) = vbNo Then
                            Exit Sub
                        End If
                    End If
                End If
               
               
               If wEnter Then
                  sPedido = sDescrip
                  'Control de mesas
                  Dim RsEstadoMesa As Recordset
                  Set RsEstadoMesa = Lib.OpenRecordset("select tCodigoMesa, tDetallado, tEstadoMesa from TMESA WHERE tEstadoMesa in ('02','03','05','06') and  TCODIGOMESA IN(SELECT TMESA FROM MPEDIDO WHERE TCODIGOPEDIDO='" & sPedido & "')", Cn)
                  If RsEstadoMesa.RecordCount > 0 Then
                     If MsgBox("Pedido asociado a una mesa ocupada" & Chr(13) & "Deseas continuar con la recuperaci�n del documento?", vbExclamation + vbYesNo, sMensaje) = vbNo Then
                        Exit Sub
                     End If
                     Cn.Execute "Update TMESA set tEstadoMesa='05' where tCodigoMesa='" & RsEstadoMesa!tCodigoMesa & "'"
                  End If
                  Set RsEstadoMesa = Nothing
                  
                  'Pin Pad
                  Dim sReferencia As String
                  sReferencia = ""
                  If lVisaNet Then
                     sReferencia = Calcular("select tReferencia as Codigo from DPAGOTARJETA where tDocumento='" & sCodigo & "'", Cn)
                  End If
                  
                  'XX AGOSTO 2011 LG ACTIVACION DEL INGRESO SI ES PAGADO POR RECIBO
                  Dim tTipoPago As String
                  Dim tnumerorecibo As String
                  tTipoPago = Calcular("SELECT ISNULL(TOTROTIPOPAGO,'') AS CODIGO FROM DPAGODOCUMENTO WHERE TDOCUMENTO='" & sCodigo & "' and totrotipopago='001'", Cn)
                  If tTipoPago = "001" Then
                      tnumerorecibo = Calcular("SELECT ISNULL(TNUMERO,'') AS CODIGO FROM DPAGODOCUMENTO WHERE TDOCUMENTO ='" & sCodigo & "' and totrotipopago='001' ", Cn)
                      Cn.Execute "UPDATE MINGRESO SET TESTADODOCUMENTO='01' WHERE TRECIBO in ( SELECT ISNULL(TNUMERO,'') AS CODIGO FROM DPAGODOCUMENTO WHERE TDOCUMENTO ='" & sCodigo & "' and totrotipopago='001')"
                      'Cn.Execute "UPDATE MINGRESO SET TESTADODOCUMENTO='01' WHERE TRECIBO='" & tnumerorecibo & "'"
                  End If
                  'FIN XX
                  
                  '-- Recuperacion de la nota de credito cuando se Anula el pago del documento
                  Dim tnotacreditorecuperar As String
                  tnotacreditorecuperar = Calcular("select isnull(tNumero,'') as codigo from DPAGODOCUMENTO where  tDocumento in ('" & sCodigo & "') and totrotipopago='002'", Cn)
                  If tnotacreditorecuperar <> "0" Then
                  Cn.Execute "update MNOTACREDITO SET tEstadoDocumento='05' where tNotaCredito = '" & tnotacreditorecuperar & "'"
                  End If
                  '-- Fin Recuperacion de nota de credito
                  
                  If sReferencia <> "" Then
                     Cn.Execute "update DPAGOTARJETA set tEstadoDocumento='04' where tDocumento='" & sCodigo & "' and tReferencia='" & sReferencia & "'"
                     Cn.Execute "delete from DPAGODOCUMENTO where tDocumento='" & sCodigo & "' and tReferencia <> '" & sReferencia & "'"
                  Else
                     Cn.Execute "delete from DPAGODOCUMENTO where tDocumento='" & sCodigo & "'"
                  End If
                                                                        
                  Cn.Execute "update MDOCUMENTO set tClientePago='', tEstadoDocumento='01',lreplica=1,  tcortesia=''  where tDocumento='" & sCodigo & "'"
                  Cn.Execute "update DPEDIDO set tCortesia='', tFacturado='F' where tDocumento='" & sCodigo & "' and tEstadoItem='N'"
                                    
                  If lInfhotel Then
                     CnInfhotel.Execute "Update MCOMANDA set tEstado='02' where tNotaPedido='" & sPedido & "' and tDocumento ='" & IIf(pais = "002", Mid(sDocumento, 1, 1) + Mid(sDocumento, 3), sDocumento) & "'"
                  End If
                  RsDocumento.Requery
               End If
          
          Case Is = 12  ' cierre
            frmRepDeliveryTicket.Show vbModal
            
        ' cambios para anulacion con nota de credito
          Case Is = 13

            
             
            
            
            
        '-------------------------------------------------
               
   End Select
End Sub

Private Sub Form_Load()
   'Detalle
   sTemporal = dbTemporal(sCaja, 48, "tCodigoPedido", "nVarChar(10)", _
                                     "tItem", "nVarChar(3)", _
                                     "tTipoPedido", "nVarChar(2)", _
                                     "tCodigoProducto", "nVarChar(7)", _
                                     "tCodigoGrupo", "nVarChar(2)", _
                                     "tCodigoSubGrupo", "nVarChar(4)", _
                                     "tMoneda", "nVarChar(3)", _
                                     "nPrecioNeto", "Float", "nPrecioImpuesto1", "Float", "nPrecioImpuesto2", "Float", "nPrecioImpuesto3", "Float", "nPrecioVenta", "Float", _
                                     "nRecargo", "Float", "nDescuento", "Float", "nPrecioOficial", "Float", _
                                     "nCantidad", "Float", _
                                     "nImpuesto1", "Float", "nImpuesto2", "Float", "nImpuesto3", "Float", _
                                     "nVenta", "Float", _
                                     "tObservacion", "nVarChar(255)", _
                                     "tCortesia", "nVarChar(4)", _
                                     "lImprime", "Bit", _
                                     "tEstadoItem", "nVarChar(3)", _
                                     "tArea", "nVarChar(3)", "lCombinacion", "Bit", "nCombinacion", "Smallint", "lImprimeArea", "Bit", _
                                     "tFacturado", "nVarChar(1)", "tDocumento", "nVarChar(20)", "lTransferido", "Bit", "tComanda", "nVarChar(10)", "fRegistro", "smalldatetime", "tMozoD", "nVarChar(4)", "tUsuarioD", "nVarChar(15)", _
                                     "nInsumo", "Float", "nGasto", "Float", "nManoObra", "Float", "nOrden", "Int", "lCorte", "bit", "tPosicion", "nVarChar(2)", "fEnvio", "smalldatetime", "nEnvio", "int", "tUnidadNegocio", "nVarChar(2)", "tOferta", "nVarChar(5)", "tAutorizaOferta", "nVarChar(15)", "tsubalmacen", "nvarchar(3)", "Estado", "nVarChar(1)")
                                      
   'Impresoras
   Isql = "select * from vTipoDocumentoImpresora where tCaja ='" & sCaja & "' and Transporte=0"
   Set RsEmision = Lib.OpenRecordset(Isql, Cn)
   If RsEmision.RecordCount = 0 Then
      MsgBox "No se ha ingresado los Documentos por Caja", vbCritical, "Error Fatal"
      Unload Me
   End If
                                      
   
'   Cn.Execute "Insert into " & sTemporal & " select *, ' ' as Estado From DPEDIDO where tCodigoPedido='" & sPedido & "' and (isnull(tFacturado,'0')='0' or len(ltrim(tFacturado)) = 0) and tEstadoItem='N' and tCodigoPedido<>''"
    Cn.Execute "Insert into " & sTemporal & " select tCodigoPedido, tItem, tTipoPedido, tCodigoProducto, tCodigoGrupo, tCodigoSubGrupo,   tMoneda, nPrecioNeto,nPrecioImpuesto1, nPrecioImpuesto2,  nPrecioImpuesto3,  nPrecioVenta,    nRecargo,  nDescuento,  nPrecioOficial, nCantidad,nImpuesto1, nImpuesto2,  nImpuesto3,   nVenta, tObservacion, tCortesia, lImprime, tEstadoItem, tArea, lCombinacion,  nCombinacion, lImprimeArea,   tFacturado,  tDocumento, lTransferido,  tComanda, fRegistro, tMozoD,tUsuarioD, nInsumo, nGasto, nManoObra,  nOrden,lCorte,tPosicion, fEnvio,nEnvio, tUnidadNegocio,tOferta, tAutorizaOferta,  tsubalmacen, ' ' as Estado From DPEDIDO where tCodigoPedido='" & sPedido & "' and (isnull(tFacturado,'0')='0' or len(ltrim(tFacturado)) = 0) and tEstadoItem='N' and tCodigoPedido<>''"
    
   Isql = "SELECT [" & sTemporal & "].*, TPRODUCTO.tResumido AS Producto, [nVenta]*[nCantidad] AS Subtotal " & _
          "FROM TPRODUCTO INNER JOIN [" & sTemporal & "] ON TPRODUCTO.tCodigoProducto = [" & sTemporal & "].tCodigoProducto where Estado=' '"
          
   Set RsDetalle1 = Lib.OpenRecordset(Isql, Cn)
   
   Call ConfGrilla(5, grdDetalle1, "Producto", 2, "Producto", 1600, 0, 0, "", _
                                   "Precio", 2, "nPrecioVenta", 810, 1, 0, "###,###,##0.00", _
                                   "Cant", 2, "nCantidad", 500, 1, 0, "#,##0.00", _
                                   "SubTotal", 2, "nVenta", 830, 1, 0, "###,###,##0.00", _
                                   "Pedido", 2, "tCodigoPedido", 1000, 0, 0, "")
                                   
   Set grdDetalle1.DataSource = RsDetalle1
   If Len(Trim(frmVenta.grdCabecera.Columns(1).Text)) = 0 Then
      grdDetalle1.Caption = "Por Facturar Observaci�n : " & frmVenta.grdCabecera.Columns(2).Text
   Else
      grdDetalle1.Caption = "Por Facturar Mesa : " & frmVenta.grdCabecera.Columns(1).Text
   End If
   
   'anulacion por nota de credito
   Isql = "SELECT * FROM TPARAMETRO"
   Set RsTparametro = Lib.OpenRecordset(Isql, Cn)
   '------------------------------------------------------------------------------
   'Documento
   If lMCPV Then
    Isql = "SELECT dbo.MDOCUMENTO.tDocumento, dbo.MDOCUMENTO.tCodigoCliente, dbo.TCLIENTE.tEmpresa AS Cliente, dbo.MDOCUMENTO.tEstadoDocumento , dbo.MDOCUMENTO.nVenta, dbo.MDOCUMENTO.tTipoDocumento, MAX(dbo.MDOCUMENTO.fRegistro) AS fFecha, MAX(dbo.TMESA.tResumido) AS Mesa, MAX(dbo.MPEDIDO.tObservacion) AS tObservacion, MAX(dbo.DDOCUMENTO.tCodigoPedido) AS tCodigoPedido, Max(dbo.MPEDIDO.tTipoPedido) as tTipoPedido, dbo.vMotorizado.Descripcion AS Motorizado " & _
          "FROM dbo.MPEDIDO LEFT OUTER JOIN dbo.vMotorizado ON dbo.MPEDIDO.tMotorizado = dbo.vMotorizado.Codigo LEFT OUTER JOIN dbo.TMESA ON dbo.MPEDIDO.tMesa = dbo.TMESA.tCodigoMesa RIGHT OUTER JOIN dbo.DDOCUMENTO ON dbo.MPEDIDO.tCodigoPedido = dbo.DDOCUMENTO.tCodigoPedido RIGHT OUTER JOIN dbo.vCompania RIGHT OUTER JOIN dbo.MDOCUMENTO ON dbo.vCompania.Codigo = dbo.MDOCUMENTO.tClientePago LEFT OUTER JOIN dbo.TCLIENTE ON dbo.MDOCUMENTO.tCodigoCliente = dbo.TCLIENTE.tCodigoCliente ON dbo.DDOCUMENTO.tDocumento = dbo.MDOCUMENTO.tDocumento " & _
          "where tEstadoDocumento ='01' and dbo.MDOCUMENTO.tUsuario ='" & sUsuario & "' " & _
          "GROUP BY dbo.MDOCUMENTO.tDocumento, dbo.MDOCUMENTO.tCodigoCliente, dbo.TCLIENTE.tEmpresa, dbo.MDOCUMENTO.tEstadoDocumento, dbo.MDOCUMENTO.nVenta , dbo.MDOCUMENTO.tTipoDocumento, dbo.vMotorizado.Descripcion " & _
          "ORDER BY dbo.MDOCUMENTO.tDocumento"
    Else
    Isql = "SELECT dbo.MDOCUMENTO.tDocumento, dbo.MDOCUMENTO.tCodigoCliente, dbo.TCLIENTE.tEmpresa AS Cliente, dbo.MDOCUMENTO.tEstadoDocumento , dbo.MDOCUMENTO.nVenta, dbo.MDOCUMENTO.tTipoDocumento, MAX(dbo.MDOCUMENTO.fRegistro) AS fFecha, MAX(dbo.TMESA.tResumido) AS Mesa, MAX(dbo.MPEDIDO.tObservacion) AS tObservacion, MAX(dbo.DDOCUMENTO.tCodigoPedido) AS tCodigoPedido, Max(dbo.MPEDIDO.tTipoPedido) as tTipoPedido, dbo.vMotorizado.Descripcion AS Motorizado " & _
          "FROM dbo.MPEDIDO LEFT OUTER JOIN dbo.vMotorizado ON dbo.MPEDIDO.tMotorizado = dbo.vMotorizado.Codigo LEFT OUTER JOIN dbo.TMESA ON dbo.MPEDIDO.tMesa = dbo.TMESA.tCodigoMesa RIGHT OUTER JOIN dbo.DDOCUMENTO ON dbo.MPEDIDO.tCodigoPedido = dbo.DDOCUMENTO.tCodigoPedido RIGHT OUTER JOIN dbo.vCompania RIGHT OUTER JOIN dbo.MDOCUMENTO ON dbo.vCompania.Codigo = dbo.MDOCUMENTO.tClientePago LEFT OUTER JOIN dbo.TCLIENTE ON dbo.MDOCUMENTO.tCodigoCliente = dbo.TCLIENTE.tCodigoCliente ON dbo.DDOCUMENTO.tDocumento = dbo.MDOCUMENTO.tDocumento " & _
          "where tEstadoDocumento ='01' and dbo.MDOCUMENTO.tCaja ='" & sCaja & "' " & _
          "GROUP BY dbo.MDOCUMENTO.tDocumento, dbo.MDOCUMENTO.tCodigoCliente, dbo.TCLIENTE.tEmpresa, dbo.MDOCUMENTO.tEstadoDocumento, dbo.MDOCUMENTO.nVenta , dbo.MDOCUMENTO.tTipoDocumento, dbo.vMotorizado.Descripcion " & _
          "ORDER BY dbo.MDOCUMENTO.tDocumento"
    End If

   Set RsDocumento = Lib.OpenRecordset(Isql, Cn)
   Call ConfGrilla(8, grdDocumento, "Documento", 2, "tDocumento", 1750, 0, 0, "@-@@@@@-@@@@@@@@@", _
                                    "Mesa", 2, "Mesa", 900, 0, 0, "", _
                                    "Motorizado", 2, "Motorizado", 1100, 0, 0, "", _
                                    "Observaci�n", 2, "tObservacion", 1400, 0, 0, "", _
                                    "Pedido", 2, "tCodigoPedido", 980, 0, 0, "", _
                                    "Total", 2, "nVenta", 900, 1, 0, "#,###,##0.00", _
                                    "Fecha", 2, "fFecha", 650, 2, 0, "dd-MMM", _
                                    "Cliente", 0, "Cliente", 3600, 0, 0, "")
                                      
   Set grdDocumento.DataSource = RsDocumento
   If RsDocumento.RecordCount = 0 Then
      sDocumento = ""
   Else
      sDocumento = IIf(IsNull(RsDocumento!tDocumento), "", RsDocumento!tDocumento)
   End If
  
   Isql = "SELECT [" & sTemporal & "].*, TPRODUCTO.tResumido AS Producto, [nVenta]*[nCantidad] AS Subtotal " & _
          "FROM TPRODUCTO INNER JOIN [" & sTemporal & "] ON TPRODUCTO.tCodigoProducto = [" & sTemporal & "].tCodigoProducto where Estado = 'X'"
   
   Set RsDetalle2 = Lib.OpenRecordset(Isql, Cn)

   Call ConfGrilla(5, grdDetalle2, "Producto", 2, "Producto", 1600, 0, 0, "", _
                                   "Precio", 2, "nPrecioVenta", 810, 1, 0, "###,###,##0.00", _
                                   "Cant", 2, "nCantidad", 500, 1, 0, "#,##0.00", _
                                   "SubTotal", 2, "nVenta", 830, 1, 0, "###,###,##0.00", _
                                   "Pedido", 2, "tCodigoPedido", 1000, 0, 0, "")

   Set grdDetalle2.DataSource = RsDetalle2
               
  'Imprime el Totales
   CalculaMonto
   
   'anulacion de documentos por nota de credito
   
   If RsTparametro!lanula = True Then
    cmdNotasCredito.Item(13).Visible = True
    Else
    cmdNotasCredito.Item(13).Visible = False
   End If
   '--------------------------------------------
   
   
   Screen.MousePointer = vbDefault
End Sub

Public Sub ActivaDetalle(Activa As Boolean)
   
   cmdMovimiento(0).Enabled = Activa
   cmdMovimiento(1).Enabled = Activa
   cmdMovimiento(2).Enabled = Activa
   cmdMovimiento(3).Enabled = Activa
   
   cmdNavegar(0).Enabled = Not Activa
   cmdNavegar(1).Enabled = Not Activa
   cmdNavegar(2).Enabled = Not Activa
   cmdNavegar(3).Enabled = Not Activa
   cmdNavegar(4).Enabled = Not Activa
   cmdNavegar(5).Enabled = Not Activa
   cmdNavegar(6).Enabled = Not Activa
   cmdNavegar(7).Enabled = Not Activa
   cmdNavegar(8).Enabled = Not Activa
   cmdNavegar(9).Enabled = Not Activa
   cmdNavegar(10).Enabled = Not Activa
   cmdNavegar(11).Enabled = Not Activa
   
   cmdOpcion(0).Enabled = Activa     ' Salir
   cmdOpcion(1).Enabled = Not Activa ' Emisi�n R�pida
   cmdOpcion(2).Enabled = Activa     ' Generar Documento
   cmdOpcion(3).Enabled = Activa     ' Agregar Pedido
   cmdOpcion(4).Enabled = Activa     ' Aumentar
   cmdOpcion(5).Enabled = Activa     ' Disminuir
   cmdOpcion(6).Enabled = Not Activa ' Pagos
   cmdOpcion(7).Enabled = Not Activa ' Re emision
   cmdOpcion(8).Enabled = Not Activa ' Cuentas po cobrar
   cmdOpcion(9).Enabled = Not Activa ' Eliminar Documento
End Sub

Private Sub Form_Unload(Cancel As Integer)
   Cn.Execute "drop table " & sTemporal
   
   Set RsDetalle1 = Nothing
   Set RsDetalle2 = Nothing
   Set RsDocumento = Nothing
   Set rst1 = Nothing
   Set RsT2 = Nothing
   Set RsImpresion = Nothing
   Set RsEmision = Nothing
   Set frmDocumento = Nothing
End Sub

Public Sub CalculaMonto()
   txtTotal1.Caption = Format(Calcular("SELECT Sum(nVenta) AS Codigo From " & sTemporal & " where Estado = ' '", Cn), "###,###,###,##0.00")
   nTotal = Calcular("SELECT Sum(nVenta) AS Codigo From " & sTemporal & " where Estado = 'X'", Cn)
   txtTotal2.Caption = Format(nTotal, "###,###,###,##0.00")
   txtTotal.Caption = Format(nTotal / Val(txtDivision.Caption), "###,###,###,##0.00")
End Sub


Private Sub grdDocumento_HeadClick(ByVal ColIndex As Integer)
    Screen.MousePointer = vbHourglass
    RsDocumento.Sort = grdDocumento.Columns(ColIndex).DataField & " ASC"
    RsDocumento.Requery
    Screen.MousePointer = vbDefault
End Sub
Private Sub grdDocumento_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
   If RsDocumento.RecordCount <> 0 Then
      If RsDocumento!tTipoPedido = "02" Then
         cmdOpcion(1).Enabled = True
         cmdOpcion(10).Enabled = True
      Else
         cmdOpcion(1).Enabled = False
         cmdOpcion(10).Enabled = False
      End If
      If lNCAnula Then
        If RsDocumento!tTipoDocumento = "00" Then
          cmdNotasCredito.Item(13).Visible = False
        Else
          cmdNotasCredito.Item(13).Visible = True
        End If
      End If
   End If
   ' anulacion por nota de credito
   'anroDocumento = grdDocumento.Columns(0).Text
   
End Sub

Sub MuestraDP()
   If ChkDetalle.value = 1 Then
        grdDocumento.Top = 50
        grdDocumento.Width = 11805
        grdDocumento.Height = 7550
   Else
        grdDocumento.Top = 4470
        grdDocumento.Width = 9670
        grdDocumento.Height = 3150
   End If
End Sub

Private Sub ChkDetalle_Click()
    MuestraDP
End Sub


Private Sub ImprimeFormatoA()
                    Dim Reporte As New dsrBoleta

                    If RsEmision!lImprimeImageCab Then
                       iImagenCab = Generar_Imagen(Cn, "select iImagenCabDoc As imagen from TCAJA where tCaja='" & sCaja & "'", "imagen", "\cliente.jpg")
                    End If
                
                    Reporte.DiscardSavedData
                    Reporte.Database.SetDataSource RsImpDocumentoE
                    
                    If xImpresioDE = "B" Then
                       Reporte.Text13.SetText "BOLETA DE VENTA ELECTRONICA"
                    ElseIf xImpresioDE = "F" Then
                       Reporte.Text13.SetText "FACTURA ELECTRONICA"
                    End If
                                                        
                    Reporte.Text8.SetText sRazonSocial
                    Reporte.ReportTitle = sDireccion
                    Reporte.Text15.SetText sTelefono
                    Reporte.Text14.SetText sFax
                    Reporte.Text16.SetText sRUC
                    Reporte.Text50.SetText sWeb
                    
                    If Calcular(" SELECT case when  lImpresionRetencion=1 then 1 else 0 end  as codigo FROM vTipodocumento WHERE Codigo=(SELECT TTIPODOCUMENTO FROM MDOCUMENTO WHERE tDocumento='" & sDocumento & "')", Cn) = 1 Then
                    Reporte.ReportComments = tTextoAgenteRetencion
                    End If
                    
                    xMontoTexto = "SON: " & NumeroCadena(str(RsImpDocumentoE!nVenta)) & " " & sMonedaN
                    Reporte.Text4.SetText xMontoTexto
                    Reporte.Text31.SetText xImpresionFE

'                        frmEmite.CRViewer.DisplayGroupTree = False
'                        frmEmite.CRViewer.ReportSource = Reporte
'                        frmEmite.CRViewer.ViewReport
'                        frmEmite.Show vbModal
                    
                    Reporte.SelectPrinter Printer.DriverName, Printer.DeviceName, Printer.Port
                    Reporte.PaperOrientation = crPortrait
                    Reporte.PrintOut False, 1, False, 1, 1
                    '----------------
                    
                    If iImagenCab Then
                       Kill App.Path & "\cliente.jpg"
                    End If
End Sub

Private Sub ImprimeFormatoAConsumo()

        Dim ReporteC As New dsrBoletaC
        
        If RsEmision!lImprimeImageCab Then
           iImagenCab = Generar_Imagen(Cn, "select iImagenCabDoc As imagen from TCAJA where tCaja='" & sCaja & "'", "imagen", "\cliente.jpg")
        End If
        
        ReporteC.DiscardSavedData
        ReporteC.Database.SetDataSource RsImpDocumentoE
        
        If xImpresioDE = "B" Then
           ReporteC.Text13.SetText "BOLETA DE VENTA ELECTRONICA"
        ElseIf xImpresioDE = "F" Then
           ReporteC.Text13.SetText "FACTURA ELECTRONICA"
        End If
                                            
        ReporteC.Text8.SetText sRazonSocial
        ReporteC.ReportTitle = sDireccion
        ReporteC.Text15.SetText sTelefono
        ReporteC.Text33.SetText sFax
        ReporteC.Text16.SetText sRUC
        ReporteC.Text50.SetText sWeb
        
        ReporteC.Text31.SetText sDescrip
        
        If Calcular(" SELECT case when  lImpresionRetencion=1 then 1 else 0 end  as codigo FROM vTipodocumento WHERE Codigo=(SELECT TTIPODOCUMENTO FROM MDOCUMENTO WHERE tDocumento='" & sDocumento & "')", Cn) = 1 Then
           ReporteC.ReportComments = tTextoAgenteRetencion
        End If
        
        xMontoTexto = "SON: " & NumeroCadena(str(RsImpDocumentoE!nVenta)) & " " & sMonedaN
        ReporteC.Text4.SetText xMontoTexto
        ReporteC.Text32.SetText xImpresionFE
    
    '                        frmEmite.CRViewer.DisplayGroupTree = False
    '                        frmEmite.CRViewer.ReportSource = ReporteC
    '                        frmEmite.CRViewer.ViewReport
    '                        frmEmite.Show vbModal
        
        ReporteC.SelectPrinter Printer.DriverName, Printer.DeviceName, Printer.Port
        ReporteC.PaperOrientation = crPortrait
        ReporteC.PrintOut False, 1, False, 1, 1
        '----------------

        If iImagenCab Then
           Kill App.Path & "\cliente.jpg"
        End If
End Sub



