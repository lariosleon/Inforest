VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{562E3E04-2C31-4ECE-83F4-4017EEE51D40}#8.0#0"; "todg8.ocx"
Begin VB.Form frmDocumentoCorrelativo 
   BackColor       =   &H00C0C0C0&
   ClientHeight    =   8820
   ClientLeft      =   -3525
   ClientTop       =   -1125
   ClientWidth     =   14520
   Icon            =   "frmDocumentoCorrelativo.frx":0000
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   8790.303
   ScaleMode       =   0  'User
   ScaleWidth      =   14520
   WindowState     =   2  'Maximized
   Begin VB.Frame fraGrilla 
      Height          =   7125
      Left            =   0
      TabIndex        =   28
      Top             =   0
      Width           =   13695
      Begin TrueOleDBGrid80.TDBGrid grdGrilla 
         Height          =   6870
         Left            =   90
         TabIndex        =   0
         Top             =   180
         Width           =   13455
         _ExtentX        =   23733
         _ExtentY        =   12118
         _LayoutType     =   4
         _RowHeight      =   21
         _WasPersistedAsPixels=   0
         Columns(0)._VlistStyle=   0
         Columns(0)._MaxComboItems=   5
         Columns(0).FooterText=   "Footer Text"
         Columns(0).DataField=   ""
         Columns(0).NumberFormat=   "True/False"
         Columns(0)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns.Count   =   1
         Splits(0)._UserFlags=   0
         Splits(0).AnchorRightColumn=   -1  'True
         Splits(0).MarqueeStyle=   3
         Splits(0).RecordSelectorWidth=   503
         Splits(0)._SavedRecordSelectors=   -1  'True
         Splits(0).ScrollBars=   3
         Splits(0).AllowColMove=   -1  'True
         Splits(0).AllowRowSelect=   0   'False
         Splits(0).FetchRowStyle=   -1  'True
         Splits(0).DividerStyle=   3
         Splits(0).DividerColor=   32768
         Splits(0).FilterBar=   -1  'True
         Splits(0).SpringMode=   0   'False
         Splits(0)._PropDict=   "_ColumnProps,515,0;_UserFlags,518,3"
         Splits(0)._ColumnProps(0)=   "Columns.Count=1"
         Splits(0)._ColumnProps(1)=   "Column(0).Width=2752"
         Splits(0)._ColumnProps(2)=   "Column(0).DividerColor=0"
         Splits(0)._ColumnProps(3)=   "Column(0)._WidthInPix=2646"
         Splits(0)._ColumnProps(4)=   "Column(0)._EditAlways=0"
         Splits(0)._ColumnProps(5)=   "Column(0)._ColStyle=260"
         Splits(0)._ColumnProps(6)=   "Column(0).Order=1"
         Splits.Count    =   1
         PrintInfos(0)._StateFlags=   0
         PrintInfos(0).Name=   "piInternal 0"
         PrintInfos(0).PageHeaderFont=   "Size=9,Charset=0,Weight=400,Underline=0,Italic=0,Strikethrough=0,Name=Calibri"
         PrintInfos(0).PageFooterFont=   "Size=9,Charset=0,Weight=400,Underline=0,Italic=0,Strikethrough=0,Name=Calibri"
         PrintInfos(0).PageHeaderHeight=   0
         PrintInfos(0).PageFooterHeight=   0
         PrintInfos(0).NeedTotalPageCount=   -1  'True
         PrintInfos.Count=   1
         AllowUpdate     =   0   'False
         DefColWidth     =   0
         HeadLines       =   1
         FootLines       =   1
         Caption         =   "Correlativo"
         TabAction       =   2
         MultipleLines   =   0
         CellTips        =   1
         CellTipsWidth   =   0
         InsertMode      =   0   'False
         MultiSelect     =   0
         DataView        =   2
         GroupByCaption  =   "Arrastre aqui la columna a agrupar"
         DeadAreaBackColor=   -2147483637
         ScrollTrack     =   -1  'True
         RowDividerColor =   12632256
         RowSubDividerColor=   -2147483648
         DirectionAfterEnter=   1
         DirectionAfterTab=   1
         MaxRows         =   250000
         ViewColumnCaptionWidth=   0
         ViewColumnWidth =   0
         CollapseColor   =   16711935
         _PropDict       =   "_ExtentX,2003,3;_ExtentY,2004,3;_LayoutType,512,2;_RowHeight,16,3;_StyleDefs,513,0;_WasPersistedAsPixels,516,2"
         _StyleDefs(0)   =   "_StyleRoot:id=0,.parent=-1,.alignment=3,.valignment=0,.bgcolor=&H80000005&"
         _StyleDefs(1)   =   ":id=0,.fgcolor=&H80000008&,.wraptext=0,.locked=0,.transparentBmp=0"
         _StyleDefs(2)   =   ":id=0,.fgpicPosition=0,.bgpicMode=0,.appearance=0,.borderSize=0,.ellipsis=0"
         _StyleDefs(3)   =   ":id=0,.borderColor=&H80000005&,.borderType=0,.bold=0,.fontsize=825,.italic=0"
         _StyleDefs(4)   =   ":id=0,.underline=0,.strikethrough=0,.charset=0"
         _StyleDefs(5)   =   ":id=0,.fontname=MS Sans Serif"
         _StyleDefs(6)   =   "Style:id=1,.parent=0,.namedParent=33"
         _StyleDefs(7)   =   "CaptionStyle:id=4,.parent=2,.namedParent=33"
         _StyleDefs(8)   =   "HeadingStyle:id=2,.parent=1,.namedParent=34"
         _StyleDefs(9)   =   "FooterStyle:id=3,.parent=1,.bgcolor=&H80000018&,.bold=0,.fontsize=900,.italic=0"
         _StyleDefs(10)  =   ":id=3,.underline=0,.strikethrough=0,.charset=0"
         _StyleDefs(11)  =   ":id=3,.fontname=Calibri"
         _StyleDefs(12)  =   "InactiveStyle:id=5,.parent=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
         _StyleDefs(13)  =   "SelectedStyle:id=6,.parent=1,.namedParent=36"
         _StyleDefs(14)  =   "EditorStyle:id=7,.parent=1"
         _StyleDefs(15)  =   "HighlightRowStyle:id=8,.parent=1,.namedParent=36"
         _StyleDefs(16)  =   "EvenRowStyle:id=9,.parent=1,.namedParent=39"
         _StyleDefs(17)  =   "OddRowStyle:id=10,.parent=1,.namedParent=40"
         _StyleDefs(18)  =   "RecordSelectorStyle:id=11,.parent=2,.namedParent=41"
         _StyleDefs(19)  =   "FilterBarStyle:id=12,.parent=1,.namedParent=42"
         _StyleDefs(20)  =   "Splits(0).Style:id=63,.parent=1,.namedParent=33"
         _StyleDefs(21)  =   "Splits(0).CaptionStyle:id=84,.parent=4"
         _StyleDefs(22)  =   "Splits(0).HeadingStyle:id=64,.parent=2"
         _StyleDefs(23)  =   "Splits(0).FooterStyle:id=65,.parent=3"
         _StyleDefs(24)  =   "Splits(0).InactiveStyle:id=66,.parent=5"
         _StyleDefs(25)  =   "Splits(0).SelectedStyle:id=80,.parent=6,.bgcolor=&HE7FAB6&,.fgcolor=&H80000009&"
         _StyleDefs(26)  =   "Splits(0).EditorStyle:id=79,.parent=7"
         _StyleDefs(27)  =   "Splits(0).HighlightRowStyle:id=81,.parent=8"
         _StyleDefs(28)  =   "Splits(0).EvenRowStyle:id=82,.parent=9"
         _StyleDefs(29)  =   "Splits(0).OddRowStyle:id=83,.parent=10"
         _StyleDefs(30)  =   "Splits(0).RecordSelectorStyle:id=85,.parent=11"
         _StyleDefs(31)  =   "Splits(0).FilterBarStyle:id=86,.parent=12"
         _StyleDefs(32)  =   "Splits(0).Columns(0).Style:id=90,.parent=63"
         _StyleDefs(33)  =   "Splits(0).Columns(0).HeadingStyle:id=87,.parent=64,.alignment=0"
         _StyleDefs(34)  =   "Splits(0).Columns(0).FooterStyle:id=88,.parent=65"
         _StyleDefs(35)  =   "Splits(0).Columns(0).EditorStyle:id=89,.parent=79"
         _StyleDefs(36)  =   "Named:id=33:Normal"
         _StyleDefs(37)  =   ":id=33,.parent=0,.bgcolor=&H80000018&,.bold=0,.fontsize=900,.italic=0"
         _StyleDefs(38)  =   ":id=33,.underline=0,.strikethrough=0,.charset=0"
         _StyleDefs(39)  =   ":id=33,.fontname=Calibri"
         _StyleDefs(40)  =   "Named:id=34:Heading"
         _StyleDefs(41)  =   ":id=34,.parent=33,.alignment=2,.valignment=2,.bgcolor=&H80000003&"
         _StyleDefs(42)  =   ":id=34,.fgcolor=&H800000&,.wraptext=-1,.locked=0,.borderSize=1,.bold=-1"
         _StyleDefs(43)  =   ":id=34,.fontsize=900,.italic=0,.underline=0,.strikethrough=0,.charset=0"
         _StyleDefs(44)  =   ":id=34,.fontname=Arial Black"
         _StyleDefs(45)  =   "Named:id=35:Footing"
         _StyleDefs(46)  =   ":id=35,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
         _StyleDefs(47)  =   "Named:id=36:Selected"
         _StyleDefs(48)  =   ":id=36,.parent=33,.bgcolor=&HE7FAB6&,.fgcolor=&H0&,.borderColor=&H808000&"
         _StyleDefs(49)  =   ":id=36,.bold=-1,.fontsize=900,.italic=0,.underline=0,.strikethrough=0,.charset=0"
         _StyleDefs(50)  =   ":id=36,.fontname=Calibri"
         _StyleDefs(51)  =   "Named:id=37:Caption"
         _StyleDefs(52)  =   ":id=37,.parent=34,.alignment=2,.bgcolor=&HFFFF&"
         _StyleDefs(53)  =   "Named:id=38:HighlightRow"
         _StyleDefs(54)  =   ":id=38,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H800000&,.locked=0,.bold=-1"
         _StyleDefs(55)  =   ":id=38,.fontsize=2025,.italic=0,.underline=0,.strikethrough=0,.charset=0"
         _StyleDefs(56)  =   ":id=38,.fontname=Calibri"
         _StyleDefs(57)  =   "Named:id=39:EvenRow"
         _StyleDefs(58)  =   ":id=39,.parent=33,.bgcolor=&HFFFFFF&,.bold=0,.fontsize=975,.italic=0"
         _StyleDefs(59)  =   ":id=39,.underline=0,.strikethrough=0,.charset=0"
         _StyleDefs(60)  =   ":id=39,.fontname=Calibri"
         _StyleDefs(61)  =   "Named:id=40:OddRow"
         _StyleDefs(62)  =   ":id=40,.parent=33,.bgcolor=&H80000005&,.borderColor=&H80000002&,.bold=0"
         _StyleDefs(63)  =   ":id=40,.fontsize=975,.italic=0,.underline=0,.strikethrough=0,.charset=0"
         _StyleDefs(64)  =   ":id=40,.fontname=Calibri"
         _StyleDefs(65)  =   "Named:id=41:RecordSelector"
         _StyleDefs(66)  =   ":id=41,.parent=34,.bgcolor=&H80000003&,.fgcolor=&H80000009&,.bold=-1"
         _StyleDefs(67)  =   ":id=41,.fontsize=975,.italic=0,.underline=0,.strikethrough=0,.charset=0"
         _StyleDefs(68)  =   ":id=41,.fontname=Calibri"
         _StyleDefs(69)  =   "Named:id=42:FilterBar"
         _StyleDefs(70)  =   ":id=42,.parent=33,.bgcolor=&H80000018&,.bold=-1,.fontsize=975,.italic=0"
         _StyleDefs(71)  =   ":id=42,.underline=0,.strikethrough=0,.charset=0"
         _StyleDefs(72)  =   ":id=42,.fontname=Calibri"
         _StyleDefs(73)  =   "Named:id=13:CabeceraTitulo"
         _StyleDefs(74)  =   ":id=13,.parent=34,.bold=-1,.fontsize=1125,.italic=0,.underline=0"
         _StyleDefs(75)  =   ":id=13,.strikethrough=0,.charset=0"
         _StyleDefs(76)  =   ":id=13,.fontname=Arial Black"
      End
      Begin MSComDlg.CommonDialog cmdSave 
         Left            =   135
         Top             =   270
         _ExtentX        =   847
         _ExtentY        =   847
         _Version        =   393216
      End
   End
   Begin VB.PictureBox Picture1 
      Align           =   2  'Align Bottom
      Height          =   1380
      Left            =   0
      ScaleHeight     =   1320
      ScaleWidth      =   14460
      TabIndex        =   12
      Top             =   7440
      Width           =   14520
      Begin VB.CommandButton cmdOpcion 
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
         Index           =   6
         Left            =   10470
         Picture         =   "frmDocumentoCorrelativo.frx":1782
         Style           =   1  'Graphical
         TabIndex        =   11
         ToolTipText     =   "Salir"
         Top             =   45
         Width           =   555
      End
      Begin VB.CommandButton cmdOpcion 
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
         Index           =   11
         Left            =   8790
         Picture         =   "frmDocumentoCorrelativo.frx":1874
         Style           =   1  'Graphical
         TabIndex        =   7
         ToolTipText     =   "Pagos de documentos"
         Top             =   45
         Width           =   555
      End
      Begin VB.CommandButton cmdOpcion 
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
         Index           =   10
         Left            =   9345
         Picture         =   "frmDocumentoCorrelativo.frx":1976
         Style           =   1  'Graphical
         TabIndex        =   8
         ToolTipText     =   "Ver detalle de pagos"
         Top             =   45
         Width           =   555
      End
      Begin VB.CommandButton cmdOpcion 
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
         Index           =   5
         Left            =   8235
         Picture         =   "frmDocumentoCorrelativo.frx":1AB8
         Style           =   1  'Graphical
         TabIndex        =   6
         ToolTipText     =   "Modifica los Pagos"
         Top             =   45
         Width           =   555
      End
      Begin VB.CommandButton cmdOpcion 
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
         Index           =   8
         Left            =   6580
         Picture         =   "frmDocumentoCorrelativo.frx":1BB2
         Style           =   1  'Graphical
         TabIndex        =   3
         ToolTipText     =   "Cambio de Documentos"
         Top             =   45
         Width           =   555
      End
      Begin VB.PictureBox Picture3 
         BackColor       =   &H80000004&
         Height          =   615
         Left            =   60
         ScaleHeight     =   555
         ScaleWidth      =   5325
         TabIndex        =   20
         Top             =   30
         Width           =   5385
         Begin VB.CommandButton cmdNavegar 
            Height          =   555
            Index           =   1
            Left            =   540
            Picture         =   "frmDocumentoCorrelativo.frx":1FF4
            Style           =   1  'Graphical
            TabIndex        =   26
            TabStop         =   0   'False
            Top             =   0
            Width           =   540
         End
         Begin VB.CommandButton cmdNavegar 
            Height          =   555
            Index           =   2
            Left            =   1080
            Picture         =   "frmDocumentoCorrelativo.frx":2536
            Style           =   1  'Graphical
            TabIndex        =   25
            TabStop         =   0   'False
            Top             =   0
            Width           =   540
         End
         Begin VB.CommandButton cmdNavegar 
            Height          =   555
            Index           =   0
            Left            =   0
            Picture         =   "frmDocumentoCorrelativo.frx":2A78
            Style           =   1  'Graphical
            TabIndex        =   24
            TabStop         =   0   'False
            Top             =   0
            Width           =   540
         End
         Begin VB.CommandButton cmdNavegar 
            Height          =   555
            Index           =   5
            Left            =   4770
            Picture         =   "frmDocumentoCorrelativo.frx":2FBA
            Style           =   1  'Graphical
            TabIndex        =   23
            TabStop         =   0   'False
            Top             =   0
            Width           =   540
         End
         Begin VB.CommandButton cmdNavegar 
            Height          =   555
            Index           =   4
            Left            =   4230
            Picture         =   "frmDocumentoCorrelativo.frx":34FC
            Style           =   1  'Graphical
            TabIndex        =   22
            TabStop         =   0   'False
            Top             =   0
            Width           =   540
         End
         Begin VB.CommandButton cmdNavegar 
            Height          =   555
            Index           =   3
            Left            =   3690
            Picture         =   "frmDocumentoCorrelativo.frx":3A3E
            Style           =   1  'Graphical
            TabIndex        =   21
            TabStop         =   0   'False
            Top             =   0
            Width           =   540
         End
         Begin VB.Label cmdTexto 
            Alignment       =   2  'Center
            BackColor       =   &H80000004&
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
            Height          =   315
            Left            =   1650
            TabIndex        =   27
            Top             =   150
            Width           =   2025
         End
      End
      Begin VB.CommandButton cmdEmite 
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
         Left            =   7125
         Picture         =   "frmDocumentoCorrelativo.frx":3F80
         Style           =   1  'Graphical
         TabIndex        =   4
         ToolTipText     =   "Emite"
         Top             =   45
         Width           =   555
      End
      Begin VB.CommandButton cmdOpcion 
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
         Index           =   14
         Left            =   17685
         Picture         =   "frmDocumentoCorrelativo.frx":44B2
         Style           =   1  'Graphical
         TabIndex        =   19
         ToolTipText     =   "Salir"
         Top             =   -450
         Width           =   555
      End
      Begin VB.CommandButton cmdProcesa 
         Height          =   555
         Left            =   5310
         Picture         =   "frmDocumentoCorrelativo.frx":45A4
         Style           =   1  'Graphical
         TabIndex        =   18
         ToolTipText     =   "Filtrar"
         Top             =   720
         Width           =   555
      End
      Begin VB.Frame fraFechas 
         Caption         =   " Rango de Fechas "
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   645
         Left            =   90
         TabIndex        =   13
         Top             =   630
         Width           =   5235
         Begin MSComCtl2.DTPicker dtpFecFin 
            Height          =   345
            Left            =   3195
            TabIndex        =   14
            Top             =   225
            Width           =   1770
            _ExtentX        =   3122
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
            Format          =   84869121
            CurrentDate     =   37539
         End
         Begin MSComCtl2.DTPicker dtpFecIni 
            Height          =   345
            Left            =   720
            TabIndex        =   15
            Top             =   225
            Width           =   1770
            _ExtentX        =   3122
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
            Format          =   84869121
            CurrentDate     =   37539
         End
         Begin VB.Label Label2 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   " Del "
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
            TabIndex        =   17
            Top             =   300
            Width           =   405
            WordWrap        =   -1  'True
         End
         Begin VB.Label Label2 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   " al "
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
            Left            =   2700
            TabIndex        =   16
            Top             =   300
            Width           =   225
            WordWrap        =   -1  'True
         End
      End
      Begin VB.CommandButton cmdExporta 
         Height          =   555
         Left            =   7680
         Picture         =   "frmDocumentoCorrelativo.frx":46A6
         Style           =   1  'Graphical
         TabIndex        =   5
         ToolTipText     =   "Exportar HTML"
         Top             =   45
         Width           =   555
      End
      Begin VB.CommandButton cmdOpcion 
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
         Index           =   9
         Left            =   6015
         Picture         =   "frmDocumentoCorrelativo.frx":4BD8
         Style           =   1  'Graphical
         TabIndex        =   2
         ToolTipText     =   "Reimprime el documento"
         Top             =   45
         Width           =   555
      End
      Begin VB.CommandButton cmdOpcion 
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
         Left            =   5460
         Picture         =   "frmDocumentoCorrelativo.frx":510A
         Style           =   1  'Graphical
         TabIndex        =   1
         ToolTipText     =   "Detalle del Pedido"
         Top             =   45
         Width           =   555
      End
      Begin VB.CommandButton cmdOpcion 
         Appearance      =   0  'Flat
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
         Left            =   9900
         Picture         =   "frmDocumentoCorrelativo.frx":520C
         Style           =   1  'Graphical
         TabIndex        =   10
         ToolTipText     =   "Anula Pagos"
         Top             =   45
         Width           =   555
      End
      Begin VB.CommandButton cmdPropina 
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
         Left            =   9900
         Picture         =   "frmDocumentoCorrelativo.frx":530E
         Style           =   1  'Graphical
         TabIndex        =   9
         ToolTipText     =   "Registra las Propinas"
         Top             =   45
         Visible         =   0   'False
         Width           =   555
      End
   End
   Begin VB.Image imageHash 
      Height          =   495
      Left            =   13680
      Top             =   5040
      Visible         =   0   'False
      Width           =   135
   End
   Begin VB.Image imageCab 
      Height          =   375
      Left            =   0
      Top             =   4200
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Image imagepIE 
      Height          =   135
      Left            =   0
      Top             =   0
      Visible         =   0   'False
      Width           =   135
   End
End
Attribute VB_Name = "frmDocumentoCorrelativo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public RsCabecera As Recordset
Dim RsImpresion As Recordset
Dim rstFuente As ADODB.Recordset
Dim nColumna As Integer
Dim RsEmision As Recordset
Dim sUsuarioAutoriza As String
Dim flagTipoCambio As Boolean
Dim tTipoPago As String
Dim tnumerorecibo As String
Dim nRespuesta As Integer
Dim lImprimeAlternativa As Boolean
Dim col             As TrueOleDBGrid80.Column
Dim cols            As TrueOleDBGrid80.Columns
Dim lAsc            As Boolean

Dim lPagoPV As Boolean

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
Private Sub cmdEmite_Click()
    With grdGrilla.PrintInfo
        'Cabecera
        .SettingsOrientation = 2
        .PageHeaderFont.Italic = True
        .RepeatColumnHeaders = True
        .PreviewMaximize = True
        .SettingsMarginLeft = 400
        .SettingsMarginRight = 400
        .SettingsMarginTop = 400
        .SettingsMarginBottom = 400
        .VariableRowHeight = True
        .PageFooter = "\tPage: \p"
        .PreviewCaption = "Emite de " & Me.Caption
        .PrintPreview
    End With
End Sub

Private Sub cmdExporta_Click()
    cmdSave.Filter = "Hyper Text Markup Language|*.html"
    cmdSave.ShowSave
    If cmdSave.FileName = "" Then
       Exit Sub
    End If
    Screen.MousePointer = vbHourglass
    grdGrilla.ExportToFile cmdSave.FileName, False, False
    Screen.MousePointer = vbNormal
End Sub

Private Sub Form_Resize()
   fraGrilla.Height = IIf(Me.Height - 2000 > 0, Me.Height - 2000, 0)
   fraGrilla.Width = IIf(Me.Width - 300 > 0, Me.Width - 300, 0)
   
   grdGrilla.Height = IIf(fraGrilla.Height - 300 > 0, fraGrilla.Height - 300, 0)
   grdGrilla.Width = IIf(fraGrilla.Width - 300 > 0, fraGrilla.Width - 300, 0)
End Sub

Private Sub grdGrilla_HeadClick(ByVal ColIndex As Integer)
    nColumna = ColIndex
    lAsc = Not lAsc
    If lAsc Then
       RsCabecera.Sort = grdGrilla.Columns(ColIndex).DataField & " ASC"
    Else
       RsCabecera.Sort = grdGrilla.Columns(ColIndex).DataField & " DESC"
    End If
End Sub

Private Sub grdGrilla_FilterChange()
    On Error GoTo errHandler
    Set cols = grdGrilla.Columns
    Dim c As Integer
    c = grdGrilla.col
    grdGrilla.HoldFields
    RsCabecera.Filter = getFilter(cols)
    grdGrilla.col = c
    grdGrilla.EditActive = True
    Exit Sub

errHandler:
    MsgBox "Error: Caracter no valido para esta columna" & vbCrLf & err.Description, vbCritical + vbOKOnly, sMensaje
    
    Call cmdClearFilter_Click(cols)
    RsCabecera.Filter = adFilterNone
End Sub

Private Sub cmdNavegar_Click(Index As Integer)
    Select Case Index
           Case Is = 0 'Primero
                MoverPuntero Primero, grdGrilla
           Case Is = 1 'PgUp
                MoverPuntero pgup, grdGrilla
           Case Is = 2 'Previo
                MoverPuntero previo, grdGrilla
           Case Is = 3 'Siguiente
                MoverPuntero siguiente, grdGrilla
           Case Is = 4 'PgDn
                MoverPuntero pgdn, grdGrilla
           Case Is = 5 'Ultimo
                MoverPuntero Ultimo, grdGrilla
    End Select
End Sub

Private Sub cmdOpcion_Click(Index As Integer)

    Select Case Index
           Case Is = 0  'Anular Documento
                If RsCabecera.RecordCount = 0 Then
                   Exit Sub
                End If
           
                If Periodo(RsCabecera!fFecha) Then
                   Exit Sub
                End If
                Dim EstadoDoc As String
                '// validacion ingresada 05/05/2018 ELDC
                EstadoDoc = Calcular("select testadodocumento as Codigo from MDOCUMENTO where tDocumento='" & RsCabecera!tDocumento & "'", Cn)
                '// fin validacion
                'If RsCabecera!tEstadoDocumento = "01" And Not lMultiCajero Then
                If EstadoDoc = "01" And Not lMultiCajero Then
                   MsgBox "Es un Documento Emitido", vbExclamation, sMensaje
                   Exit Sub
                End If
           
                If EstadoDoc = "04" Then
                   MsgBox "Es un Documento Anulado", vbExclamation, sMensaje
                   Exit Sub
                End If
           
                If EstadoDoc = "03" Then
                   MsgBox "Es un Documento por Cobrar", vbExclamation, sMensaje
                   Exit Sub
                End If
                          
                If RsCabecera!tCaja <> sCaja Then
                   MsgBox "No se puede anular el pago porque pertenece a otra Caja (Caja " & sCaja & ") ", vbCritical, sMensaje
                   Exit Sub
                End If
                
                If Calcular("select count(tdocumento) as Codigo from DPAGOTARJETA where tdocumento='" & RsCabecera!tDocumento & "'", Cn) > 0 Then
                   MsgBox "Documento con Pago TC Pinpad, anular primero el pago", vbExclamation, sMensaje
                   Exit Sub
                End If
                
                 If Calcular("select count(*) as Codigo from MNOTACREDITO where tDocumento= '" & RsCabecera!tDocumento & "'", Cn) > 0 Then
                   MsgBox "Documento con Nota de Credito Generado", vbExclamation, sMensaje
                   Exit Sub
                End If
                
                If RsCabecera!tTurno = sTurno Then
                   'Password
                   If Supervisor("05") = False Then
                      MsgBox "Clave no permitida", vbExclamation, sMensaje
                      Exit Sub
                   End If
                   
                Else
                   'Password
                   If MsgBox("El Documento es de un turno Anterior, deseas continuar?", vbQuestion + vbOKCancel, sMensaje) = vbCancel Then
                      Exit Sub
                   End If
                   
                   If Supervisor("06") = False Then
                      MsgBox "Clave no permitida", vbExclamation, sMensaje
                      Exit Sub
                   End If
                End If
                sUsuarioAutoriza = sVar1

                If Not lMultiCajero Then
                    If MsgBox("Seguro de Anular el Pago? " & Format(grdGrilla.Columns(2).value, "@-@@@@@-@@@@@@@@@") & _
                              "?", vbQuestion + vbOKCancel, sMensaje) = vbCancel Then
                       Exit Sub
                    End If
                                   
                    If EstadoDoc = "02" Then
                       sPedido = Calcular("select tcodigoPedido as Codigo from DDOCUMENTO where tDocumento='" & RsCabecera!tDocumento & "'", Cn)
                       
                       'Control de Mesas
                       Dim RsEstadoMesa As Recordset
                       Set RsEstadoMesa = Lib.OpenRecordset("select tCodigoMesa, tDetallado, tEstadoMesa from TMESA WHERE TCODIGOMESA IN(SELECT TMESA FROM MPEDIDO WHERE TCODIGOPEDIDO='" & sPedido & "')", Cn)
                       If RsEstadoMesa.RecordCount > 0 Then
                          If RsEstadoMesa!tEstadoMesa <> "01" And RsEstadoMesa!tEstadoMesa <> "04" Then
                             If MsgBox("Pedido asociado a una mesa ocupada" & Chr(13) & "Deseas continuar con la recuperación del documento?", vbExclamation + vbYesNo, sMensaje) = vbNo Then
                                Cn.Execute "Update TMESA set tEstadoMesa='02' where tCodigoMesa='" & RsEstadoMesa!tCodigoMesa & "'"
                                Exit Sub
                             End If
                          End If
                       End If
                       Set RsEstadoMesa = Nothing
                       
                       Screen.MousePointer = vbHourglass
                       
                        'XX AGOSTO 2011 LG ACTIVACION DEL INGRESO SI ES PAGADO POR RECIBO

                        tTipoPago = Calcular("SELECT ISNULL(TOTROTIPOPAGO,'') AS CODIGO FROM DPAGODOCUMENTO WHERE TDOCUMENTO='" & RsCabecera!tDocumento & "' and totrotipopago='001'", Cn)
                        If tTipoPago = "001" Then
                            tnumerorecibo = Calcular("SELECT ISNULL(TNUMERO,'') AS CODIGO FROM DPAGODOCUMENTO WHERE TDOCUMENTO ='" & RsCabecera!tDocumento & "' and TOTROTIPOPAGO='001' ", Cn)
                            Cn.Execute "UPDATE MINGRESO SET TESTADODOCUMENTO='01' WHERE TRECIBO in (SELECT ISNULL(TNUMERO,'') AS CODIGO FROM DPAGODOCUMENTO WHERE TDOCUMENTO ='" & RsCabecera!tDocumento & "' and TOTROTIPOPAGO='001' )"
                             'Cn.Execute "UPDATE MINGRESO SET TESTADODOCUMENTO='01' WHERE TRECIBO='" & tnumerorecibo & "'"
                        End If
                        'FIN XX
                        '-- Recuperacion de la nota de credito cuando se Anula el pago del documento
                        Dim tnotacreditorecuperar As String
                        tnotacreditorecuperar = Calcular("select isnull(tNumero,'') as codigo from DPAGODOCUMENTO where  tDocumento in ('" & RsCabecera!tDocumento & "') and totrotipopago='002'", Cn)
                        If tnotacreditorecuperar <> "0" Then
                        Cn.Execute "update MNOTACREDITO SET tEstadoDocumento='05' where tNotaCredito = '" & tnotacreditorecuperar & "'"
                        End If
                        '-- Fin Recuperacion de nota de credito

                       Cn.Execute "delete from DPAGODOCUMENTO where tDocumento = '" & RsCabecera!tDocumento & "'"
                       Cn.Execute "Update MDOCUMENTO set tClientePago = '', tEstadoDocumento = '01',lreplica = 1,tcortesia = ''  where tDocumento = '" & RsCabecera!tDocumento & "'"
                       Cn.Execute "Update DPEDIDO set tFacturado = 'F' where tDocumento ='" & RsCabecera!tDocumento & "'"
                       Cn.Execute "Update DPEDIDO set tFacturado='C' where substring(tDocumento,1,1)='0' and tDocumento ='" & RsCabecera!tDocumento & "'"

                       If lInfhotel Then
                          CnInfhotel.Execute "Update MCOMANDA set tEstado='02' where tNotaPedido='" & sPedido & "' and tDocumento ='" & RsCabecera!tDocumento & "'"
                       End If
                          
                       RsCabecera.Requery
                       nPos = RsCabecera.AbsolutePosition
                       RsCabecera.Requery
                       RsCabecera.AbsolutePosition = nPos
                   End If
                Else
                    If MsgBox("Seguro de Anular el Documento? " & Format(grdGrilla.Columns(2).value, "@-@@@@@-@@@@@@@@@") & _
                              "?", vbQuestion + vbOKCancel, sMensaje) = vbCancel Then
                       Exit Sub
                    End If
                                   
                    If EstadoDoc = "02" Then
                       sPedido = Calcular("select tcodigoPedido as Codigo from DDOCUMENTO where tDocumento='" & RsCabecera!tDocumento & "'", Cn)
                       Screen.MousePointer = vbHourglass
                       
                    'XX AGOSTO 2011 LG ACTIVACION DEL INGRESO SI ES PAGADO POR RECIBO


                        tTipoPago = Calcular("SELECT ISNULL(TOTROTIPOPAGO,'') AS CODIGO FROM DPAGODOCUMENTO WHERE TDOCUMENTO='" & RsCabecera!tDocumento & "' and totrotipopago='001'", Cn)
                        If tTipoPago = "001" Then
                            tnumerorecibo = Calcular("SELECT ISNULL(TNUMERO,'') AS CODIGO FROM DPAGODOCUMENTO WHERE TDOCUMENTO ='" & RsCabecera!tDocumento & "' and TOTROTIPOPAGO='001' ", Cn)
                            Cn.Execute "UPDATE MINGRESO SET TESTADODOCUMENTO='01' WHERE TRECIBO in (SELECT ISNULL(TNUMERO,'') AS CODIGO FROM DPAGODOCUMENTO WHERE TDOCUMENTO ='" & RsCabecera!tDocumento & "' and TOTROTIPOPAGO='001') "
                            'Cn.Execute "UPDATE MINGRESO SET TESTADODOCUMENTO='01' WHERE TRECIBO='" & tnumerorecibo & "'"
                        End If
                        'FIN XX
                       Cn.Execute "delete from DPAGODOCUMENTO where tDocumento = '" & RsCabecera!tDocumento & "'"
                       Cn.Execute "Update MDOCUMENTO set tClientePago = '', tEstadoDocumento = '04', fPago=null, tUsuarioAnulado='" & sUsuarioAutoriza & "', fRegistroAnulado=getdate(), tObservacion='Anulacion en opción Correlativo Documentos',lreplica = 1 where tDocumento = '" & RsCabecera!tDocumento & "'"
                                              
                       Cn.Execute "Update MPEDIDO set tEstadoPedido ='03', tMotivoAnulacion='000', tUsuarioAnulado='" & sUsuarioAutoriza & "', fRegAnulado= getdate(), tTurnoAnulado='" & sTurno & "', tObservacionAnulado='Anulacion en opción Correlativo Documentos' where tCodigoPedido ='" & sPedido & "'"
                       Cn.Execute "Update DPEDIDO Set tEstadoItem = 'A' where tCodigoPedido = '" & sPedido & "'"
                       Cn.Execute "delete TPRODUCTOPROPIEDAD where tCodigoPedido='" & sPedido & "'"
                       Cn.Execute "delete CPEDIDO where tCodigoPedido='" & sPedido & "'"
                       Cn.Execute "delete TCOMBOPROPIEDAD where tCodigoPedido='" & sPedido & "'"
                          
                       RsCabecera.Requery
                       If RsCabecera.RecordCount <> 0 Then
                          RsCabecera.MoveLast
                          cmdTexto.Caption = "Registro " & RsCabecera.AbsolutePosition & " de " & RsCabecera.RecordCount
                       Else
                          cmdTexto.Caption = "Registro 0 de 0"
                       End If
                   End If
                   
                End If
                Screen.MousePointer = vbDefault

           Case Is = 1  'Detalle
                If RsCabecera.RecordCount > 0 Then
                   sDocumento = grdGrilla.Columns(2).value
                   frmDocumentoCorrelativoDetalle.Show vbModal
                Else
                   MsgBox "No Existe Datos Ingresados", vbExclamation, sMensaje
                End If
                      
           Case Is = 5  'Cambio de Pagos
                If RsCabecera.RecordCount > 0 Then
                    If RsCabecera!tTipoDocumento <> "00" Then
                        If RsCabecera!tEstadoDocumento = "02" Then
                           nTC = Calcular("SELECT nVenta as Codigo From TTIPOCAMBIO WHERE fFecha = '" & Format(RsCabecera!fFecha, "yyyy/mm/dd") & "'", Cn)
                           nCargo = Round(RsCabecera!nVenta, 2)
                           sDocumento = RsCabecera!tDocumento
                           'sTurno = RsCabecera!tTurno
                           sTurnoModificacion = RsCabecera!tTurno
                           xTipo = "Modificacion"
                           
                           If MsgBox("Deseas cambiar la forma de pago del documento " & sDocumento & " ?", vbQuestion + vbOKCancel, sMensaje) = vbCancel Then
                              Exit Sub
                           End If
                           
                            '--- validar turnos
                            If Supervisor("10") = False Then
                               MsgBox "Clave no permitida", vbExclamation, sMensaje
                               Exit Sub
                            End If

                           If MsgBox("Seguro de Cambiar el tipo de Pago?", vbQuestion + vbOKCancel, sMensaje) = vbCancel Then
                              Exit Sub
                           End If
                           sFormulario = ""
                           sTipo = ""
                           
                           'CESAR TAVERNA
                           lPagoPV = Calcular("Select ISNULL(lPagoRapidoMod,0) as Codigo From TCAJA Where tCaja='" & sCaja & "'", Cn)
                           
                           If lPagoPV = True Then
                                frmPagoRapido.Show vbModal
                           Else
                                frmPago.Show vbModal
                           End If
                           '-------------
                           
                         Else
                           MsgBox "No Es un Documento Pagado", vbExclamation, sMensaje
                         End If
                    Else
                        MsgBox "Este Tipo de Documento no registra Pagos"
                    End If
                End If
                
           Case Is = 6  'Salir
                Unload Me
                     
           Case Is = 8  'Cambio Documento
                If RsCabecera.RecordCount = 0 Then
                   Exit Sub
                End If
                
                If Periodo(RsCabecera!fFecha) Then
                   Exit Sub
                End If
           
                If RsCabecera!tCaja <> sCaja Then
                   MsgBox "El Documento no pertenece a esta caja", vbExclamation, sMensaje
                   Exit Sub
                End If
           
                If RsCabecera!tEstadoDocumento <> "02" Then
                   MsgBox "No Es un Documento Cancelado", vbExclamation, sMensaje
                   Exit Sub
                End If
                
                If MsgBox("Seguro de Cambiar el Documento? " & Format(grdGrilla.Columns(2).value, "@-@@@@@-@@@@@@@@@") & _
                          "?", vbQuestion + vbOKCancel, sMensaje) = vbCancel Then
                   Exit Sub
                End If
                
                If RsCabecera!tTurno = sTurno Then
                   'Password
                   If Supervisor("05") = False Then
                      MsgBox "Clave no permitida", vbExclamation, sMensaje
                      Exit Sub
                   End If
                Else
                   'Password
                   If MsgBox("El Documento es de un turno Anterior, deseas continuar?", vbQuestion + vbOKCancel, sMensaje) = vbCancel Then
                      Exit Sub
                   End If
                   
                   If Supervisor("06") = False Then
                      MsgBox "Clave no permitida", vbExclamation, sMensaje
                      Exit Sub
                   End If
                End If
                sUsuarioAutoriza = sVar1
                
                scodigo = IIf(IsNull(RsCabecera!tCodigoCliente), "", RsCabecera!tCodigoCliente)
                sDescrip = IIf(IsNull(RsCabecera!Cliente), "", RsCabecera!Cliente)
                sDocumento = RsCabecera!tDocumento
                frmCambio.Show vbModal
                If wEnter = True Then
                   Screen.MousePointer = vbHourglass
                   RsCabecera.Requery
                   RsCabecera.Find ("tDocumento = '" & sDocumento & "'")
                   Screen.MousePointer = vbDefault
                End If
                                                
           Case Is = 9  'Reimpresion
               'Dim RsEmision As Recordset
               lReimpresion = True
               lImprimeAlternativa = False
                              
               If RsCabecera.RecordCount = 0 Then
                  Exit Sub
               End If
               
               sDocumento = RsCabecera!tDocumento
               Set RsEmision = Lib.OpenRecordset("select tFormulario, tImpresora, tTipoEmision, Cliente, lResumen, Monto, lFacturacionElectronica,lImprimeImageCab from vTipoDocumentoImpresora where tCaja='" & RsCabecera!tCaja & "' and tTipoEmision='" & RsCabecera!tTipoDocumento & "' and tCaja='" & sCaja & "'", Cn)
               If RsEmision.RecordCount = 0 Then
                  MsgBox "Error: Configuración en Caja o El documento pertenece a otra caja", vbCritical, sMensaje
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

                Case Else 'Peru, Ecuador
                     estadoReimpresion = False
                    
                End Select

                'Imprime Documentos
                  Dim wConsumo As Boolean
                  wConsumo = False
                  
                  nRespuesta = MsgBox("Por Consumo? ", vbQuestion + vbYesNoCancel + vbDefaultButton2, sMensaje)
                  If nRespuesta = vbYes Then
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
                     If validaImpresionAlternativa(sDocumento) = False Then
                         If MsgBox("Desea imprimir descripción Alternativa? ", vbQuestion + vbYesNo + vbDefaultButton2, sMensaje) = vbYes Then
                            lImprimeAlternativa = True
                         End If
                     End If
                  End If
                  
                  TimpresionDolaresDelivery = False
                      '-------- impresion en dolares si esta activo el check en el cliente delivery.
                If Calcular("select isnull(lEmisionMonedaExtranjera,0) as codigo from TDELIVERY where tCodigoDelivery=(select tClienteDelivery from MPEDIDO where tCodigoPedido=(select top 1 tCodigoPedido from DPEDIDO where tDocumento='" & sDocumento & "'))", Cn) Then
                    If MsgBox("¿Desea Imprimir en " & sMonedaE & "?", vbQuestion + vbYesNo + vbDefaultButton2) = vbYes Then
                        TimpresionDolaresDelivery = True
                    Else
                        TimpresionDolaresDelivery = False
                    End If
                Else
                    TimpresionDolaresDelivery = False
                End If
                '-------------------------------  SE INTRDUJO LAS CONSULTAS A  UN STORE PROCEDURE --------------------------
                ' ELDCQ 22/11/2017
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
               '--------------------------------------------------------------------------------------------------------------
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
                
                
                
                 'FACTURA ELECTRONICA
                  Dim xlFacturacionE As Boolean
                  Dim xlTipoDoc As String
                  
                  xlTipoDoc = Calcular("Select tTipoDocumento As Codigo From MDOCUMENTO Where tDocumento = '" & sDocumento & "'", Cn)
                  RsEmision.Find "tTipoEmision='" & xlTipoDoc & "'"
                  xlFacturacionE = IIf(RsEmision!lFacturacionElectronica = True, 1, 0)
                  xImpresionFE = Calcular(" SELECT tImpresionFE as codigo FROM vTipodocumento WHERE Codigo=(SELECT TTIPODOCUMENTO FROM MDOCUMENTO WHERE tDocumento='" & sDocumento & "')", Cn)
                  xImpresioDE = Mid(sDocumento, 1, 1)
                  
                  'FACTURACION_E_PERU
                  If pais = "000" Then
                     If lFacturacionE Then
                            If lFEOfisis Then
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
                                             'cadenaCodigoHash = Calcular("select CO_HASH as codigo from TCFACT_ELEC where NU_DOCU='" & fDocumento & "' and (TI_DOCU='B' or TI_DOCU ='F')", CnFE)
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
                                xi = 1
                            Else
                                    'VALIDAR RESPUESTA DE CODIGO HASH
                                    If lQRFE Then
                                        Set imageHash.Picture = LoadPicture(ImagenQR(sDocumento))
                                    Else
                                        If lImpresionCodigoBarras Then
                                            Set imageHash.Picture = LoadPicture(lValidaCodBarra(lImpresionCodigoBarras, sDocumento))
                                        Else
                                            cadenaCodigoHash = lValidaCodBarra(lImpresionCodigoBarras, sDocumento)
                                        End If
                                    End If
                                    xi = 1
                                    fDocumento = Mid(sDocumento, 1, 1) + Mid(sDocumento, 4, 3) + Mid(sDocumento, 8, 8)
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
                        If lFacturacionE And xlFacturacionE Then
                        
                                If Generar_Imagen(CnFE, "select imagen from IMAGENCODIGOBARRA where nro_efact='" & fDocumento & "'", "imagen", "\fact.bmp") = True Then
                                    ImprimeFormatoAConsumo
                                    Kill App.Path & "\fact.bmp"
                                Else
                                    ImprimeFormatoAConsumo
                                End If
                        Else
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
                                    ImprimeBoletaT RsImpresion, Calcular("select ttipodocumento  as codigo from mdocumento where tdocumento='" & sDocumento & "'", Cn), imageCab, imagepIE, TimpresionDolaresDelivery
                                End If
                            
                            End If
                        
                     Else
                            If lFacturacionE And xlFacturacionE Then
                                    If Generar_Imagen(CnFE, "select imagen from IMAGENCODIGOBARRA where nro_efact='" & fDocumento & "'", "imagen", "\fact.bmp") = True Then
                                        ImprimeFormatoA
                                        Kill App.Path & "\fact.bmp"
                                    Else
                                        ImprimeFormatoA
                                    End If
                            Else
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
               lReimpresion = False
               LimpiaRs
               Screen.MousePointer = vbDefault
               If flagTipoCambio = True Then: nTC = 0
               
           Case Is = 10  'Detalle de pago
                If RsCabecera.RecordCount = 0 Then
                   Exit Sub
                End If
           
                If RsCabecera.RecordCount > 0 Or RsCabecera!tEstadoDocumento = "02" Then
                   sDocumento = RsCabecera!tDocumento
                   frmDocumentoPago.Show vbModal
                Else
                   MsgBox "No Existe Datos Ingresados", vbExclamation, sMensaje
                End If
                
           Case Is = 11  'Pago del documento
                If RsCabecera.RecordCount > 0 Then
                   If RsCabecera!tEstadoDocumento = "01" Then
                      nTC = Calcular("SELECT nVenta as Codigo From TTIPOCAMBIO WHERE fFecha = '" & Format(RsCabecera!fFecha, "yyyy/mm/dd") & "'", Cn)
                      nCargo = Round(RsCabecera!nVenta, 2)
                      sDocumento = RsCabecera!tDocumento
                      sTurno = RsCabecera!tTurno
                      xTipo = ""
                      sFormulario = ""
                      sTipo = ""
                      frmPago.Show vbModal
                      nPos = RsCabecera.AbsolutePosition
                      RsCabecera.Requery
                      RsCabecera.AbsolutePosition = nPos
                    Else
                      MsgBox "No Es un Documento Emitido", vbExclamation, sMensaje
                    End If
                End If
                
               
    End Select
End Sub

Private Sub cmdProcesa_Click()
   Screen.MousePointer = vbHourglass
    Isql = "SELECT dbo.MDOCUMENTO.tDocumento, dbo.MDOCUMENTO.tCaja, dbo.MDOCUMENTO.tCodigoCliente, dbo.TCLIENTE.tEmpresa AS Cliente, dbo.MDOCUMENTO.tEstadoDocumento, dbo.MDOCUMENTO.nVenta, dbo.MDOCUMENTO.tTipoDocumento, MAX(dbo.MDOCUMENTO.fRegistro) AS fFecha, dbo.MDOCUMENTO.tTurno, dbo.MDOCUMENTO.tUsuario, dbo.vCortesia.Descripcion AS Cortesia, MAX(dbo.TMESA.tResumido) AS Mesa, MAX(dbo.DDOCUMENTO.tCodigoPedido) AS tCodigoPedido, dbo.vEstadoDocumento.Descripcion AS EstadoDocumento, str(dbo.MDOCUMENTO.nVenta,10,2) as xVenta, convert(varchar,  max(dbo.MDOCUMENTO.fRegistro), 103) as xFecha, " & _
           "dbo.MDOCUMENTO.tClientePago, dbo.vCompania.Descripcion AS ClientePago, dbo.MDOCUMENTO.nNeto, dbo.MDOCUMENTO.nPrecioImpuesto1, dbo.MDOCUMENTO.nPrecioImpuesto2, dbo.MDOCUMENTO.nPrecioImpuesto3, dbo.MDOCUMENTO.nRecargo, dbo.MDOCUMENTO.nDescuento, MAX(dbo.MPEDIDO.tTipoPedido) as tTipoPedido, MAX(dbo.MPEDIDO.tObservacion) AS tObservacion, MAX(dbo.vMotorizado.Descripcion) as Motorizado " & _
           "FROM dbo.MPEDIDO LEFT OUTER JOIN dbo.vMotorizado ON dbo.MPEDIDO.tMotorizado = dbo.vMotorizado.Codigo LEFT OUTER JOIN dbo.TMESA ON dbo.MPEDIDO.tMesa = dbo.TMESA.tCodigoMesa RIGHT OUTER JOIN dbo.DDOCUMENTO ON dbo.MPEDIDO.tCodigoPedido = dbo.DDOCUMENTO.tCodigoPedido RIGHT OUTER JOIN dbo.vCompania RIGHT OUTER JOIN dbo.MDOCUMENTO ON dbo.vCompania.Codigo = dbo.MDOCUMENTO.tClientePago LEFT OUTER JOIN " & _
           "dbo.vEstadoDocumento ON dbo.MDOCUMENTO.tEstadoDocumento = dbo.vEstadoDocumento.Codigo LEFT OUTER JOIN dbo.vCortesia ON dbo.MDOCUMENTO.tCortesia = dbo.vCortesia.Codigo LEFT OUTER JOIN dbo.TCLIENTE ON dbo.MDOCUMENTO.tCodigoCliente = dbo.TCLIENTE.tCodigoCliente ON dbo.DDOCUMENTO.tDocumento = dbo.MDOCUMENTO.tDocumento " & _
           "where dbo.MDOCUMENTO.fRegistro >= '" & Format(dtpFecIni.value, "yyyy/MM/dd") & " 00:00' and dbo.MDOCUMENTO.fRegistro <= '" & Format(dtpFecFin.value, "yyyy/MM/dd") & " 23:59' " & _
           "GROUP BY dbo.MDOCUMENTO.tDocumento, dbo.MDOCUMENTO.tCaja, dbo.MDOCUMENTO.tCodigoCliente, dbo.TCLIENTE.tEmpresa, dbo.MDOCUMENTO.tEstadoDocumento, dbo.MDOCUMENTO.nVenta, dbo.MDOCUMENTO.tTipoDocumento, dbo.MDOCUMENTO.tTurno, dbo.MDOCUMENTO.tUsuario, dbo.vCortesia.Descripcion, dbo.vEstadoDocumento.Descripcion, dbo.MDOCUMENTO.tClientePago, dbo.vCompania.Descripcion, dbo.MDOCUMENTO.nNeto, dbo.MDOCUMENTO.nPrecioImpuesto1, " & _
           "dbo.MDOCUMENTO.nPrecioImpuesto2, dbo.MDOCUMENTO.nPrecioImpuesto3, dbo.MDOCUMENTO.nRecargo, dbo.MDOCUMENTO.nDescuento " & _
           "order by dbo.MDOCUMENTO.tDocumento "
          
   Set RsCabecera = Lib.OpenRecordset(Isql, Cn)
   Set grdGrilla.DataSource = RsCabecera
   cmdTexto.Caption = "Registro " & IIf(RsCabecera.RecordCount = 0, 0, RsCabecera.AbsolutePosition) & " de " & RsCabecera.RecordCount
   Screen.MousePointer = vbDefault
End Sub

Private Sub cmdPropina_Click()
    If RsCabecera.RecordCount = 0 Then
       Exit Sub
    End If
    
    If Periodo(RsCabecera!fFecha) Then
       Exit Sub
    End If
    
    If RsCabecera!tEstadoDocumento = "04" Then
       MsgBox "Es un Documento Anulado", vbExclamation, sMensaje
       Exit Sub
    End If
    
    If RsCabecera!tEstadoDocumento = "03" Then
       MsgBox "Es un Documento por Cobrar", vbExclamation, sMensaje
       Exit Sub
    End If
    
    If RsCabecera!tEstadoDocumento = "01" Then
       MsgBox "Es un Documento Emitido", vbExclamation, sMensaje
       Exit Sub
    End If
    
    If RsCabecera!tCaja <> sCaja Then
       MsgBox "El documento pertenece a otra Caja (Caja " & sCaja & ") ", vbCritical, sMensaje
       Exit Sub
    End If
    
    If RsCabecera!tTurno = sTurno Then
       'Password
       If Supervisor("05") = False Then
          MsgBox "Clave no permitida", vbExclamation, sMensaje
          Exit Sub
       End If
    Else
       'Password
       If MsgBox("El Documento es de un turno Anterior, deseas continuar?", vbQuestion + vbOKCancel, sMensaje) = vbCancel Then
          Exit Sub
       End If
       
       If Supervisor("06") = False Then
          MsgBox "Clave no permitida", vbExclamation, sMensaje
          Exit Sub
       End If
    End If
    
    frmCambioPropina.Show
    
    
End Sub

Private Sub dtpFecfin_LostFocus()
   If dtpFecIni.value > dtpFecFin.value Then
      MsgBox "Error en Rango de Fechas", vbCritical, sMensaje
      dtpFecFin.SetFocus
   End If
End Sub

Private Sub Form_Load()
    Screen.MousePointer = vbHourglass
    
    If lMultiCajero Then
       cmdOpcion(0).Caption = "Anula Doc"
    End If
    
    dtpFecIni.value = FechaServidor() - nDias
    dtpFecFin.value = FechaServidor()
    
    If lMultiCajero Then
       cmdOpcion(8).Enabled = False
    End If
    
    
    If sModulo = "INFOREST" Then
       cmdOpcion(0).Enabled = True
       cmdOpcion(5).Enabled = False
       cmdOpcion(0).Enabled = IIf(wInicio, True, False)
       cmdOpcion(8).Enabled = IIf(wInicio, True, False)
       cmdOpcion(5).Enabled = IIf(wInicio, True, False)
    Else
       cmdOpcion(0).Enabled = False
       cmdOpcion(5).Enabled = True
       cmdOpcion(8).Enabled = False
    End If
    
    'Ingrese el Titulo
    Me.Caption = " Documentos Pagados "
    
    If lFacturacionE Then
       cmdOpcion(8).Enabled = False
    End If
    
    Centrar Me
    nColumna = 0
   
    'Ingrese la Vista
'    Isql = "SELECT dbo.MDOCUMENTO.tDocumento, dbo.MDOCUMENTO.tCaja, dbo.MDOCUMENTO.tCodigoCliente, dbo.TCLIENTE.tEmpresa AS Cliente, dbo.MDOCUMENTO.tEstadoDocumento, dbo.MDOCUMENTO.nVenta, dbo.MDOCUMENTO.tTipoDocumento, MAX(dbo.MDOCUMENTO.fRegistro) AS fFecha, dbo.MDOCUMENTO.tTurno, dbo.MDOCUMENTO.tUsuario, dbo.vCortesia.Descripcion AS Cortesia, MAX(dbo.TMESA.tResumido) AS Mesa, MAX(dbo.DDOCUMENTO.tCodigoPedido) AS tCodigoPedido, dbo.vEstadoDocumento.Descripcion AS EstadoDocumento, str(dbo.MDOCUMENTO.nVenta,10,2) as xVenta, convert(varchar,  max(dbo.MDOCUMENTO.fRegistro), 103) as xFecha, " & _
'           "dbo.MDOCUMENTO.tClientePago, dbo.vCompania.Descripcion AS ClientePago, dbo.MDOCUMENTO.nNeto, dbo.MDOCUMENTO.nPrecioImpuesto1, dbo.MDOCUMENTO.nPrecioImpuesto2, dbo.MDOCUMENTO.nPrecioImpuesto3, dbo.MDOCUMENTO.nRecargo, dbo.MDOCUMENTO.nDescuento, MAX(dbo.MPEDIDO.tTipoPedido) as tTipoPedido, MAX(dbo.MPEDIDO.tObservacion) AS tObservacion, MAX(dbo.vMotorizado.Descripcion) as Motorizado, convert(decimal(10,2),isnull(dbo.mdocumento.nVuelto,0)) as nVuelto " & _
'           "FROM dbo.MPEDIDO LEFT OUTER JOIN dbo.vMotorizado ON dbo.MPEDIDO.tMotorizado = dbo.vMotorizado.Codigo LEFT OUTER JOIN dbo.TMESA ON dbo.MPEDIDO.tMesa = dbo.TMESA.tCodigoMesa RIGHT OUTER JOIN dbo.DDOCUMENTO ON dbo.MPEDIDO.tCodigoPedido = dbo.DDOCUMENTO.tCodigoPedido RIGHT OUTER JOIN dbo.vCompania RIGHT OUTER JOIN dbo.MDOCUMENTO ON dbo.vCompania.Codigo = dbo.MDOCUMENTO.tClientePago LEFT OUTER JOIN " & _
'           "dbo.vEstadoDocumento ON dbo.MDOCUMENTO.tEstadoDocumento = dbo.vEstadoDocumento.Codigo LEFT OUTER JOIN dbo.vCortesia ON dbo.MDOCUMENTO.tCortesia = dbo.vCortesia.Codigo LEFT OUTER JOIN dbo.TCLIENTE ON dbo.MDOCUMENTO.tCodigoCliente = dbo.TCLIENTE.tCodigoCliente ON dbo.DDOCUMENTO.tDocumento = dbo.MDOCUMENTO.tDocumento " & _
'           "where dbo.MDOCUMENTO.fRegistro >= '" & Format(dtpFecIni.value, "yyyy/MM/dd") & " 00:00' and dbo.MDOCUMENTO.fRegistro <= '" & Format(dtpFecFin.value, "yyyy/MM/dd") & " 23:59' and dbo.MDOCUMENTO.tCaja='" & sCaja & "' " & _
'           "GROUP BY dbo.MDOCUMENTO.tDocumento, dbo.MDOCUMENTO.tCaja, dbo.MDOCUMENTO.tCodigoCliente, dbo.TCLIENTE.tEmpresa, dbo.MDOCUMENTO.tEstadoDocumento, dbo.MDOCUMENTO.nVenta, dbo.MDOCUMENTO.tTipoDocumento, dbo.MDOCUMENTO.tTurno, dbo.MDOCUMENTO.tUsuario, dbo.vCortesia.Descripcion, dbo.vEstadoDocumento.Descripcion, dbo.MDOCUMENTO.tClientePago, dbo.vCompania.Descripcion, dbo.MDOCUMENTO.nNeto, dbo.MDOCUMENTO.nPrecioImpuesto1, " & _
'           "dbo.MDOCUMENTO.nPrecioImpuesto2, dbo.MDOCUMENTO.nPrecioImpuesto3, dbo.MDOCUMENTO.nRecargo, dbo.MDOCUMENTO.nDescuento,dbo.mdocumento.nVuelto " & _
'           "order by dbo.MDOCUMENTO.tDocumento "
          
    Isql = "exec usp_Inforest_ObtieneDocumentos '" & Format(dtpFecIni.value, "yyyy/MM/dd") & " 00:00', '" & Format(dtpFecFin.value, "yyyy/MM/dd") & " 23:59', '" & sCaja & "'"
    
    Set RsCabecera = Lib.OpenRecordset(Isql, Cn)
    nColumna = 0
   
    'Configuración de la Grilla
    Call ConfGrilla(11, grdGrilla, "Caja", 2, "tCaja", 500, 2, 0, "", _
                                   "Turno", 2, "tTurno", 1100, 0, 0, "", _
                                   "Documento", 2, "tDocumento", 1700, 0, 0, "@-@@@@@-@@@@@@@@@", _
                                   "Fec.Emis", 2, "xFecha", 1100, 0, 0, "dd/mm/yyyy", _
                                   "Mesa", 2, "Mesa", 1100, 0, 0, "", _
                                   "Observacion", 2, "tObservacion", 1800, 0, 0, "", _
                                   "Monto", 2, "xVenta", 1100, 1, 0, "###,###,##0.00", _
                                   "Vuelto", 2, "nVuelto", 1000, 1, 0, "###,##0.00", _
                                   "Cliente", 2, "Cliente", 2400, 0, 0, "", _
                                   "Estado", 2, "EstadoDocumento", 1200, 0, 0, "", _
                                   "Usu.Emis", 2, "tUsuario", 1000, 0, 0, "", _
                                   "Cortesia", 2, "Cortesia", 1900, 0, 0, "")
 
    If nTC = 0 Then: flagTipoCambio = True
    Set grdGrilla.DataSource = RsCabecera
    cmdTexto.Caption = "Registro " & IIf(RsCabecera.RecordCount = 0, 0, RsCabecera.AbsolutePosition) & " de " & RsCabecera.RecordCount
    Screen.MousePointer = vbDefault
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set RsImpresion = Nothing
    Set RsEmision = Nothing
    Set RsCabecera = Nothing
    Set frmDocumentoCorrelativo = Nothing
End Sub

Private Sub grdGrilla_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
    cmdTexto.Caption = "Registro " & IIf(RsCabecera.RecordCount = 0, 0, RsCabecera.AbsolutePosition) & " de " & RsCabecera.RecordCount
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
                        
                        'Miguel rivera----------------------
                        ReporteC.Text44.SetText sMonN
                        ReporteC.Text45.SetText sMonN
                        ReporteC.Text46.SetText sMonN
                        ReporteC.Text47.SetText sMonN
                        ReporteC.Text48.SetText sMonN
                        ReporteC.Text49.SetText sMonN
                        '----------------------------------
                        
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

