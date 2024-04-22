VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{562E3E04-2C31-4ECE-83F4-4017EEE51D40}#8.0#0"; "todg8.ocx"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "msdatlst.ocx"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "tabctl32.ocx"
Begin VB.Form frmCajaDetalle 
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   9375
   ClientLeft      =   2010
   ClientTop       =   1890
   ClientWidth     =   11865
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   8.25
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ForeColor       =   &H80000001&
   Icon            =   "frmCajaDetalle.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   9375
   ScaleWidth      =   11865
   Begin VB.Frame fraDetalle 
      Height          =   8610
      Left            =   1200
      TabIndex        =   23
      Top             =   0
      Width           =   10575
      Begin VB.TextBox txtDetallado 
         Height          =   285
         Left            =   1995
         MaxLength       =   25
         TabIndex        =   143
         Text            =   " "
         Top             =   567
         Width           =   4050
      End
      Begin TabDlg.SSTab tabOpcion 
         Height          =   6210
         Left            =   120
         TabIndex        =   25
         Top             =   2280
         Width           =   10380
         _ExtentX        =   18309
         _ExtentY        =   10954
         _Version        =   393216
         Tabs            =   9
         TabHeight       =   520
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         TabCaption(0)   =   "Activaciones"
         TabPicture(0)   =   "frmCajaDetalle.frx":0442
         Tab(0).ControlEnabled=   -1  'True
         Tab(0).Control(0)=   "Frame6"
         Tab(0).Control(0).Enabled=   0   'False
         Tab(0).Control(1)=   "Frame7"
         Tab(0).Control(1).Enabled=   0   'False
         Tab(0).Control(2)=   "Frame8"
         Tab(0).Control(2).Enabled=   0   'False
         Tab(0).ControlCount=   3
         TabCaption(1)   =   "Obligatoriedades"
         TabPicture(1)   =   "frmCajaDetalle.frx":045E
         Tab(1).ControlEnabled=   0   'False
         Tab(1).Control(0)=   "chkConsumo4"
         Tab(1).Control(1)=   "chkConsumo3"
         Tab(1).Control(2)=   "chkConsumo2"
         Tab(1).Control(3)=   "chkConsumo1"
         Tab(1).Control(4)=   "chkObservacion"
         Tab(1).Control(5)=   "chkObligaPrecuenta"
         Tab(1).Control(6)=   "chkCancelacion"
         Tab(1).Control(7)=   "chkObligaPrinter"
         Tab(1).Control(8)=   "chkComanda"
         Tab(1).ControlCount=   9
         TabCaption(2)   =   "Impresiones"
         TabPicture(2)   =   "frmCajaDetalle.frx":047A
         Tab(2).ControlEnabled=   0   'False
         Tab(2).Control(0)=   "Frame10"
         Tab(2).Control(0).Enabled=   0   'False
         Tab(2).Control(1)=   "Frame11"
         Tab(2).Control(1).Enabled=   0   'False
         Tab(2).Control(2)=   "Frame9"
         Tab(2).Control(2).Enabled=   0   'False
         Tab(2).Control(3)=   "Frame15"
         Tab(2).Control(3).Enabled=   0   'False
         Tab(2).ControlCount=   4
         TabCaption(3)   =   "Emisión de Documentos"
         TabPicture(3)   =   "frmCajaDetalle.frx":0496
         Tab(3).ControlEnabled=   0   'False
         Tab(3).Control(0)=   "grdGrilla"
         Tab(3).Control(1)=   "cmdOpcionGrilla(0)"
         Tab(3).Control(2)=   "cmdOpcionGrilla(1)"
         Tab(3).Control(3)=   "cmdOpcionGrilla(2)"
         Tab(3).Control(4)=   "fraGrilla"
         Tab(3).ControlCount=   5
         TabCaption(4)   =   "Areas de Impresión"
         TabPicture(4)   =   "frmCajaDetalle.frx":04B2
         Tab(4).ControlEnabled=   0   'False
         Tab(4).Control(0)=   "grdAI"
         Tab(4).Control(1)=   "cmdOpcionGrilla(6)"
         Tab(4).Control(2)=   "cmdOpcionGrilla(5)"
         Tab(4).Control(3)=   "cmdOpcionGrilla(7)"
         Tab(4).Control(4)=   "fraArea"
         Tab(4).ControlCount=   5
         TabCaption(5)   =   "Periféricos Adicionales"
         TabPicture(5)   =   "frmCajaDetalle.frx":04CE
         Tab(5).ControlEnabled=   0   'False
         Tab(5).Control(0)=   "Frame5"
         Tab(5).Control(1)=   "Frame4"
         Tab(5).Control(2)=   "Frame2"
         Tab(5).Control(3)=   "Frame3"
         Tab(5).Control(4)=   "Frame1"
         Tab(5).Control(5)=   "Frame12"
         Tab(5).ControlCount=   6
         TabCaption(6)   =   "Multi Area de Producción"
         TabPicture(6)   =   "frmCajaDetalle.frx":04EA
         Tab(6).ControlEnabled=   0   'False
         Tab(6).Control(0)=   "fra1"
         Tab(6).Control(1)=   "chkMulti1"
         Tab(6).Control(2)=   "fra2"
         Tab(6).Control(3)=   "chkMulti2"
         Tab(6).ControlCount=   4
         TabCaption(7)   =   "Areas Chef Control"
         TabPicture(7)   =   "frmCajaDetalle.frx":0506
         Tab(7).ControlEnabled=   0   'False
         Tab(7).Control(0)=   "fraAreaChef"
         Tab(7).Control(1)=   "cmdOpcionGrilla(12)"
         Tab(7).Control(2)=   "cmdOpcionGrilla(13)"
         Tab(7).Control(3)=   "cmdOpcionGrilla(14)"
         Tab(7).Control(4)=   "grdAChef"
         Tab(7).ControlCount=   5
         TabCaption(8)   =   "Imágenes en Documentos"
         TabPicture(8)   =   "frmCajaDetalle.frx":0522
         Tab(8).ControlEnabled=   0   'False
         Tab(8).Control(0)=   "cmdQuitarFotoPie"
         Tab(8).Control(1)=   "cmdQuitarFotoCabecera"
         Tab(8).Control(2)=   "cmdAgregarFotoPie"
         Tab(8).Control(3)=   "cmdAgregarFoto"
         Tab(8).Control(4)=   "dlgFoto"
         Tab(8).Control(5)=   "dlgFotoPie"
         Tab(8).Control(6)=   "imgFotoPie"
         Tab(8).Control(7)=   "imgFoto"
         Tab(8).ControlCount=   8
         Begin VB.Frame Frame15 
            Caption         =   "Comandas"
            Height          =   1215
            Left            =   -74880
            TabIndex        =   243
            Top             =   4800
            Width           =   5460
            Begin VB.CheckBox chkComandaF2 
               Caption         =   "Imprimir Comanda Formato 2"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   225
               Left            =   180
               TabIndex        =   244
               Top             =   240
               Width           =   2595
            End
         End
         Begin VB.CheckBox chkConsumo4 
            Caption         =   "Emisión por consumo en Cuenta Corriente"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   225
            Left            =   -74730
            TabIndex        =   192
            Top             =   3880
            Width           =   3570
         End
         Begin VB.CommandButton cmdQuitarFotoPie 
            Caption         =   "Quitar Imagen"
            Height          =   350
            Left            =   -69480
            TabIndex        =   185
            Top             =   4680
            Width           =   4000
         End
         Begin VB.CommandButton cmdQuitarFotoCabecera 
            Caption         =   "Quitar Imagen"
            Height          =   350
            Left            =   -74520
            TabIndex        =   184
            Top             =   4680
            Width           =   4000
         End
         Begin VB.CommandButton cmdAgregarFotoPie 
            Caption         =   "Editar Imagen Pie"
            Height          =   350
            Left            =   -69480
            TabIndex        =   183
            Top             =   4320
            Width           =   4000
         End
         Begin VB.CommandButton cmdAgregarFoto 
            Caption         =   "Editar Imagen Cabecera"
            Height          =   350
            Left            =   -74520
            TabIndex        =   182
            Top             =   4320
            Width           =   4000
         End
         Begin VB.Frame fraAreaChef 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   3675
            Left            =   -74880
            TabIndex        =   175
            Top             =   1320
            Width           =   9830
            Begin VB.CommandButton cmdOpcionGrilla 
               Caption         =   "Grabar"
               Height          =   645
               Index           =   10
               Left            =   1515
               Picture         =   "frmCajaDetalle.frx":053E
               Style           =   1  'Graphical
               TabIndex        =   178
               Top             =   1650
               Width           =   1215
            End
            Begin VB.CommandButton cmdOpcionGrilla 
               Caption         =   "Cancelar"
               Height          =   645
               Index           =   11
               Left            =   2775
               Picture         =   "frmCajaDetalle.frx":0A70
               Style           =   1  'Graphical
               TabIndex        =   177
               Top             =   1650
               Width           =   1215
            End
            Begin VB.CheckBox chkAreaChef 
               Alignment       =   1  'Right Justify
               Caption         =   "Area Chef :  "
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   195
               Left            =   300
               TabIndex        =   176
               Top             =   1005
               Width           =   1215
            End
            Begin MSDataListLib.DataCombo cboAreaChef 
               Height          =   315
               Left            =   1305
               TabIndex        =   179
               Top             =   465
               Width           =   2970
               _ExtentX        =   5239
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
            Begin VB.Label Label 
               AutoSize        =   -1  'True
               Caption         =   "Area :"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   195
               Index           =   21
               Left            =   735
               TabIndex        =   180
               Top             =   525
               Width           =   420
            End
         End
         Begin VB.CommandButton cmdOpcionGrilla 
            Caption         =   "Eliminar"
            Height          =   645
            Index           =   12
            Left            =   -66960
            Picture         =   "frmCajaDetalle.frx":0B72
            Style           =   1  'Graphical
            TabIndex        =   173
            Top             =   2880
            Width           =   1215
         End
         Begin VB.CommandButton cmdOpcionGrilla 
            Caption         =   "Agregar"
            Height          =   645
            Index           =   13
            Left            =   -66960
            Picture         =   "frmCajaDetalle.frx":0C74
            Style           =   1  'Graphical
            TabIndex        =   172
            Top             =   1320
            Width           =   1215
         End
         Begin VB.CommandButton cmdOpcionGrilla 
            Caption         =   "Modificar"
            Height          =   645
            Index           =   14
            Left            =   -66960
            Picture         =   "frmCajaDetalle.frx":11A6
            Style           =   1  'Graphical
            TabIndex        =   171
            Top             =   2085
            Width           =   1215
         End
         Begin VB.CheckBox chkMulti2 
            Caption         =   "Areas de Producción por Sub Grupos"
            Height          =   255
            Left            =   -74640
            TabIndex        =   158
            Top             =   1920
            Width           =   3855
         End
         Begin VB.Frame fra2 
            Enabled         =   0   'False
            Height          =   2775
            Left            =   -74760
            TabIndex        =   159
            Top             =   1920
            Width           =   9375
            Begin VB.Frame fraAreaProduccion 
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   2460
               Left            =   120
               TabIndex        =   160
               Top             =   240
               Width           =   9120
               Begin VB.CommandButton cmdOpcionGrilla 
                  Caption         =   "Cancelar"
                  Height          =   645
                  Index           =   16
                  Left            =   6120
                  Picture         =   "frmCajaDetalle.frx":16D8
                  Style           =   1  'Graphical
                  TabIndex        =   162
                  Top             =   600
                  Width           =   1215
               End
               Begin VB.CommandButton cmdOpcionGrilla 
                  Caption         =   "Grabar"
                  Height          =   645
                  Index           =   15
                  Left            =   4920
                  Picture         =   "frmCajaDetalle.frx":17DA
                  Style           =   1  'Graphical
                  TabIndex        =   161
                  Top             =   600
                  Width           =   1215
               End
               Begin MSDataListLib.DataCombo cboSubGrupo 
                  Height          =   315
                  Left            =   1665
                  TabIndex        =   163
                  Top             =   480
                  Width           =   2970
                  _ExtentX        =   5239
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
               Begin MSDataListLib.DataCombo cboAreaProd 
                  Height          =   315
                  Left            =   1665
                  TabIndex        =   164
                  Top             =   945
                  Width           =   2970
                  _ExtentX        =   5239
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
               Begin VB.Label Label 
                  AutoSize        =   -1  'True
                  Caption         =   "Area Producción :"
                  BeginProperty Font 
                     Name            =   "MS Sans Serif"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   195
                  Index           =   23
                  Left            =   255
                  TabIndex        =   166
                  Top             =   1005
                  Width           =   1275
               End
               Begin VB.Label Label 
                  AutoSize        =   -1  'True
                  Caption         =   "Sub Grupo :"
                  BeginProperty Font 
                     Name            =   "MS Sans Serif"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   195
                  Index           =   22
                  Left            =   540
                  TabIndex        =   165
                  Top             =   525
                  Width           =   855
               End
            End
            Begin TrueOleDBGrid80.TDBGrid grdGrillaSubgrupos 
               Height          =   2295
               Left            =   120
               TabIndex        =   167
               Top             =   360
               Width           =   7455
               _ExtentX        =   13150
               _ExtentY        =   4048
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
            Begin VB.CommandButton cmdOpcionGrilla 
               Caption         =   "Modificar"
               Height          =   645
               Index           =   19
               Left            =   7680
               Picture         =   "frmCajaDetalle.frx":1D0C
               Style           =   1  'Graphical
               TabIndex        =   168
               Top             =   1200
               Width           =   1215
            End
            Begin VB.CommandButton cmdOpcionGrilla 
               Caption         =   "Agregar"
               Height          =   645
               Index           =   18
               Left            =   7680
               Picture         =   "frmCajaDetalle.frx":223E
               Style           =   1  'Graphical
               TabIndex        =   169
               Top             =   480
               Width           =   1215
            End
            Begin VB.CommandButton cmdOpcionGrilla 
               Caption         =   "Eliminar"
               Height          =   645
               Index           =   17
               Left            =   7680
               Picture         =   "frmCajaDetalle.frx":2770
               Style           =   1  'Graphical
               TabIndex        =   170
               Top             =   1920
               Width           =   1215
            End
         End
         Begin VB.CheckBox chkMulti1 
            Caption         =   "Una sola Area de Producción"
            Height          =   375
            Left            =   -74640
            TabIndex        =   154
            Top             =   960
            Width           =   3015
         End
         Begin VB.Frame fra1 
            Enabled         =   0   'False
            Height          =   735
            Left            =   -74760
            TabIndex        =   155
            Top             =   1080
            Width           =   9375
            Begin MSDataListLib.DataCombo cboAreaProduccion 
               Height          =   315
               Left            =   1740
               TabIndex        =   156
               Top             =   360
               Width           =   2595
               _ExtentX        =   4577
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
            Begin VB.Label Label 
               AutoSize        =   -1  'True
               Caption         =   "Área de Producción :"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   195
               Index           =   16
               Left            =   120
               TabIndex        =   157
               Top             =   420
               Width           =   1500
            End
         End
         Begin VB.Frame Frame12 
            Caption         =   " Enlace SIAB "
            Height          =   735
            Left            =   -69285
            TabIndex        =   139
            Top             =   2520
            Width           =   4185
            Begin VB.CheckBox chkSiab 
               Caption         =   "Caja conectada al POS SIAB"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   195
               Left            =   225
               TabIndex        =   140
               Top             =   375
               Width           =   3825
            End
         End
         Begin VB.Frame Frame9 
            Caption         =   " Otros Documentos "
            Height          =   2055
            Left            =   -69330
            TabIndex        =   135
            Top             =   3915
            Width           =   4290
            Begin VB.CheckBox chkCodigoReciboIngreso 
               Caption         =   "Impresión Código de Barras en Recibo de Ingreso"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   225
               Left            =   180
               TabIndex        =   189
               Top             =   975
               Width           =   4020
            End
            Begin VB.TextBox txtLimiteReimpresion 
               Alignment       =   1  'Right Justify
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   285
               Left            =   720
               TabIndex        =   148
               Text            =   "0"
               Top             =   1305
               Width           =   600
            End
            Begin VB.CheckBox chkValor 
               Caption         =   "Impresión Valorizada de las Cortesías"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   225
               Left            =   180
               TabIndex        =   138
               Top             =   675
               Width           =   3570
            End
            Begin VB.CheckBox chkCambioMesa 
               Caption         =   "Impresion de Cambio de Mesa"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   225
               Left            =   180
               TabIndex        =   137
               Top             =   360
               Width           =   3570
            End
            Begin VB.Label Label4 
               Caption         =   "Permite"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   195
               Left            =   120
               TabIndex        =   150
               Top             =   1350
               Width           =   555
            End
            Begin VB.Label Label5 
               Caption         =   "reimpresiones por Pedido (0 sin limite)"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   195
               Left            =   1410
               TabIndex        =   149
               Top             =   1350
               Width           =   2625
            End
         End
         Begin VB.Frame fraArea 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   3915
            Left            =   -74865
            TabIndex        =   48
            Top             =   1020
            Width           =   9855
            Begin VB.CommandButton cmdOpcionGrilla 
               Caption         =   "Grabar"
               Height          =   645
               Index           =   8
               Left            =   7155
               Picture         =   "frmCajaDetalle.frx":2872
               Style           =   1  'Graphical
               TabIndex        =   50
               Top             =   2970
               Width           =   1215
            End
            Begin VB.CommandButton cmdOpcionGrilla 
               Caption         =   "Cancelar"
               Height          =   645
               Index           =   9
               Left            =   8415
               Picture         =   "frmCajaDetalle.frx":2DA4
               Style           =   1  'Graphical
               TabIndex        =   49
               Top             =   2970
               Width           =   1215
            End
            Begin MSDataListLib.DataCombo cboArea 
               Height          =   315
               Left            =   1305
               TabIndex        =   51
               Top             =   465
               Width           =   2970
               _ExtentX        =   5239
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
            Begin MSDataListLib.DataCombo cboImpArea 
               Height          =   315
               Left            =   1305
               TabIndex        =   52
               Top             =   945
               Width           =   2970
               _ExtentX        =   5239
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
            Begin VB.Label Label 
               AutoSize        =   -1  'True
               Caption         =   "Area :"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   195
               Index           =   11
               Left            =   735
               TabIndex        =   54
               Top             =   525
               Width           =   420
            End
            Begin VB.Label Label 
               AutoSize        =   -1  'True
               Caption         =   "Impresora :"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   195
               Index           =   9
               Left            =   375
               TabIndex        =   53
               Top             =   1005
               Width           =   780
            End
         End
         Begin VB.Frame Frame11 
            Caption         =   " Documentos "
            Height          =   2700
            Left            =   -69330
            TabIndex        =   119
            Top             =   975
            Width           =   4290
            Begin VB.CheckBox chkMotDesc 
               Caption         =   "Impresión de Motivo de Descuento en Documentos"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   225
               Left            =   120
               TabIndex        =   240
               Top             =   2160
               Width           =   4050
            End
            Begin VB.CheckBox chkObservacionCabDoc 
               Caption         =   "Impresión de Observacion en cabecera del Documento"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   345
               Left            =   120
               TabIndex        =   213
               Top             =   1560
               Width           =   3930
            End
            Begin VB.CheckBox chkDescripcionAlternativa 
               Caption         =   "Impresión Descripción Alternativa en Documentos"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   225
               Left            =   120
               TabIndex        =   131
               Top             =   1920
               Visible         =   0   'False
               Width           =   3930
            End
            Begin VB.CheckBox chkPropiedadDocumento 
               Caption         =   "Impresión de Propiedades en Documentos"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   225
               Left            =   120
               TabIndex        =   130
               Top             =   840
               Width           =   3870
            End
            Begin VB.CheckBox chkComboDocumento 
               Caption         =   "Impresion de Combos en Documentos"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   225
               Left            =   120
               TabIndex        =   129
               Top             =   600
               Width           =   3570
            End
            Begin VB.CheckBox chkDocumentoAgrupado 
               Caption         =   "Impresion Agrupada de Items de Documentos"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   225
               Left            =   120
               TabIndex        =   128
               Top             =   315
               Width           =   3615
            End
            Begin VB.CheckBox chkObservacionDocumento 
               Caption         =   "Impresión de Observaciones en Documentos (Combos y Propiedades)"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   465
               Left            =   120
               TabIndex        =   127
               Top             =   1080
               Width           =   3810
            End
         End
         Begin VB.Frame Frame10 
            Caption         =   " Precuenta "
            Height          =   3780
            Left            =   -74865
            TabIndex        =   115
            Top             =   975
            Width           =   5460
            Begin VB.CheckBox chkImpPropina 
               Caption         =   "Solicitar Propina en Precuenta"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   225
               Left            =   180
               TabIndex        =   242
               Top             =   3000
               Width           =   3570
            End
            Begin VB.CheckBox chkPrecuentaNoValorizada 
               Caption         =   "Impresión de Precuenta no valorizada"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   225
               Left            =   180
               TabIndex        =   193
               Top             =   2760
               Width           =   3570
            End
            Begin VB.CheckBox chkImprimeImagCabPrecuenta 
               Caption         =   "Impresión Imagen en Cabecera de Documento en Precuenta"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   195
               Left            =   180
               TabIndex        =   124
               Top             =   2040
               Width           =   4995
            End
            Begin VB.CheckBox chkImprimeImagPiePrecuenta 
               Caption         =   "Impresión Imagen en Pie de Documento en Precuenta"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   225
               Left            =   180
               TabIndex        =   125
               Top             =   2280
               Width           =   4395
            End
            Begin VB.CheckBox chkBloqueaPrecuenta 
               Caption         =   "No Permitir Emisión de Precuentas"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   225
               Left            =   180
               TabIndex        =   126
               Top             =   2520
               Width           =   3570
            End
            Begin VB.TextBox txtLimitePrecuenta 
               Alignment       =   1  'Right Justify
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   285
               Left            =   840
               TabIndex        =   145
               Text            =   "0"
               Top             =   3360
               Width           =   600
            End
            Begin VB.CheckBox chkEquivaPrecuenta 
               Caption         =   "Impresión Equivalencia Dolares en Precuenta"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   225
               Left            =   180
               TabIndex        =   123
               Top             =   1800
               Width           =   3570
            End
            Begin VB.CheckBox chkPrecioNetoPrecuenta 
               Caption         =   "Impresión de Precio Neto en Precuenta"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   225
               Left            =   180
               TabIndex        =   121
               Top             =   1320
               Width           =   3570
            End
            Begin VB.CheckBox chkImpuestoPrecuenta 
               Caption         =   "Impresion Impuestos desglos. en Precuenta"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   195
               Left            =   180
               TabIndex        =   122
               Top             =   1560
               Width           =   3570
            End
            Begin VB.CheckBox chkPropiedadPrecuenta 
               Caption         =   "Impresión de Propiedades  en Precuentas"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   225
               Left            =   180
               TabIndex        =   118
               Top             =   820
               Width           =   3990
            End
            Begin VB.CheckBox chkAgrupada 
               Caption         =   "Impresion Agrupada de Items de Precuenta"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   225
               Left            =   180
               TabIndex        =   116
               Top             =   315
               Width           =   3480
            End
            Begin VB.CheckBox chkComboPrecuenta 
               Caption         =   "Impresion de Combos en Precuenta"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   225
               Left            =   180
               TabIndex        =   117
               Top             =   560
               Width           =   3570
            End
            Begin VB.CheckBox chkObservacionPrecuenta 
               Caption         =   "Impresión de Observaciones en Precuentas"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   225
               Left            =   180
               TabIndex        =   120
               Top             =   1080
               Width           =   3990
            End
            Begin VB.Label Label2 
               Caption         =   "precuentas por Pedido (0 sin limite)"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   195
               Left            =   1560
               TabIndex        =   147
               Top             =   3480
               Width           =   2580
            End
            Begin VB.Label Label3 
               Caption         =   "Permite"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   195
               Left            =   180
               TabIndex        =   146
               Top             =   3405
               Width           =   555
            End
         End
         Begin VB.Frame Frame8 
            Caption         =   " Activaciones Formas de Venta"
            Height          =   2865
            Left            =   5265
            TabIndex        =   106
            Top             =   3270
            Width           =   4695
            Begin VB.CheckBox chkCajaContingencia 
               Caption         =   "Activar Caja Contingencia"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00000080&
               Height          =   195
               Left            =   120
               TabIndex        =   241
               Top             =   2400
               Width           =   2520
            End
            Begin VB.CheckBox chkMesa247 
               Caption         =   "Caja Mesa247"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00000080&
               Height          =   225
               Left            =   135
               TabIndex        =   205
               Top             =   2100
               Width           =   2715
            End
            Begin VB.CheckBox chkWebAp 
               Caption         =   "Caja WebAp"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00000080&
               Height          =   225
               Left            =   135
               TabIndex        =   191
               Top             =   1830
               Width           =   2715
            End
            Begin VB.CheckBox chkCajaMobile 
               Caption         =   "Caja Mobile"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00000080&
               Height          =   225
               Left            =   135
               TabIndex        =   108
               Top             =   1560
               Width           =   2715
            End
            Begin VB.CheckBox chkCD 
               Caption         =   "Activa Caja Central de Delivery"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00000080&
               Height          =   225
               Left            =   135
               TabIndex        =   112
               Top             =   250
               Width           =   3570
            End
            Begin VB.CheckBox chkMultiCajero 
               Caption         =   "Activa Multicajero Caja Rápida"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00000080&
               Height          =   225
               Left            =   135
               TabIndex        =   111
               Top             =   800
               Width           =   3570
            End
            Begin VB.CheckBox chkMCPV 
               Caption         =   "Activa Multicajero Salon"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00000080&
               Height          =   225
               Left            =   135
               TabIndex        =   110
               Top             =   1050
               Width           =   3570
            End
            Begin VB.CheckBox chkCCVOX 
               Caption         =   "Activa Caja Delivery con CCVOX"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00000080&
               Height          =   225
               Left            =   135
               TabIndex        =   109
               Top             =   510
               Width           =   3570
            End
            Begin VB.CheckBox chkCompatibilidadTVS 
               Caption         =   "Permite Compatibilidad con TVS"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00000080&
               Height          =   225
               Left            =   135
               TabIndex        =   107
               Top             =   1300
               Width           =   2715
            End
         End
         Begin VB.Frame Frame7 
            Caption         =   " Activaciones Básicas "
            Height          =   5220
            Left            =   225
            TabIndex        =   96
            Top             =   910
            Width           =   5010
            Begin VB.CheckBox chkClaveEnvio 
               Caption         =   "Solicitar clave para envio a Producción"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   225
               Left            =   120
               TabIndex        =   202
               ToolTipText     =   "Activado: Pedidos con Mesa/ Desactivado: Pedidos con Cliente"
               Top             =   4850
               Width           =   3810
            End
            Begin VB.CheckBox chkBuscaPedidoFiltrarMesa 
               Caption         =   "Buscar Pedido: Filtrar Pedidos con Mesa asignada en Grilla"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   225
               Left            =   135
               TabIndex        =   196
               ToolTipText     =   "Activado: Pedidos con Mesa/ Desactivado: Pedidos con Cliente"
               Top             =   4600
               Width           =   4770
            End
            Begin VB.CheckBox chkBuscaPedidoVisualizaGrilla 
               Caption         =   "Buscar Pedido: Ingreso directo a Visualización en Grilla"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   225
               Left            =   135
               TabIndex        =   195
               ToolTipText     =   "Activado: Acceso a visualizar grilla / Desactivado: Acceso a visualizar mapa de mesas"
               Top             =   4400
               Width           =   4770
            End
            Begin VB.CheckBox chkPagoRapidoMod 
               Caption         =   "Ingreso a Pago Rápido desde Modifica Pago"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   255
               Left            =   135
               TabIndex        =   190
               Top             =   3675
               Width           =   3735
            End
            Begin VB.CheckBox chkBuscaPedido 
               Caption         =   "Búsqueda Predeterminada por Número de Pedido al Transferir"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   225
               Left            =   135
               TabIndex        =   103
               ToolTipText     =   "Activado: Por Número Pedido / Desactivado: Por Mesa"
               Top             =   1950
               Width           =   4770
            End
            Begin VB.CheckBox chkAccesoDespachoPedido 
               Caption         =   "Activa Acceso a Despachos en Central de Pedidos"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   225
               Left            =   135
               TabIndex        =   187
               Top             =   4170
               Width           =   4770
            End
            Begin VB.CheckBox chkHuella 
               Caption         =   "Activa Ingreso Directo a Huella Digital"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   225
               Left            =   135
               TabIndex        =   181
               Top             =   3930
               Width           =   3570
            End
            Begin VB.CheckBox chkCajaRapida 
               Caption         =   "Ingreso Directo a la Caja Rápida"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   225
               Left            =   135
               TabIndex        =   134
               Top             =   2985
               Width           =   3570
            End
            Begin VB.CheckBox chkPagoRapido 
               Caption         =   "Ingreso a Pago Rápido desde Caja Rápida"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   225
               Left            =   135
               TabIndex        =   133
               Top             =   3195
               Width           =   3570
            End
            Begin VB.CheckBox chkPagoRapidopv 
               Caption         =   "Ingreso a Pago Rápido desde Punto de Venta"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   225
               Left            =   135
               TabIndex        =   132
               Top             =   3450
               Width           =   3570
            End
            Begin VB.CheckBox chkPreCuenta 
               Caption         =   "Activa el poder cambiar de impresora de PreCuentas"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   225
               Left            =   135
               TabIndex        =   114
               Top             =   2475
               Width           =   4560
            End
            Begin VB.CheckBox chkDisgrega 
               Caption         =   "Activa el disgregar el Producto en dos partes"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   225
               Left            =   135
               TabIndex        =   113
               Top             =   2730
               Width           =   3930
            End
            Begin VB.CheckBox chkFiltroTipoPedido 
               Caption         =   "Activa  la Importación de Pedidos por Canales de Venta"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   225
               Left            =   135
               TabIndex        =   104
               Top             =   1440
               Width           =   4560
            End
            Begin VB.CheckBox chkAdicion 
               Caption         =   "Activa las Transferencias (Importar Pedidos)"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   225
               Left            =   135
               TabIndex        =   102
               Top             =   1710
               Width           =   3570
            End
            Begin VB.CheckBox chkModificaTipoPedido 
               Caption         =   "Activa el poder Modificar el Tipo de Pedido"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   225
               Left            =   135
               TabIndex        =   105
               Top             =   2205
               Width           =   3705
            End
            Begin VB.CheckBox chkOrden 
               Caption         =   "Activa el Control de Enumeración Automática"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   225
               Left            =   135
               TabIndex        =   101
               Top             =   195
               Width           =   3570
            End
            Begin VB.CheckBox chkElimina 
               Caption         =   "Pide un Motivo de Eliminación en el Producto"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   225
               Left            =   135
               TabIndex        =   100
               Top             =   1185
               Width           =   3570
            End
            Begin VB.CheckBox chkEliminaC 
               Caption         =   "Pide un Motivo de Eliminación en el Pedido"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   225
               Left            =   135
               TabIndex        =   99
               Top             =   915
               Width           =   3570
            End
            Begin VB.CheckBox chkDirecto 
               Caption         =   "Activa el Control de Envíos Directos"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   225
               Left            =   135
               TabIndex        =   98
               Top             =   450
               Width           =   3570
            End
            Begin VB.CheckBox chkVComanda 
               Caption         =   "Activa el ingreso de Comandas Manuales"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   225
               Left            =   135
               TabIndex        =   97
               Top             =   680
               Width           =   3570
            End
         End
         Begin VB.Frame Frame6 
            Caption         =   " Activa Password "
            Height          =   2325
            Left            =   5265
            TabIndex        =   89
            Top             =   910
            Width           =   4695
            Begin VB.CheckBox chkPassOtrosPagos 
               Caption         =   "Activa Password para Otras Formas de Pago"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   225
               Left            =   135
               TabIndex        =   245
               Top             =   1960
               Width           =   3705
            End
            Begin VB.CheckBox chkPasswordTransferencia 
               Caption         =   "Activa Password para transferencias"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   225
               Left            =   135
               TabIndex        =   95
               Top             =   1143
               Width           =   3705
            End
            Begin VB.CheckBox chkPassword 
               Caption         =   "Activa Password de Eliminacion en el Producto"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   225
               Left            =   135
               TabIndex        =   94
               Top             =   591
               Width           =   3660
            End
            Begin VB.CheckBox chkPasswordC 
               Caption         =   "Activa Password de Eliminación de Pedidos"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   225
               Left            =   135
               TabIndex        =   93
               Top             =   315
               Width           =   3570
            End
            Begin VB.CheckBox chkObligaCierre 
               Caption         =   "Activa Password al Cierre del Turno"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   225
               Left            =   135
               TabIndex        =   92
               Top             =   867
               Width           =   3570
            End
            Begin VB.CheckBox chkPasswordImportar 
               Caption         =   "Activa Password para Importar Pedidos"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   225
               Left            =   135
               TabIndex        =   91
               Top             =   1695
               Width           =   3705
            End
            Begin VB.CheckBox chkPasswordPorCobrar 
               Caption         =   "Activa Password para Enviar a Cuentas Por Cobrar"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   225
               Left            =   135
               TabIndex        =   90
               Top             =   1419
               Width           =   3945
            End
         End
         Begin VB.Frame fraGrilla 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   4995
            Left            =   -74865
            TabIndex        =   34
            Top             =   960
            Width           =   9825
            Begin VB.CheckBox chkMayorCero 
               Alignment       =   1  'Right Justify
               Caption         =   "Impresion de platos con monto mayor ""0"":"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   435
               Left            =   7680
               TabIndex        =   239
               Top             =   1560
               Width           =   2055
            End
            Begin VB.CheckBox chkCodProdDes 
               Alignment       =   1  'Right Justify
               Caption         =   "Impr. CodigoProducto, Descuento/unitario:"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   435
               Left            =   7680
               TabIndex        =   238
               Top             =   1080
               Width           =   2055
            End
            Begin VB.Frame Frame14 
               Caption         =   "Desglose en Impresión :"
               Height          =   1815
               Left            =   5280
               TabIndex        =   214
               Top             =   240
               Width           =   2295
               Begin VB.CheckBox chkImpuesto1 
                  Caption         =   "Check1"
                  BeginProperty Font 
                     Name            =   "MS Sans Serif"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   195
                  Left            =   120
                  TabIndex        =   219
                  Top             =   240
                  Visible         =   0   'False
                  Width           =   1725
               End
               Begin VB.CheckBox chkImpuesto2 
                  Caption         =   "Check2"
                  BeginProperty Font 
                     Name            =   "MS Sans Serif"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   195
                  Left            =   120
                  TabIndex        =   218
                  Top             =   480
                  Visible         =   0   'False
                  Width           =   1725
               End
               Begin VB.CheckBox chkImpuesto3 
                  Caption         =   "Check3"
                  BeginProperty Font 
                     Name            =   "MS Sans Serif"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   195
                  Left            =   120
                  TabIndex        =   217
                  Top             =   720
                  Visible         =   0   'False
                  Width           =   1725
               End
               Begin VB.CheckBox chkopGravInaf 
                  Caption         =   "Op. Gravada, Inafecta"
                  BeginProperty Font 
                     Name            =   "MS Sans Serif"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   195
                  Left            =   120
                  TabIndex        =   216
                  Top             =   960
                  Width           =   1965
               End
               Begin VB.Label Label 
                  AutoSize        =   -1  'True
                  Caption         =   "Desglose en Impresión :"
                  BeginProperty Font 
                     Name            =   "MS Sans Serif"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   195
                  Index           =   4
                  Left            =   360
                  TabIndex        =   215
                  Top             =   1320
                  Visible         =   0   'False
                  Width           =   1695
               End
            End
            Begin VB.CheckBox chkImpResumido 
               Alignment       =   1  'Right Justify
               Caption         =   "Impr. Detallado:"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   315
               Left            =   7680
               TabIndex        =   212
               Top             =   760
               Width           =   2055
            End
            Begin VB.TextBox txtPrefijoEnlace 
               Height          =   285
               Left            =   8880
               MaxLength       =   1
               TabIndex        =   204
               Top             =   2235
               Width           =   795
            End
            Begin VB.TextBox txtAutorizacion 
               Height          =   285
               Left            =   2040
               MaxLength       =   20
               TabIndex        =   197
               Top             =   2760
               Width           =   3120
            End
            Begin VB.CheckBox chkFacturacionOfisis 
               Alignment       =   1  'Right Justify
               Caption         =   "Documento Electrónico Ofisis :"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   255
               Left            =   5520
               TabIndex        =   194
               Top             =   2880
               Width           =   2565
            End
            Begin VB.CheckBox chkFacturacionE 
               Alignment       =   1  'Right Justify
               Caption         =   "Facturación Electrónica:"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   195
               Left            =   5520
               TabIndex        =   188
               Top             =   2280
               Width           =   2055
            End
            Begin VB.CheckBox chLImprimeImageCab 
               Alignment       =   1  'Right Justify
               Caption         =   "En Cabecera"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   195
               Left            =   2040
               TabIndex        =   6
               Top             =   2280
               Width           =   1335
            End
            Begin VB.CheckBox chLImprimeImagePie 
               Alignment       =   1  'Right Justify
               Caption         =   "En Pie"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   195
               Left            =   4080
               TabIndex        =   7
               Top             =   2280
               Width           =   1095
            End
            Begin VB.CommandButton cmdOpcionGrilla 
               Caption         =   "Cancelar"
               Height          =   645
               Index           =   4
               Left            =   8370
               Picture         =   "frmCajaDetalle.frx":2EA6
               Style           =   1  'Graphical
               TabIndex        =   12
               Top             =   4230
               Width           =   1215
            End
            Begin VB.TextBox txtSerie 
               Height          =   285
               Left            =   2040
               MaxLength       =   5
               TabIndex        =   1
               Top             =   735
               Width           =   1410
            End
            Begin VB.TextBox txtDescripcion 
               Height          =   285
               Left            =   2040
               MaxLength       =   50
               TabIndex        =   5
               Top             =   1815
               Width           =   3120
            End
            Begin VB.CheckBox chkResumen 
               Alignment       =   1  'Right Justify
               Caption         =   "Resumen :"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   195
               Left            =   7680
               TabIndex        =   8
               Top             =   240
               Width           =   2055
            End
            Begin VB.CommandButton cmdOpcionGrilla 
               Caption         =   "Grabar"
               Height          =   645
               Index           =   3
               Left            =   7065
               Picture         =   "frmCajaDetalle.frx":2FA8
               Style           =   1  'Graphical
               TabIndex        =   11
               Top             =   4230
               Width           =   1215
            End
            Begin VB.CheckBox chkDocEquivDolares 
               Alignment       =   1  'Right Justify
               Caption         =   "Impresión Equi. Dolares"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   315
               Left            =   7680
               TabIndex        =   9
               Top             =   480
               Width           =   2055
            End
            Begin VB.TextBox txtCorrelativo 
               Height          =   285
               Left            =   3525
               MaxLength       =   9
               TabIndex        =   2
               Top             =   735
               Width           =   1635
            End
            Begin VB.TextBox txtSerie2 
               Height          =   285
               Left            =   3315
               MaxLength       =   3
               TabIndex        =   37
               Text            =   "000"
               Top             =   735
               Width           =   645
            End
            Begin VB.TextBox txtCorrelativo2 
               Height          =   285
               Left            =   3900
               MaxLength       =   9
               TabIndex        =   35
               Text            =   "00000000"
               Top             =   735
               Width           =   1140
            End
            Begin MSDataListLib.DataCombo cboLocal 
               Height          =   315
               Left            =   2040
               TabIndex        =   36
               Top             =   735
               Width           =   1185
               _ExtentX        =   2090
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
            Begin MSDataListLib.DataCombo cboTipoDocumento 
               Height          =   315
               Left            =   2040
               TabIndex        =   0
               Top             =   315
               Width           =   3120
               _ExtentX        =   5503
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
               Height          =   315
               Left            =   2040
               TabIndex        =   3
               Top             =   1065
               Width           =   3120
               _ExtentX        =   5503
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
            Begin MSDataListLib.DataCombo cboFormulario 
               Height          =   315
               Left            =   2040
               TabIndex        =   4
               Top             =   1440
               Width           =   3120
               _ExtentX        =   5503
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
            Begin MSComCtl2.DTPicker dtpFechaInicio 
               Height          =   345
               Left            =   2040
               TabIndex        =   200
               Top             =   3120
               Width           =   2130
               _ExtentX        =   3757
               _ExtentY        =   609
               _Version        =   393216
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               CustomFormat    =   "dd/MM/yyyy"
               Format          =   88473603
               CurrentDate     =   37795
            End
            Begin MSComCtl2.DTPicker dtpFechaCaducida 
               Height          =   345
               Left            =   2040
               TabIndex        =   201
               Top             =   3600
               Width           =   2130
               _ExtentX        =   3757
               _ExtentY        =   609
               _Version        =   393216
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               CustomFormat    =   "dd/MM/yyyy"
               Format          =   88473603
               CurrentDate     =   37795
            End
            Begin VB.Frame Frame13 
               Caption         =   "(solo aplica en el enlace con OFISIS)"
               ForeColor       =   &H00800000&
               Height          =   1455
               Left            =   5400
               TabIndex        =   206
               Top             =   2640
               Width           =   4335
               Begin VB.TextBox txtCompVenta 
                  Height          =   285
                  Left            =   2160
                  MaxLength       =   50
                  TabIndex        =   210
                  Top             =   960
                  Width           =   2040
               End
               Begin VB.TextBox txtFormVenta 
                  Height          =   285
                  Left            =   2160
                  MaxLength       =   50
                  TabIndex        =   209
                  Top             =   555
                  Width           =   2040
               End
               Begin VB.Label Label 
                  AutoSize        =   -1  'True
                  Caption         =   "Cod. Comprobante Venta :"
                  BeginProperty Font 
                     Name            =   "MS Sans Serif"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   195
                  Index           =   30
                  Left            =   120
                  TabIndex        =   208
                  Top             =   1005
                  Width           =   1875
               End
               Begin VB.Label Label 
                  AutoSize        =   -1  'True
                  Caption         =   "Cod. Formulario Venta :"
                  BeginProperty Font 
                     Name            =   "MS Sans Serif"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   195
                  Index           =   29
                  Left            =   360
                  TabIndex        =   207
                  Top             =   600
                  Width           =   1650
               End
            End
            Begin VB.Label Label 
               AutoSize        =   -1  'True
               Caption         =   "Prefijo Enlace:"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   195
               Index           =   28
               Left            =   7800
               TabIndex        =   203
               Top             =   2280
               Width           =   1020
            End
            Begin VB.Label Label 
               AutoSize        =   -1  'True
               Caption         =   "Fecha Caducidad :"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   195
               Index           =   27
               Left            =   600
               TabIndex        =   199
               Top             =   3675
               Width           =   1350
            End
            Begin VB.Label Label 
               AutoSize        =   -1  'True
               Caption         =   "Fecha Inicio :"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   195
               Index           =   26
               Left            =   960
               TabIndex        =   198
               Top             =   3195
               Width           =   960
            End
            Begin VB.Label Label 
               AutoSize        =   -1  'True
               Caption         =   "Impresión de Imágenes  :"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   195
               Index           =   24
               Left            =   165
               TabIndex        =   186
               Top             =   2280
               Width           =   1770
            End
            Begin VB.Label Label 
               AutoSize        =   -1  'True
               Caption         =   "Tipo Documento :"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   195
               Index           =   2
               Left            =   660
               TabIndex        =   43
               Top             =   375
               Width           =   1275
            End
            Begin VB.Label Label 
               AutoSize        =   -1  'True
               Caption         =   "Serie y Correlativo :"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   195
               Index           =   3
               Left            =   570
               TabIndex        =   42
               Top             =   825
               Width           =   1365
            End
            Begin VB.Label Label 
               AutoSize        =   -1  'True
               Caption         =   "Impresora :"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   195
               Index           =   5
               Left            =   1155
               TabIndex        =   41
               Top             =   1125
               Width           =   780
            End
            Begin VB.Label Label 
               AutoSize        =   -1  'True
               Caption         =   "Formulario :"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   195
               Index           =   6
               Left            =   1125
               TabIndex        =   40
               Top             =   1500
               Width           =   810
            End
            Begin VB.Label Label 
               AutoSize        =   -1  'True
               Caption         =   "Descripción :"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   195
               Index           =   7
               Left            =   1005
               TabIndex        =   39
               Top             =   1860
               Width           =   930
            End
            Begin VB.Label Label 
               AutoSize        =   -1  'True
               Caption         =   "Autorización :"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   195
               Index           =   19
               Left            =   975
               TabIndex        =   38
               Top             =   2805
               Width           =   960
            End
         End
         Begin VB.Frame Frame1 
            Caption         =   " Lectora de Barras  "
            Height          =   1395
            Left            =   -69285
            TabIndex        =   73
            Top             =   1065
            Width           =   4185
            Begin VB.CheckBox chkEAN13 
               Caption         =   "EAN13"
               Height          =   195
               Left            =   2280
               TabIndex        =   211
               Top             =   720
               Width           =   1095
            End
            Begin VB.OptionButton opcCapturaPeso 
               Caption         =   "Captura el Peso"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   195
               Left            =   2190
               TabIndex        =   153
               Top             =   1040
               Width           =   1575
            End
            Begin VB.OptionButton opcCapturaPrecio 
               Caption         =   "Captura el Precio"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   195
               Left            =   270
               TabIndex        =   152
               Top             =   1040
               Value           =   -1  'True
               Width           =   1575
            End
            Begin VB.CheckBox chkRotulado 
               Caption         =   "Enlace Rotulado"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   195
               Left            =   270
               TabIndex        =   151
               Top             =   720
               Width           =   1665
            End
            Begin VB.TextBox txtLongitudBarra 
               Height          =   330
               Left            =   2250
               TabIndex        =   74
               Top             =   307
               Width           =   870
            End
            Begin VB.Label Label 
               AutoSize        =   -1  'True
               Caption         =   "Longitud Código Barras"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   195
               Index           =   17
               Left            =   270
               TabIndex        =   75
               Top             =   375
               Width           =   1650
            End
         End
         Begin VB.Frame Frame3 
            Caption         =   " Enlace VisaNet "
            Height          =   780
            Left            =   -69285
            TabIndex        =   71
            Top             =   3360
            Visible         =   0   'False
            Width           =   4185
            Begin VB.CheckBox chkVisaNet 
               Caption         =   "Caja conectada al POS Pin Pad"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   195
               Left            =   225
               TabIndex        =   72
               Top             =   375
               Width           =   3825
            End
         End
         Begin VB.Frame Frame2 
            Caption         =   " Visor de Precios "
            Height          =   1695
            Left            =   -74850
            TabIndex        =   64
            Top             =   1065
            Width           =   5430
            Begin VB.TextBox txtMensaje2 
               Height          =   285
               Left            =   1695
               MaxLength       =   19
               TabIndex        =   67
               Top             =   1200
               Width           =   3585
            End
            Begin VB.TextBox txtMensaje1 
               Height          =   285
               Left            =   1695
               MaxLength       =   19
               TabIndex        =   66
               Top             =   840
               Width           =   3585
            End
            Begin VB.TextBox txtPuerto 
               Alignment       =   1  'Right Justify
               Height          =   285
               Left            =   1695
               MaxLength       =   1
               TabIndex        =   65
               Top             =   420
               Width           =   735
            End
            Begin VB.Label Label 
               AutoSize        =   -1  'True
               Caption         =   "Mensaje 2 :"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   195
               Index           =   14
               Left            =   705
               TabIndex        =   70
               Top             =   1245
               Width           =   825
            End
            Begin VB.Label Label 
               AutoSize        =   -1  'True
               Caption         =   "Mensaje 1 :"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   195
               Index           =   13
               Left            =   705
               TabIndex        =   69
               Top             =   885
               Width           =   825
            End
            Begin VB.Label Label 
               AutoSize        =   -1  'True
               Caption         =   "Puerto :"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   195
               Index           =   12
               Left            =   975
               TabIndex        =   68
               Top             =   465
               Width           =   555
            End
         End
         Begin VB.Frame Frame4 
            Caption         =   " Balanza Electrónica "
            Height          =   2175
            Left            =   -74880
            TabIndex        =   61
            Top             =   2760
            Width           =   5385
            Begin VB.TextBox txtbaltiempo 
               Alignment       =   1  'Right Justify
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   285
               Left            =   4680
               MaxLength       =   2
               TabIndex        =   236
               Top             =   960
               Width           =   615
            End
            Begin VB.TextBox txtBalcomando 
               Alignment       =   1  'Right Justify
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   285
               Left            =   4680
               MaxLength       =   3
               TabIndex        =   235
               Top             =   600
               Width           =   615
            End
            Begin VB.CommandButton cmdGuardarBalanza 
               Caption         =   "Guardar"
               Height          =   375
               Left            =   4440
               TabIndex        =   233
               Top             =   1680
               Width           =   855
            End
            Begin VB.ComboBox cboBal2 
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   330
               ItemData        =   "frmCajaDetalle.frx":34DA
               Left            =   1440
               List            =   "frmCajaDetalle.frx":34ED
               TabIndex        =   232
               Top             =   600
               Width           =   2055
            End
            Begin VB.ComboBox cboBal5 
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   330
               ItemData        =   "frmCajaDetalle.frx":3505
               Left            =   1440
               List            =   "frmCajaDetalle.frx":3512
               TabIndex        =   231
               Top             =   1680
               Visible         =   0   'False
               Width           =   2055
            End
            Begin VB.ComboBox cboBal3 
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   330
               ItemData        =   "frmCajaDetalle.frx":3535
               Left            =   1440
               List            =   "frmCajaDetalle.frx":3548
               TabIndex        =   230
               Top             =   960
               Width           =   2055
            End
            Begin VB.ComboBox cboBal4 
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   330
               ItemData        =   "frmCajaDetalle.frx":3571
               Left            =   1440
               List            =   "frmCajaDetalle.frx":357E
               TabIndex        =   229
               Top             =   1320
               Width           =   2055
            End
            Begin VB.ComboBox cboBal1 
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   330
               ItemData        =   "frmCajaDetalle.frx":3591
               Left            =   1440
               List            =   "frmCajaDetalle.frx":35CB
               TabIndex        =   228
               Top             =   240
               Width           =   2055
            End
            Begin VB.CheckBox chkActivoBal 
               Alignment       =   1  'Right Justify
               Caption         =   "Activo"
               Height          =   195
               Left            =   4080
               TabIndex        =   222
               Top             =   1320
               Width           =   975
            End
            Begin VB.TextBox txtBalanzaPuerto 
               Alignment       =   1  'Right Justify
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   285
               Left            =   4680
               MaxLength       =   1
               TabIndex        =   62
               Top             =   240
               Width           =   615
            End
            Begin VB.Label Label 
               Alignment       =   1  'Right Justify
               Caption         =   "Tiempo Espera Bal.(Seg):"
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
               Index           =   37
               Left            =   3480
               TabIndex        =   237
               Top             =   840
               Width           =   1095
            End
            Begin VB.Label Label 
               Alignment       =   1  'Right Justify
               Caption         =   "Comando:"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   255
               Index           =   36
               Left            =   3720
               TabIndex        =   234
               Top             =   600
               Width           =   855
            End
            Begin VB.Label Label 
               Alignment       =   1  'Right Justify
               Caption         =   "Control de flujo :"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   255
               Index           =   35
               Left            =   120
               TabIndex        =   227
               Top             =   1800
               Visible         =   0   'False
               Width           =   1185
            End
            Begin VB.Label Label 
               AutoSize        =   -1  'True
               Caption         =   "Bits de parada :"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   195
               Index           =   34
               Left            =   240
               TabIndex        =   226
               Top             =   1440
               Width           =   1110
            End
            Begin VB.Label Label 
               AutoSize        =   -1  'True
               Caption         =   "Paridad :"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   195
               Index           =   33
               Left            =   720
               TabIndex        =   225
               Top             =   1080
               Width           =   630
            End
            Begin VB.Label Label 
               AutoSize        =   -1  'True
               Caption         =   "Bits de datos :"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   195
               Index           =   32
               Left            =   360
               TabIndex        =   224
               Top             =   720
               Width           =   1005
            End
            Begin VB.Label Label 
               Alignment       =   1  'Right Justify
               Caption         =   "Bits por segundo :"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   315
               Index           =   31
               Left            =   120
               TabIndex        =   223
               Top             =   360
               Width           =   1275
            End
            Begin VB.Label Label 
               AutoSize        =   -1  'True
               Caption         =   "Puerto Serial :"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   195
               Index           =   18
               Left            =   3600
               TabIndex        =   63
               Top             =   240
               Width           =   990
            End
         End
         Begin VB.Frame Frame5 
            Caption         =   " Texto Predeterminado Por Consumo "
            Height          =   825
            Left            =   -74880
            TabIndex        =   59
            Top             =   5040
            Width           =   9705
            Begin VB.TextBox txtTextoConsumo 
               Height          =   285
               Left            =   180
               MaxLength       =   200
               TabIndex        =   60
               Top             =   360
               Width           =   9345
            End
         End
         Begin VB.CommandButton cmdOpcionGrilla 
            Caption         =   "Eliminar"
            Height          =   645
            Index           =   7
            Left            =   -66270
            Picture         =   "frmCajaDetalle.frx":363C
            Style           =   1  'Graphical
            TabIndex        =   58
            Top             =   2730
            Width           =   1215
         End
         Begin VB.CommandButton cmdOpcionGrilla 
            Caption         =   "Agregar"
            Height          =   645
            Index           =   5
            Left            =   -66270
            Picture         =   "frmCajaDetalle.frx":373E
            Style           =   1  'Graphical
            TabIndex        =   57
            Top             =   1200
            Width           =   1215
         End
         Begin VB.CommandButton cmdOpcionGrilla 
            Caption         =   "Modificar"
            Height          =   645
            Index           =   6
            Left            =   -66270
            Picture         =   "frmCajaDetalle.frx":3C70
            Style           =   1  'Graphical
            TabIndex        =   56
            Top             =   1965
            Width           =   1215
         End
         Begin VB.CommandButton cmdOpcionGrilla 
            Caption         =   "Eliminar"
            Height          =   645
            Index           =   2
            Left            =   -66030
            Picture         =   "frmCajaDetalle.frx":41A2
            Style           =   1  'Graphical
            TabIndex        =   47
            Top             =   4200
            Width           =   975
         End
         Begin VB.CommandButton cmdOpcionGrilla 
            Caption         =   "Modificar"
            Height          =   645
            Index           =   1
            Left            =   -67080
            Picture         =   "frmCajaDetalle.frx":42A4
            Style           =   1  'Graphical
            TabIndex        =   46
            Top             =   4200
            Width           =   975
         End
         Begin VB.CommandButton cmdOpcionGrilla 
            Caption         =   "Agregar"
            Height          =   645
            Index           =   0
            Left            =   -68160
            Picture         =   "frmCajaDetalle.frx":47D6
            Style           =   1  'Graphical
            TabIndex        =   45
            Top             =   4200
            Width           =   975
         End
         Begin VB.CheckBox chkConsumo3 
            Caption         =   "Emisión por consumo en Caja Rápida"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   225
            Left            =   -74730
            TabIndex        =   33
            Top             =   3570
            Width           =   3570
         End
         Begin VB.CheckBox chkConsumo2 
            Caption         =   "Emisión por consumo en Pagos y División "
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   225
            Left            =   -74730
            TabIndex        =   32
            Top             =   3255
            Width           =   3570
         End
         Begin VB.CheckBox chkConsumo1 
            Caption         =   "Emisión por consumo en Emisión Rápida"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   225
            Left            =   -74730
            TabIndex        =   31
            Top             =   2925
            Width           =   3570
         End
         Begin VB.CheckBox chkObservacion 
            Caption         =   "Obligatoriedad de Observación"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   225
            Left            =   -74730
            TabIndex        =   30
            Top             =   2625
            Width           =   3570
         End
         Begin VB.CheckBox chkObligaPrecuenta 
            Caption         =   "Obligatoriedad de Impresión de Precuentas"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   225
            Left            =   -74730
            TabIndex        =   28
            Top             =   1935
            Width           =   3570
         End
         Begin VB.CheckBox chkCancelacion 
            Caption         =   "Obligatoriedad de Cancelación"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   225
            Left            =   -74730
            TabIndex        =   29
            Top             =   2265
            Width           =   3570
         End
         Begin VB.CheckBox chkObligaPrinter 
            Caption         =   "Obligatoriedad de Impresión de Pedido"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   225
            Left            =   -74730
            TabIndex        =   27
            Top             =   1590
            Width           =   3570
         End
         Begin VB.CheckBox chkComanda 
            Caption         =   "Obligatoriedad de Comandas Manuales"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   225
            Left            =   -74730
            TabIndex        =   26
            Top             =   1230
            Width           =   3570
         End
         Begin TrueOleDBGrid80.TDBGrid grdGrilla 
            Height          =   3105
            Left            =   -74865
            TabIndex        =   44
            Top             =   1065
            Width           =   9840
            _ExtentX        =   17357
            _ExtentY        =   5477
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
         Begin TrueOleDBGrid80.TDBGrid grdAI 
            Height          =   3870
            Left            =   -74865
            TabIndex        =   55
            Top             =   1110
            Width           =   8415
            _ExtentX        =   14843
            _ExtentY        =   6826
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
         Begin TrueOleDBGrid80.TDBGrid grdAChef 
            Height          =   3135
            Left            =   -74760
            TabIndex        =   174
            Top             =   1320
            Width           =   7335
            _ExtentX        =   12938
            _ExtentY        =   5530
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
         Begin MSComDlg.CommonDialog dlgFoto 
            Left            =   -71880
            Top             =   1560
            _ExtentX        =   847
            _ExtentY        =   847
            _Version        =   393216
         End
         Begin MSComDlg.CommonDialog dlgFotoPie 
            Left            =   -68280
            Top             =   1560
            _ExtentX        =   847
            _ExtentY        =   847
            _Version        =   393216
         End
         Begin VB.Image imgFotoPie 
            BorderStyle     =   1  'Fixed Single
            Height          =   3000
            Left            =   -69480
            Stretch         =   -1  'True
            ToolTipText     =   "Imagen Para Pie de Documentos"
            Top             =   1320
            Width           =   4005
         End
         Begin VB.Image imgFoto 
            BorderStyle     =   1  'Fixed Single
            Height          =   3000
            Left            =   -74520
            Stretch         =   -1  'True
            ToolTipText     =   "Imagen Para Cabecera de Documentos"
            Top             =   1320
            Width           =   4005
         End
      End
      Begin VB.TextBox txtCodigo 
         BackColor       =   &H00E0E0E0&
         Enabled         =   0   'False
         ForeColor       =   &H000000C0&
         Height          =   285
         Left            =   1995
         Locked          =   -1  'True
         TabIndex        =   76
         TabStop         =   0   'False
         Top             =   225
         Width           =   1170
      End
      Begin MSDataListLib.DataCombo cboPreCuenta 
         Height          =   315
         Left            =   1995
         TabIndex        =   77
         Top             =   909
         Width           =   2595
         _ExtentX        =   4577
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
      Begin MSDataListLib.DataCombo cboGrupo 
         Height          =   315
         Left            =   1995
         TabIndex        =   78
         Top             =   1281
         Width           =   2595
         _ExtentX        =   4577
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
      Begin MSDataListLib.DataCombo cboTipoPedido 
         Height          =   315
         Left            =   1995
         TabIndex        =   79
         Top             =   1653
         Width           =   2595
         _ExtentX        =   4577
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
      Begin MSDataListLib.DataCombo cboSucursal 
         Height          =   315
         Left            =   7695
         TabIndex        =   80
         Top             =   1275
         Width           =   2595
         _ExtentX        =   4577
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
      Begin VB.CheckBox chkActivo 
         Alignment       =   1  'Right Justify
         Caption         =   "Activo :"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   3480
         TabIndex        =   88
         Top             =   240
         Width           =   840
      End
      Begin MSDataListLib.DataCombo cboUnidadNegocio 
         Height          =   315
         Left            =   7695
         TabIndex        =   142
         Top             =   900
         Width           =   2595
         _ExtentX        =   4577
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
      Begin MSDataListLib.DataCombo cboSectorVenta 
         Height          =   315
         Left            =   7695
         TabIndex        =   144
         Top             =   480
         Width           =   2595
         _ExtentX        =   4577
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
      Begin MSDataListLib.DataCombo cboComprobante 
         Height          =   315
         Left            =   7695
         TabIndex        =   220
         Top             =   1650
         Width           =   2595
         _ExtentX        =   4577
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
      Begin VB.Label Label 
         Alignment       =   1  'Right Justify
         Caption         =   "Impresora Comprobante MESA24/7 :"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   435
         Index           =   25
         Left            =   5655
         TabIndex        =   221
         Top             =   1650
         Width           =   1905
      End
      Begin VB.Label Label 
         Alignment       =   2  'Center
         Caption         =   "Agrupación de Punto de Venta :"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   435
         Index           =   20
         Left            =   6350
         TabIndex        =   141
         Top             =   480
         Width           =   1305
      End
      Begin VB.Label lblSucursal 
         Alignment       =   1  'Right Justify
         Caption         =   "Sucursal :"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   6100
         TabIndex        =   87
         Top             =   1305
         Width           =   1455
      End
      Begin VB.Label Label 
         AutoSize        =   -1  'True
         Caption         =   "Unidad de Negocio :"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   15
         Left            =   6105
         TabIndex        =   86
         Top             =   960
         Width           =   1470
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Impresora Pre Cuenta :"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   2
         Left            =   210
         TabIndex        =   85
         Top             =   1035
         Width           =   1620
      End
      Begin VB.Label Label 
         AutoSize        =   -1  'True
         Caption         =   "Código :"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   0
         Left            =   1290
         TabIndex        =   84
         Top             =   270
         Width           =   585
      End
      Begin VB.Label Label 
         AutoSize        =   -1  'True
         Caption         =   "Descripción :"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   1
         Left            =   945
         TabIndex        =   83
         Top             =   615
         Width           =   930
      End
      Begin VB.Label Label 
         AutoSize        =   -1  'True
         Caption         =   "Grupo predeterminado :"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   8
         Left            =   210
         TabIndex        =   82
         Top             =   1365
         Width           =   1665
      End
      Begin VB.Label Label 
         AutoSize        =   -1  'True
         Caption         =   "Tipo Pedido predeterm :"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   10
         Left            =   180
         TabIndex        =   81
         Top             =   1740
         Width           =   1695
      End
   End
   Begin VB.PictureBox Picture 
      Align           =   2  'Align Bottom
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   750
      Left            =   0
      ScaleHeight     =   690
      ScaleWidth      =   11805
      TabIndex        =   15
      TabStop         =   0   'False
      Top             =   8625
      Width           =   11865
      Begin VB.CommandButton cmdOpcion 
         Caption         =   "Agregar"
         Height          =   615
         Index           =   0
         Left            =   7005
         Picture         =   "frmCajaDetalle.frx":4D08
         Style           =   1  'Graphical
         TabIndex        =   136
         Top             =   60
         Width           =   1170
      End
      Begin VB.PictureBox PicNavegacion 
         BackColor       =   &H80000004&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   30
         ScaleHeight     =   555
         ScaleWidth      =   6660
         TabIndex        =   16
         Top             =   60
         Width           =   6720
         Begin VB.CommandButton cmdNavegar 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   555
            Index           =   3
            Left            =   5220
            Picture         =   "frmCajaDetalle.frx":523A
            Style           =   1  'Graphical
            TabIndex        =   22
            TabStop         =   0   'False
            Top             =   0
            Width           =   480
         End
         Begin VB.CommandButton cmdNavegar 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   555
            Index           =   4
            Left            =   5700
            Picture         =   "frmCajaDetalle.frx":577C
            Style           =   1  'Graphical
            TabIndex        =   21
            TabStop         =   0   'False
            Top             =   0
            Width           =   480
         End
         Begin VB.CommandButton cmdNavegar 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   555
            Index           =   5
            Left            =   6180
            Picture         =   "frmCajaDetalle.frx":5CBE
            Style           =   1  'Graphical
            TabIndex        =   20
            TabStop         =   0   'False
            Top             =   0
            Width           =   480
         End
         Begin VB.CommandButton cmdNavegar 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   555
            Index           =   0
            Left            =   0
            Picture         =   "frmCajaDetalle.frx":6200
            Style           =   1  'Graphical
            TabIndex        =   19
            TabStop         =   0   'False
            Top             =   0
            Width           =   480
         End
         Begin VB.CommandButton cmdNavegar 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   555
            Index           =   2
            Left            =   960
            Picture         =   "frmCajaDetalle.frx":6742
            Style           =   1  'Graphical
            TabIndex        =   18
            TabStop         =   0   'False
            Top             =   0
            Width           =   480
         End
         Begin VB.CommandButton cmdNavegar 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   555
            Index           =   1
            Left            =   480
            Picture         =   "frmCajaDetalle.frx":6C84
            Style           =   1  'Graphical
            TabIndex        =   17
            TabStop         =   0   'False
            Top             =   0
            Width           =   480
         End
         Begin VB.Label cmdTexto 
            Alignment       =   2  'Center
            Caption         =   "Registro 0 de 0"
            Height          =   195
            Left            =   1590
            TabIndex        =   24
            Top             =   180
            Width           =   3495
         End
      End
      Begin VB.CommandButton cmdOpcion 
         Caption         =   "Grabar"
         Height          =   615
         Index           =   1
         Left            =   8160
         Picture         =   "frmCajaDetalle.frx":71C6
         Style           =   1  'Graphical
         TabIndex        =   10
         Top             =   60
         Width           =   1170
      End
      Begin VB.CommandButton cmdOpcion 
         Caption         =   "Eliminar"
         Height          =   615
         Index           =   2
         Left            =   9330
         Picture         =   "frmCajaDetalle.frx":76F8
         Style           =   1  'Graphical
         TabIndex        =   13
         Top             =   60
         Width           =   1170
      End
      Begin VB.CommandButton cmdOpcion 
         Caption         =   "Salir"
         Height          =   615
         Index           =   3
         Left            =   10500
         Picture         =   "frmCajaDetalle.frx":77FA
         Style           =   1  'Graphical
         TabIndex        =   14
         Top             =   60
         Width           =   1170
      End
   End
   Begin VB.Image Image 
      Height          =   8520
      Left            =   15
      Picture         =   "frmCajaDetalle.frx":78EC
      Stretch         =   -1  'True
      Top             =   45
      Width           =   1140
   End
End
Attribute VB_Name = "frmCajaDetalle"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim RsGrupo As Recordset
Dim RsGrilla As Recordset
Dim RsImpresora As Recordset
Dim RsPreCuenta As Recordset
Dim RsTipoDocumento As Recordset
Dim RsArea As Recordset
Dim RsTipoPedido As Recordset
Dim RsLocal As Recordset
Dim RsUnidadNegocio As Recordset
Dim rsAreaProduccion As Recordset
Dim RsAI As Recordset
Dim RsImpArea As Recordset
Dim RsFormulario As Recordset
Dim wAgrega As Boolean
Dim Rssucursal As Recordset
Dim RsSectorVenta As Recordset
'LG
Dim RsSubGrupo As Recordset
Dim rsAreaProducccionSubgrupo As Recordset
Dim rsAreaSubGrupo As Recordset
Dim strFilenameRuta As String
Dim strFilenameRutaPie As String
'CESAR AREA CHEF
Dim RsAChef As Recordset

Sub LlenaCombos()
    
    With cboTipoDocumento
         Isql = "Select * from vTipoDocumento where lActivo = 1 order by Codigo"
         Set RsTipoDocumento = Lib.OpenRecordset(Isql, Cn)
         Set .RowSource = RsTipoDocumento
             .DataField = "Descripcion"
             .ListField = "Descripcion"
             .BoundColumn = "Codigo"
    End With
    
    With cboGrupo
         Isql = "Select * from vGrupo order by Descripcion"
         Set RsGrupo = Lib.OpenRecordset(Isql, Cn)
         Set .RowSource = RsGrupo
             .DataField = "Descripcion"
             .ListField = "Descripcion"
             .BoundColumn = "Codigo"
    End With
    
    With cboFormulario
        If lFESpring Then
            Isql = "Select * from vFormulario where lActivo = 1 and codigo = '01' order by Codigo"
        Else
            Isql = "Select * from vFormulario where lActivo = 1 order by Codigo"
        End If
        Set RsFormulario = Lib.OpenRecordset(Isql, Cn)
        Set .RowSource = RsFormulario
            .DataField = "Descripcion"
            .ListField = "Descripcion"
            .BoundColumn = "Codigo"
    End With
        
    With cboArea
         Isql = "Select * from vArea where lActivo = 1 order by Codigo"
         Set RsArea = Lib.OpenRecordset(Isql, Cn)
         Set .RowSource = RsArea
             .DataField = "Descripcion"
             .ListField = "Descripcion"
             .BoundColumn = "Codigo"
    End With
    
    With cboTipoPedido
         Isql = "Select * from vTipoPedido order by Codigo"
         Set RsTipoPedido = Lib.OpenRecordset(Isql, Cn)
         Set .RowSource = RsTipoPedido
             .DataField = "Descripcion"
             .ListField = "Descripcion"
             .BoundColumn = "Codigo"
    End With
    
    With cboUnidadNegocio
         Isql = "Select * from vUnidadNegocio order by Codigo"
         Set RsUnidadNegocio = Lib.OpenRecordset(Isql, Cn)
         Set .RowSource = RsUnidadNegocio
             .DataField = "Descripcion"
             .ListField = "Descripcion"
             .BoundColumn = "Codigo"
    End With
    
    'multiareas
        With cboAreaProduccion
         Isql = "select * from vArea where lActivo=1 union select 'ABC','SIN AREA','SIN AREA','999',1,'',0,NULL,0 order by 1 "
         Set rsAreaProduccion = Lib.OpenRecordset(Isql, Cn)
         Set .RowSource = rsAreaProduccion
             .DataField = "Descripcion"
             .ListField = "Descripcion"
             .BoundColumn = "Codigo"
    End With
    
    'SUCURSALES
        With cboSucursal
         Isql = "Select * from vSucursal where lActivo = 1 order by Codigo"
         Set Rssucursal = Lib.OpenRecordset(Isql, Cn)
         Set .RowSource = Rssucursal
             .DataField = "Descripcion"
             .ListField = "Descripcion"
             .BoundColumn = "Codigo"
    End With
    
    With cboLocal
         Isql = "Select * from vLocal order by Codigo"
         Set RsLocal = Lib.OpenRecordset(Isql, Cn)
         Set .RowSource = RsLocal
             .DataField = "Codigo"
             .ListField = "Codigo"
             .BoundColumn = "Codigo"
    End With
    
    'CESAR Sector
    With cboSectorVenta
         Isql = "Select * from vSectorVenta order by Codigo"
         Set RsSectorVenta = Lib.OpenRecordset(Isql, Cn)
         Set .RowSource = RsSectorVenta
             .DataField = "Detallado"
             .ListField = "Detallado"
             .BoundColumn = "Codigo"
    End With
    
    With Me.cboSubGrupo
         Isql = "select codigo, tresumido AS descripcion from vSubGrupo where lActivo=1 order by tresumido"
         Set RsSubGrupo = Lib.OpenRecordset(Isql, Cn)
         Set .RowSource = RsSubGrupo
             .DataField = "Descripcion"
             .ListField = "Descripcion"
             .BoundColumn = "Codigo"
    End With
    
    With cboAreaProd
         Isql = "select * from vArea where lActivo=1 order by 1 "
         Set rsAreaProducccionSubgrupo = Lib.OpenRecordset(Isql, Cn)
         Set .RowSource = rsAreaProducccionSubgrupo
             .DataField = "Descripcion"
             .ListField = "Descripcion"
             .BoundColumn = "Codigo"
    End With
    
    'CESAR AREA CHEF
    With cboAreaChef
         Isql = "Select * from vArea where lActivo = 1 order by Codigo"
         Set RsArea = Lib.OpenRecordset(Isql, Cn)
         Set .RowSource = RsArea
             .DataField = "Descripcion"
             .ListField = "Descripcion"
             .BoundColumn = "Codigo"
    End With
    
    
End Sub

Sub Asignar()
    txtCodigo = IIf(IsNull(frmCaja.RsCabecera!tCaja), "", frmCaja.RsCabecera!tCaja)
   
    With cboPreCuenta
         Isql = "Select * from TIMPRESORA where tCaja = '" & txtCodigo.Text & "' order by tImpresora"
         
         Set RsPreCuenta = Lib.OpenRecordset(Isql, Cn)
         Set .RowSource = RsPreCuenta
             .DataField = "tDescripcion"
             .ListField = "tDescripcion"
             .BoundColumn = "tImpresora"
    End With
    
    With cboComprobante
         Isql = "Select * from TIMPRESORA where tCaja = '" & txtCodigo.Text & "' order by tImpresora"
         
         Set RsPreCuenta = Lib.OpenRecordset(Isql, Cn)
         Set .RowSource = RsPreCuenta
             .DataField = "tDescripcion"
             .ListField = "tDescripcion"
             .BoundColumn = "tImpresora"
    End With
        
    With frmCaja.RsCabecera
        'Cuadro de Texto
        txtDetallado = IIf(IsNull(!tDescripcion), "", !tDescripcion)
        cboPreCuenta.BoundText = IIf(IsNull(!tPrecuenta), "", Trim(!tPrecuenta))
        cboComprobante.BoundText = IIf(IsNull(!tCompMesa247), "", Trim(!tCompMesa247))
        cboGrupo.BoundText = IIf(IsNull(!tgrupo), "", Trim(!tgrupo))
        cboTipoPedido.BoundText = IIf(IsNull(!tTipoPedido), "", Trim(!tTipoPedido))
        cboUnidadNegocio.BoundText = IIf(IsNull(!tUnidadNegocio), "", Trim(!tUnidadNegocio))
        cboSucursal.BoundText = IIf(IsNull(!tSucursal), "", Trim(!tSucursal))
        cboSectorVenta.BoundText = IIf(IsNull(!tSectorVenta), "", Trim(!tSectorVenta))
        'multiarea
        If IsNull(!tSubAlmacen) Or !tSubAlmacen = "" Then
            cboAreaProduccion.BoundText = "ABC"
        Else
            cboAreaProduccion.BoundText = !tSubAlmacen
        End If
        'Check Box
        chkComanda = IIf(!lComanda = True, 1, 0)
        chkVComanda = IIf(!vComanda = True, 1, 0)
        chkComboPrecuenta = IIf(!lComboPrecuenta = True, 1, 0)
        chkComboDocumento = IIf(!lComboDocumento = True, 1, 0)
        chkAdicion = IIf(!lAdicion = True, 1, 0)
        chkActivo = IIf(!lActivo = True, 1, 0)
        chkConsumo1 = IIf(!lConsumo1 = True, 1, 0)
        chkConsumo2 = IIf(!lConsumo2 = True, 1, 0)
        chkConsumo3 = IIf(!lConsumo3 = True, 1, 0)
        'codigoReciboIngreso
        chkCodigoReciboIngreso = IIf(!lCodigoReciboIngreso = True, 1, 0)
        
        
        chkPreCuenta = IIf(!lPrecuenta = True, 1, 0)
        chkAgrupada = IIf(!lPrecuentaAgrupada = True, 1, 0)
        chkDocumentoAgrupado = IIf(!lDocumentoAgrupado = True, 1, 0)
        chkOrden = IIf(!lOrden = True, 1, 0)
        
        Me.chkBuscaPedido.value = IIf(!lBuscaPedidoNumero = True, 1, 0)
        chkImprimeImagCabPrecuenta.value = IIf(IsNull(!lImprimeImagCabPrecuenta), 0, IIf(!lImprimeImagCabPrecuenta = True, 1, 0))
        chkImprimeImagPiePrecuenta.value = IIf(IsNull(!lImprimeImagPiePrecuenta), 0, IIf(!lImprimeImagPiePrecuenta = True, 1, 0))
        
        chkAccesoDespachoPedido.value = IIf(IsNull(!lAccesoDespachoPedido), 0, IIf(!lAccesoDespachoPedido = True, 1, 0))
        
        
        chkEliminaC = IIf(!lMotivoEliminaC = True, 1, 0)
        chkPasswordC = IIf(!lPasswordC = True, 1, 0)
        
        chkElimina = IIf(!lMotivoElimina = True, 1, 0)
        chkPassword = IIf(!lPassword = True, 1, 0)
        chkObligaPrinter = IIf(IsNull(!lObligaPrinter), 0, IIf(!lObligaPrinter = True, 1, 0))
        chkObligaPrecuenta = IIf(IsNull(!lObligaPrecuenta), 0, IIf(!lObligaPrecuenta = True, 1, 0))
        chkObligaCierre.value = IIf(!lObligaCierre = True, 1, 0)
        chkFiltroTipoPedido.value = IIf(!lFiltroTipoPedido = True, 1, 0)
        chkCancelacion.value = IIf(!lCancelacion = True, 1, 0)
        chkDirecto.value = IIf(!lDirecto = True, 1, 0)
        chkCambioMesa.value = IIf(!lCambioMesa = True, 1, 0)
        txtPuerto = IIf(IsNull(!nPuerto), 0, !nPuerto)
        txtMensaje1 = IIf(IsNull(!tMensaje1), "", !tMensaje1)
        txtMensaje2 = IIf(IsNull(!tMensaje2), "", !tMensaje2)
        txtLongitudBarra = Val(IIf(IsNull(!nLongitudBarra), 0, !nLongitudBarra))
        chkVisaNet.value = IIf(!lVisaNet = True, 1, 0)
        chkImpuestoPrecuenta.value = IIf(!lImpuestoPrecuenta = True, 1, 0)
        chkValor.value = IIf(!lValorCortesia = True, 1, 0)
        chkEquivaPrecuenta = IIf(!lEquivaDolaPrecuenta = True, 1, 0)
        chkObservacion.value = IIf(!lObservacion = True, 1, 0)
        chkCajaRapida.value = IIf(!lCajaRapida = True, 1, 0)
        chkPropiedadPrecuenta.value = IIf(!lPropiedadPrecuenta = True, 1, 0)
        chkPropiedadDocumento.value = IIf(!lPropiedadDocumento = True, 1, 0)
        chkPrecioNetoPrecuenta.value = IIf(!lPrecioNetoPrecuenta = True, 1, 0)
        
        chkImprimeImagCabPrecuenta.value = IIf(!lImprimeImagCabPrecuenta = True, 1, 0)
        chkImprimeImagPiePrecuenta.value = IIf(!lImprimeImagPiePrecuenta = True, 1, 0)
        
        'FACTURACION ELECTRONICA
        'chkFacturacionE.value = IIf(!lFacturacionElectronica = True, 1, 0)
        
        txtLimitePrecuenta.Text = IIf(IsNull(!nLimitePrecuenta), 0, !nLimitePrecuenta)
        txtLimiteReimpresion.Text = IIf(IsNull(!nLimiteReimpresion), 0, !nLimiteReimpresion)
        chkPasswordTransferencia.value = IIf(!lPasswordTransferencia = True, 1, 0)
        chkPasswordImportar.value = IIf(!lPasswordImportarPedido = True, 1, 0)
        chkCD.value = IIf(!lCD = True, 1, 0)
        chkMultiCajero.value = IIf(!lMultiCajero = True, 1, 0)
        chkMCPV.value = IIf(!lMCPV = True, 1, 0)
        chkCCVOX.value = IIf(!lCCVOX = True, 1, 0)
        chkObservacionPrecuenta.value = IIf(!lObservacionPrecuenta = True, 1, 0)
        chkObservacionDocumento.value = IIf(!lObservacionDocumento = True, 1, 0)
        chkObservacionCabDoc.value = IIf(!lObservacionCabDoc = True, 1, 0)
        chkDescripcionAlternativa.value = IIf(!lActivaImpDscAlternativa = True, 1, 0)
        Me.chkMotDesc.value = IIf(IsNull(!lMotivoDescuento), 0, IIf(!lMotivoDescuento = True, 1, 0))
        Me.chkCajaContingencia.value = IIf(IsNull(!lCajaContingencia), 0, IIf(!lCajaContingencia = True, 1, 0))
        Me.chkImpPropina.value = IIf(IsNull(!lImpPropina), 0, IIf(!lImpPropina = True, 1, 0))
        Me.chkComandaF2.value = IIf(IsNull(!lImpComandaf2), 0, IIf(!lImpComandaf2 = True, 1, 0))
        chkDisgrega.value = IIf(!lDisgrega = True, 1, 0)
        chkSiab.value = IIf(!lSiab = True, 1, 0)
        
        chkBloqueaPrecuenta = IIf(!lBloqueaPrecuenta = True, 1, 0)
        If chkBloqueaPrecuenta.value = 1 Then
                txtLimitePrecuenta.Enabled = False
            Else
                txtLimitePrecuenta.Enabled = True
        End If
        
        'TVS============================
        chkCompatibilidadTVS.value = IIf(!lCompatibilidadTVS = True, 1, 0)
        '===============================
        
        chkPagoRapido.value = IIf(!lPagoRapido = True, 1, 0)
        chkPasswordPorCobrar.value = IIf(!lPasswordPorCobrar = True, 1, 0)
        chkModificaTipoPedido.value = IIf(!lmodificatipoPedido = True, 1, 0)
        txtBalanzaPuerto = IIf(IsNull(!nBalanzaPuerto), 0, !nBalanzaPuerto)

        chkCajaMobile.value = IIf(!lCajaMobile = True, 1, 0)
        chkRotulado.value = IIf(!lRotulado = True, 1, 0)
        
        '0084-2013 CESAR
        chkPagoRapidopv.value = IIf(!lPagoRapidoPV = True, 1, 0)
        txtTextoConsumo = IIf(IsNull(!tTextoConsumo), "", !tTextoConsumo)
        chkPagoRapidoMod.value = IIf(!lPagoRapidoMod = True, 1, 0)
        
        chkWebAp.value = IIf(!lWebAp = True, 1, 0)
        chkMesa247.value = IIf(!lMesa247 = True, 1, 0)
        
        chkConsumo4.value = IIf(!lConsumo4 = True, 1, 0)
        
        chkPrecuentaNoValorizada.value = IIf(!lPrecuentaNoValorizada = True, 1, 0)
        chkEAN13.value = IIf(!EAN13 = True, 1, 0)
         
        Dim xCapturaPeso As Boolean
        xCapturaPeso = IIf(IsNull(!lCapturaPeso), 0, !lCapturaPeso)
        
        If xCapturaPeso = True Then
           opcCapturaPeso.value = True
           opcCapturaPrecio.value = False
        Else
           opcCapturaPeso.value = False
           opcCapturaPrecio.value = True
        End If
        
        
         If !lMultiAreaSubGrupo = True Then
            chkMulti2.value = 1
            Me.fra2.Enabled = True
            Else
            chkMulti2.value = 0
            Me.fra2.Enabled = False
        End If
        
        
        If !lMultiAreaCaja = True Then
            chkMulti1.value = 1
            Me.fra1.Enabled = True
        Else
            chkMulti1.value = 0
            Me.fra1.Enabled = False
        End If
        
        'HUELLA
        If !lHuella = True Then
            chkHuella.value = 1
        Else
            chkHuella.value = 0
        End If
    
       Me.chkBuscaPedidoVisualizaGrilla.value = IIf(!lBuscarPedidoVisualizarGrilla = True, 1, 0)
       Me.chkBuscaPedidoFiltrarMesa.value = IIf(!lBuscarPedidoFiltrarMesa = True, 1, 0)

       Me.chkClaveEnvio.value = IIf(!lClaveEnvioProduccion = True, 1, 0)
       Me.chkPassOtrosPagos.value = IIf(!lPassOtrosPagos = True, 1, 0)
       
       On Error GoTo err
       Dim rst1 As New ADODB.Recordset
       imgFoto.DataField = "foto"
       cmdAgregarFoto.Caption = "Editar Imagen Cabecera"
       Set rst1 = Lib.OpenRecordset("select iimagencabdoc as foto from tcaja where tcaja='" & txtCodigo.Text & "'", Cn)
       Set imgFoto.DataSource = rst1
       imgFotoPie.DataField = "foto"
       cmdAgregarFotoPie.Caption = "Editar Imagen Pie"
       Set rst1 = Lib.OpenRecordset("select iimagenpiedoc as foto from tcaja where tcaja='" & txtCodigo.Text & "'", Cn)
       Set imgFotoPie.DataSource = rst1
       Call asignarBalanza
err:
       
    End With
    
    'Cambiar el Filtro
    RsGrilla.Filter = "tCaja ='" & txtCodigo.Text & "'"
    RsAI.Filter = "tCaja ='" & txtCodigo.Text & "'"
    'CESAR AREA CHEF
    RsAChef.Filter = "tCaja ='" & txtCodigo.Text & "'"
    '----------------------
    rsAreaSubGrupo.Filter = "tCaja ='" & txtCodigo.Text & "'"
End Sub

Private Sub cboTipoDocumento_Change()
    If cboTipoDocumento.BoundText = "00" Then
       'Label(4).Visible = False
       Frame14.Visible = False
    Else
       'Label(4).Visible = True
       Frame14.Visible = True
    End If
    
    If sImpuesto1 <> "" And cboTipoDocumento.BoundText <> "00" Then
       chkImpuesto1.Visible = True
       chkImpuesto1.Caption = sImpuesto1
       chkImpuesto1.value = 1
    Else
       chkImpuesto1.Visible = False
       chkImpuesto1.value = 0
    End If
        
    If sImpuesto2 <> "" And cboTipoDocumento.BoundText <> "00" Then
       chkImpuesto2.Visible = True
       chkImpuesto2.Caption = sImpuesto2
       chkImpuesto2.value = 1
    Else
       chkImpuesto2.Visible = False
       chkImpuesto2.value = 0
    End If
        
    If sImpuesto3 <> "" And cboTipoDocumento.BoundText <> "00" Then
       chkImpuesto3.Visible = True
       chkImpuesto3.Caption = sImpuesto3
       chkImpuesto3.value = 1
    Else
      chkImpuesto3.Visible = False
      chkImpuesto3.value = 0
    End If
End Sub

Private Sub chkAgrupada_Click()
   If chkAgrupada.value Then
      chkComboPrecuenta.value = False
      chkComboPrecuenta.Enabled = False
      chkPropiedadPrecuenta.Enabled = False
      chkPropiedadPrecuenta.value = 0
      chkObservacionPrecuenta.value = 0
      chkObservacionPrecuenta.Enabled = False
   Else
      chkPropiedadPrecuenta.Enabled = True
      chkComboPrecuenta.Enabled = True
      chkObservacionPrecuenta.Enabled = True
      chkPropiedadPrecuenta.value = 0
      chkObservacionPrecuenta.value = 0
   End If
End Sub

Private Sub chkBloqueaPrecuenta_Click()
    If chkBloqueaPrecuenta.value = 1 Then
            If chkObligaPrecuenta.value = 1 Then
                MsgBox "Está activada la Obligatoriedad de Emisión de Precuentas", vbCritical, sMensaje
                chkBloqueaPrecuenta.value = 0
                Exit Sub
            Else
        
                txtLimitePrecuenta.Enabled = False
            End If
        Else
            txtLimitePrecuenta.Enabled = True
    End If
End Sub


Private Sub chkDocumentoAgrupado_Click()
   If chkDocumentoAgrupado.value Then
      chkComboDocumento.value = False
      chkComboDocumento.Enabled = False
      
      chkPropiedadDocumento.Enabled = False
      chkObservacionDocumento.Enabled = False
      chkPropiedadDocumento.value = 0
      chkObservacionDocumento.value = 0
   Else
      chkComboDocumento.Enabled = True
      chkPropiedadDocumento.Enabled = True
      chkObservacionDocumento.Enabled = True
      chkPropiedadDocumento.value = 0
      chkObservacionDocumento.value = 0
      
   End If
End Sub

Private Sub chkMCPV_Click()
  If chkMCPV.value = 1 Then
     chkMultiCajero.Enabled = False
     chkMultiCajero.value = 0
  Else
     chkMultiCajero.Enabled = True
  End If
End Sub

Private Sub chkMesa247_Click()
    If chkMesa247.value = 1 Then
        Label(25).Visible = True
        cboComprobante.Visible = True
    Else
        Label(25).Visible = False
        cboComprobante.Visible = False
    End If
End Sub

Private Sub chkMultiCajero_Click()
  If chkMultiCajero.value = 1 Then
     chkMCPV.Enabled = False
     chkMCPV.value = 0
  Else
     chkMCPV.Enabled = True
  End If
End Sub

Private Sub chkObligaPrecuenta_Click()
    If chkObligaPrecuenta.value = 1 Then
        If chkBloqueaPrecuenta.value = 1 Then
            MsgBox "La Emisión de Precuentas esta bloqueada", vbCritical, sMensaje
            chkObligaPrecuenta.value = 0
            Exit Sub
        End If
    End If
        
End Sub

Private Sub chkPrecioNetoPrecuenta_Click()
    If chkPrecioNetoPrecuenta.value = 1 Then
        chkImpuestoPrecuenta.value = 1
        chkImpuestoPrecuenta.Enabled = False
    Else
        chkImpuestoPrecuenta.Enabled = True
    End If
End Sub

Private Sub cmdAgregarFoto_Click()
On Error GoTo errHandler

If txtCodigo.Text <> "" Then
    dlgFoto.CancelError = False
    With cmdAgregarFoto
        If .Caption = "Editar Imagen Cabecera" Then
            dlgFoto.Filter = "Image(*.jpg)|*.jpg|Image(*.gif)| *.gif" '"archivos (*.bmp)|*.bmp"
            dlgFoto.FileName = ""
            dlgFoto.ShowOpen
            imgFoto.Visible = True
            If dlgFoto.FileName <> "" Then
                .Caption = "Guardar Imagen Cabecera"
                strFilenameRuta = dlgFoto.FileName
                imgFoto.Picture = LoadPicture(strFilenameRuta)
            End If
        Else
            
             GuardarFoto strFilenameRuta, "1"
            .Caption = "Editar Imagen Cabecera"
        End If
    End With
    Exit Sub
Else

    MsgBox "Debe generar un codigo para la Caja"
    Exit Sub
End If
errHandler:
MsgBox "Dimensiones Recomendada para la Imagen 350*250 pixeles"
strFilenameRuta = ""
cmdAgregarFoto.Caption = "Editar Imagen Cabecera"
    Exit Sub
End Sub


Public Sub GuardarFoto(ByVal Ruta As String, ByVal tTipo As String)
        
        Dim imgTeacher()      As Byte
        Dim varPhoto          As Variant
        Dim numfile           As Long
        If (Ruta <> "") Then
            varPhoto = FileLen(Ruta)
            ReDim bufimages(varPhoto - 1) As Byte
            numfile = FreeFile
            Open Ruta For Binary As #numfile
            Get #numfile, , bufimages
            Close #numfile
             imgTeacher = bufimages
        End If
        If (Ruta = "") Then
            imgTeacher = LoadResData(101, "CUSTOM")
            varPhoto = UBound(imgTeacher)
        End If
        Dim lnfoto As Variant
        lnfoto = varPhoto
        Dim Cmd As New ADODB.Command
        Dim prm As New ADODB.Parameter
        With Cmd
                .ActiveConnection = Cn
                .CommandText = "sp_UpdImagenCaja"
                .CommandType = adCmdStoredProc
        End With
        Set prm = Cmd.CreateParameter("@tCodigo", adChar, adParamInput, 10, txtCodigo.Text)
        Cmd.Parameters.Append prm
         Set prm = Cmd.CreateParameter("@tTipo", adChar, adParamInput, 1, tTipo)
        Cmd.Parameters.Append prm
        Set prm = Cmd.CreateParameter("@oFoto", adLongVarBinary, adParamInput, lnfoto + 1)
        Cmd.Parameters.Append prm
        
        If Not IsNull(imgTeacher) Then
            prm.AppendChunk imgTeacher
        Else
            prm.value = Null
        End If
        Cmd.Execute
End Sub

Private Sub cmdAgregarFotoPie_Click()
'on error GoTo ErrHandler
If txtCodigo.Text <> "" Then
    dlgFotoPie.CancelError = False
    With cmdAgregarFotoPie
        If .Caption = "Editar Imagen Pie" Then
            dlgFotoPie.Filter = "Image(*.jpg)|*.jpg|Image(*.gif)| *.gif" '"archivos (*.bmp)|*.bmp"
            dlgFotoPie.FileName = ""
            dlgFotoPie.ShowOpen
            imgFotoPie.Visible = True
            If dlgFotoPie.FileName <> "" Then
                .Caption = "Guardar Imagen Pie"
                strFilenameRutaPie = dlgFotoPie.FileName
                imgFotoPie.Picture = LoadPicture(strFilenameRutaPie)
            End If
        Else
            
             GuardarFoto strFilenameRutaPie, "2"
            .Caption = "Editar Imagen Pie"
        End If
    End With
    Exit Sub
Else

    MsgBox "Debe generar un codigo para la Caja"
    Exit Sub
End If
errHandler:
MsgBox "Dimensiones Recomendada para la Imagen 350*250 pixeles"
strFilenameRutaPie = ""
cmdAgregarFoto.Caption = "Editar Imagen Pie"
End Sub

Private Sub cmdGuardarBalanza_Click()
On Error GoTo fin
    If Calcular("select COUNT(*) AS CODIGO from TCONFIGURAPERIFERICO where tcaja='" & txtCodigo.Text & "' and tTabla='BALANZA'", Cn) = 0 Then
        Isql = "insert into TCONFIGURAPERIFERICO (tTabla,tcaja,nDato1,nDato2,nDato3,nDato4,nDato5,nDato6,nDato7,nDato8, lActivo)" & _
                "Values ('BALANZA', '" & txtCodigo.Text & "', '" & Me.cboBal1.Text & "', '" & Me.cboBal2.Text & "', '" & Me.cboBal3.Text & _
                "', '" & Me.cboBal4.Text & "', '" & Me.cboBal5.Text & "', '" & txtBalanzaPuerto.Text & "', '" & Me.txtBalcomando.Text & "', '" & Me.txtbaltiempo.Text & "'," & Me.chkActivoBal.value & ")"
                
        Cn.Execute Isql
    Else
        Isql = "update TCONFIGURAPERIFERICO set nDato1='" & Me.cboBal1.Text & "',nDato2='" & Me.cboBal2.Text & "',nDato3='" & Me.cboBal3.Text & _
        "',nDato4='" & Me.cboBal4.Text & "',nDato5='" & Me.cboBal5.Text & "',nDato6='" & txtBalanzaPuerto.Text & "',nDato7='" & Me.txtBalcomando.Text & "'" & _
        ",nDato8='" & Me.txtbaltiempo.Text & "', lActivo= " & Me.chkActivoBal.value & "  where tTabla='BALANZA' and tcaja='" & txtCodigo.Text & "'"
        Cn.Execute Isql
    End If
    MsgBox "Proceso Correcto", vbInformation, "Inforest"
Exit Sub
fin:
End Sub
Private Sub asignarBalanza()
On Error GoTo fin
    Dim RsBalanza As Recordset
    Isql = "Select * from vBalanza where tCaja = '" & txtCodigo.Text & "'"
    Set RsBalanza = Lib.OpenRecordset(Isql, Cn)
    
    If Not (RsBalanza.EOF Or RsBalanza.BOF) Then
        RsBalanza.MoveFirst
        With RsBalanza
             Me.cboBal1.Text = IIf(IsNull(!nDato1), "", !nDato1)
             Me.cboBal2.Text = IIf(IsNull(!nDato2), "", !nDato2)
             Me.cboBal3.Text = IIf(IsNull(!nDato3), "", !nDato3)
             Me.cboBal4.Text = IIf(IsNull(!nDato4), "", !nDato4)
             Me.cboBal5.Text = IIf(IsNull(!nDato5), "", !nDato5)
             Me.txtBalanzaPuerto.Text = IIf(IsNull(!nDato6), "", !nDato6)
             Me.txtBalcomando.Text = IIf(IsNull(!nDato7), "", !nDato7)
             Me.txtbaltiempo.Text = IIf(IsNull(!nDato8), "", !nDato8)
             Me.chkActivoBal.value = IIf(IsNull(!lActivo), 0, IIf(!lActivo = True, 1, 0))
        End With
    Else
         Me.cboBal1.Text = ""
         Me.cboBal2.Text = ""
         Me.cboBal3.Text = ""
         Me.cboBal4.Text = ""
         Me.cboBal5.Text = ""
         Me.txtBalanzaPuerto.Text = ""
         Me.txtbaltiempo.Text = ""
         Me.txtBalcomando.Text = ""
         Me.chkActivoBal.value = 0
    End If
Exit Sub
fin:

End Sub

Private Sub cmdNavegar_Click(Index As Integer)
    Select Case Index
           Case Is = 0 'Primero
                MoverPuntero Primero, frmCaja.grdGrilla
           Case Is = 1 'PgUp
                MoverPuntero pgup, frmCaja.grdGrilla
           Case Is = 2 'Previo
                MoverPuntero previo, frmCaja.grdGrilla
           Case Is = 3 'Siguiente
                MoverPuntero siguiente, frmCaja.grdGrilla
           Case Is = 4 'PgDn
                MoverPuntero pgdn, frmCaja.grdGrilla
           Case Is = 5 'Ultimo
                MoverPuntero Ultimo, frmCaja.grdGrilla
    End Select
   Asignar
   cmdTexto.Caption = "Registro " & IIf(frmCaja.RsCabecera.RecordCount = 0, 0, frmCaja.RsCabecera.AbsolutePosition) & " de " & frmCaja.RsCabecera.RecordCount
End Sub

Private Sub cmdOpcion_Click(Index As Integer)

   Select Case Index
          Case Is = 0 ' Agregar
               Sw = True
               ActivarBotones (False)
               Blanquear Me
               inicio
               'Cambia el Nombre del Primer Text
               txtDetallado.SetFocus
          
          Case Is = 1 ' Grabar
               Dim nCorrela As String
               'Chequea Datos
               If txtDetallado.Text = "" Then MsgBox "Ingrese la Descripción Detallada", vbExclamation, sMensaje: txtDetallado.SetFocus: Exit Sub
                                   
               If Sw Then
                  Sw = False
                  
                  'Asignar El Campo de Codificación
                  nCorrela = Calcular("select max(tCaja) as Codigo from TCAJA", Cn)
                  If IsNull(nCorrela) Or nCorrela = "" Then
                     txtCodigo = "001"
                  Else
                     txtCodigo = Lib.Correlativo(nCorrela, 3)
                  End If
                                     
               sPasa = txtCodigo.Text
               
               'Inserta Movimiento auditoria
               lAuditoria = RegistraMovimientoAuditoria(tModuloSeg, sMDB, "TCAJA", "CAJA", "01", sUsuario, sPasa, "", _
                            "tcaja", "CodigoCaja", sPasa, "tdescripcion", "Descripcion Detallada", txtDetallado.Text, "tPrecuenta", "Impresora Precuenta", cboPreCuenta.BoundText, "lSiab", "Flag Activacion enlace Sistema SIAB", IIf(chkSiab.value, "Verdadero", "Falso"), _
                            "tGrupo", "Grupo Predeterminado", cboGrupo.BoundText, "tTipoPedido", "Tipo Pedido Predeterminado", cboTipoPedido.BoundText, _
                            "tUnidadNegocio", "Unidad de Negocio", cboUnidadNegocio.BoundText, "tSucursal", "Sucursal", cboSucursal.BoundText, _
                            "tSubAlmacen", "Area de Produccion", IIf(cboAreaProduccion.BoundText = "ABC", "", cboAreaProduccion.BoundText), "lActivo", "Flag Activo", IIf(chkActivo.value, "Verdadero", "Falso"), _
                            "nPuerto", "Visor Puerto", Val(txtPuerto.Text), "tMensaje1", "Visor Mensaje 1", txtMensaje1.Text, "tMensaje2", "Visor Mensaje 2", txtMensaje2.Text, _
                            "nBalanzaPuerto", "Balanza Electronica Puerto", Val(txtBalanzaPuerto.Text), "lVisaNet", "Flag Enlace VisaNet", IIf(chkVisaNet.value, "Verdadero", "Falso"), "nLongitudBarra", "Lector Barra Longitud", txtLongitudBarra.Text, "lcapturaPeso", "Lector de Barra", IIf(opcCapturaPeso.value, "Verdadero", "Falso"), _
                            "lComanda", "Flag Obligatoriedad Comanda Manual", IIf(chkComanda.value, "Verdadero", "Falso"), "lObligaPrinter", "Flag Obliga Impresion Pedido", IIf(chkObligaPrinter.value, "Verdadero", "Falso"), "lObligaPrecuenta", "Flag Obliga Impresion Precuenta", IIf(chkObligaPrecuenta.value, "Verdadero", "Falso"), "lCancelacion", "Flag Obligatoriedad de Cancelacion", IIf(chkCancelacion.value, "Verdadero", "Falso"), "lObservacion", "Flag Obligatoriedad Observacion", IIf(chkObservacion.value, "Verdadero", "Falso"), "lConsumo1", "Flag Emision Rapida por Consumo", IIf(chkConsumo1.value, "Verdadero", "Falso"), "lConsumo2", "Flag Emision por Consumo En Pagos y Division", IIf(chkConsumo2.value, "Verdadero", "Falso"), "lConsumo3", "Flag Emision por Consumo en Caja Rapida", IIf(chkConsumo3.value, "Verdadero", "Falso"), _
                            "lPrecuentaAgrupada", "Flag Impresion Agrupada en Precuenta", IIf(chkAgrupada.value, "Verdadero", "Falso"), "lDocumentoAgrupado", "Flag Impresion Agrupada en Documentos", IIf(chkDocumentoAgrupado.value, "Verdadero", "Falso"), "lComboPrecuenta", "Flag Impresion de Combos en Precuenta", IIf(chkComboPrecuenta.value, "Verdadero", "Falso"), "lComboDocumento", "Flag Impresion de Combos en Documentos", IIf(chkComboDocumento.value, "Verdadero", "Falso"), "lPropiedadPrecuenta", "Flag Impresion Propiedad en Precuenta", IIf(chkPropiedadPrecuenta.value, "Verdadero", "Falso"), "lObservacionPrecuenta", "Flag Impresion de Observacion en Precuenta", IIf(chkObservacionPrecuenta.value, "Verdadero", "Falso"), "lPropiedadDocumento", "Flag Impresion Propiedad en Documento", IIf(chkPropiedadDocumento.value, "Verdadero", "Falso"), "lObservacionDocumento", "Flag Impresion Observacion en Documento", IIf(chkObservacionDocumento.value, "Verdadero", "Falso"), _
                            "lPrecioNetoPrecuenta", "Flag Impresion Prec Neto en Precuenta", IIf(chkPrecioNetoPrecuenta.value, "Verdadero", "Falso"), "lPrecuenta", "Flag Permite Cambiar Impresora Precuenta", IIf(chkPreCuenta.value, "Verdadero", "Falso"), "lImpuestoPrecuenta", "Flag Impresion Impuesto Desglos. Precuenta", IIf(chkImpuestoPrecuenta.value, "Verdadero", "Falso"), "lCambioMesa", "Flag Impresion de Cambio de Mesa", IIf(chkCambioMesa.value, "Verdadero", "Falso"), "lValorCortesia", "Flag Impresion Valorizada de Cortesias", IIf(chkValor.value, "Verdadero", "Falso"), "lequivadolaprecuenta", "Flag Impresion de Equivalencia Dolares en Precuenta", IIf(chkEquivaPrecuenta.value, "Verdadero", "Falso"), "lActivaImpDscAlternativa", "Flag Impresion Descripcion Alternativa", IIf(chkDescripcionAlternativa.value, "Verdadero", "Falso"), "nLimitePrecuenta", "Limite de Precuentas", Val(txtLimitePrecuenta.Text), "nLimiteReimpresion", "Limite de Re Impresiones Pedido", Val(txtLimiteReimpresion.Text), _
                            "vComanda", "Flag Activa Ingreso Comanda Manual", IIf(chkVComanda.value, "Verdadero", "Falso"), "lMotivoEliminaC", "Flag Pide Motivo Elimina Pedido", IIf(chkEliminaC.value, "Verdadero", "Falso"), "lPasswordC", "Flag Activa Password Eliminacion Pedido", IIf(chkPasswordC.value, "Verdadero", "Falso"), "lMotivoElimina", "Flag Pide Motivo Elimina Producto", IIf(chkElimina.value, "Verdadero", "Falso"), "lPassword", "Flag Activa Password Elimina Producto", IIf(chkPassword.value, "Verdadero", "Falso"), "lObligacierre", "Flag Activa Password Cierre Turno", IIf(chkObligaCierre.value, "Verdadero", "Falso"), "lPasswordTransferencia", "Flag Activa Password Transferencia", IIf(chkPasswordTransferencia.value, "Verdadero", "Falso"), "lpasswordporcobrar", "Flag Activa Password Por Cobrar", IIf(chkPasswordPorCobrar.value, "Verdadero", "Falso"), "lPasswordImportarPedido", "Flag Activa Password Importar Pedido", IIf(chkPasswordImportar.value, "Verdadero", "Falso"), _
                            "lFiltroTipoPedido", "Flag Permite Importar Pedidos por Canal", IIf(chkFiltroTipoPedido.value, "Verdadero", "Falso"), "lAdicion", "Flag Permite Transferencias", IIf(chkAdicion.value, "Verdadero", "Falso"), "lmodificatipopedido", "Flag Permite Modificar Tipo de Pedido", IIf(chkModificaTipoPedido.value, "Verdadero", "Falso"), "lCajaRapida", "Flag Ingreso Directo a Caja Rapida", IIf(chkCajaRapida.value, "Verdadero", "Falso"), "lPagoRapido", "Flag Ingreso a Pago Rapido desde Caja Rapida", IIf(chkPagoRapido.value, "Verdadero", "Falso"), "lOrden", "Flag Activa Control Enum. Automatica", IIf(chkOrden.value, "Verdadero", "Falso"), "lDirecto", "Flag Activa Control de Envios Directos", IIf(chkDirecto.value, "Verdadero", "Falso"), "lDisgrega", "Flag Disgregar en Dos Partes", IIf(chkDisgrega.value, "Verdadero", "Falso"), "lCD", "Flag Activa Caja Central Delivery", IIf(chkCD.value, "Verdadero", "Falso"), "lCCVOX", "Flag Activa Caja Delivery CCVOX", IIf(chkCCVOX.value, "Verdadero", "Falso"), _
                            "lMultiCajero", "Flag Activa Multicajero Caja Rapida", IIf(chkMultiCajero.value, "Verdadero", "Falso"), "lMCPV", "Flag Activa Multicajero Salon", IIf(chkMCPV.value, "Verdadero", "Falso"), "lCompatibilidadTVS", "Flag Permite Compatibilidad con TVS", IIf(chkCompatibilidadTVS.value, "Verdadero", "Falso"), "lPagoRapidoPV", "Flag Ingreso a Pago Rapido desde Punto Venta", IIf(chkPagoRapidopv.value, "Verdadero", "Falso"), "tTextoConsumo", "Motivo de Consumo Predeterminado", txtTextoConsumo.Text, "tSectorVenta", "SectorVenta", cboSectorVenta.BoundText, "lCajaMobile", "Flag Caja Mobile", IIf(chkCajaMobile.value, "Verdadero", "Falso"), "lBloqueaPrecuenta", "Bloquear Precuenta", IIf(chkBloqueaPrecuenta.value, "Verdadero", "Falso"), "lRotulado", "Enlace Rotulado", IIf(chkRotulado.value, "Verdadero", "Falso"), "lMultiAreaSubGrupo", "Flag Multi Area Por SubGrupo", IIf(Me.chkMulti2.value, "Verdadero", "Falso"), _
                            "lMultiAreaCaja", "Flag Multi Area ", IIf(Me.chkMulti1.value, "Verdadero", "Falso"), "lHuella", "Flag Huella ", IIf(Me.chkHuella.value, "Verdadero", "Falso"), "lImprimeImagCabPrecuenta", "Imagen Cabecera Precuenta", IIf(Me.chkImprimeImagCabPrecuenta.value, "Verdadero", "Falso"), "lImprimeImagpiePrecuenta", "Imagen Pie Precuenta", IIf(Me.chkImprimeImagPiePrecuenta.value, "Verdadero", "Falso"), "lAccesoDespachoPedido", "Acceso Despacho Pedido", IIf(Me.chkAccesoDespachoPedido.value, "Verdadero", "Falso"), "LBuscarpedidovisualizargrilla", "Buscar Pedido Visualizar Grilla", IIf(Me.chkBuscaPedidoVisualizaGrilla.value, "Verdadero", "Falso"), "lbuscarpedidofiltrarmesa", "Buscar Pedido Filtrar Mesa", IIf(Me.chkBuscaPedidoFiltrarMesa.value, "Verdadero", "Falso"), "lMotivoDescuento", "Imprime Motivo Descuento", IIf(Me.chkMotDesc.value, "Verdadero", "Falso"), "lCajaContingencia", "Activa Caja Contingencia", IIf(Me.chkCajaContingencia.value, "Verdadero", "Falso"), _
                            "lImpPropina", "Solicita Propina Imp Prec", IIf(Me.chkImpPropina.value, "Verdadero", "Falso"), "lImpComandaf2", "imprime comanda formato 2", IIf(Me.chkComandaF2.value, "Verdadero", "Falso"), "lPassOtrosPagos", "Activa Password otros pagos", IIf(Me.chkPassOtrosPagos.value, "Verdadero", "Falso"))
                
                'La Funcion RegistraMovimientoAuditoria devuelve true si se ejecuto correctamente.
                  If lAuditoria = False Then
                    Screen.MousePointer = vbDefault
                    Exit Sub
                  End If
                                                                            
                  'Cambiar el SQL
                  Isql = "insert into TCAJA( " & _
                  "tCaja, tDescripcion, tPrecuenta, vComanda, lComanda, lMotivoEliminaC, lMotivoElimina, lComboPrecuenta, lComboDocumento, lPasswordC, lPassword, tGrupo, lConsumo1, lConsumo2, lConsumo3, lPrecuenta, lAdicion, lPrecuentaAgrupada, tTipoPedido, lObliga, lMozo, lObligaPrinter, lObligaPrecuenta, lPax, lObligaCierre, lFiltroTipoPedido, nPuerto, tMensaje1, tMensaje2, lCancelacion, lDirecto, lCambioMesa, lVisaNet, lImpuestoPrecuenta, lDocumentoAgrupado, lOrden, lActivo, lValorCortesia, lObservacion, lCajaRapida, lPropiedadDocumento, lPropiedadPrecuenta, lPrecioNetoPrecuenta, nLimitePrecuenta, tUnidadNegocio, lPasswordTransferencia, nLimiteReimpresion, lCD, lMultiCajero, lFechaEntregaDelivery, lMCPV, lCCVOX, lMotorizado,lequivadolaprecuenta,tsubalmacen, lObservacionPrecuenta, lObservacionDocumento,lPasswordImportarPedido, lActivaImpDscAlternativa, lCompatibilidadTVS, nLongitudBarra,lpagorapido, lDisgrega,lpasswordporcobrar,lmodificatipopedido,TSUCURSAL, nBalanzaPuerto, lCapturaPeso, " & _
                  "lPagoRapidoPV, tTextoConsumo, lSiab, tSectorVenta,lcajamobile, lbloqueaprecuenta, lRotulado, lmultiAreaSubGrupo, lMultiAreaCaja, lHuella, lImprimeImagCabPrecuenta, lImprimeImagpiePrecuenta, laccesodespachopedido,lBuscaPedidoNumero,lCodigoReciboIngreso,lPagoRapidoMod,lWebAp,lMesa247,lConsumo4, lPrecuentaNoValorizada, LBuscarpedidovisualizargrilla ,lbuscarpedidofiltrarmesa, lClaveEnvioProduccion, EAN13,lObservacionCabDoc,tCompMesa247,lMotivoDescuento, lCajaContingencia, lImpPropina, lImpComandaf2, lPassOtrosPagos) " & _
                   "values ('" & txtCodigo.Text & "', " & _
                         " '" & txtDetallado.Text & "', " & _
                         " '" & cboPreCuenta.BoundText & "', " & _
                                chkVComanda.value & ", " & chkComanda.value & ", " & chkEliminaC.value & ", " & chkElimina.value & ", " & _
                                chkComboPrecuenta.value & ", " & chkComboDocumento.value & ", " & _
                                chkPasswordC.value & ", " & _
                                chkPassword.value & ", " & _
                         " '" & cboGrupo.BoundText & "', " & _
                                chkConsumo1.value & ", " & _
                                chkConsumo2.value & ", " & _
                                chkConsumo3.value & ", " & _
                                chkPreCuenta.value & ", " & _
                                chkAdicion.value & ", " & chkAgrupada.value & ", " & _
                         " '" & cboTipoPedido.BoundText & "',null, null,  " & _
                                chkObligaPrinter.value & ", " & chkObligaPrecuenta.value & ",null, " & _
                                chkObligaCierre.value & ", " & chkFiltroTipoPedido.value & ", " & _
                                Val(txtPuerto.Text) & ", '" & txtMensaje1 & "', '" & txtMensaje1 & "', " & _
                                chkCancelacion.value & ", " & chkDirecto.value & ", " & chkCambioMesa.value & ", " & chkVisaNet.value & ", " & chkImpuestoPrecuenta.value & "," & chkDocumentoAgrupado.value & ", " & chkOrden.value & ", " & chkActivo.value & "," & chkValor.value & ", " & _
                                chkObservacion.value & ", " & chkCajaRapida.value & ", " & chkPropiedadDocumento.value & ", " & chkPropiedadPrecuenta.value & ", " & chkPrecioNetoPrecuenta.value & ", " & Val(txtLimitePrecuenta.Text) & ", '" & cboUnidadNegocio.BoundText & "', " & Val(txtLimitePrecuenta.Text) & ", " & chkPassword.value & ", " & chkCD.value & ", " & chkMultiCajero.value & ", null, " & chkMCPV.value & ", " & chkCCVOX.value & ", null," & chkEquivaPrecuenta.value & ",'" & IIf(cboAreaProduccion.BoundText = "ABC", "", cboAreaProduccion.BoundText) & "'," & chkObservacionPrecuenta.value & "," & chkObservacionDocumento.value & "," & chkPasswordImportar.value & "," & chkDescripcionAlternativa.value & "," & chkCompatibilidadTVS.value & ", '" & txtLongitudBarra.Text & "'," & chkPagoRapido.value & "," & chkDisgrega.value & "," & chkPasswordPorCobrar.value & "," & chkModificaTipoPedido.value & ", '" & _
                                cboSucursal.BoundText & "', " & Val(txtBalanzaPuerto) & ", " & IIf(opcCapturaPeso.value, 1, 0) & ", " & chkPagoRapidopv.value & ",'" & txtTextoConsumo.Text & "', " & chkSiab.value & ", '" & cboSectorVenta.BoundText & "', " & chkCajaMobile.value & ", " & chkBloqueaPrecuenta.value & "," & chkRotulado.value & "," & chkMulti2.value & ", " & chkMulti1.value & ", " & chkHuella.value & "," & chkImprimeImagCabPrecuenta.value & "," & chkImprimeImagPiePrecuenta.value & ", " & chkAccesoDespachoPedido.value & "," & Me.chkBuscaPedido.value & ", " & Me.chkCodigoReciboIngreso.value & "," & chkPagoRapidoMod.value & "," & chkWebAp.value & "," & chkMesa247.value & "," & chkConsumo4.value & ", " & chkPrecuentaNoValorizada.value & "," & Me.chkBuscaPedidoVisualizaGrilla.value & "," & Me.chkBuscaPedidoFiltrarMesa.value & ", " & Me.chkClaveEnvio.value & " , " & IIf(chkEAN13.value, 1, 0) & "," & _
                                IIf(chkObservacionCabDoc.value, 1, 0) & ", '" & cboComprobante.BoundText & "', " & Me.chkMotDesc.value & "," & Me.chkCajaContingencia.value & ", " & Me.chkImpPropina.value & ", " & Me.chkComandaF2.value & ", " & Me.chkPassOtrosPagos.value & ")"
                  Cn.Execute Isql
                  
                  frmCaja.RsCabecera.Requery
                  frmCaja.RsCabecera.MoveLast
                  ActivarBotones (True)
                  cmdTexto.Caption = "Registro " & IIf(frmCaja.RsCabecera.RecordCount = 0, 0, frmCaja.RsCabecera.AbsolutePosition) & " de " & frmCaja.RsCabecera.RecordCount
                               
'                  'Inserta Movimiento auditoria
'                  lAuditoria = RegistraMovimientoAuditoria(tModuloSeg, sMDB, "TTIPODOCUMENTOIMPRESORA", "TIPODOCUMENTOIMPRESORA", "01", sUsuario, sPasa, "", _
'                  "tcaja", "CodigoCaja", sPasa, "tTipoEmision", "Tipo Emision", "01", "tImpresora", "Impresora", "", _
'                  "tDescripcion", "Descripcion", "Cortesia", "tFormulario", "Formulario", "01", _
'                  "tSerie", "Serie", "00000", "tUltimoNumero", "UltimoNumero", "000000000", _
'                  "tUsuario", "Usuario", sUsuario, "lResumen", "Flag Resumen", "Verdadero", _
'                  "lEquivaleDolares", "Equivale Dolares", "Falso")
'
'                  If lAuditoria = False Then
'                     Screen.MousePointer = vbDefault
'                     Exit Sub
'                  End If
'
'                  Isql = "insert into TTIPODOCUMENTOIMPRESORA( " & _
'                         "tCaja, tTipoEmision, tImpresora, tDescripcion, tFormulario, tSerie, tUltimoNumero, tUsuario, lResumen, fRegistro,lEquivaDolares) " & _
'                         "values (  '" & txtCodigo.Text & "', " & _
'                                  " '00', " & _
'                                  " '', " & _
'                                  " 'Cortesía', " & _
'                                  " '01', " & _
'                                  " '00000', " & _
'                                  " '000000000', " & _
'                                  " '" & sUsuario & "', " & _
'                                  " 1, " & _
'                                  " getdate(),0 )"
'
'                   Cn.Execute Isql

                   RsGrilla.Requery
                   RsGrilla.Filter = "tCaja ='" & txtCodigo.Text & "'"
                   RsAI.Filter = "tCaja ='" & txtCodigo.Text & "'"

                   If Not RsGrilla.EOF Then
                      RsGrilla.MoveFirst
                   End If
                   MsgBox "Registro Agregado", vbInformation, sMensaje
                       
               Else
                   sPasa = txtCodigo.Text
               
                   'Inserta Movimiento auditoria
                   lAuditoria = RegistraMovimientoAuditoria(tModuloSeg, sMDB, "TCAJA", "CAJA", "02", sUsuario, sPasa, "", _
                            "tcaja", "CodigoCaja", sPasa, "tdescripcion", "Descripcion Detallada", txtDetallado.Text, "tPrecuenta", "Impresora Precuenta", cboPreCuenta.BoundText, "lSiab", "Flag Activacion enlace Sistema SIAB", IIf(chkSiab.value, "Verdadero", "Falso"), _
                            "tGrupo", "Grupo Predeterminado", cboGrupo.BoundText, "tTipoPedido", "Tipo Pedido Predeterminado", cboTipoPedido.BoundText, _
                            "tUnidadNegocio", "Unidad de Negocio", cboUnidadNegocio.BoundText, "tSucursal", "Sucursal", cboSucursal.BoundText, _
                            "tSubAlmacen", "Area de Produccion", IIf(cboAreaProduccion.BoundText = "ABC", "", cboAreaProduccion.BoundText), "lActivo", "Flag Activo", IIf(chkActivo.value, "Verdadero", "Falso"), _
                            "nPuerto", "Visor Puerto", Val(txtPuerto.Text), "tMensaje1", "Visor Mensaje 1", txtMensaje1.Text, "tMensaje2", "Visor Mensaje 2", txtMensaje2.Text, _
                            "nBalanzaPuerto", "Balanza Electronica Puerto", Val(txtBalanzaPuerto.Text), "lVisaNet", "Flag Enlace VisaNet", IIf(chkVisaNet.value, "Verdadero", "Falso"), "nLongitudBarra", "Lector Barra Longitud", txtLongitudBarra.Text, "lcapturaPeso", "Lector de Barra", IIf(opcCapturaPeso.value, "Verdadero", "Falso"), _
                            "lComanda", "Flag Obligatoriedad Comanda Manual", IIf(chkComanda.value, "Verdadero", "Falso"), "lObligaPrinter", "Flag Obliga Impresion Pedido", IIf(chkObligaPrinter.value, "Verdadero", "Falso"), "lObligaPrecuenta", "Flag Obliga Impresion Precuenta", IIf(chkObligaPrecuenta.value, "Verdadero", "Falso"), "lCancelacion", "Flag Obligatoriedad de Cancelacion", IIf(chkCancelacion.value, "Verdadero", "Falso"), "lObservacion", "Flag Obligatoriedad Observacion", IIf(chkObservacion.value, "Verdadero", "Falso"), "lConsumo1", "Flag Emision Rapida por Consumo", IIf(chkConsumo1.value, "Verdadero", "Falso"), "lConsumo2", "Flag Emision por Consumo En Pagos y Division", IIf(chkConsumo2.value, "Verdadero", "Falso"), "lConsumo3", "Flag Emision por Consumo en Caja Rapida", IIf(chkConsumo3.value, "Verdadero", "Falso"), _
                            "lPrecuentaAgrupada", "Flag Impresion Agrupada en Precuenta", IIf(chkAgrupada.value, "Verdadero", "Falso"), "lDocumentoAgrupado", "Flag Impresion Agrupada en Documentos", IIf(chkDocumentoAgrupado.value, "Verdadero", "Falso"), "lComboPrecuenta", "Flag Impresion de Combos en Precuenta", IIf(chkComboPrecuenta.value, "Verdadero", "Falso"), "lComboDocumento", "Flag Impresion de Combos en Documentos", IIf(chkComboDocumento.value, "Verdadero", "Falso"), "lPropiedadPrecuenta", "Flag Impresion Propiedad en Precuenta", IIf(chkPropiedadPrecuenta.value, "Verdadero", "Falso"), "lObservacionPrecuenta", "Flag Impresion de Observacion en Precuenta", IIf(chkObservacionPrecuenta.value, "Verdadero", "Falso"), "lPropiedadDocumento", "Flag Impresion Propiedad en Documento", IIf(chkPropiedadDocumento.value, "Verdadero", "Falso"), "lObservacionDocumento", "Flag Impresion Observacion en Documento", IIf(chkObservacionDocumento.value, "Verdadero", "Falso"), _
                            "lPrecioNetoPrecuenta", "Flag Impresion Prec Neto en Precuenta", IIf(chkPrecioNetoPrecuenta.value, "Verdadero", "Falso"), "lPrecuenta", "Flag Permite Cambiar Impresora Precuenta", IIf(chkPreCuenta.value, "Verdadero", "Falso"), "lImpuestoPrecuenta", "Flag Impresion Impuesto Desglos. Precuenta", IIf(chkImpuestoPrecuenta.value, "Verdadero", "Falso"), "lCambioMesa", "Flag Impresion de Cambio de Mesa", IIf(chkCambioMesa.value, "Verdadero", "Falso"), "lValorCortesia", "Flag Impresion Valorizada de Cortesias", IIf(chkValor.value, "Verdadero", "Falso"), "lequivadolaprecuenta", "Flag Impresion de Equivalencia Dolares en Precuenta", IIf(chkEquivaPrecuenta.value, "Verdadero", "Falso"), "lActivaImpDscAlternativa", "Flag Impresion Descripcion Alternativa", IIf(chkDescripcionAlternativa.value, "Verdadero", "Falso"), "nLimitePrecuenta", "Limite de Precuentas", Val(txtLimitePrecuenta.Text), "nLimiteReimpresion", "Limite de Re Impresiones Pedido", Val(txtLimiteReimpresion.Text), _
                            "vComanda", "Flag Activa Ingreso Comanda Manual", IIf(chkVComanda.value, "Verdadero", "Falso"), "lMotivoEliminaC", "Flag Pide Motivo Elimina Pedido", IIf(chkEliminaC.value, "Verdadero", "Falso"), "lPasswordC", "Flag Activa Password Eliminacion Pedido", IIf(chkPasswordC.value, "Verdadero", "Falso"), "lMotivoElimina", "Flag Pide Motivo Elimina Producto", IIf(chkElimina.value, "Verdadero", "Falso"), "lPassword", "Flag Activa Password Elimina Producto", IIf(chkPassword.value, "Verdadero", "Falso"), "lObligacierre", "Flag Activa Password Cierre Turno", IIf(chkObligaCierre.value, "Verdadero", "Falso"), "lPasswordTransferencia", "Flag Activa Password Transferencia", IIf(chkPasswordTransferencia.value, "Verdadero", "Falso"), "lpasswordporcobrar", "Flag Activa Password Por Cobrar", IIf(chkPasswordPorCobrar.value, "Verdadero", "Falso"), "lPasswordImportarPedido", "Flag Activa Password Importar Pedido", IIf(chkPasswordImportar.value, "Verdadero", "Falso"), _
                            "lFiltroTipoPedido", "Flag Permite Importar Pedidos por Canal", IIf(chkFiltroTipoPedido.value, "Verdadero", "Falso"), "lAdicion", "Flag Permite Transferencias", IIf(chkAdicion.value, "Verdadero", "Falso"), "lmodificatipopedido", "Flag Permite Modificar Tipo de Pedido", IIf(chkModificaTipoPedido.value, "Verdadero", "Falso"), "lCajaRapida", "Flag Ingreso Directo a Caja Rapida", IIf(chkCajaRapida.value, "Verdadero", "Falso"), "lPagoRapido", "Flag Ingreso a Pago Rapido desde Caja Rapida", IIf(chkPagoRapido.value, "Verdadero", "Falso"), "lOrden", "Flag Activa Control Enum. Automatica", IIf(chkOrden.value, "Verdadero", "Falso"), "lDirecto", "Flag Activa Control de Envios Directos", IIf(chkDirecto.value, "Verdadero", "Falso"), "lDisgrega", "Flag Disgregar en Dos Partes", IIf(chkDisgrega.value, "Verdadero", "Falso"), "lCD", "Flag Activa Caja Central Delivery", IIf(chkCD.value, "Verdadero", "Falso"), "lCCVOX", "Flag Activa Caja Delivery CCVOX", IIf(chkCCVOX.value, "Verdadero", "Falso"), _
                            "lMultiCajero", "Flag Activa Multicajero Caja Rapida", IIf(chkMultiCajero.value, "Verdadero", "Falso"), "lMCPV", "Flag Activa Multicajero Salon", IIf(chkMCPV.value, "Verdadero", "Falso"), "lCompatibilidadTVS", "Flag Permite Compatibilidad con TVS", IIf(chkCompatibilidadTVS.value, "Verdadero", "Falso"), "lPagoRapidoPV", "Flag Ingreso a Pago Rapido desde Punto Venta", IIf(chkPagoRapidopv.value, "Verdadero", "Falso"), "tTextoConsumo", "Motivo de Consumo Predeterminado", txtTextoConsumo.Text, "tSectorVenta", "SectorVenta", cboSectorVenta.BoundText, "lCajaMobile", "Flag Caja Mobile", IIf(chkCajaMobile.value, "Verdadero", "Falso"), "lBloqueaPrecuenta", "Bloquea Precuenta", IIf(chkBloqueaPrecuenta.value, "Verdadero", "Falso"), "lRotulado", "Enlace Rotulado", IIf(chkRotulado.value, "Verdadero", "Falso"), "lMultiAreaSubGrupo", "Flag Multi Area Por SubGrupo", IIf(Me.chkMulti2.value, "Verdadero", "Falso"), _
                            "lMultiAreaCaja", "Flag Multi Area ", IIf(Me.chkMulti1.value, "Verdadero", "Falso"), "lHuella", "Flag Huella ", IIf(Me.chkHuella.value, "Verdadero", "Falso"), "lImprimeImagCabPrecuenta", "Imagen Cabecera Precuenta", IIf(Me.chkImprimeImagCabPrecuenta.value, "Verdadero", "Falso"), "lImprimeImagpiePrecuenta", "Imagen Pie Precuenta", IIf(Me.chkImprimeImagPiePrecuenta.value, "Verdadero", "Falso"), "lAccesoDespachoPedido", "Acceso Despacho Pedido", IIf(Me.chkAccesoDespachoPedido.value, "Verdadero", "Falso"), "LBuscarpedidovisualizargrilla", "Buscar Pedido Visualizar Grilla", IIf(Me.chkBuscaPedidoVisualizaGrilla.value, "Verdadero", "Falso"), "lbuscarpedidofiltrarmesa", "Buscar Pedido Filtrar Mesa", IIf(Me.chkBuscaPedidoFiltrarMesa.value, "Verdadero", "Falso"), "lMotivoDescuento", "Imprime Motivo Descuento", IIf(Me.chkMotDesc.value, "Verdadero", "Falso"), "lCajaContingencia", "Activa Caja Contingencia", IIf(Me.chkCajaContingencia.value, "Verdadero", "Falso"), _
                            "lImpPropina", "Solicita Propina Imp Prec", IIf(Me.chkImpPropina.value, "Verdadero", "Falso"), "lImpComandaf2", "imprime comanda formato 2", IIf(Me.chkComandaF2.value, "Verdadero", "Falso"), "lPassOtrosPagos", "Activa Password otros pagos", IIf(Me.chkPassOtrosPagos.value, "Verdadero", "Falso"))
                
                   If lAuditoria = False Then
                      Screen.MousePointer = vbDefault
                      Exit Sub
                   End If
                
                  'Cambiar el SQL
                  Isql = "update TCAJA set " & _
                         "tDescripcion ='" & txtDetallado.Text & "', tPrecuenta ='" & cboPreCuenta.BoundText & "', vComanda =" & chkVComanda.value & ", " & _
                         "lComanda =" & chkComanda.value & ", lMotivoEliminaC =" & chkEliminaC.value & ", lMotivoElimina =" & chkElimina.value & ", " & _
                         "lComboPrecuenta =" & chkComboPrecuenta.value & ", lComboDocumento =" & chkComboDocumento.value & ", lPasswordC =" & chkPasswordC.value & ", " & _
                         "lPassword =" & chkPassword.value & ", lPasswordTransferencia =" & chkPasswordTransferencia.value & ", " & _
                         "tGrupo ='" & cboGrupo.BoundText & "', tUnidadNegocio ='" & cboUnidadNegocio.BoundText & "',tsubalmacen ='" & IIf(cboAreaProduccion.BoundText = "ABC", "", cboAreaProduccion.BoundText) & "', " & _
                         "lConsumo1 =" & chkConsumo1.value & ", lConsumo2 =" & chkConsumo2.value & ", lConsumo3 =" & chkConsumo3.value & ", " & _
                         "lPrecuenta =" & chkPreCuenta.value & ", lDisgrega =" & chkDisgrega.value & "," & _
                         "lActivaImpDscAlternativa =" & chkDescripcionAlternativa.value & ",lcajamobile=" & chkCajaMobile.value & ",  " & _
                         "lAdicion =" & chkAdicion.value & ", lCCVOX =" & chkCCVOX.value & ", lbloqueaprecuenta=" & chkBloqueaPrecuenta.value & ", " & _
                         "lPrecuentaAgrupada =" & chkAgrupada.value & ", lObservacionPrecuenta =" & chkObservacionPrecuenta.value & ",lObservacionDocumento =" & chkObservacionDocumento.value & ",lBuscaPedidoNumero=" & Me.chkBuscaPedido.value & ",  " & _
                         "lObliga =  0 , lMozo =  0 ,lMotorizado =  0 ,lObligaPrinter =" & IIf(chkObligaPrinter.value, 1, 0) & ", lObligaPrecuenta =" & IIf(chkObligaPrecuenta.value, 1, 0) & ", lCodigoReciboIngreso=" & Me.chkCodigoReciboIngreso.value & ", " & _
                         "tTipoPedido ='" & cboTipoPedido.BoundText & "', lPax =0 ,lMultiAreaSubGrupo=" & Me.chkMulti2.value & ", lmultiAreaCaja=" & Me.chkMulti1.value & ",  " & _
                         "lObligacierre =" & chkObligaCierre.value & ", lCancelacion =" & chkCancelacion.value & ", lDirecto =" & chkDirecto.value & ", lDocumentoAgrupado =" & chkDocumentoAgrupado.value & ", " & _
                         "lFiltroTipoPedido=" & chkFiltroTipoPedido.value & ", lMCPV =" & chkMCPV.value & ", lImprimeImagCabPrecuenta=" & chkImprimeImagCabPrecuenta.value & ", lImprimeImagpiePrecuenta=" & chkImprimeImagPiePrecuenta.value & ",  " & _
                         "lActivo =" & chkActivo.value & ", lCambioMesa =" & chkCambioMesa.value & ", lVisaNet =" & chkVisaNet.value & ", lImpuestoPrecuenta =" & chkImpuestoPrecuenta.value & ", " & _
                         "nPuerto =" & Val(txtPuerto.Text) & ", tMensaje1='" & txtMensaje1.Text & "', tMensaje2='" & txtMensaje2.Text & "', lOrden= " & chkOrden.value & ", lValorCortesia=" & chkValor.value & ", nLongitudBarra ='" & txtLongitudBarra.Text & "', " & _
                         "lObservacion =" & chkObservacion.value & ", lCajaRapida =" & chkCajaRapida.value & ", lPropiedadDocumento =" & chkPropiedadDocumento.value & ", lPrecioNetoPrecuenta =" & chkPrecioNetoPrecuenta.value & ", lPropiedadPrecuenta =" & chkPropiedadPrecuenta.value & ", " & "nLimitePrecuenta=" & Val(txtLimitePrecuenta.Text) & ", nLimiteReimpresion=" & Val(txtLimiteReimpresion.Text) & ",lCD =" & chkCD.value & ", lMultiCajero =" & chkMultiCajero.value & ", lFechaEntregaDelivery =0,lequivadolaprecuenta =" & chkEquivaPrecuenta.value & ",lPasswordImportarPedido =" & chkPasswordImportar.value & ", lCompatibilidadTVS =" & chkCompatibilidadTVS.value & ", lPagoRapido =" & chkPagoRapido.value & ", lpasswordporcobrar =" & chkPasswordPorCobrar.value & ", lmodificatipopedido =" & chkModificaTipoPedido.value & ", tSucursal ='" & cboSucursal.BoundText & "', nBalanzaPuerto= " & Val(txtBalanzaPuerto.Text) & ", " & _
                         "lCapturaPeso = " & IIf(opcCapturaPeso.value, 1, 0) & ", laccesodespachopedido=" & chkAccesoDespachoPedido.value & ",  " & _
                         "lPagoRapidoPV = " & chkPagoRapidopv.value & ", LBuscarpedidovisualizargrilla=" & Me.chkBuscaPedidoVisualizaGrilla.value & ", lbuscarpedidofiltrarmesa=" & Me.chkBuscaPedidoFiltrarMesa.value & " ,  " & _
                         "tTextoConsumo = '" & txtTextoConsumo.Text & "', lSiab =" & chkSiab.value & ",lRotulado=" & chkRotulado.value & ", tSectorVenta ='" & cboSectorVenta.BoundText & "', " & _
                         "lPagoRapidoMod = " & chkPagoRapidoMod.value & ", lClaveEnvioProduccion = " & chkClaveEnvio.value & ", EAN13 = " & chkEAN13.value & ", " & _
                         "lWebAp = " & chkWebAp.value & ",lMesa247 = " & chkMesa247.value & ", lConsumo4 = " & chkConsumo4.value & ", lPrecuentaNoValorizada = " & chkPrecuentaNoValorizada.value & ", " & _
                         "lHuella = " & chkHuella.value & ", lObservacionCabDoc = " & IIf(chkObservacionCabDoc.value, 1, 0) & ", tCompMesa247='" & cboComprobante.BoundText & "', lMotivoDescuento=" & Me.chkMotDesc.value & ", lCajaContingencia= " & Me.chkCajaContingencia.value & ", lImpPropina= " & Me.chkImpPropina.value & ", lImpComandaf2=" & Me.chkComandaF2.value & ", lPassOtrosPagos = " & Me.chkPassOtrosPagos.value & " " & _
                         " where tCaja = '" & txtCodigo & "'"
                                                
                   Cn.Execute Isql
                   nPos = frmCaja.RsCabecera.Bookmark
                   frmCaja.RsCabecera.Requery
                   If frmCaja.RsCabecera.RecordCount = 0 Then
                      frmCaja.RsCabecera.Filter = adFilterNone
                   End If
                   frmCaja.RsCabecera.Bookmark = nPos
                   Screen.MousePointer = vbDefault
                   MsgBox "Registro Modificado", vbInformation, sMensaje
               End If
               cmdOpcionGrilla(0).Enabled = True
               cmdOpcionGrilla(1).Enabled = True
               cmdOpcionGrilla(2).Enabled = True
               
'               On Error Resume Next
'               Dim CnSyBase As Connection
'               Dim sSyBASE As String

'               sSyBASE = Trim(LeerIni(App.Path + "\INFOREST.INI", "CONEXION", "SYBASE", ""))
'               Set CnSyBase = New Connection
'
'               CnSyBase.Provider = "ASAProv.80"
'               CnSyBase.CursorLocation = adUseServer
'               CnSyBase.ConnectionString = "DSN=" & sSyBASE
'               CnSyBase.CommandTimeout = 250
'               CnSyBase.Open
'               Dim X1 As String
'               Dim X2 As Date
'               Dim X3 As Date
'               X1 = Calcular("select top 1 autocombte_Fac as Codigo from tb_sri_tipo_ven_tram_cab order by id_tramite_auto desc", CnSyBase)
'               X2 = Calcular("select top 1 fecemicbte_Fac as Codigo from tb_sri_tipo_ven_tram_cab order by id_tramite_auto desc", CnSyBase)
'               X3 = Calcular("select top 1 fechacaducacmbte as Codigo from tb_sri_tipo_ven_tram_cab order by id_tramite_auto desc", CnSyBase)
'               Cn.Execute "update TCAJA set tNumeroAutorizacion='" & X1 & "', fInicio=" & X2 & ",fValido=" & X3
'               CnSyBase.Close
               
          Case Is = 2 ' Eliminar
               
               If frmCaja.RsCabecera.RecordCount = 0 Then
                  Exit Sub
               End If
            
               'Cambia el MsgBox
               If MsgBox("Seguro de Eliminar la Caja " & txtCodigo & "?", vbQuestion + vbOKCancel, sMensaje) = vbCancel Then
                  Exit Sub
               End If
               
               If txtCodigo.Text = sCaja Then
                  MsgBox "No se puede eliminar la Caja activa", vbCritical, sMensaje
                  Exit Sub
               End If
               
               sPasa = txtCodigo.Text
               
               'Inserta Movimiento auditoria
               lAuditoria = RegistraMovimientoAuditoria(tModuloSeg, sMDB, "TCAJA", "CAJA", "03", sUsuario, sPasa, "", _
                            "tcaja", "CodigoCaja", sPasa, "tdescripcion", "Descripcion Detallada", txtDetallado.Text)
               
                If lAuditoria = False Then
                        Screen.MousePointer = vbDefault
                        Exit Sub
                   End If
                                  
               
               'Cambia el Delete
               Cn.Execute "delete from TCAJA where tCaja = '" & txtCodigo & "'"
               Cn.Execute "delete from TTIPODOCUMENTOIMPRESORA where tCaja ='" & txtCodigo.Text & "'"
               frmCaja.RsCabecera.Requery
               
               If frmCaja.RsCabecera.RecordCount <> 0 Then
                  frmCaja.RsCabecera.MoveLast
                  Asignar
                  cmdTexto.Caption = "Registro " & IIf(frmCaja.RsCabecera.RecordCount = 0, 0, frmCaja.RsCabecera.AbsolutePosition) & " de " & frmCaja.RsCabecera.RecordCount
               Else
                  ActivarBotones False
                  Blanquear Me
                  Sw = True
               End If
          
          Case Is = 3 ' Salir
               Unload Me
   End Select
End Sub

Private Sub cmdOpcionGrilla_Click(Index As Integer)
   Select Case Index
          Case Is = 0 ' Agregar
               'Cambiar los Controles
                With RsGrilla
                     'Cuadro de Texto
                     cboTipoDocumento.Text = ""
                     cboImpresora.Text = ""
                     cboFormulario.Text = ""
                     txtSerie.Text = ""
                     txtSerie2.Text = ""
                     txtCorrelativo.Text = ""
                     txtCorrelativo2.Text = ""
                     txtDescripcion.Text = ""
                     txtAutorizacion.Text = ""
                     chLImprimeImageCab.value = 0
                     chLImprimeImagePie.value = 0
                     chkFacturacionE.value = 0
                     chkFacturacionOfisis.value = 0
                     txtPrefijoEnlace.Text = ""
                End With
                SubDetalle False
                wAgrega = True
                cboTipoDocumento.Enabled = True
          
          Case Is = 1 ' Modificar
               If RsGrilla.RecordCount = 0 Then
                  Exit Sub
               End If
               SubDetalle False
               wAgrega = False
               SubAsignar
               cboTipoDocumento.Enabled = False
          
          Case Is = 2 ' Eliminar
               If RsGrilla.RecordCount = 0 Then
                  Exit Sub
               End If
               
               'Cambia el MsgBox
               SubAsignar
               If MsgBox("Seguro de Eliminar este Tipo Documento " & cboTipoDocumento.Text & "?", vbQuestion + vbOKCancel, sMensaje) = vbCancel Then
                  Exit Sub
               End If
               
                   sPasa = txtCodigo.Text
                   
                   If pais = "002" Then 'Ecuador
                      'Inserta Movimiento auditoria
                      lAuditoria = RegistraMovimientoAuditoria(tModuloSeg, sMDB, "TTIPODOCUMENTOIMPRESORA", "TIPODOCUMENTOIMPRESORA", "03", sUsuario, sPasa, "", _
                      "tcaja", "CodigoCaja", sPasa, "tTipoEmision", "Tipo Emision", cboTipoDocumento.BoundText, "tImpresora", "Impresora", cboImpresora.BoundText, _
                      "tDescripcion", "Descripcion", txtDescripcion.Text, "tFormulario", "Formulario", cboFormulario.BoundText, _
                      "tSerie", "Serie", txtSerie2.Text, "tUltimoNumero", "UltimoNumero", txtCorrelativo2.Text, _
                      "tUsuario", "Usuario", sUsuario)
                   Else
                      'Inserta Movimiento auditoria
                      lAuditoria = RegistraMovimientoAuditoria(tModuloSeg, sMDB, "TTIPODOCUMENTOIMPRESORA", "TIPODOCUMENTOIMPRESORA", "03", sUsuario, sPasa, "", _
                      "tcaja", "CodigoCaja", sPasa, "tTipoEmision", "Tipo Emision", cboTipoDocumento.BoundText, "tImpresora", "Impresora", cboImpresora.BoundText, _
                      "tDescripcion", "Descripcion", txtDescripcion.Text, "tFormulario", "Formulario", cboFormulario.BoundText, _
                      "tSerie", "Serie", txtSerie.Text, "tUltimoNumero", "UltimoNumero", txtCorrelativo.Text, _
                      "tUsuario", "Usuario", sUsuario)
                   End If
          
                   If lAuditoria = False Then
                      Screen.MousePointer = vbDefault
                      Exit Sub
                   End If
                   
               'Cambia el Delete
               Cn.Execute "delete from TTIPODOCUMENTOIMPRESORA where tCaja ='" & txtCodigo.Text & "' and tTipoEmision ='" & cboTipoDocumento.BoundText & "'"
               RsGrilla.Requery
               If RsGrilla.RecordCount <> 0 Then
                  RsGrilla.MoveLast
               End If
          
          Case Is = 3 ' Grabar
          
               If pais = "000" Then
                     If Len(Trim(txtSerie.Text)) <> 5 Then
                          MsgBox "El número de serie debe ser de 5 caracteres", vbCritical, sMensaje
                          Exit Sub
                     End If
               End If
          
               If wAgrega Then
                   RsGrilla.Find ("tTipoEmision ='" & cboTipoDocumento.BoundText & "'")
                   If Not RsGrilla.EOF Then
                      MsgBox "Tipo de Documento " & cboTipoDocumento.Text & " ya ingresado", vbCritical, sMensaje
                      Exit Sub
                   End If
                          
                   sPasa = txtCodigo.Text
                   
                   If pais = "002" Then 'Ecuador
                      'Inserta Movimiento auditoria
                      lAuditoria = RegistraMovimientoAuditoria(tModuloSeg, sMDB, "TTIPODOCUMENTOIMPRESORA", "TIPODOCUMENTOIMPRESORA", "01", sUsuario, sPasa, "", _
                      "tcaja", "CodigoCaja", sPasa, "tTipoEmision", "Tipo Emision", cboTipoDocumento.BoundText, "tImpresora", "Impresora", cboImpresora.BoundText, _
                      "tDescripcion", "Descripcion", txtDescripcion.Text, "tFormulario", "Formulario", cboFormulario.BoundText, _
                      "tSerie", "Serie", cboLocal.Text & txtSerie2.Text, "tUltimoNumero", "UltimoNumero", txtCorrelativo2.Text, _
                      "tUsuario", "Usuario", sUsuario, "lResumen", "Flag Resumen", IIf(chkResumen.value, "Verdadero", "Falso"), "limpuesto1", "Impuesto 1", IIf(chkImpuesto1.value, "Verdadero", "Falso"), "limpuesto2", "Impuesto 2", IIf(chkImpuesto2.value, "Verdadero", "Falso"), "limpuesto3", "Impuesto 3", IIf(chkImpuesto3.value, "Verdadero", "Falso"), _
                      "lEquivaDolares", "Equivale Dolares", IIf(chkDocEquivDolares.value, "Verdadero", "Falso"), _
                      "lImprimeImageCab", "Imagen Cabecera", IIf(chLImprimeImageCab.value, "Verdadero", "Falso"), _
                      "lImprimeImagePie", "Imagen Pie", IIf(chLImprimeImagePie.value, "Verdadero", "Falso"), "lImpProdDesc", "Impresion de CodPlatos y descuento unitario", IIf(Me.chkCodProdDes.value, "Verdadero", "Falso"), "lImpDocMayorCero", "Impresion de Platos Mayor a Cero", IIf(Me.chkMayorCero.value, "Verdadero", "Falso"))
                      
                   Else
                      'Inserta Movimiento auditoria
                      lAuditoria = RegistraMovimientoAuditoria(tModuloSeg, sMDB, "TTIPODOCUMENTOIMPRESORA", "TIPODOCUMENTOIMPRESORA", "01", sUsuario, sPasa, "", _
                      "tcaja", "CodigoCaja", sPasa, "tTipoEmision", "Tipo Emision", cboTipoDocumento.BoundText, "tImpresora", "Impresora", cboImpresora.BoundText, _
                      "tDescripcion", "Descripcion", txtDescripcion.Text, "tFormulario", "Formulario", cboFormulario.BoundText, _
                      "tSerie", "Serie", txtSerie.Text, "tUltimoNumero", "UltimoNumero", txtCorrelativo.Text, _
                      "tUsuario", "Usuario", sUsuario, "lResumen", "Flag Resumen", IIf(chkResumen.value, "Verdadero", "Falso"), "limpuesto1", "Impuesto 1", IIf(chkImpuesto1.value, "Verdadero", "Falso"), "limpuesto2", "Impuesto 2", IIf(chkImpuesto2.value, "Verdadero", "Falso"), "limpuesto3", "Impuesto 3", IIf(chkImpuesto3.value, "Verdadero", "Falso"), _
                      "lEquivaDolares", "Equivale Dolares", IIf(chkDocEquivDolares.value, "Verdadero", "Falso"), _
                      "lImprimeImageCab", "Imagen Cabecera", IIf(chLImprimeImageCab.value, "Verdadero", "Falso"), _
                      "lImprimeImagePie", "Imagen Pie", IIf(chLImprimeImagePie.value, "Verdadero", "Falso"), "lImpProdDesc", "Impresion de CodPlatos y descuento unitario", IIf(Me.chkCodProdDes.value, "Verdadero", "Falso"), "lImpDocMayorCero", "Impresion de Platos Mayor a Cero", IIf(Me.chkMayorCero.value, "Verdadero", "Falso"))
                   End If
                   
                   If lAuditoria = False Then
                        Screen.MousePointer = vbDefault
                        Exit Sub
                   End If
                   If pais = "002" Then 'Ecuador
                            Isql = "insert into TTIPODOCUMENTOIMPRESORA( " & _
                                   "tCaja, tTipoEmision, tImpresora, tDescripcion, tFormulario, tSerie, tUltimoNumero, tUsuario, lResumen, lImpuesto1, lImpuesto2, lImpuesto3, fRegistro,lEquivaDolares, lImprimeImageCab,lImprimeImagePie, lFacturacionElectronica,tNumeroAutorizacion,fInicio,fCaducidad,tPrefijoEnlace,lDocumentoElectronicoOfisis,limprimeresumen,lOpGravInaf, lImpProdDesc, limpDocMayorCero) " & _
                                   "values (  '" & txtCodigo.Text & "', " & _
                                            " '" & cboTipoDocumento.BoundText & "', " & _
                                            " '" & cboImpresora.BoundText & "', " & _
                                            " '" & txtDescripcion.Text & "', " & _
                                            " '" & cboFormulario.BoundText & "', " & _
                                            " '" & cboLocal.Text & txtSerie2.Text & "', " & _
                                            " '" & txtCorrelativo2.Text & "', " & _
                                            " '" & sUsuario & "', " & _
                                                   chkResumen.value & ", " & _
                                                   chkImpuesto1.value & ", " & _
                                                   chkImpuesto2.value & ", " & _
                                                   chkImpuesto3.value & ", " & _
                                            " getdate()," & chkDocEquivDolares.value & ", " & chLImprimeImageCab.value & ", " & chLImprimeImagePie.value & ", " & chkFacturacionE.value & ",'" & txtAutorizacion.Text & "','" & Format(dtpFechaInicio.value, "yyyy/MM/dd") & "','" & Format(dtpFechaCaducida.value, "yyyy/MM/dd") & "','" & txtPrefijoEnlace.Text & "'," & chkFacturacionOfisis.value & " ," & chkImpResumido.value & "," & chkopGravInaf.value & "," & Me.chkCodProdDes.value & "," & Me.chkMayorCero.value & " )"
                       Else
                            Isql = "insert into TTIPODOCUMENTOIMPRESORA( " & _
                            "tCaja, tTipoEmision, tImpresora, tDescripcion, tFormulario, tSerie, tUltimoNumero, tUsuario, lResumen, lImpuesto1, lImpuesto2, lImpuesto3, fRegistro,lEquivaDolares,  lImprimeImageCab,lImprimeImagePie, lFacturacionElectronica,tPrefijoEnlace,lDocumentoElectronicoOfisis,tFormVenta,tCompVenta,limprimeresumen,lOpGravInaf, lImpProdDesc, limpDocMayorCero) " & _
                            "values (  '" & txtCodigo.Text & "', " & _
                                     " '" & cboTipoDocumento.BoundText & "', " & _
                                     " '" & cboImpresora.BoundText & "', " & _
                                     " '" & txtDescripcion.Text & "', " & _
                                     " '" & cboFormulario.BoundText & "', " & _
                                     " '" & txtSerie.Text & "', " & _
                                     " '" & txtCorrelativo.Text & "', " & _
                                     " '" & sUsuario & "', " & _
                                            chkResumen.value & ", " & _
                                            chkImpuesto1.value & ", " & _
                                            chkImpuesto2.value & ", " & _
                                            chkImpuesto3.value & ", " & _
                                     " getdate()," & chkDocEquivDolares.value & ", " & chLImprimeImageCab.value & ", " & chLImprimeImagePie.value & ", " & chkFacturacionE.value & ",'" & txtPrefijoEnlace.Text & "'," & chkFacturacionOfisis.value & ",'" & txtFormVenta.Text & "','" & txtCompVenta.Text & "'," & chkImpResumido.value & "," & chkopGravInaf.value & "," & Me.chkCodProdDes.value & ", " & Me.chkMayorCero.value & " )"
                       End If
                       Cn.Execute Isql
                       RsGrilla.Filter = "tCaja ='" & txtCodigo.Text & "'"
                       RsGrilla.Requery
                       RsGrilla.MoveLast
                       MsgBox "Registro Agregado", vbInformation, sMensaje
                       
               Else
                   sPasa = txtCodigo.Text
                                      
                   'Inserta Movimiento auditoria
                   If pais = "002" Then 'Ecuador
                        lAuditoria = RegistraMovimientoAuditoria(tModuloSeg, sMDB, "TTIPODOCUMENTOIMPRESORA", "TIPODOCUMENTOIMPRESORA", "02", sUsuario, sPasa, "", _
                        "tcaja", "Codigo de Caja", sPasa, "tTipoEmision", "Tipo Emision", cboTipoDocumento.BoundText, "tImpresora", "Impresora", cboImpresora.BoundText, _
                        "tDescripcion", "Descripcion", txtDescripcion.Text, "tFormulario", "Formulario", cboFormulario.BoundText, _
                        "tSerie", "Serie", cboLocal.Text & txtSerie2.Text, "tUltimoNumero", "UltimoNumero", txtCorrelativo2.Text, _
                        "tUsuario", "Usuario", sUsuario, "lResumen", "Flag Resumen", IIf(chkResumen.value, "Verdadero", "Falso"), "limpuesto1", "Impuesto 1", IIf(chkImpuesto1.value, "Verdadero", "Falso"), "limpuesto2", "Impuesto 2", IIf(chkImpuesto2.value, "Verdadero", "Falso"), "limpuesto3", "Impuesto 3", IIf(chkImpuesto3.value, "Verdadero", "Falso"), _
                        "lEquivaDolares", "Equivale Dolares", IIf(chkDocEquivDolares.value, "Verdadero", "Falso"), _
                        "lImprimeImageCab", "Imagen Cabecera", IIf(chLImprimeImageCab.value, "Verdadero", "Falso"), _
                        "lImprimeImagePie", "Imagen Pie", IIf(chLImprimeImagePie.value, "Verdadero", "Falso"), "lImpProdDesc", "Impresion de CodPlatos y descuento unitario", IIf(Me.chkCodProdDes.value, "Verdadero", "Falso"))
                  Else
                        lAuditoria = RegistraMovimientoAuditoria(tModuloSeg, sMDB, "TTIPODOCUMENTOIMPRESORA", "TIPODOCUMENTOIMPRESORA", "02", sUsuario, sPasa, "", _
                        "tcaja", "Codigo de Caja", sPasa, "tTipoEmision", "Tipo Emision", cboTipoDocumento.BoundText, "tImpresora", "Impresora", cboImpresora.BoundText, _
                        "tDescripcion", "Descripcion", txtDescripcion.Text, "tFormulario", "Formulario", cboFormulario.BoundText, _
                        "tSerie", "Serie", txtSerie.Text, "tUltimoNumero", "UltimoNumero", txtCorrelativo.Text, _
                        "tUsuario", "Usuario", sUsuario, "lResumen", "Flag Resumen", IIf(chkResumen.value, "Verdadero", "Falso"), "limpuesto1", "Impuesto 1", IIf(chkImpuesto1.value, "Verdadero", "Falso"), "limpuesto2", "Impuesto 2", IIf(chkImpuesto2.value, "Verdadero", "Falso"), "limpuesto3", "Impuesto 3", IIf(chkImpuesto3.value, "Verdadero", "Falso"), _
                        "lEquivaDolares", "Equivale Dolares", IIf(chkDocEquivDolares.value, "Verdadero", "Falso"), _
                        "lImprimeImageCab", "Imagen Cabecera", IIf(chLImprimeImageCab.value, "Verdadero", "Falso"), _
                        "lImprimeImagePie", "Imagen Pie", IIf(chLImprimeImagePie.value, "Verdadero", "Falso"), "lImpProdDesc", "Impresion de CodPlatos y descuento unitario", IIf(Me.chkCodProdDes.value, "Verdadero", "Falso"))
                  End If
                  If lAuditoria = False Then
                        Screen.MousePointer = vbDefault
                        Exit Sub
                   End If
               
                  'Cambiar el SQL
                  If pais = "002" Then 'Ecuador
                     Isql = "update TTIPODOCUMENTOIMPRESORA set " & _
                            "tImpresora ='" & cboImpresora.BoundText & "', " & _
                            "tDescripcion ='" & txtDescripcion.Text & "', " & _
                            "tFormulario ='" & cboFormulario.BoundText & "', " & _
                            "tSerie ='" & cboLocal.Text & txtSerie2.Text & "', " & _
                            "tUltimoNumero ='" & txtCorrelativo2.Text & "', " & _
                            "lResumen =" & chkResumen.value & ", " & _
                            "lImpuesto1 =" & chkImpuesto1.value & ", " & _
                            "lImpuesto2 =" & chkImpuesto2.value & ", " & _
                            "lImpuesto3 =" & chkImpuesto3.value & ", " & _
                            "lImprimeImageCab=" & chLImprimeImageCab.value & ",  " & _
                            "lImprimeImagepie=" & chLImprimeImagePie.value & ",  " & _
                            "lFacturacionElectronica=" & chkFacturacionE.value & ",  " & _
                            "limprimeresumen=" & chkImpResumido.value & ",  " & _
                            "lOpGravInaf=" & chkopGravInaf.value & ",  " & _
                            "lDocumentoElectronicoOfisis =" & Me.chkFacturacionOfisis.value & ", " & _
                            "lEquivaDolares =" & chkDocEquivDolares.value & ", " & _
                            "fInicio = '" & Format(dtpFechaInicio.value, "yyyy/MM/dd") & "', " & _
                            "tPrefijoEnlace ='" & txtPrefijoEnlace.Text & "', " & _
                            "fCaducidad = '" & Format(dtpFechaCaducida.value, "yyyy/MM/dd") & "', " & _
                            "tNumeroAutorizacion ='" & txtAutorizacion.Text & "', " & _
                            "lImpProdDesc =" & Me.chkCodProdDes.value & ", " & _
                            "lImpDocMayorCero =" & Me.chkMayorCero.value & " " & _
                            " where tCaja = '" & txtCodigo.Text & "' and tTipoEmision = '" & cboTipoDocumento.BoundText & "'"
                   Else 'lImpProdDesc
                     Isql = "update TTIPODOCUMENTOIMPRESORA set " & _
                            "tImpresora ='" & cboImpresora.BoundText & "', " & _
                            "tDescripcion ='" & txtDescripcion.Text & "', " & _
                            "tFormulario ='" & cboFormulario.BoundText & "', " & _
                            "tSerie ='" & txtSerie.Text & "', " & _
                            "tUltimoNumero ='" & txtCorrelativo.Text & "', " & _
                            "lResumen =" & chkResumen.value & ", " & _
                            "lImpuesto1 =" & chkImpuesto1.value & ", " & _
                            "lImpuesto2 =" & chkImpuesto2.value & ", " & _
                            "lImprimeImageCab=" & chLImprimeImageCab.value & ",  " & _
                            "lImprimeImagepie=" & chLImprimeImagePie.value & ",  " & _
                            "lImpuesto3 =" & chkImpuesto3.value & ", " & _
                            "tPrefijoEnlace ='" & txtPrefijoEnlace.Text & "', " & _
                            "tFormVenta ='" & txtFormVenta.Text & "', " & _
                            "tCompVenta ='" & txtCompVenta.Text & "', " & _
                            "lFacturacionElectronica=" & chkFacturacionE.value & ",  " & _
                            "limprimeresumen=" & chkImpResumido.value & ",  " & _
                            "lOpGravInaf=" & chkopGravInaf.value & ",  " & _
                            "lDocumentoElectronicoOfisis =" & Me.chkFacturacionOfisis.value & ", " & _
                            "lEquivaDolares =" & chkDocEquivDolares.value & ", " & _
                            "lImpProdDesc =" & Me.chkCodProdDes.value & ", " & _
                            "lImpDocMayorCero =" & Me.chkMayorCero.value & " " & _
                            " where tCaja = '" & txtCodigo.Text & "' and tTipoEmision = '" & cboTipoDocumento.BoundText & "'"
                   End If
                   Cn.Execute Isql
                   nPos = RsGrilla.AbsolutePosition 'chkImpResumido
                   RsGrilla.Requery
                   RsGrilla.AbsolutePosition = nPos
                   MsgBox "Registro Modificado", vbInformation, sMensaje
               End If
               SubDetalle True
          
          Case Is = 4 ' Cancelar
               SubDetalle True
               
          Case Is = 5 ' Agregar Area
               'Cambiar los Controles
                With RsAI
                     'Cuadro de Texto
                     cboArea.Text = ""
                     cboImpArea.Text = ""
                     txtDescripcion.Text = ""
                End With
                SubDetArea False
                wAgrega = True
          
          Case Is = 6 ' Modificar Area
               If RsAI.RecordCount = 0 Then
                  Exit Sub
               End If
               SubDetArea False
               wAgrega = False
               SubArea
          
          Case Is = 7 ' Eliminar Area
               If RsAI.RecordCount = 0 Then
                  Exit Sub
               End If
               
               'Cambia el MsgBox
               SubArea
               If MsgBox("Seguro de Eliminar esta Area " & cboArea.Text & "?", vbQuestion + vbOKCancel, sMensaje) = vbCancel Then
                  Exit Sub
               End If
                              
               'Cambia el Delete
               Cn.Execute "delete from TAREAIMPRESORA where tCaja ='" & txtCodigo.Text & "' and tArea ='" & cboArea.BoundText & "'"
               RsAI.Requery
               If RsAI.RecordCount <> 0 Then
                  RsAI.MoveLast
               End If
          
          Case Is = 8 ' Grabar ARea
               If wAgrega Then
                  If RsAI.RecordCount > 0 Then
                     RsAI.MoveFirst
                     RsAI.Find ("tArea ='" & cboArea.BoundText & "'")
                     If Not RsAI.EOF Then
                        MsgBox "Area " & cboArea.Text & " ya ingresada", vbCritical, sMensaje
                        Exit Sub
                     End If
                  End If
                  
                   Isql = "insert into TAREAIMPRESORA( " & _
                          "tCaja, tArea, tImpresora, tUsuario, fRegistro) " & _
                          "values (  '" & txtCodigo.Text & "', " & _
                                   " '" & cboArea.BoundText & "', " & _
                                   " '" & cboImpArea.BoundText & "', " & _
                                   " '" & sUsuario & "', " & _
                                   " getdate() )"
            
                       Cn.Execute Isql
          
                       RsAI.Filter = "tCaja ='" & txtCodigo.Text & "'"
                       RsAI.Requery
                       RsAI.MoveLast
                       MsgBox "Registro Agregado", vbInformation, sMensaje
               Else
                              
                  'Cambiar el SQL
                  Isql = "update TAREAIMPRESORA set " & _
                         "tImpresora ='" & cboImpArea.BoundText & "' " & _
                         " where tCaja = '" & txtCodigo.Text & "' and tArea = '" & cboArea.BoundText & "'"
                       
                   Cn.Execute Isql
                   nPos = RsArea.AbsolutePosition
                   RsAI.Requery
                   RsAI.AbsolutePosition = nPos
                   MsgBox "Registro Modificado", vbInformation, sMensaje
               End If
               SubDetArea True
          
          Case Is = 9 ' Cancelar
               SubDetArea True
               
          'CESAR CHEF CONTROL
          'AREA CHEF
          Case Is = 10 ' Grabar AreaChef
               If wAgrega Then
                  If RsAChef.RecordCount > 0 Then
                     RsAChef.MoveFirst
                     RsAChef.Find ("Area ='" & cboAreaChef.BoundText & "'")
                     If Not RsAChef.EOF Then
                        MsgBox "Area " & cboAreaChef.Text & " ya ingresada", vbCritical, sMensaje
                        Exit Sub
                     End If
                  End If
                  
                   Isql = "insert into TAREACHEF( " & _
                          "tCaja, tArea, lArea, tUsuario, fRegistro) " & _
                          "values (  '" & txtCodigo.Text & "', " & _
                                   " '" & cboAreaChef.BoundText & "', " & _
                                   " " & chkAreaChef.value & " , " & _
                                   " '" & sUsuario & "', " & _
                                   " getdate() )"
            
                       Cn.Execute Isql
          
                       RsAChef.Filter = "tCaja ='" & txtCodigo.Text & "'"
                       RsAChef.Requery
                       RsAChef.MoveLast
                       MsgBox "Registro Agregado", vbInformation, sMensaje
               Else
                              
                  'Cambiar el SQL
                  Isql = "update TAREACHEF set " & _
                         "lArea = " & chkAreaChef.value & " " & _
                         " where tCaja = '" & txtCodigo.Text & "' and tArea = '" & cboAreaChef.BoundText & "'"
                       
                   Cn.Execute Isql
                   nPos = RsArea.AbsolutePosition
                   RsAChef.Requery
                   RsAChef.AbsolutePosition = nPos
                   MsgBox "Registro Modificado", vbInformation, sMensaje
               End If
               SubDetAreaChef True
               
          Case Is = 11 'Cancelar Area Chef
               SubDetAreaChef True
               
          Case Is = 13 'Agregar Area Chef
                With RsAChef
                     cboAreaChef.Text = ""
                     chkAreaChef.value = 0
                End With
                SubDetAreaChef False
                wAgrega = True
               
          Case Is = 14 ' Modificar Area Chef
               If RsAChef.RecordCount = 0 Then
                  Exit Sub
               End If
               SubDetAreaChef False
               wAgrega = False
               SubAreaChef
               
          Case Is = 12 ' Eliminar Area Chef
               If RsAChef.RecordCount = 0 Then
                  Exit Sub
               End If
               
               SubAreaChef
               If MsgBox("Seguro de Eliminar esta Area " & cboAreaChef.Text & "?", vbQuestion + vbOKCancel, sMensaje) = vbCancel Then
                  Exit Sub
               End If
                              
               'Cambia el Delete
               Cn.Execute "delete from TAREACHEF where tCaja ='" & txtCodigo.Text & "' and tArea ='" & cboAreaChef.BoundText & "'"
               RsAChef.Requery
               If RsAChef.RecordCount <> 0 Then
                  RsAChef.MoveLast
               End If
          '---------------------------------------
               
          Case 15 ' graba sub grupos
                If cboSubGrupo.Text = "" Then
                        MsgBox "Seleccionar un Sub Grupo", vbInformation, sMensaje
                        Exit Sub
                End If
                If cboAreaProd.Text = "" Then
                        MsgBox "Seleccionar un Area", vbInformation, sMensaje
                        Exit Sub
                End If
                 If wAgrega Then
                  If rsAreaSubGrupo.RecordCount > 0 Then
                     rsAreaSubGrupo.MoveFirst
                     rsAreaSubGrupo.Find ("tSubGrupo ='" & cboSubGrupo.BoundText & "'")
                     If Not rsAreaSubGrupo.EOF Then
                        MsgBox "Sub Grupo " & cboSubGrupo.Text & " ya ingresado", vbCritical, sMensaje
                        Exit Sub
                     End If
                  End If
                  
                   Isql = "insert into TAREASUBGRUPO( " & _
                          "TCAJA, TSUBGRUPO, TAREA,TUSUARIO,FREGISTRO) " & _
                          "values (  '" & txtCodigo.Text & "', " & _
                                   " '" & cboSubGrupo.BoundText & "', " & _
                                   " '" & cboAreaProd.BoundText & "', " & _
                                   " '" & sUsuario & "', " & _
                                   " getdate() )"
            
                       Cn.Execute Isql
          
                       rsAreaSubGrupo.Filter = "tCaja ='" & txtCodigo.Text & "'"
                       rsAreaSubGrupo.Requery
                       rsAreaSubGrupo.MoveLast
                       MsgBox "Registro Agregado", vbInformation, sMensaje
               Else
                              
                  'Cambiar el SQL
                  Isql = "update TAREASUBGRUPO set " & _
                         "tarea ='" & cboAreaProd.BoundText & "' " & _
                         " where tCaja = '" & txtCodigo.Text & "' and tSubGrupo = '" & cboSubGrupo.BoundText & "'"
                       
                   Cn.Execute Isql
                   nPos = rsAreaSubGrupo.AbsolutePosition
                   rsAreaSubGrupo.Requery
                   rsAreaSubGrupo.AbsolutePosition = nPos
                   MsgBox "Registro Modificado", vbInformation, sMensaje
               End If
               SubDetAreaSubGrupo True
            
           Case 16
               SubDetAreaSubGrupo True
        
         Case 17 ' lg eliminar:
          If rsAreaSubGrupo.RecordCount = 0 Then
                  Exit Sub
               End If
               
               'Cambia el MsgBox
               SubAreaSubGrupo
               If MsgBox("Seguro de Eliminar este Sub Grupo " & cboSubGrupo.Text & "?", vbQuestion + vbOKCancel, sMensaje) = vbCancel Then
                  Exit Sub
               End If
                              
               'Cambia el Delete
               Cn.Execute "delete from tareasubgrupo where tCaja ='" & txtCodigo.Text & "' and tsubgrupo ='" & cboSubGrupo.BoundText & "'"
               rsAreaSubGrupo.Requery
               If rsAreaSubGrupo.RecordCount <> 0 Then
                  rsAreaSubGrupo.MoveLast
               End If
          Case 18 ' lg nuevp
              With rsAreaSubGrupo
                     cboSubGrupo.Text = ""
                     cboAreaProd.Text = ""
              End With
                SubDetAreaSubGrupo False
                wAgrega = True
        Case 19 ' lg modificacion area subgrupo
                If rsAreaSubGrupo.RecordCount = 0 Then
                  Exit Sub
               End If
               SubDetAreaSubGrupo False
               wAgrega = False
               SubAreaSubGrupo
   End Select
   
End Sub


Sub SubAreaSubGrupo()
    With rsAreaSubGrupo
         cboSubGrupo.BoundText = IIf(IsNull(!tSubGrupo), "", !tSubGrupo)
         cboAreaProd.BoundText = IIf(IsNull(!tArea), "", !tArea)
    End With
End Sub

Public Sub SubDetAreaSubGrupo(Activa As Boolean)
   fraAreaProduccion.Visible = Not Activa
   ActivarBotones Activa
   cmdOpcion(1).Enabled = Activa
   cmdOpcion(3).Enabled = Activa
   'A-M-E
   cmdOpcionGrilla(18).Enabled = Activa
   cmdOpcionGrilla(19).Enabled = Activa
   cmdOpcionGrilla(17).Enabled = Activa
 
End Sub

Private Sub cmdQuitarFotoCabecera_Click()
    imgFoto.Picture = Nothing
    Cn.Execute "update tcaja set iimagencabdoc=null where tcaja='" & txtCodigo.Text & "'"
End Sub

Private Sub cmdQuitarFotoPie_Click()
    imgFotoPie.Picture = Nothing
    Cn.Execute "update tcaja set iimagenpiedoc=null where tcaja='" & txtCodigo.Text & "'"
End Sub

Private Sub Form_Load()
    Screen.MousePointer = vbHourglass
    Centrar Me
    
    'Ingrese el Titulo
    Me.Caption = " Mantenimiento de Cajas "
    fraDetalle.Caption = Me.Caption
    
    'Ingrese el SubTitulo
    grdGrilla.Caption = " Configuración de Documentos "
    fraGrilla.Visible = False
    
    grdAI.Caption = " Configuración de Areas "
    fraArea.Visible = False
    fraArea.Caption = grdAI.Caption
    
     If chkMesa247.value = 1 Then
        Label(25).Visible = True
        cboComprobante.Visible = True
    Else
        Label(25).Visible = False
        cboComprobante.Visible = False
    End If
       
    'lucho areas por subgrupos
    Me.grdGrillaSubgrupos.Caption = " Configuración de Areas Por Sub Grupos "
    Me.fraAreaProduccion.Visible = False
    Me.fraAreaProduccion.Caption = grdGrillaSubgrupos.Caption
    
    
    'CESAR AREA CHEF
    grdAChef.Caption = " Configuración de Areas Chef Control "
    fraAreaChef.Visible = False
    fraAreaChef.Caption = grdAChef.Caption
    
    
    
    
    
    
    'Llena todos los Combos
    LlenaCombos
    
    'Ingresar la Vista de la Grilla
    Isql = "select * from vTipoDocumentoImpresora"
    Set RsGrilla = Lib.OpenRecordset(Isql, Cn)
    
    Call ConfGrilla(8, grdGrilla, "Descripción", 2, "Descripcion", 1900, 0, 0, "", _
                                  "Prefijo", 2, "Prefijo", 600, 2, 0, "", _
                                  "Serie", 2, "tSerie", 800, 2, 0, "", _
                                  "Número", 2, "tUltimoNumero", 1200, 0, 0, "", _
                                  "Impresora", 2, "Impresora", 1500, 0, 0, "", _
                                  "Autorizacion", 2, "tNumeroAutorizacion", 1400, 0, 0, "", _
                                  "Imag Cab", 2, "lImprimeImageCab", 920, 2, 4, "", _
                                  "Imag Pie", 2, "lImprimeImagePie", 920, 2, 4, "")
    Set grdGrilla.DataSource = RsGrilla
    
    'Ingresar la Vista de la Grilla 2
    Isql = "select * from vAreaImpresora"
    Set RsAI = Lib.OpenRecordset(Isql, Cn)
    
    Call ConfGrilla(2, grdAI, "Area", 2, "Area", 3000, 0, 0, "", _
                              "Impresora", 2, "Impresora", 3000, 0, 0, "")
    Set grdAI.DataSource = RsAI
    
    
    'lucho area subgrupo
        
    Isql = "Select * From vAreaSubGrupo"
    Set rsAreaSubGrupo = Lib.OpenRecordset(Isql, Cn)
    
    Call ConfGrilla(3, grdGrillaSubgrupos, "Caja", 2, "tCaja", 0, 0, 0, "", _
                                 "SubGrupo", 2, "SubGrupo", 3400, 0, 0, "", _
                                 "Area", 2, "Area", 3400, 0, 0, "")
                                 
    Set Me.grdGrillaSubgrupos.DataSource = rsAreaSubGrupo
    '------------------------------------
       
    'CESAR AREA CHEF
    Isql = "Select * From vAreaChef"
    Set RsAChef = Lib.OpenRecordset(Isql, Cn)
    
    Call ConfGrilla(3, grdAChef, "Caja", 2, "Caja", 2000, 0, 0, "", _
                                 "Area", 2, "Area", 3000, 0, 0, "", _
                                 "AreaChef", 2, "AreaChef", 1000, 2, 4, "")
    Set grdAChef.DataSource = RsAChef
    '------------------------------------
    
    
    If Sw = True Then
       ActivarBotones (False)
       Blanquear Me
       inicio
    Else
       'Cambiar la Busqueda y Nombre del formulario Cabecera
       ActivarBotones (True)
       Asignar
    End If
    
    If pais = "002" Then 'Ecuador
       txtSerie.Visible = False
       txtCorrelativo.Visible = False
       cboLocal.Visible = True
       txtSerie2.Visible = True
       txtCorrelativo2.Visible = True
              
       If lFacturacionE Then
            txtAutorizacion.Enabled = False
            dtpFechaInicio.Enabled = False
            dtpFechaCaducida.Enabled = False
       Else
            txtAutorizacion.Enabled = True
            dtpFechaInicio.Enabled = True
            dtpFechaCaducida.Enabled = True
       End If
   
    Else
       txtSerie.Visible = True
       txtCorrelativo.Visible = True
       cboLocal.Visible = True
       txtSerie2.Visible = False
       txtCorrelativo2.Visible = False
       cboLocal.Visible = False
       Label(19).Visible = False
       Label(26).Visible = False
       Label(27).Visible = False
       
       txtAutorizacion.Visible = False
       dtpFechaInicio.Visible = False
       dtpFechaCaducida.Visible = False
       
    End If
    Me.tabOpcion.Tab = 0
    cmdTexto.Caption = "Registro " & IIf(frmCaja.RsCabecera.RecordCount = 0, 0, frmCaja.RsCabecera.AbsolutePosition) & " de " & frmCaja.RsCabecera.RecordCount
    Screen.MousePointer = vbDefault
End Sub

Private Sub Form_Unload(Cancel As Integer)
   Set RsGrupo = Nothing
   Set RsImpresora = Nothing
   Set RsPreCuenta = Nothing
   Set RsTipoDocumento = Nothing
   Set RsArea = Nothing
   Set RsAI = Nothing
   Set RsImpArea = Nothing
   Set RsFormulario = Nothing
   Set RsGrilla = Nothing
   Set RsTipoPedido = Nothing
   Set frmCajaDetalle = Nothing
End Sub

Sub ActivarBotones(ByVal Activa As Boolean)
    cmdNavegar(0).Enabled = Activa
    cmdNavegar(1).Enabled = Activa
    cmdNavegar(2).Enabled = Activa
    cmdNavegar(3).Enabled = Activa
    cmdNavegar(4).Enabled = Activa
    cmdNavegar(5).Enabled = Activa
    cmdOpcion(0).Enabled = Activa
    cmdOpcion(2).Enabled = Activa
End Sub

Private Sub grdGrilla_DblClick()
   cmdOpcionGrilla_Click (1)
End Sub

Public Sub SubDetalle(Activa As Boolean)
   With cboImpresora
        Isql = "Select * from TIMPRESORA where tCaja = '" & txtCodigo.Text & "' order by tImpresora"
        Set RsImpresora = Lib.OpenRecordset(Isql, Cn)
        Set .RowSource = RsImpresora
             .DataField = "tDescripcion"
             .ListField = "tDescripcion"
             .BoundColumn = "tImpresora"
   End With

   fraGrilla.Visible = Not Activa
   ActivarBotones Activa
   cmdOpcion(1).Enabled = Activa
   cmdOpcion(3).Enabled = Activa
   cmdOpcionGrilla(0).Enabled = Activa
   cmdOpcionGrilla(1).Enabled = Activa
   cmdOpcionGrilla(2).Enabled = Activa

   'Controles
   txtDetallado.Enabled = Activa
   chkActivo.Enabled = Activa
End Sub
















Private Sub txtBalanzaPuerto_Change()
    If Not IsNumeric(Me.txtBalanzaPuerto.Text) And Trim(Me.txtBalanzaPuerto.Text) <> "" Then
        MsgBox "Valor no valido!!!", vbInformation, sMensaje
        Me.txtBalanzaPuerto.Text = ""
        Me.txtBalanzaPuerto.SetFocus
    End If
End Sub

Private Sub txtbaltiempo_Change()
    If Not IsNumeric(Me.txtbaltiempo.Text) And Trim(Me.txtbaltiempo.Text) <> "" Then
        MsgBox "Valor no valido!!!", vbInformation, sMensaje
        Me.txtbaltiempo.Text = ""
        Me.txtbaltiempo.SetFocus
    End If
End Sub

Private Sub txtbaltiempo_LostFocus()
'    If Not IsNumeric(Me.txtbaltiempo.Text) And Trim(Me.txtbaltiempo.Text) <> "" Then
'        MsgBox "Valor no valido!!!", vbInformation, sMensaje
'    End If
End Sub

Private Sub txtCorrelativo_Lostfocus()
   txtCorrelativo.Text = Mid("000000000", 1, 9 - Len(Trim(str(Val(txtCorrelativo.Text))))) + Trim(str(Val(txtCorrelativo.Text)))
End Sub

Private Sub txtSerie_LostFocus()
   'txtSerie.Text = Mid("00000", 1, 5 - Len(Trim(str(Val(txtSerie.Text))))) + Trim(str(Val(txtSerie.Text)))
   txtSerie.Text = UCase(txtSerie.Text)
End Sub

Private Sub txtCorrelativo2_Lostfocus()
   txtCorrelativo2.Text = Mid("00000000", 1, 9 - Len(Trim(str(Val(txtCorrelativo2.Text))))) + Trim(str(Val(txtCorrelativo2.Text)))
End Sub

Private Sub txtSerie2_LostFocus()
   txtSerie2.Text = Mid("000", 1, 3 - Len(Trim(str(Val(txtSerie2.Text))))) + Trim(str(Val(txtSerie2.Text)))
End Sub

Sub SubAsignar()
    With RsGrilla
         'Cuadro de Texto
         cboTipoDocumento.BoundText = IIf(IsNull(!TTipoEmision), "", !TTipoEmision)
         cboImpresora.BoundText = IIf(IsNull(!timpresora), "", Trim(!timpresora))
         cboFormulario.BoundText = IIf(IsNull(!tFormulario), "", Trim(!tFormulario))
         txtDescripcion.Text = IIf(IsNull(!tDescripcion), "", !tDescripcion)
         chkResumen.value = IIf(IsNull(!lResumen), 0, IIf(!lResumen, 1, 0))
         chLImprimeImageCab.value = IIf(IsNull(!lImprimeImageCab), 0, IIf(!lImprimeImageCab, 1, 0))
         chLImprimeImagePie.value = IIf(IsNull(!lImprimeImagepie), 0, IIf(!lImprimeImagepie, 1, 0))
         chkImpResumido.value = IIf(IsNull(!lImprimeResumen), 0, IIf(!lImprimeResumen, 1, 0))
         'FACTURACION ELECTRONICA
         chkFacturacionE.value = IIf(IsNull(!lFacturacionElectronica), 0, IIf(!lFacturacionElectronica, 1, 0))
         
         chkFacturacionOfisis.value = IIf(IsNull(!lDocumentoElectronicoOfisis), 0, IIf(!lDocumentoElectronicoOfisis, 1, 0))
         chkopGravInaf.value = IIf(IsNull(!lOpGravInaf), 0, IIf(!lOpGravInaf, 1, 0))
         Me.chkCodProdDes.value = IIf(IsNull(!lImpProdDesc), 0, IIf(!lImpProdDesc, 1, 0))  'IIf(IsNull(!lImpProdDesc), 0, IIf(!lImpProdDesc, 1, 0))
         Me.chkMayorCero.value = IIf(IsNull(!lImpDocMayorCero), 0, IIf(!lImpDocMayorCero, 1, 0))  'IIf(IsNull(!lImpProdDesc), 0, IIf(!lImpProdDesc, 1, 0))
         If pais = "002" Then 'Ecuador
            cboLocal.Text = Mid(IIf(IsNull(!tSerie), "001", !tSerie), 1, 3)
            txtSerie2.Text = Mid(IIf(IsNull(!tSerie), "001", !tSerie), 4, 3)
            txtCorrelativo2.Text = IIf(IsNull(!tUltimoNumero), "", !tUltimoNumero)
            txtAutorizacion.Text = IIf(IsNull(!tNumeroAutorizacion), "", !tNumeroAutorizacion)
            dtpFechaInicio.value = IIf(IsNull(!fInicio), FechaServidor(), !fInicio)
            dtpFechaCaducida.value = IIf(IsNull(!fCaducidad), FechaServidor(), !fCaducidad)
         Else
            txtSerie.Text = IIf(IsNull(!tSerie), "", !tSerie)
            txtCorrelativo.Text = IIf(IsNull(!tUltimoNumero), "", !tUltimoNumero)
            txtAutorizacion.Text = ""
         End If
         
         txtPrefijoEnlace.Text = IIf(IsNull(!tPrefijoEnlace), "", !tPrefijoEnlace)
         
         txtFormVenta.Text = IIf(IsNull(!tFormVenta), "", !tFormVenta)
         txtCompVenta.Text = IIf(IsNull(!tCompVenta), "", !tCompVenta)
         
         If IsNull(!LEQUIVADOLARES) = True Then
            chkDocEquivDolares.value = 0
         ElseIf !LEQUIVADOLARES = 0 Then
            chkDocEquivDolares.value = 0
        Else
            chkDocEquivDolares.value = 1
         End If
                  
         If !TTipoEmision = "00" Then
            Frame14.Visible = False
            'Label(4).Visible = False
         Else
            'Label(4).Visible = true
            Frame14.Visible = True
         End If
        
         If sImpuesto1 <> "" And !TTipoEmision <> "00" Then
            chkImpuesto1.Visible = True
            chkImpuesto1.Caption = sImpuesto1
            chkImpuesto1.value = IIf(IsNull(!lImpuesto1), 0, IIf(!lImpuesto1, 1, 0))
         Else
            chkImpuesto1.Visible = False
            chkImpuesto1.value = 0
         End If
             
         If sImpuesto2 <> "" And !TTipoEmision <> "00" Then
            chkImpuesto2.Visible = True
            chkImpuesto2.Caption = sImpuesto2
            chkImpuesto2.value = IIf(IsNull(!lImpuesto2), 0, IIf(!lImpuesto2, 1, 0))
         Else
            chkImpuesto2.Visible = False
            chkImpuesto2.value = 0
         End If
             
         If sImpuesto3 <> "" And !TTipoEmision <> "00" Then
            chkImpuesto3.Visible = True
            chkImpuesto3.Caption = sImpuesto3
            chkImpuesto3.value = IIf(IsNull(!lImpuesto3), 0, IIf(!lImpuesto3, 1, 0))
        Else
           chkImpuesto3.Visible = False
           chkImpuesto3.value = 0
        End If
    End With
End Sub

Public Sub SubDetArea(Activa As Boolean)
   With cboImpArea
        Isql = "Select * from TIMPRESORA where tCaja = '" & txtCodigo.Text & "' order by tImpresora"
        Set RsImpArea = Lib.OpenRecordset(Isql, Cn)
        
         Set .RowSource = RsImpArea
             .DataField = "tDescripcion"
             .ListField = "tDescripcion"
             .BoundColumn = "tImpresora"
   End With

   fraArea.Visible = Not Activa
   ActivarBotones Activa
   cmdOpcion(1).Enabled = Activa
   cmdOpcion(3).Enabled = Activa
   
   cmdOpcionGrilla(5).Enabled = Activa
   cmdOpcionGrilla(6).Enabled = Activa
   cmdOpcionGrilla(7).Enabled = Activa
   
   'Controles
   txtDetallado.Enabled = Activa
   chkActivo.Enabled = Activa
End Sub

Sub SubArea()
    With RsAI
         cboArea.BoundText = IIf(IsNull(!tArea), "", !tArea)
         cboImpArea.BoundText = IIf(IsNull(!timpresora), "", Trim(!timpresora))
    End With
End Sub

Public Sub inicio()
    RsGrilla.Filter = "tCaja ='" & txtCodigo.Text & "'"
    RsAI.Filter = "tCaja ='" & txtCodigo.Text & "'"
    
      
    rsAreaSubGrupo.Filter = "tcaja='" & txtCodigo.Text & "'"
    
    cmdOpcionGrilla(0).Enabled = False
    cmdOpcionGrilla(1).Enabled = False
    cmdOpcionGrilla(2).Enabled = False
    cmdOpcion(0).Enabled = False
    cmdOpcion(2).Enabled = False
    
    chkVComanda.value = 0
    chkComanda.value = 0
    chkObligaPrinter.value = 0
    chkObligaPrecuenta.value = 0
   ' chkObliga.value = 0
   'chkMozo.value = 0
  '  chkMotorizado.value = 0
    chkAdicion.value = 0
   ' chkPax.value = 0
    chkConsumo1.value = 0
    chkConsumo2.value = 0
    chkConsumo3.value = 0
    
    chkCodigoReciboIngreso.value = 0
    
    chkComboPrecuenta.value = 0
    chkEliminaC.value = 0
    chkPasswordC.value = 0
    chkElimina.value = 0
    chkPassword.value = 0
    chkAccesoDespachoPedido.value = 0
    chkObligaCierre.value = 0
    chkFiltroTipoPedido.value = 0
    chkEquivaPrecuenta = 0
    chkCancelacion.value = 0
    chkDirecto.value = 0
    chkCambioMesa.value = 0
    chkPreCuenta.value = 0
    chkBuscaPedido.value = 0
    chkImprimeImagCabPrecuenta.value = 0
    chkImprimeImagPiePrecuenta.value = 0
    chkAgrupada.value = 0
    chkSiab.value = 0
    chkComboDocumento.value = 0
    chkVisaNet.value = 0
    chkImpuestoPrecuenta.value = 0
    chkDocumentoAgrupado.value = 0
    chkOrden.value = 0
    chkActivo.value = 1
    chkValor.value = 0
    chkCajaMobile.value = 0
    txtLimitePrecuenta.Text = "0"
    txtLimiteReimpresion.Text = "0"
    chkPasswordTransferencia.value = 0
    chkPasswordImportar.value = 0
    chkDescripcionAlternativa.value = 0
    Me.chkMotDesc.value = 0
    Me.chkCajaContingencia.value = 0
    Me.chkImpPropina.value = 0
    chkCompatibilidadTVS.value = 0 'TVS
    chkCD.value = 0
    chkDisgrega.value = 0
    chkMultiCajero.value = 0
    chkMCPV.value = 0
   ' chkFechaDelivery.value = 0
    chkCCVOX.value = 0
    chkObservacionPrecuenta.value = 0
    chkObservacionDocumento.value = 0
    chkObservacionCabDoc.value = 0
    txtLongitudBarra.Text = "0"
    chkPagoRapido.value = 0
    
    chkPasswordPorCobrar.value = 0
    chkModificaTipoPedido.value = 0
    txtBalanzaPuerto.Text = ""
    chkBloqueaPrecuenta.value = 0
    txtLimitePrecuenta.Enabled = True
        chkMulti1.value = 0
    chkMulti2.value = 0
        Me.chkBuscaPedidoFiltrarMesa.value = 0
    Me.chkBuscaPedidoVisualizaGrilla.value = 0
    Me.chkPassOtrosPagos.value = 0
End Sub

'lg

Private Sub chkMulti1_Click()
    If chkMulti1.value Then
        chkMulti2.value = 0
        fra2.Enabled = False
        fra1.Enabled = True
    Else
            fra1.Enabled = False
    End If
End Sub

Private Sub chkMulti2_Click()
    If chkMulti2.value Then
        chkMulti1.value = 0
        fra1.Enabled = False
        fra2.Enabled = True
    Else
        fra2.Enabled = False
    End If
End Sub

'CESAR AREA CHEF
Public Sub SubDetAreaChef(Activa As Boolean)
   fraAreaChef.Visible = Not Activa
   ActivarBotones Activa
   cmdOpcion(1).Enabled = Activa
   cmdOpcion(3).Enabled = Activa
   'A-M-E
   cmdOpcionGrilla(13).Enabled = Activa
   cmdOpcionGrilla(14).Enabled = Activa
   cmdOpcionGrilla(12).Enabled = Activa
   
   'Controles
   txtDetallado.Enabled = Activa
   chkActivo.Enabled = Activa
End Sub
Sub SubAreaChef()
    With RsAChef
         cboAreaChef.BoundText = IIf(IsNull(!tArea), "", !tArea)
         chkAreaChef = IIf(IsNull(!AreaChef), 0, IIf(!AreaChef, 1, 0))
    End With
End Sub
'------------------------
