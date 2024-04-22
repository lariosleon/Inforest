VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "tabctl32.ocx"
Begin VB.Form frmParametro 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Parámetros Generales"
   ClientHeight    =   8805
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   10545
   ForeColor       =   &H00808080&
   Icon            =   "frmParametro.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8805
   ScaleWidth      =   10545
   Begin VB.CheckBox chkKDS 
      Height          =   240
      Left            =   120
      TabIndex        =   148
      Top             =   8400
      Width           =   225
   End
   Begin VB.CommandButton btnKDS 
      Caption         =   "KDS"
      Enabled         =   0   'False
      Height          =   480
      Left            =   360
      TabIndex        =   147
      Top             =   8280
      Width           =   645
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   8130
      Left            =   0
      TabIndex        =   49
      Top             =   30
      Width           =   10485
      _ExtentX        =   18494
      _ExtentY        =   14340
      _Version        =   393216
      Tabs            =   4
      Tab             =   3
      TabsPerRow      =   4
      TabHeight       =   520
      ForeColor       =   -2147483630
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      TabCaption(0)   =   "Datos Generales"
      TabPicture(0)   =   "frmParametro.frx":000C
      Tab(0).ControlEnabled=   0   'False
      Tab(0).Control(0)=   "Frame1"
      Tab(0).Control(1)=   "Frame5"
      Tab(0).Control(2)=   "Frame2"
      Tab(0).Control(3)=   "fraRuc"
      Tab(0).ControlCount=   4
      TabCaption(1)   =   "Otros Datos"
      TabPicture(1)   =   "frmParametro.frx":0028
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Frame28"
      Tab(1).Control(1)=   "Frame35"
      Tab(1).Control(2)=   "Frame33"
      Tab(1).Control(3)=   "Frame29"
      Tab(1).Control(4)=   "Frame23"
      Tab(1).Control(5)=   "Frame22"
      Tab(1).Control(6)=   "Frame21"
      Tab(1).Control(7)=   "Frame15"
      Tab(1).Control(8)=   "Frame16"
      Tab(1).Control(9)=   "Frame12"
      Tab(1).Control(10)=   "Frame11"
      Tab(1).Control(11)=   "Frame9"
      Tab(1).Control(12)=   "Frame7"
      Tab(1).Control(13)=   "Frame4"
      Tab(1).Control(14)=   "Frame6"
      Tab(1).Control(15)=   "Frame8"
      Tab(1).ControlCount=   16
      TabCaption(2)   =   "Complementos"
      TabPicture(2)   =   "frmParametro.frx":0044
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "Frame24"
      Tab(2).Control(1)=   "FrmVisor"
      Tab(2).Control(2)=   "Frame18"
      Tab(2).Control(3)=   "Frame17"
      Tab(2).Control(4)=   "Frame14"
      Tab(2).Control(5)=   "Frame13"
      Tab(2).Control(6)=   "frmMobile"
      Tab(2).ControlCount=   7
      TabCaption(3)   =   "Facturacion Electronica"
      TabPicture(3)   =   "frmParametro.frx":0060
      Tab(3).ControlEnabled=   -1  'True
      Tab(3).Control(0)=   "Frame10"
      Tab(3).Control(0).Enabled=   0   'False
      Tab(3).ControlCount=   1
      Begin VB.Frame Frame28 
         Caption         =   "Pedidos"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1035
         Left            =   -71400
         TabIndex        =   243
         Top             =   5760
         Width           =   3165
         Begin VB.CheckBox chkBloqInafecto 
            Alignment       =   1  'Right Justify
            Caption         =   "No permitir comandar Platos Inafectos y afectos en un mismo pedido."
            Height          =   615
            Left            =   120
            TabIndex        =   244
            ToolTipText     =   "Si esta Activo el check no se permitira comandar Items Inafectos en pedido Afecto."
            Top             =   240
            Width           =   2805
         End
      End
      Begin VB.Frame Frame35 
         Caption         =   "Descargo de Ventas"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   660
         Left            =   -71400
         TabIndex        =   239
         Top             =   4440
         Width           =   3165
         Begin VB.CheckBox chkValidaStock 
            Alignment       =   1  'Right Justify
            Caption         =   "Validar Stock de Insumos en Descargo"
            Height          =   375
            Left            =   120
            TabIndex        =   240
            Top             =   240
            Width           =   2805
         End
      End
      Begin VB.Frame Frame33 
         Caption         =   "Anticipos"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1020
         Left            =   -68160
         TabIndex        =   233
         Top             =   4080
         Width           =   3525
         Begin VB.CommandButton cmdBuscaAnticipo 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   330
            Left            =   2880
            Picture         =   "frmParametro.frx":007C
            Style           =   1  'Graphical
            TabIndex        =   242
            Top             =   600
            Width           =   630
         End
         Begin VB.CheckBox chkActivaAnticipo 
            Alignment       =   1  'Right Justify
            Caption         =   "Activar Anticipo"
            Height          =   255
            Left            =   120
            TabIndex        =   236
            Top             =   240
            Width           =   2055
         End
         Begin VB.TextBox txtCodigoAnticipo 
            Alignment       =   1  'Right Justify
            Enabled         =   0   'False
            Height          =   285
            Left            =   1680
            TabIndex        =   235
            Tag             =   "02155454555"
            Top             =   600
            Width           =   1215
         End
         Begin VB.Label Label10 
            Caption         =   "Cod de Item Anticipo:"
            Height          =   255
            Left            =   120
            TabIndex        =   234
            Top             =   600
            Width           =   1575
         End
      End
      Begin VB.Frame Frame29 
         Caption         =   "Caja Rapida - Pago de Documentos "
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1065
         Left            =   -74880
         TabIndex        =   223
         Top             =   6960
         Width           =   3360
         Begin VB.CheckBox chkPagoCheque 
            Alignment       =   1  'Right Justify
            Caption         =   "Desactivar Forma de pago Cheque/Depositos :"
            Height          =   435
            Left            =   120
            TabIndex        =   225
            ToolTipText     =   "Si esta seleccionado, orienta el sistema a los clientes."
            Top             =   240
            Width           =   3090
         End
         Begin VB.CheckBox chkPagoOtra 
            Alignment       =   1  'Right Justify
            Caption         =   "Desactivar Otras Formas de Pago :"
            Height          =   195
            Left            =   120
            TabIndex        =   224
            ToolTipText     =   "Si esta seleccionado, orienta el sistema a los clientes."
            Top             =   720
            Width           =   3090
         End
      End
      Begin VB.Frame Frame10 
         Caption         =   "Facturación Electrónica"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   7395
         Left            =   150
         TabIndex        =   201
         Top             =   465
         Width           =   10185
         Begin VB.Frame FrmFacEcuador 
            Height          =   6975
            Left            =   5040
            TabIndex        =   268
            Top             =   240
            Visible         =   0   'False
            Width           =   5055
            Begin VB.Frame Frame36 
               Caption         =   "Integracion Estupendo"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H8000000D&
               Height          =   600
               Left            =   120
               TabIndex        =   269
               Top             =   240
               Width           =   4820
               Begin VB.CheckBox chkFEEstupendo 
                  Caption         =   "Facturación Electrónica Estupendo"
                  BeginProperty Font 
                     Name            =   "Arial"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   210
                  Left            =   120
                  TabIndex        =   270
                  Top             =   300
                  Width           =   3060
               End
            End
         End
         Begin VB.Frame FrmFacPeru 
            Height          =   6975
            Left            =   5040
            TabIndex        =   245
            Top             =   240
            Visible         =   0   'False
            Width           =   5055
            Begin VB.Frame FrameOfiisis 
               Caption         =   "Integración Ofisis"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H8000000D&
               Height          =   750
               Left            =   120
               TabIndex        =   265
               Top             =   120
               Width           =   4820
               Begin VB.CheckBox chkFEOfisis 
                  Caption         =   "Facturación Electrónica Ofisis"
                  BeginProperty Font 
                     Name            =   "Arial"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   195
                  Left            =   120
                  TabIndex        =   266
                  Top             =   480
                  Width           =   2535
               End
               Begin VB.Label Label 
                  Alignment       =   1  'Right Justify
                  AutoSize        =   -1  'True
                  Caption         =   "Conexion a base de datos Sql Server"
                  BeginProperty Font 
                     Name            =   "Arial Narrow"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   -1  'True
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   225
                  Index           =   61
                  Left            =   1750
                  TabIndex        =   267
                  Top             =   165
                  Width           =   2865
                  WordWrap        =   -1  'True
               End
            End
            Begin VB.Frame Frame26 
               Caption         =   "Integracion Spring"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H8000000D&
               Height          =   825
               Left            =   120
               TabIndex        =   262
               Top             =   960
               Width           =   4820
               Begin VB.CheckBox chkFESpring 
                  Caption         =   "Facturación Electrónica Spring"
                  BeginProperty Font 
                     Name            =   "Arial"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   210
                  Left            =   120
                  TabIndex        =   263
                  Top             =   480
                  Width           =   2775
               End
               Begin VB.Label Label 
                  Alignment       =   1  'Right Justify
                  AutoSize        =   -1  'True
                  Caption         =   "Conexion a base de datos Sql Server"
                  BeginProperty Font 
                     Name            =   "Arial Narrow"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   -1  'True
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   225
                  Index           =   62
                  Left            =   1800
                  TabIndex        =   264
                  Top             =   165
                  Width           =   2865
                  WordWrap        =   -1  'True
               End
            End
            Begin VB.Frame Frame30 
               Caption         =   "Integracion Carvajal"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H8000000D&
               Height          =   480
               Left            =   120
               TabIndex        =   260
               Top             =   2295
               Width           =   4820
               Begin VB.CheckBox chkFEGesa 
                  Caption         =   "Grupo GESA"
                  BeginProperty Font 
                     Name            =   "Arial"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   195
                  Left            =   2400
                  TabIndex        =   271
                  Top             =   240
                  Width           =   1740
               End
               Begin VB.CheckBox chkFECarbajal 
                  Caption         =   "InfoRest - Carvajal"
                  BeginProperty Font 
                     Name            =   "Arial"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   195
                  Left            =   120
                  TabIndex        =   261
                  Top             =   240
                  Width           =   1740
               End
            End
            Begin VB.Frame Frame25 
               Caption         =   "Integracion Paperlees"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H8000000D&
               Height          =   480
               Left            =   120
               TabIndex        =   256
               Top             =   1815
               Width           =   4820
               Begin VB.CheckBox chkFEpape 
                  Caption         =   "Facturación Electrónica Paperlees"
                  BeginProperty Font 
                     Name            =   "Arial"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   195
                  Left            =   120
                  TabIndex        =   258
                  Top             =   240
                  Width           =   2820
               End
               Begin VB.CheckBox chkFEubl21 
                  Caption         =   "Activa UBL 2.1"
                  Enabled         =   0   'False
                  BeginProperty Font 
                     Name            =   "Arial"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   195
                  Left            =   3240
                  TabIndex        =   257
                  Top             =   240
                  Width           =   1500
               End
               Begin VB.Label Label 
                  Alignment       =   1  'Right Justify
                  AutoSize        =   -1  'True
                  Caption         =   "Conexion TCP IP"
                  BeginProperty Font 
                     Name            =   "Arial Narrow"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   -1  'True
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   225
                  Index           =   68
                  Left            =   3075
                  TabIndex        =   259
                  Top             =   0
                  Width           =   1545
                  WordWrap        =   -1  'True
               End
            End
            Begin VB.Frame fraPaCarvajal 
               Caption         =   "Parametros"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H8000000D&
               Height          =   2415
               Left            =   120
               TabIndex        =   252
               Top             =   4335
               Width           =   4815
               Begin VB.TextBox txtParamCarv 
                  Height          =   1575
                  Left            =   120
                  MultiLine       =   -1  'True
                  TabIndex        =   254
                  Top             =   240
                  Width           =   4575
               End
               Begin VB.TextBox txtCarvajalCorreos 
                  Height          =   405
                  Left            =   720
                  TabIndex        =   253
                  Top             =   1920
                  Width           =   3975
               End
               Begin VB.Label Label9 
                  Caption         =   "Correos:"
                  Height          =   255
                  Left            =   120
                  TabIndex        =   255
                  Top             =   2040
                  Width           =   735
               End
            End
            Begin VB.Frame Frame31 
               Caption         =   "Integracion TCI"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H8000000D&
               Height          =   480
               Left            =   120
               TabIndex        =   250
               Top             =   2775
               Width           =   4820
               Begin VB.CheckBox chkfeTCI 
                  Caption         =   "Facturación Electrónica TCI"
                  BeginProperty Font 
                     Name            =   "Arial"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   195
                  Left            =   120
                  TabIndex        =   251
                  Top             =   240
                  Width           =   3060
               End
            End
            Begin VB.Frame Frame32 
               Caption         =   "Integracion Bizlinks"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H8000000D&
               Height          =   480
               Left            =   120
               TabIndex        =   248
               Top             =   3255
               Width           =   4820
               Begin VB.CheckBox chkFEBiz 
                  Caption         =   "Facturación Electrónica Bizlinks"
                  BeginProperty Font 
                     Name            =   "Arial"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   195
                  Left            =   120
                  TabIndex        =   249
                  Top             =   240
                  Width           =   3060
               End
            End
            Begin VB.Frame Frame34 
               Caption         =   "Integracion Good Hope"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H8000000D&
               Height          =   480
               Left            =   120
               TabIndex        =   246
               Top             =   3735
               Width           =   4820
               Begin VB.CheckBox chkFEGood 
                  Caption         =   "Facturación Electrónica Good Hope"
                  BeginProperty Font 
                     Name            =   "Arial"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   195
                  Left            =   120
                  TabIndex        =   247
                  Top             =   240
                  Width           =   3060
               End
            End
         End
         Begin VB.CheckBox chkInNC 
            Caption         =   "Incluir Notas de Credito en Liquidacion de Cajeros"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Left            =   360
            TabIndex        =   232
            Top             =   2040
            Width           =   4215
         End
         Begin VB.CheckBox chkAnulacionNC 
            Caption         =   "Activar notas de credito por anulacion de documentos"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Left            =   360
            TabIndex        =   226
            Top             =   1800
            Width           =   4455
         End
         Begin VB.Frame Frame27 
            Caption         =   "Conexión a Base de Datos"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   2040
            Left            =   135
            TabIndex        =   212
            Top             =   2280
            Width           =   4820
            Begin VB.CommandButton cmdValidar 
               Caption         =   "Validar"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   675
               Index           =   2
               Left            =   3255
               Picture         =   "frmParametro.frx":017E
               Style           =   1  'Graphical
               TabIndex        =   222
               Top             =   1230
               UseMaskColor    =   -1  'True
               Width           =   1410
            End
            Begin VB.TextBox txtClaveFE 
               Alignment       =   2  'Center
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
               IMEMode         =   3  'DISABLE
               Left            =   1410
               PasswordChar    =   "*"
               TabIndex        =   221
               Top             =   1560
               Width           =   1800
            End
            Begin VB.TextBox txtUsuarioFE 
               Alignment       =   2  'Center
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
               Left            =   1410
               TabIndex        =   220
               Top             =   1245
               Width           =   1800
            End
            Begin VB.TextBox txtServidorFE 
               Alignment       =   2  'Center
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
               Left            =   1410
               TabIndex        =   215
               Top             =   615
               Width           =   3250
            End
            Begin VB.TextBox txtBDFE 
               Alignment       =   2  'Center
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
               Left            =   1410
               TabIndex        =   214
               Top             =   930
               Width           =   3250
            End
            Begin VB.Label Label 
               Alignment       =   1  'Right Justify
               AutoSize        =   -1  'True
               Caption         =   "Clave :"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   210
               Index           =   64
               Left            =   285
               TabIndex        =   219
               Top             =   1590
               Width           =   1050
               WordWrap        =   -1  'True
            End
            Begin VB.Label Label 
               Alignment       =   1  'Right Justify
               AutoSize        =   -1  'True
               Caption         =   "Usuario :"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   210
               Index           =   63
               Left            =   285
               TabIndex        =   218
               Top             =   1290
               Width           =   1050
               WordWrap        =   -1  'True
            End
            Begin VB.Label Label 
               Alignment       =   1  'Right Justify
               AutoSize        =   -1  'True
               Caption         =   "Servidor :"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   210
               Index           =   58
               Left            =   240
               TabIndex        =   217
               Top             =   660
               Width           =   1110
               WordWrap        =   -1  'True
            End
            Begin VB.Label Label 
               Alignment       =   1  'Right Justify
               AutoSize        =   -1  'True
               Caption         =   "Base Datos :"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   210
               Index           =   59
               Left            =   300
               TabIndex        =   216
               Top             =   975
               Width           =   1050
               WordWrap        =   -1  'True
            End
            Begin VB.Label Label 
               Alignment       =   1  'Right Justify
               AutoSize        =   -1  'True
               Caption         =   "Motor BD: Sql Server"
               BeginProperty Font 
                  Name            =   "Arial Narrow"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   -1  'True
                  Italic          =   -1  'True
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   225
               Index           =   65
               Left            =   300
               TabIndex        =   213
               Top             =   285
               Width           =   1740
               WordWrap        =   -1  'True
            End
         End
         Begin VB.Frame FrameTipoImpresion 
            Caption         =   "Tipo Impresión"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   1080
            Left            =   3060
            TabIndex        =   206
            Top             =   210
            Width           =   1800
            Begin VB.OptionButton optOpcion 
               Caption         =   "Código Barras"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   200
               Index           =   0
               Left            =   120
               TabIndex        =   209
               Top             =   300
               Value           =   -1  'True
               Width           =   1485
            End
            Begin VB.OptionButton optOpcion 
               Caption         =   "Código Hash"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   255
               Index           =   1
               Left            =   120
               TabIndex        =   208
               Top             =   540
               Width           =   1275
            End
            Begin VB.OptionButton optOpcion 
               Caption         =   "Código QR"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   255
               Index           =   2
               Left            =   120
               TabIndex        =   207
               Top             =   800
               Width           =   1275
            End
         End
         Begin VB.CheckBox chkFacturacionE 
            Alignment       =   1  'Right Justify
            Caption         =   "Facturación Electrónica  :"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Left            =   225
            TabIndex        =   205
            Top             =   345
            Width           =   2145
         End
         Begin VB.CheckBox chkAmbienteFE 
            Alignment       =   1  'Right Justify
            Caption         =   "Ambiente Producción  :"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Left            =   390
            TabIndex        =   204
            Top             =   600
            Width           =   1980
         End
         Begin VB.TextBox txtCodigoFE 
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
            Left            =   2175
            MaxLength       =   3
            TabIndex        =   203
            Top             =   945
            Width           =   750
         End
         Begin VB.TextBox txtRutaImgFE 
            Alignment       =   1  'Right Justify
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
            Left            =   1560
            TabIndex        =   202
            Top             =   1485
            Width           =   3315
         End
         Begin VB.Label Label 
            Alignment       =   1  'Right Justify
            Caption         =   "Código Facturación :"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   210
            Index           =   50
            Left            =   570
            TabIndex        =   211
            Top             =   960
            Width           =   1515
         End
         Begin VB.Label lblImgFE 
            Alignment       =   1  'Right Justify
            Caption         =   "Ruta de Imagen :"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   270
            Left            =   225
            TabIndex        =   210
            Top             =   1515
            Width           =   1275
         End
      End
      Begin VB.Frame Frame24 
         Caption         =   "Cheff Control"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   1095
         Left            =   -74910
         TabIndex        =   198
         Top             =   3540
         Width           =   5535
         Begin VB.CheckBox chkCheffFiltroSubGrupo 
            Caption         =   "Permite el Filtro de Pedidos por Sub-Grupos."
            Height          =   285
            Left            =   240
            TabIndex        =   200
            Top             =   675
            Width           =   3975
         End
         Begin VB.CheckBox chkCheffFiltroSalon 
            Caption         =   "Permite el Filtro de Pedidos por Salones."
            Height          =   315
            Left            =   240
            TabIndex        =   199
            Top             =   300
            Width           =   3975
         End
      End
      Begin VB.Frame Frame23 
         Caption         =   "Configuracion - Nota de Credito"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1245
         Left            =   -74880
         TabIndex        =   196
         Top             =   5640
         Width           =   3360
         Begin VB.CheckBox chkDesNCPG 
            Alignment       =   1  'Right Justify
            Caption         =   "Bloquear NC como Forma de Pago :"
            Height          =   195
            Left            =   120
            TabIndex        =   231
            ToolTipText     =   "Si esta seleccionado, orienta el sistema a los clientes."
            Top             =   960
            Width           =   3090
         End
         Begin VB.CheckBox chkNCElimina 
            Alignment       =   1  'Right Justify
            Caption         =   "Bloquear Eliminar NC"
            Height          =   195
            Left            =   120
            TabIndex        =   228
            ToolTipText     =   "Si esta seleccionado, orienta el sistema a los clientes."
            Top             =   720
            Width           =   3090
         End
         Begin VB.CheckBox chkNCParcial 
            Alignment       =   1  'Right Justify
            Caption         =   "Bloquear NC Parciales"
            Height          =   195
            Left            =   120
            TabIndex        =   227
            ToolTipText     =   "Si esta seleccionado, orienta el sistema a los clientes."
            Top             =   480
            Width           =   3090
         End
         Begin VB.CheckBox chkNCFecha 
            Alignment       =   1  'Right Justify
            Caption         =   "Bloquedo de Fecha"
            Height          =   195
            Left            =   120
            TabIndex        =   197
            ToolTipText     =   "Si esta seleccionado, orienta el sistema a los clientes."
            Top             =   240
            Width           =   3090
         End
      End
      Begin VB.Frame Frame22 
         Caption         =   "Configuracion -  BAR"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1695
         Left            =   -68160
         TabIndex        =   188
         Top             =   2400
         Width           =   3495
         Begin VB.CommandButton cmdbuscarItemCover 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   330
            Left            =   2760
            Picture         =   "frmParametro.frx":05C0
            Style           =   1  'Graphical
            TabIndex        =   241
            Top             =   1200
            Width           =   630
         End
         Begin VB.TextBox txtCodigoItemCover 
            Alignment       =   1  'Right Justify
            Enabled         =   0   'False
            Height          =   285
            Left            =   1680
            TabIndex        =   193
            Tag             =   "02155454555"
            Top             =   1200
            Width           =   1095
         End
         Begin VB.TextBox txtMontoMinCover 
            Alignment       =   1  'Right Justify
            Height          =   285
            Left            =   1680
            TabIndex        =   191
            Top             =   720
            Width           =   1095
         End
         Begin VB.CheckBox chkCover 
            Alignment       =   1  'Right Justify
            Caption         =   "Activar cargo automático de Cover a Pedido"
            Height          =   375
            Left            =   240
            TabIndex        =   189
            Top             =   240
            Width           =   2565
         End
         Begin VB.Label Label7 
            Caption         =   "Codigo de Item Cover"
            Height          =   255
            Left            =   120
            TabIndex        =   192
            Top             =   1200
            Width           =   1695
         End
         Begin VB.Label Label6 
            Caption         =   "Monto Minimo / PAX (Cliente)"
            Height          =   375
            Left            =   120
            TabIndex        =   190
            Top             =   720
            Width           =   1575
         End
      End
      Begin VB.Frame FrmVisor 
         Caption         =   "Visor de 8"""
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   1095
         Left            =   -74880
         TabIndex        =   183
         Top             =   6525
         Width           =   5535
         Begin VB.TextBox txtvisortiempo 
            Alignment       =   1  'Right Justify
            Height          =   285
            Left            =   4320
            TabIndex        =   186
            Top             =   600
            Visible         =   0   'False
            Width           =   975
         End
         Begin VB.CheckBox chktactil 
            Alignment       =   1  'Right Justify
            Caption         =   "Visor Tactil"
            Height          =   195
            Left            =   4080
            TabIndex        =   185
            Top             =   240
            Visible         =   0   'False
            Width           =   1215
         End
         Begin VB.CheckBox chkVisor8 
            Alignment       =   1  'Right Justify
            Caption         =   "Activar Visor de 8"" (AMC)"
            Height          =   255
            Left            =   120
            TabIndex        =   184
            Top             =   255
            Width           =   2295
         End
         Begin VB.Label Label5 
            Caption         =   "Tiempo de espera de mensaje e Inactividad. (Segundos) :"
            Height          =   255
            Left            =   120
            TabIndex        =   187
            Top             =   600
            Visible         =   0   'False
            Width           =   4215
         End
      End
      Begin VB.Frame Frame21 
         Caption         =   "Transferencia Gratuita"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1335
         Left            =   -71400
         TabIndex        =   180
         Top             =   3120
         Width           =   3165
         Begin VB.TextBox txtGlosaImpresion 
            Height          =   495
            Left            =   1440
            MultiLine       =   -1  'True
            TabIndex        =   195
            Top             =   720
            Width           =   1695
         End
         Begin VB.TextBox TxtCuentaContable 
            Height          =   285
            Left            =   1440
            TabIndex        =   182
            Top             =   300
            Width           =   1695
         End
         Begin VB.Label Label8 
            Caption         =   "Glosa Impresion:"
            Height          =   255
            Left            =   120
            TabIndex        =   194
            Top             =   720
            Width           =   1335
         End
         Begin VB.Label Label4 
            Caption         =   "Cuenta Contable:"
            Height          =   255
            Left            =   120
            TabIndex        =   181
            Top             =   360
            Width           =   1335
         End
      End
      Begin VB.Frame Frame18 
         Caption         =   "Enlace SAP "
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   1665
         Left            =   -74880
         TabIndex        =   168
         Top             =   4695
         Width           =   5535
         Begin VB.Frame Frame20 
            Caption         =   "Datos Local"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   855
            Left            =   120
            TabIndex        =   175
            Top             =   600
            Width           =   5295
            Begin VB.TextBox TxtCodAlmcSAP 
               Alignment       =   1  'Right Justify
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
               Left            =   1320
               TabIndex        =   177
               Top             =   360
               Width           =   2895
            End
            Begin VB.Label Label3 
               Caption         =   "Codigo :"
               Height          =   255
               Left            =   600
               TabIndex        =   176
               Top             =   360
               Width           =   1095
            End
         End
         Begin VB.Frame Frame19 
            Caption         =   "Datos del Servidor"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   1215
            Left            =   120
            TabIndex        =   170
            Top             =   2910
            Width           =   5295
            Begin VB.CommandButton cmdConSAP 
               Caption         =   "Probar Conexion"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   735
               Left            =   4230
               Style           =   1  'Graphical
               TabIndex        =   178
               Top             =   330
               Width           =   975
            End
            Begin VB.TextBox TxtBaseSAP 
               Alignment       =   1  'Right Justify
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
               Left            =   1320
               TabIndex        =   174
               Top             =   720
               Width           =   2895
            End
            Begin VB.TextBox txtServidorSAP 
               Alignment       =   1  'Right Justify
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
               Left            =   1320
               TabIndex        =   171
               Top             =   360
               Width           =   2895
            End
            Begin VB.Label Label2 
               Caption         =   "Base de Datos :"
               Height          =   255
               Left            =   120
               TabIndex        =   173
               Top             =   720
               Width           =   1335
            End
            Begin VB.Label Label1 
               Caption         =   "Servidor :"
               Height          =   255
               Left            =   600
               TabIndex        =   172
               Top             =   360
               Width           =   855
            End
         End
         Begin VB.CheckBox ChkSAP 
            Caption         =   "Integración SAP"
            Height          =   375
            Left            =   240
            TabIndex        =   169
            Top             =   240
            Width           =   2295
         End
      End
      Begin VB.Frame Frame17 
         Caption         =   "Enlace Mesa247"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   1140
         Left            =   -74880
         TabIndex        =   163
         Top             =   4665
         Visible         =   0   'False
         Width           =   5505
         Begin VB.TextBox txtAdicionMesa247 
            Alignment       =   1  'Right Justify
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
            Left            =   2520
            MaxLength       =   3
            TabIndex        =   167
            Top             =   705
            Width           =   1320
         End
         Begin VB.TextBox txtCajaMesa247 
            Alignment       =   1  'Right Justify
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
            Left            =   2520
            MaxLength       =   3
            TabIndex        =   164
            Top             =   360
            Width           =   1320
         End
         Begin VB.Label Label 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "Adición Recepción de Pedido :"
            Height          =   255
            Index           =   60
            Left            =   120
            TabIndex        =   166
            Top             =   720
            Width           =   2250
            WordWrap        =   -1  'True
         End
         Begin VB.Label Label 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "Caja Recepción de Pedido :"
            Height          =   195
            Index           =   57
            Left            =   240
            TabIndex        =   165
            Top             =   405
            Width           =   2130
            WordWrap        =   -1  'True
         End
      End
      Begin VB.Frame Frame15 
         Caption         =   "Configuración Empresa"
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
         Left            =   -71400
         TabIndex        =   152
         Top             =   1440
         Width           =   3165
         Begin VB.TextBox txtCodigoMarca 
            Alignment       =   1  'Right Justify
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
            Left            =   1440
            MaxLength       =   2
            TabIndex        =   160
            Top             =   600
            Width           =   975
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
            Height          =   330
            Left            =   2400
            Picture         =   "frmParametro.frx":06C2
            Style           =   1  'Graphical
            TabIndex        =   159
            Top             =   1200
            Width           =   630
         End
         Begin VB.TextBox txtCodigoUbigeo 
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
            Left            =   1440
            TabIndex        =   155
            Top             =   1200
            Width           =   990
         End
         Begin VB.TextBox txtCodigoEmpresa 
            Alignment       =   1  'Right Justify
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
            Left            =   1440
            MaxLength       =   3
            TabIndex        =   154
            Top             =   300
            Width           =   975
         End
         Begin VB.TextBox txtCodigoTienda 
            Alignment       =   1  'Right Justify
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
            Left            =   1440
            MaxLength       =   4
            TabIndex        =   153
            Top             =   900
            Width           =   975
         End
         Begin VB.Label Label 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "Código Marca :"
            Height          =   195
            Index           =   56
            Left            =   0
            TabIndex        =   161
            Top             =   640
            Width           =   1365
            WordWrap        =   -1  'True
         End
         Begin VB.Label Label 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "Ubigeo :"
            Height          =   210
            Index           =   55
            Left            =   720
            TabIndex        =   158
            Top             =   1200
            Width           =   645
            WordWrap        =   -1  'True
         End
         Begin VB.Label Label 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "Código Tienda :"
            Height          =   195
            Index           =   53
            Left            =   120
            TabIndex        =   157
            Top             =   920
            Width           =   1245
            WordWrap        =   -1  'True
         End
         Begin VB.Label Label 
            AutoSize        =   -1  'True
            Caption         =   "Código Empresa :"
            Height          =   195
            Index           =   54
            Left            =   120
            TabIndex        =   156
            Top             =   360
            Width           =   1245
         End
      End
      Begin VB.Frame Frame16 
         Caption         =   "Activación Club"
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
         Left            =   -71400
         TabIndex        =   144
         Top             =   5160
         Width           =   3165
         Begin VB.CheckBox chkClub 
            Alignment       =   1  'Right Justify
            Caption         =   "Activa Club :"
            Height          =   195
            Left            =   120
            TabIndex        =   145
            ToolTipText     =   "Si esta seleccionado, orienta el sistema a los clientes."
            Top             =   240
            Width           =   2805
         End
      End
      Begin VB.Frame Frame14 
         Caption         =   "Anfitrionas"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   2925
         Left            =   -74925
         TabIndex        =   127
         Top             =   495
         Width           =   5565
         Begin VB.CheckBox chkAgradecimiento 
            Caption         =   "Enviar Email de Agradecimiento por Reserva"
            Height          =   240
            Left            =   75
            TabIndex        =   141
            Top             =   5220
            Visible         =   0   'False
            Width           =   4440
         End
         Begin VB.CheckBox chkRecordatorio 
            Caption         =   "Enviar Email de Recordatorio de Reserva"
            Height          =   240
            Left            =   75
            TabIndex        =   140
            Top             =   2970
            Visible         =   0   'False
            Width           =   4440
         End
         Begin VB.CheckBox chkConfirmacion 
            Caption         =   "Enviar Email de Confirmación de Reserva"
            Height          =   240
            Left            =   75
            TabIndex        =   139
            Top             =   600
            Width           =   4440
         End
         Begin VB.TextBox txtAgradecimiento 
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
            Height          =   1950
            Left            =   75
            MaxLength       =   3500
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   138
            Top             =   5445
            Visible         =   0   'False
            Width           =   5385
         End
         Begin VB.TextBox txtRecordatorio 
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
            Height          =   1950
            Left            =   75
            MaxLength       =   3500
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   137
            Top             =   3195
            Visible         =   0   'False
            Width           =   5385
         End
         Begin VB.TextBox txtConfirmacion 
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
            Height          =   1950
            Left            =   75
            MaxLength       =   3500
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   136
            Top             =   825
            Width           =   5385
         End
         Begin VB.TextBox txtToleranciaReserva 
            Alignment       =   1  'Right Justify
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
            Left            =   4275
            MaxLength       =   3
            TabIndex        =   134
            Top             =   225
            Width           =   1170
         End
         Begin VB.Label Label 
            Alignment       =   1  'Right Justify
            Caption         =   "Min Toleracia Reserva :"
            Height          =   210
            Index           =   52
            Left            =   1650
            TabIndex        =   135
            Top             =   255
            Width           =   2235
         End
      End
      Begin VB.Frame Frame13 
         Caption         =   "Central de Delivery"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   2130
         Left            =   -69240
         TabIndex        =   119
         Top             =   1965
         Width           =   4545
         Begin VB.TextBox txtMaxMotorizado 
            Alignment       =   2  'Center
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
            Left            =   2925
            MaxLength       =   5
            TabIndex        =   237
            Top             =   1480
            Width           =   1320
         End
         Begin VB.TextBox txtAsignacionMotorizado 
            Alignment       =   1  'Right Justify
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
            Left            =   2925
            MaxLength       =   8
            TabIndex        =   123
            Top             =   885
            Width           =   1320
         End
         Begin VB.TextBox txtTiempoDelivery 
            Alignment       =   1  'Right Justify
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
            Left            =   2925
            TabIndex        =   122
            Top             =   555
            Width           =   1320
         End
         Begin VB.TextBox txtDiaDelivery 
            Alignment       =   1  'Right Justify
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
            Left            =   2925
            TabIndex        =   121
            Top             =   225
            Width           =   1320
         End
         Begin VB.CheckBox chkHoraEntrega 
            Alignment       =   1  'Right Justify
            Caption         =   "Asignar hora de entrega desde Despachador"
            Height          =   255
            Left            =   225
            TabIndex        =   120
            Top             =   1200
            Width           =   4035
         End
         Begin VB.Label Label 
            AutoSize        =   -1  'True
            Caption         =   "Maximo de Motorizados a Asignar:"
            Height          =   195
            Index           =   69
            Left            =   240
            TabIndex        =   238
            Top             =   1560
            Width           =   2625
            WordWrap        =   -1  'True
         End
         Begin VB.Label Label 
            AutoSize        =   -1  'True
            Caption         =   "Asignación a Motorizados S/."
            Height          =   195
            Index           =   12
            Left            =   600
            TabIndex        =   126
            Top             =   930
            Width           =   2100
            WordWrap        =   -1  'True
         End
         Begin VB.Label Label 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "Días de búsqueda Delivery :"
            Height          =   195
            Index           =   42
            Left            =   450
            TabIndex        =   125
            Top             =   270
            Width           =   2250
            WordWrap        =   -1  'True
         End
         Begin VB.Label Label 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "Min promedio de entrega del :"
            Height          =   195
            Index           =   45
            Left            =   510
            TabIndex        =   124
            Top             =   600
            Width           =   2190
            WordWrap        =   -1  'True
         End
      End
      Begin VB.Frame frmMobile 
         Caption         =   "Mobile Inforest"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   1380
         Left            =   -69225
         TabIndex        =   116
         Top             =   495
         Width           =   4545
         Begin VB.CheckBox chkMCCaja 
            Alignment       =   1  'Right Justify
            Caption         =   "Activa solicitud de Autorización al  cambiar de Caja"
            Height          =   495
            Left            =   240
            TabIndex        =   118
            Top             =   840
            Width           =   3975
         End
         Begin VB.CheckBox chkMUnidadNegocio 
            Alignment       =   1  'Right Justify
            Caption         =   "Permite el Filtro por Unidad de Negocio en los dispositivos Móviles"
            Height          =   495
            Left            =   240
            TabIndex        =   117
            Top             =   360
            Width           =   3975
         End
      End
      Begin VB.Frame Frame12 
         Caption         =   "Equipos Biométricos (Huella Dactilar)"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   900
         Left            =   -68160
         TabIndex        =   110
         Top             =   1440
         Width           =   3525
         Begin VB.CheckBox chkSecugen 
            Alignment       =   1  'Right Justify
            Caption         =   "Hamster Plus (SecuGen) :      "
            Height          =   255
            Left            =   240
            TabIndex        =   112
            ToolTipText     =   "Modelo HSDU03P"
            Top             =   480
            Width           =   2565
         End
         Begin VB.CheckBox chkDigital 
            Alignment       =   1  'Right Justify
            Caption         =   "Digital Persona 4500  :"
            Height          =   255
            Left            =   240
            TabIndex        =   111
            Top             =   240
            Width           =   2565
         End
      End
      Begin VB.Frame Frame11 
         Caption         =   " Configuración de Factura Variable"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1215
         Left            =   -74880
         TabIndex        =   104
         Top             =   4440
         Width           =   3360
         Begin VB.TextBox txtCabeceraV 
            Alignment       =   1  'Right Justify
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
            Left            =   1800
            TabIndex        =   133
            Top             =   520
            Width           =   1470
         End
         Begin VB.TextBox txtItemV 
            Alignment       =   1  'Right Justify
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
            Left            =   1800
            TabIndex        =   132
            Top             =   240
            Width           =   1470
         End
         Begin VB.TextBox txtPieV 
            Alignment       =   1  'Right Justify
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
            Left            =   1800
            TabIndex        =   131
            Top             =   810
            Width           =   1470
         End
         Begin VB.Label Label 
            Alignment       =   1  'Right Justify
            Caption         =   "Posición Pie de Pag :"
            Height          =   210
            Index           =   41
            Left            =   150
            TabIndex        =   107
            Top             =   840
            Width           =   1515
         End
         Begin VB.Label Label 
            Alignment       =   1  'Right Justify
            Caption         =   "Posición Cabecera :"
            Height          =   210
            Index           =   46
            Left            =   180
            TabIndex        =   106
            Top             =   560
            Width           =   1485
         End
         Begin VB.Label Label 
            Alignment       =   1  'Right Justify
            Caption         =   "Cantidad de Items :"
            Height          =   210
            Index           =   47
            Left            =   255
            TabIndex        =   105
            Top             =   280
            Width           =   1410
         End
      End
      Begin VB.Frame Frame9 
         Caption         =   " Varios "
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1020
         Left            =   -71400
         TabIndex        =   31
         Top             =   360
         Width           =   3165
         Begin VB.TextBox txtDia 
            Alignment       =   1  'Right Justify
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
            Left            =   1440
            TabIndex        =   100
            Top             =   570
            Width           =   1470
         End
         Begin VB.TextBox txtCorrelativo 
            Alignment       =   1  'Right Justify
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
            Left            =   1440
            TabIndex        =   99
            Top             =   180
            Width           =   1470
         End
         Begin VB.Label Label 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "Días en la grilla :"
            Height          =   195
            Index           =   40
            Left            =   45
            TabIndex        =   102
            Top             =   585
            Width           =   1245
            WordWrap        =   -1  'True
         End
         Begin VB.Label Label 
            AutoSize        =   -1  'True
            Caption         =   "Correlativo :"
            Height          =   195
            Index           =   29
            Left            =   450
            TabIndex        =   101
            Top             =   225
            Width           =   840
         End
      End
      Begin VB.Frame Frame7 
         Caption         =   " Configuración de Factura Manual "
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1215
         Left            =   -74880
         TabIndex        =   33
         Top             =   3240
         Width           =   3360
         Begin VB.TextBox txtItem 
            Alignment       =   1  'Right Justify
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
            Left            =   1800
            TabIndex        =   130
            Top             =   225
            Width           =   1470
         End
         Begin VB.TextBox txtCabecera 
            Alignment       =   1  'Right Justify
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
            Left            =   1800
            TabIndex        =   129
            Top             =   525
            Width           =   1470
         End
         Begin VB.TextBox txtDetalle 
            Alignment       =   1  'Right Justify
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
            Left            =   1800
            TabIndex        =   128
            Top             =   825
            Width           =   1470
         End
         Begin VB.Label Label 
            Alignment       =   1  'Right Justify
            Caption         =   "Posición Detalle :"
            Height          =   210
            Index           =   17
            Left            =   255
            TabIndex        =   98
            Top             =   885
            Width           =   1410
         End
         Begin VB.Label Label 
            Alignment       =   1  'Right Justify
            Caption         =   "Posición Cabecera :"
            Height          =   210
            Index           =   15
            Left            =   180
            TabIndex        =   97
            Top             =   585
            Width           =   1485
         End
         Begin VB.Label Label 
            Alignment       =   1  'Right Justify
            Caption         =   "Cantidad de Items :"
            Height          =   210
            Index           =   1
            Left            =   255
            TabIndex        =   96
            Top             =   285
            Width           =   1410
         End
      End
      Begin VB.Frame Frame4 
         Caption         =   " Puntos  "
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1020
         Left            =   -68160
         TabIndex        =   32
         Top             =   360
         Width           =   3525
         Begin VB.TextBox txtClub 
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
            Left            =   1785
            TabIndex        =   95
            Top             =   210
            Width           =   1470
         End
         Begin VB.TextBox txtPunto 
            Alignment       =   1  'Right Justify
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
            Left            =   1785
            TabIndex        =   94
            Top             =   580
            Width           =   1470
         End
         Begin VB.Label Label 
            AutoSize        =   -1  'True
            Caption         =   "Descripción Puntos :"
            Height          =   195
            Index           =   38
            Left            =   300
            TabIndex        =   151
            Top             =   255
            Width           =   1470
         End
         Begin VB.Label Label 
            AutoSize        =   -1  'True
            Caption         =   "Valor de un Punto S/. :"
            Height          =   195
            Index           =   37
            Left            =   135
            TabIndex        =   150
            Top             =   600
            Width           =   1635
         End
      End
      Begin VB.Frame Frame6 
         Caption         =   " Dia Contable"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1620
         Left            =   -74880
         TabIndex        =   29
         Top             =   360
         Width           =   3360
         Begin VB.OptionButton optDCAutomatico 
            Caption         =   "Automático / Hora de Cierre"
            Height          =   465
            Left            =   195
            TabIndex        =   92
            Top             =   240
            Value           =   -1  'True
            Width           =   1485
         End
         Begin VB.OptionButton optDCManual 
            Caption         =   "Manual (Cierre de Turno)"
            Height          =   195
            Left            =   180
            TabIndex        =   91
            Top             =   840
            Width           =   2055
         End
         Begin VB.CheckBox chkImprimeDiaContable 
            Caption         =   "Impresión de Dia Contable en Documentos"
            Height          =   435
            Left            =   195
            TabIndex        =   90
            Top             =   1080
            Width           =   2895
         End
         Begin MSComCtl2.DTPicker dtpHoraDC 
            Height          =   315
            Left            =   1800
            TabIndex        =   93
            Top             =   360
            Width           =   1335
            _ExtentX        =   2355
            _ExtentY        =   556
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
            CustomFormat    =   "HH:mm 'HRS'"
            Format          =   87293955
            UpDown          =   -1  'True
            CurrentDate     =   38587.2083333333
         End
      End
      Begin VB.Frame Frame8 
         Caption         =   " Configuración de Guía"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1215
         Left            =   -74880
         TabIndex        =   30
         Top             =   2040
         Width           =   3360
         Begin VB.TextBox txtItemGuia 
            Alignment       =   1  'Right Justify
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
            Left            =   1815
            TabIndex        =   86
            Top             =   210
            Width           =   1470
         End
         Begin VB.TextBox txtCabeceraGuia 
            Alignment       =   1  'Right Justify
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
            Left            =   1800
            TabIndex        =   85
            Top             =   510
            Width           =   1470
         End
         Begin VB.TextBox txtDetalleGuia 
            Alignment       =   1  'Right Justify
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
            Left            =   1800
            TabIndex        =   84
            Top             =   810
            Width           =   1470
         End
         Begin VB.Label Label 
            Alignment       =   1  'Right Justify
            Caption         =   "Posición Cabecera :"
            Height          =   210
            Index           =   34
            Left            =   180
            TabIndex        =   88
            Top             =   540
            Width           =   1485
         End
         Begin VB.Label Label 
            Alignment       =   1  'Right Justify
            Caption         =   "Cantidad de Items :"
            Height          =   210
            Index           =   35
            Left            =   270
            TabIndex        =   89
            Top             =   240
            Width           =   1410
         End
         Begin VB.Label Label 
            Alignment       =   1  'Right Justify
            Caption         =   "Posición Detalle :"
            Height          =   210
            Index           =   33
            Left            =   255
            TabIndex        =   87
            Top             =   840
            Width           =   1410
         End
      End
      Begin VB.Frame fraRuc 
         Caption         =   "Identificador Tributario "
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
         Left            =   -68040
         TabIndex        =   28
         Top             =   6315
         Width           =   3345
         Begin VB.TextBox txtLongitud 
            Alignment       =   1  'Right Justify
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
            Left            =   2550
            TabIndex        =   81
            Top             =   420
            Width           =   615
         End
         Begin VB.OptionButton opcLongitud 
            Caption         =   " = Longitud"
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
            Index           =   0
            Left            =   1830
            TabIndex        =   80
            Top             =   900
            Width           =   1350
         End
         Begin VB.OptionButton opcLongitud 
            Caption         =   " > = Longitud"
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
            Index           =   1
            Left            =   1830
            TabIndex        =   79
            Top             =   1275
            Width           =   1455
         End
         Begin VB.Label Label 
            AutoSize        =   -1  'True
            Caption         =   "Longitud Identificador Tributario"
            Height          =   195
            Index           =   31
            Left            =   120
            TabIndex        =   83
            Top             =   465
            Width           =   2235
         End
         Begin VB.Label Label 
            AutoSize        =   -1  'True
            Caption         =   "Condición de Longitud"
            Height          =   195
            Index           =   32
            Left            =   120
            TabIndex        =   82
            Top             =   885
            Width           =   1590
         End
      End
      Begin VB.Frame Frame2 
         Caption         =   " Impuestos y Extras"
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
         Left            =   -74880
         TabIndex        =   27
         Top             =   6315
         Width           =   6615
         Begin VB.TextBox txtIImp3 
            Alignment       =   1  'Right Justify
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
            Left            =   4890
            TabIndex        =   24
            Top             =   1275
            Width           =   690
         End
         Begin VB.TextBox txtIImp2 
            Alignment       =   1  'Right Justify
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
            Left            =   4890
            TabIndex        =   22
            Top             =   840
            Width           =   690
         End
         Begin VB.TextBox txtIImp1 
            Alignment       =   1  'Right Justify
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
            Left            =   4890
            TabIndex        =   20
            Top             =   405
            Width           =   690
         End
         Begin VB.TextBox txtDImp3 
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
            Left            =   1095
            MaxLength       =   50
            TabIndex        =   23
            Top             =   1275
            Width           =   3630
         End
         Begin VB.TextBox txtDImp2 
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
            Left            =   1095
            MaxLength       =   50
            TabIndex        =   21
            Top             =   840
            Width           =   3630
         End
         Begin VB.TextBox txtDImp1 
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
            Left            =   1095
            MaxLength       =   50
            TabIndex        =   19
            Top             =   405
            Width           =   3630
         End
         Begin VB.Label Label 
            AutoSize        =   -1  'True
            Caption         =   "Porcentaje"
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
            Index           =   24
            Left            =   4875
            TabIndex        =   78
            Top             =   240
            Width           =   690
         End
         Begin VB.Label Label 
            AutoSize        =   -1  'True
            Caption         =   "Descripción"
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
            Index           =   25
            Left            =   2565
            TabIndex        =   77
            Top             =   240
            Width           =   750
         End
         Begin VB.Label Label 
            AutoSize        =   -1  'True
            Caption         =   "Extra 2 :"
            Height          =   195
            Index           =   28
            Left            =   330
            TabIndex        =   76
            Top             =   1320
            Width           =   585
         End
         Begin VB.Label Label 
            AutoSize        =   -1  'True
            Caption         =   "Extra 1 :"
            Height          =   195
            Index           =   27
            Left            =   330
            TabIndex        =   75
            Top             =   885
            Width           =   585
         End
         Begin VB.Label Label 
            AutoSize        =   -1  'True
            Caption         =   "Impuesto :"
            Height          =   195
            Index           =   26
            Left            =   195
            TabIndex        =   74
            Top             =   450
            Width           =   735
         End
         Begin VB.Label Label 
            AutoSize        =   -1  'True
            Caption         =   "%"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   240
            Index           =   23
            Left            =   5640
            TabIndex        =   73
            Top             =   1290
            Width           =   210
         End
         Begin VB.Label Label 
            AutoSize        =   -1  'True
            Caption         =   "%"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   240
            Index           =   22
            Left            =   5640
            TabIndex        =   72
            Top             =   855
            Width           =   210
         End
         Begin VB.Label Label 
            AutoSize        =   -1  'True
            Caption         =   "%"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   240
            Index           =   21
            Left            =   5640
            TabIndex        =   71
            Top             =   420
            Width           =   210
         End
      End
      Begin VB.Frame Frame5 
         Caption         =   " Activaciones "
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   5865
         Left            =   -66840
         TabIndex        =   26
         Top             =   375
         Width           =   2250
         Begin VB.CheckBox chkValidaDNI 
            Caption         =   "Validar DNI"
            Height          =   255
            Left            =   70
            TabIndex        =   230
            Top             =   5040
            Width           =   1215
         End
         Begin VB.CheckBox chkTCenImp 
            Caption         =   "Ver T.C. en Imp."
            Height          =   255
            Left            =   70
            TabIndex        =   229
            Top             =   5280
            Width           =   1815
         End
         Begin VB.CheckBox ChkActCuentaCorriente 
            Caption         =   "Act. de Cuentas C."
            Height          =   315
            Left            =   75
            TabIndex        =   179
            Top             =   5520
            Visible         =   0   'False
            Width           =   1935
         End
         Begin VB.CheckBox chkPagoAntesImpresion 
            Caption         =   "Registro de Pago antes de emisión de Comprobantes"
            Height          =   675
            Left            =   75
            TabIndex        =   162
            Top             =   4420
            Width           =   1935
         End
         Begin VB.CheckBox chkEventos 
            Caption         =   "Enlace Eventos"
            Height          =   195
            Left            =   75
            TabIndex        =   149
            Top             =   1150
            Width           =   1650
         End
         Begin VB.CheckBox chkTarjeta 
            Caption         =   "Validación Tarjeta"
            Height          =   195
            Left            =   75
            TabIndex        =   146
            Top             =   2100
            Width           =   1665
         End
         Begin VB.CheckBox chkControlEnviosProduccion 
            Caption         =   "Activa envios a producción por usuario"
            Height          =   435
            Left            =   75
            TabIndex        =   143
            Top             =   3960
            Width           =   1935
         End
         Begin VB.CheckBox chkEnvioAutomatico 
            Caption         =   "Envio a producción automático"
            Height          =   435
            Left            =   75
            TabIndex        =   142
            Top             =   3480
            Width           =   1815
         End
         Begin VB.CheckBox chkControlUsuario 
            Caption         =   "Control de Usuarios Por Nivel"
            Height          =   435
            Left            =   75
            TabIndex        =   109
            Top             =   3000
            Width           =   1815
         End
         Begin VB.CheckBox chkConsultaDescargo 
            Caption         =   "Activa Consulta de Descargo de Venta al Cierre de Turno"
            Height          =   615
            Left            =   75
            TabIndex        =   103
            Top             =   2350
            Width           =   2025
         End
         Begin VB.CheckBox chkMultiLocal 
            Caption         =   "Enlace Multilocal"
            Height          =   195
            Left            =   75
            TabIndex        =   70
            Top             =   1860
            Width           =   1665
         End
         Begin VB.CheckBox chkComboGeneral 
            Caption         =   "Listado General(Combos)"
            Height          =   300
            Left            =   75
            TabIndex        =   69
            Top             =   1580
            Width           =   2070
         End
         Begin VB.CheckBox ChkEquivalencia 
            Caption         =   "Muestra Equivalencia"
            Height          =   195
            Left            =   75
            TabIndex        =   68
            Top             =   1380
            Width           =   1890
         End
         Begin VB.CheckBox chkCierre 
            Caption         =   "Cierre a Ciegas"
            Height          =   195
            Left            =   75
            TabIndex        =   67
            Top             =   210
            Width           =   1365
         End
         Begin VB.CheckBox chkInfhotel 
            Caption         =   "Enlace Infhotel"
            Height          =   195
            Left            =   75
            TabIndex        =   66
            Top             =   930
            Width           =   1410
         End
         Begin VB.CheckBox chkAlmacen 
            Caption         =   "Enlace Almacén"
            Height          =   195
            Left            =   75
            TabIndex        =   65
            Top             =   690
            Width           =   1590
         End
         Begin VB.CheckBox chkPrinter 
            Caption         =   "Kitchen Printer"
            Height          =   195
            Left            =   75
            TabIndex        =   64
            Top             =   450
            Width           =   1365
         End
      End
      Begin VB.Frame Frame1 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   5865
         Left            =   -74880
         TabIndex        =   25
         Top             =   375
         Width           =   7965
         Begin VB.TextBox txtFax 
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
            Left            =   5680
            MaxLength       =   30
            TabIndex        =   115
            Top             =   1470
            Width           =   2175
         End
         Begin VB.TextBox txtRetencion 
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
            Left            =   2475
            MaxLength       =   150
            MultiLine       =   -1  'True
            TabIndex        =   18
            Top             =   5040
            Width           =   5385
         End
         Begin VB.TextBox txtPieFE 
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
            Left            =   2475
            MaxLength       =   255
            MultiLine       =   -1  'True
            TabIndex        =   17
            Top             =   4680
            Width           =   5385
         End
         Begin VB.TextBox txtPie 
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
            Left            =   2475
            MaxLength       =   255
            MultiLine       =   -1  'True
            TabIndex        =   15
            Top             =   4050
            Width           =   5385
         End
         Begin VB.TextBox txtDireccion 
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
            Left            =   2475
            MaxLength       =   150
            TabIndex        =   2
            Top             =   820
            Width           =   5385
         End
         Begin VB.TextBox txtEmail 
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
            Left            =   2475
            MaxLength       =   50
            TabIndex        =   5
            Top             =   1785
            Width           =   3345
         End
         Begin VB.TextBox txtMonedaE 
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
            Left            =   3450
            MaxLength       =   30
            TabIndex        =   11
            Top             =   3075
            Width           =   2355
         End
         Begin VB.TextBox txtMonE 
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
            Left            =   2475
            MaxLength       =   3
            TabIndex        =   10
            Top             =   3075
            Width           =   915
         End
         Begin VB.TextBox txtTelefono 
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
            Left            =   2475
            MaxLength       =   30
            TabIndex        =   4
            Top             =   1480
            Width           =   2385
         End
         Begin VB.TextBox txtRUC 
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
            Left            =   2475
            MaxLength       =   15
            TabIndex        =   7
            Top             =   2430
            Width           =   3345
         End
         Begin VB.TextBox txtWebPage 
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
            Left            =   2475
            MaxLength       =   50
            TabIndex        =   6
            Top             =   2100
            Width           =   5385
         End
         Begin VB.TextBox txtSocial 
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
            Left            =   2475
            MaxLength       =   50
            TabIndex        =   1
            Top             =   510
            Width           =   5385
         End
         Begin VB.TextBox txtMonN 
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
            Left            =   2475
            MaxLength       =   3
            TabIndex        =   8
            Top             =   2760
            Width           =   915
         End
         Begin VB.TextBox txtMonedaN 
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
            Left            =   3450
            MaxLength       =   30
            TabIndex        =   9
            Top             =   2760
            Width           =   2355
         End
         Begin VB.TextBox txtComercial 
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
            Left            =   2475
            MaxLength       =   50
            TabIndex        =   0
            Top             =   200
            Width           =   5385
         End
         Begin VB.TextBox txtPiePreCuenta 
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
            Left            =   2475
            MaxLength       =   255
            MultiLine       =   -1  'True
            TabIndex        =   14
            Top             =   3735
            Width           =   5385
         End
         Begin VB.TextBox txtElimina 
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
            Left            =   2475
            MaxLength       =   15
            TabIndex        =   16
            Top             =   4380
            Width           =   5385
         End
         Begin VB.TextBox txtContribuyenteEspecial 
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
            Left            =   2475
            MaxLength       =   10
            TabIndex        =   12
            Top             =   3405
            Width           =   915
         End
         Begin VB.TextBox txtDireccion2 
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
            Left            =   2475
            MaxLength       =   150
            TabIndex        =   3
            Top             =   1160
            Width           =   5385
         End
         Begin MSComCtl2.DTPicker dtpContribuyenteEspecial 
            Height          =   315
            Left            =   3450
            TabIndex        =   13
            Top             =   3390
            Width           =   2355
            _ExtentX        =   4154
            _ExtentY        =   556
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
            Format          =   87293953
            UpDown          =   -1  'True
            CurrentDate     =   2.20833333333333
         End
         Begin VB.Label Label 
            Alignment       =   1  'Right Justify
            Caption         =   "Fax  :"
            Height          =   210
            Index           =   51
            Left            =   4920
            TabIndex        =   114
            Top             =   1500
            Width           =   645
         End
         Begin VB.Label Label 
            Alignment       =   1  'Right Justify
            Caption         =   "Texto para Agentes de Retención :"
            Height          =   450
            Index           =   49
            Left            =   195
            TabIndex        =   113
            Top             =   5000
            Width           =   2205
         End
         Begin VB.Label Label 
            Alignment       =   1  'Right Justify
            Caption         =   "Texto en Pie Facturación :"
            Height          =   210
            Index           =   48
            Left            =   195
            TabIndex        =   108
            Top             =   4695
            Width           =   2205
         End
         Begin VB.Label Label 
            Alignment       =   1  'Right Justify
            Caption         =   "Texto en Pie del Documento :"
            Height          =   210
            Index           =   10
            Left            =   75
            TabIndex        =   63
            Top             =   4095
            Width           =   2325
         End
         Begin VB.Label Label 
            Alignment       =   1  'Right Justify
            Caption         =   "Moneda Extranjera :"
            Height          =   210
            Index           =   9
            Left            =   75
            TabIndex        =   62
            Top             =   3120
            Width           =   2325
         End
         Begin VB.Label Label 
            Alignment       =   1  'Right Justify
            Caption         =   "WebPage :"
            Height          =   210
            Index           =   6
            Left            =   75
            TabIndex        =   61
            Top             =   2145
            Width           =   2325
         End
         Begin VB.Label Label 
            Alignment       =   1  'Right Justify
            Caption         =   "E-Mail :"
            Height          =   210
            Index           =   5
            Left            =   75
            TabIndex        =   60
            Top             =   1815
            Width           =   2325
         End
         Begin VB.Label Label 
            Alignment       =   1  'Right Justify
            Caption         =   "Teléfonos  :"
            Height          =   210
            Index           =   4
            Left            =   75
            TabIndex        =   59
            Top             =   1485
            Width           =   2325
         End
         Begin VB.Label Label 
            Alignment       =   1  'Right Justify
            Caption         =   "Razón Social (Bussines Name) :"
            Height          =   210
            Index           =   2
            Left            =   75
            TabIndex        =   58
            Top             =   560
            Width           =   2325
         End
         Begin VB.Label Label 
            Alignment       =   1  'Right Justify
            Caption         =   "Moneda Nacional :"
            Height          =   210
            Index           =   8
            Left            =   75
            TabIndex        =   57
            Top             =   2790
            Width           =   2325
         End
         Begin VB.Label Label 
            Alignment       =   1  'Right Justify
            Caption         =   "Id. Tributario (Federal Id.) :"
            Height          =   210
            Index           =   7
            Left            =   75
            TabIndex        =   56
            Top             =   2460
            Width           =   2325
         End
         Begin VB.Label Label 
            Alignment       =   1  'Right Justify
            Caption         =   "Dirección :"
            Height          =   210
            Index           =   3
            Left            =   75
            TabIndex        =   55
            Top             =   870
            Width           =   2325
         End
         Begin VB.Label Label 
            Alignment       =   1  'Right Justify
            Caption         =   "Razón Comercial (Legal Name) :"
            Height          =   210
            Index           =   0
            Left            =   75
            TabIndex        =   54
            Top             =   240
            Width           =   2325
         End
         Begin VB.Label Label 
            Alignment       =   1  'Right Justify
            Caption         =   "Texto en Pie de la Precuenta :"
            Height          =   210
            Index           =   36
            Left            =   75
            TabIndex        =   53
            Top             =   3765
            Width           =   2325
         End
         Begin VB.Label Label 
            Alignment       =   1  'Right Justify
            Caption         =   "Mensaje de Eliminación Pedido :"
            Height          =   210
            Index           =   11
            Left            =   75
            TabIndex        =   52
            Top             =   4410
            Width           =   2325
         End
         Begin VB.Label Label 
            Alignment       =   1  'Right Justify
            Caption         =   "Ley Contribuyente Especial :"
            Height          =   210
            Index           =   39
            Left            =   75
            TabIndex        =   51
            Top             =   3435
            Width           =   2325
         End
         Begin VB.Label Label 
            Alignment       =   1  'Right Justify
            Height          =   210
            Index           =   30
            Left            =   75
            TabIndex        =   50
            Top             =   1180
            Width           =   2325
         End
      End
   End
   Begin VB.Frame Frame3 
      Caption         =   " Porcentajes por T/P "
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
      Left            =   3360
      TabIndex        =   40
      Top             =   2760
      Visible         =   0   'False
      Width           =   2280
      Begin VB.TextBox txtCanal5 
         Alignment       =   1  'Right Justify
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
         Left            =   1245
         TabIndex        =   37
         Top             =   1185
         Width           =   720
      End
      Begin VB.TextBox txtCanal4 
         Alignment       =   1  'Right Justify
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
         Left            =   1245
         TabIndex        =   36
         Top             =   870
         Width           =   720
      End
      Begin VB.TextBox txtllevar 
         Alignment       =   1  'Right Justify
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
         Left            =   1245
         TabIndex        =   35
         Top             =   540
         Width           =   720
      End
      Begin VB.TextBox txtDelivery 
         Alignment       =   1  'Right Justify
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
         Left            =   1245
         TabIndex        =   34
         Top             =   225
         Width           =   720
      End
      Begin VB.Label Label 
         AutoSize        =   -1  'True
         Caption         =   "Tipo Pedido 5:"
         Height          =   195
         Index           =   44
         Left            =   135
         TabIndex        =   48
         Top             =   1230
         Width           =   1035
      End
      Begin VB.Label Label 
         AutoSize        =   -1  'True
         Caption         =   "Tipo Pedido 4:"
         Height          =   195
         Index           =   43
         Left            =   135
         TabIndex        =   47
         Top             =   915
         Width           =   1035
      End
      Begin VB.Label Label 
         AutoSize        =   -1  'True
         Caption         =   "Tipo Pedido 3:"
         Height          =   195
         Index           =   19
         Left            =   135
         TabIndex        =   46
         Top             =   585
         Width           =   1035
      End
      Begin VB.Label Label 
         AutoSize        =   -1  'True
         Caption         =   "%"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   18
         Left            =   2025
         TabIndex        =   45
         Top             =   1200
         Width           =   210
      End
      Begin VB.Label Label 
         AutoSize        =   -1  'True
         Caption         =   "%"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   16
         Left            =   2025
         TabIndex        =   44
         Top             =   885
         Width           =   210
      End
      Begin VB.Label Label 
         AutoSize        =   -1  'True
         Caption         =   "%"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   13
         Left            =   2025
         TabIndex        =   43
         Top             =   240
         Width           =   210
      End
      Begin VB.Label Label 
         AutoSize        =   -1  'True
         Caption         =   "%"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   14
         Left            =   2025
         TabIndex        =   42
         Top             =   570
         Width           =   210
      End
      Begin VB.Label Label 
         AutoSize        =   -1  'True
         Caption         =   "Tipo Pedido 2:"
         Height          =   195
         Index           =   20
         Left            =   135
         TabIndex        =   41
         Top             =   270
         Width           =   1035
      End
   End
   Begin VB.CommandButton cmdOpcion 
      Caption         =   "Cancelar"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Index           =   1
      Left            =   8040
      Picture         =   "frmParametro.frx":07C4
      Style           =   1  'Graphical
      TabIndex        =   39
      Top             =   8160
      Width           =   1170
   End
   Begin VB.CommandButton cmdOpcion 
      Caption         =   "Grabar"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Index           =   0
      Left            =   9300
      Picture         =   "frmParametro.frx":08C6
      Style           =   1  'Graphical
      TabIndex        =   38
      Top             =   8160
      Width           =   1170
   End
End
Attribute VB_Name = "frmParametro"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Dim RsParametro As Recordset
Dim RsProducto As Recordset
Dim lImprimeCodigoBarras As Boolean

'KDS
Private Sub btnKDS_Click()
    frmKDSConfiguracion.Show vbModal
End Sub
Private Sub chkActivaAnticipo_Click()
    If Me.chkActivaAnticipo.value = 1 Then
        MsgBox "Tener Presente:" + vbNewLine + "Activar Anticipos puede causar problemas con las interfaces contables, Favor de Contactarse con Infhotel Servicios Informaticos S.A.C!!!"
    End If
End Sub

Private Sub chkAgradecimiento_Click()
 If chkAgradecimiento.value = 1 Then
        Me.txtAgradecimiento.Enabled = True
 Else
    txtAgradecimiento.Enabled = True
 End If
End Sub

Private Sub chkConfirmacion_Click()
 If chkConfirmacion.value = 1 Then
        txtConfirmacion.Enabled = True
 Else
    txtConfirmacion.Enabled = True
 End If
End Sub

Private Sub chkDigital_Click()
If chkDigital.value = 1 Then
    chkSecugen.value = 0
End If
End Sub

Private Sub chkFEGesa_Click()
    If Me.chkFEGesa.value = 1 Then
        Me.chkfeTCI.value = 0
        chkFEOfisis.value = 0
        'chkInNC.value = 0
        chkFEpape.value = 0
        chkFESpring.value = 0
        chkFECarbajal.value = 0
        Me.chkFEBiz.value = 0
        chkFacturacionE.value = 1
        Me.fraPaCarvajal.Visible = False
        Me.chkFEubl21.Enabled = False
        Me.chkFEubl21.value = 0
        chkFEGood.value = 0
    End If

End Sub

Private Sub chkFEGood_Click()
   If Me.chkFEGood.value = 1 Then
        Me.chkfeTCI.value = 0
        chkFEOfisis.value = 0
        'chkInNC.value = 0
        chkFEpape.value = 0
        chkFESpring.value = 0
        chkFECarbajal.value = 0
        Me.chkFEBiz.value = 0
        chkFacturacionE.value = 1
        Me.fraPaCarvajal.Visible = True
        Me.chkFEubl21.Enabled = False
        Me.chkFEubl21.value = 0
        chkFEGesa.value = 0
    Else
        Me.fraPaCarvajal.Visible = False
    End If
End Sub
Private Sub chkFEBiz_Click()
   If Me.chkFEBiz.value = 1 Then
        Me.chkfeTCI.value = 0
        Me.chkFEGood.value = 0
        chkFEOfisis.value = 0
        'chkInNC.value = 0
        chkFEpape.value = 0
        chkFESpring.value = 0
        chkFECarbajal.value = 0
        chkFacturacionE.value = 1
        Me.fraPaCarvajal.Visible = True
        Me.chkFEubl21.Enabled = False
        Me.chkFEubl21.value = 0
        chkFEGesa.value = 0
    Else
        Me.fraPaCarvajal.Visible = False
    End If
End Sub

Private Sub chkFECarbajal_Click()
    If chkFECarbajal.value = 1 Then
        chkFEOfisis.value = 0
        Me.chkFEGood.value = 0
        'chkInNC.value = 0
        chkFEpape.value = 0
        chkFESpring.value = 0
        chkfeTCI.value = 0
        Me.chkFEBiz.value = 0
        chkFacturacionE.value = 1
        Me.fraPaCarvajal.Visible = True
        Me.chkFEubl21.Enabled = False
        Me.chkFEubl21.value = 0
        chkFEGesa.value = 0
    Else
        Me.fraPaCarvajal.Visible = False
    End If
End Sub
Private Sub chkFETCI_Click()
    If chkfeTCI.value = 1 Then
        chkFEOfisis.value = 0
        Me.chkFEGood.value = 0
        'chkInNC.value = 0
        chkFEpape.value = 0
        chkFESpring.value = 0
        chkFECarbajal.value = 0
        Me.chkFEBiz.value = 0
        Me.fraPaCarvajal.Visible = False
        Me.chkFEubl21.Enabled = False
        Me.chkFEubl21.value = 0
        chkFEGesa.value = 0
    End If
End Sub
Private Sub chkFEOfisis_Click()
    If chkFEOfisis.value = 1 Then
        chkFESpring.value = 0
        Me.chkFEGood.value = 0
        chkFEpape.value = 0
        chkFECarbajal.value = 0
        chkfeTCI.value = 0
        Me.chkFEBiz.value = 0
        chkFacturacionE.value = 1
        Me.chkFEubl21.Enabled = False
        Me.chkFEubl21.value = 0
        chkFEGesa.value = 0
    End If
End Sub

Private Sub chkFEpape_Click()
    If chkFEpape.value = 1 Then
        chkFESpring.value = 0
        Me.chkFEGood.value = 0
        chkFEOfisis.value = 0
        chkfeTCI.value = 0
        chkFEGesa.value = 0
        'chkInNC.value = 0
        chkFECarbajal.value = 0
        Me.chkFEBiz.value = 0
        If optOpcion(0).value = True Then
            optOpcion(0).value = False
            If optOpcion(1).value = False Then
                If optOpcion(2).value = False Then
                    optOpcion(1).value = True
                End If
            Else
            
            End If
        End If
        optOpcion(0).Visible = False
        chkFacturacionE.value = 1
        Me.chkFEubl21.Enabled = True
    Else
        optOpcion(0).Visible = True
        Me.chkFEubl21.Enabled = False
        Me.chkFEubl21.value = 0
    End If
End Sub

Private Sub chkFESpring_Click()
    If chkFESpring.value = 1 Then
        chkFEOfisis.value = 0
        Me.chkFEGood.value = 0
        'chkInNC.value = 0
        chkFEpape.value = 0
        chkFECarbajal.value = 0
        chkfeTCI.value = 0
        Me.chkFEBiz.value = 0
        chkFacturacionE.value = 1
        Me.chkFEubl21.Enabled = False
        Me.chkFEubl21.value = 0
        chkFEGesa.value = 0
    End If
End Sub

'KDS
Private Sub chkKDS_Click()
    btnKDS.Enabled = chkKds.value
End Sub

Private Sub chkRecordatorio_Click()
 If chkRecordatorio.value = 1 Then
    Me.txtRecordatorio.Enabled = True
 Else
    txtRecordatorio.Enabled = True
 End If
End Sub

Private Sub ChkSAP_Click()

    If ChkSAP.value = 1 Then
        Frame20.Visible = True
    Else
        Frame20.Visible = False
    End If

End Sub

Private Sub chkSecugen_Click()
If chkSecugen.value = 1 Then
    chkDigital.value = 0
End If
End Sub

Private Sub cmdBusca_Click()
   Dim xCriterio As String
   Isql = "Select tCodigo as Codigo, tDistrito as Descripcion, tProvincia as Provincia, tDepartamento as Departamento from TUBIGEO order by tCodigo asc"
   
   frmBusca.cboCriterio.Enabled = True
   frmBusca.nPredeterm = 1
   Call ConfGrilla(4, frmBusca.grdGrilla, "Codigo", 2, "Codigo", 1200, 0, 0, "", _
                                          "Distrito", 2, "Descripcion", 1500, 0, 0, "", _
                                          "Provincia", 2, "Provincia", 2500, 0, 0, "", _
                                          "Departamento", 2, "Departamento", 3000, 0, 0, "")
   frmBusca.Show vbModal
   If Not wEnter Then
      Exit Sub
   End If
   txtCodigoUbigeo.Text = sCodigo
End Sub

Private Sub cmdBuscaAnticipo_Click()
On Error GoTo fin
 
   'If IIf(chkCover.value, 1, 0) = 1 Then
       Isql = "select codigo, descripcion  from vproducto where lactivo=1 and nprecioventa=1 " '"exec sp_VinculacionSAP '" & sServidorSAp & "','" & sBdSAP & "','" & sCodSap & "','','',2"
     
       Call ConfGrilla(2, frmBusquedaRapida.grdGrilla, "Codigo", 1, "Codigo", 1300, 2, 0, "", _
                                                    "Descripción", 2, "Descripcion", 4800, 0, 0, "")
        frmBusquedaRapida.grdGrilla.Caption = "Solo Productos con Precio Venta = 1 "
       frmBusquedaRapida.Show vbModal
       Sw = True

       If sCodigo <> "" Then
            txtCodigoAnticipo.Text = sCodigo
       End If

   ' End If

Exit Sub
fin:
MsgBox (error)
End Sub

Private Sub cmdbuscarItemCover_Click()
On Error GoTo fin
 
   If IIf(chkCover.value, 1, 0) = 1 Then
        Isql = "select codigo, descripcion  from vproducto where lactivo=1 and nprecioventa=1 " '"exec sp_VinculacionSAP '" & sServidorSAp & "','" & sBdSAP & "','" & sCodSap & "','','',2"
     
       Call ConfGrilla(2, frmBusquedaRapida.grdGrilla, "Codigo", 1, "Codigo", 1300, 2, 0, "", _
                                                    "Descripción", 2, "Descripcion", 4800, 0, 0, "")
        frmBusquedaRapida.grdGrilla.Caption = "Solo Productos con Precio Venta = 1 "
       frmBusquedaRapida.Show vbModal
       Sw = True

       If sCodigo <> "" Then
            txtCodigoItemCover.Text = sCodigo
       Else
            txtCodigoItemCover.Text = ""
       End If

    End If

Exit Sub
fin:
MsgBox (error)
End Sub

Private Sub cmdOpcion_Click(Index As Integer)
  
    If Index = 0 Then
    '--- Sap ------
     If ChkSAP Then
         If TxtCodAlmcSAP.Text = "" Then MsgBox "Ingrese el codigo de tienda.", vbExclamation, sMensaje: TxtCodAlmcSAP.SetFocus: Exit Sub
    End If
    
    Dim xImprimeCodigoBarras As Integer
    If optOpcion(0).value = True Then
         xImprimeCodigoBarras = 1
    Else
         xImprimeCodigoBarras = 0
    End If
    
    If chkFEOfisis.value = 1 And optOpcion(0).value = True Then
        MsgBox "Tipo Impresión (Ofisis): Habilitada unicamente para Hash o QR.", vbExclamation, "Integración Ofisis": optOpcion(1).SetFocus: Exit Sub
    End If
    
    If chkFESpring.value = 1 Then
        If optOpcion(0).value = True Then
            MsgBox "Tipo Impresión (Spring): Habilitada unicamente para Hash y QR.", vbExclamation, "Integración Spring": optOpcion(1).SetFocus: Exit Sub
        End If
    End If
    
    If chkFECarbajal.value = 1 Then
        If optOpcion(0).value = True Or optOpcion(1).value = True Then
            MsgBox "Tipo Impresión (Carvajal): Habilitada unicamente para QR.", vbExclamation, "Integración Carvajal": optOpcion(2).SetFocus: Exit Sub
        End If
    End If
    
    If chkFEOfisis.value = 1 Or chkFESpring.value = 1 Or chkFEpape.value = 1 Or chkFECarbajal.value = 1 Then
        chkFacturacionE.value = 1
    End If
   
      Screen.MousePointer = vbHourglass
      'KDS
      
           sPasa = txtRuc.Text
            'Inserta Movimiento auditoria
            lAuditoria = RegistraMovimientoAuditoria(tModuloSeg, sMDB, "TPARAMETRO", "PARAMETRO", "02", sUsuario, sPasa, "", _
                "tidentificacionTributaria", "Id Tributario", Me.txtRuc.Text, "tRazonSocial", "Razon Social", txtSocial.Text, "tRazonComercial", "Razon Comercial", txtComercial.Text, _
                "tDireccion", "Direccion", txtDireccion.Text, "tTelefono", "Telefono", txtTelefono.Text, "tEmail", "Email", txtEmail.Text, "tWebPage", "Pagina Web", txtWebPage.Text, "tDireccion2", "Direccion2", txtDireccion2.Text, _
                "tMonN", "Simb Moneda Nacional", txtMonN.Text, "tMonedaN", "Desc Moneda Nacional", txtMonedaN.Text, "tMonE", "Simb Moneda Extranjera", txtMonE.Text, "tMonedaE", "Desc Moneda Extranjera", txtMonedaE.Text, _
                "tPiePrecuenta", "Pie Precuenta", txtPiePreCuenta.Text, "tPie", "Pie Documentos", txtPie.Text, "tElimina", "Glosa de Eliminacion", txtElimina.Text, _
                "tImpuesto1", "Descripcion Impuesto 1", txtDImp1.Text, "Impuesto1", "Porcentaje Impuesto 1", Val(txtIImp1.Text), "tImpuesto2", "Descripcion Impuesto 2", txtDImp2.Text, "Impuesto2", "Porcentaje Impuesto 2", Val(txtIImp2.Text), "tImpuesto3", "Descripcion Impuesto 3", txtDImp3.Text, "Impuesto3", "Porcentaje Impuesto 3", Val(txtIImp3.Text), _
                "nDelivery", "Porc Recargo por Tipo de Pedido 2", Val(txtDelivery.Text), "nLlevar", "Porc Recargo por Tipo de Pedido 3", Val(txtllevar.Text), "nCanal4", "Porc Recargo por Tipo de Pedido 4", Val(txtCanal4.Text), "nCanal5", "Porc Recargo por Tipo de Pedido 5", Val(txtCanal5.Text), _
                "nCorrelativo", "Correlativo de Pedidos", Val(txtCorrelativo.Text), "nItem", "Impresion Factura Manual Items", txtItem.Text, "nCabecera", "Impresion Factura Manual Cabecera", txtCabecera.Text, "nDetalle", "Impresion Factura Manual Detalle", txtDetalle.Text, "nItemGuia", "Impresion Guia Items", txtItemGuia.Text, "nCabeceraGuia", "Impresion Guia Cabecera", txtCabeceraGuia.Text, "nDetalleGuia", "Impresion Guia Detalle", txtDetalleGuia.Text, _
                "nDias", "Dias en Grilla", txtDia.Text, "nDiasDelivery", "Dias Busqueda Delivery", txtDiaDelivery.Text, "nTiempoMinutoCD", "Tiempo Entrega Delivery", txtTiempoDelivery.Text, "nAsignacionMotorizado", "Monto Maximo Motorizado", Val(txtAsignacionMotorizado.Text), _
                "nLongitud", "Longitud Identificador Tributario", Val(txtLongitud.Text), "lLongitud", "Condicion de Longitud", IIf(opcLongitud(0).value, "Verdadero", "Falso"), "tClub", "Nombre de Punto Club", txtClub.Text, "nPunto", "Valor de Punto Club", Val(txtPunto.Text), _
                "lDiaContableAutomatico", "Flag Dia Contable Automatico", IIf(optDCAutomatico.value, "Verdadero", "Falso"), "lDiaContableManual", "Flag Dia Contable Manual", IIf(optDCManual.value, "Verdadero", "Falso"), "thoracierrediacontable", "Dia Contable Hora Cierre Automatico", Format(Me.dtpHoraDC.value, "HH:mm"), "lImprimeDiacontable", "Dia Contable Impresion Documentos", IIf(Me.chkImprimeDiaContable.value, "Verdadero", "Falso"), _
                "lKds", "Flag Kds", IIf(chkKds.value, "Verdadero", "Falso"), "lCierre", "Flag Cierre a Ciegas", IIf(Me.chkCierre.value, "Verdadero", "Falso"), "lPrinter", "Flag Kitchen Printer", IIf(Me.chkPrinter.value, "Verdadero", "Falso"), "lAlmacen", "Flag Enlace Almacen", IIf(chkAlmacen.value, "Verdadero", "Falso"), "lInfhotel", "Flag Enlace Infhotel", IIf(chkInfhotel.value, "Verdadero", "Falso"), _
                "lequivalencia", "Flag Muestra Equivalencia", IIf(ChkEquivalencia.value, "Verdadero", "Falso"), "lComboGeneral", "Flag Listado General Combos", IIf(chkComboGeneral.value, "Verdadero", "Falso"), "lMultiLocal", "Flag Multi Local", IIf(chkMultiLocal.value, "Verdadero", "Falso"), "lClub", "Flag Club", IIf(chkClub.value, "Verdadero", "Falso"), "tContribuyenteEspecial", "Contribuyente Especial", txtContribuyenteEspecial.Text, "fContribuyenteEspecial", "Fecha Contribuyente Especial", dtpContribuyenteEspecial.value, "lMobileUnidadNegocio", "Mobile Filtro por Unidad Negocio", IIf(Me.chkMUnidadNegocio.value, "Verdadero", "Falso"), "lMobilePasswordCCaja", "Mobile Contrasenia Cambio Caja", IIf(Me.chkMCCaja.value, "Verdadero", "Falso"), "lActivaConsultaDescargo", "Consulta Descargo al Cierre Turno", IIf(Me.chkConsultaDescargo.value, "Verdadero", "Falso"), _
                "nCabeceraV", "Impresion Factura Variable Cabecera", txtCabeceraV.Text, "nItemV", "Impresion Factura Variable Items", txtItemV.Text, "nPieV", "Impresion Factura Variable Pie", txtPieV.Text, "lFacturacionE", "Flag Facturacion Electronica", IIf(Me.chkFacturacionE.value, "Verdadero", "Falso"), "lControlUsuario", "Flag Control Usuario", IIf(Me.chkControlUsuario.value, "Verdadero", "Falso"), "lHoraEntregaDelivery", "Flag Hora Entrega", IIf(Me.chkHoraEntrega.value, "Verdadero", "Falso"), _
                "lHuellaDigital", "Flag Digital Persona", IIf(Me.chkDigital.value, "Verdadero", "Falso"), "lHuellaSecugen", "Flag Secugen", IIf(Me.chkSecugen.value, "Verdadero", "Falso"), "tAgenteRetencion", "Texto Agente Retencion", txtRetencion.Text, "lCheffFiltroSalon", "Cheff-Flag Filtro Salon", IIf(Me.chkCheffFiltroSalon.value, "Verdadero", "Falso"), "lCheffFiltroSubGrupo", "Cheff-Flag Filtro Sub-Grupo", IIf(Me.chkCheffFiltroSubGrupo.value, "Verdadero", "Falso"), "lFESpring", "Cheff-Flag Facturación Electronica Spring", IIf(Me.chkFESpring.value, "Verdadero", "Falso"), "lFECarbajal", "Cheff-Flag Facturación Electronica Carbajal", IIf(Me.chkFECarbajal.value, "Verdadero", "Falso"), "lFETCI", "Facturación Electronica TCI", IIf(Me.chkfeTCI.value, "Verdadero", "Falso"), _
                "lDesactivaNCFP", "Desactiva NC forma de Pago", IIf(Me.chkDesNCPG.value, "Verdadero", "Falso"), "lFEBiz", "Facturacion Electronica Bizlinks", IIf(Me.chkFEBiz.value, "Verdadero", "Falso"), "tCodAnticipo", "Codigo de Producto Anticipo", Me.txtCodigoAnticipo.Text, "lActivaAnticipo", "Activa Anticipo", IIf(Me.chkActivaAnticipo.value, "Verdadero", "Falso"), "lFEGood", "Activa FE Good Hope", IIf(Me.chkFEGood.value, "Verdadero", "Falso"), "tMaxMotorizado", "Maximo de Motorizado", Me.txtMaxMotorizado.Text, "lStockDescargo", "Valida Stock en Descargo", IIf(Me.chkValidaStock.value, "Verdadero", "Falso"), "lFEubl21", "Activa UBL 2.1 FE paperlees", IIf(Me.chkFEubl21.value, "Verdadero", "Falso"), "lBloqInafecto", "Bloquear inafecto", IIf(Me.chkBloqInafecto.value, "Verdadero", "Falso"), _
                "lEstupendoFE", "Facturacion Estupendo Ecuador", IIf(Me.chkFEEstupendo.value, "Verdadero", "Falso"), "LFEGesa", "Facturacion electronica GESA", IIf(Me.chkFEGesa.value, "Verdadero", "Falso"))
                
                If lAuditoria = False Then
                    Screen.MousePointer = vbDefault
                        Exit Sub
                End If
                
                'La Funcion RegistraMovimientoAuditoria devuelve true si se ejecuto correctamente.
      Isql = "Update TPARAMETRO Set " & _
              "tRazonSocial = '" & txtSocial.Text & "',  tRazonComercial ='" & txtComercial.Text & "',  tDireccion ='" & txtDireccion.Text & "', tDireccion2 ='" & txtDireccion2.Text & "', " & _
              "tIdentificacionTributaria ='" & txtRuc.Text & "',  tMonedaN ='" & txtMonedaN.Text & "',  lActivaConsultaDescargo =" & IIf(chkConsultaDescargo.value, 1, 0) & " ,  " & _
              "tMonN ='" & txtMonN.Text & "',  tMonedaE ='" & txtMonedaE.Text & "', tMonE ='" & txtMonE.Text & "', " & _
              "tTelefono ='" & txtTelefono.Text & "', tEmail ='" & txtEmail.Text & "', tWebPage ='" & txtWebPage.Text & "', " & _
              "tPie ='" & txtPie.Text & "',  tPieprecuenta ='" & txtPiePreCuenta.Text & "', " & _
              "lemailconfirmacion=" & IIf(Me.chkConfirmacion.value, 1, 0) & ", lemailrecordatorio=" & IIf(Me.chkRecordatorio.value, 1, 0) & ", lemailagradecimiento=" & IIf(Me.chkAgradecimiento.value, 1, 0) & " , " & _
              "temailconfirmacion='" & Trim(Me.txtConfirmacion.Text) & "', temailrecordatorio='" & Trim(Me.txtRecordatorio.Text) & "', temailagradecimiento='" & Trim(Me.txtAgradecimiento.Text) & "' , " & _
              "tImpuesto1 ='" & txtDImp1.Text & "', tImpuesto2 ='" & txtDImp2.Text & "', tImpuesto3 ='" & txtDImp3.Text & "', " & _
              "nDelivery =" & Val(txtDelivery.Text) & ", nLlevar =" & Val(txtllevar.Text) & ", nCanal4 =" & Val(txtCanal4.Text) & ", nCanal5 =" & Val(txtCanal5.Text) & ", " & _
              "nTiempoToleranciaAnf=" & Val(Me.txtToleranciaReserva.Text) & ",nCorrelativo =" & Val(txtCorrelativo.Text) & ",lMobileUnidadNegocio =" & IIf(Me.chkMUnidadNegocio.value, 1, 0) & ", lMobilePasswordCCaja =" & IIf(Me.chkMCCaja.value, 1, 0) & ",  " & _
              "Impuesto1 =" & Val(txtIImp1.Text) & ", Impuesto2 =" & Val(txtIImp2.Text) & ", Impuesto3 =" & Val(txtIImp3.Text) & ", " & _
              "tElimina ='" & txtElimina.Text & "', nAsignacionMotorizado=" & Val(Me.txtAsignacionMotorizado.Text) & "," & _
              "lCierre =" & IIf(chkCierre.value, 1, 0) & ", lequivalencia =" & IIf(ChkEquivalencia.value, 1, 0) & ", tAgenteRetencion='" & txtRetencion & "', " & _
              "lAlmacen =" & IIf(chkAlmacen.value, 1, 0) & ", nItemGuia =" & txtItemGuia.Text & ", nCabeceraGuia =" & txtCabeceraGuia.Text & ", nDetalleGuia =" & txtDetalleGuia.Text & ", " & _
              "nLongitud = " & Val(txtLongitud.Text) & ", lLongitud=" & IIf(opcLongitud(0).value, 1, 0) & ", tContribuyenteEspecial='" & txtContribuyenteEspecial.Text & "'," & _
              "lPrinter =" & IIf(chkPrinter.value, 1, 0) & ", lInfhotel =" & IIf(chkInfhotel.value, 1, 0) & ",lhuelladigital=" & IIf(chkDigital.value, 1, 0) & ",lhuellasecugen=" & IIf(chkSecugen.value, 1, 0) & ", " & _
              "lMultiLocal =" & IIf(chkMultiLocal.value, 1, 0) & ", lDiaContableAutomatico=" & IIf(optDCAutomatico.value, 1, 0) & " , lDiaContableManual=" & IIf(optDCManual.value, 1, 0) & " ,thoracierrediacontable='" & Format(Me.dtpHoraDC.value, "HH:mm") & "', " & _
              "nItem =" & txtItem.Text & ", nCabecera =" & txtCabecera.Text & ", nDetalle =" & txtDetalle.Text & ", tPassword ='', tClub ='" & txtClub.Text & "', nPunto=" & Val(txtPunto.Text) & ", nDias= " & Val(txtDia.Text) & ", nDiasDelivery= " & Val(txtDiaDelivery.Text) & ", lComboGeneral=" & IIf(chkComboGeneral.value, 1, 0) & ", nTiempoMinutoCD =" & txtTiempoDelivery.Text & ", lKDS ='" & chkKds.value & "' , lClub = '" & chkClub.value & "', lImprimeDiacontable='" & chkImprimeDiaContable.value & "', fContribuyenteEspecial= '" & Format(dtpContribuyenteEspecial.value, "yyyy/MM/dd") & "', " & _
              "nCabeceraV =" & txtCabeceraV.Text & ", nItemV =" & txtItemV.Text & ", nPieV =" & txtPieV.Text & ", lFacturacionE='" & chkFacturacionE.value & "', lControlUsuario='" & chkControlUsuario.value & "', lHoraEntregaDelivery='" & chkHoraEntrega.value & "', " & _
              "tCodigoFE='" & txtCodigoFE.Text & "',tPieDocumento1='" & txtPieFE.Text & "',lAmbienteFE='" & chkAmbienteFE.value & "',tFax='" & txtFax.Text & "', lImprimeCodigoBarras = " & xImprimeCodigoBarras & ",lEnvioAutomatico = '" & chkEnvioAutomatico.value & "', lControlEnviosProduccion = '" & chkControlEnviosProduccion.value & "', lActivaTarjeta = '" & chkTarjeta.value & "', lEventos = '" & chkEventos.value & "',lFEOfisis = '" & chkFEOfisis.value & "',tCodigoEmpresa='" & txtCodigoEmpresa.Text & "',tCodigoTienda='" & txtCodigoTienda.Text & "',tCodigoMarca='" & txtCodigoMarca.Text & "'," & _
              "tCodigoUbigeo='" & txtCodigoUbigeo.Text & "', lPagoAntesImpresion='" & chkPagoAntesImpresion.value & "',tCajaMesa247='" & txtCajaMesa247.Text & "',tServidorFE='" & txtServidorFE.Text & "',tBDFE='" & txtBDFE.Text & "',tAdicionMesa247='" & txtAdicionMesa247.Text & "',linteSAP ='" & ChkSAP.value & "',tservidorSAP ='" & txtServidorSAP.Text & "', tBDSAP = '" & TxtBaseSAP.Text & "', tCodAlmcSAP = '" & TxtCodAlmcSAP.Text & "', lActivaCuenCorrienteAut= '" & ChkActCuentaCorriente.value & "' , tCuentaContableCort = '" & txtCuentaContable.Text & "', lVisor8 = '" & chkVisor8.value & "', lVisortactil = '" & chktactil.value & "', lGlosaTransGratuita='" & txtGlosaImpresion.Text & "', lvisortiempo='" & txtvisortiempo.Text & "'," & _
              "lActivaCover=" & IIf(chkCover.value, 1, 0) & " , tMontoMinCover='" & IIf(Trim(txtMontoMinCover.Text) = "", 0, Trim(txtMontoMinCover.Text)) & "', tCodItemCover= '" & IIf(Trim(txtCodigoItemCover.Text) = "", "", Trim(txtCodigoItemCover.Text)) & "', lNcOfisis=" & IIf(chkInNC.value, 1, 0) & ", tRutaFE='" & Trim(txtRutaImgFE.Text) & "', lCodigoQrFE = " & IIf(optOpcion(2).value, 1, 0) & " , lactivaFechaNC= " & IIf(chkNCFecha.value, 1, 0) & ", lCheffFiltroSalon=" & IIf(chkCheffFiltroSalon.value, 1, 0) & ", lCheffFiltroSubGrupo=" & IIf(chkCheffFiltroSubGrupo.value, 1, 0) & ", lFEpape=" & IIf(chkFEpape.value, 1, 0) & ", lAnula=" & IIf(chkAnulacionNC.value, 1, 0) & ", lDesPagoCheque= " & IIf(chkPagoCheque.value, 1, 0) & ", lDesPagoOtro=" & IIf(chkPagoOtra.value, 1, 0) & ", lFESpring=" & IIf(chkFESpring.value, 1, 0) & ", tUsuarioFE='" & txtUsuarioFE.Text & "', tClaveFE='" & txtClaveFE.Text & "', lFECarbajal=" & IIf(chkFECarbajal.value, 1, 0) & "," & "paramCarvajal = '" & Me.txtParamCarv.Text & "' " & _
              ", tCarvajalCorreos='" & Me.txtCarvajalCorreos.Text & "'" & _
              ", lParcialNC = " & IIf(Me.chkNCParcial.value, 1, 0) & ", lNCElimina = " & IIf(Me.chkNCElimina.value, 1, 0) & ", lValidaDNI = " & IIf(Me.chkValidaDNI.value, 1, 0) & ", lVerTCImp = " & IIf(Me.chkTCenImp.value, 1, 0) & ", lFETCI= " & IIf(Me.chkfeTCI.value, 1, 0) & ", lDesactivaNCFP=" & IIf(Me.chkDesNCPG.value, 1, 0) & ", lFEBiz=" & IIf(Me.chkFEBiz.value, 1, 0) & ", tCodAnticipo='" & Trim(Me.txtCodigoAnticipo.Text) & "', lActivaAnticipo=" & IIf(Me.chkActivaAnticipo.value, 1, 0) & ", lFEGood= " & IIf(Me.chkFEGood.value, 1, 0) & ", tMaxMotorizado='" & Me.txtMaxMotorizado.Text & "', lStockDescargo=" & IIf(Me.chkValidaStock.value, 1, 0) & " , lFEubl21=" & IIf(Me.chkFEubl21.value, 1, 0) & " , lBloqInafecto=" & IIf(Me.chkBloqInafecto.value, 1, 0) & " , lEstupendoFE=" & IIf(Me.chkFEEstupendo.value, 1, 0) & " , lFEGesa=" & IIf(Me.chkFEGesa.value, 1, 0)
              'parametro carvajal arriba

      
      Cn.Execute Isql
      
      Cn.Execute "Update TTABLA Set tResumido ='" & Trim(txtMonN.Text) & "' where TTABLA = 'MONEDA' and tCodigo = '01'"
      Cn.Execute "Update TTABLA Set tResumido ='" & Trim(txtMonE.Text) & "' where TTABLA = 'MONEDA' and tCodigo = '02'"

      sRazonSocial = txtSocial.Text
      sRazonComercial = txtComercial.Text
      sDireccion = txtDireccion.Text
      sDireccion2 = txtDireccion2.Text
      sRUC = txtRuc.Text
      sMonN = txtMonN.Text
      sMonedaN = txtMonedaN.Text
      sMonE = txtMonE.Text
      sMonedaE = txtMonedaE.Text
      sImpuesto1 = txtDImp1.Text
      sImpuesto2 = txtDImp2.Text
      sImpuesto3 = txtDImp3.Text
      nPorcentaje1 = Val(txtIImp1.Text)
      nPorcentaje2 = Val(txtIImp2.Text)
      nPorcentaje3 = Val(txtIImp3.Text)
      nDELIVERY = Val(txtDelivery.Text)
      nLlevar = Val(txtllevar.Text)
      sElimina = txtElimina.Text
      lAlmacen = chkAlmacen.value
      lInfhotel = chkInfhotel.value
      lPrinter = chkPrinter.value
      '---SAP----
      lSAP = ChkSAP.value
      sServidorSAp = Me.txtServidorSAP.Text
      sBdSAP = Me.TxtBaseSAP.Text
      sCodSap = Me.TxtCodAlmcSAP.Text
      '-----visor-----
      
      lvisor = chkVisor8.value
      
      '------
      'huella
      
      lActivaConsultaDescargo = chkConsultaDescargo.value
      lHuellaDigitalPersona = chkDigital.value
      lHuellaSecugen = chkSecugen.value
      
      'FACTURACION ELECTRONICA
      lFacturacionE = chkFacturacionE.value
      tCodigoFE = txtCodigoFE.Text
      tPieDocumento1 = txtPieFE.Text
      lAmbienteProduccion = chkAmbienteFE.value
      
      Screen.MousePointer = vbDefault
      MsgBox "Parámetros Actualizados", vbInformation, sMensaje
      Unload Me
   Else
      Unload Me
   End If
End Sub

Private Sub cmdValidar_Click(Index As Integer)
    If txtServidorFE = "" Then MsgBox "Ingrese el Nombre del Servidor...", vbExclamation, sMensaje: txtServidorFE.SetFocus: Exit Sub
    If txtBDFE = "" Then MsgBox "Ingrese el Nombre de la Base de Datos...", vbExclamation, sMensaje: txtBDFE.SetFocus: Exit Sub
    If txtUsuarioFE = "" Then MsgBox "Ingrese el Nombre del Usuario Sql...", vbExclamation, sMensaje: txtUsuarioFE.SetFocus: Exit Sub
    If txtClaveFE = "" Then MsgBox "Ingrese el Password del Usuario Sql...", vbExclamation, sMensaje: txtClaveFE.SetFocus: Exit Sub
    If (validaConexionSistemaExterno(txtServidorFE, txtBDFE, txtUsuarioFE, txtClaveFE)) = False Then
        MsgBox "No se puede establecer conexión con: " & txtServidorFE, vbCritical, sMensaje
    Else
        MsgBox "Prueba de conexón satisfactoria con el Servidor " & txtServidorFE, vbInformation, sMensaje
    End If
End Sub

Private Sub Form_Load()
   Centrar Me
   'parametro carvajal
   Me.fraPaCarvajal.Visible = False
   
   If pais = "000" Then
    Me.FrmFacPeru.Visible = True
   End If
   If pais = "002" Then
    Me.FrmFacEcuador.Visible = True
   End If
   
   Isql = "select * from TPARAMETRO"
   Set RsParametro = Lib.OpenRecordset(Isql, Cn)
      
   txtComercial.Text = IIf(IsNull(RsParametro!tRazonComercial), "", RsParametro!tRazonComercial)
   txtSocial.Text = IIf(IsNull(RsParametro!tRazonSocial), "", RsParametro!tRazonSocial)
   txtDireccion.Text = IIf(IsNull(RsParametro!tDireccion), "", RsParametro!tDireccion)
   txtDireccion2.Text = IIf(IsNull(RsParametro!tDireccion2), "", RsParametro!tDireccion2)
   txtRuc.Text = IIf(IsNull(RsParametro!tIdentificacionTributaria), "", RsParametro!tIdentificacionTributaria)
   txtMonN.Text = IIf(IsNull(RsParametro!tMonN), "", RsParametro!tMonN)
   txtMonedaN.Text = IIf(IsNull(RsParametro!tMonedaN), "", RsParametro!tMonedaN)
   txtMonE.Text = IIf(IsNull(RsParametro!tMonE), "", RsParametro!tMonE)
   txtMonedaE.Text = IIf(IsNull(RsParametro!tMonedaE), "", RsParametro!tMonedaE)
   txtTelefono.Text = IIf(IsNull(RsParametro!tTelefono), "", RsParametro!tTelefono)
   txtEmail.Text = IIf(IsNull(RsParametro!temail), "", RsParametro!temail)
   txtWebPage.Text = IIf(IsNull(RsParametro!tWebPage), "", RsParametro!tWebPage)
   txtPie.Text = IIf(IsNull(RsParametro!tPie), "", RsParametro!tPie)
   txtPiePreCuenta.Text = IIf(IsNull(RsParametro!tPiePreCuenta), "", RsParametro!tPiePreCuenta)
   txtDImp1.Text = IIf(IsNull(RsParametro!tImpuesto1), "", RsParametro!tImpuesto1)
   txtDImp2.Text = IIf(IsNull(RsParametro!tImpuesto2), "", RsParametro!tImpuesto2)
   txtDImp3.Text = IIf(IsNull(RsParametro!tImpuesto3), "", RsParametro!tImpuesto3)
   txtIImp1.Text = IIf(IsNull(RsParametro!IMPUESTO1), 0, RsParametro!IMPUESTO1)
   txtIImp2.Text = IIf(IsNull(RsParametro!IMPUESTO2), 0, RsParametro!IMPUESTO2)
   txtIImp3.Text = IIf(IsNull(RsParametro!IMPUESTO3), 0, RsParametro!IMPUESTO3)
   txtCorrelativo.Text = IIf(IsNull(RsParametro!nCorrelativo), "", RsParametro!nCorrelativo)
   txtDelivery.Text = IIf(IsNull(RsParametro!nDELIVERY), 0, RsParametro!nDELIVERY)
   txtllevar.Text = IIf(IsNull(RsParametro!nLlevar), 0, RsParametro!nLlevar)
   txtCanal4.Text = IIf(IsNull(RsParametro!nCanal4), 0, RsParametro!nCanal4)
   txtCanal5.Text = IIf(IsNull(RsParametro!nCanal5), 0, RsParametro!nCanal5)
   txtElimina.Text = IIf(IsNull(RsParametro!tElimina), "", RsParametro!tElimina)
   txtItem.Text = IIf(IsNull(RsParametro!nItem), 0, RsParametro!nItem)
   chkPrinter.value = IIf(IsNull(RsParametro!lPrinter), 0, IIf(RsParametro!lPrinter = True, 1, 0))
   txtLongitud.Text = IIf(IsNull(RsParametro!nLongitud), 11, RsParametro!nLongitud)
   chkAlmacen.value = IIf(IsNull(RsParametro!lAlmacen), 0, IIf(RsParametro!lAlmacen = True, 1, 0))
   chkInfhotel.value = IIf(IsNull(RsParametro!lInfhotel), 0, IIf(RsParametro!lInfhotel = True, 1, 0))
   chkMultiLocal.value = IIf(IsNull(RsParametro!lmultilocal), 0, IIf(RsParametro!lmultilocal = True, 1, 0))
   chkCierre = IIf(IsNull(RsParametro!lCierre), 0, IIf(RsParametro!lCierre = True, 1, 0))
   txtClub.Text = IIf(IsNull(RsParametro!tClub), "", RsParametro!tClub)
   txtPunto.Text = Format(IIf(IsNull(RsParametro!nPunto), 1, RsParametro!nPunto), "#,##0.00")
   txtDia.Text = Format(IIf(IsNull(RsParametro!nDias), 1, RsParametro!nDias), "#,##0")
   txtTiempoDelivery.Text = Format(IIf(IsNull(RsParametro!nTiempoMinutoCD), 0, RsParametro!nTiempoMinutoCD), "##0")
   txtDiaDelivery.Text = Format(IIf(IsNull(RsParametro!nDiasDelivery), 1, RsParametro!nDiasDelivery), "#,##0")
   'parametro cravajal
   Me.txtParamCarv = IIf(IsNull(RsParametro!paramCarvajal), "", RsParametro!paramCarvajal)
   
   txtFax.Text = IIf(IsNull(RsParametro!tFax), "", RsParametro!tFax)
   ' glosa de impresion transferencia gratuita
   Me.txtGlosaImpresion.Text = IIf(IsNull(RsParametro!lGlosaTransGratuita), "", RsParametro!lGlosaTransGratuita)
   
   '-------------
   Me.txtRetencion.Text = IIf(IsNull(RsParametro!tAgenteRetencion), "", RsParametro!tAgenteRetencion)
   
   txtCabecera.Text = IIf(IsNull(RsParametro!nCabecera), 0, RsParametro!nCabecera)
   txtDetalle.Text = IIf(IsNull(RsParametro!nDetalle), 0, RsParametro!nDetalle)
   txtItemGuia.Text = IIf(IsNull(RsParametro!nItemGuia), 0, RsParametro!nItemGuia)
   txtCabeceraGuia.Text = IIf(IsNull(RsParametro!nCabeceraGuia), 0, RsParametro!nCabeceraGuia)
   txtDetalleGuia.Text = IIf(IsNull(RsParametro!nDetalleGuia), 0, RsParametro!nDetalleGuia)
   ChkEquivalencia.value = IIf(IsNull(RsParametro!lEquivalencia), 0, IIf(RsParametro!lEquivalencia = True, 1, 0))
   chkComboGeneral.value = IIf(IsNull(RsParametro!lComboGeneral), 0, IIf(RsParametro!lComboGeneral = True, 1, 0))
   txtContribuyenteEspecial = IIf(IsNull(RsParametro!tContribuyenteEspecial), 0, (RsParametro!tContribuyenteEspecial))
   dtpContribuyenteEspecial = IIf(IsNull(RsParametro!fContribuyenteEspecial), 2, (RsParametro!fContribuyenteEspecial))
   
   Me.chkMUnidadNegocio.value = IIf(IsNull(RsParametro!lMobileUnidadNegocio), 0, IIf(RsParametro!lMobileUnidadNegocio = True, 1, 0))
   Me.chkMCCaja.value = IIf(IsNull(RsParametro!lMobilePasswordCCaja), 0, IIf(RsParametro!lMobilePasswordCCaja = True, 1, 0))
   
   chkConsultaDescargo.value = IIf(IsNull(RsParametro!lActivaConsultaDescargo), 0, IIf(RsParametro!lActivaConsultaDescargo = True, 1, 0))
   'huella
   chkDigital.value = IIf(IsNull(RsParametro!lHUELLADIGITAL), 0, IIf(RsParametro!lHUELLADIGITAL = True, 1, 0))
   chkSecugen.value = IIf(IsNull(RsParametro!lHuellaSecugen), 0, IIf(RsParametro!lHuellaSecugen = True, 1, 0))
   Me.txtCarvajalCorreos.Text = IIf(IsNull(RsParametro!tCarvajalCorreos), "", RsParametro!tCarvajalCorreos)
   'KDS
   Me.chkKds.value = IIf(IsNull(RsParametro!lKDS), 0, IIf(RsParametro!lKDS = True, 1, 0))
   
   chkValidaDNI.value = IIf(IsNull(RsParametro!lValidaDNI), 0, IIf(RsParametro!lValidaDNI = True, 1, 0))
   Me.chkTCenImp.value = IIf(IsNull(RsParametro!lVerTCImp), 0, IIf(RsParametro!lVerTCImp = True, 1, 0))
   
   Me.chkFEubl21.value = IIf(IsNull(RsParametro!lFEubl21), 0, IIf(RsParametro!lFEubl21 = True, 1, 0))
   
   If IsNull(RsParametro!lLongitud) Or RsParametro!lLongitud = 0 Then
     opcLongitud(0).value = 0
     opcLongitud(1).value = 1
   Else
     opcLongitud(0).value = 1
     opcLongitud(1).value = 0
   End If
   
   Me.optDCAutomatico.value = IIf(IsNull(RsParametro!lDiaContableAutomatico), 0, RsParametro!lDiaContableAutomatico)
   Me.optDCManual.value = IIf(IsNull(RsParametro!lDiaContablemanual), 0, RsParametro!lDiaContablemanual)
   Me.dtpHoraDC.value = IIf(IsNull(RsParametro!tHoraCierreDiaContable), "00:00", RsParametro!tHoraCierreDiaContable)
   chkImprimeDiaContable.value = IIf(IsNull(RsParametro!lImprimeDiaContable), 0, IIf(RsParametro!lImprimeDiaContable = True, 1, 0))
   
   'Club
   Me.chkClub.value = IIf(IsNull(RsParametro!lClub), 0, IIf(RsParametro!lClub = True, 1, 0))
   
   'motorizados
   txtAsignacionMotorizado.Text = Format(IIf(IsNull(RsParametro!nAsignacionMotorizado), 0, RsParametro!nAsignacionMotorizado), "###,##0.00")

   'Formnato Variabe
   txtCabeceraV.Text = IIf(IsNull(RsParametro!nCabeceraV), 0, RsParametro!nCabeceraV)
   txtItemV.Text = IIf(IsNull(RsParametro!nItemV), 0, RsParametro!nItemV)
   txtPieV.Text = IIf(IsNull(RsParametro!nPieV), 0, RsParametro!nPieV)
   
   'FACTURACION ELECTRONICA
   Me.chkFacturacionE.value = IIf(IsNull(RsParametro!lFacturacionE), 0, IIf(RsParametro!lFacturacionE = True, 1, 0))
   Me.txtCodigoFE.Text = IIf(IsNull(RsParametro!tCodigoFE), "000", RsParametro!tCodigoFE)
   Me.txtPieFE.Text = IIf(IsNull(RsParametro!tPieDocumento1), "", RsParametro!tPieDocumento1)
   Me.chkAmbienteFE.value = IIf(IsNull(RsParametro!lAmbienteFE), 0, IIf(RsParametro!lAmbienteFE = True, 1, 0))
   
  'Control Usuarios
   Me.chkControlUsuario.value = IIf(IsNull(RsParametro!lControlUsuario), 0, IIf(RsParametro!lControlUsuario = True, 1, 0))
   Me.chkHoraEntrega.value = IIf(IsNull(RsParametro!lHoraEntregaDelivery), 0, IIf(RsParametro!lHoraEntregaDelivery = True, 1, 0))
   
   'anfitriona
   Me.chkConfirmacion.value = IIf(IsNull(RsParametro!lEmailConfirmacion), 0, IIf(RsParametro!lEmailConfirmacion = True, 1, 0))
   Me.chkRecordatorio.value = IIf(IsNull(RsParametro!lEmailRecordatorio), 0, IIf(RsParametro!lEmailRecordatorio = True, 1, 0))
   Me.chkAgradecimiento.value = IIf(IsNull(RsParametro!lEmailAgradecimiento), 0, IIf(RsParametro!lEmailAgradecimiento = True, 1, 0))
   
   Me.txtConfirmacion.Text = IIf(IsNull(RsParametro!tEmailConfirmacion), "", RsParametro!tEmailConfirmacion)
   Me.txtRecordatorio.Text = IIf(IsNull(RsParametro!tEmailRecordatorio), "", RsParametro!tEmailRecordatorio)
   Me.txtAgradecimiento.Text = IIf(IsNull(RsParametro!tEmailAgradecimiento), "", RsParametro!tEmailAgradecimiento)
   Me.txtToleranciaReserva.Text = Format(IIf(IsNull(RsParametro!nTiempoToleranciaAnf), 0, RsParametro!nTiempoToleranciaAnf), "##0")
    
   lImprimeCodigoBarras = IIf(IsNull(RsParametro!lImprimeCodigoBarras), 0, IIf(RsParametro!lImprimeCodigoBarras = True, 1, 0))
   
   Me.chkEnvioAutomatico.value = IIf(IsNull(RsParametro!lEnvioAutomatico), 0, IIf(RsParametro!lEnvioAutomatico = True, 1, 0))
   
   Me.chkControlEnviosProduccion.value = IIf(IsNull(RsParametro!lControlEnviosProduccion), 0, IIf(RsParametro!lControlEnviosProduccion = True, 1, 0))
   
   
   Me.chkTarjeta.value = IIf(IsNull(RsParametro!lActivaTarjeta), 0, IIf(RsParametro!lActivaTarjeta = True, 1, 0))
   
   Me.chkEventos.value = IIf(IsNull(RsParametro!lEventos), 0, IIf(RsParametro!lEventos = True, 1, 0))
   
   ' --- activacion de cuentas corrientes Automatico
   ChkActCuentaCorriente.value = IIf(IsNull(RsParametro!lActivaCuenCorrienteAut), 0, IIf(RsParametro!lActivaCuenCorrienteAut = True, 1, 0))
   
   Me.chkFEOfisis.value = IIf(IsNull(RsParametro!lFEOfisis), 0, IIf(RsParametro!lFEOfisis = True, 1, 0))
   
   Me.txtCodigoEmpresa.Text = IIf(IsNull(RsParametro!tCodigoEmpresa), "", RsParametro!tCodigoEmpresa)
   Me.txtCodigoTienda.Text = IIf(IsNull(RsParametro!tCodigoTienda), "", RsParametro!tCodigoTienda)
   Me.txtCodigoMarca.Text = IIf(IsNull(RsParametro!tCodigoMarca), "", RsParametro!tCodigoMarca)
   Me.txtCodigoUbigeo.Text = IIf(IsNull(RsParametro!tCodigoUbigeo), "", RsParametro!tCodigoUbigeo)
   Me.txtCuentaContable.Text = IIf(IsNull(RsParametro!tCuentaContableCort), "", RsParametro!tCuentaContableCort)
      
   Me.chkPagoAntesImpresion.value = IIf(IsNull(RsParametro!lPagoAntesImpresion), 0, IIf(RsParametro!lPagoAntesImpresion = True, 1, 0))
   
   Me.txtCajaMesa247.Text = IIf(IsNull(RsParametro!tCajaMesa247), "", RsParametro!tCajaMesa247)
   Me.txtAdicionMesa247.Text = IIf(IsNull(RsParametro!tAdicionMesa247), "", RsParametro!tAdicionMesa247)
   Me.txtServidorFE.Text = IIf(IsNull(RsParametro!tServidorFE), "", RsParametro!tServidorFE)
   Me.txtBDFE.Text = IIf(IsNull(RsParametro!tBDFE), "", RsParametro!tBDFE)
   '--- SAP
   Me.ChkSAP.value = IIf(IsNull(RsParametro!lInteSAP), 0, IIf(RsParametro!lInteSAP = True, 1, 0))
   ChkSAP_Click
    If ChkSAP.value = 1 Then
     'Me.txtServidorSAP.Text = IIf(IsNull(RsParametro!tservidorSAP), "", RsParametro!tservidorSAP)
     'Me.TxtBaseSAP.Text = IIf(IsNull(RsParametro!tBDSAP), "", RsParametro!tBDSAP)
      Me.TxtCodAlmcSAP.Text = IIf(IsNull(RsParametro!tCodAlmcSAP), "", RsParametro!tCodAlmcSAP)
    Else
     'Me.txtServidorSAP.Text = "" 'IIf(IsNull(RsParametro!tservidorSAP), "", RsParametro!tservidorSAP)
     'Me.TxtBaseSAP.Text = "" ' IIf(IsNull(RsParametro!tBDSAP), "", RsParametro!tBDSAP)
     Me.TxtCodAlmcSAP.Text = "" 'IIf(IsNull(RsParametro!tCodAlmcSAP), "", RsParametro!tCodAlmcSAP)
    End If
    '----visor-----

    Me.chkVisor8.value = IIf(IsNull(RsParametro!lvisor8), 0, IIf(RsParametro!lvisor8 = True, 1, 0))
    Me.chktactil.value = IIf(IsNull(RsParametro!lvisortactil), 0, IIf(RsParametro!lvisortactil = True, 1, 0))
    Me.txtvisortiempo.Text = IIf(IsNull(RsParametro!lvisortiempo), "", RsParametro!lvisortiempo)
   '-------
   '---- bar - cover ecuador----'
   Me.chkCover.value = IIf(IsNull(RsParametro!lActivaCover), 0, IIf(RsParametro!lActivaCover = True, 1, 0))
   Me.txtMontoMinCover.Text = IIf(IsNull(RsParametro!tMontoMinCover), "", RsParametro!tMontoMinCover)
   Me.txtCodigoItemCover.Text = IIf(IsNull(RsParametro!tCodItemCover), "", RsParametro!tCodItemCover)
   
   Me.txtRutaImgFE.Text = IIf(IsNull(RsParametro!tRutaFE), "", RsParametro!tRutaFE)
   '------------------------
   ' ofisis
   Me.chkInNC.value = IIf(IsNull(RsParametro!lNcOfisis), 0, IIf(RsParametro!lNcOfisis = True, 1, 0))
   
   ' notas de credito
   Me.chkNCFecha.value = IIf(IsNull(RsParametro!lactivaFechaNC), 0, IIf(RsParametro!lactivaFechaNC = True, 1, 0))
   Me.chkNCParcial.value = IIf(IsNull(RsParametro!lParcialNC), 0, IIf(RsParametro!lParcialNC = True, 1, 0))
   Me.chkNCElimina.value = IIf(IsNull(RsParametro!lNCElimina), 0, IIf(RsParametro!lNCElimina = True, 1, 0))
   
   'cheff control
   Me.chkCheffFiltroSalon.value = IIf(IsNull(RsParametro!lCheffFiltroSalon), 0, IIf(RsParametro!lCheffFiltroSalon = True, 1, 0))
   Me.chkCheffFiltroSubGrupo.value = IIf(IsNull(RsParametro!lCheffFiltroSubGrupo), 0, IIf(RsParametro!lCheffFiltroSubGrupo = True, 1, 0))
   
   'FE Paperlees
   Me.chkFEpape.value = IIf(IsNull(RsParametro!lFEpape), 0, IIf(RsParametro!lFEpape = True, 1, 0))
   'anulacion de documentos por nota de credito
   Me.chkAnulacionNC.value = IIf(IsNull(RsParametro!lAnula), 0, IIf(RsParametro!lAnula = True, 1, 0))
'   If Me.chkFEpape.value = False Then
'        Me.chkAnulacionNC.Enabled = False
'
'   End If
   '--------------------------------------------
   
   Me.chkPagoCheque.value = IIf(IsNull(RsParametro!lDesPagoCheque), 0, IIf(RsParametro!lDesPagoCheque = True, 1, 0))
   Me.chkPagoOtra.value = IIf(IsNull(RsParametro!lDesPagoOtro), 0, IIf(RsParametro!lDesPagoOtro = True, 1, 0))
   
   'FE Spring
   Me.chkFESpring.value = IIf(IsNull(RsParametro!lFESpring), 0, IIf(RsParametro!lFESpring = True, 1, 0))
   Me.txtUsuarioFE.Text = IIf(IsNull(RsParametro!tUsuarioFE), "", RsParametro!tUsuarioFE)
   Me.txtClaveFE.Text = IIf(IsNull(RsParametro!tClaveFE), "", RsParametro!tClaveFE)
   
   'FE Carbajal
   Me.chkFECarbajal.value = IIf(IsNull(RsParametro!lFECarbajal), 0, IIf(RsParametro!lFECarbajal = True, 1, 0))
   
   'FE TCI
   Me.chkfeTCI.value = IIf(IsNull(RsParametro!lFETCI), 0, IIf(RsParametro!lFETCI = True, 1, 0))
   
   Me.chkDesNCPG.value = IIf(IsNull(RsParametro!lDesactivaNCFP), 0, IIf(RsParametro!lDesactivaNCFP = True, 1, 0))
   Me.chkFEBiz.value = IIf(IsNull(RsParametro!lFEBiz), 0, IIf(RsParametro!lFEBiz = True, 1, 0))
   Me.txtCodigoAnticipo.Text = IIf(IsNull(RsParametro!tCodAnticipo), "", RsParametro!tCodAnticipo)
   Me.chkActivaAnticipo.value = IIf(IsNull(RsParametro!lActivaAnticipo), 0, IIf(RsParametro!lActivaAnticipo = True, 1, 0))
   Me.chkFEGood.value = IIf(IsNull(RsParametro!lFEGood), 0, IIf(RsParametro!lFEGood = True, 1, 0))
   Me.txtMaxMotorizado.Text = IIf(IsNull(RsParametro!tMaxMotorizado), "", RsParametro!tMaxMotorizado)
   Me.chkValidaStock.value = IIf(IsNull(RsParametro!lStockDescargo), 0, IIf(RsParametro!lStockDescargo = True, 1, 0))
   Me.chkBloqInafecto.value = IIf(IsNull(RsParametro!lBloqInafecto), 0, IIf(RsParametro!lBloqInafecto = True, 1, 0))
   Me.chkFEEstupendo.value = IIf(IsNull(RsParametro!lEstupendoFE), 0, IIf(RsParametro!lEstupendoFE = True, 1, 0))
   
   Me.chkFEGesa.value = IIf(IsNull(RsParametro!lFEGesa), 0, IIf(RsParametro!lFEGesa = True, 1, 0))
   
   If lImprimeCodigoBarras Then
        optOpcion(0).value = True
        optOpcion(1).value = False
   Else
        optOpcion(0).value = False
        optOpcion(1).value = True
   End If
   
   If pais = "002" Or pais = "001" Then
        FrameTipoImpresion.Visible = False
   End If
    If IIf(IsNull(RsParametro!lCodigoQrFE), 0, IIf(RsParametro!lCodigoQrFE = True, 1, 0)) = 1 Then
        optOpcion(2).value = True
    Else
        optOpcion(2).value = False
    End If
    
    If Not lAlmacen Then
        Frame35.Enabled = False
    End If
    
    SSTab1.Tab = 0
End Sub

Private Sub Form_Unload(Cancel As Integer)
   RsParametro.Close
   Set RsParametro = Nothing
   Set frmParametro = Nothing
End Sub

Private Sub optDCAutomatico_Click()
        If optDCAutomatico.value = True Then
            Me.dtpHoraDC.Enabled = True
            Me.dtpHoraDC.SetFocus
        Else
            Me.dtpHoraDC.Enabled = False
        End If
End Sub

Private Sub optDCManual_Click()
        If optDCManual.value = True Then
              Me.dtpHoraDC.value = "00:00"
              Me.dtpHoraDC.Enabled = False
        Else
         Me.dtpHoraDC.Enabled = True
          Me.dtpHoraDC.SetFocus
            
        End If
End Sub

Private Sub txtChk_LostFocus()
  If Val(txtChk.Text) < 0 Or Val(txtChk.Text) > 60 Then
     MsgBox "Rango Erroneo", vbExclamation, sMensaje
     txtChk.SetFocus
  End If
End Sub


Private Sub txtAsignacionMotorizado_KeyPress(KeyAscii As Integer)
   TabNext KeyAscii
   Numerico KeyAscii, txtAsignacionMotorizado
End Sub

Private Sub txtMaxMotorizado_Change()
    If Not IsNumeric(Me.txtMaxMotorizado.Text) Then
        Me.txtMaxMotorizado.Text = ""
    End If
    Me.txtMaxMotorizado.Text = Trim(Me.txtMaxMotorizado.Text)
    Me.txtMaxMotorizado.SelStart = Len(Me.txtMaxMotorizado)
End Sub

Private Sub txtPunto_LostFocus()
   If Val(txtPunto.Text) > 1 Then
      txtPunto.Text = Format(txtPunto.Text, "#,##0.00")
   Else
     txtPunto.Text = "1.00"
   End If
End Sub

Private Sub txtToleranciaReserva_KeyPress(KeyAscii As Integer)
    TabNext KeyAscii
   Numerico KeyAscii, txtToleranciaReserva
End Sub


