VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Begin VB.Form frmRepPedido 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Reporte de Pedidos"
   ClientHeight    =   6780
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7620
   Icon            =   "frmRepPedido.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6780
   ScaleWidth      =   7620
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
      Left            =   4695
      Picture         =   "frmRepPedido.frx":000C
      Style           =   1  'Graphical
      TabIndex        =   23
      Top             =   6150
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
      Height          =   6090
      Left            =   30
      TabIndex        =   26
      Top             =   30
      Width           =   7545
      Begin VB.CheckBox chkOrigenVenta 
         Caption         =   "Todos los Origenes de Venta"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Left            =   5145
         TabIndex        =   43
         Top             =   1920
         Value           =   1  'Checked
         Width           =   2175
      End
      Begin VB.CheckBox chkFranjaHoraria 
         Caption         =   "Franja Horaria"
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
         Left            =   5130
         TabIndex        =   42
         Top             =   4680
         Width           =   1635
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
         Height          =   765
         Left            =   1890
         TabIndex        =   41
         Top             =   5280
         Width           =   3165
         Begin VB.OptionButton optOpcion 
            Caption         =   "Resumido"
            Height          =   240
            Index           =   0
            Left            =   225
            TabIndex        =   20
            Top             =   480
            Width           =   2700
         End
         Begin VB.OptionButton optOpcion 
            Caption         =   "Detallado"
            Height          =   240
            Index           =   1
            Left            =   240
            TabIndex        =   19
            Top             =   240
            Value           =   -1  'True
            Width           =   2790
         End
      End
      Begin VB.TextBox txtEnvio 
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
         Height          =   315
         Left            =   1905
         MaxLength       =   3
         TabIndex        =   12
         Text            =   "0"
         Top             =   3630
         Width           =   1635
      End
      Begin VB.TextBox txtPrecuenta 
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
         Height          =   315
         Left            =   1905
         MaxLength       =   3
         TabIndex        =   11
         Text            =   "0"
         Top             =   3210
         Width           =   1635
      End
      Begin VB.TextBox txtAdulto 
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
         Height          =   315
         Left            =   1905
         MaxLength       =   3
         TabIndex        =   10
         Text            =   "0"
         Top             =   2820
         Width           =   1635
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
         Left            =   4275
         Picture         =   "frmRepPedido.frx":082E
         Style           =   1  'Graphical
         TabIndex        =   13
         Top             =   4080
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
         Left            =   1905
         Locked          =   -1  'True
         TabIndex        =   25
         Top             =   4110
         Width           =   2265
      End
      Begin VB.CheckBox chkLocal 
         Caption         =   "Todos los Locales"
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
         Left            =   5130
         TabIndex        =   1
         Top             =   345
         Value           =   1  'Checked
         Width           =   1905
      End
      Begin VB.CheckBox chkEstadoPedido 
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
         Left            =   5130
         TabIndex        =   9
         Top             =   2430
         Value           =   1  'Checked
         Width           =   2175
      End
      Begin VB.CheckBox chkTipoPedido 
         Caption         =   "Todos los Pedidos"
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
         Left            =   5130
         TabIndex        =   7
         Top             =   1545
         Value           =   1  'Checked
         Width           =   1905
      End
      Begin VB.CheckBox chkCaja 
         Caption         =   "Todos las Cajas"
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
         Left            =   5130
         TabIndex        =   5
         Top             =   1145
         Value           =   1  'Checked
         Width           =   1875
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
         Left            =   5130
         TabIndex        =   14
         Top             =   4125
         Value           =   1  'Checked
         Width           =   1875
      End
      Begin VB.CheckBox chkSalon 
         Caption         =   "Todos los Salones"
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
         Left            =   5130
         TabIndex        =   3
         Top             =   745
         Value           =   1  'Checked
         Width           =   1905
      End
      Begin MSComCtl2.DTPicker dtpFecFin 
         Height          =   345
         Left            =   1905
         TabIndex        =   17
         Top             =   4950
         Width           =   1635
         _ExtentX        =   2884
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
         Format          =   84738049
         CurrentDate     =   37541.9993055556
      End
      Begin MSComCtl2.DTPicker dtpFecIni 
         Height          =   345
         Left            =   1905
         TabIndex        =   15
         Top             =   4515
         Width           =   1635
         _ExtentX        =   2884
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
         Format          =   84738049
         CurrentDate     =   37539.2083333333
      End
      Begin MSComCtl2.DTPicker dtpHorIni 
         Height          =   375
         Left            =   3810
         TabIndex        =   16
         Top             =   4500
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
         Format          =   84738051
         UpDown          =   -1  'True
         CurrentDate     =   37539
      End
      Begin MSComCtl2.DTPicker dtpHorFin 
         Height          =   375
         Left            =   3810
         TabIndex        =   18
         Top             =   4935
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
         Format          =   84738051
         UpDown          =   -1  'True
         CurrentDate     =   37541.9993055556
      End
      Begin MSDataListLib.DataCombo cboLocal 
         Height          =   315
         Left            =   1905
         TabIndex        =   0
         Top             =   330
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
      Begin MSDataListLib.DataCombo cboSalon 
         Height          =   315
         Left            =   1905
         TabIndex        =   2
         Top             =   735
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
      Begin MSDataListLib.DataCombo cboTipoPedido 
         Height          =   315
         Left            =   1905
         TabIndex        =   6
         Top             =   1530
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
      Begin MSDataListLib.DataCombo cboCaja 
         Height          =   315
         Left            =   1905
         TabIndex        =   4
         Top             =   1125
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
      Begin MSDataListLib.DataCombo cboEstadoPedido 
         Height          =   315
         Left            =   1905
         TabIndex        =   8
         Top             =   2400
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
      Begin MSDataListLib.DataCombo cboOrigenVenta 
         Height          =   315
         Left            =   1905
         TabIndex        =   44
         Top             =   1920
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
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Origenes Venta :"
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
         Index           =   12
         Left            =   360
         TabIndex        =   45
         Top             =   1980
         Width           =   1440
      End
      Begin VB.Label Label 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Envios Producción :"
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
         Index           =   11
         Left            =   105
         TabIndex        =   40
         Top             =   3690
         Width           =   1725
      End
      Begin VB.Label Label 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Cantidad mínima a Filtrar (solo para detallado)"
         ForeColor       =   &H00000000&
         Height          =   195
         Index           =   10
         Left            =   3585
         TabIndex        =   39
         Top             =   3690
         Width           =   3210
      End
      Begin VB.Label Label 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Cantidad mínima a Filtrar"
         ForeColor       =   &H00000000&
         Height          =   195
         Index           =   9
         Left            =   3600
         TabIndex        =   38
         Top             =   3270
         Width           =   1740
      End
      Begin VB.Label Label 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Cantidad mínima a Filtrar "
         ForeColor       =   &H00000000&
         Height          =   195
         Index           =   7
         Left            =   3600
         TabIndex        =   37
         Top             =   2880
         Width           =   1785
      End
      Begin VB.Label Label 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Precuentas :"
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
         Left            =   735
         TabIndex        =   36
         Top             =   3270
         Width           =   1095
      End
      Begin VB.Label Label 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Adultos :"
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
         Left            =   1065
         TabIndex        =   35
         Top             =   2880
         Width           =   765
      End
      Begin VB.Label Label 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Local :"
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
         Index           =   8
         Left            =   1230
         TabIndex        =   34
         Top             =   390
         Width           =   600
      End
      Begin VB.Label Label 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Estado Pedido :"
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
         Left            =   465
         TabIndex        =   33
         Top             =   2475
         Width           =   1365
      End
      Begin VB.Label Label 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Tipo de Pedido :"
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
         Left            =   405
         TabIndex        =   32
         Top             =   1590
         Width           =   1425
      End
      Begin VB.Label Label 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Caja :"
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
         Left            =   1320
         TabIndex        =   31
         Top             =   1185
         Width           =   510
      End
      Begin VB.Label Label3 
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
         Left            =   705
         TabIndex        =   30
         Top             =   5025
         Width           =   1125
      End
      Begin VB.Label Label2 
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
         Left            =   600
         TabIndex        =   29
         Top             =   4590
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
         Left            =   1200
         TabIndex        =   28
         Top             =   4170
         Width           =   630
      End
      Begin VB.Label Label 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Salón :"
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
         Left            =   1215
         TabIndex        =   27
         Top             =   795
         Width           =   615
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
      Left            =   6150
      Picture         =   "frmRepPedido.frx":0930
      Style           =   1  'Graphical
      TabIndex        =   24
      Top             =   6150
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
      Left            =   3240
      Picture         =   "frmRepPedido.frx":0A22
      Style           =   1  'Graphical
      TabIndex        =   22
      Top             =   6150
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
      Left            =   1785
      Picture         =   "frmRepPedido.frx":0F54
      Style           =   1  'Graphical
      TabIndex        =   21
      Top             =   6150
      Width           =   1455
   End
   Begin MSComDlg.CommonDialog cmdSave 
      Left            =   1215
      Top             =   6210
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
End
Attribute VB_Name = "frmRepPedido"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim rsReporte As Recordset
Dim RsLocal As Recordset
Dim RsSalon As Recordset
Dim RsCaja As Recordset
Dim RsTipoPedido As Recordset
Dim RsEstadoPedido As Recordset
Dim Rsturno As Recordset
Dim ReporteD As New dsrPedidosD
Dim ReporteR As New dsrPedidosR
Dim sCriterio As String
Dim sTitulo As String
Dim RsOrigen As Recordset
Dim fInicio As Date
Dim fFinal As Date

Sub LlenaCombos()
    With cboLocal
         Isql = "Select * from vLocal"
         Set RsLocal = Lib.OpenRecordset(Isql, Cn)
         Set .RowSource = RsLocal
         .DataField = "Descripcion"
         .ListField = "Descripcion"
         .BoundColumn = "Codigo"
    End With
    With cboSalon
         Isql = "Select * from vSalon"
         Set RsSalon = Lib.OpenRecordset(Isql, Cn)
         Set .RowSource = RsSalon
         .DataField = "Descripcion"
         .ListField = "Descripcion"
         .BoundColumn = "Codigo"
    End With
    With Me.cboOrigenVenta
         Isql = "Select CodOrigenVenta as codigo, Descripcion   from vOrigenVenta where Activo=1"
         Set RsOrigen = Lib.OpenRecordset(Isql, Cn)
         Set .RowSource = RsOrigen
             .DataField = "Descripcion"
             .ListField = "Descripcion"
             .BoundColumn = "Codigo"
    End With
    With cboTipoPedido
         Isql = "Select * from vTipoPedido"
         Set RsTipoPedido = Lib.OpenRecordset(Isql, Cn)
         Set .RowSource = RsTipoPedido
             .DataField = "Descripcion"
             .ListField = "Descripcion"
             .BoundColumn = "Codigo"
    End With
    With cboEstadoPedido
         Isql = "Select * from vEstadoPedido"
         Set RsEstadoPedido = Lib.OpenRecordset(Isql, Cn)
         Set .RowSource = RsEstadoPedido
             .DataField = "Descripcion"
             .ListField = "Descripcion"
             .BoundColumn = "Codigo"
    End With
    With cboCaja
         Isql = "Select tCaja as codigo, tDescripcion as Descripcion from TCAJA"
         Set RsEstadoPedido = Lib.OpenRecordset(Isql, Cn)
         Set .RowSource = RsEstadoPedido
             .DataField = "Descripcion"
             .ListField = "Descripcion"
             .BoundColumn = "Codigo"
    End With

End Sub

Private Sub cboLocal_Click(Area As Integer)
    cboSalon.Text = ""
    With cboSalon
         Isql = "Select * from vSalon " & IIf(cboLocal.BoundText = "", "", "where tLocal = '" & cboLocal.BoundText & "'") & " order by Descripcion "
         Set RsSalon = Lib.OpenRecordset(Isql, Cn)
         Set .RowSource = RsSalon
    End With
End Sub



Private Sub cboTipoPedido_Change()
    Me.cboOrigenVenta.BoundText = ""
    With cboOrigenVenta
         Isql = "Select CodOrigenVenta as codigo, Descripcion   from vOrigenVenta  " & IIf(cboTipoPedido.BoundText = "", "", "where CodCanalVenta = '" & cboTipoPedido.BoundText & "' and Activo=1") & " order by Descripcion "
         Set RsOrigen = Lib.OpenRecordset(Isql, Cn)
         Set .RowSource = RsOrigen
    End With
End Sub

Private Sub chkOrigenVenta_Click()
    If chkOrigenVenta.value = 1 Then
        Me.cboOrigenVenta.Enabled = False
        Me.cboOrigenVenta.Text = ""
    Else
        Me.cboOrigenVenta.Enabled = True
    End If
End Sub

Private Sub ChkCaja_Click()
   If ChkCaja.value = 1 Then
      cboCaja.Enabled = False
      cboCaja.Text = ""
   Else
      cboCaja.Enabled = True
   End If
End Sub

Private Sub cmdBusca_Click()
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
   sTurno = sCodigo
   txtTurno.Text = sCodigo
End Sub

Private Sub chkTipoPedido_Click()
   If chkTipoPedido.value = 1 Then
      cboTipoPedido.Enabled = False
      cboTipoPedido.Text = ""
   Else
      cboTipoPedido.Enabled = True
   End If
End Sub

Private Sub chkLocal_Click()
   If chkLocal.value = 1 Then
      cboLocal.Enabled = False
      cboLocal.Text = ""
      RsSalon.Filter = adFilterNone
   Else
      cboLocal.Enabled = True
   End If
End Sub

Private Sub chkSalon_Click()
   If chkSalon.value = 1 Then
      cboSalon.Enabled = False
      cboSalon.Text = ""
   Else
      cboSalon.Enabled = True
   End If
End Sub

Private Sub chkEstadoPedido_Click()
   If chkEstadoPedido.value = 1 Then
      cboEstadoPedido.Enabled = False
      cboEstadoPedido.Text = ""
   Else
      cboEstadoPedido.Enabled = True
   End If
End Sub

Private Sub cmdOpcion_Click(Index As Integer)
   If Index = 2 Then
      Unload Me
      Exit Sub
   End If
   
   sCriterio = ""
   sTitulo = ""
   
   If ChkTurno.value = 0 Then
      If sTurno = "" Then
         MsgBox "Debe escoger un Turno", vbCritical, sMensaje
         Exit Sub
      End If
            
      Set Rsturno = Lib.OpenRecordset("select * from MTURNO where tTurno='" & sTurno & "'", Cn)
      If Rsturno.RecordCount > 0 Then
         sTitulo = " Del " & Format(Rsturno!finicial, "dd/MMM/yyyy HH:mm") & " Al " & Format(IIf(IsNull(Rsturno!fFinal), Format(Rsturno!finicial, "dd/mmm/yyyy") & " 23:59", Rsturno!fFinal), "dd/mmm/yyyy HH:mm")
         sCriterio = " MPEDIDO.tTurno ='" & sTurno & "'"
      Else
         MsgBox "Error Fatal en Turnos", vbCritical, sMensaje
      End If
      
   Else
      If dtpFecIni.value + dtpHorIni.value > dtpFecFin.value + dtpHorFin.value Then
         MsgBox "Error en Rango de Fechas", vbCritical, sMensaje
         Exit Sub
         dtpFecFin.SetFocus
      End If
      
            If ChkFranjaHoraria.value = 0 Then
                    sTitulo = " Del " & Format(dtpFecIni.value, "dd/mmm/yyyy") & " " & Format(dtpHorIni.value, "HH:mm") & " Al " & Format(dtpFecFin.value, "dd/mmm/yyyy") & " " & Format(dtpHorFin.value, "HH:mm")
                Else
                    sTitulo = " Del " & Format(dtpFecIni.value, "dd/mmm/yyyy") & " Al " & Format(dtpFecFin.value, "dd/mmm/yyyy") & " Franja Horaria:  " & Format(dtpHorIni.value, "HH:mm") & " - " & Format(dtpHorFin.value, "HH:mm")
            End If
        
    '  sTitulo = " Del " & Format(dtpFecIni.value, "dd/mmm/yyyy") & " " & Format(dtpHorIni.value, "HH:mm") & " Al " & Format(dtpFecFin.value, "dd/mmm/yyyy") & " " & Format(dtpHorFin.value, "HH:mm")
      sCriterio = " MPEDIDO.fRegistro >= '" & Format(dtpFecIni.value, "yyyy/mm/dd") & " " & Format(dtpHorIni.value, "HH:mm") & "' and MPEDIDO.fRegistro <= '" & Format(dtpFecFin.value, "yyyy/mm/dd") & " " & Format(dtpHorFin.value, "HH:mm") & "' "
   End If
   
   If chkLocal.value = 0 Then
      If cboLocal.Text = "" Then
         MsgBox "Debe escoger un Local", vbCritical, sMensaje
         Exit Sub
      End If
      sCriterio = sCriterio & " and vSalon.tLocal ='" & cboLocal.BoundText & "'"
   End If
   
   If chkSalon.value = 0 Then
      If cboSalon.Text = "" Then
         MsgBox "Debe escoger un Salón", vbCritical, sMensaje
         Exit Sub
      End If
      sCriterio = sCriterio & " and MPEDIDO.tSalon ='" & cboSalon.BoundText & "'"
   End If
      
   If chkTipoPedido.value = 0 Then
      If cboTipoPedido.Text = "" Then
         MsgBox "Debe escoger un Tipo de Pedido", vbCritical, sMensaje
         Exit Sub
      End If
      sCriterio = sCriterio & " and MPEDIDO.tTipoPedido ='" & cboTipoPedido.BoundText & "'"
   End If
    If Me.chkOrigenVenta.value = 0 Then
      If Me.cboOrigenVenta.Text = "" Then
         MsgBox "Debe escoger un Origen Venta", vbCritical, sMensaje
         Exit Sub
      End If
    End If
   If chkEstadoPedido.value = 0 Then
      If cboEstadoPedido.Text = "" Then
         MsgBox "Debe escoger un Estado", vbCritical, sMensaje
         Exit Sub
      End If
      sCriterio = sCriterio & " and tEstadoPedido ='" & cboEstadoPedido.BoundText & "'"
   End If
            
   If ChkCaja.value = 0 Then
      If cboCaja.Text = "" Then
         MsgBox "Debe escoger una caja", vbCritical, sMensaje
         Exit Sub
      End If
      sCriterio = sCriterio & " and tCaja='" & cboCaja.BoundText & "'"
   End If
   
   If CDbl(txtAdulto.Text) > 0 Then
      sCriterio = sCriterio & " and nAdulto >=" & CDbl(txtAdulto.Text)
   End If
         
   If CDbl(txtPrecuenta.Text) > 0 Then
      sCriterio = sCriterio & " and nPrecuenta >=" & CDbl(txtPrecuenta.Text)
   End If
         
   If CDbl(txtEnvio.Text) > 0 And optOpcion(1).value = True Then
      sCriterio = sCriterio & " and nEnvio >=" & CDbl(txtEnvio.Text)
   End If
         
   Select Case Index
          Case Is = 0 ' Preview
          
               Genera
               
               If rsReporte.EOF = True Then
               
                  MsgBox "No hay Datos para Mostrar", vbExclamation, "Mensaje del Sistema"
                  Screen.MousePointer = vbDefault
                  Exit Sub
                  
               End If
               If optOpcion(1).value = True Then
                  ReporteD.SelectPrinter Printer.DriverName, Printer.DeviceName, Printer.Port
                  ReporteD.PaperOrientation = crLandscape
               Else
                  ReporteR.SelectPrinter Printer.DriverName, Printer.DeviceName, Printer.Port
                  ReporteR.PaperOrientation = crLandscape
               End If
               frmEmite.CRViewer.ViewReport
               
               frmEmite.Show vbModal
          
          
          Case Is = 1 ' Imprimir
               Genera
               Screen.MousePointer = vbDefault
               If rsReporte.EOF = True Then
                   MsgBox "No hay Datos para Mostrar", vbExclamation, "Mensaje del Sistema"
                   Exit Sub
               End If
               If optOpcion(1).value = True Then
                  ReporteD.SelectPrinter Printer.DriverName, Printer.DeviceName, Printer.Port
                  ReporteD.PaperOrientation = crLandscape
                  ReporteD.PrintOut
               Else
                  ReporteR.SelectPrinter Printer.DriverName, Printer.DeviceName, Printer.Port
                  ReporteR.PaperOrientation = crLandscape
                  ReporteR.PrintOut
               End If
          
          Case Is = 2 ' Salir
               Unload Me
               
          Case Is = 3 ' Exportar
               Genera
               Screen.MousePointer = vbDefault
               If rsReporte.EOF = True Then
                  MsgBox "No hay Datos para Mostrar", vbExclamation, "Mensaje del Sistema"
                  Exit Sub
               End If
               If optOpcion(1).value = True Then
                  ReporteD.ExportOptions.FormatType = 21
                  ReporteD.ExportOptions.DestinationType = 1
                  cmdSave.Filter = "Libro de Microsoft Excel|*.xls"
                  cmdSave.ShowSave
                  If cmdSave.FileName = "" Then
                     Exit Sub
                  End If
                  ReporteD.ExportOptions.DiskFileName = cmdSave.FileName
                  ReporteD.Export False
               Else
                  ReporteR.ExportOptions.FormatType = 21
                  ReporteR.ExportOptions.DestinationType = 1
                  cmdSave.Filter = "Libro de Microsoft Excel|*.xls"
                  cmdSave.ShowSave
                  If cmdSave.FileName = "" Then
                     Exit Sub
                  End If
                  ReporteR.ExportOptions.DiskFileName = cmdSave.FileName
                  ReporteR.Export False
               End If
   End Select

End Sub

Private Sub dtpFecfin_LostFocus()
   If dtpFecIni.value > dtpFecFin.value Then
      MsgBox "Error en Rango de Fechas", vbCritical, sMensaje
      dtpFecFin.SetFocus
   End If
End Sub

Private Sub ChkTurno_Click()
   If ChkTurno.value = 1 Then
      cmdBusca.Enabled = False
      sTurno = ""
      dtpFecIni.Enabled = True
      dtpFecFin.Enabled = True
      dtpHorIni.Enabled = True
      dtpHorFin.Enabled = True
      ChkFranjaHoraria.Enabled = True
      ChkFranjaHoraria.value = 0
   Else
      cmdBusca.Enabled = True
      dtpFecIni.Enabled = False
      dtpFecFin.Enabled = False
      dtpHorIni.Enabled = False
      dtpHorFin.Enabled = False
       ChkFranjaHoraria.value = 0
      ChkFranjaHoraria.Enabled = False
   End If
End Sub

Private Sub Form_Load()
    Centrar Me
    LlenaCombos
    dtpFecIni.value = Date
    dtpFecFin.value = Date
    cmdBusca.Enabled = False
    sTurno = ""
    
    cboLocal.Enabled = False
    cboSalon.Enabled = False
    cboTipoPedido.Enabled = False
    cboEstadoPedido.Enabled = False
    cboCaja.Enabled = False
    
    cboLocal.Text = ""
    cboSalon.Text = ""
    cboCaja.Text = ""
    cboTipoPedido.Text = ""
    cboEstadoPedido.Text = ""
End Sub

Private Sub Form_Unload(Cancel As Integer)
   Set rsReporte = Nothing
   Set Rsturno = Nothing
   Set RsTipoPedido = Nothing
   Set RsEstadoPedido = Nothing
   Set RsCaja = Nothing
   Set RsLocal = Nothing
   Set RsSalon = Nothing
   Set frmRepPedido = Nothing
End Sub

Public Sub Genera()
   Screen.MousePointer = vbHourglass
   Dim oComando As clsComando
   Dim tipopedido As String
    Set oComando = New clsComando
    
    If Not oComando.CreateCmdSp("spRep_Pedido", Cn) Then
       Set oComando = Nothing
       Exit Sub
    End If
    fInicio = Format(dtpFecIni.value, "yyyy/MM/dd") & " " & Format(dtpHorIni.value, "HH:mm")
    fFinal = Format(dtpFecFin.value, "yyyy/MM/dd") & " " & Format(dtpHorFin.value, "HH:mm")
    If Me.cboTipoPedido.Text <> "" And Me.chkTipoPedido.value = 0 Then
        tipopedido = cboTipoPedido.BoundText
        Else
        tipopedido = ""
    End If
    oComando.CreateParameter "@flagFranjaHoraria", adBoolean, adParamInput, 1, IIf(ChkFranjaHoraria.value = 1, True, False)
    oComando.CreateParameter "@flagTipo", adBoolean, adParamInput, 1, optOpcion(1).value
    oComando.CreateParameter "@flagTurnoFecha", adBoolean, adParamInput, 1, Me.ChkTurno.value
    oComando.CreateParameter "@tturno", adVarChar, adParamInput, 30, sTurno
    oComando.CreateParameter "@tlocal", adVarChar, adParamInput, 30, cboLocal.BoundText
    oComando.CreateParameter "@tSalon", adVarChar, adParamInput, 30, cboSalon.BoundText
    oComando.CreateParameter "@tTipoPedido", adVarChar, adParamInput, 30, tipopedido
    oComando.CreateParameter "@tEstadoPedido", adVarChar, adParamInput, 30, cboEstadoPedido.BoundText
    oComando.CreateParameter "@tcaja", adVarChar, adParamInput, 30, cboCaja.BoundText
    oComando.CreateParameter "@nAdulto", adInteger, adParamInput, 5, CDbl(txtAdulto.Text)
    oComando.CreateParameter "@nPrecuenta", adInteger, adParamInput, 5, CDbl(txtPrecuenta.Text)
    oComando.CreateParameter "@nEnvio", adInteger, adParamInput, 5, CDbl(txtEnvio.Text)
    oComando.CreateParameter "@fInicio", adDBDate, adParamInput, 10, fInicio
    oComando.CreateParameter "@fFinal", adDBDate, adParamInput, 10, fFinal
    oComando.CreateParameter "@OrigenVenta", adVarChar, adParamInput, 30, Me.cboOrigenVenta.BoundText
    If Not oComando.GetParamOK Then
       Set oComando = Nothing
       Exit Sub
    End If
    Set rsReporte = oComando.GetSP()
   If optOpcion(0).value = True Then
      ReporteR.DiscardSavedData
      ReporteR.Database.SetDataSource rsReporte
      ReporteR.Text14.SetText localConectado
      ReporteR.ReportTitle = sTitulo
      ReporteR.Text22.SetText "Correlativo Pedidos Resumido"
      ReporteR.Text13.SetText sRazonSocial
      frmEmite.CRViewer.ReportSource = ReporteR
   Else
      ReporteD.DiscardSavedData
      ReporteD.Database.SetDataSource rsReporte
      ReporteD.ReportTitle = sTitulo
      ReporteD.Text22.SetText "Correlativo Pedidos Detallado"
      ReporteD.Text13.SetText sRazonSocial
      ReporteD.Text15.SetText localConectado
      frmEmite.CRViewer.ReportSource = ReporteD
   End If
End Sub
