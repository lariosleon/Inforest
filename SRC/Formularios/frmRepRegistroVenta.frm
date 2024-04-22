VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "msdatlst.ocx"
Begin VB.Form frmRepRegistroVenta 
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Registro de Ventas"
   ClientHeight    =   8925
   ClientLeft      =   2895
   ClientTop       =   2190
   ClientWidth     =   9045
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   9.75
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmRepRegistroVenta.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8925
   ScaleWidth      =   9045
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
      Index           =   2
      Left            =   4515
      Picture         =   "frmRepRegistroVenta.frx":000C
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   8175
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
      Height          =   8025
      Left            =   45
      TabIndex        =   28
      Top             =   45
      Width           =   8970
      Begin VB.CheckBox chkTipoPago 
         Caption         =   "Todos los Pagos"
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
         Left            =   6480
         TabIndex        =   44
         Top             =   1410
         Value           =   1  'Checked
         Width           =   2325
      End
      Begin VB.CheckBox chkDiaContable 
         Caption         =   "Evaluar por Dia Contable"
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
         Left            =   1560
         TabIndex        =   27
         Top             =   7140
         Width           =   2595
      End
      Begin VB.Frame Frame3 
         BorderStyle     =   0  'None
         Height          =   1215
         Left            =   240
         TabIndex        =   39
         Top             =   5760
         Visible         =   0   'False
         Width           =   8535
         Begin VB.ComboBox CmbMes 
            Height          =   360
            Left            =   5160
            Style           =   2  'Dropdown List
            TabIndex        =   20
            Top             =   120
            Width           =   2265
         End
         Begin MSComCtl2.DTPicker dtpAnual 
            Height          =   405
            Left            =   2280
            TabIndex        =   19
            Top             =   120
            Width           =   2265
            _ExtentX        =   3995
            _ExtentY        =   714
            _Version        =   393216
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            CustomFormat    =   "yyyy"
            Format          =   94240771
            UpDown          =   -1  'True
            CurrentDate     =   37539
         End
         Begin MSComCtl2.DTPicker dtpHora 
            Height          =   375
            Left            =   2280
            TabIndex        =   21
            Top             =   600
            Width           =   2265
            _ExtentX        =   3995
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
            CustomFormat    =   "HH"
            Format          =   94240771
            UpDown          =   -1  'True
            CurrentDate     =   38818
         End
         Begin VB.Label Label 
            AutoSize        =   -1  'True
            Caption         =   "A�o :"
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
            Index           =   9
            Left            =   1680
            TabIndex        =   42
            Top             =   240
            Width           =   465
         End
         Begin VB.Label Label 
            AutoSize        =   -1  'True
            Caption         =   "Hora de Cierre y Apertura  :"
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
            Index           =   8
            Left            =   0
            TabIndex        =   41
            Top             =   600
            Width           =   2355
         End
         Begin VB.Label Label 
            AutoSize        =   -1  'True
            Caption         =   "Mes :"
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
            Index           =   4
            Left            =   4680
            TabIndex        =   40
            Top             =   240
            Width           =   480
         End
      End
      Begin VB.CheckBox chkCaja 
         Caption         =   "Todas las Cajas"
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
         Left            =   6465
         TabIndex        =   12
         Top             =   1785
         Value           =   1  'Checked
         Width           =   1905
      End
      Begin VB.CheckBox chkRegistroVenta 
         Caption         =   "Mostrar Documentos Solo Que Aparecen En Registro de Ventas"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   465
         Left            =   1560
         TabIndex        =   26
         Top             =   7500
         Value           =   1  'Checked
         Width           =   6015
      End
      Begin VB.ComboBox cboRedondeo 
         Height          =   360
         ItemData        =   "frmRepRegistroVenta.frx":082E
         Left            =   2520
         List            =   "frmRepRegistroVenta.frx":0838
         Style           =   2  'Dropdown List
         TabIndex        =   14
         Top             =   2655
         Width           =   1860
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
         Height          =   2490
         Left            =   2520
         TabIndex        =   36
         Top             =   3120
         Width           =   5655
         Begin VB.OptionButton optOpcion 
            Caption         =   "Correlativo de Documento Detallado"
            Height          =   240
            Index           =   6
            Left            =   210
            TabIndex        =   48
            Top             =   930
            Width           =   4935
         End
         Begin VB.OptionButton optOpcion 
            Caption         =   "Detallado por Comprobante"
            Height          =   240
            Index           =   5
            Left            =   210
            TabIndex        =   47
            Top             =   2130
            Width           =   3375
         End
         Begin VB.OptionButton optOpcion 
            Caption         =   "Correlativo de Documento"
            Height          =   240
            Index           =   4
            Left            =   210
            TabIndex        =   43
            Top             =   640
            Width           =   3015
         End
         Begin VB.OptionButton optOpcion 
            Caption         =   "Agrupado Por Tipo Documento"
            Height          =   240
            Index           =   3
            Left            =   210
            TabIndex        =   18
            Top             =   1815
            Width           =   3375
         End
         Begin VB.OptionButton optOpcion 
            Caption         =   "Agrupado por Fechas"
            Height          =   240
            Index           =   2
            Left            =   210
            TabIndex        =   17
            Top             =   1500
            Width           =   3020
         End
         Begin VB.OptionButton optOpcion 
            Caption         =   "Estado de Documentos"
            Height          =   240
            Index           =   1
            Left            =   210
            TabIndex        =   16
            Top             =   1220
            Width           =   2860
         End
         Begin VB.OptionButton optOpcion 
            Caption         =   "Correlativo de Documento Sunat"
            Height          =   240
            Index           =   0
            Left            =   210
            TabIndex        =   15
            Top             =   330
            Value           =   -1  'True
            Width           =   3375
         End
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
         Left            =   2520
         Locked          =   -1  'True
         TabIndex        =   4
         Top             =   195
         Width           =   2775
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
         Left            =   5460
         Picture         =   "frmRepRegistroVenta.frx":085A
         Style           =   1  'Graphical
         TabIndex        =   5
         Top             =   165
         Width           =   765
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
         Left            =   6465
         TabIndex        =   6
         Top             =   210
         Value           =   1  'Checked
         Width           =   2115
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
         Left            =   6465
         TabIndex        =   10
         Top             =   1005
         Value           =   1  'Checked
         Width           =   2325
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
         Left            =   6465
         TabIndex        =   8
         Top             =   630
         Value           =   1  'Checked
         Width           =   2325
      End
      Begin MSComDlg.CommonDialog cmdSave 
         Left            =   7170
         Top             =   -510
         _ExtentX        =   847
         _ExtentY        =   847
         _Version        =   393216
         DefaultExt      =   "Xls"
      End
      Begin VB.ComboBox cboOrden 
         Height          =   360
         ItemData        =   "frmRepRegistroVenta.frx":095C
         Left            =   2520
         List            =   "frmRepRegistroVenta.frx":0969
         Style           =   2  'Dropdown List
         TabIndex        =   13
         Top             =   2220
         Width           =   1860
      End
      Begin MSComCtl2.DTPicker dtpFecFin 
         Height          =   330
         Left            =   2520
         TabIndex        =   24
         Top             =   6360
         Width           =   1860
         _ExtentX        =   3281
         _ExtentY        =   582
         _Version        =   393216
         Format          =   94240769
         CurrentDate     =   37539
      End
      Begin MSComCtl2.DTPicker dtpFecIni 
         Height          =   330
         Left            =   2520
         TabIndex        =   22
         Top             =   5805
         Width           =   1860
         _ExtentX        =   3281
         _ExtentY        =   582
         _Version        =   393216
         Format          =   94240769
         CurrentDate     =   37539
      End
      Begin MSComCtl2.DTPicker dtpHorIni 
         Height          =   330
         Left            =   4440
         TabIndex        =   23
         Top             =   5805
         Width           =   1590
         _ExtentX        =   2805
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
         Format          =   94240771
         UpDown          =   -1  'True
         CurrentDate     =   37539
      End
      Begin MSComCtl2.DTPicker dtpHorFin 
         Height          =   330
         Left            =   4440
         TabIndex        =   25
         Top             =   6360
         Width           =   1590
         _ExtentX        =   2805
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
         Format          =   94240771
         UpDown          =   -1  'True
         CurrentDate     =   37541.9993055556
      End
      Begin MSDataListLib.DataCombo cboTipoDocumento 
         Height          =   315
         Left            =   2520
         TabIndex        =   7
         Top             =   585
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
         Left            =   2520
         TabIndex        =   9
         Top             =   990
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
      Begin MSDataListLib.DataCombo cboCaja 
         Height          =   315
         Left            =   2520
         TabIndex        =   11
         Top             =   1815
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
      Begin MSDataListLib.DataCombo cboTipoPago 
         Height          =   315
         Left            =   2535
         TabIndex        =   45
         Top             =   1400
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
      Begin VB.Label Label 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Tipo de Pago :"
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
         Left            =   1185
         TabIndex        =   46
         Top             =   1455
         Width           =   1275
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
         Index           =   11
         Left            =   1935
         TabIndex        =   38
         Top             =   1860
         Width           =   510
      End
      Begin VB.Label Label 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Redondeo :"
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
         Left            =   1440
         TabIndex        =   37
         Top             =   2745
         Width           =   1005
      End
      Begin VB.Label Label 
         Alignment       =   1  'Right Justify
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
         Left            =   1725
         TabIndex        =   35
         Top             =   255
         Width           =   720
      End
      Begin VB.Label Label 
         Alignment       =   1  'Right Justify
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
         Left            =   705
         TabIndex        =   34
         Top             =   1050
         Width           =   1740
      End
      Begin VB.Label Label 
         Caption         =   "( Factura = Nota de Cr�dito )"
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   6
         Left            =   6855
         TabIndex        =   33
         Top             =   825
         Width           =   1815
      End
      Begin VB.Label Label 
         Alignment       =   1  'Right Justify
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
         Left            =   915
         TabIndex        =   32
         Top             =   645
         Width           =   1530
      End
      Begin VB.Label Label 
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
         Index           =   1
         Left            =   1320
         TabIndex        =   31
         Top             =   6405
         Width           =   1125
      End
      Begin VB.Label Label 
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
         Index           =   0
         Left            =   1215
         TabIndex        =   30
         Top             =   5865
         Width           =   1230
      End
      Begin VB.Label Label 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Ordenado por :"
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
         Left            =   1155
         TabIndex        =   29
         Top             =   2295
         Width           =   1290
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
      Index           =   3
      Left            =   5970
      Picture         =   "frmRepRegistroVenta.frx":098A
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   8175
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
      Left            =   3060
      Picture         =   "frmRepRegistroVenta.frx":0A7C
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   8175
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
      Left            =   1605
      Picture         =   "frmRepRegistroVenta.frx":0FAE
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   8175
      Width           =   1455
   End
End
Attribute VB_Name = "frmRepRegistroVenta"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim RsCompania As Recordset
Dim RsCliente As Recordset
Dim RsEstado As Recordset
Dim RsTipoDocumento As Recordset
Dim RsTipoPago As Recordset
Dim Resumido As New dsrRegistroVentaR
Dim Detallado As New dsrRegistroVentaD
Dim Detallado2 As New dsrRegistroVentaD2
Dim Consolidado As New dsrRegistroVentaC
Dim especial As New dsrRegistroVentaE

'-----Sunat
Dim RsSunat As New dsrRegistroVentaSunat
'-----------

'-----Sunat_AD
Dim RsSunatAD As New dsrRegistroVentaSunatAD
'-----------

'-----Pastipan
Dim RsComprobanteDetallado As New dsrRegistroVentaF
'-----------

Dim Identidad As String
Dim RsPrinter As Recordset
Dim RsCaja    As Recordset
Dim sReporte As String
Dim sCriterio As String
Dim xCriterio As String
Dim vCriterio As String
Dim sCliente As String
Dim fInicio As Date
Dim fFinal As Date
Dim meses(11) As String


Sub LlenaCombos()
   With cboTipoDocumento
      'Compania
      Isql = "Select * from vTipoDocumento where Codigo<>'00' Order By Codigo"
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
   
    With cboCaja
         Isql = "Select tCaja as Codigo, tDescripcion as Descripcion from tCaja order by 1"
         Set RsCaja = Lib.OpenRecordset(Isql, Cn)
         Set .RowSource = RsCaja
             .DataField = "Descripcion"
             .ListField = "Descripcion"
             .BoundColumn = "Codigo"
    End With
    
    With cboTipoPago
      'Estado
      Isql = "select * from vTipoPago where lActivo=1 union select '06','Cortesia','Cortesia',1,'','' order by 1"
      Set RsTipoPago = Lib.OpenRecordset(Isql, Cn)
      Set .RowSource = RsTipoPago
      .DataField = "Descripcion"
      .ListField = "Descripcion"
      .BoundColumn = "Codigo"
   End With

    Dim i As Integer
        For i = 0 To 11
            CmbMes.AddItem meses(i)
        Next i
        CmbMes.ListIndex = Month(Date) - 1
End Sub


Private Sub ChkCaja_Click()
   If ChkCaja.value = 1 Then
      cboCaja.Enabled = False
      cboCaja.Text = ""
   Else
      cboCaja.Enabled = True
   End If
End Sub

'DiaContable
Private Sub chkDiaContable_Click()
    If chkDiaContable.value = 1 Then
        Me.dtpHorFin.Visible = False
        Me.dtpHorIni.Visible = False
    Else
        Me.dtpHorFin.Visible = True
        Me.dtpHorIni.Visible = True
    End If
End Sub

Private Sub cmdBusca_Click()
   Dim xCriterio As String
   Isql = "Select tCodigoCliente as Codigo, tIdentidad as Identidad, tEmpresa as Descripcion from TCLIENTE order by tEmpresa"
   frmBusca.cboCriterio.Enabled = True
   frmBusca.nPredeterm = 2
   Call ConfGrilla(3, frmBusca.grdGrilla, "Codigo", 2, "Codigo", 1200, 0, 0, "", _
                                          "Identificador", 2, "Identidad", 1500, 0, 0, "", _
                                          "Raz�n Comercial", 2, "Descripcion", 4500, 0, 0, "")
   frmBusca.Show vbModal
   If Not wEnter Then
      Exit Sub
   End If
   'sCliente = Calcular("Select tCodigoCliente As Codigo From tCliente Where tIdentidad ='" & sCodigo & "'", Cn)
   sCliente = scodigo
   txtCliente.Text = sDescrip
End Sub

Private Sub cmdOpcion_Click(Index As Integer)
   
   If Index = 3 Then
      Unload Me
      Exit Sub
   End If
   
   sCriterio = ""
   xCriterio = ""
   vCriterio = ""

'DiaContable
If chkDiaContable.value = 0 Then
   If dtpFecIni.value + dtpHorIni.value > dtpFecFin.value + dtpHorFin.value Then
      MsgBox "Error en Rango de Fechas", vbCritical, sMensaje
      Exit Sub
      dtpFecFin.SetFocus
   End If
Else
   If dtpFecIni.value > dtpFecFin.value Then
      MsgBox "Error en Rango de Fechas", vbCritical, sMensaje
      Exit Sub
      dtpFecFin.SetFocus
   End If
End If
      
   If chkCliente.value = 0 Then
      If sCliente = "" Then
         MsgBox "Debe escoger el Cliente", vbCritical, sMensaje
         Exit Sub
      End If
      sCriterio = sCriterio & " AND tCodigoCliente ='" & sCliente & "'"
      xCriterio = xCriterio & " AND tCodigoCliente ='" & sCliente & "'"
      vCriterio = vCriterio & " AND CodigoCliente ='" & sCliente & "'"                          'jesus 160512
   End If
   
   If chkTipoDocumento.value = 0 Then
      If cboTipoDocumento.Text = "" Then
         MsgBox "Debe escoger un Tipo de Documento", vbCritical, sMensaje
         Exit Sub
      End If
      sCriterio = sCriterio & " AND tTipoDocumento ='" & cboTipoDocumento.BoundText & "'"
      xCriterio = xCriterio & " AND Codigo ='" & cboTipoDocumento.BoundText & "'"
      vCriterio = vCriterio & " AND CodigoDocumento ='" & cboTipoDocumento.BoundText & "'"      'jesus 160512
   End If
      
   If chkEstado.value = 0 Then
      If cboEstado.Text = "" Then
         MsgBox "Debe escoger un Estado del Documento", vbCritical, sMensaje
         Exit Sub
      End If
      sCriterio = sCriterio & " AND MDOCUMENTO.tEstadoDocumento ='" & cboEstado.BoundText & "'"
      xCriterio = xCriterio & " AND MNOTACREDITO.tEstadoDocumento ='" & cboEstado.BoundText & "'"
      vCriterio = vCriterio & " AND EstadoDocumento ='" & cboEstado.BoundText & "'"             'jesus 160512
   End If
      
   If ChkCaja.value = 0 Then
      If cboCaja.Text = "" Then
         MsgBox "Debe escoger una Caja", vbCritical, sMensaje
         Exit Sub
      End If
      sCriterio = sCriterio & " AND MDOCUMENTO.tcaja ='" & cboCaja.BoundText & "'"
      xCriterio = xCriterio & " AND MNOTACREDITO.tcaja ='" & cboCaja.BoundText & "'"
      vCriterio = vCriterio & " AND Caja ='" & cboCaja.BoundText & "'"                          'jesus 160512
   End If
   
   'TipoPago
   If chkTipoPago.value = 0 Then
      If cboTipoPago.Text = "" Then
         MsgBox "Debe escoger un Tipo de Pago", vbCritical, sMensaje
         Exit Sub
      End If
      sCriterio = sCriterio & " AND vtipopago.Codigo ='" & cboTipoPago.BoundText & "'"
      'xCriterio = xCriterio & " AND vtipopago.Descripcion ='" & cboTipoPago.BoundText & "'"
   End If
      
   cmdOpcion(Index).Enabled = False
   Select Case Index
   
          Case Is = 0 ' Preview
                 If optOpcion(0).value = True Then '----SUNAT
                            Genera1
                 ElseIf optOpcion(3).value = True Then '----SUNAT_AD
                            Genera2
                 ElseIf optOpcion(5).value = True Then '----PASTIPAN
                            Genera3
                 Else
                            Genera
                 End If

               
               If RsPrinter.EOF = True Then
                  Screen.MousePointer = vbDefault
                  MsgBox "No hay Datos para Mostrar", vbExclamation, "Mensaje del Sistema"
                  cmdOpcion(Index).Enabled = True
                  Exit Sub
               End If
               
               If optOpcion(0).value = True Then
                  RsSunat.SelectPrinter Printer.DriverName, Printer.DeviceName, Printer.Port
                  RsSunat.PaperOrientation = crLandscape
               End If
               If optOpcion(1).value = True Then
                  Resumido.SelectPrinter Printer.DriverName, Printer.DeviceName, Printer.Port
                  Resumido.PaperOrientation = crPortrait
               End If
               If optOpcion(2).value = True Then
                  Consolidado.SelectPrinter Printer.DriverName, Printer.DeviceName, Printer.Port
                  Consolidado.PaperOrientation = crLandscape
               End If
               If optOpcion(3).value = True Then
                  RsSunatAD.SelectPrinter Printer.DriverName, Printer.DeviceName, Printer.Port
                  RsSunatAD.PaperOrientation = crLandscape
               End If
               'CESAR RVAntiguo
               If optOpcion(4).value = True Then
                  Detallado.SelectPrinter Printer.DriverName, Printer.DeviceName, Printer.Port
                  Detallado.PaperOrientation = crLandscape
               End If
               
               'PASTIPAN
               If optOpcion(5).value = True Then
                  RsComprobanteDetallado.SelectPrinter Printer.DriverName, Printer.DeviceName, Printer.Port
                  RsComprobanteDetallado.PaperOrientation = crLandscape
               End If
                'PARDOS
               If optOpcion(6).value = True Then
                  Detallado2.SelectPrinter Printer.DriverName, Printer.DeviceName, Printer.Port
                  Detallado2.PaperOrientation = crLandscape
               End If
               frmEmite.CRViewer.ViewReport
               frmEmite.Show vbModal
          
          Case Is = 1 ' Imprimir
          
                 If optOpcion(0).value = True Then '----SUNAT
                             Genera1
                 ElseIf optOpcion(3).value = True Then '----SUNAT_AD
                             Genera2
                 ElseIf optOpcion(5).value = True Then '----PASTIPAN
                             Genera3
                 Else
                             Genera
                 End If
                 
                
               Screen.MousePointer = vbDefault
               If RsPrinter.EOF = True Then
                  MsgBox "No hay Datos para Mostrar", vbExclamation, "Mensaje del Sistema"
                  cmdOpcion(Index).Enabled = True
                  Exit Sub
               End If
               If optOpcion(0).value = True Then
                  RsSunat.SelectPrinter Printer.DriverName, Printer.DeviceName, Printer.Port
                  RsSunat.PaperOrientation = crLandscape
                  RsSunat.PrintOut
               ElseIf optOpcion(1).value = True Then
                  Resumido.SelectPrinter Printer.DriverName, Printer.DeviceName, Printer.Port
                  Resumido.PaperOrientation = crPortrait
                  Resumido.PrintOut
               ElseIf optOpcion(2).value = True Then
                  Consolidado.SelectPrinter Printer.DriverName, Printer.DeviceName, Printer.Port
                  Consolidado.PaperOrientation = crLandscape
                  Consolidado.PrintOut
               
               ElseIf optOpcion(3).value = True Then
                  RsSunatAD.SelectPrinter Printer.DriverName, Printer.DeviceName, Printer.Port
                  RsSunatAD.PaperOrientation = crLandscape
                  RsSunatAD.PrintOut
               'PASTIPAN
               ElseIf optOpcion(5).value = True Then
                  RsComprobanteDetallado.SelectPrinter Printer.DriverName, Printer.DeviceName, Printer.Port
                  RsComprobanteDetallado.PaperOrientation = crLandscape
                  RsComprobanteDetallado.PrintOut
               'Pardos - erick reporte Correlativo Documento Detallado 160718
               ElseIf optOpcion(6).value = True Then
                  Detallado2.SelectPrinter Printer.DriverName, Printer.DeviceName, Printer.Port
                  Detallado2.PaperOrientation = crLandscape
                  Detallado2.PrintOut
               'CESAR RVAntiguo
               Else
                  Detallado.SelectPrinter Printer.DriverName, Printer.DeviceName, Printer.Port
                  Detallado.PaperOrientation = crLandscape
                  Detallado.PrintOut
               End If
               
          
          Case Is = 2 ' Exportar
          
                 If optOpcion(0).value = True Then '----SUNAT
                             Genera1
                 ElseIf optOpcion(3).value = True Then '----SUNAT_AD
                             Genera2
                 ElseIf optOpcion(5).value = True Then '----PASTIPAN
                             Genera3
                 Else
                             Genera
                 End If
                
               Screen.MousePointer = vbDefault
               If RsPrinter.EOF = True Then
                  MsgBox "No hay Datos para Mostrar", vbExclamation, "Mensaje del Sistema"
                  cmdOpcion(Index).Enabled = True
                  Exit Sub
               End If
               Screen.MousePointer = vbDefault
               'On Error GoTo Cancelar

               If optOpcion(0).value = True Then
                  RsSunat.SelectPrinter Printer.DriverName, Printer.DeviceName, Printer.Port
                  RsSunat.PaperOrientation = crLandscape
                  
                  RsSunat.ExportOptions.FormatType = 21
                  RsSunat.ExportOptions.DestinationType = 1
                  cmdSave.Filter = "Libro de Microsoft Excel|*.xls"
                  cmdSave.ShowSave
                  If cmdSave.FileName = "" Then
                     cmdOpcion(Index).Enabled = True
                     Exit Sub
                  End If
                  RsSunat.ExportOptions.DiskFileName = cmdSave.FileName
                  RsSunat.Export False
                  
               ElseIf optOpcion(1).value = True Then
                  Resumido.SelectPrinter Printer.DriverName, Printer.DeviceName, Printer.Port
                  Resumido.PaperOrientation = crPortrait
                  Resumido.ExportOptions.FormatType = 21
                  Resumido.ExportOptions.DestinationType = 1
                  cmdSave.Filter = "Libro de Microsoft Excel|*.xls"
                  cmdSave.ShowSave
                  If cmdSave.FileName = "" Then
                     cmdOpcion(Index).Enabled = True
                     Exit Sub
                  End If
                  Resumido.ExportOptions.DiskFileName = cmdSave.FileName
                  Resumido.Export False
                  
               ElseIf optOpcion(2).value = True Then
                  Consolidado.SelectPrinter Printer.DriverName, Printer.DeviceName, Printer.Port
                  Consolidado.PaperOrientation = crLandscape
                  Consolidado.ExportOptions.FormatType = 21
                  Consolidado.ExportOptions.DestinationType = 1
                  cmdSave.Filter = "Libro de Microsoft Excel|*.xls"
                  cmdSave.ShowSave
                  If cmdSave.FileName = "" Then
                     cmdOpcion(Index).Enabled = True
                     Exit Sub
                  End If
                  Consolidado.ExportOptions.DiskFileName = cmdSave.FileName
                  Consolidado.Export False
                  
                ElseIf optOpcion(3).value = True Then
                  RsSunatAD.SelectPrinter Printer.DriverName, Printer.DeviceName, Printer.Port
                  RsSunatAD.PaperOrientation = crLandscape
                  RsSunatAD.ExportOptions.FormatType = 21
                  RsSunatAD.ExportOptions.DestinationType = 1
                  cmdSave.Filter = "Libro de Microsoft Excel|*.xls"
                  cmdSave.ShowSave
                  If cmdSave.FileName = "" Then
                     cmdOpcion(Index).Enabled = True
                     Exit Sub
                  End If
                  RsSunatAD.ExportOptions.DiskFileName = cmdSave.FileName
                  RsSunatAD.Export False
                
                'PASTIPAN
                ElseIf optOpcion(5).value = True Then
                  RsComprobanteDetallado.SelectPrinter Printer.DriverName, Printer.DeviceName, Printer.Port
                  RsComprobanteDetallado.PaperOrientation = crPortrait
                  RsComprobanteDetallado.ExportOptions.FormatType = 21
                  RsComprobanteDetallado.ExportOptions.DestinationType = 1
                  cmdSave.Filter = "Libro de Microsoft Excel|*.xls"
                  cmdSave.ShowSave
                  If cmdSave.FileName = "" Then
                     cmdOpcion(Index).Enabled = True
                     Exit Sub
                  End If
                  RsComprobanteDetallado.ExportOptions.DiskFileName = cmdSave.FileName
                  RsComprobanteDetallado.Export False
                'CESAR RVAntiguo
                ElseIf optOpcion(6).value = True Then
                  Detallado2.SelectPrinter Printer.DriverName, Printer.DeviceName, Printer.Port
                  Detallado2.PaperOrientation = crLandscape
                  Detallado2.ExportOptions.FormatType = 21
                  Detallado2.ExportOptions.DestinationType = 1
                  cmdSave.Filter = "Libro de Microsoft Excel|*.xls"
                  cmdSave.ShowSave
                  If cmdSave.FileName = "" Then
                     cmdOpcion(Index).Enabled = True
                     Exit Sub
                  End If
                  Detallado2.ExportOptions.DiskFileName = cmdSave.FileName
                  Detallado2.Export False
                Else
                  Detallado.SelectPrinter Printer.DriverName, Printer.DeviceName, Printer.Port
                  Detallado.PaperOrientation = crLandscape
                  Detallado.ExportOptions.FormatType = 21
                  Detallado.ExportOptions.DestinationType = 1
                  cmdSave.Filter = "Libro de Microsoft Excel|*.xls"
                  cmdSave.ShowSave
                  If cmdSave.FileName = "" Then
                     cmdOpcion(Index).Enabled = True
                     Exit Sub
                  End If
                  Detallado.ExportOptions.DiskFileName = cmdSave.FileName
                  Detallado.Export False
                  
               End If
         
          Case Is = 3 ' Salir
               Unload Me
   End Select
   cmdOpcion(Index).Enabled = True
Cancelar:
   Exit Sub
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

Private Sub chkTipoPago_Click()
   If chkTipoPago.value = 1 Then
      cboTipoPago.Enabled = False
      cboTipoPago.Text = ""
   Else
      cboTipoPago.Enabled = True
   End If
End Sub

Private Sub dtpFecfin_LostFocus()
   If dtpFecIni.value > dtpFecFin.value Then
      MsgBox "Error en Rango de Fechas", vbCritical, sMensaje
      dtpFecFin.SetFocus
   End If
End Sub

Private Sub Form_Load()
   Centrar Me
   dtpFecIni.value = Date
   dtpFecFin.value = Date
   cmdBusca.Enabled = False
   cboOrden.Text = cboOrden.List(0)
    
   cboTipoDocumento.Enabled = False
   cboEstado.Enabled = False
   cboTipoPago.Enabled = False
   cboCaja.Enabled = False
  
    meses(0) = "Enero"
    meses(1) = "Febrero"
    meses(2) = "Marzo"
    meses(3) = "Abril"
    meses(4) = "Mayo"
    meses(5) = "Junio"
    meses(6) = "Julio"
    meses(7) = "Agosto"
    meses(8) = "setiembre"
    meses(9) = "Octubre"
    meses(10) = "Noviembre"
    meses(11) = "Diciembre"
    dtpAnual.value = Date

    LlenaCombos
  
  ' sReporte = dbTemporal(sCaja, 20, "fRegistro", "datetime", "tDocumento", "varchar(15)", "tCodigoCliente", "nVarchar(5)", _
                        "nNeto", "float", "nImpuesto1", "float", "nImpuesto2", "float", "nImpuesto3", "float", "nVenta", "float", "nRecargo", "float", _
                        "nDescuento", "float", "tEstadoDocumento", "varchar(2)", "tTipoDocumento", "varchar(2)", "tTemporal", "varchar(15)", _
                        "nNetoSuma", "float", "nImpuestoSuma1", "float", "nImpuestoSuma2", "float", "nImpuestoSuma3", "float", "nVentaSuma", "float", "nRecargoSuma", "float", "nDescuentoSuma", "float")
End Sub

'-------PASTIPAN----
Public Sub Genera3()
    Dim oComando As clsComando
    
    'DiaContable
    If chkDiaContable.value = 0 Then
    fInicio = Format(dtpFecIni.value, "yyyy/mm/dd") & " " & Format(dtpHorIni.value, "HH:mm")
    fFinal = Format(dtpFecFin.value, "yyyy/mm/dd") & " " & Format(dtpHorFin.value, "HH:mm")
    Else
    fInicio = Format(dtpFecIni.value, "yyyy/mm/dd")
    fFinal = Format(dtpFecFin.value, "yyyy/mm/dd")
    End If
    
    Screen.MousePointer = vbHourglass
    Set oComando = New clsComando
    
    If optOpcion(5).value = True Then
        If Not oComando.CreateCmdSp("spRep_ComprobanteDetallado", Cn) Then
            Set oComando = Nothing
            Exit Sub
        End If
        oComando.CreateParameter "@fInicio", adDBDate, adParamInput, 10, fInicio
        oComando.CreateParameter "@fFinal", adDBDate, adParamInput, 10, fFinal
        oComando.CreateParameter "@tCliente", adVarChar, adParamInput, 20, sCliente
        oComando.CreateParameter "@tTipoDoc", adVarChar, adParamInput, 20, cboTipoDocumento.BoundText
        oComando.CreateParameter "@tEstadoDoc", adVarChar, adParamInput, 20, cboEstado.BoundText
        oComando.CreateParameter "@tCaja", adVarChar, adParamInput, 5, cboCaja.BoundText
        oComando.CreateParameter "@flagRegVenta", adBoolean, adParamInput, 1, Me.chkRegistroVenta.value
        oComando.CreateParameter "@diaContable", adBoolean, adParamInput, 1, Me.chkDiaContable.value
    End If
    
    If Not oComando.GetParamOK Then
       Set oComando = Nothing
       Exit Sub
    End If
    Set RsPrinter = oComando.GetSP()
    
    RsComprobanteDetallado.DiscardSavedData
    RsComprobanteDetallado.Database.SetDataSource RsPrinter
    RsComprobanteDetallado.ReportTitle = "Del " & Format(dtpFecIni.value, "dd/mmm/yyyy") & " Al " & Format(dtpFecFin.value, "dd/mmm/yyyy")
    RsComprobanteDetallado.Text8.SetText sRazonSocial
    RsComprobanteDetallado.Text5.SetText sRUC
    'RsComprobanteDetallado.Text23.SetText localConectado
    
     frmEmite.CRViewer.ReportSource = RsComprobanteDetallado
    
     Screen.MousePointer = vbDefault
End Sub

'-------SUNAT_AD----
Public Sub Genera2()
    Dim oComando As clsComando
    
        'DiaContable
    If chkDiaContable.value = 0 Then
    fInicio = Format(dtpFecIni.value, "yyyy/mm/dd") & " " & Format(dtpHorIni.value, "HH:mm")
    fFinal = Format(dtpFecFin.value, "yyyy/mm/dd") & " " & Format(dtpHorFin.value, "HH:mm")
    Else
    fInicio = Format(dtpFecIni.value, "yyyy/mm/dd")
    fFinal = Format(dtpFecFin.value, "yyyy/mm/dd")
    End If
    
    Screen.MousePointer = vbHourglass
    Set oComando = New clsComando
    
    If optOpcion(3).value = True Then
       If Not oComando.CreateCmdSp("spRep_RegVentaSunatAD", Cn) Then
         Set oComando = Nothing
         Exit Sub
       End If
    oComando.CreateParameter "@fInicio", adDBDate, adParamInput, 10, fInicio
    oComando.CreateParameter "@fFinal", adDBDate, adParamInput, 10, fFinal
    oComando.CreateParameter "@tCliente", adVarChar, adParamInput, 20, sCliente
    oComando.CreateParameter "@tTipoDoc", adVarChar, adParamInput, 20, cboTipoDocumento.BoundText
    oComando.CreateParameter "@tEstadoDoc", adVarChar, adParamInput, 20, cboEstado.BoundText
    oComando.CreateParameter "@tCaja", adVarChar, adParamInput, 5, cboCaja.BoundText
    oComando.CreateParameter "@sOrden", adVarChar, adParamInput, 125, cboOrden.Text
    oComando.CreateParameter "@flagRegVenta", adBoolean, adParamInput, 1, Me.chkRegistroVenta.value
    oComando.CreateParameter "@flagRedondeo", adBoolean, adParamInput, 1, IIf(UCase(cboRedondeo.Text) = "DOS DECIMALES", True, False)
    oComando.CreateParameter "@diaContable", adBoolean, adParamInput, 1, Me.chkDiaContable.value
    
    End If
    
    If Not oComando.GetParamOK Then
       Set oComando = Nothing
       Exit Sub
    End If
    Set RsPrinter = oComando.GetSP()
    
     RsSunatAD.DiscardSavedData
     RsSunatAD.Database.SetDataSource RsPrinter
     RsSunatAD.ReportTitle = "Del " & Format(dtpFecIni.value, "dd/mmm/yyyy") & " Al " & Format(dtpFecFin.value, "dd/mmm/yyyy")
     RsSunatAD.Text21.SetText sRazonSocial
     RsSunatAD.Text22.SetText sRUC
     RsSunatAD.Text23.SetText localConectado

     frmEmite.CRViewer.ReportSource = RsSunatAD
    
     Screen.MousePointer = vbDefault

End Sub




Public Sub Genera1()
    '--------CESAR-------Registro Sunat
    Dim oComando As clsComando
    
    'DiaContable
    If chkDiaContable.value = 0 Then
    fInicio = Format(dtpFecIni.value, "yyyy/mm/dd") & " " & Format(dtpHorIni.value, "HH:mm")
    fFinal = Format(dtpFecFin.value, "yyyy/mm/dd") & " " & Format(dtpHorFin.value, "HH:mm")
    Else
    fInicio = Format(dtpFecIni.value, "yyyy/mm/dd")
    fFinal = Format(dtpFecFin.value, "yyyy/mm/dd")
    End If
    
    Screen.MousePointer = vbHourglass
    Set oComando = New clsComando
    
   '--------Emite el reporte sunat
    If optOpcion(0).value = True Then
       If Not oComando.CreateCmdSp("spRep_RegVentaSunat", Cn) Then
         Set oComando = Nothing
         Exit Sub
       End If
    oComando.CreateParameter "@fInicio", adDBDate, adParamInput, 10, fInicio
    oComando.CreateParameter "@fFinal", adDBDate, adParamInput, 10, fFinal
    oComando.CreateParameter "@tCliente", adVarChar, adParamInput, 20, sCliente
    oComando.CreateParameter "@tTipoDoc", adVarChar, adParamInput, 20, cboTipoDocumento.BoundText
    oComando.CreateParameter "@tEstadoDoc", adVarChar, adParamInput, 20, cboEstado.BoundText
    oComando.CreateParameter "@tCaja", adVarChar, adParamInput, 5, cboCaja.BoundText
    oComando.CreateParameter "@sOrden", adVarChar, adParamInput, 125, cboOrden.Text
    oComando.CreateParameter "@flagRegVenta", adBoolean, adParamInput, 1, Me.chkRegistroVenta.value
    oComando.CreateParameter "@flagRedondeo", adBoolean, adParamInput, 1, IIf(UCase(cboRedondeo.Text) = "DOS DECIMALES", True, False)
    oComando.CreateParameter "@diaContable", adBoolean, adParamInput, 1, Me.chkDiaContable.value
            
    End If
    
    If Not oComando.GetParamOK Then
       Set oComando = Nothing
       Exit Sub
    End If
    Set RsPrinter = oComando.GetSP()
    
     RsSunat.DiscardSavedData
     RsSunat.Database.SetDataSource RsPrinter
     RsSunat.ReportTitle = IIf(chkDiaContable.value = 1, "Por Dia Contable", "") & " Del " & Format(dtpFecIni.value, "dd/mmm/yyyy") & " Al " & Format(dtpFecFin.value, "dd/mmm/yyyy")
     RsSunat.Text21.SetText sRazonSocial
     RsSunat.Text22.SetText sRUC
     RsSunat.Text23.SetText localConectado

     frmEmite.CRViewer.ReportSource = RsSunat
    
     Screen.MousePointer = vbDefault
    
    '/////////////////////////////////
End Sub

Public Sub Genera()
   Screen.MousePointer = vbHourglass
   
   Dim oComando As clsComando
   Dim sOrden As String
   
   
    Dim sFecha As String
    Dim sFecha2 As String
    Dim sAno As String
    Dim sMes As String
 

    sFecha = ""
    sAno = LTrim(str(Year(dtpAnual.value)))
    sMes = IIf(CmbMes.ListIndex + 1 > 9, LTrim(str(CmbMes.ListIndex + 1)), "0" & LTrim(str(CmbMes.ListIndex + 1)))
        
    'DiaContable
    If chkDiaContable.value = 0 Then
    Select Case CmbMes.ListIndex + 1
           Case 2
                If Year(dtpAnual.value) Mod 4 = 0 Then
                   sFecha = "dbo.MDOCUMENTO.fRegistro >= DATEADD(HH," & Hour(dtpHora) & ", '" & sAno & "/" & sMes & "/01') and dbo.MDOCUMENTO.fRegistro <= DATEADD(HH," & 24 + Hour(dtpHora) & ", '" & sAno & "/" & sMes & "/29')"
                   sFecha2 = "dbo.MNOTACREDITO.FFECHA >= DATEADD(HH," & Hour(dtpHora) & ", '" & sAno & "/" & sMes & "/01') and dbo.MNOTACREDITO.FFECHA <= DATEADD(HH," & 24 + Hour(dtpHora) & ", '" & sAno & "/" & sMes & "/29')"
                Else
                   sFecha = "dbo.MDOCUMENTO.fRegistro >= DATEADD(HH," & Hour(dtpHora) & ", '" & sAno & "/" & sMes & "/01') and dbo.MDOCUMENTO.fRegistro <= DATEADD(HH," & 24 + Hour(dtpHora) & ", '" & sAno & "/" & sMes & "/28')"
                   sFecha2 = "dbo.MNOTACREDITO.FFECHA >= DATEADD(HH," & Hour(dtpHora) & ", '" & sAno & "/" & sMes & "/01') and dbo.MNOTACREDITO.FFECHA <= DATEADD(HH," & 24 + Hour(dtpHora) & ", '" & sAno & "/" & sMes & "/28')"
                End If
                
           Case 1, 3, 5, 7, 8, 10, 12
                   sFecha = "dbo.MDOCUMENTO.fRegistro >= DATEADD(HH," & Hour(dtpHora) & ", '" & sAno & "/" & sMes & "/01') and dbo.MDOCUMENTO.fRegistro <= DATEADD(HH," & 24 + Hour(dtpHora) & ", '" & sAno & "/" & sMes & "/31')"
                   sFecha2 = "dbo.MNOTACREDITO.FFECHA >= DATEADD(HH," & Hour(dtpHora) & ", '" & sAno & "/" & sMes & "/01') and dbo.MNOTACREDITO.FFECHA <= DATEADD(HH," & 24 + Hour(dtpHora) & ", '" & sAno & "/" & sMes & "/31')"
           Case 4, 6, 9, 11
                   sFecha = "dbo.MDOCUMENTO.fRegistro >= DATEADD(HH," & Hour(dtpHora) & ", '" & sAno & "/" & sMes & "/01') and dbo.MDOCUMENTO.fRegistro <= DATEADD(HH," & 24 + Hour(dtpHora) & ", '" & sAno & "/" & sMes & "/30')"
                   sFecha2 = "dbo.MNOTACREDITO.FFECHA >= DATEADD(HH," & Hour(dtpHora) & ", '" & sAno & "/" & sMes & "/01') and dbo.MNOTACREDITO.FFECHA <= DATEADD(HH," & 24 + Hour(dtpHora) & ", '" & sAno & "/" & sMes & "/30')"
    End Select
    Else
    Select Case CmbMes.ListIndex + 1
           Case 2
                If Year(dtpAnual.value) Mod 4 = 0 Then
                   sFecha = "dbo.MDOCUMENTO.fDiaContable >= DATEADD(HH," & Hour(dtpHora) & ", '" & sAno & "/" & sMes & "/01') and dbo.MDOCUMENTO.fDiaContable <= DATEADD(HH," & 24 + Hour(dtpHora) & ", '" & sAno & "/" & sMes & "/29')"
                   sFecha2 = "dbo.MNOTACREDITO.fDiaContable >= DATEADD(HH," & Hour(dtpHora) & ", '" & sAno & "/" & sMes & "/01') and dbo.MNOTACREDITO.fDiaContable <= DATEADD(HH," & 24 + Hour(dtpHora) & ", '" & sAno & "/" & sMes & "/29')"
                Else
                   sFecha = "dbo.MDOCUMENTO.fDiaContable >= DATEADD(HH," & Hour(dtpHora) & ", '" & sAno & "/" & sMes & "/01') and dbo.MDOCUMENTO.fDiaContable <= DATEADD(HH," & 24 + Hour(dtpHora) & ", '" & sAno & "/" & sMes & "/28')"
                   sFecha2 = "dbo.MNOTACREDITO.fDiaContable >= DATEADD(HH," & Hour(dtpHora) & ", '" & sAno & "/" & sMes & "/01') and dbo.MNOTACREDITO.fDiaContable <= DATEADD(HH," & 24 + Hour(dtpHora) & ", '" & sAno & "/" & sMes & "/28')"
                End If

           Case 1, 3, 5, 7, 8, 10, 12
                   sFecha = "dbo.MDOCUMENTO.fDiaContable >= DATEADD(HH," & Hour(dtpHora) & ", '" & sAno & "/" & sMes & "/01') and dbo.MDOCUMENTO.fDiaContable <= DATEADD(HH," & 24 + Hour(dtpHora) & ", '" & sAno & "/" & sMes & "/31')"
                   sFecha2 = "dbo.MNOTACREDITO.fDiaContable >= DATEADD(HH," & Hour(dtpHora) & ", '" & sAno & "/" & sMes & "/01') and dbo.MNOTACREDITO.fDiaContable <= DATEADD(HH," & 24 + Hour(dtpHora) & ", '" & sAno & "/" & sMes & "/31')"
           Case 4, 6, 9, 11
                   sFecha = "dbo.MDOCUMENTO.fDiaContable >= DATEADD(HH," & Hour(dtpHora) & ", '" & sAno & "/" & sMes & "/01') and dbo.MDOCUMENTO.fDiaContable <= DATEADD(HH," & 24 + Hour(dtpHora) & ", '" & sAno & "/" & sMes & "/30')"
                   sFecha2 = "dbo.MNOTACREDITO.fDiaContable >= DATEADD(HH," & Hour(dtpHora) & ", '" & sAno & "/" & sMes & "/01') and dbo.MNOTACREDITO.fDiaContable <= DATEADD(HH," & 24 + Hour(dtpHora) & ", '" & sAno & "/" & sMes & "/30')"
    End Select
    End If

    Set oComando = New clsComando
    If optOpcion(6).value = True Then
        If Not oComando.CreateCmdSp("spRep_RegVentaPagos", Cn) Then
           Set oComando = Nothing
           Exit Sub
        End If
   ElseIf optOpcion(4).value = False Then
        If Not oComando.CreateCmdSp("spRep_RegVenta", Cn) Then
           Set oComando = Nothing
           Exit Sub
        End If
   Else
        If Not oComando.CreateCmdSp("spRep_RegVentaDetallado", Cn) Then
           Set oComando = Nothing
           Exit Sub
        End If
   End If
    'DiaContable
    If chkDiaContable.value = 0 Then
    fInicio = Format(dtpFecIni.value, "yyyy/mm/dd") & " " & Format(dtpHorIni.value, "HH:mm")
    fFinal = Format(dtpFecFin.value, "yyyy/mm/dd") & " " & Format(dtpHorFin.value, "HH:mm")
    Else
    fInicio = Format(dtpFecIni.value, "yyyy/mm/dd")
    fFinal = Format(dtpFecFin.value, "yyyy/mm/dd")
    End If
    
    oComando.CreateParameter "@flagTipo", adBoolean, adParamInput, 1, Me.optOpcion(3).value
    oComando.CreateParameter "@flagRegVenta", adBoolean, adParamInput, 1, Me.chkRegistroVenta.value
    'CESAR RVAntiguo
    oComando.CreateParameter "@flagCorrelativo", adBoolean, adParamInput, 1, optOpcion(4).value
    oComando.CreateParameter "@flagEstado", adBoolean, adParamInput, 1, optOpcion(1).value
    oComando.CreateParameter "@flagAgrupado", adBoolean, adParamInput, 1, optOpcion(2).value
    oComando.CreateParameter "@flagRedondeo", adBoolean, adParamInput, 1, IIf(UCase(cboRedondeo.Text) = "DOS DECIMALES", True, False)
    oComando.CreateParameter "@tCliente", adVarChar, adParamInput, 20, sCliente
    oComando.CreateParameter "@tTipoDoc", adVarChar, adParamInput, 20, cboTipoDocumento.BoundText
    oComando.CreateParameter "@tEstadoDoc", adVarChar, adParamInput, 20, cboEstado.BoundText
    oComando.CreateParameter "@tCaja", adVarChar, adParamInput, 5, cboCaja.BoundText
    oComando.CreateParameter "@sOrden", adVarChar, adParamInput, 125, cboOrden.Text
    oComando.CreateParameter "@fInicio", adDBDate, adParamInput, 10, fInicio
    oComando.CreateParameter "@fFinal", adDBDate, adParamInput, 10, fFinal
    oComando.CreateParameter "@sAno", adVarChar, adParamInput, 4, sAno
    oComando.CreateParameter "@sMes", adVarChar, adParamInput, 4, sMes
    oComando.CreateParameter "@sFecha", adVarChar, adParamInput, 400, sFecha
    oComando.CreateParameter "@sFecha2", adVarChar, adParamInput, 400, sFecha2
    oComando.CreateParameter "@dHour", adDouble, adParamInput, 8, Val(Hour(dtpHora))
    oComando.CreateParameter "@flagAnoMes", adBoolean, adParamInput, 1, Me.optOpcion(2).value
    oComando.CreateParameter "@diaContable", adBoolean, adParamInput, 1, Me.chkDiaContable.value
    'TIPOPAGO
     oComando.CreateParameter "@tTipoPago", adVarChar, adParamInput, 20, cboTipoPago.Text
     
    If Not oComando.GetParamOK Then
       Set oComando = Nothing
       Exit Sub
    End If
    Set RsPrinter = oComando.GetSP()
   
   If optOpcion(1).value = True Then
      Resumido.DiscardSavedData
      Resumido.Database.SetDataSource RsPrinter
      
      Resumido.ReportTitle = IIf(Me.chkDiaContable.value = 1, "Por Dia Contable.", "") & " Agrupado por Fechas Del " & Format(dtpFecIni.value, "dd/mmm/yyyy") & " Al " & Format(dtpFecFin.value, "dd/mmm/yyyy")
      Resumido.Text10.SetText IIf(sImpuesto1 = "", "No definido", sImpuesto1)
      Resumido.Text14.SetText IIf(sImpuesto2 = "", "No definido", sImpuesto2)
      Resumido.Text13.SetText IIf(sImpuesto3 = "", "No definido", sImpuesto3)
      Resumido.Text8.SetText sRazonSocial
      Resumido.Text5.SetText localConectado
      frmEmite.CRViewer.DisplayGroupTree = False
      frmEmite.CRViewer.ReportSource = Resumido
            
   ElseIf optOpcion(2).value = True Then
      Consolidado.DiscardSavedData
      Consolidado.Database.SetDataSource RsPrinter
      Consolidado.ReportTitle = IIf(Me.chkDiaContable.value = 1, "Por Dia Contable.", "") & " Agrupado de : " & Year(dtpAnual.value) & " - " & UCase(Me.CmbMes.Text)
      Consolidado.Text10.SetText IIf(sImpuesto1 = "", "No definido", sImpuesto1)
      Consolidado.Text14.SetText IIf(sImpuesto2 = "", "No definido", sImpuesto2)
      Consolidado.Text13.SetText IIf(sImpuesto3 = "", "No definido", sImpuesto3)
      Consolidado.Text8.SetText sRazonSocial
      Consolidado.Text17.SetText localConectado
      frmEmite.CRViewer.DisplayGroupTree = False
      frmEmite.CRViewer.ReportSource = Consolidado
      'CESAR RV Antiguo
   ElseIf optOpcion(4).value = True Then
      Detallado.DiscardSavedData
      Detallado.Database.SetDataSource RsPrinter
      Detallado.ReportTitle = "Del " & Format(dtpFecIni.value, "dd/mmm/yyyy") & " Al " & Format(dtpFecFin.value, "dd/mmm/yyyy")
      Detallado.Text13.SetText sImpuesto1
      Detallado.Text19.SetText sImpuesto2
      Detallado.Text17.SetText sImpuesto3
      Detallado.Text5.SetText sRazonSocial
      
      frmEmite.CRViewer.DisplayGroupTree = False
      frmEmite.CRViewer.ReportSource = Detallado
    ElseIf optOpcion(6).value = True Then
      Detallado2.DiscardSavedData
      Detallado2.Database.SetDataSource RsPrinter
      Detallado2.ReportTitle = "Del " & Format(dtpFecIni.value, "dd/mmm/yyyy") & " Al " & Format(dtpFecFin.value, "dd/mmm/yyyy")
      Detallado2.Text13.SetText sImpuesto1
      Detallado2.Text19.SetText sImpuesto2
      Detallado2.Text17.SetText sImpuesto3
      Detallado2.Text5.SetText sRazonSocial
      
      frmEmite.CRViewer.DisplayGroupTree = False
      frmEmite.CRViewer.ReportSource = Detallado2
   End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
   Set RsPrinter = Nothing
   Set frmRepRegistroVenta = Nothing
   Set RsCliente = Nothing
   Set RsCaja = Nothing
   Set RsEstado = Nothing
   Set RsTipoDocumento = Nothing
   Set RsTipoPago = Nothing
End Sub
'DiaContable
Private Sub optOpcion_Click(Index As Integer)
 If Index = 3 Then
    Me.cboOrden.Enabled = False
    Me.chkTipoPago.Enabled = False
    Me.chkTipoPago.value = 1
    Me.cboTipoPago.Text = ""
    Me.cboTipoPago.Enabled = False
 ElseIf Index = 0 Then
    'Me.cboOrden.Enabled = True
    Me.chkTipoPago.Enabled = False
    Me.chkTipoPago.value = 1
    Me.cboTipoPago.Text = ""
    Me.cboTipoPago.Enabled = False
 Else
    Me.cboOrden.Enabled = True
    Me.chkTipoPago.Enabled = True
 End If
 
 If Index = 2 Then
    Frame3.Visible = True
    Me.chkTipoPago.Enabled = True
 ElseIf Index = 3 Or Index = 0 Then
    Frame3.Visible = False
    Me.chkTipoPago.Enabled = False
    Me.chkTipoPago.value = 1
    Me.cboTipoPago.Text = ""
    Me.cboTipoPago.Enabled = False
  Else
    Frame3.Visible = False
    Me.chkTipoPago.Enabled = True
 End If
 If Index = 5 Then
    Me.cboOrden.Enabled = False
    Me.cboRedondeo.Enabled = False
    Me.chkTipoPago.Enabled = False
    Me.chkTipoPago.value = 1
    Me.cboTipoPago.Text = ""
    Me.cboTipoPago.Enabled = False
 Else
    Me.cboRedondeo.Enabled = True
 End If
End Sub
