VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form frmRepPaloteo 
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   7080
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8505
   ForeColor       =   &H8000000C&
   Icon            =   "frmRepPaloteo.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7080
   ScaleWidth      =   8505
   ShowInTaskbar   =   0   'False
   Begin TabDlg.SSTab SSTab1 
      Height          =   6255
      Left            =   120
      TabIndex        =   65
      Top             =   120
      Width           =   8205
      _ExtentX        =   14473
      _ExtentY        =   11033
      _Version        =   393216
      Tabs            =   2
      TabsPerRow      =   2
      TabHeight       =   520
      TabCaption(0)   =   "Opciones"
      TabPicture(0)   =   "frmRepPaloteo.frx":000C
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Label2"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Label3"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "dtpHorFin"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "dtpHorIni"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "dtpFecIni"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "dtpFecFin"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "chkDiaContable"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).Control(7)=   "Frame3"
      Tab(0).Control(7).Enabled=   0   'False
      Tab(0).Control(8)=   "Frame4"
      Tab(0).Control(8).Enabled=   0   'False
      Tab(0).Control(9)=   "Frame2"
      Tab(0).Control(9).Enabled=   0   'False
      Tab(0).Control(10)=   "chkFranjaHoraria"
      Tab(0).Control(10).Enabled=   0   'False
      Tab(0).Control(11)=   "Frame5"
      Tab(0).Control(11).Enabled=   0   'False
      Tab(0).Control(12)=   "Frame6"
      Tab(0).Control(12).Enabled=   0   'False
      Tab(0).ControlCount=   13
      TabCaption(1)   =   "Filtros"
      TabPicture(1)   =   "frmRepPaloteo.frx":0028
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Label(14)"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).Control(1)=   "Label(1)"
      Tab(1).Control(1).Enabled=   0   'False
      Tab(1).Control(2)=   "Label(0)"
      Tab(1).Control(2).Enabled=   0   'False
      Tab(1).Control(3)=   "Label(2)"
      Tab(1).Control(3).Enabled=   0   'False
      Tab(1).Control(4)=   "Label(4)"
      Tab(1).Control(4).Enabled=   0   'False
      Tab(1).Control(5)=   "Label(5)"
      Tab(1).Control(5).Enabled=   0   'False
      Tab(1).Control(6)=   "Label(6)"
      Tab(1).Control(6).Enabled=   0   'False
      Tab(1).Control(7)=   "Label(7)"
      Tab(1).Control(7).Enabled=   0   'False
      Tab(1).Control(8)=   "Label(9)"
      Tab(1).Control(8).Enabled=   0   'False
      Tab(1).Control(9)=   "Label(3)"
      Tab(1).Control(9).Enabled=   0   'False
      Tab(1).Control(10)=   "Label(10)"
      Tab(1).Control(10).Enabled=   0   'False
      Tab(1).Control(11)=   "Label(11)"
      Tab(1).Control(11).Enabled=   0   'False
      Tab(1).Control(12)=   "Label(8)"
      Tab(1).Control(12).Enabled=   0   'False
      Tab(1).Control(13)=   "Label(12)"
      Tab(1).Control(13).Enabled=   0   'False
      Tab(1).Control(14)=   "Label(13)"
      Tab(1).Control(14).Enabled=   0   'False
      Tab(1).Control(15)=   "Label(15)"
      Tab(1).Control(15).Enabled=   0   'False
      Tab(1).Control(16)=   "cboOrigenVenta"
      Tab(1).Control(16).Enabled=   0   'False
      Tab(1).Control(17)=   "cboSectorVenta"
      Tab(1).Control(17).Enabled=   0   'False
      Tab(1).Control(18)=   "cboUnidadNegocio"
      Tab(1).Control(18).Enabled=   0   'False
      Tab(1).Control(19)=   "cboUsuario"
      Tab(1).Control(19).Enabled=   0   'False
      Tab(1).Control(20)=   "cboCaja"
      Tab(1).Control(20).Enabled=   0   'False
      Tab(1).Control(21)=   "cboArea"
      Tab(1).Control(21).Enabled=   0   'False
      Tab(1).Control(22)=   "cboSubGrupo"
      Tab(1).Control(22).Enabled=   0   'False
      Tab(1).Control(23)=   "cboGrupo"
      Tab(1).Control(23).Enabled=   0   'False
      Tab(1).Control(24)=   "cboMozo"
      Tab(1).Control(24).Enabled=   0   'False
      Tab(1).Control(25)=   "cboTipoPedido"
      Tab(1).Control(25).Enabled=   0   'False
      Tab(1).Control(26)=   "cboSalon"
      Tab(1).Control(26).Enabled=   0   'False
      Tab(1).Control(27)=   "cboTipoProducto"
      Tab(1).Control(27).Enabled=   0   'False
      Tab(1).Control(28)=   "chkSectorVenta"
      Tab(1).Control(28).Enabled=   0   'False
      Tab(1).Control(29)=   "chkTipo"
      Tab(1).Control(29).Enabled=   0   'False
      Tab(1).Control(30)=   "chkSalon"
      Tab(1).Control(30).Enabled=   0   'False
      Tab(1).Control(31)=   "chkTurno"
      Tab(1).Control(31).Enabled=   0   'False
      Tab(1).Control(32)=   "chkGrupo"
      Tab(1).Control(32).Enabled=   0   'False
      Tab(1).Control(33)=   "chkTipoPedido"
      Tab(1).Control(33).Enabled=   0   'False
      Tab(1).Control(34)=   "chkSubGrupo"
      Tab(1).Control(34).Enabled=   0   'False
      Tab(1).Control(35)=   "chkProducto"
      Tab(1).Control(35).Enabled=   0   'False
      Tab(1).Control(36)=   "chkMozo"
      Tab(1).Control(36).Enabled=   0   'False
      Tab(1).Control(37)=   "cmdBusca(0)"
      Tab(1).Control(37).Enabled=   0   'False
      Tab(1).Control(38)=   "cmdBusca(1)"
      Tab(1).Control(38).Enabled=   0   'False
      Tab(1).Control(39)=   "txtProducto"
      Tab(1).Control(39).Enabled=   0   'False
      Tab(1).Control(40)=   "txtTurno"
      Tab(1).Control(40).Enabled=   0   'False
      Tab(1).Control(41)=   "chkArea"
      Tab(1).Control(41).Enabled=   0   'False
      Tab(1).Control(42)=   "txtCliente"
      Tab(1).Control(42).Enabled=   0   'False
      Tab(1).Control(43)=   "cmdBusca(2)"
      Tab(1).Control(43).Enabled=   0   'False
      Tab(1).Control(44)=   "chkCliente"
      Tab(1).Control(44).Enabled=   0   'False
      Tab(1).Control(45)=   "chkCaja"
      Tab(1).Control(45).Enabled=   0   'False
      Tab(1).Control(46)=   "chkUsuario"
      Tab(1).Control(46).Enabled=   0   'False
      Tab(1).Control(47)=   "chkUnidadNegocio"
      Tab(1).Control(47).Enabled=   0   'False
      Tab(1).Control(48)=   "cmdBusca(3)"
      Tab(1).Control(48).Enabled=   0   'False
      Tab(1).Control(49)=   "txtTienda"
      Tab(1).Control(49).Enabled=   0   'False
      Tab(1).Control(50)=   "chkTienda"
      Tab(1).Control(50).Enabled=   0   'False
      Tab(1).Control(51)=   "chkOrigenVenta"
      Tab(1).Control(51).Enabled=   0   'False
      Tab(1).ControlCount=   52
      Begin VB.CheckBox chkOrigenVenta 
         Caption         =   "Todos los Origen Venta"
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
         Left            =   -69400
         TabIndex        =   91
         Top             =   1560
         Value           =   1  'Checked
         Width           =   2505
      End
      Begin VB.Frame Frame6 
         Caption         =   "Tipo de Emision"
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
         Left            =   3885
         TabIndex        =   87
         Top             =   4880
         Width           =   3285
         Begin VB.OptionButton optEmision 
            Caption         =   "Resumido"
            Enabled         =   0   'False
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   240
            Index           =   1
            Left            =   1800
            TabIndex        =   89
            Top             =   360
            Width           =   1335
         End
         Begin VB.OptionButton optEmision 
            Caption         =   "Detallado"
            Enabled         =   0   'False
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   240
            Index           =   0
            Left            =   120
            TabIndex        =   88
            Top             =   360
            Value           =   -1  'True
            Width           =   1260
         End
      End
      Begin VB.CheckBox chkTienda 
         Caption         =   "Todos las Tiendas"
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
         Left            =   -69420
         TabIndex        =   61
         Top             =   5430
         Value           =   1  'Checked
         Width           =   2115
      End
      Begin VB.TextBox txtTienda 
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
         Left            =   -73050
         Locked          =   -1  'True
         TabIndex        =   59
         Top             =   5415
         Width           =   2745
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
         Index           =   3
         Left            =   -70320
         Picture         =   "frmRepPaloteo.frx":0044
         Style           =   1  'Graphical
         TabIndex        =   60
         Top             =   5385
         Width           =   765
      End
      Begin VB.CheckBox chkUnidadNegocio 
         Caption         =   "Todas las Unidad Negocio"
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
         Left            =   -69420
         TabIndex        =   46
         Top             =   3285
         Value           =   1  'Checked
         Width           =   2580
      End
      Begin VB.CheckBox chkUsuario 
         Caption         =   "Todas los Usuarios"
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
         Left            =   -69420
         TabIndex        =   48
         Top             =   3630
         Value           =   1  'Checked
         Width           =   2265
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
         Left            =   -69420
         TabIndex        =   44
         Top             =   2940
         Value           =   1  'Checked
         Width           =   1905
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
         Left            =   -69420
         TabIndex        =   58
         Top             =   5025
         Value           =   1  'Checked
         Width           =   2115
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
         Index           =   2
         Left            =   -70320
         Picture         =   "frmRepPaloteo.frx":0146
         Style           =   1  'Graphical
         TabIndex        =   57
         Top             =   5010
         Width           =   765
      End
      Begin VB.TextBox txtCliente 
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
         Left            =   -73050
         Locked          =   -1  'True
         TabIndex        =   56
         Top             =   5040
         Width           =   2745
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
         Left            =   -69420
         TabIndex        =   40
         Top             =   2235
         Value           =   1  'Checked
         Width           =   1905
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
         Left            =   -73050
         Locked          =   -1  'True
         TabIndex        =   62
         Top             =   5790
         Width           =   2745
      End
      Begin VB.TextBox txtProducto 
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
         Left            =   -73050
         Locked          =   -1  'True
         TabIndex        =   53
         Top             =   4665
         Width           =   2745
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
         Index           =   1
         Left            =   -70320
         Picture         =   "frmRepPaloteo.frx":0248
         Style           =   1  'Graphical
         TabIndex        =   63
         Top             =   5760
         Width           =   765
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
         Left            =   -70320
         Picture         =   "frmRepPaloteo.frx":034A
         Style           =   1  'Graphical
         TabIndex        =   54
         Top             =   4635
         Width           =   765
      End
      Begin VB.CheckBox chkMozo 
         Caption         =   "Todos los Meseros"
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
         Left            =   -69420
         TabIndex        =   38
         Top             =   1890
         Value           =   1  'Checked
         Width           =   2505
      End
      Begin VB.CheckBox chkProducto 
         Caption         =   "Todos los Productos"
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
         Left            =   -69420
         TabIndex        =   55
         Top             =   4680
         Value           =   1  'Checked
         Width           =   2115
      End
      Begin VB.CheckBox chkSubGrupo 
         Caption         =   "Todos los SubGrupos"
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
         Left            =   -69420
         TabIndex        =   52
         Top             =   4335
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
         Left            =   -69420
         TabIndex        =   36
         Top             =   1185
         Value           =   1  'Checked
         Width           =   1905
      End
      Begin VB.CheckBox chkGrupo 
         Caption         =   "Todos los Grupos"
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
         Left            =   -69420
         TabIndex        =   50
         Top             =   3990
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
         Left            =   -69420
         TabIndex        =   64
         Top             =   5805
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
         Left            =   -69420
         TabIndex        =   32
         Top             =   495
         Value           =   1  'Checked
         Width           =   1905
      End
      Begin VB.CheckBox chkTipo 
         Caption         =   "Todos los Tipos"
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
         Left            =   -69420
         TabIndex        =   34
         Top             =   840
         Value           =   1  'Checked
         Width           =   1875
      End
      Begin VB.CheckBox chkSectorVenta 
         Caption         =   "Todas las Agrupaciones"
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
         Left            =   -69420
         TabIndex        =   42
         Top             =   2595
         Value           =   1  'Checked
         Width           =   2500
      End
      Begin VB.Frame Frame5 
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
         Height          =   825
         Left            =   3885
         TabIndex        =   71
         Top             =   3960
         Width           =   3285
         Begin VB.CheckBox chkPorPropiedades 
            Caption         =   "Separado por Propiedades"
            Height          =   345
            Left            =   135
            TabIndex        =   30
            Top             =   315
            Width           =   2820
         End
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
         Height          =   195
         Left            =   5415
         TabIndex        =   9
         Top             =   1005
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
         Height          =   2430
         Left            =   765
         TabIndex        =   68
         Top             =   1440
         Width           =   2730
         Begin VB.OptionButton optOpcion 
            Caption         =   "Pedidos"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   240
            Index           =   2
            Left            =   45
            TabIndex        =   13
            Top             =   1035
            Width           =   1755
         End
         Begin VB.OptionButton optOpcion 
            Caption         =   "Local / Salon"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   240
            Index           =   0
            Left            =   45
            TabIndex        =   10
            Top             =   270
            Value           =   -1  'True
            Width           =   1545
         End
         Begin VB.OptionButton optOpcion 
            Caption         =   "Grupo / SubGrupo"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   240
            Index           =   1
            Left            =   45
            TabIndex        =   11
            Top             =   525
            Width           =   2415
         End
         Begin VB.OptionButton optOpcion 
            Caption         =   "Area de Producción"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   240
            Index           =   3
            Left            =   45
            TabIndex        =   12
            Top             =   780
            Width           =   2265
         End
         Begin VB.OptionButton optOpcion 
            Caption         =   "Tipo de Pedido"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   240
            Index           =   4
            Left            =   45
            TabIndex        =   14
            Top             =   1290
            Width           =   1755
         End
         Begin VB.OptionButton optOpcion 
            Caption         =   "Producto"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   240
            Index           =   5
            Left            =   45
            TabIndex        =   15
            Top             =   1545
            Width           =   1545
         End
         Begin VB.OptionButton optOpcion 
            Caption         =   "Agrupacion (Sub Grupo)"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   240
            Index           =   6
            Left            =   45
            TabIndex        =   16
            Top             =   1800
            Width           =   2475
         End
         Begin VB.OptionButton optOpcion 
            Caption         =   "Comparativo"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   240
            Index           =   7
            Left            =   45
            TabIndex        =   17
            Top             =   2055
            Width           =   1995
         End
      End
      Begin VB.Frame Frame4 
         Caption         =   " Valores "
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
         Left            =   765
         TabIndex        =   67
         Top             =   3960
         Width           =   2730
         Begin VB.CheckBox chkCFacturados 
            Caption         =   "Valorizar con Cero la Pre Venta Facturada"
            Height          =   360
            Left            =   75
            TabIndex        =   29
            Top             =   1140
            Width           =   2415
         End
         Begin VB.OptionButton optValor 
            Caption         =   "Precio Costo"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   240
            Index           =   2
            Left            =   90
            TabIndex        =   28
            Top             =   825
            Width           =   1455
         End
         Begin VB.OptionButton optValor 
            Caption         =   "Precio Neto"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   240
            Index           =   1
            Left            =   90
            TabIndex        =   27
            Top             =   547
            Width           =   1455
         End
         Begin VB.OptionButton optValor 
            Caption         =   "Precio Venta"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   240
            Index           =   0
            Left            =   90
            TabIndex        =   26
            Top             =   270
            Value           =   -1  'True
            Width           =   1500
         End
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
         Height          =   2430
         Left            =   3885
         TabIndex        =   66
         Top             =   1440
         Width           =   3285
         Begin VB.OptionButton optOrigen 
            Caption         =   "Cancelados como Cortesias"
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
            Left            =   135
            TabIndex        =   23
            Top             =   1550
            Width           =   2745
         End
         Begin VB.OptionButton optOrigen 
            Caption         =   "Todos los Pedidos Facturados"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000000C0&
            Height          =   195
            Index           =   6
            Left            =   135
            TabIndex        =   22
            Top             =   1300
            Width           =   2910
         End
         Begin VB.OptionButton optOrigen 
            Caption         =   "Cargos"
            Height          =   195
            Index           =   5
            Left            =   135
            TabIndex        =   25
            Top             =   2055
            Width           =   1935
         End
         Begin VB.OptionButton optOrigen 
            Caption         =   "Combinación"
            Height          =   195
            Index           =   4
            Left            =   135
            TabIndex        =   24
            Top             =   1800
            Width           =   1935
         End
         Begin VB.OptionButton optOrigen 
            Caption         =   "(c) Ctas. Corrientes"
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
            Index           =   3
            Left            =   135
            TabIndex        =   21
            Top             =   1050
            Width           =   2025
         End
         Begin VB.OptionButton optOrigen 
            Caption         =   "(b) Cortesías"
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
            Index           =   2
            Left            =   135
            TabIndex        =   20
            Top             =   800
            Width           =   1935
         End
         Begin VB.OptionButton optOrigen 
            Caption         =   "(a) Ventas en el Rango"
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
            Left            =   135
            TabIndex        =   19
            Top             =   550
            Width           =   2295
         End
         Begin VB.OptionButton optOrigen 
            Caption         =   "Producción (a+b+c)"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00C00000&
            Height          =   195
            Index           =   0
            Left            =   135
            TabIndex        =   18
            Top             =   300
            Value           =   -1  'True
            Width           =   2025
         End
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
         Height          =   195
         Left            =   5415
         TabIndex        =   6
         Top             =   555
         Width           =   2715
      End
      Begin MSComCtl2.DTPicker dtpFecFin 
         Height          =   345
         Left            =   1785
         TabIndex        =   7
         Top             =   990
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
         Format          =   85131265
         CurrentDate     =   37541.9993055556
      End
      Begin MSComCtl2.DTPicker dtpFecIni 
         Height          =   345
         Left            =   1785
         TabIndex        =   4
         Top             =   480
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
         Format          =   85131265
         CurrentDate     =   37539.2083333333
      End
      Begin MSComCtl2.DTPicker dtpHorIni 
         Height          =   375
         Left            =   3840
         TabIndex        =   5
         Top             =   480
         Width           =   1200
         _ExtentX        =   2117
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
         Format          =   85131267
         UpDown          =   -1  'True
         CurrentDate     =   37539
      End
      Begin MSComCtl2.DTPicker dtpHorFin 
         Height          =   375
         Left            =   3840
         TabIndex        =   8
         Top             =   975
         Width           =   1200
         _ExtentX        =   2117
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
         Format          =   85131267
         UpDown          =   -1  'True
         CurrentDate     =   37541.9993055556
      End
      Begin MSDataListLib.DataCombo cboTipoProducto 
         Height          =   315
         Left            =   -73050
         TabIndex        =   33
         Top             =   825
         Width           =   3495
         _ExtentX        =   6165
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
      Begin MSDataListLib.DataCombo cboSalon 
         Height          =   315
         Left            =   -73050
         TabIndex        =   31
         Top             =   480
         Width           =   3495
         _ExtentX        =   6165
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
         Left            =   -73050
         TabIndex        =   35
         Top             =   1155
         Width           =   3495
         _ExtentX        =   6165
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
      Begin MSDataListLib.DataCombo cboMozo 
         Height          =   315
         Left            =   -73050
         TabIndex        =   37
         Top             =   1860
         Width           =   3495
         _ExtentX        =   6165
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
         Left            =   -73050
         TabIndex        =   49
         Top             =   3915
         Width           =   3495
         _ExtentX        =   6165
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
      Begin MSDataListLib.DataCombo cboSubGrupo 
         Height          =   315
         Left            =   -73050
         TabIndex        =   51
         Top             =   4260
         Width           =   3495
         _ExtentX        =   6165
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
      Begin MSDataListLib.DataCombo cboArea 
         Height          =   315
         Left            =   -73050
         TabIndex        =   39
         Top             =   2205
         Width           =   3495
         _ExtentX        =   6165
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
         Left            =   -73050
         TabIndex        =   43
         Top             =   2895
         Width           =   3495
         _ExtentX        =   6165
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
      Begin MSDataListLib.DataCombo cboUsuario 
         Height          =   315
         Left            =   -73050
         TabIndex        =   47
         Top             =   3585
         Width           =   3495
         _ExtentX        =   6165
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
      Begin MSDataListLib.DataCombo cboUnidadNegocio 
         Height          =   315
         Left            =   -73050
         TabIndex        =   45
         Top             =   3240
         Width           =   3495
         _ExtentX        =   6165
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
         Left            =   -73050
         TabIndex        =   41
         Top             =   2550
         Width           =   3495
         _ExtentX        =   6165
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
      Begin MSDataListLib.DataCombo cboOrigenVenta 
         Height          =   315
         Left            =   -73050
         TabIndex        =   90
         Top             =   1500
         Width           =   3495
         _ExtentX        =   6165
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
         BackStyle       =   0  'Transparent
         Caption         =   "Origen de Venta :"
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
         Index           =   15
         Left            =   -74640
         TabIndex        =   92
         Top             =   1560
         Width           =   1515
      End
      Begin VB.Label Label 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Tienda :"
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
         Index           =   13
         Left            =   -73845
         TabIndex        =   86
         Top             =   5475
         Width           =   720
      End
      Begin VB.Label Label 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Unid. Negocio :"
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
         Left            =   -74700
         TabIndex        =   85
         Top             =   3330
         Width           =   1575
      End
      Begin VB.Label Label 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Usuario :"
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
         Left            =   -73905
         TabIndex        =   84
         Top             =   3675
         Width           =   780
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
         Left            =   -73635
         TabIndex        =   83
         Top             =   2985
         Width           =   510
      End
      Begin VB.Label Label 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Cliente Fidelizado :"
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
         Left            =   -74760
         TabIndex        =   82
         Top             =   5100
         Width           =   1635
      End
      Begin VB.Label Label 
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
         Left            =   -74670
         TabIndex        =   81
         Top             =   2280
         Width           =   1545
      End
      Begin VB.Label Label 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Meseros :"
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
         Index           =   9
         Left            =   -73965
         TabIndex        =   80
         Top             =   1935
         Width           =   840
      End
      Begin VB.Label Label 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Producto :"
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
         Left            =   -74025
         TabIndex        =   79
         Top             =   4725
         Width           =   900
      End
      Begin VB.Label Label 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "SubGrupo :"
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
         Left            =   -74100
         TabIndex        =   78
         Top             =   4380
         Width           =   975
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
         Left            =   -74550
         TabIndex        =   77
         Top             =   1230
         Width           =   1425
      End
      Begin VB.Label Label 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Grupo :"
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
         Left            =   -73770
         TabIndex        =   76
         Top             =   4035
         Width           =   645
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
         Left            =   -73755
         TabIndex        =   75
         Top             =   5850
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
         Left            =   -73740
         TabIndex        =   74
         Top             =   540
         Width           =   615
      End
      Begin VB.Label Label 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Tipo de Producto :"
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
         Left            =   -74730
         TabIndex        =   73
         Top             =   885
         Width           =   1605
      End
      Begin VB.Label Label 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Agrupacion P. Venta :"
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
         Index           =   14
         Left            =   -75000
         TabIndex        =   72
         Top             =   2640
         Width           =   1890
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
         Left            =   585
         TabIndex        =   70
         Top             =   1065
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
         Left            =   480
         TabIndex        =   69
         Top             =   690
         Width           =   1230
      End
   End
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
      Left            =   4035
      Picture         =   "frmRepPaloteo.frx":044C
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   6405
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
      Left            =   5475
      Picture         =   "frmRepPaloteo.frx":0C6E
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   6405
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
      Left            =   2595
      Picture         =   "frmRepPaloteo.frx":0D60
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   6405
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
      Left            =   1155
      Picture         =   "frmRepPaloteo.frx":1292
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   6405
      Width           =   1455
   End
   Begin MSComDlg.CommonDialog cmdSave 
      Left            =   45
      Top             =   8235
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
End
Attribute VB_Name = "frmRepPaloteo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim sNombre     As String
Dim rsReporte   As Recordset
Dim RsTempoE    As Recordset
Dim RsTempoT    As Recordset
Dim RsTempoC    As Recordset
Dim RsTempoP    As Recordset
Dim Rsturno     As Recordset
Dim RsSalon     As Recordset
Dim RsMozo      As Recordset
Dim RsTipoPedido As Recordset
Dim RsTipo      As Recordset
Dim RsOrigen     As Recordset
Dim RsArea      As Recordset
Dim RsCaja      As Recordset
Dim RsGrupo     As Recordset
Dim RsSubgrupo  As Recordset
Dim RsProducto  As Recordset
Dim RsSectorVenta  As Recordset
Dim RsCliente   As Recordset
Dim RsTienda    As Recordset
Dim RsUsuario   As Recordset
Dim RsUnidadNegocio  As Recordset

Dim Detallado   As New dsrPaloteoD
Dim Resumido    As New dsrPaloteoR
Dim Pormenorizado As New dsrPaloteoP
Dim PorArea     As New dsrPaloteoA
Dim PorProducto As New dsrPaloteoPP
Dim Comparativo As New dsrPaloteoC

Dim ComparativoResumido As New dsrPaloteoCResumido

Dim sCriterio   As String
Dim xCriterio   As String
Dim sProducto   As String
Dim sTurno      As String
Dim sCliente    As String
Dim sTienda     As String
Dim sFiltro     As String
Dim sTitulo     As String
Dim sPrecio     As String
Dim sTexto      As String
Dim sTexto2     As String

Dim fInicio As Date
Dim fFinal As Date

Sub LlenaCombos()
    With cboSalon
         Isql = "Select * from vSalon"
         Set RsSalon = Lib.OpenRecordset(Isql, Cn)
         Set .RowSource = RsSalon
         .DataField = "Descripcion"
         .ListField = "Descripcion"
         .BoundColumn = "Codigo"
    End With

 With cboSectorVenta
         Isql = "Select codigo, resumido as descripcion from vSectorVenta"
         Set RsSectorVenta = Lib.OpenRecordset(Isql, Cn)
         Set .RowSource = RsSectorVenta
         .DataField = "Descripcion"
         .ListField = "Descripcion"
         .BoundColumn = "Codigo"
    End With
    
    With cboTipoProducto
         Isql = "Select * from vTipoProducto"
         Set RsTipo = Lib.OpenRecordset(Isql, Cn)
         Set .RowSource = RsTipo
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
    
    With cboMozo
         Isql = "Select * from vMozo where substring(Codigo,1,1)<>'*'"
         Set RsMozo = Lib.OpenRecordset(Isql, Cn)
         Set .RowSource = RsMozo
             .DataField = "Descripcion"
             .ListField = "Descripcion"
             .BoundColumn = "Codigo"
    End With
    
    With cboArea
         Isql = "Select * from vArea"
         Set RsArea = Lib.OpenRecordset(Isql, Cn)
         Set .RowSource = RsArea
             .DataField = "Descripcion"
             .ListField = "Descripcion"
             .BoundColumn = "Codigo"
    End With
    
    With cboCaja
         Isql = "Select tCaja as Codigo, tDescripcion as Descripcion from tCaja"  '?????????
         Set RsCaja = Lib.OpenRecordset(Isql, Cn)
         Set .RowSource = RsCaja
             .DataField = "Descripcion"
             .ListField = "Descripcion"
             .BoundColumn = "Codigo"
    End With
    
    With cboUsuario
         Isql = "Select tResumido as Codigo, tDetallado as Descripcion from tUsuario where lActivo = 1 and tGrupoUsuario<>'00'"
         Set RsUsuario = Lib.OpenRecordset(Isql, Cn)
         Set .RowSource = RsUsuario
             .DataField = "Descripcion"
             .ListField = "Descripcion"
             .BoundColumn = "Codigo"
    End With
    
    With cboUnidadNegocio
         Isql = "Select * from vUnidadNegocio"
         Set RsUnidadNegocio = Lib.OpenRecordset(Isql, Cn)
         Set .RowSource = RsUnidadNegocio
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
           
    With cboSubGrupo
         Isql = "Select * from vSubGrupo order by Descripcion"
         Set RsSubgrupo = Lib.OpenRecordset(Isql, Cn)
         Set .RowSource = RsSubgrupo
             .DataField = "Descripcion"
             .ListField = "Descripcion"
             .BoundColumn = "Codigo"
    End With
End Sub

Private Sub cboGrupo_Change()
    cboSubGrupo.BoundText = ""
    With cboSubGrupo
         Isql = "Select * from vSubGrupo " & IIf(cboGrupo.BoundText = "", "", "where tGrupo = '" & cboGrupo.BoundText & "'") & " order by Descripcion "
         Set RsSubgrupo = Lib.OpenRecordset(Isql, Cn)
         Set .RowSource = RsSubgrupo
    End With
    
    sProducto = ""
    txtProducto.Text = ""
End Sub

Private Sub cboTipoPedido_Change()
    Me.cboOrigenVenta.BoundText = ""
    With cboOrigenVenta
         Isql = "Select CodOrigenVenta as codigo, Descripcion   from vOrigenVenta  " & IIf(cboTipoPedido.BoundText = "", "", "where CodCanalVenta = '" & cboTipoPedido.BoundText & "' and Activo=1") & " order by Descripcion "
         Set RsOrigen = Lib.OpenRecordset(Isql, Cn)
         Set .RowSource = RsOrigen
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

Private Sub ChkCaja_Click()
   If chkCaja.value = 1 Then
      cboCaja.Enabled = False
      cboCaja.Text = ""
   Else
      cboCaja.Enabled = True
   End If
End Sub

Private Sub chkDiaContable_Click()
    If chkDiaContable.value = 1 Then
        Me.dtpHorFin.Visible = False
        Me.dtpHorIni.Visible = False
        Me.chkFranjaHoraria.value = 0
        Me.chkFranjaHoraria.Visible = False
    Else
        Me.dtpHorFin.Visible = True
        Me.dtpHorIni.Visible = True
        Me.chkFranjaHoraria.value = 0
        Me.chkFranjaHoraria.Visible = True
    End If
End Sub

Private Sub chkOrigenVenta_Click()
    If chkOrigenVenta.value = 1 Then
        Me.cboOrigenVenta.Enabled = False
        Me.cboOrigenVenta.Text = ""
    Else
        Me.cboOrigenVenta.Enabled = True
    End If
End Sub

Private Sub chkSectorVenta_Click()
If chkSectorVenta.value = 1 Then
      cboSectorVenta.Enabled = False
      cboSectorVenta.Text = ""
   Else
      cboSectorVenta.Enabled = True
   End If
End Sub

Private Sub chkUsuario_Click()
   If chkUsuario.value = 1 Then
      cboUsuario.Enabled = False
      cboUsuario.Text = ""
   Else
      cboUsuario.Enabled = True
   End If
End Sub

Private Sub chkUnidadNegocio_Click()
   If chkUnidadNegocio.value = 1 Then
      cboUnidadNegocio.Enabled = False
      cboUnidadNegocio.Text = ""
   Else
      cboUnidadNegocio.Enabled = True
   End If
End Sub

Private Sub chkTipoPedido_Click()
   If chkTipoPedido.value = 1 Then
      cboTipoPedido.Enabled = False
      cboTipoPedido.Text = ""
   Else
      cboTipoPedido.Enabled = True
   End If
End Sub

Private Sub chkMozo_Click()
   If chkMozo.value = 1 Then
      cboMozo.Enabled = False
      cboMozo.Text = ""
   Else
      cboMozo.Enabled = True
   End If
End Sub

Private Sub chkGrupo_Click()
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

Private Sub chkSalon_Click()
   If chkSalon.value = 1 Then
      cboSalon.Enabled = False
      cboSalon.Text = ""
   Else
      cboSalon.Enabled = True
   End If
End Sub

Private Sub chkTipo_Click()
   If chkTipo.value = 1 Then
      cboTipoProducto.Enabled = False
      cboTipoProducto.Text = ""
   Else
      cboTipoProducto.Enabled = True
   End If
End Sub

Private Sub chkproducto_Click()
   If chkProducto.value = 1 Then
      sProducto = ""
      txtProducto.Text = ""
      cmdBusca(0).Enabled = False
   Else
      cmdBusca(0).Enabled = True
   End If
End Sub

Private Sub chkTienda_Click()
   If chkTienda.value = 1 Then
      sTienda = ""
      txtTienda.Text = ""
      cmdBusca(3).Enabled = False
   Else
      cmdBusca(3).Enabled = True
   End If
End Sub

Private Sub chkCliente_Click()
   If chkCliente.value = 1 Then
      sCliente = ""
      txtCliente.Text = ""
      cmdBusca(2).Enabled = False
   Else
      cmdBusca(2).Enabled = True
   End If
End Sub

Private Sub cmdOpcion_Click(Index As Integer)
   If Index = 2 Then
      Unload Me
      Exit Sub
   End If
   sTitulo = ""
   sFiltro = ""
   
   If chkTurno.value = 0 Then
      If sTurno = "" Then
         MsgBox "Debe escoger un Turno", vbCritical, sMensaje
         Exit Sub
      End If
      
      Set Rsturno = Lib.OpenRecordset("select * from MTURNO where tTurno='" & sTurno & "'", Cn)
      If Rsturno.RecordCount > 0 Then
         sTitulo = " Del " & Format(Rsturno!finicial, "dd/MMM/yyyy HH:mm") & " Al " & Format(IIf(IsNull(Rsturno!fFinal), Format(Rsturno!finicial, "dd/mmm/yyyy") & " 23:59", Rsturno!fFinal), "dd/mmm/yyyy HH:mm")
      Else
         MsgBox "Error Fatal en Turnos", vbCritical, sMensaje
      End If
   Else
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
        If chkDiaContable.value = 0 Then
            If chkFranjaHoraria.value = 0 Then
              sTitulo = " Del " & Format(dtpFecIni.value, "dd/mmm/yyyy") & " " & Format(dtpHorIni.value, "HH:mm") & " Al " & Format(dtpFecFin.value, "dd/mmm/yyyy") & " " & Format(dtpHorFin.value, "HH:mm")
            Else
              sTitulo = " Del " & Format(dtpFecIni.value, "dd/mmm/yyyy") & " Al " & Format(dtpFecFin.value, "dd/mmm/yyyy") & " Franja Horaria : " & Format(dtpHorIni.value, "HH:mm") & " - " & Format(dtpHorFin.value, "HH:mm")
            End If
        Else
            sTitulo = "Por Dia Contable Del " & Format(dtpFecIni.value, "dd/mmm/yyyy") & " Al " & Format(dtpFecFin.value, "dd/mmm/yyyy") & " "
        End If
   End If
  
   If chkSalon.value = 0 Then
      If cboSalon.Text = "" Then
         MsgBox "Debe escoger un Salón", vbCritical, sMensaje
         Exit Sub
      End If
   End If
   
   If chkTipo.value = 0 Then
      If cboTipoProducto.Text = "" Then
         MsgBox "Debe escoger un Tipo de Producto", vbCritical, sMensaje
         Exit Sub
      End If
   End If
   
   If chkMozo.value = 0 Then
      If cboMozo.Text = "" Then
         MsgBox "Debe escoger un Mesero", vbCritical, sMensaje
         Exit Sub
      End If
   End If
   
   If chkTipoPedido.value = 0 Then
      If cboTipoPedido.Text = "" Then
         MsgBox "Debe escoger un Tipo de Pedido", vbCritical, sMensaje
         Exit Sub
      End If
   End If
   
   If chkArea.value = 0 Then
      If cboArea.Text = "" Then
         MsgBox "Debe escoger un Area", vbCritical, sMensaje
        Exit Sub
      End If
   End If

   If chkSectorVenta.value = 0 Then
      If cboSectorVenta.Text = "" Then
         MsgBox "Debe escoger una Agrupacion de Puntos de Ventas", vbCritical, sMensaje
        Exit Sub
      End If
   End If
            
      
      
   If chkCaja.value = 0 Then
      If cboCaja.Text = "" Then
         MsgBox "Debe escoger una Caja", vbCritical, sMensaje
        Exit Sub
      End If
   End If
        
   If chkUsuario.value = 0 Then
      If cboUsuario.Text = "" Then
         MsgBox "Debe escoger una Usuario", vbCritical, sMensaje
        Exit Sub
      End If
       sTitulo = sTitulo & " filtrado por el Usuario : " & cboUsuario.Text
   End If
    
   If chkUnidadNegocio.value = 0 Then
      If cboUnidadNegocio.Text = "" Then
         MsgBox "Debe escoger una Unidad de negocio", vbCritical, sMensaje
        Exit Sub
      End If
      sTitulo = sTitulo & " filtrado por la Unidad de Negocio : " & cboUnidadNegocio.Text
   End If
        
   If chkGrupo.value = 0 Then
      If cboGrupo.Text = "" Then
         MsgBox "Debe escoger el grupo", vbCritical, sMensaje
         Exit Sub
      End If
   End If
   
   If Me.chkOrigenVenta.value = 0 Then
      If Me.cboOrigenVenta.Text = "" Then
         MsgBox "Debe escoger un Origen Venta", vbCritical, sMensaje
         Exit Sub
      End If
   End If

   If chkSubGrupo.value = 0 Then
      If cboSubGrupo.Text = "" Then
         MsgBox "Debe escoger el subgrupo", vbCritical, sMensaje
         Exit Sub
      End If
   End If
   
   If chkProducto.value = 0 Then
      If sProducto = "" Then
         MsgBox "Debe escoger el producto", vbCritical, sMensaje
         Exit Sub
      End If
   End If
   
   If chkCliente.value = 0 Then
      If sCliente = "" Then
         MsgBox "Debe escoger un Cliente", vbCritical, sMensaje
         Exit Sub
      End If
   End If
   
   If chkTienda.value = 0 Then
      If sTienda = "" Then
         MsgBox "Debe escoger una Tienda", vbCritical, sMensaje
         Exit Sub
      End If
   End If
   
   Select Case Index
          Case Is = 0 ' Preview
               Genera
               If rsReporte.EOF = True Then
                  MsgBox "No hay Datos para Mostrar", vbExclamation, "Mensaje del Sistema"
                  Screen.MousePointer = vbDefault
                  Exit Sub
               End If

               If optOpcion(0).value = True Then
                  frmEmite.CRViewer.DisplayGroupTree = True
                  Detallado.SelectPrinter Printer.DriverName, Printer.DeviceName, Printer.Port
                  Detallado.PaperOrientation = crPortrait
               ElseIf optOpcion(1).value = True Then
                  frmEmite.CRViewer.DisplayGroupTree = False
                  Resumido.SelectPrinter Printer.DriverName, Printer.DeviceName, Printer.Port
                  Resumido.PaperOrientation = crPortrait
               ElseIf optOpcion(2).value = True Then
                  frmEmite.CRViewer.DisplayGroupTree = False
                  Pormenorizado.SelectPrinter Printer.DriverName, Printer.DeviceName, Printer.Port
                  Pormenorizado.PaperOrientation = crPortrait
               ElseIf optOpcion(5).value = True Then
                  frmEmite.CRViewer.DisplayGroupTree = False
                  PorProducto.SelectPrinter Printer.DriverName, Printer.DeviceName, Printer.Port
                  PorProducto.PaperOrientation = crPortrait
               ElseIf optOpcion(7).value = True Then
               
                  If optEmision(0).value = True Then
                        frmEmite.CRViewer.DisplayGroupTree = False
                        Comparativo.SelectPrinter Printer.DriverName, Printer.DeviceName, Printer.Port
                        Comparativo.PaperOrientation = crPortrait
                  Else
                        frmEmite.CRViewer.DisplayGroupTree = False
                        ComparativoResumido.SelectPrinter Printer.DriverName, Printer.DeviceName, Printer.Port
                        ComparativoResumido.PaperOrientation = crPortrait
                  End If
                  
               Else
                  frmEmite.CRViewer.DisplayGroupTree = False
                  PorArea.SelectPrinter Printer.DriverName, Printer.DeviceName, Printer.Port
                  PorArea.PaperOrientation = crPortrait
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
               If optOpcion(0).value = True Then
                  Detallado.SelectPrinter Printer.DriverName, Printer.DeviceName, Printer.Port
                  Detallado.PaperOrientation = crPortrait
                  Detallado.PrintOut
               ElseIf optOpcion(1).value = True Then
                  Resumido.SelectPrinter Printer.DriverName, Printer.DeviceName, Printer.Port
                  Resumido.PaperOrientation = crPortrait
                  Resumido.PrintOut
               ElseIf optOpcion(2).value = True Then
                  Pormenorizado.SelectPrinter Printer.DriverName, Printer.DeviceName, Printer.Port
                  Pormenorizado.PaperOrientation = crPortrait
                  Pormenorizado.PrintOut
               ElseIf optOpcion(5).value = True Then
                  PorProducto.SelectPrinter Printer.DriverName, Printer.DeviceName, Printer.Port
                  PorProducto.PaperOrientation = crPortrait
                  PorProducto.PrintOut
               ElseIf optOpcion(7).value = True Then
               
                  If optEmision(0).value = True Then
                        Comparativo.SelectPrinter Printer.DriverName, Printer.DeviceName, Printer.Port
                        Comparativo.PaperOrientation = crPortrait
                        Comparativo.PrintOut
                  Else
                        ComparativoResumido.SelectPrinter Printer.DriverName, Printer.DeviceName, Printer.Port
                        ComparativoResumido.PaperOrientation = crPortrait
                        ComparativoResumido.PrintOut
                  End If
                  
               Else
                  PorArea.SelectPrinter Printer.DriverName, Printer.DeviceName, Printer.Port
                  PorArea.PaperOrientation = crPortrait
                  PorArea.PrintOut
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
               ElseIf optOpcion(1).value = True Then
                  Resumido.ExportOptions.FormatType = 21
                  Resumido.ExportOptions.DestinationType = 1
                  cmdSave.Filter = "Libro de Microsoft Excel|*.xls"
                  cmdSave.ShowSave
                  If cmdSave.FileName = "" Then
                     Exit Sub
                  End If
                  Resumido.ExportOptions.DiskFileName = cmdSave.FileName
                  Resumido.Export False
               ElseIf optOpcion(2).value = True Then
                  Pormenorizado.ExportOptions.FormatType = 21
                  Pormenorizado.ExportOptions.DestinationType = 1
                  cmdSave.Filter = "Libro de Microsoft Excel|*.xls"
                  cmdSave.ShowSave
                  If cmdSave.FileName = "" Then
                     Exit Sub
                  End If
                  Pormenorizado.ExportOptions.DiskFileName = cmdSave.FileName
                  Pormenorizado.Export False
               ElseIf optOpcion(5).value = True Then
                  PorProducto.ExportOptions.FormatType = 21
                  PorProducto.ExportOptions.DestinationType = 1
                  cmdSave.Filter = "Libro de Microsoft Excel|*.xls"
                  cmdSave.ShowSave
                  If cmdSave.FileName = "" Then
                     Exit Sub
                  End If
                  PorProducto.ExportOptions.DiskFileName = cmdSave.FileName
                  PorProducto.Export False
               ElseIf optOpcion(7).value = True Then
               
                  If optEmision(0).value = True Then
                        Comparativo.ExportOptions.FormatType = 21
                        Comparativo.ExportOptions.DestinationType = 1
                        cmdSave.Filter = "Libro de Microsoft Excel|*.xls"
                        cmdSave.ShowSave
                        If cmdSave.FileName = "" Then
                           Exit Sub
                        End If
                        Comparativo.ExportOptions.DiskFileName = cmdSave.FileName
                        Comparativo.Export False
                  Else
                        ComparativoResumido.ExportOptions.FormatType = 21
                        ComparativoResumido.ExportOptions.DestinationType = 1
                        cmdSave.Filter = "Libro de Microsoft Excel|*.xls"
                        cmdSave.ShowSave
                        If cmdSave.FileName = "" Then
                           Exit Sub
                        End If
                        ComparativoResumido.ExportOptions.DiskFileName = cmdSave.FileName
                        ComparativoResumido.Export False
                  End If
                  
               Else
                  PorArea.ExportOptions.FormatType = 21
                  PorArea.ExportOptions.DestinationType = 1
                  cmdSave.Filter = "Libro de Microsoft Excel|*.xls"
                  cmdSave.ShowSave
                  If cmdSave.FileName = "" Then
                     Exit Sub
                  End If
                  PorArea.ExportOptions.DiskFileName = cmdSave.FileName
                  PorArea.Export False
               End If
   End Select
   
End Sub

Private Sub cmdBusca_Click(Index As Integer)
   Dim xCriterio As String
   Select Case Index
          Case Is = 0
               xCriterio = "lActivo = 1 "
               If cboGrupo.Text <> "" Then
                  xCriterio = xCriterio & " and tGrupo ='" & cboGrupo.BoundText & "'"
               End If
               
               If cboSubGrupo.Text <> "" Then
                  xCriterio = xCriterio & " and tSubGrupo ='" & cboSubGrupo.BoundText & "'"
               End If
               
               Isql = "select Codigo, Descripcion, Grupo, SubGrupo, nPrecioVenta from vProducto where " & xCriterio & " order by Descripcion"
               frmBusca.nPredeterm = 3
               Call ConfGrilla(5, frmBusca.grdGrilla, "Codigo", 2, "Codigo", 800, 2, 0, "", _
                                                      "Grupo", 2, "Grupo", 1300, 0, 0, "", _
                                                      "SubGrupo", 2, "SubGrupo", 1400, 0, 0, "", _
                                                      "Plato", 2, "Descripcion", 2900, 0, 0, "", _
                                                      "Precio", 2, "nPrecioVenta", 900, 1, 0, "##,##0.00")
               frmBusca.Show vbModal
               If Not wEnter Then
                  Exit Sub
               End If
               sProducto = sCodigo
               txtProducto.Text = sDescrip
          
          Case Is = 1
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
          
          Case Is = 2
               Isql = "SELECT dbo.TDELIVERY.tCodigoDelivery AS Codigo, dbo.TDELIVERY.tTelefono, LTRIM(dbo.TDELIVERY.tApellido) + ' ' + LTRIM(dbo.TDELIVERY.tNombre) AS Descripcion From dbo.TDELIVERY Where (dbo.TDELIVERY.lActivo = 1)"
               frmBusca.cboCriterio.Enabled = True
               frmBusca.nPredeterm = 2
               Call ConfGrilla(3, frmBusca.grdGrilla, "Codigo", 2, "Codigo", 1200, 0, 0, "", _
                                                      "Identificador", 2, "tTelefono", 1500, 0, 0, "", _
                                                      "Cliente", 2, "Descripcion", 4500, 0, 0, "")
               frmBusca.Show vbModal
               If Not wEnter Then
                  Exit Sub
               End If
               sTienda = ""
               txtTienda.Text = ""
               sCliente = sCodigo
               txtCliente.Text = sDescrip
   
          Case Is = 3
               Isql = "select tCodigoTienda as Codigo, tNombre as Descripcion, tDireccion from TTIENDA where tCodigoDelivery='" & sCliente & "'"
               
               frmBusca.cboCriterio.Enabled = True
               frmBusca.nPredeterm = 2
               Call ConfGrilla(3, frmBusca.grdGrilla, "Codigo", 2, "Codigo", 1200, 0, 0, "", _
                                                      "Identificador", 2, "Descripcion", 1500, 0, 0, "", _
                                                      "Razón Comercial", 2, "tDireccion", 4500, 0, 0, "")
               frmBusca.Show vbModal
               If Not wEnter Then
                  Exit Sub
               End If
               sTienda = sCodigo
               txtTienda.Text = sDescrip
   
   End Select

End Sub

Private Sub dtpFecfin_LostFocus()
   If dtpFecIni.value > dtpFecFin.value Then
      MsgBox "Error en Rango de Fechas", vbCritical, sMensaje
      dtpFecFin.SetFocus
   End If
End Sub

Private Sub ChkTurno_Click()
   If chkTurno.value = 1 Then
      cmdBusca(1).Enabled = False
      txtTurno.Text = ""
      sTurno = ""
      txtTurno.Text = ""
      dtpFecIni.Enabled = True
      dtpFecFin.Enabled = True
      dtpHorIni.Enabled = True
      dtpHorFin.Enabled = True
      chkFranjaHoraria.Enabled = True
      chkFranjaHoraria.value = 0
      Me.chkDiaContable.Enabled = True
      Me.chkDiaContable.value = 0
   Else
      cmdBusca(1).Enabled = True
      dtpFecIni.Enabled = False
      dtpFecFin.Enabled = False
      dtpHorIni.Enabled = False
      dtpHorFin.Enabled = False
      chkFranjaHoraria.Enabled = False
      chkFranjaHoraria.value = 0
      chkDiaContable.Enabled = False
      Me.chkDiaContable.value = 0
   End If
End Sub

Private Sub Form_Load()
    Centrar Me
    LlenaCombos
    dtpFecIni.value = Date
    dtpFecFin.value = Date
    cboTipoProducto.Enabled = False
    cboTipoProducto.Text = ""
    cboSalon.Enabled = False
    cboTipoPedido.Enabled = False
    cboSalon.Text = ""
    cboGrupo.Enabled = False
    cboGrupo.Text = ""
    cboSubGrupo.Enabled = False
    cboSubGrupo.Text = ""
    cboMozo.Enabled = False
    cboMozo.Text = ""
    cboArea.Enabled = False
    cboArea.Text = ""
    cboCaja.Enabled = False
    cboCaja.Text = ""
    cboSectorVenta.Enabled = False
    cboSectorVenta.Text = ""
    cboUsuario.Enabled = False
    cboUnidadNegocio.Enabled = False
    cboUsuario.Text = ""
    cboUnidadNegocio.Text = ""
    Me.cboOrigenVenta.Enabled = False
    Me.cboOrigenVenta.Text = ""
    cmdBusca(0).Enabled = False
    cmdBusca(1).Enabled = False
    cmdBusca(2).Enabled = False
    cmdBusca(3).Enabled = False
End Sub

Private Sub Form_Unload(Cancel As Integer)
   Set rsReporte = Nothing
   Set RsTempoE = Nothing
   Set RsTempoT = Nothing
   Set RsTempoC = Nothing
   Set RsSectorVenta = Nothing
   Set RsTempoP = Nothing
   Set Rsturno = Nothing
   Set RsTipoPedido = Nothing
   Set RsGrupo = Nothing
   Set RsSubgrupo = Nothing
   Set RsCaja = Nothing
   Set RsMozo = Nothing
   Set RsTipo = Nothing
   Set RsUsuario = Nothing
   Set RsUnidadNegocio = Nothing
   Set frmRepPaloteo = Nothing
End Sub

Public Sub Genera()
   Dim sCostoPropiedad As String
   Dim sCombo As String
   Dim sCostoComboPropiedad As String
   
   Dim sPrecio1 As String
   Screen.MousePointer = vbHourglass
   
   
   If chkCFacturados.value = 0 Then
   'Produccion
   If optOrigen(0).value = True Then
   sPrecio1 = "(isnull(dbo.DPEDIDO.nInsumo,0)+isnull(dbo.DPEDIDO.nGasto,0)+isnull(dbo.DPEDIDO.nManoObra,0)) as nCosto, dbo.DPEDIDO.nPrecioNeto as nNeto"
      If optOpcion(7).value Then
        sPrecio = "(isnull(dbo.DPEDIDO.nInsumo,0)+isnull(dbo.DPEDIDO.nGasto,0)+isnull(dbo.DPEDIDO.nManoObra,0)) * dbo.DPEDIDO.nCantidad as nCosto, dbo.DPEDIDO.nPrecioNeto * dbo.DPEDIDO.nCantidad as nNeto"
        sTexto = "Precio Neto vs Costo"
      ElseIf optValor(0).value Then
         sPrecio = "dbo.DPEDIDO.nPrecioVenta * dbo.DPEDIDO.nCantidad"
         sTexto = "Precio Venta"
         sTexto2 = "Total Venta"
      ElseIf optValor(1).value Then
         sPrecio = "dbo.DPEDIDO.nPrecioNeto * dbo.DPEDIDO.nCantidad"
         sTexto = "Precio Neto"
         sTexto2 = "Total Neto"
      Else
         sPrecio = "(isnull(dbo.DPEDIDO.nInsumo,0)+isnull(dbo.DPEDIDO.nGasto,0)+isnull(dbo.DPEDIDO.nManoObra,0)) * dbo.DPEDIDO.nCantidad"
         sTexto = "Precio Costo"
         sTexto2 = "Total Costo"
      End If
            
    End If
'   'Venta
   If optOrigen(1).value = True Then
   sPrecio1 = "(isnull(dbo.DPEDIDO.nInsumo,0)+isnull(dbo.DPEDIDO.nGasto,0)+isnull(dbo.DPEDIDO.nManoObra,0)) as nCosto, dbo.DPEDIDO.nPrecioNeto as nNeto"
      If optOpcion(7).value Then
        sPrecio = "(isnull(dbo.DPEDIDO.nInsumo,0)+isnull(dbo.DPEDIDO.nGasto,0)+isnull(dbo.DPEDIDO.nManoObra,0)) * dbo.DPEDIDO.nCantidad as nCosto, dbo.DPEDIDO.nPrecioNeto * dbo.DPEDIDO.nCantidad as nNeto"
        sTexto = "Precio Neto vs Costo"
      ElseIf optValor(0).value Then
         sPrecio = "dbo.DPEDIDO.nPrecioVenta * dbo.DPEDIDO.nCantidad"
         sTexto = "Precio Venta"
      ElseIf optValor(1).value Then
         sPrecio = "dbo.DPEDIDO.nPrecioNeto * dbo.DPEDIDO.nCantidad"
         sTexto = "Precio Neto"
      Else
         sPrecio = "(isnull(dbo.DPEDIDO.nInsumo,0)+isnull(dbo.DPEDIDO.nGasto,0)+isnull(dbo.DPEDIDO.nManoObra,0)) * dbo.DPEDIDO.nCantidad"
         sTexto = "Precio Costo"
      End If
   End If

   'Cortesia
   If optOrigen(2).value = True Then
   sPrecio1 = "(isnull(dbo.DPEDIDO.nInsumo,0)+isnull(dbo.DPEDIDO.nGasto,0)+isnull(dbo.DPEDIDO.nManoObra,0)) as nCosto, dbo.DPEDIDO.nPrecioNeto as nNeto"
      If optOpcion(7).value Then
        sPrecio = "(isnull(dbo.DPEDIDO.nInsumo,0)+isnull(dbo.DPEDIDO.nGasto,0)+isnull(dbo.DPEDIDO.nManoObra,0)) * dbo.DPEDIDO.nCantidad as nCosto, dbo.DPEDIDO.nPrecioNeto * dbo.DPEDIDO.nCantidad as nNeto"
        sTexto = "Precio Neto vs Costo"
      ElseIf optValor(0).value Then
         sPrecio = "dbo.DPEDIDO.nPrecioVenta * dbo.DPEDIDO.nCantidad"
         sTexto = "Precio Venta"
      ElseIf optValor(1).value Then
         sPrecio = "dbo.DPEDIDO.nPrecioNeto  * dbo.DPEDIDO.nCantidad"
         sTexto = "Precio Neto"
      Else
         sPrecio = "(isnull(dbo.DPEDIDO.nInsumo,0)+isnull(dbo.DPEDIDO.nGasto,0)+isnull(dbo.DPEDIDO.nManoObra,0)) * dbo.DPEDIDO.nCantidad"
         sTexto = "Precio Costo"
      End If
   End If

   'Cuentas Corrientes
   If optOrigen(3).value = True Then
   sPrecio1 = "(isnull(dbo.DPEDIDO.nInsumo,0)+isnull(dbo.DPEDIDO.nGasto,0)+isnull(dbo.DPEDIDO.nManoObra,0)) as nCosto, dbo.DPEDIDO.nPrecioNeto as nNeto"
      If optOpcion(7).value Then
        sPrecio = "(isnull(dbo.DPEDIDO.nInsumo,0)+isnull(dbo.DPEDIDO.nGasto,0)+isnull(dbo.DPEDIDO.nManoObra,0)) * dbo.DPEDIDO.nCantidad as nCosto, dbo.DPEDIDO.nPrecioNeto * dbo.DPEDIDO.nCantidad as nNeto"
        sTexto = "Precio Neto vs Costo"
      ElseIf optValor(0).value Then
         sPrecio = "dbo.DPEDIDO.nPrecioVenta * dbo.DPEDIDO.nCantidad"
         sTexto = "Precio Venta"
      ElseIf optValor(1).value Then
         sPrecio = "dbo.DPEDIDO.nPrecioNeto*dbo.DPEDIDO.nCantidad"
         sTexto = "Precio Neto"
      Else
         sPrecio = "(isnull(dbo.DPEDIDO.nInsumo,0)+isnull(dbo.DPEDIDO.nGasto,0)+isnull(dbo.DPEDIDO.nManoObra,0)) * dbo.DPEDIDO.nCantidad"
         sTexto = "Precio Costo"
      End If
   End If

   'Combinacion
   If optOrigen(4).value = True Then
   sPrecio1 = "(isnull(dbo.CPEDIDO.nInsumo,0)+isnull(dbo.CPEDIDO.nGasto,0)+isnull(dbo.CPEDIDO.nManoObra,0)) as nCosto, dbo.CPEDIDO.nPrecioNeto as nNeto"
      If optOpcion(7).value Then
        'sPrecio = "(isnull(dbo.vProducto.nInsumo,0)+isnull(dbo.vProducto.nGasto,0)+isnull(dbo.vProducto.nManoObra,0)) * dbo.CPEDIDO.nCantidad as nCosto, dbo.CPEDIDO.nPrecioNeto * dbo.CPEDIDO.nCantidad as nNeto"
        sPrecio = "(isnull(dbo.CPEDIDO.nInsumo,0)+isnull(dbo.CPEDIDO.nGasto,0)+isnull(dbo.CPEDIDO.nManoObra,0)) * dbo.CPEDIDO.nCantidad as nCosto , dbo.CPEDIDO.nPrecioNeto * dbo.CPEDIDO.nCantidad as nNeto"
        sTexto = "Precio Neto vs Costo"
      ElseIf optValor(0).value Then
        sPrecio = "dbo.CPEDIDO.nVenta * dbo.CPEDIDO.nCantidad"
        sTexto = "Precio Venta"
      ElseIf optValor(1).value Then
        sPrecio = "dbo.CPEDIDO.nPrecioNeto  * dbo.DPEDIDO.nCantidad"
        sTexto = "Precio Neto"
      Else
        sPrecio = "(isnull(dbo.CPEDIDO.nInsumo,0)+isnull(dbo.CPEDIDO.nGasto,0)+isnull(dbo.CPEDIDO.nManoObra,0)) * dbo.CPEDIDO.nCantidad"
        sTexto = "Precio Costo"
      End If
   End If

   'Cargos
   If optOrigen(5).value = True Then
   sPrecio1 = "(isnull(dbo.DPEDIDO.nInsumo,0)+isnull(dbo.DPEDIDO.nGasto,0)+isnull(dbo.DPEDIDO.nManoObra,0)) as nCosto, dbo.DPEDIDO.nPrecioNeto as nNeto"
      If optOpcion(7).value Then
        sPrecio = "(isnull(dbo.DPEDIDO.nInsumo,0)+isnull(dbo.DPEDIDO.nGasto,0)+isnull(dbo.DPEDIDO.nManoObra,0)) * dbo.DPEDIDO.nCantidad as nCosto, dbo.DPEDIDO.nPrecioNeto * dbo.DPEDIDO.nCantidad as nNeto"
        sTexto = "Precio Neto vs Costo"
      ElseIf optValor(0).value Then
        sPrecio = "dbo.DPEDIDO.nPrecioVenta * dbo.DPEDIDO.nCantidad"
        sTexto = "Precio Venta"
     ElseIf optValor(1).value Then
        sPrecio = "dbo.DPEDIDO.nPrecioNeto*dbo.DPEDIDO.nCantidad"
        sTexto = "Precio Neto"
     Else
        sPrecio = "(isnull(dbo.DPEDIDO.nInsumo,0)+isnull(dbo.DPEDIDO.nGasto,0)+isnull(dbo.DPEDIDO.nManoObra,0)) * dbo.DPEDIDO.nCantidad"
        sTexto = "Precio Costo"
     End If
   End If

   'Pedidos Facturados
   If optOrigen(6).value = True Then
   sPrecio1 = "(isnull(dbo.DPEDIDO.nInsumo,0)+isnull(dbo.DPEDIDO.nGasto,0)+isnull(dbo.DPEDIDO.nManoObra,0)) as nCosto, dbo.DPEDIDO.nPrecioNeto as nNeto"
      If optOpcion(7).value Then
        sPrecio = "(isnull(dbo.DPEDIDO.nInsumo,0)+isnull(dbo.DPEDIDO.nGasto,0)+isnull(dbo.DPEDIDO.nManoObra,0)) * dbo.DPEDIDO.nCantidad as nCosto, dbo.DPEDIDO.nPrecioNeto * dbo.DPEDIDO.nCantidad as nNeto"
        sTexto = "Precio Neto vs Costo"
      ElseIf optValor(0).value Then
         sPrecio = "dbo.DPEDIDO.nPrecioVenta * dbo.DPEDIDO.nCantidad"
         sTexto = "Precio Venta"
      ElseIf optValor(1).value Then
         sPrecio = "dbo.DPEDIDO.nPrecioNeto * dbo.DPEDIDO.nCantidad"
         sTexto = "Precio Neto"
      Else
         sPrecio = "(isnull(dbo.DPEDIDO.nInsumo,0)+isnull(dbo.DPEDIDO.nGasto,0)+isnull(dbo.DPEDIDO.nManoObra,0)) * dbo.DPEDIDO.nCantidad"
         sTexto = "Precio Costo"
      End If
   End If
    Dim tipopedido As String

   'Pedidos Facturados como cortesias
   If optOrigen(8).value = True Then
   sPrecio1 = "(isnull(dbo.DPEDIDO.nInsumo,0)+isnull(dbo.DPEDIDO.nGasto,0)+isnull(dbo.DPEDIDO.nManoObra,0)) as nCosto, dbo.DPEDIDO.nPrecioNeto as nNeto"
      If optOpcion(7).value Then
        sPrecio = "(isnull(dbo.DPEDIDO.nInsumo,0)+isnull(dbo.DPEDIDO.nGasto,0)+isnull(dbo.DPEDIDO.nManoObra,0)) * dbo.DPEDIDO.nCantidad as nCosto, dbo.DPEDIDO.nPrecioNeto * dbo.DPEDIDO.nCantidad as nNeto"
        sTexto = "Precio Neto vs Costo"
      ElseIf optValor(0).value Then
         sPrecio = "dbo.DPEDIDO.nPrecioVenta * dbo.DPEDIDO.nCantidad"
         sTexto = "Precio Venta"
      ElseIf optValor(1).value Then
         sPrecio = "dbo.DPEDIDO.nPrecioNeto * dbo.DPEDIDO.nCantidad"
         sTexto = "Precio Neto"
      Else
         sPrecio = "(isnull(dbo.DPEDIDO.nInsumo,0)+isnull(dbo.DPEDIDO.nGasto,0)+isnull(dbo.DPEDIDO.nManoObra,0)) * dbo.DPEDIDO.nCantidad"
         sTexto = "Precio Costo"
      End If
    End If
    Else ' si valoriza con 0
    
    'Produccion
   If optOrigen(0).value = True Then
   sPrecio1 = "case when dpedido.lregistroventa=0 then 0 else (isnull(dbo.DPEDIDO.nInsumo,0)+isnull(dbo.DPEDIDO.nGasto,0)+isnull(dbo.DPEDIDO.nManoObra,0)) end as nCosto, case when dpedido.lregistroventa=0 then 0 else dbo.DPEDIDO.nPrecioNeto end as nNeto"
      If optOpcion(7).value Then
       sPrecio = "case when dpedido.lregistroventa=0 then 0 else (isnull(dbo.DPEDIDO.nInsumo,0)+isnull(dbo.DPEDIDO.nGasto,0)+isnull(dbo.DPEDIDO.nManoObra,0)) * dbo.DPEDIDO.nCantidad end as nCosto, case when dpedido.lregistroventa=0 then 0 else dbo.DPEDIDO.nPrecioNeto * dbo.DPEDIDO.nCantidad end as nNeto"
       sTexto = "Precio Neto vs Costo"
      ElseIf optValor(0).value Then
         sPrecio = "case when dpedido.lregistroventa=0 then 0 else dbo.DPEDIDO.nPrecioVenta * dbo.DPEDIDO.nCantidad end "
         sTexto = "Precio Venta"
         sTexto2 = "Total Venta"
      ElseIf optValor(1).value Then
         sPrecio = "case when dpedido.lregistroventa=0 then 0 else dbo.DPEDIDO.nPrecioNeto * dbo.DPEDIDO.nCantidad end "
         sTexto = "Precio Neto"
         sTexto2 = "Total Neto"
      Else
         sPrecio = "case when dpedido.lregistroventa=0 then 0 else (isnull(dbo.DPEDIDO.nInsumo,0)+isnull(dbo.DPEDIDO.nGasto,0)+isnull(dbo.DPEDIDO.nManoObra,0)) * dbo.DPEDIDO.nCantidad end "
         sTexto = "Precio Costo"
         sTexto2 = "Total Costo"
      End If
            
    End If
'   'Venta
   If optOrigen(1).value = True Then
   sPrecio1 = "case when dpedido.lregistroventa=0 then 0 else (isnull(dbo.DPEDIDO.nInsumo,0)+isnull(dbo.DPEDIDO.nGasto,0)+isnull(dbo.DPEDIDO.nManoObra,0)) end as nCosto, case when dpedido.lregistroventa=0 then 0 else dbo.DPEDIDO.nPrecioNeto end as nNeto"
      If optOpcion(7).value Then
        sPrecio = "case when dpedido.lregistroventa=0 then 0 else (isnull(dbo.DPEDIDO.nInsumo,0)+isnull(dbo.DPEDIDO.nGasto,0)+isnull(dbo.DPEDIDO.nManoObra,0)) * dbo.DPEDIDO.nCantidad end as nCosto,case when dpedido.lregistroventa=0 then 0 else  dbo.DPEDIDO.nPrecioNeto * dbo.DPEDIDO.nCantidad end as nNeto"
        sTexto = "Precio Neto vs Costo"
      ElseIf optValor(0).value Then
         sPrecio = "case when dpedido.lregistroventa=0 then 0 else dbo.DPEDIDO.nPrecioVenta * dbo.DPEDIDO.nCantidad end "
         sTexto = "Precio Venta"
      ElseIf optValor(1).value Then
         sPrecio = "case when dpedido.lregistroventa=0 then 0 else dbo.DPEDIDO.nPrecioNeto * dbo.DPEDIDO.nCantidad end "
         sTexto = "Precio Neto"
      Else
         sPrecio = "case when dpedido.lregistroventa=0 then 0 else (isnull(dbo.DPEDIDO.nInsumo,0)+isnull(dbo.DPEDIDO.nGasto,0)+isnull(dbo.DPEDIDO.nManoObra,0)) * dbo.DPEDIDO.nCantidad end "
         sTexto = "Precio Costo"
      End If
   End If

   'Cortesia
   If optOrigen(2).value = True Then
   sPrecio1 = "case when dpedido.lregistroventa=0 then 0 else (isnull(dbo.DPEDIDO.nInsumo,0)+isnull(dbo.DPEDIDO.nGasto,0)+isnull(dbo.DPEDIDO.nManoObra,0)) end as nCosto, case when dpedido.lregistroventa=0 then 0 else dbo.DPEDIDO.nPrecioNeto end as nNeto"
      If optOpcion(7).value Then
        sPrecio = "case when dpedido.lregistroventa=0 then 0 else (isnull(dbo.DPEDIDO.nInsumo,0)+isnull(dbo.DPEDIDO.nGasto,0)+isnull(dbo.DPEDIDO.nManoObra,0)) * dbo.DPEDIDO.nCantidad end as nCosto, case when dpedido.lregistroventa=0 then 0 else dbo.DPEDIDO.nPrecioNeto * dbo.DPEDIDO.nCantidad end as nNeto"
        sTexto = "Precio Neto vs Costo"
      ElseIf optValor(0).value Then
         sPrecio = "case when dpedido.lregistroventa=0 then 0 else dbo.DPEDIDO.nPrecioVenta * dbo.DPEDIDO.nCantidad end "
         sTexto = "Precio Venta"
      ElseIf optValor(1).value Then
         sPrecio = "case when dpedido.lregistroventa=0 then 0 else dbo.DPEDIDO.nPrecioNeto  * dbo.DPEDIDO.nCantidad end "
         sTexto = "Precio Neto"
      Else
         sPrecio = "case when dpedido.lregistroventa=0 then 0 else (isnull(dbo.DPEDIDO.nInsumo,0)+isnull(dbo.DPEDIDO.nGasto,0)+isnull(dbo.DPEDIDO.nManoObra,0)) * dbo.DPEDIDO.nCantidad end "
         sTexto = "Precio Costo"
      End If
   End If

   'Cuentas Corrientes
   If optOrigen(3).value = True Then
   sPrecio1 = "case when dpedido.lregistroventa=0 then 0 else (isnull(dbo.DPEDIDO.nInsumo,0)+isnull(dbo.DPEDIDO.nGasto,0)+isnull(dbo.DPEDIDO.nManoObra,0)) end as nCosto, case when dpedido.lregistroventa=0 then 0 else dbo.DPEDIDO.nPrecioNeto end as nNeto"
      If optOpcion(7).value Then
        sPrecio = "case when dpedido.lregistroventa=0 then 0 else (isnull(dbo.DPEDIDO.nInsumo,0)+isnull(dbo.DPEDIDO.nGasto,0)+isnull(dbo.DPEDIDO.nManoObra,0)) * dbo.DPEDIDO.nCantidad end as nCosto, case when dpedido.lregistroventa=0 then 0 else dbo.DPEDIDO.nPrecioNeto * dbo.DPEDIDO.nCantidad end as nNeto"
        sTexto = "Precio Neto vs Costo"
      ElseIf optValor(0).value Then
         sPrecio = "case when dpedido.lregistroventa=0 then 0 else dbo.DPEDIDO.nPrecioVenta * dbo.DPEDIDO.nCantidad end "
         sTexto = "Precio Venta"
      ElseIf optValor(1).value Then
         sPrecio = "case when dpedido.lregistroventa=0 then 0 else dbo.DPEDIDO.nPrecioNeto*dbo.DPEDIDO.nCantidad end "
         sTexto = "Precio Neto"
      Else
         sPrecio = "case when dpedido.lregistroventa=0 then 0 else (isnull(dbo.DPEDIDO.nInsumo,0)+isnull(dbo.DPEDIDO.nGasto,0)+isnull(dbo.DPEDIDO.nManoObra,0)) * dbo.DPEDIDO.nCantidad end "
         sTexto = "Precio Costo"
      End If
   End If

   'Combinacion
   If optOrigen(4).value = True Then
   sPrecio1 = "case when dpedido.lregistroventa=0 then 0 else (isnull(dbo.CPEDIDO.nInsumo,0)+isnull(dbo.CPEDIDO.nGasto,0)+isnull(dbo.CPEDIDO.nManoObra,0)) end as nCosto, case when pedido.lregistroventa=0 then 0 else dbo.CPEDIDO.nPrecioNeto end as nNeto"
      If optOpcion(7).value Then
        sPrecio = "case when dpedido.lregistroventa=0 then 0 else (isnull(dbo.vProducto.nInsumo,0)+isnull(dbo.vProducto.nGasto,0)+isnull(dbo.vProducto.nManoObra,0)) * dbo.CPEDIDO.nCantidad end as nCosto, case when dpedido.lregistroventa=0 then 0 else dbo.CPEDIDO.nPrecioNeto * dbo.CPEDIDO.nCantidad end as nNeto"
        sTexto = "Precio Neto vs Costo"
      ElseIf optValor(0).value Then
        sPrecio = "case when dpedido.lregistroventa=0 then 0 else dbo.CPEDIDO.nVenta * dbo.CPEDIDO.nCantidad end "
        sTexto = "Precio Venta"
      ElseIf optValor(1).value Then
        sPrecio = "case when dpedido.lregistroventa=0 then 0 else dbo.CPEDIDO.nPrecioNeto  * dbo.DPEDIDO.nCantidad end"
        sTexto = "Precio Neto"
      Else
        sPrecio = "case when dpedido.lregistroventa=0 then 0 else (isnull(dbo.CPEDIDO.nInsumo,0)+isnull(dbo.CPEDIDO.nGasto,0)+isnull(dbo.CPEDIDO.nManoObra,0)) * dbo.CPEDIDO.nCantidad end "
        sTexto = "Precio Costo"
      End If
   End If

   'Cargos
   If optOrigen(5).value = True Then
   sPrecio1 = "case when dpedido.lregistroventa=0 then 0 else (isnull(dbo.DPEDIDO.nInsumo,0)+isnull(dbo.DPEDIDO.nGasto,0)+isnull(dbo.DPEDIDO.nManoObra,0)) end as nCosto, case when dpedido.lregistroventa=0 then 0 else dbo.DPEDIDO.nPrecioNeto end as nNeto"
      If optOpcion(7).value Then
        sPrecio = "case when dpedido.lregistroventa=0 then 0 else (isnull(dbo.DPEDIDO.nInsumo,0)+isnull(dbo.DPEDIDO.nGasto,0)+isnull(dbo.DPEDIDO.nManoObra,0)) * dbo.DPEDIDO.nCantidad end as nCosto,case when dpedido.lregistroventa=0 then 0 else dbo.DPEDIDO.nPrecioNeto * dbo.DPEDIDO.nCantidad end as nNeto"
        sTexto = "Precio Neto vs Costo"
      ElseIf optValor(0).value Then
        sPrecio = "case when dpedido.lregistroventa=0 then 0 else dbo.DPEDIDO.nPrecioVenta * dbo.DPEDIDO.nCantidad end "
        sTexto = "Precio Venta"
     ElseIf optValor(1).value Then
        sPrecio = "case when dpedido.lregistroventa=0 then 0 else dbo.DPEDIDO.nPrecioNeto*dbo.DPEDIDO.nCantidad end "
        sTexto = "Precio Neto"
     Else
        sPrecio = "case when dpedido.lregistroventa=0 then 0 else (isnull(dbo.DPEDIDO.nInsumo,0)+isnull(dbo.DPEDIDO.nGasto,0)+isnull(dbo.DPEDIDO.nManoObra,0)) * dbo.DPEDIDO.nCantidad end "
        sTexto = "Precio Costo"
     End If
   End If

   'Pedidos Facturados
   If optOrigen(6).value = True Then
   sPrecio1 = "case when dpedido.lregistroventa=0 then 0 else (isnull(dbo.DPEDIDO.nInsumo,0)+isnull(dbo.DPEDIDO.nGasto,0)+isnull(dbo.DPEDIDO.nManoObra,0)) end as nCosto, case when dpedido.lregistroventa=0 then 0 else dbo.DPEDIDO.nPrecioNeto end as nNeto"
      If optOpcion(7).value Then
        sPrecio = "case when dpedido.lregistroventa=0 then 0 else (isnull(dbo.DPEDIDO.nInsumo,0)+isnull(dbo.DPEDIDO.nGasto,0)+isnull(dbo.DPEDIDO.nManoObra,0)) * dbo.DPEDIDO.nCantidad end as nCosto, case when dpedido.lregistroventa=0 then 0 else dbo.DPEDIDO.nPrecioNeto * dbo.DPEDIDO.nCantidad end as nNeto"
        sTexto = "Precio Neto vs Costo"
      ElseIf optValor(0).value Then
         sPrecio = "case when dpedido.lregistroventa=0 then 0 else dbo.DPEDIDO.nPrecioVenta * dbo.DPEDIDO.nCantidad end "
         sTexto = "Precio Venta"
      ElseIf optValor(1).value Then
         sPrecio = "case when dpedido.lregistroventa=0 then 0 else dbo.DPEDIDO.nPrecioNeto * dbo.DPEDIDO.nCantidad end "
         sTexto = "Precio Neto"
      Else
         sPrecio = "case when dpedido.lregistroventa=0 then 0 else (isnull(dbo.DPEDIDO.nInsumo,0)+isnull(dbo.DPEDIDO.nGasto,0)+isnull(dbo.DPEDIDO.nManoObra,0)) * dbo.DPEDIDO.nCantidad end "
         sTexto = "Precio Costo"
      End If
   End If
    

   'Pedidos Facturados como cortesias
   If optOrigen(8).value = True Then
   sPrecio1 = "case when dpedido.lregistroventa=0 then 0 else (isnull(dbo.DPEDIDO.nInsumo,0)+isnull(dbo.DPEDIDO.nGasto,0)+isnull(dbo.DPEDIDO.nManoObra,0)) end as nCosto, case when dpedido.lregistroventa=0 then 0 else dbo.DPEDIDO.nPrecioNeto end as nNeto"
      If optOpcion(7).value Then
        sPrecio = "case when dpedido.lregistroventa=0 then 0 else (isnull(dbo.DPEDIDO.nInsumo,0)+isnull(dbo.DPEDIDO.nGasto,0)+isnull(dbo.DPEDIDO.nManoObra,0)) * dbo.DPEDIDO.nCantidad end as nCosto,case when dpedido.lregistroventa=0 then 0 else  dbo.DPEDIDO.nPrecioNeto * dbo.DPEDIDO.nCantidad end as nNeto"
        sTexto = "Precio Neto vs Costo"
      ElseIf optValor(0).value Then
         sPrecio = "case when dpedido.lregistroventa=0 then dbo.DPEDIDO.nPrecioVenta * dbo.DPEDIDO.nCantidad end "
         sTexto = "Precio Venta"
      ElseIf optValor(1).value Then
         sPrecio = "case when dpedido.lregistroventa=0 then dbo.DPEDIDO.nPrecioNeto * dbo.DPEDIDO.nCantidad end "
         sTexto = "Precio Neto"
      Else
         sPrecio = "case when dpedido.lregistroventa=0 then (isnull(dbo.DPEDIDO.nInsumo,0)+isnull(dbo.DPEDIDO.nGasto,0)+isnull(dbo.DPEDIDO.nManoObra,0)) * dbo.DPEDIDO.nCantidad end "
         sTexto = "Precio Costo"
      End If
    End If
    End If
        If Me.cboTipoPedido.Text <> "" And Me.chkTipoPedido.value = 0 Then
        tipopedido = cboTipoPedido.BoundText
        Else
        tipopedido = ""
        End If
    Dim oComando As clsComando
    Set oComando = New clsComando
    If Not oComando.CreateCmdSp("spRep_PaloteoProduccion", Cn) Then
       Set oComando = Nothing
       Exit Sub
    End If
    fFinal = Format(dtpFecFin.value, "yyyy/MM/DD") & " " & Format(dtpHorFin.value, "HH:mm")
    fInicio = Format(dtpFecIni.value, "yyyy/MM/DD") & " " & Format(dtpHorIni.value, "HH:mm")
    oComando.CreateParameter "@flagFranjaHoraria", adBoolean, adParamInput, 1, IIf(chkFranjaHoraria.value = 1, True, False)
    oComando.CreateParameter "@flagConPropiedades", adBoolean, adParamInput, 1, IIf(chkPorPropiedades.value = 1, True, False)
    oComando.CreateParameter "@flagProduccion", adBoolean, adParamInput, 1, optOrigen(0).value
    oComando.CreateParameter "@flagVenta", adBoolean, adParamInput, 1, optOrigen(1).value
    oComando.CreateParameter "@flagCortesia", adBoolean, adParamInput, 1, optOrigen(2).value
    oComando.CreateParameter "@flagCuentaCte", adBoolean, adParamInput, 1, optOrigen(3).value
    oComando.CreateParameter "@flagCombinacion", adBoolean, adParamInput, 1, optOrigen(4).value
    oComando.CreateParameter "@flagCargo", adBoolean, adParamInput, 1, optOrigen(5).value
    oComando.CreateParameter "@flagPedidosFacturados", adBoolean, adParamInput, 1, optOrigen(6).value
    oComando.CreateParameter "@flagPedidosFacturadosCortesia", adBoolean, adParamInput, 1, optOrigen(8).value
    oComando.CreateParameter "@flagOpcionCom", adBoolean, adParamInput, 1, optOpcion(7).value
    oComando.CreateParameter "@flagTurnoOFecha", adBoolean, adParamInput, 1, chkTurno.value
    oComando.CreateParameter "@flagPrecioCosto", adBoolean, adParamInput, 1, optValor(2).value
    oComando.CreateParameter "@flagSeleccion0", adBoolean, adParamInput, 1, optOpcion(0).value
    oComando.CreateParameter "@flagSeleccion1", adBoolean, adParamInput, 1, optOpcion(1).value
    oComando.CreateParameter "@flagSeleccion2", adBoolean, adParamInput, 1, optOpcion(2).value
    oComando.CreateParameter "@flagSeleccion3", adBoolean, adParamInput, 1, optOpcion(3).value
    oComando.CreateParameter "@flagSeleccion4", adBoolean, adParamInput, 1, optOpcion(4).value
    oComando.CreateParameter "@flagSeleccion5", adBoolean, adParamInput, 1, optOpcion(5).value
    oComando.CreateParameter "@flagSeleccion6", adBoolean, adParamInput, 1, optOpcion(6).value
    oComando.CreateParameter "@flagSeleccion7", adBoolean, adParamInput, 1, optOpcion(7).value
    oComando.CreateParameter "@turno", adVarChar, adParamInput, 30, sTurno
    oComando.CreateParameter "@tSalon", adVarChar, adParamInput, 30, cboSalon.BoundText
    oComando.CreateParameter "@tTipoProducto", adVarChar, adParamInput, 30, cboTipoProducto.BoundText
    oComando.CreateParameter "@tMozo", adVarChar, adParamInput, 30, cboMozo.BoundText
    oComando.CreateParameter "@tTipoPedido", adVarChar, adParamInput, 30, tipopedido
    oComando.CreateParameter "@tArea", adVarChar, adParamInput, 30, cboArea.Text
    oComando.CreateParameter "@tCaja", adVarChar, adParamInput, 30, cboCaja.BoundText
    oComando.CreateParameter "@tUsuario", adVarChar, adParamInput, 30, cboUsuario.BoundText
    oComando.CreateParameter "@tUnidadNegocio", adVarChar, adParamInput, 30, cboUnidadNegocio.BoundText
    oComando.CreateParameter "@tGrupo", adVarChar, adParamInput, 30, cboGrupo.Text
    oComando.CreateParameter "@tSubGrupo", adVarChar, adParamInput, 30, cboSubGrupo.Text
    oComando.CreateParameter "@tCodigoProducto", adVarChar, adParamInput, 30, sProducto
    oComando.CreateParameter "@tCodigoCliente", adVarChar, adParamInput, 30, sCliente
    oComando.CreateParameter "@tCodigoTienda", adVarChar, adParamInput, 30, sTienda
    oComando.CreateParameter "@tCodigoSectorVenta", adVarChar, adParamInput, 30, cboSectorVenta.BoundText
    oComando.CreateParameter "@sPrecio", adVarChar, adParamInput, 400, sPrecio
    oComando.CreateParameter "@sBoton2", adVarChar, adParamInput, 30, sBoton2
    oComando.CreateParameter "@sBoton3", adVarChar, adParamInput, 30, sBoton3
    oComando.CreateParameter "@sBoton4", adVarChar, adParamInput, 30, sBoton4
    oComando.CreateParameter "@sBoton5", adVarChar, adParamInput, 30, sBoton5
    oComando.CreateParameter "@fInicio", adDBDate, adParamInput, 10, fInicio
    oComando.CreateParameter "@fFinal", adDBDate, adParamInput, 10, fFinal
    oComando.CreateParameter "@flagCFacturado", adBoolean, adParamInput, 1, IIf(chkCFacturados.value = 1, 1, 0)
    oComando.CreateParameter "@flagDiaContable", adBoolean, adParamInput, 1, IIf(chkDiaContable.value = 1, 1, 0)
    
    oComando.CreateParameter "@flagTipoEmision", adBoolean, adParamInput, 1, optEmision(0).value
    oComando.CreateParameter "@sPrecio1", adVarChar, adParamInput, 400, sPrecio1
    oComando.CreateParameter "@OrigenVenta", adVarChar, adParamInput, 30, Me.cboOrigenVenta.BoundText
    
    If Not oComando.GetParamOK Then
       Set oComando = Nothing
       Exit Sub
    End If
    Set rsReporte = oComando.GetSP()
    
   If optOpcion(0).value = True Then
      Detallado.DiscardSavedData
      Detallado.Text8.SetText localConectado
      Detallado.Database.SetDataSource rsReporte
      Detallado.Text10.SetText "Paloteo de " & IIf(optOrigen(0).value = True, "Producción", IIf(optOrigen(1).value = True, "Ventas", IIf(optOrigen(2).value = True, "Cortesías", IIf(optOrigen(3).value = True, "Cuentas Corrientes", IIf(optOrigen(6).value = True, "Pedidos Facturados", IIf(optOrigen(8).value = True, "Pedidos Cancelados como Cortesia", "Combinaciones")))))) & " Agrupado por Local/Salon"
      Detallado.Text14.SetText sRazonSocial
      Detallado.ReportTitle = IIf(chkTurno.value = 1, "Todos los Turnos", "Turno : " & sTurno) & sTitulo & Chr(13) & sTexto & " Expresado en " & sMonedaN
      Detallado.Text9.SetText sTexto2
      frmEmite.CRViewer.ReportSource = Detallado
   End If
   
   If optOpcion(1).value = True Then
      Resumido.DiscardSavedData
      Resumido.Text14.SetText localConectado
      Resumido.Database.SetDataSource rsReporte
      Resumido.Text10.SetText "Paloteo de " & IIf(optOrigen(0).value = True, "Producción", IIf(optOrigen(1).value = True, "Ventas", IIf(optOrigen(2).value = True, "Cortesías", IIf(optOrigen(3).value = True, "Cuentas Corrientes", IIf(optOrigen(6).value = True, "Pedidos Facturados", IIf(optOrigen(8).value = True, "Pedidos Cancelados como Cortesia", "Combinaciones")))))) & " Agrupado por Grupo/SubGrupo"
      Resumido.Text13.SetText sRazonSocial
      Resumido.ReportTitle = IIf(chkTurno.value = 1, "Todos los Turnos", "Turno : " & sTurno) & sTitulo & Chr(13) & sTexto & " Expresado en " & sMonedaN
      Resumido.Text9.SetText sTexto2
      frmEmite.CRViewer.ReportSource = Resumido
   End If
   
   If optOpcion(2).value = True Then
      Pormenorizado.DiscardSavedData
      Pormenorizado.Database.SetDataSource rsReporte
      Pormenorizado.Text7.SetText localConectado
      Pormenorizado.Text13.SetText sRazonSocial
      Pormenorizado.Text10.SetText "Paloteo de " & IIf(optOrigen(0).value = True, "Producción", IIf(optOrigen(1).value = True, "Ventas", IIf(optOrigen(2).value = True, "Cortesías", IIf(optOrigen(3).value = True, "Cuentas Corrientes", IIf(optOrigen(6).value = True, "Pedidos Facturados", IIf(optOrigen(8).value = True, "Pedidos Cancelados como Cortesia", "Combinaciones")))))) & " Agrupado Por Pedidos"
      Pormenorizado.Text4.SetText "Pedido"
      Pormenorizado.Text6.SetText "Documento"
      Pormenorizado.ReportTitle = IIf(chkTurno.value = 1, "Todos los Turnos", "Turno : " & sTurno) & sTitulo & Chr(13) & sTexto & " Expresado en " & sMonedaN
      Pormenorizado.Text9.SetText sTexto2
      frmEmite.CRViewer.ReportSource = Pormenorizado
   End If
            
   If optOpcion(3).value = True Then
      PorArea.DiscardSavedData
      PorArea.Text7.SetText localConectado
      PorArea.Database.SetDataSource rsReporte
      PorArea.Text10.SetText "Paloteo de " & IIf(optOrigen(0).value = True, "Producción", IIf(optOrigen(1).value = True, "Ventas", IIf(optOrigen(2).value = True, "Cortesías", IIf(optOrigen(3).value = True, "Cuentas Corrientes", IIf(optOrigen(6).value = True, "Pedidos Facturados", IIf(optOrigen(8).value = True, "Pedidos Cancelados como Cortesia", "Combinaciones")))))) & " Agrupado por Areas de Producción"
      PorArea.Text13.SetText sRazonSocial
      PorArea.ReportTitle = IIf(chkTurno.value = 1, "Todos los Turnos", "Turno : " & sTurno) & sTitulo & Chr(13) & sTexto & " Expresado en " & sMonedaN
      PorArea.Text9.SetText sTexto2
      frmEmite.CRViewer.ReportSource = PorArea
   End If
      
   If optOpcion(4).value = True Then
      PorArea.DiscardSavedData
      PorArea.Text7.SetText localConectado
      PorArea.Database.SetDataSource rsReporte
      PorArea.Text10.SetText "Paloteo de " & IIf(optOrigen(0).value = True, "Producción", IIf(optOrigen(1).value = True, "Ventas", IIf(optOrigen(2).value = True, "Cortesías", IIf(optOrigen(3).value = True, "Cuentas Corrientes", IIf(optOrigen(6).value = True, "Pedidos Facturados", IIf(optOrigen(8).value = True, "Pedidos Cancelados como Cortesia", "Combinaciones")))))) & " Agrupado por Tipo de Pedidos"
      PorArea.Text13.SetText sRazonSocial
      PorArea.ReportTitle = IIf(chkTurno.value = 1, "Todos los Turnos", "Turno : " & sTurno) & sTitulo & Chr(13) & sTexto & " Expresado en " & sMonedaN
      PorArea.Text9.SetText sTexto2
      frmEmite.CRViewer.ReportSource = PorArea
   End If
   
   If optOpcion(5).value = True Then
      PorProducto.DiscardSavedData
      PorProducto.Text6.SetText localConectado
      PorProducto.Database.SetDataSource rsReporte
      PorProducto.Text10.SetText "Paloteo de " & IIf(optOrigen(0).value = True, "Producción", IIf(optOrigen(1).value = True, "Ventas", IIf(optOrigen(2).value = True, "Cortesías", IIf(optOrigen(3).value = True, "Cuentas Corrientes", IIf(optOrigen(6).value = True, "Pedidos Facturados", IIf(optOrigen(8).value = True, "Pedidos Cancelados como Cortesia", "Combinaciones")))))) & " Agrupado por Producto"
      PorProducto.Text13.SetText sRazonSocial
      PorProducto.ReportTitle = IIf(chkTurno.value = 1, "Todos los Turnos", "Turno : " & sTurno) & sTitulo & Chr(13) & sTexto & " Expresado en " & sMonedaN
      PorProducto.Text9.SetText sTexto2
      frmEmite.CRViewer.ReportSource = PorProducto
      
   End If

   If optOpcion(6).value = True Then
      PorArea.DiscardSavedData
      PorArea.Text7.SetText localConectado
      PorArea.Database.SetDataSource rsReporte
      PorArea.Text10.SetText "Paloteo de " & IIf(optOrigen(0).value = True, "Producción", IIf(optOrigen(1).value = True, "Ventas", IIf(optOrigen(2).value = True, "Cortesías", IIf(optOrigen(3).value = True, "Cuentas Corrientes", IIf(optOrigen(6).value = True, "Pedidos Facturados", IIf(optOrigen(8).value = True, "Pedidos Cancelados como Cortesia", "Combinaciones")))))) & " Agrupado por Criterios de Agrupación"
      PorArea.Text13.SetText sRazonSocial
      PorArea.ReportTitle = IIf(chkTurno.value = 1, "Todos los Turnos", "Turno : " & sTurno) & sTitulo & Chr(13) & sTexto & " Expresado en " & sMonedaN
      PorArea.Text9.SetText sTexto2
      frmEmite.CRViewer.ReportSource = PorArea
   End If
   
   If optOpcion(7).value = True Then
        If optEmision(0).value = True Then
            Comparativo.DiscardSavedData
            Comparativo.Text14.SetText localConectado
            Comparativo.Database.SetDataSource rsReporte
            Comparativo.Text10.SetText "Paloteo de " & IIf(optOrigen(0).value = True, "Producción", IIf(optOrigen(1).value = True, "Ventas", IIf(optOrigen(2).value = True, "Cortesías", IIf(optOrigen(3).value = True, "Cuentas Corrientes", IIf(optOrigen(6).value = True, "Pedidos Facturados", IIf(optOrigen(8).value = True, "Pedidos Cancelados como Cortesia", "Combinaciones")))))) & " Comparativos"
            Comparativo.Text13.SetText sRazonSocial
            Comparativo.ReportTitle = IIf(chkTurno.value = 1, "Todos los Turnos", "Turno : " & sTurno) & sTitulo & Chr(13) & sTexto & " Expresado en " & sMonedaN
            frmEmite.CRViewer.ReportSource = Comparativo
        Else
            ComparativoResumido.DiscardSavedData
            ComparativoResumido.Text14.SetText localConectado
            ComparativoResumido.Database.SetDataSource rsReporte
            'ComparativoResumido.Text10.SetText "Paloteo de " & IIf(optOrigen(0).value = True, "Producción", IIf(optOrigen(1).value = True, "Ventas", IIf(optOrigen(2).value = True, "Cortesías", IIf(optOrigen(3).value = True, "Cuentas Corrientes", IIf(optOrigen(6).value = True, "Pedidos Facturados", IIf(optOrigen(8).value = True, "Pedidos Cancelados como Cortesia", "Combinaciones")))))) & " Comparativos"
            ComparativoResumido.Text13.SetText sRazonSocial
            ComparativoResumido.ReportTitle = IIf(chkTurno.value = 1, "Todos los Turnos", "Turno : " & sTurno) & sTitulo & Chr(13) & sTexto & " Expresado en " & sMonedaN
            frmEmite.CRViewer.ReportSource = ComparativoResumido
        End If
   End If
   
End Sub

Private Sub optOpcion_Click(Index As Integer)
   If Index = 7 Then
      optValor(0).Enabled = False
      optValor(1).Enabled = False
      optValor(2).Enabled = False
      'REPORTE TANTA
      optEmision(0).Enabled = True
      optEmision(1).Enabled = True
   Else
      optValor(0).Enabled = True
      optValor(1).Enabled = True
      optValor(2).Enabled = True
      'REPORTE TANTA
      optEmision(0).Enabled = False
      optEmision(1).Enabled = False
   End If
End Sub

