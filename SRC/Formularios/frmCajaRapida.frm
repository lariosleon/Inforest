VERSION 5.00
Object = "{562E3E04-2C31-4ECE-83F4-4017EEE51D40}#8.0#0"; "todg8.ocx"
Begin VB.Form frmCajaRapida 
   BorderStyle     =   0  'None
   Caption         =   "Caja Rápida"
   ClientHeight    =   9225
   ClientLeft      =   0
   ClientTop       =   -105
   ClientWidth     =   12285
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   9225
   ScaleWidth      =   12285
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame FrameFeSpring 
      BackColor       =   &H00C0E0FF&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1680
      Left            =   3360
      TabIndex        =   340
      Top             =   3315
      Visible         =   0   'False
      Width           =   6315
      Begin VB.Image imgProceso 
         Height          =   240
         Index           =   0
         Left            =   840
         Picture         =   "frmCajaRapida.frx":0000
         Top             =   1155
         Width           =   240
      End
      Begin VB.Label lblPaso2 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H00000000&
         BackStyle       =   0  'Transparent
         Caption         =   "Obteniendo codigo XXXX almacenado."
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   210
         Left            =   1185
         TabIndex        =   343
         Top             =   1155
         Visible         =   0   'False
         Width           =   3105
      End
      Begin VB.Label lblPaso1 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H00000000&
         BackStyle       =   0  'Transparent
         Caption         =   "Enviando información de documento a XXXX."
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   210
         Left            =   1170
         TabIndex        =   342
         Top             =   870
         Visible         =   0   'False
         Width           =   3660
      End
      Begin VB.Label Label4 
         Appearance      =   0  'Flat
         BackColor       =   &H00404080&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "  Facturación Electronica"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   300
         Left            =   15
         TabIndex        =   341
         Top             =   15
         Width           =   2490
      End
      Begin VB.Image imgProceso 
         Height          =   240
         Index           =   1
         Left            =   840
         Picture         =   "frmCajaRapida.frx":0213
         Top             =   855
         Width           =   240
      End
      Begin VB.Image imgProceso 
         Height          =   240
         Index           =   2
         Left            =   840
         Picture         =   "frmCajaRapida.frx":0426
         Top             =   855
         Width           =   240
      End
      Begin VB.Image imgProceso 
         Height          =   240
         Index           =   3
         Left            =   840
         Picture         =   "frmCajaRapida.frx":0768
         Top             =   1140
         Width           =   240
      End
      Begin VB.Label Label5 
         BackColor       =   &H00C0FFFF&
         Caption         =   "   Proceso de envio de documento a XXXXX......."
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   12.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00004000&
         Height          =   1110
         Left            =   210
         TabIndex        =   344
         Top             =   435
         Width           =   5910
      End
   End
   Begin VB.CommandButton cmdCabecera 
      Caption         =   "Transferencia Pedidos"
      Height          =   630
      Index           =   7
      Left            =   1365
      TabIndex        =   335
      Top             =   5770
      Width           =   1230
   End
   Begin VB.Frame tabProducto 
      Height          =   7665
      Left            =   6720
      TabIndex        =   209
      Top             =   1320
      Width           =   5280
      Begin VB.CommandButton cmdBoton 
         Caption         =   "4"
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   720
         Index           =   4
         Left            =   2280
         Style           =   1  'Graphical
         TabIndex        =   264
         TabStop         =   0   'False
         Top             =   3015
         Width           =   720
      End
      Begin VB.CommandButton cmdBoton 
         Caption         =   "5"
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   720
         Index           =   5
         Left            =   3015
         Style           =   1  'Graphical
         TabIndex        =   263
         TabStop         =   0   'False
         Top             =   3015
         Width           =   720
      End
      Begin VB.CommandButton cmdBoton 
         Caption         =   "6"
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   720
         Index           =   6
         Left            =   3750
         Style           =   1  'Graphical
         TabIndex        =   262
         TabStop         =   0   'False
         Top             =   3015
         Width           =   720
      End
      Begin VB.CommandButton cmdBoton 
         Caption         =   "7"
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   720
         Index           =   7
         Left            =   4485
         Style           =   1  'Graphical
         TabIndex        =   261
         TabStop         =   0   'False
         Top             =   3015
         Width           =   720
      End
      Begin VB.CommandButton cmdBoton 
         Caption         =   "8"
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   720
         Index           =   8
         Left            =   90
         Style           =   1  'Graphical
         TabIndex        =   260
         TabStop         =   0   'False
         Top             =   3780
         Width           =   720
      End
      Begin VB.CommandButton cmdBoton 
         Caption         =   "9"
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   720
         Index           =   9
         Left            =   825
         Style           =   1  'Graphical
         TabIndex        =   259
         TabStop         =   0   'False
         Top             =   3780
         Width           =   720
      End
      Begin VB.CommandButton cmdBoton 
         Caption         =   "10"
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   720
         Index           =   10
         Left            =   1560
         Style           =   1  'Graphical
         TabIndex        =   258
         TabStop         =   0   'False
         Top             =   3780
         Width           =   720
      End
      Begin VB.CommandButton cmdBoton 
         Caption         =   "11"
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   720
         Index           =   11
         Left            =   2280
         Style           =   1  'Graphical
         TabIndex        =   257
         TabStop         =   0   'False
         Top             =   3780
         Width           =   720
      End
      Begin VB.CommandButton cmdBoton 
         Caption         =   "12"
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   720
         Index           =   12
         Left            =   3015
         Style           =   1  'Graphical
         TabIndex        =   256
         TabStop         =   0   'False
         Top             =   3780
         Width           =   720
      End
      Begin VB.CommandButton cmdBoton 
         Caption         =   "13"
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   720
         Index           =   13
         Left            =   3750
         Style           =   1  'Graphical
         TabIndex        =   255
         TabStop         =   0   'False
         Top             =   3780
         Width           =   720
      End
      Begin VB.CommandButton cmdBoton 
         Caption         =   "14"
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   720
         Index           =   14
         Left            =   4485
         Style           =   1  'Graphical
         TabIndex        =   254
         TabStop         =   0   'False
         Top             =   3780
         Width           =   720
      End
      Begin VB.CommandButton cmdBoton 
         Caption         =   "15"
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   720
         Index           =   15
         Left            =   90
         Style           =   1  'Graphical
         TabIndex        =   253
         TabStop         =   0   'False
         Top             =   4530
         Width           =   720
      End
      Begin VB.CommandButton cmdBoton 
         Caption         =   "16"
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   720
         Index           =   16
         Left            =   825
         Style           =   1  'Graphical
         TabIndex        =   252
         TabStop         =   0   'False
         Top             =   4530
         Width           =   720
      End
      Begin VB.CommandButton cmdBoton 
         Caption         =   "17"
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   720
         Index           =   17
         Left            =   1560
         Style           =   1  'Graphical
         TabIndex        =   251
         TabStop         =   0   'False
         Top             =   4530
         Width           =   720
      End
      Begin VB.CommandButton cmdBoton 
         Caption         =   "18"
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   720
         Index           =   18
         Left            =   2280
         Style           =   1  'Graphical
         TabIndex        =   250
         TabStop         =   0   'False
         Top             =   4530
         Width           =   720
      End
      Begin VB.CommandButton cmdBoton 
         Caption         =   "19"
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   720
         Index           =   19
         Left            =   3015
         Style           =   1  'Graphical
         TabIndex        =   249
         TabStop         =   0   'False
         Top             =   4530
         Width           =   720
      End
      Begin VB.CommandButton cmdBoton 
         Caption         =   "20"
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   720
         Index           =   20
         Left            =   3750
         Style           =   1  'Graphical
         TabIndex        =   248
         TabStop         =   0   'False
         Top             =   4530
         Width           =   720
      End
      Begin VB.CommandButton cmdBoton 
         Caption         =   "3"
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   720
         Index           =   3
         Left            =   1560
         Style           =   1  'Graphical
         TabIndex        =   247
         TabStop         =   0   'False
         Top             =   3015
         Width           =   720
      End
      Begin VB.CommandButton cmdBoton 
         Caption         =   "2"
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   720
         Index           =   2
         Left            =   825
         Style           =   1  'Graphical
         TabIndex        =   246
         TabStop         =   0   'False
         Top             =   3015
         Width           =   720
      End
      Begin VB.CommandButton cmdBoton 
         Caption         =   "1"
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   720
         Index           =   1
         Left            =   90
         Style           =   1  'Graphical
         TabIndex        =   245
         TabStop         =   0   'False
         Top             =   3015
         Width           =   720
      End
      Begin VB.CommandButton cmdBoton 
         Caption         =   "21"
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   720
         Index           =   21
         Left            =   4500
         Style           =   1  'Graphical
         TabIndex        =   244
         TabStop         =   0   'False
         Top             =   4530
         Width           =   720
      End
      Begin VB.CommandButton cmdBoton 
         Caption         =   "22"
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   720
         Index           =   22
         Left            =   90
         Style           =   1  'Graphical
         TabIndex        =   243
         TabStop         =   0   'False
         Top             =   5295
         Width           =   720
      End
      Begin VB.CommandButton cmdBoton 
         Caption         =   "23"
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   720
         Index           =   23
         Left            =   825
         Style           =   1  'Graphical
         TabIndex        =   242
         TabStop         =   0   'False
         Top             =   5295
         Width           =   720
      End
      Begin VB.CommandButton cmdBoton 
         Caption         =   "24"
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   720
         Index           =   24
         Left            =   1560
         Style           =   1  'Graphical
         TabIndex        =   241
         TabStop         =   0   'False
         Top             =   5295
         Width           =   720
      End
      Begin VB.CommandButton cmdBoton 
         Caption         =   "25"
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   720
         Index           =   25
         Left            =   2280
         Style           =   1  'Graphical
         TabIndex        =   240
         TabStop         =   0   'False
         Top             =   5295
         Width           =   720
      End
      Begin VB.CommandButton cmdBoton 
         Caption         =   "26"
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   720
         Index           =   26
         Left            =   3015
         Style           =   1  'Graphical
         TabIndex        =   239
         TabStop         =   0   'False
         Top             =   5295
         Width           =   720
      End
      Begin VB.CommandButton cmdBoton 
         Caption         =   "27"
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   720
         Index           =   27
         Left            =   3750
         Style           =   1  'Graphical
         TabIndex        =   238
         TabStop         =   0   'False
         Top             =   5295
         Width           =   720
      End
      Begin VB.CommandButton cmdBoton 
         Caption         =   "28"
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   720
         Index           =   28
         Left            =   4485
         Style           =   1  'Graphical
         TabIndex        =   237
         TabStop         =   0   'False
         Top             =   5295
         Width           =   720
      End
      Begin VB.CommandButton cmdBoton 
         Caption         =   "29"
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   720
         Index           =   29
         Left            =   90
         Style           =   1  'Graphical
         TabIndex        =   236
         TabStop         =   0   'False
         Top             =   6060
         Width           =   720
      End
      Begin VB.CommandButton cmdBoton 
         Caption         =   "30"
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   720
         Index           =   30
         Left            =   825
         Style           =   1  'Graphical
         TabIndex        =   235
         TabStop         =   0   'False
         Top             =   6060
         Width           =   720
      End
      Begin VB.CommandButton cmdBoton 
         Caption         =   "31"
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   720
         Index           =   31
         Left            =   1560
         Style           =   1  'Graphical
         TabIndex        =   234
         TabStop         =   0   'False
         Top             =   6060
         Width           =   720
      End
      Begin VB.CommandButton cmdBoton 
         Caption         =   "32"
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   720
         Index           =   32
         Left            =   2280
         Style           =   1  'Graphical
         TabIndex        =   233
         TabStop         =   0   'False
         Top             =   6060
         Width           =   720
      End
      Begin VB.CommandButton cmdBoton 
         Caption         =   "33"
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   720
         Index           =   33
         Left            =   3015
         Style           =   1  'Graphical
         TabIndex        =   232
         TabStop         =   0   'False
         Top             =   6060
         Width           =   720
      End
      Begin VB.CommandButton cmdBoton 
         Caption         =   "34"
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   720
         Index           =   34
         Left            =   3750
         Style           =   1  'Graphical
         TabIndex        =   231
         TabStop         =   0   'False
         Top             =   6060
         Width           =   720
      End
      Begin VB.CommandButton cmdBoton 
         Caption         =   "35"
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   720
         Index           =   35
         Left            =   4485
         Style           =   1  'Graphical
         TabIndex        =   230
         TabStop         =   0   'False
         Top             =   6060
         Width           =   720
      End
      Begin VB.CommandButton cmdBoton 
         Caption         =   "36"
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   720
         Index           =   36
         Left            =   90
         Style           =   1  'Graphical
         TabIndex        =   229
         TabStop         =   0   'False
         Top             =   6825
         Width           =   720
      End
      Begin VB.CommandButton cmdBoton 
         Caption         =   "37"
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   720
         Index           =   37
         Left            =   825
         Style           =   1  'Graphical
         TabIndex        =   228
         TabStop         =   0   'False
         Top             =   6825
         Width           =   720
      End
      Begin VB.CommandButton cmdBoton 
         Caption         =   "38"
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   720
         Index           =   38
         Left            =   1560
         Style           =   1  'Graphical
         TabIndex        =   227
         TabStop         =   0   'False
         Top             =   6825
         Width           =   720
      End
      Begin VB.CommandButton cmdBoton 
         Caption         =   "39"
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   720
         Index           =   39
         Left            =   2280
         Style           =   1  'Graphical
         TabIndex        =   226
         TabStop         =   0   'False
         Top             =   6825
         Width           =   720
      End
      Begin VB.CommandButton cmdBoton 
         Caption         =   "40"
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   720
         Index           =   40
         Left            =   3015
         Style           =   1  'Graphical
         TabIndex        =   225
         TabStop         =   0   'False
         Top             =   6825
         Width           =   720
      End
      Begin VB.CommandButton cmdBoton 
         Caption         =   "41"
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   720
         Index           =   41
         Left            =   3750
         Style           =   1  'Graphical
         TabIndex        =   224
         TabStop         =   0   'False
         Top             =   6825
         Width           =   720
      End
      Begin VB.CommandButton cmdBoton 
         Caption         =   "42"
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   720
         Index           =   42
         Left            =   4485
         Style           =   1  'Graphical
         TabIndex        =   223
         TabStop         =   0   'False
         Top             =   6825
         Width           =   720
      End
      Begin VB.CommandButton cmdEtiqueta 
         Caption         =   "Command1"
         Height          =   645
         Index           =   3
         Left            =   3540
         Style           =   1  'Graphical
         TabIndex        =   222
         Top             =   225
         Width           =   1680
      End
      Begin VB.CommandButton cmdEtiqueta 
         Caption         =   "Command1"
         Height          =   645
         Index           =   2
         Left            =   1815
         Style           =   1  'Graphical
         TabIndex        =   221
         Top             =   225
         Width           =   1680
      End
      Begin VB.CommandButton cmdEtiqueta 
         Caption         =   "Command1"
         Height          =   645
         Index           =   1
         Left            =   90
         Style           =   1  'Graphical
         TabIndex        =   220
         Top             =   225
         Width           =   1680
      End
      Begin VB.CommandButton cmdAgrupacion 
         Caption         =   "Command1"
         Height          =   600
         Index           =   9
         Left            =   3015
         Style           =   1  'Graphical
         TabIndex        =   219
         Top             =   2295
         Width           =   1455
      End
      Begin VB.CommandButton cmdAgrupacion 
         Caption         =   "Command1"
         Height          =   600
         Index           =   8
         Left            =   1552
         Style           =   1  'Graphical
         TabIndex        =   218
         Top             =   2295
         Width           =   1455
      End
      Begin VB.CommandButton cmdAgrupacion 
         Caption         =   "Command1"
         Height          =   600
         Index           =   7
         Left            =   90
         Style           =   1  'Graphical
         TabIndex        =   217
         Top             =   2295
         Width           =   1455
      End
      Begin VB.CommandButton cmdAgrupacion 
         Caption         =   "Command1"
         Height          =   600
         Index           =   6
         Left            =   3015
         Style           =   1  'Graphical
         TabIndex        =   216
         Top             =   1635
         Width           =   1455
      End
      Begin VB.CommandButton cmdAgrupacion 
         Caption         =   "Command1"
         Height          =   600
         Index           =   5
         Left            =   1552
         Style           =   1  'Graphical
         TabIndex        =   215
         Top             =   1635
         Width           =   1455
      End
      Begin VB.CommandButton cmdAgrupacion 
         Caption         =   "Command1"
         Height          =   600
         Index           =   4
         Left            =   90
         Style           =   1  'Graphical
         TabIndex        =   214
         Top             =   1635
         Width           =   1455
      End
      Begin VB.CommandButton cmdAgrupacion 
         Caption         =   "Command1"
         Height          =   600
         Index           =   3
         Left            =   3015
         Style           =   1  'Graphical
         TabIndex        =   213
         Top             =   990
         Width           =   1455
      End
      Begin VB.CommandButton cmdAgrupacion 
         Caption         =   "Command1"
         Height          =   600
         Index           =   2
         Left            =   1552
         Style           =   1  'Graphical
         TabIndex        =   212
         Top             =   990
         Width           =   1455
      End
      Begin VB.CommandButton cmdAgrupacion 
         Height          =   600
         Index           =   1
         Left            =   90
         Style           =   1  'Graphical
         TabIndex        =   211
         Top             =   990
         Width           =   1455
      End
      Begin VB.CommandButton cmdSinBoton 
         Caption         =   "Búsq Rápida"
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   6.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1890
         Left            =   4500
         Style           =   1  'Graphical
         TabIndex        =   210
         TabStop         =   0   'False
         Top             =   990
         Width           =   720
      End
      Begin VB.Shape Shape6 
         Height          =   780
         Left            =   45
         Top             =   6795
         Width           =   5190
      End
      Begin VB.Shape Shape5 
         Height          =   780
         Left            =   45
         Top             =   5265
         Width           =   5190
      End
      Begin VB.Shape Shape4 
         Height          =   780
         Left            =   45
         Top             =   3735
         Width           =   5190
      End
      Begin VB.Shape Shape3 
         Height          =   4605
         Left            =   3735
         Top             =   2970
         Width           =   735
      End
      Begin VB.Shape Shape2 
         Height          =   4605
         Left            =   2250
         Top             =   2970
         Width           =   780
      End
      Begin VB.Shape Shape1 
         Height          =   4605
         Left            =   810
         Top             =   2970
         Width           =   735
      End
      Begin VB.Shape Shape 
         BorderColor     =   &H000000C0&
         BorderWidth     =   2
         Height          =   4605
         Left            =   45
         Top             =   2970
         Width           =   5190
      End
   End
   Begin VB.Frame Frame1 
      Height          =   8895
      Left            =   0
      TabIndex        =   1
      Top             =   240
      Width           =   12135
      Begin VB.CommandButton cmdNotasCredito 
         Caption         =   "Nota Credito"
         Height          =   630
         Index           =   50
         Left            =   90
         TabIndex        =   387
         Top             =   7520
         Width           =   1230
      End
      Begin VB.CommandButton cmdCabecera 
         Caption         =   "Fecha Entrega"
         Height          =   630
         Index           =   6
         Left            =   90
         TabIndex        =   334
         Top             =   5540
         Width           =   1230
      End
      Begin VB.CommandButton cmdCabecera 
         Caption         =   "Punto de Venta"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   630
         Index           =   0
         Left            =   2640
         TabIndex        =   333
         Top             =   5540
         Visible         =   0   'False
         Width           =   1230
      End
      Begin VB.CommandButton cmdDetalle 
         BackColor       =   &H00FFFFC0&
         Caption         =   "Entregar A"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Index           =   14
         Left            =   3375
         Style           =   1  'Graphical
         TabIndex        =   330
         Top             =   520
         Width           =   1230
      End
      Begin VB.CommandButton cmdOpcion 
         Caption         =   "Cargos"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   630
         Index           =   14
         Left            =   3915
         Picture         =   "frmCajaRapida.frx":0AAA
         Style           =   1  'Graphical
         TabIndex        =   319
         Top             =   5540
         Visible         =   0   'False
         Width           =   1230
      End
      Begin VB.CommandButton cmdOpcion 
         BackColor       =   &H00FFC0C0&
         Caption         =   "Cliente Frecuente"
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
         Index           =   9
         Left            =   2100
         MaskColor       =   &H00FFFFFF&
         Style           =   1  'Graphical
         TabIndex        =   323
         Top             =   120
         Width           =   1230
      End
      Begin VB.CommandButton cmdDetalle 
         Caption         =   "Cuentas Corrientes"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   630
         Index           =   10
         Left            =   3915
         Style           =   1  'Graphical
         TabIndex        =   322
         Top             =   7520
         Width           =   1230
      End
      Begin VB.CommandButton cmdDetalle 
         Caption         =   "Descuento"
         Height          =   630
         Index           =   12
         Left            =   1365
         TabIndex        =   320
         Top             =   7520
         Width           =   1230
      End
      Begin VB.Frame fraProductoCombo 
         Caption         =   " Productos de Combos "
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
         Height          =   7650
         Left            =   6720
         TabIndex        =   265
         Top             =   1080
         Width           =   5235
         Begin VB.CommandButton cmdProductoCombo 
            BeginProperty Font 
               Name            =   "Small Fonts"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   720
            Index           =   43
            Left            =   180
            Style           =   1  'Graphical
            TabIndex        =   318
            Top             =   6020
            Width           =   720
         End
         Begin VB.CommandButton cmdProductoCombo 
            BeginProperty Font 
               Name            =   "Small Fonts"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   720
            Index           =   44
            Left            =   1008
            Style           =   1  'Graphical
            TabIndex        =   317
            Top             =   6020
            Width           =   720
         End
         Begin VB.CommandButton cmdProductoCombo 
            BeginProperty Font 
               Name            =   "Small Fonts"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   720
            Index           =   45
            Left            =   1836
            Style           =   1  'Graphical
            TabIndex        =   316
            Top             =   6020
            Width           =   720
         End
         Begin VB.CommandButton cmdProductoCombo 
            BeginProperty Font 
               Name            =   "Small Fonts"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   720
            Index           =   46
            Left            =   2664
            Style           =   1  'Graphical
            TabIndex        =   315
            Top             =   6020
            Width           =   720
         End
         Begin VB.CommandButton cmdProductoCombo 
            BeginProperty Font 
               Name            =   "Small Fonts"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   720
            Index           =   47
            Left            =   3492
            Style           =   1  'Graphical
            TabIndex        =   314
            Top             =   6020
            Width           =   720
         End
         Begin VB.CommandButton cmdProductoCombo 
            BeginProperty Font 
               Name            =   "Small Fonts"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   720
            Index           =   48
            Left            =   4320
            Style           =   1  'Graphical
            TabIndex        =   313
            Top             =   6020
            Width           =   720
         End
         Begin VB.CommandButton cmdBuscar 
            BeginProperty Font 
               Name            =   "Small Fonts"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   720
            Index           =   2
            Left            =   3492
            Picture         =   "frmCajaRapida.frx":0BAC
            Style           =   1  'Graphical
            TabIndex        =   312
            Top             =   6840
            Width           =   1575
         End
         Begin VB.CommandButton cmdProductoCombo 
            BeginProperty Font 
               Name            =   "Small Fonts"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   720
            Index           =   42
            Left            =   4320
            Style           =   1  'Graphical
            TabIndex        =   311
            Top             =   5205
            Width           =   720
         End
         Begin VB.CommandButton cmdProductoCombo 
            BeginProperty Font 
               Name            =   "Small Fonts"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   720
            Index           =   41
            Left            =   3492
            Style           =   1  'Graphical
            TabIndex        =   310
            Top             =   5205
            Width           =   720
         End
         Begin VB.CommandButton cmdProductoCombo 
            BeginProperty Font 
               Name            =   "Small Fonts"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   720
            Index           =   40
            Left            =   2664
            Style           =   1  'Graphical
            TabIndex        =   309
            Top             =   5205
            Width           =   720
         End
         Begin VB.CommandButton cmdProductoCombo 
            BeginProperty Font 
               Name            =   "Small Fonts"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   720
            Index           =   39
            Left            =   1836
            Style           =   1  'Graphical
            TabIndex        =   308
            Top             =   5205
            Width           =   720
         End
         Begin VB.CommandButton cmdProductoCombo 
            BeginProperty Font 
               Name            =   "Small Fonts"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   720
            Index           =   38
            Left            =   1008
            Style           =   1  'Graphical
            TabIndex        =   307
            Top             =   5205
            Width           =   720
         End
         Begin VB.CommandButton cmdProductoCombo 
            BeginProperty Font 
               Name            =   "Small Fonts"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   720
            Index           =   37
            Left            =   180
            Style           =   1  'Graphical
            TabIndex        =   306
            Top             =   5205
            Width           =   720
         End
         Begin VB.CommandButton cmdProductoCombo 
            BeginProperty Font 
               Name            =   "Small Fonts"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   720
            Index           =   36
            Left            =   4320
            Style           =   1  'Graphical
            TabIndex        =   305
            Top             =   4390
            Width           =   720
         End
         Begin VB.CommandButton cmdProductoCombo 
            BeginProperty Font 
               Name            =   "Small Fonts"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   720
            Index           =   35
            Left            =   3492
            Style           =   1  'Graphical
            TabIndex        =   304
            Top             =   4390
            Width           =   720
         End
         Begin VB.CommandButton cmdProductoCombo 
            BeginProperty Font 
               Name            =   "Small Fonts"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   720
            Index           =   34
            Left            =   2664
            Style           =   1  'Graphical
            TabIndex        =   303
            Top             =   4390
            Width           =   720
         End
         Begin VB.CommandButton cmdProductoCombo 
            BeginProperty Font 
               Name            =   "Small Fonts"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   720
            Index           =   33
            Left            =   1836
            Style           =   1  'Graphical
            TabIndex        =   302
            Top             =   4390
            Width           =   720
         End
         Begin VB.CommandButton cmdProductoCombo 
            BeginProperty Font 
               Name            =   "Small Fonts"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   720
            Index           =   32
            Left            =   1008
            Style           =   1  'Graphical
            TabIndex        =   301
            Top             =   4390
            Width           =   720
         End
         Begin VB.CommandButton cmdProductoCombo 
            BeginProperty Font 
               Name            =   "Small Fonts"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   720
            Index           =   31
            Left            =   180
            Style           =   1  'Graphical
            TabIndex        =   300
            Top             =   4390
            Width           =   720
         End
         Begin VB.CommandButton cmdProductoCombo 
            BeginProperty Font 
               Name            =   "Small Fonts"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   720
            Index           =   30
            Left            =   4320
            Style           =   1  'Graphical
            TabIndex        =   299
            Top             =   3575
            Width           =   720
         End
         Begin VB.CommandButton cmdProductoCombo 
            BeginProperty Font 
               Name            =   "Small Fonts"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   720
            Index           =   29
            Left            =   3492
            Style           =   1  'Graphical
            TabIndex        =   298
            Top             =   3575
            Width           =   720
         End
         Begin VB.CommandButton cmdProductoCombo 
            BeginProperty Font 
               Name            =   "Small Fonts"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   720
            Index           =   28
            Left            =   2664
            Style           =   1  'Graphical
            TabIndex        =   297
            Top             =   3575
            Width           =   720
         End
         Begin VB.CommandButton cmdProductoCombo 
            BeginProperty Font 
               Name            =   "Small Fonts"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   720
            Index           =   27
            Left            =   1836
            Style           =   1  'Graphical
            TabIndex        =   296
            Top             =   3575
            Width           =   720
         End
         Begin VB.CommandButton cmdProductoCombo 
            BeginProperty Font 
               Name            =   "Small Fonts"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   720
            Index           =   26
            Left            =   1008
            Style           =   1  'Graphical
            TabIndex        =   295
            Top             =   3575
            Width           =   720
         End
         Begin VB.CommandButton cmdProductoCombo 
            BeginProperty Font 
               Name            =   "Small Fonts"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   720
            Index           =   25
            Left            =   180
            Style           =   1  'Graphical
            TabIndex        =   294
            Top             =   3575
            Width           =   720
         End
         Begin VB.CommandButton cmdProductoCombo 
            BeginProperty Font 
               Name            =   "Small Fonts"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   720
            Index           =   24
            Left            =   4320
            Style           =   1  'Graphical
            TabIndex        =   293
            Top             =   2760
            Width           =   720
         End
         Begin VB.CommandButton cmdProductoCombo 
            BeginProperty Font 
               Name            =   "Small Fonts"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   720
            Index           =   23
            Left            =   3492
            Style           =   1  'Graphical
            TabIndex        =   292
            Top             =   2760
            Width           =   720
         End
         Begin VB.CommandButton cmdProductoCombo 
            BeginProperty Font 
               Name            =   "Small Fonts"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   720
            Index           =   22
            Left            =   2664
            Style           =   1  'Graphical
            TabIndex        =   291
            Top             =   2760
            Width           =   720
         End
         Begin VB.CommandButton cmdProductoCombo 
            BeginProperty Font 
               Name            =   "Small Fonts"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   720
            Index           =   21
            Left            =   1836
            Style           =   1  'Graphical
            TabIndex        =   290
            Top             =   2760
            Width           =   720
         End
         Begin VB.CommandButton cmdProductoCombo 
            BeginProperty Font 
               Name            =   "Small Fonts"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   720
            Index           =   20
            Left            =   1008
            Style           =   1  'Graphical
            TabIndex        =   289
            Top             =   2760
            Width           =   720
         End
         Begin VB.CommandButton cmdProductoCombo 
            BeginProperty Font 
               Name            =   "Small Fonts"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   720
            Index           =   19
            Left            =   180
            Style           =   1  'Graphical
            TabIndex        =   288
            Top             =   2760
            Width           =   720
         End
         Begin VB.CommandButton cmdProductoCombo 
            BeginProperty Font 
               Name            =   "Small Fonts"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   720
            Index           =   18
            Left            =   4320
            Style           =   1  'Graphical
            TabIndex        =   287
            Top             =   1945
            Width           =   720
         End
         Begin VB.CommandButton cmdProductoCombo 
            BeginProperty Font 
               Name            =   "Small Fonts"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   720
            Index           =   17
            Left            =   3492
            Style           =   1  'Graphical
            TabIndex        =   286
            Top             =   1945
            Width           =   720
         End
         Begin VB.CommandButton cmdProductoCombo 
            BeginProperty Font 
               Name            =   "Small Fonts"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   720
            Index           =   16
            Left            =   2664
            Style           =   1  'Graphical
            TabIndex        =   285
            Top             =   1945
            Width           =   720
         End
         Begin VB.CommandButton cmdProductoCombo 
            BeginProperty Font 
               Name            =   "Small Fonts"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   720
            Index           =   15
            Left            =   1836
            Style           =   1  'Graphical
            TabIndex        =   284
            Top             =   1945
            Width           =   720
         End
         Begin VB.CommandButton cmdProductoCombo 
            BeginProperty Font 
               Name            =   "Small Fonts"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   720
            Index           =   14
            Left            =   1008
            Style           =   1  'Graphical
            TabIndex        =   283
            Top             =   1945
            Width           =   720
         End
         Begin VB.CommandButton cmdProductoCombo 
            BeginProperty Font 
               Name            =   "Small Fonts"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   720
            Index           =   13
            Left            =   180
            Style           =   1  'Graphical
            TabIndex        =   282
            Top             =   1945
            Width           =   720
         End
         Begin VB.CommandButton cmdProductoCombo 
            BeginProperty Font 
               Name            =   "Small Fonts"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   720
            Index           =   12
            Left            =   4320
            Style           =   1  'Graphical
            TabIndex        =   281
            Top             =   1130
            Width           =   720
         End
         Begin VB.CommandButton cmdProductoCombo 
            BeginProperty Font 
               Name            =   "Small Fonts"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   720
            Index           =   11
            Left            =   3492
            Style           =   1  'Graphical
            TabIndex        =   280
            Top             =   1130
            Width           =   720
         End
         Begin VB.CommandButton cmdProductoCombo 
            BeginProperty Font 
               Name            =   "Small Fonts"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   720
            Index           =   10
            Left            =   2664
            Style           =   1  'Graphical
            TabIndex        =   279
            Top             =   1130
            Width           =   720
         End
         Begin VB.CommandButton cmdProductoCombo 
            BeginProperty Font 
               Name            =   "Small Fonts"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   720
            Index           =   9
            Left            =   1836
            Style           =   1  'Graphical
            TabIndex        =   278
            Top             =   1130
            Width           =   720
         End
         Begin VB.CommandButton cmdProductoCombo 
            BeginProperty Font 
               Name            =   "Small Fonts"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   720
            Index           =   8
            Left            =   1008
            Style           =   1  'Graphical
            TabIndex        =   277
            Top             =   1130
            Width           =   720
         End
         Begin VB.CommandButton cmdProductoCombo 
            BeginProperty Font 
               Name            =   "Small Fonts"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   720
            Index           =   7
            Left            =   180
            Style           =   1  'Graphical
            TabIndex        =   276
            Top             =   1130
            Width           =   720
         End
         Begin VB.CommandButton cmdProductoCombo 
            BeginProperty Font 
               Name            =   "Small Fonts"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   720
            Index           =   6
            Left            =   4320
            Style           =   1  'Graphical
            TabIndex        =   275
            Top             =   315
            Width           =   720
         End
         Begin VB.CommandButton cmdProductoCombo 
            BeginProperty Font 
               Name            =   "Small Fonts"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   720
            Index           =   5
            Left            =   3492
            Style           =   1  'Graphical
            TabIndex        =   274
            Top             =   315
            Width           =   720
         End
         Begin VB.CommandButton cmdProductoCombo 
            BeginProperty Font 
               Name            =   "Small Fonts"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   720
            Index           =   4
            Left            =   2664
            Style           =   1  'Graphical
            TabIndex        =   273
            Top             =   315
            Width           =   720
         End
         Begin VB.CommandButton cmdProductoCombo 
            BeginProperty Font 
               Name            =   "Small Fonts"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   720
            Index           =   3
            Left            =   1836
            Style           =   1  'Graphical
            TabIndex        =   272
            Top             =   315
            Width           =   720
         End
         Begin VB.CommandButton cmdProductoCombo 
            BeginProperty Font 
               Name            =   "Small Fonts"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   720
            Index           =   2
            Left            =   1008
            Style           =   1  'Graphical
            TabIndex        =   271
            Top             =   315
            Width           =   720
         End
         Begin VB.CommandButton cmdProductoCombo 
            BeginProperty Font 
               Name            =   "Small Fonts"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   720
            Index           =   1
            Left            =   180
            Style           =   1  'Graphical
            TabIndex        =   270
            Top             =   315
            Width           =   720
         End
         Begin VB.CommandButton cmdProductoCombo 
            BeginProperty Font 
               Name            =   "Small Fonts"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   720
            Index           =   49
            Left            =   180
            Style           =   1  'Graphical
            TabIndex        =   269
            Top             =   6840
            Width           =   720
         End
         Begin VB.CommandButton cmdProductoCombo 
            BeginProperty Font 
               Name            =   "Small Fonts"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   720
            Index           =   50
            Left            =   1008
            Style           =   1  'Graphical
            TabIndex        =   268
            Top             =   6840
            Width           =   720
         End
         Begin VB.CommandButton cmdProductoCombo 
            BeginProperty Font 
               Name            =   "Small Fonts"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   720
            Index           =   51
            Left            =   1836
            Style           =   1  'Graphical
            TabIndex        =   267
            Top             =   6840
            Width           =   720
         End
         Begin VB.CommandButton cmdProductoCombo 
            BeginProperty Font 
               Name            =   "Small Fonts"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   720
            Index           =   52
            Left            =   2664
            Style           =   1  'Graphical
            TabIndex        =   266
            Top             =   6840
            Width           =   720
         End
      End
      Begin VB.Frame fraEliminacion 
         Caption         =   " Motivo de Eliminación "
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
         Height          =   7650
         Left            =   6720
         TabIndex        =   122
         Top             =   1080
         Width           =   5220
         Begin VB.CommandButton cmdEliminacion 
            BeginProperty Font 
               Name            =   "Small Fonts"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   810
            Index           =   38
            Left            =   2130
            TabIndex        =   161
            Top             =   6630
            Width           =   810
         End
         Begin VB.CommandButton cmdEliminacion 
            BeginProperty Font 
               Name            =   "Small Fonts"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   810
            Index           =   37
            Left            =   1140
            TabIndex        =   160
            Top             =   6630
            Width           =   810
         End
         Begin VB.CommandButton cmdEliminacion 
            BeginProperty Font 
               Name            =   "Small Fonts"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   810
            Index           =   36
            Left            =   150
            TabIndex        =   159
            Top             =   6630
            Width           =   810
         End
         Begin VB.CommandButton cmdEliminacion 
            BeginProperty Font 
               Name            =   "Small Fonts"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   810
            Index           =   35
            Left            =   4110
            TabIndex        =   158
            Top             =   5730
            Width           =   810
         End
         Begin VB.CommandButton cmdEliminacion 
            BeginProperty Font 
               Name            =   "Small Fonts"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   810
            Index           =   31
            Left            =   150
            TabIndex        =   157
            Top             =   5730
            Width           =   810
         End
         Begin VB.CommandButton cmdEliminacion 
            BeginProperty Font 
               Name            =   "Small Fonts"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   810
            Index           =   32
            Left            =   1140
            TabIndex        =   156
            Top             =   5730
            Width           =   810
         End
         Begin VB.CommandButton cmdEliminacion 
            BeginProperty Font 
               Name            =   "Small Fonts"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   810
            Index           =   33
            Left            =   2130
            TabIndex        =   155
            Top             =   5730
            Width           =   810
         End
         Begin VB.CommandButton cmdEliminacion 
            BeginProperty Font 
               Name            =   "Small Fonts"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   810
            Index           =   34
            Left            =   3120
            TabIndex        =   154
            Top             =   5730
            Width           =   810
         End
         Begin VB.CommandButton cmdEliminacion 
            BeginProperty Font 
               Name            =   "Small Fonts"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   810
            Index           =   30
            Left            =   4110
            TabIndex        =   153
            Top             =   4830
            Width           =   810
         End
         Begin VB.CommandButton cmdEliminacion 
            BeginProperty Font 
               Name            =   "Small Fonts"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   810
            Index           =   26
            Left            =   150
            TabIndex        =   152
            Top             =   4830
            Width           =   810
         End
         Begin VB.CommandButton cmdEliminacion 
            BeginProperty Font 
               Name            =   "Small Fonts"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   810
            Index           =   27
            Left            =   1140
            TabIndex        =   151
            Top             =   4830
            Width           =   810
         End
         Begin VB.CommandButton cmdEliminacion 
            BeginProperty Font 
               Name            =   "Small Fonts"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   810
            Index           =   28
            Left            =   2130
            TabIndex        =   150
            Top             =   4830
            Width           =   810
         End
         Begin VB.CommandButton cmdEliminacion 
            BeginProperty Font 
               Name            =   "Small Fonts"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   810
            Index           =   29
            Left            =   3120
            TabIndex        =   149
            Top             =   4830
            Width           =   810
         End
         Begin VB.CommandButton cmdEliminacion 
            BeginProperty Font 
               Name            =   "Small Fonts"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   810
            Index           =   25
            Left            =   4110
            TabIndex        =   148
            Top             =   3930
            Width           =   810
         End
         Begin VB.CommandButton cmdEliminacion 
            BeginProperty Font 
               Name            =   "Small Fonts"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   810
            Index           =   21
            Left            =   150
            TabIndex        =   147
            Top             =   3930
            Width           =   810
         End
         Begin VB.CommandButton cmdEliminacion 
            BeginProperty Font 
               Name            =   "Small Fonts"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   810
            Index           =   22
            Left            =   1140
            TabIndex        =   146
            Top             =   3930
            Width           =   810
         End
         Begin VB.CommandButton cmdEliminacion 
            BeginProperty Font 
               Name            =   "Small Fonts"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   810
            Index           =   23
            Left            =   2130
            TabIndex        =   145
            Top             =   3930
            Width           =   810
         End
         Begin VB.CommandButton cmdEliminacion 
            BeginProperty Font 
               Name            =   "Small Fonts"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   810
            Index           =   24
            Left            =   3120
            TabIndex        =   144
            Top             =   3930
            Width           =   810
         End
         Begin VB.CommandButton cmdEliminacion 
            BeginProperty Font 
               Name            =   "Small Fonts"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   810
            Index           =   20
            Left            =   4110
            TabIndex        =   143
            Top             =   3030
            Width           =   810
         End
         Begin VB.CommandButton cmdEliminacion 
            BeginProperty Font 
               Name            =   "Small Fonts"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   810
            Index           =   16
            Left            =   150
            TabIndex        =   142
            Top             =   3030
            Width           =   810
         End
         Begin VB.CommandButton cmdEliminacion 
            BeginProperty Font 
               Name            =   "Small Fonts"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   810
            Index           =   17
            Left            =   1140
            TabIndex        =   141
            Top             =   3030
            Width           =   810
         End
         Begin VB.CommandButton cmdEliminacion 
            BeginProperty Font 
               Name            =   "Small Fonts"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   810
            Index           =   18
            Left            =   2130
            TabIndex        =   140
            Top             =   3030
            Width           =   810
         End
         Begin VB.CommandButton cmdEliminacion 
            BeginProperty Font 
               Name            =   "Small Fonts"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   810
            Index           =   19
            Left            =   3120
            TabIndex        =   139
            Top             =   3030
            Width           =   810
         End
         Begin VB.CommandButton cmdEliminacion 
            BeginProperty Font 
               Name            =   "Small Fonts"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   810
            Index           =   15
            Left            =   4110
            TabIndex        =   138
            Top             =   2130
            Width           =   810
         End
         Begin VB.CommandButton cmdEliminacion 
            BeginProperty Font 
               Name            =   "Small Fonts"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   810
            Index           =   11
            Left            =   150
            TabIndex        =   137
            Top             =   2130
            Width           =   810
         End
         Begin VB.CommandButton cmdEliminacion 
            BeginProperty Font 
               Name            =   "Small Fonts"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   810
            Index           =   12
            Left            =   1140
            TabIndex        =   136
            Top             =   2130
            Width           =   810
         End
         Begin VB.CommandButton cmdEliminacion 
            BeginProperty Font 
               Name            =   "Small Fonts"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   810
            Index           =   13
            Left            =   2130
            TabIndex        =   135
            Top             =   2130
            Width           =   810
         End
         Begin VB.CommandButton cmdEliminacion 
            BeginProperty Font 
               Name            =   "Small Fonts"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   810
            Index           =   14
            Left            =   3120
            TabIndex        =   134
            Top             =   2130
            Width           =   810
         End
         Begin VB.CommandButton cmdEliminacion 
            BeginProperty Font 
               Name            =   "Small Fonts"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   810
            Index           =   10
            Left            =   4110
            TabIndex        =   133
            Top             =   1230
            Width           =   810
         End
         Begin VB.CommandButton cmdEliminacion 
            BeginProperty Font 
               Name            =   "Small Fonts"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   810
            Index           =   1
            Left            =   150
            TabIndex        =   132
            Top             =   330
            Width           =   810
         End
         Begin VB.CommandButton cmdEliminacion 
            BeginProperty Font 
               Name            =   "Small Fonts"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   810
            Index           =   2
            Left            =   1140
            TabIndex        =   131
            Top             =   330
            Width           =   810
         End
         Begin VB.CommandButton cmdEliminacion 
            BeginProperty Font 
               Name            =   "Small Fonts"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   810
            Index           =   3
            Left            =   2130
            TabIndex        =   130
            Top             =   330
            Width           =   810
         End
         Begin VB.CommandButton cmdEliminacion 
            BeginProperty Font 
               Name            =   "Small Fonts"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   810
            Index           =   4
            Left            =   3120
            TabIndex        =   129
            Top             =   330
            Width           =   810
         End
         Begin VB.CommandButton cmdEliminacion 
            BeginProperty Font 
               Name            =   "Small Fonts"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   810
            Index           =   5
            Left            =   4110
            TabIndex        =   128
            Top             =   330
            Width           =   810
         End
         Begin VB.CommandButton cmdEliminacion 
            BeginProperty Font 
               Name            =   "Small Fonts"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   810
            Index           =   6
            Left            =   150
            TabIndex        =   127
            Top             =   1230
            Width           =   810
         End
         Begin VB.CommandButton cmdEliminacion 
            BeginProperty Font 
               Name            =   "Small Fonts"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   810
            Index           =   7
            Left            =   1140
            TabIndex        =   126
            Top             =   1230
            Width           =   810
         End
         Begin VB.CommandButton cmdEliminacion 
            BeginProperty Font 
               Name            =   "Small Fonts"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   810
            Index           =   8
            Left            =   2130
            TabIndex        =   125
            Top             =   1230
            Width           =   810
         End
         Begin VB.CommandButton cmdEliminacion 
            BeginProperty Font 
               Name            =   "Small Fonts"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   810
            Index           =   9
            Left            =   3120
            TabIndex        =   124
            Top             =   1230
            Width           =   810
         End
         Begin VB.CommandButton cmdOpcion 
            Caption         =   "Cancelar"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   810
            Index           =   17
            Left            =   3150
            Picture         =   "frmCajaRapida.frx":0FEE
            Style           =   1  'Graphical
            TabIndex        =   123
            Top             =   6630
            Width           =   1785
         End
      End
      Begin VB.Frame fraCombo 
         Caption         =   " Combo"
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
         Height          =   4665
         Left            =   45
         TabIndex        =   3
         Top             =   840
         Visible         =   0   'False
         Width           =   6660
         Begin VB.CommandButton cmdCombo 
            Caption         =   "Propiedad y Observación"
            Height          =   630
            Index           =   5
            Left            =   5220
            TabIndex        =   18
            Top             =   2625
            Width           =   1230
         End
         Begin VB.CommandButton cmdCombo 
            Caption         =   "Cantidad"
            Height          =   630
            Index           =   4
            Left            =   5220
            TabIndex        =   17
            Top             =   105
            Width           =   1230
         End
         Begin VB.CommandButton cmdNavegar 
            Height          =   600
            Index           =   17
            Left            =   4440
            Picture         =   "frmCajaRapida.frx":1578
            Style           =   1  'Graphical
            TabIndex        =   16
            Top             =   790
            Width           =   615
         End
         Begin VB.CommandButton cmdNavegar 
            Height          =   600
            Index           =   16
            Left            =   4440
            Picture         =   "frmCajaRapida.frx":1E42
            Style           =   1  'Graphical
            TabIndex        =   15
            Top             =   2690
            Width           =   615
         End
         Begin VB.CommandButton cmdCombo 
            Caption         =   "&Aumentar"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   630
            Index           =   2
            Left            =   5220
            Picture         =   "frmCajaRapida.frx":270C
            Style           =   1  'Graphical
            TabIndex        =   14
            Top             =   735
            Width           =   1230
         End
         Begin VB.CommandButton cmdCombo 
            Caption         =   "&Disminuir"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   630
            Index           =   3
            Left            =   5220
            Picture         =   "frmCajaRapida.frx":280E
            Style           =   1  'Graphical
            TabIndex        =   13
            Top             =   1365
            Width           =   1230
         End
         Begin VB.CommandButton cmdNavegar 
            Height          =   600
            Index           =   12
            Left            =   4440
            Picture         =   "frmCajaRapida.frx":2910
            Style           =   1  'Graphical
            TabIndex        =   12
            Top             =   180
            Width           =   615
         End
         Begin VB.CommandButton cmdNavegar 
            Height          =   600
            Index           =   15
            Left            =   4440
            Picture         =   "frmCajaRapida.frx":31DA
            Style           =   1  'Graphical
            TabIndex        =   11
            Top             =   3320
            Width           =   615
         End
         Begin VB.CommandButton cmdNavegar 
            Height          =   600
            Index           =   13
            Left            =   4440
            Picture         =   "frmCajaRapida.frx":3AA4
            Style           =   1  'Graphical
            TabIndex        =   10
            Top             =   1420
            Width           =   615
         End
         Begin VB.CommandButton cmdNavegar 
            Height          =   600
            Index           =   14
            Left            =   4440
            Picture         =   "frmCajaRapida.frx":436E
            Style           =   1  'Graphical
            TabIndex        =   9
            Top             =   2055
            Width           =   615
         End
         Begin VB.CommandButton cmdCombo 
            Caption         =   "Elimina"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   630
            Index           =   1
            Left            =   5220
            Picture         =   "frmCajaRapida.frx":4C38
            Style           =   1  'Graphical
            TabIndex        =   8
            Top             =   1995
            Width           =   1230
         End
         Begin VB.CommandButton cmdCombo 
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
            Height          =   630
            Index           =   0
            Left            =   5220
            Picture         =   "frmCajaRapida.frx":4D3A
            Style           =   1  'Graphical
            TabIndex        =   7
            Top             =   3915
            Width           =   1230
         End
         Begin VB.CommandButton cmdCombo 
            Caption         =   "Orden"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   630
            Index           =   6
            Left            =   5070
            Picture         =   "frmCajaRapida.frx":4E2C
            Style           =   1  'Graphical
            TabIndex        =   6
            Top             =   3255
            Width           =   780
         End
         Begin VB.CommandButton cmdCombo 
            Caption         =   "Orden"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   630
            Index           =   7
            Left            =   5850
            Picture         =   "frmCajaRapida.frx":4F2E
            Style           =   1  'Graphical
            TabIndex        =   5
            Top             =   3255
            Width           =   780
         End
         Begin VB.CommandButton cmdCombo 
            Caption         =   "---"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   24
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   600
            Index           =   8
            Left            =   4440
            Style           =   1  'Graphical
            TabIndex        =   4
            Top             =   3940
            Width           =   615
         End
         Begin TrueOleDBGrid80.TDBGrid grdCombo 
            Height          =   4350
            Left            =   30
            TabIndex        =   19
            Top             =   210
            Width           =   4380
            _ExtentX        =   7726
            _ExtentY        =   7673
            _LayoutType     =   4
            _RowHeight      =   26
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
      End
      Begin VB.Frame fraItem 
         Height          =   4260
         Left            =   45
         TabIndex        =   54
         Top             =   855
         Width           =   5985
         Begin TrueOleDBGrid80.TDBGrid grdDetalle 
            Height          =   4005
            Left            =   60
            TabIndex        =   55
            Top             =   180
            Width           =   5865
            _ExtentX        =   10345
            _ExtentY        =   7064
            _LayoutType     =   4
            _RowHeight      =   26
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
            Caption         =   "Detalle a Facturar"
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
      End
      Begin VB.CommandButton cmdDetalle 
         Caption         =   "---"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   24
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   580
         Index           =   8
         Left            =   6075
         Style           =   1  'Graphical
         TabIndex        =   205
         Top             =   4500
         Width           =   615
      End
      Begin VB.Frame Frame3 
         Caption         =   " Tipo de Pedido "
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   690
         Left            =   6720
         TabIndex        =   57
         Top             =   420
         Width           =   5235
         Begin VB.CommandButton cmdCabecera 
            Caption         =   "En el &Local"
            Height          =   510
            Index           =   5
            Left            =   4170
            TabIndex        =   337
            Top             =   165
            Width           =   1020
         End
         Begin VB.CommandButton cmdCabecera 
            Caption         =   "En el &Local"
            Height          =   510
            Index           =   4
            Left            =   3140
            TabIndex        =   336
            Top             =   165
            Width           =   1020
         End
         Begin VB.CommandButton cmdCabecera 
            Caption         =   "En el &Local"
            Height          =   510
            Index           =   3
            Left            =   2110
            TabIndex        =   328
            Top             =   165
            Width           =   1020
         End
         Begin VB.CommandButton cmdCabecera 
            Caption         =   "&Para Llevar"
            Height          =   510
            Index           =   1
            Left            =   30
            TabIndex        =   59
            Top             =   165
            Width           =   1020
         End
         Begin VB.CommandButton cmdCabecera 
            Caption         =   "En el &Local"
            Height          =   510
            Index           =   2
            Left            =   1070
            TabIndex        =   58
            Top             =   165
            Width           =   1020
         End
      End
      Begin VB.CommandButton cmdDetalle 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   630
         Index           =   3
         Left            =   3915
         Picture         =   "frmCajaRapida.frx":5030
         Style           =   1  'Graphical
         TabIndex        =   53
         Top             =   6200
         Width           =   1230
      End
      Begin VB.CommandButton cmdDetalle 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   630
         Index           =   2
         Left            =   2640
         Picture         =   "frmCajaRapida.frx":5132
         Style           =   1  'Graphical
         TabIndex        =   52
         Top             =   6200
         Width           =   1230
      End
      Begin VB.CommandButton cmdDetalle 
         Caption         =   "Cantidad"
         Height          =   630
         Index           =   1
         Left            =   1365
         TabIndex        =   51
         Top             =   6200
         Width           =   1230
      End
      Begin VB.CommandButton cmdDetalle 
         Caption         =   "Elimina"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   630
         Index           =   0
         Left            =   90
         Picture         =   "frmCajaRapida.frx":5234
         Style           =   1  'Graphical
         TabIndex        =   50
         Top             =   6200
         Width           =   1230
      End
      Begin VB.CommandButton cmdNavegar 
         Height          =   580
         Index           =   6
         Left            =   6075
         Picture         =   "frmCajaRapida.frx":5336
         Style           =   1  'Graphical
         TabIndex        =   49
         Top             =   3900
         Width           =   615
      End
      Begin VB.CommandButton cmdNavegar 
         Height          =   580
         Index           =   5
         Left            =   6075
         Picture         =   "frmCajaRapida.frx":5C00
         Style           =   1  'Graphical
         TabIndex        =   48
         Top             =   3300
         Width           =   615
      End
      Begin VB.CommandButton cmdNavegar 
         Height          =   580
         Index           =   4
         Left            =   6075
         Picture         =   "frmCajaRapida.frx":64CA
         Style           =   1  'Graphical
         TabIndex        =   47
         Top             =   2700
         Width           =   615
      End
      Begin VB.CommandButton cmdNavegar 
         Height          =   580
         Index           =   3
         Left            =   6075
         Picture         =   "frmCajaRapida.frx":6D94
         Style           =   1  'Graphical
         TabIndex        =   46
         Top             =   2100
         Width           =   615
      End
      Begin VB.CommandButton cmdNavegar 
         Height          =   580
         Index           =   2
         Left            =   6075
         Picture         =   "frmCajaRapida.frx":765E
         Style           =   1  'Graphical
         TabIndex        =   45
         Top             =   1500
         Width           =   615
      End
      Begin VB.CommandButton cmdNavegar 
         Height          =   580
         Index           =   1
         Left            =   6075
         Picture         =   "frmCajaRapida.frx":7F28
         Style           =   1  'Graphical
         TabIndex        =   44
         Top             =   880
         Width           =   615
      End
      Begin VB.CommandButton cmdDetalle 
         BackColor       =   &H00C0FFFF&
         Caption         =   "Observación"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Index           =   6
         Left            =   3390
         Style           =   1  'Graphical
         TabIndex        =   43
         Top             =   120
         Width           =   1230
      End
      Begin VB.CommandButton cmdDetalle 
         Caption         =   "Propiedad  y Observación"
         Height          =   630
         Index           =   5
         Left            =   1365
         TabIndex        =   42
         Top             =   6860
         Width           =   1230
      End
      Begin VB.CommandButton cmdOpcion 
         Caption         =   "Cancelar"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   630
         Index           =   1
         Left            =   5295
         Picture         =   "frmCajaRapida.frx":87F2
         Style           =   1  'Graphical
         TabIndex        =   41
         Top             =   6860
         Width           =   1300
      End
      Begin VB.CommandButton cmdOpcion 
         Caption         =   "Ir al Punto de Venta"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   630
         Index           =   2
         Left            =   5295
         Style           =   1  'Graphical
         TabIndex        =   40
         Top             =   7520
         Width           =   1300
      End
      Begin VB.CommandButton cmdDetalle 
         Caption         =   "Mozos"
         Height          =   630
         Index           =   9
         Left            =   5295
         Style           =   1  'Graphical
         TabIndex        =   39
         Top             =   6200
         Width           =   1300
      End
      Begin VB.CommandButton cmdOpcion 
         Caption         =   "Enviar Pedido"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   630
         Index           =   8
         Left            =   2640
         TabIndex        =   38
         Top             =   6860
         Width           =   1230
      End
      Begin VB.CommandButton cmdOpcion 
         Caption         =   "Combos"
         Height          =   630
         Index           =   10
         Left            =   3915
         TabIndex        =   37
         Top             =   6860
         Width           =   1230
      End
      Begin VB.CommandButton cmdDetalle 
         Caption         =   "Visualizar Pedido"
         CausesValidation=   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   630
         Index           =   7
         Left            =   90
         Style           =   1  'Graphical
         TabIndex        =   36
         Top             =   6860
         Width           =   1230
      End
      Begin VB.CommandButton cmdTipoDocumento 
         Height          =   630
         Index           =   4
         Left            =   3915
         TabIndex        =   35
         Top             =   8175
         Width           =   1230
      End
      Begin VB.CommandButton cmdTipoDocumento 
         Height          =   630
         Index           =   1
         Left            =   90
         TabIndex        =   34
         Top             =   8160
         Width           =   1230
      End
      Begin VB.CommandButton cmdTipoDocumento 
         Height          =   630
         Index           =   2
         Left            =   1365
         TabIndex        =   33
         Top             =   8175
         Width           =   1230
      End
      Begin VB.CommandButton cmdTipoDocumento 
         Height          =   630
         Index           =   3
         Left            =   2640
         TabIndex        =   32
         Top             =   8175
         Width           =   1230
      End
      Begin VB.CommandButton cmdOpcion 
         Caption         =   "Anulación Documento"
         Height          =   630
         Index           =   0
         Left            =   90
         TabIndex        =   31
         Top             =   7520
         Width           =   1230
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
         Height          =   630
         Index           =   5
         Left            =   5295
         Picture         =   "frmCajaRapida.frx":88F4
         Style           =   1  'Graphical
         TabIndex        =   2
         Top             =   8175
         Width           =   1300
      End
      Begin VB.TextBox txtBarra 
         Height          =   435
         Left            =   180
         TabIndex        =   56
         Top             =   4320
         Width           =   1875
      End
      Begin VB.Frame fraDetalle 
         Caption         =   " Detalle del Plato "
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800080&
         Height          =   7650
         Left            =   6720
         TabIndex        =   162
         Top             =   1080
         Width           =   5220
         Begin VB.Frame fraDescuento 
            Caption         =   " Descuentos / Recargos "
            ForeColor       =   &H00800080&
            Height          =   1605
            Left            =   1665
            TabIndex        =   200
            Top             =   180
            Width           =   3090
            Begin VB.CommandButton cmdDescuento 
               Caption         =   "Monto del Descuento"
               Height          =   555
               Index           =   0
               Left            =   195
               TabIndex        =   204
               Top             =   270
               Width           =   1245
            End
            Begin VB.CommandButton cmdDescuento 
               Caption         =   "Monto del Recargo"
               Height          =   555
               Index           =   2
               Left            =   195
               TabIndex        =   203
               Top             =   915
               Width           =   1245
            End
            Begin VB.CommandButton cmdDescuento 
               Caption         =   "( % ) del Descuento"
               Height          =   555
               Index           =   1
               Left            =   1650
               TabIndex        =   202
               Top             =   270
               Width           =   1245
            End
            Begin VB.CommandButton cmdDescuento 
               Caption         =   "( % ) del Recargo"
               Height          =   555
               Index           =   3
               Left            =   1650
               TabIndex        =   201
               Top             =   915
               Width           =   1245
            End
         End
         Begin VB.Frame fraPrecio 
            Caption         =   " Precio de Venta "
            ForeColor       =   &H00800080&
            Height          =   3840
            Left            =   120
            TabIndex        =   171
            Top             =   2880
            Width           =   4890
            Begin VB.Label txtObserva 
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
               Height          =   510
               Left            =   1065
               TabIndex        =   199
               Top             =   3165
               Width           =   2895
            End
            Begin VB.Label txtCortesia 
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
               Height          =   540
               Left            =   2475
               TabIndex        =   198
               Top             =   2550
               Width           =   1485
            End
            Begin VB.Label txtVenta 
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
               ForeColor       =   &H00000000&
               Height          =   240
               Left            =   1065
               TabIndex        =   197
               Top             =   2850
               Width           =   1365
            End
            Begin VB.Label txtOficial 
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
               ForeColor       =   &H00000000&
               Height          =   240
               Left            =   1065
               TabIndex        =   196
               Top             =   240
               Width           =   1365
            End
            Begin VB.Label txtPVenta 
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
               ForeColor       =   &H00000000&
               Height          =   240
               Left            =   1065
               TabIndex        =   195
               Top             =   2310
               Width           =   1365
            End
            Begin VB.Label txtNeto 
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
               ForeColor       =   &H00808080&
               Height          =   240
               Left            =   1065
               TabIndex        =   194
               Top             =   1140
               Width           =   1365
            End
            Begin VB.Label txtDImporte 
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
               ForeColor       =   &H000000FF&
               Height          =   240
               Left            =   1785
               TabIndex        =   193
               Top             =   540
               Width           =   645
            End
            Begin VB.Label txtRImporte 
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
               Left            =   1785
               TabIndex        =   192
               Top             =   810
               Width           =   645
            End
            Begin VB.Label txtRPorcentaje 
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
               Left            =   1065
               TabIndex        =   191
               Top             =   810
               Width           =   510
            End
            Begin VB.Label txtDPorcentaje 
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
               ForeColor       =   &H000000FF&
               Height          =   240
               Left            =   1065
               TabIndex        =   190
               Top             =   540
               Width           =   510
            End
            Begin VB.Label txtImpuesto3 
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
               ForeColor       =   &H00404040&
               Height          =   240
               Left            =   1065
               TabIndex        =   189
               Top             =   1965
               Width           =   1365
            End
            Begin VB.Label txtImpuesto2 
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
               ForeColor       =   &H00404040&
               Height          =   240
               Left            =   1065
               TabIndex        =   188
               Top             =   1695
               Width           =   1365
            End
            Begin VB.Label txtImpuesto1 
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
               ForeColor       =   &H00404040&
               Height          =   240
               Left            =   1065
               TabIndex        =   187
               Top             =   1410
               Width           =   1365
            End
            Begin VB.Label Label1 
               Alignment       =   1  'Right Justify
               AutoSize        =   -1  'True
               Caption         =   "Cortesía"
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
               Index           =   21
               Left            =   2490
               TabIndex        =   186
               Top             =   2325
               Width           =   555
            End
            Begin VB.Label Label1 
               Alignment       =   1  'Right Justify
               AutoSize        =   -1  'True
               Caption         =   "%"
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
               Index           =   20
               Left            =   1590
               TabIndex        =   185
               Top             =   840
               Width           =   150
            End
            Begin VB.Label Label1 
               Alignment       =   1  'Right Justify
               AutoSize        =   -1  'True
               Caption         =   "%"
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
               Index           =   19
               Left            =   1590
               TabIndex        =   184
               Top             =   570
               Width           =   150
            End
            Begin VB.Label Label1 
               Alignment       =   1  'Right Justify
               AutoSize        =   -1  'True
               Caption         =   "Precio Neto :"
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
               Index           =   9
               Left            =   165
               TabIndex        =   183
               Top             =   1185
               Width           =   825
            End
            Begin VB.Label Label1 
               Alignment       =   1  'Right Justify
               AutoSize        =   -1  'True
               Caption         =   "Observación :"
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
               Index           =   22
               Left            =   180
               TabIndex        =   182
               Top             =   3165
               Width           =   855
            End
            Begin VB.Label Label1 
               Alignment       =   1  'Right Justify
               AutoSize        =   -1  'True
               Caption         =   "Total :"
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
               Index           =   18
               Left            =   600
               TabIndex        =   181
               Top             =   2925
               Width           =   390
            End
            Begin VB.Label Label1 
               Alignment       =   1  'Right Justify
               AutoSize        =   -1  'True
               Caption         =   "Impuesto 3 :"
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
               Index           =   12
               Left            =   240
               TabIndex        =   180
               Top             =   2010
               Width           =   750
            End
            Begin VB.Label Label1 
               Alignment       =   1  'Right Justify
               AutoSize        =   -1  'True
               Caption         =   "Impuesto 1 :"
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
               Index           =   10
               Left            =   240
               TabIndex        =   179
               Top             =   1455
               Width           =   750
            End
            Begin VB.Label Label1 
               Alignment       =   1  'Right Justify
               AutoSize        =   -1  'True
               Caption         =   "Impuesto 2 :"
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
               Index           =   11
               Left            =   240
               TabIndex        =   178
               Top             =   1725
               Width           =   750
            End
            Begin VB.Label Label1 
               Alignment       =   1  'Right Justify
               AutoSize        =   -1  'True
               Caption         =   "Precio Oficial :"
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
               Index           =   15
               Left            =   90
               TabIndex        =   177
               Top             =   300
               Width           =   900
            End
            Begin VB.Label Label1 
               Alignment       =   1  'Right Justify
               AutoSize        =   -1  'True
               Caption         =   "Descuento :"
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
               Index           =   13
               Left            =   240
               TabIndex        =   176
               Top             =   585
               Width           =   750
            End
            Begin VB.Label Label1 
               Alignment       =   1  'Right Justify
               AutoSize        =   -1  'True
               Caption         =   "Precio Venta :"
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
               Index           =   16
               Left            =   150
               TabIndex        =   175
               Top             =   2340
               Width           =   870
            End
            Begin VB.Label Label1 
               Alignment       =   1  'Right Justify
               AutoSize        =   -1  'True
               Caption         =   "Recargo :"
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
               Index           =   14
               Left            =   390
               TabIndex        =   174
               Top             =   855
               Width           =   600
            End
            Begin VB.Label Label1 
               Alignment       =   1  'Right Justify
               AutoSize        =   -1  'True
               Caption         =   "Cantidad :"
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
               Index           =   17
               Left            =   375
               TabIndex        =   173
               Top             =   2625
               Width           =   615
            End
            Begin VB.Label txtCantidad 
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
               ForeColor       =   &H00000000&
               Height          =   240
               Left            =   1065
               TabIndex        =   172
               Top             =   2580
               Width           =   1365
            End
         End
         Begin VB.Frame fraImpuesto 
            Caption         =   " Impuestos "
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00800080&
            Height          =   1065
            Left            =   120
            TabIndex        =   167
            Top             =   1785
            Width           =   4905
            Begin VB.CommandButton cmdImpuesto 
               Height          =   630
               Index           =   0
               Left            =   150
               TabIndex        =   170
               Top             =   270
               Width           =   1245
            End
            Begin VB.CommandButton cmdImpuesto 
               Height          =   630
               Index           =   1
               Left            =   1725
               TabIndex        =   169
               Top             =   270
               Width           =   1245
            End
            Begin VB.CommandButton cmdImpuesto 
               Height          =   630
               Index           =   2
               Left            =   3315
               TabIndex        =   168
               Top             =   270
               Width           =   1245
            End
         End
         Begin VB.CommandButton cmdPrecio 
            Caption         =   "Precio"
            Height          =   585
            Left            =   285
            TabIndex        =   166
            Top             =   300
            Width           =   1200
         End
         Begin VB.CommandButton cmdCortesia 
            Caption         =   "Cortesía"
            Height          =   585
            Left            =   285
            TabIndex        =   165
            Top             =   990
            Width           =   1200
         End
         Begin VB.CommandButton cmdOpcion 
            Caption         =   "Aceptar"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   630
            Index           =   3
            Left            =   3585
            Picture         =   "frmCajaRapida.frx":89E6
            Style           =   1  'Graphical
            TabIndex        =   164
            Top             =   6870
            Width           =   1410
         End
         Begin VB.CommandButton cmdOpcion 
            Caption         =   "Cancelar"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   630
            Index           =   4
            Left            =   1995
            Picture         =   "frmCajaRapida.frx":8AE8
            Style           =   1  'Graphical
            TabIndex        =   163
            Top             =   6870
            Width           =   1410
         End
      End
      Begin VB.Frame fraMozo 
         Caption         =   " Mozo "
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
         Height          =   4335
         Left            =   6750
         TabIndex        =   60
         Top             =   1080
         Width           =   5235
         Begin VB.Frame fraOrigenVentas 
            Caption         =   "Origen de Ventas"
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
            Height          =   3015
            Left            =   0
            TabIndex        =   345
            Top             =   0
            Width           =   5235
            Begin VB.Frame fraMorotizado 
               Caption         =   "Motorizado"
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
               Height          =   3015
               Left            =   0
               TabIndex        =   366
               Top             =   0
               Width           =   5235
               Begin VB.CommandButton cmdMotorizado 
                  BeginProperty Font 
                     Name            =   "Small Fonts"
                     Size            =   6.75
                     Charset         =   0
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   450
                  Index           =   17
                  Left            =   120
                  TabIndex        =   386
                  Top             =   2520
                  Width           =   1125
               End
               Begin VB.CommandButton cmdMotorizado 
                  BeginProperty Font 
                     Name            =   "Small Fonts"
                     Size            =   6.75
                     Charset         =   0
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   450
                  Index           =   19
                  Left            =   2640
                  TabIndex        =   385
                  Top             =   2520
                  Width           =   1125
               End
               Begin VB.CommandButton cmdMotorizado 
                  BeginProperty Font 
                     Name            =   "Small Fonts"
                     Size            =   6.75
                     Charset         =   0
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   450
                  Index           =   10
                  Left            =   1320
                  TabIndex        =   384
                  Top             =   1400
                  Width           =   1125
               End
               Begin VB.CommandButton cmdMotorizado 
                  BeginProperty Font 
                     Name            =   "Small Fonts"
                     Size            =   6.75
                     Charset         =   0
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   450
                  Index           =   18
                  Left            =   1320
                  TabIndex        =   383
                  Top             =   2520
                  Width           =   1125
               End
               Begin VB.CommandButton cmdMotorizado 
                  BeginProperty Font 
                     Name            =   "Small Fonts"
                     Size            =   6.75
                     Charset         =   0
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   450
                  Index           =   16
                  Left            =   3960
                  TabIndex        =   382
                  Top             =   1980
                  Width           =   1125
               End
               Begin VB.CommandButton cmdMotorizado 
                  BeginProperty Font 
                     Name            =   "Small Fonts"
                     Size            =   6.75
                     Charset         =   0
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   450
                  Index           =   15
                  Left            =   2640
                  TabIndex        =   381
                  Top             =   1980
                  Width           =   1125
               End
               Begin VB.CommandButton cmdMotorizado 
                  BeginProperty Font 
                     Name            =   "Small Fonts"
                     Size            =   6.75
                     Charset         =   0
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   450
                  Index           =   14
                  Left            =   1320
                  TabIndex        =   380
                  Top             =   1980
                  Width           =   1125
               End
               Begin VB.CommandButton cmdMotorizado 
                  BeginProperty Font 
                     Name            =   "Small Fonts"
                     Size            =   6.75
                     Charset         =   0
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   450
                  Index           =   13
                  Left            =   120
                  TabIndex        =   379
                  Top             =   1980
                  Width           =   1125
               End
               Begin VB.CommandButton cmdMotorizado 
                  BeginProperty Font 
                     Name            =   "Small Fonts"
                     Size            =   6.75
                     Charset         =   0
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   450
                  Index           =   12
                  Left            =   3960
                  TabIndex        =   378
                  Top             =   1400
                  Width           =   1125
               End
               Begin VB.CommandButton cmdMotorizado 
                  BeginProperty Font 
                     Name            =   "Small Fonts"
                     Size            =   6.75
                     Charset         =   0
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   450
                  Index           =   11
                  Left            =   2640
                  TabIndex        =   377
                  Top             =   1400
                  Width           =   1125
               End
               Begin VB.CommandButton cmdBuscar 
                  BeginProperty Font 
                     Name            =   "Small Fonts"
                     Size            =   6.75
                     Charset         =   0
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   450
                  Index           =   4
                  Left            =   3960
                  Style           =   1  'Graphical
                  TabIndex        =   376
                  Top             =   2520
                  Width           =   1125
               End
               Begin VB.CommandButton cmdMotorizado 
                  BeginProperty Font 
                     Name            =   "Small Fonts"
                     Size            =   6.75
                     Charset         =   0
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   450
                  Index           =   8
                  Left            =   3960
                  TabIndex        =   375
                  Top             =   840
                  Width           =   1125
               End
               Begin VB.CommandButton cmdMotorizado 
                  BeginProperty Font 
                     Name            =   "Small Fonts"
                     Size            =   6.75
                     Charset         =   0
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   450
                  Index           =   7
                  Left            =   2640
                  TabIndex        =   374
                  Top             =   840
                  Width           =   1125
               End
               Begin VB.CommandButton cmdMotorizado 
                  BeginProperty Font 
                     Name            =   "Small Fonts"
                     Size            =   6.75
                     Charset         =   0
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   450
                  Index           =   6
                  Left            =   1320
                  TabIndex        =   373
                  Top             =   840
                  Width           =   1125
               End
               Begin VB.CommandButton cmdMotorizado 
                  BeginProperty Font 
                     Name            =   "Small Fonts"
                     Size            =   6.75
                     Charset         =   0
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   450
                  Index           =   9
                  Left            =   120
                  TabIndex        =   372
                  Top             =   1400
                  Width           =   1125
               End
               Begin VB.CommandButton cmdMotorizado 
                  BeginProperty Font 
                     Name            =   "Small Fonts"
                     Size            =   6.75
                     Charset         =   0
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   450
                  Index           =   5
                  Left            =   120
                  TabIndex        =   371
                  Top             =   840
                  Width           =   1125
               End
               Begin VB.CommandButton cmdMotorizado 
                  BeginProperty Font 
                     Name            =   "Small Fonts"
                     Size            =   6.75
                     Charset         =   0
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   435
                  Index           =   4
                  Left            =   3960
                  TabIndex        =   370
                  Top             =   300
                  Width           =   1125
               End
               Begin VB.CommandButton cmdMotorizado 
                  BeginProperty Font 
                     Name            =   "Small Fonts"
                     Size            =   6.75
                     Charset         =   0
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   435
                  Index           =   3
                  Left            =   2640
                  TabIndex        =   369
                  Top             =   300
                  Width           =   1125
               End
               Begin VB.CommandButton cmdMotorizado 
                  BeginProperty Font 
                     Name            =   "Small Fonts"
                     Size            =   6.75
                     Charset         =   0
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   435
                  Index           =   2
                  Left            =   1320
                  TabIndex        =   368
                  Top             =   300
                  Width           =   1125
               End
               Begin VB.CommandButton cmdMotorizado 
                  BeginProperty Font 
                     Name            =   "Small Fonts"
                     Size            =   6.75
                     Charset         =   0
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   435
                  Index           =   1
                  Left            =   120
                  TabIndex        =   367
                  Top             =   300
                  Width           =   1125
               End
            End
            Begin VB.CommandButton cmdOrigen 
               BeginProperty Font 
                  Name            =   "Small Fonts"
                  Size            =   6.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   435
               Index           =   1
               Left            =   120
               TabIndex        =   365
               Top             =   300
               Width           =   1125
            End
            Begin VB.CommandButton cmdOrigen 
               BeginProperty Font 
                  Name            =   "Small Fonts"
                  Size            =   6.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   435
               Index           =   2
               Left            =   1320
               TabIndex        =   364
               Top             =   300
               Width           =   1125
            End
            Begin VB.CommandButton cmdOrigen 
               BeginProperty Font 
                  Name            =   "Small Fonts"
                  Size            =   6.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   435
               Index           =   3
               Left            =   2640
               TabIndex        =   363
               Top             =   300
               Width           =   1125
            End
            Begin VB.CommandButton cmdOrigen 
               BeginProperty Font 
                  Name            =   "Small Fonts"
                  Size            =   6.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   435
               Index           =   4
               Left            =   3960
               TabIndex        =   362
               Top             =   300
               Width           =   1125
            End
            Begin VB.CommandButton cmdOrigen 
               BeginProperty Font 
                  Name            =   "Small Fonts"
                  Size            =   6.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   450
               Index           =   5
               Left            =   120
               TabIndex        =   361
               Top             =   840
               Width           =   1125
            End
            Begin VB.CommandButton cmdOrigen 
               BeginProperty Font 
                  Name            =   "Small Fonts"
                  Size            =   6.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   450
               Index           =   9
               Left            =   120
               TabIndex        =   360
               Top             =   1400
               Width           =   1125
            End
            Begin VB.CommandButton cmdOrigen 
               BeginProperty Font 
                  Name            =   "Small Fonts"
                  Size            =   6.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   450
               Index           =   6
               Left            =   1320
               TabIndex        =   359
               Top             =   840
               Width           =   1125
            End
            Begin VB.CommandButton cmdOrigen 
               BeginProperty Font 
                  Name            =   "Small Fonts"
                  Size            =   6.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   450
               Index           =   7
               Left            =   2640
               TabIndex        =   358
               Top             =   840
               Width           =   1125
            End
            Begin VB.CommandButton cmdOrigen 
               BeginProperty Font 
                  Name            =   "Small Fonts"
                  Size            =   6.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   450
               Index           =   8
               Left            =   3960
               TabIndex        =   357
               Top             =   840
               Width           =   1125
            End
            Begin VB.CommandButton cmdBuscar 
               BeginProperty Font 
                  Name            =   "Small Fonts"
                  Size            =   6.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   450
               Index           =   3
               Left            =   3960
               Style           =   1  'Graphical
               TabIndex        =   356
               Top             =   2520
               Width           =   1125
            End
            Begin VB.CommandButton cmdOrigen 
               BeginProperty Font 
                  Name            =   "Small Fonts"
                  Size            =   6.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   450
               Index           =   11
               Left            =   2640
               TabIndex        =   355
               Top             =   1400
               Width           =   1125
            End
            Begin VB.CommandButton cmdOrigen 
               BeginProperty Font 
                  Name            =   "Small Fonts"
                  Size            =   6.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   450
               Index           =   12
               Left            =   3960
               TabIndex        =   354
               Top             =   1400
               Width           =   1125
            End
            Begin VB.CommandButton cmdOrigen 
               BeginProperty Font 
                  Name            =   "Small Fonts"
                  Size            =   6.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   450
               Index           =   13
               Left            =   120
               TabIndex        =   353
               Top             =   1980
               Width           =   1125
            End
            Begin VB.CommandButton cmdOrigen 
               BeginProperty Font 
                  Name            =   "Small Fonts"
                  Size            =   6.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   450
               Index           =   14
               Left            =   1320
               TabIndex        =   352
               Top             =   1980
               Width           =   1125
            End
            Begin VB.CommandButton cmdOrigen 
               BeginProperty Font 
                  Name            =   "Small Fonts"
                  Size            =   6.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   450
               Index           =   15
               Left            =   2640
               TabIndex        =   351
               Top             =   1980
               Width           =   1125
            End
            Begin VB.CommandButton cmdOrigen 
               BeginProperty Font 
                  Name            =   "Small Fonts"
                  Size            =   6.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   450
               Index           =   16
               Left            =   3960
               TabIndex        =   350
               Top             =   1980
               Width           =   1125
            End
            Begin VB.CommandButton cmdOrigen 
               BeginProperty Font 
                  Name            =   "Small Fonts"
                  Size            =   6.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   450
               Index           =   18
               Left            =   1320
               TabIndex        =   349
               Top             =   2520
               Width           =   1125
            End
            Begin VB.CommandButton cmdOrigen 
               BeginProperty Font 
                  Name            =   "Small Fonts"
                  Size            =   6.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   450
               Index           =   10
               Left            =   1320
               TabIndex        =   348
               Top             =   1400
               Width           =   1125
            End
            Begin VB.CommandButton cmdOrigen 
               BeginProperty Font 
                  Name            =   "Small Fonts"
                  Size            =   6.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   450
               Index           =   19
               Left            =   2640
               TabIndex        =   347
               Top             =   2520
               Width           =   1125
            End
            Begin VB.CommandButton cmdOrigen 
               BeginProperty Font 
                  Name            =   "Small Fonts"
                  Size            =   6.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   450
               Index           =   17
               Left            =   120
               TabIndex        =   346
               Top             =   2520
               Width           =   1125
            End
         End
         Begin VB.CommandButton cmdBuscar 
            BeginProperty Font 
               Name            =   "Small Fonts"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   810
            Index           =   0
            Left            =   4080
            Picture         =   "frmCajaRapida.frx":8BEA
            Style           =   1  'Graphical
            TabIndex        =   80
            Top             =   3270
            Width           =   810
         End
         Begin VB.CommandButton cmdMozo 
            BeginProperty Font 
               Name            =   "Small Fonts"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   810
            Index           =   19
            Left            =   3090
            TabIndex        =   79
            Top             =   3270
            Width           =   810
         End
         Begin VB.CommandButton cmdMozo 
            BeginProperty Font 
               Name            =   "Small Fonts"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   810
            Index           =   18
            Left            =   2100
            TabIndex        =   78
            Top             =   3270
            Width           =   810
         End
         Begin VB.CommandButton cmdMozo 
            BeginProperty Font 
               Name            =   "Small Fonts"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   810
            Index           =   17
            Left            =   1110
            TabIndex        =   77
            Top             =   3270
            Width           =   810
         End
         Begin VB.CommandButton cmdMozo 
            BeginProperty Font 
               Name            =   "Small Fonts"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   810
            Index           =   16
            Left            =   120
            TabIndex        =   76
            Top             =   3270
            Width           =   810
         End
         Begin VB.CommandButton cmdMozo 
            BeginProperty Font 
               Name            =   "Small Fonts"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   810
            Index           =   15
            Left            =   4080
            TabIndex        =   75
            Top             =   2280
            Width           =   810
         End
         Begin VB.CommandButton cmdMozo 
            BeginProperty Font 
               Name            =   "Small Fonts"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   810
            Index           =   14
            Left            =   3090
            TabIndex        =   74
            Top             =   2280
            Width           =   810
         End
         Begin VB.CommandButton cmdMozo 
            BeginProperty Font 
               Name            =   "Small Fonts"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   810
            Index           =   13
            Left            =   2100
            TabIndex        =   73
            Top             =   2280
            Width           =   810
         End
         Begin VB.CommandButton cmdMozo 
            BeginProperty Font 
               Name            =   "Small Fonts"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   810
            Index           =   12
            Left            =   1110
            TabIndex        =   72
            Top             =   2280
            Width           =   810
         End
         Begin VB.CommandButton cmdMozo 
            BeginProperty Font 
               Name            =   "Small Fonts"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   810
            Index           =   11
            Left            =   120
            TabIndex        =   71
            Top             =   2280
            Width           =   810
         End
         Begin VB.CommandButton cmdMozo 
            BeginProperty Font 
               Name            =   "Small Fonts"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   810
            Index           =   10
            Left            =   4080
            TabIndex        =   70
            Top             =   1290
            Width           =   810
         End
         Begin VB.CommandButton cmdMozo 
            BeginProperty Font 
               Name            =   "Small Fonts"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   810
            Index           =   9
            Left            =   3090
            TabIndex        =   69
            Top             =   1290
            Width           =   810
         End
         Begin VB.CommandButton cmdMozo 
            BeginProperty Font 
               Name            =   "Small Fonts"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   810
            Index           =   8
            Left            =   2100
            TabIndex        =   68
            Top             =   1290
            Width           =   810
         End
         Begin VB.CommandButton cmdMozo 
            BeginProperty Font 
               Name            =   "Small Fonts"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   810
            Index           =   7
            Left            =   1110
            TabIndex        =   67
            Top             =   1290
            Width           =   810
         End
         Begin VB.CommandButton cmdMozo 
            BeginProperty Font 
               Name            =   "Small Fonts"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   810
            Index           =   6
            Left            =   120
            TabIndex        =   66
            Top             =   1290
            Width           =   810
         End
         Begin VB.CommandButton cmdMozo 
            BeginProperty Font 
               Name            =   "Small Fonts"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   810
            Index           =   5
            Left            =   4080
            TabIndex        =   65
            Top             =   300
            Width           =   810
         End
         Begin VB.CommandButton cmdMozo 
            BeginProperty Font 
               Name            =   "Small Fonts"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   810
            Index           =   4
            Left            =   3090
            TabIndex        =   64
            Top             =   300
            Width           =   810
         End
         Begin VB.CommandButton cmdMozo 
            BeginProperty Font 
               Name            =   "Small Fonts"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   810
            Index           =   3
            Left            =   2100
            TabIndex        =   63
            Top             =   300
            Width           =   810
         End
         Begin VB.CommandButton cmdMozo 
            BeginProperty Font 
               Name            =   "Small Fonts"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   810
            Index           =   2
            Left            =   1110
            TabIndex        =   62
            Top             =   300
            Width           =   810
         End
         Begin VB.CommandButton cmdMozo 
            BeginProperty Font 
               Name            =   "Small Fonts"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   810
            Index           =   1
            Left            =   120
            TabIndex        =   61
            Top             =   300
            Width           =   810
         End
      End
      Begin VB.Frame fraPropiedad 
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
         Height          =   7665
         Left            =   6750
         TabIndex        =   81
         Top             =   1080
         Width           =   5235
         Begin VB.CommandButton cmdBuscar 
            BeginProperty Font 
               Name            =   "Small Fonts"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   720
            Index           =   8
            Left            =   4320
            Picture         =   "frmCajaRapida.frx":902C
            Style           =   1  'Graphical
            TabIndex        =   329
            Top             =   3498
            Width           =   720
         End
         Begin VB.TextBox lblObservacion 
            Height          =   915
            Left            =   135
            Locked          =   -1  'True
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   119
            Top             =   5895
            Width           =   4965
         End
         Begin VB.TextBox lblResumen 
            Height          =   1500
            Left            =   1980
            Locked          =   -1  'True
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   118
            Top             =   4290
            Width           =   3075
         End
         Begin VB.Frame Frame4 
            Caption         =   "Frame4"
            Height          =   5460
            Left            =   1755
            TabIndex        =   117
            Top             =   225
            Width           =   60
         End
         Begin VB.CommandButton cmdBusca 
            BeginProperty Font 
               Name            =   "Small Fonts"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   720
            Left            =   915
            Picture         =   "frmCajaRapida.frx":946E
            Style           =   1  'Graphical
            TabIndex        =   116
            Top             =   5085
            Width           =   720
         End
         Begin VB.CommandButton cmdOpcion 
            Caption         =   "Observación"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   720
            Index           =   7
            Left            =   2475
            Picture         =   "frmCajaRapida.frx":98B0
            Style           =   1  'Graphical
            TabIndex        =   115
            Top             =   6885
            Width           =   1470
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
            Height          =   720
            Index           =   6
            Left            =   4005
            Picture         =   "frmCajaRapida.frx":99F2
            Style           =   1  'Graphical
            TabIndex        =   114
            Top             =   6885
            Width           =   1110
         End
         Begin VB.CommandButton cmdPropiedad 
            BeginProperty Font 
               Name            =   "Small Fonts"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   720
            Index           =   5
            Left            =   1995
            TabIndex        =   113
            Top             =   1122
            Width           =   720
         End
         Begin VB.CommandButton cmdPropiedad 
            BeginProperty Font 
               Name            =   "Small Fonts"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   720
            Index           =   4
            Left            =   4320
            TabIndex        =   112
            Top             =   330
            Width           =   720
         End
         Begin VB.CommandButton cmdOperador 
            BeginProperty Font 
               Name            =   "Small Fonts"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   720
            Index           =   9
            Left            =   915
            Style           =   1  'Graphical
            TabIndex        =   111
            Top             =   1122
            Width           =   720
         End
         Begin VB.CommandButton cmdOperador 
            BeginProperty Font 
               Name            =   "Small Fonts"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   720
            Index           =   2
            Left            =   135
            Style           =   1  'Graphical
            TabIndex        =   110
            Top             =   1122
            Width           =   720
         End
         Begin VB.CommandButton cmdPropiedad 
            BeginProperty Font 
               Name            =   "Small Fonts"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   720
            Index           =   3
            Left            =   3540
            TabIndex        =   109
            Top             =   330
            Width           =   720
         End
         Begin VB.CommandButton cmdPropiedad 
            BeginProperty Font 
               Name            =   "Small Fonts"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   720
            Index           =   2
            Left            =   2775
            TabIndex        =   108
            Top             =   330
            Width           =   720
         End
         Begin VB.CommandButton cmdPropiedad 
            BeginProperty Font 
               Name            =   "Small Fonts"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   720
            Index           =   1
            Left            =   1995
            TabIndex        =   107
            Top             =   330
            Width           =   720
         End
         Begin VB.CommandButton cmdOperador 
            BeginProperty Font 
               Name            =   "Small Fonts"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   720
            Index           =   8
            Left            =   915
            Style           =   1  'Graphical
            TabIndex        =   106
            Top             =   330
            Width           =   720
         End
         Begin VB.CommandButton cmdOperador 
            BeginProperty Font 
               Name            =   "Small Fonts"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   720
            Index           =   1
            Left            =   135
            Style           =   1  'Graphical
            TabIndex        =   105
            Top             =   330
            Width           =   720
         End
         Begin VB.CommandButton cmdPropiedad 
            BeginProperty Font 
               Name            =   "Small Fonts"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   720
            Index           =   6
            Left            =   2775
            TabIndex        =   104
            Top             =   1122
            Width           =   720
         End
         Begin VB.CommandButton cmdPropiedad 
            BeginProperty Font 
               Name            =   "Small Fonts"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   720
            Index           =   8
            Left            =   4320
            TabIndex        =   103
            Top             =   1122
            Width           =   720
         End
         Begin VB.CommandButton cmdPropiedad 
            BeginProperty Font 
               Name            =   "Small Fonts"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   720
            Index           =   7
            Left            =   3540
            TabIndex        =   102
            Top             =   1122
            Width           =   720
         End
         Begin VB.CommandButton cmdOperador 
            BeginProperty Font 
               Name            =   "Small Fonts"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   720
            Index           =   10
            Left            =   915
            Style           =   1  'Graphical
            TabIndex        =   101
            Top             =   1914
            Width           =   720
         End
         Begin VB.CommandButton cmdOperador 
            BeginProperty Font 
               Name            =   "Small Fonts"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   720
            Index           =   3
            Left            =   135
            Style           =   1  'Graphical
            TabIndex        =   100
            Top             =   1914
            Width           =   720
         End
         Begin VB.CommandButton cmdPropiedad 
            BeginProperty Font 
               Name            =   "Small Fonts"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   720
            Index           =   9
            Left            =   1995
            TabIndex        =   99
            Top             =   1914
            Width           =   720
         End
         Begin VB.CommandButton cmdPropiedad 
            BeginProperty Font 
               Name            =   "Small Fonts"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   720
            Index           =   11
            Left            =   3540
            TabIndex        =   98
            Top             =   1914
            Width           =   720
         End
         Begin VB.CommandButton cmdPropiedad 
            BeginProperty Font 
               Name            =   "Small Fonts"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   720
            Index           =   10
            Left            =   2775
            TabIndex        =   97
            Top             =   1914
            Width           =   720
         End
         Begin VB.CommandButton cmdOperador 
            BeginProperty Font 
               Name            =   "Small Fonts"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   720
            Index           =   11
            Left            =   915
            Style           =   1  'Graphical
            TabIndex        =   96
            Top             =   2706
            Width           =   720
         End
         Begin VB.CommandButton cmdOperador 
            BeginProperty Font 
               Name            =   "Small Fonts"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   720
            Index           =   4
            Left            =   135
            Style           =   1  'Graphical
            TabIndex        =   95
            Top             =   2706
            Width           =   720
         End
         Begin VB.CommandButton cmdPropiedad 
            BeginProperty Font 
               Name            =   "Small Fonts"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   720
            Index           =   12
            Left            =   4320
            TabIndex        =   94
            Top             =   1914
            Width           =   720
         End
         Begin VB.CommandButton cmdPropiedad 
            BeginProperty Font 
               Name            =   "Small Fonts"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   720
            Index           =   14
            Left            =   2775
            TabIndex        =   93
            Top             =   2706
            Width           =   720
         End
         Begin VB.CommandButton cmdPropiedad 
            BeginProperty Font 
               Name            =   "Small Fonts"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   720
            Index           =   13
            Left            =   1995
            TabIndex        =   92
            Top             =   2706
            Width           =   720
         End
         Begin VB.CommandButton cmdOperador 
            BeginProperty Font 
               Name            =   "Small Fonts"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   720
            Index           =   12
            Left            =   900
            Style           =   1  'Graphical
            TabIndex        =   91
            Top             =   3498
            Width           =   720
         End
         Begin VB.CommandButton cmdOperador 
            BeginProperty Font 
               Name            =   "Small Fonts"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   720
            Index           =   5
            Left            =   135
            Style           =   1  'Graphical
            TabIndex        =   90
            Top             =   3498
            Width           =   720
         End
         Begin VB.CommandButton cmdPropiedad 
            BeginProperty Font 
               Name            =   "Small Fonts"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   720
            Index           =   15
            Left            =   3540
            TabIndex        =   89
            Top             =   2706
            Width           =   720
         End
         Begin VB.CommandButton cmdPropiedad 
            BeginProperty Font 
               Name            =   "Small Fonts"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   720
            Index           =   17
            Left            =   1995
            TabIndex        =   88
            Top             =   3498
            Width           =   720
         End
         Begin VB.CommandButton cmdPropiedad 
            BeginProperty Font 
               Name            =   "Small Fonts"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   720
            Index           =   16
            Left            =   4320
            TabIndex        =   87
            Top             =   2706
            Width           =   720
         End
         Begin VB.CommandButton cmdOperador 
            BeginProperty Font 
               Name            =   "Small Fonts"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   720
            Index           =   13
            Left            =   900
            Style           =   1  'Graphical
            TabIndex        =   86
            Top             =   4290
            Width           =   720
         End
         Begin VB.CommandButton cmdOperador 
            BeginProperty Font 
               Name            =   "Small Fonts"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   720
            Index           =   6
            Left            =   135
            Style           =   1  'Graphical
            TabIndex        =   85
            Top             =   4290
            Width           =   720
         End
         Begin VB.CommandButton cmdPropiedad 
            BeginProperty Font 
               Name            =   "Small Fonts"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   720
            Index           =   18
            Left            =   2775
            TabIndex        =   84
            Top             =   3498
            Width           =   720
         End
         Begin VB.CommandButton cmdPropiedad 
            BeginProperty Font 
               Name            =   "Small Fonts"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   720
            Index           =   19
            Left            =   3540
            TabIndex        =   83
            Top             =   3498
            Width           =   720
         End
         Begin VB.CommandButton cmdOperador 
            BeginProperty Font 
               Name            =   "Small Fonts"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   720
            Index           =   7
            Left            =   135
            Style           =   1  'Graphical
            TabIndex        =   82
            Top             =   5085
            Width           =   720
         End
         Begin VB.Label lblPropiedad 
            AutoSize        =   -1  'True
            Caption         =   "  Propiedad "
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
            Height          =   195
            Index           =   3
            Left            =   2025
            TabIndex        =   121
            Top             =   45
            Width           =   1050
         End
         Begin VB.Label lblPropiedad 
            AutoSize        =   -1  'True
            Caption         =   "  Operador   "
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
            Height          =   195
            Index           =   2
            Left            =   315
            TabIndex        =   120
            Top             =   0
            Width           =   1095
         End
      End
      Begin VB.Frame fraPuntoVenta 
         Caption         =   " Punto de Venta "
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
         Height          =   2535
         Left            =   6750
         TabIndex        =   20
         Top             =   1080
         Width           =   5235
         Begin VB.CommandButton cmdPunto 
            BeginProperty Font 
               Name            =   "Small Fonts"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   810
            Index           =   1
            Left            =   120
            TabIndex        =   30
            Top             =   300
            Width           =   810
         End
         Begin VB.CommandButton cmdPunto 
            BeginProperty Font 
               Name            =   "Small Fonts"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   810
            Index           =   2
            Left            =   1110
            TabIndex        =   29
            Top             =   300
            Width           =   810
         End
         Begin VB.CommandButton cmdPunto 
            BeginProperty Font 
               Name            =   "Small Fonts"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   810
            Index           =   3
            Left            =   2100
            TabIndex        =   28
            Top             =   300
            Width           =   810
         End
         Begin VB.CommandButton cmdPunto 
            BeginProperty Font 
               Name            =   "Small Fonts"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   810
            Index           =   4
            Left            =   3090
            TabIndex        =   27
            Top             =   300
            Width           =   810
         End
         Begin VB.CommandButton cmdPunto 
            BeginProperty Font 
               Name            =   "Small Fonts"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   810
            Index           =   5
            Left            =   4080
            TabIndex        =   26
            Top             =   300
            Width           =   810
         End
         Begin VB.CommandButton cmdPunto 
            BeginProperty Font 
               Name            =   "Small Fonts"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   810
            Index           =   6
            Left            =   120
            TabIndex        =   25
            Top             =   1290
            Width           =   810
         End
         Begin VB.CommandButton cmdPunto 
            BeginProperty Font 
               Name            =   "Small Fonts"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   810
            Index           =   7
            Left            =   1110
            TabIndex        =   24
            Top             =   1290
            Width           =   810
         End
         Begin VB.CommandButton cmdPunto 
            BeginProperty Font 
               Name            =   "Small Fonts"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   810
            Index           =   8
            Left            =   2100
            TabIndex        =   23
            Top             =   1290
            Width           =   810
         End
         Begin VB.CommandButton cmdPunto 
            BeginProperty Font 
               Name            =   "Small Fonts"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   810
            Index           =   9
            Left            =   3090
            TabIndex        =   22
            Top             =   1290
            Width           =   810
         End
         Begin VB.CommandButton cmdBuscar 
            BeginProperty Font 
               Name            =   "Small Fonts"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   810
            Index           =   1
            Left            =   4080
            Picture         =   "frmCajaRapida.frx":9AE4
            Style           =   1  'Graphical
            TabIndex        =   21
            Top             =   1290
            Width           =   810
         End
      End
      Begin VB.CommandButton cmdDetalle 
         Caption         =   "Ofertas"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   630
         Index           =   13
         Left            =   2640
         Picture         =   "frmCajaRapida.frx":9F26
         Style           =   1  'Graphical
         TabIndex        =   321
         Top             =   7520
         Width           =   1230
      End
      Begin VB.Label Label3 
         Caption         =   "Fecha Entrega:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   5290
         TabIndex        =   339
         Top             =   5520
         Width           =   1220
      End
      Begin VB.Label txtFechaEntrega 
         BackColor       =   &H00C0FFFF&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   6.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   370
         Left            =   5290
         TabIndex        =   338
         Top             =   5770
         Width           =   1340
      End
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         BackColor       =   &H00C0C0FF&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Insumos Críticos"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   6.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   435
         Left            =   6675
         TabIndex        =   332
         Top             =   0
         Width           =   5355
      End
      Begin VB.Label txtEntregar 
         BackColor       =   &H00FFFFC0&
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
         ForeColor       =   &H00000000&
         Height          =   300
         Left            =   4680
         TabIndex        =   331
         Top             =   540
         Width           =   1950
      End
      Begin VB.Image imagepIE 
         Height          =   135
         Left            =   0
         Top             =   8280
         Visible         =   0   'False
         Width           =   135
      End
      Begin VB.Image imageCab 
         Height          =   135
         Left            =   0
         Top             =   8520
         Visible         =   0   'False
         Width           =   135
      End
      Begin VB.Label txtDescuento 
         Alignment       =   1  'Right Justify
         BorderStyle     =   1  'Fixed Single
         Caption         =   "1,500.00"
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
         Height          =   300
         Left            =   4140
         TabIndex        =   327
         Top             =   5190
         Width           =   990
      End
      Begin VB.Label txtCliente 
         BackColor       =   &H00FFC0C0&
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
         ForeColor       =   &H00000000&
         Height          =   350
         Left            =   90
         TabIndex        =   326
         Top             =   520
         Width           =   1950
      End
      Begin VB.Label txtObservacion 
         BackColor       =   &H00C0FFFF&
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
         ForeColor       =   &H00000000&
         Height          =   300
         Left            =   4680
         TabIndex        =   325
         Top             =   160
         Width           =   1950
      End
      Begin VB.Label txtTelefono 
         BackColor       =   &H00FFC0C0&
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
         ForeColor       =   &H00000000&
         Height          =   270
         Left            =   90
         TabIndex        =   208
         Top             =   180
         Width           =   1950
      End
      Begin VB.Label txtMonto 
         Alignment       =   1  'Right Justify
         BorderStyle     =   1  'Fixed Single
         Caption         =   "250,500.00"
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
         Height          =   300
         Left            =   5235
         TabIndex        =   207
         Top             =   5190
         Width           =   1230
      End
      Begin VB.Label txtMontoLetras 
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
         Height          =   285
         Left            =   90
         TabIndex        =   206
         Top             =   5190
         Width           =   3930
      End
      Begin VB.Label txtTipoDocumento 
         Caption         =   "TipoDocumento"
         Height          =   240
         Left            =   10860
         TabIndex        =   324
         Top             =   1500
         Width           =   1005
      End
   End
   Begin VB.Label txtTitulo 
      BackColor       =   &H00800000&
      Caption         =   " Caja Rápida 001"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   350
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   12045
   End
   Begin VB.Image imageHash 
      Height          =   735
      Left            =   12120
      Top             =   3720
      Visible         =   0   'False
      Width           =   135
   End
End
Attribute VB_Name = "frmCajaRapida"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)

Private Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" ( _
                    ByVal hwnd As Long, _
                    ByVal lpOperation As String, _
                    ByVal lpFile As String, _
                    ByVal lpParameters As String, _
                    ByVal lpDirectory As String, _
                    ByVal nShowCmd As Long) As Long

Private Const SW_HIDE As Long = 0
Private Const SW_SHOWNORMAL As Long = 1
Private Const SW_SHOWMAXIMIZED As Long = 3
Private Const SW_SHOWMINIMIZED As Long = 2

Option Explicit
Dim nMontoPedidoFacturar As Double
' anulacion por nota de credito
Dim RsTparametro As Recordset
'--------------------------------
Dim xOperador As String
Dim wAgregarPropiedad As Boolean
Dim numeroSerieImpresora As String 'SUNAT
Dim codigoImpresora As String 'SUNAT
Dim rstFuente As ADODB.Recordset
Dim RsDetalle As Recordset
Dim RsProducto As Recordset
Dim RsPropiedad As Recordset
Dim RsArea As Recordset
Dim RsTipoDocumento As Recordset
Dim RsImpresion As Recordset
Dim RsCombo As Recordset
Dim RsMozo As Recordset
Dim RsMotivoEliminacion As Recordset
Dim RsOferta As Recordset
Dim RsProductoPropiedad As Recordset
Dim RsComboPropiedad As Recordset
Dim RsOperador As Recordset
Dim rsPuntoVenta As Recordset
Dim RsProductoCombo As Recordset
Dim RsCajaRapida As Recordset

'Canales de Venta
Dim lActivaMozo As Boolean
Dim lActivaMotorizado As Boolean
Dim lObligaMesa As Boolean
'Origen de ventas
Dim lOrigenVentas As Boolean
'Public lObligaPax As Boolean
Dim lObligaMotorizado As Boolean
Dim lObligaMozo As Boolean
Dim lObligaFechaEntrega As Boolean
Dim lObligaClienteFrecuente As Boolean
Dim lCanalDelivery As Boolean
Dim lCanalCentralPedidos As Boolean
Dim Tienda As String
Dim RsMotorizado As Recordset


'origen de ventas
Dim RsOrigenVentas As Recordset
Dim RscanalOrigenVentas As Recordset
Dim vOrigenVentas As String
'Dim lOrigenVentas As Boolean
Dim RsCanalesVenta As Recordset
'-----------------------------
Dim sProducto As String
Dim sProductoCombo As String

Dim nMonto As Double

Public sDetalle As String
Public sProductoPropiedad As String
'validacionMontoMInimo
Dim nMontoPedidoFacturarMInimo As Double
'validacionMontoMInimo


Dim i As Integer
Dim sCortesia As String
Dim sCombo As String
Public sComboDetalle As String
Public sComboPropiedad As String
Dim sDetalleConsumo As String
Dim Index As Integer
Dim sTipoPedido As String
Dim sMonedaBase As String
Dim sPuntoVenta As String
Dim sComandaInfhotel As String

'Variables Combo
Dim wCombo As Boolean
Dim wAgregaCombo As Boolean
Dim nCombo As Integer


Dim nPVenta As Double
Dim nPBase As Double
Dim nImpuesto1 As Double
Dim nImpuesto2 As Double
Dim nImpuesto3 As Double
Dim nRecargo As Double
Dim nDescuento As Double
Dim nOficial As Double
Dim nCantidad As Double
Dim sitem As String
Dim xItem As String
Dim sMotorizado As String
Dim sObser As String
Dim sSubGrupo As String
Dim sGrupo As String
Dim nPos As Integer
Dim nCCombo As Double
Dim sUsuarioAutoriza As String
Dim xDescuento As Double
Public Pedido As String
Dim nOperadorPropiedad As Integer
Dim nOrden As Integer
Dim lPropiedad As Boolean

Dim nRet As Integer
Dim sOperacion As String
Dim sRetorno As String * 512
Dim sClave As String
Dim sMonto As String
Dim xError As String
Dim sRefer As String
Dim nCorrela As String
Dim lEmisor As Boolean
Dim lLoop As Boolean
Dim nContador As Integer
Dim sPrefijo As String

Dim sTD As String
Dim xSuma As Double
Dim sCompania As String
Dim sContacto As String
Dim UltimaComanda As String

Dim tAutorizaDescuento As String
Dim sCodigoDescuento As String
Dim sDescripcionDescuento As String
Dim sClienteFrecuente As String

Dim ltope As Boolean
Dim nTope As Double
Dim lRatio As Boolean
Dim Acumulado As Double
Dim lImprimeAlternativa As Boolean
Dim lAplicablePedido As Boolean

'============================================= extranjero bolivia
Dim tAutorizacion As String
Dim tcodigoControl As String
Dim tDosificacion As String
Dim tIdentidadNIT As String

Dim muestra As String
Dim variableEmite As Boolean
Dim nTotalDescuento As Double
Dim sXML As String

'insumocombo
Dim sInsumoCombo As String


'------VALIDA CORREO----------
Dim sTipoDocum As String
Dim lValidaEmail As Boolean
Dim sEmail As String


'FACTURACION_E_PERU
Dim RsImpDocumentoE As New Recordset
Dim RsCodigoHash As New ADODB.Recordset
Dim fDocumento As String
Dim xMontoTexto As String
Dim iImagenCab As Boolean
Dim xImpresionFE As String
Dim xImpresioDE As String
    
'
Dim tUsuActua As String
Dim TimpresionDolaresDelivery  As Boolean

Private Sub ImportarPedido(tempPedido As String)
    
         Isql = "insert into [" & sProductoPropiedad & "] " & _
               "(tItem,tCodigoPropiedad,tProducto,tEnlace,nInsumo,nGasto,nManoObra,nCantidad,nInsumoUnitario,nGastounitario,nManoObraUnitario) " & _
               "select tItem,tCodigoPropiedad,tProducto,tEnlace,nInsumo,nGasto,nManoObra,nCantidad,nInsumoUnitario,nGastounitario,nManoObraUnitario from TPRODUCTOPROPIEDAD where tcodigopedido='" & tempPedido & "' "
        
         Cn.Execute Isql
         RsProductoPropiedad.Requery
         
         Isql = "insert into [" & sDetalle & "] " & _
                "(tCodigoPedido, tTipoPedido, tItem, tCodigoProducto, tCodigoGrupo, tCodigoSubGrupo, " & _
                "nPrecioNeto, nRecargo, nDescuento, nPrecioOficial, nPrecioImpuesto1, nPrecioImpuesto2, nPrecioImpuesto3, nPrecioVenta, " & _
                "nCantidad, nVenta, nImpuesto1, nImpuesto2, nImpuesto3, " & _
                "lImprime, tArea, lImprimeArea, lCombinacion, nCombinacion, nInsumo, nGasto, nManoObra, nOrden, tEstadoItem,tsubalmacen,toferta,tCajaD,tObservacion) " & _
                "select tCodigoPedido,tTipoPedido,tItem,tCodigoProducto, tCodigoGrupo,tCodigoSubGrupo,nPrecioNeto,nRecargo,nDescuento, " & _
                "nPrecioOficial,nPrecioImpuesto1, nPrecioImpuesto2, nPrecioImpuesto3, nPrecioVenta,nCantidad, nVenta, nImpuesto1, nImpuesto2, nImpuesto3, " & _
                "lImprime, tArea, lImprimeArea, lCombinacion, nCombinacion, nInsumo, nGasto, nManoObra, nOrden, tEstadoItem,tsubalmacen,toferta,tCajaD,tObservacion from dpedido where tcodigopedido='" & tempPedido & "' "
         Cn.Execute Isql
         RsDetalle.Requery
                             
         Isql = "insert into [" & sComboPropiedad & "] " & _
               "(tItem,tItemCombo,tCodigoPropiedad,tProducto,tEnlace,nInsumo,nGasto,nManoObra,nCantidad,nInsumoUnitario,nGastoUnitario,nManoObraUnitario) " & _
               "select tItem,tItemCombo,tCodigoPropiedad,tProducto,tEnlace,nInsumo,nGasto,nManoObra,nCantidad,nInsumoUnitario,nGastoUnitario,nManoObraUnitario from TCOMBOPROPIEDAD where tcodigopedido='" & tempPedido & "' "

         Cn.Execute Isql
         RsComboPropiedad.Requery
         
         Isql = "insert into [" & sComboDetalle & "] " & _
               "(tItem, tItemCombo, tProducto, tProductoCombo, nCantidad, tCodigoGrupo, tCodigoSubGrupo, nPrecioNeto, nImpuesto1, nImpuesto2, nImpuesto3, nVenta, nInsumo, nGasto, nManoObra, lImprimeArea, lImprime, nOrden, tObservacion, lCorte) " & _
               "select tItem, tItemCombo, tProducto, tProductoCombo, nCantidad, tCodigoGrupo, tCodigoSubGrupo, nPrecioNeto, nImpuesto1, nImpuesto2, nImpuesto3, nVenta, nInsumo, nGasto, nManoObra, lImprimeArea, lImprime, nOrden, tObservacion, lCorte from CPEDIDO where tcodigopedido='" & tempPedido & "' "
        
         Cn.Execute Isql
         RsCombo.Requery
        
         cargarDatosCabecera (tempPedido)
         
         nMonto = Format(Calcular("select sum(nVenta) as Codigo FROM [" & sDetalle & "] where tEstadoItem = 'N'", Cn), "#,###,##0.00")
         VisualizaMonto
End Sub

Private Sub LimpiarData()
        Cn.Execute "delete " & sDetalle
        Cn.Execute "delete " & sComboDetalle
        Cn.Execute "delete " & sComboPropiedad
        Cn.Execute "delete " & sProductoPropiedad

        RsDetalle.Requery
        RsComboPropiedad.Requery
        RsProductoPropiedad.Requery
        Inicializar
        Screen.MousePointer = vbDefault
End Sub

Private Sub cmdAgrupacion_Click(Index As Integer)
   For i = 1 To 42
       cmdBoton(i).Visible = False
   Next i
   RsProducto.Filter = "tCajaRapida=" & sPrefijo & Index
   Shape.BorderColor = cmdAgrupacion(Index).backColor
   Shape1.BorderColor = cmdAgrupacion(Index).backColor
   Shape2.BorderColor = cmdAgrupacion(Index).backColor
   Shape3.BorderColor = cmdAgrupacion(Index).backColor
   Shape4.BorderColor = cmdAgrupacion(Index).backColor
   Shape5.BorderColor = cmdAgrupacion(Index).backColor
   Shape6.BorderColor = cmdAgrupacion(Index).backColor
   
   If Not RsProducto.EOF Then
      RsProducto.MoveFirst
      Do While Not RsProducto.EOF
         If RsProducto!nBotonRapido > 0 Then
            cmdBoton(RsProducto!nBotonRapido).Visible = True
            cmdBoton(RsProducto!nBotonRapido).Enabled = True
            cmdBoton(RsProducto!nBotonRapido).backColor = cmdAgrupacion(Index).backColor
            cmdBoton(RsProducto!nBotonRapido).Caption = RsProducto!tResumido
            If ((sTipoPedido = "01" And RsProducto!lLocal = False) Or (sTipoPedido = "02" And RsProducto!lDelivery = False) Or (sTipoPedido = "03" And RsProducto!lLlevar = False) Or (RsProducto!tUnidadNegocio <> sUnidadNegocio)) Then
               cmdBoton(RsProducto!nBotonRapido).Enabled = False
            End If
         End If
         RsProducto.MoveNext
      Loop
   End If
   If txtBarra.Visible = True Then
      txtBarra.SetFocus
   End If
'   sCajaRapida = sPrefijo & Index
End Sub

Private Sub cmdBoton_Click(Index As Integer)
   txtBarra.SetFocus
      
   RsProducto.MoveFirst
   RsProducto.Find "nbotonRapido = " & Trim(str(Index))
   sProducto = RsProducto!codigo
   
     'INSUMOCRITICO23
        If validadIngresoProducto(sProducto) = False Then
            Exit Sub
        End If
    'INSUMOCRITICO23
   
    If lBal And RsProducto!lBalanza Then
       Dim nResultado As Double
       nResultado = Pesar(nBalanzaPuerto)
       nResultado = Format(nResultado, "#,##0.000")
       If nResultado > 0 Then
          InsertaProducto nResultado
       End If
    Else
       nCantidad = 1
       InsertaProducto 1
    End If

   If IIf(IsNull(RsProducto!lPropiedad), False, RsProducto!lPropiedad) Then
      lPropiedad = True
   End If
End Sub

Private Sub cmdBusca_Click()
    sTipo = ""
    sTemp = ""
    'Isql = "select * from vOperador where lActivo = 1 Order by Descripcion "
    ListarOperadoresConFiltro (sProducto)
    Call ConfGrilla(2, frmBusquedaRapida.grdGrilla, "Código", 2, "Codigo", 1200, 2, 0, "", _
                                                    "Descripcion", 2, "Descripcion", 7000, 0, 0, "")
    frmBusquedaRapida.nPredeterm = 1
    frmBusquedaRapida.Show vbModal
    
    If wEnter = True Then
        Screen.MousePointer = vbHourglass
        For i = 1 To 13
            cmdOperador(i).backColor = vbButtonFace
        Next i
        RsOperador.MoveFirst
        RsOperador.Find "Codigo='" & sCodigo & "'"
        If Not RsOperador.EOF And RsOperador!nBoton > 0 Then
           cmdOperador(RsOperador!nBoton).backColor = vbRed
        End If
        AsignaPropiedad
        Screen.MousePointer = vbDefault
    End If
    txtBarra.SetFocus
End Sub

Private Sub cmdBuscar_Click(Index As Integer)
   Select Case Index
   
   Case Is = 0
      sTemp = ""
      Isql = "select * from vMozo where substring(Codigo,1,1)<>'*' AND lActivo = 1 Order by Descripcion"
      Call ConfGrilla(2, frmBusquedaRapida.grdGrilla, "Código", 2, "Codigo", 1200, 2, 0, "", _
                                                      "Descripcion", 2, "Descripcion", 7000, 0, 0, "")
                        
      frmBusquedaRapida.nPredeterm = 1
      frmBusquedaRapida.Show vbModal
      If wEnter = True Then
         sMozo = sCodigo
         txtTitulo.Caption = " Caja Rápida : " & sCaja & " Mesero: " & sDescrip
      End If
      tabProducto.Visible = True
      fraMozo.Visible = False
   
   Case Is = 1
      sTemp = ""
      sTipo = "Infhotel"
      Isql = "Select tPuntoVenta as Codigo, tDescripcion as Descripcion, nUltimoComanda, tmoneda" & _
             " From tPuntoVenta " & _
             " where tHotel='" & sHotel & "' AND lActivo=1 and lInforest=1"
      Call ConfGrilla(2, frmBusquedaRapida.grdGrilla, "Código", 2, "Codigo", 1200, 2, 0, "", _
                                                      "Descripcion", 2, "Descripcion", 7000, 0, 0, "")
                         
      frmBusquedaRapida.nPredeterm = 1
      frmBusquedaRapida.Show vbModal
      If wEnter = True Then
         sPuntoVenta = sCodigo
         rsPuntoVenta.MoveFirst
         rsPuntoVenta.Find "Codigo= '" & sCodigo & "'"
         cmdCabecera(0).Caption = rsPuntoVenta!Descripcion
      End If
      
      tabProducto.Visible = True
      fraPuntoVenta.Visible = False
      
   Case Is = 2 'ProductoCombo
         sTemp = ""
         If lComboGeneral Then
            Isql = "select * from vProducto where lActivo = 1 and " & IIf(sTipoPedido = "01", "lLocal=1", IIf(sTipoPedido = "02", "lDelivery=1", "lLlevar=1")) & " Order by Descripcion"
         Else
            Isql = "SELECT * FROM dbo.vProducto INNER JOIN dbo.TCOMBO ON dbo.vProducto.Codigo = dbo.TCOMBO.tCodigoProducto Where (dbo.vProducto.lActivo = 1) And tCombo='" & sProductoCombo & "' And " & IIf(sTipoPedido = "01", "lLocal=1", IIf(sTipoPedido = "02", "lDelivery=1", "lLlevar=1")) & " ORDER BY dbo.vProducto.Descripcion "
         End If
         Call ConfGrilla(5, frmBusquedaRapida.grdGrilla, "Grupo", 2, "Grupo", 1600, 0, 0, "", _
                                                         "Producto", 2, "Descripcion", 3600, 0, 0, "", _
                                                         "Precio", 2, "nPrecioVenta", 1000, 1, 0, "###,##0.00", _
                                                         "Bot", 2, "nBoton", 500, 1, 0, "", _
                                                         "SubGrupo", 2, "SubGrupo", 1500, 0, 0, "")
         frmBusquedaRapida.nPredeterm = 1
         frmBusquedaRapida.Show vbModal

         If wEnter Then
            sProducto = sCodigo
            Dim xxx As String
            xxx = RsProducto.Filter
            RsProducto.Filter = adFilterNone
            RsProducto.MoveFirst
            RsProducto.Find "Codigo = '" & sProducto & "'"
            
            nCCombo = Calcular("select sum(nCantidad) as Codigo " & _
                               "FROM " & sComboDetalle & "  WHERE    tItem='" & sitem & "'", Cn)
            If nCCombo < nCombo * RsDetalle!nCantidad Then
                            'Oscar Ortega----------------------------------------------
                            Dim oRsProductoDeCombo As Recordset
                            Set oRsProductoDeCombo = Obtener_ProductoDeCombo(RsDetalle!tCodigoProducto, sProducto)
                            If oRsProductoDeCombo.RecordCount > 0 Then
                                If IIf(IsNull(oRsProductoDeCombo!lUnico), False, oRsProductoDeCombo!lUnico) Then
                                    'Obtener Suma de cantidades
                                    Dim nCantidadEnElCombo As Integer
                                    nCantidadEnElCombo = ObtenerSumaCantidadesEnElCombo(sitem, oRsProductoDeCombo!tEtiqueta)
                                    'Suma de cantidades < que nCantidad
                                    If nCantidadEnElCombo < nCantidad Then
                                        InsertaCombo sProducto
                                    Else
                                        MsgBox "Solo es permitido " & nCantidad & " elemento(s) de tipo " & oRsProductoDeCombo!tEtiqueta, vbExclamation, sMensaje
                                    End If
                                Else
                                    InsertaCombo sProducto
                                End If
                            Else
                                InsertaCombo sProducto
                            End If

            Else
               MsgBox "La cantidad máxima de items para este producto es de " & nCombo * RsDetalle!nCantidad, vbExclamation, sMensaje
            End If
            RsProducto.Filter = IIf(xxx = "0", "", xxx)
          End If

    Case 8
                sTipo = ""
                sTemp = ""
                'Isql = "select * from vOperador where lActivo = 1 Order by Descripcion "
                'ListarOperadoresConFiltro (sProducto)
                Dim sPropiedad As String
                
                
                sPropiedad = dbTemporal(sCaja, 11, "Codigo", "nVarChar(20)", _
                                                    "tProducto", "nVarChar(10)", _
                                                    "Operador", "nVarChar(150)", _
                                                    "Descripcion", "nVarChar(150)", _
                                                    "tOperador", "nVarChar(2)", _
                                                    "nPrecio", "float", _
                                                    "tEnlace", "nVarChar(15)", _
                                                    "nInsumo", "float", _
                                                    "nGasto", "float", _
                                                    "nManoObra", "float", _
                                                    "tEstado", "nvarchar(50)")
                                                    
              If wAgregaCombo = False Then
                    Isql = " insert into " & sPropiedad & " select tCodigoPropiedad as Codigo,tProducto, " & _
                           " tOperador.tDetallado AS Operador, TPROPIEDAD.tDetallado as Propiedad, " & _
                           " TPROPIEDAD.tOperador, nPrecio, tEnlace, nInsumo, nGasto, nManoObra, 'Agregar' " & _
                           " FROM dbo.TPROPIEDAD LEFT OUTER JOIN dbo.TOPERADOR ON dbo.TPROPIEDAD.tOperador =  " & _
                           " dbo.TOPERADOR.tOperador " & _
                           " Where TOPERADOR.tOperador='" & xOperador & "' AND TPROPIEDAD.tProducto='" & sProducto & "' AND TPROPIEDAD.lActivo = 1 And IsNull(tOperador.lStockMenos, 0) <> 1"
                    
                    Cn.Execute Isql
                    If lAlmacen = True Then
                    
                        If Calcular("select count(*) as codigo from vOperador where lStockMenos=1  and Codigo='" & xOperador & "'", Cn) > 0 Then
                                Isql = "  insert into " & sPropiedad & " select '9999' as Codigo,  tCodigoPlato as tProducto, 'Sin' as Operador,  " & _
                                        "  tDetallado as Propiedad,    '" & xOperador & "' as tOperador, 0, " & sAlmacenMDB & ".dbo.DRECETAVENTA.tCodigoProducto as tEnlace, nCantidad * nPrecio as nInsumo, " & _
                                        "  0, 0 ,  'Agregar' FROM " & sAlmacenMDB & ".dbo.DRECETAVENTA INNER JOIN " & sAlmacenMDB & ".dbo.MRECETAVENTA ON " & sAlmacenMDB & ".dbo.DRECETAVENTA.tLocal = " & sAlmacenMDB & ".dbo.MRECETAVENTA.tLocal AND " & sAlmacenMDB & ".dbo.DRECETAVENTA.tRecetaVenta = " & sAlmacenMDB & ".dbo.MRECETAVENTA.tRecetaVenta  " & _
                                        " INNER JOIN " & sAlmacenMDB & ".dbo.TPRODUCTO ON " & sAlmacenMDB & ".dbo.DRECETAVENTA.tCodigoProducto = " & sAlmacenMDB & ".dbo.TPRODUCTO.tCodigoProducto   " & _
                                        " Where lNoDescargo = 1 and " & sAlmacenMDB & ".dbo.DRECETAVENTA.tLocal='" & sLocal & "' and tCodigoPlato='" & sProducto & "'"
           
                                Cn.Execute Isql
                        End If
                    
                    End If
                    
                    Isql = "  update " & sPropiedad & " SET tEstado='Quitar' from " & sPropiedad & " inner join (SELECT " & sProductoPropiedad & ".tCodigoPropiedad, " & sProductoPropiedad & ".tProducto , " & sProductoPropiedad & ".tEnlace     " & _
                           "  FROM  " & sProductoPropiedad & "   where  tItem='" & sitem & "' and  " & sProductoPropiedad & ".TPRODUCTO='" & sProducto & "') t1 on " & sPropiedad & ".Codigo=t1.tCodigoPropiedad and " & sPropiedad & ".tProducto=t1.tProducto and " & sPropiedad & ".tEnlace=t1.tEnlace "
 
                    
'
'                    Isql = " update " & sPropiedad & " SET tEstado='Quitar' from " & sPropiedad & " inner join (SELECT " & sProductoPropiedad & ".tCodigoPropiedad , " & sProductoPropiedad & ".tProducto   From  dbo.TOPERADOR INNER JOIN " & sProductoPropiedad & " INNER JOIN (  select tCodigoPropiedad as Codigo, TPROPIEDAD.tDetallado as Descripcion,   tProducto, TPROPIEDAD.tOperador, nPrecio, tEnlace, nInsumo,    nGasto, nManoObra, tOperador.tDetallado AS Operador    FROM dbo.TPROPIEDAD LEFT OUTER JOIN dbo.TOPERADOR ON dbo.TPROPIEDAD.tOperador =     dbo.TOPERADOR.tOperador Where TPROPIEDAD.lActivo = 1 And      IsNull(TOPERADOR.lStockMenos, 0) <> 1) " & _
'                           " T1 ON  " & sProductoPropiedad & ".tCodigoPropiedad = T1.Codigo AND " & sProductoPropiedad & ".tProducto = T1.tProducto AND " & sProductoPropiedad & ".tEnlace = T1.tEnlace ON dbo.tOperador.tOperador = T1.tOperador        COLLATE Modern_Spanish_CI_AS LEFT OUTER JOIN dbo.TPROPIEDAD ON dbo.TOPERADOR.tOperador = dbo.TPROPIEDAD.tOperador AND " & sProductoPropiedad & ".tCodigoPropiedad = dbo.TPROPIEDAD.tCodigoPropiedad AND " & sProductoPropiedad & ".tProducto = dbo.TPROPIEDAD.tProducto where   tItem='" & sitem & "' and " & sProductoPropiedad & ".TPRODUCTO='" & sProducto & "' and TOPERADOR.tOperador='" & xOperador & "' ) t1 on " & sPropiedad & ".Codigo=t1.tCodigoPropiedad and " & sPropiedad & ".tProducto=t1.tProducto"
                           
                    Cn.Execute Isql
                    Isql = "SELECT * FROM " & sPropiedad
             Else
             
                    Isql = " insert into " & sPropiedad & " select tCodigoPropiedad as Codigo,tProducto,tOperador.tDetallado AS Operador, TPROPIEDAD.tDetallado as Propiedad, TPROPIEDAD.tOperador, nPrecio, tEnlace, nInsumo, nGasto, " & _
                           " nManoObra, 'Agregar' FROM dbo.TPROPIEDAD LEFT OUTER JOIN dbo.TOPERADOR ON dbo.TPROPIEDAD.tOperador = dbo.TOPERADOR.tOperador Where TOPERADOR.tOperador='" & xOperador & "' AND TPROPIEDAD.tProducto='" & sCombo & "' AND TPROPIEDAD.lActivo = 1 And IsNull(tOperador.lStockMenos, 0) <> 1 "
                    Cn.Execute Isql


                    If lAlmacen = True Then
                    
                        If Calcular("select count(*) as codigo from vOperador where lStockMenos=1  and Codigo='" & xOperador & "'", Cn) > 0 Then
                                Isql = "  insert into " & sPropiedad & " select '9999' as Codigo,  tCodigoPlato as tProducto, 'Sin' as Operador,  " & _
                                        "  tDetallado as Propiedad,    '" & xOperador & "' as tOperador, 0, " & sAlmacenMDB & ".dbo.DRECETAVENTA.tCodigoProducto as tEnlace, nCantidad * nPrecio as nInsumo, " & _
                                        "  0, 0 ,  'Agregar' FROM " & sAlmacenMDB & ".dbo.DRECETAVENTA INNER JOIN " & sAlmacenMDB & ".dbo.MRECETAVENTA ON " & sAlmacenMDB & ".dbo.DRECETAVENTA.tLocal = " & sAlmacenMDB & ".dbo.MRECETAVENTA.tLocal AND " & sAlmacenMDB & ".dbo.DRECETAVENTA.tRecetaVenta = " & sAlmacenMDB & ".dbo.MRECETAVENTA.tRecetaVenta  " & _
                                        " INNER JOIN " & sAlmacenMDB & ".dbo.TPRODUCTO ON " & sAlmacenMDB & ".dbo.DRECETAVENTA.tCodigoProducto = " & sAlmacenMDB & ".dbo.TPRODUCTO.tCodigoProducto   " & _
                                        " Where lNoDescargo = 1 and " & sAlmacenMDB & ".dbo.DRECETAVENTA.tLocal='" & sLocal & "' and tCodigoPlato='" & sCombo & "'"
           
                                Cn.Execute Isql
                        End If
                    
                    End If

                    Isql = "  update " & sPropiedad & " SET tEstado='Quitar' from " & sPropiedad & " inner join (SELECT " & sComboPropiedad & ".tCodigoPropiedad, " & sComboPropiedad & ".tProducto , " & sComboPropiedad & ".tEnlace     " & _
                           "  FROM  " & sComboPropiedad & "   where  " & sComboPropiedad & ".tItem='" & sitem & "' and  " & sComboPropiedad & ".tItemcombo='" & xItem & "' and  " & sComboPropiedad & ".TPRODUCTO='" & sCombo & " ') t1 on " & sPropiedad & ".Codigo=t1.tCodigoPropiedad and " & sPropiedad & ".tProducto=t1.tProducto and " & sPropiedad & ".tEnlace=t1.tEnlace  "
 

'                    Isql = "update " & sPropiedad & " SET tEstado='Quitar' from " & sPropiedad & " inner join (SELECT   " & sComboPropiedad & ".tItem, " & sComboPropiedad & ".tItemCombo,tpropiedad.tCodigoPropiedad, " & sComboPropiedad & ".TPRODUCTO,  T1.Descripcion, T1.Operador  FROM         dbo.TOPERADOR INNER JOIN dbo.TPROPIEDAD ON   dbo.TOPERADOR.tOperador = dbo.TPROPIEDAD.tOperador RIGHT OUTER JOIN " & sComboPropiedad & " INNER JOIN    (select tCodigoPropiedad as Codigo, TPROPIEDAD.tDetallado as Descripcion, tProducto, TPROPIEDAD.tOperador, nPrecio, tEnlace, nInsumo, nGasto, nManoObra, tOperador.tDetallado AS Operador FROM dbo.TPROPIEDAD LEFT OUTER JOIN dbo.TOPERADOR ON dbo.TPROPIEDAD.tOperador = dbo.tOperador.tOperador Where TPROPIEDAD.lActivo = 1 And IsNull(TOPERADOR.lStockMenos, 0) <> 1) " & _
'                            " T1 ON " & sComboPropiedad & ".tCodigoPropiedad = T1.Codigo AND " & sComboPropiedad & ".tProducto =  T1.tProducto AND " & sComboPropiedad & ".tEnlace = T1.tEnlace ON dbo.TOPERADOR.tOperador = T1.tOperador COLLATE Modern_Spanish_CI_AS AND dbo.TPROPIEDAD.tCodigoPropiedad =" & sComboPropiedad & ".tCodigoPropiedad AND dbo.TPROPIEDAD.tProducto = " & sComboPropiedad & ".tProducto
                        'where TOPERADOR.tOperador='" & xOperador & "' and  " & sComboPropiedad & ".tItem='" & sitem & "' and " & sComboPropiedad & ".tItemCombo='" & xItem & "' and " & sComboPropiedad & ".tProducto='" & sCombo & "' ) t1 on " & sPropiedad & ".Codigo=t1.tCodigoPropiedad and " & sPropiedad & ".tProducto=t1.tProducto"

                    Cn.Execute Isql
             
                    Isql = "SELECT * FROM " & sPropiedad
             End If

                Call ConfGrilla(3, frmBusquedaRapida.grdGrilla, "Código", 2, "Codigo", 1200, 2, 0, "", _
                                                                "Estado", 2, "tEstado", 1500, 0, 0, "", _
                                                                "Descripcion", 2, "Descripcion", 5500, 0, 0, "")
                                                                
'                frmBusquedaRapida.nPredeterm = 1
'                 If tipoBusqueda = "BusquedaCajaRapida" Then
'                    tenlacebusqueda = IIf(RsGrilla.EOF = True, "", RsGrilla!tenlace)
'                    tipoBusqueda = ""
'                End If
                frmBusquedaRapida.tipoBusqueda = "BusquedaCajaRapida"
                
                frmBusquedaRapida.Show vbModal
                
                If wEnter = True Then
                  '  Screen.MousePointer = vbHourglass
                    If wAgregaCombo = False Then
                            If Calcular("SELECT COUNT(*) AS CODIGO FROM " & sProductoPropiedad & " WHERE tItem='" & sitem & "' AND tCodigoPropiedad='" & sCodigo & "' AND TPRODUCTO='" & sProducto & "' and tEnlace='" & tenlacebusqueda & "'", Cn) = 0 Then
                                    wAgregarPropiedad = True
                            Else
                                    wAgregarPropiedad = False
                            End If
                             AgregarPropiedadBusqueda sCodigo, sDescrip
                    Else
                    
                           If Calcular("SELECT COUNT(*) AS CODIGO FROM " & sComboPropiedad & " WHERE tItem='" & sitem & "' AND tCodigoPropiedad='" & sCodigo & "' AND TPRODUCTO='" & sCombo & "' and titemcombo='" & xItem & "'", Cn) = 0 Then
                                    wAgregarPropiedad = True
                            Else
                                    wAgregarPropiedad = False
                            End If
                           AgregarPropiedadBusqueda sCodigo, sDescrip
                    End If
                    
                    
                   ' Screen.MousePointer = vbDefault
                End If
              '  txtBarra.SetFocus
        
            Case Is = 3 'Origen de ventas
                sTemp = ""
                Isql = "select * from vOrigenVenta where Activo = 1 Order by Descripcion"
                Call ConfGrilla(2, frmBusquedaRapida.grdGrilla, "CodOrigenVenta", 2, "CodOrigenVenta", 1200, 2, 0, "", _
                                                                "Descripcion", 2, "Descripcion", 7000, 0, 0, "")
                frmBusquedaRapida.nPredeterm = 1
                frmBusquedaRapida.tipoBusqueda = "OrigenVentas"
                frmBusquedaRapida.Show vbModal
                If wEnter = True Then
                   vOrigenVentas = sCodigo
                End If
                txtBarra.SetFocus
        

          
   End Select
   
   txtBarra.SetFocus
End Sub
Public Sub AgregarPropiedadBusqueda(ByVal CodigoPropiedad As String, ByVal DescripcionPropiedad As String)
 

   '  HabilitaTimerColor (False)
  Dim ncantidadPropiedad As Double
   Dim Cantidad As Double
    Dim nInsumo As Double
    Dim nGasto As Double
    Dim nMObra As Double
    Dim nBotonPropiedad As Double
    nBotonPropiedad = 999
    RsPropiedad.MoveFirst
    RsPropiedad.Find ("Descripcion = '" & DescripcionPropiedad & "'")
     RsPropiedad.Find ("Codigo = '" & CodigoPropiedad & "'")
   
     If Not (RsOperador.EOF Or RsOperador.BOF) Then
        nOperadorPropiedad = Calcular("select isnull(ncontrol,0) as codigo from voperador where codigo='" & RsOperador!codigo & "'", Cn)
     End If
    Dim k As Integer
    For k = 1 To 19
        If cmdPropiedad(k).Caption = DescripcionPropiedad Then
        nBotonPropiedad = k
        Exit For
        End If
    Next k
    
    If nBotonPropiedad <> 999 Then
        If cmdPropiedad(nBotonPropiedad).FontBold = True Then
            cmdPropiedad(nBotonPropiedad).FontBold = False
        Else
            cmdPropiedad(nBotonPropiedad).FontBold = True

        End If
    
    End If
    
                 
        If wAgregarPropiedad = False Then
           ' cmdPropiedad(nBotonPropiedad).FontBold = False
            If Not RsPropiedad.EOF Then
                If wAgregaCombo Then
                    Cantidad = Calcular("select isnull(ncantidad,1) as codigo from " & sComboPropiedad & " where   titem='" & sitem & "' and titemcombo='" & xItem & "' and  tproducto='" & sCombo & "' and tcodigopropiedad='" & RsPropiedad!codigo & "' ", Cn)

                    Cn.Execute "delete " & sComboPropiedad & " where tItem = '" & sitem & "' and tItemCombo='" & xItem & "' and tProducto='" & sCombo & "' and tCodigoPropiedad='" & RsPropiedad!codigo & "'"
                Else
                    Cantidad = Calcular("select isnull(ncantidad,1) as codigo from " & sProductoPropiedad & " where   titem='" & sitem & "' and tproducto='" & sProducto & "' and tcodigopropiedad='" & RsPropiedad!codigo & "' and tenlace='" & RsPropiedad!tEnlace & "'", Cn)
                
                    Cn.Execute "delete " & sProductoPropiedad & "  where tItem = '" & sitem & "' and tProducto='" & sProducto & "' and tCodigoPropiedad='" & RsPropiedad!codigo & "' and tEnlace='" & RsPropiedad!tEnlace & "'"
                     If RsPropiedad!nPrecio <> 0 Then
                            nMonto = CambiaPrecio(nPVenta - RsPropiedad!nPrecio)
                            txtMonto.Caption = Format(nMonto, "###,##0.00")
                     End If
                End If
                If Cantidad <> 1 Then
                           lblResumen.Text = Replace(lblResumen.Text, RsOperador!Descripcion & " " & DescripcionPropiedad & ": (" & Cantidad & "), ", "")
                Else
                           lblResumen.Text = Replace(lblResumen.Text, RsOperador!Descripcion & " " & DescripcionPropiedad & ", ", "")
                End If
                
'                lblResumen.Text = Replace(lblResumen.Text, RsOperador!Descripcion & " " & DescripcionPropiedad & ", ", "")
                
            End If
        Else
        
            ncantidadPropiedad = 1
            If RsPropiedad!lsolicitacantidad = 1 Or RsPropiedad!lsolicitacantidad = True Then
                sTipo = "Prepintado"
            
                sCodigo = ncantidadPropiedad
            
                frmNumPad.Show vbModal
                If wEnter And Val(sDescrip) > 0 Then
            
                            ncantidadPropiedad = sDescrip
                        
                End If
            End If
        
            If nOperadorPropiedad > 0 Then
                If wAgregaCombo Then
                       Isql = "SELECT COUNT(" & sComboPropiedad & ".tCodigoPropiedad) AS codigo " & _
                              "FROM " & sComboPropiedad & " INNER JOIN dbo.TPROPIEDAD ON " & sComboPropiedad & ".tCodigoPropiedad = dbo.TPROPIEDAD.tCodigoPropiedad AND " & sComboPropiedad & ".tProducto = dbo.TPROPIEDAD.tProducto " & _
                              "where tItem = '" & sitem & "' and tItemCombo='" & xItem & "' and " & sComboPropiedad & ".tProducto='" & sCombo & "'  and tOperador='" & RsOperador!codigo & "'"
                   If nOperadorPropiedad <= Calcular(Isql, Cn) Then
                      MsgBox "Ha llegado a la Cantidad máxima de " & nOperadorPropiedad & " Propiedad(es) por Operador", vbExclamation, sMensaje
                      Exit Sub
                   End If
                Else
                    Isql = "SELECT COUNT(" & sProductoPropiedad & ".tCodigoPropiedad) AS codigo FROM " & sProductoPropiedad & " INNER JOIN " & _
                    "dbo.TPROPIEDAD ON " & sProductoPropiedad & ".tCodigoPropiedad = dbo.TPROPIEDAD.tCodigoPropiedad AND " & sProductoPropiedad & ".tProducto = dbo.TPROPIEDAD.tProducto " & _
                    "where tItem = '" & sitem & "' and tOperador='" & RsOperador!codigo & "'"
                    If nOperadorPropiedad <= Calcular(Isql, Cn) Then
                       MsgBox "Ha llegado a la Cantidad máxima de " & nOperadorPropiedad & " Propiedad(es) por Operador", vbExclamation, sMensaje
                       Exit Sub
                    End If
                End If
            End If
       
    
          '  cmdPropiedad(nBotonPropiedad).FontBold = True
            If Not RsPropiedad.EOF Then
               nInsumo = IIf(IsNull(RsPropiedad!nInsumo), 0, RsPropiedad!nInsumo)
               nGasto = IIf(IsNull(RsPropiedad!nGasto), 0, RsPropiedad!nGasto)
               nMObra = IIf(IsNull(RsPropiedad!nManoObra), 0, RsPropiedad!nManoObra)
               If wAgregaCombo Then
                    Cn.Execute "Insert into " & sComboPropiedad & " values ('" & sitem & "', '" & xItem & "', '" & RsPropiedad!codigo & "', '" & sCombo & "', '" & RsPropiedad!tEnlace & "', " & IIf(IsNull(RsPropiedad!nInsumo), 0, ncantidadPropiedad * RsPropiedad!nInsumo) & ", " & IIf(IsNull(RsPropiedad!nGasto), 0, ncantidadPropiedad * RsPropiedad!nGasto) & ", " & IIf(IsNull(RsPropiedad!nManoObra), 0, ncantidadPropiedad * RsPropiedad!nManoObra) & ", " & ncantidadPropiedad & ", " & IIf(IsNull(RsPropiedad!nInsumo), 0, RsPropiedad!nInsumo) & ", " & IIf(IsNull(RsPropiedad!nGasto), 0, RsPropiedad!nGasto) & ", " & IIf(IsNull(RsPropiedad!nManoObra), 0, RsPropiedad!nManoObra) & ") "
                    
                Else
                    Cn.Execute "Insert into " & sProductoPropiedad & " values ('" & sitem & "', '" & RsPropiedad!codigo & "', '" & sProducto & "', '" & RsPropiedad!tEnlace & "', " & IIf(IsNull(RsPropiedad!nInsumo), 0, ncantidadPropiedad * RsPropiedad!nInsumo) & ", " & IIf(IsNull(RsPropiedad!nGasto), 0, ncantidadPropiedad * RsPropiedad!nGasto) & ", " & IIf(IsNull(RsPropiedad!nManoObra), 0, ncantidadPropiedad * RsPropiedad!nManoObra) & ", " & ncantidadPropiedad & "," & IIf(IsNull(RsPropiedad!nInsumo), 0, RsPropiedad!nInsumo) & ", " & IIf(IsNull(RsPropiedad!nGasto), 0, RsPropiedad!nGasto) & ", " & IIf(IsNull(RsPropiedad!nManoObra), 0, RsPropiedad!nManoObra) & " )"
                    If RsPropiedad!nPrecio <> 0 Then
                       nMonto = CambiaPrecio(nPVenta + (RsPropiedad!nPrecio * ncantidadPropiedad))
                       txtMonto.Caption = Format(nMonto, "###,##0.00")
                    End If

                End If
          End If
  '        lblResumen.Text = lblResumen.Text & RsOperador!Descripcion & " " & DescripcionPropiedad & ", "
        If ncantidadPropiedad <> 1 Then
          
                lblResumen.Text = lblResumen.Text & RsOperador!Descripcion & " " & DescripcionPropiedad & ": (" & ncantidadPropiedad & "), "
          Else
                lblResumen.Text = lblResumen.Text & RsOperador!Descripcion & " " & DescripcionPropiedad & ", "
          End If
          
    End If
    
    
    If wAgregaCombo Then
       RsComboPropiedad.Requery
    Else
       RsProductoPropiedad.Requery
    End If
    
 '  HabilitaTimerColor (True)
End Sub

Private Sub cmdCabecera_Click(Index As Integer)
   Select Case Index
   
      Case Is = 5 'Canal5
           If RsDetalle.RecordCount > 0 Then
                MsgBox "No se puede cambiar el canal de venta"
           Else
                sTipoPedido = "05"
                sMotorizado = ""
                
                cmdCabecera(1).FontBold = False
                cmdCabecera(2).FontBold = False
                cmdCabecera(3).FontBold = False
                cmdCabecera(4).FontBold = False
                cmdCabecera(5).FontBold = True
           End If
   
      Case Is = 4 'Canal 4
           If RsDetalle.RecordCount > 0 Then
                MsgBox "No se puede cambiar el canal de venta"
           Else
                sTipoPedido = "04"
                sMotorizado = ""
                
                cmdCabecera(1).FontBold = False
                cmdCabecera(2).FontBold = False
                cmdCabecera(3).FontBold = False
                cmdCabecera(4).FontBold = True
                cmdCabecera(5).FontBold = False
           End If
   
      Case Is = 3 'Para llevar
      
            'origen de ventas
            Me.fraOrigenVentas.Visible = False
            '--------------------------------
      
           If RsDetalle.RecordCount > 0 Then
                MsgBox "No se puede cambiar el canal de venta"
           Else
               sTipoPedido = "03"
               sMotorizado = ""
    
               cmdCabecera(1).FontBold = False
               cmdCabecera(2).FontBold = False
               cmdCabecera(3).FontBold = True
               cmdCabecera(4).FontBold = False
               cmdCabecera(5).FontBold = False
           End If
      Case Is = 1 'En el Local
            
            'origen de ventas
            Me.fraOrigenVentas.Visible = False
            '--------------------------------
           If RsDetalle.RecordCount > 0 Then
                MsgBox "No se puede cambiar el canal de venta"
           Else
                sTipoPedido = "01"
                sMotorizado = ""
                
                cmdCabecera(1).FontBold = True
                cmdCabecera(2).FontBold = False
                cmdCabecera(3).FontBold = False
                cmdCabecera(4).FontBold = False
                cmdCabecera(5).FontBold = False
           End If
                        
      Case Is = 2 'En Delivery
           If RsDetalle.RecordCount > 0 Then
                MsgBox "No se puede cambiar el canal de venta"
           Else
                sTipoPedido = "02"
                sMotorizado = ""
                
                cmdCabecera(1).FontBold = False
                cmdCabecera(2).FontBold = True
                cmdCabecera(3).FontBold = False
                cmdCabecera(4).FontBold = False
                cmdCabecera(5).FontBold = False
                
                'origen de ventas
                RsCanalesVenta.Filter = "tCodigoCanalVenta = '" & sTipoPedido & "'"
                lOrigenVentas = IIf(IsNull(RsCanalesVenta!lCanalDelivery), False, RsCanalesVenta!lCanalDelivery)
                
                If lOrigenVentas Then
                    Me.fraOrigenVentas.Visible = True
                    Else
                        Me.fraOrigenVentas.Visible = False
                End If
                
           End If

           
      Case Is = 0 'Punto de Venta
            tabProducto.Visible = False
            fraPuntoVenta.Visible = True
            
      Case Is = 6
            frmPrograma.Show vbModal
            If wEnter = True Then
                txtFechaEntrega.Caption = sCodigo
            Else
                txtFechaEntrega.Caption = ""
            End If
            
      Case Is = 7
      
            If RsDetalle.RecordCount > 0 And Pedido = "" Then
                Exit Sub
            End If
      
            If lPasswordImportarPedido Then
               If Supervisor("15") = False Then
                  MsgBox "Clave no permitida", vbExclamation, sMensaje
                  Exit Sub
               End If
            End If
            
            If lFiltroTipoPedido Then
                Isql = "select *, Caso = case vpedidoGrilla.tCaja when '" & sCaja & "' then 'Exportar' ELSE 'Importar' END " & _
                       "from vPedidoGrilla INNER JOIN dbo.MPEDIDO ON dbo.vPedidoGrilla.Codigo = dbo.MPEDIDO.tCodigoPedido " & _
                       "where MPEDIDO.tTipoPedido='" & sTipoPedidoPD & "' and tCodigoPedido not in (select distinct dbo.DPEDIDO.tcodigopedido FROM dbo.MPEDIDO INNER JOIN dbo.DPEDIDO ON dbo.MPEDIDO.tCodigoPedido = dbo.DPEDIDO.tCodigoPedido where  tEstadoPedido='01' and (tFacturado = 'F' or tFacturado='P')) and " & _
                       "vpedidoGrilla.tEstadoPedido ='01' and " & _
                       "(vpedidoGrilla.tCaja <>'" & sCaja & "' or (vpedidoGrilla.tCaja='" & sCaja & "' and len(ltrim(tCajaAnterior))<>0 )) " & _
                       "order by Mesa, vPedidoGrilla.tObservacion"
            Else
                If lMCPV Then
                    Isql = "select *, 'Importar' as Caso " & _
                           "from vPedidoGrilla INNER JOIN dbo.MPEDIDO ON dbo.vPedidoGrilla.Codigo = dbo.MPEDIDO.tCodigoPedido " & _
                           "where tCodigoPedido not in (select distinct dbo.DPEDIDO.tcodigopedido FROM dbo.MPEDIDO INNER JOIN dbo.DPEDIDO ON dbo.MPEDIDO.tCodigoPedido = dbo.DPEDIDO.tCodigoPedido where  tEstadoPedido='01' and (tFacturado = 'F' or tFacturado='P')) and " & _
                           "vpedidoGrilla.tEstadoPedido ='01' and " & _
                           "vpedidoGrilla.tUsuario <>'" & sUsuario & "' " & _
                           "order by Mesa, vPedidoGrilla.tObservacion"
                Else
                    Isql = "select *, Caso = case vpedidoGrilla.tCaja when '" & sCaja & "' then 'Exportar' ELSE 'Importar' END " & _
                           "from vPedidoGrilla INNER JOIN dbo.MPEDIDO ON dbo.vPedidoGrilla.Codigo = dbo.MPEDIDO.tCodigoPedido " & _
                           "where tCodigoPedido not in (select distinct dbo.DPEDIDO.tcodigopedido FROM dbo.MPEDIDO INNER JOIN dbo.DPEDIDO ON dbo.MPEDIDO.tCodigoPedido = dbo.DPEDIDO.tCodigoPedido where  tEstadoPedido='01' and (tFacturado = 'F' or tFacturado='P')) and " & _
                           "vpedidoGrilla.tEstadoPedido ='01' and " & _
                           "(vpedidoGrilla.tCaja <>'" & sCaja & "' or (vpedidoGrilla.tCaja='" & sCaja & "' and len(ltrim(tCajaAnterior))<>0 )) " & _
                           "order by Mesa, vPedidoGrilla.tObservacion"
                End If
            End If
            
            Call ConfGrilla(8, frmBusquedaRapida.grdGrilla, "Función", 2, "Caso", 800, 0, 0, "", _
                                                            "Caja", 2, "tCaja", 550, 2, 0, "", _
                                                            "Mesa", 2, "Mesa", 1000, 0, 0, "", _
                                                            "Observacion", 2, "tObservacion", 1800, 0, 0, "", _
                                                            "Pedido", 2, "Descripcion", 1000, 0, 0, "", _
                                                            "Monto", 2, "Suma", 900, 1, 0, "###,##0.00", _
                                                            "Mesero", 2, "Mozo", 1100, 0, 0, "", _
                                                            "Usuario", 2, "tUsuario", 1100, 0, 0, "")
                                                            
            If lBuscaPedidoNumero = True Then
                frmBusquedaRapida.nPredeterm = 4
            Else
                frmBusquedaRapida.nPredeterm = 2
            End If
                
            
            frmBusquedaRapida.Show vbModal
            
            If wEnter Then
                                                     
               sPedido = sCodigo
               
               'Refrescar
               Dim RsRefresca As Recordset
               Set RsRefresca = Lib.OpenRecordset("select tTurno, tCaja, tCajaAnterior, tTurnoAnterior from MPEDIDO where tCodigoPedido='" & sPedido & "'", Cn)
  
               If lMCPV Then
                  ActualizaPedido
                  Cn.Execute "update MPEDIDO set tUsuario = '" & sUsuario & "', tTurno='" & sTurno & "' where tCodigoPedido='" & sPedido & "' "
                  LimpiarData
                  
               ElseIf RsRefresca!tTurno = "MOZO" Then 'Importar desde Mozos
                  If lInfhotel Then
                     Cn.Execute "update MPEDIDO set tPuntoVenta='" & sPuntoVentaInfhotel & "', tCaja = '" & sCaja & "', tTurno='" & sTurno & "' where tCodigoPedido='" & sPedido & "' "
                  Else
                     Cn.Execute "update MPEDIDO set tCaja = '" & sCaja & "', tTurno='" & sTurno & "' where tCodigoPedido='" & sPedido & "' "
                  End If
                  
                  If RsDetalle.RecordCount > 0 And Pedido <> sPedido Then
                     Exit Sub
                  End If
                  
                  If Pedido <> "" Then
                     Exit Sub
                  End If
                         
                  ActualizaPedido
                  Cn.Execute "update TCAJA set lRefresca=1 where tCaja='" & RsRefresca!tCaja & "'"
                  
                  Pedido = sPedido
                        
                  ImportarPedido Pedido
                  
               ElseIf RsRefresca!tTurno = sTurno And Not IsNull(RsRefresca!tTurnoAnterior) Then  'Exportar Mozo
                    
                  If sCaja = RsRefresca!tCajaAnterior Then
                        Exit Sub
                  End If
                  
                  If RsDetalle.RecordCount > 0 And Pedido <> sPedido Then
                        Exit Sub
                  End If

                  ActualizaPedido
                  Cn.Execute "update MPEDIDO set tTurno = '" & RsRefresca!tTurnoAnterior & "', tCaja='" & RsRefresca!tCajaAnterior & "' where tCodigoPedido='" & sPedido & "' "
               
                  LimpiarData
                    
               ElseIf RsRefresca!tTurno = sTurno And IsNull(RsRefresca!tTurnoAnterior) Then 'Exportar
                      
                  If sCaja = RsRefresca!tCajaAnterior Then
                        Exit Sub
                  End If
                  
                  If RsDetalle.RecordCount > 0 And Pedido <> sPedido Then
                      Exit Sub
                  End If
                  
                  ActualizaPedido
               
                  Cn.Execute "update TCAJA set lRefresca=1 where tCaja='" & RsRefresca!tCajaAnterior & "'"
                  Cn.Execute "update MPEDIDO set tTurno = 'MOZO', tCaja='" & RsRefresca!tCajaAnterior & "' where tCodigoPedido='" & sPedido & "' "
                  
                  LimpiarData

               Else  'Importar
                      If Calcular("select count(ddocumento.tDocumento) as Codigo from DDOCUMENTO inner join mdocumento on ddocumento.tdocumento= mdocumento.tdocumento where tCodigoPedido='" & sPedido & "' and mdocumento.testadodocumento<>'04'", Cn) > 0 Then
                         MsgBox "Error: No se puede importar pedido con Documentos", vbExclamation, sMensaje
                         Exit Sub
                      Else
                                 
                         If RsDetalle.RecordCount > 0 And Pedido <> sPedido Then
                            Exit Sub
                         End If
                         
                         If Pedido <> "" Then
                            Exit Sub
                         End If
                            
                         Cn.Execute "update MPEDIDO set tCaja = '" & sCaja & "', tTurno='" & sTurno & "' where tCodigoPedido='" & sPedido & "' "
                         Pedido = sPedido
                        
                         ImportarPedido Pedido
                        
                      End If
                End If
            
            End If
            
    
   End Select
   cmdEtiqueta_Click (1)
   txtBarra.SetFocus
End Sub


Private Sub cargarDatosCabecera(Pedido As String)

    Dim RsCabeceraPedido As Recordset
    Dim xTipoPedido As String
    
    Isql = "SELECT M.tCodigoPedido,T.tCodigoDelivery,T.tNombre + ' ' + t.tApellido as cliente,TM.tDetallado,M.fProgramacion,M.tEntregarA, M.tObservacion, M.tTipoPedido, M.nDescuento, M.tMozo, M.tDescuento, M.tClienteDelivery " & _
           "FROM MPEDIDO M LEFT JOIN TDELIVERY T ON M.tClienteDelivery = T.tCodigoDelivery " & _
           "LEFT JOIN TMESA TM ON M.tMesa = TM.tCodigoMesa " & _
           "WHERE M.tCodigoPedido = '" & Pedido & "'"
    Set RsCabeceraPedido = Lib.OpenRecordset(Isql, Cn)
    
    txtTelefono.Caption = IIf(IsNull(RsCabeceraPedido!tCodigoDelivery), "", RsCabeceraPedido!tCodigoDelivery)
    txtCliente.Caption = IIf(IsNull(RsCabeceraPedido!Cliente), "", RsCabeceraPedido!Cliente)
    txtObservacion.Caption = IIf(IsNull(RsCabeceraPedido!tObservacion), "", RsCabeceraPedido!tObservacion)
    txtEntregar.Caption = IIf(IsNull(RsCabeceraPedido!TEntregarA), "", RsCabeceraPedido!TEntregarA)
    txtFechaEntrega.Caption = IIf(IsNull(RsCabeceraPedido!fProgramacion), "", RsCabeceraPedido!fProgramacion)
    xTipoPedido = IIf(IsNull(RsCabeceraPedido!tTipoPedido), "", RsCabeceraPedido!tTipoPedido)
    
    xDescuento = IIf(IsNull(RsCabeceraPedido!nDescuento), 0, RsCabeceraPedido!nDescuento)
    sCodigoDescuento = IIf(IsNull(RsCabeceraPedido!tDescuento), 0, RsCabeceraPedido!tDescuento)
    sMozo = IIf(IsNull(RsCabeceraPedido!tMozo), "", RsCabeceraPedido!tMozo)
    sClienteFrecuente = IIf(IsNull(RsCabeceraPedido!tClienteDelivery), "", RsCabeceraPedido!tClienteDelivery)
    
    Select Case xTipoPedido
        Case "01"
            sTipoPedido = "01"
            cmdCabecera(1).FontBold = True
            cmdCabecera(2).FontBold = False
            cmdCabecera(3).FontBold = False
            cmdCabecera(4).FontBold = False
            cmdCabecera(5).FontBold = False
        Case "02"
            sTipoPedido = "02"
            cmdCabecera(2).FontBold = True
            cmdCabecera(1).FontBold = False
            cmdCabecera(3).FontBold = False
            cmdCabecera(4).FontBold = False
            cmdCabecera(5).FontBold = False
        Case "03"
            sTipoPedido = "03"
            cmdCabecera(3).FontBold = True
            cmdCabecera(1).FontBold = False
            cmdCabecera(2).FontBold = False
            cmdCabecera(4).FontBold = False
            cmdCabecera(5).FontBold = False
        Case "04"
            sTipoPedido = "04"
            cmdCabecera(4).FontBold = True
            cmdCabecera(1).FontBold = False
            cmdCabecera(2).FontBold = False
            cmdCabecera(3).FontBold = False
            cmdCabecera(5).FontBold = False
        Case "05"
            sTipoPedido = "05"
            cmdCabecera(5).FontBold = True
            cmdCabecera(1).FontBold = False
            cmdCabecera(2).FontBold = False
            cmdCabecera(3).FontBold = False
            cmdCabecera(4).FontBold = False
    End Select
    
    
End Sub

Private Sub cmdCombo_Click(Index As Integer)
   txtBarra.SetFocus
   Dim nPos As Integer
   Dim nOrd As Integer
   Select Case Index
          Case Is = 0 ' Salir
               fraCombo.Visible = False
               fraProductoCombo.Visible = False
               wAgregaCombo = False
               ActivaCabecera True
               
                If fraPropiedad.Visible = True Then
                  cmdOpcion_Click (6)
               End If
               AsignaProducto
               RsDetalle.Requery
          
          Case Is = 1 ' Elimina
               If RsCombo.RecordCount = 0 Then
                  Exit Sub
               End If
                If obtieneEliminaItemFijoCombo(RsCombo.Fields("tproducto"), RsCombo.Fields("tproductocombo")) = True Then
                    MsgBox "No se puede quitar este producto del combo. Consulte con el Administrador"
                    Exit Sub
               End If
               If MsgBox("Seguro de Eliminar el Producto?", vbQuestion + vbYesNo, sMensaje) = vbYes Then
                  sUsuarioAutoriza = sUsuario
                  If (lPassword And lPrinter = False) Or (lPassword And lPrinter And RsCombo!lImprime) Then
                        tUsuActua = sUsuario
                     If Supervisor("02") = False Then
                        MsgBox "Clave no permitida", vbExclamation, sMensaje
                        Exit Sub
                     End If
                     sUsuarioAutoriza = sVar1
                     sUsuario = tUsuActua
                  End If
                                                                            
                  If lPrinter = False Or (lPrinter = True And RsCombo!lImprime) Then
                     'Impresion del Pedidos Anulados
                     Isql = "SELECT TPRODUCTO_1.tDetallado AS Producto, dbo.vSalon.tResumido + ' - ' + dbo.TMESA.tResumido AS Mesa, dbo.TPRODUCTOAREA.tArea, dbo.MPEDIDO.tTipoPedido AS TipoPedido, dbo.MPEDIDO.nAdulto, dbo.MPEDIDO.lPrioridad AS Prioridad, dbo.MPEDIDO.tObservacion AS Observacion, dbo.vMozo.Descripcion AS Mozo, dbo.CPEDIDO.nCantidad AS nCombo, dbo.CPEDIDO.tItem, dbo.CPEDIDO.tItemCombo, dbo.CPEDIDO.tObservacion AS tObservacionCombo, TPRODUCTO_2.tDetallado AS Combo, dbo.vDelivery.Cliente " & _
                            "FROM dbo.TPRODUCTO TPRODUCTO_2 LEFT OUTER JOIN dbo.TPRODUCTOAREA ON TPRODUCTO_2.tCodigoProducto = dbo.TPRODUCTOAREA.tCodigoProducto RIGHT OUTER JOIN dbo.TMESA LEFT OUTER JOIN dbo.vSalon ON dbo.TMESA.tSalon = dbo.vSalon.Codigo RIGHT OUTER JOIN dbo.vMozo RIGHT OUTER JOIN dbo.vDelivery RIGHT OUTER JOIN dbo.MPEDIDO ON dbo.vDelivery.Codigo = dbo.MPEDIDO.tClienteDelivery LEFT OUTER JOIN dbo.CPEDIDO ON dbo.MPEDIDO.tCodigoPedido = dbo.CPEDIDO.tCodigoPedido ON dbo.vMozo.Codigo = dbo.MPEDIDO.tMozo ON " & _
                            "dbo.TMESA.tCodigoMesa = dbo.MPEDIDO.tMesa ON TPRODUCTO_2.tCodigoProducto = dbo.CPEDIDO.tProductoCombo LEFT OUTER JOIN dbo.TPRODUCTO TPRODUCTO_1 ON dbo.CPEDIDO.tProducto = TPRODUCTO_1.tCodigoProducto " & _
                            "Where dbo.CPEDIDO.lImprime = 1 And dbo.CPEDIDO.lImprimeArea = 1 and dbo.CPEDIDO.tCodigoPedido = '" & Pedido & "' and dbo.CPEDIDO.tItem ='" & RsCombo!tItem & "' and dbo.CPEDIDO.tItemCombo='" & RsCombo!tItemCombo & "' ORDER BY dbo.CPEDIDO.tItem"
              
                     Set RsImpresion = Lib.OpenRecordset(Isql, Cn)
                     Dim i As Integer
                     If RsImpresion.RecordCount = 0 Then
                        LimpiaRs
                     Else
                     
                        If Not RsImpresion.EOF Then
                           RsArea.MoveFirst
                           For i = 1 To RsArea.RecordCount
                               RsImpresion.Filter = "tArea='" & RsArea!tArea & "'"
                               If RsArea!tIcono = "" Or RsArea!nValor = 1 Then
                                  If RsImpresion.RecordCount <> 0 Then
                                     RsImpresion.MoveFirst
                                     sPedido = Pedido
                                     ImprimePedido RsImpresion, "A", RsArea!timpresora, RsArea!Area, False, RsProductoPropiedad, RsComboPropiedad, "Rapido"
                                     sPedido = ""
                                  End If
                               End If
                               RsArea.MoveNext
                           Next i
                        End If
                      End If
                      LimpiaRs
                   End If
                        
                        'Oscar Ortega----------------------------------------------
                        Dim RstCombo2 As Recordset
                        Isql = "Select c.nCantidad, t.nAumento From CPEDIDO As c Left Join TCombo as t On c.tProducto = t.tCombo And c.tProductoCombo = t.tCodigoProducto Where c.tCodigoPedido = '" & Pedido & "' And c.tItem = '" & sitem & "' And c.tItemCombo = '" & xItem & "'"
                         Isql = "Select t.nAumento From [" & sComboDetalle & "] As c Left Join TCombo as t On c.tProducto = t.tCombo And c.tProductoCombo = t.tCodigoProducto Where c.tItem = '" & sitem & "' And c.tItemCombo = '" & RsCombo!tItemCombo & "'"
                        Set RstCombo2 = Lib.OpenRecordset(Isql, Cn)
                        If RstCombo2.RecordCount > 0 Then
                                If IIf(IsNull(RstCombo2!nAumento), 0, RstCombo2!nAumento) > 0 Then
                                        txtMonto.Caption = Format(CambiaPrecio(nPVenta - ((RstCombo2!nAumento / nCantidad) * RsCombo!nCantidad)), "#,###,##0.00")
                                End If
                        End If
                        'Fin Oscar Ortega------------------------------------------
                        
                        'KDS2
                        If lKDS Then
                            Dim kdsRsCabecera As Recordset
                            Isql = "SELECT * From vPedidoCabecera Where Codigo = '" & Pedido & "' Order By codigo "
                            Set kdsRsCabecera = Lib.OpenRecordset(Isql, Cn)
                            Call KDS_EliminarProductoDeCombo(kdsRsCabecera, sitem, xItem)
                        End If
                        
                                  
                        
                        'insumoCOMBO2013
                        'INSUMOCRITICO23
                    Dim rstItems As New ADODB.Recordset
                    Set rstItems = New ADODB.Recordset
                    Set rstItems = Lib.OpenRecordset("SELECT     dbo.TPRODUCTO.tCodigoInsumo, " & sDetalle & ".nCantidad * " & sComboDetalle & ".nCantidad AS nCantidad FROM         " & sDetalle & " INNER JOIN     " & sComboDetalle & " ON   " & sDetalle & ".tItem = " & sComboDetalle & ".tItem INNER JOIN                       dbo.TPRODUCTO ON " & sComboDetalle & ".tProductoCombo = dbo.TPRODUCTO.tCodigoProducto WHERE      " & sComboDetalle & ".tItem ='" & sitem & "' and tItemCombo='" & xItem & "' and  (dbo.TPRODUCTO.lControlInsumoCritico = 1) AND (ISNULL(dbo.TPRODUCTO.tCodigoInsumo, '') <> '') AND (ISNULL(" & sComboDetalle & ".lImprime, 0) = 1) ", Cn)
                
                    If Not (rstItems.EOF Or rstItems.BOF) Then
                        modificaStockInsumo rstItems.Fields(0), rstItems.Fields(1), "I"
                    End If
                    
                    
                   Dim cMax As String
                   cMax = Calcular("select max(tItem) as Codigo from APEDIDO where tCodigoPedido='" & Pedido & "'", Cn)
                   cMax = Lib.Correlativo(cMax, 3)
                   Isql = "insert into APEDIDO (tCodigoPedido, tItem, tCodigoProducto, tCodigoGrupo, tCodigoSubGrupo, " & _
                          "nPrecioNeto, nPrecioImpuesto1, nPrecioImpuesto2, nPrecioImpuesto3, nPrecioVenta, " & _
                          "nDescuento, nPrecioOficial, nCantidad, nImpuesto1, nImpuesto2, nImpuesto3, nVenta, " & _
                          "tComanda, lImprime, tUsuario, fRegistro, tUsuarioAnulado, fRegistroAnulado, " & _
                          "tObservacion, tObservacionAnulado, tEstadoItem, lImprimeArea, tArea, tMotivoEliminacion, tTurnoAnulado,fDiaContable) " & _
                          "select '" & Pedido & "' as tCodigoPedido, '" & cMax & "' as tItem, cpedido.tProductocombo, cpedido.tCodigoGRupo, cpedido.tCodigoSubGrupo, " & _
                          "cpedido.nPrecioNeto, cpedido.nImpuesto1/cpedido.ncantidad, cpedido.nImpuesto2/cpedido.ncantidad, cpedido.nImpuesto3/cpedido.ncantidad, cpedido.nVenta/cpedido.ncantidad, " & _
                          "0, cpedido.nPrecioNeto, cpedido.nCantidad, cpedido.nImpuesto1, cpedido.nImpuesto2, cpedido.nImpuesto3, cpedido.nVenta, '', cpedido.lImprime, " & _
                          "'" & sUsuario & "' as tUsuario, dpedido.fregistro as fRegistro, " & _
                          "'" & sUsuarioAutoriza & "' as tUsuarioAnulado, getDate() as fRegistroAnulado, " & _
                          "'Anulado de Combo' as tObservacion, 'Anul. de Combo:' + t.tResumido as tObservacionAnulado, 'N', cpedido.lImprimeArea, '', '000', '" & sTurno & "','" & Format(obtieneDiaContable, "yyyyMMdd") & "' " & _
                          "from  dpedido inner join cpedido on dpedido.tCodigoPedido=cpedido.tcodigopedido and  dpedido.tItem = CPEDIDO.tItem inner join tproducto t on t.tcodigoproducto = dpedido.tCodigoProducto " & _
                          "where cpedido.tCodigoPedido = '" & Pedido & "' and cpedido.tItem ='" & sitem & "' and cpedido.tItemCombo='" & xItem & "'"
                    Cn.Execute Isql
    
                  
                        
                   Cn.Execute "DELETE " & sComboPropiedad & " where tItem ='" & sitem & "' and tItemCombo='" & xItem & "'"
                   Cn.Execute "DELETE " & sComboDetalle & " where tItem ='" & sitem & "' and tItemCombo='" & xItem & "'"
                
                   Cn.Execute "delete from CPEDIDO where tCodigoPedido ='" & Pedido & "' and tItem ='" & sitem & "' and tItemCombo='" & xItem & "'"
                   Cn.Execute "delete from TCOMBOPROPIEDAD where tCodigoPedido='" & Pedido & "' and tItem ='" & sitem & "' and tItemCombo='" & xItem & "'"
                   
                   RsComboPropiedad.Requery
                   RsCombo.Requery
                   If RsCombo.RecordCount > 0 Then
                      RsCombo.MoveLast
                   End If
               End If
                
          Case Is = 2 'Aumentar
               If RsCombo.RecordCount = 0 Then
                  Exit Sub
               End If
               
               If RsCombo!lImprime = True Then
                  MsgBox "El Item ya fue Impreso", vbExclamation, sMensaje
                  Exit Sub
               End If
               
               nCCombo = Calcular("select sum(nCantidad) as Codigo " & _
                                  "FROM " & sComboDetalle & " WHERE tItem='" & sitem & "'", Cn)
               If nCCombo + 1 > nCombo * RsDetalle!nCantidad Then
                  MsgBox "La cantidad máxima de items para este producto es de " & nCombo * RsDetalle!nCantidad, vbExclamation, sMensaje
                  Exit Sub
               End If
                'OO----------------------------------------------
                Dim oRsProductoDeCombo As Recordset
                Set oRsProductoDeCombo = Obtener_ProductoDeCombo(RsDetalle!tCodigoProducto, sCombo)
                If oRsProductoDeCombo.RecordCount > 0 Then
                    If IIf(IsNull(oRsProductoDeCombo!lUnico), False, oRsProductoDeCombo!lUnico) Then
                         Dim nCantidadEnElCombo As Integer
                         nCantidadEnElCombo = ObtenerSumaCantidadesEnElCombo(sitem, oRsProductoDeCombo!tEtiqueta)
                         If nCantidadEnElCombo >= RsDetalle!nCantidad Then
                             MsgBox "Solo es permitido " & nCantidad & " elemento(s) de tipo " & oRsProductoDeCombo!tEtiqueta, vbExclamation, sMensaje
                             Exit Sub
                         End If
                    End If
                End If
                '----------------------------------------------------------
               nPos = RsCombo.AbsolutePosition
               Cn.Execute "update " & sComboDetalle & " set nCantidad = " & RsCombo!nCantidad + 1 & " where tItem ='" & sitem & "' and tItemCombo='" & RsCombo!tItemCombo & "'"
               'OO------------------------------------------------------------
               Isql = "Select t.nAumento From [" & sComboDetalle & "] As c Left Join TCombo as t On c.tProducto = t.tCombo And c.tProductoCombo = t.tCodigoProducto Where c.tItem = '" & sitem & "' And c.tItemCombo = '" & RsCombo!tItemCombo & "'"
               Dim RstCombo As Recordset
               Set RstCombo = Lib.OpenRecordset(Isql, Cn)
               If IIf(IsNull(RstCombo!nAumento), 0, RstCombo!nAumento) > 0 Then
                    txtMonto.Caption = Format(CambiaPrecio(nPVenta + RstCombo!nAumento / nCantidad), "#,###,##0.00")
               End If
               '--------------------------------------------------------
               RsCombo.Requery
               RsCombo.AbsolutePosition = nPos
               
          Case Is = 3 'Disminuir
               If RsCombo.RecordCount = 0 Then
                  Exit Sub
               End If
               If RsDetalle!lImprime = True Then
                  MsgBox "El Item ya fue Impreso", vbExclamation, sMensaje
                  Exit Sub
               End If
          
               If grdCombo.Columns(2).Text > 1 Then
                  nPos = RsCombo.AbsolutePosition
                  Cn.Execute "update " & sComboDetalle & " set nCantidad = " & RsCombo!nCantidad - 1 & " where tItem ='" & sitem & "' and tItemCombo='" & RsCombo!tItemCombo & "'"
                  'Oscar Ortega------------------------------------------------------------
                  Isql = "Select t.nAumento From [" & sComboDetalle & "] As c Left Join TCombo as t On c.tProducto = t.tCombo And c.tProductoCombo = t.tCodigoProducto Where c.tCodigoPedido = '" & Pedido & "' And c.tItem = '" & sitem & "' And c.tItemCombo = '" & RsCombo!tItemCombo & "'"
                  Set RstCombo = Lib.OpenRecordset(Isql, Cn)
                  If IIf(IsNull(RstCombo!nAumento), 0, RstCombo!nAumento) > 0 Then
                    txtMonto.Caption = Format(CambiaPrecio(nPVenta - RstCombo!nAumento / nCantidad), "#,###,##0.00")
                  End If
                  'Fin Oscar Ortega--------------------------------------------------------
                  RsCombo.Requery
                  RsCombo.AbsolutePosition = nPos
               End If
               
          Case Is = 4 'Cantidad
               If RsCombo.RecordCount = 0 Then
                  Exit Sub
               End If
               If RsCombo!lImprime = True Then
                  MsgBox "El Item ya fue Impreso", vbExclamation, sMensaje
                  Exit Sub
               End If
               sTipo = ""
               frmNumPad.Show vbModal
               If wEnter And Val(sDescrip) > 0 Then
                  nCantidad = Val(sDescrip)
                  nCCombo = Calcular("select sum(nCantidad) as Codigo " & _
                                     "FROM " & sComboDetalle & " WHERE tItem='" & sitem & "'", Cn)
                  If nCCombo + nCantidad - RsCombo!nCantidad > nCombo * RsDetalle!nCantidad Then
                     MsgBox "La cantidad máxima de items para este producto es de " & nCombo, vbExclamation, sMensaje
                     nCantidad = 1
                     Exit Sub
                  End If
                  
                  'Oscar Ortega----------------------------------------------
                  Set oRsProductoDeCombo = Obtener_ProductoDeCombo(RsDetalle!tCodigoProducto, sCombo)
                  If oRsProductoDeCombo.RecordCount > 0 Then
                     If IIf(IsNull(oRsProductoDeCombo!lUnico), False, oRsProductoDeCombo!lUnico) Then
                         nCantidadEnElCombo = ObtenerSumaCantidadesEnElCombo(sitem, oRsProductoDeCombo!tEtiqueta)
                         If nCantidad > RsDetalle!nCantidad Then
                             MsgBox "Solo es permitido " & RsDetalle!nCantidad & " elemento(s) de tipo " & oRsProductoDeCombo!tEtiqueta, vbExclamation, sMensaje
                             nCantidad = 1
                             Exit Sub
                         End If
                     End If
                  End If
                  '----------------------------------------------------------
                  'OO----------------------------------------------
                  nCantidad = RsDetalle!nCantidad
                  Isql = "Select c.nCantidad, t.nAumento From [" & sComboDetalle & "] As c Left Join TCombo as t On c.tProducto = t.tCombo And c.tProductoCombo = t.tCodigoProducto Where c.tCodigoPedido = '" & sPedido & "' And c.tItem = '" & sitem & "' And c.tItemCombo = '" & xItem & "'"
                  Set RstCombo = Lib.OpenRecordset(Isql, Cn)
                  If RstCombo.RecordCount > 0 Then
                    If IIf(IsNull(RstCombo!nAumento), 0, RstCombo!nAumento) > 0 Then
                        txtMonto.Caption = Format(CambiaPrecio(nPVenta - ((RstCombo!nAumento / nCantidad) * RstCombo!nCantidad)), "#,###,##0.00")
                        txtMonto.Caption = Format(CambiaPrecio(nPVenta + ((RstCombo!nAumento / nCantidad) * Val(sDescrip))), "#,###,##0.00")
                    End If
                    
                  End If
                  '------------------------------------------
                  
                  nPos = RsDetalle.AbsolutePosition
                  Cn.Execute "update " & sComboDetalle & " set nCantidad = " & Val(sDescrip) & " where tItem ='" & sitem & "' and tItemCombo='" & RsCombo!tItemCombo & "'"
                  RsCombo.Requery
                  RsCombo.AbsolutePosition = nPos
               End If
          
          Case Is = 5 'Propiedad Combos
               If RsCombo.RecordCount = 0 Then
                  Exit Sub
               End If
          
               If RsCombo!lImprime = True Then
                  MsgBox "El Item ya fue Impreso", vbExclamation, sMensaje
                  Exit Sub
               End If
               fraProductoCombo.Visible = False
               fraPropiedad.Visible = True
               
               ListarOperadoresConFiltro (sCombo)
   
           Case Is = 6  'Orden +
                If RsCombo.RecordCount = 0 Then
                   Exit Sub
                End If
               If RsCombo!lImprime = True Then
                  MsgBox "El Item ya fue Impreso", vbExclamation, sMensaje
                  Exit Sub
               End If
                nPos = RsCombo.AbsolutePosition
                nOrd = IIf(IsNull(RsCombo!nOrden), 0, RsCombo!nOrden)
                Cn.Execute "update " & sComboDetalle & " set nOrden = " & nOrd + 1 & " where tItem ='" & sitem & "' and tItemCombo='" & RsCombo!tItemCombo & "'"
                RsCombo.Requery
                RsCombo.AbsolutePosition = nPos
                
           Case Is = 7  'Orden -
                If RsCombo.RecordCount = 0 Then
                   Exit Sub
                End If
                If RsCombo!lImprime = True Then
                  MsgBox "El Item ya fue Impreso", vbExclamation, sMensaje
                  Exit Sub
               End If
                nPos = RsCombo.AbsolutePosition
                nOrd = IIf(IsNull(RsCombo!nOrden), 0, RsCombo!nOrden)
                If nOrd > 1 Then
                   Cn.Execute "update " & sComboDetalle & " set nOrden = " & nOrd - 1 & " where tItem ='" & sitem & "' and tItemCombo='" & RsCombo!tItemCombo & "'"
                   RsCombo.Requery
                   RsCombo.AbsolutePosition = nPos
                End If
               
          Case Is = 8  'Linea Corte
               If RsCombo.RecordCount = 0 Then
                  Exit Sub
               End If
          
               nPos = RsCombo.AbsolutePosition
               If IIf(IsNull(RsCombo!lCorte), False, RsCombo!lCorte) Then
                  Cn.Execute "update  " & sComboDetalle & " set lCorte = 0 where tCodigoPedido='" & Pedido & "' and tItem ='" & sitem & "' and tItemCombo='" & RsCombo!tItemCombo & "'"
               Else
                  Cn.Execute "update  " & sComboDetalle & " set lCorte = 1 where tCodigoPedido='" & Pedido & "' and tItem ='" & sitem & "' and tItemCombo='" & RsCombo!tItemCombo & "'"
               End If
               RsCombo.Requery
               RsCombo.AbsolutePosition = nPos
               
   End Select
End Sub

Private Sub cmdCortesia_Click()
    sTemp = ""
    Isql = "select * from vCortesia where lActivo = 1 Order by Descripcion"
    Call ConfGrilla(2, frmBusquedaRapida.grdGrilla, "Código", 2, "Codigo", 1200, 2, 0, "", _
                                                    "Descripcion", 2, "Descripcion", 7000, 0, 0, "")
                         
    frmBusquedaRapida.nPredeterm = 1
    frmBusquedaRapida.Show vbModal
    If wEnter = True Then
       sCortesia = sCodigo
       txtCortesia.Caption = sDescrip
       
       txtDImporte.Caption = "0.00"
       txtRImporte.Caption = "0.00"
       txtImpuesto1.Caption = "0.00"
       txtImpuesto2.Caption = "0.00"
       txtImpuesto3.Caption = "0.00"
       txtPVenta.Caption = "0.00"
       txtVenta.Caption = "0.00"
       txtDPorcentaje.Caption = "0.00"
       txtRPorcentaje.Caption = "0.00"
       
       nPBase = 0
       nRecargo = 0
       nDescuento = 0
       nPVenta = 0
       nImpuesto1 = 0
       nImpuesto2 = 0
       nImpuesto3 = 0
    Else
       sCortesia = ""
       txtCortesia.Caption = ""
    End If
    txtBarra.SetFocus
End Sub

Private Sub cmdDescuento_Click(Index As Integer)
   Select Case Index
          Case Is = 0 ' Dscto. Monto
               If nPBase > 0 Then
                  sTipo = ""
                  frmNumPad.Show vbModal
                  If wEnter Then
                     nDescuento = Val(sDescrip)
                     txtDImporte.Caption = Format(nDescuento, "###,###,###,##0.00")
                     CalculaPrecio
                  End If
               End If
          
          Case Is = 1 ' Dscto. Porcentaje
               If nPBase > 0 Then
                  sTipo = ""
                  frmNumPad.Show vbModal
                  If wEnter Then
                     txtDPorcentaje.Caption = Format(sDescrip, "###,###,###,##0.00")
                     nDescuento = nOficial * Val(sDescrip) / 100
                     txtDImporte.Caption = Format(nDescuento, "###,###,###,##0.00")
                     CalculaPrecio
                  End If
               End If
          
          Case Is = 2 ' Recargo Monto
               If nPBase > 0 Then
                  sTipo = ""
                  frmNumPad.Show vbModal
                  If wEnter Then
                     nRecargo = Val(sDescrip)
                     txtRImporte.Caption = Format(nRecargo, "###,###,###,##0.00")
                     CalculaPrecio
                  End If
               End If
          
          Case Is = 3 ' Recargo Porcentaje
               If nPBase > 0 Then
                  sTipo = ""
                  frmNumPad.Show vbModal
                  If wEnter Then
                     txtRPorcentaje.Caption = Format(sDescrip, "###,###,###,##0.00")
                     nRecargo = nOficial * Val(sDescrip) / 100
                     txtRImporte.Caption = Format(nRecargo, "###,###,###,##0.00")
                     CalculaPrecio
                  End If
               End If
          
    End Select
    txtBarra.SetFocus
End Sub

Private Sub cmdDetalle_Click(Index As Integer)
   txtBarra.SetFocus
   wAgregaCombo = False
   fraCombo.Visible = False
   Select Case Index
          Case Is = 0 'Eliminar
               If RsDetalle.RecordCount = 0 Then
                  Exit Sub
               End If
               
               If MsgBox("Seguro de Eliminar el Producto?", vbQuestion + vbYesNo, sMensaje) = vbYes Then
                  sUsuarioAutoriza = sUsuario
                  If (lPassword And lPrinter = False) Or (lPassword And lPrinter And RsDetalle!lImprime) Then
                        tUsuActua = sUsuario
                     If Supervisor("02") = False Then
                        MsgBox "Clave no permitida", vbExclamation, sMensaje
                        Exit Sub
                     End If
                     sUsuario = tUsuActua
                     sUsuarioAutoriza = sVar1
                  End If
                                                                            
                  If (lElimina And lPrinter = False) Or (lElimina And lPrinter = True And RsDetalle!lImprime) Then
                     fraEliminacion.Visible = True
                     tabProducto.Visible = False
                     ActivaCabecera False
                  Else
                     sCodigo = ""
                     sDescrip = ""
                     SoloEliminaItem
                  End If
               End If
                                                                                                                                                                                
          Case Is = 1 ' Cantidad
               If RsDetalle.RecordCount = 0 Then
                  Exit Sub
               End If
          
               If RsDetalle!lImprime = True Then
                  MsgBox "El Item ya fue Impreso", vbExclamation, sMensaje
                  Exit Sub
               End If
               sTipo = ""
               frmNumPad.Show vbModal
               If wEnter And Val(sDescrip) > 0 Then
                    'Oscar Ortega------------
                    Dim oRsDetalleProducto As Recordset
                    Set oRsDetalleProducto = ObtenerDetalleProducto(sitem)
                    
                    'CESAR ROTULADO
                    If oRsDetalleProducto!tCodigoEtiqueta <> "" Then
                       MsgBox "No es posible aplicar los cambios", vbCritical + vbInformation
                       Exit Sub
                    End If
                    
                    Dim Xse As Integer
                    Dim ix As Integer
                    
                    If (IIf(IsNull(oRsDetalleProducto!lCombinacion), 0, oRsDetalleProducto!lCombinacion)) Then
                       If verificaCantidadDeItemsCombos(sitem, oRsDetalleProducto!nCombinacion, Val(sDescrip)) = False Then
                           MsgBox "No es posible aplicar los cambios. Verifique la cantidad de productos dentro del combo", vbCritical + vbInformation
                           Exit Sub
                        End If
                        
                        If oRsDetalleProducto.Fields("NCANTIDAD") > Val(sDescrip) Then 'DISMINUYE
                           Dim nCantidadMax As Double
                           nCantidadMax = Obtener_CantidadMaximaDeUnicoEtiqueta(sitem, nCantidad)
                           If nCantidad > nCantidadMax Then
                              'nCantidad = Val(sDescrip)
                              'Cn.Execute "update [" & sDetalle & "] set nCantidad = " & Val(sDescrip) & ", nVenta = " & nCantidad * nPVenta & ",nImpuesto1 = nPrecioImpuesto1 * " & nCantidad & ", nImpuesto2 = nPrecioImpuesto2 * " & nCantidad & ", nImpuesto3 = nPrecioImpuesto3 * " & nCantidad & " where tItem ='" & sitem & "'"
                              Xse = nCantidad - Val(sDescrip)
                              For i = 1 To Xse
                                    nPos = RsDetalle.AbsolutePosition
                                    nCantidad = nCantidad - 1
                                    Set oRsDetalleProducto = ObtenerDetalleProducto(sitem)
                                    Dim DDcombo As String
                                    DDcombo = CambiaPrecioCombo((oRsDetalleProducto!nVenta - (oRsDetalleProducto!nPrecioOficial - oRsDetalleProducto!nDescuento)) / nCantidad)
                              Next i
                           Else
                              MsgBox ("No puedes reducir la cantidad de combos con elementos únicos" & Chr(13) & "Disminuya primero la cantidad de productos dentro del combo"), vbExclamation
                              Exit Sub
                           End If
                           
                        Else 'AUMENTA
                           'nCantidad = Val(sDescrip)
                           'Cn.Execute "update [" & sDetalle & "] set nCantidad = " & Val(sDescrip) & ", nVenta = " & nCantidad * nPVenta & ",nImpuesto1 = nPrecioImpuesto1 * " & nCantidad & ", nImpuesto2 = nPrecioImpuesto2 * " & nCantidad & ", nImpuesto3 = nPrecioImpuesto3 * " & nCantidad & " where tItem ='" & sitem & "'"
                           Xse = Val(sDescrip) - nCantidad
                           For i = 1 To Xse
                                nPos = RsDetalle.AbsolutePosition
                                nCantidad = nCantidad + 1
                                Set oRsDetalleProducto = ObtenerDetalleProducto(sitem)
                                Dim AAcombo As String
                                AAcombo = CambiaPrecioCombo(((oRsDetalleProducto!nVenta - oRsDetalleProducto!nDescuento) + oRsDetalleProducto!nPrecioOficial) / nCantidad)
                           Next i
                        End If
                    Else
                        nCantidad = Val(sDescrip)
                        nPos = RsDetalle.AbsolutePosition
                        Cn.Execute "update [" & sDetalle & "] set nCantidad = " & Val(sDescrip) & ", nVenta = " & nCantidad * nPVenta & ",nImpuesto1 = nPrecioImpuesto1 * " & nCantidad & ", nImpuesto2 = nPrecioImpuesto2 * " & nCantidad & ", nImpuesto3 = nPrecioImpuesto3 * " & nCantidad & " where tItem ='" & sitem & "'"
                    End If
                  
                    RsDetalle.Requery
                    RsDetalle.MoveFirst
                    RsDetalle.Find "tItem = '" & sitem & "'"
                    fxCombo "M", nCantidad, sProducto
                    nMonto = Calcular("select sum(nVenta) as Codigo FROM [" & sDetalle & "] where tEstadoItem = 'N'", Cn)
                                    
                    verificatitulo
               End If
               
          Case Is = 2 ' Aumentar
               If RsDetalle.RecordCount = 0 Then
                  Exit Sub
               End If
          
               If RsDetalle!lImprime = True Then
                  MsgBox "El Item ya fue Impreso", vbExclamation, sMensaje
                  Exit Sub
               End If
               nCantidad = nCantidad + 1
               'Oscar Ortega Aumenta Combo---------------
               Set oRsDetalleProducto = ObtenerDetalleProducto(sitem)
               If (oRsDetalleProducto!lCombinacion) Then
                  If verificaCantidadDeItemsCombos(sitem, oRsDetalleProducto!nCombinacion, nCantidad) = False Then
                     MsgBox "No es posible aplicar los cambios. Verifique la cantidad de productos dentro del combo", vbCritical + vbInformation
                     Exit Sub
                  End If
                  'txtMonto.Caption = CambiaPrecio((oRsDetalleProducto!nVenta) / (nCantidad - 1))
                  'Cn.Execute "update [" & sDetalle & "] set nCantidad = " & nCantidad & ", nVenta = " & nCantidad * nPVenta & ",nImpuesto1 = nPrecioImpuesto1* " & nCantidad & ", nImpuesto2 = nPrecioImpuesto2* " & nCantidad & ", nImpuesto3 = nPrecioImpuesto3* " & nCantidad & " where tItem ='" & sitem & "'"
                  Dim Acombo As String
                  Acombo = CambiaPrecioCombo(((oRsDetalleProducto!nPrecioOficial - oRsDetalleProducto!nDescuento) + oRsDetalleProducto!nVenta) / nCantidad)
               Else
                  'CESAR ROTULADO
                  If oRsDetalleProducto!tCodigoEtiqueta <> "" Then
                       MsgBox "No es posible aplicar los cambios", vbCritical + vbInformation
                    Else
                  Cn.Execute "update [" & sDetalle & "] set nCantidad = " & nCantidad & ", nVenta = " & nCantidad * nPVenta & ",nImpuesto1 = nPrecioImpuesto1* " & nCantidad & ", nImpuesto2 = nPrecioImpuesto2* " & nCantidad & ", nImpuesto3 = nPrecioImpuesto3* " & nCantidad & " where tItem ='" & sitem & "'"
                  End If
                  
               End If
               RsDetalle.Requery
               RsDetalle.MoveFirst
               RsDetalle.Find "tItem = '" & sitem & "'"
               fxCombo "M", nCantidad, sProducto
               nMonto = Calcular("select sum(nVenta) as Codigo FROM [" & sDetalle & "] where tEstadoItem = 'N'", Cn)
               verificatitulo
                    
          Case Is = 3 ' Disminuir
               If RsDetalle.RecordCount = 0 Then
                  Exit Sub
               End If
          
               If RsDetalle!lImprime = True Then
                  MsgBox "El Item ya fue Impreso", vbExclamation, sMensaje
                  Exit Sub
               End If
          
               If grdDetalle.Columns(4).Text > 1 Then
               
                  nCantidad = nCantidad - 1
                  'Oscar Ortega Disminuir Combo---------------
                  Set oRsDetalleProducto = ObtenerDetalleProducto(sitem)
                  If (oRsDetalleProducto!lCombinacion) Then
                            If verificaCantidadDeItemsCombos(sitem, oRsDetalleProducto!nCombinacion, nCantidad) = False Then
                                MsgBox "No es posible aplicar los cambios. Verifique la cantidad de productos dentro del combo", vbCritical + vbInformation
                                Exit Sub
                            End If
                            
                            'Dim nCantidadMax As Double
                            nCantidadMax = Obtener_CantidadMaximaDeUnicoEtiqueta(sitem, nCantidad + 1)
                            If nCantidad + 1 > nCantidadMax Then
                               'Disminuir Combo
                               'txtMonto.Caption = CambiaPrecio((oRsDetalleProducto!nVenta) / (nCantidad + 1))
                               'Cn.Execute "update [" & sDetalle & "] set nCantidad = " & nCantidad & ", nVenta = " & nCantidad * nPVenta & ",nImpuesto1 = nPrecioImpuesto1* " & nCantidad & ", nImpuesto2 = nPrecioImpuesto2* " & nCantidad & ", nImpuesto3 = nPrecioImpuesto3* " & nCantidad & " where tItem ='" & sitem & "'"
                               Dim Dcombo As String
                               Dcombo = CambiaPrecioCombo((oRsDetalleProducto!nVenta - (oRsDetalleProducto!nPrecioOficial - oRsDetalleProducto!nDescuento)) / nCantidad)
                            Else
                                MsgBox ("El combo tiene demasiados elementos únicos.")
                                nCantidad = nCantidad + 1
                                Exit Sub
                            End If
                  Else
                        'CESAR ROTULADO
                        If oRsDetalleProducto!tCodigoEtiqueta <> "" Then
                           MsgBox "No es posible aplicar los cambios", vbCritical + vbInformation
                        Else
                           Cn.Execute "update [" & sDetalle & "] set nCantidad = " & nCantidad & ", nVenta = " & nCantidad * nPVenta & ",nImpuesto1 = nPrecioImpuesto1* " & nCantidad & ", nImpuesto2 = nPrecioImpuesto2* " & nCantidad & ", nImpuesto3 = nPrecioImpuesto3* " & nCantidad & " where tItem ='" & sitem & "'"
                        End If
                  
                  End If
                  RsDetalle.Requery
                  RsDetalle.MoveFirst
                  RsDetalle.Find "tItem = '" & sitem & "'"
                  fxCombo "M", nCantidad, sProducto
                  nMonto = Calcular("select sum(nVenta) as Codigo FROM [" & sDetalle & "] where tEstadoItem = 'N'", Cn)
                  verificatitulo
               End If
               
          Case Is = 4 ' Precios
               If RsDetalle.RecordCount = 0 Then
                  Exit Sub
               End If
                tUsuActua = sUsuario
               If Supervisor("03") = False Then
                  MsgBox "Clave no permitida", vbExclamation, sMensaje
                  Exit Sub
               End If
               sUsuario = tUsuActua
               If RsDetalle!lImprime = True Then
                  MsgBox "El Item ya fue Impreso", vbExclamation, sMensaje
                  Exit Sub
               End If
               
               tabProducto.Visible = False
               fraDetalle.Visible = True
               ActivaCabecera False
               
          Case Is = 5 ' Propiedades
          
            'origen de ventas
            Me.fraOrigenVentas.Visible = False
            '--------------------------------
            
               If RsDetalle.RecordCount = 0 Then
                  Exit Sub
               End If
               
               If RsDetalle!lImprime = True Then
                  MsgBox "El Item ya fue Impreso", vbExclamation, sMensaje
                  Exit Sub
               End If
               
               ActivaCabecera False
               tabProducto.Visible = False
               ListarOperadoresConFiltro (sProducto)
               AsignaPropiedad
               fraPropiedad.Visible = True
          
          Case Is = 6 ' Observacion
               
                'origen de ventas
                Me.fraOrigenVentas.Visible = False
                '--------------------------------
               
               
               frmKeyBoard.Caption = "Nombre / Observación"
               frmKeyBoard.txtResultado.Text = sObser
               frmKeyBoard.Show vbModal
               
               If wEnter = True Then
                  sObser = sDescrip
                  txtObservacion.Caption = sObser
               End If
                    
          Case Is = 7 ' Visualizacion de Pedido
               If RsDetalle.RecordCount = 0 Then
                  MsgBox "No existen Datos a Visualizar", vbExclamation, sMensaje
                  Exit Sub
               End If
               sTipo = "CajaRapida"
               frmPedido.Show vbModal
               If wEnter Then
                  RsDetalle.Requery
               End If
          
          Case Is = 8
               If RsDetalle.RecordCount = 0 Then
                  Exit Sub
               End If
          
               If RsDetalle!lImprime = True Then
                  MsgBox "El Item ya fue Impreso", vbExclamation, sMensaje
                  Exit Sub
               End If
               
               nPos = RsDetalle.AbsolutePosition
               If IIf(IsNull(RsDetalle!lCorte), False, RsDetalle!lCorte) Then
                  Cn.Execute "update " & sDetalle & " set lCorte = 0 where tCodigoPedido='" & Pedido & "' and tItem ='" & sitem & "'"
               Else
                  Cn.Execute "update " & sDetalle & " set lCorte = 1 where tCodigoPedido='" & Pedido & "' and tItem ='" & sitem & "'"
               End If
               RsDetalle.Requery
               RsDetalle.AbsolutePosition = nPos
          
          Case Is = 9 ' Mozos
          
                'origen de ventas
                Me.fraOrigenVentas.Visible = False
                '--------------------------------
          
               tabProducto.Visible = False
               fraMozo.Visible = True
          
          Case Is = 10  'cuenta Corriente
          
            If RsDetalle.RecordCount = 0 Then
               Exit Sub
            End If
            
               'VALIDACION CANAL DE VENTA
                Dim rsCanalVentas As Recordset
                Dim lObligaMozo As Boolean
                Dim lObligaMotorizado As Boolean
                Dim lObligaClienteFrecuente As Boolean
                Dim lObligaFechaEntrega As Boolean
                Dim lObligaEntregarA As Boolean
                
                Set rsCanalVentas = Lib.OpenRecordset("select * from vTipoPedido", Cn)
                rsCanalVentas.Filter = "Codigo = '" & sTipoPedido & "'"
                
                lObligaMozo = IIf(IsNull(rsCanalVentas!lObligaMozo), False, rsCanalVentas!lObligaMozo)
                lObligaMotorizado = IIf(IsNull(rsCanalVentas!lObligaMotorizado), False, rsCanalVentas!lObligaMotorizado)
                lObligaClienteFrecuente = IIf(IsNull(rsCanalVentas!lObligaClienteFrecuente), False, rsCanalVentas!lObligaClienteFrecuente)
                lObligaFechaEntrega = IIf(IsNull(rsCanalVentas!lObligaIngresoFechaEntrega), False, rsCanalVentas!lObligaIngresoFechaEntrega)
                lObligaEntregarA = IIf(IsNull(rsCanalVentas!lObligaEntregarA), False, rsCanalVentas!lObligaEntregarA)
   
               'Obligatoriedad de Mozo
               If lObligaMozo Then
                  If lMCPV Then
                      sMozo = ObtenerCodigoMozo(sVar1)
                      If sMozo = "" Then
                         Exit Sub
                      End If
                  Else
                      If sMozo = "" Or sMozo = "0000" Then
                         MsgBox "Asigne al Mesero", vbExclamation, sMensaje
                         Exit Sub
                      End If
                  End If
               End If
               
               'Obligatoriedad de Motorizado
'               If lObligaMotorizado Then
'                  If sMotorizado = "" Or sMotorizado = "0000" Then
'                     MsgBox "Asigne al Motorizado", vbExclamation, sMensaje
'                     Exit Sub
'                  End If
'               End If
               
               'Obligatoriedad de Mesa
'               If lObligaMesa And sMesa = "" And txtObservacion.Caption = "" Then
'                  MsgBox "Asigne una Mesa", vbExclamation, sMensaje
'                  cmdCabecera_Click (13)
'                  Exit Sub
'               End If
               
               'Obligatoriedad de Cliente Frecuente
               If sClienteFrecuente = "" And lObligaClienteFrecuente Then
                  MsgBox "Asigne el Cliente Delivery", vbExclamation, sMensaje
                  cmdOpcion_Click (9)
                  Exit Sub
               End If
               
               'Obligatoriedad de Fecha de Entrega
               If Me.txtFechaEntrega.Caption = "" And lObligaFechaEntrega Then
                  MsgBox "Asigne la Fecha de Entrega", vbExclamation, sMensaje
                  cmdCabecera_Click (6)
                  Exit Sub
               End If
               
               'Entregar A
               If lObligaEntregarA = True And Me.txtEntregar.Caption = "" Then
                  MsgBox "Asigne información en Entregar A", vbExclamation, sMensaje
                  cmdDetalle_Click (14)
                  Exit Sub
               End If
               
            
            If lMCPV Then
                If Not ValidaExistenciaProducto() Then
                    MsgBox "El Pedido ya fue importado", vbExclamation, sMensaje
                    RsDetalle.Requery
                    Exit Sub
                End If
            End If
            variableEmite = False
                   
            If lPrinter And lObligaPrinter Then
               i = Calcular("select count(tCodigoPedido) as codigo from " & sDetalle & " where lImprime=0", Cn)
               If i > 0 Then
                  cmdOpcion_Click (8)
                        'insumocritico
                           If variableEmite = False Then: Exit Sub
                        'insumocritico
                   
               End If
            End If
                        
                   'insumocritico
                        variableEmite = False
                   'insumocritico
                        
            If Calcular("select count(tFacturado) as Codigo from " & sDetalle & " where isnull(tFacturado,'0') <> '0' and len(ltrim(tFacturado)) <> 0", Cn) > 0 Then
               MsgBox "Imposible pasar el pedido a Cuenta Corrientes, pedidos con items Facturados", vbExclamation, sMensaje
               Exit Sub
            End If
            tUsuActua = sUsuario
            sUsuarioAutoriza = sUsuario
            If Supervisor("09") = False Then
               MsgBox "Clave no permitida", vbExclamation, sMensaje
               Exit Sub
            End If
            sUsuarioAutoriza = sVar1
            sUsuario = tUsuActua
            'Chequea si existe platos a facturar
            sTD = "N"
            RsDetalle.MoveFirst
            Do While Not RsDetalle.EOF
               If (Len(Trim(RsDetalle!tFacturado)) = 0 Or IsNull(RsDetalle!tFacturado)) Then
                  sTD = "S"
                  Exit Do
               End If
               RsDetalle.MoveNext
            Loop
    
            If sTD <> "S" Then
               MsgBox "Error: No existen Productos a Facturar", vbCritical, sMensaje
               Exit Sub
            End If
                                                                        
            sTemp = ""
            Isql = "select * from vCompania where lActivo = 1 order by Descripcion"
            Call ConfGrilla(6, frmBusquedaRapida.grdGrilla, "Código", 2, "Codigo", 800, 2, 0, "", _
                                                            "Identidad", 2, "Identidad", 1500, 0, 0, "", _
                                                            "Cliente", 2, "Descripcion", 2800, 0, 0, "", _
                                                            "Linea", 2, "nLinea", 1050, 1, 0, "##,##0.00", _
                                                            "Consumo", 2, "nConsumo", 1050, 1, 0, "##,##0.00", _
                                                            "Fecha Venc", 2, "fFechaVence", 1050, 0, 0, "dd/mm/yyyy")
            frmBusquedaRapida.nPredeterm = 2
            frmBusquedaRapida.Show vbModal
                                                                         
            If Not wEnter Or sCodigo = "" Then
               Exit Sub
            End If

            sCliente = sCodigo
            xSuma = Calcular("select sum(nVenta) as Codigo FROM " & sDetalle & " where tEstadoItem = 'N' and isnull(tFacturado,'') = ''", Cn)
            
            'Validacion de escoger segun estadoFrecuente
            Dim lValidaEstado As Boolean
            lValidaEstado = False
            lValidaEstado = Calcular("select ISNULL(tb.nValor,0) as codigo from TDELIVERY t INNER JOIN TTABLA tb on t.tEstadoFrecuente = tb.TCODIGO where  t.tCodigoDelivery='" & sCliente & "' and tb.TTABLA='ESTADOFRECUENTE'", Cn)
            If lValidaEstado Then
                MsgBox "No es posible seleccionar al cliente, estado no permitido", vbCritical, sMensaje
                Exit Sub
            End If
            
            
            'centralizada
            Dim xLinea As Double
            Dim xConsumo As Double
            
'            If lCentral = False Then
'                    xLinea = Calcular("select nLinea as Codigo FROM TCOMPANIA where tCodigoCliente = '" & sCliente & "'", Cn)
'                    xConsumo = Calcular("select nConsumo as Codigo FROM TCOMPANIA where tCodigoCliente = '" & sCliente & "'", Cn)
                    xLinea = Calcular("select nLinea as Codigo FROM TDELIVERY where TCODIGODELIVERY = '" & sCliente & "'", Cn)
                    xConsumo = Calcular("select nConsumo as Codigo FROM TDELIVERY where TCODIGODELIVERY = '" & sCliente & "'", Cn)
'            Else
'                Dim conServidor As ADODB.Connection
'                Set conServidor = devuelveConexionCentral(sServidorCentral, bdInforestCentral)
'                If conServidor.State Then
''                       xLinea = Calcular("select isnull(nLinea,0) as Codigo FROM TCOMPANIA where tCodigoCliente = '" & sCliente & "'", conServidor)
''                       xConsumo = Calcular("select isnull(nConsumo,0) as Codigo FROM TCOMPANIA where tCodigoCliente = '" & sCliente & "'", conServidor)
'                        xLinea = Calcular("select isnull(nLinea,0) as Codigo FROM TDELIVERY where TCODIGODELIVERY = '" & sCliente & "'", conServidor)
'                        xConsumo = Calcular("select isnull(nConsumo,0) as Codigo FROM TDELIVERY where TCODIGODELIVERY = '" & sCliente & "'", conServidor)
'                 Else
'                        MsgBox "No es posible conectar con el servidor central" & vbCrLf & "No se puede trabajar con la cuenta corriente", vbCritical, sMensaje
'                        Exit Sub
'
'                End If
'
'
'            End If
            
            If xSuma > xLinea - xConsumo Then
               MsgBox "El Cliente no tiene linea suficiente " & Chr(13) & _
                      "Linea : " & Format(xLinea, "###,##0.00") & "  Consumo : " & Format(xConsumo, "###,##0.00") & Chr(13) & _
                      "Saldo : " & Format(xLinea - xConsumo, "###,##0.00"), vbCritical, sMensaje
               Exit Sub
            End If

            If MsgBox("Esta seguro de Enviar el Pedido Nro: " & Pedido & _
               Chr(13) & "a Cuentas Corrientes al Cliente " & sDescrip & "?", vbQuestion + vbOKCancel, sMensaje) = vbCancel Then
               Exit Sub
            End If
                        
            Cn.Execute "Update MPEDIDO set tClienteCtaCte ='" & sCliente & "', tEstadoPedido = '04' where tCodigoPedido='" & Pedido & "'"
'            If lCentral = False Then
            
                Cn.Execute "Update TDELIVERY set nConsumo = " & xConsumo + xSuma & " where TCODIGODELIVERY ='" & sCliente & "'"
'            Else
'                conServidor.Execute "Update TDELIVERY set nConsumo = " & xConsumo + xSuma & " where TCODIGODELIVERY ='" & sCliente & "'"
'            End If
            
            Isql = "select * from vCtaCte " & _
                   "WHERE Codigo='" & Pedido & "'"
            Set RsImpresion = Lib.OpenRecordset(Isql, Cn)
            
            If RsImpresion.RecordCount = 0 Then
               LimpiaRs
               MsgBox "No existen Datos a Imprimir", vbExclamation, sMensaje
            Else
               ImprimeCtaCte RsImpresion
            End If
            LimpiaRs
                        
            Cn.Execute "delete " & sDetalle
            Cn.Execute "delete " & sComboDetalle
            Cn.Execute "delete " & sComboPropiedad
            Cn.Execute "delete " & sProductoPropiedad
            
            RsDetalle.Requery
            RsComboPropiedad.Requery
            RsProductoPropiedad.Requery
            Inicializar
            Screen.MousePointer = vbDefault
                                                                           
         Case Is = 12 '
                Dim wCalcula As Boolean
                Dim sDescripcionDescuento  As String
                Dim ltope As Boolean
                Dim procedeDescuento As Boolean
              
                If RsDetalle.RecordCount > 0 Then
                         tUsuActua = sUsuario
                         If Supervisor("10") = False Then
                            MsgBox "Clave no permitida", vbExclamation, sMensaje
                            Exit Sub
                         End If
                         sUsuario = tUsuActua
                         sUsuarioAutoriza = sVar1
                         tAutorizaDescuento = sUsuarioAutoriza
                         sTemp = ""
                         
                         Isql = "SELECT Codigo, LTRIM(RTRIM(Descripcion)) as Descripcion, case lRatio when 1 then nRatio else 0 END as nRatio, case lRatio when 0 then nRatio else 0 END as nMonto FROM vMotivoDescuento WHERE lActivo = '1' and AplicaAnticipo=0 ORDER BY Descripcion"
                         Call ConfGrilla(4, frmBusquedaRapida.grdGrilla, "Código", 2, "Codigo", 1500, 2, 0, "", _
                                                                         "Descripción", 2, "Descripcion", 4300, 0, 0, "", _
                                                                         "Porcentaje", 2, "nRatio", 1200, 1, 0, "###,##0.00", _
                                                                         "Monto", 2, "nMonto", 1200, 1, 0, "###,##0.00")
                         frmBusquedaRapida.nPredeterm = 1
                         frmBusquedaRapida.Show vbModal
                         wCalcula = True
                
                         If wEnter Then
                            sCodigoDescuento = sCodigo
                            Dim RsDesc As ADODB.Recordset
                            Set RsDesc = Lib.OpenRecordset("select * from vMotivoDescuento where Codigo = '" & sCodigo & "'", Cn)
                            If RsDesc.EOF Then
                               Exit Sub
                            End If
                            RsDesc.MoveFirst
                            xDescuento = RsDesc!nRatio
                            lRatio = RsDesc!lRatio
                            
                            If sCodigo = "000" Then
                                If xDescuento > 99 Then
                                   xDescuento = 0
                                   MsgBox "Porcentaje no válido", vbCritical, sMensaje
                                   Exit Sub
                                 End If
                      
                               frmKeyBoard.Caption = "Descripcion del Descuento"
                               frmKeyBoard.Show vbModal
                               sDescripcionDescuento = sDescrip
                            End If
                         Else
                            Exit Sub
                         End If
                         
                         Dim SumTotalPedido As Double
                         If Pedido = "" Then
                            SumTotalPedido = Calcular("select sum(d.nventa) as codigo from " & sDetalle & " d inner join TPRODUCTO p on d.tCodigoProducto = p.tCodigoProducto where p.lDescuento = 1", Cn)
                         Else
                            SumTotalPedido = Calcular("select sum(d.nventa) as codigo from DPEDIDO d inner join TPRODUCTO p on d.tCodigoProducto = p.tCodigoProducto where d.tCodigoPedido='" & Pedido & "' and p.lDescuento = 1", Cn)
                         End If
                         
                         'CDbl(txtMonto.Caption)
                         If Not RsDesc!lRatio And (RsDesc!nRatio > SumTotalPedido) Then
                            sCodigoDescuento = ""
                            xDescuento = 0
                            MsgBox "Descuento mayor al Pedido", vbCritical, sMensaje
                            Exit Sub
                         End If
                         
                         If RsDesc!lBloqueo Then
                            sTipo = "Prepintado"
                            sCodigo = xDescuento
                            frmNumPad.Show vbModal
                         Else
                            wEnter = False
                         End If
                                        
                         If wEnter Then
                            xDescuento = Val(sDescrip)
                            
                            If (RsDesc!lRatio) Then
                                If xDescuento > 99 Then
                                    MsgBox "Descuento Incorrecto"
                                    Exit Sub
                                End If
                            Else
                                If xDescuento > SumTotalPedido Then
                                    MsgBox "Descuento mayor al Pedido", vbCritical, sMensaje
                                    Exit Sub
                                End If
                            End If
                         End If
                         
                         CalculaDescuento
                         RsDetalle.Requery
                         nMonto = Calcular("select sum(nventa) as codigo from " & sDetalle & "", Cn)
                         VisualizaMonto
                End If
         
         Case Is = 13  ' ofertaaaaaaaaaaaaa dic 2010
               tUsuActua = sUsuario
               If Supervisor("10") = False Then
                   MsgBox "Clave no permitida", vbExclamation, sMensaje
                   Exit Sub
                End If
                sUsuario = tUsuActua
                Dim sCriterio As String
                Dim nOferta As Double
                
                If sTipoPedido = "01" Then
                   sCriterio = " and lLocal=1"
                ElseIf sTipoPedido = "02" Then
                   sCriterio = " and lDelivery=1"
                ElseIf sTipoPedido = "03" Then
                   sCriterio = " and lLlevar=1"
                ElseIf sTipoPedido = "04" Then
                   sCriterio = " and lCanal4=1"
                Else
                   sCriterio = " and lCanal5=1"
                End If
                
                Isql = "SELECT tOferta as Codigo, tNombre as Descripcion, " & _
                       "case when nRatio>0 then 'Descuento del ' + str(nRatio,2) + '%' else " & _
                       "case when nMonto>0 then 'Descuento de ' + str(nMonto,2) + ' " & sMonedaN & "' else 'Producto al Precio de ' + ' " & sMonN & " ' + str(nPrecio,2) end end as Oferta " & _
                       "From dbo.TOFERTA WHERE lAutomatica=0 and tCodigoProducto = '" & sProducto & "' and lActivo=1 " & _
                       " and (substring(tFrecuencia," & Weekday(FechaServidor(), vbMonday) & "+1,1) = '1' or (substring(tFrecuencia,1,1)='1') and MONTH(fFecha) = " & Month(FechaServidor()) & " AND DAY(fFecha)= " & Day(FechaServidor()) & ") and tHoraInicial<='" & Format(Time, "HH:mm") & "' and tHoraFinal>='" & Format(Time, "HH:mm") & "'" & _
                       " and (lPermanente=1 or (lPermanente=0 and fFechaInicial<='" & Format(FechaServidor(), "yyyy/mm/dd") & "' and fFechaFinal>='" & Format(FechaServidor(), "yyyy/mm/dd") & "')) " & sCriterio
                
                Call ConfGrilla(3, frmBusquedaRapida.grdGrilla, "Código", 2, "Codigo", 1000, 2, 0, "", _
                                                                "Oferta", 2, "Descripcion", 3000, 0, 0, "", _
                                                                "Descripcion de la Oferta", 2, "Oferta", 4200, 0, 0, "")
                frmBusquedaRapida.nPredeterm = 1
                frmBusquedaRapida.Show vbModal
                wCalcula = True
                
                
                If wEnter Then
                   Isql = "select * from TOFERTA where tOferta='" & sCodigo & "' and tCodigoProducto='" & sProducto & "'"
                   Set RsOferta = Lib.OpenRecordset(Isql, Cn)
                   If RsOferta.RecordCount > 0 Then
                      RsOferta.MoveFirst
                      If RsOferta!nPrecio > 0 Then
                         nOferta = nOficial - IIf(IsNull(RsOferta!nPrecio), 0, RsOferta!nPrecio)
                      ElseIf RsOferta!nMonto > 0 Then
                         nOferta = RsOferta!nMonto
                      Else
                         nOferta = nOficial * IIf(IsNull(RsOferta!nRatio), 1, RsOferta!nRatio) / 100
                      End If
                   End If
                   
                   nPVenta = nOficial - nOferta
                   nDescuento = nOficial - nPVenta
                           
                    Dim Acumulado As Double
                    Select Case pais ' ok
                        Case "001" 'Bolivia
                                Acumulado = 0
                                Acumulado = IIf(nImpuesto1 > 0, Acumulado + nPorcentaje1, Acumulado)
                                Acumulado = IIf(nImpuesto2 > 0, Acumulado + nPorcentaje2, Acumulado)
                                Acumulado = IIf(nImpuesto3 > 0, Acumulado + nPorcentaje3, Acumulado)
                                Acumulado = (Acumulado / 100)
                                nImpuesto1 = IIf(nImpuesto1 > 0, nPVenta * nPorcentaje1 / 100, 0)
                                nImpuesto2 = IIf(nImpuesto2 > 0, nPVenta * nPorcentaje2 / 100, 0)
                                nImpuesto3 = IIf(nImpuesto3 > 0, nPVenta * nPorcentaje3 / 100, 0)
                                nPBase = nPVenta - nImpuesto1 - nImpuesto2 - nImpuesto3
                        
                        Case Else 'Peru, Ecuador
                                Acumulado = 0
                                Acumulado = IIf(nImpuesto1 > 0, Acumulado + nPorcentaje1, Acumulado)
                                Acumulado = IIf(nImpuesto2 > 0, Acumulado + nPorcentaje2, Acumulado)
                                Acumulado = IIf(nImpuesto3 > 0, Acumulado + nPorcentaje3, Acumulado)
                                Acumulado = 1 + (Acumulado / 100)
                                nImpuesto1 = IIf(nImpuesto1 > 0, nPVenta / Acumulado * nPorcentaje1 / 100, 0)
                                nImpuesto2 = IIf(nImpuesto2 > 0, nPVenta / Acumulado * nPorcentaje2 / 100, 0)
                                nImpuesto3 = IIf(nImpuesto3 > 0, nPVenta / Acumulado * nPorcentaje3 / 100, 0)
                                nPBase = nPVenta - nImpuesto1 - nImpuesto2 - nImpuesto3
                    End Select
                    
                    Isql = "Update " & sDetalle & " Set nPrecioNeto = " & nPBase & ", " & _
                    "nDescuento = " & nDescuento & ", " & _
                    "nRecargo = " & nRecargo & ", " & _
                    "nPrecioOficial = " & nOficial & ", " & _
                    "nprecioImpuesto1 = " & nImpuesto1 & ", " & _
                    "nprecioImpuesto2 = " & nImpuesto2 & ", " & _
                    "nprecioImpuesto3 = " & nImpuesto3 & ", " & _
                    "nPrecioVenta = " & nPVenta & ", " & _
                    "nventa = " & nPVenta * nCantidad & ", " & _
                    "nCantidad = " & nCantidad & ", " & _
                    "nImpuesto1 = " & nImpuesto1 * nCantidad & ", " & _
                    "nImpuesto2 = " & nImpuesto2 * nCantidad & ", " & _
                    "nImpuesto3 = " & nImpuesto3 * nCantidad & ", tOferta='" & sCodigo & "', tAutorizaOferta='" & sVar1 & "' " & _
                    "where tItem = '" & sitem & "'"
                    Cn.Execute Isql
                    
                    nPos = RsDetalle.AbsolutePosition
                    RsDetalle.Requery
                    RsDetalle.AbsolutePosition = nPos
                   
                    nMonto = Calcular("select sum(nVenta) as Codigo FROM [" & sDetalle & "] where tEstadoItem = 'N'", Cn)
                   
                End If
         '====================
         Case Is = 14 'entregar a
                frmKeyBoard.Caption = "Entregar A"
                frmKeyBoard.txtResultado.Text = txtEntregar.Caption
                frmKeyBoard.Show vbModal
                If wEnter = True Then
                    txtEntregar.Caption = sDescrip
                End If
                              
   End Select
   
End Sub

Private Sub cmdDirecto_Click(Index As Integer)
    sTemp = ""
    Isql = "select * from vProducto where lActivo = 1 and " & IIf(sTipoPedido = "01", "lLocal=1", IIf(sTipoPedido = "02", "lDelivery=1", "lLlevar=1")) & " Order by Descripcion"
    Call ConfGrilla(5, frmBusquedaRapida.grdGrilla, "Grupo", 2, "Grupo", 1600, 0, 0, "", _
                                                    "Producto", 2, "Descripcion", 3600, 0, 0, "", _
                                                    "Precio", 2, "nPrecioVenta", 1000, 1, 0, "###,##0.00", _
                                                    "Bot", 2, "nBoton", 500, 1, 0, "", _
                                                    "SubGrupo", 2, "SubGrupo", 1500, 0, 0, "")
    frmBusquedaRapida.nPredeterm = 1
    frmBusquedaRapida.Show vbModal
    
    If wEnter Then
       sProducto = sCodigo
       
        'INSUMOCRITICO23
        If validadIngresoProducto(sProducto) = False Then
            Exit Sub
        End If
      'INSUMOCRITICO23
    
       Dim xxx As String
       xxx = RsProducto.Filter
       RsProducto.Filter = adFilterNone
       RsProducto.MoveFirst
       RsProducto.Find ("Codigo='" & sProducto & "'")
    
       If Not RsProducto.EOF() Then
          If wAgregaCombo Then
             nCCombo = Calcular("select sum(nCantidad) as Codigo " & _
                               "FROM " & sComboDetalle & " WHERE tItem='" & sitem & "'", Cn)
        
             If nCCombo < nCombo * RsDetalle!nCantidad Then
                InsertaCombo sProducto
             Else
                MsgBox "La cantidad máxima de items para este producto es de " & nCombo * RsDetalle!nCantidad, vbExclamation, sMensaje
             End If
          Else
             If lBal And RsProducto!lBalanza Then
                Dim nResultado As Double
                nResultado = Pesar(nBalanzaPuerto)
                nResultado = Format(nResultado, "#,##0.00")
                If nResultado > 0 Then
                   InsertaProducto nResultado
                End If
             Else
             nCantidad = 1
                InsertaProducto 1
             End If
             
             If IIf(IsNull(RsProducto!lPropiedad), False, RsProducto!lPropiedad) Then
                lPropiedad = True
             End If
          End If
       End If
       RsProducto.Filter = IIf(xxx = "0", "", xxx)
    End If
    txtBarra.SetFocus
End Sub

Private Sub cmdEtiqueta_Click(Index As Integer)
    sPrefijo = Index
    RsCajaRapida.Filter = "Prefijo='" & sPrefijo & "'"
    RsCajaRapida.MoveFirst
    cmdEtiqueta(1).FontBold = IIf(Index = 1, True, False)
    cmdEtiqueta(2).FontBold = IIf(Index = 2, True, False)
    cmdEtiqueta(3).FontBold = IIf(Index = 3, True, False)
    
    For i = 1 To 9
        cmdAgrupacion(i).backColor = IIf(RsCajaRapida!nValor = 0, -2147483633, RsCajaRapida!nValor)
        If LTrim(RsCajaRapida!tDetallado) = "" Then
           'cmdAgrupacion(i).Caption = "(no utilizado)"
           cmdAgrupacion(i).Visible = False
        Else
           cmdAgrupacion(i).Caption = RsCajaRapida!tDetallado
           cmdAgrupacion(i).Visible = True
        End If
        RsCajaRapida.MoveNext
    Next i
    cmdAgrupacion_Click (1)
    If txtBarra.Visible = True Then
        txtBarra.SetFocus
    End If
End Sub

Private Sub cmdImpuesto_Click(Index As Integer)
    Select Case Index
        Case Is = 0
             nImpuesto1 = IIf(nImpuesto1 = 0, nPBase * nPorcentaje1 / 100, 0)
             txtImpuesto1.Caption = Format(nImpuesto1, "###,###,###,##0.00")
        Case Is = 1
             nImpuesto2 = IIf(nImpuesto2 = 0, nPBase * nPorcentaje2 / 100, 0)
             txtImpuesto2.Caption = Format(nImpuesto2, "###,###,###,##0.00")
        Case Is = 2
             nImpuesto3 = IIf(nImpuesto3 = 0, nPBase * nPorcentaje3 / 100, 0)
             txtImpuesto3.Caption = Format(nImpuesto3, "###,###,###,##0.00")
    End Select
    nPVenta = nPBase + nImpuesto1 + nImpuesto2 + nImpuesto3
    txtPVenta.Caption = Format(nPVenta, "###,###,##0.00")
    txtVenta.Caption = Format((nPVenta * nCantidad), "###,###,###,##0.00")
    txtBarra.SetFocus
End Sub

Private Sub cmdMotorizado_Click(Index As Integer)
   'origen de ventas
    'HabilitaTimerColor (False)
    
   RsMotorizado.MoveFirst
   RsMotorizado.Find "nboton = " & Trim(str(Index))
   frmVenta.txtMotorizado.Caption = RsMotorizado!Descripcion
   sMotorizado = RsMotorizado!codigo
   Me.fraMorotizado.Visible = False
   'HabilitaTimerColor (True)
End Sub
Private Sub cmdMozo_Click(Index As Integer)
   RsMozo.MoveFirst
   RsMozo.Find "nboton = " & Trim(str(Index))
   txtTitulo.Caption = " Caja Rápida : " & sCaja & " Mesero: " & RsMozo!Descripcion
   sMozo = RsMozo!codigo
   tabProducto.Visible = True
   fraMozo.Visible = False
   txtBarra.SetFocus
End Sub

Private Sub cmdNavegar_Click(Index As Integer)
    Select Case Index
           Case Is = 1 'Primero
                MoverPuntero Primero, grdDetalle
           Case Is = 2 'PgUp
                MoverPuntero pgup, grdDetalle
           Case Is = 3 'Previo
                MoverPuntero previo, grdDetalle
           Case Is = 4 'Siguiente
                MoverPuntero siguiente, grdDetalle
           Case Is = 5 'PgDn
                MoverPuntero pgdn, grdDetalle
           Case Is = 6 'Ultimo
                MoverPuntero Ultimo, grdDetalle
           Case Is = 12 'Primero
                MoverPuntero Primero, grdCombo
           Case Is = 13 'Previo
                MoverPuntero previo, grdCombo
           Case Is = 14 'Siguiente
                MoverPuntero siguiente, grdCombo
           Case Is = 15 'Ultimo
                MoverPuntero Ultimo, grdCombo
           Case Is = 16 'PgDn
                MoverPuntero pgdn, grdCombo
           Case Is = 17 'PgUp
                MoverPuntero pgup, grdCombo
    End Select
    txtBarra.SetFocus
End Sub

Private Sub cmdNotasCredito_Click(Index As Integer)
  'anulacion por nota de credito
    Select Case Index
           Case Is = 50  'Nuevo
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
   txtBarra.SetFocus
   Select Case Index
   
          Case Is = 0  'Anulacion Documento
               i = Calcular("select count(tCodigoPedido) as codigo from " & sDetalle & " where lImprime=1", Cn)
               If i > 0 Then
                  MsgBox "El pedido actual esta activo", vbExclamation, sMensaje
                  Exit Sub
               End If
               
               If Pedido <> "" Then
                  MsgBox "El pedido actual esta activo, cancele y vuelva a entrar", vbExclamation, sMensaje
                  Exit Sub
               End If
               tUsuActua = sUsuario
               If Supervisor("05") = False Then
                  MsgBox "Clave no permitida", vbExclamation, sMensaje
                  Exit Sub
               End If
               sUsuario = tUsuActua
               Isql = "SELECT MAX(dbo.DDOCUMENTO.tCodigoPedido) AS Descripcion, dbo.MDOCUMENTO.tDocumento AS Codigo, MAX(dbo.MDOCUMENTO.fRegistro) AS fFecha, dbo.MDOCUMENTO.tUsuario, dbo.TCLIENTE.tEmpresa AS Cliente, dbo.MDOCUMENTO.nVenta, MAX(dbo.MPEDIDO.tObservacion) As tObservacion " & _
                      "FROM dbo.MPEDIDO RIGHT OUTER JOIN dbo.DDOCUMENTO ON dbo.MPEDIDO.tCodigoPedido = dbo.DDOCUMENTO.tCodigoPedido RIGHT OUTER JOIN dbo.MDOCUMENTO LEFT OUTER JOIN dbo.TCLIENTE ON dbo.MDOCUMENTO.tCodigoCliente = dbo.TCLIENTE.tCodigoCliente ON dbo.DDOCUMENTO.tDocumento = dbo.MDOCUMENTO.tDocumento " & _
                      "where dbo.MDOCUMENTO.tTurno='" & sTurno & "' and tEstadoDocumento = '02' " & _
                      "GROUP BY dbo.MDOCUMENTO.tDocumento, dbo.TCLIENTE.tEmpresa, dbo.MDOCUMENTO.nVenta, dbo.MDOCUMENTO.tUsuario " & _
                      "ORDER BY dbo.MDOCUMENTO.tDocumento"
                      
               Call ConfGrilla(6, frmBusquedaRapida.grdGrilla, "Documento", 2, "Codigo", 1500, 0, 0, "", _
                                                               "Fec.Emis", 2, "fFecha", 1000, 0, 0, "", _
                                                               "Monto", 2, "nVenta", 1000, 1, 0, "###,###,##0.00", _
                                                               "Observacion", 2, "tObservacion", 2000, 0, 0, "", _
                                                               "Cliente", 2, "Cliente", 1700, 0, 0, "", _
                                                               "Usu.Emis", 2, "tUsuario", 1000, 0, 0, "")
               frmBusquedaRapida.nPredeterm = 0
               frmBusquedaRapida.Show vbModal
               
               If wEnter Then
                  Dim lContinua As Boolean
                  lContinua = True
                  'Pin Pad
                  Dim RsPinPad As Recordset
                  Set RsPinPad = Lib.OpenRecordset("select nMonto, tReferencia from DPAGOTARJETA where tDocumento='" & sCodigo & "'", Cn)
                  If RsPinPad.RecordCount > 0 Then
                     Dim sMonto As String
                     lContinua = False
                     sMonto = Format(str(RsPinPad!nMonto), "0000000000.00")
                     sMonto = Mid(sMonto, 1, 10) & Mid(sMonto, 12, 2)
                     sOperacion = OP_FINANCIERA & "A" & sMonto & Chr$(FS) & _
                                                  "B" & "000000000000" & Chr$(FS) & _
                                                  "C" & "0" & Chr$(FS) & _
                                                  "D" & sEmpresa & Chr$(FS) & _
                                                  "E" & sCaja
                     nRet = fiStartOperation(sOperacion, 2, sMensaje)
                                                                        
                     If nRet = RET_OK Or nRet = RET_RUNNING Then
                        If Not Imprimir(sPreCuenta) Then
                           Exit Sub
                        End If
                        Printer.FontName = sFont
                        Printer.FontBold = False
                        sClave = ""
                        nContador = 0
                        lEmisor = True
                        lLoop = True
                        Do
                          sRetorno = ""
                          nRet = fiGetStatus(sRetorno, 512)
                          lEmisor = ImprimeCabecera(sRetorno, lEmisor)
                          sClave = MensajePinPad(sRetorno)
                          If Mid(sClave, 1, 3) = "A00" Or Mid(sClave, 1, 3) = "A11" Then
                             sRefer = BuscaRetornoPinPad(sClave, "E")
                             Cn.Execute "update DPAGOTARJETA set tEstadoDocumento='04' where tDocumento='" & sCodigo & "' and tReferencia='" & sRefer & "'"
                             lContinua = True
                          Else
                             xError = BuscaRetornoPinPad(sClave, "B")
                             If Len(xError) > 0 Then
                                Mensaje xError, "VisaNet", 1000
                             End If
                          End If
                          
                          Mensaje "PinPad Listo. Esperando...", "PinPad", 500
                          nContador = nContador + 1
                          If nContador >= nTimeOut Then
                             If MsgBox("Tiempo de espera agotado, deseas mas tiempo?", vbExclamation + vbOKCancel, "VisaNet") = vbOK Then
                                lLoop = True
                                nContador = nTimeOut / 2
                             Else
                                 lLoop = False
                             End If
                          End If
                           
                           If nRet <> "0" Then
                              nContador = 0
                           End If
                        Loop While (Mid$(sRetorno, 5, 2) <> "C1") And lLoop
                     Else
                        MsgBox "Error de conectividad", vbCritical, sMensaje
                        Exit Sub
                     End If
                  End If
                                    
                  If Not lContinua Then
                     Exit Sub
                  End If
                  
                  Pedido = sDescrip
                  sPedido = Pedido
                  Cn.Execute "delete from " & sDetalle
                  Cn.Execute "UPDATE DPEDIDO SET TDOCUMENTO='' , TFACTURADO='' WHERE TCODIGOPEDIDO='" & Pedido & "'"
                  Cn.Execute "Insert into " & sDetalle & " (tCodigoPedido, tItem, tTipoPedido, tCodigoProducto, tCodigoGrupo, tCodigoSubGrupo, tMoneda, " & _
                  "nPrecioNeto, nPrecioImpuesto1, nPrecioImpuesto2, nPrecioImpuesto3, nPrecioVenta, nRecargo, nDescuento, nPrecioOficial, " & _
                  "nCantidad, nImpuesto1, nImpuesto2, nImpuesto3, nVenta, tObservacion, tCortesia, lImprime, tEstadoItem, tArea, lImprimeArea, nInsumo, nGasto, nManoObra, nOrden, lCombinacion, nCombinacion, lCorte,tsubalmacen) " & _
                  "select tCodigoPedido, tItem, tTipoPedido, tCodigoProducto, tCodigoGrupo, tCodigoSubGrupo, tMoneda, " & _
                  "nPrecioNeto, nPrecioImpuesto1, nPrecioImpuesto2, nPrecioImpuesto3, nPrecioVenta, nRecargo, nDescuento, nPrecioOficial, " & _
                  "nCantidad, nImpuesto1, nImpuesto2, nImpuesto3, nVenta, tObservacion, tCortesia, lImprime, tEstadoItem, tArea, lImprimeArea, nInsumo, nGasto, nManoObra, nOrden, lCombinacion, nCombinacion, lCorte,tsubalmacen " & _
                  "From DPEDIDO where tEstadoItem='N' and tCodigoPedido='" & Pedido & "'"
                  
                  Cn.Execute "delete from " & sComboDetalle
                  Cn.Execute "insert into " & sComboDetalle & "(tItem, tItemCombo, tProducto, tProductoCombo, nCantidad, tCodigoGrupo, tCodigoSubGrupo, nPrecioNeto, nImpuesto1, nImpuesto2, nImpuesto3, nVenta, nInsumo, nGasto, nManoObra, lImprimeArea, lImprime, nOrden, tObservacion, lCorte) " & _
                  "select tItem, tItemCombo, tProducto, tProductoCombo, nCantidad, tCodigoGrupo, tCodigoSubGrupo, nPrecioNeto, nImpuesto1, nImpuesto2, nImpuesto3, nVenta, nInsumo, nGasto, nManoObra, lImprimeArea, lImprime, nOrden, tObservacion, lCorte " & _
                  "From CPEDIDO where tCodigoPedido='" & Pedido & "'"

                  Cn.Execute "delete from " & sProductoPropiedad
                  Cn.Execute "insert into " & sProductoPropiedad & "(tItem, tCodigoPropiedad, tProducto, tEnlace, nInsumo, nGasto, nManoObra) " & _
                  "select tItem, tCodigoPropiedad, tProducto, tEnlace, nInsumo, nGasto, nManoObra " & _
                  "From TPRODUCTOPROPIEDAD where tCodigoPedido='" & Pedido & "'"
                  
                  Cn.Execute "delete from " & sComboPropiedad
                  Cn.Execute "insert into " & sComboPropiedad & "(tItem, tItemCombo, tCodigoPropiedad, tProducto, tEnlace, nInsumo, nGasto, nManoObra) " & _
                  "select tItem, tItemCombo, tCodigoPropiedad, tProducto, tEnlace, nInsumo, nGasto, nManoObra " & _
                  "From TCOMBOPROPIEDAD where tCodigoPedido='" & Pedido & "'"
                             
                  Cn.Execute "Update MDOCUMENTO set tObservacion ='ANULACION RAPIDA' where tDocumento ='" & sDocumento & "'"
                  If Not lFECarbajal Then
                    Cn.Execute "delete from DPAGODOCUMENTO where tDocumento='" & sCodigo & "'"
                    Cn.Execute "update MDOCUMENTO set tEstadoDocumento='04', tUsuarioAnulado='" & sVar1 & "', fRegistroAnulado=getdate(), tObservacion='ANULACION RAPIDA' ,lreplica=1  where tDocumento='" & sCodigo & "'"
                    Cn.Execute "update MPEDIDO set nReimpresion = isnull(nReimpresion,0) + 1, tEstadoPedido='01'  where tCodigoPedido='" & Pedido & "'"
                    RsDetalle.Requery
                  End If
                  'FACTURACION_E_PERU
                  If pais = "000" Then
                    If lFacturacionE Then
                    
                           If lFEOfisis Then 'OFISIS
                                 Dim lDocElecOfisis As Boolean
                                 lDocElecOfisis = Calcular("select isnull(tdi.lDocumentoElectronicoOfisis,0) as codigo from TTIPODOCUMENTOIMPRESORA tdi inner join MDOCUMENTO m on tdi.tTipoEmision = m.tTipoDocumento and tdi.tCaja = m.tCaja  where m.tDocumento= '" & sCodigo & "'", Cn)
                                 
                                 If lDocElecOfisis Then 'DOC ELECTRONICO OFISIS
                                        Dim xCDROfisis As String
                                        Dim RsDocumentoOfisis As Recordset
                                        Dim xContOfisis As Integer
                                        
                                        fDocumento = Mid(sCodigo, 1, 1) + Mid(sCodigo, 4, 3) + "-" + CStr(CLng(Mid(sCodigo, 8, 8)))
                                        Isql = "Select * From dbo.TCFACT_ELEC where NU_DOCU='" & fDocumento & "'"
                                        Set RsDocumentoOfisis = Lib.OpenRecordset(Isql, CnFE)
                                        
                                        If RsDocumentoOfisis.RecordCount > 0 Then
                                            CnFE.Execute "Update TCFACT_ELEC set CO_ESTA_DOCU = 'ANU' Where NU_DOCU = '" & fDocumento & "' and TI_DOCU <> 'D'"
                                        End If
                                 End If
                           ElseIf lFESpring Then 'SPRING
                           
                           ElseIf lFEpape Then 'PAPERLESS
                           
                           ElseIf lFECarbajal Then 'CARBAJAL
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
                                Cn.Execute "delete from DPAGODOCUMENTO where tDocumento='" & sCodigo & "'"
                                Cn.Execute "update MDOCUMENTO set tEstadoDocumento='04', tUsuarioAnulado='" & sVar1 & "', fRegistroAnulado=getdate(), tObservacion='ANULACION RAPIDA' ,lreplica=1  where tDocumento='" & sCodigo & "'"
                                Cn.Execute "update MPEDIDO set nReimpresion = isnull(nReimpresion,0) + 1, tEstadoPedido='01'  where tCodigoPedido='" & Pedido & "'"
                                RsDetalle.Requery
                           Else ' INFOFACT
                           
                                 Dim lDocElecInfofact As Boolean
                                 lDocElecInfofact = Calcular("select isnull(tdi.lFacturacionElectronica,0) as codigo from TTIPODOCUMENTOIMPRESORA tdi inner join MDOCUMENTO m on tdi.tTipoEmision = m.tTipoDocumento and tdi.tCaja = m.tCaja  where m.tDocumento= '" & sCodigo & "'", Cn)
                                 
                                 If lDocElecInfofact Then 'DOC ELECTRONICO INFOFACT
                                        Dim xCDR As String
                                        Dim RsDocumentoVenta As Recordset
                                        Dim xCont As Integer
                                        
                                        fDocumento = Mid(sCodigo, 1, 1) + Mid(sCodigo, 4, 3) + Mid(sCodigo, 8, 8)
                                        Isql = "Select * From dbo.DOCUMENTOVENTA where nro_efact='" & fDocumento & "'"
                                        Set RsDocumentoVenta = Lib.OpenRecordset(Isql, CnFE)
                                        
                                        If RsDocumentoVenta.RecordCount > 0 Then
                                            
                                                 xCDR = IIf(IsNull(RsDocumentoVenta!cdr), "", RsDocumentoVenta!cdr)
                                            
                                                 Dim oComandoBaja As clsComando
                                                 Set oComandoBaja = New clsComando
                                                 
                                                 If Mid(sCodigo, 1, 1) = "F" Then
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
                                                                    oComandoBaja.CreateParameter "@CodigoDocumento", adVarChar, adParamInput, 20, sCodigo
                                                    
                                                                    If Not oComandoBaja.GetParamOK Then
                                                                         Set oComandoBaja = Nothing
                                                                         Exit Sub
                                                                    End If
                                                                    If Not oComandoBaja.ExecSP Then
                                                                         Set oComandoBaja = Nothing
                                                                         Exit Sub
                                                                    End If
                                                                    
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
                                                             oComandoBaja.CreateParameter "@CodigoDocumento", adVarChar, adParamInput, 20, sCodigo
                                             
                                                             If Not oComandoBaja.GetParamOK Then
                                                                  Set oComandoBaja = Nothing
                                                                  Exit Sub
                                                             End If
                                                             If Not oComandoBaja.ExecSP Then
                                                                  Set oComandoBaja = Nothing
                                                                  Exit Sub
                                                             End If
                                                             
                                                        'Else
                                                        '     MsgBox "Documento no puede ser Anulado, se supero el limite de dias desde su emisión", vbExclamation, sMensaje
                                                        '     Exit Sub
                                                        'End If
                                                        '----------------------
                                                 End If
                                        End If
                                  End If
                           
                           End If
                    End If
                  End If
                  
                  If RsDetalle.RecordCount = 0 Then
                     nMonto = 0
                  Else
                     nMonto = Calcular("select sum(nVenta) as Codigo FROM [" & sDetalle & "] where tEstadoItem = 'N'", Cn)
                  End If
                  VisualizaMonto
                  
                  Dim RsTemp As Recordset
                  Set RsTemp = Lib.OpenRecordset("select tTipoPedido, tMozo, tObservacion, vMozo.Descripcion as Mozo FROM dbo.MPEDIDO LEFT OUTER JOIN dbo.vMozo ON dbo.MPEDIDO.tMozo = dbo.vMozo.Codigo Where tcodigopedido='" & Pedido & "'", Cn)
                                    
                  sTipoPedido = IIf(IsNull(RsTemp!tTipoPedido), "01", RsTemp!tTipoPedido)
                  If sTipoPedido = "01" Then
                     'cmdCabecera_Click (2)
                  Else
                     'cmdCabecera_Click (1)
                  End If
                  sMozo = IIf(IsNull(RsTemp!tMozo), "0000", RsTemp!tMozo)
                  txtTitulo.Caption = " Caja Rápida : " & sCaja & " Mesero: " & IIf(IsNull(RsTemp!Mozo), "", RsTemp!Mozo)
                  sObser = IIf(IsNull(RsTemp!tObservacion), "", RsTemp!tObservacion)
                  txtObservacion.Caption = sObser
                  Set RsTemp = Nothing
                  
                  If lInfhotel Then
                     sComandaInfhotel = Calcular("select tComanda as Codigo From MPEDIDO where tCodigoPedido='" & Pedido & "'", Cn)
                     CnInfhotel.Execute "update MCOMANDA set TESTADO='04' where tComanda ='" & sComandaInfhotel & "'  and tPuntoVenta='" & sPuntoVenta & "'"
                     CnInfhotel.Execute "delete from DCOMANDA where tcomanda='" & sComandaInfhotel & "' and tcodigoitem='100000' and tPuntoVenta='" & sPuntoVenta & "'"
                     CnInfhotel.Execute "delete from WMCOMANDA where tComanda ='" & sComandaInfhotel & "' and tPuntoVenta='" & sPuntoVenta & "'"
                     CnInfhotel.Execute "delete from WDCOMANDA where tComanda ='" & sComandaInfhotel & "' and tPuntoVenta='" & sPuntoVenta & "'"
                  End If
               End If
          
          Case Is = 1  'Cancelar
               If RsDetalle.RecordCount = 0 Then
                  Exit Sub
               End If
                  
               If MsgBox("Seguro de Cancelar el Pedido?", vbQuestion + vbYesNo, sMensaje) = vbYes Then
               
                  If Pedido <> "" Then
                     sUsuarioAutoriza = sUsuario
                     tUsuActua = sUsuario
                      If lPasswordC Then
                         If Supervisor("01") = False Then
                            MsgBox "Clave no permitida", vbExclamation, sMensaje
                            Exit Sub
                         End If
                         sUsuarioAutoriza = sVar1
                      End If
                      sUsuario = tUsuActua
                      Sw = True
                      If lEliminaC Then
                         fraEliminacion.Visible = True
                         tabProducto.Visible = False
                         ActivaCabecera False
                      Else
                         sCodigo = ""
                         sDescrip = ""
                         EliminaCabecera
                      End If
                       Inicializar
                  Else
                      Cn.Execute "delete from " & sDetalle
                      Cn.Execute "delete from " & sComboDetalle
                      Cn.Execute "delete from " & sProductoPropiedad
                      Cn.Execute "delete from " & sComboPropiedad
                      Cn.Execute "Update MPEDIDO set tEstadoPedido ='03', tMotivoAnulacion='" & sCodigo & "', tUsuarioAnulado='" & sUsuarioAutoriza & "', fRegAnulado= getdate(), tTurnoAnulado='" & sTurno & "', tObservacionAnulado='" & sDescrip & "'   where tCodigoPedido ='" & Pedido & "'"
                      Cn.Execute "Update DPEDIDO Set tEstadoItem = 'A' where tCodigoPedido = '" & Pedido & "'"
                      Cn.Execute "delete TPRODUCTOPROPIEDAD where tCodigoPedido='" & Pedido & "'"
                      Cn.Execute "delete CPEDIDO where tCodigoPedido='" & Pedido & "'"
                      Cn.Execute "delete TCOMBOPROPIEDAD where tCodigoPedido='" & Pedido & "'"
                                         
                      RsDetalle.Requery
                      RsCombo.Requery
                      RsPropiedad.Requery
                      RsProductoPropiedad.Requery
                      RsComboPropiedad.Requery
                    
                      nMonto = 0
                      Pedido = ""
                      sProducto = ""
                      wCombo = False
                      nCombo = 0
                      sitem = ""
                      sCodigoDescuento = ""
                      tAutorizaDescuento = ""
                      ltope = False
                      nTope = 0
                      xDescuento = 0
                      Inicializar
                      VisualizaMonto
                  End If
               End If
                              
          Case Is = 2  'Aceptar
               wEnter = False
               If Pedido = "" Then
                  If nPuerto > 0 Then
                     Visor String(Int((19 - Len(tMensaje1)) / 2), " ") & tMensaje1, String(Int((19 - Len(tMensaje2)) / 2), " ") & tMensaje2, nPuerto, "N"
                  End If
                  Unload Me
               Else
                  If sModulo = "INFOREST" Or sModulo = "ADICION" Then
                     Cn.Execute "delete from DPEDIDO where tCodigoPedido='" & Pedido & "'"
                     'Inserta el Detalle
                     Cn.Execute "Insert into DPEDIDO (tCodigoPedido, tItem, tTipoPedido, tCodigoProducto, tCodigoGrupo, tCodigoSubGrupo, tMoneda, " & _
                                "nPrecioNeto, nPrecioImpuesto1, nPrecioImpuesto2, nPrecioImpuesto3, nPrecioVenta, nRecargo, nDescuento, nPrecioOficial, " & _
                                "nCantidad, nImpuesto1, nImpuesto2, nImpuesto3, nVenta, tObservacion, tCortesia, lImprime, tEstadoItem, tArea, lImprimeArea,fregistro, nInsumo, nGasto, nManoObra, nOrden, lCombinacion, nCombinacion, lCorte,toferta,tautorizaoferta,tSubAlmacen,tCodigoEtiqueta,tunidadnegocio,nenvio,fenvio,fdiacontable) " & _
                                "select tCodigoPedido, tItem, tTipoPedido, tCodigoProducto, tCodigoGrupo, tCodigoSubGrupo,  tMoneda, " & _
                                "nPrecioNeto, nPrecioImpuesto1, nPrecioImpuesto2, nPrecioImpuesto3, nPrecioVenta, nRecargo, nDescuento, nPrecioOficial, " & _
                                "nCantidad, nImpuesto1, nImpuesto2, nImpuesto3, nVenta, tObservacion, tCortesia, lImprime, tEstadoItem, tArea, lImprimeArea,getdate(), nInsumo, nGasto, nManoObra, nOrden, lCombinacion, nCombinacion, lCorte,toferta,tautorizaoferta,tSubAlmacen,tCodigoEtiqueta,'" & sUnidadNegocio & "' ,nenvio,fenvio,'" & Format(obtieneDiaContable, "yyyyMMdd") & "' " & _
                                "From [" & sDetalle & "] where tEstadoItem='N'"
                     If nPuerto > 0 Then
                        Visor String(Int((19 - Len(tMensaje1)) / 2), " ") & tMensaje1, String(Int((19 - Len(tMensaje2)) / 2), " ") & tMensaje2, nPuerto, "N"
                     End If
                     Unload Me
                  Else
                     If RsDetalle.RecordCount = 0 Then
                        Cn.Execute "Update MPEDIDO set tEstadoPedido ='03', tMotivoAnulacion='000', tUsuarioAnulado='" & sUsuarioAutoriza & "', fRegAnulado= getdate(), tTurnoAnulado='" & sTurno & "', tObservacionAnulado='Cancelación de un Pedido en Blanco desde Caja Rapida'  where tCodigoPedido ='" & Pedido & "'"
                        Unload Me
                     Else
                        MsgBox "No debes tener pedidos sin atender", vbCritical, sMensaje
                     End If
                  End If
               End If
               
          Case Is = 3 ' Aceptar Precios
               GrabaProducto
               tabProducto.Visible = True
               fraDetalle.Visible = False
               ActivaCabecera True
               
          Case Is = 4 ' Cancelar Precios
               tabProducto.Visible = True
               fraDetalle.Visible = False
               ActivaCabecera True
                         
          Case Is = 5 ' Salir
               wEnter = True
               
               i = Calcular("select count(tCodigoPedido) as codigo from " & sDetalle & " where lImprime=1", Cn)
               If i > 0 Or RsDetalle.RecordCount > 0 Then
                  MsgBox "El pedido actual esta activo", vbExclamation, sMensaje
                  Exit Sub
               End If
               
               If Pedido = "" Then
                  If nPuerto > 0 Then
                     Visor String(Int((19 - Len(tMensaje1)) / 2), " ") & tMensaje1, String(Int((19 - Len(tMensaje2)) / 2), " ") & tMensaje2, nPuerto, "N"
                  End If
                  Unload Me
               Else
                  If sModulo = "INFOREST" Or sModulo = "ADICION" Then
                     Cn.Execute "delete from DPEDIDO where tCodigoPedido='" & Pedido & "'"
                     'Inserta el Detalle
                     Cn.Execute "Insert into DPEDIDO (tCodigoPedido, tItem, tTipoPedido, tCodigoProducto, tCodigoGrupo, tCodigoSubGrupo, tMoneda, " & _
                                "nPrecioNeto, nPrecioImpuesto1, nPrecioImpuesto2, nPrecioImpuesto3, nPrecioVenta, nRecargo, nDescuento, nPrecioOficial, " & _
                                "nCantidad, nImpuesto1, nImpuesto2, nImpuesto3, nVenta, tObservacion, tCortesia, lImprime, tEstadoItem, tArea, lImprimeArea,fregistro, nInsumo, nGasto, nManoObra, nOrden, lCombinacion, nCombinacion, lCorte,toferta,tautorizaoferta,tSubAlmacen,tCodigoEtiqueta,tunidadnegocio) " & _
                                "select tCodigoPedido, tItem, tTipoPedido, tCodigoProducto, tCodigoGrupo, tCodigoSubGrupo,  tMoneda, " & _
                                "nPrecioNeto, nPrecioImpuesto1, nPrecioImpuesto2, nPrecioImpuesto3, nPrecioVenta, nRecargo, nDescuento, nPrecioOficial, " & _
                                "nCantidad, nImpuesto1, nImpuesto2, nImpuesto3, nVenta, tObservacion, tCortesia, lImprime, tEstadoItem, tArea, lImprimeArea,getdate(), nInsumo, nGasto, nManoObra, nOrden, lCombinacion, nCombinacion, lCorte,toferta,tautorizaoferta,tSubAlmacen,tCodigoEtiqueta,'" & sUnidadNegocio & "' " & _
                                " From [" & sDetalle & "] where tEstadoItem='N'"
                     
                     If nPuerto > 0 Then
                        Visor String(Int((19 - Len(tMensaje1)) / 2), " ") & tMensaje1, String(Int((19 - Len(tMensaje2)) / 2), " ") & tMensaje2, nPuerto, "N"
                     End If
                     Unload Me
                  Else
                     If RsDetalle.RecordCount = 0 Then
                        Cn.Execute "Update MPEDIDO set tEstadoPedido ='03', tMotivoAnulacion='000', tUsuarioAnulado='" & sUsuarioAutoriza & "', fRegAnulado= getdate(), tTurnoAnulado='" & sTurno & "', tObservacionAnulado='Cancelación de un Pedido en Blanco desde Caja Rapida'  where tCodigoPedido ='" & Pedido & "'"
                        Unload Me
                     Else
                        MsgBox "No debes tener pedidos sin atender", vbCritical, sMensaje
                     End If
                     
                  End If
               End If
                         
          Case Is = 6 ' Aceptar Propiedades
          
               If wAgregaCombo Then
               
                    If ObligaPropiedad(sCombo) = False Then
                        Exit Sub
                    Else
                        grdDetalle.Enabled = True
                    End If
                    cmdOpcion(1).Enabled = False
                    RsCombo.Requery
                    RsCombo.MoveFirst
                    RsCombo.Find "titemCombo = '" & xItem & "'"
                    fraProductoCombo.Visible = True
                    fraPropiedad.Visible = False
                    
               Else
                    'Oscar Ortega---------------------------------
                    If ObligaPropiedad(sProducto) = False Then
                        Exit Sub
                    Else
                        grdDetalle.Enabled = True
                    End If
                    RsDetalle.Requery
                    RsDetalle.MoveFirst
                    RsDetalle.Find "titem = '" & sitem & "'"
                    ActivaCabecera True
                    tabProducto.Visible = True
                    fraPropiedad.Visible = False
               End If
               
          Case Is = 7 ' Observaciones
               If wAgregaCombo Then
                  If RsCombo!lImprime = True Then
                     MsgBox "El Item ya fue Impreso", vbExclamation, sMensaje
                     Exit Sub
                  End If
               Else
                  If RsDetalle!lImprime = True Then
                     MsgBox "El Item ya fue Impreso", vbExclamation, sMensaje
                     Exit Sub
                  End If
               End If
          
               frmKeyBoard.Caption = "Observación del Producto"
               frmKeyBoard.txtResultado.Text = lblObservacion.Text
               frmKeyBoard.Show vbModal
               If wEnter = True Then
                  If wAgregaCombo Then
                     nPos = RsCombo.AbsolutePosition
                     Cn.Execute "Update " & sComboDetalle & " set tObservacion = '" & sDescrip & "' where tItem ='" & sitem & "' and tItemCombo='" & xItem & "'"
                  Else
                     nPos = RsDetalle.AbsolutePosition
                     Cn.Execute "Update " & sDetalle & " set tObservacion = '" & sDescrip & "' where tItem ='" & sitem & "'"
                  End If
                  lblObservacion.Text = sDescrip
               End If
          
       Case Is = 8  'Imp.Pedido
             If lObservacion And Trim(txtObservacion.Caption) = "" Then
               MsgBox "Debes ingresar la Observación", vbInformation, sMensaje
               cmdDetalle_Click (6)
               If Trim(txtObservacion.Caption) = "" Then
                  Exit Sub
               End If
            End If
       
            If RsArea.RecordCount = 0 Then
               MsgBox "No existe area definida", vbInformation, sMensaje
               Exit Sub
            End If
            
            If RsDetalle.RecordCount = 0 Then
               Exit Sub
            End If
            
            'OO
            If ExistenPropiedadesPendientesEnPedido() Then
                If ExistenPropiedadesPendientesEnCombos() Then
                    Screen.MousePointer = vbHourglass
                                        
                     'insuimo2013
                    Cn.Execute "delete from " & sInsumoCombo
                    Dim X As Integer
                    Dim rstItems As New Recordset
                    Dim cadenaInsumos As String
                    Dim cadenaAEnviar As String
                    Dim cmdInsumo As New ADODB.Command
                    Dim resultado As String
                    Cn.Execute "insert into " & sInsumoCombo & " select sum(ncantidad) ncantidad, TPRODUCTO.TCODIGOINSUMO from " & sDetalle & " inner join tproducto on " & sDetalle & ".tcodigoproducto=tproducto.tcodigoproducto INNER JOIN  dbo.TINSUMO ON dbo.TPRODUCTO.tcodigoInsumo = dbo.TINSUMO.tcodigo  where lcontrolinsumocritico=1 and isnull(limprime,0)=0 AND ISNULL(TCODIGOINSUMO,''  )<>'' and (tinsumo.lactivo=1) group by tcodigoinsumo"
                    Cn.Execute "insert into " & sInsumoCombo & " sELECT     SUM(" & sDetalle & ".nCantidad * " & sComboDetalle & ".nCantidad) AS ncantidad, dbo.TPRODUCTO.tCodigoInsumo FROM  " & sComboDetalle & " INNER JOIN " & sDetalle & " ON  " & sComboDetalle & ".tItem = " & sDetalle & ".tItem INNER JOIN dbo.TINSUMO INNER JOIN dbo.TPRODUCTO ON dbo.TINSUMO.tcodigo = dbo.TPRODUCTO.tCodigoInsumo ON " & sComboDetalle & ".tProductoCombo = dbo.TPRODUCTO.tCodigoProducto WHERE     (dbo.TPRODUCTO.lControlInsumoCritico = 1) AND (dbo.TINSUMO.lactivo = 1) AND (ISNULL(" & sComboDetalle & ".lImprime, 0) = 0) AND (ISNULL(dbo.TPRODUCTO.tCodigoInsumo, N'') <> '') AND (" & sDetalle & ".lCombinacion = 1) GROUP BY dbo.TPRODUCTO.tCodigoInsumo "
                    cadenaInsumos = "select  SUM(ncantidad), tCodigoInsumo from " & sInsumoCombo & "  group by tCodigoInsumo order by 2"
                       Set rstItems = Lib.OpenRecordset(cadenaInsumos, Cn)
                                        If Not (rstItems.EOF Or rstItems.BOF) Then
                                             rstItems.MoveFirst
                                             For X = 0 To rstItems.RecordCount - 1
                                                 cadenaAEnviar = cadenaAEnviar + rstItems.Fields(1) + "|" + str(rstItems.Fields(0)) + "$"
                                                 rstItems.MoveNext
                                             Next X
                                           
                                            Set cmdInsumo = New ADODB.Command
 
                                            With cmdInsumo
                                                 .ActiveConnection = Cn
                                                 .CommandType = adCmdStoredProc
                                                 .CommandText = "USP_actualizaStockInsumo"
                                                 .Parameters.Refresh
                                                 .Parameters("@vi_detalles") = cadenaAEnviar
                                                 .Parameters("@vi_numdet") = rstItems.RecordCount
                                                 .Parameters("@vch_Salida") = ""
                                            End With
                                            cmdInsumo.Execute
                                            resultado = cmdInsumo.Parameters("@vch_Salida").value
                                            If resultado <> "1" Then:  MsgBox "No hay cantidad disponible de : " & resultado: variableEmite = False: Screen.MousePointer = vbDefault: Exit Sub
                                        End If
                                        variableEmite = True
                    'InsumosCriticos
                    
                    
                    If Pedido = "" Then
                       GeneraPedido
                    Else
                       ActualizaPedido
                    End If
                                
                    If nPuerto > 0 Then
                       Visor "Enviando Pedido...", "", nPuerto, "N"
                    End If
                                                        
                    If lOrden Then
                       Isql = "select * from vPedido " & _
                              "Where Codigo = '" & Pedido & "' and nOrden in (select nOrden from DPEDIDO where tCodigoPedido='" & Pedido & "' and (lImprime = 0 or (isnull(lImprimeAreaCombo,0) = 1  and isnull(lImprimeCombo,0) = 0 ))) " & _
                              "ORDER BY nOrden, tItem, nOrdenCombo,tetiqueta,combo " 'tItemCombo"
                    Else
                       Isql = "select * from vPedido " & _
                              "Where Codigo = '" & Pedido & "' And lImprimeArea = 1 and (lImprime = 0 or (isnull(lImprimeAreaCombo,0) = 1  and isnull(lImprimeCombo,0) = 0 ))" & _
                              "ORDER BY nOrden, tItem, nOrdenCombo, tetiqueta,combo " ' tItemCombo"
                    End If
                                                                    
                    Set RsImpresion = Lib.OpenRecordset(Isql, Cn)
        
                    If Not RsImpresion.EOF Then
                       RsArea.MoveFirst
                       For i = 1 To RsArea.RecordCount
                           RsImpresion.Filter = "tArea='" & RsArea!tArea & "'"
                           If RsArea!tIcono = "" Or RsArea!nValor = 1 Then
                              If RsImpresion.RecordCount <> 0 Then
                                 RsImpresion.MoveFirst
                                 sPedido = Pedido
                                 ImprimePedido RsImpresion, "N", RsArea!timpresora, RsArea!Area, False, RsProductoPropiedad, RsComboPropiedad, "Rapido"
                                 sPedido = ""
                              End If
                           End If
        
                           RsArea.MoveNext
                       Next i
                    End If
                     
                    If lKDS Then
                        Dim kdsRsCabecera As Recordset
                        Isql = "SELECT * From vPedidoCabecera Where Codigo = '" & Pedido & "' Order By codigo "
                        Set kdsRsCabecera = Lib.OpenRecordset(Isql, Cn)
                       Call KDS_AnadirNuevaOrden(kdsRsCabecera)
                    End If
                    
                    Cn.Execute "Update CPEDIDO Set lImprime = 1 where tCodigoPedido = '" & Pedido & "'"
                    Cn.Execute "Update " & sComboDetalle & "  Set lImprime = 1"
                    RsCombo.Requery
                    
                    'CESAR----CHEF CONTROL
                    Dim ChefEnvio As Boolean
                    ChefEnvio = Calcular("select ISNULL(lEnvioChef,0) as Codigo FROM TPARAMETRO", Cn)
                    
                    Cn.Execute "Update DPEDIDO Set lNoCantado=0 where tCodigoPedido = '" & Pedido & "' and lNoCantado IS NULL"
                    
                    If ChefEnvio Then
                    Cn.Execute "Update DPEDIDO Set lCantadoc=1,fCantadoC=GetDate(), lTipoEnvio=0 where tCodigoPedido = '" & Pedido & "' and lImprime<>1"
                    End If
                    '---------------------------------
        
                    Cn.Execute "Update DPEDIDO Set lImprime = 1, fenvio=getdate(), nEnvio= isnull(nEnvio,0) + 1 where tCodigoPedido = '" & Pedido & "' and limprime<>1 "
                    Cn.Execute "update MPEDIDO set nReimpresion = isnull(nReimpresion,0) + 1 where tCodigoPedido='" & Pedido & "'"
                    Cn.Execute "Update " & sDetalle & "  Set fenvio = getdate(), nEnvio = isnull(nEnvio,0) + 1 where limprime <> 1"
                    Cn.Execute "Update " & sDetalle & "  Set lImprime = 1"
                    sPedido = Pedido
            
                End If
            End If
            RsDetalle.Requery
            LimpiaRs
             Label2.Caption = muestra
            Screen.MousePointer = vbDefault
                        
       Case Is = 9  'Cliente Frecuente
                If Not Sw Then
                   'sTemp = txtTelefono.Caption
                End If

                wEnter = False
                sTipo = sTipoPedido
                sCodigo = ""
                sCodigoParienteSeleccionado = ""
                sCodigoInvitado = ""
                If lClub Then
                    frmBusquedaSocio.Show vbModal
                Else
                
                    frmBusquedaDelivery.Show vbModal
                End If
                If wEnter = True Then
                
                   sClienteFrecuente = sCodigo
                   txtCliente.Caption = sDescrip
                   txtTelefono.Caption = sClienteFrecuente
                   xDescuento = nVar1
                   sCodigoDescuento = "000"
                   sDescripcionDescuento = "Descuento por Cliente Frecuente"
                End If
               
       Case Is = 10  'Combos
            If wCombo Then
               tabProducto.Visible = False
               fraCombo.Visible = True
               fraProductoCombo.Visible = True
               wAgregaCombo = True
               ActivaCabecera False
               
               If Not RsCombo.EOF Then
                  RsCombo.MoveFirst
               End If
               AsignaProductoCombo
               
               txtBarra.SetFocus
            End If
            
       'OO--------------------------------------------------------------------------------------------------------
       Case Is = 14  'Cargos
            If RsDetalle.RecordCount = 0 Then
               Exit Sub
            End If
            
            If sPuntoVenta = "" Then
               MsgBox "Falta ingresar el punto de venta", vbExclamation, sMensaje
               Exit Sub
            End If
            
            If Calcular("select count(tFacturado) as Codigo from DPEDIDO where tCodigoPedido ='" & sPedido & "' and isnull(tFacturado,'0') <> '0' and len(ltrim(tFacturado)) <> 0", Cn) > 0 Then
               MsgBox "Imposible pasar el pedido a Cargos, pedidos con items Facturados", vbExclamation, sMensaje
               Exit Sub
            End If
            
            'Clave de Multi Cajero
            tUsuActua = sUsuario
            If lMultiCajero Then
               If Supervisor("16") = False Then
                  MsgBox "Clave no permitida", vbExclamation, sMensaje
                  Exit Sub
               End If
               sUsuario = sVar1
            End If
            sUsuario = tUsuActua
            If lPrinter And lObligaPrinter Then
               i = Calcular("select count(tCodigoPedido) as codigo from " & sDetalle & " where lImprime=0", Cn)
               If i > 0 Then
                  cmdOpcion_Click (8)
               End If
            End If
            sPedido = Pedido
                                                                  
            'Chequea si existe platos a facturar
            sTD = "N"
            RsDetalle.MoveFirst
            Do While Not RsDetalle.EOF
               If (Len(Trim(RsDetalle!tFacturado)) = 0 Or IsNull(RsDetalle!tFacturado)) Then
                  sTD = "S"
                  Exit Do
               End If
               RsDetalle.MoveNext
            Loop
    
            If sTD <> "S" Then
               MsgBox "Error: No existen Productos a Facturar", vbCritical, sMensaje
               Exit Sub
            End If

            frmCargo.Show vbModal

            If Not wEnter Then
               Exit Sub
            End If
           
            Dim tItem As Integer
            Dim Correlativo As Integer
            Dim CorrelaProp As Integer
            Dim nMovimiento As Integer
            Dim sAsignado   As String
            Dim MonPuntoventa As String
            
            sCliente = sCodigo
            xSuma = Calcular("select sum(nVenta) as Codigo FROM DPEDIDO where tEstadoItem = 'N' and (isnull(tFacturado,'0') = '0' or len(ltrim(tFacturado)) = 0) and tCodigoPedido='" & sPedido & "'", Cn)
            
            If sDescrip = "Infhotel" Then
               If MsgBox("Esta seguro de Generar el Pedido Nro: " & sPedido & " en Infhotel?", vbQuestion + vbOKCancel, sMensaje) = vbCancel Then
                  Exit Sub
               End If
               sHabitacion = ""
               sReserva = ""
               sPasajero = ""
            Else
               If MsgBox("Esta seguro de Enviar el Pedido Nro: " & sPedido & _
                  Chr(13) & "a la " & Trim(sDescrip) & " " & IIf(sDescrip = "Reserva", sReserva, sHabitacion) & "?", vbQuestion + vbOKCancel, sMensaje) = vbCancel Then
                  Exit Sub
               End If
            End If
            
            MonPuntoventa = Calcular("select max(tmoneda) as codigo from vpuntoventa where tpuntoventa='" & sPuntoVenta & "'", CnInfhotel)
               Call PUltimaComanda
               sComandaInfhotel = Calcular("select left(MAX(tComanda),8) as Codigo from MCOMANDA where tPuntoVenta='" & sPuntoVenta & "'", CnInfhotel)
               sComandaInfhotel = Lib.Correlativo(sComandaInfhotel, 8)
               CnInfhotel.Execute "Update TPUNTOVENTA Set nUltimoComanda = '" & sComandaInfhotel & "' where tPuntoVenta='" & sPuntoVenta & "'"
               sComandaInfhotel = sComandaInfhotel & "-" & UCase(Mid(rsPuntoVenta!Descripcion, 1, 3))
               rsPuntoVenta.Requery
                              
               'Genero las comandas en Infhotel
               'Cabecera
               If sDescrip = "Infhotel" Then sAsignado = "01"
               If sDescrip = "Reserva" Then sAsignado = "03"
               If sDescrip = "Habitacion" Then sAsignado = "02"
               If sTipoComanda = "01" Then
                    Isql = "Insert into MCOMANDA " & _
                           "(tComanda, tPuntoVenta, tHotel, nMovimiento, fFecha, hHora, nTotal, tEstado, " & _
                           "tEmitido, tAsignacion, tCodigoReserva, tNumeroHabitacion, tCodigoFuncionario, " & _
                           "tCaja, tDocumento, tUsuario, nTCambio, tCodigoCompania, tCliente, tMoneda, fFechaE, hHoraE, tUsuarioE, tIncluido, nRoomSer, nDescuento, tNotaPedido, tCompania, tContacto, lInforest ) " & _
                           "values('" & sComandaInfhotel & "', '" & sPuntoVenta & "', '" & sHotel & "', 1,  getdate(), getdate(), " & IIf(MonPuntoventa = "01", xSuma, xSuma / nTC) & ", '01', " & _
                           "1,'" & sAsignado & "', '" & sReserva & "', '" & sHabitacion & "', '', " & _
                           "'', '', '" & xUsuario & "', " & nTC & ", '', '" & sPasajero & "', '" & MonPuntoventa & "', getdate(), getdate(), '" & xUsuario & "','" & sTipoComanda & "', " & IIf(sTipoPedido = "03", nLlevar, 0) & ", " & xDescuento & ", '" & sPedido & "', '" & sCompania & "', '" & sContacto & "', 1)"
                    CnInfhotel.Execute Isql
               Else
                    Isql = "Insert into MCOMANDA " & _
                          "(tComanda, tPuntoVenta, tHotel, nMovimiento, fFecha, hHora, nTotal, tEstado, " & _
                          "tEmitido, tAsignacion, tCodigoReserva, tNumeroHabitacion, tCodigoFuncionario, " & _
                          "tCaja, tDocumento, tUsuario, nTCambio, tCodigoCompania, tCliente, tMoneda, fFechaE, hHoraE, tUsuarioE, tIncluido, nRoomSer, nDescuento, tNotaPedido, tCompania, tContacto, lInforest) " & _
                          "values('" & sComandaInfhotel & "', '" & sPuntoVenta & "', '" & sHotel & "', 1,  getdate(), getdate(), 0, '01', " & _
                          "1,'" & sAsignado & "', '" & sReserva & "', '" & sHabitacion & "', '', " & _
                          "'', '', '" & xUsuario & "', " & nTC & ", '', '" & sPasajero & "', '" & MonPuntoventa & "', getdate(), getdate(), '" & xUsuario & "','" & sTipoComanda & "', " & IIf(sTipoPedido = "03", nLlevar, 0) & ", " & xDescuento & ", '" & sPedido & "', '" & sCompania & "', '" & sContacto & "', 1)"
                    CnInfhotel.Execute Isql
               End If
            
            If sDescrip = "Habitacion" Then
               txtObservacion.Caption = "Hab: " & sHabitacion
            ElseIf sDescrip = "Reserva" Then
               txtObservacion.Caption = "Res: " & sReserva
               sHabitacion = ""
            End If
            Isql = "Update MPEDIDO Set " & _
                    "tEstadoPedido ='05', " & _
                    "tComanda ='" & sComandaInfhotel & "', " & _
                    "tPuntoVenta ='" & sPuntoVenta & "', " & _
                    "tReserva ='" & sReserva & "', " & _
                    "tHabitacion ='" & sHabitacion & "', " & _
                    "tObservacion='" & txtObservacion.Caption & "', " & _
                    "tFichaPasajero ='" & sFichaPasajero & "', " & _
                    "tTipoComanda ='" & sTipoComanda & "', " & _
                    "tPasajero ='" & sPasajero & "' " & _
                    "  where tCodigoPedido = '" & sPedido & "'"
            Cn.Execute Isql
            
            'Detalle
''            CnInfhotel.Execute "delete from DCOMANDA where tComanda ='" & RsCabecera!tComanda & "'"
            nMovimiento = Calcular("select max(nmovimiento) as codigo from dcomanda where tcomanda='" & sComandaInfhotel & "'", CnInfhotel) + 1
            If sTipoComanda = "01" Then
                Isql = "Insert into DCOMANDA " & _
                       "(tComanda, tPuntoVenta, tHotel, tItem, nMovimiento, tNotaPedido, tCodigoItem, " & _
                       "nPrecioUnitario, nCantidad, nTotal, nPrecioCos, tCodigoReserva, tNumeroHabitacion, " & _
                       "tCuenta, tCaja, tDocumento, tAsignado, tUsuario, fFecha, hHora) " & _
                       "select '" & sComandaInfhotel & "' as tComanda, '" & sPuntoVenta & "' as tPuntoVenta, '" & sHotel & "' as tHotel, tItem , " & nMovimiento & " as nMovimiento, '" & sPedido & "' as tNotaPedido, tInfhotel as tCodigoItem, " & _
                       IIf(MonPuntoventa = "01", "T1.nPrecioVenta", "T1.nPrecioVenta / " & nTC) & " as nPrecioUnitario, nCantidad, " & IIf(MonPuntoventa = "01", "nVenta", "nVenta / " & nTC) & " as nTotal, T1.nInsumo+T1.nGasto+T1.nManoObra as nPrecioCos, '" & sReserva & "' as tCodigoReserva, '" & sHabitacion & "' as tNumeroHabitacion, " & _
                       "'' as tCuenta, '' as tCaja, '' as tDocuemento, '" & sAsignado & "' as tAsignado, '" & xUsuario & "' as  tUsuario, getdate() as fFecha, getdate() as hHoraMovimiento " & _
                       "FROM OPENROWSET('SQLOLEDB','" & sRuta & "';'" & sUserName & "';'" & sUserPassword & "', " & sMDB & ".dbo.DPEDIDO) T1 INNER JOIN OPENROWSET('SQLOLEDB','" & sRuta & "';'" & sUserName & "';'" & sUserPassword & "', " & sMDB & ".dbo.TPRODUCTO) T2 ON T1.tCodigoProducto = T2.tCodigoProducto " & _
                       "where (isnull(tFacturado,'0') = '0' or len(ltrim(tFacturado)) = 0) and tCodigoPedido ='" & sPedido & "'"
                CnInfhotel.Execute Isql
            Else
                Isql = "Insert into DCOMANDA " & _
                       "(tComanda, tPuntoVenta, tHotel, tItem, nMovimiento, tNotaPedido, tCodigoItem, " & _
                       "nPrecioUnitario, nCantidad, nTotal, nPrecioCos, tCodigoReserva, tNumeroHabitacion, " & _
                       "tCuenta, tCaja, tDocumento, tAsignado, tUsuario, fFecha, hHora) " & _
                       "select '" & sComandaInfhotel & "' as tComanda, '" & sPuntoVenta & "' as tPuntoVenta, '" & sHotel & "' as tHotel, tItem ," & nMovimiento & " as nMovimiento, '" & sPedido & "' as tNotaPedido, tInfhotel as tCodigoItem, " & _
                       " 0 as nPrecioUnitario, nCantidad, 0 as nTotal, T1.nInsumo+T1.nGasto+T1.nManoObra as nPrecioCos, '" & sReserva & "' as tCodigoReserva, '" & sHabitacion & "' as tNumeroHabitacion, " & _
                       "'' as tCuenta, '' as tCaja, '' as tDocuemento, '" & sAsignado & "' as tAsignado, '" & xUsuario & "' as  tUsuario, getdate() as fFecha, getdate() as hHoraMovimiento " & _
                       "FROM OPENROWSET('SQLOLEDB','" & sRuta & "';'" & sUserName & "';'" & sUserPassword & "', " & sMDB & ".dbo.DPEDIDO) T1 INNER JOIN OPENROWSET('SQLOLEDB','" & sRuta & "';'" & sUserName & "';'" & sUserPassword & "', " & sMDB & ".dbo.TPRODUCTO) T2 ON T1.tCodigoProducto = T2.tCodigoProducto " & _
                       "where (isnull(tFacturado,'0') = '0' or len(ltrim(tFacturado)) = 0) and tCodigoPedido ='" & sPedido & "'"
                CnInfhotel.Execute Isql
            End If
            
            'Propinas
            If Val(sPropina) > 0 Then
                'Grabo la propina en el Mpropina del Inforest
                Isql = "insert into MPROPINA " & _
                      "(tcodigopedido,fregistro,tmoneda,nmonto,tusuario, tComanda) " & _
                      "values('" & sPedido & "',getdate(),'" & sMonPropina & "'," & sPropina & ",'" & xUsuario & "', '" & sComandaInfhotel & "')"
                Cn.Execute Isql
                
                'Grabo la propina en el Detalle de la comanda
                tItem = Calcular("select max(titem) as codigo from dcomanda where tcomanda='" & sComandaInfhotel & "'", CnInfhotel) + 1
                Isql = "Insert Into dcomanda " & _
                       "(tComanda, tPuntoVenta, tHotel, tItem, nMovimiento, tNotaPedido, tCodigoItem, " & _
                       "nPrecioUnitario, nCantidad, nTotal, nPrecioCos, tCodigoReserva, tNumeroHabitacion, " & _
                       "tCuenta, tCaja, tDocumento, tAsignado, tUsuario, fFecha, hHora) " & _
                       "values('" & sComandaInfhotel & "','" & sPuntoVenta & "','00','" & tItem & "','1','" & sPedido & "', " & _
                       "'100000'," & sPropina & ",'1'," & sPropina & ",'','" & sReserva & "','" & sHabitacion & "','', " & _
                       " '','','" & sAsignado & "','" & xUsuario & "',getdate(),getdate())"
                CnInfhotel.Execute Isql
               
                'Grabo la propina en el Mpropina del Infhotel
                CorrelaProp = Calcular("select max(ncorrela) as codigo from mpropina", CnInfhotel) + 1
            
                Isql = "Insert Into MPROPINA " & _
                       "(ncorrela,tcodigoreserva,tnumerohabitacion,tcomanda,tcodigoitem,tmoneda,nmonto,tdocumento,tresponsable,testado,ffecha,tusuario, tPuntoVenta) " & _
                       "values(" & CorrelaProp & ",'" & sReserva & "','" & sHabitacion & "','" & sComandaInfhotel & "','100000','" & sMonPropina & "'," & sPropina & ", " & _
                       "'','" & Mid(sMozo, 2, 3) & "','01',getdate(),'" & xUsuario & "','" & sPuntoVenta & "')"
                CnInfhotel.Execute Isql
                
                Isql = "Update MCOMANDA set ncorrelaprop=" & CorrelaProp & " where tcomanda='" & sComandaInfhotel & "' and tPuntoVenta='" & sPuntoVenta & "'"
                CnInfhotel.Execute Isql
            End If
                        
            'Actualiza las Cuentas Corrientes Infhotel
            If sDescrip = "Reserva" Then
               i = Calcular("select max(tNumeroCorrelativo) as Codigo from TCUENTARESERVA where tCodigoReserva='" & sReserva & "'", CnInfhotel)
               Isql = "Insert into TCUENTARESERVA " & _
                      "(tCodigoReserva, tNumeroHabitacion, fFecha, hHoraMovimiento, tComanda, tNotaPedido, tCodigoItem, nPrecioUnitario, nCantidad,testado, ttipo,tHotel, " & _
                      " nTotal, tNumeroCorrelativo,tpuntoventa, tItem, tUsuario) " & _
                      "select '" & sReserva & "' as tCodigoReserva, '" & sHabitacion & "' as tNumeroHabitacion , getdate() as fFecha, getdate() as hHoraMovimiento, '" & sComandaInfhotel & "' as tComanda, '" & sPedido & "' as tNotaPedido, tInfhotel as tCodigoItem, " & _
                      IIf(sMonedaBase = "01", "T1.nPrecioVenta", "T1.nPrecioVenta / " & nTC) & " as nPrecioUnitario, nCantidad,'' as testado,'' as ttipo,'" & sHotel & "' as tHotel," & IIf(sMonedaBase = "01", "nVenta", "nVenta / " & nTC) & " as nTotal, tItem + " & i & ",'" & sPuntoVenta & "' as tpuntoventa ,tItem, '" & xUsuario & "' as tUsuario " & _
                      "FROM OPENROWSET('SQLOLEDB','" & sRuta & "';'" & sUserName & "';'" & sUserPassword & "', " & sMDB & ".dbo.DPEDIDO) T1 INNER JOIN OPENROWSET('SQLOLEDB','" & sRuta & "';'" & sUserName & "';'" & sUserPassword & "', " & sMDB & ".dbo.TPRODUCTO) T2 ON T1.tCodigoProducto = T2.tCodigoProducto " & _
                      "where (isnull(tFacturado,'0') = '0' or len(ltrim(tFacturado)) = 0) and tCodigoPedido ='" & sPedido & "'"
               CnInfhotel.Execute Isql
                             
               If Val(sPropina) > 0 Then
                    'Graba la propina en la Cta Cte de la Reserva
                    i = Calcular("select max(titem) as Codigo from TCUENTARESERVA where tCodigoReserva='" & sReserva & "'and tcomanda='" & sComandaInfhotel & "'", CnInfhotel) + 1
                    Correlativo = Calcular("select max(tnumerocorrelativo) as Codigo from TCUENTARESERVA where tCodigoReserva='" & sReserva & "'", CnInfhotel) + 1
                    Isql = "insert into tcuentareserva " & _
                          "(tCodigoReserva, tNumeroHabitacion, fFecha, hHoraMovimiento, tComanda, tNotaPedido, tCodigoItem, nPrecioUnitario, nCantidad,testado,ttipo,tHotel, " & _
                          " nTotal, tNumeroCorrelativo,tpuntoventa ,tItem, tUsuario,ncorrelaprop) " & _
                          " values('" & sReserva & "','" & sHabitacion & "',getdate(),getdate(),'" & sComandaInfhotel & "','" & sPedido & "','100000'," & sPropina & ",'1','','','00', " & _
                          " " & sPropina & "," & Correlativo & ",'" & sPuntoVenta & "'," & i & ",'" & xUsuario & "'," & CorrelaProp & ")"
                    CnInfhotel.Execute Isql
                    
                    If sMonedaBase = sMonPropina Then
                        Isql = "UPDATE tCuentaReserva " & _
                              "SET nPrecioUnitario=" & CDbl(sPropina) & "," & _
                              "nCantidad=1," & _
                              "nTotal=" & CDbl(sPropina) & " " & _
                              "WHERE ncorrelaprop='" & CorrelaProp & "'"
                         CnInfhotel.Execute Isql
                    Else
                        If (sMonedaBase = "02" And sMonPropina = "01") Then
                        'La moneda base esta en $ y la propina esta en S/.
                            Isql = "UPDATE tcuentareserva " & _
                                 "SET nPrecioUnitario=" & CDbl(sPropina) / nTC & "," & _
                                 "nCantidad=1," & _
                                 "ntotal=" & CDbl(sPropina) / nTC & "" & _
                                 "WHERE nCorrelaProp='" & CorrelaProp & "'"
                            CnInfhotel.Execute Isql
                         Else
                            'La moneda base esta en S/. y la propina esta en $
                            Isql = "UPDATE tcuentaReserva" & _
                                 "SET nPrecioUnitario=" & CDbl(sPropina) * nTC & "," & _
                                 "nCantidad=1," & _
                                 "ntotal=" & CDbl(sPropina) * nTC & "" & _
                                 "WHERE nCorrelaProp='" & CorrelaProp & "'"
                            CnInfhotel.Execute Isql
                         End If
                    End If
                End If

            ElseIf sDescrip = "Habitacion" Then
               i = Calcular("select max(tNumeroCorrelativo) as Codigo from TCUENTAHABITACION where tNumeroHabitacion='" & sHabitacion & "' and tCodigoReserva='" & sReserva & "'", CnInfhotel)
               Isql = "Insert into TCUENTAHABITACION " & _
                      "(tCodigoReserva, tNumeroHabitacion, fFecha, hHoraMovimiento,testado,ttipo, tComanda, tNotaPedido, tCodigoItem, nPrecioUnitario, nCantidad, tHotel, " & _
                      " nTotal, tNumeroCorrelativo,tpuntoventa, tItem, tUsuario) " & _
                      "select '" & sReserva & "' as tCodigoReserva, '" & sHabitacion & "' as tNumeroHabitacion , getdate() as fFecha, getdate() as hHoraMovimiento,'' as testado,'' as ttipo ,'" & sComandaInfhotel & "' as tComanda, '" & sPedido & "' as tNotaPedido, tInfhotel as tCodigoItem, " & _
                      IIf(sMonedaBase = "01", "T1.nPrecioVenta", "T1.nPrecioVenta / " & nTC) & " as nPrecioUnitario, nCantidad, '" & sHotel & "' as tHotel," & IIf(sMonedaBase = "01", "nVenta", "nVenta / " & nTC) & " as nTotal, tItem + " & i & ",'" & sPuntoVenta & "' as tpuntoventa ,tItem, '" & xUsuario & "' as tUsuario " & _
                      "FROM OPENROWSET('SQLOLEDB','" & sRuta & "';'" & sUserName & "';'" & sUserPassword & "', " & sMDB & ".dbo.DPEDIDO) T1 INNER JOIN OPENROWSET('SQLOLEDB','" & sRuta & "';'" & sUserName & "';'" & sUserPassword & "', " & sMDB & ".dbo.TPRODUCTO) T2 ON T1.tCodigoProducto = T2.tCodigoProducto " & _
                      "where (isnull(tFacturado,'0') = '0' or len(ltrim(tFacturado)) = 0) and tCodigoPedido ='" & sPedido & "'"
               CnInfhotel.Execute Isql
               
               If Val(sPropina) > 0 Then
                  'Graba la propina en la Cta Cte de la Habitacion
                  i = Calcular("select max(titem) as Codigo from TCUENTAHABITACION where tCodigoReserva='" & sReserva & "'and tcomanda='" & sComandaInfhotel & "'", CnInfhotel) + 1
                  Correlativo = Calcular("select max(tnumerocorrelativo) as Codigo from TCUENTAHABITACION where tCodigoReserva='" & sReserva & "'", CnInfhotel) + 1
                  Isql = "Insert into TCUENTAHABITACION " & _
                         "(tCodigoReserva, tNumeroHabitacion, fFecha, hHoraMovimiento,testado,ttipo, tComanda, tNotaPedido, tCodigoItem, nPrecioUnitario, nCantidad, tHotel, " & _
                         " nTotal, tNumeroCorrelativo,tpuntoventa ,tItem, tUsuario,ncorrelaprop) " & _
                         " Values('" & sReserva & "','" & sHabitacion & "',getdate(),getdate(),'','','" & sComandaInfhotel & "','" & sPedido & "','100000'," & sPropina & ",'1','00', " & _
                         " " & sPropina & "," & Correlativo & ",'" & sPuntoVenta & "'," & i & ",'" & xUsuario & "'," & CorrelaProp & ")"
                  CnInfhotel.Execute Isql
                  
               If sMonedaBase = sMonPropina Then
                   Isql = "UPDATE tCuentaHabitacion " & _
                         "SET nPrecioUnitario=" & CDbl(sPropina) & "," & _
                         "nCantidad=1," & _
                         "nTotal=" & CDbl(sPropina) & " " & _
                         "WHERE ncorrelaprop='" & CorrelaProp & "'"
                    CnInfhotel.Execute Isql
                  Else
                    If (sMonedaBase = "02" And sMonPropina = "01") Then
                    'La moneda base esta en $ y la propina esta en S/.
                        Isql = "UPDATE tcuentahabitacion " & _
                             "SET nPrecioUnitario=" & CDbl(sPropina) / nTC & "," & _
                             "nCantidad=1," & _
                             "ntotal=" & CDbl(sPropina) / nTC & "" & _
                             "WHERE nCorrelaProp='" & CorrelaProp & "'"
                        CnInfhotel.Execute Isql
                    Else
                        'La moneda base esta en S/. y la propina esta en $
                        Isql = "UPDATE tcuentahabitacion " & _
                             "SET nPrecioUnitario=" & CDbl(sPropina) * nTC & "," & _
                             "nCantidad=1," & _
                             "ntotal=" & CDbl(sPropina) * nTC & "" & _
                             "WHERE nCorrelaProp='" & CorrelaProp & "'"
                        CnInfhotel.Execute Isql
                    End If
                  End If
               End If
            End If

           'CarlosD 13/11/2006
            CorrelativoC = 0
            CorrelativoC = Val(Calcular("select max(nmovimiento) as codigo from wmcomanda where tcomanda='" & sComandaInfhotel & "' and tpuntoventa='" & sPuntoVenta & "' and thotel='" & sHotel & "'", CnInfhotel)) + 1
            
            Isql = "INSERT INTO WMCOMANDA([tComanda],[tPuntoVenta],[fFecha],[hHora],[tMoneda],[nTotal],[tEstado],[tEmitido], [tAsignacion],[tCodigoReserva],[tNumeroHabitacion],[tUsuario],[tCodigoCompania],[tCliente],[tCodigoFuncionario],[tHotel],[NTCAMBIO],[NMOVIMIENTO],[NDESCUENTO],[NROOMSER],[TMOZO],[tIncluido],[tMesa], tNotaPedido, tCompania, tContacto ) " & _
                   "Values (" & _
                    "'" & sComandaInfhotel & "', " & _
                    "'" & sPuntoVenta & "', " & _
                    "Getdate(), " & _
                    "GetDate(), " & _
                    "'" & MonPuntoventa & "'," & _
                    "" & IIf(MonPuntoventa = "01", xSuma, xSuma / nTC) & ", " & _
                    "'01',0, " & _
                    "'" & sAsignado & "', " & _
                    "'" & sReserva & "', " & _
                    "'" & sHabitacion & "', " & _
                    "'" & xUsuario & "', " & _
                    "'', " & _
                    "'" & sCliente & "', " & _
                    "'', " & _
                    "'" & sHotel & "', " & _
                    "'" & nTC & "', '" & CorrelativoC & "', " & nDescuento & ",'','','" & sTipoComanda & "', '" & sMesa & "', '" & sPedido & "', '" & sCompania & "', '" & sContacto & "')"
                CnInfhotel.Execute Isql

            Isql = "INSERT INTO WDCOMANDA SELECT * From dbo.DComanda where tComanda='" & sComandaInfhotel & "' and tPuntoVenta='" & sPuntoVenta & "'"
            CnInfhotel.Execute Isql
            
            'Libera la Mesa
            Cn.Execute "Update TMESA set tEstadoMesa = '04' where tCodigoMesa ='" & sMesa & "'"
            'Juntar Mesa
            Cn.Execute "update TMESA set tEstadoMesa='01' where tCodigoMesa in (select tMesa from TPEDIDOMESA where tCodigoPedido='" & sPedido & "')"
            
            wEnter = False
            
            Dim xPrecuenta As Boolean
            xPrecuenta = False
            If lObligaPrecuenta Then
               xPrecuenta = True
            Else
               If MsgBox("Deseas imprimir la Precuenta?", vbQuestion + vbYesNo, sMensaje) = vbYes Then
                  xPrecuenta = True
               End If
            End If
            
            If xPrecuenta Then
                Screen.MousePointer = vbHourglass
                'Imprime Precuenta
                If lPrecuentaImpresora Then
                   frmPrecuentaImpresora.Show vbModal
                   If Not wEnter Then
                      sPropina = ""
                      sTipoComanda = ""
                      Exit Sub
                   End If
                Else
                   sCodigo = sPreCuenta
                End If
                                        
                If lPrecuentaAgrupada Then
                   Isql = "select * from vPrecuentaAgrupada WHERE Codigo='" & sPedido & "' order by tItem"
                Else
                   Isql = "select * from vPrecuenta WHERE Codigo='" & sPedido & "' order by tItem"
                End If
    
                Set RsImpresion = Lib.OpenRecordset(Isql, Cn)
                ImprimeInfhotel RsImpresion, sCodigo
                Cn.Execute "update MPEDIDO set fRegCuenta = getDate() where tCodigoPedido='" & sPedido & "'"
                LimpiaRs
            End If
            
            sPropina = ""
            sTipoComanda = ""
            
            RsDetalle.Requery
            RsCombo.Requery
            
            sHabitacion = ""
            sReserva = ""
            sPasajero = ""
            Pedido = ""
            sPedido = ""

            Cn.Execute "delete " & sDetalle
            Cn.Execute "delete " & sComboDetalle
            Cn.Execute "delete " & sComboPropiedad
            Cn.Execute "delete " & sProductoPropiedad
    
            RsDetalle.Requery
            RsComboPropiedad.Requery
            RsProductoPropiedad.Requery
            Inicializar
            Screen.MousePointer = vbDefault

        'OO Fin--------------------------------------------------------------------------------------------------------
                        
       Case Is = 17  'Cancelacion
            sCodigo = ""
            sDescrip = ""
            fraEliminacion.Visible = False
            tabProducto.Visible = True
            ActivaCabecera True
   End Select
End Sub

Private Sub cmdOperador_Click(Index As Integer)
   Dim i As Integer
   Screen.MousePointer = vbHourglass
   For i = 1 To 13
       cmdOperador(i).backColor = vbButtonFace
   Next i
   RsOperador.MoveFirst
   RsOperador.Find "nboton = " & Trim(str(Index))
   nOperadorPropiedad = RsOperador!nControl
    xOperador = RsOperador!codigo

   cmdOperador(Index).backColor = vbRed
   If wAgregaCombo Then
      AsignaComboPropiedad
   Else
      AsignaPropiedad
   End If
   Screen.MousePointer = vbDefault
End Sub

Private Sub cmdOrigen_Click(Index As Integer)
' origen de ventas
   'HabilitaTimerColor (False)
   RsOrigenVentas.MoveFirst
   RsOrigenVentas.Find "boton = " & Trim(str(Index))
   vOrigenVentas = RsOrigenVentas!CodOrigenVenta
   'sMotorizado = RsMotorizado!codigo
   'HabilitaTimerColor (True)
   validarOrigenVentas
End Sub
Private Sub cmdPrecio_Click()
   Dim Acumulado As Double
   sTipo = ""
   frmNumPad.Show vbModal
   If wEnter And Val(nPVenta) > 0 Then
      nPVenta = Val(sDescrip)
      nOficial = nPVenta
      nDescuento = 0
      nRecargo = 0
      txtDPorcentaje.Caption = Format(0, "###,###,###,##0.00")
      txtRPorcentaje.Caption = Format(0, "###,###,###,##0.00")
      txtDImporte.Caption = Format(nDescuento, "###,###,###,##0.00")
      txtRImporte.Caption = Format(nRecargo, "###,###,###,##0.00")
      
       'extranjero bolivia
      Select Case pais ' ok
        Case "001" 'Bolivia
                Acumulado = IIf(txtImpuesto1.Caption <> 0, Acumulado + nPorcentaje1, Acumulado)
                Acumulado = IIf(txtImpuesto2.Caption <> 0, Acumulado + nPorcentaje2, Acumulado)
                Acumulado = IIf(txtImpuesto3.Caption <> 0, Acumulado + nPorcentaje3, Acumulado)
                Acumulado = (Acumulado / 100)
                nImpuesto1 = IIf(txtImpuesto1.Caption <> 0, nPVenta * nPorcentaje1 / 100, 0)
                nImpuesto2 = IIf(txtImpuesto2.Caption <> 0, nPVenta * nPorcentaje2 / 100, 0)
                nImpuesto3 = IIf(txtImpuesto3.Caption <> 0, nPVenta * nPorcentaje3 / 100, 0)
        
        Case Else 'Peru, Ecuador
                Acumulado = IIf(txtImpuesto1.Caption <> 0, Acumulado + nPorcentaje1, Acumulado)
                Acumulado = IIf(txtImpuesto2.Caption <> 0, Acumulado + nPorcentaje2, Acumulado)
                Acumulado = IIf(txtImpuesto3.Caption <> 0, Acumulado + nPorcentaje3, Acumulado)
                Acumulado = 1 + (Acumulado / 100)
                nImpuesto1 = IIf(txtImpuesto1.Caption <> 0, nPVenta / Acumulado * nPorcentaje1 / 100, 0)
                nImpuesto2 = IIf(txtImpuesto2.Caption <> 0, nPVenta / Acumulado * nPorcentaje2 / 100, 0)
                nImpuesto3 = IIf(txtImpuesto3.Caption <> 0, nPVenta / Acumulado * nPorcentaje3 / 100, 0)
      
      End Select
      nPBase = nPVenta - nImpuesto1 - nImpuesto2 - nImpuesto3
      txtNeto.Caption = Format(nPBase, "###,###,###,##0.00")
      txtImpuesto1.Caption = Format(nImpuesto1, "###,###,###,##0.00")
      txtImpuesto2.Caption = Format(nImpuesto2, "###,###,###,##0.00")
      txtImpuesto3.Caption = Format(nImpuesto3, "###,###,###,##0.00")
      txtOficial.Caption = Format(nOficial, "###,###,##0.00")
      txtPVenta.Caption = Format(nPVenta, "###,###,##0.00")
      txtVenta.Caption = Format((nPVenta * nCantidad), "###,###,###,##0.00")
   End If
   txtBarra.SetFocus
End Sub

Private Sub cmdProducto_Click(Index As Integer)
   txtBarra.SetFocus
      
   RsProducto.MoveFirst
   RsProducto.Find "nbotonRapido = " & Trim(str(Index))
   sProducto = RsProducto!codigo
   
   If validadIngresoProducto(sProducto) = False Then
      Exit Sub
   End If
    
   If wAgregaCombo Then
      nCCombo = Calcular("select sum(nCantidad) as Codigo " & _
                        "FROM " & sComboDetalle & " WHERE tItem='" & sitem & "'", Cn)

      If nCCombo < nCombo * RsDetalle!nCantidad Then
         InsertaCombo sProducto
      Else
         MsgBox "La cantidad máxima de items para este producto es de " & nCombo * RsDetalle!nCantidad, vbExclamation, sMensaje
      End If
   Else
        If lBal And RsProducto!lBalanza Then
           Dim nResultado As Double
           nResultado = Pesar(nBalanzaPuerto)
           nResultado = Format(nResultado, "#,##0.00")
           If nResultado > 0 Then
              InsertaProducto nResultado
           End If
        Else
        nCantidad = 1
           InsertaProducto 1
        End If
     
      If IIf(IsNull(RsProducto!lPropiedad), False, RsProducto!lPropiedad) Then
         lPropiedad = True
      End If
   End If
End Sub

Private Sub cmdProductoCombo_Click(Index As Integer)
    Dim xFiltro As String
    xFiltro = RsProducto.Filter
    RsProducto.Filter = adFilterNone
    txtBarra.SetFocus
    RsProducto.MoveFirst
    RsProducto.Find "tResumido = '" & cmdProductoCombo(Index).Caption & "'"
    sProducto = RsProducto!codigo
    
    nCCombo = Calcular("select sum(nCantidad) as Codigo " & _
                       "FROM " & sComboDetalle & " WHERE tItem='" & sitem & "'", Cn)
 
    If nCCombo < nCombo * RsDetalle!nCantidad Then
        'Oscar Ortega----------------------------------------------
        Dim oRsProductoDeCombo As Recordset
        Set oRsProductoDeCombo = Obtener_ProductoDeCombo(RsDetalle!tCodigoProducto, sProducto)
        If oRsProductoDeCombo.RecordCount > 0 Then
            If IIf(IsNull(oRsProductoDeCombo!lUnico), False, oRsProductoDeCombo!lUnico) Then
                'Obtener Suma de cantidades
                Dim nCantidadEnElCombo As Integer
                nCantidadEnElCombo = ObtenerSumaCantidadesEnElCombo(sitem, oRsProductoDeCombo!tEtiqueta)
                'Suma de cantidades < que nCantidad
                If nCantidadEnElCombo < nCantidad Then
                    InsertaCombo sProducto
                Else
                    MsgBox "Solo es permitido " & nCantidad & " elemento(s) de tipo " & oRsProductoDeCombo!tEtiqueta, vbExclamation, sMensaje
                End If
            Else
                InsertaCombo sProducto
            End If
        Else
            InsertaCombo sProducto
        End If
       '----------------------------------------------------------
        'InsertaCombo sProducto
    Else
       MsgBox "La cantidad máxima de items para este producto es de " & nCombo * RsDetalle!nCantidad, vbExclamation, sMensaje
    End If
    RsProducto.Filter = xFiltro
End Sub

Private Sub cmdSinBoton_Click()
    sTemp = ""
    If Calcular("Select count(*) as Codigo from tclienteproducto where tcodigodelivery='" & sCliente & "' ", Cn) = 0 Then
        Isql = "select * from vProducto where lActivo = 1 and " & IIf(sTipoPedido = "01", "lLocal=1", IIf(sTipoPedido = "02", "lDelivery=1", "lLlevar=1")) & " Order by Descripcion"
        frmBusquedaRapida.cmdOpcion(4).Visible = False
    Else
        Isql = "select vProducto.Grupo, vProducto.Descripcion , tclienteproducto.nprecio As nPrecioVenta , vProducto.nBoton, vProducto.SubGrupo, vProducto.Codigo from vProducto inner join tclienteproducto on vProducto.codigo=tclienteproducto.tcodigoproducto where vProducto.lActivo = 1 and (vProducto.tUnidadNegocio='" & sUnidadNegocio & "' ) Order by vProducto.Descripcion"
        frmBusquedaRapida.cmdOpcion(4).Visible = True
        frmBusquedaRapida.cmdOpcion(4).FontBold = True
    End If
    
    'Isql = "select * from vProducto where lActivo = 1 and " & IIf(sTipoPedido = "01", "lLocal=1", IIf(sTipoPedido = "02", "lDelivery=1", "lLlevar=1")) & " Order by Descripcion"
    Call ConfGrilla(5, frmBusquedaRapida.grdGrilla, "Grupo", 2, "Grupo", 1600, 0, 0, "", _
                                                    "Producto", 2, "Descripcion", 3600, 0, 0, "", _
                                                    "Precio", 2, "nPrecioVenta", 1000, 1, 0, "###,##0.00", _
                                                    "Bot", 2, "nBoton", 500, 1, 0, "", _
                                                    "SubGrupo", 2, "SubGrupo", 1500, 0, 0, "")
    frmBusquedaRapida.nPredeterm = 1
    frmBusquedaRapida.Show vbModal
    
    If wEnter Then
       sProducto = sCodigo
        'INSUMOCRITICO23
        If validadIngresoProducto(sProducto) = False Then
            Exit Sub
        End If
        'INSUMOCRITICO23
 
       Dim xxx As String
       xxx = RsProducto.Filter
       RsProducto.Filter = adFilterNone
       RsProducto.MoveFirst
       RsProducto.Find ("Codigo='" & sProducto & "'")
    
       If Not RsProducto.EOF() Then
          If wAgregaCombo Then
             nCCombo = Calcular("select sum(nCantidad) as Codigo " & _
                               "FROM " & sComboDetalle & " WHERE tItem='" & sitem & "'", Cn)
        
             If nCCombo < nCombo * RsDetalle!nCantidad Then
                InsertaCombo sProducto
             Else
                MsgBox "La cantidad máxima de items para este producto es de " & nCombo * RsDetalle!nCantidad, vbExclamation, sMensaje
             End If
          Else
                          
             If lBal And RsProducto!lBalanza Then
                Dim nResultado As Double
                nResultado = Pesar(nBalanzaPuerto)
                nResultado = Format(nResultado, "#,##0.00")
                If nResultado > 0 Then
                   InsertaProducto nResultado
                End If
             Else
                nCantidad = 1
                InsertaProducto 1
             End If
             
             If IIf(IsNull(RsProducto!lPropiedad), False, RsProducto!lPropiedad) Then
                lPropiedad = True
             End If
          End If
       End If
       RsProducto.Filter = IIf(xxx = "0", "", xxx)
    End If
    txtBarra.SetFocus

End Sub

Private Sub cmdTipoDocumento_Click(Index As Integer)
   sCliente = ""
   txtTipoDocumento.Caption = cmdTipoDocumento(Index).Caption
   txtBarra.SetFocus
   
   nTotalPR = txtMonto.Caption

   If lObservacion And Trim(txtObservacion.Caption) = "" Then
      MsgBox "Debes ingresar la Observación", vbInformation, sMensaje
      cmdDetalle_Click (6)
      If Trim(txtObservacion.Caption) = "" Then
         Exit Sub
      End If
   End If
   Label2.Caption = muestra
   
   'Clave de Multi Cajero
    tUsuActua = sUsuario
    If lMultiCajero Then
      If Supervisor("16") = False Then
         MsgBox "Clave no permitida", vbExclamation, sMensaje
         Exit Sub
      End If
      sUsuario = sVar1
    End If
    sUsuario = tUsuActua
    'Chequea Descuento
    nTotalDescuento = CDbl(Calcular("select sum(nDescuento*nCantidad) as Codigo from " & sDetalle, Cn))
    If nTotalDescuento > 0 Then
       Dim nTope As Double
       Dim nTotalMes As Double
       Dim nConsumo As Double
       Dim aplicaTope As Boolean
       
       lAplicablePedido = Calcular("select lAplicablePedido as Codigo FROM vMotivoDescuento where lActivo=1 and Codigo='" & sCodigoDescuento & "'", Cn)
       nTope = Calcular("select nTope as Codigo from vMotivoDescuento where lActivo=1 and Codigo='" & sCodigoDescuento & "'", Cn)
       
         
       If nTope > 0 Then
          If Calcular("select lTopePedido as Codigo from vMotivoDescuento where lActivo=1 and Codigo='" & sCodigoDescuento & "'", Cn) Then
             If nTotalDescuento > nTope Then
                If MsgBox("El Descuento a aplicar Supera El Tope Registrado por Pedido" & Chr(13) & "¿Desea aplicar el Tope de " & sMonN & " " & nTope & "?", vbQuestion + vbYesNo) = vbYes Then
                   CalculaAplicaTope (nTope)
                Else
                   Exit Sub
                End If
             End If
          Else
             nTotalMes = Calcular("select sum(DPEDIDO.nDescuento*nCantidad) as Codigo FROM dbo.MPEDIDO INNER JOIN dbo.DPEDIDO ON dbo.MPEDIDO.tCodigoPedido = dbo.DPEDIDO.tCodigoPedido " & _
                                  "WHERE month(MPEDIDO.fFecha) = month(getdate()) and year(MPEDIDO.fFecha)=year(getdate()) and mPedido.tDescuento='" & sCodigoDescuento & "' and tEstadoPedido<>'01' and tEstadoPedido<>'03'", Cn)
             
             If nTotalDescuento + nTotalMes > nTope Then
                If nTotalDescuento < nTope Then
                   If MsgBox("El Descuento a aplicar Supera El Tope Registrado dentro de un mes" & Chr(13) & "¿Desea aplicar el saldo " & sMonN & " " & nTope - nTotalMes & "?", vbQuestion + vbYesNo) = vbYes Then
                      CalculaAplicaTope (nTope - nTotalMes)
                   Else
                      Exit Sub
                   End If
                Else
                    MsgBox "El Descuento a aplicar Supera El Tope Registrado dentro de un mes", vbExclamation
                    Exit Sub
                End If
             End If
          End If
       End If
    End If
    sCodigoDescuento = IIf(lAplicablePedido, "", sCodigoDescuento)
    nMonto = Calcular("select sum(nventa) as codigo from " & sDetalle & "", Cn)
    VisualizaMonto
    variableEmite = False
   
   'VALIDACION CANAL DE VENTA
   Dim rsCanalVentas As Recordset
   Dim lObligaMozo As Boolean
   Dim lObligaMotorizado As Boolean
   Dim lObligaClienteFrecuente As Boolean
   Dim lObligaFechaEntrega As Boolean
   Dim lObligaEntregarA As Boolean
   
   Set rsCanalVentas = Lib.OpenRecordset("select * from vTipoPedido", Cn)
   rsCanalVentas.Filter = "Codigo = '" & sTipoPedido & "'"
   
   lObligaMozo = IIf(IsNull(rsCanalVentas!lObligaMozo), False, rsCanalVentas!lObligaMozo)
   lObligaMotorizado = IIf(IsNull(rsCanalVentas!lObligaMotorizado), False, rsCanalVentas!lObligaMotorizado)
   lObligaClienteFrecuente = IIf(IsNull(rsCanalVentas!lObligaClienteFrecuente), False, rsCanalVentas!lObligaClienteFrecuente)
   lObligaFechaEntrega = IIf(IsNull(rsCanalVentas!lObligaIngresoFechaEntrega), False, rsCanalVentas!lObligaIngresoFechaEntrega)
   lObligaEntregarA = IIf(IsNull(rsCanalVentas!lObligaEntregarA), False, rsCanalVentas!lObligaEntregarA)
   
               'Obligatoriedad de Mozo
               If lObligaMozo Then
                  If lMCPV Then
                      sMozo = ObtenerCodigoMozo(sVar1)
                      If sMozo = "" Then
                         Exit Sub
                      End If
                  Else
                      If sMozo = "" Or sMozo = "0000" Then
                         MsgBox "Asigne al Mesero", vbExclamation, sMensaje
                         Exit Sub
                      End If
                  End If
               End If
               
               'Obligatoriedad de Motorizado
'               If lObligaMotorizado Then
'                  If sMotorizado = "" Or sMotorizado = "0000" Then
'                     MsgBox "Asigne al Motorizado", vbExclamation, sMensaje
'                     Exit Sub
'                  End If
'               End If
               
               'Obligatoriedad de Mesa
'               If lObligaMesa And sMesa = "" And txtObservacion.Caption = "" Then
'                  MsgBox "Asigne una Mesa", vbExclamation, sMensaje
'                  cmdCabecera_Click (13)
'                  Exit Sub
'               End If
               
               'Obligatoriedad de Cliente Frecuente
               If sClienteFrecuente = "" And lObligaClienteFrecuente Then
                  MsgBox "Asigne el Cliente Delivery", vbExclamation, sMensaje
                  cmdOpcion_Click (9)
                  Exit Sub
               End If
               
               'Obligatoriedad de Fecha de Entrega
               If Me.txtFechaEntrega.Caption = "" And lObligaFechaEntrega Then
                  MsgBox "Asigne la Fecha de Entrega", vbExclamation, sMensaje
                  cmdCabecera_Click (6)
                  Exit Sub
               End If
               
               'Entregar A
               If lObligaEntregarA = True And Me.txtEntregar.Caption = "" Then
                  MsgBox "Asigne información en Entregar A", vbExclamation, sMensaje
                  cmdDetalle_Click (14)
                  Exit Sub
               End If

   Call Facturar
   
   If wEnter = True Then
        variableEmite = False
        
        Inicializar
        sPedido = ""
        If nPuerto > 0 Then
           Visor String(Int((19 - Len(tMensaje1)) / 2), " ") & tMensaje1, String(Int((19 - Len(tMensaje2)) / 2), " ") & tMensaje2, nPuerto, "N"
        End If
   End If
   
End Sub

Private Sub Form_Activate()
   If txtBarra.Enabled = True Then
      txtBarra.SetFocus
   End If
End Sub
Private Function validarOrigenVentas()
' origen de ventas
    'lActivaMozo = IIf(IsNull(RsCanalesVenta!lActivaMozo), False, RsCanalesVenta!lActivaMozo)
    lActivaMotorizado = IIf(IsNull(RsCanalesVenta!lActivaMotorizado), False, RsCanalesVenta!lActivaMotorizado)
    lCanalDelivery = IIf(IsNull(RsCanalesVenta!lCanalDelivery), False, RsCanalesVenta!lCanalDelivery)
    lCanalCentralPedidos = IIf(IsNull(RsCanalesVenta!lCanalCentralPedidos), False, RsCanalesVenta!lCanalCentralPedidos)
    'entregarA
    'lObligaEntregarA = IIf(IsNull(RsCanalesVenta!lObligaEntregarA), False, RsCanalesVenta!lObligaEntregarA)
    
    'origen de ventas
     lOrigenVentas = IIf(IsNull(RsCanalesVenta!lCanalDelivery), False, RsCanalesVenta!lCanalDelivery)
    
    If lMCPV Then
        'lObligaMozo = False
        'lActivaMozo = False
    Else
        'lObligaMozo = IIf(IsNull(RsCanalesVenta!lObligaMozo), False, RsCanalesVenta!lObligaMozo)
    End If
    lObligaMotorizado = IIf(IsNull(RsCanalesVenta!lObligaMotorizado), False, RsCanalesVenta!lObligaMotorizado)
    'lObligaMesa = IIf(IsNull(RsCanalesVenta!lObligaMesa), False, RsCanalesVenta!lObligaMesa)
    'lObligaPax = IIf(IsNull(RsCanalesVenta!lObligaPax), False, RsCanalesVenta!lObligaPax)
    'lObligaFechaEntrega = IIf(IsNull(RsCanalesVenta!lObligaIngresoFechaEntrega), False, RsCanalesVenta!lObligaIngresoFechaEntrega)
   ' lObligaClienteFrecuente = IIf(IsNull(RsCanalesVenta!lObligaClienteFrecuente), False, RsCanalesVenta!lObligaClienteFrecuente)
    
    
    'If lActivaMozo Then
        'fraMozo.Visible = True
    'Else
      '  fraMozo.Visible = False
   ' End If
                    
    If lActivaMotorizado Then
        Me.fraMorotizado.Visible = True
    Else
        Me.fraMorotizado.Visible = False
    End If
    
    Me.fraOrigenVentas.Visible = False
    
    
    
End Function

Private Sub Form_Load()

       'anulacion por nota de credito
   Isql = "SELECT * FROM TPARAMETRO"
   Set RsTparametro = Lib.OpenRecordset(Isql, Cn)
   '------------------------------------------------------------------------------
      'anulacion de documentos por nota de credito
   
   If RsTparametro!lanula = True Then
    cmdNotasCredito.Item(50).Visible = True
        Else
            cmdNotasCredito.Item(50).Visible = False
            
   End If
   '--------------------------------------------
    'origen de ventas
    
    Me.fraOrigenVentas.Visible = False
    Me.fraMorotizado.Visible = False
    
    '--------------------------------------------


    'InsumosCriticos ' 23 2013
   sInsumoCombo = dbTemporal(sCaja, 2, "nCantidad", "float", _
                                    "tCodigoInsumo", "nVarChar(20)")
   
   sDetalle = dbTemporal(sCaja, 45, "tCodigoPedido", "nVarChar(10)", _
                                    "tItem", "nVarChar(3)", _
                                    "tTipoPedido", "nVarChar(2)", _
                                    "tCodigoProducto", "nVarChar(7)", _
                                    "tCodigoGrupo", "nVarChar(2)", _
                                    "tCodigoSubGrupo", "nVarChar(4)", _
                                    "tMoneda", "nVarChar(3)", _
                                    "nPrecioNeto", "Float", _
                                    "nPrecioImpuesto1", "Float", "nPrecioImpuesto2", "Float", "nPrecioImpuesto3", "Float", _
                                    "nPrecioVenta", "Float", _
                                    "nRecargo", "Float", "nDescuento", "Float", _
                                    "nPrecioOficial", "Float", _
                                    "nCantidad", "Float", _
                                    "nImpuesto1", "Float", "nImpuesto2", "Float", "nImpuesto3", "Float", _
                                      "nVenta", "Float", _
                                    "tObservacion", "nVarChar(255)", _
                                    "tCortesia", "nVarChar(4)", _
                                    "lImprime", "Bit", _
                                    "tEstadoItem", "nVarChar(3)", _
                                    "tArea", "nVarChar(3)", _
                                    "lCombinacion", "Bit", "nCombinacion", "Smallint", "lImprimeArea", "Bit", _
                                    "tFacturado", "nVarChar(1)", "tDocumento", "nVarChar(20)", "lTransferido", "Bit", "tComanda", "nVarChar(10)", _
                                    "nInsumo", "Float", "nGasto", "Float", "nManoObra", "Float", "nOrden", "int", "lCorte", "bit", "Estado", "nVarChar(1)", "toferta", "nvarchar(5)", "tautorizaoferta", "nvarchar(15)", "tSubAlmacen", "nvarchar(6)", "tCodigoEtiqueta", "nvarchar(50)", "fenvio", "datetime", "nenvio", "int", "tCajaD", "nvarchar(3)")
      
   Centrar Me
   Dim sTemp1 As String
   Dim sTemp2 As String
   Dim sTemp3 As String
      
   muestra = Label2.Caption
'   If cmdCabecera(0).Visible = False Then
'        Label2.Left = 3200
'        Label2.Width = 1980
'   Else
'        Label2.Left = 4200
'        Label2.Width = 1005
'   End If
      
   cmdCabecera(1).Caption = sBoton1
   cmdCabecera(2).Caption = sBoton2
   cmdCabecera(3).Caption = sBoton3
   cmdCabecera(4).Caption = sBoton4
   cmdCabecera(5).Caption = sBoton5
    
    If sBoton1 = "" Then
        cmdCabecera(1).Enabled = False
        cmdCabecera(1).Caption = "N/D"
    End If
    If sBoton2 = "" Then
        cmdCabecera(2).Enabled = False
        cmdCabecera(2).Caption = "N/D"
    End If
    If sBoton3 = "" Then
        cmdCabecera(3).Enabled = False
        cmdCabecera(3).Caption = "N/D"
    End If
    If sBoton4 = "" Then
        cmdCabecera(4).Enabled = False
        cmdCabecera(4).Caption = "N/D"
    End If
    If sBoton5 = "" Then
        cmdCabecera(5).Enabled = False
        cmdCabecera(5).Caption = "N/D"
    End If
   'cmdCabecera(1).FontBold = True
   'sTipoPedido = "01"
   
   Select Case sTipoPedidoPD
        Case Is = "01"
            sTipoPedido = "01"
            If sBoton1 = "" Then
                cmdCabecera(1).Enabled = False
                cmdCabecera(1).Caption = "N/D"
            Else
                cmdCabecera(1).FontBold = True
            End If

        Case Is = "02"
            sTipoPedido = "02"
            If sBoton2 = "" Then
                cmdCabecera(2).Enabled = False
                cmdCabecera(2).Caption = "N/D"
            Else
                cmdCabecera(2).FontBold = True
            End If
        Case Is = "03"
            sTipoPedido = "03"
            If sBoton3 = "" Then
                cmdCabecera(3).Enabled = False
                cmdCabecera(3).Caption = "N/D"
            Else
                cmdCabecera(3).FontBold = True
            End If
        Case Is = "04"
            sTipoPedido = "04"
            If sBoton4 = "" Then
                cmdCabecera(4).Enabled = False
                cmdCabecera(4).Caption = "N/D"
            Else
                cmdCabecera(4).FontBold = True
            End If
        Case Is = "05"
            sTipoPedido = "05"
            If sBoton5 = "" Then
                cmdCabecera(5).Enabled = False
                cmdCabecera(5).Caption = "N/D"
            Else
                cmdCabecera(5).FontBold = True
            End If
        Case Else
   End Select
   
   
   
   nOperadorPropiedad = 0
            
   sTemp1 = Calcular("select tDetallado as Codigo from TTABLA where tTabla='ETIQUETA' and tCodigo='01'", Cn)
   sTemp2 = Calcular("select tDetallado as Codigo from TTABLA where tTabla='ETIQUETA' and tCodigo='02'", Cn)
   sTemp3 = Calcular("select tDetallado as Codigo from TTABLA where tTabla='ETIQUETA' and tCodigo='03'", Cn)
   cmdEtiqueta(1).Caption = IIf(sTemp1 = "0", "", sTemp1)
   cmdEtiqueta(2).Caption = IIf(sTemp2 = "0", "", sTemp2)
   cmdEtiqueta(3).Caption = IIf(sTemp3 = "0", "", sTemp3)
   sMozo = ""
                
'  If lBal Then
'    With frmMsComm.MSCommBalanza
'         If .PortOpen Then
'            .PortOpen = False
'         End If
'
'          .CommPort = nBalanzaPuerto
'          .Settings = nBalanzaBS & "," & nBalanzaParidad & "," & nBalanzaBD & "," & nBalanzaBP '"9600,n,8,1"'"9600,n,8,1"
'          .InBufferSize = 1024
'          .OutBufferSize = 512
'          .RThreshold = 15
'          .SThreshold = 1
'          .InputLen = 15
'          .InputMode = comInputModeText
'          .RTSEnable = True
'          .PortOpen = True
'     End With
'   End If
                
   'Operador
   Isql = "select * from vOperador where lActivo = 1 order by Codigo"
   Set RsOperador = Lib.OpenRecordset(Isql, Cn)
      
   'Propiedades
   sProductoPropiedad = dbTemporal(sCaja, 11, "tItem", "nVarChar(3)", _
                                             "tCodigoPropiedad", "nVarChar(4)", _
                                             "tProducto", "nVarChar(7)", _
                                             "tEnlace", "nVarChar(7)", _
                                             "nInsumo", "float", _
                                             "nGasto", "float", _
                                             "nManoObra", "float", _
                                             "nCantidad", "float", _
                                             "nInsumoUnitario", "float", _
                                             "nGastounitario", "float", _
                                             "nManoObraUnitario", "float")
   Dim xSql As String
   If lAlmacen Then
      Dim RsOp As Recordset
      Set RsOp = Lib.OpenRecordset("select Codigo, Descripcion from vOperador where lStockMenos=1", Cn)
      If RsOp.RecordCount > 0 Then
         xSql = "select tCodigoPropiedad as Codigo, TPROPIEDAD.tDetallado as Descripcion, tProducto, TOPERADOR.tOperador as tOperador, TOPERADOR.tDetallado as Operador, nPrecio, tEnlace, " & _
                "nInsumo, nGasto, nManoObra, ISNULL(tpropiedad.lsolicitacantidad,0) lsolicitacantidad  " & _
                "FROM dbo.TPROPIEDAD LEFT OUTER JOIN dbo.TOPERADOR ON dbo.TPROPIEDAD.tOperador = dbo.TOPERADOR.tOperador " & _
                "Where TPROPIEDAD.lActivo = 1 And IsNull(TOPERADOR.lStockMenos, 0) <> 1 union " & _
                "select '9999' as Codigo, tDetallado as Descripcion, tCodigoPlato as tProducto, '" & RsOp!codigo & "' as tOperador, '" & RsOp!Descripcion & "' as Operador, 0, " & sAlmacenMDB & ".dbo.DRECETAVENTA.tCodigoProducto as tEnlace, nCantidad * nPrecio as nInsumo, 0, 0,0 " & _
                "FROM " & sAlmacenMDB & ".dbo.DRECETAVENTA INNER JOIN " & sAlmacenMDB & ".dbo.MRECETAVENTA ON " & sAlmacenMDB & ".dbo.DRECETAVENTA.tLocal = " & sAlmacenMDB & ".dbo.MRECETAVENTA.tLocal AND " & sAlmacenMDB & ".dbo.DRECETAVENTA.tRecetaVenta = " & sAlmacenMDB & ".dbo.MRECETAVENTA.tRecetaVenta INNER JOIN " & sAlmacenMDB & ".dbo.TPRODUCTO ON " & sAlmacenMDB & ".dbo.DRECETAVENTA.tCodigoProducto = " & sAlmacenMDB & ".dbo.TPRODUCTO.tCodigoProducto " & _
                "Where lNoDescargo = 1"
      Else
         xSql = "select tCodigoPropiedad as Codigo, TPROPIEDAD.tDetallado as Descripcion, tProducto, TPROPIEDAD.tOperador, nPrecio, tEnlace, " & _
                "nInsumo, nGasto, nManoObra, toperador.tDetallado AS Operador, ISNULL(tpropiedad.lsolicitacantidad,0) lsolicitacantidad  " & _
                "FROM dbo.TPROPIEDAD LEFT OUTER JOIN dbo.TOPERADOR ON dbo.TPROPIEDAD.tOperador = dbo.TOPERADOR.tOperador " & _
                "Where TPROPIEDAD.lActivo = 1 And IsNull(TOPERADOR.lStockMenos, 0) <> 1"
      End If
   Else
         xSql = "select tCodigoPropiedad as Codigo, TPROPIEDAD.tDetallado as Descripcion, tProducto, TPROPIEDAD.tOperador, nPrecio, tEnlace, " & _
                "nInsumo, nGasto, nManoObra, toperador.tDetallado AS Operador, ISNULL(tpropiedad.lsolicitacantidad,0) lsolicitacantidad  " & _
                "FROM dbo.TPROPIEDAD LEFT OUTER JOIN dbo.TOPERADOR ON dbo.TPROPIEDAD.tOperador = dbo.TOPERADOR.tOperador " & _
                "Where TPROPIEDAD.lActivo = 1 And IsNull(TOPERADOR.lStockMenos, 0) <> 1"
   End If
   Set RsPropiedad = Lib.OpenRecordset(xSql, Cn)
                      
   Isql = "SELECT [" & sDetalle & "].*, dbo.TPRODUCTO.tDetallado AS Producto, dbo.vCortesia.Descripcion AS Cortesia, dbo.TPRODUCTO.lDescuento AS lDescuento, CASE [" & sDetalle & "].nDescuento WHEN 0 THEN 0 ELSE [" & sDetalle & "].nDescuento * 100 / [" & sDetalle & "].nPrecioOficial END AS Descuento, " & _
          "dbo.TPRODUCTO.lModificable AS lModificable, CONVERT(bit, ISNULL(DATALENGTH([" & sDetalle & "].tObservacion), 0)) AS lObservacion, ISNULL(T1.nPropiedad, 0) AS lPropiedad " & _
          "FROM [" & sDetalle & "] LEFT OUTER JOIN (SELECT tItem, CASE WHEN COUNT(tProducto) > 0 THEN 1 ELSE 0 END AS nPropiedad FROM [" & sProductoPropiedad & "] Group by tItem) T1 " & _
          "ON [" & sDetalle & "].tItem = T1.tItem LEFT OUTER JOIN dbo.vCortesia ON [" & sDetalle & "].tCortesia = dbo.vCortesia.Codigo LEFT OUTER JOIN " & _
          "dbo.TPRODUCTO ON [" & sDetalle & "].tCodigoProducto = dbo.TPRODUCTO.tCodigoProducto " & _
          "ORDER BY [" & sDetalle & "].tItem"
   Set RsDetalle = Lib.OpenRecordset(Isql, Cn)
            
   'Producto Propiedad
   Isql = "SELECT " & sProductoPropiedad & ".tItem, " & sProductoPropiedad & ".tCodigoPropiedad, " & sProductoPropiedad & ".tProducto, t1.Descripcion AS Descripcion, Operador, " & sProductoPropiedad & ".ncantidad  " & _
          "FROM " & sProductoPropiedad & " INNER JOIN (" & xSql & ") T1 ON " & sProductoPropiedad & ".tCodigoPropiedad = T1.Codigo and " & sProductoPropiedad & ".tProducto = T1.tProducto and " & sProductoPropiedad & ".tenlace= t1.tenlace "
   Set RsProductoPropiedad = Lib.OpenRecordset(Isql, Cn)
            
   'Combo Propiedad
    sComboPropiedad = dbTemporal(sCaja, 12, "tItem", "nVarChar(3)", _
                                          "tItemCombo", "nVarChar(3)", _
                                          "tCodigoPropiedad", "nVarChar(4)", _
                                          "tProducto", "nVarChar(7)", _
                                          "tEnlace", "nVarChar(7)", _
                                          "nInsumo", "float", _
                                          "nGasto", "float", _
                                          "nManoObra", "float", _
                                          "nCantidad", "float", _
                                          "nInsumoUnitario", "float", _
                                          "nGastoUnitario", "float", _
                                          "nManoObraUnitario", "float")

   Isql = "SELECT " & sComboPropiedad & ".tItem, " & sComboPropiedad & ".tItemCombo, T1.Descripcion, T1.Operador , " & sComboPropiedad & ".ncantidad " & _
          "FROM " & sComboPropiedad & " INNER JOIN (" & xSql & ") T1 ON " & sComboPropiedad & ".tCodigoPropiedad = T1.Codigo AND " & sComboPropiedad & ".tProducto = T1.tProducto AND " & sComboPropiedad & ".tEnlace = T1.tEnlace "
   Set RsComboPropiedad = Lib.OpenRecordset(Isql, Cn)
         
   'Combos
   Isql = "SELECT dbo.TCOMBO.tCombo, dbo.TCOMBO.tCodigoProducto AS Codigo, dbo.TPRODUCTO.tResumido AS Descripcion ,ISNULL(TCOMBO.NVALOR,-2147483633) NVALOR " & _
          "FROM dbo.TCOMBO INNER JOIN dbo.TPRODUCTO ON dbo.TCOMBO.tCodigoProducto = dbo.TPRODUCTO.tCodigoProducto " & _
          "where lActivo=1 ORDER BY TCOMBO.TETIQUETA,dbo.TPRODUCTO.tResumido  "
   Set RsProductoCombo = Lib.OpenRecordset(Isql, Cn)
         
   'Producto
   Isql = "select * from vProducto where lActivo = 1 Order by nBoton"
   Set RsProducto = Lib.OpenRecordset(Isql, Cn)

   'Areas
   Set RsArea = Lib.OpenRecordset("select * from vAreaImpresora where tCaja ='" & sCaja & "'", Cn)
                                    
   'Mozos
   Isql = "select * from vMozo where substring(Codigo,1,1)<>'*' AND lActivo = 1 Order by nBoton"
   Set RsMozo = Lib.OpenRecordset(Isql, Cn)
   AsignaBoton 19, RsMozo, cmdMozo()
   
   
      'Origen de ventas
   Isql = "select * from vOrigenVenta where Activo = 1 and Visible = 1 Order by Boton"
   Set RsOrigenVentas = Lib.OpenRecordset(Isql, Cn)
   
   Isql = "select * from vTipoPedido where Codigo = '02'"
   Set RscanalOrigenVentas = Lib.OpenRecordset(Isql, Cn)
   
   AsignaBotonOrigenVentas 19, RsOrigenVentas, Me.cmdOrigen()
   Set RsCanalesVenta = Lib.OpenRecordset("select * from TCANALVENTA", Cn)
   
   'Motorizado
   Isql = "select * from vMotorizado where lActivo = 1 Order by nBoton"
   Set RsMotorizado = Lib.OpenRecordset(Isql, Cn)
   AsignaBoton 19, RsMotorizado, cmdMotorizado()
   '-----------------------------------------------------
   
   
   'Motivo de Eliminacion
   Isql = "select * from vMotivoEliminacion where lActivo = 1 order by Codigo"
   Set RsMotivoEliminacion = Lib.OpenRecordset(Isql, Cn)
   AsignaComando 38, RsMotivoEliminacion, cmdEliminacion()
   
   'Tipo de Documentos
'   If pais = "002" Then 'Ecuador
'      Set RsTipoDocumento = Lib.OpenRecordset("select * from vTipoDocumentoImpresora where tCaja ='" & sCaja & "' and Transporte=0 and isnull(tNumeroAutorizacion,'')<>'' order by tTipoEmision", Cn)
'   Else
'      Set RsTipoDocumento = Lib.OpenRecordset("select * from vTipoDocumentoImpresora where tCaja ='" & sCaja & "' and Transporte=0 order by tTipoEmision", Cn)
'   End If

   If pais = "002" Then 'Ecuador
      Set RsTipoDocumento = Lib.OpenRecordset("select * from vTipoDocumentoImpresora where tCaja ='" & sCaja & "' and Transporte=0 and isnull(tNumeroAutorizacion,'')<>'' And lNotaCredito = 0 And lActivo = 1 UNION Select * From vTipoDocumentoImpresora where tCaja ='" & sCaja & "' and Transporte=0 And lNotaCredito = 0 And lFacturacionElectronica=1 and lActivo =1 order by tTipoEmision", Cn)
   Else
      Set RsTipoDocumento = Lib.OpenRecordset("select * from vTipoDocumentoImpresora where tCaja ='" & sCaja & "' and Transporte=0 And lNotaCredito = 0 And lActivo = 1 order by tTipoEmision", Cn)
   End If
   
   'LlenaProducto "01"
   Call ConfGrilla(9, grdDetalle, "Or", 2, "nOrden", 300, 1, 0, "#0", _
                                  "-", 2, "lCorte", 250, 0, 4, "", _
                                  "Producto", 2, "Producto", 1950, 0, 0, "", _
                                  "Precio", 2, "nPrecioVenta", 700, 1, 0, "###,###,##0.00", _
                                  "Cant.", 2, "nCantidad", 650, 1, 0, "#,##0.00", _
                                  "SubTotal", 2, "nVenta", 900, 1, 0, "###,###,##0.00", _
                                  "E", 2, "lImprime", 250, 0, 4, "", _
                                  "P", 2, "lPropiedad", 250, 0, 4, "", _
                                  "O", 2, "lObservacion", 250, 0, 4, "")
   Set grdDetalle.DataSource = RsDetalle
         
   nMonto = 0
   txtMonto.Caption = Format(nMonto, "###,##0.00")
         
   'Asigna Operador
   AsignaBoton 13, RsOperador, cmdOperador()
   If RsOperador.RecordCount > 0 Then
      RsOperador.MoveFirst
      If Not IsNull(RsOperador!nBoton) And RsOperador!nBoton > 0 Then
         cmdOperador_Click (RsOperador!nBoton)
      End If
   End If
         
   With RsTipoDocumento
       If .RecordCount > 0 Then
          .MoveFirst
          For i = 1 To IIf(.RecordCount >= 4, 4, .RecordCount)
              cmdTipoDocumento(i).Visible = True
              cmdTipoDocumento(i).Caption = !Descripcion
              .MoveNext
          Next i

          For i = .RecordCount + 1 To 4
              cmdTipoDocumento(i).Visible = False
          Next i
       Else
          For i = 1 To 4
              cmdTipoDocumento(i).Visible = False
          Next i
       End If
  End With
  
  If Not lPrinter Then
     cmdOpcion(8).Visible = False
  End If
  
   'Combo
   Call ConfGrilla(7, grdCombo, "-", 2, "lCorte", 250, 0, 4, "", _
                                "Producto", 2, "Producto", 1950, 0, 0, "", _
                                "Cant.", 2, "nCantidad", 650, 1, 0, "#,##0.00", _
                                "E", 2, "lImprime", 250, 0, 4, "", _
                                "P", 2, "lPropiedad", 250, 0, 4, "", _
                                "O", 2, "lObservacion", 250, 0, 4, "", _
                                "Ord", 2, "nOrden", 400, 1, 0, "#0")
      
   sComboDetalle = dbTemporal(sCaja, 24, "tCodigoPedido", "nVarchar(10)", "tItem", "nVarchar(3)", "tItemCombo", "nVarchar(3)", "tProducto", "nVarchar(7)", "tProductoCombo", "nVarchar(7)", "nCantidad", "float", "tCodigoGrupo", "nVarchar(2)", "tCodigoSubGrupo", "nVarchar(4)", _
                                         "nPrecioNeto", "float", "nImpuesto1", "float", "nImpuesto2", "float", "nImpuesto3", "float", "nVenta", "float", "nInsumo", "float", "nGasto", "float", "nManoObra", "float", "lImprimeArea", "bit", "lImprime", "bit", "nOrden", "int", "tObservacion", "nVarchar(250)", "lCorte", "bit", "lAtendidoC", "BIT", "fAtendidoC", "DATETIME", "tUsuarioAtendio", "nvarchar(15)")
   
   Isql = "SELECT dbo." & sComboDetalle & ".tProducto, dbo." & sComboDetalle & ".tItem, dbo." & sComboDetalle & ".tItemCombo, dbo." & sComboDetalle & ".tProductoCombo, dbo." & sComboDetalle & ".nCantidad, dbo." & sComboDetalle & ".tCodigoGrupo, dbo." & sComboDetalle & ".tCodigoSubGrupo, dbo.TPRODUCTO.tDetallado AS Producto, " & _
          "dbo.MPEDIDO.tEstadoPedido, dbo.MPEDIDO.tCaja, dbo." & sComboDetalle & ".lImprimeArea, dbo." & sComboDetalle & ".lImprime, dbo." & sComboDetalle & ".nOrden, CONVERT(bit,ISNULL(DATALENGTH(dbo." & sComboDetalle & ".tObservacion), 0)) AS lObservacion, ISNULL(T1.nPropiedad, 0) AS lPropiedad, dbo." & sComboDetalle & ".tObservacion, dbo." & sComboDetalle & ".lCorte " & _
          "FROM dbo." & sComboDetalle & " LEFT OUTER JOIN (SELECT tItem, tItemCombo, CASE WHEN COUNT(tProducto) > 0 THEN 1 ELSE 0 END AS nPropiedad From " & sComboPropiedad & " " & _
          "GROUP BY tItem, tItemCombo) AS T1 ON dbo." & sComboDetalle & ".tItemCombo = T1.tItemCombo AND dbo." & sComboDetalle & ".tItem = T1.tItem LEFT OUTER JOIN dbo.TPRODUCTO ON dbo." & sComboDetalle & ".tProductoCombo = dbo.TPRODUCTO.tCodigoProducto LEFT OUTER JOIN dbo.MPEDIDO ON dbo." & sComboDetalle & ".tCodigoPedido = dbo.MPEDIDO.tCodigoPedido"
                
   Set RsCombo = Lib.OpenRecordset(Isql, Cn)
   Set grdCombo.DataSource = RsCombo
  
  Impuesto
  
  If sMozo = "" Then
     sMozo = "0000"
     txtTitulo.Caption = " Caja Rápida : " & sCaja & " Mesero: Sin Mesero"
  Else
     cmdDetalle(9).Enabled = False
     txtTitulo.Caption = " Caja Rápida : " & sCaja & " Mesero: " & Calcular("select descripcion as codigo from vMozo where Codigo='" & sMozo & "'", Cn)
  End If
  
  fraMozo.Visible = False
  fraDetalle.Visible = False
  fraPropiedad.Visible = False
  fraEliminacion.Visible = False
  fraPuntoVenta.Visible = False
  fraProductoCombo.Visible = False
  Pedido = ""
  sObser = ""
  txtObservacion.Caption = sObser
  wCombo = False
  wAgregaCombo = False
  nCombo = 0
  sComandaInfhotel = ""
  Sw = False
  lPropiedad = False
  Set RsCajaRapida = Lib.OpenRecordset("select tCodigo, substring(tCodigo,1,1) as Prefijo, tDetallado, nValor from TTABLA where tTabla='CAJARAPIDA'", Cn)

  If lInfhotel Then
     'Moneda Base
     sMonedaBase = Calcular("select tMoneda as Codigo from TPARAMETRO", CnInfhotel)
     
     'Punto de Venta
     Isql = "Select tPuntoVenta as Codigo, tDescripcion as Descripcion, nUltimoComanda, tmoneda" & _
            " From tPuntoVenta " & _
            " where tHotel='" & sHotel & "' AND lActivo=1 and lInforest=1"
     Set rsPuntoVenta = Lib.OpenRecordset(Isql, CnInfhotel)
     AsignaComando 9, rsPuntoVenta, cmdPunto()
     cmdCabecera(0).Visible = True
     
     For i = 1 To 9
         cmdPunto(i).FontBold = False
     Next i
    
     For i = 1 To 9
         If cmdPunto(i).Caption = Calcular("select tDescripcion as codigo from tPuntoventa where tPuntoVenta='" & sPuntoVentaInfhotel & "'", CnInfhotel) Then
           cmdPunto_Click (i)
        End If
    Next i
    sPuntoVenta = sPuntoVentaInfhotel
     
  End If
  
  If lMultiCajero Then
      cmdOpcion(2).Enabled = False
  End If
  
  If lInfhotel Then
      cmdOpcion(14).Visible = True
  End If
  
  cmdEtiqueta_Click (1)
  txtDescuento.Caption = "0.00"
  Screen.MousePointer = vbDefault
  


   
End Sub

Public Sub InsertaProducto(xCantidad As Double)
    Screen.MousePointer = vbHourglass
    Dim nValor As Double
    Dim precioventa As Double
    Dim lImp1 As Boolean
    Dim lImp2 As Boolean
    Dim lImp3 As Boolean
    Dim nOrden As Integer
    Dim RsOrd As Recordset
        
    ' variables MULTIAREAPRODUCCION
    Dim lProductoMultiArea As Boolean
    Dim tsubalmacen As String
    Dim tAreaProduccion As String
    
        'CPvalicacion central d pedido LG
   ' Dim codigoClienteF As String
    Dim lClienteExcluyeProducto As Boolean
    Dim lProductoPermiteDescuento As Boolean
    Dim lClienteControlaProducto As Boolean
  '  codigoClienteF = Calcular("select tclientedelivery as codigo from mpedido where tcodigopedido='" & sPedido & "'", Cn)
    
'    If Calcular("select isnull(treservainf,'') as codigo from mpedido where tcodigopedido='" & sPedido & "'", Cn) <> "" Then
'        MsgBox "Se ha aplicado Anticipo al Pedido!!!, no se puede Ingresar mas productos!!", vbInformation, sMensaje
'         Screen.MousePointer = vbDefault
'        Exit Sub
'    End If
    
    
    If sCliente <> "" Then ' verificamos si tiene cliente frecuente
            If Calcular("select count(tcodigodelivery) as codigo from TCLIENTEPRODUCTO where tcodigoDelivery='" & sCliente & "'", Cn) > 0 Then
                lClienteExcluyeProducto = Calcular("select isnull(lexcluyeproductos,0) as codigo from tdelivery where tcodigodelivery='" & sCliente & "' ", Cn)
                If lClienteExcluyeProducto Then
                    If Calcular("select count(tcodigodelivery) as codigo from tclienteproducto where tcodigodelivery='" & sCliente & "' AND TCODIGOPRODUCTO='" & sProducto & "'  ", Cn) = 0 Then ' si el cliente tiene productos asociados
                                MsgBox "Según Configuración, este Producto no puede seleccionarse para el Cliente Frecuente Indicado"
                                Screen.MousePointer = vbDefault
                                Exit Sub
                    End If
                End If
                
                lClienteControlaProducto = False
                If Calcular("select count(tcodigoproducto) as codigo from tclienteproducto where tcodigodelivery='" & sCliente & "' AND TCODIGOPRODUCTO='" & sProducto & "'", Cn) > 0 Then
                    lClienteControlaProducto = True
                End If
                
            End If
    End If
    'INSUMOCRITICO23
    
    Dim rsInsumo As New ADODB.Recordset
    If Calcular("select isnull(lControlInsumoCritico,0) as codigo from tproducto where tcodigoproducto='" & sProducto & "'", Cn) = True Then
                    Set rsInsumo = Lib.OpenRecordset("select isnull(tcodigoinsumo,'') tcodigoinsumo , isnull(tinsumo.descripcion,'') ,isnull(nstock,0) from tproducto inner join tinsumo on tproducto.tcodigoinsumo =tinsumo.tcodigo where tproducto.tcodigoproducto='" & sProducto & "' and tinsumo.lactivo=1", Cn)
                    If Not (rsInsumo.EOF Or rsInsumo.BOF) Then
                        Label2.Caption = " Insumo Crítico--> " & rsInsumo.Fields(1) & " =  Stock: " & str(rsInsumo.Fields(2)) & "     Solicitado: " + str(xCantidad)
                    End If
    Else
        Label2.Caption = muestra
    End If
    'INSUMOCRITICO
    
    sitem = Lib.Correlativo(Calcular("select max(tItem) as codigo from [" & sDetalle & "]", Cn), 3)
    If RsDetalle.RecordCount = 0 Then
       'sitem = "001"
       nOrden = 1
    Else
       'sitem = Lib.Correlativo(Calcular("select max(tItem) as codigo from [" & sDetalle & "]", Cn), 3)
       If lOrden Then
          Set RsOrd = Lib.OpenRecordset("select nOrden, lImprime from " & sDetalle & " Order by nOrden DESC", Cn)
          If RsOrd.RecordCount > 0 Then
             If IIf(IsNull(RsOrd!lImprime), False, RsOrd!lImprime) Then
                nOrden = RsOrd!nOrden + 1
             Else
                nOrden = RsOrd!nOrden
             End If
          Else
             nOrden = 1
          End If
       Else
          nOrden = RsProducto!nOrden
       End If
       
    End If
                                
    'Precios con Recargos / Descargos por Tipo de Pedido
    nRecargo = 0
    nDescuento = 0
    nValor = 0
    nValor = nValor + IIf(RsProducto!lImpuesto1, nPorcentaje1, 0)
    nValor = nValor + IIf(RsProducto!lImpuesto2, nPorcentaje2, 0)
    nValor = nValor + IIf(RsProducto!lImpuesto3, nPorcentaje3, 0)
            
    lImp1 = RsProducto!lImpuesto1
    lImp2 = RsProducto!lImpuesto2
    lImp3 = RsProducto!lImpuesto3

    If sTipoPedido = "02" Then
       If IsNull(RsProducto!nPrecioDelivery) Or RsProducto!nPrecioDelivery = 0 Then
          nOficial = IIf(IsNull(RsProducto!nprecioVenta), 0, RsProducto!nprecioVenta) + (nDELIVERY * IIf(IsNull(RsProducto!nprecioVenta), 0, RsProducto!nprecioVenta) / 100)
       Else
          nOficial = IIf(IsNull(RsProducto!nPrecioDelivery), 0, RsProducto!nPrecioDelivery)
          nValor = 0
          nValor = nValor + IIf(RsProducto!lImpuesto4, nPorcentaje1, 0)
          nValor = nValor + IIf(RsProducto!lImpuesto5, nPorcentaje2, 0)
          nValor = nValor + IIf(RsProducto!lImpuesto6, nPorcentaje3, 0)
          lImp1 = RsProducto!lImpuesto4
          lImp2 = RsProducto!lImpuesto5
          lImp3 = RsProducto!lImpuesto6
       End If
    ElseIf sTipoPedido = "03" Then
       If IsNull(RsProducto!nPreciollevar) Or RsProducto!nPreciollevar = 0 Then
          nOficial = IIf(IsNull(RsProducto!nprecioVenta), 0, RsProducto!nprecioVenta) + (nLlevar * IIf(IsNull(RsProducto!nprecioVenta), 0, RsProducto!nprecioVenta) / 100)
       Else
          nOficial = IIf(IsNull(RsProducto!nPreciollevar), 0, RsProducto!nPreciollevar)
          nValor = 0
          nValor = nValor + IIf(RsProducto!lImpuesto7, nPorcentaje1, 0)
          nValor = nValor + IIf(RsProducto!lImpuesto8, nPorcentaje2, 0)
          nValor = nValor + IIf(RsProducto!lImpuesto9, nPorcentaje3, 0)
          lImp1 = RsProducto!lImpuesto7
          lImp2 = RsProducto!lImpuesto8
          lImp3 = RsProducto!lImpuesto9
       End If
    ElseIf sTipoPedido = "04" Then
       If IsNull(RsProducto!nPrecioCanal4) Or RsProducto!nPrecioCanal4 = 0 Then
          nOficial = IIf(IsNull(RsProducto!nprecioVenta), 0, RsProducto!nprecioVenta) + (nLlevar * IIf(IsNull(RsProducto!nprecioVenta), 0, RsProducto!nprecioVenta) / 100)
       Else
          nOficial = IIf(IsNull(RsProducto!nPrecioCanal4), 0, RsProducto!nPrecioCanal4)
          nValor = 0
          nValor = nValor + IIf(RsProducto!lImpuesto10, nPorcentaje1, 0)
          nValor = nValor + IIf(RsProducto!lImpuesto11, nPorcentaje2, 0)
          nValor = nValor + IIf(RsProducto!lImpuesto12, nPorcentaje3, 0)
          lImp1 = RsProducto!lImpuesto10
          lImp2 = RsProducto!lImpuesto11
          lImp3 = RsProducto!lImpuesto12
       End If
    ElseIf sTipoPedido = "05" Then
       If IsNull(RsProducto!nPrecioCanal5) Or RsProducto!nPrecioCanal5 = 0 Then
          nOficial = IIf(IsNull(RsProducto!nprecioVenta), 0, RsProducto!nprecioVenta) + (nLlevar * IIf(IsNull(RsProducto!nprecioVenta), 0, RsProducto!nprecioVenta) / 100)
       Else
          nOficial = IIf(IsNull(RsProducto!nPrecioCanal5), 0, RsProducto!nPrecioCanal5)
          nValor = 0
          nValor = nValor + IIf(RsProducto!lImpuesto10, nPorcentaje1, 0)
          nValor = nValor + IIf(RsProducto!lImpuesto11, nPorcentaje2, 0)
          nValor = nValor + IIf(RsProducto!lImpuesto12, nPorcentaje3, 0)
          lImp1 = RsProducto!lImpuesto10
          lImp2 = RsProducto!lImpuesto11
          lImp3 = RsProducto!lImpuesto12
       End If
    Else
       nOficial = IIf(IsNull(RsProducto!nprecioVenta), 0, RsProducto!nprecioVenta)
    End If
    
    nOficial = IIf(RsProducto!tMoneda = "02", nOficial * nTC, nOficial)
    nPVenta = IIf(IsNull(RsProducto!nprecioVenta), 0, RsProducto!nprecioVenta)
    
    'CPvalicacion LG
    If lClienteControlaProducto Then
        nOficial = Calcular("SELECT ISNULL(NPRECIO,0) AS CODIGO FROM TCLIENTEPRODUCTO WHERE TCODIGODELIVERY='" & sCliente & "' AND TCODIGOPRODUCTO='" & sProducto & "'", Cn)
        nOficial = IIf(RsProducto!tMoneda = "02", nOficial * nTC, nOficial)
        lProductoPermiteDescuento = Calcular("select isnull(lPermiteDescuentos,0) as codigo from tclienteproducto where TCODIGODELIVERY='" & sCliente & "' AND TCODIGOPRODUCTO='" & sProducto & "'", Cn)
    End If
    'CPvalicacion LG
    
    'Busca Oferta
    Dim sCriterio As String
    Dim nOferta As Double
    Dim tOferta As String
    Dim lAcumulable As Boolean
    sCriterio = "tCodigoProducto ='" & sProducto & "' and lActivo=1 and lAutomatica=1"
    sCriterio = sCriterio & " and (substring(tFrecuencia," & Weekday(FechaServidor(), vbMonday) & "+1,1) = '1' or (substring(tFrecuencia,1,1)='1') and MONTH(fFecha) = " & Month(FechaServidor()) & " AND DAY(fFecha)= " & Day(FechaServidor()) & ") and tHoraInicial<='" & Format(Time, "HH:mm") & "' and tHoraFinal>='" & Format(Time, "HH:mm") & "'"
    sCriterio = sCriterio & " and (lPermanente=1 or (lPermanente=0 and fFechaInicial<='" & Format(FechaServidor(), "yyyy/mm/dd") & "' and fFechaFinal>='" & Format(FechaServidor(), "yyyy/mm/dd") & "')) "
    If sTipoPedido = "01" Then
       sCriterio = sCriterio & " and lLocal=1"
    ElseIf sTipoPedido = "02" Then
       sCriterio = sCriterio & " and lDelivery=1"
    ElseIf sTipoPedido = "03" Then
       sCriterio = sCriterio & " and lLlevar=1"
    ElseIf sTipoPedido = "04" Then
       sCriterio = sCriterio & " and lCanal4=1"
    Else
       sCriterio = sCriterio & " and lCanal5=1"
    End If
       
    Isql = "select * from TOFERTA where " & sCriterio
    Set RsOferta = Lib.OpenRecordset(Isql, Cn)
        
    'inserta descto
'    nOferta = 0
'    lAcumulable = False
'    If RsOferta.RecordCount > 0 And RsProducto!lDescuento Then
'       RsOferta.MoveFirst
'       If RsOferta!nRatio > 0 Then
'          nOferta = nOficial * IIf(IsNull(RsOferta!nRatio), 1, RsOferta!nRatio) / 100
'       Else
'          nOferta = nOficial - IIf(IsNull(RsOferta!nMonto), 0, RsOferta!nMonto)
'       End If
'    End If
       nOferta = 0
    lAcumulable = False
    If RsOferta.RecordCount > 0 And RsProducto!lDescuento Then
       RsOferta.MoveFirst
       tOferta = RsOferta!tOferta
       If RsOferta!nPrecio > 0 Then
          nOferta = nOficial - IIf(IsNull(RsOferta!nPrecio), 0, RsOferta!nPrecio)
       ElseIf RsOferta!nMonto > 0 Then
          nOferta = RsOferta!nMonto
       Else
          nOferta = nOficial * IIf(IsNull(RsOferta!nRatio), 1, RsOferta!nRatio) / 100
       End If
    Else
       If Calcular("select lExcluyente as Codigo from TOFERTA where tCodigoProducto ='" & sProducto & "' and lActivo=1", Cn) Then
          Screen.MousePointer = vbDefault
          MsgBox "Este producto no puede ser cargado en esta franja horaria" & Chr(13) & "Consulte con el Manager", vbCritical, sMensaje
          Exit Sub
       End If
    End If
     If lClienteControlaProducto Then
        nPVenta = nOficial
        If lProductoPermiteDescuento Then
            If xDescuento <> 0 And RsProducto!lDescuento Then
               If RsOferta.RecordCount > 0 Then
                  If RsOferta!lAcumulable Then
                     nPVenta = (nPVenta - nOferta) - ((nPVenta - nOferta) * xDescuento / 100)
                     nDescuento = nOficial - nPVenta
                  Else
                     nPVenta = nPVenta - nOferta
                     nDescuento = nOficial - nPVenta
                  End If
               Else
                  nPVenta = nPVenta - (nPVenta * xDescuento / 100)
                  nDescuento = nOficial - nPVenta
               End If
               
            Else
               nPVenta = nPVenta - nOferta
               nDescuento = nOficial - nPVenta
            End If
        Else
            nOferta = 0
            
        End If
    Else
        If xDescuento <> 0 And RsProducto!lDescuento Then
           If RsOferta.RecordCount > 0 Then
              If RsOferta!lAcumulable Then
                 nPVenta = (nOficial - nOferta) - ((nOficial - nOferta) * xDescuento / 100)
                 nDescuento = nOficial - nPVenta
              Else
                 nPVenta = nOficial - nOferta
                 nDescuento = nOficial - nPVenta
              End If
           Else
              nPVenta = nOficial - (nOficial * xDescuento / 100)
              nDescuento = nOficial - nPVenta
           End If
        Else
           nPVenta = nOficial - nOferta
           nDescuento = nOficial - nPVenta
        End If
    
    End If
        'CPvalicacion LG
    
'    If xDescuento <> 0 And RsProducto!lDescuento Then
'       If RsOferta.RecordCount > 0 Then
'          If RsOferta!lAcumulable Then
'             nPVenta = (nOficial - nOferta) - ((nOficial - nOferta) * xDescuento / 100)
'             nDescuento = nOficial - nPVenta
'          Else
'             nPVenta = nOficial - nOferta
'             nDescuento = nOficial - nPVenta
'          End If
'       Else
'          nPVenta = nOficial - (nOficial * xDescuento / 100)
'          nDescuento = nOficial - nPVenta
'       End If
'    Else
'       nPVenta = nOficial - nOferta
'       nDescuento = nOficial - nPVenta
'    End If
'
     Select Case pais ' ok
        Case "001" 'Bolivia
                nValor = (nValor / 100)
                nImpuesto1 = IIf(lImp1, nPVenta * nPorcentaje1 / 100, 0)
                nImpuesto2 = IIf(lImp2, nPVenta * nPorcentaje2 / 100, 0)
                nImpuesto3 = IIf(lImp3, nPVenta * nPorcentaje3 / 100, 0)
                nPBase = nPVenta - nImpuesto1 - nImpuesto2 - nImpuesto3
        
        Case Else 'Peru, Ecuador
                nValor = 1 + (nValor / 100)
                nImpuesto1 = IIf(lImp1, nPVenta / nValor * nPorcentaje1 / 100, 0)
                nImpuesto2 = IIf(lImp2, nPVenta / nValor * nPorcentaje2 / 100, 0)
                nImpuesto3 = IIf(lImp3, nPVenta / nValor * nPorcentaje3 / 100, 0)
                nPBase = nPVenta - nImpuesto1 - nImpuesto2 - nImpuesto3
    End Select
    
    
    Dim nInsumo As Double
    Dim nGasto As Double
    Dim nMObra As Double
    
    If sTipoPedido = "01" Then
       nInsumo = IIf(IsNull(RsProducto!nInsumo), 0, RsProducto!nInsumo)
       nGasto = IIf(IsNull(RsProducto!nGasto), 0, RsProducto!nGasto)
       nMObra = IIf(IsNull(RsProducto!nManoObra), 0, RsProducto!nManoObra)
    ElseIf sTipoPedido = "02" Then
       nInsumo = IIf(IsNull(RsProducto!nInsumo2), 0, RsProducto!nInsumo2)
       nGasto = IIf(IsNull(RsProducto!nGasto2), 0, RsProducto!nGasto2)
       nMObra = IIf(IsNull(RsProducto!nManoObra2), 0, RsProducto!nManoObra2)
    ElseIf sTipoPedido = "03" Then
       nInsumo = IIf(IsNull(RsProducto!nInsumo3), 0, RsProducto!nInsumo3)
       nGasto = IIf(IsNull(RsProducto!nGasto3), 0, RsProducto!nGasto3)
       nMObra = IIf(IsNull(RsProducto!nManoObra3), 0, RsProducto!nManoObra3)
    ElseIf sTipoPedido = "04" Then
       nInsumo = IIf(IsNull(RsProducto!nInsumo4), 0, RsProducto!nInsumo4)
       nGasto = IIf(IsNull(RsProducto!nGasto4), 0, RsProducto!nGasto4)
       nMObra = IIf(IsNull(RsProducto!nManoObra4), 0, RsProducto!nManoObra4)
    Else
       nInsumo = IIf(IsNull(RsProducto!nInsumo5), 0, RsProducto!nInsumo5)
       nGasto = IIf(IsNull(RsProducto!nGasto5), 0, RsProducto!nGasto5)
       nMObra = IIf(IsNull(RsProducto!nManoObra5), 0, RsProducto!nManoObra5)
    End If
    
    'Dim tAreaProduccion As String
       'multiarea produccion
    lProductoMultiArea = Calcular("select isnull(lmultiarea,0) as codigo from tproducto where tcodigoproducto='" & RsProducto.Fields("codigo") & "'", Cn)
    
    If lProductoMultiArea = False Then
        tsubalmacen = ""
    Else
        tsubalmacen = ""
        If lMultiAreaSubGrupo = True Then
            tAreaProduccion = Calcular("select isnull(tarea,'') codigo from TAREASUBGRUPO where tcaja='" & sCaja & "' and tSubGrupo='" & IIf(IsNull(RsProducto!tSubGrupo), "", RsProducto!tSubGrupo) & "'", Cn)
            tsubalmacen = Calcular("select isnull(tvalor,'')  as codigo from varea where codigo='" & tAreaProduccion & "'", Cn)
        End If
        If lMultiAreaCaja = True Then
            tAreaProduccion = Calcular("select isnull(tsubalmacen,'') as codigo from tcaja where tcaja='" & sCaja & "'", Cn)
            tsubalmacen = Calcular("select isnull(tvalor,'')  as codigo from varea where codigo='" & tAreaProduccion & "'", Cn)
        End If
        
    
        
        If tsubalmacen = "0" Then
            tsubalmacen = ""
        End If
    
    End If
    
    fxCombo "A", 1, sProducto
    Isql = "insert into [" & sDetalle & "] " & _
           "(tCodigoPedido, tTipoPedido, tItem, tCodigoProducto, tCodigoGrupo, tCodigoSubGrupo, " & _
           "nPrecioNeto, nRecargo, nDescuento, nPrecioOficial, nPrecioImpuesto1, nPrecioImpuesto2, nPrecioImpuesto3, nPrecioVenta, " & _
           "nCantidad, nVenta, nImpuesto1, nImpuesto2, nImpuesto3, " & _
           "lImprime, tArea, lImprimeArea, lCombinacion, nCombinacion, nInsumo, nGasto, nManoObra, nOrden, tEstadoItem,tsubalmacen,toferta,tCajaD) " & _
           "Values( '" & Pedido & "', '" & sTipoPedido & "', " _
                   & "'" & sitem & "', " _
                   & "'" & sProducto & "', " _
                   & "'" & IIf(IsNull(RsProducto!tgrupo), "", RsProducto!tgrupo) & "', " _
                   & "'" & IIf(IsNull(RsProducto!tSubGrupo), "", RsProducto!tSubGrupo) & "', " _
                   & nPBase & ", " & nRecargo & ", " _
                   & nDescuento & ", " _
                   & nOficial & ", " _
                   & nImpuesto1 & ", " & nImpuesto2 & ", " & nImpuesto3 & ", " _
                   & nPVenta & ", " & xCantidad & ", " _
                   & nPVenta * xCantidad & ", " _
                   & nImpuesto1 * xCantidad & ", " & nImpuesto2 * xCantidad & ", " & nImpuesto3 * xCantidad & ", " _
                   & "0, '" & RsProducto!tArea & "', " & IIf(RsProducto!lImprimeArea, -1, 0) & "," _
                   & IIf(RsProducto!lCombinacion, 1, 0) & ", " & RsProducto!nCombinacion & ", " _
                   & nInsumo & ", " _
                   & nGasto & ", " _
                   & nMObra & ", " _
                   & nOrden & ", " _
                   & "'N','" & tsubalmacen & "','" & tOferta & "', '" & sCaja & "') "
    Cn.Execute Isql
    RsDetalle.Requery
     nMonto = Format(Calcular("select sum(nVenta) as Codigo FROM [" & sDetalle & "] where tEstadoItem = 'N'", Cn), "#,###,##0.00")
    RsDetalle.MoveLast
    Screen.MousePointer = vbDefault

End Sub

Private Sub Form_Unload(Cancel As Integer)
'    If lBal Then
'       frmMsComm.MSCommBalanza.PortOpen = False
'    End If
   
   Cn.Execute "drop table " & sDetalle
   Cn.Execute "drop table " & sComboDetalle
   Cn.Execute "drop table " & sComboPropiedad
   Cn.Execute "drop table " & sProductoPropiedad
   Set frmCajaRapida = Nothing
End Sub

Public Sub AsignaProducto()
   If Not RsDetalle.EOF Then
      cmdPrecio.Enabled = IIf(IsNull(RsDetalle!lModificable), False, RsDetalle!lModificable)
      sProducto = IIf(IsNull(RsDetalle!tCodigoProducto), "", RsDetalle!tCodigoProducto)
      sSubGrupo = IIf(IsNull(RsDetalle!tCodigoSubGrupo), "", RsDetalle!tCodigoSubGrupo)
      sGrupo = IIf(IsNull(RsDetalle!tCodigoGrupo), "", RsDetalle!tCodigoGrupo)
      sitem = IIf(IsNull(RsDetalle!tItem), "001", RsDetalle!tItem)
      nOrden = IIf(IsNull(RsDetalle!nOrden), 0, RsDetalle!nOrden)
      sCortesia = IIf(IsNull(RsDetalle!tCortesia), "", RsDetalle!tCortesia)
      
      nOficial = IIf(IsNull(RsDetalle!nPrecioOficial), 0, RsDetalle!nPrecioOficial)
      nDescuento = IIf(IsNull(RsDetalle!nDescuento), 0, RsDetalle!nDescuento)
      nRecargo = IIf(IsNull(RsDetalle!nRecargo), 0, RsDetalle!nRecargo)
      nPBase = IIf(IsNull(RsDetalle!nPrecioNeto), 0, RsDetalle!nPrecioNeto)
      nImpuesto1 = IIf(IsNull(RsDetalle!nprecioImpuesto1), 0, RsDetalle!nprecioImpuesto1)
      nImpuesto2 = IIf(IsNull(RsDetalle!nprecioImpuesto2), 0, RsDetalle!nprecioImpuesto2)
      nImpuesto3 = IIf(IsNull(RsDetalle!nprecioImpuesto3), 0, RsDetalle!nprecioImpuesto3)
      nPVenta = IIf(IsNull(RsDetalle!nprecioVenta), 0, RsDetalle!nprecioVenta)
      nCantidad = IIf(IsNull(RsDetalle!nCantidad), 0, RsDetalle!nCantidad)
    
      txtOficial.Caption = Format(nOficial, "###,###,###,##0.00")
      txtNeto.Caption = Format(nPBase, "###,###,###,##0.00")
      txtDImporte.Caption = Format(nDescuento, "###,###,###,##0.00")
      txtRImporte.Caption = Format(nRecargo, "###,###,###,##0.00")
      txtImpuesto1.Caption = Format(nImpuesto1, "###,###,###,##0.00")
      txtImpuesto2.Caption = Format(nImpuesto2, "###,###,###,##0.00")
      txtImpuesto3.Caption = Format(nImpuesto3, "###,###,###,##0.00")
      txtPVenta.Caption = Format(nPVenta, "###,###,###,##0.00")
      txtCantidad.Caption = Format(nCantidad, "##,##0.00")
      txtVenta.Caption = Format(nPVenta * nCantidad, "###,###,###,##0.00")
      lblObservacion.Text = IIf(IsNull(RsDetalle!tObservacion), "", RsDetalle!tObservacion)
        If IIf(IsNull(RsDetalle!lImprime), False, RsDetalle!lImprime) = False Then
            'luchoinsumos
             verificatitulo
             'luchoinsumos
        Else
            Label2.Caption = muestra
        End If
                                        
                    
      If nOficial = 0 Then
         txtDPorcentaje.Caption = "0.00"
         txtRPorcentaje.Caption = "0.00"
      Else
         txtDPorcentaje.Caption = Format(nDescuento * 100 / nOficial, "###,###,###,##0.00")
         txtRPorcentaje.Caption = Format(nRecargo * 100 / nOficial, "###,###,###,##0.00")
      End If
           
           
      'Llena el Combo
      fraCombo.Caption = IIf(IsNull(RsDetalle!Producto), "", " " & RsDetalle!Producto & " ")
      wCombo = IIf(IsNull(RsDetalle!lCombinacion), False, RsDetalle!lCombinacion)
      nCombo = IIf(IsNull(RsDetalle!nCombinacion), 1, RsDetalle!nCombinacion)
      RsCombo.Filter = "[tItem]='" & sitem & "'"
      fraCombo.Visible = False
      wAgregaCombo = False
      
       If wCombo = True Then
        sProductoCombo = sProducto
      End If
      
      txtCortesia.Caption = IIf(IsNull(RsDetalle!Cortesia), "", RsDetalle!Cortesia)
      txtObserva.Caption = IIf(IsNull(RsDetalle!tObservacion), "", RsDetalle!tObservacion)
      VisualizaMonto
      
      tabProducto.Visible = True
      fraPropiedad.Visible = False
      ActivaCabecera True
      'ojoooooooo
      ListarOperadoresConFiltro sProducto
      AsignaPropiedad
   End If
End Sub

Public Sub Facturar()
On Error GoTo fin
    Dim sSerie As String
    Dim sCorrela As String
    Dim sPrefijo As String
    Dim RsSuma As Recordset
    Dim sTipoDocumento As String
    Dim sImp As String
    Dim wConsumo As Boolean
    Dim i As Integer
    Dim j As Integer
    Dim X As Integer
    Dim nRespuesta As Integer
    Dim imprimeDolaDocumentos As String
    
    Dim RscadenaCodigoHash As Recordset
    
    Dim xUltimoCorrelativo As String
    sDetalleConsumo = ""
    
    'FACTURACION_E_PERU
    Dim cadenaCodigoHash As String

    lImprimeAlternativa = False
    sCliente = ""
    wConsumo = False
    
    tAutorizacion = ""
    tcodigoControl = ""
    tDosificacion = ""
    tIdentidadNIT = ""
    
    'FACTURACION OFISIS
    Dim oComandoCabeceraOfisis As clsComando
    Dim oComandoDetalleOfisis As clsComando
    Dim oComandoFirmaDocumentoOfisis As clsComando
    
    Dim oComandoCabeceraOfisis1 As clsComando
    Dim oComandoDetalleOfisis1 As clsComando
    Dim oComandoFirmaDocumentoOfisis1 As clsComando
    
    Dim rdi As Integer
    
    
    
    lblPaso1.Visible = False
    lblPaso2.Visible = False
    imgProceso(0).Visible = False
    imgProceso(1).Visible = False
    imgProceso(2).Visible = False
    imgProceso(3).Visible = False
    FrameFeSpring.Visible = False
    
    
    
    
    sUsuarioAutoriza = sUsuario
    If RsDetalle.RecordCount = 0 Then
       Exit Sub
    End If

    'Chequea Consistencia
    RsTipoDocumento.Requery
    RsTipoDocumento.MoveFirst
    RsTipoDocumento.Find ("Descripcion='" & txtTipoDocumento.Caption & "'")
    If RsTipoDocumento.EOF Then
       MsgBox "Error: Configure los Documentos", vbCritical, sMensaje
       Exit Sub
    Else
       xlTipoDocumento = Calcular("Select lValidaRuc As Codigo From TTIPODOCUMENTO Where tCodigoTipoDocumento = '" & RsTipoDocumento!TTipoEmision & "'", Cn)
    End If
    
    If nPuerto > 0 Then
       Visor txtTipoDocumento.Caption, "", nPuerto, "N"
    End If

    X = Calcular("select count(tItem) as codigo from [" & sDetalle & "] where (isnull(tFacturado,'0') = '0' or len(ltrim(tFacturado)) = 0)", Cn)
    If X <= 0 Then
       MsgBox "Error: No existen Productos a Facturar", vbCritical, sMensaje
       Exit Sub
    End If

     'validacionmontominimo
    nMontoPedidoFacturarMInimo = 0
       nMontoPedidoFacturarMInimo = Calcular("select isnull(nMontoMinimo,0) as codigo   from vTipoDocumento where Codigo='" & RsTipoDocumento!TTipoEmision & "'", Cn)
       If nMontoPedidoFacturarMInimo > 0 Then
           If nMontoPedidoFacturarMInimo >= CDbl(txtMonto.Caption) Then
               MsgBox "El Monto a Facturar no llega al Minimo Permitido al Tipo de Documento"
               wEnter = False
               Exit Sub
           End If
       End If
           
    'validacionMontoMaximo
    nMontoPedidoFacturar = 0
    nMontoPedidoFacturar = Calcular("select isnull(nMontoMaximo,0) as codigo   from vTipoDocumento where Codigo='" & RsTipoDocumento!TTipoEmision & "'", Cn)
    If nMontoPedidoFacturar > 0 Then
    If nMontoPedidoFacturar < CDbl(txtMonto.Caption) Then
        MsgBox "El Monto a Facturar supera al Máximo Permitido al Tipo de Documento"
        wEnter = False
        Exit Sub
    End If
    End If
    
'     'validacion de Descuento para Facturas - peru
'    If xlTipoDocumento = True And pais = "000" Then
'        If Calcular("select sum(nimpuesto1) as codigo from " + sDetalle + " where tcodigopedido='" + sPedido + "'", Cn) <= 0 Then
'            MsgBox "Este Documento no se puede Emitir sin IGV!!! ", vbInformation, "Inforest"
'            Exit Sub
'        End If
'    End If

    'Consistencia Cortesia
    sCortesia = ""
    If RsTipoDocumento!TTipoEmision = "00" Then
       tUsuActua = sUsuario
       If Supervisor("04") = False Then
          MsgBox "Clave no permitida", vbExclamation, sMensaje
          Exit Sub
       End If
       sUsuario = tUsuActua
       sUsuarioAutoriza = sVar1
       sTemp = ""
       Isql = "select * from vCortesia where lActivo = 1 Order by Descripcion"
       Call ConfGrilla(2, frmBusquedaRapida.grdGrilla, "Código", 2, "Codigo", 1200, 2, 0, "", _
                                                       "Descripcion", 2, "Descripcion", 7000, 0, 0, "")
       frmBusquedaRapida.nPredeterm = 1
       frmBusquedaRapida.Show vbModal
       
       If wEnter = True Then
          sCortesia = sCodigo
          If nPuerto > 0 Then
             Visor "Cortesia", sDescrip, nPuerto, "N"
          End If
       Else
          Exit Sub
       End If
    End If

    If RsTipoDocumento!TTipoEmision = "00" And sCortesia <> "" Then
        Dim nTotalCortesiaActual As Double
        Dim nTopeCortesia As Double
        Dim nTotalDocActual As Double
        nTopeCortesia = Calcular("select isnull(tope,0) as codigo from vcortesia where codigo='" & sCortesia & "'", Cn)
        If nTopeCortesia > 0 Then
                nTotalCortesiaActual = Calcular("select sum(isnull(nventa,0))  as codigo From mDocumento where ttipodocumento='00' and tcortesia='" & sCortesia & "' and month(fregistro)=month(getdate()) ", Cn)
                nTotalDocActual = Val(txtMonto.Caption)
                    If nTotalCortesiaActual + nTotalDocActual > nTopeCortesia Then
                            MsgBox "Con esta Emisión se supera el Tope Mensual asignado para la Cortesia " & UCase(sDescrip) & vbCrLf & "Tope Mensual: " & nTopeCortesia & ". Ya Asignado : " & nTotalCortesiaActual, vbCritical
                            wEnter = False
                            variableEmite = False
                            Exit Sub
                    End If
        End If
    End If

       'impresion imagen
       Set rstFuente = New ADODB.Recordset
       imageCab.Picture = Nothing
       imagepIE.Picture = Nothing
       Set rstFuente = Lib.OpenRecordset("select iImagenCabDoc AS foto, iImagenPieDoc as fotoPie  from tcaja where tcaja='" & sCaja & "'", Cn)
       imageCab.DataField = "foto"
       Set imageCab.DataSource = rstFuente
       imagepIE.DataField = "fotoPie"
       Set imagepIE.DataSource = rstFuente
        
    
    'Por Consumo
    If lConsumo3 = True Then
       If RsTipoDocumento!TTipoEmision <> "00" Then
          nRespuesta = MsgBox("Por Consumo? ", vbQuestion + vbYesNoCancel + vbDefaultButton2, sMensaje)
          If nRespuesta = vbYes Then
             frmKeyBoard.txtResultado = "POR CONSUMO"
             frmKeyBoard.Show vbModal
             If sDescrip = "" Or Not wEnter Then
                MsgBox "Error: La descripcion no puede ser en blanco", vbCritical, sMensaje
                Exit Sub
             End If
             sDetalleConsumo = sDescrip
             wConsumo = True
          ElseIf nRespuesta = vbCancel Then
             Exit Sub
          End If
       End If
    End If
    
    TimpresionDolaresDelivery = False
    '-------- impresion en dolares si esta activo el check en el cliente delivery.
    If Calcular("select isnull(lEmisionMonedaExtranjera,0) as codigo from tdelivery where tcodigodelivery='" & sClienteFrecuente & "'", Cn) Then
        If MsgBox("¿Desea Imprimir en " & sMonedaE & "?", vbQuestion + vbYesNo + vbDefaultButton2) = vbYes Then
            TimpresionDolaresDelivery = True
        Else
            TimpresionDolaresDelivery = False
        End If
    Else
        TimpresionDolaresDelivery = False
    End If
    
    
    'Tipo de Emision
    If Not wConsumo And RsTipoDocumento!tFormulario <> "01" Then
       'Factura
       'Consistencia Factura

       If RsTipoDocumento!Cliente Then
          sTemp = ""
          Isql = "SELECT * from vCliente where lActivo = 1 Order by Descripcion"
                    Isql = "exec usp_Inforest_ObtieneClientesFactura '" & sClienteFrecuente & "','" & RsTipoDocumento!TTipoEmision & "'"

          frmBusquedaRapida.cmdOpcion(1).Enabled = True
          frmBusquedaRapida.cmdOpcion(2).Enabled = True
          frmBusquedaRapida.cmdOpcion(3).Enabled = True
          
          Select Case pais
            Case "001" 'Bolivia
                 Call ConfGrilla(3, frmBusquedaRapida.grdGrilla, "Código", 2, "Codigo", 1100, 2, 0, "", _
                                                          "Ident", 2, "tIdentidad", 1600, 2, 0, "", _
                                                          "Cliente", 2, "Descripcion", 5500, 0, 0, "")
            Case Else 'Peru, Ecuador
                        If lClub Then
                            Call ConfGrilla(4, frmBusquedaRapida.grdGrilla, "Código", 2, "Codigo", 1000, 2, 0, "", _
                                                                          "Ident", 2, "tIdentidad", 1600, 2, 0, "", _
                                                                          "Cliente", 2, "Descripcion", 4500, 0, 0, "", _
                                                                          "Enlace", 2, "tEnlace", 1100, 0, 0, "")
                        Else
                            Call ConfGrilla(3, frmBusquedaRapida.grdGrilla, "Código", 2, "Codigo", 1100, 2, 0, "", _
                                                                          "Ident", 2, "tIdentidad", 1600, 2, 0, "", _
                                                                          "Cliente", 2, "Descripcion", 5500, 0, 0, "")
                        End If
          End Select
          
          sTipoDocum = RsTipoDocumento!TTipoEmision

          frmBusquedaRapida.nPredeterm = 1
          frmBusquedaRapida.Show vbModal

          If wEnter = True And sCodigo <> "" Then
             sCliente = sCodigo
                     'imprimedni
                      Dim RsTc1 As ADODB.Recordset
                      Set RsTc1 = New ADODB.Recordset
                      Set RsTc1 = Lib.OpenRecordset("usp_Inforest_ValidaClienteSel '" & sTipoDocum & "','" & sCliente & "'", Cn)
                      If Not (RsTc1.EOF Or RsTc1.BOF) Then
                       RsTc1.MoveFirst
                       If RsTc1.Fields(0) <> "ok" Then
                           MsgBox "Error: El tipo de Identidad del Cliente no Corresponde al Tipo de Documento", vbCritical, sMensaje
                           Exit Sub
                       End If
                      End If
                 lValidaEmail = Calcular("Select lValidaEmail As codigo From vTipoDocumento where Codigo='" & sTipoDocum & "'", Cn)
                 
                 If lValidaEmail = True Then
                    sEmail = Calcular("Select ISNULL(tcorreo,'') As codigo From vCLIENTE where Codigo ='" & sCodigo & "' ", Cn)
                 
                    If sEmail = "" Then
                       MsgBox "El cliente no tiene Email registrado", vbCritical, sMensaje
                       Exit Sub
                    End If
                 End If
             
          Else
             Exit Sub
          End If
       End If
       
       If Pedido = "" Then
            GeneraPedido
            If lPrinter And lObligaPrinter Then
               i = Calcular("select count(tCodigoPedido) as codigo from " & sDetalle & " where lImprime=0", Cn)
               If i > 0 Then
                  cmdOpcion_Click (8)
                  If variableEmite = False Then: Exit Sub
               End If
            Else
               GeneraPedido
            End If
       Else
          sPedido = Pedido
          Cn.Execute "delete from DPEDIDO where tCodigoPedido='" & Pedido & "'"
          'Inserta el Detalle
          Cn.Execute "Insert into DPEDIDO (tCodigoPedido, tItem, tTipoPedido, tCodigoProducto, tCodigoGrupo, tCodigoSubGrupo, tMoneda, " & _
                     "nPrecioNeto, nPrecioImpuesto1, nPrecioImpuesto2, nPrecioImpuesto3, nPrecioVenta, nRecargo, nDescuento, nPrecioOficial, " & _
                     "nCantidad, nImpuesto1, nImpuesto2, nImpuesto3, nVenta, tObservacion, tCortesia, lImprime, tEstadoItem, tArea, lImprimeArea,fregistro, nInsumo, nGasto, nManoObra, nOrden, lCombinacion, nCombinacion, lCorte,toferta,tautorizaoferta,tsubalmacen,tCodigoEtiqueta,tunidadnegocio,fenvio,nenvio,fdiacontable) " & _
                     "select tCodigoPedido, tItem, tTipoPedido, tCodigoProducto, tCodigoGrupo, tCodigoSubGrupo, tMoneda, " & _
                     "nPrecioNeto, nPrecioImpuesto1, nPrecioImpuesto2, nPrecioImpuesto3, nPrecioVenta, nRecargo, nDescuento, nPrecioOficial, " & _
                     "nCantidad, nImpuesto1, nImpuesto2, nImpuesto3, nVenta, tObservacion, tCortesia, lImprime, tEstadoItem, tArea, lImprimeArea,getdate(), nInsumo, nGasto, nManoObra, nOrden, lCombinacion, nCombinacion, lCorte,toferta,tautorizaoferta,tsubalmacen,tCodigoEtiqueta,'" & sUnidadNegocio & "' ,fenvio,nenvio,'" & Format(obtieneDiaContable, "yyyyMMdd") & "' " & _
                     "From [" & sDetalle & "] where tEstadoItem='N'"
          Cn.Execute "Update MPEDIDO set tObservacion='" & txtObservacion.Caption & "', tMozo='" & sMozo & "' where tCodigoPedido='" & Pedido & "'"
       End If

       'Genera y Actualiza los Numero de Documento
       RsDetalle.MoveFirst
       
       If RsTipoDocumento!tFormulario = "03" Then
          nFItem = nItemV
       End If

       For i = 1 To IIf(X Mod nFItem = 0, Int(X / nFItem), Int(X / nFItem) + 1)
           RsTipoDocumento.Requery
           RsTipoDocumento.MoveFirst
           RsTipoDocumento.Find ("Descripcion='" & txtTipoDocumento.Caption & "'")
           If RsTipoDocumento.EOF Then
              MsgBox "Error: Configure los Documentos", vbCritical, sMensaje
              Exit Sub
           End If

           sSerie = RsTipoDocumento!tSerie
           sCorrela = Lib.Correlativo(RsTipoDocumento!tUltimoNumero, 9)
           sPrefijo = RsTipoDocumento!prefijo
           sTipoDocumento = RsTipoDocumento!TTipoEmision
           sImp = RsTipoDocumento!timpresora
           sDocumento = sPrefijo & sSerie & sCorrela
           sResumen = RsTipoDocumento!lResumen
           
             Select Case pais
                 Case "001" 'Bolivia
                         tAutorizacion = obtieneAutorizacionDosificacion(sCaja, "1")
                         tDosificacion = obtieneAutorizacionDosificacion(sCaja, "2")
                         If tAutorizacion <> "" And tDosificacion <> "" Then
                             
                         Else
                             MsgBox "Error al obtener Número de Autorización o Dosificación. Verifique.", vbCritical, sMensaje
                             Exit Sub
                         End If
                 Case "002" 'ECUADOR
                     tAutorizacion = RsTipoDocumento!tNumeroAutorizacion
                 Case Else 'Peru
                     tAutorizacion = ""
                     tcodigoControl = ""
                     tDosificacion = ""
                         
             End Select
           
           'Genera el Detalle de DDOCUMENTO
           Dim xClave As String
           For j = 1 To nFItem
               xClave = RsDetalle!tItem
               Isql = "Update DPEDIDO set tDocumento = '" & sDocumento & "' where tItem = '" & xClave & "' and tCodigoPedido = '" & Pedido & "' and (isnull(tFacturado,'0')='0' or len(ltrim(tFacturado)) = 0) "
               Cn.Execute Isql
               RsDetalle.MoveFirst
               RsDetalle.Find ("tItem ='" & xClave & "'")
               RsDetalle.MoveNext
               If RsDetalle.EOF Then
                  Exit For
               End If
           Next j

           'Inserta Detalle de Documento
           Isql = "Insert into DDOCUMENTO " & _
                  "       ( tDocumento, tItem, tCodigoPedido, tCodigoProducto, nPrecioNeto, nPrecioImpuesto1, nPrecioImpuesto2, nPrecioImpuesto3, " & _
                  "nPrecioVenta, nRecargo, nDescuento, nCantidad, nPrecioOficial, nImpuesto1, nImpuesto2, nImpuesto3, nVenta ) " & _
                  "select  '" & sDocumento & "' as tDocumento , tItem, tCodigoPedido, tCodigoProducto, nPrecioNeto, nPrecioImpuesto1, nPrecioImpuesto2, nPrecioImpuesto3, " & _
                  "nPrecioVenta, nRecargo, nDescuento, nCantidad, nPrecioOficial, nImpuesto1, nImpuesto2, nImpuesto3, nVenta From DPEDIDO " & _
                  "where (isnull(tFacturado,'0') = '0' or len(ltrim(tFacturado)) = 0) and tDocumento ='" & sDocumento & "'"
           Cn.Execute Isql

           'Genera el Detalle de MDOCUMENTO
           Isql = "Insert into MDOCUMENTO " & _
                  "     ( tDocumento, tTipoDocumento, tCodigoCliente, tEstadoDocumento, tCaja, tSalon, tTurno, tUsuario, tUsuarioAutoriza, fRegistro, fDiaContable, tConsumo, lImpresionMonedaExtranjera) " & _
                  "Values(   '" & sDocumento & "', " _
                          & "'" & sTipoDocumento & "', " _
                          & "'" & IIf(sCliente = "", "", sCliente) & "', " _
                          & "'01', " _
                          & "'" & sCaja & "', " _
                          & "'" & sSalon & "', " _
                          & "'" & sTurno & "', " _
                          & "'" & Mid(sUsuario, 1, 15) & "', " _
                          & "'" & sUsuarioAutoriza & "', " _
                          & " getdate(), '" & Format(obtieneDiaContable, "yyyyMMdd") & "','" & sDetalleConsumo & "', " & IIf(TimpresionDolaresDelivery, 1, 0) & " ) "
           Cn.Execute Isql

           'Calcula el total de la cabecera
           Set RsSuma = Lib.OpenRecordset("select sum(nPrecioNeto*nCantidad) as nNeto, sum(nImpuesto1) as nImpuesto1, sum(nImpuesto2) as nImpuesto2, sum(nImpuesto3) as nImpuesto3, sum(nVenta) as nVenta, isnull(sum(nDescuento*nCantidad),0) nDescuento " & _
                                          " from DPEDIDO where tDocumento ='" & sDocumento & "' group by tDocumento", Cn)

           'Actualiza el Documento con el Temporal
           nCargo = Round(RsSuma!nVenta, 2)

           Select Case pais
               Case "001"
                   tcodigoControl = devuelveCodigoControl(sCaja, sCorrela, tAutorizacion, tDosificacion, sCliente, nCargo)
           End Select
           
           Isql = "Update MDOCUMENTO set nNeto = " & RsSuma!nNeto & " , " & _
                                        "nRecargo = 0, " & _
                                        "nDescuento = " & IIf(lAplicablePedido, 0, RsSuma!nDescuento) & ", " & _
                                        "nPrecioOficial = 0 , " & _
                                        "nPrecioImpuesto1 = " & RsSuma!nImpuesto1 & " , " & _
                                        "nPrecioImpuesto2 = " & RsSuma!nImpuesto2 & " , " & _
                                        "nPrecioImpuesto3 = " & RsSuma!nImpuesto3 & " , " & _
                                        "tautorizacion = '" & tAutorizacion & "' , " & _
                                        "tcodigocontrol = '" & tcodigoControl & "' , " & _
                                        "nVenta = " & RsSuma!nVenta & _
                                        " ,lreplica=1 where tDocumento = '" & sDocumento & "'"
           Cn.Execute Isql
           
           
           wEnter = True
           
           If lPagoAntesImpresion Then
                 xTipo = ""
                 sFormulario = "CajaRapida"
                
                 If sFormulario = "CajaRapida" And lPagoRapido = True Then
                    frmPagoRapido.Show vbModal
                 Else
                    frmPago.Show vbModal
                 End If
                 
                 If wEnter = False Then
                    Dim RsCantDocumentos As Recordset
                    Set RsCantDocumentos = Lib.OpenRecordset("select distinct tDocumento from DDOCUMENTO where tCodigoPedido ='" & Pedido & "'", Cn)
                    
                    For rdi = 0 To RsCantDocumentos.RecordCount - 1
                        Cn.Execute "Delete MDOCUMENTO Where tDocumento= '" & RsCantDocumentos!tDocumento & "'"
                        Cn.Execute "Delete DDOCUMENTO Where tDocumento= '" & RsCantDocumentos!tDocumento & "'"
                        Cn.Execute "Delete DPAGODOCUMENTO Where tDocumento= '" & RsCantDocumentos!tDocumento & "'"
                        Cn.Execute "update DPEDIDO set tFacturado = '' , tDocumento = '' where tCodigoPedido = '" & Pedido & "'"
                        RsCantDocumentos.MoveNext
                    Next rdi
                    
                    xUltimoCorrelativo = Calcular("select MAX(tDocumento) as codigo from MDOCUMENTO where tcaja='" & sCaja & "' and tTipoDocumento='" & sTipoDocumento & "'", Cn)
                    xUltimoCorrelativo = Right(xUltimoCorrelativo, 9)
                    
                    Cn.Execute "Update TTIPODOCUMENTOIMPRESORA Set tUltimoNumero = '" & xUltimoCorrelativo & "' where tTipoEmision ='" & sTipoDocumento & "' and tCaja ='" & sCaja & "'"
                    Exit Sub
                 End If
           End If
           
           
           
           Cn.Execute "Update TTIPODOCUMENTOIMPRESORA Set tUltimoNumero = '" & sCorrela & "' where tTipoEmision ='" & sTipoDocumento & "' and tCaja ='" & sCaja & "'"
           
           If lPagoAntesImpresion = False Then
           
                   nMonto = RsSuma!nVenta
                   
                   'Actualiza Base de Datos Detalle del Pedido
                   Cn.Execute "Update DPEDIDO set tFacturado = 'F' where tDocumento ='" & sDocumento & "'"
                   Cn.Execute "Update MPEDIDO set tEstadoPedido = '02' where tCodigoPedido ='" & Pedido & "'"
                   
                   'PARA NO FISCALES
                   Cn.Execute "UPDATE DPEDIDO SET lregistroventa=(select case when registroventa=0 then 0 else 1 end from vtipodocumento where codigo='" & sTipoDocumento & "') where  tCodigoPedido ='" & Pedido & "' and tDocumento ='" & sDocumento & "'"
                   
                   'Imprime Documentos
                    If wConsumo = False And lDescripcionAlternativa = True Then
                     If validaImpresionAlternativa(sDocumento) = False Then
                            If MsgBox("Desea imprimir descripción Alternativa? ", vbQuestion + vbYesNo + vbDefaultButton2, sMensaje) = vbYes Then
                                  lImprimeAlternativa = True
                            End If
                      End If
                    End If
                    '-------------------------------  SE INTRDUJO LAS CONSULTAS A  UN STORE PROCEDURE -------------------------
                    ' ELDCQ 22/11/2017
                    If lImprimeAlternativa = False Then
                        Isql = "EXEC usp_Inforest_Impresion '" & sDocumento & "',3"
                        'FACTURACION_E_PERU
                        IsqlFact = "EXEC usp_Inforest_Impresion '" & sDocumento & "',4"
                    Else
                        Isql = "EXEC usp_Inforest_Impresion '" & sDocumento & "',7"
                        'FACTURACION_E_PERU
                        IsqlFact = "EXEC usp_Inforest_Impresion '" & sDocumento & "',8"
                    End If
                    '-------------------------------------------------------------------------------------------------------------
                    'FACTURACION_E_PERU
                    Set RsImpDocumentoE = Lib.OpenRecordset(IsqlFact, Cn)
                    xImpresionFE = Calcular(" SELECT tImpresionFE as codigo FROM vTipodocumento WHERE Codigo=(SELECT TTIPODOCUMENTO FROM MDOCUMENTO WHERE tDocumento='" & sDocumento & "')", Cn)
                    xImpresioDE = Mid(sDocumento, 1, 1)
                    '---------------------------------------
                
                    Set RsImpresion = Lib.OpenRecordset(Isql, Cn)
        
                        imprimeDolaDocumentos = Calcular("select isnull(lequivadolares,0) as codigo from vtipodocumentoimpresora where tcaja='" & sCaja & "' and ttipoemision='" & sTipoDocumento & "' ", Cn)
                        If imprimeDolaDocumentos = "Verdadero" Then
                            lDocumEquivaPrecuenta = True
                        Else
                            lDocumEquivaPrecuenta = False
                        End If
                                      
                    If RsImpresion.RecordCount = 0 Then
                       LimpiaRs
                       MsgBox "No existen Datos a Imprimir", vbExclamation, sMensaje
                    Else
                            'SUNAT
                            numeroSerieImpresora = obtieneNumeroSerieImpresora(sCaja, sImp)
                            codigoImpresora = sImp
                            'SUNAT
                            Cn.Execute " update mdocumento set timpresora='" & codigoImpresora & "', tSerieImpresora='" & numeroSerieImpresora & "' where tdocumento ='" & sDocumento & "' "

                           'FACTURACION_E_PERU
                           If pais = "000" Then
                               If lFacturacionE Then
                                    If lFEOfisis Then 'OFISIS
                        
                                            '----CABECERA
                                            Set oComandoCabeceraOfisis = New clsComando
                                            If Not oComandoCabeceraOfisis.CreateCmdSp("USP_FactDocumentoOfisis", Cn) Then
                                                 Set oComandoCabeceraOfisis = Nothing
                                                 Exit Sub
                                            End If
                                            oComandoCabeceraOfisis.CreateParameter "@CodigoDocumento", adVarChar, adParamInput, 20, sDocumento
                        
                                            If Not oComandoCabeceraOfisis.GetParamOK Then
                                                 Set oComandoCabeceraOfisis = Nothing
                                                 Exit Sub
                                            End If
                                            If Not oComandoCabeceraOfisis.ExecSP Then
                                                 Set oComandoCabeceraOfisis = Nothing
                                                 Exit Sub
                                            End If
                                        
                                            '----FIRMA DOCUMENTO OFISIS
                                            If RsTipoDocumento!lDocumentoElectronicoOfisis Then
                                                Set oComandoFirmaDocumentoOfisis = New clsComando
                                                If Not oComandoFirmaDocumentoOfisis.CreateCmdSp("USP_FactFirmaDocumentoOfisis", Cn) Then
                                                     Set oComandoFirmaDocumentoOfisis = Nothing
                                                     Exit Sub
                                                End If
                                                oComandoFirmaDocumentoOfisis.CreateParameter "@CodigoDocumento", adVarChar, adParamInput, 20, sDocumento
                        
                                                If Not oComandoFirmaDocumentoOfisis.GetParamOK Then
                                                     Set oComandoFirmaDocumentoOfisis = Nothing
                                                     Exit Sub
                                                End If
                                                If Not oComandoFirmaDocumentoOfisis.ExecSP Then
                                                     Set oComandoFirmaDocumentoOfisis = Nothing
                                                     Exit Sub
                                                End If
                                                
                                                'VALIDAR RESPUESTA CODIGO DE BARRA
                                                fDocumento = Mid(sDocumento, 1, 1) + Mid(sDocumento, 4, 3) + "-" + CStr(CLng(Mid(sDocumento, 8, 8)))
                                                Sleep 3000
                                                If lImpresionCodigoBarras Then
                                                    imageHash.DataField = "foto"
                                                    Set RsCodigoHash = Lib.OpenRecordset("USP_FactObtenerCodigoBarraOfisis '" & fDocumento & "','" & Mid(sDocumento, 1, 1) & "','' ", Cn)
                                                    Set imageHash.DataSource = RsCodigoHash
                                                    
                                                ElseIf lQRFE Then
                                                    Set imageHash.Picture = LoadPicture(ImagenQR_Ofisis(fDocumento, sDocumento))
                                                Else
                                                    Set RscadenaCodigoHash = Lib.OpenRecordset("USP_FactConsultaHash '" & fDocumento & "','0' ", Cn)
                                                    If RscadenaCodigoHash.RecordCount > 0 Then
                                                        cadenaCodigoHash = RscadenaCodigoHash!codigo
                                                    End If
                                                    'cadenaCodigoHash = Calcular("select CO_HASH as codigo from TCFACT_ELEC where NU_DOCU='" & fDocumento & "' and (TI_DOCU='B' or TI_DOCU ='F')", CnFE)
                                                End If
                                            End If
                                    
                                    ElseIf lFESpring Then
                                    
                                    ElseIf lFECarbajal Then
                                    
                                    ElseIf lFEpape Then
                                        If IIf(RsTipoDocumento!lFacturacionElectronica = True, 1, 0) Then
                                            If tCodigoFE = "000" Then
                                                 If lQRFE Then
                                                     Set imageHash.Picture = LoadPicture(CrearImagenQR(PapeTermico))
                                                 Else
                                                     If lImpresionCodigoBarras Then
                                                         'Set imageHash.Picture = LoadPicture(ImagenQR(sDocumento))
                                                     Else
                                                         cadenaCodigoHash = PapeMatricial
                                                     End If
                                                 End If
                                             End If
                                        End If
                                    ElseIf lFEGesa Then
                                        If IIf(RsTipoDocumento!lFacturacionElectronica = True, 1, 0) Then
                                            If Not INSERTAFE(sDocumento, "", 1, "") Then '----CABECERA
                                                MsgBox "No se pudo enviar el documento a facturacion electronica", vbInformation
                                            End If
                                            If lQRFE Then
                                                Set imageHash.Picture = LoadPicture(QRHASH_FE_INFOREST(3, sDocumento, 0))
                                            Else
                                                If lImpresionCodigoBarras Then
                                                    Set imageHash.Picture = LoadPicture(QRHASH_FE_INFOREST(1, sDocumento, 0))
                                                Else
                                                    cadenaCodigoHash = QRHASH_FE_INFOREST(2, sDocumento, 0)
                                                End If
                                            End If
                                        End If
                                    Else 'INFOFACT
                                        If IIf(RsTipoDocumento!lFacturacionElectronica = True, 1, 0) Then
                                             If Not INSERTAFE(sDocumento, "", 1, "") Then '----CABECERA
                                                 Exit Sub
                                             End If
                                             If RsImpDocumentoE!Ruc <> "" Then
                                                 If Not INSERTAFE(sDocumento, "", 2, RsImpDocumentoE!Ruc) Then '----CLIENTE
                                                     Exit Sub
                                                 End If
                                             End If
                                             'VALIDAR RESPUESTA DE CODIGO HASH Y CODIGO DE BARRA
                                             fDocumento = Mid(sDocumento, 1, 1) + Mid(sDocumento, 4, 3) + Mid(sDocumento, 8, 8)
                                             If tCodigoFE = "000" Then
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
                                                   
                        End If
                        '---------------------------------------
        
                        'Configura la Impresora
                        Imprimir (sImp)
                        Printer.FontName = sFont
                        Printer.FontBold = False
        
                       'TVARIABLE CESAR
                       'FORMATO TICKET VARIABLE
                       If RsTipoDocumento!tFormulario = "03" Then
                               If RsTipoDocumento!Cliente And RsTipoDocumento!Monto = 0 Then
                               
                                     ImprimeFacturaVariable RsImpresion, sEmpresa
                                                   
                                  NFactura = sCorrela
                                  frmVenta.lblFactura.Caption = NFactura
                               ElseIf RsTipoDocumento!TTipoEmision = "00" Then
                                  If MsgBox("Deseas imprimir el Voucher", vbQuestion + vbYesNo, sMensaje) = vbYes Then
                                     ImprimeCortesia RsImpresion, sTipoDocumento, imageCab, imagepIE
                                  End If
                               Else
                                  ImprimeBoletaN RsImpresion, sEmpresa, sTipoDocumento
                               End If
                              
                       Else
                       
                               'FORMATO VARIABLE
                               If lFacturacionE And lFEOfisis = False Then
                                       'FACTURACION_E_PERU
                                       'FORMATO A4
                                       If Generar_Imagen(CnFE, "select imagen from IMAGENCODIGOBARRA where nro_efact='" & fDocumento & "'", "imagen", "\fact.bmp") = True Then
                                              ImprimeFormatoA sDocumento
                                              Kill App.Path & "\fact.bmp"
                                       Else
                                              ImprimeFormatoA sDocumento
                                       End If
                                                                                 
                               Else
                                           If sTipoDocumento = "01" Then
                                              If wConsumo Then
                                                 ImprimeFacturaConsumoN RsImpresion, sDetalleConsumo, sEmpresa
                                              Else
                                                 ImprimeFacturaN RsImpresion, sEmpresa, sTipoDocumento
                                              End If
                                              NFactura = sCorrela
                                              frmVenta.lblFactura.Caption = NFactura
                                           Else
                                              If wConsumo Then
                                                 ImprimeBoletaConsumoN RsImpresion, sDetalleConsumo, sEmpresa
                                              Else
                                                 ImprimeBoletaN RsImpresion, sEmpresa, sTipoDocumento
                                              End If
                                           
                                           End If
                               
                               End If
            
            
                        End If
                      
                   End If
                   
'                   If pais = "002" Then
'                        sXML = Trim(LeerIni(App.Path + "\INFOREST.INI", "Configuracion", "RutaXML", "."))
'                        GeneraFacturaElectronica sXML, sDocumento
'                   End If
                    If pais = "002" And lFEEcuador = False Then
                       sXML = Trim(LeerIni(App.Path + "\INFOREST.INI", "Configuracion", "RutaXML", "."))
                       GeneraFacturaElectronica sXML, sDocumento
                    End If
                    
                    If lFEEcuador Then
                     If INSERTA_FE_INFOREST(sDocumento, 1, DateTime.Now) = False Then
                         MsgBox "No se pudo enviar el documento a Facturacion Electronica!!! Verificar con su area de sistemas!!!"
                     End If
                    End If
                   
           End If
           
           If lPagoAntesImpresion = False Then
                xTipo = ""
                sFormulario = "CajaRapida"
                
                If sFormulario = "CajaRapida" And lPagoRapido = True Then
                   frmPagoRapido.Show vbModal
                Else
                   frmPago.Show vbModal
                End If
           End If
        Next i
        
        
        
        If lPagoAntesImpresion Then
            Set RsCantDocumentos = Lib.OpenRecordset("select distinct tDocumento from DDOCUMENTO where tCodigoPedido ='" & Pedido & "'", Cn)
            For rdi = 0 To RsCantDocumentos.RecordCount - 1
                sDocumento = RsCantDocumentos!tDocumento
                
                   'nMonto = RsSuma!nVenta
                   
                   'Actualiza Base de Datos Detalle del Pedido
                   Cn.Execute "Update DPEDIDO set tFacturado = 'F' where tDocumento ='" & sDocumento & "'"
                   Cn.Execute "Update MPEDIDO set tEstadoPedido = '02' where tCodigoPedido ='" & Pedido & "'"
                   
                   'PARA NO FISCALES
                   Cn.Execute "UPDATE DPEDIDO SET lregistroventa=(select case when registroventa=0 then 0 else 1 end from vtipodocumento where codigo='" & sTipoDocumento & "') where  tCodigoPedido ='" & Pedido & "' and tDocumento ='" & sDocumento & "'"
                   
                   'Imprime Documentos
                    If wConsumo = False And lDescripcionAlternativa = True Then
                     If validaImpresionAlternativa(sDocumento) = False Then
                            If MsgBox("Desea imprimir descripción Alternativa? ", vbQuestion + vbYesNo + vbDefaultButton2, sMensaje) = vbYes Then
                                  lImprimeAlternativa = True
                            End If
                      End If
                    End If
                    '-------------------------------  SE INTRDUJO LAS CONSULTAS A  UN STORE PROCEDURE -------------------------
                    ' ELDCQ 22/11/2017
                    If lImprimeAlternativa = False Then
                        Isql = "EXEC usp_Inforest_Impresion '" & sDocumento & "',3"
                        'FACTURACION_E_PERU
                        IsqlFact = "EXEC usp_Inforest_Impresion '" & sDocumento & "',4"
                    Else
                        Isql = "EXEC usp_Inforest_Impresion '" & sDocumento & "',7"
                        'FACTURACION_E_PERU
                        IsqlFact = "EXEC usp_Inforest_Impresion '" & sDocumento & "',8"
                    End If
                    '----------------------------------------------------------------------------------------------------------------------
                    'FACTURACION_E_PERU
                    Set RsImpDocumentoE = Lib.OpenRecordset(IsqlFact, Cn)
                    xImpresionFE = Calcular(" SELECT tImpresionFE as codigo FROM vTipodocumento WHERE Codigo=(SELECT TTIPODOCUMENTO FROM MDOCUMENTO WHERE tDocumento='" & sDocumento & "')", Cn)
                    xImpresioDE = Mid(sDocumento, 1, 1)
                    '---------------------------------------
                
                    Set RsImpresion = Lib.OpenRecordset(Isql, Cn)
        
                        imprimeDolaDocumentos = Calcular("select isnull(lequivadolares,0) as codigo from vtipodocumentoimpresora where tcaja='" & sCaja & "' and ttipoemision='" & sTipoDocumento & "' ", Cn)
                        If imprimeDolaDocumentos = "Verdadero" Then
                            lDocumEquivaPrecuenta = True
                        Else
                            lDocumEquivaPrecuenta = False
                        End If
                                      
                    If RsImpresion.RecordCount = 0 Then
                       LimpiaRs
                       MsgBox "No existen Datos a Imprimir", vbExclamation, sMensaje
                    Else
                   
                            'SUNAT
                            numeroSerieImpresora = obtieneNumeroSerieImpresora(sCaja, sImp)
                            codigoImpresora = sImp
                            'SUNAT
                            Cn.Execute " update mdocumento set timpresora='" & codigoImpresora & "', tSerieImpresora='" & numeroSerieImpresora & "' where tdocumento ='" & sDocumento & "' "
                
        
                           'FACTURACION_E_PERU
                           If pais = "000" Then
                               If lFacturacionE Then
                                    If lFEOfisis Then 'OFISIS
                        
                                            '----CABECERA
                                            Set oComandoCabeceraOfisis = New clsComando
                                            If Not oComandoCabeceraOfisis.CreateCmdSp("USP_FactDocumentoOfisis", Cn) Then
                                                 Set oComandoCabeceraOfisis = Nothing
                                                 Exit Sub
                                            End If
                                            oComandoCabeceraOfisis.CreateParameter "@CodigoDocumento", adVarChar, adParamInput, 20, sDocumento
                        
                                            If Not oComandoCabeceraOfisis.GetParamOK Then
                                                 Set oComandoCabeceraOfisis = Nothing
                                                 Exit Sub
                                            End If
                                            If Not oComandoCabeceraOfisis.ExecSP Then
                                                 Set oComandoCabeceraOfisis = Nothing
                                                 Exit Sub
                                            End If
                                           
                                            '----FIRMA DOCUMENTO OFISIS
                                            If RsTipoDocumento!lDocumentoElectronicoOfisis Then
                                                Set oComandoFirmaDocumentoOfisis = New clsComando
                                                If Not oComandoFirmaDocumentoOfisis.CreateCmdSp("USP_FactFirmaDocumentoOfisis", Cn) Then
                                                     Set oComandoFirmaDocumentoOfisis = Nothing
                                                     Exit Sub
                                                End If
                                                oComandoFirmaDocumentoOfisis.CreateParameter "@CodigoDocumento", adVarChar, adParamInput, 20, sDocumento
                        
                                                If Not oComandoFirmaDocumentoOfisis.GetParamOK Then
                                                     Set oComandoFirmaDocumentoOfisis = Nothing
                                                     Exit Sub
                                                End If
                                                If Not oComandoFirmaDocumentoOfisis.ExecSP Then
                                                     Set oComandoFirmaDocumentoOfisis = Nothing
                                                     Exit Sub
                                                End If
                                                
                                                'VALIDAR RESPUESTA CODIGO DE BARRA
                                                fDocumento = Mid(sDocumento, 1, 1) + Mid(sDocumento, 4, 3) + "-" + CStr(CLng(Mid(sDocumento, 8, 8)))
                                                Sleep 3000
                                                If lImpresionCodigoBarras Then
                                                    imageHash.DataField = "foto"
                                                    Set RsCodigoHash = Lib.OpenRecordset("USP_FactObtenerCodigoBarraOfisis '" & fDocumento & "','" & Mid(sDocumento, 1, 1) & "','' ", Cn)
                                                    Set imageHash.DataSource = RsCodigoHash
                                                Else
                                                    cadenaCodigoHash = Calcular("select CO_HASH as codigo from TCFACT_ELEC where NU_DOCU='" & fDocumento & "' and (TI_DOCU='B' or TI_DOCU ='F')", CnFE)
                                                End If
                                            End If
                                     
                                    ElseIf lFESpring Then
                                    
                                    ElseIf lFECarbajal Then
                        
                                    ElseIf lFEpape Then
                                        If IIf(RsTipoDocumento!lFacturacionElectronica = True, 1, 0) Then
                                            If tCodigoFE = "000" Then
                                                 If lQRFE Then
                                                     Set imageHash.Picture = LoadPicture(CrearImagenQR(PapeTermico))
                                                 Else
                                                     If lImpresionCodigoBarras Then
                                                         'Set imageHash.Picture = LoadPicture(ImagenQR(sDocumento))
                                                     Else
                                                         cadenaCodigoHash = PapeMatricial
                                                     End If
                                                 End If
                                             End If
                                        End If
                                    Else 'INFOFACT
                                        If IIf(RsTipoDocumento!lFacturacionElectronica = True, 1, 0) Then
                                             If Not INSERTAFE(sDocumento, "", 1, "") Then '----CABECERA
                                                 Exit Sub
                                             End If
                                             If RsImpDocumentoE!Ruc <> "" Then
                                                 If Not INSERTAFE(sDocumento, "", 2, RsImpDocumentoE!Ruc) Then '----CLIENTE
                                                     Exit Sub
                                                 End If
                                             End If
                                             'VALIDAR RESPUESTA DE CODIGO HASH Y CODIGO DE BARRA
                                             fDocumento = Mid(sDocumento, 1, 1) + Mid(sDocumento, 4, 3) + Mid(sDocumento, 8, 8)
                                             If tCodigoFE = "000" Then
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
                        End If
                        '---------------------------------------
        
                        'Configura la Impresora
                        Imprimir (sImp)
                        Printer.FontName = sFont
                        Printer.FontBold = False
        
                       'TVARIABLE CESAR
                       'FORMATO TICKET VARIABLE
                       If RsTipoDocumento!tFormulario = "03" Then
                               If RsTipoDocumento!Cliente And RsTipoDocumento!Monto = 0 Then
                               
                                     ImprimeFacturaVariable RsImpresion, sEmpresa
                                                   
                                  NFactura = sCorrela
                                  frmVenta.lblFactura.Caption = NFactura
                               ElseIf RsTipoDocumento!TTipoEmision = "00" Then
                                  If MsgBox("Deseas imprimir el Voucher", vbQuestion + vbYesNo, sMensaje) = vbYes Then
                                     ImprimeCortesia RsImpresion, sTipoDocumento, imageCab, imagepIE
                                  End If
                               Else
                                  ImprimeBoletaN RsImpresion, sEmpresa, sTipoDocumento
                               End If
                              
                       Else
                       
                               'FORMATO VARIABLE
                               If lFacturacionE And lFEOfisis = False Then
                                       'FACTURACION_E_PERU
                                       'FORMATO A4
                                       If Generar_Imagen(CnFE, "select imagen from IMAGENCODIGOBARRA where nro_efact='" & fDocumento & "'", "imagen", "\fact.bmp") = True Then
                                              ImprimeFormatoA sDocumento
                                              Kill App.Path & "\fact.bmp"
                                       Else
                                              ImprimeFormatoA sDocumento
                                       End If
                                                                                 
                               Else
                                           If sTipoDocumento = "01" Then
                                              If wConsumo Then
                                                 ImprimeFacturaConsumoN RsImpresion, sDetalleConsumo, sEmpresa
                                              Else
                                                 ImprimeFacturaN RsImpresion, sEmpresa, sTipoDocumento
                                              End If
                                              NFactura = sCorrela
                                              frmVenta.lblFactura.Caption = NFactura
                                           Else
                                              If wConsumo Then
                                                 ImprimeBoletaConsumoN RsImpresion, sDetalleConsumo, sEmpresa
                                              Else
                                                 ImprimeBoletaN RsImpresion, sEmpresa, sTipoDocumento
                                              End If
                                           
                                           End If
                               
                               End If
            
            
                        End If
                      
                   End If
                   
'                   If pais = "002" Then
'                        sXML = Trim(LeerIni(App.Path + "\INFOREST.INI", "Configuracion", "RutaXML", "."))
'                        GeneraFacturaElectronica sXML, sDocumento
'                   End If
                    If pais = "002" And lFEEcuador = False Then
                       sXML = Trim(LeerIni(App.Path + "\INFOREST.INI", "Configuracion", "RutaXML", "."))
                       GeneraFacturaElectronica sXML, sDocumento
                    End If
                    
                    If lFEEcuador Then
                     If INSERTA_FE_INFOREST(sDocumento, 1, DateTime.Now) = False Then
                         MsgBox "No se pudo enviar el documento a Facturacion Electronica!!! Verificar con su area de sistemas!!!"
                     End If
                    End If

            RsCantDocumentos.MoveNext
            Next rdi
        End If

        LimpiaRs
       
       
    '------------------- FORMATO TICKET
    Else
    
       If RsTipoDocumento!Cliente And (RsTipoDocumento!Monto <= nMonto Or RsTipoDocumento!Monto = 0) Then
          'Factura
          sTemp = ""
          Isql = "SELECT * from vCliente where lActivo = 1 Order by Descripcion"
          'imprimedni
          Isql = "exec usp_Inforest_ObtieneClientesFactura '" & sClienteFrecuente & "','" & RsTipoDocumento!TTipoEmision & "'"

          frmBusquedaRapida.cmdOpcion(1).Enabled = True
          frmBusquedaRapida.cmdOpcion(2).Enabled = True
          frmBusquedaRapida.cmdOpcion(3).Enabled = True
          
          Select Case pais
            Case "001" 'Bolivia
                Call ConfGrilla(3, frmBusquedaRapida.grdGrilla, "Código", 2, "Codigo", 1100, 2, 0, "", _
                                                          "Ident", 2, "tIdentidad", 1600, 2, 0, "", _
                                                          "Cliente", 2, "Descripcion", 5500, 0, 0, "")
            Case Else 'Peru, Ecuador
                        If lClub Then
                            Call ConfGrilla(4, frmBusquedaRapida.grdGrilla, "Código", 2, "Codigo", 1000, 2, 0, "", _
                                                                          "Ident", 2, "tIdentidad", 1600, 2, 0, "", _
                                                                          "Cliente", 2, "Descripcion", 4500, 0, 0, "", _
                                                                          "Enlace", 2, "tEnlace", 1100, 0, 0, "")
                        Else
                            Call ConfGrilla(3, frmBusquedaRapida.grdGrilla, "Código", 2, "Codigo", 1100, 2, 0, "", _
                                                                          "Ident", 2, "tIdentidad", 1600, 2, 0, "", _
                                                                          "Cliente", 2, "Descripcion", 5500, 0, 0, "")
                        End If
          End Select
           '------VALIDA CORREO----------
          sTipoDocum = RsTipoDocumento!TTipoEmision
          frmBusquedaRapida.nPredeterm = 1
          frmBusquedaRapida.Show vbModal
          
          If wEnter = True And sCodigo <> "" Then
             sCliente = sCodigo
                    Dim RsTc As ADODB.Recordset
                    Set RsTc = New ADODB.Recordset
                    Set RsTc = Lib.OpenRecordset("usp_Inforest_ValidaClienteSel '" & sTipoDocum & "','" & sCliente & "'", Cn)
                    If Not (RsTc.EOF Or RsTc.BOF) Then
                     RsTc.MoveFirst
                     If RsTc.Fields(0) <> "ok" Then
                         MsgBox "Error: El tipo de Identidad del Cliente no Corresponde al Tipo de Documento", vbCritical, sMensaje
                         Exit Sub
                     End If
                    End If
                    
                    lValidaEmail = Calcular("Select lValidaEmail As codigo From vTipoDocumento where Codigo='" & sTipoDocum & "'", Cn)
                    
                    If lValidaEmail = True Then
                       sEmail = Calcular("Select ISNULL(tcorreo,'') As codigo From vCLIENTE where Codigo ='" & sCodigo & "' ", Cn)
                    
                       If sEmail = "" Then
                          MsgBox "El cliente no tiene Email registrado", vbCritical, sMensaje
                          Exit Sub
                       End If
                    End If
                    
                    If Calcular("Select lValidaUbigeo As codigo From vTipoDocumento where Codigo='" & sTipoDocum & "'", Cn) = True Then
                        Dim TempUbigeo As String
                        Dim TempUrbaniza As String
                        TempUbigeo = Calcular("Select ISNULL(CodigoUbigeo,'') As codigo From vCLIENTE where Codigo ='" & sCodigo & "' ", Cn)
                        TempUrbaniza = Calcular("Select ISNULL(Urbanizacion,'') As codigo From vCLIENTE where Codigo ='" & sCodigo & "' ", Cn)
                        If Trim(TempUbigeo) = "" Or Trim(TempUrbaniza) = "" Then
                            MsgBox "El cliente no tiene Ubigeo ó Urbanizacion registrado, Favor de verificar!!!", vbCritical, sMensaje
                            Exit Sub
                        End If
                    End If
                    
          Else
             'MsgBox "Proceso Cancelado", vbCritical, sMensaje
             Exit Sub
          End If
           
           
       End If
       If Pedido = "" Then
          'GeneraPedido
            If lPrinter And lObligaPrinter Then
               i = Calcular("select count(tCodigoPedido) as codigo from " & sDetalle & " where lImprime=0", Cn)
               If i > 0 Then
                  cmdOpcion_Click (8)
                  If variableEmite = False Then: Exit Sub
               End If
            Else
               GeneraPedido
            End If
          
       Else
          sPedido = Pedido

          'IMPRESION DE PRODUCTOS NO ENVIADOS
          If lPrinter And lObligaPrinter Then
            i = Calcular("select count(tCodigoPedido) as codigo from " & sDetalle & " where lImprime=0", Cn)
            If i > 0 Then
               cmdOpcion_Click (8)
            End If
          End If
          
          ActualizaPedido

          Cn.Execute "update MPEDIDO set tObservacion='" & txtObservacion.Caption & "' where tCodigoPedido='" & sPedido & "'"
          
       End If

       'impresion imagen
       Set rstFuente = New ADODB.Recordset
       imageCab.Picture = Nothing
       imagepIE.Picture = Nothing
       Set rstFuente = Lib.OpenRecordset("select iImagenCabDoc AS foto, iImagenPieDoc as fotoPie  from tcaja where tcaja='" & sCaja & "'", Cn)
       imageCab.DataField = "foto"
       Set imageCab.DataSource = rstFuente
       imagepIE.DataField = "fotoPie"
       Set imagepIE.DataSource = rstFuente


       'Genera y Actualiza los Numero de Documento
       sSerie = RsTipoDocumento!tSerie
       sCorrela = Lib.Correlativo(RsTipoDocumento!tUltimoNumero, 9)
       sPrefijo = RsTipoDocumento!prefijo
       sTipoDocumento = RsTipoDocumento!TTipoEmision
       sImp = RsTipoDocumento!timpresora
       sDocumento = sPrefijo & sSerie & sCorrela
       sResumen = RsTipoDocumento!lResumen
       
       'Cn.Execute "Update TTIPODOCUMENTOIMPRESORA Set tUltimoNumero = '" & sCorrela & "' where tTipoEmision ='" & sTipoDocumento & "' and tCaja ='" & sCaja & "'"
       
       'Inserta Detalle de Documento
       Isql = "Insert into DDOCUMENTO " & _
              "       ( tDocumento, tItem, tCodigoPedido, tCodigoProducto, nPrecioNeto, nPrecioImpuesto1, nPrecioImpuesto2, nPrecioImpuesto3, " & _
              "nPrecioVenta, nRecargo, nDescuento, nCantidad, nPrecioOficial, nImpuesto1, nImpuesto2, nImpuesto3, nVenta ) " & _
              "select  '" & sDocumento & "' as tDocumento , tItem, tCodigoPedido, tCodigoProducto, nPrecioNeto, nPrecioImpuesto1, nPrecioImpuesto2, nPrecioImpuesto3, " & _
              "nPrecioVenta, nRecargo, nDescuento, nCantidad, nPrecioOficial, nImpuesto1, nImpuesto2, nImpuesto3, nVenta From DPEDIDO " & _
              "where (isnull(tFacturado,'0') = '0' or len(ltrim(tFacturado)) = 0) and tCodigoPedido ='" & Pedido & "'"
       Cn.Execute Isql

       'Calcula el total de la cabecera
       Set RsSuma = Lib.OpenRecordset("select sum(nPrecioNeto*nCantidad) as nNeto, sum(nImpuesto1) as nImpuesto1, sum(nImpuesto2) as nImpuesto2, sum(nImpuesto3) as nImpuesto3, sum(nVenta) as nVenta " & _
                                      " from DPEDIDO where (isnull(tFacturado,'0') = '0' or len(ltrim(tFacturado)) = 0) and tCodigoPedido ='" & Pedido & "' group by tCodigoPedido ", Cn)

       'Inserta el Documento
       nCargo = Round(RsSuma!nVenta, 2)

       Select Case pais
            Case "001"
                    tAutorizacion = obtieneAutorizacionDosificacion(sCaja, "1")
                    tDosificacion = obtieneAutorizacionDosificacion(sCaja, "2")
                    If tAutorizacion <> "" And tDosificacion <> "" Then
                        tcodigoControl = devuelveCodigoControl(sCaja, sCorrela, tAutorizacion, tDosificacion, sCliente, nCargo)
                        If tcodigoControl = "" Then
                           MsgBox "Error al generar Código de Control", vbCritical, sMensaje
                           Exit Sub
                        End If
                    Else
                        MsgBox "Error al obtener Número de Autorización o Dosificación. Verifique.", vbCritical, sMensaje
                        Exit Sub
                    End If
                    
            Case "002" 'Ecuador
                    tAutorizacion = RsTipoDocumento!tNumeroAutorizacion

            Case Else 'Peru
                tAutorizacion = ""
                tcodigoControl = ""
                tDosificacion = ""

       End Select

       If lAplicablePedido Then
          nTotalDescuento = 0
       Else
          nTotalDescuento = Calcular("select sum(nDescuento*nCantidad) as Codigo From " & sDetalle, Cn)
       End If
       Isql = "Insert into MDOCUMENTO " & _
              "     ( tDocumento, tTipoDocumento, tCortesia, tcodigoCliente, tEstadoDocumento, tCaja, tTurno, nNeto, nRecargo, nDescuento, nPrecioOficial, nPrecioImpuesto1, nPrecioImpuesto2, nPrecioImpuesto3, nVenta, tSalon, tUsuario, tUsuarioAutoriza, fRegistro,TAUTORIZACION,TCODIGOCONTROL, fdiacontable, tDescuento, tConsumo, lImpresionMonedaExtranjera) " & _
              "Values(   '" & sDocumento & "', " _
                      & "'" & sTipoDocumento & "', " _
                      & "'" & sCortesia & "', " _
                      & "'" & IIf(sCliente = "", "", sCliente) & "', " _
                      & "'01', " _
                      & "'" & sCaja & "', " _
                      & "'" & sTurno & "', " _
                      & RsSuma!nNeto & ", " _
                      & "0, " & nTotalDescuento & ", 0, " _
                      & RsSuma!nImpuesto1 & ", " _
                      & RsSuma!nImpuesto2 & ", " _
                      & RsSuma!nImpuesto3 & ", " _
                      & RsSuma!nVenta & ", " _
                      & "'" & sSalon & "', " _
                      & "'" & Mid(sUsuario, 1, 15) & "', " _
                      & "'" & Mid(sUsuarioAutoriza, 1, 15) & "', " _
                       & "getdate(),'" & tAutorizacion & "','" & tcodigoControl & "','" & Format(obtieneDiaContable, "yyyyMMdd") & "','" & IIf(lAplicablePedido, "", sCodigoDescuento) & "', '" & sDetalleConsumo & "', " & IIf(TimpresionDolaresDelivery, 1, 0) & " ) "
       Cn.Execute Isql
       

       Cn.Execute "Update TTIPODOCUMENTOIMPRESORA Set tUltimoNumero = '" & sCorrela & "' where tTipoEmision ='" & sTipoDocumento & "' and tCaja ='" & sCaja & "'"
       
       wEnter = True
       If lPagoAntesImpresion Then
       
            If RsTipoDocumento!TTipoEmision <> "00" Then
               Cn.Execute "Update MPEDIDO set tEstadoPedido = '02'  where tCodigoPedido ='" & Pedido & "'"
               xTipo = ""
               sFormulario = "CajaRapida"
               If sFormulario = "CajaRapida" And lPagoRapido = True Then
                  frmPagoRapido.Show vbModal
                  
                    If wEnter = False Then
                        Cn.Execute "Update TMESA set tEstadoMesa = '02' where tCodigoMesa ='" & sMesa & "'"
                        Cn.Execute "Update MPEDIDO set tEstadoPedido = '01', lReplica = 0  where tCodigoPedido ='" & Pedido & "'"
                        Cn.Execute "Delete MDOCUMENTO Where tDocumento= '" & sDocumento & "'"
                        Cn.Execute "Delete DDOCUMENTO Where tDocumento= '" & sDocumento & "'"
                        Cn.Execute "Delete DPAGODOCUMENTO Where tDocumento= '" & sDocumento & "'"
                        
                        xUltimoCorrelativo = Calcular("select MAX(tDocumento) as codigo from MDOCUMENTO where tcaja='" & sCaja & "' and tTipoDocumento='" & sTipoDocumento & "'", Cn)
                        xUltimoCorrelativo = Right(xUltimoCorrelativo, 9)
                    
                        Cn.Execute "Update TTIPODOCUMENTOIMPRESORA Set tUltimoNumero = '" & xUltimoCorrelativo & "' where tTipoEmision ='" & sTipoDocumento & "' and tCaja ='" & sCaja & "'"
                        Exit Sub
                    End If
               Else
                  frmPago.Show vbModal
                  
                    If wEnter = False Then
                        Cn.Execute "Update TMESA set tEstadoMesa = '02' where tCodigoMesa ='" & sMesa & "'"
                        Cn.Execute "Update MPEDIDO set tEstadoPedido = '01', lReplica = 0  where tCodigoPedido ='" & Pedido & "'"
                        Cn.Execute "Delete MDOCUMENTO Where tDocumento= '" & sDocumento & "'"
                        Cn.Execute "Delete DDOCUMENTO Where tDocumento= '" & sDocumento & "'"
                        Cn.Execute "Delete DPAGODOCUMENTO Where tDocumento= '" & sDocumento & "'"
                        
                        xUltimoCorrelativo = Calcular("select MAX(tDocumento) as codigo from MDOCUMENTO where tcaja='" & sCaja & "' and tTipoDocumento='" & sTipoDocumento & "'", Cn)
                        xUltimoCorrelativo = Right(xUltimoCorrelativo, 9)
                    
                        Cn.Execute "Update TTIPODOCUMENTOIMPRESORA Set tUltimoNumero = '" & xUltimoCorrelativo & "' where tTipoEmision ='" & sTipoDocumento & "' and tCaja ='" & sCaja & "'"
                        Exit Sub
                    End If
               End If
            End If
       
       End If
       
            '-----------------------
           If pais = "000" And lFEpape And IIf(RsTipoDocumento!lFacturacionElectronica = True, 1, 0) Then
                If Not FacturarTCPIP(2, sDocumento, 0) Then
                    Cn.Execute "Update TMESA set tEstadoMesa = '02' where tCodigoMesa ='" & sMesa & "'"
                    Cn.Execute "Update MPEDIDO set tEstadoPedido = '01', lReplica = 0  where tCodigoPedido ='" & sPedido & "'"
                    Cn.Execute "Delete MDOCUMENTO Where tDocumento= '" & sDocumento & "'"
                    Cn.Execute "Delete DDOCUMENTO Where tDocumento= '" & sDocumento & "'"
                    Cn.Execute "Delete DPAGODOCUMENTO Where tDocumento= '" & sDocumento & "'"
                    
                    xUltimoCorrelativo = Calcular("select MAX(tDocumento) as codigo from MDOCUMENTO where tcaja='" & sCaja & "' and tTipoDocumento='" & sTipoDocumento & "'", Cn)
                    xUltimoCorrelativo = Right(xUltimoCorrelativo, 9)
                
                    Cn.Execute "Update TTIPODOCUMENTOIMPRESORA Set tUltimoNumero = '" & xUltimoCorrelativo & "' where tTipoEmision ='" & sTipoDocumento & "' and tCaja ='" & sCaja & "'"
                    
                    wEnter = False
                    MsgBox "Ocurrio un Problema al Procesar el Documento!!!", vbInformation, sMensaje
                   Exit Sub
                End If
           End If
           '------------------------
       'Actualiza Base de Datos Detalle del Pedido
        If sPrefijo = "0" Then
           Cn.Execute "Update DPEDIDO set tFacturado = 'C', tDocumento = '" & sDocumento & "' where tCodigoPedido ='" & Pedido & "' and (isnull(tFacturado,'0')='0' or len(ltrim(tfacturado))=0)"
           Cn.Execute "Update MPEDIDO set tEstadoPedido = '02' where tCodigoPedido = '" & Pedido & "'"
           Cn.Execute "Update MDOCUMENTO set tEstadoDocumento ='02',lreplica=1  where tDocumento = '" & sDocumento & "'"
        Else
        '// cambio realizado el 05/05/2018 ELDC - actualiza a "P" en Pago antes de Impresion
            If lPagoAntesImpresion Then
                Cn.Execute "Update DPEDIDO set tFacturado = 'P', tDocumento = '" & sDocumento & "' where tCodigoPedido ='" & Pedido & "' and (isnull(tFacturado,'0')='0' or len(ltrim(tFacturado))=0)"
            Else
                Cn.Execute "Update DPEDIDO set tFacturado = 'F', tDocumento = '" & sDocumento & "' where tCodigoPedido ='" & Pedido & "' and (isnull(tFacturado,'0')='0' or len(ltrim(tFacturado))=0)"
            End If
           'Cn.Execute "Update DPEDIDO set tFacturado = 'F', tDocumento = '" & sDocumento & "' where tCodigoPedido ='" & Pedido & "' and (isnull(tFacturado,'0')='0' or len(ltrim(tFacturado))=0)"
           Cn.Execute "Update DPEDIDO set tFacturado = 'C' where tDocumento ='" & sDocumento & "' and len(ltrim(tCortesia)) = 4 "
           Cn.Execute "Update MPEDIDO set tEstadoPedido = '02' where tCodigoPedido ='" & Pedido & "'"
         '// fin de cambio
        End If
                
        'LOG
        If lLogCajaRapida Then
            Cn.Execute "INSERT INTO TLOG_IMPRESION (TDOCUMENTO,TPOSICION1) VALUES('" & sDocumento & "','CREACION DOCUMENTO')"
        End If
        
        'PARA NO FISCALES
        Cn.Execute "UPDATE DPEDIDO SET lregistroventa=(select case when registroventa=0 then 0 else 1 end from vtipodocumento where codigo='" & sTipoDocumento & "') where  tCodigoPedido ='" & Pedido & "' and tDocumento ='" & sDocumento & "'"
        
        If lInfhotel Then
           Dim xSuma As Double
           xSuma = Calcular("select sum(nVenta) as Codigo FROM DPEDIDO where tEstadoItem = 'N' and tDocumento ='" & sDocumento & "' and tCodigoPedido='" & Pedido & "'", Cn)
           
           If sComandaInfhotel = "" Then
              sComandaInfhotel = Calcular("select left(MAX(tComanda),8) as Codigo from MCOMANDA where tPuntoVenta='" & sPuntoVenta & "'", CnInfhotel)
              sComandaInfhotel = Lib.Correlativo(sComandaInfhotel, 8)
              CnInfhotel.Execute "Update TPUNTOVENTA Set nUltimoComanda = '" & sComandaInfhotel & "' where tPuntoVenta='" & sPuntoVenta & "'"
              sComandaInfhotel = sComandaInfhotel & "-" & UCase(Mid(rsPuntoVenta!Descripcion, 1, 3))
              rsPuntoVenta.Requery
              rsPuntoVenta.MoveFirst
              rsPuntoVenta.Find "Codigo='" & sPuntoVenta & "'"
                                          
              'Genero las comandas en Infhotel
              'Cabecera
              Isql = "Insert into MCOMANDA " & _
                     "(tComanda, tPuntoVenta, tHotel, nMovimiento, fFecha, hHora, nTotal, tEstado, " & _
                     "tEmitido, tAsignacion, tCodigoReserva, tNumeroHabitacion, tCodigoFuncionario, " & _
                     "tCaja, tDocumento, tUsuario, nTCambio, tCodigoCompania, tCliente, tMoneda, fFechaE, hHoraE, tUsuarioE,TNOTAPEDIDO) " & _
                     "values('" & sComandaInfhotel & "', '" & sPuntoVenta & "', '" & sHotel & "', 1,  getdate(), getdate(), " & xSuma & ", '01', " & _
                     "1, '" & IIf(RsTipoDocumento!TTipoEmision = "00", "05", "01") & "', '', '', '" & IIf(RsTipoDocumento!TTipoEmision = "00", Mid(sCortesia, 3, 2), "") & "', " & _
                     "'" & sCajaInfhotel & "', '" & IIf(pais = "002", Mid(sDocumento, 1, 1) + Mid(sDocumento, 3), sDocumento) & "', '" & xUsuario & "', " & nTC & ", '', '" & sPasajero & "', '01', getdate(), getdate(), '" & xUsuario & "','" & Pedido & "')"
              CnInfhotel.Execute Isql
           Else
              'sComandaInfhotel = RsCabecera!tComanda
              CnInfhotel.Execute "update MCOMANDA set TASIGNACION='" & IIf(RsTipoDocumento!TTipoEmision = "00", "05", "01") & "', TCODIGORESERVA='', TNUMEROHABITACION='', TCLIENTE='', nTotal= " & xSuma & ", tEstado='01' " & _
                                 "where tComanda ='" & sComandaInfhotel & "' and tPuntoVenta='" & sPuntoVenta & "'"
           End If
           
           'Detalle
           Dim nMovimiento As Integer
           CnInfhotel.Execute "delete from DCOMANDA where tComanda ='" & sComandaInfhotel & "' and tPuntoVenta='" & sPuntoVenta & "'"
           nMovimiento = Calcular("select max(nmovimiento) as codigo from dcomanda where tcomanda='" & sComandaInfhotel & "'", CnInfhotel) + 1
           
           Isql = "Insert into DCOMANDA " & _
                  "(tComanda, tPuntoVenta, tHotel, tItem, nMovimiento, tNotaPedido, tCodigoItem, " & _
                  "nPrecioUnitario, nCantidad, nTotal, nPrecioCos, tCodigoReserva, tNumeroHabitacion, " & _
                  "tCuenta, tCaja, tDocumento, tAsignado, tUsuario, fFecha, hHora) " & _
                  "select '" & sComandaInfhotel & "' as tComanda, '" & sPuntoVenta & "' as tPuntoVenta, '" & sHotel & "' as tHotel, tItem , " & nMovimiento & ", '" & Pedido & "' as tNotaPedido, tInfhotel as tCodigoItem, " & _
                  "T1.nPrecioVenta as nPrecioUnitario, nCantidad, nVenta as nTotal, T1.nInsumo+T1.nGasto+T1.nManoObra as nPrecioCos, '" & sReserva & "' as tCodigoReserva, '" & sHabitacion & "' as tNumeroHabitacion, " & _
                  "'' as tCuenta, '" & sCajaInfhotel & "' as tCaja, '" & sDocumento & "' as tDocumento, '" & IIf(sDescrip = "Reserva", "03", "02") & "' as tAsignado, '" & xUsuario & "' as  tUsuario, getdate() as fFecha, getdate() as hHoraMovimiento " & _
                  "FROM OPENROWSET('SQLOLEDB','" & sRuta & "';'" & sUserName & "';'" & sUserPassword & "', " & sMDB & ".dbo.DPEDIDO) T1 INNER JOIN OPENROWSET('SQLOLEDB','" & sRuta & "';'" & sUserName & "';'" & sUserPassword & "', " & sMDB & ".dbo.TPRODUCTO) T2 ON T1.tCodigoProducto = T2.tCodigoProducto " & _
                  "where tDocumento='" & sDocumento & "' and tCodigoPedido ='" & Pedido & "'"
           CnInfhotel.Execute Isql
           Cn.Execute "update MPEDIDO set tComanda = '" & sComandaInfhotel & "', tPuntoVenta='" & sPuntoVenta & "'  where tCodigoPedido='" & Pedido & "'"
           sComandaInfhotel = ""
        End If
        
        
       'Imprime Documentos
           'Imprime documentos
           If wConsumo = False And lDescripcionAlternativa = True Then
              If validaImpresionAlternativa(sDocumento) = False Then
                    If MsgBox("Desea imprimir descripción Alternativa? ", vbQuestion + vbYesNo + vbDefaultButton2, sMensaje) = vbYes Then
                          lImprimeAlternativa = True
                    End If
              End If
           End If
             '-------------------------------  SE INTRDUJO LAS CONSULTAS A  UN STORE PROCEDURE -------------------------
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
            '-------------------------------------------------------------------------------------------------------------
        'FACTURACION_E_PERU
        Set RsImpDocumentoE = Lib.OpenRecordset(IsqlFact, Cn)
        xImpresionFE = Calcular(" SELECT tImpresionFE as codigo FROM vTipodocumento WHERE Codigo=(SELECT TTIPODOCUMENTO FROM MDOCUMENTO WHERE tDocumento='" & sDocumento & "')", Cn)
        xImpresioDE = Mid(sDocumento, 1, 1)
        '---------------------------------------
        
        Set RsImpresion = Lib.OpenRecordset(Isql, Cn)
        
        'Log
        If lLogCajaRapida Then
            Cn.Execute "UPDATE TLOG_IMPRESION SET TPOSICION2='RECORDSET IMPRESION CARGADO ' + '" & RsImpresion.RecordCount & "' WHERE TDOCUMENTO='" & sDocumento & "'"
        End If
        
        imprimeDolaDocumentos = Calcular("select isnull(lequivadolares,0) as codigo from vtipodocumentoimpresora where tcaja='" & sCaja & "' and ttipoemision='" & sTipoDocumento & "' ", Cn)
        If imprimeDolaDocumentos = "Verdadero" Then
            lDocumEquivaPrecuenta = True
        Else
            lDocumEquivaPrecuenta = False
        End If
        
        
       If RsImpresion.RecordCount = 0 Then
          LimpiaRs
          MsgBox "No existen Datos a Imprimir", vbExclamation, sMensaje
       Else
            'SUNAT
            numeroSerieImpresora = obtieneNumeroSerieImpresora(sCaja, sImp)
            codigoImpresora = sImp
            'SUNAT
            Cn.Execute " update mdocumento set timpresora='" & codigoImpresora & "', tSerieImpresora='" & numeroSerieImpresora & "' where tdocumento ='" & sDocumento & "' "

            'Log
            If lLogCajaRapida Then
            Cn.Execute "UPDATE TLOG_IMPRESION SET TPOSICION3='IMPRESORA' + '" & sImp & "'+' CAJA' +'" & sCaja & "' WHERE TDOCUMENTO='" & sDocumento & "'"
            End If

               'FACTURACION_E_PERU
               If pais = "000" Then
                   If lFacturacionE Then
                   
                       If lFEOfisis Then
                                '----CABECERA
                                Set oComandoCabeceraOfisis = New clsComando
                                If Not oComandoCabeceraOfisis.CreateCmdSp("USP_FactDocumentoOfisis", Cn) Then
                                     Set oComandoCabeceraOfisis = Nothing
                                     Exit Sub
                                End If
                                oComandoCabeceraOfisis.CreateParameter "@CodigoDocumento", adVarChar, adParamInput, 20, sDocumento
    
                                If Not oComandoCabeceraOfisis.GetParamOK Then
                                     Set oComandoCabeceraOfisis = Nothing
                                     Exit Sub
                                End If
                                If Not oComandoCabeceraOfisis.ExecSP Then
                                     Set oComandoCabeceraOfisis = Nothing
                                     Exit Sub
                                End If

                                '----FIRMA DOCUMENTO OFISIS
                                If RsTipoDocumento!lDocumentoElectronicoOfisis Then
                                    Set oComandoFirmaDocumentoOfisis = New clsComando
                                    If Not oComandoFirmaDocumentoOfisis.CreateCmdSp("USP_FactFirmaDocumentoOfisis", Cn) Then
                                         Set oComandoFirmaDocumentoOfisis = Nothing
                                         Exit Sub
                                    End If
                                    oComandoFirmaDocumentoOfisis.CreateParameter "@CodigoDocumento", adVarChar, adParamInput, 20, sDocumento
        
                                    If Not oComandoFirmaDocumentoOfisis.GetParamOK Then
                                         Set oComandoFirmaDocumentoOfisis = Nothing
                                         Exit Sub
                                    End If
                                    If Not oComandoFirmaDocumentoOfisis.ExecSP Then
                                         Set oComandoFirmaDocumentoOfisis = Nothing
                                         Exit Sub
                                    End If
                                    
                                    'VALIDAR RESPUESTA CODIGO DE BARRA
                                    fDocumento = Mid(sDocumento, 1, 1) + Mid(sDocumento, 4, 3) + "-" + CStr(CLng(Mid(sDocumento, 8, 8)))
                                    Sleep 3000
                                    If lImpresionCodigoBarras Then
                                        imageHash.DataField = "foto"
                                        Set RsCodigoHash = Lib.OpenRecordset("USP_FactObtenerCodigoBarraOfisis '" & fDocumento & "','" & Mid(sDocumento, 1, 1) & "','' ", Cn)
                                        Set imageHash.DataSource = RsCodigoHash
                                        
                                    ElseIf lQRFE Then
                                        Set imageHash.Picture = LoadPicture(ImagenQR_Ofisis(fDocumento, sDocumento))
                                    Else
                                        Set RscadenaCodigoHash = Lib.OpenRecordset("USP_FactConsultaHash '" & fDocumento & "','0' ", Cn)
                                        If RscadenaCodigoHash.RecordCount > 0 Then
                                            cadenaCodigoHash = RscadenaCodigoHash!codigo
                                        End If
                                        'cadenaCodigoHash = Calcular("select CO_HASH as codigo from TCFACT_ELEC where NU_DOCU='" & fDocumento & "' and (TI_DOCU='B' or TI_DOCU ='F')", CnFE)
                                    End If
                                End If
                               
                               
                       ElseIf lFESpring Then
                            If IIf(RsTipoDocumento!lFacturacionElectronica = True, 1, 0) Then

                                 frmMensajeFeSpring.sDocumento = sDocumento
                                 frmMensajeFeSpring.oVenta = 4 ' 3: "Formulario Caja Rapida"
                                 frmMensajeFeSpring.Show vbModal
                                 If frmMensajeFeSpring.lEnvio = False Then
                                    Exit Sub
                                 End If
                                                     
                                 'VALIDAR RESPUESTA DE CODIGO HASH Y CODIGO DE BARRA
                                 fDocumento = Mid(sDocumento, 1, 1) + Mid(sDocumento, 4, 3) + Mid(sDocumento, 8, 8)
                                 If tCodigoFE = "000" Then
                                     If lQRFE Then
                                         If frmMensajeFeSpring.lQrInf Then
                                            Set imageHash.Picture = LoadPicture(ImagenQR(sDocumento))
                                         Else
                                            Set imageHash.Picture = LoadPicture(ImagenFeSpring(lQRFE, sDocumento))
                                         End If
                                     Else
                                         If lImpresionCodigoBarras Then
                                             
                                         Else
                                             cadenaCodigoHash = ImagenFeSpring(lQRFE, sDocumento)
                                         End If
                                     End If
                                 End If
                             End If
                               
                             
                       ElseIf lFECarbajal Then
                            Label5.Caption = "   Proceso de envio de documento a InfoFact......."
                            lblPaso1.Caption = "Enviando información de documento a InfoFact."
                            lblPaso2.Caption = "Obteniendo codigo " & IIf(lQRFE, "QR", IIf(lImpresionCodigoBarras, "de barras", " hash")) & " almacenado."
                            If IIf(RsTipoDocumento!lFacturacionElectronica = True, 1, 0) Then
                                Dim sImporteLetra As String
                                sImporteLetra = NumeroCadena(str(RsImpDocumentoE!nVenta)) + " " + sMonedaN
                                FrameFeSpring.Visible = True
                                lblPaso1.Visible = True
                                lblPaso2.Visible = True
                                imgProceso(0).Visible = False
                                imgProceso(1).Visible = False
                                imgProceso(2).Visible = False
                                imgProceso(3).Visible = False
                                Sleep 1000
                                If Not INSERTAFE_CARVAJAL(sDocumento, sImporteLetra, 0, 0) Then '----CABECERA
                                        Cn.Execute "Update TMESA set tEstadoMesa = '02' where tCodigoMesa ='" & sMesa & "'"
                                        Cn.Execute "Update MPEDIDO set tEstadoPedido = '01', lReplica = 0  where tCodigoPedido ='" & sPedido & "'"
                                        Cn.Execute "Delete MDOCUMENTO Where tDocumento= '" & sDocumento & "'"
                                        Cn.Execute "Delete DDOCUMENTO Where tDocumento= '" & sDocumento & "'"
                                        Cn.Execute "Delete DPAGODOCUMENTO Where tDocumento= '" & sDocumento & "'"
                                        xUltimoCorrelativo = Calcular("select MAX(tDocumento) as codigo from MDOCUMENTO where tcaja='" & sCaja & "' and tTipoDocumento='" & sTipoDocumento & "'", Cn)
                                        xUltimoCorrelativo = Right(xUltimoCorrelativo, 9)
                                        Cn.Execute "Update TTIPODOCUMENTOIMPRESORA Set tUltimoNumero = '" & xUltimoCorrelativo & "' where tTipoEmision ='" & sTipoDocumento & "' and tCaja ='" & sCaja & "'"
                                        Cn.Execute "Update DPEDIDO set tFacturado = NULL, tDocumento = NULL where tCodigoPedido ='" & sPedido & "' "
                                        Cn.Execute "Update MPEDIDO set tEstadoPedido = '01', lReplica=1 where tCodigoPedido = '" & sPedido & "'"
                                        Cn.Execute "UPDATE DPEDIDO SET lregistroventa = NULL where tCodigoPedido ='" & sPedido & "' and  tDocumento ='" & sDocumento & "'"
                                        imgProceso(2).Visible = True
                                        imgProceso(3).Visible = True
                                        Sleep 1000
                                        FrameFeSpring.Visible = False
                                        Exit Sub
                                 End If
                                 imgProceso(0).Visible = True
                                 'VALIDAR RESPUESTA DE CODIGO HASH Y CODIGO DE BARRA
                                 fDocumento = Mid(sDocumento, 1, 1) + Mid(sDocumento, 4, 3) + Mid(sDocumento, 8, 8)
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
                                 imgProceso(1).Visible = True
                                 Sleep 1000
                                 FrameFeSpring.Visible = False
                            End If
                                
                       ElseIf lFEpape Then
                            If IIf(RsTipoDocumento!lFacturacionElectronica = True, 1, 0) Then
                                If tCodigoFE = "000" Then
                                     If lQRFE Then
                                         Set imageHash.Picture = LoadPicture(CrearImagenQR(PapeTermico))
                                     Else
                                         If lImpresionCodigoBarras Then
                                             'Set imageHash.Picture = LoadPicture(ImagenQR(sDocumento))
                                         Else
                                             cadenaCodigoHash = PapeMatricial
                                         End If
                                     End If
                                 End If
                            End If
                       ElseIf lFEBiz Then
                       
                            If IIf(RsTipoDocumento!lFacturacionElectronica = True, 1, 0) Then
                                If Not INSERTA_FE_INFOREST(sDocumento, 1, DateTime.Date) Then '----CABECERA
                                    Cn.Execute "Update TMESA set tEstadoMesa = '02' where tCodigoMesa ='" & sMesa & "'"
                                    Cn.Execute "Update MPEDIDO set tEstadoPedido = '01', lReplica = 0  where tCodigoPedido ='" & sPedido & "'"
                                    Cn.Execute "Delete MDOCUMENTO Where tDocumento= '" & sDocumento & "'"
                                    Cn.Execute "Delete DDOCUMENTO Where tDocumento= '" & sDocumento & "'"
                                    Cn.Execute "Delete DPAGODOCUMENTO Where tDocumento= '" & sDocumento & "'"
                                    xUltimoCorrelativo = Calcular("select MAX(tDocumento) as codigo from MDOCUMENTO where tcaja='" & sCaja & "' and tTipoDocumento='" & sTipoDocumento & "'", Cn)
                                    xUltimoCorrelativo = Right(xUltimoCorrelativo, 9)
                                    Cn.Execute "Update TTIPODOCUMENTOIMPRESORA Set tUltimoNumero = '" & xUltimoCorrelativo & "' where tTipoEmision ='" & sTipoDocumento & "' and tCaja ='" & sCaja & "'"
                                    Cn.Execute "Update DPEDIDO set tFacturado = NULL, tDocumento = NULL where tCodigoPedido ='" & sPedido & "' "
                                    Cn.Execute "Update MPEDIDO set tEstadoPedido = '01', lReplica=1 where tCodigoPedido = '" & sPedido & "'"
                                    Cn.Execute "UPDATE DPEDIDO SET lregistroventa = NULL where tCodigoPedido ='" & sPedido & "' and  tDocumento ='" & sDocumento & "'"
                                    GoTo fin
                                 End If
                                 'VALIDAR RESPUESTA DE CODIGO HASH Y CODIGO DE BARRA
                                 fDocumento = Mid(sDocumento, 1, 1) + Mid(sDocumento, 4, 3) + Mid(sDocumento, 8, 8)
                                 If tCodigoFE = "000" Then
                                     If lQRFE Then
                                         Set imageHash.Picture = LoadPicture(QRHASH_FE_INFOREST(3, sDocumento, 0))
                                     Else
                                         If lImpresionCodigoBarras Then
                                             Set imageHash.Picture = LoadPicture(QRHASH_FE_INFOREST(1, sDocumento, 0))
                                         Else
                                             cadenaCodigoHash = QRHASH_FE_INFOREST(2, sDocumento, 0)
                                         End If
                                     End If
                                 End If
                            End If
                        ElseIf lFEGesa Then
                            If IIf(RsTipoDocumento!lFacturacionElectronica = True, 1, 0) Then
                                If Not INSERTAFE(sDocumento, "", 1, "") Then '----CABECERA
                                    MsgBox "No se pudo enviar el documento a facturacion electronica", vbInformation
                                End If
                                If lQRFE Then
                                    Set imageHash.Picture = LoadPicture(QRHASH_FE_INFOREST(3, sDocumento, 0))
                                Else
                                    If lImpresionCodigoBarras Then
                                        Set imageHash.Picture = LoadPicture(QRHASH_FE_INFOREST(1, sDocumento, 0))
                                    Else
                                        cadenaCodigoHash = QRHASH_FE_INFOREST(2, sDocumento, 0)
                                    End If
                                End If
                            End If
                       Else 'INFOFACT
                            If IIf(RsTipoDocumento!lFacturacionElectronica = True, 1, 0) Then
                                If Not INSERTAFE(sDocumento, "", 1, "") Then '----CABECERA
                                    Exit Sub
                                End If
                                If RsImpDocumentoE!Ruc <> "" Then
                                    If Not INSERTAFE(sDocumento, "", 2, RsImpDocumentoE!Ruc) Then '----CLIENTE
                                        Exit Sub
                                    End If
                                End If
                                'VALIDAR RESPUESTA DE CODIGO HASH Y CODIGO DE BARRA
                                fDocumento = Mid(sDocumento, 1, 1) + Mid(sDocumento, 4, 3) + Mid(sDocumento, 8, 8)
                                If tCodigoFE = "000" Then
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
            End If
            '---------------------------------------


            'Configura la Impresora
            Imprimir (sImp)
            Printer.FontName = sFont
            Printer.FontBold = False

            If wConsumo Then
               If RsTipoDocumento!tFormulario = "01" Then
                      If RsTipoDocumento!Cliente Then
                         'FACTURACION ELECTRONICA
                         If IIf(RsTipoDocumento!lFacturacionElectronica = True, 1, 0) And tCodigoFE <> "999" Then
                              ImprimeFacturaConsumoElectronico RsImpresion, sDetalleConsumo, imageHash, sTipoDocumento, imageCab, imagepIE, cadenaCodigoHash, TimpresionDolaresDelivery
                         Else
                              ImprimeFacturaConsumoT RsImpresion, sDetalleConsumo, sTipoDocumento, imageCab, imagepIE, TimpresionDolaresDelivery
                         End If
                         
                         NFactura = sCorrela
                         frmVenta.lblFactura.Caption = NFactura
                      Else
                          'FACTURACION ELECTRONICA
                          If IIf(RsTipoDocumento!lFacturacionElectronica = True, 1, 0) And tCodigoFE <> "999" Then
                              ImprimeBoletaConsumoElectronico RsImpresion, sDetalleConsumo, imageHash, sTipoDocumento, imageCab, imagepIE, cadenaCodigoHash, TimpresionDolaresDelivery
                          Else
                              ImprimeBoletaConsumoT RsImpresion, sDetalleConsumo, sTipoDocumento, imageCab, imagepIE, TimpresionDolaresDelivery
                          End If
                      End If
               Else
                      If lFacturacionE Then
                                                                  
                              If Generar_Imagen(CnFE, "select imagen from IMAGENCODIGOBARRA where nro_efact='" & fDocumento & "'", "imagen", "\fact.bmp") = True Then
                                  ImprimeFormatoAConsumo sDocumento
                                  Kill App.Path & "\fact.bmp"
                              Else
                                  ImprimeFormatoAConsumo sDocumento
                              End If
                      Else
                      
                          If RsTipoDocumento!Cliente Then
                             ImprimeFacturaConsumoN RsImpresion, sDetalleConsumo, sEmpresa
                             NFactura = sCorrela
                             frmVenta.lblFactura.Caption = NFactura
                          Else
                             ImprimeBoletaConsumoN RsImpresion, sDetalleConsumo, sEmpresa
                          End If
                      End If
                  
               End If
             
          Else
               If RsTipoDocumento!tFormulario = "01" Then
                   If RsTipoDocumento!Cliente Then
                      
                       'FACTURACION ELECTRONICA
                       If IIf(RsTipoDocumento!lFacturacionElectronica = True, 1, 0) And tCodigoFE <> "999" Then
                          ImprimeFacturaElectronica RsImpresion, imageHash, sTipoDocumento, imageCab, imagepIE, cadenaCodigoHash, TimpresionDolaresDelivery
                       Else
                            'Log
                            If lLogCajaRapida = True Then
                             Cn.Execute "UPDATE TLOG_IMPRESION SET TPOSICION4='INIC IMPRE ' + '" & RsImpresion.RecordCount & "' +'/' +'" & sTipoDocumento & "' WHERE TDOCUMENTO='" & sDocumento & "'"
                            End If
                
                               ImprimeFacturaT RsImpresion, sTipoDocumento, imageCab, imagepIE, TimpresionDolaresDelivery
                           If lLogCajaRapida = True Then
                               Cn.Execute "UPDATE TLOG_IMPRESION SET TPOSICION7='FIN IMPRE ' + '" & RsImpresion.RecordCount & "' +'/' +'" & sTipoDocumento & "' WHERE TDOCUMENTO='" & sDocumento & "'"
                           End If
                       End If
                      
                      NFactura = sCorrela
                      frmVenta.lblFactura.Caption = NFactura
                      
                   ElseIf RsTipoDocumento!TTipoEmision = "00" Then
                      If MsgBox("Deseas imprimir el Voucher", vbQuestion + vbYesNo, sMensaje) = vbYes Then
                         ImprimeCortesia RsImpresion, sTipoDocumento, imageCab, imagepIE
                      End If
                      
                   Else
                
                       'FACTURACION ELECTRONICA
                       If IIf(RsTipoDocumento!lFacturacionElectronica = True, 1, 0) And tCodigoFE <> "999" Then
                          ImprimeBoletaElectronica RsImpresion, imageHash, sTipoDocumento, imageCab, imagepIE, cadenaCodigoHash, TimpresionDolaresDelivery
                       Else
                       'Log
                       If lLogCajaRapida = True Then
                             Cn.Execute "UPDATE TLOG_IMPRESION SET TPOSICION4='INIC IMPRE ' + '" & RsImpresion.RecordCount & "' +'/' +'" & sTipoDocumento & "' WHERE TDOCUMENTO='" & sDocumento & "'"
                       End If
                
                          ImprimeBoletaT RsImpresion, sTipoDocumento, imageCab, imagepIE, TimpresionDolaresDelivery
                       If lLogCajaRapida = True Then
                             Cn.Execute "UPDATE TLOG_IMPRESION SET TPOSICION7='FIN IMPRE ' + '" & RsImpresion.RecordCount & "' +'/' +'" & sTipoDocumento & "' WHERE TDOCUMENTO='" & sDocumento & "'"
                       End If
                          
                       End If
                  End If
                
             Else
                  If RsTipoDocumento!Cliente Then
                       ImprimeFacturaConsumoN RsImpresion, sDetalleConsumo, sEmpresa
                       NFactura = sCorrela
                       frmVenta.lblFactura.Caption = NFactura
                    ElseIf RsTipoDocumento!TTipoEmision = "00" Then
                       If MsgBox("Deseas imprimir el Voucher", vbQuestion + vbYesNo, sMensaje) = vbYes Then
                          ImprimeCortesia RsImpresion, sTipoDocumento, imageCab, imagepIE
                       End If
                    Else
                        'Log
                        If lLogCajaRapida = True Then
                            Cn.Execute "UPDATE TLOG_IMPRESION SET TPOSICION4='INIC IMPRE ' + '" & RsImpresion.RecordCount & "' +'/' +'" & sTipoDocumento & "' WHERE TDOCUMENTO='" & sDocumento & "'"
                        End If
                       ImprimeBoletaT RsImpresion, sTipoDocumento, imageCab, imagepIE, TimpresionDolaresDelivery
                       If lLogCajaRapida = True Then
                            Cn.Execute "UPDATE TLOG_IMPRESION SET TPOSICION7='FIN IMPRE ' + '" & RsImpresion.RecordCount & "' +'/' +'" & sTipoDocumento & "' WHERE TDOCUMENTO='" & sDocumento & "'"
    
                       End If
                  End If
                
             End If
             
          End If
               'CESAR FACTURACION ELECTRONICA
               If pais = "002" Then
                    If lFacturacionE Then
                     ' PARA FACTURACION DE ECUADOR
                    End If
               End If
               '---------------------------------
'               If pais = "002" Then
'                   sXML = Trim(LeerIni(App.Path + "\INFOREST.INI", "Configuracion", "RutaXML", "."))
'                   GeneraFacturaElectronica sXML, sDocumento
'               End If
                If pais = "002" And lFEEcuador = False Then
                   sXML = Trim(LeerIni(App.Path + "\INFOREST.INI", "Configuracion", "RutaXML", "."))
                   GeneraFacturaElectronica sXML, sDocumento
                End If
                
                If lFEEcuador Then
                 If INSERTA_FE_INFOREST(sDocumento, 1, DateTime.Now) = False Then
                     MsgBox "No se pudo enviar el documento a Facturacion Electronica!!! Verificar con su area de sistemas!!!"
                 End If
                End If
          
       End If

       LimpiaRs
        If lPagoAntesImpresion = False Then
                'Cancelacion del Documento
                If RsTipoDocumento!TTipoEmision <> "00" Then
                   Cn.Execute "Update MPEDIDO set tEstadoPedido = '02'  where tCodigoPedido ='" & Pedido & "'"
                   xTipo = ""
                   sFormulario = "CajaRapida"
                   If sFormulario = "CajaRapida" And lPagoRapido = True Then
                      frmPagoRapido.Show vbModal
                   Else
                      frmPago.Show vbModal
                   End If
                End If
        End If
        '-----------------------
        If pais = "000" And lFEpape And IIf(RsTipoDocumento!lFacturacionElectronica = True, 1, 0) Then
             If Not FacturarTCPIP(3, sDocumento, 0) Then
                MsgBox ("La confirmacion ha fallado, reenviar Documento!!!"), vbInformation, sMensaje
             End If
        End If
        '------------------------
    End If

    Cn.Execute "delete " & sDetalle
    Cn.Execute "delete " & sComboDetalle
    Cn.Execute "delete " & sComboPropiedad
    Cn.Execute "delete " & sProductoPropiedad

    RsDetalle.Requery
    RsComboPropiedad.Requery
    RsProductoPropiedad.Requery
    Inicializar
    Screen.MousePointer = vbDefault
    Exit Sub
fin:
wEnter = False
RsDetalle.Requery
RsComboPropiedad.Requery
RsProductoPropiedad.Requery
Inicializar
Screen.MousePointer = vbDefault
Call Log_Inforest("CAJA RAPIDA", "EMISION DE DOCUMENTO", sPedido, "", sDocumento, error, "", "FALLA AL GENERAR DOCUMENTO CAJA RAPIDA ", sUsuario)
MsgBox "Error: Emision de Documento / " + error, vbInformation, sMensaje
    
End Sub


Private Sub ImprimeFormatoAConsumo(ByVal nDocumento As String)

                    Dim ReporteC As New dsrBoletaC
                    
                    If RsTipoDocumento!lImprimeImageCab Then
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
                    
                    ReporteC.Text31.SetText sDetalleConsumo
                    
                    If Calcular(" SELECT case when  lImpresionRetencion=1 then 1 else 0 end  as codigo FROM vTipodocumento WHERE Codigo=(SELECT TTIPODOCUMENTO FROM MDOCUMENTO WHERE tDocumento='" & nDocumento & "')", Cn) = 1 Then
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


Private Sub ImprimeFormatoA(ByVal nDocumento As String)
                    Dim Reporte As New dsrBoleta

                    If RsTipoDocumento!lImprimeImageCab Then
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
                    
                    If Calcular(" SELECT case when  lImpresionRetencion=1 then 1 else 0 end  as codigo FROM vTipodocumento WHERE Codigo=(SELECT TTIPODOCUMENTO FROM MDOCUMENTO WHERE tDocumento='" & nDocumento & "')", Cn) = 1 Then
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

Public Sub GeneraPedido()
    Dim oComando As clsComando
    Set oComando = New clsComando
    If Not oComando.CreateCmdSp("spIns_MPEDIDO", Cn) Then
       Set oComando = Nothing
       Exit Sub
    End If
    
    oComando.CreateParameter "@tCliente", adVarChar, adParamInput, 7, sClienteFrecuente
    oComando.CreateParameter "@tTipoPedido", adVarChar, adParamInput, 2, sTipoPedido
    oComando.CreateParameter "@lPrioridad", adBoolean, adParamInput, 1, 0
    oComando.CreateParameter "@tTipoAtencion", adVarChar, adParamInput, 2, "01"
    oComando.CreateParameter "@tMesa", adVarChar, adParamInput, 3, ""
    oComando.CreateParameter "@tMozo", adVarChar, adParamInput, 4, Right(sMozo, 4)
    oComando.CreateParameter "@tMotorizado", adVarChar, adParamInput, 4, "0000"
    oComando.CreateParameter "@tCaja", adVarChar, adParamInput, 3, sCaja
    oComando.CreateParameter "@tSalon", adVarChar, adParamInput, 2, sSalon
    oComando.CreateParameter "@tTurno", adVarChar, adParamInput, 10, sTurno
    oComando.CreateParameter "@tObservacion", adVarChar, adParamInput, 250, sObser
    oComando.CreateParameter "@nTiempo", adInteger, adParamInput, 10, 0
    oComando.CreateParameter "@tUsuario", adVarChar, adParamInput, 15, Mid(sUsuario, 1, 15)
    oComando.CreateParameter "@nAdulto", adInteger, adParamInput, 10, 0
    oComando.CreateParameter "@nNino", adInteger, adParamInput, 10, 0
    oComando.CreateParameter "@nMesa", adInteger, adParamInput, 10, 0
    oComando.CreateParameter "@tPuntoVenta", adVarChar, adParamInput, 2, ""
    oComando.CreateParameter "@tHabitacion", adVarChar, adParamInput, 6, ""
    oComando.CreateParameter "@tReserva", adVarChar, adParamInput, 6, ""
    oComando.CreateParameter "@tPasajero", adVarChar, adParamInput, 50, ""
    oComando.CreateParameter "@tCompania", adVarChar, adParamInput, 5, ""
    oComando.CreateParameter "@tContacto", adVarChar, adParamInput, 4, ""
    oComando.CreateParameter "@nDescuento", adDouble, adParamInput, 10, xDescuento
    oComando.CreateParameter "@tDescuento", adVarChar, adParamInput, 3, sCodigoDescuento
    oComando.CreateParameter "@tObservacionDescuento", adVarChar, adParamInput, 250, sDescripcionDescuento
    oComando.CreateParameter "@tAutorizaDescuento", adVarChar, adParamInput, 15, IIf(sCodigoDescuento = "", "", tAutorizaDescuento)
    oComando.CreateParameter "@nTiempoDelivery", adInteger, adParamInput, 10, nTiempoDelivery
    oComando.CreateParameter "@tTienda", adVarChar, adParamInput, 3, ""
    oComando.CreateParameter "@fDiaContable", adDate, adParamInput, 10, obtieneDiaContable
    oComando.CreateParameter "@fProgramacion", adDate, adParamInput, 20, IIf(txtFechaEntrega.Caption = "", Null, Format(txtFechaEntrega.Caption, "dd/MM/yyyy HH:mm"))
    'invitado2013
    oComando.CreateParameter "@tCodigoInvitado", adVarChar, adParamInput, 10, sCodigoInvitado
    'pariente2014
    oComando.CreateParameter "@tCodigopariente", adVarChar, adParamInput, 7, sCodigoParienteSeleccionado
    'entregara
    oComando.CreateParameter "@tEntregarA", adVarChar, adParamInput, 20, IIf(Len(txtEntregar.Caption) = 0, "", Left(Me.txtEntregar.Caption, 20))
    oComando.CreateParameter "@nTiempoAntesEnvio", adInteger, adParamInput, 10, 0
    oComando.CreateParameter "@nMontoMaximo", adInteger, adParamInput, 250, 0
    oComando.CreateParameter "@tPedido", adVarChar, adParamOutput, 10, Pedido
    
    'origen de ventas
    If vOrigenVentas = Null Then
        vOrigenVentas = "00"
    End If
    oComando.CreateParameter "@codigoOrigenVentas", adVarChar, adParamInput, 2, vOrigenVentas
           
    If Not oComando.GetParamOK Then
       Set oComando = Nothing
       Exit Sub
    End If
    If Not oComando.ExecSP Then
       Set oComando = Nothing
       Exit Sub
    Else
       Pedido = oComando.GetParameterValue("@tPedido")
       
       Cn.Execute "UPDATE MPEDIDO SET FDIACONTABLE='" & Format(obtieneDiaContable, "yyyyMMdd") & "'  where tcodigopedido='" & Pedido & "' "
    End If
                                  
    'Actualiza el Numero de Pedido en el Detalle Temporal
    Cn.Execute "Update [" & sDetalle & "] Set tCodigoPedido = '" & Pedido & "'"
    
    'Inserta el Detalle
    
    Cn.Execute "Insert into DPEDIDO (tCodigoPedido, tItem, tTipoPedido, tCodigoProducto, tCodigoGrupo, tCodigoSubGrupo, tMoneda, " & _
               "nPrecioNeto, nPrecioImpuesto1, nPrecioImpuesto2, nPrecioImpuesto3, nPrecioVenta, nRecargo, nDescuento, nPrecioOficial, " & _
               "nCantidad, nImpuesto1, nImpuesto2, nImpuesto3, nVenta, tObservacion, tCortesia, lImprime, tEstadoItem, tArea, lImprimeArea,fregistro, nInsumo, nGasto, nManoObra, nOrden, lCombinacion, nCombinacion, lCorte,toferta,tautorizaoferta,tsubalmacen,tCodigoEtiqueta,tunidadnegocio,fDiaContable, tCajaD) " & _
               "select tCodigoPedido, tItem, tTipoPedido, tCodigoProducto, tCodigoGrupo, tCodigoSubGrupo, tMoneda, " & _
               "nPrecioNeto, nPrecioImpuesto1, nPrecioImpuesto2, nPrecioImpuesto3, nPrecioVenta, nRecargo, nDescuento, nPrecioOficial, " & _
               "nCantidad, nImpuesto1, nImpuesto2, nImpuesto3, nVenta, tObservacion, tCortesia, lImprime, tEstadoItem, tArea, lImprimeArea,getDate(), nInsumo, nGasto, nManoObra, nOrden, lCombinacion, nCombinacion, lCorte,toferta,tautorizaoferta,tSubalmacen,tCodigoEtiqueta,'" & sUnidadNegocio & "','" & Format(obtieneDiaContable, "yyyyMMdd") & "', '" & sCaja & "' " & _
               "From [" & sDetalle & "] where tEstadoItem='N'"
    

    'Actualiza el Numero de Pedido en el Detalle Combos
    Cn.Execute "Update [" & sComboDetalle & "] Set tCodigoPedido = '" & Pedido & "'"
    
    'Inserta Combo
    Cn.Execute "Insert into CPEDIDO select * from " & sComboDetalle
    
    'Inserta las propiedades de los Combos
    Cn.Execute "Insert into TCOMBOPROPIEDAD select '" & Pedido & "', tItem, tItemCombo, tCodigoPropiedad, tProducto, tEnlace, nInsumo, nGasto, nManoObra, 1,ncantidad,ninsumounitario,ngastounitario,nmanoobraunitario from " & sComboPropiedad
    
    'Inserta las propiedades
    Cn.Execute "Insert into TPRODUCTOPROPIEDAD select '" & Pedido & "', tItem, tCodigoPropiedad, tProducto, tEnlace, nInsumo, nGasto, nManoObra, 1,ncantidad,ninsumounitario,ngastounitario,nmanoobraunitario from " & sProductoPropiedad
    sPedido = Pedido
End Sub



Public Sub GrabaProducto()
   Screen.MousePointer = vbHourglass
   
   Isql = "Update [" & sDetalle & "] Set nPrecioNeto = " & nPBase & ", " & _
           "nDescuento = " & nDescuento & ", " & _
           "nRecargo = " & nRecargo & ", " & _
           "nPrecioOficial = " & nOficial & ", " & _
           "nprecioImpuesto1 = " & nImpuesto1 & ", " & _
           "nprecioImpuesto2 = " & nImpuesto2 & ", " & _
           "nprecioImpuesto3 = " & nImpuesto3 & ", " & _
           "nPrecioVenta = " & nPVenta & ", " & _
           "nventa = " & nPVenta * nCantidad & ", " & _
           "nCantidad = " & nCantidad & ", " & _
           "nImpuesto1 = " & nImpuesto1 * nCantidad & ", " & _
           "nImpuesto2 = " & nImpuesto2 * nCantidad & ", " & _
           "nImpuesto3 = " & nImpuesto3 * nCantidad & ", " & _
           "tCortesia = '" & sCortesia & "', " & _
           "lImprime = 0 " & _
           "where tItem = '" & sitem & "'"
           
           Cn.Execute Isql
           RsDetalle.Requery
           RsDetalle.MoveFirst
           RsDetalle.Find "tItem = '" & sitem & "'"
   nMonto = Format(Calcular("select sum(nVenta) as Codigo FROM [" & sDetalle & "] where tEstadoItem = 'N'", Cn), "#,###,##0.00")
   Screen.MousePointer = vbDefault
End Sub

Public Sub CalculaPrecio()
    Dim Acumulado As Double
    
    If nPVenta = 0 Then
       txtDPorcentaje.Caption = "0.00"
       txtRPorcentaje.Caption = "0.00"
       nRecargo = 0
       nDescuento = 0
       nImpuesto1 = 0
       nImpuesto2 = 0
       nImpuesto3 = 0
    Else
        Select Case pais 'ok
            Case "001" 'Bolivia
                    nPVenta = nOficial - nDescuento + nRecargo
                    
                    Acumulado = IIf(txtImpuesto1.Caption <> 0, Acumulado + nPorcentaje1, Acumulado)
                    Acumulado = IIf(txtImpuesto2.Caption <> 0, Acumulado + nPorcentaje2, Acumulado)
                    Acumulado = IIf(txtImpuesto3.Caption <> 0, Acumulado + nPorcentaje3, Acumulado)
                    Acumulado = (Acumulado / 100)
                    nImpuesto1 = IIf(txtImpuesto1.Caption <> 0, nPVenta * nPorcentaje1 / 100, 0)
                    nImpuesto2 = IIf(txtImpuesto2.Caption <> 0, nPVenta * nPorcentaje2 / 100, 0)
                    nImpuesto3 = IIf(txtImpuesto3.Caption <> 0, nPVenta * nPorcentaje3 / 100, 0)
                    nPBase = nPVenta - nImpuesto1 - nImpuesto2 - nImpuesto3
                    txtDPorcentaje.Caption = Format(nDescuento * 100 / nOficial, "###,###,###,##0.00")
                    txtRPorcentaje.Caption = Format(nRecargo * 100 / nOficial, "###,###,###,##0.00")
        
            Case Else 'Peru, Ecuador
                    nPVenta = nOficial - nDescuento + nRecargo
                    
                    Acumulado = IIf(txtImpuesto1.Caption <> 0, Acumulado + nPorcentaje1, Acumulado)
                    Acumulado = IIf(txtImpuesto2.Caption <> 0, Acumulado + nPorcentaje2, Acumulado)
                    Acumulado = IIf(txtImpuesto3.Caption <> 0, Acumulado + nPorcentaje3, Acumulado)
                    Acumulado = 1 + (Acumulado / 100)
                    nImpuesto1 = IIf(txtImpuesto1.Caption <> 0, nPVenta / Acumulado * nPorcentaje1 / 100, 0)
                    nImpuesto2 = IIf(txtImpuesto2.Caption <> 0, nPVenta / Acumulado * nPorcentaje2 / 100, 0)
                    nImpuesto3 = IIf(txtImpuesto3.Caption <> 0, nPVenta / Acumulado * nPorcentaje3 / 100, 0)
                    nPBase = nPVenta - nImpuesto1 - nImpuesto2 - nImpuesto3
                    txtDPorcentaje.Caption = Format(nDescuento * 100 / nOficial, "###,###,###,##0.00")
                    txtRPorcentaje.Caption = Format(nRecargo * 100 / nOficial, "###,###,###,##0.00")
        
        End Select
    End If
    txtImpuesto1.Caption = Format(nImpuesto1, "###,###,###,##0.00")
    txtImpuesto2.Caption = Format(nImpuesto2, "###,###,###,##0.00")
    txtImpuesto3.Caption = Format(nImpuesto3, "###,###,###,##0.00")
    
    txtNeto.Caption = Format(nPBase, "###,###,##0.00")
    txtPVenta.Caption = Format(nPVenta, "###,###,##0.00")
    txtVenta.Caption = Format((nPVenta * nCantidad), "###,###,###,##0.00")
End Sub

Public Sub Impuesto()
   Label1(10).Caption = sImpuesto1 & " :"
   Label1(11).Caption = sImpuesto2 & " :"
   Label1(12).Caption = sImpuesto3 & " :"
   
   Label1(10).Visible = IIf(sImpuesto1 = "", False, True)
   Label1(11).Visible = IIf(sImpuesto2 = "", False, True)
   Label1(12).Visible = IIf(sImpuesto3 = "", False, True)
   
   txtImpuesto1.Visible = IIf(sImpuesto1 = "", False, True)
   txtImpuesto2.Visible = IIf(sImpuesto2 = "", False, True)
   txtImpuesto3.Visible = IIf(sImpuesto3 = "", False, True)
   
   cmdImpuesto(0).Caption = sImpuesto1
   cmdImpuesto(1).Caption = sImpuesto2
   cmdImpuesto(2).Caption = sImpuesto3
   
   cmdImpuesto(0).Visible = IIf(sImpuesto1 = "", False, True)
   cmdImpuesto(1).Visible = IIf(sImpuesto2 = "", False, True)
   cmdImpuesto(2).Visible = IIf(sImpuesto3 = "", False, True)
End Sub

Private Sub grdCombo_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
  AsignaCombo
  If fraPropiedad.Visible = True Then
     cmdOpcion_Click (6)
  End If
End Sub

Private Sub grdDetalle_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
   
   If fraPropiedad.Visible = True Then
      nPos = RsDetalle.AbsolutePosition
      RsDetalle.Requery
      RsDetalle.AbsolutePosition = nPos
   End If
   AsignaProducto
   
   If lPropiedad Then
      lPropiedad = False
      cmdDetalle_Click (5)
   End If

   On Error Resume Next
   txtBarra.SetFocus
End Sub

Public Sub AsignaCombo()
   If Not RsCombo.EOF Then
      sCombo = IIf(IsNull(RsCombo!tProductoCombo), "", RsCombo!tProductoCombo)
      xItem = IIf(IsNull(RsCombo!tItemCombo), "001", RsCombo!tItemCombo)
      lblObservacion.Text = IIf(IsNull(RsCombo!tObservacion), "", RsCombo!tObservacion)
      ListarOperadoresConFiltro (sCombo) 'OO
      AsignaComboPropiedad
   End If
End Sub

Private Sub cmdPropiedad_Click(Index As Integer)
    Dim nInsumo As Double
    Dim nGasto As Double
    Dim nMObra As Double
    
    Dim Cantidad As Double
    Dim ncantidadPropiedad As Double
    
    RsPropiedad.MoveFirst
    RsPropiedad.Find ("Descripcion = '" & cmdPropiedad(Index).Caption & "'")
      If Not (RsOperador.EOF Or RsOperador.BOF) Then
     nOperadorPropiedad = Calcular("select isnull(ncontrol,0) as codigo from voperador where codigo='" & RsOperador!codigo & "'", Cn)
     End If
  
    
    
    If cmdPropiedad(Index).FontBold = True Then
       cmdPropiedad(Index).FontBold = False
       If Not RsPropiedad.EOF Then
          If wAgregaCombo Then
             Cantidad = Calcular("select isnull(ncantidad,1) as codigo from " & sComboPropiedad & " where      titem='" & sitem & "' and titemcombo='" & xItem & "' and  tproducto='" & sCombo & "' and tcodigopropiedad='" & RsPropiedad!codigo & "' ", Cn)
             If RsPropiedad!nPrecio <> 0 Then
                nMonto = CambiaPrecio(nPVenta - (RsPropiedad!nPrecio * Cantidad))
                txtMonto.Caption = Format(nMonto, "###,##0.00")
             End If
             Cn.Execute "delete " & sComboPropiedad & " where tItem = '" & sitem & "' and tItemCombo='" & xItem & "' and tProducto='" & sCombo & "' and tCodigoPropiedad='" & RsPropiedad!codigo & "'"
          Else
             Cantidad = Calcular("select isnull(ncantidad,1) as codigo from " & sProductoPropiedad & " where     titem='" & sitem & "' and tproducto='" & sProducto & "' and tcodigopropiedad='" & RsPropiedad!codigo & "'  ", Cn)
             If RsPropiedad!nPrecio <> 0 Then
                nMonto = CambiaPrecio(nPVenta - (RsPropiedad!nPrecio * Cantidad))
                txtMonto.Caption = Format(nMonto, "###,##0.00")
             End If
             Cn.Execute "delete " & sProductoPropiedad & " where tItem = '" & sitem & "' and tProducto='" & sProducto & "' and tCodigoPropiedad='" & RsPropiedad!codigo & "'"
          End If
          
          If Cantidad <> 1 Then
            'lblResumen.Text = Replace(lblResumen.Text, RsOperador!Descripcion & " " & cmdPropiedad(Index).Caption & ": (" & Cantidad & "), ", "")
            lblResumen.Text = Replace(lblResumen.Text, RsOperador!Descripcion & " " & cmdPropiedad(Index).Caption & ", ", "")
          Else
            lblResumen.Text = Replace(lblResumen.Text, RsOperador!Descripcion & " " & cmdPropiedad(Index).Caption & ", ", "")
          End If
          
          'lblResumen.Text = Replace(lblResumen.Text, RsOperador!Descripcion & " " & cmdPropiedad(Index).Caption & ", ", "")
       End If
    Else
    
            ncantidadPropiedad = 1
             If RsPropiedad!lsolicitacantidad = 1 Or RsPropiedad!lsolicitacantidad = True Then
                sTipo = "Prepintado"
                
                sCodigo = ncantidadPropiedad
                
                frmNumPad.Show vbModal
                If wEnter And Val(sDescrip) > 0 Then
                
                            ncantidadPropiedad = sDescrip
                            
                End If
            End If
               
        
    
    
       If nOperadorPropiedad > 0 Then
          If wAgregaCombo Then
             Isql = "SELECT COUNT(" & sComboPropiedad & ".tCodigoPropiedad) AS codigo " & _
                    "FROM " & sComboPropiedad & " INNER JOIN dbo.TPROPIEDAD ON " & sComboPropiedad & ".tCodigoPropiedad = dbo.TPROPIEDAD.tCodigoPropiedad AND " & sComboPropiedad & ".tProducto = dbo.TPROPIEDAD.tProducto " & _
                    "where tItem = '" & sitem & "' and tItemCombo='" & xItem & "' and " & sComboPropiedad & ".tProducto='" & sCombo & "'  and tOperador='" & RsOperador!codigo & "'"
             If nOperadorPropiedad <= Calcular(Isql, Cn) Then
                MsgBox "Ha llegado a la Cantidad máxima de " & nOperadorPropiedad & " Propiedad(es) por Operador", vbExclamation, sMensaje
                Exit Sub
             End If
          Else
             Isql = "SELECT COUNT(" & sProductoPropiedad & ".tCodigoPropiedad) AS codigo FROM " & sProductoPropiedad & " INNER JOIN " & _
                    "dbo.TPROPIEDAD ON " & sProductoPropiedad & ".tCodigoPropiedad = dbo.TPROPIEDAD.tCodigoPropiedad AND " & sProductoPropiedad & ".tProducto = dbo.TPROPIEDAD.tProducto " & _
                    "where tItem = '" & sitem & "' and tOperador='" & RsOperador!codigo & "'"
             If nOperadorPropiedad <= Calcular(Isql, Cn) Then
                MsgBox "Ha llegado a la Cantidad máxima de " & nOperadorPropiedad & " Propiedad(es) por Operador", vbExclamation, sMensaje
                Exit Sub
             End If
          End If
       End If
       
       cmdPropiedad(Index).FontBold = True
       If Not RsPropiedad.EOF Then
          nInsumo = IIf(IsNull(RsPropiedad!nInsumo), 0, RsPropiedad!nInsumo)
          nGasto = IIf(IsNull(RsPropiedad!nGasto), 0, RsPropiedad!nGasto)
          nMObra = IIf(IsNull(RsPropiedad!nManoObra), 0, RsPropiedad!nManoObra)
       
          If wAgregaCombo Then
             Cn.Execute "Insert into " & sComboPropiedad & " values ('" & sitem & "', '" & xItem & "', '" & RsPropiedad!codigo & "', '" & sCombo & "', '" & RsPropiedad!tEnlace & "', " & IIf(IsNull(RsPropiedad!nInsumo), 0, ncantidadPropiedad * RsPropiedad!nInsumo) & ", " & IIf(IsNull(RsPropiedad!nGasto), 0, ncantidadPropiedad * RsPropiedad!nGasto) & ", " & IIf(IsNull(RsPropiedad!nManoObra), 0, ncantidadPropiedad * RsPropiedad!nManoObra) & ", " & ncantidadPropiedad & ", " & IIf(IsNull(RsPropiedad!nInsumo), 0, RsPropiedad!nInsumo) & ", " & IIf(IsNull(RsPropiedad!nGasto), 0, RsPropiedad!nGasto) & ", " & IIf(IsNull(RsPropiedad!nManoObra), 0, RsPropiedad!nManoObra) & ") "
             If RsPropiedad!nPrecio <> 0 Then
                nMonto = CambiaPrecio(nPVenta + (RsPropiedad!nPrecio * ncantidadPropiedad))
                txtMonto.Caption = Format(nMonto, "###,##0.00")
             End If
          Else
             Cn.Execute "Insert into " & sProductoPropiedad & " values ('" & sitem & "', '" & RsPropiedad!codigo & "', '" & sProducto & "', '" & RsPropiedad!tEnlace & "', " & IIf(IsNull(RsPropiedad!nInsumo), 0, ncantidadPropiedad * RsPropiedad!nInsumo) & ", " & IIf(IsNull(RsPropiedad!nGasto), 0, ncantidadPropiedad * RsPropiedad!nGasto) & ", " & IIf(IsNull(RsPropiedad!nManoObra), 0, ncantidadPropiedad * RsPropiedad!nManoObra) & ", " & ncantidadPropiedad & "," & IIf(IsNull(RsPropiedad!nInsumo), 0, RsPropiedad!nInsumo) & ", " & IIf(IsNull(RsPropiedad!nGasto), 0, RsPropiedad!nGasto) & ", " & IIf(IsNull(RsPropiedad!nManoObra), 0, RsPropiedad!nManoObra) & " )"
             If RsPropiedad!nPrecio <> 0 Then
                nMonto = CambiaPrecio(nPVenta + (RsPropiedad!nPrecio * ncantidadPropiedad))
                txtMonto.Caption = Format(nMonto, "###,##0.00")
             End If
          End If
          'lblResumen.Text = lblResumen.Text & RsOperador!Descripcion & " " & cmdPropiedad(Index).Caption & ", "
          
           If ncantidadPropiedad <> 1 Then
          
                lblResumen.Text = lblResumen.Text & RsOperador!Descripcion & " " & cmdPropiedad(Index).Caption & ": (" & ncantidadPropiedad & "), "
          Else
                lblResumen.Text = lblResumen.Text & RsOperador!Descripcion & " " & cmdPropiedad(Index).Caption & ", "
          End If
          
       End If
       
    End If
    
    If wAgregaCombo Then
       RsComboPropiedad.Requery
    Else
       RsProductoPropiedad.Requery
    End If
End Sub


'
Public Sub AsignaPropiedad()
    Dim i As Integer
    If RsOperador.RecordCount > 0 Then
       RsPropiedad.Filter = "tOperador = '" & RsOperador!codigo & "' and tProducto='" & sProducto & "'"
       nOperadorPropiedad = IIf(IsNull(RsOperador!nControl), 0, RsOperador!nControl)
    Else
       RsPropiedad.Filter = "tOperador = '  ' and tProducto='" & sProducto & "'"
       nOperadorPropiedad = 0
    End If
    AsignaComando 19, RsPropiedad, cmdPropiedad()

    For i = 1 To 19
        cmdPropiedad(i).FontBold = False
    Next i
    
    lblResumen.Text = ""
    RsProductoPropiedad.Filter = "tItem='" & sitem & "'"
    If Not RsProductoPropiedad.EOF Then
       RsProductoPropiedad.MoveFirst
       Do While Not RsProductoPropiedad.EOF
          For i = 1 To 19
              If cmdPropiedad(i).Caption = RsProductoPropiedad!Descripcion And RsOperador!Descripcion = RsProductoPropiedad!Operador Then
                 cmdPropiedad(i).FontBold = True
                 Exit For
              End If
          Next i
          If RsProductoPropiedad!nCantidad = 1 Then
            lblResumen.Text = lblResumen.Text & LTrim(RsProductoPropiedad!Operador) & " " & LTrim(RsProductoPropiedad!Descripcion) & ", "
          Else
             lblResumen.Text = lblResumen.Text & LTrim(RsProductoPropiedad!Operador) & " " & LTrim(RsProductoPropiedad!Descripcion) & ": (" & RsProductoPropiedad!nCantidad & "), "
          End If
          RsProductoPropiedad.MoveNext
       Loop
    End If

End Sub

Private Sub txtBarra_KeyPress(KeyAscii As Integer)

   If KeyAscii = 13 And txtBarra.Text <> "" Then
        Dim xxx As String
        
        If lRotulado = True Then
              'CESAR ROTULADO
              Dim rCodigoEtiqueta As String
              Dim rCodigoProducto As String
              Dim rLenBarra As String
              Dim rCantidad As Double
              Dim X As Integer
              
              xxx = RsProducto.Filter
              
              rLenBarra = Len(Trim(txtBarra.Text))
              X = rLenBarra - 31
              
              If nLongitudBarra <> 0 Then
                rCodigoProducto = Mid(txtBarra.Text, nLongitudBarra + 1, 7)
                rCodigoEtiqueta = Mid(txtBarra.Text, 1, nLongitudBarra)
                
                If lCapturaPeso Then
                   rCantidad = Val(Mid(txtBarra.Text, 31 + 1, X))
                   InsertaProductoRotulado rCodigoProducto, rCantidad, rCodigoEtiqueta
                Else
                   rCantidad = 1
                   InsertaProductoRotulado rCodigoProducto, rCantidad, rCodigoEtiqueta
                End If
              Else
                  MsgBox "Error: Longitud de barra no registrada", vbCritical, sMensaje
              End If
              txtBarra.Text = ""
              sProducto = ""
              RsProducto.Filter = IIf(xxx = "0", "", xxx)
        Else
     
            xxx = RsProducto.Filter
            RsProducto.Filter = adFilterNone
            RsProducto.MoveFirst
            RsProducto.Find "tbarra = '" & Trim(txtBarra.Text) & "'"
            
            If Not RsProducto.EOF Then
               sProducto = RsProducto!codigo
               
               If lBal And RsProducto!lBalanza Then
                  Dim nResultado As Double
                  nResultado = Pesar(nBalanzaPuerto)
                  nResultado = Format(nResultado, "#,##0.00")
                  If nResultado > 0 Then
                     InsertaProducto nResultado
                  End If
               Else
               nCantidad = 1
                  InsertaProducto 1
               End If
               
               
               If IIf(IsNull(RsProducto!lPropiedad), False, RsProducto!lPropiedad) Then
                  lPropiedad = True
               End If
            Else
               If nLongitudBarra > 0 Then
                  RsProducto.MoveFirst
                  RsProducto.Find "tbarra = '" & Trim(Mid(txtBarra.Text, 1, nLongitudBarra)) & "'"
                  If Not RsProducto.EOF Then
                     sProducto = RsProducto!codigo
                     Dim nCantidadBarra As Double
                       
                        If lCapturaPeso Then
                        
                         If EAN13 Then
                                nCantidadBarra = Val(Mid(txtBarra.Text, nLongitudBarra + 1, 1) + "." + Mid(txtBarra.Text, nLongitudBarra + 2, 3))
                             Else
                                nCantidadBarra = Val(Mid(txtBarra.Text, nLongitudBarra + 1, 2) + "." + Mid(txtBarra.Text, nLongitudBarra + 3, 4))
                             End If
                            'nCantidadBarra = Val(Mid(txtBarra.Text, nLongitudBarra + 1, 2) + "." + Mid(txtBarra.Text, nLongitudBarra + 3, 4))
                           InsertaProducto nCantidadBarra
                        Else
                           nCantidadBarra = Val(Mid(txtBarra.Text, nLongitudBarra + 1, 3) + "." + Mid(txtBarra.Text, nLongitudBarra + 4, 3))
                           InsertaProducto CalculaCantidad(nCantidadBarra)
                        End If
                     
                     If IIf(IsNull(RsProducto!lPropiedad), False, RsProducto!lPropiedad) Then
                        lPropiedad = True
                     End If
                  Else
                     MsgBox "Producto no encontrado", vbCritical, sMensaje
                  End If
               Else
                  MsgBox "Producto no encontrado", vbCritical, sMensaje
               End If
            End If
            txtBarra.Text = ""
            sProducto = ""
            RsProducto.Filter = IIf(xxx = "0", "", xxx)
        End If

   End If
   
End Sub
Public Sub InsertaProductoRotulado(codigoProducto As String, xCantidad As Double, codigoEtiqueta As String)
    'CESAR ROTULADO
    Dim nValor As Double
    Dim lImp1 As Boolean
    Dim lImp2 As Boolean
    Dim lImp3 As Boolean
    'ORDEN
    Dim RsOrd As Recordset
    Dim nOrden As Integer
    'DETALLE
    Dim lProductoMultiArea As Boolean
    Dim tsubalmacen As String
    Dim tAreaProduccion As String
    'OFERTA
    Dim tOferta As String
    Dim nOferta As Double
    tOferta = ""
    nOferta = 0
    
    RsProducto.Filter = adFilterNone
    RsProducto.MoveFirst
    RsProducto.Find "Codigo = '" & Trim(codigoProducto) & "'"
    
    
    If Not RsProducto.EOF Then

        nRecargo = 0
        nDescuento = 0
        nValor = 0
        nValor = nValor + IIf(RsProducto!lImpuesto1, nPorcentaje1, 0)
        nValor = nValor + IIf(RsProducto!lImpuesto2, nPorcentaje2, 0)
        nValor = nValor + IIf(RsProducto!lImpuesto3, nPorcentaje3, 0)

        lImp1 = RsProducto!lImpuesto1
        lImp2 = RsProducto!lImpuesto2
        lImp3 = RsProducto!lImpuesto3

        If sTipoPedido = "02" Then
           If IsNull(RsProducto!nPrecioDelivery) Or RsProducto!nPrecioDelivery = 0 Then
              nOficial = IIf(IsNull(RsProducto!nprecioVenta), 0, RsProducto!nprecioVenta) + (nDELIVERY * IIf(IsNull(RsProducto!nprecioVenta), 0, RsProducto!nprecioVenta) / 100)
           Else
              nOficial = IIf(IsNull(RsProducto!nPrecioDelivery), 0, RsProducto!nPrecioDelivery)
              nValor = 0
              nValor = nValor + IIf(RsProducto!lImpuesto4, nPorcentaje1, 0)
              nValor = nValor + IIf(RsProducto!lImpuesto5, nPorcentaje2, 0)
              nValor = nValor + IIf(RsProducto!lImpuesto6, nPorcentaje3, 0)
              lImp1 = RsProducto!lImpuesto4
              lImp2 = RsProducto!lImpuesto5
              lImp3 = RsProducto!lImpuesto6
           End If
        ElseIf sTipoPedido = "03" Then
           If IsNull(RsProducto!nPreciollevar) Or RsProducto!nPreciollevar = 0 Then
              nOficial = IIf(IsNull(RsProducto!nprecioVenta), 0, RsProducto!nprecioVenta) + (nLlevar * IIf(IsNull(RsProducto!nprecioVenta), 0, RsProducto!nprecioVenta) / 100)
           Else
              nOficial = IIf(IsNull(RsProducto!nPreciollevar), 0, RsProducto!nPreciollevar)
              nValor = 0
              nValor = nValor + IIf(RsProducto!lImpuesto7, nPorcentaje1, 0)
              nValor = nValor + IIf(RsProducto!lImpuesto8, nPorcentaje2, 0)
              nValor = nValor + IIf(RsProducto!lImpuesto9, nPorcentaje3, 0)
              lImp1 = RsProducto!lImpuesto7
              lImp2 = RsProducto!lImpuesto8
              lImp3 = RsProducto!lImpuesto9
           End If
        ElseIf sTipoPedido = "04" Then
           If IsNull(RsProducto!nPrecioCanal4) Or RsProducto!nPrecioCanal4 = 0 Then
              nOficial = IIf(IsNull(RsProducto!nprecioVenta), 0, RsProducto!nprecioVenta) + (nLlevar * IIf(IsNull(RsProducto!nprecioVenta), 0, RsProducto!nprecioVenta) / 100)
           Else
              nOficial = IIf(IsNull(RsProducto!nPrecioCanal4), 0, RsProducto!nPrecioCanal4)
              nValor = 0
              nValor = nValor + IIf(RsProducto!lImpuesto10, nPorcentaje1, 0)
              nValor = nValor + IIf(RsProducto!lImpuesto11, nPorcentaje2, 0)
              nValor = nValor + IIf(RsProducto!lImpuesto12, nPorcentaje3, 0)
              lImp1 = RsProducto!lImpuesto10
              lImp2 = RsProducto!lImpuesto11
              lImp3 = RsProducto!lImpuesto12
           End If
        ElseIf sTipoPedido = "05" Then
           If IsNull(RsProducto!nPrecioCanal5) Or RsProducto!nPrecioCanal5 = 0 Then
              nOficial = IIf(IsNull(RsProducto!nprecioVenta), 0, RsProducto!nprecioVenta) + (nLlevar * IIf(IsNull(RsProducto!nprecioVenta), 0, RsProducto!nprecioVenta) / 100)
           Else
              nOficial = IIf(IsNull(RsProducto!nPrecioCanal5), 0, RsProducto!nPrecioCanal5)
              nValor = 0
              nValor = nValor + IIf(RsProducto!lImpuesto13, nPorcentaje1, 0)
              nValor = nValor + IIf(RsProducto!lImpuesto14, nPorcentaje2, 0)
              nValor = nValor + IIf(RsProducto!lImpuesto15, nPorcentaje3, 0)
              lImp1 = RsProducto!lImpuesto13
              lImp2 = RsProducto!lImpuesto14
              lImp3 = RsProducto!lImpuesto15
           End If
        Else
           nOficial = IIf(IsNull(RsProducto!nprecioVenta), 0, RsProducto!nprecioVenta)
        End If
        
        'PRECIO OFICIAL
        nOficial = IIf(RsProducto!tMoneda = "02", nOficial * nTC, nOficial)
        
        'DESCUENTO
        If xDescuento <> 0 And RsProducto!lDescuento Then
              nPVenta = nOficial - (nOficial * xDescuento / 100)
              nDescuento = nOficial - nPVenta
        Else
              nPVenta = nOficial - nOferta
              nDescuento = nOficial - nPVenta
        End If


        Select Case pais
            Case "001" 'Bolivia
                    nValor = (nValor / 100)
                    nImpuesto1 = IIf(lImp1, nPVenta * nPorcentaje1 / 100, 0)
                    nImpuesto2 = IIf(lImp2, nPVenta * nPorcentaje2 / 100, 0)
                    nImpuesto3 = IIf(lImp3, nPVenta * nPorcentaje3 / 100, 0)
                    nPBase = nPVenta - nImpuesto1 - nImpuesto2 - nImpuesto3

            Case Else 'Peru, Ecuador
                    nValor = 1 + (nValor / 100)
                    nImpuesto1 = IIf(lImp1, nPVenta / nValor * nPorcentaje1 / 100, 0)
                    nImpuesto2 = IIf(lImp2, nPVenta / nValor * nPorcentaje2 / 100, 0)
                    nImpuesto3 = IIf(lImp3, nPVenta / nValor * nPorcentaje3 / 100, 0)
                    nPBase = nPVenta - nImpuesto1 - nImpuesto2 - nImpuesto3
        End Select

        Dim nInsumo As Double
        Dim nGasto As Double
        Dim nMObra As Double
        
            If sTipoPedido = "01" Then
               nInsumo = IIf(IsNull(RsProducto!nInsumo), 0, RsProducto!nInsumo)
               nGasto = IIf(IsNull(RsProducto!nGasto), 0, RsProducto!nGasto)
               nMObra = IIf(IsNull(RsProducto!nManoObra), 0, RsProducto!nManoObra)
            ElseIf sTipoPedido = "02" Then
               nInsumo = IIf(IsNull(RsProducto!nInsumo2), 0, RsProducto!nInsumo2)
               nGasto = IIf(IsNull(RsProducto!nGasto2), 0, RsProducto!nGasto2)
               nMObra = IIf(IsNull(RsProducto!nManoObra2), 0, RsProducto!nManoObra2)
            ElseIf sTipoPedido = "03" Then
               nInsumo = IIf(IsNull(RsProducto!nInsumo3), 0, RsProducto!nInsumo3)
               nGasto = IIf(IsNull(RsProducto!nGasto3), 0, RsProducto!nGasto3)
               nMObra = IIf(IsNull(RsProducto!nManoObra3), 0, RsProducto!nManoObra3)
            ElseIf sTipoPedido = "04" Then
               nInsumo = IIf(IsNull(RsProducto!nInsumo4), 0, RsProducto!nInsumo4)
               nGasto = IIf(IsNull(RsProducto!nGasto4), 0, RsProducto!nGasto4)
               nMObra = IIf(IsNull(RsProducto!nManoObra4), 0, RsProducto!nManoObra4)
            Else
               nInsumo = IIf(IsNull(RsProducto!nInsumo5), 0, RsProducto!nInsumo5)
               nGasto = IIf(IsNull(RsProducto!nGasto5), 0, RsProducto!nGasto5)
               nMObra = IIf(IsNull(RsProducto!nManoObra5), 0, RsProducto!nManoObra5)
            End If
            
            sitem = Lib.Correlativo(Calcular("select max(tItem) as codigo from [" & sDetalle & "]", Cn), 3)
            'CALCULAR ITEM
            If RsDetalle.RecordCount = 0 Then
               'sitem = "001"
               nOrden = 1
            Else
               'sitem = Lib.Correlativo(Calcular("select max(tItem) as codigo from [" & sDetalle & "]", Cn), 3)
               If lOrden Then
                  Set RsOrd = Lib.OpenRecordset("select nOrden, lImprime from " & sDetalle & " Order by nOrden DESC", Cn)
                  If RsOrd.RecordCount > 0 Then
                     If IIf(IsNull(RsOrd!lImprime), False, RsOrd!lImprime) Then
                        nOrden = RsOrd!nOrden + 1
                     Else
                        nOrden = RsOrd!nOrden
                     End If
                  Else
                     nOrden = 1
                  End If
               Else
                  nOrden = RsProducto!nOrden
               End If
               
            End If
            
            
            'TSUBALMACEN
            lProductoMultiArea = Calcular("select isnull(lmultiarea,0) as codigo from tproducto where tcodigoproducto='" & RsProducto.Fields("codigo") & "'", Cn)
    
            If lProductoMultiArea = False Then
                tsubalmacen = ""
            Else
                tAreaProduccion = Calcular("select isnull(tsubalmacen,'') as codigo from tcaja where tcaja='" & sCaja & "'", Cn)
                
                tsubalmacen = Calcular("select isnull(tvalor,'')  as codigo from varea where codigo='" & tAreaProduccion & "'", Cn)
                
                If tsubalmacen = "0" Then
                    tsubalmacen = ""
                End If
            End If
            

            'DETALLE TEMPORAL
             Isql = "insert into [" & sDetalle & "] " & _
                    "(tCodigoPedido, tTipoPedido, tItem, tCodigoProducto, tCodigoGrupo, tCodigoSubGrupo, " & _
                    "nPrecioNeto, nRecargo, nDescuento, nPrecioOficial, nPrecioImpuesto1, nPrecioImpuesto2, nPrecioImpuesto3, nPrecioVenta, " & _
                    "nCantidad, nVenta, nImpuesto1, nImpuesto2, nImpuesto3, " & _
                    "lImprime, tArea, lImprimeArea, lCombinacion, nCombinacion, nInsumo, nGasto, nManoObra, nOrden, tEstadoItem, tCodigoEtiqueta, tsubalmacen, toferta) " & _
                    "Values( '" & Pedido & "', '01', " _
                            & "'" & sitem & "', " _
                            & "'" & codigoProducto & "', " _
                            & "'" & IIf(IsNull(RsProducto!tgrupo), "", RsProducto!tgrupo) & "', " _
                            & "'" & IIf(IsNull(RsProducto!tSubGrupo), "", RsProducto!tSubGrupo) & "', " _
                            & nPBase & ", " & nRecargo & ", " _
                            & nDescuento & ", " _
                            & nOficial & ", " _
                            & nImpuesto1 & ", " & nImpuesto2 & ", " & nImpuesto3 & ", " _
                            & nPVenta & ", " & xCantidad & ", " _
                            & nPVenta * xCantidad & ", " _
                            & nImpuesto1 * xCantidad & ", " & nImpuesto2 * xCantidad & ", " & nImpuesto3 * xCantidad & ", " _
                            & "0, '" & RsProducto!tArea & "', " & IIf(RsProducto!lImprimeArea, -1, 0) & "," _
                            & IIf(RsProducto!lCombinacion, 1, 0) & ", " & RsProducto!nCombinacion & ", " _
                            & nInsumo & ", " _
                            & nGasto & ", " _
                            & nMObra & ", " _
                            & nOrden & ", " _
                            & "'N','" & codigoEtiqueta & "','" & tsubalmacen & "','" & tOferta & "') "
             Cn.Execute Isql
             RsDetalle.Requery
              nMonto = Format(Calcular("select sum(nVenta) as Codigo FROM [" & sDetalle & "] where tEstadoItem = 'N'", Cn), "#,###,##0.00")
             RsDetalle.MoveLast
             Screen.MousePointer = vbDefault
             
    Else
       MsgBox "Producto no encontrado", vbCritical, sMensaje
    End If

End Sub


Public Function CalculaCantidad(nPrecio As Double) As Double
    Select Case sTipoPedido
           Case "01"
                CalculaCantidad = nPrecio / RsProducto!nprecioVenta
           Case "02"
                If RsProducto!nPrecioDelivery > 0 Then
                   CalculaCantidad = nPrecio / RsProducto!nPrecioDelivery
                Else
                   CalculaCantidad = nPrecio / RsProducto!nprecioVenta
                End If
           
           Case "03"
                If RsProducto!nPreciollevar > 0 Then
                   CalculaCantidad = nPrecio / RsProducto!nPreciollevar
                Else
                   CalculaCantidad = nPrecio / RsProducto!nprecioVenta
                End If
                
           Case "04"
                If RsProducto!nPrecioCanal4 > 0 Then
                   CalculaCantidad = nPrecio / RsProducto!nPrecioCanal4
                Else
                   CalculaCantidad = nPrecio / RsProducto!nprecioVenta
                End If
                   
           Case "05"
                If RsProducto!nPrecioCanal5 > 0 Then
                   CalculaCantidad = nPrecio / RsProducto!nPrecioCanal5
                Else
                   CalculaCantidad = nPrecio / RsProducto!nprecioVenta
                End If
    
    End Select
      
End Function

Public Sub InsertaCombo(wProducto As String)
    Screen.MousePointer = vbHourglass
    Dim xItem As String
    Dim nValor As Double
    Dim nCNeto As Double
    Dim nCImp1 As Double
    Dim nCImp2 As Double
    Dim nCImp3 As Double
    Dim nCVenta As Double
    Dim lImp1 As Boolean
    Dim lImp2 As Boolean
    Dim lImp3 As Boolean
    Dim nInsumo As Double
    Dim nGasto As Double
    Dim nMano As Double
    
    If RsCombo.RecordCount = 0 Then
       xItem = "001"
    Else
       xItem = Lib.Correlativo(Calcular("select max(tItemCombo) as codigo from " & sComboDetalle & " where tItem = '" & sitem & "'", Cn), 3)
    End If
                          
    nValor = 0
    nValor = nValor + IIf(RsProducto!lImpuesto1, nPorcentaje1, 0)
    nValor = nValor + IIf(RsProducto!lImpuesto2, nPorcentaje2, 0)
    nValor = nValor + IIf(RsProducto!lImpuesto3, nPorcentaje3, 0)
      
    lImp1 = RsProducto!lImpuesto1
    lImp2 = RsProducto!lImpuesto2
    lImp3 = RsProducto!lImpuesto3
          
    If sTipoPedido = "02" Then
       If IsNull(RsProducto!nPrecioDelivery) Or RsProducto!nPrecioDelivery = 0 Then
          nCVenta = IIf(IsNull(RsProducto!nprecioVenta), 0, RsProducto!nprecioVenta) + (nDELIVERY * IIf(IsNull(RsProducto!nprecioVenta), 0, RsProducto!nprecioVenta) / 100)
       Else
          nCVenta = IIf(IsNull(RsProducto!nPrecioDelivery), 0, RsProducto!nPrecioDelivery)
          nValor = 0
          nValor = nValor + IIf(RsProducto!lImpuesto4, nPorcentaje1, 0)
          nValor = nValor + IIf(RsProducto!lImpuesto5, nPorcentaje2, 0)
          nValor = nValor + IIf(RsProducto!lImpuesto6, nPorcentaje3, 0)
          lImp1 = RsProducto!lImpuesto4
          lImp2 = RsProducto!lImpuesto5
          lImp3 = RsProducto!lImpuesto6
       End If
    ElseIf sTipoPedido = "03" Then
       If IsNull(RsProducto!nPreciollevar) Or RsProducto!nPreciollevar = 0 Then
          nCVenta = IIf(IsNull(RsProducto!nprecioVenta), 0, RsProducto!nprecioVenta) + (nLlevar * IIf(IsNull(RsProducto!nprecioVenta), 0, RsProducto!nprecioVenta) / 100)
       Else
          nCVenta = IIf(IsNull(RsProducto!nPreciollevar), 0, RsProducto!nPreciollevar)
          nValor = 0
          nValor = nValor + IIf(RsProducto!lImpuesto7, nPorcentaje1, 0)
          nValor = nValor + IIf(RsProducto!lImpuesto8, nPorcentaje2, 0)
          nValor = nValor + IIf(RsProducto!lImpuesto9, nPorcentaje3, 0)
          lImp1 = RsProducto!lImpuesto7
          lImp2 = RsProducto!lImpuesto8
          lImp3 = RsProducto!lImpuesto9
       End If
    ElseIf sTipoPedido = "04" Then
       If IsNull(RsProducto!nPrecioCanal4) Or RsProducto!nPrecioCanal4 = 0 Then
          nCVenta = IIf(IsNull(RsProducto!nprecioVenta), 0, RsProducto!nprecioVenta) + (nLlevar * IIf(IsNull(RsProducto!nprecioVenta), 0, RsProducto!nprecioVenta) / 100)
       Else
          nCVenta = IIf(IsNull(RsProducto!nPrecioCanal4), 0, RsProducto!nPrecioCanal4)
          nValor = 0
          nValor = nValor + IIf(RsProducto!lImpuesto10, nPorcentaje1, 0)
          nValor = nValor + IIf(RsProducto!lImpuesto11, nPorcentaje2, 0)
          nValor = nValor + IIf(RsProducto!lImpuesto12, nPorcentaje3, 0)
          lImp1 = RsProducto!lImpuesto10
          lImp2 = RsProducto!lImpuesto11
          lImp3 = RsProducto!lImpuesto12
       End If
    ElseIf sTipoPedido = "05" Then
       If IsNull(RsProducto!nPrecioCanal5) Or RsProducto!nPrecioCanal5 = 0 Then
          nCVenta = IIf(IsNull(RsProducto!nprecioVenta), 0, RsProducto!nprecioVenta) + (nLlevar * IIf(IsNull(RsProducto!nprecioVenta), 0, RsProducto!nprecioVenta) / 100)
       Else
          nCVenta = IIf(IsNull(RsProducto!nPrecioCanal5), 0, RsProducto!nPrecioCanal5)
          nValor = 0
          nValor = nValor + IIf(RsProducto!lImpuesto10, nPorcentaje1, 0)
          nValor = nValor + IIf(RsProducto!lImpuesto11, nPorcentaje2, 0)
          nValor = nValor + IIf(RsProducto!lImpuesto12, nPorcentaje3, 0)
          lImp1 = RsProducto!lImpuesto10
          lImp2 = RsProducto!lImpuesto11
          lImp3 = RsProducto!lImpuesto12
       End If
    Else
       nCVenta = IIf(IsNull(RsProducto!nprecioVenta), 0, RsProducto!nprecioVenta)
    End If
    
    If sTipoPedido = "01" Then
       nInsumo = IIf(IsNull(RsProducto!nInsumo), 0, RsProducto!nInsumo)
       nGasto = IIf(IsNull(RsProducto!nGasto), 0, RsProducto!nGasto)
       nMano = IIf(IsNull(RsProducto!nManoObra), 0, RsProducto!nManoObra)
    ElseIf sTipoPedido = "02" Then
       nInsumo = IIf(IsNull(RsProducto!nInsumo2), 0, RsProducto!nInsumo2)
       nGasto = IIf(IsNull(RsProducto!nGasto2), 0, RsProducto!nGasto2)
       nMano = IIf(IsNull(RsProducto!nManoObra2), 0, RsProducto!nManoObra2)
    ElseIf sTipoPedido = "03" Then
       nInsumo = IIf(IsNull(RsProducto!nInsumo3), 0, RsProducto!nInsumo3)
       nGasto = IIf(IsNull(RsProducto!nGasto3), 0, RsProducto!nGasto3)
       nMano = IIf(IsNull(RsProducto!nManoObra3), 0, RsProducto!nManoObra3)
    ElseIf sTipoPedido = "04" Then
       nInsumo = IIf(IsNull(RsProducto!nInsumo4), 0, RsProducto!nInsumo4)
       nGasto = IIf(IsNull(RsProducto!nGasto4), 0, RsProducto!nGasto4)
       nMano = IIf(IsNull(RsProducto!nManoObra4), 0, RsProducto!nManoObra4)
    Else
       nInsumo = IIf(IsNull(RsProducto!nInsumo5), 0, RsProducto!nInsumo5)
       nGasto = IIf(IsNull(RsProducto!nGasto5), 0, RsProducto!nGasto5)
       nMano = IIf(IsNull(RsProducto!nManoObra5), 0, RsProducto!nManoObra5)
    End If
    nCVenta = IIf(RsProducto!tMoneda = "02", nCVenta * nTC, nCVenta)
    
   Select Case pais 'ok
        Case "001" 'Bolivia
            nValor = (nValor / 100)
            nCImp1 = IIf(lImp1, nCVenta * nPorcentaje1 / 100, 0)
            nCImp2 = IIf(lImp2, nCVenta * nPorcentaje2 / 100, 0)
            nCImp3 = IIf(lImp3, nCVenta * nPorcentaje3 / 100, 0)
            nCNeto = nCVenta - nCImp1 - nCImp2 - nCImp3
            
        Case Else 'Peru, Ecuador
            nValor = 1 + (nValor / 100)
            nCImp1 = IIf(lImp1, nCVenta / nValor * nPorcentaje1 / 100, 0)
            nCImp2 = IIf(lImp2, nCVenta / nValor * nPorcentaje2 / 100, 0)
            nCImp3 = IIf(lImp3, nCVenta / nValor * nPorcentaje3 / 100, 0)
            nCNeto = nCVenta - nCImp1 - nCImp2 - nCImp3
    End Select
    
    Isql = "insert into " & sComboDetalle & " " & _
           "(tCodigoPedido, tProducto, tItem, tItemCombo, tProductoCombo, nCantidad, tCodigoGrupo, tCodigoSubGrupo, nPrecioNeto, nImpuesto1, nImpuesto2, nImpuesto3, nVenta, nInsumo, nGasto, nManoObra, lImprimeArea, lImprime, nOrden ) " & _
           "Values(   '" & Pedido & "', " _
                   & "'" & RsDetalle!tCodigoProducto & "', " _
                   & "'" & sitem & "', " _
                   & "'" & xItem & "', " _
                   & "'" & wProducto & "', 1, " _
                   & "'" & IIf(IsNull(RsProducto!tgrupo), "", RsProducto!tgrupo) & "', " _
                   & "'" & IIf(IsNull(RsProducto!tSubGrupo), "", RsProducto!tSubGrupo) & "', " _
                   & nCNeto & ", " _
                   & nCImp1 & ", " _
                   & nCImp2 & ", " _
                   & nCImp3 & ", " _
                   & nCVenta & ", " _
                   & nInsumo & ", " & nGasto & ", " & nMano & ", " _
                   & IIf(RsProducto!lImprimeArea, -1, 0) & ", 0, " _
                   & RsProducto!nOrden & ") "
    Cn.Execute Isql
    'Oscar Ortega------------------------------------------------------------
    Isql = "select * from TCombo Where tCombo = '" & RsDetalle!tCodigoProducto & "' And tCodigoProducto = '" & wProducto & "'"
    Dim RstCombo As Recordset
    Set RstCombo = Lib.OpenRecordset(Isql, Cn)
    If RstCombo.RecordCount > 0 Then
        If IIf(IsNull(RstCombo!nAumento), 0, RstCombo!nAumento) > 0 Then
            txtMonto.Caption = Format(CambiaPrecio(nPVenta + RstCombo!nAumento / nCantidad), "#,###,##0.00")
        End If
    End If
    'Fin Oscar Ortega--------------------------------------------------------
    RsCombo.Requery
    RsCombo.MoveLast
    AsignaCombo
    Screen.MousePointer = vbDefault
End Sub

Public Sub fxCombo(Funcion As String, Cantidad As Double, Combo As String)
   Dim i As Integer
   Dim xItem As String
   Dim RsTemp As Recordset
   Dim nValor As Double
   Dim nCNeto As Double
   Dim nCImp1 As Double
   Dim nCImp2 As Double
   Dim nCImp3 As Double
   Dim nCVenta As Double
   Dim lImp1 As Boolean
   Dim lImp2 As Boolean
   Dim lImp3 As Boolean
   Dim nInsumo As Double
   Dim nGasto As Double
   Dim nMano As Double
      
   If sTipoPedido = "01" Then
        Isql = "SELECT dbo.TCOMBO.tCombo, dbo.TCOMBO.tCodigoProducto, dbo.TPRODUCTO.tGrupo, dbo.TPRODUCTO.tSubGrupo, dbo.TPRODUCTO.nPrecioVenta, dbo.TPRODUCTO.nPrecioLlevar, dbo.TPRODUCTO.nPrecioDelivery, dbo.TPRODUCTO.lImpuesto1, dbo.TPRODUCTO.lImpuesto2, dbo.TPRODUCTO.lImpuesto3, dbo.TPRODUCTO.lImpuesto4, dbo.TPRODUCTO.lImpuesto5, dbo.TPRODUCTO.lImpuesto6, dbo.TPRODUCTO.lImpuesto7, dbo.TPRODUCTO.lImpuesto8, dbo.TPRODUCTO.lImpuesto9, dbo.TPRODUCTO.tMoneda, dbo.TPRODUCTO.lImprimeArea, dbo.TSUBGRUPO.nOrden, dbo.TPRODUCTO.nInsumo As nInsumo, dbo.TPRODUCTO.nGasto As nGasto, dbo.TPRODUCTO.nManoObra As nManoObra, TCOMBO.nCantidad " & _
               "FROM dbo.TSUBGRUPO RIGHT OUTER JOIN dbo.TPRODUCTO ON dbo.TSUBGRUPO.tCodigoSubGrupo = dbo.TPRODUCTO.tSubGrupo RIGHT OUTER JOIN dbo.TCOMBO ON dbo.TPRODUCTO.tCodigoProducto = dbo.TCOMBO.tCodigoProducto Where dbo.TCOMBO.tCombo = '" & Combo & "' and dbo.TCOMBO.lFijo=1"
         
   ElseIf sTipoPedido = "02" Then
        Isql = "SELECT dbo.TCOMBO.tCombo, dbo.TCOMBO.tCodigoProducto, dbo.TPRODUCTO.tGrupo, dbo.TPRODUCTO.tSubGrupo, dbo.TPRODUCTO.nPrecioVenta, dbo.TPRODUCTO.nPrecioLlevar, dbo.TPRODUCTO.nPrecioDelivery, dbo.TPRODUCTO.lImpuesto1, dbo.TPRODUCTO.lImpuesto2, dbo.TPRODUCTO.lImpuesto3, dbo.TPRODUCTO.lImpuesto4, dbo.TPRODUCTO.lImpuesto5, dbo.TPRODUCTO.lImpuesto6, dbo.TPRODUCTO.lImpuesto7, dbo.TPRODUCTO.lImpuesto8, dbo.TPRODUCTO.lImpuesto9, dbo.TPRODUCTO.tMoneda, dbo.TPRODUCTO.lImprimeArea, dbo.TSUBGRUPO.nOrden, dbo.TPRODUCTO.nInsumo2 As nInsumo, dbo.TPRODUCTO.nGasto2 As nGasto, dbo.TPRODUCTO.nManoObra2 As nManoObra, TCOMBO.nCantidad " & _
               "FROM dbo.TSUBGRUPO RIGHT OUTER JOIN dbo.TPRODUCTO ON dbo.TSUBGRUPO.tCodigoSubGrupo = dbo.TPRODUCTO.tSubGrupo RIGHT OUTER JOIN dbo.TCOMBO ON dbo.TPRODUCTO.tCodigoProducto = dbo.TCOMBO.tCodigoProducto Where dbo.TCOMBO.tCombo = '" & Combo & "' and dbo.TCOMBO.lFijo=1"
   
   ElseIf sTipoPedido = "03" Then
        Isql = "SELECT dbo.TCOMBO.tCombo, dbo.TCOMBO.tCodigoProducto, dbo.TPRODUCTO.tGrupo, dbo.TPRODUCTO.tSubGrupo, dbo.TPRODUCTO.nPrecioVenta, dbo.TPRODUCTO.nPrecioLlevar, dbo.TPRODUCTO.nPrecioDelivery, dbo.TPRODUCTO.lImpuesto1, dbo.TPRODUCTO.lImpuesto2, dbo.TPRODUCTO.lImpuesto3, dbo.TPRODUCTO.lImpuesto4, dbo.TPRODUCTO.lImpuesto5, dbo.TPRODUCTO.lImpuesto6, dbo.TPRODUCTO.lImpuesto7, dbo.TPRODUCTO.lImpuesto8, dbo.TPRODUCTO.lImpuesto9, dbo.TPRODUCTO.tMoneda, dbo.TPRODUCTO.lImprimeArea, dbo.TSUBGRUPO.nOrden, dbo.TPRODUCTO.nInsumo3 As nInsumo, dbo.TPRODUCTO.nGasto3 As nGasto, dbo.TPRODUCTO.nManoObra3 As nManoObra, TCOMBO.nCantidad " & _
               "FROM dbo.TSUBGRUPO RIGHT OUTER JOIN dbo.TPRODUCTO ON dbo.TSUBGRUPO.tCodigoSubGrupo = dbo.TPRODUCTO.tSubGrupo RIGHT OUTER JOIN dbo.TCOMBO ON dbo.TPRODUCTO.tCodigoProducto = dbo.TCOMBO.tCodigoProducto Where dbo.TCOMBO.tCombo = '" & Combo & "' and dbo.TCOMBO.lFijo=1"
         
   ElseIf sTipoPedido = "04" Then
        Isql = "SELECT dbo.TCOMBO.tCombo, dbo.TCOMBO.tCodigoProducto, dbo.TPRODUCTO.tGrupo, dbo.TPRODUCTO.tSubGrupo, dbo.TPRODUCTO.nPrecioVenta, dbo.TPRODUCTO.nPrecioLlevar, dbo.TPRODUCTO.nPrecioDelivery, dbo.TPRODUCTO.lImpuesto1, dbo.TPRODUCTO.lImpuesto2, dbo.TPRODUCTO.lImpuesto3, dbo.TPRODUCTO.lImpuesto4, dbo.TPRODUCTO.lImpuesto5, dbo.TPRODUCTO.lImpuesto6, dbo.TPRODUCTO.lImpuesto7, dbo.TPRODUCTO.lImpuesto8, dbo.TPRODUCTO.lImpuesto9, dbo.TPRODUCTO.tMoneda, dbo.TPRODUCTO.lImprimeArea, dbo.TSUBGRUPO.nOrden, dbo.TPRODUCTO.nInsumo4 As nInsumo, dbo.TPRODUCTO.nGasto4 As nGasto, dbo.TPRODUCTO.nManoObra4 As nManoObra, TCOMBO.nCantidad " & _
               "FROM dbo.TSUBGRUPO RIGHT OUTER JOIN dbo.TPRODUCTO ON dbo.TSUBGRUPO.tCodigoSubGrupo = dbo.TPRODUCTO.tSubGrupo RIGHT OUTER JOIN dbo.TCOMBO ON dbo.TPRODUCTO.tCodigoProducto = dbo.TCOMBO.tCodigoProducto Where dbo.TCOMBO.tCombo = '" & Combo & "' and dbo.TCOMBO.lFijo=1"
         
   ElseIf sTipoPedido = "05" Then
        Isql = "SELECT dbo.TCOMBO.tCombo, dbo.TCOMBO.tCodigoProducto, dbo.TPRODUCTO.tGrupo, dbo.TPRODUCTO.tSubGrupo, dbo.TPRODUCTO.nPrecioVenta, dbo.TPRODUCTO.nPrecioLlevar, dbo.TPRODUCTO.nPrecioDelivery, dbo.TPRODUCTO.lImpuesto1, dbo.TPRODUCTO.lImpuesto2, dbo.TPRODUCTO.lImpuesto3, dbo.TPRODUCTO.lImpuesto4, dbo.TPRODUCTO.lImpuesto5, dbo.TPRODUCTO.lImpuesto6, dbo.TPRODUCTO.lImpuesto7, dbo.TPRODUCTO.lImpuesto8, dbo.TPRODUCTO.lImpuesto9, dbo.TPRODUCTO.tMoneda, dbo.TPRODUCTO.lImprimeArea, dbo.TSUBGRUPO.nOrden, dbo.TPRODUCTO.nInsumo5 As nInsumo, dbo.TPRODUCTO.nGasto5 As nGasto, dbo.TPRODUCTO.nManoObra5 As nManoObra, TCOMBO.nCantidad " & _
               "FROM dbo.TSUBGRUPO RIGHT OUTER JOIN dbo.TPRODUCTO ON dbo.TSUBGRUPO.tCodigoSubGrupo = dbo.TPRODUCTO.tSubGrupo RIGHT OUTER JOIN dbo.TCOMBO ON dbo.TPRODUCTO.tCodigoProducto = dbo.TCOMBO.tCodigoProducto Where dbo.TCOMBO.tCombo = '" & Combo & "' and dbo.TCOMBO.lFijo=1"
         
   End If
   
   Set RsTemp = Lib.OpenRecordset(Isql, Cn)
   If RsTemp.RecordCount = 0 Then
      Exit Sub
   End If
   RsCombo.Filter = "tItem='" & sitem & "'"
   
   Select Case Funcion
          Case Is = "A"
               Do While Not RsTemp.EOF
                  If RsCombo.RecordCount = 0 Then
                     xItem = "001"
                  Else
                     xItem = Lib.Correlativo(Calcular("select max(tItemCombo) as codigo from " & sComboDetalle & " where tItem = '" & sitem & "'", Cn), 3)
                  End If
               
                  nValor = 0
                  nValor = nValor + IIf(RsTemp!lImpuesto1, nPorcentaje1, 0)
                  nValor = nValor + IIf(RsTemp!lImpuesto2, nPorcentaje2, 0)
                  nValor = nValor + IIf(RsTemp!lImpuesto3, nPorcentaje3, 0)
                    
                  lImp1 = RsTemp!lImpuesto1
                  lImp2 = RsTemp!lImpuesto2
                  lImp3 = RsTemp!lImpuesto3
                  If sTipoPedido = "02" Then
                     If IsNull(RsProducto!nPrecioDelivery) Or RsProducto!nPrecioDelivery = 0 Then
                        nCVenta = IIf(IsNull(RsTemp!nprecioVenta), 0, RsTemp!nprecioVenta) + (nDELIVERY * IIf(IsNull(RsTemp!nprecioVenta), 0, RsTemp!nprecioVenta) / 100)
                     Else
                        nCVenta = IIf(IsNull(RsTemp!nPrecioDelivery), 0, RsTemp!nPrecioDelivery)
                        nValor = 0
                        nValor = nValor + IIf(RsTemp!lImpuesto4, nPorcentaje1, 0)
                        nValor = nValor + IIf(RsTemp!lImpuesto5, nPorcentaje2, 0)
                        nValor = nValor + IIf(RsTemp!lImpuesto6, nPorcentaje3, 0)
                        lImp1 = RsTemp!lImpuesto4
                        lImp2 = RsTemp!lImpuesto5
                        lImp3 = RsTemp!lImpuesto6
                     End If
                  ElseIf sTipoPedido = "03" Then
                     If IsNull(RsTemp!nPreciollevar) Or RsTemp!nPreciollevar = 0 Then
                        nCVenta = IIf(IsNull(RsTemp!nprecioVenta), 0, RsTemp!nprecioVenta) + (nLlevar * IIf(IsNull(RsTemp!nprecioVenta), 0, RsTemp!nprecioVenta) / 100)
                     Else
                        nCVenta = IIf(IsNull(RsTemp!nPreciollevar), 0, RsTemp!nPreciollevar)
                        nValor = 0
                        nValor = nValor + IIf(RsTemp!lImpuesto7, nPorcentaje1, 0)
                        nValor = nValor + IIf(RsTemp!lImpuesto8, nPorcentaje2, 0)
                        nValor = nValor + IIf(RsTemp!lImpuesto9, nPorcentaje3, 0)
                        lImp1 = RsTemp!lImpuesto7
                        lImp2 = RsTemp!lImpuesto8
                        lImp3 = RsTemp!lImpuesto9
                     End If
                  Else
                     nCVenta = IIf(IsNull(RsTemp!nprecioVenta), 0, RsTemp!nprecioVenta)
                  End If
                  nInsumo = IIf(IsNull(RsTemp!nInsumo), 0, RsTemp!nInsumo)
                  nGasto = IIf(IsNull(RsTemp!nGasto), 0, RsTemp!nGasto)
                  nMano = IIf(IsNull(RsTemp!nManoObra), 0, RsTemp!nManoObra)
               
                  nCVenta = IIf(RsTemp!tMoneda = "02", nCVenta * nTC, nCVenta)
               
                  Select Case pais 'ok
                        Case "001" 'Bolivia
                                nValor = (nValor / 100)
                                nCImp1 = IIf(lImp1, nCVenta * nPorcentaje1 / 100, 0)
                                nCImp2 = IIf(lImp2, nCVenta * nPorcentaje2 / 100, 0)
                                nCImp3 = IIf(lImp3, nCVenta * nPorcentaje3 / 100, 0)
                                nCNeto = nCVenta - nCImp1 - nCImp2 - nCImp3
                        
                        Case Else 'Peru, Ecuador
                                nValor = 1 + (nValor / 100)
                                nCImp1 = IIf(lImp1, nCVenta / nValor * nPorcentaje1 / 100, 0)
                                nCImp2 = IIf(lImp2, nCVenta / nValor * nPorcentaje2 / 100, 0)
                                nCImp3 = IIf(lImp3, nCVenta / nValor * nPorcentaje3 / 100, 0)
                                nCNeto = nCVenta - nCImp1 - nCImp2 - nCImp3
                        
                    End Select
                                     
                  Isql = "insert into " & sComboDetalle & _
                       " (tCodigoPedido, tProducto, tItem, tItemCombo, tProductoCombo, nCantidad, tCodigoGrupo, tCodigoSubGrupo, nPrecioNeto, nImpuesto1, nImpuesto2, nImpuesto3, nVenta, nInsumo, nGasto, nManoObra, lImprimeArea, lImprime, nOrden ) " & _
                         "Values(   '" & Pedido & "', " _
                                 & "'" & Combo & "', " _
                                 & "'" & sitem & "', " _
                                 & "'" & xItem & "', " _
                                 & "'" & RsTemp!tCodigoProducto & "', " & RsTemp!nCantidad & ", " _
                                 & "'" & IIf(IsNull(RsTemp!tgrupo), "", RsTemp!tgrupo) & "', " _
                                 & "'" & IIf(IsNull(RsTemp!tSubGrupo), "", RsTemp!tSubGrupo) & "', " _
                                 & nCNeto & ", " _
                                 & nCImp1 & ", " _
                                 & nCImp2 & ", " _
                                 & nCImp3 & ", " _
                                 & nCVenta & ", " _
                                 & nInsumo & ", " & nGasto & ", " & nMano & ", " _
                                 & IIf(RsProducto!lImprimeArea, -1, 0) & ", 0, " _
                                 & RsProducto!nOrden & ") "

                  Cn.Execute Isql
                  RsCombo.Requery
                  RsTemp.MoveNext
               Loop
          Case Is = "M"
               Do While Not RsTemp.EOF
                  Dim X As Double
                  X = Calcular("select nCantidad as Codigo FROM TCOMBO where tCombo='" & RsTemp!tCombo & "' and tCodigoproducto ='" & RsTemp!tCodigoProducto & "'", Cn)
                  Isql = "update " & sComboDetalle & " set nCantidad = " & X * Cantidad & " where tCodigoPedido='" & Pedido & "' and tItem='" & sitem & "' and tProductocombo='" & RsTemp!tCodigoProducto & "'"
                  Cn.Execute Isql
                  RsCombo.Requery
                  RsTemp.MoveNext
               Loop
          
          Case Is = "D"
               For i = 1 To Cantidad
                   Isql = "DELETE from " & sComboDetalle & " " & _
                          "where tCodigoPedido ='" & Pedido & "' and tProducto='" & Combo & "' and tItem='" & sitem & "'"
                   Cn.Execute Isql
                   RsCombo.Requery
               Next i
   End Select
   Set RsTemp = Nothing
End Sub

Public Sub Inicializar()
  'sMozo = "0000"
   sCodigoDescuento = ""
   sClienteFrecuente = ""
   sCodigoParienteSeleccionado = ""
   sCodigoInvitado = ""
   xDescuento = 0
   nTope = 0
   ltope = False
   tAutorizaDescuento = ""
   sObser = ""
   sDescripcionDescuento = ""
   xDescuento = 0
   txtTitulo.Caption = " Caja Rápida : " & sCaja & " Mesero: Sin Mesero"
   txtObservacion.Caption = sObser
   txtEntregar.Caption = ""
   fraMozo.Visible = False
   fraDetalle.Visible = False
   fraPropiedad.Visible = False
   Pedido = ""
   wCombo = False
   wAgregaCombo = False
   nCombo = 0
   nMonto = 0
   Pedido = ""
   txtTelefono.Caption = ""
   txtCliente.Caption = ""
   VisualizaMonto
   Cn.Execute "delete from " & sComboDetalle
   RsCombo.Requery
   txtFechaEntrega.Caption = ""
End Sub

Private Sub cmdEliminacion_Click(Index As Integer)
    RsMotivoEliminacion.MoveFirst
    RsMotivoEliminacion.Find ("Descripcion = '" & cmdEliminacion(Index).Caption & "'")
    
    If RsMotivoEliminacion.EOF Then
       RsMotivoEliminacion.MoveFirst
    End If
    
    If RsMotivoEliminacion!codigo = "000" Then
       frmKeyBoard.txtResultado = ""
       frmKeyBoard.Show vbModal
       If Not wEnter Then
          Exit Sub
       End If
       sCodigo = "000"
    Else
       sCodigo = RsMotivoEliminacion!codigo
       sDescrip = ""
    End If
    
    fraEliminacion.Visible = False
    tabProducto.Visible = True
    If Sw Then
       EliminaCabecera
    Else
         'KDS2
        If lKDS Then
            Dim kdsRsCabecera As Recordset
            Isql = "SELECT * From vPedidoCabecera Where Codigo = '" & Pedido & "' Order By codigo "
            Set kdsRsCabecera = Lib.OpenRecordset(Isql, Cn)
            Call KDS_EliminarProducto(kdsRsCabecera, sitem)
        End If
       EliminaItem
    End If
    Sw = False
    ActivaCabecera True
End Sub

Public Sub ActivaCabecera(Activa As Boolean)
   Dim i As Integer
   For i = 1 To 4
       cmdTipoDocumento(i).Enabled = Activa
       cmdNavegar(i).Enabled = Activa
   Next i
   
   cmdDetalle(1).Enabled = Activa
   cmdDetalle(2).Enabled = Activa
   cmdDetalle(3).Enabled = Activa
   
   cmdNavegar(5).Enabled = Activa
   cmdNavegar(6).Enabled = Activa
   cmdDetalle(0).Enabled = Activa
   cmdDetalle(5).Enabled = Activa
   cmdDetalle(6).Enabled = Activa
   cmdDetalle(7).Enabled = Activa
   cmdDetalle(8).Enabled = Activa
   'cmdDetalle(9).Enabled = Activa
   cmdDetalle(10).Enabled = Activa
   cmdDetalle(12).Enabled = Activa
   cmdDetalle(13).Enabled = Activa
   cmdOpcion(0).Enabled = Activa
   cmdOpcion(1).Enabled = Activa
   cmdOpcion(2).Enabled = IIf(lMultiCajero, False, Activa)
   cmdOpcion(5).Enabled = Activa
   cmdOpcion(10).Enabled = Activa
   cmdCabecera(1).Enabled = Activa
   cmdCabecera(2).Enabled = Activa
   
   cmdOpcion(14).Enabled = Activa
   
   If lPrinter Then
      cmdOpcion(8).Enabled = Activa
   End If

End Sub

Public Sub EliminaItem()
    Dim xMax As String
    Dim sMotivo As String
    
    If lPrinter Then
       'Impresion del Pedidos Anulados
       sMotivo = Calcular("select Descripcion as Codigo from vMotivoEliminacion where Codigo='" & sCodigo & "'", Cn)
       
       Isql = "select *, '" & sMotivo & "' as MotivoEliminacion FROM dbo.vPedido " & _
              "WHERE Codigo = '" & Pedido & "' and tItem = '" & sitem & "' and lImprime = 1 And lImprimeArea = 1 " & _
              "ORDER BY tItem,tetiqueta,combo"
                                                                                             
       Set RsImpresion = Lib.OpenRecordset(Isql, Cn)
       Dim i As Integer
       If RsImpresion.RecordCount = 0 Then
          LimpiaRs
       Else
          RsArea.MoveFirst
          For i = 1 To RsArea.RecordCount
              RsImpresion.Filter = "tArea = '" & RsArea!tArea & "'"
              If RsArea!tIcono = "" Or sSalon = RsArea!tIcono Or (sSalon = "" And RsArea!nValor = 1) Then
                 If RsImpresion.RecordCount <> 0 Then
                    RsImpresion.MoveFirst
                    sPedido = Pedido
                    ImprimePedido RsImpresion, "A", RsArea!timpresora, RsArea!Area, False, RsProductoPropiedad, RsProductoPropiedad, "Rapido"
                    sPedido = ""
                 End If
              End If
              RsArea.MoveNext
          Next i
          RsDetalle.Requery
       End If
    End If
    
    xMax = Calcular("select max(tItem) as Codigo from APEDIDO where tCodigoPedido='" & Pedido & "'", Cn)
    xMax = Lib.Correlativo(xMax, 3)
    Isql = "insert into APEDIDO (tCodigoPedido, tItem, tCodigoProducto, tCodigoGrupo, tCodigoSubGrupo, " & _
           "nPrecioNeto, nPrecioImpuesto1, nPrecioImpuesto2, nPrecioImpuesto3, nPrecioVenta, nRecargo, " & _
           "nDescuento, nPrecioOficial, nCantidad, nImpuesto1, nImpuesto2, nImpuesto3, nVenta, " & _
           "tComanda, lImprime, tUsuario, fRegistro, tUsuarioAnulado, fRegistroAnulado, " & _
           "tObservacion, tObservacionAnulado, tEstadoItem, lImprimeArea, tArea, tMotivoEliminacion, tTurnoAnulado,fDiaContable) " & _
           "select '" & Pedido & "' as tCodigoPedido, '" & xMax & "' as tItem, tCodigoProducto, tCodigoGRupo, tCodigoSubGrupo, " & _
           "nPrecioNeto, nPrecioImpuesto1, nPrecioImpuesto2, nPrecioImpuesto3, nPrecioVenta, nRecargo, " & _
           "nDescuento, nPrecioOficial, nCantidad, nImpuesto1, nImpuesto2, nImpuesto3, nVenta, tComanda, lImprime, " & _
           "'" & Mid(sUsuario, 1, 15) & "' as tUsuario, getDate() as fRegistro, " & _
           "'" & sUsuarioAutoriza & "' as tUsuarioAnulado, getDate() as fRegistroAnulado, " & _
           "tObservacion, '" & sDescrip & "' as tObservacion, tEstadoItem, lImprimeArea, tArea, '" & sCodigo & "', '" & sTurno & "','" & Format(obtieneDiaContable, "yyyyMMdd") & "' " & _
           "from " & sDetalle & _
           " where tCodigoPedido = '" & Pedido & "' and tItem = '" & sitem & "'"
    Cn.Execute Isql
    
        
 
      'INSUMOCRITICO23
    Dim rstItems As New ADODB.Recordset
    Set rstItems = New ADODB.Recordset
    'Set rstItems = Lib.OpenRecordset("select tcodigoinsumo,ncantidad from dpedido inner join tproducto on dpedido.tcodigoproducto=tproducto.tcodigoproducto where tcodigopedido='" & sPedido & "' and titem='" & sitem & "' and tproducto.lControlInsumoCritico=1 and isnull(tproducto.tcodigoinsumo,'')<>''  and isnull(dpedido.limprime,0)=1 ", Cn)
    Set rstItems = Lib.OpenRecordset("  usp_Inforest_RevertirInsumosCriticos '" & sPedido & "','" & sitem & "' ", Cn)
    If Not (rstItems.EOF Or rstItems.BOF) Then
        modificaStockInsumo rstItems.Fields(0), rstItems.Fields(1), "I"
    End If
       
    Cn.Execute "delete from DPEDIDO where tCodigoPedido = '" & Pedido & "' and tItem = '" & sitem & "'"
    Cn.Execute "delete TPRODUCTOPROPIEDAD where tCodigoPedido='" & Pedido & "' and tItem = '" & sitem & "'"
    Cn.Execute "delete CPEDIDO where tCodigoPedido='" & Pedido & "' and tItem = '" & sitem & "'"
    Cn.Execute "delete TCOMBOPROPIEDAD where tCodigoPedido='" & Pedido & "' and tItem = '" & sitem & "'"
    
    Cn.Execute "delete from " & sDetalle & " where tItem = '" & sitem & "'"
    Cn.Execute "delete from " & sComboDetalle & " where tItem ='" & sitem & "'"
    Cn.Execute "delete from " & sProductoPropiedad & " where tItem ='" & sitem & "'"
    Cn.Execute "delete from " & sComboPropiedad & " where tItem ='" & sitem & "'"
                          
    RsDetalle.Requery
    RsCombo.Requery
    RsPropiedad.Requery
    RsProductoPropiedad.Requery
        
    nMonto = Calcular("select sum(nVenta) as Codigo FROM [" & sDetalle & "] where tEstadoItem = 'N'", Cn)
    If RsDetalle.RecordCount = 0 Then
       nMonto = 0
       sProducto = ""
       wCombo = False
       nCombo = 0
       sitem = ""
    Else
       RsDetalle.MoveLast
       sitem = RsDetalle!tItem
       nMonto = Calcular("select sum(nVenta) as Codigo FROM [" & sDetalle & "] where tEstadoItem = 'N'", Cn)
    End If
    VisualizaMonto
End Sub

Public Sub SoloEliminaItem()
    'INSUMOCRITICO2013
    Dim rstItems As New ADODB.Recordset
    Set rstItems = New ADODB.Recordset
    Cn.Execute "DELETE FROM " & sInsumoCombo
    Cn.Execute "INSERT INTO " & sInsumoCombo & " select tcodigoinsumo,ncantidad from " & sDetalle & " inner join tproducto on " & sDetalle & ".tcodigoproducto=tproducto.tcodigoproducto where titem='" & sitem & "' and tproducto.lControlInsumoCritico=1 and isnull(tproducto.tcodigoinsumo,'')<>''  and isnull(" & sDetalle & ".limprime,0)=1"
    Cn.Execute "insert into " & sInsumoCombo & " SELECT     " & sDetalle & ".nCantidad * " & sComboDetalle & ".nCantidad AS nCantidad , dbo.TPRODUCTO.tCodigoInsumo FROM   " & sDetalle & " INNER JOIN  " & sComboDetalle & " ON " & sDetalle & ".tItem = " & sComboDetalle & ".tItem INNER JOIN dbo.TPRODUCTO ON " & sComboDetalle & ".tProductoCombo = dbo.TPRODUCTO.tCodigoProducto WHERE   " & sComboDetalle & ".tItem = '" & sitem & "' and  (dbo.TPRODUCTO.lControlInsumoCritico = 1) AND (ISNULL(dbo.TPRODUCTO.tCodigoInsumo, '') <> '') AND (ISNULL(" & sComboDetalle & ".lImprime, 0) = 1) "



    Set rstItems = Lib.OpenRecordset("select tCodigoInsumo,  SUM(ncantidad) as ncantidad from " & sInsumoCombo & "  group by tCodigoInsumo order by 2 ", Cn)
    
    If Not (rstItems.EOF Or rstItems.BOF) Then
        modificaStockInsumo rstItems.Fields(0), rstItems.Fields(1), "I"
    End If
    'INSUMOCRITICO
    

    Cn.Execute "delete from " & sDetalle & " where tItem = '" & sitem & "'"
    Cn.Execute "delete from " & sComboDetalle & " where tItem ='" & sitem & "'"
    Cn.Execute "delete from " & sComboPropiedad & " where tItem ='" & sitem & "'"
    Cn.Execute "delete from " & sProductoPropiedad & " where tItem='" & sitem & "'"
                              
    If RsDetalle.RecordCount <> 0 Then
       nMonto = nMonto - (grdDetalle.Columns(4).Text * nPVenta)
    Else
       nMonto = 0
    End If
    
    RsProductoPropiedad.Requery
    RsComboPropiedad.Requery
    RsCombo.Requery
    RsDetalle.Requery
    
    If RsDetalle.RecordCount = 0 Then
       txtMonto.Caption = "0.00"
       sProducto = ""
       wCombo = False
       nCombo = 0
       sitem = ""
    Else
       nMonto = Calcular("select sum(nVenta) as Codigo FROM [" & sDetalle & "] where tEstadoItem = 'N'", Cn)
       RsDetalle.MoveLast
       sitem = RsDetalle!tItem
    End If
    VisualizaMonto
End Sub

Public Sub VisualizaMonto()
   txtMonto.Caption = Format(nMonto, "#,###,##0.00")
   txtMontoLetras.Caption = NumeroCadena(str(nMonto))
   If nPuerto > 0 Then
      If RsDetalle.RecordCount > 0 Then
         Visor "Total:" & sMonN & Right(String(10, " ") & Format(nMonto, "###,##0.00,"), 10), RsDetalle!nCantidad & " " & RsDetalle!Producto, nPuerto, "N"
      End If
   End If
   txtDescuento.Caption = Format(Calcular("select sum(nDescuento*nCantidad) as Codigo FROM " & sDetalle, Cn), "#,###,##0.00")
    If lvisor Then
        Call InsertVisor8
    End If
End Sub

Public Sub ActualizaPedido()
      'Actualiza el Numero de Pedido en el Detalle Temporal
                  Dim oComando As clsComando
                  Set oComando = New clsComando
                  If Not oComando.CreateCmdSp("spUpd_MPEDIDO", Cn) Then
                     Set oComando = Nothing
                     Exit Sub
                  End If
                        oComando.CreateParameter "@tCliente", adVarChar, adParamInput, 7, sClienteFrecuente
                        oComando.CreateParameter "@tTipoPedido", adVarChar, adParamInput, 2, sTipoPedido
                        oComando.CreateParameter "@lPrioridad", adBoolean, adParamInput, 1, 1
                        oComando.CreateParameter "@tTipoAtencion", adVarChar, adParamInput, 2, "01"
                        oComando.CreateParameter "@tMozo", adVarChar, adParamInput, 4, Right(sMozo, 4)
                        oComando.CreateParameter "@tMotorizado", adVarChar, adParamInput, 4, "000"
                        oComando.CreateParameter "@tObservacion", adVarChar, adParamInput, 250, txtObservacion.Caption
                        oComando.CreateParameter "@nTiempo", adInteger, adParamInput, 10, 0
                        oComando.CreateParameter "@tPuntoVenta", adVarChar, adParamInput, 2, ""
                        oComando.CreateParameter "@tHabitacion", adVarChar, adParamInput, 6, ""
                        oComando.CreateParameter "@tReserva", adVarChar, adParamInput, 6, ""
                        oComando.CreateParameter "@tPasajero", adVarChar, adParamInput, 50, ""
                        oComando.CreateParameter "@tCompania", adVarChar, adParamInput, 5, ""
                        oComando.CreateParameter "@tContacto", adVarChar, adParamInput, 4, ""
                        oComando.CreateParameter "@nDescuento", adDouble, adParamInput, 10, xDescuento
                        oComando.CreateParameter "@tDescuento", adVarChar, adParamInput, 3, sCodigoDescuento
                        oComando.CreateParameter "@tObservacionDescuento", adVarChar, adParamInput, 250, IIf(sCodigoDescuento = "000", sDescripcionDescuento, "")
                        oComando.CreateParameter "@tAutorizaDescuento", adVarChar, adParamInput, 15, IIf(sCodigoDescuento = "", "", tAutorizaDescuento)
                        oComando.CreateParameter "@tTienda", adVarChar, adParamInput, 3, ""
                        oComando.CreateParameter "@fProgramacion", adDate, adParamInput, 20, IIf(txtFechaEntrega.Caption = "", Null, Format(txtFechaEntrega.Caption, "dd/MM/yyyy HH:mm"))
                        oComando.CreateParameter "@tCodigoInvitado", adVarChar, adParamInput, 10, sCodigoInvitado
                        oComando.CreateParameter "@tCodigopariente", adVarChar, adParamInput, 7, sCodigoParienteSeleccionado
                        oComando.CreateParameter "@tEntregarA", adVarChar, adParamInput, 20, IIf(Len(txtEntregar.Caption) = 0, "", Left(Me.txtEntregar.Caption, 20))
                        oComando.CreateParameter "@nTiempoAntesEnvio", adInteger, adParamInput, 10, 0
                        oComando.CreateParameter "@nMontoMaximo", adInteger, adParamInput, 250, 0
                        oComando.CreateParameter "@tPedido", adVarChar, adParamInput, 10, Pedido
                        
                        'origen de ventas
                        If vOrigenVentas = Null Then
                            vOrigenVentas = "00"
                        End If
                        oComando.CreateParameter "@codigoOrigenVentas", adVarChar, adParamInput, 2, vOrigenVentas
                      
                  If Not oComando.GetParamOK Then
                     Set oComando = Nothing
                     Exit Sub
                  End If
                  If Not oComando.ExecSP Then
                     Set oComando = Nothing
                     Exit Sub
                  End If
                  

      'Cn.Execute "Update MPEDIDO set fProgramacion = '" & IIf(Len(txtFechaEntrega.Caption) = 0, "", Left(Me.txtFechaEntrega.Caption, 16)) & "' , tclientedelivery='" & sClienteFrecuente & "',tcodigopariente='" & sCodigoParienteSeleccionado & "', tcodigoinvitado='" & sCodigoInvitado & "' , tObservacion='" & txtObservacion.Caption & "', tMozo='" & sMozo & "', tentregara='" & IIf(Len(txtEntregar.Caption) = 0, "", Left(Me.txtEntregar.Caption, 20)) & "' where tCodigoPedido='" & Pedido & "'"
      
      Cn.Execute "Update [" & sDetalle & "] Set tCodigoPedido = '" & Pedido & "'"
      
      'Inserta el Detalle
      Cn.Execute "delete DPEDIDO where tCodigoPedido = '" & Pedido & "' and lImprime=0"
      
      Cn.Execute "Insert into DPEDIDO (tCodigoPedido, tItem, tTipoPedido, tCodigoProducto, tCodigoGrupo, tCodigoSubGrupo, tMoneda, " & _
                 "nPrecioNeto, nPrecioImpuesto1, nPrecioImpuesto2, nPrecioImpuesto3, nPrecioVenta, nRecargo, nDescuento, nPrecioOficial, " & _
                 "nCantidad, nImpuesto1, nImpuesto2, nImpuesto3, nVenta, tObservacion, tCortesia, lImprime, tEstadoItem, tArea, lImprimeArea,fregistro, nInsumo, nGasto, nManoObra, nOrden, lCombinacion, nCombinacion, lCorte,toferta,tautorizaoferta,tSubalmacen,tCodigoEtiqueta,tunidadnegocio,fDiaContable, fEnvio, nEnvio, tCajaD ) " & _
                 "select tCodigoPedido, tItem, tTipoPedido, tCodigoProducto, tCodigoGrupo, tCodigoSubGrupo, tMoneda, " & _
                 "nPrecioNeto, nPrecioImpuesto1, nPrecioImpuesto2, nPrecioImpuesto3, nPrecioVenta, nRecargo, nDescuento, nPrecioOficial, " & _
                 "nCantidad, nImpuesto1, nImpuesto2, nImpuesto3, nVenta, tObservacion, tCortesia, lImprime, tEstadoItem, tArea, lImprimeArea,getdate(), nInsumo, nGasto, nManoObra, nOrden, lCombinacion, nCombinacion, lCorte,toferta,tautorizaoferta,tsubalmacen,tCodigoEtiqueta,'" & sUnidadNegocio & "' ,'" & Format(obtieneDiaContable, "yyyyMMdd") & "', fEnvio, nEnvio, '" & sCaja & "' " & _
                 "From [" & sDetalle & "] where tEstadoItem='N' and lImprime=0"
      
      'Actualiza el Numero de Pedido en el Detalle Combos
      Cn.Execute "Update [" & sComboDetalle & "] Set tCodigoPedido = '" & Pedido & "'"
      
      'Inserta Combo
      Cn.Execute "delete CPEDIDO where tCodigoPedido = '" & Pedido & "'"
      Cn.Execute "Insert into CPEDIDO select * from " & sComboDetalle
      
      'Inserta las propiedades
      Cn.Execute "delete TPRODUCTOPROPIEDAD where tCodigoPedido = '" & Pedido & "'"
      Cn.Execute "Insert into TPRODUCTOPROPIEDAD select '" & Pedido & "', tItem,  tCodigoPropiedad, tProducto, tEnlace, nInsumo, nGasto, nManoObra, 1,ncantidad,ninsumounitario,ngastounitario,nmanoobraunitario from " & sProductoPropiedad
      
      'Inserta las propiedades de los Combos
      Cn.Execute "delete tcombopropiedad where tCodigoPedido = '" & Pedido & "'"
      Cn.Execute "Insert into TCOMBOPROPIEDAD select '" & Pedido & "', tItem, tItemCombo, tCodigoPropiedad, tProducto, tEnlace, nInsumo, nGasto, nManoObra, 1,ncantidad,ninsumounitario,ngastounitario,nmanoobraunitario from " & sComboPropiedad
End Sub

Private Sub cmdPunto_Click(Index As Integer)
   Dim i As Integer
   For i = 1 To 9
       cmdPunto(i).FontBold = False
   Next i
   cmdPunto(Index).FontBold = True
   rsPuntoVenta.MoveFirst
   rsPuntoVenta.Find "Descripcion = '" & cmdPunto(Index).Caption & "'"
   cmdCabecera(0).Caption = rsPuntoVenta!Descripcion
   sPuntoVenta = rsPuntoVenta!codigo
   tabProducto.Visible = True
   fraPuntoVenta.Visible = False
End Sub

Public Sub AsignaComboPropiedad()
    Dim i As Integer
    If RsOperador.RecordCount > 0 Then
        If RsOperador.EOF Then
            RsOperador.MoveFirst
        End If
       RsPropiedad.Filter = "tOperador = '" & RsOperador!codigo & "' and tProducto='" & sCombo & "'"
    Else
       RsPropiedad.Filter = "tOperador = '  ' and tProducto='" & sCombo & "'"
    End If
    AsignaComando 19, RsPropiedad, cmdPropiedad()

    For i = 1 To 19
        cmdPropiedad(i).FontBold = False
    Next i
    
    lblResumen.Text = ""
    RsComboPropiedad.Filter = "tItem='" & sitem & "' and tItemCombo='" & xItem & "'"
    If Not RsComboPropiedad.EOF Then
       RsComboPropiedad.MoveFirst
       Do While Not RsComboPropiedad.EOF
          For i = 1 To 19
              If cmdPropiedad(i).Caption = RsComboPropiedad!Descripcion And RsOperador!Descripcion = RsComboPropiedad!Operador Then
                 cmdPropiedad(i).FontBold = True
                 Exit For
              End If
          Next i
          
          If RsComboPropiedad!nCantidad = 1 Then
            lblResumen.Text = lblResumen.Text & LTrim(RsComboPropiedad!Operador) & " " & LTrim(RsComboPropiedad!Descripcion) & ", "
          Else
            lblResumen.Text = lblResumen.Text & LTrim(RsComboPropiedad!Operador) & " " & LTrim(RsComboPropiedad!Descripcion) & ": (" & RsComboPropiedad!nCantidad & "), "
          End If
          
          'lblResumen.Text = lblResumen.Text & LTrim(RsComboPropiedad!Operador) & " " & LTrim(RsComboPropiedad!Descripcion) & ", "
          RsComboPropiedad.MoveNext
       Loop
    End If
End Sub

Public Function CambiaPrecio(Valor As Double)
    nPVenta = Val(Valor)
    nOficial = nPVenta
    Dim Acumulado As Double
    Select Case pais 'ok
        Case "001" 'Bolivia
                    Acumulado = 0
                    Acumulado = IIf(txtImpuesto1.Caption <> 0, Acumulado + nPorcentaje1, Acumulado)
                    Acumulado = IIf(txtImpuesto2.Caption <> 0, Acumulado + nPorcentaje2, Acumulado)
                    Acumulado = IIf(txtImpuesto3.Caption <> 0, Acumulado + nPorcentaje3, Acumulado)
                    Acumulado = (Acumulado / 100)
                    
                    nImpuesto1 = IIf(txtImpuesto1.Caption <> 0, nPVenta * nPorcentaje1 / 100, 0)
                    nImpuesto2 = IIf(txtImpuesto2.Caption <> 0, nPVenta * nPorcentaje2 / 100, 0)
                    nImpuesto3 = IIf(txtImpuesto3.Caption <> 0, nPVenta * nPorcentaje3 / 100, 0)
                    
                    nPBase = nPVenta - nImpuesto1 - nImpuesto2 - nImpuesto3
        
        Case Else 'Peru, Ecuador
                    Acumulado = 0
                    Acumulado = IIf(txtImpuesto1.Caption <> 0, Acumulado + nPorcentaje1, Acumulado)
                    Acumulado = IIf(txtImpuesto2.Caption <> 0, Acumulado + nPorcentaje2, Acumulado)
                    Acumulado = IIf(txtImpuesto3.Caption <> 0, Acumulado + nPorcentaje3, Acumulado)
                    Acumulado = 1 + (Acumulado / 100)
                    
                    nImpuesto1 = IIf(txtImpuesto1.Caption <> 0, nPVenta / Acumulado * nPorcentaje1 / 100, 0)
                    nImpuesto2 = IIf(txtImpuesto2.Caption <> 0, nPVenta / Acumulado * nPorcentaje2 / 100, 0)
                    nImpuesto3 = IIf(txtImpuesto3.Caption <> 0, nPVenta / Acumulado * nPorcentaje3 / 100, 0)
                    
                    nPBase = nPVenta - nImpuesto1 - nImpuesto2 - nImpuesto3
        
        
    End Select
    
    Isql = "Update " & sDetalle & " Set nPrecioNeto = " & nPBase & ", " & _
           "nPrecioOficial = " & nOficial & ", " & _
           "nprecioImpuesto1 = " & nImpuesto1 & ", " & _
           "nprecioImpuesto2 = " & nImpuesto2 & ", " & _
           "nprecioImpuesto3 = " & nImpuesto3 & ", " & _
           "nPrecioVenta = " & nPVenta & ", " & _
           "nventa = " & nPVenta * nCantidad & ", " & _
           "nCantidad = " & nCantidad & ", " & _
           "nImpuesto1 = " & nImpuesto1 * nCantidad & ", " & _
           "nImpuesto2 = " & nImpuesto2 * nCantidad & ", " & _
           "nImpuesto3 = " & nImpuesto3 * nCantidad & " " & _
           "where tItem = '" & sitem & "'"
           Cn.Execute Isql
    CambiaPrecio = Calcular("select sum(nVenta) as Codigo FROM " & sDetalle, Cn)
    txtMonto.Caption = Format(CambiaPrecio, "#,###,##0.00")
    nMonto = CambiaPrecio
End Function



Public Function CambiaPrecioCombo(Valor As Double)
    nPVenta = Val(Valor)
    'nOficial = nPVenta
    Dim Acumulado As Double
    Select Case pais 'ok
        Case "001" 'Bolivia
                    Acumulado = 0
                    Acumulado = IIf(txtImpuesto1.Caption <> 0, Acumulado + nPorcentaje1, Acumulado)
                    Acumulado = IIf(txtImpuesto2.Caption <> 0, Acumulado + nPorcentaje2, Acumulado)
                    Acumulado = IIf(txtImpuesto3.Caption <> 0, Acumulado + nPorcentaje3, Acumulado)
                    Acumulado = (Acumulado / 100)
                    
                    nImpuesto1 = IIf(txtImpuesto1.Caption <> 0, nPVenta * nPorcentaje1 / 100, 0)
                    nImpuesto2 = IIf(txtImpuesto2.Caption <> 0, nPVenta * nPorcentaje2 / 100, 0)
                    nImpuesto3 = IIf(txtImpuesto3.Caption <> 0, nPVenta * nPorcentaje3 / 100, 0)
                    
                    nPBase = nPVenta - nImpuesto1 - nImpuesto2 - nImpuesto3
        
        Case Else 'Peru, Ecuador
                    Acumulado = 0
                    Acumulado = IIf(txtImpuesto1.Caption <> 0, Acumulado + nPorcentaje1, Acumulado)
                    Acumulado = IIf(txtImpuesto2.Caption <> 0, Acumulado + nPorcentaje2, Acumulado)
                    Acumulado = IIf(txtImpuesto3.Caption <> 0, Acumulado + nPorcentaje3, Acumulado)
                    Acumulado = 1 + (Acumulado / 100)
                    
                    nImpuesto1 = IIf(txtImpuesto1.Caption <> 0, nPVenta / Acumulado * nPorcentaje1 / 100, 0)
                    nImpuesto2 = IIf(txtImpuesto2.Caption <> 0, nPVenta / Acumulado * nPorcentaje2 / 100, 0)
                    nImpuesto3 = IIf(txtImpuesto3.Caption <> 0, nPVenta / Acumulado * nPorcentaje3 / 100, 0)
                    
                    nPBase = nPVenta - nImpuesto1 - nImpuesto2 - nImpuesto3
        
        
    End Select
    '"nPrecioOficial = " & nOficial & ", " &
    
    Isql = "Update " & sDetalle & " Set nPrecioNeto = " & nPBase & ", " & _
           "nprecioImpuesto1 = " & nImpuesto1 & ", " & _
           "nprecioImpuesto2 = " & nImpuesto2 & ", " & _
           "nprecioImpuesto3 = " & nImpuesto3 & ", " & _
           "nPrecioVenta = " & nPVenta & ", " & _
           "nventa = " & nPVenta * nCantidad & ", " & _
           "nCantidad = " & nCantidad & ", " & _
           "nImpuesto1 = " & nImpuesto1 * nCantidad & ", " & _
           "nImpuesto2 = " & nImpuesto2 * nCantidad & ", " & _
           "nImpuesto3 = " & nImpuesto3 * nCantidad & " " & _
           "where tItem = '" & sitem & "'"
           Cn.Execute Isql
           
    CambiaPrecioCombo = Calcular("select sum(nVenta) as Codigo FROM " & sDetalle, Cn)
    txtMonto.Caption = Format(CambiaPrecioCombo, "#,###,##0.00")
    nMonto = CambiaPrecioCombo
End Function
Public Sub EliminaCabecera()
    'KDS2
    If (lKDS = True) Then
        Dim kdsRsCabecera As Recordset
        Isql = "SELECT * From vPedidoCabecera Where Codigo = '" & Pedido & "' Order By codigo "
        Set kdsRsCabecera = Lib.OpenRecordset(Isql, Cn)
        Call KDS_EliminarOrden(kdsRsCabecera)
    End If
    
   Dim i As Integer
   Screen.MousePointer = vbHourglass
   Dim sMotivo As String
   
   If lPrinter Then
      sMotivo = Calcular("select Descripcion as Codigo from vMotivoEliminacion where Codigo='" & sCodigo & "'", Cn)
      Isql = "select *, Descripcion as MotivoEliminacion FROM dbo.vPedido LEFT OUTER JOIN dbo.vMotivoEliminacion ON dbo.vPedido.tMotivoEliminacion = dbo.vMotivoEliminacion.Codigo " & _
             "WHERE vPedido.Codigo='" & sPedido & "' AND lImprime=1 AND lImprimeArea=1 " & _
             "ORDER BY tItem,tetiqueta,combo"
                     
       Set RsImpresion = Lib.OpenRecordset(Isql, Cn)
       
       If Not RsImpresion.EOF Then
          RsArea.MoveFirst
          For i = 1 To RsArea.RecordCount
              RsImpresion.Filter = "tArea = '" & RsArea!tArea & "'"
              If RsArea!tIcono = "" Or sSalon = RsArea!tIcono Or (sSalon = "" And RsArea!nValor = 1) Then
                 If RsImpresion.RecordCount <> 0 Then
                    RsImpresion.MoveFirst
                    'sPedido = Pedido
                    ImprimePedido RsImpresion, "A", RsArea!timpresora, RsArea!Area, False, RsProductoPropiedad, RsProductoPropiedad, "Rapido"
                    'sPedido = ""
                 End If
              End If
              RsArea.MoveNext
          Next i
       End If
   End If
   
   If lInfhotel Then
      CnInfhotel.Execute "update MCOMANDA set TESTADO ='04', TOBSERVACIONANULA = 'Anulado por Inforest - " & sUsuarioAutoriza & " " & Pedido & " - " & Trim(sDescrip) & "' where tComanda ='" & sComandaInfhotel & "'"
   End If

   'INSUMOCRITICO
   Cn.Execute "delete from " & sInsumoCombo
   
    Dim rstItems As New ADODB.Recordset
    Dim j As Integer
    Set rstItems = New ADODB.Recordset
    Cn.Execute "insert into " & sInsumoCombo & " select tcodigoinsumo,ncantidad from " & sDetalle & " inner join tproducto on " & sDetalle & ".tcodigoproducto=tproducto.tcodigoproducto where  tproducto.lControlInsumoCritico=1 and isnull(tproducto.tcodigoinsumo,'')<>''  and isnull(" & sDetalle & ".limprime,0)=1 "
    Cn.Execute "insert into " & sInsumoCombo & " SELECT " & sDetalle & ".nCantidad * " & sComboDetalle & ".nCantidad AS nCantidad , dbo.TPRODUCTO.tCodigoInsumo FROM " & sDetalle & " INNER JOIN " & sComboDetalle & "  on " & sDetalle & ".tItem = " & sComboDetalle & ".tItem INNER JOIN dbo.TPRODUCTO ON " & sComboDetalle & ".tProductoCombo = dbo.TPRODUCTO.tCodigoProducto WHERE     (dbo.TPRODUCTO.lControlInsumoCritico = 1) AND (ISNULL(dbo.TPRODUCTO.tCodigoInsumo, '') <> '') AND (ISNULL(" & sComboDetalle & ".lImprime, 0) = 1)"
    
    Set rstItems = Lib.OpenRecordset("select tCodigoInsumo,  SUM(ncantidad) as ncantidad from " & sInsumoCombo & "  group by tCodigoInsumo order by 2", Cn)
    If Not (rstItems.EOF Or rstItems.BOF) Then
        rstItems.MoveFirst
        For j = 0 To rstItems.RecordCount - 1
            modificaStockInsumo rstItems.Fields(0), rstItems.Fields(1), "I"
            rstItems.MoveNext
        Next j
    End If
   'INSUMOCRITIC

   Cn.Execute "delete from " & sDetalle
   Cn.Execute "delete from " & sComboDetalle
   Cn.Execute "delete from " & sProductoPropiedad
   Cn.Execute "delete from " & sComboPropiedad
   Cn.Execute "Update MPEDIDO set tEstadoPedido ='03', tMotivoAnulacion='" & sCodigo & "', tUsuarioAnulado='" & sUsuarioAutoriza & "', fRegAnulado= getdate(), tTurnoAnulado='" & sTurno & "', tObservacionAnulado='" & sDescrip & "'  where tCodigoPedido ='" & sPedido & "'"
   Cn.Execute "Update DPEDIDO Set tEstadoItem = 'A' where tCodigoPedido = '" & sPedido & "'"
   Cn.Execute "delete TPRODUCTOPROPIEDAD where tCodigoPedido='" & sPedido & "'"
   Cn.Execute "delete CPEDIDO where tCodigoPedido='" & sPedido & "'"
   Cn.Execute "delete TCOMBOPROPIEDAD where tCodigoPedido='" & sPedido & "'"
                      
   RsDetalle.Requery
   RsCombo.Requery
   RsComboPropiedad.Requery
   RsProductoPropiedad.Requery
 
   nMonto = 0
   Pedido = ""
   sProducto = ""
   wCombo = False
   nCombo = 0
   sitem = ""
   VisualizaMonto
   Screen.MousePointer = vbDefault
End Sub

Public Sub AsignaProductoCombo()
    Dim i As Integer
    RsProductoCombo.Filter = "tCombo = '" & sProducto & "'"
    AsignaComandoColor 52, RsProductoCombo, cmdProductoCombo()
    'AsignaComando 52, RsProductoCombo, cmdProductoCombo()
End Sub

Sub PUltimaComanda()
    Dim RsPuntoVentaU   As ADODB.Recordset
    Isql = "Select tPuntoVenta as Codigo, tDescripcion as Descripcion, nUltimoComanda, tmoneda" & _
           " From tPuntoVenta " & _
           " where tHotel='" & sHotel & "' AND lActivo=1 and lInforest=1 and tPuntoVenta='" & sPuntoVenta & "'"
    Set RsPuntoVentaU = Lib.OpenRecordset(Isql, CnInfhotel)
    UltimaComanda = IIf(IsNull(RsPuntoVentaU!nUltimoComanda), "", RsPuntoVentaU!nUltimoComanda)
    Set RsPuntoVentaU = Nothing
End Sub

Private Sub ListarOperadoresConFiltro(ByVal tProducto As String)
    Isql = "select * from vOperador where lActivo = 1 " & _
           "AND ((select Count(tCodigoPropiedad) " & _
           "FROM dbo.TPROPIEDAD LEFT OUTER JOIN dbo.TOPERADOR ON dbo.TPROPIEDAD.tOperador = dbo.TOPERADOR.tOperador " & _
           "Where TPROPIEDAD.lActivo = 1 And IsNull(TOPERADOR.lStockMenos, 0) <> 1 " & _
           "And TPROPIEDAD.tOperador = vOperador.Codigo and tProducto='" & tProducto & "') > 0 OR lStockMenos > 0 ) " & _
           "order by Codigo"
    Set RsOperador = Lib.OpenRecordset(Isql, Cn)
    AsignaBoton 13, RsOperador, cmdOperador()
    If RsOperador.RecordCount > 0 Then
        RsOperador.MoveFirst
        Dim i As Integer
        
        For i = 1 To RsOperador.RecordCount
            If RsOperador!nBoton <> 0 Then
                cmdOperador(RsOperador!nBoton).backColor = vbButtonFace
            End If
            RsOperador.MoveNext
        Next i
        RsOperador.MoveFirst
        xOperador = RsOperador!codigo

        If RsOperador!nBoton <> 0 Then
            cmdOperador(RsOperador!nBoton).backColor = vbRed
        End If
    End If
End Sub

Private Function ObligaPropiedad(ByVal tProducto As String) As Boolean
    Dim j As Integer
    Dim i As Integer
    Dim RstProductoPropiedad As Recordset
    Dim flag As Boolean
    Dim oPos As Integer
    oPos = RsOperador.AbsolutePosition
    Dim mensajeOperador As String
    flag = True 'Si permite salir
    ObligaPropiedad = True
    If RsOperador.RecordCount > 0 Then
        If RsOperador.EOF Then
            RsOperador.MoveFirst
            For i = 1 To 13
                cmdOperador(i).backColor = vbButtonFace
            Next i
        End If
        RsOperador.MoveFirst
        While RsOperador.EOF = False
            If RsOperador!lObligaPropiedad = True Then
                If wAgregaCombo Then
                    Isql = "Select * From " & sComboPropiedad & " Where tCodigoPropiedad In (Select tCodigoPropiedad from TPropiedad Where tOperador = '" & RsOperador!codigo & "' And tProducto = '" & tProducto & "' ) And tItem = '" & sitem & "' And tItemCombo = '" & xItem & "'"
                Else
                    Isql = "Select * From " & sProductoPropiedad & " Where tCodigoPropiedad In (Select tCodigoPropiedad from TPropiedad Where tOperador = '" & RsOperador!codigo & "' And tProducto = '" & tProducto & "' ) And tItem = '" & sitem & "' "
                End If
                Set RstProductoPropiedad = Lib.OpenRecordset(Isql, Cn)
                If RstProductoPropiedad.RecordCount = 0 Then
                    flag = False 'Esta Obligado y no ha elegido Propiedad
                    mensajeOperador = mensajeOperador + "(" + RsOperador!Descripcion + ")"
                End If
            End If
            RsOperador.MoveNext
        Wend
        If flag = False Then
            MsgBox "Propiedades obligadas " & mensajeOperador, vbExclamation, sMensaje
        End If
        ObligaPropiedad = flag
        RsOperador.AbsolutePosition = oPos
        RsOperador.Find "nboton = " & Trim(str(RsOperador!nBoton))
        nOperadorPropiedad = RsOperador!nControl
        For i = 1 To 13
            cmdOperador(i).backColor = vbButtonFace
        Next i
        cmdOperador(RsOperador!nBoton).backColor = vbRed
        'AsignaPropiedad
    End If
End Function

'OO
Private Function ExistenPropiedadesPendientesEnPedido() As Boolean
    Dim oRsDPedidoNoImp As Recordset 'Lista de Productos no impresos
    Set oRsDPedidoNoImp = Obtener_ProductosNoImpresosPorPedido()
    Dim oi As Integer
    Dim oj As Integer
    Dim oflag As Boolean
    Dim oMensaje As String
    oMensaje = "Item(s) con obligatoriedad de propiedad: "
    oflag = True
    'Para cada`producto de DPedido cual lImprime = '0'
    For oi = 1 To oRsDPedidoNoImp.RecordCount
    'Obtener Operadores Obligatorios Filtrados
        Dim oRsOperadoresObligados As Recordset ' Lista de Operadores obligados de un producto
        Set oRsOperadoresObligados = Obtener_OperadoresObligatoriosPorProducto(oRsDPedidoNoImp!tCodigoProducto)
        'Para cada operador Obtener la lista de propiedades
        For oj = 1 To oRsOperadoresObligados.RecordCount
            Dim oRsPropiedadesDeOperador As Recordset ' Lista de Propiedades por Operador
            Set oRsPropiedadesDeOperador = Obtener_PropiedadesSeleccionadasPorProducto(oRsDPedidoNoImp!tItem, oRsDPedidoNoImp!tCodigoProducto, oRsOperadoresObligados!codigo)
            'Verificar si en TProductoPropiedad existe para el item y tCodigoPropiedad IN (lista de propiedades)
            If oRsPropiedadesDeOperador.RecordCount = 0 Then
                oflag = False
                oMensaje = oMensaje + "(" + oRsDPedidoNoImp!tDetallado + ")"
                oj = oRsOperadoresObligados.RecordCount
            End If
            oRsOperadoresObligados.MoveNext
        Next oj
        oRsDPedidoNoImp.MoveNext
    Next oi
    If oflag = False Then
        MsgBox (oMensaje)
    End If
    ExistenPropiedadesPendientesEnPedido = oflag
End Function

'OO
Private Function Obtener_ProductosNoImpresosPorPedido() As Recordset
    If wAgregaCombo Then
        Isql = "Select [" & sDetalle & "].*,TProducto.tDetallado from [" & sDetalle & "] Left Join TProducto On [" & sDetalle & "].tProductoCombo = TProducto.tCodigoProducto " & _
               "Where [" & sDetalle & "].lImprime = '0' And tItem = '" & sitem & "' "
    Else
        Isql = "Select [" & sDetalle & "].*,TProducto.tDetallado from [" & sDetalle & "] Left Join TProducto On [" & sDetalle & "].tCodigoProducto = TProducto.tCodigoProducto Where [" & sDetalle & "].lImprime = '0'"
    End If
    Set Obtener_ProductosNoImpresosPorPedido = Lib.OpenRecordset(Isql, Cn)
End Function

'OO
Private Function Obtener_OperadoresObligatoriosPorProducto(ByVal tProducto) As Recordset
    Isql = "select * from vOperador where lActivo = 1 AND lObligaPropiedad = 1" & _
       "AND ((select Count(tCodigoPropiedad) " & _
       "FROM dbo.TPROPIEDAD LEFT OUTER JOIN dbo.TOPERADOR ON dbo.TPROPIEDAD.tOperador = dbo.TOPERADOR.tOperador " & _
       "Where TPROPIEDAD.lActivo = 1 And IsNull(TOPERADOR.lStockMenos, 0) <> 1 " & _
       "And TPROPIEDAD.tOperador = vOperador.Codigo and tProducto='" & tProducto & "') > 0 OR lStockMenos > 0 ) " & _
       "order by Codigo"
    Set Obtener_OperadoresObligatoriosPorProducto = Lib.OpenRecordset(Isql, Cn)
End Function

'OO
Private Function Obtener_PropiedadesSeleccionadasPorProducto(ByVal tItem, ByVal tProducto, ByVal tOperador) As Recordset
    If wAgregaCombo Then
        Isql = "Select * From TComboPropiedad " & _
            "Where tItem = '" & tItem & "' And tCodigoPropiedad In ( " & _
            "Select tCodigoPropiedad from TPropiedad " & _
            "Where tProducto = '" & tProducto & "' And tOperador = '" & tOperador & "') "
    Else
        Isql = "Select * From " & sProductoPropiedad & " " & _
            "Where tItem = '" & tItem & "' And tCodigoPropiedad In ( " & _
            "Select tCodigoPropiedad from TPropiedad " & _
            "Where tProducto = '" & tProducto & "' And tOperador = '" & tOperador & "') "
    End If
    Set Obtener_PropiedadesSeleccionadasPorProducto = Lib.OpenRecordset(Isql, Cn)
End Function

'OO
Private Function ExistenPropiedadesPendientesEnCombos() As Boolean
    Dim oRsCombosPedido As Recordset
    Dim oi, oj, ok As Integer
    Dim oMensaje As String
    oMensaje = "Combo(s) con productos con propiedades obligatorias: "
    Set oRsCombosPedido = Obtener_TodosLosCombosDelPedido()
    Dim oflag As Boolean
    oflag = True
    For oi = 1 To oRsCombosPedido.RecordCount
        Dim oRsProductoNoImpCombo As Recordset
        Set oRsProductoNoImpCombo = Obtener_LosProductosNoImpDelCombo(oRsCombosPedido!tItem)
        For oj = 1 To oRsProductoNoImpCombo.RecordCount
            Dim oRsOperadores As Recordset
            Set oRsOperadores = Obtener_OperadoresObligatoriosDeUnProducto(oRsProductoNoImpCombo!tProductoCombo)
            For ok = 1 To oRsOperadores.RecordCount
                Dim oRsPropiedadesDeOperador As Recordset
                Set oRsPropiedadesDeOperador = Obtener_PropiedadesSeleccionadasPorProductoDeCombo(oRsCombosPedido!tItem, oRsProductoNoImpCombo!tItemCombo, oRsProductoNoImpCombo!tProductoCombo, oRsOperadores!codigo)
                If oRsPropiedadesDeOperador.RecordCount = 0 Then
                    oflag = False
                    oMensaje = oMensaje + "(" + oRsCombosPedido!tDetallado + ")"
                    oj = oRsProductoNoImpCombo.RecordCount
                    ok = oRsOperadores.RecordCount
                End If
                oRsOperadores.MoveNext
            Next ok
            oRsProductoNoImpCombo.MoveNext
        Next oj
        oRsCombosPedido.MoveNext
    Next oi
    
    If (oflag = False) Then
        MsgBox (oMensaje)
    End If
    ExistenPropiedadesPendientesEnCombos = oflag
End Function

'OO
Private Function Obtener_TodosLosCombosDelPedido() As Recordset
    Isql = "Select  [" & sDetalle & "].*,TProducto.tDetallado From [" & sDetalle & "] Left Join TProducto On [" & sDetalle & "].tCodigoProducto = TProducto.tCodigoProducto " & _
           "Where   [" & sDetalle & "].lCombinacion = '1'"
    Set Obtener_TodosLosCombosDelPedido = Lib.OpenRecordset(Isql, Cn)
End Function

'OO
Private Function Obtener_LosProductosNoImpDelCombo(ByVal tItem As String) As Recordset
    Isql = "Select * from [" & sComboDetalle & "] Where tItem = '" & tItem & "' And lImprime = '0'"
    Set Obtener_LosProductosNoImpDelCombo = Lib.OpenRecordset(Isql, Cn)
End Function

'OO
Private Function Obtener_PropiedadesSeleccionadasPorProductoDeCombo(ByVal tItem As String, ByVal tItemCombo As String, ByVal tProducto As String, ByVal tOperador As String) As Recordset
    Isql = "Select * From " & sComboPropiedad & " " & _
           "Where tItem = '" & tItem & "' And tItemCombo = '" & tItemCombo & "' And tCodigoPropiedad  In ( " & _
           "Select tCodigoPropiedad from TPropiedad Where tProducto = '" & tProducto & "' And tOperador = '" & tOperador & "') "
    Set Obtener_PropiedadesSeleccionadasPorProductoDeCombo = Lib.OpenRecordset(Isql, Cn)
End Function

'OO
Private Function Obtener_OperadoresObligatoriosDeUnProducto(ByVal tProducto As String) As Recordset
    Isql = "select * from vOperador where lActivo = 1 AND lObligaPropiedad = 1 AND ( " & _
           "(select Count(tCodigoPropiedad) FROM dbo.TPROPIEDAD LEFT OUTER JOIN dbo.TOPERADOR ON dbo.TPROPIEDAD.tOperador = dbo.TOPERADOR.tOperador " & _
           "Where TPROPIEDAD.lActivo = 1 And IsNull(tOperador.lStockMenos, 0) <> 1 And TPROPIEDAD.tOperador = vOperador.Codigo " & _
           "And tProducto='" & tProducto & "') > 0 OR lStockMenos > 0 ) order by Codigo"
    Set Obtener_OperadoresObligatoriosDeUnProducto = Lib.OpenRecordset(Isql, Cn)
End Function

'OO
Private Function Obtener_ProductoDeCombo(ByVal tCombo As String, ByVal tCodigoProducto As String) As Recordset
    Isql = "Select lUnico,tEtiqueta From TCOMBO Where tCombo = '" & tCombo & "' And tCodigoProducto = '" & tCodigoProducto & "'"
    Set Obtener_ProductoDeCombo = Lib.OpenRecordset(Isql, Cn)
End Function

'OO
Private Function ObtenerSumaCantidadesEnElCombo(ByVal tItem As String, ByVal tEtiqueta As String) As Double
    Isql = "Select ISNULL(Sum(nCantidad),0) as nCantidad from [" & sComboDetalle & "] " & _
           "Where   tItem = '" & tItem & "' And tProductoCombo IN ( " & _
           "Select tCodigoProducto From TCOMBO Where tEtiqueta = '" & tEtiqueta & "' And lUnico = '1') "
    Dim Suma As Double
    Dim oRsResultado As Recordset
    Set oRsResultado = Lib.OpenRecordset(Isql, Cn)
    Suma = oRsResultado!nCantidad
    ObtenerSumaCantidadesEnElCombo = Suma
End Function

'OO
Private Function ObtenerDetalleProducto(ByVal tItem As String) As Recordset
    'Isql = "Select D.* ,P.nPrecioVenta as 'PrecioProducto' from [" & sDetalle & "] As D left Join TProducto As P On D.tCodigoProducto = P.tCodigoProducto where D.tItem ='" & tItem & "'"
    'Isql = "Select D.* ,CASE (D.tTipoPedido)  WHEN '01' THEN P.nprecioventa when '02' then p.npreciodelivery when '03' then p.npreciollevar when '04' then p.npreciocanal4 when '05' then p.npreciocanal5 END as 'PrecioProducto' from [" & sDetalle & "] As D left Join TProducto As P On D.tCodigoProducto = P.tCodigoProducto where D.tItem ='" & tItem & "'"
        Isql = " Select D.* , " & _
           " case when ( " & _
           " CASE (D.tTipoPedido)  WHEN '01' THEN P.nprecioventa " & _
           " when '02' then p.npreciodelivery " & _
           "            when '03' then p.npreciollevar " & _
           "            when '04' then p.npreciocanal4 " & _
           "            when '05' then p.npreciocanal5 " & _
           " END)=0 then p.nprecioventa else " & _
           " ( CASE (D.tTipoPedido)  WHEN '01' THEN P.nprecioventa " & _
           "            when '02' then p.npreciodelivery " & _
           "            when '03' then p.npreciollevar " & _
           "            when '04' then p.npreciocanal4 " & _
           "            when '05' then p.npreciocanal5 End) end as 'PrecioProducto',P.tMoneda as 'tMonedaProducto' from [" & sDetalle & "] As D left Join TProducto As P On D.tCodigoProducto = P.tCodigoProducto where D.tItem ='" & tItem & "'"
'    Isql = " Select D.* ,d.nprecioventa   as 'PrecioProducto',P.tMoneda as 'tMonedaProducto' from [" & sDetalle & "] As D left Join TProducto As P On D.tCodigoProducto = P.tCodigoProducto where D.tItem ='" & tItem & "'"

    Set ObtenerDetalleProducto = Lib.OpenRecordset(Isql, Cn)
End Function
'OO
Private Function Obtener_CantidadMaximaDeUnicoEtiqueta(ByVal tItem As String, cantidadActual As Double) As Double
    
    Isql = "Select Sum(P.nCantidad) as Cantidad " & _
           "from [" & sComboDetalle & "] as P Left Join TCOMBO as C ON P.tProducto = C.tCombo AND   P.tProductoCombo = C.tCodigoProducto " & _
           "where P.tItem ='" & tItem & "' And C.lUnico = '1' " & _
           "Group By C.tEtiqueta,c.tcodigoproducto order by 1 desc"
    Dim oRsResultado As Recordset
    Dim oi As Integer
    Dim CantMax As Double
    CantMax = 0
    CantMax = cantidadActual
    Set oRsResultado = Lib.OpenRecordset(Isql, Cn)
    If oRsResultado.RecordCount > 0 Then
        For oi = 1 To oRsResultado.RecordCount
            If oRsResultado!Cantidad <> CantMax Then
                CantMax = oRsResultado!Cantidad
            End If
            oRsResultado.MoveNext
        Next oi
        Obtener_CantidadMaximaDeUnicoEtiqueta = CantMax
    Else
        Obtener_CantidadMaximaDeUnicoEtiqueta = 0
    End If
End Function

Private Function ObtenerCodigoMozo(ByVal tResumido As String) As String
    Isql = "select * from vMozo where substring(Codigo,1,1)<>'*' AND lActivo = 1 Order by nBoton"
    Set RsMozo = Lib.OpenRecordset(Isql, Cn)
    RsMozo.Filter = "tResumido = '" & tResumido & "'"
    If RsMozo.RecordCount = 0 Then
       txtTitulo.Caption = " Caja Rápida : " & sCaja & " Mesero: Sin Mesero"
       ObtenerCodigoMozo = "0000"
    Else
       txtTitulo.Caption = " Caja Rápida : " & sCaja & " Mesero: " & tResumido
       ObtenerCodigoMozo = RsMozo!codigo
    End If
End Function

Private Function CalculaDescuento() As Boolean
    Dim sCriterio As String
    Dim lAcumulable As Boolean
    Dim nOferta As Double
    Dim nSuma As Double
    
    nSuma = Calcular("SELECT sum(nPrecioOficial*nCantidad) as Codigo FROM " & sDetalle & " LEFT OUTER JOIN dbo.TPRODUCTO ON " & sDetalle & ".tCodigoProducto = dbo.TPRODUCTO.tCodigoProducto where lDescuento=1", Cn)

    If RsDetalle.RecordCount <> 0 Then
       RsDetalle.MoveFirst
       
       Do While Not RsDetalle.EOF
          'Busca Oferta
          nPVenta = 0
          sCriterio = "tCodigoProducto ='" & RsDetalle!tCodigoProducto & "' and lActivo=1"
          sCriterio = sCriterio & " and (tFrecuencia='00' or tFrecuencia='0" & Weekday(FechaServidor(), vbMonday) & "' or (tFrecuencia='99' and fFecha='" & Format(FechaServidor(), "yyyy/MM/dd 00:00") & "') and tHoraInicial<='" & Format(Time, "HH:mm") & "' and tHoraFinal>='" & Format(Time, "HH:mm") & "')"
          sCriterio = sCriterio & " and (lPermanente=1 or (lPermanente=0 and fFechaInicial<='" & Format(FechaServidor(), "yyyy/mm/dd") & "' and fFechaFinal>='" & Format(FechaServidor(), "yyyy/mm/dd") & "'))"
            
          Isql = "select * from TOFERTA where " & sCriterio
          Set RsOferta = Lib.OpenRecordset(Isql, Cn)
          
          lAcumulable = True
          nOferta = 0
          Acumulado = 0
       
          If RsOferta.RecordCount > 0 Then
             RsOferta.MoveFirst
             lAcumulable = RsOferta!lAcumulable
             nOferta = RsDetalle!nPrecioOficial * IIf(IsNull(RsOferta!nRatio), 1, RsOferta!nRatio) / 100
          End If
          
         If RsDetalle!lDescuento And lAcumulable = True Then
            If Calcular("select lRatio as Codigo FROM vMotivoDescuento where Codigo='" & sCodigoDescuento & "'", Cn) Then
               nPVenta = (RsDetalle!nPrecioOficial - nOferta) - ((RsDetalle!nPrecioOficial - nOferta) * xDescuento / 100)
            Else
               Dim xPorc As Double
               xPorc = (RsDetalle!nPrecioOficial - nOferta) * RsDetalle!nCantidad * 100 / nSuma
               nPVenta = (RsDetalle!nPrecioOficial - nOferta) - ((xPorc * xDescuento / 100) / RsDetalle!nCantidad)
            End If
            
            Select Case pais ' ok
                Case "001" 'Bolivia
                         Acumulado = IIf(RsDetalle!nprecioImpuesto1 <> 0, Acumulado + nPorcentaje1, Acumulado)
                         Acumulado = IIf(RsDetalle!nprecioImpuesto2 <> 0, Acumulado + nPorcentaje2, Acumulado)
                         Acumulado = IIf(RsDetalle!nprecioImpuesto3 <> 0, Acumulado + nPorcentaje3, Acumulado)
                         Acumulado = (Acumulado / 100)
                        
                         nImpuesto1 = IIf(RsDetalle!nprecioImpuesto1 <> 0, nPVenta * nPorcentaje1 / 100, 0)
                         nImpuesto2 = IIf(RsDetalle!nprecioImpuesto2 <> 0, nPVenta * nPorcentaje2 / 100, 0)
                         nImpuesto3 = IIf(RsDetalle!nprecioImpuesto3 <> 0, nPVenta * nPorcentaje3 / 100, 0)
                         nPBase = nPVenta - nImpuesto1 - nImpuesto2 - nImpuesto3
                         
                Case Else 'Peru, Ecuador
                         Acumulado = IIf(RsDetalle!nprecioImpuesto1 <> 0, Acumulado + nPorcentaje1, Acumulado)
                         Acumulado = IIf(RsDetalle!nprecioImpuesto2 <> 0, Acumulado + nPorcentaje2, Acumulado)
                         Acumulado = IIf(RsDetalle!nprecioImpuesto3 <> 0, Acumulado + nPorcentaje3, Acumulado)
                         Acumulado = 1 + (Acumulado / 100)
                        
                         nImpuesto1 = IIf(RsDetalle!nprecioImpuesto1 <> 0, nPVenta / Acumulado * nPorcentaje1 / 100, 0)
                         nImpuesto2 = IIf(RsDetalle!nprecioImpuesto2 <> 0, nPVenta / Acumulado * nPorcentaje2 / 100, 0)
                         nImpuesto3 = IIf(RsDetalle!nprecioImpuesto3 <> 0, nPVenta / Acumulado * nPorcentaje3 / 100, 0)
                         nPBase = nPVenta - nImpuesto1 - nImpuesto2 - nImpuesto3
                         
            End Select
             Isql = "Update " & sDetalle & " Set nPrecioNeto = " & nPBase & ", " & _
                    "nDescuento = " & RsDetalle!nPrecioOficial - nPVenta & ", " & _
                    "nRecargo = " & nRecargo & ", " & _
                    "nPrecioOficial = " & RsDetalle!nPrecioOficial & ", " & _
                    "nprecioImpuesto1 = " & nImpuesto1 & ", " & _
                    "nprecioImpuesto2 = " & nImpuesto2 & ", " & _
                    "nprecioImpuesto3 = " & nImpuesto3 & ", " & _
                    "nPrecioVenta = " & nPVenta & ", " & _
                    "nventa = " & nPVenta * RsDetalle!nCantidad & ", " & _
                    "nCantidad = " & RsDetalle!nCantidad & ", " & _
                    "nImpuesto1 = " & nImpuesto1 * RsDetalle!nCantidad & ", " & _
                    "nImpuesto2 = " & nImpuesto2 * RsDetalle!nCantidad & ", " & _
                    "nImpuesto3 = " & nImpuesto3 * RsDetalle!nCantidad & ", " & _
                    "tCortesia = '" & sCortesia & "' " & _
                    "where tItem = '" & RsDetalle!tItem & "' "
                    Cn.Execute Isql
          End If
       RsDetalle.MoveNext
       Loop
    End If

End Function

Public Function verificaCantidadDeItemsCombos(ByVal tItem As String, ByVal numeroDeItemCombos As Double, ByVal ncantidadNueva As Double) As Boolean
    Dim cantidadMaximoNueva As Double
    Dim cantidadMaximoPosible As Double
    Dim X As Integer
    Dim oRstRecorriendoCombo As New Recordset
    Dim oRstCantidadNUnicos As New Recordset
    
    verificaCantidadDeItemsCombos = False
    cantidadMaximoNueva = numeroDeItemCombos * ncantidadNueva
    cantidadMaximoPosible = 0
    Set oRstRecorriendoCombo = Lib.OpenRecordset("select tcombo.ncantidad ,cpedido.tproductocombo from [" & sComboDetalle & "]  AS cpedido inner join tcombo on cpedido.tproducto=tcombo.tcombo and cpedido.tproductocombo=tcombo.tcodigoproducto where  cpedido.titem='" & tItem & "' and lfijo=1", Cn)
    If Not (oRstRecorriendoCombo.EOF Or oRstRecorriendoCombo.BOF) Then
        oRstRecorriendoCombo.MoveFirst
        For X = 0 To oRstRecorriendoCombo.RecordCount - 1
            cantidadMaximoPosible = cantidadMaximoPosible + (oRstRecorriendoCombo!nCantidad * ncantidadNueva)
            oRstRecorriendoCombo.MoveNext
        Next X
    End If
    
    Set oRstCantidadNUnicos = Lib.OpenRecordset("select isnull(sum(cpedido.ncantidad),0) from [" & sComboDetalle & "]  AS cpedido inner join tcombo on cpedido.tproducto=tcombo.tcombo and cpedido.tproductocombo=tcombo.tcodigoproducto where   titem='" & tItem & "' and tcombo.lfijo=0 ", Cn)
    If Not (oRstCantidadNUnicos.EOF Or oRstCantidadNUnicos.BOF) Then
        cantidadMaximoPosible = cantidadMaximoPosible + oRstCantidadNUnicos.Fields(0)
    End If
    
    If cantidadMaximoNueva >= cantidadMaximoPosible Then
        verificaCantidadDeItemsCombos = True
    End If
    
End Function


'insumo critico23
Public Function obtieneProductos(tcodigoinsumo As String) As String
 obtieneProductos = ""
 Dim rstProductos As New ADODB.Recordset
 Dim k As Integer
 Set rstProductos = Lib.OpenRecordset("select isnull(tdetallado,'') as producto FROM         dbo.TPRODUCTO INNER JOIN  " & sDetalle & " ON dbo.TPRODUCTO.tCodigoProducto = " & sDetalle & ".tCodigoProducto INNER JOIN  dbo.TINSUMO ON dbo.TPRODUCTO.tcodigoInsumo = dbo.TINSUMO.tcodigo where   tproducto.tcodigoinsumo='" & tcodigoinsumo & "' and lcontrolinsumocritico=1 and isnull(limprime,0)=0 AND   (tinsumo.lactivo=1) group by isnull(tdetallado,'') ", Cn)
 If Not (rstProductos.EOF Or rstProductos.BOF) Then
        rstProductos.MoveFirst
        For k = 0 To rstProductos.RecordCount - 1
                    obtieneProductos = IIf(Len(obtieneProductos) = 0, rstProductos.Fields(0), obtieneProductos & " / " & rstProductos.Fields(0))
            rstProductos.MoveNext
        Next k
 End If
 
 
End Function
'insumo critico23
 
'luchiinsumo
Public Sub verificatitulo()
        'INSUMOCRITICO23
        Dim rsInsumo As New ADODB.Recordset
        If Calcular("select isnull(lControlInsumoCritico,0) as codigo from tproducto  INNER JOIN " & sDetalle & "  on tproducto.tcodigoproducto=" & sDetalle & ".tcodigoproducto where  titem='" & sitem & "'", Cn) = True Then
                        Set rsInsumo = Lib.OpenRecordset("select isnull(tcodigoinsumo,'') tcodigoinsumo , isnull(tinsumo.descripcion,'') ,isnull(nstock,0) , " & sDetalle & ".ncantidad from tproducto inner join tinsumo on tproducto.tcodigoinsumo =tinsumo.tcodigo inner join " & sDetalle & "  on tproducto.tcodigoproducto=" & sDetalle & ".tcodigoproducto  where     titem='" & sitem & "' and tinsumo.lactivo=1", Cn)
                        If Not (rsInsumo.EOF Or rsInsumo.BOF) Then
                                Label2.Caption = "   Insumo Crítico ->   " & rsInsumo.Fields(1) & " =  Stock: " & str(rsInsumo.Fields(2)) & "      Solicitado: " + str(rsInsumo.Fields(3))
                        End If
                Else
                        Label2.Caption = muestra
        End If
        'INSUMOCRITICO
    If lvisor Then
        Call InsertVisor8
    End If
End Sub
'luchoinsumo

'diaContable
Public Function obtieneDiaContable() As Date
   Dim oComando As New clsComando
   Dim DiaContable As Date
   Dim rst1 As New ADODB.Recordset
   Set oComando = New clsComando
                  If Not oComando.CreateCmdSp("usp_GenObtieneDiaContable", Cn) Then
                     Set oComando = Nothing
                     Exit Function
                  End If
                  oComando.CreateParameter "@lDiaContable", adBoolean, adParamInput, 1, lDiaContable
                  oComando.CreateParameter "@sHoraCierre", adVarChar, adParamInput, 5, tHoraCierreDiaContable
                  oComando.CreateParameter "@tUsuario", adVarChar, adParamInput, 15, sUsuario
                 oComando.CreateParameter "@fDiaContable", adDBDate, adParamOutput, 10, DiaContable
                If Not oComando.GetParamOK Then
                   Set oComando = Nothing
                   Exit Function
                End If
                    Set rst1 = oComando.GetSP()
                obtieneDiaContable = oComando.GetParameterValue("@fDiaContable")
                fImpresionDiaContable = obtieneDiaContable
End Function
'diaContable
Private Sub CalculaAplicaTope(nTope As Double)
    Dim sCriterio As String
    Dim lAcumulable As Boolean
    Dim nOferta As Double
    Dim nSuma As Double
    
    nSuma = Calcular("SELECT sum(nPrecioOficial*nCantidad) as Codigo FROM " & sDetalle & " LEFT OUTER JOIN dbo.TPRODUCTO ON " & sDetalle & ".tCodigoProducto = dbo.TPRODUCTO.tCodigoProducto where lDescuento=1", Cn)

If RsDetalle.RecordCount <> 0 Then
   RsDetalle.MoveFirst
   
   Do While Not RsDetalle.EOF
      'Busca Oferta
      nPVenta = 0
      sCriterio = "tCodigoProducto ='" & RsDetalle!tCodigoProducto & "' and lActivo=1"
      sCriterio = sCriterio & " and (tFrecuencia='00' or tFrecuencia='0" & Weekday(FechaServidor(), vbMonday) & "' or (tFrecuencia='99' and fFecha='" & Format(FechaServidor(), "yyyy/MM/dd 00:00") & "') and tHoraInicial<='" & Format(Time, "HH:mm") & "' and tHoraFinal>='" & Format(Time, "HH:mm") & "')"
      sCriterio = sCriterio & " and (lPermanente=1 or (lPermanente=0 and fFechaInicial<='" & Format(FechaServidor(), "yyyy/mm/dd") & "' and fFechaFinal>='" & Format(FechaServidor(), "yyyy/mm/dd") & "'))"
        
      Isql = "select * from TOFERTA where " & sCriterio
      Set RsOferta = Lib.OpenRecordset(Isql, Cn)
      
      lAcumulable = True
      nOferta = 0
      Acumulado = 0
      
      If RsOferta.RecordCount > 0 Then
         RsOferta.MoveFirst
         lAcumulable = RsOferta!lAcumulable
         nOferta = RsDetalle!nPrecioOficial * IIf(IsNull(RsOferta!nRatio), 1, RsOferta!nRatio) / 100
      End If
      
      If RsDetalle!lDescuento And lAcumulable = True Then
         
         xDescuento = (RsDetalle!nPrecioOficial - nOferta) * (RsDetalle!nCantidad * 100 / nSuma)
         nPVenta = (RsDetalle!nPrecioOficial - nOferta) - ((nTope * xDescuento / 100) / RsDetalle!nCantidad)
         
          Select Case pais ' ok
            Case "001" 'Bolivia
                         Acumulado = IIf(RsDetalle!nprecioImpuesto1 <> 0, Acumulado + nPorcentaje1, Acumulado)
                         Acumulado = IIf(RsDetalle!nprecioImpuesto2 <> 0, Acumulado + nPorcentaje2, Acumulado)
                         Acumulado = IIf(RsDetalle!nprecioImpuesto3 <> 0, Acumulado + nPorcentaje3, Acumulado)
                         Acumulado = (Acumulado / 100)
                        
                         nImpuesto1 = IIf(RsDetalle!nprecioImpuesto1 <> 0, nPVenta * nPorcentaje1 / 100, 0)
                         nImpuesto2 = IIf(RsDetalle!nprecioImpuesto2 <> 0, nPVenta * nPorcentaje2 / 100, 0)
                         nImpuesto3 = IIf(RsDetalle!nprecioImpuesto3 <> 0, nPVenta * nPorcentaje3 / 100, 0)
                         nPBase = nPVenta - nImpuesto1 - nImpuesto2 - nImpuesto3
            Case Else 'Peru, Ecuador
                         Acumulado = IIf(RsDetalle!nprecioImpuesto1 <> 0, Acumulado + nPorcentaje1, Acumulado)
                         Acumulado = IIf(RsDetalle!nprecioImpuesto2 <> 0, Acumulado + nPorcentaje2, Acumulado)
                         Acumulado = IIf(RsDetalle!nprecioImpuesto3 <> 0, Acumulado + nPorcentaje3, Acumulado)
                         Acumulado = 1 + (Acumulado / 100)
                        
                         nImpuesto1 = IIf(RsDetalle!nprecioImpuesto1 <> 0, nPVenta / Acumulado * nPorcentaje1 / 100, 0)
                         nImpuesto2 = IIf(RsDetalle!nprecioImpuesto2 <> 0, nPVenta / Acumulado * nPorcentaje2 / 100, 0)
                         nImpuesto3 = IIf(RsDetalle!nprecioImpuesto3 <> 0, nPVenta / Acumulado * nPorcentaje3 / 100, 0)
                         nPBase = nPVenta - nImpuesto1 - nImpuesto2 - nImpuesto3
          
          End Select
         Isql = "Update " & sDetalle & " Set nPrecioNeto = " & nPBase & ", " & _
                "nDescuento = " & RsDetalle!nPrecioOficial - nPVenta & ", " & _
                "nRecargo = " & nRecargo & ", " & _
                "nPrecioOficial = " & RsDetalle!nPrecioOficial & ", " & _
                "nprecioImpuesto1 = " & nImpuesto1 & ", " & _
                "nprecioImpuesto2 = " & nImpuesto2 & ", " & _
                "nprecioImpuesto3 = " & nImpuesto3 & ", " & _
                "nPrecioVenta = " & nPVenta & ", " & _
                "nventa = " & nPVenta * RsDetalle!nCantidad & ", " & _
                "nCantidad = " & RsDetalle!nCantidad & ", " & _
                "nImpuesto1 = " & nImpuesto1 * RsDetalle!nCantidad & ", " & _
                "nImpuesto2 = " & nImpuesto2 * RsDetalle!nCantidad & ", " & _
                "nImpuesto3 = " & nImpuesto3 * RsDetalle!nCantidad & ", " & _
                "tCortesia = '" & sCortesia & "' " & _
                "where tItem = '" & RsDetalle!tItem & "'"
                Cn.Execute Isql
      End If
   RsDetalle.MoveNext
   Loop
End If

End Sub
Private Sub InsertVisor8()
On Error GoTo fin
    If lvisor And sCaja <> "" Then
        Cn.Execute "delete from infovisor where tcaja='" & sCaja & "'"
        Cn.Execute "insert into infovisor(id,tcaja,Pedido,estado) values(1,'" & sCaja & "','" & sCaja & "',1)"
        Cn.Execute "delete from Visor_Dpedido where tcajad='" & sCaja & "'"
         Cn.Execute "insert into visor_dpedido " & _
           "(tCodigoPedido, tTipoPedido, tItem, tCodigoProducto, tCodigoGrupo, tCodigoSubGrupo, " & _
           "nPrecioNeto, nRecargo, nDescuento, nPrecioOficial, nPrecioImpuesto1, nPrecioImpuesto2, nPrecioImpuesto3, nPrecioVenta, " & _
           "nCantidad, nVenta, nImpuesto1, nImpuesto2, nImpuesto3, " & _
           "lImprime, tArea, lImprimeArea, lCombinacion, nCombinacion, nInsumo, nGasto, nManoObra, nOrden, tEstadoItem,tsubalmacen,toferta,tCajaD) " & _
            " Select '" & sCaja & "'+tItem, tTipoPedido, tItem, tCodigoProducto, tCodigoGrupo, tCodigoSubGrupo, " & _
           "nPrecioNeto, nRecargo, nDescuento, nPrecioOficial, nPrecioImpuesto1, nPrecioImpuesto2, nPrecioImpuesto3, nPrecioVenta, " & _
           "nCantidad, nVenta, nImpuesto1, nImpuesto2, nImpuesto3, " & _
           "lImprime, tArea, lImprimeArea, lCombinacion, nCombinacion, nInsumo, nGasto, nManoObra, nOrden, tEstadoItem,tsubalmacen,toferta,'" & sCaja & "' " & _
           " From [" & sDetalle & "]"
    End If
fin:
    
End Sub








