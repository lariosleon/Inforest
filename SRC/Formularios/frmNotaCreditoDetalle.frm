VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmNotaCreditoDetalle 
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   5520
   ClientLeft      =   2520
   ClientTop       =   2640
   ClientWidth     =   11025
   Icon            =   "frmNotaCreditoDetalle.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "frmNotaCreditoDetalle.frx":030A
   ScaleHeight     =   5520
   ScaleWidth      =   11025
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
      Left            =   2745
      TabIndex        =   62
      Top             =   2040
      Visible         =   0   'False
      Width           =   6315
      Begin VB.Image imgProceso 
         Height          =   240
         Index           =   0
         Left            =   840
         Picture         =   "frmNotaCreditoDetalle.frx":040C
         Top             =   1155
         Width           =   240
      End
      Begin VB.Label lblPaso2 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H00000000&
         BackStyle       =   0  'Transparent
         Caption         =   "Obteniendo codigo hash almacenado."
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
         TabIndex        =   65
         Top             =   1155
         Visible         =   0   'False
         Width           =   3090
      End
      Begin VB.Label lblPaso1 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H00000000&
         BackStyle       =   0  'Transparent
         Caption         =   "Enviando información de documento a spring."
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
         TabIndex        =   64
         Top             =   870
         Visible         =   0   'False
         Width           =   3780
      End
      Begin VB.Label Label3 
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
         TabIndex        =   63
         Top             =   15
         Width           =   2490
      End
      Begin VB.Image imgProceso 
         Height          =   240
         Index           =   1
         Left            =   840
         Picture         =   "frmNotaCreditoDetalle.frx":061F
         Top             =   855
         Width           =   240
      End
      Begin VB.Image imgProceso 
         Height          =   240
         Index           =   2
         Left            =   840
         Picture         =   "frmNotaCreditoDetalle.frx":0832
         Top             =   855
         Width           =   240
      End
      Begin VB.Image imgProceso 
         Height          =   240
         Index           =   3
         Left            =   840
         Picture         =   "frmNotaCreditoDetalle.frx":0B74
         Top             =   1140
         Width           =   240
      End
      Begin VB.Label Label2 
         BackColor       =   &H00C0FFFF&
         Caption         =   "   Proceso de envio de documento a Spring......."
         BeginProperty Font 
            Name            =   "Arial"
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
         TabIndex        =   66
         Top             =   435
         Width           =   5910
      End
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
      Height          =   4665
      Left            =   2490
      TabIndex        =   26
      Top             =   60
      Width           =   8475
      Begin MSComCtl2.DTPicker dtpFecha 
         Height          =   315
         Left            =   1740
         TabIndex        =   2
         Top             =   615
         Width           =   2295
         _ExtentX        =   4048
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
         Format          =   16908289
         CurrentDate     =   38096
      End
      Begin VB.TextBox txtNC3 
         BackColor       =   &H00E0E0E0&
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
         ForeColor       =   &H000000C0&
         Height          =   285
         Left            =   2940
         Locked          =   -1  'True
         TabIndex        =   60
         TabStop         =   0   'False
         Top             =   300
         Width           =   1095
      End
      Begin VB.TextBox txtNC2 
         BackColor       =   &H00E0E0E0&
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
         ForeColor       =   &H000000C0&
         Height          =   285
         Left            =   2025
         Locked          =   -1  'True
         TabIndex        =   59
         TabStop         =   0   'False
         Top             =   300
         Width           =   870
      End
      Begin VB.CommandButton cmdNotaCredito 
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
         Left            =   4140
         Style           =   1  'Graphical
         TabIndex        =   0
         Top             =   270
         Width           =   1170
      End
      Begin VB.CommandButton cmdNotaCredito 
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
         Left            =   5325
         Style           =   1  'Graphical
         TabIndex        =   1
         Top             =   270
         Width           =   1170
      End
      Begin VB.TextBox txtPrefijo 
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
         ForeColor       =   &H00800000&
         Height          =   285
         Left            =   1740
         MaxLength       =   17
         TabIndex        =   58
         Top             =   960
         Width           =   255
      End
      Begin VB.CommandButton cmdOpcion 
         Caption         =   "Correlativo"
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
         Index           =   6
         Left            =   4274
         Style           =   1  'Graphical
         TabIndex        =   5
         Top             =   1650
         Width           =   1170
      End
      Begin VB.CommandButton cmdOpcion 
         Caption         =   "Tipo de Documento"
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
         Index           =   4
         Left            =   1740
         Style           =   1  'Graphical
         TabIndex        =   3
         Top             =   1650
         Width           =   1170
      End
      Begin VB.TextBox txtCorrela 
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
         ForeColor       =   &H00800000&
         Height          =   285
         Left            =   2940
         MaxLength       =   17
         TabIndex        =   57
         Top             =   960
         Width           =   1095
      End
      Begin VB.CommandButton cmdOpcion 
         Caption         =   "Motivo"
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
         Index           =   8
         Left            =   5541
         Style           =   1  'Graphical
         TabIndex        =   6
         Top             =   1650
         Width           =   1170
      End
      Begin VB.CommandButton cmdOpcion 
         Caption         =   "Monto"
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
         Index           =   9
         Left            =   6810
         Style           =   1  'Graphical
         TabIndex        =   7
         Top             =   1650
         Width           =   1260
      End
      Begin VB.Frame Frame 
         Caption         =   " Resultados "
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2265
         Left            =   210
         TabIndex        =   33
         Top             =   2280
         Width           =   7875
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
            Height          =   345
            Index           =   15
            Left            =   3270
            Picture         =   "frmNotaCreditoDetalle.frx":0EB6
            Style           =   1  'Graphical
            TabIndex        =   10
            Top             =   1455
            Width           =   390
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
            Height          =   345
            Index           =   14
            Left            =   3270
            Picture         =   "frmNotaCreditoDetalle.frx":0FB8
            Style           =   1  'Graphical
            TabIndex        =   9
            Top             =   1110
            Width           =   390
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
            Height          =   345
            Index           =   13
            Left            =   3270
            Picture         =   "frmNotaCreditoDetalle.frx":10BA
            Style           =   1  'Graphical
            TabIndex        =   8
            Top             =   765
            Width           =   390
         End
         Begin VB.TextBox txtResTotal 
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
            Left            =   6060
            Locked          =   -1  'True
            MaxLength       =   15
            TabIndex        =   53
            Text            =   "0.00"
            Top             =   1830
            Width           =   1395
         End
         Begin VB.TextBox txtResImp3 
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
            Left            =   6060
            Locked          =   -1  'True
            MaxLength       =   15
            TabIndex        =   52
            Text            =   "0.00"
            Top             =   1485
            Width           =   1395
         End
         Begin VB.TextBox txtResImp2 
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
            Left            =   6060
            Locked          =   -1  'True
            MaxLength       =   15
            TabIndex        =   51
            Text            =   "0.00"
            Top             =   1140
            Width           =   1395
         End
         Begin VB.TextBox txtResImp1 
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
            Left            =   6060
            Locked          =   -1  'True
            MaxLength       =   15
            TabIndex        =   50
            Text            =   "0.00"
            Top             =   795
            Width           =   1395
         End
         Begin VB.TextBox txtResNeto 
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
            Left            =   6060
            Locked          =   -1  'True
            MaxLength       =   15
            TabIndex        =   49
            Text            =   "0.00"
            Top             =   450
            Width           =   1395
         End
         Begin VB.TextBox txtDocTotal 
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
            ForeColor       =   &H00800000&
            Height          =   285
            Left            =   3930
            Locked          =   -1  'True
            MaxLength       =   15
            TabIndex        =   48
            Text            =   "0.00"
            Top             =   1830
            Width           =   1395
         End
         Begin VB.TextBox txtDocImp3 
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
            ForeColor       =   &H00800000&
            Height          =   285
            Left            =   3930
            Locked          =   -1  'True
            MaxLength       =   15
            TabIndex        =   47
            Text            =   "0.00"
            Top             =   1485
            Width           =   1395
         End
         Begin VB.TextBox txtDocImp2 
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
            ForeColor       =   &H00800000&
            Height          =   285
            Left            =   3930
            Locked          =   -1  'True
            MaxLength       =   15
            TabIndex        =   46
            Text            =   "0.00"
            Top             =   1140
            Width           =   1395
         End
         Begin VB.TextBox txtDocImp1 
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
            ForeColor       =   &H00800000&
            Height          =   285
            Left            =   3930
            Locked          =   -1  'True
            MaxLength       =   15
            TabIndex        =   45
            Text            =   "0.00"
            Top             =   795
            Width           =   1395
         End
         Begin VB.TextBox txtDocNeto 
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
            ForeColor       =   &H00800000&
            Height          =   285
            Left            =   3930
            Locked          =   -1  'True
            MaxLength       =   15
            TabIndex        =   44
            Text            =   "0.00"
            Top             =   450
            Width           =   1395
         End
         Begin VB.TextBox txtNCNeto 
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
            ForeColor       =   &H000000C0&
            Height          =   285
            Left            =   1800
            Locked          =   -1  'True
            MaxLength       =   15
            TabIndex        =   38
            Text            =   "0.00"
            Top             =   450
            Width           =   1395
         End
         Begin VB.TextBox txtNCImp1 
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
            ForeColor       =   &H000000C0&
            Height          =   285
            Left            =   1800
            Locked          =   -1  'True
            MaxLength       =   15
            TabIndex        =   37
            Text            =   "0.00"
            Top             =   795
            Width           =   1395
         End
         Begin VB.TextBox txtNCImp2 
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
            ForeColor       =   &H000000C0&
            Height          =   285
            Left            =   1800
            Locked          =   -1  'True
            MaxLength       =   15
            TabIndex        =   36
            Text            =   "0.00"
            Top             =   1140
            Width           =   1395
         End
         Begin VB.TextBox txtNCImp3 
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
            ForeColor       =   &H000000C0&
            Height          =   285
            Left            =   1800
            Locked          =   -1  'True
            MaxLength       =   15
            TabIndex        =   35
            Text            =   "0.00"
            Top             =   1485
            Width           =   1395
         End
         Begin VB.TextBox txtNCTotal 
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
            ForeColor       =   &H000000C0&
            Height          =   285
            Left            =   1800
            Locked          =   -1  'True
            MaxLength       =   15
            TabIndex        =   34
            Text            =   "0.00"
            Top             =   1830
            Width           =   1395
         End
         Begin VB.Label Label 
            Alignment       =   2  'Center
            AutoSize        =   -1  'True
            Caption         =   "Resultado"
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
            Index           =   11
            Left            =   6315
            TabIndex        =   56
            Top             =   210
            Width           =   870
         End
         Begin VB.Label Label 
            Alignment       =   2  'Center
            AutoSize        =   -1  'True
            Caption         =   "Documento"
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
            Index           =   10
            Left            =   4125
            TabIndex        =   55
            Top             =   210
            Width           =   1005
         End
         Begin VB.Label Label 
            Alignment       =   2  'Center
            AutoSize        =   -1  'True
            Caption         =   "Nota de Crédito"
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
            Left            =   1815
            TabIndex        =   54
            Top             =   210
            Width           =   1350
         End
         Begin VB.Label Label 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "Total :"
            Height          =   195
            Index           =   8
            Left            =   990
            TabIndex        =   43
            Top             =   1875
            Width           =   450
         End
         Begin VB.Label Label 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "Impuesto 3 :"
            Height          =   195
            Index           =   7
            Left            =   555
            TabIndex        =   42
            Top             =   1530
            Width           =   870
         End
         Begin VB.Label Label 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "Impuesto 2 :"
            Height          =   195
            Index           =   6
            Left            =   555
            TabIndex        =   41
            Top             =   1185
            Width           =   870
         End
         Begin VB.Label Label 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "Impuesto 1 :"
            Height          =   195
            Index           =   5
            Left            =   570
            TabIndex        =   40
            Top             =   840
            Width           =   870
         End
         Begin VB.Label Label 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "Neto :"
            Height          =   195
            Index           =   4
            Left            =   990
            TabIndex        =   39
            Top             =   495
            Width           =   435
         End
      End
      Begin VB.TextBox txtSerie 
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
         ForeColor       =   &H00800000&
         Height          =   285
         Left            =   2025
         MaxLength       =   17
         TabIndex        =   32
         Top             =   960
         Width           =   870
      End
      Begin VB.CommandButton cmdOpcion 
         Caption         =   "Número de Serie"
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
         Index           =   5
         Left            =   3007
         Style           =   1  'Graphical
         TabIndex        =   4
         Top             =   1650
         Width           =   1170
      End
      Begin VB.TextBox txtObservacion 
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
         Left            =   1740
         MaxLength       =   50
         TabIndex        =   30
         Top             =   1290
         Width           =   6315
      End
      Begin VB.TextBox txtNC1 
         BackColor       =   &H00E0E0E0&
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
         ForeColor       =   &H000000C0&
         Height          =   285
         Left            =   1740
         Locked          =   -1  'True
         TabIndex        =   16
         TabStop         =   0   'False
         Top             =   300
         Width           =   255
      End
      Begin VB.Label lblEstado 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Emitido"
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
         Left            =   7380
         TabIndex        =   61
         Top             =   390
         Width           =   630
      End
      Begin VB.Label Label 
         AutoSize        =   -1  'True
         Caption         =   "Motivo :"
         Height          =   195
         Index           =   2
         Left            =   1050
         TabIndex        =   31
         Top             =   1335
         Width           =   570
      End
      Begin VB.Label Label 
         AutoSize        =   -1  'True
         Caption         =   "Documento :"
         Height          =   195
         Index           =   3
         Left            =   705
         TabIndex        =   29
         Top             =   1005
         Width           =   915
      End
      Begin VB.Label Label 
         AutoSize        =   -1  'True
         Caption         =   "Fecha :"
         Height          =   195
         Index           =   1
         Left            =   1080
         TabIndex        =   28
         Top             =   675
         Width           =   540
      End
      Begin VB.Label Label 
         AutoSize        =   -1  'True
         Caption         =   "Nota de Crédito:"
         Height          =   195
         Index           =   0
         Left            =   465
         TabIndex        =   27
         Top             =   345
         Width           =   1155
      End
   End
   Begin VB.PictureBox Picture 
      Align           =   2  'Align Bottom
      Height          =   750
      Left            =   0
      ScaleHeight     =   690
      ScaleWidth      =   10965
      TabIndex        =   17
      TabStop         =   0   'False
      Top             =   4770
      Width           =   11025
      Begin VB.CommandButton cmdOpcion 
         Caption         =   "Procesar"
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
         Index           =   12
         Left            =   7440
         Picture         =   "frmNotaCreditoDetalle.frx":11BC
         Style           =   1  'Graphical
         TabIndex        =   13
         Top             =   60
         Width           =   1170
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
         Left            =   9780
         Picture         =   "frmNotaCreditoDetalle.frx":14FE
         Style           =   1  'Graphical
         TabIndex        =   15
         Top             =   60
         Width           =   1170
      End
      Begin VB.CommandButton cmdOpcion 
         Caption         =   "Eliminar"
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
         Left            =   8610
         Picture         =   "frmNotaCreditoDetalle.frx":15F0
         Style           =   1  'Graphical
         TabIndex        =   14
         Top             =   60
         Width           =   1170
      End
      Begin VB.CommandButton cmdOpcion 
         Caption         =   "Agregar"
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
         Left            =   5100
         Picture         =   "frmNotaCreditoDetalle.frx":16F2
         Style           =   1  'Graphical
         TabIndex        =   11
         Top             =   60
         Width           =   1170
      End
      Begin VB.CommandButton cmdOpcion 
         Caption         =   "Grabar"
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
         Left            =   6270
         Picture         =   "frmNotaCreditoDetalle.frx":1C24
         Style           =   1  'Graphical
         TabIndex        =   12
         Top             =   60
         Width           =   1170
      End
      Begin VB.PictureBox PicNavegacion 
         BackColor       =   &H80000004&
         Height          =   615
         Left            =   30
         ScaleHeight     =   555
         ScaleWidth      =   4950
         TabIndex        =   18
         Top             =   30
         Width           =   5010
         Begin VB.CommandButton cmdNavegar 
            Height          =   555
            Index           =   1
            Left            =   480
            Picture         =   "frmNotaCreditoDetalle.frx":2156
            Style           =   1  'Graphical
            TabIndex        =   24
            TabStop         =   0   'False
            Top             =   0
            Width           =   480
         End
         Begin VB.CommandButton cmdNavegar 
            Height          =   555
            Index           =   2
            Left            =   960
            Picture         =   "frmNotaCreditoDetalle.frx":2698
            Style           =   1  'Graphical
            TabIndex        =   23
            TabStop         =   0   'False
            Top             =   0
            Width           =   480
         End
         Begin VB.CommandButton cmdNavegar 
            Height          =   555
            Index           =   0
            Left            =   0
            Picture         =   "frmNotaCreditoDetalle.frx":2BDA
            Style           =   1  'Graphical
            TabIndex        =   22
            TabStop         =   0   'False
            Top             =   0
            Width           =   480
         End
         Begin VB.CommandButton cmdNavegar 
            Height          =   555
            Index           =   5
            Left            =   4470
            Picture         =   "frmNotaCreditoDetalle.frx":311C
            Style           =   1  'Graphical
            TabIndex        =   21
            TabStop         =   0   'False
            Top             =   0
            Width           =   480
         End
         Begin VB.CommandButton cmdNavegar 
            Height          =   555
            Index           =   4
            Left            =   3990
            Picture         =   "frmNotaCreditoDetalle.frx":365E
            Style           =   1  'Graphical
            TabIndex        =   20
            TabStop         =   0   'False
            Top             =   0
            Width           =   480
         End
         Begin VB.CommandButton cmdNavegar 
            Height          =   555
            Index           =   3
            Left            =   3510
            Picture         =   "frmNotaCreditoDetalle.frx":3BA0
            Style           =   1  'Graphical
            TabIndex        =   19
            TabStop         =   0   'False
            Top             =   0
            Width           =   480
         End
         Begin VB.Label cmdTexto 
            Alignment       =   2  'Center
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
            Height          =   255
            Left            =   1470
            TabIndex        =   25
            Top             =   150
            Width           =   1935
         End
      End
   End
   Begin VB.Image imagepIE 
      Height          =   135
      Left            =   0
      Top             =   0
      Visible         =   0   'False
      Width           =   135
   End
   Begin VB.Image imageCab 
      Height          =   135
      Left            =   0
      Top             =   0
      Visible         =   0   'False
      Width           =   135
   End
   Begin VB.Image imageHash 
      Height          =   615
      Left            =   11040
      Top             =   1920
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.Image Image 
      Height          =   4695
      Left            =   30
      Picture         =   "frmNotaCreditoDetalle.frx":40E2
      Stretch         =   -1  'True
      Top             =   30
      Width           =   2385
   End
End
Attribute VB_Name = "frmNotaCreditoDetalle"
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

Dim RsDetalle As Recordset
Dim RsDocumento As Recordset
Dim RsTipoDocumento As Recordset

Dim sDetalle As String
Dim sDocumento As String
Dim sNC As String

Dim nNeto As Double
Dim nImpuesto1 As Double
Dim nImpuesto2 As Double
Dim nImpuesto3 As Double
Dim nTotal As Double

Dim nNCNeto As Double
Dim nNCImp1 As Double
Dim nNCImp2 As Double
Dim nNCImp3 As Double
Dim nNCTotal As Double
Dim nCantidad As Double

'FACTURACION E
Dim RsNotaCredito As Recordset
Dim nEmision As Integer
Dim xTipoNotaCredito As String
Dim xCodigoMotivo As String
Dim cadenaCodigoHash As String
'-----
Dim sImp As String
Dim impTipo As String
Dim iImagenCab As Boolean
Dim RsImpDocumentoE As New Recordset
Dim xMontoTexto As String

Dim lNotaCreditoEmitirFE As Boolean
Dim tPrefijoEnlace As String

Dim TimpresionDolaresDelivery As Boolean

Sub Asignar()
    With frmNotaCredito.RsCabecera
        'Cuadro de Texto
        sNC = IIf(IsNull(!tNotaCredito), "", !tNotaCredito)
        txtNC1.Text = Mid(sNC, 1, 1)

        dtpFecha = IIf(IsNull(!fFecha), "", !fFecha)
        sDocumento = IIf(IsNull(!tDocumento), "", !tDocumento)
        txtPrefijo = Mid(sDocumento, 1, 1)
        
        If pais = "002" Then
            txtNC2.Text = Mid(sNC, 2, 6)
            txtNC3.Text = Mid(sNC, 8, 9)
            txtSerie = Mid(sDocumento, 2, 6)
            txtCorrela = Mid(sDocumento, 8, 9)
        Else
            txtNC2.Text = Mid(sNC, 2, 5)
            txtNC3.Text = Mid(sNC, 7, 9)
            txtSerie = Mid(sDocumento, 2, 5)
            txtCorrela = Mid(sDocumento, 7, 9)
        End If
        
        txtObservacion = IIf(IsNull(!tObservacion), "", !tObservacion)
        lblEstado.Caption = IIf(IsNull(!EstadoDocumento), "", !EstadoDocumento)
        
        nNCNeto = IIf(IsNull(!nNeto), 0, !nNeto)
        nNCImp1 = IIf(IsNull(!nImpuesto1), 0, !nImpuesto1)
        nNCImp2 = IIf(IsNull(!nImpuesto2), 0, !nImpuesto2)
        nNCImp3 = IIf(IsNull(!nImpuesto3), 0, !nImpuesto3)
        nNCTotal = IIf(IsNull(!nVenta), 0, !nVenta)
        
        nNeto = IIf(IsNull(!nDocNeto), 0, !nDocNeto)
        nImpuesto1 = IIf(IsNull(!nDocImpuesto1), 0, !nDocImpuesto1)
        nImpuesto2 = IIf(IsNull(!nDocImpuesto2), 0, !nDocImpuesto2)
        nImpuesto3 = IIf(IsNull(!nDocImpuesto3), 0, !nDocImpuesto3)
        nTotal = IIf(IsNull(!nDocVenta), 0, !nDocVenta)
                        
        txtNCNeto.Text = Format(nNCNeto, "#,###,##0.00")
        txtNCImp1.Text = Format(nNCImp1, "#,###,##0.00")
        txtNCImp2.Text = Format(nNCImp2, "#,###,##0.00")
        txtNCImp3.Text = Format(nNCImp3, "#,###,##0.00")
        txtNCTotal.Text = Format(nNCTotal, "#,###,##0.00")

        txtDocNeto.Text = Format(nNeto, "#,###,##0.00")
        txtDocImp1.Text = Format(nImpuesto1, "#,###,##0.00")
        txtDocImp2.Text = Format(nImpuesto2, "#,###,##0.00")
        txtDocImp3.Text = Format(nImpuesto3, "#,###,##0.00")
        txtDocTotal.Text = Format(nTotal, "#,###,##0.00")
        
        txtResNeto.Text = Format(nNeto - nNCNeto, "#,###,##0.00")
        txtResImp1.Text = Format(nImpuesto1 - nNCImp1, "#,###,##0.00")
        txtResImp2.Text = Format(nImpuesto2 - nNCImp2, "#,###,##0.00")
        txtResImp3.Text = Format(nImpuesto3 - nNCImp3, "#,###,##0.00")
        txtResTotal.Text = Format(nTotal - nNCTotal, "#,###,##0.00")
        
        cmdOpcion(4).Enabled = False
        cmdOpcion(5).Enabled = False
        cmdOpcion(6).Enabled = False
        cmdOpcion(8).Enabled = True
        cmdOpcion(9).Enabled = True
        'cmdOpcion(10).Enabled = False
        'cmdOpcion(11).Enabled = False
        cmdNotaCredito(1).Enabled = False
        cmdNotaCredito(2).Enabled = False

    End With
    
    If lblEstado.Caption = "ANULADO" Then
       cmdOpcion(1).Enabled = False
       cmdOpcion(2).Enabled = False
       cmdOpcion(12).Enabled = False
       cmdOpcion(8).Enabled = False
       cmdOpcion(9).Enabled = False
       dtpFecha.Enabled = False
    ElseIf lblEstado.Caption = "PROCESADO" Then
       cmdOpcion(1).Enabled = False
       cmdOpcion(12).Enabled = False
       cmdOpcion(8).Enabled = False
       cmdOpcion(9).Enabled = False
       dtpFecha.Enabled = False
    ElseIf lblEstado.Caption = "PAGADO" Then
       cmdOpcion(1).Enabled = False
       cmdOpcion(12).Enabled = False
       cmdOpcion(8).Enabled = False
       cmdOpcion(9).Enabled = False
       dtpFecha.Enabled = False
    Else
       cmdOpcion(1).Enabled = True
       cmdOpcion(2).Enabled = True
       cmdOpcion(12).Enabled = True
        If lNCElimina Then
         cmdOpcion(2).Enabled = False
        End If
        If lParcialNC Then
             Me.Frame.Enabled = False
             cmdOpcion(9).Enabled = False
        End If
        If lactivaFechaNC Then
          dtpFecha.Enabled = False
        Else
          dtpFecha.Enabled = True
        End If
    End If
    cmdTexto.Caption = "Registro " & frmNotaCredito.RsCabecera.AbsolutePosition & " de " & frmNotaCredito.RsCabecera.RecordCount
End Sub

Private Sub cmdNavegar_Click(Index As Integer)
    Select Case Index
           Case Is = 0 'Primero
                MoverPuntero Primero, frmNotaCredito.grdGrilla
                Asignar
                cmdTexto.Caption = "Registro " & IIf(frmNotaCredito.RsCabecera.RecordCount = 0, 0, frmNotaCredito.RsCabecera.AbsolutePosition) & " de " & frmNotaCredito.RsCabecera.RecordCount
           Case Is = 1 'PgUp
                MoverPuntero pgup, frmNotaCredito.grdGrilla
                Asignar
                cmdTexto.Caption = "Registro " & IIf(frmNotaCredito.RsCabecera.RecordCount = 0, 0, frmNotaCredito.RsCabecera.AbsolutePosition) & " de " & frmNotaCredito.RsCabecera.RecordCount
           Case Is = 2 'Previo
                MoverPuntero previo, frmNotaCredito.grdGrilla
                Asignar
                cmdTexto.Caption = "Registro " & IIf(frmNotaCredito.RsCabecera.RecordCount = 0, 0, frmNotaCredito.RsCabecera.AbsolutePosition) & " de " & frmNotaCredito.RsCabecera.RecordCount
           Case Is = 3 'Siguiente
                MoverPuntero siguiente, frmNotaCredito.grdGrilla
                Asignar
                cmdTexto.Caption = "Registro " & IIf(frmNotaCredito.RsCabecera.RecordCount = 0, 0, frmNotaCredito.RsCabecera.AbsolutePosition) & " de " & frmNotaCredito.RsCabecera.RecordCount
           Case Is = 4 'PgDn
                MoverPuntero pgdn, frmNotaCredito.grdGrilla
                Asignar
                cmdTexto.Caption = "Registro " & IIf(frmNotaCredito.RsCabecera.RecordCount = 0, 0, frmNotaCredito.RsCabecera.AbsolutePosition) & " de " & frmNotaCredito.RsCabecera.RecordCount
           Case Is = 5 'Ultimo
                MoverPuntero Ultimo, frmNotaCredito.grdGrilla
                Asignar
                cmdTexto.Caption = "Registro " & IIf(frmNotaCredito.RsCabecera.RecordCount = 0, 0, frmNotaCredito.RsCabecera.AbsolutePosition) & " de " & frmNotaCredito.RsCabecera.RecordCount
    End Select
End Sub

Private Sub cmdNotaCredito_Click(Index As Integer)
   Dim xDescripcion As String
   Dim xRsNotaCredito As Recordset
   
        If pais = "002" Then 'Ecuador
           Set xRsNotaCredito = Lib.OpenRecordset("select * from vTipoDocumentoImpresora where tCaja ='" & sCaja & "' and Transporte=0 and isnull(tNumeroAutorizacion,'')<>'' And lNotaCredito = 1 And lActivo = 1 UNION Select * From vTipoDocumentoImpresora where tCaja ='" & sCaja & "' and Transporte=0 And lNotaCredito = 1 And lFacturacionElectronica=1 and lActivo =1 order by tTipoEmision", Cn)
        Else
           Set xRsNotaCredito = Lib.OpenRecordset("select * from vTipoDocumentoImpresora where tCaja ='" & sCaja & "' and Transporte=0 And lNotaCredito = 1 And lActivo = 1 order by tTipoEmision", Cn)
        End If
   
   Select Case Index
   
       Case Is = 1
           xDescripcion = Trim(cmdNotaCredito(Index).Caption)
           xRsNotaCredito.Find "Descripcion='" & xDescripcion & "'"
           txtNC1.Text = xRsNotaCredito!prefijo
           xTipoNotaCredito = xRsNotaCredito!TTipoEmision
           lNotaCreditoEmitirFE = xRsNotaCredito!lFacturacionElectronica
           cmdNotaCredito(Index).Enabled = False
           cmdNotaCredito(2).Enabled = True
           tPrefijoEnlace = xRsNotaCredito!tPrefijoEnlace
       Case Is = 2
           xDescripcion = Trim(cmdNotaCredito(Index).Caption)
           xRsNotaCredito.Find "Descripcion='" & xDescripcion & "'"
           txtNC1.Text = xRsNotaCredito!prefijo
           xTipoNotaCredito = xRsNotaCredito!TTipoEmision
           lNotaCreditoEmitirFE = xRsNotaCredito!lFacturacionElectronica
           cmdNotaCredito(Index).Enabled = False
           cmdNotaCredito(1).Enabled = True
           tPrefijoEnlace = xRsNotaCredito!tPrefijoEnlace
   End Select
   
   If lNotaCreditoEmitirFE And tPrefijoEnlace <> "" Then
        cmdOpcion(4).Enabled = False
        txtPrefijo.Text = tPrefijoEnlace
   End If

End Sub

Private Sub cmdOpcion_Click(Index As Integer)

    Dim oComando As clsComando
    Dim oComandoDetalle As clsComando
    Dim sImporteLetra As String
    Dim RsCodigoHash As New ADODB.Recordset
    Dim fDocumento As String
    Dim lcodigoHash As Boolean
    
    Dim oComandoCabeceraOfisis As clsComando
    Dim oComandoDetalleOfisis As clsComando
    Dim oComandoFirmaDocumentoOfisis As clsComando
    Dim fso1 As Object
                                            
   Select Case Index
          Case Is = 0 'Agregar
          
               If Supervisor("27") = False Then
                      MsgBox "Clave no permitida", vbExclamation, sMensaje
                      Exit Sub
               End If
          
               ActivarBotones (False)
               Blanquear Me
               Sw = True
               nNeto = 0
               nImpuesto1 = 0
               nImpuesto2 = 0
               nImpuesto3 = 0
               nTotal = 0
               
               nNCNeto = 0
               nNCImp1 = 0
               nNCImp2 = 0
               nNCImp3 = 0
               nNCTotal = 0
               
               txtNCNeto.Text = "0.00"
               txtNCImp1.Text = "0.00"
               txtNCImp2.Text = "0.00"
               txtNCImp3.Text = "0.00"
               txtNCTotal.Text = "0.00"
               
               txtDocNeto.Text = "0.00"
               txtDocImp1.Text = "0.00"
               txtDocImp2.Text = "0.00"
               txtDocImp3.Text = "0.00"
               txtDocTotal.Text = "0.00"
               
               txtResNeto.Text = "0.00"
               txtResImp1.Text = "0.00"
               txtResImp2.Text = "0.00"
               txtResImp3.Text = "0.00"
               txtResTotal.Text = "0.00"
               
               cmdOpcion(1).Enabled = True
               cmdOpcion(3).Enabled = True
               cmdOpcion(4).Enabled = True
               cmdOpcion(5).Enabled = True
               cmdOpcion(6).Enabled = True
               cmdOpcion(8).Enabled = False
               cmdOpcion(9).Enabled = False
               dtpFecha.Enabled = True
               
               RsNotaCredito.Requery
               AsignaComando 2, RsNotaCredito, cmdNotaCredito()
               If RsNotaCredito.RecordCount = 1 Then
                  cmdNotaCredito(1).Enabled = True
               Else
                  cmdNotaCredito(1).Enabled = True
                  cmdNotaCredito(2).Enabled = True
               End If

               
               'cmdOpcion(10).Enabled = True
               'cmdOpcion(11).Enabled = True
               'txtPrefijo.Text = "F"
               'txtNC1.Text = "N"
               
          Case Is = 1 'Grabar
               'Chequea Datos
               Dim nCorrela As String
               Dim nPos As Integer
               
                       
                'FACTURACION E
               If txtNC1.Text = "" Then MsgBox "Seleccione el Tipo de Documento", vbExclamation, sMensaje: Exit Sub
               If txtPrefijo.Text = "" Then MsgBox "Ingrese el Documento a Afectar", vbExclamation, sMensaje: Exit Sub
               If txtSerie.Text = "" Then MsgBox "Ingrese la Serie del Documento a Afectar", vbExclamation, sMensaje: Exit Sub
               If txtCorrela.Text = "" Then MsgBox "Ingrese el Correlativo del Documento a Afectar", vbExclamation, sMensaje: Exit Sub
               If nNCTotal = 0 Then MsgBox "El valor de la Nota de Crédito debe ser mayor a cero", vbExclamation, sMensaje: Exit Sub
               If txtObservacion.Text = "" Then MsgBox "Ingrese el Motivo de la Nota de Crédito", vbExclamation, sMensaje: Exit Sub
                                          
               Dim sNSerie As String
               Dim sNPrefijo As String
               Dim sNCorrela As String
               Dim sNTipoEmision As String
               Dim sNDocumento As String
               '---------------
               
                'VALIDAR CDR DE DOCUMENTO
'                If pais = "000" Then
'                    If lFacturacionE Then
'                        Dim xCDR As String
'                        Dim RsDocumentoVenta As Recordset
'                        Dim xCont As Integer
'
'                        If lFEOfisis Then
'                                 'CORRECION  EDL
'                                fDocumento = Mid(sDocumento, 1, 1) + Mid(sDocumento, 4, 3) + "-" + CStr(CLng(Mid(sDocumento, 8, 8))) 'fDocumento = Mid(sDocumento, 1, 1) + Mid(sDocumento, 4, 3) + "-" + CStr(CINT(Mid(sDocumento, 8, 8)))
'                        ElseIf lFECarbajal Then
'                                 'CORRECION  EDL
'                                fDocumento = Mid(sDocumento, 1, 1) + Mid(sDocumento, 4, 3) + "-" + CStr(CLng(Mid(sDocumento, 8, 8))) 'fDocumento = Mid(sDocumento, 1, 1) + Mid(sDocumento, 4, 3) + "-" + CStr(CINT(Mid(sDocumento, 8, 8)))
'                        ElseIf lFEpape Then
'
'                        ElseIf lFESpring Then
'                                fDocumento = Mid(sDocumento, 1, 1) + Mid(sDocumento, 4, 3) + "-" + CStr(CLng(Mid(sDocumento, 8, 8))) 'fDocumento = Mid(sDocumento, 1, 1) + Mid(sDocumento, 4, 3) + "-" + CStr(CINT(Mid(sDocumento, 8, 8)))
'                        ElseIf lFacturacionE Then  'INFOFACT
'                                fDocumento = Mid(sDocumento, 1, 1) + Mid(sDocumento, 4, 3) + Mid(sDocumento, 8, 8)
'                                Isql = "Select * From dbo.DOCUMENTOVENTA where nro_efact='" & fDocumento & "' and tipodocu <>'07'"
'                                Set RsDocumentoVenta = Lib.OpenRecordset(Isql, CnFE)
'
'                                If RsDocumentoVenta.RecordCount > 0 Then
'                                    xCDR = IIf(IsNull(RsDocumentoVenta!cdr), "", RsDocumentoVenta!cdr)
'                                    xCont = Calcular("Select COUNT(*) As Codigo From dbo.DOCUMENTOVENTA where numerorefe='" & fDocumento & "'", CnFE)
'                                End If
'                        End If
'                    End If
'                End If
    
               If Sw Then
                  Sw = False
                   
                'FACTURACION E
                RsNotaCredito.Requery
                RsNotaCredito.MoveFirst
                RsNotaCredito.Find ("tTipoEmision='" & xTipoNotaCredito & "'")
    
                sNSerie = RsNotaCredito!tSerie
                sNPrefijo = RsNotaCredito!prefijo
                sNCorrela = Lib.Correlativo(RsNotaCredito!tUltimoNumero, 9)
                sNTipoEmision = RsNotaCredito!TTipoEmision
                sNC = sNPrefijo & sNSerie & sNCorrela
                Cn.Execute "Update TTIPODOCUMENTOIMPRESORA Set tUltimoNumero = '" & sNCorrela & "' where tTipoEmision ='" & sNTipoEmision & "' and tCaja ='" & sCaja & "'"
                '-------------
                                                      
                'Cambiar el SQL
                Isql = "insert into MNOTACREDITO( " & _
                         "tNotaCredito, fFecha, tDocumento, nNeto, nImpuesto1, nImpuesto2, nImpuesto3, nVenta, " & _
                         "tEstadoDocumento, tTurno, tCaja, tTipoDocumento, tMotivo, tUsuario, tObservacion, fDiaContable, fRegistro) " & _
                         "values ('" & sNC & "', '" & _
                                Format(dtpFecha.value, "yyyy/mm/dd") + " " + Format(Time, "hh:mm:ss") & "' , " & _
                                "'" & sDocumento & "', " & _
                                nNCNeto & ", " & _
                                nNCImp1 & ", " & _
                                nNCImp2 & ", " & _
                                nNCImp3 & ", " & _
                                nNCTotal & ", " & _
                                "'01', " & _
                                " '" & sTurno & "', " & _
                                " '" & sCaja & "', " & _
                                " '" & xTipoNotaCredito & "', " & _
                                " '" & xCodigoMotivo & "', " & _
                                "'" & sUsuario & "'," & _
                                " '" & txtObservacion.Text & "', '" & Format(obtieneDiaContable, "yyyyMMdd") & "', " & _
                                "getdate() )"
                  Cn.Execute Isql
                  
                  
                  'cambio anulacion por notas de credito
                  If modProcedimiento.pasa = True Then
                  frmNotaCredito.RsCabecera.Requery
                  frmNotaCredito.RsCabecera.Find "tNotaCredito ='" & sNC & "'"
                  End If
                  
                  MsgBox "Registro Guardado", vbInformation, sMensaje
                  
                  ' cambio anulacion por nota de credito
                  ActivarBotones (True)
                  If modProcedimiento.pasa = True Then
                  cmdTexto.Caption = "Registro " & IIf(frmNotaCredito.RsCabecera.RecordCount = 0, 0, frmNotaCredito.RsCabecera.AbsolutePosition) & " de " & frmNotaCredito.RsCabecera.RecordCount
                  frmNotaCredito.RsCabecera.Requery
                  frmNotaCredito.RsCabecera.MoveLast
                  End If

               Else
                  'Cambiar el SQL
                  Isql = "update MNOTACREDITO set " & _
                         "nNeto =" & nNCNeto & ", " & _
                         "nImpuesto2 =" & nNCImp2 & ", " & _
                         "nImpuesto3 =" & nNCImp3 & ", " & _
                         "nVenta =" & nNCTotal & ", " & _
                         "fFecha ='" & Format(dtpFecha.value, "yyyy/mm/dd") + " " + Format(Time, "hh:mm:ss") & "', " & _
                         "nImpuesto1 =" & nNCImp1 & ", " & _
                         "tMotivo ='" & xCodigoMotivo & "', " & _
                         "fDiaContable ='" & Format(obtieneDiaContable, "yyyyMMdd") & "', " & _
                         "tObservacion ='" & txtObservacion.Text & "',lreplica=1" & _
                         " where tNotaCredito= '" & sNC & "'"

                  Cn.Execute Isql
                  
                  'Cambiar el Nombre del Formulario
                  'cambios nota de credito anulacion
                  If modProcedimiento.pasa = True Then
                  nPos = frmNotaCredito.RsCabecera.AbsolutePosition
                  frmNotaCredito.RsCabecera.Requery
                  frmNotaCredito.RsCabecera.AbsolutePosition = nPos
                  End If
                  MsgBox "Registro Modificado", vbInformation, sMensaje
               End If
          
          
          
          Case Is = 2 'Eliminar
          
               If frmNotaCredito.RsCabecera.RecordCount = 0 Then
                  Exit Sub
               End If
               
                frmNotaCredito.RsCabecera.Requery
                frmNotaCredito.RsCabecera.MoveLast
               
               lblPaso1.Visible = True
                lblPaso2.Visible = True
                imgProceso(0).Visible = False
                imgProceso(1).Visible = False
                imgProceso(2).Visible = False
                imgProceso(3).Visible = False
                FrameFeSpring.Visible = False
               
               If frmNotaCredito.RsCabecera!tTurno = sTurno Then
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

               If MsgBox("Seguro de Eliminar la Nota de Crédito Nro." & sNC & "?", vbQuestion + vbOKCancel, sMensaje) = vbCancel Then
                  Exit Sub
               End If
                   
               If pais = "000" Then
                    If lFacturacionE Then
                          Dim lDocElecInfofact, lDocElecInfofactOfisis As Boolean
                          Dim xDocumentoVenta As String
                          Dim RsDocNotaCreditoFE As Recordset
                          lDocElecInfofact = Calcular("select isnull(tdi.lFacturacionElectronica,0) as codigo from TTIPODOCUMENTOIMPRESORA tdi inner join MNOTACREDITO m on tdi.tTipoEmision = m.tTipoDocumento and tdi.tCaja = m.tCaja  where m.tNotaCredito= '" & sNC & "'", Cn)
                          lDocElecInfofactOfisis = Calcular("select isnull(tdi.lDocumentoElectronicoOfisis,0) as codigo from TTIPODOCUMENTOIMPRESORA tdi inner join MNOTACREDITO m on tdi.tTipoEmision = m.tTipoDocumento and tdi.tCaja = m.tCaja  where m.tNotaCredito= '" & sNC & "'", Cn)
                          xDocumentoVenta = Calcular("select tdocumento as codigo from MNOTACREDITO where tNotaCredito='" & sNC & "'", Cn)
                          
                          If lDocElecInfofact And lDocElecInfofactOfisis = False Then
                            If lFEpape Then
                                If (Calcular("select isnull(tEstadoDocumento,'01')  as codigo from mnotacredito where tnotacredito='" & sNC & "'", Cn) = "01") Then
                                    Dim tTipoDocNotacredito, xUltimoCorrelativo As String
                                    tTipoDocNotacredito = Calcular("select isnull(ttipodocumento,'')  as codigo from mnotacredito where tnotacredito='" & sNC & "'", Cn)
                                    Cn.Execute "Delete mnotacredito Where tnotacredito= '" & sNC & "'"
                                    xUltimoCorrelativo = Calcular("select MAX(tnotacredito) as codigo from mnotacredito where tcaja='" & sCaja & "' and tTipoDocumento='" & tTipoDocNotacredito & "'", Cn)
                                    xUltimoCorrelativo = Right(xUltimoCorrelativo, 9)
                                    Cn.Execute "Update TTIPODOCUMENTOIMPRESORA Set tUltimoNumero = '" & xUltimoCorrelativo & "' where tTipoEmision ='" & tTipoDocNotacredito & "' and tCaja ='" & sCaja & "'"
                                    frmNotaCredito.RsCabecera.Requery
                                    Unload Me
                                    Exit Sub
                                Else
                                    If (Calcular("select isnull(tEstadoDocumento,'01')  as codigo from mnotacredito where tnotacredito='" & sNC & "'", Cn) = "01") Then
                                        Cn.Execute "update MNOTACREDITO set tEstadoDocumento = '04',lreplica=1 where tNotaCredito = '" & sNC & "'"
                                        frmNotaCredito.RsCabecera.Requery
                                    End If
                                End If
                            
                            ElseIf lFESpring Then
                            ElseIf lFEBiz Then
                            ElseIf lFECarbajal Then
                                    Label2.Caption = "   Proceso de anulación de documento en InfoFact......."
                                    lblPaso1.Caption = "Enviando información de documento a InfoFact."
                                    lblPaso2.Caption = "Obteniendo codigo " & IIf(lQRFE, "QR", IIf(lImpresionCodigoBarras, "de barras", " hash")) & " almacenado."
                                    sImporteLetra = NumeroCadena(str(Round(nNCTotal, 2))) + " " + sMonedaN
                                    FrameFeSpring.Visible = True
                                    Sleep 1000
                                    If Not INSERTAFE_CARVAJAL(sNC, sImporteLetra, 1, 1) Then '----CABECERA
                                            imgProceso(2).Visible = True
                                            imgProceso(3).Visible = True
                                            Sleep 1000
                                            FrameFeSpring.Visible = False
                                            Exit Sub
                                     End If
                                     imgProceso(0).Visible = True
                                     imgProceso(1).Visible = True
                                     Sleep 1500
                                     FrameFeSpring.Visible = False
                                     Cn.Execute "update MNOTACREDITO set tEstadoDocumento = '04',lreplica=1 where tNotaCredito = '" & sNC & "'"
                                     frmNotaCredito.RsCabecera.Requery
                            Else
                                fDocumento = Mid(xDocumentoVenta, 1, 1) + Mid(sNC, 4, 3) + Mid(sNC, 8, 8)
                                Isql = "Select * From dbo.DOCUMENTOVENTA where nro_efact='" & fDocumento & "' and tipodocu = '07'"
                                Set RsDocNotaCreditoFE = Lib.OpenRecordset(Isql, CnFE)
                            
                                If RsDocNotaCreditoFE.RecordCount > 0 Then
                                    Dim oComandoBaja As clsComando
                                    Set oComandoBaja = New clsComando
    
                                    If Not oComandoBaja.CreateCmdSp("USP_FactDocumentoBaja", Cn) Then
                                         Set oComandoBaja = Nothing
                                         Exit Sub
                                    End If
                                    oComandoBaja.CreateParameter "@CodigoDocumento", adVarChar, adParamInput, 20, sNC
                                    If Not oComandoBaja.GetParamOK Then
                                         Set oComandoBaja = Nothing
                                         Exit Sub
                                    End If
                                    If Not oComandoBaja.ExecSP Then
                                         Set oComandoBaja = Nothing
                                         Exit Sub
                                    End If
                                    Cn.Execute "update MNOTACREDITO set tEstadoDocumento = '04',lreplica=1 where tNotaCredito = '" & sNC & "'"
                                    frmNotaCredito.RsCabecera.Requery
                                Else
                                    MsgBox "Esta NC no se encuentra Procesada o no se ha Enviado a Facturacion Electronica!!!", vbInformation
                                End If
                            End If
                          Else
                            If lDocElecInfofactOfisis Then
                                MsgBox "La interfaz con Ofisis no soporta Anulación, La anulación se realizara solo para el sistema!!!", vbInformation
                            End If
                            Cn.Execute "update MNOTACREDITO set tEstadoDocumento = '04',lreplica=1 where tNotaCredito = '" & sNC & "'"
                            frmNotaCredito.RsCabecera.Requery
                          End If
                    Else
                          Cn.Execute "update MNOTACREDITO set tEstadoDocumento = '04',lreplica=1 where tNotaCredito = '" & sNC & "'"
                          frmNotaCredito.RsCabecera.Requery
                    End If
               Else
                    Cn.Execute "update MNOTACREDITO set tEstadoDocumento = '04',lreplica=1 where tNotaCredito = '" & sNC & "'"
                    frmNotaCredito.RsCabecera.Requery
               End If
               

               If frmNotaCredito.RsCabecera.RecordCount <> 0 Then
                  frmNotaCredito.RsCabecera.Requery
                  frmNotaCredito.RsCabecera.Find "tNotaCredito ='" & sNC & "'"
                  Asignar
               Else
                  ActivarBotones False
                  Blanquear Me
                  Sw = True
               End If

          Case Is = 3 'Salir
               Unload Me
          
          Case Is = 4 'Tipo Documento
               RsTipoDocumento.MoveNext
               If RsTipoDocumento.EOF Then
                  RsTipoDocumento.MoveFirst
               End If
               cmdOpcion(4).Caption = RsTipoDocumento!Descripcion
               txtPrefijo.Text = RsTipoDocumento!prefijo
                                       
          Case Is = 5 'KB Numero de Serie
               sDescrip = ""
               frmKeyBoard.txtResultado.Text = sDescrip
               frmKeyBoard.Show vbModal
               If wEnter Then
                    If pais = "002" Then 'ECUADOR
                       sDescrip = Mid(Trim(sDescrip), 1, 6)
                       txtSerie.Text = Mid("000000", 1, 6 - Len(Trim(sDescrip))) & Trim(sDescrip)
                    Else
                       sDescrip = Mid(Trim(sDescrip), 1, 5)
                       txtSerie.Text = Mid("00000", 1, 5 - Len(Trim(sDescrip))) & Trim(sDescrip)
                    End If
               End If
               
               sDescrip = ""

          Case Is = 6 'KB correlativo
               sTipo = "Numero"
               frmNumPad.Show vbModal
               If wEnter Then
                  txtCorrela.Text = Mid("000000000", 1, 9 - Len(Trim(sDescrip))) & Trim(sDescrip)
               End If
               
             
               'Consistencia
               If txtPrefijo.Text = "" Then
                  txtCorrela.Text = ""
                  Exit Sub
               End If
               
               If txtSerie.Text = "" Then
                  txtCorrela.Text = ""
                  Exit Sub
               End If
               
               
               'Busqueda
               sDocumento = txtPrefijo.Text & txtSerie.Text & txtCorrela.Text
            
               If lFacturacionE Then
                    If lNotaCreditoEmitirFE Then
                        If tPrefijoEnlace <> txtPrefijo.Text Then
                            MsgBox "Error : El tipo de documento " & sDocumento & " no puede se puede asociar a un tipo " & tPrefijoEnlace & " de Nota de Credito ", vbExclamation, sMensaje
                            txtPrefijo.Text = ""
                            txtSerie.Text = ""
                            txtCorrela.Text = ""
                            Exit Sub
                        End If
                    End If
               End If
               
               ' validar  si esta cancelado el documento
               Set RsDocumento = Lib.OpenRecordset("select * from MDOCUMENTO where tDocumento ='" & sDocumento & "' and tEstadoDocumento = '01'", Cn)
               If RsDocumento.RecordCount = 1 Then
                  MsgBox "Error : Documento no Cancelado", vbExclamation, sMensaje
                  txtCorrela.Text = ""
                  Set RsDocumento = Nothing
                  Exit Sub
               End If

               Set RsDocumento = Lib.OpenRecordset("select * from MDOCUMENTO where tDocumento ='" & sDocumento & "' and tEstadoDocumento <>'04'", Cn)
               If RsDocumento.RecordCount = 0 Then
                  MsgBox "Error : Documento no Existe", vbExclamation, sMensaje
                  txtCorrela.Text = ""
                  Set RsDocumento = Nothing
                  Exit Sub
               End If

               Dim CantNotaCredito, CantDocumento As Double
               Dim numnotacredito As Integer
               
                CantDocumento = Calcular("SELECT nVenta as codigo FROM MDOCUMENTO WHERE TDOCUMENTO = '" & sDocumento & "'", Cn)
                CantNotaCredito = Calcular("SELECT sum(nVenta) as codigo FROM MNOTACREDITO WHERE tDocumento= '" & sDocumento & "' AND tEstadoDocumento <>'04'", Cn)
                numnotacredito = Calcular("SELECT count(nVenta) as codigo FROM MNOTACREDITO WHERE tDocumento= '" & sDocumento & "' AND tEstadoDocumento <>'04'", Cn)
                
               If CDbl(CantNotaCredito) >= CDbl(CantDocumento) Then
                MsgBox ("!Ya se ha generado (" & numnotacredito & ") Notas de Credito por el total del Documento¡")
                 txtCorrela.Text = ""
                 Set RsDocumento = Nothing
                Exit Sub
               End If
               
               'Mensaje de documento pertenece a otra caja solo FE activo
               If lFacturacionE Then
                    If RsDocumento!tCaja <> sCaja Then
                        MsgBox "Tener en consideración que el Documento a canjear no pertenece a esta Caja", vbExclamation, sMensaje
                    End If
               End If
               
               cmdOpcion(4).Enabled = False
               cmdOpcion(5).Enabled = False
               cmdOpcion(6).Enabled = False
               'cmdOpcion(10).Enabled = False
               'cmdOpcion(11).Enabled = False
               cmdOpcion(8).Enabled = True
               cmdOpcion(9).Enabled = True
               
               cmdNotaCredito(1).Enabled = False
               cmdNotaCredito(2).Enabled = False
               
               nNeto = RsDocumento!nNeto
               nImpuesto1 = RsDocumento!nprecioImpuesto1
               nImpuesto2 = RsDocumento!nprecioImpuesto2
               nImpuesto3 = RsDocumento!nprecioImpuesto3
               nTotal = RsDocumento!nVenta
               
               txtDocNeto.Text = Format(nNeto, "#,###,##0.00")
               txtDocImp1.Text = Format(nImpuesto1, "#,###,##0.00")
               txtDocImp2.Text = Format(nImpuesto2, "#,###,##0.00")
               txtDocImp3.Text = Format(nImpuesto3, "#,###,##0.00")
               txtDocTotal.Text = Format(nTotal, "#,###,##0.00")
               
               txtResNeto.Text = Format(nNeto - nNCNeto, "#,###,##0.00")
               txtResImp1.Text = Format(nImpuesto1 - nNCImp1, "#,###,##0.00")
               txtResImp2.Text = Format(nImpuesto2 - nNCImp2, "#,###,##0.00")
               txtResImp3.Text = Format(nImpuesto3 - nNCImp3, "#,###,##0.00")
               txtResTotal.Text = Format(nTotal - nNCTotal, "#,###,##0.00")
               
'            If lParcialNC Then
'                Dim Acum As Double
'                cmdOpcion(9).Enabled = False
'                  nNCTotal = nTotal
'
'                  If CDbl(nNCTotal) > CDbl(txtDocTotal.Text) Then
'                    MsgBox ("!La cantidad Asignada no puede ser Mayor al monto Del documento¡"), vbInformation
'                    nNCTotal = CDbl(txtDocTotal.Text)
'                  End If
'                  Acum = 0
'                  Acum = IIf(nPorcentaje1 > 0, Acum + nPorcentaje1, Acum)
'                  Acum = IIf(nPorcentaje2 > 0, Acum + nPorcentaje2, Acum)
'                  Acum = IIf(nPorcentaje3 > 0, Acum + nPorcentaje3, Acum)
'                  Acum = 1 + (Acum / 100)
'                    Select Case pais ' ok
'                        Case "001" 'Bolivia
'                                nNCImp1 = IIf(nPorcentaje1 > 0, nNCTotal * nPorcentaje1 / 100, 0)
'                                nNCImp2 = IIf(nPorcentaje2 > 0, nNCTotal * nPorcentaje2 / 100, 0)
'                                nNCImp3 = IIf(nPorcentaje3 > 0, nNCTotal * nPorcentaje3 / 100, 0)
'                                nNCNeto = nNCTotal - nNCImp1 - nNCImp2 - nNCImp3
'
'                        Case Else 'Peru, Ecuador
'                                nNCImp1 = IIf(nPorcentaje1 > 0, nNCTotal / Acum * nPorcentaje1 / 100, 0)
'                                nNCImp2 = IIf(nPorcentaje2 > 0, nNCTotal / Acum * nPorcentaje2 / 100, 0)
'                                nNCImp3 = IIf(nPorcentaje3 > 0, nNCTotal / Acum * nPorcentaje3 / 100, 0)
'                                nNCNeto = nNCTotal - nNCImp1 - nNCImp2 - nNCImp3
'
'                        End Select
'
'               'End If
'
'               txtNCNeto.Text = Format(nNCNeto, "#,###,##0.00")
'               txtNCImp1.Text = Format(nNCImp1, "#,###,##0.00")
'               txtNCImp2.Text = Format(nNCImp2, "#,###,##0.00")
'               txtNCImp3.Text = Format(nNCImp3, "#,###,##0.00")
'               txtNCTotal.Text = Format(nNCTotal, "#,###,##0.00")
'
'               txtResNeto.Text = Format(nNeto - nNCNeto, "#,###,##0.00")
'               txtResImp1.Text = Format(nImpuesto1 - nNCImp1, "#,###,##0.00")
'               txtResImp2.Text = Format(nImpuesto2 - nNCImp2, "#,###,##0.00")
'               txtResImp3.Text = Format(nImpuesto3 - nNCImp3, "#,###,##0.00")
'               txtResTotal.Text = Format(nTotal - nNCTotal, "#,###,##0.00")
'            End If
            If lParcialNC Then
                Dim Acum As Double
                cmdOpcion(9).Enabled = False
                  nNCTotal = nTotal
                  
                  If CDbl(nNCTotal) > CDbl(txtDocTotal.Text) Then
                    MsgBox ("!La cantidad Asignada no puede ser Mayor al monto Del documento¡"), vbInformation
                    nNCTotal = CDbl(txtDocTotal.Text)
                  End If
                  Acum = 0
                  Acum = IIf(nPorcentaje1 > 0, Acum + nPorcentaje1, Acum)
                  Acum = IIf(nPorcentaje2 > 0, Acum + nPorcentaje2, Acum)
                  Acum = IIf(nPorcentaje3 > 0, Acum + nPorcentaje3, Acum)
                  Acum = 1 + (Acum / 100)
                    Select Case pais ' ok
                        Case "001" 'Bolivia
                                nNCImp1 = IIf(nPorcentaje1 > 0, nNCTotal * nPorcentaje1 / 100, 0)
                                nNCImp2 = IIf(nPorcentaje2 > 0, nNCTotal * nPorcentaje2 / 100, 0)
                                nNCImp3 = IIf(nPorcentaje3 > 0, nNCTotal * nPorcentaje3 / 100, 0)
                                nNCNeto = nNCTotal - nNCImp1 - nNCImp2 - nNCImp3
                                
                        Case Else 'Peru, Ecuador
                                nNCImp1 = IIf(nPorcentaje1 > 0, nNCTotal / Acum * nPorcentaje1 / 100, 0)
                                nNCImp2 = IIf(nPorcentaje2 > 0, nNCTotal / Acum * nPorcentaje2 / 100, 0)
                                nNCImp3 = IIf(nPorcentaje3 > 0, nNCTotal / Acum * nPorcentaje3 / 100, 0)
                                nNCNeto = nNCTotal - nNCImp1 - nNCImp2 - nNCImp3
                                
                        End Select
                  
               'End If
               
'               txtNCNeto.Text = Format(nNCNeto, "#,###,##0.00")
'               txtNCImp1.Text = Format(nNCImp1, "#,###,##0.00")
'               txtNCImp2.Text = Format(nNCImp2, "#,###,##0.00")
'               txtNCImp3.Text = Format(nNCImp3, "#,###,##0.00")
'               txtNCTotal.Text = Format(nNCTotal, "#,###,##0.00")
'                If lParcialNC Then
'                'Dim tNeto, tImp1, tImp2, tImp3, tTtoal As Double
                
                nNCNeto = Calcular("select isnull(nNeto,0) as Codigo from mdocumento where tdocumento='" & sDocumento & "'", Cn)
                nNCImp1 = Calcular("select isnull(nPrecioImpuesto1,0) as Codigo from mdocumento where tdocumento='" & sDocumento & "'", Cn)
                nNCImp2 = Calcular("select isnull(nPrecioImpuesto2,0) as Codigo from mdocumento where tdocumento='" & sDocumento & "'", Cn)
                nNCImp3 = Calcular("select isnull(nPrecioImpuesto3,0) as Codigo from mdocumento where tdocumento='" & sDocumento & "'", Cn)
                nNCTotal = Calcular("select isnull(nVenta,0) as Codigo from mdocumento where tdocumento='" & sDocumento & "'", Cn)
                
                txtNCNeto.Text = Format(nNCNeto, "#,###,##0.00")
                txtNCImp1.Text = Format(nNCImp1, "#,###,##0.00")
                txtNCImp2.Text = Format(nNCImp2, "#,###,##0.00")
                txtNCImp3.Text = Format(nNCImp3, "#,###,##0.00")
                txtNCTotal.Text = Format(nNCTotal, "#,###,##0.00")
               'End If
               
               txtResNeto.Text = Format(nNeto - nNCNeto, "#,###,##0.00")
               txtResImp1.Text = Format(nImpuesto1 - nNCImp1, "#,###,##0.00")
               txtResImp2.Text = Format(nImpuesto2 - nNCImp2, "#,###,##0.00")
               txtResImp3.Text = Format(nImpuesto3 - nNCImp3, "#,###,##0.00")
               txtResTotal.Text = Format(nTotal - nNCTotal, "#,###,##0.00")
            End If
          Case Is = 8 'KB Observacion
          
                    Select Case pais
                    
                      Case "000"
                          If lFacturacionE Then
                                If lNotaCreditoEmitirFE Then
                                        Isql = "Select * From vMotivoNotaCredito order by Codigo"
                                        
                                        frmBusquedaRapida.cmdOpcion(1).Enabled = False
                                        frmBusquedaRapida.cmdOpcion(2).Enabled = False
                                        frmBusquedaRapida.cmdOpcion(3).Enabled = False
                                        frmBusquedaRapida.nPredeterm = 1
                                        
                                        Call ConfGrilla(2, frmBusquedaRapida.grdGrilla, "Código", 2, "Codigo", 1500, 2, 0, "", _
                                                          "Motivo", 2, "Descripcion", 6600, 0, 0, "")
                                                          
                                        sTemp = IIf(sTemp = "0", "", sTemp)
                                        xCodigoMotivo = ""
                                        frmBusquedaRapida.Show vbModal
                                        
                                        If wEnter = True And sCodigo <> "" Then
                                           xCodigoMotivo = sCodigo
                                           txtObservacion.Text = Calcular("SELECT tDetallado As codigo FROM ttabla WHERE tTabla='MOTIVONOTACREDITO' And tCodigo='" & xCodigoMotivo & "'", Cn)
                                        Else
                                           Exit Sub
                                        End If
                                Else
                                        frmKeyBoard.txtResultado.Text = txtObservacion.Text
                                        frmKeyBoard.Show vbModal
                                        txtObservacion.Text = IIf(wEnter, sDescrip, txtObservacion.Text)
                                End If
                          Else
                              frmKeyBoard.txtResultado.Text = txtObservacion.Text
                              frmKeyBoard.Show vbModal
                              txtObservacion.Text = IIf(wEnter, sDescrip, txtObservacion.Text)
                          End If
                      
                      Case Else
                              frmKeyBoard.txtResultado.Text = txtObservacion.Text
                              frmKeyBoard.Show vbModal
                              txtObservacion.Text = IIf(wEnter, sDescrip, txtObservacion.Text)
                              
                    End Select
                               
          Case Is = 9 'NumPad Cantidad
               Dim Acumulado As Double
               sTipo = ""
               frmNumPad.Show vbModal
               If wEnter Then
                  nNCTotal = CDbl(sDescrip)
                  
                  If CDbl(nNCTotal) > CDbl(txtDocTotal.Text) Then
                  MsgBox ("!La cantidad Asignada no puede ser Mayor al monto Del documento¡"), vbInformation
                  nNCTotal = CDbl(txtDocTotal.Text)
                  End If
                  
                    Dim CantNotaCredito2, CantDocumento2 As Double
                    Dim numnotacredito2 As Integer

                    CantDocumento2 = Calcular("SELECT nVenta as codigo FROM MDOCUMENTO WHERE TDOCUMENTO = '" & sDocumento & "'", Cn)
                    CantNotaCredito2 = Calcular("SELECT sum(nVenta) as codigo FROM MNOTACREDITO WHERE tDocumento= '" & sDocumento & "' AND tEstadoDocumento <>'04'", Cn)
                    numnotacredito2 = Calcular("SELECT count(nVenta) as codigo FROM MNOTACREDITO WHERE tDocumento= '" & sDocumento & "' AND tEstadoDocumento <>'04'", Cn)

                    'CantNotaCredito2 = Val(CantNotaCredito2) + Val(nNCTotal)

                    If CDbl(CantNotaCredito2) + CDbl(nNCTotal) > CantDocumento2 Then
                     MsgBox ("!La cantidad Asiganada + las cantidades de las Notas de creditos Generados al documento es Mayor al monto Del documento ¡")
                       nNCTotal = CDbl(CantDocumento2) - CDbl(CantNotaCredito2)
                     'Exit Sub
                    End If

                  Acumulado = 0
                  Acumulado = IIf(nPorcentaje1 > 0, Acumulado + nPorcentaje1, Acumulado)
                  Acumulado = IIf(nPorcentaje2 > 0, Acumulado + nPorcentaje2, Acumulado)
                  Acumulado = IIf(nPorcentaje3 > 0, Acumulado + nPorcentaje3, Acumulado)
                  Acumulado = 1 + (Acumulado / 100)
                    Select Case pais ' ok
                        Case "001" 'Bolivia
                                nNCImp1 = IIf(nPorcentaje1 > 0, nNCTotal * nPorcentaje1 / 100, 0)
                                nNCImp2 = IIf(nPorcentaje2 > 0, nNCTotal * nPorcentaje2 / 100, 0)
                                nNCImp3 = IIf(nPorcentaje3 > 0, nNCTotal * nPorcentaje3 / 100, 0)
                                nNCNeto = nNCTotal - nNCImp1 - nNCImp2 - nNCImp3
                                
                        Case Else 'Peru, Ecuador
                                nNCImp1 = IIf(nPorcentaje1 > 0, nNCTotal / Acumulado * nPorcentaje1 / 100, 0)
                                nNCImp2 = IIf(nPorcentaje2 > 0, nNCTotal / Acumulado * nPorcentaje2 / 100, 0)
                                nNCImp3 = IIf(nPorcentaje3 > 0, nNCTotal / Acumulado * nPorcentaje3 / 100, 0)
                                nNCNeto = nNCTotal - nNCImp1 - nNCImp2 - nNCImp3
                                
                    End Select
                  
               End If
               
               txtNCNeto.Text = Format(nNCNeto, "#,###,##0.00")
               txtNCImp1.Text = Format(nNCImp1, "#,###,##0.00")
               txtNCImp2.Text = Format(nNCImp2, "#,###,##0.00")
               txtNCImp3.Text = Format(nNCImp3, "#,###,##0.00")
               txtNCTotal.Text = Format(nNCTotal, "#,###,##0.00")
               
               txtResNeto.Text = Format(nNeto - nNCNeto, "#,###,##0.00")
               txtResImp1.Text = Format(nImpuesto1 - nNCImp1, "#,###,##0.00")
               txtResImp2.Text = Format(nImpuesto2 - nNCImp2, "#,###,##0.00")
               txtResImp3.Text = Format(nImpuesto3 - nNCImp3, "#,###,##0.00")
               txtResTotal.Text = Format(nTotal - nNCTotal, "#,###,##0.00")
          
'          Case Is = 10 'KB Numero de Serie
'               sTipo = "Numero"
'               frmNumPad.Show vbModal
'               If wEnter Then
'                  txtNC2.Text = Mid("00000", 1, 5 - Len(Trim(sDescrip))) & Trim(sDescrip)
'               End If
'
'          Case Is = 11 'KB Numero Correlativo
'               sTipo = "Numero"
'               frmNumPad.Show vbModal
'               If wEnter Then
'                  txtNC3.Text = Mid("000000000", 1, 9 - Len(Trim(sDescrip))) & Trim(sDescrip)
'               End If
'
'               'Consistencia
'               If txtNC2.Text = "" Then
'                  txtNC3.Text = ""
'                  Exit Sub
'               End If
'
'               'Busqueda
'               sNC = txtNC1.Text & txtNC2.Text & txtNC3.Text
'               If Not Calcular("select tNotaCredito as codigo from MNOTACREDITO where tNotaCredito ='" & sNC & "'", Cn) = "0" Then
'                  MsgBox "Error : Documento Existente", vbExclamation, sMensaje
'                  sNC = ""
'                  txtNC3.Text = ""
'                  Exit Sub
'               End If
                                   
          Case Is = 12 'Procesar
               Dim xPedido As String
               Dim xEstadoNC As String
                
                lblPaso1.Visible = True
                lblPaso2.Visible = True
                imgProceso(0).Visible = False
                imgProceso(1).Visible = False
                imgProceso(2).Visible = False
                imgProceso(3).Visible = False
                FrameFeSpring.Visible = False
               
               xEstadoNC = Calcular("Select ISNULL(tEstadoDocumento,'') As Codigo From MNOTACREDITO Where tNotaCredito = '" & sNC & "' ", Cn)
               impTipo = "0"
               
'               IsqlFact = "select tNotaCredito,tDocumento,tCodigoPedido,fRegistro,tCaja,Cliente,Direccion,ISNULL(RUC,'') As RUC,ncantidad,producto,nprecioventa,nPrecioOficial,(nPrecioOficial-nprecioventa) As descuento,venta,nNeto,nPrecioImpuesto1, " & _
'                          "nPrecioImpuesto2,nVenta,nDescuento,tItem,Mesa,Mozo," & _
'                          "(SELECT SUM(D.nPrecioNeto*nCantidad) FROM dDocumento D WHERE D.tDocumento = '" & sDocumento & "' AND ((D.nPrecioImpuesto1>0 AND D.nPrecioImpuesto2=0) OR (D.nPrecioImpuesto1>0 AND D.nPrecioImpuesto2>0))) As Gravada," & _
'                          "(SELECT ISNULL(SUM(D.nPrecioNeto*nCantidad),0) FROM dDocumento D WHERE ((D.nPrecioImpuesto1<=0 AND D.nPrecioImpuesto2>0) OR(D.nPrecioImpuesto1<=0 AND D.nPrecioImpuesto2<=0)) AND D.tDocumento = '" & sDocumento & "') As Inafecta, tObservacion" & _
'                          " from vNotaCreditoImpresora where tNotaCredito='" & sNC & "' order by tItem"

               IsqlFact = "exec usp_inforest_Impresion '" & sNC & "',10 "
               Set RsImpDocumentoE = Lib.OpenRecordset(IsqlFact, Cn)
               
               If xEstadoNC = "01" Then
                      RsNotaCredito.Requery
                      RsNotaCredito.MoveFirst
                      
                      If pais = "002" Then
                          RsNotaCredito.Find ("tSerie='" & Mid(sNC, 2, 6) & "'")
                      Else
                          RsNotaCredito.Find ("tSerie='" & Mid(sNC, 2, 5) & "'")
                      End If
                   
                      If MsgBox("Deseas Procesar la Nota de Crédito Nro: " & sNC & " ? ", vbQuestion + vbYesNo, sMensaje) = vbYes Then
                        
                      
                             TimpresionDolaresDelivery = False
                             '------ impresion en dolares para check de cliente delivery
                             If Calcular("select isnull(lImpresionMonedaExtranjera,0) as codigo from MDOCUMENTO where tDocumento='" & sDocumento & "'", Cn) Then
                                     TimpresionDolaresDelivery = True
                                     Cn.Execute "update mnotacredito set lImpresionMonedaExtranjera=1 where tnotacredito='" & sNC & "'"
                             Else
                                 TimpresionDolaresDelivery = False
                                 Cn.Execute "update mnotacredito set lImpresionMonedaExtranjera=0 where tnotacredito='" & sNC & "'"
                             End If
                            
                            '-----------------------
                            If pais = "000" And lFEpape And IIf(RsNotaCredito!lFacturacionElectronica = True, 1, 0) Then
                                 If Not FacturarTCPIP(2, sNC, 1) Then
                                    Cn.Execute "update mnotacredito set lImpresionMonedaExtranjera=0 where tnotacredito='" & sNC & "'"
                                    Exit Sub
                                 End If
                            End If
                            '------------------------
                            If Round(nNCTotal, 2) = Round(nTotal, 2) Then
                                    xPedido = Calcular("select min(tCodigoPedido) as codigo from DDOCUMENTO where tDocumento ='" & sDocumento & "'", Cn)
                                    If IsNull(xPedido) Then
                                      MsgBox "Error : Documento sin pedido", vbExclamation, sMensaje
                                      Exit Sub
                                    End If
                                    
                                    If Not lFECarbajal Then
                                        Cn.Execute "update MPEDIDO set tEstadoPedido ='01'  where tCodigoPedido in (select tcodigopedido from dpedido where tdocumento='" & sDocumento & "')"
                                        Cn.Execute "update DPEDIDO set tDocumento ='', tFacturado ='' where tdocumento='" & sDocumento & "'"
                                        Cn.Execute "update MNOTACREDITO set tEstadodocumento='05',lreplica=1 where tNotaCredito='" & sNC & "'"
                                    End If
                                    
                                    'FACTURACION ELECTRONICA ECUADOR
                                    If pais = "002" Then
                                        If lFacturacionE And lFEEcuador = False Then
                                           If IIf(RsNotaCredito!lFacturacionElectronica = True, 1, 0) Then
                                                'CABECERA
                                                Set oComando = New clsComando
                                                If Not oComando.CreateCmdSp("USP_InsertaFactNotaCredito", Cn) Then
                                                     Set oComando = Nothing
                                                     Exit Sub
                                                End If
                                                oComando.CreateParameter "@NotaCredito", adVarChar, adParamInput, 20, sNC
                        
                                                If Not oComando.GetParamOK Then
                                                     Set oComando = Nothing
                                                     Exit Sub
                                                End If
                                                If Not oComando.ExecSP Then
                                                     Set oComando = Nothing
                                                     Exit Sub
                                                End If
                                                
                                                'DETALLE
                                                Set oComandoDetalle = New clsComando
                                                If Not oComandoDetalle.CreateCmdSp("USP_InsertaFactNotaCreditoDetalle", Cn) Then
                                                     Set oComandoDetalle = Nothing
                                                     Exit Sub
                                                End If
                                                oComandoDetalle.CreateParameter "@NotaCredito", adVarChar, adParamInput, 20, sNC
                            
                                                If Not oComandoDetalle.GetParamOK Then
                                                     Set oComandoDetalle = Nothing
                                                     Exit Sub
                                                End If
                                                If Not oComandoDetalle.ExecSP Then
                                                     Set oComandoDetalle = Nothing
                                                     Exit Sub
                                                End If
                                           End If
                                           If lFacturacionE And lFEEcuador Then
                                                 If INSERTA_FE_INFOREST(sDocumento, 2, DateTime.Now) = False Then
                                                     MsgBox "No se pudo enviar el documento a Facturacion Electronica!!! Verificar con su area de sistemas!!!"
                                                 End If
                                           End If
                                        End If
                                    End If
    
                            Else
                                    If Not lFECarbajal Then
                                        Cn.Execute "update MNOTACREDITO set tEstadodocumento='05' ,lreplica=1 where tNotaCredito='" & sNC & "'"
                                    End If
                            End If
                        
                            
                            'FACTURACION_E_PERU
                            If pais = "000" Then
                                If lFacturacionE Then
                                    If lFEOfisis Then 'OFISIS
                                            '----CABECERA
                                            Set oComandoCabeceraOfisis = New clsComando
                                            If Not oComandoCabeceraOfisis.CreateCmdSp("USP_FactNotaCreditoOfisis", Cn) Then
                                                 Set oComandoCabeceraOfisis = Nothing
                                                 Exit Sub
                                            End If
                                            oComandoCabeceraOfisis.CreateParameter "@CodigoDocumento", adVarChar, adParamInput, 20, sNC
                
                                            If Not oComandoCabeceraOfisis.GetParamOK Then
                                                 Set oComandoCabeceraOfisis = Nothing
                                                 Exit Sub
                                            End If
                                            If Not oComandoCabeceraOfisis.ExecSP Then
                                                 Set oComandoCabeceraOfisis = Nothing
                                                 Exit Sub
                                            End If
                                            
                                            '----FIRMA DOCUMENTO OFISIS
                                            If RsNotaCredito!lDocumentoElectronicoOfisis Then
                                                Set oComandoFirmaDocumentoOfisis = New clsComando
                                                If Not oComandoFirmaDocumentoOfisis.CreateCmdSp("USP_FactFirmaDocumentoOfisis", Cn) Then
                                                     Set oComandoFirmaDocumentoOfisis = Nothing
                                                     Exit Sub
                                                End If
                                                oComandoFirmaDocumentoOfisis.CreateParameter "@CodigoDocumento", adVarChar, adParamInput, 20, sNC
                    
                                                If Not oComandoFirmaDocumentoOfisis.GetParamOK Then
                                                     Set oComandoFirmaDocumentoOfisis = Nothing
                                                     Exit Sub
                                                End If
                                                If Not oComandoFirmaDocumentoOfisis.ExecSP Then
                                                     Set oComandoFirmaDocumentoOfisis = Nothing
                                                     Exit Sub
                                                End If
                                                
                                                'VALIDAR RESPUESTA CODIGO DE BARRA
                                                fDocumento = Mid(sDocumento, 1, 1) + Mid(sNC, 4, 3) + "-" + CStr(CLng(Mid(sNC, 8, 8)))
                                            End If
                                            
                                    ElseIf lFESpring Then
                                        If Round(nNCTotal, 2) = Round(nTotal, 2) Then
                                            Cn.Execute "update MPEDIDO set tEstadoPedido ='01'  where tCodigoPedido in (select tcodigopedido from dpedido where tdocumento='" & sDocumento & "')"
                                            Cn.Execute "update DPEDIDO set tDocumento ='', tFacturado ='' where tdocumento='" & sDocumento & "'"
                                            Cn.Execute "update MNOTACREDITO set tEstadodocumento='05',lreplica=1 where tNotaCredito='" & sNC & "'"
                                        Else
                                            Cn.Execute "update MNOTACREDITO set tEstadodocumento='05' ,lreplica=1 where tNotaCredito='" & sNC & "'"
                                        End If
                                        
                                    ElseIf lFECarbajal Then
                                        If IIf(RsNotaCredito!lFacturacionElectronica = True, 1, 0) Then
                                            
                                            Label2.Caption = "   Proceso de envio de documento a InfoFact......."
                                            lblPaso1.Caption = "Enviando información de documento a InfoFact."
                                            lblPaso2.Caption = "Obteniendo codigo " & IIf(lQRFE, "QR", IIf(lImpresionCodigoBarras, "de barras", " hash")) & " almacenado."
                                            sImporteLetra = NumeroCadena(str(Round(nNCTotal, 2))) + " " + sMonedaN
                                            FrameFeSpring.Visible = True
                                            Sleep 1000
                                            If Not INSERTAFE_CARVAJAL(sNC, sImporteLetra, 1, 0) Then '----CABECERA
                                                    imgProceso(2).Visible = True
                                                    imgProceso(3).Visible = True
                                                    Sleep 1000
                                                    FrameFeSpring.Visible = False
                                                    Exit Sub
                                             End If
                                             imgProceso(0).Visible = True
                                             imgProceso(1).Visible = True
                                             Sleep 1500
                                             FrameFeSpring.Visible = False
                                             impTipo = "1"
                                        End If
                                        If Round(nNCTotal, 2) = Round(nTotal, 2) Then
                                            Cn.Execute "update MPEDIDO set tEstadoPedido ='01'  where tCodigoPedido in (select tcodigopedido from dpedido where tdocumento='" & sDocumento & "')"
                                            Cn.Execute "update DPEDIDO set tDocumento ='', tFacturado ='' where tdocumento='" & sDocumento & "'"
                                            Cn.Execute "update MNOTACREDITO set tEstadodocumento='05',lreplica=1 where tNotaCredito='" & sNC & "'"
                                        Else
                                            Cn.Execute "update MNOTACREDITO set tEstadodocumento='05' ,lreplica=1 where tNotaCredito='" & sNC & "'"
                                        End If
                                        
                                        
                                    ElseIf lFEpape Then
                                        impTipo = "1" 'IMPRESION FORMATO ELECTRONICO
                                    ElseIf lFEBiz Then
                                        impTipo = "1"
                                        If IIf(RsNotaCredito!lFacturacionElectronica = True, 1, 0) Then
                                            If Not INSERTA_FE_INFOREST(sNC, 2, DateTime.Date) Then '----CABECERA
                                                MsgBox "No se pudo procesar la Nota de credito!! Favor de verificar la informacion", vbInformation, sMensaje
                                                Exit Sub
                                             End If
                                             Sleep 1500
                                        End If
                                        If Round(nNCTotal, 2) = Round(nTotal, 2) Then
                                            Cn.Execute "update MPEDIDO set tEstadoPedido ='01'  where tCodigoPedido in (select tcodigopedido from dpedido where tdocumento='" & sDocumento & "')"
                                            Cn.Execute "update DPEDIDO set tDocumento ='', tFacturado ='' where tdocumento='" & sDocumento & "'"
                                            Cn.Execute "update MNOTACREDITO set tEstadodocumento='05',lreplica=1 where tNotaCredito='" & sNC & "'"
                                        Else
                                            Cn.Execute "update MNOTACREDITO set tEstadodocumento='05' ,lreplica=1 where tNotaCredito='" & sNC & "'"
                                        End If
                                        
                                    ElseIf lFEGesa Then
                                        If IIf(RsNotaCredito!lFacturacionElectronica = True, 1, 0) Then
                                            'CABECERA
                                            Set oComando = New clsComando
                                            If Not oComando.CreateCmdSp("USP_FactNotaCredito", Cn) Then
                                                 Set oComando = Nothing
                                                 Exit Sub
                                            End If
                                            oComando.CreateParameter "@NotaCredito", adVarChar, adParamInput, 20, sNC
                                            oComando.CreateParameter "@CodigoDocumento", adVarChar, adParamInput, 250, ""
                                            If Not oComando.GetParamOK Then
                                                 Set oComando = Nothing
                                                 'Exit Sub
                                            End If
                                            If Not oComando.ExecSP Then
                                                 Set oComando = Nothing
                                                 'Exit Sub
                                            End If
                                            impTipo = "1" 'IMPRESION FORMATO ELECTRONICO
                                        End If
                                    
                                    
                                    Else 'INFOFACT
                                        If IIf(RsNotaCredito!lFacturacionElectronica = True, 1, 0) Then
                                            sImporteLetra = NumeroCadena(str(Round(nNCTotal, 2))) + " " + sMonedaN
                                            'CABECERA
                                            Set oComando = New clsComando
                                            If Not oComando.CreateCmdSp("USP_FactNotaCredito", Cn) Then
                                                 Set oComando = Nothing
                                                 Exit Sub
                                            End If
                                            oComando.CreateParameter "@NotaCredito", adVarChar, adParamInput, 20, sNC
                                            oComando.CreateParameter "@CodigoDocumento", adVarChar, adParamInput, 250, sImporteLetra
                    
                                            If Not oComando.GetParamOK Then
                                                 Set oComando = Nothing
                                                 Exit Sub
                                            End If
                                            If Not oComando.ExecSP Then
                                                 Set oComando = Nothing
                                                 Exit Sub
                                            End If
                                            
                                            impTipo = "1" 'IMPRESION FORMATO ELECTRONICO
                                            fDocumento = Mid(sDocumento, 1, 1) + Mid(sNC, 4, 3) + Mid(sNC, 8, 8)
                                        End If
                                    
                                    End If

                                End If
                            End If
                            
                            'IMPRESION NC
                            Imprimir (sImp)
                            Printer.FontName = sFont
                            Printer.FontBold = False
                            
                            Dim RsImpresion As Recordset
                            Isql = "exec usp_inforest_Impresion '" & sNC & "',11 "
                            'Isql = "Select * From vNotaCreditoImpresora Where tNotaCredito='" & sNC & "'"
                            Set RsImpresion = Lib.OpenRecordset(Isql, Cn)
                            
                            Dim rstFuente As Recordset
                            Set rstFuente = New ADODB.Recordset
                            imageCab.Picture = Nothing
                            imagepIE.Picture = Nothing
                            Set rstFuente = Lib.OpenRecordset("select iImagenCabDoc AS foto, iImagenPieDoc as fotoPie  from tcaja where tcaja='" & sCaja & "'", Cn)
                            imageCab.DataField = "foto"
                            Set imageCab.DataSource = rstFuente
                            imagepIE.DataField = "fotoPie"
                            Set imagepIE.DataSource = rstFuente
                            sNTipoEmision = RsNotaCredito!TTipoEmision
                            
                            
                            If pais = "000" Then 'PERU
                                If lFacturacionE Then
                                        If lFEOfisis Then
                                                If RsNotaCredito!lDocumentoElectronicoOfisis Then
                                                  impTipo = "1"
                                                  Sleep 2000
                                                  
                                                  If lImpresionCodigoBarras Then
                                                        imageHash.DataField = "foto"
                                                        Set RsCodigoHash = Lib.OpenRecordset("USP_FactObtenerCodigoBarraOfisis '" & fDocumento & "','D','' ", Cn)
                                                        Set imageHash.DataSource = RsCodigoHash
                                                        
                                                    ElseIf lQRFE Then
                                                        Set imageHash.Picture = LoadPicture(ImagenQR_Ofisis(fDocumento, "D"))
                                                  Else
                                                        Dim RscadenaCodigoHash As Recordset

                                                        Set RscadenaCodigoHash = Lib.OpenRecordset("USP_FactConsultaHash '" & fDocumento & "','1' ", Cn)
                                                        If RscadenaCodigoHash.RecordCount > 0 Then
                                                            cadenaCodigoHash = RscadenaCodigoHash!codigo
                                                        End If
                                                        'cadenaCodigoHash = Calcular("select CO_HASH as codigo from TCFACT_ELEC where NU_DOCU='" & fDocumento & "' and TI_DOCU='D' ", CnFE)
                                                  End If
                                                  
                                                  ImprimeNotaCredito RsImpresion, imageHash, impTipo, cadenaCodigoHash, TimpresionDolaresDelivery, imageCab, imagepIE, sNTipoEmision
                                                Else
                                                  impTipo = "0"
                                                  ImprimeNotaCredito RsImpresion, imageHash, impTipo, cadenaCodigoHash, TimpresionDolaresDelivery, imageCab, imagepIE, sNTipoEmision
                                                End If
                                        
                                        ElseIf lFESpring Then
                                        
                                            ImprimeNotaCredito RsImpresion, imageHash, impTipo, cadenaCodigoHash, TimpresionDolaresDelivery, imageCab, imagepIE, sNTipoEmision
                                            
                                        ElseIf lFECarbajal Then
                                            If IIf(RsNotaCredito!lFacturacionElectronica = True, 1, 0) Then
                                                If tCodigoFE = "000" Then
                                                     If lQRFE Then
                                                         Set imageHash.Picture = LoadPicture(ImagenFeCarvajal(3, sNC, 1))
                                                     Else
                                                         If lImpresionCodigoBarras Then
                                                             Set imageHash.Picture = LoadPicture(ImagenFeCarvajal(1, sNC, 1))
                                                         Else
                                                             cadenaCodigoHash = ImagenFeCarvajal(2, sNC, 1)
                                                         End If
                                                     End If
                                                 End If
                                                If RsNotaCredito!tFormulario = "01" Then
                                                    ImprimeNotaCredito RsImpresion, imageHash, impTipo, cadenaCodigoHash, TimpresionDolaresDelivery, imageCab, imagepIE, sNTipoEmision
                                                Else
                                                    ImprimeFormatoA
                                                    Set fso1 = CreateObject("Scripting.FileSystemObject")
                                                    If fso1.FileExists(App.Path & "\fact.bmp") Then
                                                        Kill App.Path & "\fact.bmp"
                                                    End If
                                                End If
                                            End If
                                        ElseIf lFEpape Then
                                            If IIf(RsNotaCredito!lFacturacionElectronica = True, 1, 0) Then
                                               If RsNotaCredito!tFormulario = "01" Then
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
                                                     ImprimeNotaCredito RsImpresion, imageHash, impTipo, cadenaCodigoHash, TimpresionDolaresDelivery, imageCab, imagepIE, sNTipoEmision
                                                Else
                                                    CrearImagenQR (PapeTermico)
                                                    ImprimeFormatoA
                                                    Kill App.Path & "\BaseTempQr.bmp"
                                                End If
                                            End If
                                        ElseIf lFEBiz Then
                                            If IIf(RsNotaCredito!lFacturacionElectronica = True, 1, 0) Then
                                                If tCodigoFE = "000" Then
                                                     If lQRFE Then
                                                         Set imageHash.Picture = LoadPicture(QRHASH_FE_INFOREST(3, sNC, 1))
                                                     Else
                                                         If lImpresionCodigoBarras Then
                                                             Set imageHash.Picture = LoadPicture(QRHASH_FE_INFOREST(1, sNC, 1))
                                                         Else
                                                             cadenaCodigoHash = QRHASH_FE_INFOREST(2, sNC, 1)
                                                         End If
                                                     End If
                                                 End If
                                                If RsNotaCredito!tFormulario = "01" Then
                                                     ImprimeNotaCredito RsImpresion, imageHash, impTipo, cadenaCodigoHash, TimpresionDolaresDelivery, imageCab, imagepIE, sNTipoEmision
                                                Else
                                                    ImprimeFormatoA
                                                    
                                                    Set fso1 = CreateObject("Scripting.FileSystemObject")
                                                    If fso1.FileExists(App.Path & "\fact.bmp") Then
                                                        Kill App.Path & "\fact.bmp"
                                                    End If
                                                End If
                                            End If
                                        ElseIf lFEGesa Then
                                        
                                             If IIf(RsNotaCredito!lFacturacionElectronica = True, 1, 0) Then
                                             
                                                If lQRFE Then
                                                    Set imageHash.Picture = LoadPicture(QRHASH_FE_INFOREST(3, sNC, 1))
                                                Else
                                                    If lImpresionCodigoBarras Then
                                                        Set imageHash.Picture = LoadPicture(QRHASH_FE_INFOREST(1, sNC, 1))
                                                    Else
                                                        cadenaCodigoHash = QRHASH_FE_INFOREST(2, sNC, 1)
                                                    End If
                                                End If
                                                     
                                                If RsNotaCredito!tFormulario = "01" Then
                                                     ImprimeNotaCredito RsImpresion, imageHash, impTipo, cadenaCodigoHash, TimpresionDolaresDelivery, imageCab, imagepIE, sNTipoEmision
                                                Else
                                                    ImprimeFormatoA
                                                    
                                                    Set fso1 = CreateObject("Scripting.FileSystemObject")
                                                    If fso1.FileExists(App.Path & "\fact.bmp") Then
                                                        Kill App.Path & "\fact.bmp"
                                                    End If
                                                End If
                                            End If
                                        
                                        Else 'INFOFACT
                                                If IIf(RsNotaCredito!lFacturacionElectronica = True, 1, 0) Then
                                                     If RsNotaCredito!tFormulario = "01" Then
                                                         'VALIDAR RESPUESTA DE CODIGO HASH
                                                         If tCodigoFE = "000" Then
                                                         
                                                            If lQRFE Then
                                                                Set imageHash.Picture = LoadPicture(ImagenQR(sNC))
                                                            Else
                                                            
                                                                If lImpresionCodigoBarras Then
                                                                    Set imageHash.Picture = LoadPicture(lValidaCodBarra(lImpresionCodigoBarras, sNC))
                                                                Else
                                                                    cadenaCodigoHash = lValidaCodBarra(lImpresionCodigoBarras, sNC)
                                                                End If
                                                            End If

                                                         End If
                                                                                          
                                                         ImprimeNotaCredito RsImpresion, imageHash, impTipo, cadenaCodigoHash, TimpresionDolaresDelivery, imageCab, imagepIE, sNTipoEmision
                                                         
                                                     Else
                                                         'FORMATO A4
                                                         If Generar_Imagen(CnFE, "select imagen from IMAGENCODIGOBARRA where nro_efact='" & fDocumento & "' and tipodocu = '07'", "imagen", "\fact.bmp") = True Then
                                                            ImprimeFormatoA
                                                            Kill App.Path & "\fact.bmp"
                                                         Else
                                                            ImprimeFormatoA
                                                         End If
                                                     End If
                                                 
                                                Else 'NO ELECTRONICO
                                                     ImprimeNotaCredito RsImpresion, imageHash, impTipo, cadenaCodigoHash, TimpresionDolaresDelivery, imageCab, imagepIE, sNTipoEmision
                                                End If
                                        End If
                                    
                                 Else 'NO ELECTRONICO
                                      ImprimeNotaCredito RsImpresion, imageHash, impTipo, cadenaCodigoHash, TimpresionDolaresDelivery, imageCab, imagepIE, sNTipoEmision
                                 End If

                            Else 'ECUADOR
                                ImprimeNotaCredito RsImpresion, imageHash, impTipo, cadenaCodigoHash, TimpresionDolaresDelivery, imageCab, imagepIE, sNTipoEmision
                            End If
                            '------------------------------
                            '-----------------------
                            
                            If lNcOfisis Then
                                Dim MNTurno As String
                                Dim MNDocRef As String
                                'Dim nMontoDoc As Float
                                MNTurno = Calcular("select isnull(tturno,'') as codigo from mnotacredito where tNotaCredito='" & sNC & "'", Cn)
                                MNDocRef = Calcular("select isnull(tdocumento,'') as codigo from mnotacredito where tNotaCredito='" & sNC & "'", Cn)
                                If MNTurno = sTurno And Round(nNCTotal, 2) = Round(nTotal, 2) Then
                                    Cn.Execute "delete from dpagodocumento where tdocumento='" & MNDocRef & "'"
                                     Isql = "insert into DPAGODOCUMENTO " & _
                                     "( tDocumento, tCorrelativo, tTurno, tTipoPago, tOtroTipoPago, tMoneda, nTipoCambio, nMonto, tNumero, tBanco, fRegistro, tUsuario,fDiaContable ) " & _
                                     "Values(   '" & MNDocRef & "'," _
                                                & "1," _
                                                & "'" & sTurno & "'," _
                                                & "'04'," _
                                                & "'002'," _
                                                & "'01'," _
                                                & nTC & ", " _
                                                 & Round(nNCTotal, 2) & ", " _
                                                & "'" & sNC & "', " _
                                                & "'', " _
                                                & "getdate()," _
                                                & "'" & sUsuario & "','" & Format(obtieneDiaContable, "yyyyMMdd") & "')"
                                    Cn.Execute Isql
'                                    If sOtroTipoCancelacion = "001" Then
'                                        Cn.Execute "update MINGRESO set tEstadoDocumento ='02' where tRecibo ='" & sTipoDocumento & "'"
'                                    ElseIf sOtroTipoCancelacion = "002" Then
                                    Cn.Execute "update MNOTACREDITO set tEstadoDocumento ='02',lreplica=1 where tNotaCredito ='" & sNC & "'"
'                                    End If
                                    
                                    
                                End If
                            End If
                            
                            If pais = "000" And lFEpape And IIf(RsNotaCredito!lFacturacionElectronica = True, 1, 0) Then
                                 If Not FacturarTCPIP(3, sNC, 1) Then
                                    MsgBox ("La confirmacion ha fallado favor de contactarse con paperlees"), vbInformation, sMensaje
                                 End If
                            End If
                            '------------------------
                            If modProcedimiento.pasa = True Then
                            frmNotaCredito.RsCabecera.Requery
                            frmNotaCredito.RsCabecera.Find "tNotaCredito ='" & sNC & "'"
                            Asignar
                            Else
                            Unload Me
                            End If
                            
                        
                      End If
                 
              End If
               
               
          Case Is = 13 'Correccion Impuesto1
               sTipo = ""
               frmNumPad.Show vbModal
               
               If wEnter Then
                  nNCImp1 = Val(sDescrip)
                  nNCNeto = nNCTotal - nNCImp1 - nNCImp2 - nNCImp3
               End If
               
               txtNCNeto.Text = Format(nNCNeto, "#,###,##0.00")
               txtNCImp1.Text = Format(nNCImp1, "#,###,##0.00")
               txtNCImp2.Text = Format(nNCImp2, "#,###,##0.00")
               txtNCImp3.Text = Format(nNCImp3, "#,###,##0.00")
               txtNCTotal.Text = Format(nNCTotal, "#,###,##0.00")
               
               txtResNeto.Text = Format(nNeto - nNCNeto, "#,###,##0.00")
               txtResImp1.Text = Format(nImpuesto1 - nNCImp1, "#,###,##0.00")
               txtResImp2.Text = Format(nImpuesto2 - nNCImp2, "#,###,##0.00")
               txtResImp3.Text = Format(nImpuesto3 - nNCImp3, "#,###,##0.00")
               txtResTotal.Text = Format(nTotal - nNCTotal, "#,###,##0.00")
                    
                    
          Case Is = 14 'Correccion Impuesto1
               sTipo = ""
               frmNumPad.Show vbModal
               
               If wEnter Then
                  nNCImp2 = Val(sDescrip)
                  nNCNeto = nNCTotal - nNCImp1 - nNCImp2 - nNCImp3
               End If
               
               txtNCNeto.Text = Format(nNCNeto, "#,###,##0.00")
               txtNCImp1.Text = Format(nNCImp1, "#,###,##0.00")
               txtNCImp2.Text = Format(nNCImp2, "#,###,##0.00")
               txtNCImp3.Text = Format(nNCImp3, "#,###,##0.00")
               txtNCTotal.Text = Format(nNCTotal, "#,###,##0.00")
               
               txtResNeto.Text = Format(nNeto - nNCNeto, "#,###,##0.00")
               txtResImp1.Text = Format(nImpuesto1 - nNCImp1, "#,###,##0.00")
               txtResImp2.Text = Format(nImpuesto2 - nNCImp2, "#,###,##0.00")
               txtResImp3.Text = Format(nImpuesto3 - nNCImp3, "#,###,##0.00")
               txtResTotal.Text = Format(nTotal - nNCTotal, "#,###,##0.00")
                    
          Case Is = 15 'Correccion Impuesto1
               sTipo = ""
               frmNumPad.Show vbModal
               
               If wEnter Then
                  nNCImp3 = Val(sDescrip)
                  nNCNeto = nNCTotal - nNCImp1 - nNCImp2 - nNCImp3
               End If
               
               txtNCNeto.Text = Format(nNCNeto, "#,###,##0.00")
               txtNCImp1.Text = Format(nNCImp1, "#,###,##0.00")
               txtNCImp2.Text = Format(nNCImp2, "#,###,##0.00")
               txtNCImp3.Text = Format(nNCImp3, "#,###,##0.00")
               txtNCTotal.Text = Format(nNCTotal, "#,###,##0.00")
               
               txtResNeto.Text = Format(nNeto - nNCNeto, "#,###,##0.00")
               txtResImp1.Text = Format(nImpuesto1 - nNCImp1, "#,###,##0.00")
               txtResImp2.Text = Format(nImpuesto2 - nNCImp2, "#,###,##0.00")
               txtResImp3.Text = Format(nImpuesto3 - nNCImp3, "#,###,##0.00")
               txtResTotal.Text = Format(nTotal - nNCTotal, "#,###,##0.00")
                    
   End Select
End Sub

Private Sub ImprimeFormatoA()
                    
                    Dim xImpresionFE As String
                    
                    'NEW
                    Dim RsImpresionNC As Recordset
                    Isql = "Select * From MNOTACREDITO Where tNotaCredito='" & sNC & "'"
                    Set RsImpresionNC = Lib.OpenRecordset(Isql, Cn)
                    
                    Dim xMotivoNT As String
                    xMotivoNT = Calcular("Select ISNULL(tMotivo,'06') As Codigo from MNOTACREDITO Where tNotacredito = '" & sNC & "'", Cn)
                    
                    Dim xNeto As String
                    Dim xVenta As String
                    Dim xImp1 As String
                    Dim xImp2 As String
                    
                    xNeto = Format(RsImpresionNC!nNeto, "##,###,##0.00")
                    xVenta = Format(RsImpresionNC!nVenta, "##,###,##0.00")
                    xImp1 = Format(RsImpresionNC!nImpuesto1, "##,###,##0.00")
                    xImp2 = Format(RsImpresionNC!nImpuesto2, "##,###,##0.00")
                    '----------------------
                    
                    
                    xImpresionFE = Calcular("SELECT tImpresionFE as codigo FROM vTipodocumento WHERE Codigo=(SELECT TTIPODOCUMENTO FROM MNOTACREDITO WHERE tNotaCredito='" & sNC & "')", Cn)

                    If RsNotaCredito!lImprimeImageCab Then
                       iImagenCab = Generar_Imagen(Cn, "select iImagenCabDoc As imagen from TCAJA where tCaja='" & sCaja & "'", "imagen", "\cliente.jpg")
                    End If
                    
                    If xMotivoNT = "06" Then
                        Dim Reporte As New dsrNotaCredito
                        
                        Reporte.DiscardSavedData
                        Reporte.Database.SetDataSource RsImpDocumentoE
                        
                        Reporte.Text13.SetText "NOTA DE CREDITO ELECTRONICA"
                        
                        Reporte.Text8.SetText sRazonSocial
                        Reporte.ReportTitle = sDireccion
                        Reporte.Text15.SetText sTelefono
                        Reporte.Text14.SetText sFax
                        Reporte.Text16.SetText sRUC
                        Reporte.Text50.SetText sWeb
                        
                        If Calcular(" SELECT case when  lImpresionRetencion=1 then 1 else 0 end  as codigo FROM vTipodocumento WHERE Codigo=(SELECT TTIPODOCUMENTO FROM MNOTACREDITO WHERE tNotaCredito='" & sNC & "')", Cn) = 1 Then
                        Reporte.ReportComments = tTextoAgenteRetencion
                        End If
                        
                        xMontoTexto = "SON: " & NumeroCadena(str(RsImpDocumentoE!nVenta)) & " " & sMonedaN
                        Reporte.Text4.SetText xMontoTexto
                        Reporte.Text31.SetText xImpresionFE
    
'                    frmEmite.CRViewer.DisplayGroupTree = False
'                    frmEmite.CRViewer.ReportSource = Reporte
'                    frmEmite.CRViewer.ViewReport
'                    frmEmite.Show vbModal
    
                        Reporte.SelectPrinter Printer.DriverName, Printer.DeviceName, Printer.Port
                        Reporte.PaperOrientation = crPortrait
                        Reporte.PrintOut False, 1, False, 1, 1
                        '----------------
                        
                        
                    Else
                    
                        Dim Reporte1 As New dsrNotaCreditoObservacion
                        
                        Reporte1.DiscardSavedData
                        Reporte1.Database.SetDataSource RsImpDocumentoE
                        
                                         
                        Reporte1.Text13.SetText "NOTA DE CREDITO ELECTRONICA"
                        
                        Reporte1.Text8.SetText sRazonSocial
                        Reporte1.ReportTitle = sDireccion
                        Reporte1.Text15.SetText sTelefono
                        Reporte1.Text14.SetText sFax
                        Reporte1.Text16.SetText sRUC
                        
                        Reporte1.Text29.SetText xVenta
                        Reporte1.Text36.SetText xNeto
                        Reporte1.Text38.SetText xImp1
                        Reporte1.Text45.SetText xImp2
                        Reporte1.Text49.SetText xVenta
                        Reporte1.Text50.SetText sWeb
                        
                        If Calcular(" SELECT case when  lImpresionRetencion=1 then 1 else 0 end  as codigo FROM vTipodocumento WHERE Codigo=(SELECT TTIPODOCUMENTO FROM MNOTACREDITO WHERE tNotaCredito='" & sNC & "')", Cn) = 1 Then
                        Reporte1.ReportComments = tTextoAgenteRetencion
                        End If
                        
                        xMontoTexto = "SON: " & NumeroCadena(str(xVenta)) & " " & sMonedaN
                        Reporte1.Text4.SetText xMontoTexto
                        Reporte1.Text31.SetText xImpresionFE
    
'                        frmEmite.CRViewer.DisplayGroupTree = False
'                        frmEmite.CRViewer.ReportSource = Reporte1
'                        frmEmite.CRViewer.ViewReport
'                        frmEmite.Show vbModal

                        Reporte1.SelectPrinter Printer.DriverName, Printer.DeviceName, Printer.Port
                        Reporte1.PaperOrientation = crPortrait
                        Reporte1.PrintOut False, 1, False, 1, 1
                                       
                    End If
                                    
                    If iImagenCab Then
                       Kill App.Path & "\cliente.jpg"
                    End If
End Sub

Private Sub Form_Activate()
    Set RsTipoDocumento = Lib.OpenRecordset("select * from vTipoDocumento where Codigo <> '00'  and Canjear= 1 and lActivo = 1", Cn)
    If RsTipoDocumento.RecordCount = 0 Then
        MsgBox "No hay Tipos de Documentos con opción de Canje por Nota de Crédito", vbCritical
        Unload Me
    End If
        
End Sub

Private Sub Form_Load()
   Screen.MousePointer = vbHourglass
   Centrar Me
       
   'Ingrese el SubTitulo
   Me.Caption = " Mantenimiento de Notas de Crédito "
   fraDetalle.Caption = Me.Caption
      
      If lactivaFechaNC Then
        dtpFecha.Enabled = False
      Else
        dtpFecha.Enabled = True
      End If
      
    'FACTURACION E
    If pais = "002" Then 'Ecuador
      Set RsNotaCredito = Lib.OpenRecordset("select * from vTipoDocumentoImpresora where tCaja ='" & sCaja & "' and Transporte=0 and isnull(tNumeroAutorizacion,'')<>'' And lNotaCredito = 1 And lActivo = 1 UNION Select * From vTipoDocumentoImpresora where tCaja ='" & sCaja & "' and Transporte=0 And lNotaCredito = 1 And lFacturacionElectronica=1 and lActivo =1 order by tTipoEmision", Cn)
    Else
      Set RsNotaCredito = Lib.OpenRecordset("select * from vTipoDocumentoImpresora where tCaja ='" & sCaja & "' and Transporte=0 And lNotaCredito = 1 And lActivo = 1 order by tTipoEmision", Cn)
    End If
    
    sImp = RsNotaCredito!timpresora
    nEmision = RsNotaCredito.RecordCount
    
'    If nEmision = 0 Then
'      MsgBox "No se ha ingresado los Documentos de Nota de Credito por Caja", vbCritical
'      Unload Me
'    End If

    AsignaComando 2, RsNotaCredito, cmdNotaCredito()
    '------
    
   
   If Sw = True Then
      ActivarBotones (False)
      Blanquear Me
      dtpFecha.value = FechaServidor()
      nNeto = 0
      nImpuesto1 = 0
      nImpuesto2 = 0
      nImpuesto3 = 0
      nTotal = 0
      
      nNCNeto = 0
      nNCImp1 = 0
      nNCImp2 = 0
      nNCImp3 = 0
      nNCTotal = 0
      
      txtNCNeto.Text = "0.00"
      txtNCImp1.Text = "0.00"
      txtNCImp2.Text = "0.00"
      txtNCImp3.Text = "0.00"
      txtNCTotal.Text = "0.00"
      
      txtDocNeto.Text = "0.00"
      txtDocImp1.Text = "0.00"
      txtDocImp2.Text = "0.00"
      txtDocImp3.Text = "0.00"
      txtDocTotal.Text = "0.00"
      
      txtResNeto.Text = "0.00"
      txtResImp1.Text = "0.00"
      txtResImp2.Text = "0.00"
      txtResImp3.Text = "0.00"
      txtResTotal.Text = "0.00"
      
      cmdOpcion(8).Enabled = False
      cmdOpcion(9).Enabled = False
      'txtPrefijo.Text = "F"
      
   Else
      ActivarBotones (True)
      Asignar
   End If
   
   If nPorcentaje1 = 0 Then
      Label(5).Visible = False
      txtNCImp1.Visible = False
      txtDocImp1.Visible = False
      txtResImp1.Visible = False
      cmdOpcion(13).Visible = False
   Else
      Label(5).Caption = sImpuesto1 & " : "
   End If
   
   If nPorcentaje2 = 0 Then
      Label(6).Visible = False
      txtNCImp2.Visible = False
      txtDocImp2.Visible = False
      txtResImp2.Visible = False
      cmdOpcion(14).Visible = False
   Else
      Label(6).Caption = sImpuesto2 & " : "
   End If
   
   If nPorcentaje3 = 0 Then
      Label(7).Visible = False
      txtNCImp3.Visible = False
      txtDocImp3.Visible = False
      txtResImp3.Visible = False
      cmdOpcion(15).Visible = False
   Else
      Label(7).Caption = sImpuesto3 & " : "
   End If

    If lNCElimina Then
     cmdOpcion(2).Enabled = False
    End If
    If lParcialNC Then
         Me.Frame.Enabled = False
         cmdOpcion(9).Enabled = False
    End If
    If lactivaFechaNC Then
      dtpFecha.Enabled = False
    Else
      dtpFecha.Enabled = True
    End If
    
  If modProcedimiento.pasa = True Then
   cmdTexto.Caption = "Registro " & IIf(frmNotaCredito.RsCabecera.RecordCount = 0, 0, frmNotaCredito.RsCabecera.AbsolutePosition) & " de " & frmNotaCredito.RsCabecera.RecordCount
  End If
   
   
   'cmdTexto.Caption = "Registro " & IIf(frmNotaCredito.RsCabecera.RecordCount = 0, 0, frmNotaCredito.RsCabecera.AbsolutePosition) & " de " & frmNotaCredito.RsCabecera.RecordCount
   Screen.MousePointer = vbDefault
   
End Sub

Private Sub Form_Unload(Cancel As Integer)
    'Cambia el Nombre del Formulario
    Set frmNotaCreditoDetalle = Nothing
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
    cmdOpcion(12).Enabled = Activa
End Sub


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


