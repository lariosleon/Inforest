VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Begin VB.Form frmPrograma 
   BorderStyle     =   1  'Fixed Single
   Caption         =   " Fecha de Entrega"
   ClientHeight    =   4800
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8460
   Icon            =   "frmPrograma.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4800
   ScaleWidth      =   8460
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame1 
      Caption         =   "Enviar Pedido a Producci�n"
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
      Left            =   5400
      TabIndex        =   11
      Top             =   2760
      Width           =   3015
      Begin VB.CommandButton cmdNumpad 
         Caption         =   "Sin tiempo autom�tico"
         Height          =   795
         Index           =   0
         Left            =   1920
         TabIndex        =   14
         Top             =   240
         Width           =   1005
      End
      Begin VB.CommandButton cmdNumpad 
         Caption         =   "Minutos Antes"
         Height          =   795
         Index           =   1
         Left            =   940
         TabIndex        =   13
         Top             =   240
         Width           =   885
      End
      Begin VB.TextBox txtminutosAntes 
         Alignment       =   2  'Center
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
         Height          =   375
         Left            =   120
         TabIndex        =   12
         Text            =   "0"
         Top             =   420
         Width           =   735
      End
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
      Height          =   675
      Index           =   9
      Left            =   5670
      Picture         =   "frmPrograma.frx":000C
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   4080
      Width           =   1365
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
      Height          =   675
      Index           =   8
      Left            =   7080
      Picture         =   "frmPrograma.frx":010E
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   4080
      Width           =   1365
   End
   Begin VB.CommandButton cmdOpcion 
      Caption         =   "Minuto Siguiente"
      Height          =   675
      Index           =   5
      Left            =   7080
      TabIndex        =   7
      Top             =   1980
      Width           =   1365
   End
   Begin VB.CommandButton cmdOpcion 
      Caption         =   "Minuto Anterior"
      Height          =   675
      Index           =   4
      Left            =   5670
      TabIndex        =   6
      Top             =   1980
      Width           =   1365
   End
   Begin VB.CommandButton cmdOpcion 
      Caption         =   "Hora Siguiente"
      Height          =   675
      Index           =   3
      Left            =   7080
      TabIndex        =   5
      Top             =   750
      Width           =   1365
   End
   Begin VB.CommandButton cmdOpcion 
      Caption         =   "Hora Anterior"
      Height          =   675
      Index           =   2
      Left            =   5670
      TabIndex        =   4
      Top             =   750
      Width           =   1365
   End
   Begin VB.CommandButton cmdOpcion 
      Caption         =   "Mes Siguiente"
      Height          =   675
      Index           =   1
      Left            =   7080
      TabIndex        =   3
      Top             =   30
      Width           =   1365
   End
   Begin VB.CommandButton cmdOpcion 
      Caption         =   "Mes Anterior"
      Height          =   675
      Index           =   0
      Left            =   5670
      TabIndex        =   2
      Top             =   30
      Width           =   1365
   End
   Begin MSComCtl2.MonthView Calendar 
      Height          =   4770
      Left            =   0
      TabIndex        =   10
      Top             =   0
      Width           =   5325
      _ExtentX        =   9393
      _ExtentY        =   8414
      _Version        =   393216
      ForeColor       =   -2147483630
      BackColor       =   -2147483633
      Appearance      =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      StartOfWeek     =   85917697
      CurrentDate     =   37474
   End
   Begin VB.Label txtHora 
      Alignment       =   2  'Center
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "10:20"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   405
      Left            =   5730
      TabIndex        =   9
      Top             =   1500
      Width           =   1275
   End
   Begin VB.Label Label 
      Caption         =   "Horas"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   7170
      TabIndex        =   8
      Top             =   1500
      Width           =   1185
   End
End
Attribute VB_Name = "frmPrograma"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim sHora As Integer
Dim sMinuto As Integer

Private Sub cmdNumPad_Click(Index As Integer)

Select Case Index

    Case Is = 0
        txtminutosAntes.Text = "0"

    Case Is = 1
        sTipo = "Numero"
               
        frmNumPad.Show vbModal
        If wEnter And Val(sDescrip) > 0 Then
            txtminutosAntes.Text = Val(sDescrip)
        End If
End Select

        
End Sub

Private Sub cmdOpcion_Click(Index As Integer)
   If sModulo = "ADICION" Then
      frmMozoUsuario.ReseteaTimer
   End If

    Select Case Index
           Case Is = 0 ' Mes Anterior
                If Calendar.Month = 1 Then
                   Calendar.Month = 12
                   Calendar.Year = Calendar.Year - 1
                Else
                   Calendar.Month = Calendar.Month - 1
                End If
           
           Case Is = 1 ' Mes Siguiente
                If Calendar.Month = 12 Then
                   Calendar.Month = 1
                   Calendar.Year = Calendar.Year + 1
                Else
                   Calendar.Month = Calendar.Month + 1
                End If
           
           Case Is = 2 ' Hora Anterior
                If sHora = 0 Then
                   sHora = 23
                Else
                   sHora = sHora - 1
                End If
                txtHora.Caption = Format(TimeSerial(sHora, sMinuto, 0), "HH:MM")
           
           Case Is = 3 ' Hora Siguiente
                If sHora = 23 Then
                   sHora = 0
                Else
                   sHora = sHora + 1
                End If
                txtHora.Caption = Format(TimeSerial(sHora, sMinuto, 0), "HH:MM")
           
           Case Is = 4 ' Minuto Anterior
                If sMinuto = 0 Then
                   sMinuto = 59
                Else
                   sMinuto = sMinuto - 1
                End If
                txtHora.Caption = Format(TimeSerial(sHora, sMinuto, 0), "HH:MM")
           
           Case Is = 5 ' Minuto Siguiente
                If sMinuto = 59 Then
                   sMinuto = 0
                Else
                   sMinuto = sMinuto + 1
                End If
                txtHora.Caption = Format(TimeSerial(sHora, sMinuto, 0), "HH:MM")
           
           Case Is = 9 ' Cancelar
                wEnter = False
                Unload Me
                
           Case Is = 8 ' Aceptar
                wEnter = True
                sCodigo = Calendar.value & " " & txtHora.Caption
                lMinutoEnvioAntes = txtminutosAntes.Text
                If Format(FechaServidor(), "yyyyMMdd HH:mm") > Format(sCodigo, "yyyyMMdd HH:MM") Then
                    MsgBox "No puede seleccionar una Fecha pasada", vbCritical
                    Exit Sub
                End If
                
                Unload Me
                           
    End Select
End Sub

Private Sub Form_Load()
   If sModulo = "ADICION" Then
      frmMozoUsuario.ReseteaTimer
   End If
   Centrar Me
   wEnter = False
   If sModulo = "ADICION" Then
        txtHora.Caption = IIf(frmCargoMozo.txtFechaProg.Caption = "", Format(Time, "HH:00"), Format(frmCargoMozo.txtFechaProg.Caption, "HH:MM"))
        sHora = Hour(txtHora.Caption)
        sMinuto = Minute(txtHora.Caption)
        Calendar.value = IIf(frmCargoMozo.txtFechaProg.Caption = "", Format(FechaServidor(), "short date"), Format(frmCargoMozo.txtFechaProg.Caption, "short date"))
        txtminutosAntes.Text = IIf(frmCargoMozo.txtEnvioAntes.Text = "", 0, frmCargoMozo.txtEnvioAntes.Text)
   Else
        txtHora.Caption = IIf(frmVenta.txtFechaProg.Caption = "", Format(Time, "HH:00"), Format(frmVenta.txtFechaProg.Caption, "HH:MM"))
        sHora = Hour(txtHora.Caption)
        sMinuto = Minute(txtHora.Caption)
        Calendar.value = IIf(frmVenta.txtFechaProg.Caption = "", Format(FechaServidor(), "short date"), Format(frmVenta.txtFechaProg.Caption, "short date"))
        txtminutosAntes.Text = IIf(frmVenta.txtEnvioAntes.Text = "", 0, frmVenta.txtEnvioAntes.Text)
   End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
   Set frmPrograma = Nothing
End Sub



