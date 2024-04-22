VERSION 5.00
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Begin VB.Form frmReciboEgresoDetalle 
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   4830
   ClientLeft      =   2520
   ClientTop       =   2640
   ClientWidth     =   9810
   Icon            =   "frmReciboEgresoDetalle.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4830
   ScaleWidth      =   9810
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
      Height          =   4005
      Left            =   1860
      TabIndex        =   13
      Top             =   30
      Width           =   7935
      Begin VB.CommandButton cmdOpcion 
         Caption         =   "NC"
         Height          =   525
         Index           =   8
         Left            =   3480
         TabIndex        =   35
         Top             =   2880
         Width           =   855
      End
      Begin VB.TextBox txtnotacredito 
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'None
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
         Left            =   1200
         Locked          =   -1  'True
         TabIndex        =   34
         TabStop         =   0   'False
         Top             =   3000
         Width           =   2175
      End
      Begin VB.CommandButton cmdOpcion 
         Caption         =   "Dolares"
         Height          =   615
         Index           =   7
         Left            =   4200
         Style           =   1  'Graphical
         TabIndex        =   30
         Top             =   600
         Width           =   1170
      End
      Begin VB.TextBox txtDescripcion 
         BorderStyle     =   0  'None
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
         Height          =   525
         Left            =   1200
         MaxLength       =   50
         TabIndex        =   22
         Text            =   " "
         Top             =   1680
         Width           =   4020
      End
      Begin VB.TextBox txtMonto 
         Alignment       =   1  'Right Justify
         BorderStyle     =   0  'None
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
         Left            =   1200
         MaxLength       =   15
         TabIndex        =   21
         Text            =   " 0.00"
         Top             =   840
         Width           =   1695
      End
      Begin VB.TextBox txtRecibo 
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   0  'None
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
         Left            =   1200
         Locked          =   -1  'True
         TabIndex        =   20
         TabStop         =   0   'False
         Top             =   240
         Width           =   1695
      End
      Begin VB.TextBox txtTurno 
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   0  'None
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
         Left            =   1200
         Locked          =   -1  'True
         TabIndex        =   19
         TabStop         =   0   'False
         Top             =   3540
         Width           =   2175
      End
      Begin VB.TextBox txtFecha 
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   0  'None
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
         Left            =   3720
         Locked          =   -1  'True
         TabIndex        =   18
         TabStop         =   0   'False
         Top             =   240
         Width           =   1635
      End
      Begin VB.TextBox txtAutorizacion 
         BorderStyle     =   0  'None
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
         Height          =   525
         Left            =   1200
         MaxLength       =   50
         TabIndex        =   17
         Text            =   " "
         Top             =   2280
         Width           =   4020
      End
      Begin VB.CommandButton cmdOpcion 
         Caption         =   "Descripción"
         Height          =   495
         Index           =   4
         Left            =   5280
         Picture         =   "frmReciboEgresoDetalle.frx":030A
         Style           =   1  'Graphical
         TabIndex        =   16
         Top             =   1680
         Width           =   1170
      End
      Begin VB.CommandButton cmdOpcion 
         Caption         =   "Soles"
         Height          =   615
         Index           =   6
         Left            =   3000
         Style           =   1  'Graphical
         TabIndex        =   15
         Top             =   600
         Width           =   1170
      End
      Begin VB.CommandButton cmdOpcion 
         Caption         =   "Autorización"
         Height          =   495
         Index           =   5
         Left            =   5280
         Picture         =   "frmReciboEgresoDetalle.frx":040C
         Style           =   1  'Graphical
         TabIndex        =   14
         Top             =   2280
         Width           =   1170
      End
      Begin MSDataListLib.DataCombo cboTipoEgreso 
         Height          =   315
         Left            =   1200
         TabIndex        =   32
         Top             =   1320
         Width           =   5235
         _ExtentX        =   9234
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
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Nota  Credito :"
         Height          =   195
         Left            =   45
         TabIndex        =   33
         Top             =   3000
         Width           =   1020
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Tipo Egreso :"
         Height          =   195
         Left            =   135
         TabIndex        =   31
         Top             =   1320
         Width           =   945
      End
      Begin VB.Label lblEstado 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Atendido"
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
         Left            =   6240
         TabIndex        =   29
         Top             =   480
         Width           =   1155
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Fecha :"
         Height          =   195
         Left            =   3120
         TabIndex        =   28
         Top             =   240
         Width           =   540
      End
      Begin VB.Label Label10 
         AutoSize        =   -1  'True
         Caption         =   "Descripción :"
         Height          =   195
         Left            =   150
         TabIndex        =   27
         Top             =   1680
         Width           =   930
      End
      Begin VB.Label Label14 
         AutoSize        =   -1  'True
         Caption         =   "Autorización :"
         Height          =   195
         Left            =   120
         TabIndex        =   26
         Top             =   2280
         Width           =   960
      End
      Begin VB.Label Label20 
         AutoSize        =   -1  'True
         Caption         =   "Monto :"
         Height          =   195
         Left            =   540
         TabIndex        =   25
         Top             =   840
         Width           =   540
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         Caption         =   "Recibo :"
         Height          =   195
         Left            =   480
         TabIndex        =   24
         Top             =   240
         Width           =   600
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Turno :"
         Height          =   195
         Left            =   570
         TabIndex        =   23
         Top             =   3540
         Width           =   510
      End
   End
   Begin VB.PictureBox Picture 
      Align           =   2  'Align Bottom
      Height          =   750
      Left            =   0
      ScaleHeight     =   690
      ScaleWidth      =   9750
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   4080
      Width           =   9810
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
         Left            =   8520
         Picture         =   "frmReciboEgresoDetalle.frx":050E
         Style           =   1  'Graphical
         TabIndex        =   11
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
         Left            =   7350
         Picture         =   "frmReciboEgresoDetalle.frx":0600
         Style           =   1  'Graphical
         TabIndex        =   10
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
         Left            =   5010
         Picture         =   "frmReciboEgresoDetalle.frx":0702
         Style           =   1  'Graphical
         TabIndex        =   9
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
         Left            =   6180
         Picture         =   "frmReciboEgresoDetalle.frx":0C34
         Style           =   1  'Graphical
         TabIndex        =   8
         Top             =   60
         Width           =   1170
      End
      Begin VB.PictureBox PicNavegacion 
         BackColor       =   &H80000004&
         Height          =   615
         Left            =   30
         ScaleHeight     =   555
         ScaleWidth      =   4830
         TabIndex        =   1
         Top             =   60
         Width           =   4890
         Begin VB.CommandButton cmdNavegar 
            Height          =   555
            Index           =   1
            Left            =   480
            Picture         =   "frmReciboEgresoDetalle.frx":1166
            Style           =   1  'Graphical
            TabIndex        =   7
            TabStop         =   0   'False
            Top             =   0
            Width           =   480
         End
         Begin VB.CommandButton cmdNavegar 
            Height          =   555
            Index           =   2
            Left            =   960
            Picture         =   "frmReciboEgresoDetalle.frx":16A8
            Style           =   1  'Graphical
            TabIndex        =   6
            TabStop         =   0   'False
            Top             =   0
            Width           =   480
         End
         Begin VB.CommandButton cmdNavegar 
            Height          =   555
            Index           =   0
            Left            =   0
            Picture         =   "frmReciboEgresoDetalle.frx":1BEA
            Style           =   1  'Graphical
            TabIndex        =   5
            TabStop         =   0   'False
            Top             =   0
            Width           =   480
         End
         Begin VB.CommandButton cmdNavegar 
            Height          =   555
            Index           =   5
            Left            =   4350
            Picture         =   "frmReciboEgresoDetalle.frx":212C
            Style           =   1  'Graphical
            TabIndex        =   4
            TabStop         =   0   'False
            Top             =   0
            Width           =   480
         End
         Begin VB.CommandButton cmdNavegar 
            Height          =   555
            Index           =   4
            Left            =   3870
            Picture         =   "frmReciboEgresoDetalle.frx":266E
            Style           =   1  'Graphical
            TabIndex        =   3
            TabStop         =   0   'False
            Top             =   0
            Width           =   480
         End
         Begin VB.CommandButton cmdNavegar 
            Height          =   555
            Index           =   3
            Left            =   3390
            Picture         =   "frmReciboEgresoDetalle.frx":2BB0
            Style           =   1  'Graphical
            TabIndex        =   2
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
            TabIndex        =   12
            Top             =   150
            Width           =   1845
         End
      End
   End
   Begin VB.Image Image 
      Height          =   4005
      Left            =   45
      Picture         =   "frmReciboEgresoDetalle.frx":30F2
      Stretch         =   -1  'True
      Top             =   30
      Width           =   1785
   End
End
Attribute VB_Name = "frmReciboEgresoDetalle"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim sEstado As String
Dim nMonto As Double
Dim RsEgreso As Recordset
Dim vwNC As Boolean
Dim rsTipoEgreso As New ADODB.Recordset
Public Sub LlenaCombo()
        With cboTipoEgreso
             Isql = "select Codigo, tresumido from vTipoEgreso where lActivo=1 "
             Set rsTipoEgreso = Lib.OpenRecordset(Isql, Cn)
             Set .RowSource = rsTipoEgreso
                 .DataField = "tResumido"
                 .ListField = "tResumido"
                 .BoundColumn = "Codigo"
        End With
End Sub


Sub Asignar()
    With frmReciboEgreso.RsCabecera
        'Cuadro de Texto
        txtRecibo = IIf(IsNull(!tRecibo), "", !tRecibo)
        txtFecha = IIf(IsNull(!fFecha), "", !fFecha)
        nMonto = IIf(IsNull(!nMonto), 0, !nMonto)
        txtMonto.Text = Format(nMonto, "###,###,##0.00")
        txtDescripcion = IIf(IsNull(!tDescripcion), "", !tDescripcion)
        txtAutorizacion = IIf(IsNull(!tAutoriza), "", !tAutoriza)
        txtTurno = IIf(IsNull(!tTurno), "", !tTurno)
        cboTipoEgreso.BoundText = IIf(IsNull(!ttipoegreso), "", !ttipoegreso)
        cmdOpcion(6).FontBold = IIf(IsNull(!tMoneda), True, IIf(!tMoneda = "01", True, False))
        cmdOpcion(7).FontBold = IIf(IsNull(!tMoneda), False, IIf(!tMoneda = "01", False, True))
                
        lblEstado.ForeColor = IIf(IsNull(!tEstadoDocumento), vbBlack, IIf(!tEstadoDocumento = "04", vbRed, vbBlack))
        sEstado = IIf(IsNull(!tEstadoDocumento), "01", !tEstadoDocumento)
        lblEstado.Caption = IIf(IsNull(!tEstadoDocumento), "", IIf(!tEstadoDocumento = "04", "ANULADO", "EMITIDO"))
        txtnotacredito.Text = IIf(IsNull(!dReferencia), "", !dReferencia)
        'If Periodo(!fFecha) Or sEstado = "04" Or sEstado = "02" Or sTurno <> !tTurno Then
        If Periodo(!fFecha) Or sEstado = "04" Or sEstado = "02" Then
           Bloquear True
        Else
           Bloquear False
        End If
         If sEstado = "01" Then
            cmdOpcion(1).Enabled = False
        End If
    End With
    cmdTexto.Caption = "Registro " & frmReciboEgreso.RsCabecera.AbsolutePosition & " de " & frmReciboEgreso.RsCabecera.RecordCount
End Sub

Private Sub cmdNavegar_Click(Index As Integer)
    Select Case Index
           Case Is = 0 'Primero
                MoverPuntero Primero, frmReciboEgreso.grdGrilla
           Case Is = 1 'PgUp
                MoverPuntero pgup, frmReciboEgreso.grdGrilla
           Case Is = 2 'Previo
                MoverPuntero previo, frmReciboEgreso.grdGrilla
           Case Is = 3 'Siguiente
                MoverPuntero siguiente, frmReciboEgreso.grdGrilla
           Case Is = 4 'PgDn
                MoverPuntero pgdn, frmReciboEgreso.grdGrilla
           Case Is = 5 'Ultimo
                MoverPuntero Ultimo, frmReciboEgreso.grdGrilla
    End Select
    Asignar

End Sub

Private Sub cmdOpcion_Click(Index As Integer)
   Select Case Index
          Case Is = 0 'Agregar
               vwNC = False
               Sw = True
               ActivarBotones (False)
               Blanquear Me
               
               'Cambia el Nombre del Primer Text
               txtTurno.Text = sTurno
               txtFecha.Text = Date
               nMonto = 0
               txtMonto.Text = "0.00"
               lblEstado.Caption = ""
               Bloquear False
               cboTipoEgreso.Text = ""
               cmdOpcion(1).Enabled = True
                
          Case Is = 1 'Grabar
               'Chequea Datos
               Dim nCorrela As String
               Dim nPos As Integer
               If txtMonto.Text = "" Then MsgBox "Ingrese el Monto del Recibo", vbExclamation, sMensaje: txtMonto.SetFocus: Exit Sub
               If txtDescripcion.Text = "" Then MsgBox "Ingrese la Descripción del Recibo", vbExclamation, sMensaje: Exit Sub
              
               If vwNC Then
               Else
                If cboTipoEgreso.Text = "" Then MsgBox "Seleccione un Tipo de Egreso", vbExclamation, sMensaje: Exit Sub
               End If
               
               If Sw Then
               
                  'Obtiene el Numero de Orden
                  nCorrela = Calcular("select max(tRecibo) as Codigo from MEGRESO where substring(tRecibo,1,2)= substring(ltrim(str(year(getdate()))),3,2)", Cn)
                  If IsNull(nCorrela) Or Mid(nCorrela, 1, 2) <> Mid(Trim(str(Year(Date))), 3, 2) Then
                      txtRecibo.Text = Mid(Trim(str(Year(Date))), 3, 2) & "00000001"
                  Else
                      txtRecibo.Text = Mid(Trim(str(Year(Date))), 3, 2) & Lib.Correlativo(Mid(nCorrela, 3, 8), 8)
                  End If
                  Sw = False
                   
                   If vwNC = False Then
                    'Cambiar el SQL
                    Isql = "insert into MEGRESO( " & _
                           "tRecibo, tCaja, tTurno, fFecha, tMoneda, nTipoCambio, nMonto, tDescripcion, tAutoriza, tEstadoDocumento, tUsuario, fRegistro,fDiaContable,ttipoEgreso,dReferencia) " & _
                           "values ('" & txtRecibo.Text & "', " & _
                                  " '" & sCaja & "', " & _
                                  " '" & txtTurno.Text & "', " & _
                                  "{fn curdate()}, " & _
                                  IIf(cmdOpcion(6).FontBold = True, "'01'", "'02'") & ", " & _
                                  nTC & ", " & _
                                  nMonto & ", " & _
                                  " '" & txtDescripcion.Text & "', " & _
                                  " '" & txtAutorizacion.Text & "', " & _
                                  "'01', " & _
                                  "'" & sUsuario & "'," & _
                                  "getdate(),'" & Format(obtieneDiaContable, "yyyyMMdd") & "','" & cboTipoEgreso.BoundText & "','' )"
                    Else
                     'Cambiar el SQL
                    Isql = "insert into MEGRESO( " & _
                           "tRecibo, tCaja, tTurno, fFecha, tMoneda, nTipoCambio, nMonto, tDescripcion, tAutoriza, tEstadoDocumento, tUsuario, fRegistro,fDiaContable,ttipoEgreso,dReferencia) " & _
                           "values ('" & txtRecibo.Text & "', " & _
                                  " '" & sCaja & "', " & _
                                  " '" & txtTurno.Text & "', " & _
                                  "{fn curdate()}, " & _
                                  "'01'" & ", " & _
                                  nTC & ", " & _
                                  nMonto & ", " & _
                                  " '" & txtDescripcion.Text & "', " & _
                                  " '" & txtAutorizacion.Text & "', " & _
                                  "'01', " & _
                                  "'" & sUsuario & "'," & _
                                  "getdate(),'" & Format(obtieneDiaContable, "yyyyMMdd") & "','','" & txtnotacredito.Text & "' )"
                    
                                  
                    End If
                       Cn.Execute Isql
                    
                    If vwNC Then
                        Cn.Execute "update mnotacredito set testadodocumento='02' where tNotaCredito='" & Trim(txtnotacredito.Text) & "'"
                    End If
                                              
                       frmReciboEgreso.RsCabecera.Requery
                       frmReciboEgreso.RsCabecera.Find "tRecibo ='" & txtRecibo.Text & "'"
                       'RsDetalle.MoveLast
                       MsgBox "Registro Guardado", vbInformation, sMensaje
                       ActivarBotones (True)
                       cmdTexto.Caption = "Registro " & IIf(frmReciboEgreso.RsCabecera.RecordCount = 0, 0, frmReciboEgreso.RsCabecera.AbsolutePosition) & " de " & frmReciboEgreso.RsCabecera.RecordCount
                       
                       
                       Isql = "select tRecibo, tTurno, fFecha, tMoneda, nMonto, tDescripcion, tAutoriza, tEstadoDocumento, tUsuario, fRegistro, Estado, Moneda, TipoEgreso,ttipoegreso, str(nMonto,10,2) as xMonto, convert(varchar, fFecha, 103) as xFecha from vEgreso " & _
                               "where tRecibo = '" & txtRecibo.Text & "'"
                               
                       Set RsEgreso = Lib.OpenRecordset(Isql, Cn)
                                        
                       ImprimeReciboEgreso RsEgreso
                       Unload Me
               Else
               
                    'JESUS
                    If frmReciboEgreso.RsCabecera!tTurno = sTurno Then
                       'Password

                    Else
                       'Password
                       If Supervisor("06") = False Then
                          MsgBox "Clave no permitida", vbExclamation, sMensaje
                          Exit Sub
                       End If
                    End If
                    '----------Fin 24.10.2017---------
                    
                    
                  'Cambiar el SQL
                  If vwNC Then
                    Isql = "update MEGRESO set " & _
                           "tMoneda =" & IIf(cmdOpcion(6).FontBold = True, "'01'", "'02'") & ", " & _
                           "nMonto =" & nMonto & ", " & _
                           "tDescripcion ='" & txtDescripcion.Text & "', " & _
                           "tTipoEgreso='" & cboTipoEgreso.BoundText & "', " & _
                           "tAutoriza ='" & txtAutorizacion.Text & "', lReplica=1 " & _
                           " where tRecibo = '" & txtRecibo & "'"
                         
                    Else
                        Isql = "update MEGRESO set " & _
                                "tMoneda =" & IIf(cmdOpcion(6).FontBold = True, "'01'", "'02'") & ", " & _
                                "nMonto =" & nMonto & ", " & _
                                "tDescripcion ='" & txtDescripcion.Text & "', " & _
                                "tTipoEgreso='" & cboTipoEgreso.BoundText & "', " & _
                                "tAutoriza ='" & txtAutorizacion.Text & "', lReplica=1 " & _
                                " where tRecibo = '" & txtRecibo & "'"
                    End If
                       
                   Cn.Execute Isql
                   
                   'Cambiar el Nombre del Formulario
                   nPos = frmReciboEgreso.RsCabecera.AbsolutePosition
                   frmReciboEgreso.RsCabecera.Requery
                   frmReciboEgreso.RsCabecera.AbsolutePosition = nPos
                   MsgBox "Registro Modificado", vbInformation, sMensaje
               End If
               'ImprimeReciboEgreso frmReciboEgreso.RsCabecera
          
          Case Is = 2 'Eliminar
               If frmReciboEgreso.RsCabecera.RecordCount = 0 Then
                  Exit Sub
               End If
               
               If frmReciboEgreso.RsCabecera!tTurno = sTurno Then
                  'Password
                  If Supervisor("07") = False Then
                     MsgBox "Clave no permitida", vbExclamation, sMensaje
                     Exit Sub
                  End If
               Else
                  'Password
                  If Supervisor("06") = False Then
                     MsgBox "Clave no permitida", vbExclamation, sMensaje
                     Exit Sub
                  End If
               End If
                              
               'Cambia el MsgBox
               If MsgBox("Seguro de Eliminar el Recibo Nro." & txtRecibo & "?", vbQuestion + vbOKCancel, sMensaje) = vbCancel Then
                  Exit Sub
               End If
               
               'Cambia el Delete
               Cn.Execute "update MEGRESO set tEstadoDocumento = '04' where tRecibo = '" & txtRecibo & "'"
               nPos = frmReciboEgreso.RsCabecera.AbsolutePosition
               frmReciboEgreso.RsCabecera.Requery
               frmReciboEgreso.RsCabecera.AbsolutePosition = nPos
               Asignar
               
          Case Is = 3 'Salir
               Unload Me
          
          Case Is = 4 'KB Descripcion
                frmKeyBoard.txtResultado.Text = txtDescripcion.Text
                frmKeyBoard.Show vbModal
                txtDescripcion.Text = IIf(wEnter, sDescrip, txtDescripcion.Text)
               
          Case Is = 5 'KB Autorizacion
                frmKeyBoard.txtResultado.Text = txtAutorizacion.Text
                frmKeyBoard.Show vbModal
                txtAutorizacion.Text = IIf(wEnter, sDescrip, txtAutorizacion.Text)
          
          Case Is = 6 'Numpad Soles
               frmNumPad.Show vbModal
               nMonto = IIf(wEnter, Val(sDescrip), nMonto)
               txtMonto.Text = Format(nMonto, "###,##0.00")
               cmdOpcion(6).FontBold = True
               cmdOpcion(7).FontBold = False
          
          Case Is = 7 'Numpad Dolares
               frmNumPad.Show vbModal
               nMonto = IIf(wEnter, Val(sDescrip), nMonto)
               txtMonto.Text = Format(nMonto, "###,##0.00")
               cmdOpcion(7).FontBold = True
               cmdOpcion(6).FontBold = False
               
         Case Is = 8 'NOTACREDITO
               'frmNumPad.Show vbModal
'               nMonto = IIf(wEnter, Val(sDescrip), nMonto)
'               txtMonto.Text = Format(nMonto, "###,##0.00")
          
               vwNC = True
               Isql = "select *, tNotaCredito as Codigo, nVenta as Descripcion from vNotaCredito where tEstadodocumento ='05'"
               Call ConfGrilla(4, frmBusquedaRapida.grdGrilla, "Nota Credito", 2, "tNotaCredito", 1900, 2, 0, "", _
                                                               "Fecha", 2, "fFecha", 1200, 0, 0, "dd/mm/yyyy", _
                                                               "Cliente", 2, "Cliente", 3900, 0, 0, "", _
                                                               "Monto", 2, "nVenta", 1200, 1, 0, "##,##0.00")
               frmBusquedaRapida.nPredeterm = 1
               frmBusquedaRapida.Show vbModal
               If wEnter Then
                  txtnotacredito.Text = sCodigo
'                  stipodocumento = sCodigo
'                  txtReferencia.Caption = sCodigo
'                  sOtroTipoCancelacion = "002"
'                  nAbonoN = nAbonoN - nDocumento
                    txtMonto.Text = Format(CDbl(sDescrip), "###,##0.00")
                    nMonto = CDbl(sDescrip)
                    'txtMonto.Text = Val(sDescrip)
'                  sOtroTipo = sTipo
                txtMonto.Enabled = False
                cmdOpcion(7).Enabled = False
                cmdOpcion(6).Enabled = False
'                cmdOpcion(6).FontBold = True
'                cmdOpcion(7).FontBold = False
               Else
                  Exit Sub
               
               End If
               
'               txtDocMonto.Caption = Format(nDocumento, "###,###,###,##0.00")
'               nAbonoN = nAbonoN + nDocumento
'               CalculaTotal
   End Select
End Sub

Private Sub Form_Load()
   Screen.MousePointer = vbHourglass
   Centrar Me
   LlenaCombo
   If sMonE = "" Or sMonN = sMonE Then
      cmdOpcion(6).Caption = sMonedaN
      cmdOpcion(7).Visible = False
   Else
      cmdOpcion(6).Caption = sMonedaN
      cmdOpcion(7).Caption = sMonedaE
   End If
   
   cmdOpcion(6).FontBold = True
    
    'Ingrese el SubTitulo
    Me.Caption = " Mantenimiento de Recibos de Egreso "
    fraDetalle.Caption = Me.Caption
    
    'CESAR PERMISOS DE GRUPO
    Dim xGrupo As String
    Dim xRGrupoAcceso As Recordset
    xGrupo = Calcular("Select tGrupoUsuario As Codigo From TUSUARIO where tResumido = '" & sUsuario & "'", Cn)
        
        
    If Sw = True Then
       ActivarBotones (False)
       Blanquear Me
       txtTurno.Text = sTurno
       txtFecha.Text = Date
       txtMonto.Text = "0.00"
       lblEstado.Caption = ""
       cmdOpcion(1).Enabled = True
    Else
       Asignar
       ActivarBotones (True)
       
       
           If xGrupo <> "00" Then
            Isql = "Select * From TGRUPOACCESO WHERE tGrupoUsuario = '" & xGrupo & "'"
            Set xRGrupoAcceso = Lib.OpenRecordset(Isql, Cn)
            '30204010 MODIFICAR
                xRGrupoAcceso.Filter = "tCodigoAcceso=30204010"
                If xRGrupoAcceso.RecordCount > 0 Then
                   cmdOpcion(1).Enabled = False
                Else
                   cmdOpcion(1).Enabled = True
                End If
            End If
       
    End If
    
    cmdTexto.Caption = "Registro " & IIf(frmReciboEgreso.RsCabecera.RecordCount = 0, 0, frmReciboEgreso.RsCabecera.AbsolutePosition) & " de " & frmReciboEgreso.RsCabecera.RecordCount
    Screen.MousePointer = vbDefault

    
End Sub

Private Sub Form_Unload(Cancel As Integer)
    'Cambia el Nombre del Formulario
    Set frmReciboEgresoDetalle = Nothing
End Sub

Sub ActivarBotones(ByVal Activa As Boolean)
    cmdNavegar(0).Enabled = Activa
    cmdNavegar(1).Enabled = Activa
    cmdNavegar(2).Enabled = Activa
    cmdNavegar(3).Enabled = Activa
    cmdNavegar(4).Enabled = Activa
    cmdNavegar(5).Enabled = Activa
End Sub

Public Sub Bloquear(Activo As Boolean)
   cmdOpcion(1).Enabled = Not Activo
   cmdOpcion(2).Enabled = Not Activo
   cmdOpcion(4).Enabled = Not Activo
   cmdOpcion(5).Enabled = Not Activo
   cmdOpcion(6).Enabled = Not Activo
   cmdOpcion(7).Enabled = Not Activo
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
End Function
'diaContable

