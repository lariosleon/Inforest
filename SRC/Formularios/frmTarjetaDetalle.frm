VERSION 5.00
Begin VB.Form frmTarjetaDetalle 
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   4260
   ClientLeft      =   2520
   ClientTop       =   2640
   ClientWidth     =   10215
   Icon            =   "frmTarjetaDetalle.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4260
   ScaleWidth      =   10215
   Begin VB.Frame fraBoton 
      Caption         =   " Botonera "
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3435
      Left            =   8820
      TabIndex        =   22
      Top             =   0
      Width           =   1365
      Begin VB.CommandButton cmdBoton 
         Caption         =   "1"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   510
         Index           =   1
         Left            =   120
         Style           =   1  'Graphical
         TabIndex        =   36
         Top             =   360
         Width           =   510
      End
      Begin VB.CommandButton cmdBoton 
         Caption         =   "2"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   510
         Index           =   2
         Left            =   720
         Style           =   1  'Graphical
         TabIndex        =   35
         Top             =   360
         Width           =   510
      End
      Begin VB.CommandButton cmdBoton 
         Caption         =   "3"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   510
         Index           =   3
         Left            =   120
         Style           =   1  'Graphical
         TabIndex        =   34
         Top             =   975
         Width           =   510
      End
      Begin VB.CommandButton cmdBoton 
         Caption         =   "4"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   510
         Index           =   4
         Left            =   720
         Style           =   1  'Graphical
         TabIndex        =   33
         Top             =   975
         Width           =   510
      End
      Begin VB.CommandButton cmdBoton 
         Caption         =   "5"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   510
         Index           =   5
         Left            =   120
         Style           =   1  'Graphical
         TabIndex        =   32
         Top             =   1575
         Width           =   510
      End
      Begin VB.CommandButton cmdBoton 
         Caption         =   "6"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   510
         Index           =   6
         Left            =   720
         Style           =   1  'Graphical
         TabIndex        =   31
         Top             =   1575
         Width           =   510
      End
      Begin VB.CommandButton cmdBoton 
         Caption         =   "7"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   510
         Index           =   7
         Left            =   120
         Style           =   1  'Graphical
         TabIndex        =   30
         Top             =   2190
         Width           =   510
      End
      Begin VB.CommandButton cmdBoton 
         Caption         =   "8"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   510
         Index           =   8
         Left            =   720
         Style           =   1  'Graphical
         TabIndex        =   29
         Top             =   2190
         Width           =   510
      End
      Begin VB.Label Label11 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Boton :"
         Height          =   195
         Left            =   90
         TabIndex        =   38
         Top             =   2955
         Width           =   510
      End
      Begin VB.Label txtBoton 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
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
         Height          =   285
         Left            =   675
         TabIndex        =   37
         Top             =   2910
         Width           =   540
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
      Height          =   3450
      Left            =   1755
      TabIndex        =   21
      Top             =   0
      Width           =   7035
      Begin VB.TextBox txtCodTarEx 
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
         Left            =   1875
         MaxLength       =   15
         TabIndex        =   41
         Text            =   " "
         Top             =   2760
         Width           =   1230
      End
      Begin VB.CheckBox chkPinPad 
         Alignment       =   1  'Right Justify
         Caption         =   "Utiliza POS :"
         Height          =   195
         Left            =   4440
         TabIndex        =   7
         Top             =   360
         Width           =   1200
      End
      Begin VB.TextBox txtCuentaContable 
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
         Left            =   1875
         MaxLength       =   15
         TabIndex        =   6
         Text            =   " "
         Top             =   2080
         Width           =   2550
      End
      Begin VB.CheckBox chkActivo 
         Alignment       =   1  'Right Justify
         Caption         =   "Activo :"
         Height          =   195
         Left            =   6000
         TabIndex        =   8
         Top             =   360
         Width           =   885
      End
      Begin VB.TextBox txtTelefono1 
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
         Left            =   1875
         MaxLength       =   15
         TabIndex        =   5
         Text            =   " "
         Top             =   2436
         Width           =   2550
      End
      Begin VB.TextBox txtCodigo 
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
         Left            =   1875
         Locked          =   -1  'True
         TabIndex        =   0
         TabStop         =   0   'False
         Top             =   300
         Width           =   1215
      End
      Begin VB.TextBox txtResumido 
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
         Left            =   1875
         MaxLength       =   30
         TabIndex        =   2
         Text            =   " "
         Top             =   1012
         Width           =   2550
      End
      Begin VB.TextBox txtDetallado 
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
         Left            =   1875
         MaxLength       =   50
         TabIndex        =   1
         Text            =   " "
         Top             =   656
         Width           =   5025
      End
      Begin VB.TextBox txtFactor 
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
         Left            =   1875
         MaxLength       =   15
         TabIndex        =   4
         Text            =   " 0.00"
         Top             =   1724
         Width           =   1215
      End
      Begin VB.TextBox txtRepresentante 
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
         Left            =   1875
         MaxLength       =   40
         TabIndex        =   3
         Text            =   " "
         Top             =   1368
         Width           =   2550
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Cod ApiWeb :"
         Height          =   195
         Left            =   760
         TabIndex        =   42
         Top             =   2800
         Width           =   990
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Cuenta Contable :"
         Height          =   195
         Left            =   495
         TabIndex        =   40
         Top             =   2125
         Width           =   1275
      End
      Begin VB.Label Label10 
         AutoSize        =   -1  'True
         Caption         =   "Descripción Detallada :"
         Height          =   195
         Left            =   120
         TabIndex        =   28
         Top             =   701
         Width           =   1650
      End
      Begin VB.Label Label14 
         AutoSize        =   -1  'True
         Caption         =   "Descripción Resumida :"
         Height          =   195
         Left            =   90
         TabIndex        =   27
         Top             =   1057
         Width           =   1680
      End
      Begin VB.Label Label20 
         AutoSize        =   -1  'True
         Caption         =   "Factor Retención :"
         Height          =   195
         Left            =   450
         TabIndex        =   26
         Top             =   1769
         Width           =   1320
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         Caption         =   "Código :"
         Height          =   195
         Left            =   1185
         TabIndex        =   25
         Top             =   345
         Width           =   585
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Cta. Contable Comisión :"
         Height          =   195
         Left            =   45
         TabIndex        =   24
         Top             =   2481
         Width           =   1725
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "Representante :"
         Height          =   195
         Left            =   630
         TabIndex        =   23
         Top             =   1413
         Width           =   1140
      End
   End
   Begin VB.PictureBox Picture 
      Align           =   2  'Align Bottom
      Height          =   750
      Left            =   0
      ScaleHeight     =   690
      ScaleWidth      =   10155
      TabIndex        =   13
      TabStop         =   0   'False
      Top             =   3510
      Width           =   10215
      Begin VB.PictureBox PicNavegacion 
         BackColor       =   &H80000004&
         Height          =   615
         Left            =   30
         ScaleHeight     =   555
         ScaleWidth      =   5250
         TabIndex        =   14
         Top             =   90
         Width           =   5310
         Begin VB.CommandButton cmdNavegar 
            Height          =   555
            Index           =   3
            Left            =   3810
            Picture         =   "frmTarjetaDetalle.frx":0442
            Style           =   1  'Graphical
            TabIndex        =   20
            TabStop         =   0   'False
            Top             =   0
            Width           =   480
         End
         Begin VB.CommandButton cmdNavegar 
            Height          =   555
            Index           =   4
            Left            =   4290
            Picture         =   "frmTarjetaDetalle.frx":0984
            Style           =   1  'Graphical
            TabIndex        =   19
            TabStop         =   0   'False
            Top             =   0
            Width           =   480
         End
         Begin VB.CommandButton cmdNavegar 
            Height          =   555
            Index           =   5
            Left            =   4770
            Picture         =   "frmTarjetaDetalle.frx":0EC6
            Style           =   1  'Graphical
            TabIndex        =   18
            TabStop         =   0   'False
            Top             =   0
            Width           =   480
         End
         Begin VB.CommandButton cmdNavegar 
            Height          =   555
            Index           =   0
            Left            =   0
            Picture         =   "frmTarjetaDetalle.frx":1408
            Style           =   1  'Graphical
            TabIndex        =   17
            TabStop         =   0   'False
            Top             =   0
            Width           =   480
         End
         Begin VB.CommandButton cmdNavegar 
            Height          =   555
            Index           =   2
            Left            =   960
            Picture         =   "frmTarjetaDetalle.frx":194A
            Style           =   1  'Graphical
            TabIndex        =   16
            TabStop         =   0   'False
            Top             =   0
            Width           =   480
         End
         Begin VB.CommandButton cmdNavegar 
            Height          =   555
            Index           =   1
            Left            =   480
            Picture         =   "frmTarjetaDetalle.frx":1E8C
            Style           =   1  'Graphical
            TabIndex        =   15
            TabStop         =   0   'False
            Top             =   0
            Width           =   480
         End
         Begin VB.Label cmdTexto 
            Alignment       =   2  'Center
            Caption         =   "Registro 0 de 0"
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
            Left            =   1470
            TabIndex        =   39
            Top             =   150
            Width           =   2295
         End
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
         Left            =   6600
         Picture         =   "frmTarjetaDetalle.frx":23CE
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
         Left            =   5430
         Picture         =   "frmTarjetaDetalle.frx":2900
         Style           =   1  'Graphical
         TabIndex        =   9
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
         Left            =   7770
         Picture         =   "frmTarjetaDetalle.frx":2E32
         Style           =   1  'Graphical
         TabIndex        =   11
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
         Left            =   8940
         Picture         =   "frmTarjetaDetalle.frx":2F34
         Style           =   1  'Graphical
         TabIndex        =   12
         Top             =   60
         Width           =   1170
      End
   End
   Begin VB.Image Image 
      Height          =   3375
      Left            =   60
      Picture         =   "frmTarjetaDetalle.frx":3026
      Stretch         =   -1  'True
      Top             =   45
      Width           =   1650
   End
End
Attribute VB_Name = "frmTarjetaDetalle"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim RsBoton As Recordset

Sub Asignar()
    With frmTarjeta.RsCabecera
        'Cuadro de Texto
        txtCodigo = IIf(IsNull(!tCodigoTarjeta), "", !tCodigoTarjeta)
        txtDetallado = IIf(IsNull(!tDetallado), "", !tDetallado)
        txtResumido = IIf(IsNull(!tResumido), "", !tResumido)
        txtFactor = Format(IIf(IsNull(!nFactorRetencion), "0.00", !nFactorRetencion), "###,##0.00")
        txtRepresentante = IIf(IsNull(!tRepresentante), "", !tRepresentante)
        txtTelefono1 = IIf(IsNull(!ttelefono1), "", !ttelefono1)
        txtBoton = IIf(IsNull(!nBoton), "", !nBoton)
        txtCuentaContable = IIf(IsNull(!tcuentaContable), "", !tcuentaContable)
        ' se agrego el campo de tCodTarjetaEx
        txtCodTarEx.Text = IIf(IsNull(!tCodTarjetaEx), "", !tCodTarjetaEx)
        
 
        'Check Box
        chkPinPad = IIf(!lPinPad = True, 1, 0)
        chkActivo = IIf(!lActivo = True, 1, 0)
        Botonera
        
    End With
End Sub
Private Sub cmdBoton_Click(Index As Integer)
   If Val(txtBoton) <> 0 Then
      cmdBoton(Val(txtBoton)).backColor = vbButtonFace
      cmdBoton(Val(txtBoton)).Enabled = True
   End If
   cmdBoton(Index).backColor = vbRed
   cmdBoton(Index).Enabled = False
   txtBoton.Caption = Index
End Sub

Private Sub cmdNavegar_Click(Index As Integer)
    Select Case Index
           Case Is = 0 'Primero
                MoverPuntero Primero, frmTarjeta.grdGrilla
           Case Is = 1 'PgUp
                MoverPuntero pgup, frmTarjeta.grdGrilla
           Case Is = 2 'Previo
                MoverPuntero previo, frmTarjeta.grdGrilla
           Case Is = 3 'Siguiente
                MoverPuntero siguiente, frmTarjeta.grdGrilla
           Case Is = 4 'PgDn
                MoverPuntero pgdn, frmTarjeta.grdGrilla
           Case Is = 5 'Ultimo
                MoverPuntero Ultimo, frmTarjeta.grdGrilla
    End Select
   Asignar
   cmdTexto.Caption = "Registro " & IIf(frmTarjeta.RsCabecera.RecordCount = 0, 0, frmTarjeta.RsCabecera.AbsolutePosition) & " de " & frmTarjeta.RsCabecera.RecordCount
End Sub

Private Sub cmdOpcion_Click(Index As Integer)
   Select Case Index
          Case Is = 0 ' Agregar
              ' If frmTarjeta.RsCabecera.RecordCount < 8 Then
                  Sw = True
                  ActivarBotones (False)
                  Blanquear Me
                  chkActivo.value = 1
                  chkPinPad.value = 0
                  txtFactor.Text = "0.00"
                  'Cambia el Nombre del Primer Text
                  txtDetallado.SetFocus
                  Botonera
'               Else
'                  MsgBox "Ha llegado al límite de 8 Tarjetas Bancarias", vbExclamation, sMensaje
'               End If
          
          Case Is = 1 ' Grabar
               Dim nCorrela As String
               
               'Chequea Datos
               If txtDetallado.Text = "" Then MsgBox "Ingrese la Descripción Detallada", vbExclamation, sMensaje: txtDetallado.SetFocus: Exit Sub
               If txtResumido.Text = "" Then MsgBox "Ingrese la Descripción Resumida", vbExclamation, sMensaje: txtResumido.SetFocus: Exit Sub
                   
               If Sw Then
                  'Obtiene el Numero de Orden
                  nCorrela = Calcular("select max(tCodigoTarjeta) as Codigo from TTARJETACREDITO", Cn)
                  If IsNull(nCorrela) Or nCorrela = "" Then
                      txtCodigo.Text = "01"
                  Else
                      txtCodigo.Text = Lib.Correlativo(nCorrela, 2)
                  End If
                  Sw = False
                                   
                                   
                sPasa = txtCodigo.Text
                  
                'Inserta Movimiento auditoria
                lAuditoria = RegistraMovimientoAuditoria(tModuloSeg, sMDB, "TTARJETACREDITO", "TARJETA CREDITO", "01", sUsuario, sPasa, "", _
                   "TCODIGOTARJETA", "CODIGO TARJETA", sPasa, "tDetallado", "Descripcion Detallada", txtDetallado.Text, "tResumido", "Descripcion Resumida", txtResumido.Text, _
                   "tRepresentante", "Representante", txtRepresentante.Text, "nFactorRetencion", "Retencion", Val(txtFactor.Text), _
                   "tCuentaContable", "Cuenta Contable", txtCuentaContable.Text, "tTelefono1", "Cuenta Contable Comision", Val(txtTelefono1.Text), _
                   "lPinPad", "Flag Pin Pad", IIf(chkPinPad.value, "Verdadero", "Falso"), "nBoton", "Botonera", Val(txtBoton), "lActivo", "Flag Activo", IIf(chkActivo.value, "Verdadero", "Falso"))
                
                If lAuditoria = False Then
                    Screen.MousePointer = vbDefault
                        Exit Sub
                End If
                
                'La Funcion RegistraMovimientoAuditoria devuelve true si se ejecuto correctamente.
                   
                                                     
                                                     
                                                     
                                   
                  'Cambiar el SQL
                  Isql = "insert into TTARJETACREDITO ( " & _
                         "tCodigoTarjeta, tDetallado, tResumido, nFactorRetencion, tRepresentante, " & _
                         "tTelefono1, nBoton, tCuentaContable, lPinPad, lActivo, tUsuario,tCodTarjetaEx,fRegistro) " & _
                         "values ('" & txtCodigo.Text & "', " & _
                                " '" & txtDetallado.Text & "', " & _
                                " '" & txtResumido.Text & "', " & _
                                       Val(txtFactor.Text) & ", " & _
                                " '" & txtRepresentante.Text & "', " & _
                                " '" & txtTelefono1.Text & "', " & _
                                   Val(txtBoton.Caption) & ", " & _
                                " '" & txtCuentaContable.Text & "', " & _
                                       chkPinPad.value & ", " & _
                                       chkActivo.value & ", " & _
                                "'" & sUsuario & "'," & _
                                " '" & txtCodTarEx.Text & "', " & _
                                " getdate() )"
                                  
                  Cn.Execute Isql
                  

                  
                  
                  frmTarjeta.RsCabecera.Sort = "tCodigoTarjeta ASC"
                  frmTarjeta.RsCabecera.Requery
                  frmTarjeta.RsCabecera.MoveLast
                  MsgBox "Registro Guardado", vbInformation, sMensaje
                  ActivarBotones (True)
                  cmdTexto.Caption = "Registro " & frmTarjeta.RsCabecera.AbsolutePosition & " de " & frmTarjeta.RsCabecera.RecordCount
               Else
               
                    sPasa = txtCodigo.Text
                  
                'Inserta Movimiento auditoria
                lAuditoria = RegistraMovimientoAuditoria(tModuloSeg, sMDB, "TTARJETACREDITO", "TARJETA CREDITO", "02", sUsuario, sPasa, "", _
                   "TCODIGOTARJETA", "CODIGO TARJETA", sPasa, "tDetallado", "Descripcion Detallada", txtDetallado.Text, "tResumido", "Descripcion Resumida", txtResumido.Text, _
                   "tRepresentante", "Representante", txtRepresentante.Text, "nFactorRetencion", "Retencion", Val(txtFactor.Text), _
                   "tCuentaContable", "Cuenta Contable", txtCuentaContable.Text, "tTelefono1", "Cuenta Contable Comision", Val(txtTelefono1.Text), _
                   "lPinPad", "Flag Pin Pad", IIf(chkPinPad.value, "Verdadero", "Falso"), "nBoton", "Botonera", Val(txtBoton), "lActivo", "Flag Activo", IIf(chkActivo.value, "Verdadero", "Falso"))
                
                If lAuditoria = False Then
                    Screen.MousePointer = vbDefault
                        Exit Sub
                End If
                
                
                'La Funcion RegistraMovimientoAuditoria devuelve true si se ejecuto correctamente.
                   
               
                  'Cambiar el SQL
                  Isql = "update TTARJETACREDITO set " & _
                         "tDetallado ='" & txtDetallado.Text & "', " & _
                         "tResumido ='" & txtResumido.Text & "', " & _
                         "nFactorRetencion =" & Val(txtFactor.Text) & ", " & _
                         "tRepresentante ='" & txtRepresentante.Text & "', " & _
                         "tTelefono1 ='" & txtTelefono1.Text & "', " & _
                         "tCuentaContable ='" & txtCuentaContable.Text & "', " & _
                         "nBoton =" & Val(txtBoton.Caption) & ", " & _
                         "lPinPad =" & chkPinPad.value & ", " & _
                         "tCodTarjetaEx ='" & txtCodTarEx.Text & "', " & _
                         "lActivo =" & chkActivo.value & ", lReplica=1 " & _
                         " where tCodigoTarjeta = '" & txtCodigo & "'"
                       
                   Cn.Execute Isql
                   nPos = frmTarjeta.RsCabecera.Bookmark
                   frmTarjeta.RsCabecera.Requery
                   If frmTarjeta.RsCabecera.RecordCount = 0 Then
                      frmTarjeta.RsCabecera.Filter = adFilterNone
                   End If
                   frmTarjeta.RsCabecera.Bookmark = nPos
                   Screen.MousePointer = vbDefault
                   MsgBox "Registro Modificado", vbInformation, sMensaje
               End If
                              
               cmdTexto.Caption = "Registro " & IIf(frmTarjeta.RsCabecera.RecordCount = 0, 0, frmTarjeta.RsCabecera.AbsolutePosition) & " de " & frmTarjeta.RsCabecera.RecordCount
          
          Case Is = 2 ' Eliminar
               If frmTarjeta.RsCabecera.RecordCount = 0 Then
                  Exit Sub
               End If
               'Cambia el MsgBox
               If MsgBox("Seguro de Eliminar la Tarjeta " & txtCodigo & "?", vbQuestion + vbOKCancel, sMensaje) = vbCancel Then
                  Exit Sub
               End If
               
                    sPasa = txtCodigo.Text
                  
                'Inserta Movimiento auditoria
                lAuditoria = RegistraMovimientoAuditoria(tModuloSeg, sMDB, "TTARJETACREDITO", "TARJETA CREDITO", "03", sUsuario, sPasa, "", _
                   "TCODIGOTARJETA", "CODIGO TARJETA", sPasa, "tDetallado", "Descripcion Detallada", txtDetallado.Text)
                
                If lAuditoria = False Then
                    Screen.MousePointer = vbDefault
                        Exit Sub
                End If
                
                'La Funcion RegistraMovimientoAuditoria devuelve true si se ejecuto correctamente.
                   
               
               
               'Cambia el Delete
               Cn.Execute "delete from TTARJETACREDITO where tCodigoTarjeta = '" & txtCodigo & "'"
               frmTarjeta.RsCabecera.Requery
               If frmTarjeta.RsCabecera.RecordCount <> 0 Then
                  frmTarjeta.RsCabecera.MoveLast
                  Asignar
                  cmdTexto.Caption = "Registro " & IIf(frmTarjeta.RsCabecera.RecordCount = 0, 0, frmTarjeta.RsCabecera.AbsolutePosition) & " de " & frmTarjeta.RsCabecera.RecordCount
               Else
                  ActivarBotones False
                  Blanquear Me
                  Sw = True
               End If
          
          Case Is = 3 ' Salir
               Unload Me
   End Select

End Sub

Private Sub Form_Load()
    Screen.MousePointer = vbHourglass
    Centrar Me
    'Ingrese el SubTitulo
    Me.Caption = " Tarjetas Bancarias "
    fraDetalle.Caption = Me.Caption
       
    'Botones
    Isql = "select tCodigoTarjeta, nBoton from TTARJETACREDITO order by tCodigoTarjeta"
    Set RsBoton = Lib.OpenRecordset(Isql, Cn)
        
    If Sw = True Then
       ActivarBotones (False)
       Blanquear Me
       chkActivo.value = 1
       chkPinPad.value = 0
       txtFactor.Text = "0.00"
       Botonera
    Else
       'Cambiar la Busqueda y Nombre del formulario Cabecera
       ActivarBotones (True)
       Asignar
    End If
    cmdTexto.Caption = "Registro " & IIf(frmTarjeta.RsCabecera.RecordCount = 0, 0, frmTarjeta.RsCabecera.AbsolutePosition) & " de " & frmTarjeta.RsCabecera.RecordCount
    Screen.MousePointer = vbDefault
End Sub

Private Sub Form_Unload(Cancel As Integer)
    'Cambia el Nombre del Formulario
    Set frmTarjetaDetalle = Nothing
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

Private Sub Botonera()
    Dim i As Integer
    txtBoton.Caption = "NA"
    If RsBoton.RecordCount <> 0 Then
        For i = 1 To 8
            RsBoton.MoveFirst
            RsBoton.Find ("nBoton=" & i)
            If RsBoton.EOF Then
               cmdBoton(i).backColor = vbButtonFace
               cmdBoton(i).Enabled = True
            Else
               cmdBoton(i).Enabled = False
               If RsBoton!tCodigoTarjeta = txtCodigo.Text Then
                  txtBoton.Caption = str(i)
                  cmdBoton(i).backColor = vbRed
               Else
                  cmdBoton(i).backColor = vbBlue
               End If
            End If
        Next i
    Else
       For i = 1 To 8
           cmdBoton(i).backColor = vbButtonFace
           cmdBoton(i).Enabled = True
       Next i
    End If
End Sub


Private Sub txtDetallado_LostFocus()
   Call ValidaStr(txtDetallado)
End Sub

Private Sub txtResumido_LostFocus()
   Call ValidaStr(txtResumido)
End Sub


