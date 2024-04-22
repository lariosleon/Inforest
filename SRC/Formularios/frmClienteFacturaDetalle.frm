VERSION 5.00
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "msdatlst.ocx"
Begin VB.Form frmClienteFacturaDetalle 
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   4200
   ClientLeft      =   2520
   ClientTop       =   2640
   ClientWidth     =   9540
   Icon            =   "frmClienteFacturaDetalle.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4200
   ScaleWidth      =   9540
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
      Height          =   3495
      Left            =   1680
      TabIndex        =   19
      Top             =   0
      Width           =   7815
      Begin VB.TextBox txtUrbanizacion 
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
         Left            =   1515
         MaxLength       =   200
         TabIndex        =   34
         Text            =   " "
         Top             =   1920
         Width           =   6210
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Ubigeo"
         Height          =   300
         Left            =   5160
         TabIndex        =   33
         Top             =   360
         Width           =   855
      End
      Begin VB.TextBox txtCodigoUbigeo 
         Alignment       =   2  'Center
         BackColor       =   &H80000018&
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
         ForeColor       =   &H00FF0000&
         Height          =   285
         Left            =   6135
         MaxLength       =   50
         TabIndex        =   32
         Top             =   360
         Width           =   1575
      End
      Begin VB.CommandButton cmdVerifica 
         Caption         =   "Verificar Ruc SUNAT"
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
         Left            =   5640
         TabIndex        =   31
         Top             =   1440
         Width           =   2055
      End
      Begin VB.TextBox txtEnlace 
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
         MaxLength       =   50
         TabIndex        =   28
         Text            =   " "
         Top             =   3045
         Width           =   2835
      End
      Begin VB.TextBox txtCorreo 
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
         Left            =   1515
         MaxLength       =   200
         TabIndex        =   5
         Text            =   " "
         Top             =   2685
         Width           =   6210
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
         Left            =   1515
         MaxLength       =   200
         TabIndex        =   4
         Text            =   " "
         Top             =   2300
         Width           =   6210
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
         Left            =   1515
         MaxLength       =   200
         TabIndex        =   1
         Text            =   " "
         Top             =   753
         Width           =   6210
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
         Left            =   4890
         MaxLength       =   15
         TabIndex        =   3
         Text            =   " "
         Top             =   1116
         Width           =   2835
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
         Left            =   1515
         Locked          =   -1  'True
         TabIndex        =   0
         TabStop         =   0   'False
         Top             =   390
         Width           =   1170
      End
      Begin VB.CheckBox chkActivo 
         Alignment       =   1  'Right Justify
         Caption         =   "Activo :"
         Height          =   195
         Left            =   855
         TabIndex        =   6
         Top             =   3090
         Width           =   840
      End
      Begin MSDataListLib.DataCombo cboTipoIdentidad 
         Height          =   315
         Left            =   1515
         TabIndex        =   2
         Top             =   1080
         Width           =   2280
         _ExtentX        =   4022
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
      Begin MSDataListLib.DataCombo cboTipoCliente 
         Height          =   315
         Left            =   1515
         TabIndex        =   30
         Top             =   1485
         Width           =   2280
         _ExtentX        =   4022
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
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "Urbanizacion :"
         Height          =   195
         Left            =   420
         TabIndex        =   35
         Top             =   1930
         Width           =   1020
      End
      Begin VB.Label lblTipoCliente 
         AutoSize        =   -1  'True
         Caption         =   "Tipo Cliente :"
         Height          =   195
         Left            =   510
         TabIndex        =   29
         Top             =   1530
         Width           =   930
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Enlace"
         Height          =   195
         Left            =   4200
         TabIndex        =   27
         Top             =   3090
         Width           =   495
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Tipo Identificación :"
         Height          =   195
         Left            =   45
         TabIndex        =   26
         Top             =   1125
         Width           =   1395
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Correo :"
         Height          =   195
         Left            =   885
         TabIndex        =   25
         Top             =   2685
         Width           =   555
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Dirección :"
         Height          =   195
         Left            =   675
         TabIndex        =   24
         Top             =   2300
         Width           =   765
      End
      Begin VB.Label Label10 
         AutoSize        =   -1  'True
         Caption         =   "Razón Social :"
         Height          =   195
         Left            =   405
         TabIndex        =   22
         Top             =   780
         Width           =   1035
      End
      Begin VB.Label Label14 
         AutoSize        =   -1  'True
         Caption         =   "Identificador :"
         Height          =   195
         Left            =   3960
         TabIndex        =   21
         Top             =   1125
         Width           =   960
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         Caption         =   "Código :"
         Height          =   195
         Left            =   855
         TabIndex        =   20
         Top             =   435
         Width           =   585
      End
   End
   Begin VB.PictureBox Picture 
      Align           =   2  'Align Bottom
      Height          =   750
      Left            =   0
      ScaleHeight     =   690
      ScaleWidth      =   9480
      TabIndex        =   11
      TabStop         =   0   'False
      Top             =   3450
      Width           =   9540
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
         Left            =   8220
         Picture         =   "frmClienteFacturaDetalle.frx":0442
         Style           =   1  'Graphical
         TabIndex        =   10
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
         Left            =   7050
         Picture         =   "frmClienteFacturaDetalle.frx":0534
         Style           =   1  'Graphical
         TabIndex        =   9
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
         Left            =   4710
         Picture         =   "frmClienteFacturaDetalle.frx":0636
         Style           =   1  'Graphical
         TabIndex        =   7
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
         Left            =   5880
         Picture         =   "frmClienteFacturaDetalle.frx":0B68
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
         ScaleWidth      =   4590
         TabIndex        =   12
         Top             =   60
         Width           =   4650
         Begin VB.CommandButton cmdNavegar 
            Height          =   555
            Index           =   1
            Left            =   480
            Picture         =   "frmClienteFacturaDetalle.frx":109A
            Style           =   1  'Graphical
            TabIndex        =   18
            TabStop         =   0   'False
            Top             =   0
            Width           =   480
         End
         Begin VB.CommandButton cmdNavegar 
            Height          =   555
            Index           =   2
            Left            =   960
            Picture         =   "frmClienteFacturaDetalle.frx":15DC
            Style           =   1  'Graphical
            TabIndex        =   17
            TabStop         =   0   'False
            Top             =   0
            Width           =   480
         End
         Begin VB.CommandButton cmdNavegar 
            Height          =   555
            Index           =   0
            Left            =   0
            Picture         =   "frmClienteFacturaDetalle.frx":1B1E
            Style           =   1  'Graphical
            TabIndex        =   16
            TabStop         =   0   'False
            Top             =   0
            Width           =   480
         End
         Begin VB.CommandButton cmdNavegar 
            Height          =   555
            Index           =   5
            Left            =   4110
            Picture         =   "frmClienteFacturaDetalle.frx":2060
            Style           =   1  'Graphical
            TabIndex        =   15
            TabStop         =   0   'False
            Top             =   0
            Width           =   480
         End
         Begin VB.CommandButton cmdNavegar 
            Height          =   555
            Index           =   4
            Left            =   3630
            Picture         =   "frmClienteFacturaDetalle.frx":25A2
            Style           =   1  'Graphical
            TabIndex        =   14
            TabStop         =   0   'False
            Top             =   0
            Width           =   480
         End
         Begin VB.CommandButton cmdNavegar 
            Height          =   555
            Index           =   3
            Left            =   3150
            Picture         =   "frmClienteFacturaDetalle.frx":2AE4
            Style           =   1  'Graphical
            TabIndex        =   13
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
            TabIndex        =   23
            Top             =   180
            Width           =   1665
         End
      End
   End
   Begin VB.Image Image 
      Height          =   3465
      Left            =   0
      Picture         =   "frmClienteFacturaDetalle.frx":3026
      Stretch         =   -1  'True
      Top             =   30
      Width           =   1710
   End
End
Attribute VB_Name = "frmClienteFacturaDetalle"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim RsTipoIdentidad As Recordset
Dim RsTipoCliente As Recordset
Dim RsParametro As Recordset
'Dim Isql As String
Public lagregar As Boolean

Sub Asignar()
    With frmClienteFactura.RsCabecera
        'Cuadro de Texto
        txtCodigo = IIf(IsNull(!codigo), "", !codigo)
        txtDetallado = IIf(IsNull(!Descripcion), "", !Descripcion)
        txtResumido = IIf(IsNull(!tIdentidad), "", !tIdentidad)
        txtDireccion = IIf(IsNull(!tDireccion), "", !tDireccion)
        
        txtCorreo = IIf(IsNull(!tcorreo), "", !tcorreo)
        cboTipoIdentidad.BoundText = IIf(IsNull(!tTipoIdentidad), "", !tTipoIdentidad)
        
        'Check Box
        chkActivo = IIf(!lActivo = True, 1, 0)
        txtEnlace = IIf(IsNull(!tEnlace), "", !tEnlace)
        
        'Tipo cliente
        cboTipoCliente.BoundText = IIf(IsNull(!tTipoCliente), "", !tTipoCliente)
        
        Me.txtCodigoUbigeo = IIf(IsNull(!CodigoUbigeo), "", !CodigoUbigeo)
        Me.txtUrbanizacion = IIf(IsNull(!Urbanizacion), "", !Urbanizacion)
        
    End With
    
End Sub

Private Sub cmdNavegar_Click(Index As Integer)
    Select Case Index
           Case Is = 0 'Primero
                MoverPuntero Primero, frmClienteFactura.grdGrilla
           Case Is = 1 'PgUp
                MoverPuntero pgup, frmClienteFactura.grdGrilla
           Case Is = 2 'Previo
                MoverPuntero previo, frmClienteFactura.grdGrilla
           Case Is = 3 'Siguiente
                MoverPuntero siguiente, frmClienteFactura.grdGrilla
           Case Is = 4 'PgDn
                MoverPuntero pgdn, frmClienteFactura.grdGrilla
           Case Is = 5 'Ultimo
                MoverPuntero Ultimo, frmClienteFactura.grdGrilla
    End Select
   Asignar
   cmdTexto.Caption = "Registro " & frmClienteFactura.RsCabecera.AbsolutePosition & " de " & frmClienteFactura.RsCabecera.RecordCount
End Sub

Private Sub cmdOpcion_Click(Index As Integer)

   Dim xtTipoIdentidad As String
   Dim Numero As String
   Dim correo As String

   Select Case Index
          Case Is = 0 ' Agregar
               Sw = True
               ActivarBotones (False)
               Blanquear Me
               lagregar = True
               LlenaCombos
               cboTipoIdentidad.BoundText = ""
               chkActivo.value = 1
               'Cambia el Nombre del Primer Text
               txtDetallado.SetFocus
                    
          Case Is = 1 ' Grabar
               Dim nCorrela As String
               Dim nPos As Variant
               'Chequea Datos
               If Trim(txtDetallado.Text) = "" Then MsgBox "Ingrese la Razón Social", vbExclamation, sMensaje: txtDetallado.SetFocus: Exit Sub
               If txtResumido.Text = "" Then MsgBox "Ingrese el Id. Tributario", vbExclamation, sMensaje: txtResumido.SetFocus: Exit Sub
               If cboTipoIdentidad.Text = "" Then MsgBox "Seleccione el Tipo de Identidad", vbExclamation, sMensaje: cboTipoIdentidad.SetFocus: Exit Sub
               
               If lSAP Then
                    If cboTipoCliente.Text = "" Then MsgBox "Seleccione el Tipo de Cliente.", vbExclamation, sMensaje: cboTipoCliente.SetFocus: Exit Sub
               End If
               
               ' cambios para validar DNI y numeros
               If RsParametro!lValidaDNI = True Then
                  If cboTipoIdentidad.SelectedItem = 2 Then
                    Numero = modProcedimiento.ValidarDNI(LTrim(Me.txtResumido))
                    If Numero = False Then
                    MsgBox "El DNI ingresado no es valido", vbCritical, sMensaje
                    Exit Sub
                    End If
                  End If
                  ' validar correo
                 correo = modProcedimiento.Validar_Email(Me.txtCorreo.Text)
                 If correo = False Then
                 MsgBox "El correo electronico es invalido", vbCritical, sMensaje
                 Exit Sub
                 End If
                  
               End If
               '---------------------------------
                    
               If pais = "002" Then ' ECUADOR
                        If Calcular("Select isnull(nValor,0) As Codigo from vtipoidentidad where Codigo= '" & Me.cboTipoIdentidad.BoundText & "'", Cn) Then
                            If Len(Trim(txtResumido.Text)) = 13 Or Len(Trim(txtResumido.Text)) = 10 Then
        
                            Else
                               MsgBox "La longitud del Identificador debe ser 10(Cédula) ó 13(RUC)", vbCritical, sMensaje
                               Exit Sub
                            End If
                        End If
                        
                        If Len(Trim(txtResumido.Text)) = 10 Then
                            xtTipoIdentidad = "01"
                        ElseIf Len(Trim(txtResumido.Text)) = 13 Then
                            xtTipoIdentidad = "02"
                        End If
                Else
                        'PERU - BOLIVIA
                        If Calcular("Select isnull(nValor,0) As Codigo from vtipoidentidad where Codigo= '" & Me.cboTipoIdentidad.BoundText & "'", Cn) Then
                        
                                If lLongitud And Len(Trim(txtResumido.Text)) <> nLongitud Then
                                   MsgBox "La longitud del Identificador debe ser " & nLongitud, vbCritical, sMensaje
                                   Exit Sub
                                ElseIf Not lLongitud And Len(Trim(txtResumido.Text)) < nLongitud Then
                                   MsgBox "La longitud del Identificador debe ser mayor igual a " & nLongitud, vbCritical, sMensaje
                                   Exit Sub
                                End If
                                
                                If Not ValidaRuc(txtResumido.Text) Then
                                   MsgBox "El número ingresado no es válido", vbCritical, sMensaje
                                   Exit Sub
                                End If
                        End If
                        xtTipoIdentidad = ""
               End If
               
               If lFEpape And (cboTipoIdentidad.BoundText = "01" Or cboTipoIdentidad.BoundText = "02") Then
                    If txtCorreo.Text = "" Then MsgBox "Ingrese el Correo Electrónico", vbExclamation, sMensaje: txtCorreo.SetFocus: Exit Sub
                    'VALIDA MAIL
                    If Not Validar_Email(txtCorreo.Text) Then
                       MsgBox "El Correo ingresado no es válido", vbCritical, sMensaje
                       Exit Sub
                    End If
               End If

               If Val(Calcular("select tIdentidad as Codigo from TCLIENTE where tIdentidad = '" & txtResumido.Text & "'", Cn)) > 0 And Sw Then
                  MsgBox "Identificador Repetido", vbCritical, sMensaje
                  Exit Sub
               End If
                    
               If Sw Then
                  If Calcular("select count(tIdentidad) as Codigo from TCLIENTE where tIdentidad='" & txtResumido.Text & "'", Cn) > 0 Then
                     MsgBox "Error: Identificador Existente", vbCritical, sMensaje
                     Exit Sub
                  End If
                              
                  'Obtiene el Numero de Orden
                  nCorrela = Calcular("select max(tCodigoCliente) as Codigo from TCLIENTE", Cn)
                  If IsNull(nCorrela) Or nCorrela = "" Then
                      txtCodigo.Text = "00001"
                  Else
                      txtCodigo.Text = Lib.Correlativo(nCorrela, 5)
                  End If
                  Sw = False
                   
                  'Cambiar el SQL
                  Isql = "insert into TCLIENTE( " & _
                         "tCodigoCliente, tEmpresa, tIdentidad, tDireccion, tUsuario, fRegistro, tCorreo, tEnlace, tTipoIdentidad,tTipoCliente, lActivo, tUbigeo, tUrbanizacion) " & _
                         "values ( '" & txtCodigo.Text & "', " & _
                                " '" & txtDetallado.Text & "', " & _
                                " '" & txtResumido.Text & "', " & _
                                " '" & txtDireccion.Text & "', " & _
                                " '" & sUsuario & "', getdate(), " & _
                                " '" & txtCorreo.Text & "', " & _
                                " '" & txtEnlace.Text & "', " & _
                                " '" & cboTipoIdentidad.BoundText & "', " & _
                                " '" & cboTipoCliente.BoundText & "', " & _
                                       chkActivo.value & ",'" & Me.txtCodigoUbigeo.Text & "', '" & Me.txtUrbanizacion.Text & "') "
            
                  Cn.Execute Isql
                  frmClienteFactura.RsCabecera.Sort = "Codigo ASC"
                  frmClienteFactura.RsCabecera.Requery
                  frmClienteFactura.RsCabecera.MoveLast
                  ActivarBotones (True)
                  MsgBox "Registro Guardado", vbInformation, sMensaje
                  cmdTexto.Caption = "Registro " & IIf(frmClienteFactura.RsCabecera.RecordCount = 0, 0, frmClienteFactura.RsCabecera.AbsolutePosition) & " de " & frmClienteFactura.RsCabecera.RecordCount
               
               Else
               
                  If lFEpape And (cboTipoIdentidad.BoundText = "01" Or cboTipoIdentidad.BoundText = "02") Then
                         If txtCorreo.Text = "" Then MsgBox "Ingrese el Correo Electrónico", vbExclamation, sMensaje: txtCorreo.SetFocus: Exit Sub
                          'VALIDA MAIL
                         If Not Validar_Email(txtCorreo.Text) Then
                            MsgBox "El Correo ingresado no es válido", vbCritical, sMensaje
                            Exit Sub
                         End If
                  End If
                    
                  'Cambiar el SQL
                  If Calcular("select count(tIdentidad) as Codigo from TCLIENTE where tCodigoCliente <>'" & txtCodigo.Text & "' and tIdentidad='" & txtResumido.Text & "'", Cn) > 0 Then
                     MsgBox "Error: Identificador Existente", vbCritical, sMensaje
                     Exit Sub
                  End If
                  
                  Isql = "update TCLIENTE set " & _
                         "tEmpresa ='" & txtDetallado.Text & "', " & _
                         "tIdentidad ='" & txtResumido.Text & "', " & _
                         "tDireccion ='" & txtDireccion.Text & "', " & _
                         "tCorreo ='" & txtCorreo.Text & "', " & _
                         "tEnlace ='" & txtEnlace.Text & "', " & _
                         "tTipoIdentidad = '" & cboTipoIdentidad.BoundText & "', " & _
                         "tTipoCliente = '" & cboTipoCliente.BoundText & "', " & _
                         "lActivo =" & chkActivo.value & "," & _
                         "tUbigeo ='" & Me.txtCodigoUbigeo.Text & "'," & _
                         "tUrbanizacion ='" & Me.txtUrbanizacion.Text & "'," & _
                         "lreplica=1 where tCodigoCliente = '" & txtCodigo & "'"
                       
                   Cn.Execute Isql
                   nPos = frmClienteFactura.RsCabecera.Bookmark
                   frmClienteFactura.RsCabecera.Requery
                   If frmClienteFactura.RsCabecera.RecordCount = 0 Then
                      frmClienteFactura.RsCabecera.Filter = adFilterNone
                   End If
                   frmClienteFactura.RsCabecera.Bookmark = nPos
                   Screen.MousePointer = vbDefault
                   MsgBox "Registro Modificado", vbInformation, sMensaje
               End If
          
          Case Is = 2 ' Eliminar
               If frmClienteFactura.RsCabecera.RecordCount = 0 Then
                  Exit Sub
               End If
               
               'Cambia el MsgBox
               If MsgBox("Seguro de Eliminar el Cliente" & txtDetallado.Text & "?", vbQuestion + vbOKCancel, sMensaje) = vbCancel Then
                  Exit Sub
               End If
               
               If Calcular("select count(tDocumento) as Codigo from MDOCUMENTO where tCodigoCliente ='" & txtCodigo.Text & "'", Cn) > 0 Then
                  MsgBox "Error: Se han generado documentos con este Cliente" & Chr(13) & "No se puede Eliminar", vbCritical, sMensaje
                  Exit Sub
               End If
                               
               'Cambia el Delete
               Cn.Execute "delete from TCLIENTE where tCodigoCliente = '" & txtCodigo & "'"
               frmClienteFactura.RsCabecera.Requery
               If frmClienteFactura.RsCabecera.RecordCount <> 0 Then
                  frmClienteFactura.RsCabecera.MoveLast
                  Asignar
                  cmdTexto.Caption = "Registro " & IIf(frmClienteFactura.RsCabecera.RecordCount = 0, 0, frmClienteFactura.RsCabecera.AbsolutePosition) & " de " & frmClienteFactura.RsCabecera.RecordCount
               Else
                  ActivarBotones False
                  Blanquear Me
                  Sw = True
               End If
          
          Case Is = 3 ' Salir
               Unload Me
   End Select

End Sub

Private Sub cmdVerifica_Click()
On Error GoTo fin
    Dim loRUC As vfpsrucperu.vfpsruc
    Set loRUC = New vfpsrucperu.vfpsruc
    
    Dim lcNroRuc As String
 
        lcNroRuc = txtResumido.Text
        If loRUC.VFPs_ConsultarRUC(lcNroRuc, False) Then
            'DEVOLVIO LA CONSULTA CORRECTAMENTE
            'PROPIEDADES A CONSULTAR LUEGO DE LA CONSULTA DE RUC
            'loRUC.LCRUC
            txtDetallado.Text = loRUC.LCRAZONSOCIAL
            'loRUC.LCTIPOCON
            'loRUC.c
            'loRUC.LCTELEFONO
            'loRUC.LDFECHAINS
            'loRUC.LCESTADO
            'loRUC.LCCONDICION
            txtDireccion.Text = loRUC.LCDIRECCION
            'loRUC.LDFECHAINICIO
            'loRUC.LCSISEMICOMP
            'loRUC.LCACTCOMEXT
            'loRUC.LCSISCONTA
        End If
Exit Sub
fin:
    MsgBox "Error al consultar Ruc", vbInformation, sMensaje
End Sub

Private Sub Command1_Click()
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

Private Sub Form_Load()

   ' cambios validar DNI
   Isql = "select lValidaDNI from TPARAMETRO"
   Set RsParametro = Lib.OpenRecordset(Isql, Cn)
   '--------------------------------------------
   
    Screen.MousePointer = vbHourglass
    Centrar Me
    LlenaCombos
    'Ingrese el SubTitulo
    Me.Caption = " Mantenimiento de Clientes / Transportistas "
    fraDetalle.Caption = Me.Caption
    
    If lSAP Then
        lblTipoCliente.Visible = True
        cboTipoCliente.Visible = True
    Else
        lblTipoCliente.Visible = False
        cboTipoCliente.Visible = False
    End If
    
    If pais = "000" Then
        cmdVerifica.Visible = True
    Else
        cmdVerifica.Visible = False
    End If
    
    If Sw = True Then
       ActivarBotones (False)
       Blanquear Me
       chkActivo.value = 1
    Else
       'Cambiar la Busqueda y Nombre del formulario Cabecera
       'frmClienteFactura.RsCabecera.Find ("Codigo = '" & frmClienteFactura.RsCabecera!Codigo & "'")
       ActivarBotones (True)
       Asignar
    End If
    
    If Not lClub Then
        Label4.Visible = False
        txtEnlace.Visible = False
    End If
    
    cmdTexto.Caption = "Registro " & frmClienteFactura.RsCabecera.AbsolutePosition & " de " & frmClienteFactura.RsCabecera.RecordCount
    Screen.MousePointer = vbDefault
End Sub

Sub LlenaCombos()
    With cboTipoIdentidad
         If lagregar Then
            Isql = "Select * from vTipoIdentidad where lactivo=1 order by Descripcion"
         Else
            Isql = "Select * from vTipoIdentidad order by Descripcion"
         End If
         
         Set RsTipoIdentidad = Lib.OpenRecordset(Isql, Cn)
         Set .RowSource = RsTipoIdentidad
             .DataField = "tResumido"
             .ListField = "tResumido"
             .BoundColumn = "Codigo"
    End With
    
    With cboTipoCliente
         If lagregar Then
            Isql = "Select * from vTipoGrupoCliente where lactivo=1 order by Descripcion"
         Else
            Isql = "Select * from vTipoGrupoCliente order by Descripcion"
         End If
         
         Set RsTipoCliente = Lib.OpenRecordset(Isql, Cn)
         Set .RowSource = RsTipoCliente
             .DataField = "tResumido"
             .ListField = "tResumido"
             .BoundColumn = "Codigo"
    End With
    
End Sub
Private Sub Form_Unload(Cancel As Integer)
    'Cambia el Nombre del Formulario
        Set RsTipoIdentidad = Nothing
        Set RsTipoCliente = Nothing

    Set frmClienteFacturaDetalle = Nothing
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


Private Sub txtDetallado_LostFocus()
  ' Call ValidaStr(txtDetallado)
End Sub

Private Sub txtDireccion_LostFocus()
   Call ValidaStr(txtDireccion)
End Sub

Private Sub txtResumido_LostFocus()
    Call ValidaStr(txtResumido)
    If cboTipoIdentidad.BoundText = "02" Then
    'Consitencia RUC
            If lLongitud And Len(Trim(txtResumido.Text)) <> nLongitud Then
               MsgBox "La longitud del Id. Tributario debe ser " & nLongitud, vbCritical, sMensaje
               Exit Sub
            ElseIf Not lLongitud And Len(Trim(txtResumido.Text)) < nLongitud Then
               MsgBox "La longitud del Identificador debe ser mayor igual a " & nLongitud, vbCritical, sMensaje
               Exit Sub
            End If
    End If
End Sub

'cambio de validar DNI
Private Sub txtResumido_GotFocus()
If cboTipoIdentidad = "" Then
    MsgBox "Debe colocar un identificador"
    foco
    Exit Sub
End If
End Sub

'cambio de validar DNI
Private Function foco()
txtDetallado.SetFocus
End Function


'cambio de validar DNI
Private Sub txtResumido_KeyPress(KeyAscii As Integer)
    If RsParametro!lValidaDNI = True Then
        If cboTipoIdentidad.SelectedItem = 2 Then
            If KeyAscii = 13 Then
            KeyAscii = 0
            SendKeys "{tab}"
            ElseIf KeyAscii <> 8 Then
            If Not IsNumeric(Chr(KeyAscii)) Then
            Beep
            KeyAscii = 0
            End If
            End If
        End If
    End If
    
    If NadaSimbolos(KeyAscii) = False Then
        Beep
        KeyAscii = 0
    End If
End Sub

