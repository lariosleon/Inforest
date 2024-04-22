VERSION 5.00
Begin VB.Form frmTipoIdentidadDetalle 
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   3300
   ClientLeft      =   2520
   ClientTop       =   2640
   ClientWidth     =   9495
   Icon            =   "frmTipoIdentidadDetalle.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3300
   ScaleWidth      =   9495
   Begin VB.TextBox txtCodigoIdentidad 
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
      Left            =   4230
      MaxLength       =   15
      TabIndex        =   22
      Text            =   " "
      Top             =   1440
      Width           =   1215
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
      Height          =   2520
      Left            =   2280
      TabIndex        =   17
      Top             =   0
      Width           =   7155
      Begin VB.TextBox txtRefTipoPersona 
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
         Left            =   5235
         MaxLength       =   15
         TabIndex        =   25
         Text            =   " "
         Top             =   1440
         Width           =   1770
      End
      Begin VB.CheckBox chkValidacion 
         Alignment       =   1  'Right Justify
         Caption         =   "Validación :"
         Height          =   195
         Left            =   1020
         TabIndex        =   3
         Top             =   1830
         Width           =   1125
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
         Left            =   1950
         MaxLength       =   50
         TabIndex        =   1
         Text            =   " "
         Top             =   712
         Width           =   5070
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
         Left            =   1950
         MaxLength       =   15
         TabIndex        =   2
         Text            =   " "
         Top             =   1094
         Width           =   2595
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
         Left            =   1950
         Locked          =   -1  'True
         TabIndex        =   0
         TabStop         =   0   'False
         Top             =   330
         Width           =   1170
      End
      Begin VB.CheckBox chkActivo 
         Alignment       =   1  'Right Justify
         Caption         =   "Activo :"
         Height          =   195
         Left            =   1305
         TabIndex        =   4
         Top             =   2130
         Width           =   840
      End
      Begin VB.Label lblRefTipoPersona 
         AutoSize        =   -1  'True
         Caption         =   "Referencia Tipo Persona :"
         Height          =   195
         Left            =   3315
         TabIndex        =   24
         Top             =   1480
         Width           =   1860
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Código Identidad :"
         Height          =   195
         Left            =   580
         TabIndex        =   23
         Top             =   1480
         Width           =   1290
      End
      Begin VB.Label Label10 
         AutoSize        =   -1  'True
         Caption         =   "Descripción Detallada :"
         Height          =   195
         Left            =   210
         TabIndex        =   20
         Top             =   757
         Width           =   1650
      End
      Begin VB.Label Label14 
         AutoSize        =   -1  'True
         Caption         =   "Descripción Resumida :"
         Height          =   195
         Left            =   180
         TabIndex        =   19
         Top             =   1139
         Width           =   1680
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         Caption         =   "Código :"
         Height          =   195
         Left            =   1275
         TabIndex        =   18
         Top             =   375
         Width           =   585
      End
   End
   Begin VB.PictureBox Picture 
      Align           =   2  'Align Bottom
      Height          =   750
      Left            =   0
      ScaleHeight     =   690
      ScaleWidth      =   9435
      TabIndex        =   9
      TabStop         =   0   'False
      Top             =   2550
      Width           =   9495
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
         Picture         =   "frmTipoIdentidadDetalle.frx":0442
         Style           =   1  'Graphical
         TabIndex        =   8
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
         Picture         =   "frmTipoIdentidadDetalle.frx":0534
         Style           =   1  'Graphical
         TabIndex        =   7
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
         Picture         =   "frmTipoIdentidadDetalle.frx":0636
         Style           =   1  'Graphical
         TabIndex        =   5
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
         Picture         =   "frmTipoIdentidadDetalle.frx":0B68
         Style           =   1  'Graphical
         TabIndex        =   6
         Top             =   60
         Width           =   1170
      End
      Begin VB.PictureBox PicNavegacion 
         BackColor       =   &H80000004&
         Height          =   615
         Left            =   30
         ScaleHeight     =   555
         ScaleWidth      =   4590
         TabIndex        =   10
         Top             =   60
         Width           =   4650
         Begin VB.CommandButton cmdNavegar 
            Height          =   555
            Index           =   1
            Left            =   480
            Picture         =   "frmTipoIdentidadDetalle.frx":109A
            Style           =   1  'Graphical
            TabIndex        =   16
            TabStop         =   0   'False
            Top             =   0
            Width           =   480
         End
         Begin VB.CommandButton cmdNavegar 
            Height          =   555
            Index           =   2
            Left            =   960
            Picture         =   "frmTipoIdentidadDetalle.frx":15DC
            Style           =   1  'Graphical
            TabIndex        =   15
            TabStop         =   0   'False
            Top             =   0
            Width           =   480
         End
         Begin VB.CommandButton cmdNavegar 
            Height          =   555
            Index           =   0
            Left            =   0
            Picture         =   "frmTipoIdentidadDetalle.frx":1B1E
            Style           =   1  'Graphical
            TabIndex        =   14
            TabStop         =   0   'False
            Top             =   0
            Width           =   480
         End
         Begin VB.CommandButton cmdNavegar 
            Height          =   555
            Index           =   5
            Left            =   4110
            Picture         =   "frmTipoIdentidadDetalle.frx":2060
            Style           =   1  'Graphical
            TabIndex        =   13
            TabStop         =   0   'False
            Top             =   0
            Width           =   480
         End
         Begin VB.CommandButton cmdNavegar 
            Height          =   555
            Index           =   4
            Left            =   3630
            Picture         =   "frmTipoIdentidadDetalle.frx":25A2
            Style           =   1  'Graphical
            TabIndex        =   12
            TabStop         =   0   'False
            Top             =   0
            Width           =   480
         End
         Begin VB.CommandButton cmdNavegar 
            Height          =   555
            Index           =   3
            Left            =   3150
            Picture         =   "frmTipoIdentidadDetalle.frx":2AE4
            Style           =   1  'Graphical
            TabIndex        =   11
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
            Left            =   1440
            TabIndex        =   21
            Top             =   150
            Width           =   1665
         End
      End
   End
   Begin VB.Image Image 
      Height          =   2505
      Left            =   45
      Picture         =   "frmTipoIdentidadDetalle.frx":3026
      Stretch         =   -1  'True
      Top             =   15
      Width           =   2205
   End
End
Attribute VB_Name = "frmTipoIdentidadDetalle"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Sub Asignar()
    With frmTipoIdentidad.RsCabecera
        'Cuadro de Texto
        txtCodigo = IIf(IsNull(!codigo), "", !codigo)
        txtDetallado = IIf(IsNull(!Descripcion), "", !Descripcion)
        txtResumido = IIf(IsNull(!tResumido), "", !tResumido)
            
        'Check Box
        chkValidacion = IIf(IsNull(!nValor), 0, !nValor)
        chkActivo = IIf(!lActivo = True, 1, 0)
        
        txtCodigoIdentidad = IIf(IsNull(!tValor), "", !tValor)
        
        txtRefTipoPersona = IIf(IsNull(!tValor2), "", !tValor2)
        
    End With
End Sub

Private Sub cmdNavegar_Click(Index As Integer)
    Select Case Index
           Case Is = 0 'Primero
                MoverPuntero Primero, frmTipoIdentidad.grdGrilla
           Case Is = 1 'PgUp
                MoverPuntero pgup, frmTipoIdentidad.grdGrilla
           Case Is = 2 'Previo
                MoverPuntero previo, frmTipoIdentidad.grdGrilla
           Case Is = 3 'Siguiente
                MoverPuntero siguiente, frmTipoIdentidad.grdGrilla
           Case Is = 4 'PgDn
                MoverPuntero pgdn, frmTipoIdentidad.grdGrilla
           Case Is = 5 'Ultimo
                MoverPuntero Ultimo, frmTipoIdentidad.grdGrilla
    End Select
   Asignar
   cmdTexto.Caption = "Registro " & frmTipoIdentidad.RsCabecera.AbsolutePosition & " de " & frmTipoIdentidad.RsCabecera.RecordCount
End Sub

Private Sub cmdOpcion_Click(Index As Integer)
   Select Case Index
          Case Is = 0 ' Agregar
               Sw = True
               ActivarBotones (False)
               Blanquear Me
               chkActivo.value = 1
               'Cambia el Nombre del Primer Text
               txtDetallado.SetFocus
          
          Case Is = 1 ' Grabar
               Dim nCorrela As String
          
               'Chequea Datos
               If txtDetallado.Text = "" Then MsgBox "Ingrese la Descripción Detallada", vbExclamation, sMensaje: txtDetallado.SetFocus: Exit Sub
               If txtResumido.Text = "" Then MsgBox "Ingrese la Descripción Resumida", vbExclamation, sMensaje: txtResumido.SetFocus: Exit Sub
                    
               If Sw Then
                  'Obtiene el Numero de Orden
                  nCorrela = Calcular("select max(tCodigo) as Codigo from TTABLA where tTabla ='TIPOIDENTIDAD' ", Cn)
                  If IsNull(nCorrela) Or nCorrela = "" Then
                      txtCodigo.Text = "01"
                  Else
                      txtCodigo.Text = Lib.Correlativo(nCorrela, 2)
                  End If
                  Sw = False
                   
                  'Cambiar el SQL
                  Isql = "insert into TTABLA( " & _
                         "tTabla, tCodigo, tDetallado, tResumido, nValor, tValor, tValor2, lActivo) " & _
                         "values ('TIPOIDENTIDAD', " & _
                                " '" & txtCodigo.Text & "', " & _
                                " '" & txtDetallado.Text & "', " & _
                                " '" & txtResumido.Text & "', " & _
                                       chkValidacion.value & ", " & _
                                " '" & txtCodigoIdentidad.Text & "', " & _
                                " '" & txtRefTipoPersona.Text & "', " & _
                                       chkActivo.value & ") "
           
                      Cn.Execute Isql
                      frmTipoIdentidad.RsCabecera.Sort = "Codigo ASC"
                      frmTipoIdentidad.RsCabecera.Requery
                      frmTipoIdentidad.RsCabecera.MoveLast
                      MsgBox "Registro Guardado", vbInformation, sMensaje
                      ActivarBotones (True)
                      cmdTexto.Caption = "Registro " & IIf(frmTipoIdentidad.RsCabecera.RecordCount = 0, 0, frmTipoIdentidad.RsCabecera.AbsolutePosition) & " de " & frmTipoIdentidad.RsCabecera.RecordCount
              Else
                 'Cambiar el SQL
                 Isql = "update TTABLA set " & _
                        "tDetallado ='" & txtDetallado.Text & "', " & _
                        "tResumido ='" & txtResumido.Text & "', " & _
                        "nValor =" & chkValidacion.value & ", " & _
                        "tValor ='" & txtCodigoIdentidad.Text & "', " & _
                        "tValor2 ='" & txtRefTipoPersona.Text & "', " & _
                        "lActivo =" & chkActivo.value & _
                        " where tTAbla = 'TIPOIDENTIDAD' and tCodigo = '" & txtCodigo & "'"
                      
                  Cn.Execute Isql
                  nPos = frmTipoIdentidad.RsCabecera.Bookmark
                  frmTipoIdentidad.RsCabecera.Requery
                  If frmTipoIdentidad.RsCabecera.RecordCount = 0 Then
                     frmTipoIdentidad.RsCabecera.Filter = adFilterNone
                  End If
                  frmTipoIdentidad.RsCabecera.Bookmark = nPos
                  Screen.MousePointer = vbDefault
                  MsgBox "Registro Modificado", vbInformation, sMensaje
              End If
         
          Case Is = 2 ' Eliminar
               If frmTipoIdentidad.RsCabecera.RecordCount = 0 Then
                  Exit Sub
               End If
               'Cambia el MsgBox
               If MsgBox("Seguro de Eliminar el Tipo de Identidad " & txtCodigo & "?", vbQuestion + vbOKCancel, sMensaje) = vbCancel Then
                  Exit Sub
               End If
               
               'Cambia el Delete
               Cn.Execute "delete from TTABLA where tTabla = 'TIPOIDENTIDAD' and tCodigo = '" & txtCodigo & "'"
               frmTipoIdentidad.RsCabecera.Requery
               If frmTipoIdentidad.RsCabecera.RecordCount <> 0 Then
                  frmTipoIdentidad.RsCabecera.MoveLast
                  Asignar
                  cmdTexto.Caption = "Registro " & IIf(frmTipoIdentidad.RsCabecera.RecordCount = 0, 0, frmTipoIdentidad.RsCabecera.AbsolutePosition) & " de " & frmTipoIdentidad.RsCabecera.RecordCount
                  
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
    Me.Caption = " Mantenimiento de Tipo de Identidad "
    fraDetalle.Caption = Me.Caption
       
    'Ingresar la Tabla
    If Sw = True Then
       ActivarBotones (False)
       Blanquear Me
       chkActivo.value = 1
    Else
       'Cambiar la Busqueda y Nombre del formulario Cabecera
       ActivarBotones (True)
       Asignar
    End If
    
    If lSAP Then
        lblRefTipoPersona.Visible = True
        txtRefTipoPersona.Visible = True
    Else
        lblRefTipoPersona.Visible = False
        txtRefTipoPersona.Visible = False
    End If
    
    cmdTexto.Caption = "Registro " & frmTipoIdentidad.RsCabecera.AbsolutePosition & " de " & frmTipoIdentidad.RsCabecera.RecordCount
    Screen.MousePointer = vbDefault
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set frmTipoIdentidadDetalle = Nothing
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

