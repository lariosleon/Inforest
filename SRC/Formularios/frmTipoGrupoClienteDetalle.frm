VERSION 5.00
Begin VB.Form frmTipoGrupoClienteDetalle 
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   3300
   ClientLeft      =   2520
   ClientTop       =   2640
   ClientWidth     =   9495
   Icon            =   "frmTipoGrupoClienteDetalle.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3300
   ScaleWidth      =   9495
   Begin VB.TextBox txtCodigoExterno 
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
      TabIndex        =   21
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
      TabIndex        =   16
      Top             =   0
      Width           =   7155
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
         TabIndex        =   3
         Top             =   2055
         Width           =   840
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "C�digo Externo :"
         Height          =   195
         Left            =   675
         TabIndex        =   22
         Top             =   1485
         Width           =   1170
      End
      Begin VB.Label Label10 
         AutoSize        =   -1  'True
         Caption         =   "Descripci�n Detallada :"
         Height          =   195
         Left            =   210
         TabIndex        =   19
         Top             =   757
         Width           =   1650
      End
      Begin VB.Label Label14 
         AutoSize        =   -1  'True
         Caption         =   "Descripci�n Resumida :"
         Height          =   195
         Left            =   180
         TabIndex        =   18
         Top             =   1139
         Width           =   1680
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         Caption         =   "C�digo :"
         Height          =   195
         Left            =   1275
         TabIndex        =   17
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
      TabIndex        =   8
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
         Picture         =   "frmTipoGrupoClienteDetalle.frx":0442
         Style           =   1  'Graphical
         TabIndex        =   7
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
         Picture         =   "frmTipoGrupoClienteDetalle.frx":0534
         Style           =   1  'Graphical
         TabIndex        =   6
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
         Picture         =   "frmTipoGrupoClienteDetalle.frx":0636
         Style           =   1  'Graphical
         TabIndex        =   4
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
         Picture         =   "frmTipoGrupoClienteDetalle.frx":0B68
         Style           =   1  'Graphical
         TabIndex        =   5
         Top             =   60
         Width           =   1170
      End
      Begin VB.PictureBox PicNavegacion 
         BackColor       =   &H80000004&
         Height          =   615
         Left            =   30
         ScaleHeight     =   555
         ScaleWidth      =   4590
         TabIndex        =   9
         Top             =   60
         Width           =   4650
         Begin VB.CommandButton cmdNavegar 
            Height          =   555
            Index           =   1
            Left            =   480
            Picture         =   "frmTipoGrupoClienteDetalle.frx":109A
            Style           =   1  'Graphical
            TabIndex        =   15
            TabStop         =   0   'False
            Top             =   0
            Width           =   480
         End
         Begin VB.CommandButton cmdNavegar 
            Height          =   555
            Index           =   2
            Left            =   960
            Picture         =   "frmTipoGrupoClienteDetalle.frx":15DC
            Style           =   1  'Graphical
            TabIndex        =   14
            TabStop         =   0   'False
            Top             =   0
            Width           =   480
         End
         Begin VB.CommandButton cmdNavegar 
            Height          =   555
            Index           =   0
            Left            =   0
            Picture         =   "frmTipoGrupoClienteDetalle.frx":1B1E
            Style           =   1  'Graphical
            TabIndex        =   13
            TabStop         =   0   'False
            Top             =   0
            Width           =   480
         End
         Begin VB.CommandButton cmdNavegar 
            Height          =   555
            Index           =   5
            Left            =   4110
            Picture         =   "frmTipoGrupoClienteDetalle.frx":2060
            Style           =   1  'Graphical
            TabIndex        =   12
            TabStop         =   0   'False
            Top             =   0
            Width           =   480
         End
         Begin VB.CommandButton cmdNavegar 
            Height          =   555
            Index           =   4
            Left            =   3630
            Picture         =   "frmTipoGrupoClienteDetalle.frx":25A2
            Style           =   1  'Graphical
            TabIndex        =   11
            TabStop         =   0   'False
            Top             =   0
            Width           =   480
         End
         Begin VB.CommandButton cmdNavegar 
            Height          =   555
            Index           =   3
            Left            =   3150
            Picture         =   "frmTipoGrupoClienteDetalle.frx":2AE4
            Style           =   1  'Graphical
            TabIndex        =   10
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
            TabIndex        =   20
            Top             =   150
            Width           =   1665
         End
      End
   End
   Begin VB.Image Image 
      Height          =   2505
      Left            =   45
      Picture         =   "frmTipoGrupoClienteDetalle.frx":3026
      Stretch         =   -1  'True
      Top             =   15
      Width           =   2205
   End
End
Attribute VB_Name = "frmTipoGrupoClienteDetalle"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Sub Asignar()
    With frmTipoGrupoCliente.RsCabecera
        'Cuadro de Texto
        txtCodigo = IIf(IsNull(!codigo), "", !codigo)
        txtDetallado = IIf(IsNull(!Descripcion), "", !Descripcion)
        txtResumido = IIf(IsNull(!tResumido), "", !tResumido)
            
        'Check Box
        chkActivo = IIf(!lActivo = True, 1, 0)
        
        txtCodigoExterno = IIf(IsNull(!tValor), "", !tValor)
        
    End With
End Sub

Private Sub cmdNavegar_Click(Index As Integer)
    Select Case Index
           Case Is = 0 'Primero
                MoverPuntero Primero, frmTipoGrupoCliente.grdGrilla
           Case Is = 1 'PgUp
                MoverPuntero pgup, frmTipoGrupoCliente.grdGrilla
           Case Is = 2 'Previo
                MoverPuntero previo, frmTipoGrupoCliente.grdGrilla
           Case Is = 3 'Siguiente
                MoverPuntero siguiente, frmTipoGrupoCliente.grdGrilla
           Case Is = 4 'PgDn
                MoverPuntero pgdn, frmTipoGrupoCliente.grdGrilla
           Case Is = 5 'Ultimo
                MoverPuntero Ultimo, frmTipoGrupoCliente.grdGrilla
    End Select
   Asignar
   cmdTexto.Caption = "Registro " & frmTipoGrupoCliente.RsCabecera.AbsolutePosition & " de " & frmTipoGrupoCliente.RsCabecera.RecordCount
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
               If txtDetallado.Text = "" Then MsgBox "Ingrese la Descripci�n Detallada", vbExclamation, sMensaje: txtDetallado.SetFocus: Exit Sub
               If txtResumido.Text = "" Then MsgBox "Ingrese la Descripci�n Resumida", vbExclamation, sMensaje: txtResumido.SetFocus: Exit Sub
                    
               If Sw Then
                  'Obtiene el Numero de Orden
                  nCorrela = Calcular("select max(tCodigo) as Codigo from TTABLA where tTabla ='TIPOGRUPOCLIENTE' ", Cn)
                  If IsNull(nCorrela) Or nCorrela = "" Then
                      txtCodigo.Text = "01"
                  Else
                      txtCodigo.Text = Lib.Correlativo(nCorrela, 2)
                  End If
                  Sw = False
                   
                  'Cambiar el SQL
                  Isql = "insert into TTABLA( " & _
                         "tTabla, tCodigo, tDetallado, tResumido, tValor, lActivo) " & _
                         "values ('TIPOGRUPOCLIENTE', " & _
                                " '" & txtCodigo.Text & "', " & _
                                " '" & txtDetallado.Text & "', " & _
                                " '" & txtResumido.Text & "', " & _
                                " '" & txtCodigoExterno.Text & "', " & _
                                       chkActivo.value & ") "
           
                      Cn.Execute Isql
                      frmTipoGrupoCliente.RsCabecera.Sort = "Codigo ASC"
                      frmTipoGrupoCliente.RsCabecera.Requery
                      frmTipoGrupoCliente.RsCabecera.MoveLast
                      MsgBox "Registro Guardado", vbInformation, sMensaje
                      ActivarBotones (True)
                      cmdTexto.Caption = "Registro " & IIf(frmTipoGrupoCliente.RsCabecera.RecordCount = 0, 0, frmTipoGrupoCliente.RsCabecera.AbsolutePosition) & " de " & frmTipoGrupoCliente.RsCabecera.RecordCount
              Else
                 'Cambiar el SQL
                 Isql = "update TTABLA set " & _
                        "tDetallado ='" & txtDetallado.Text & "', " & _
                        "tResumido ='" & txtResumido.Text & "', " & _
                        "tValor ='" & txtCodigoExterno.Text & "', " & _
                        "lActivo =" & chkActivo.value & _
                        " where tTAbla = 'TIPOGRUPOCLIENTE' and tCodigo = '" & txtCodigo & "'"
                      
                  Cn.Execute Isql
                  nPos = frmTipoGrupoCliente.RsCabecera.Bookmark
                  frmTipoGrupoCliente.RsCabecera.Requery
                  If frmTipoGrupoCliente.RsCabecera.RecordCount = 0 Then
                     frmTipoGrupoCliente.RsCabecera.Filter = adFilterNone
                  End If
                  frmTipoGrupoCliente.RsCabecera.Bookmark = nPos
                  Screen.MousePointer = vbDefault
                  MsgBox "Registro Modificado", vbInformation, sMensaje
              End If
         
          Case Is = 2 ' Eliminar
               If frmTipoGrupoCliente.RsCabecera.RecordCount = 0 Then
                  Exit Sub
               End If
               'Cambia el MsgBox
               If MsgBox("Seguro de Eliminar el Tipo de Cliente facturado " & txtCodigo & "?", vbQuestion + vbOKCancel, sMensaje) = vbCancel Then
                  Exit Sub
               End If
               
               'Cambia el Delete
               Cn.Execute "delete from TTABLA where tTabla = 'TIPOGRUPOCLIENTE' and tCodigo = '" & txtCodigo & "'"
               frmTipoGrupoCliente.RsCabecera.Requery
               If frmTipoGrupoCliente.RsCabecera.RecordCount <> 0 Then
                  frmTipoGrupoCliente.RsCabecera.MoveLast
                  Asignar
                  cmdTexto.Caption = "Registro " & IIf(frmTipoGrupoCliente.RsCabecera.RecordCount = 0, 0, frmTipoGrupoCliente.RsCabecera.AbsolutePosition) & " de " & frmTipoGrupoCliente.RsCabecera.RecordCount
                  
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
    Me.Caption = " Mantenimiento de Tipo de Cliente Facturado "
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
    
    cmdTexto.Caption = "Registro " & frmTipoGrupoCliente.RsCabecera.AbsolutePosition & " de " & frmTipoGrupoCliente.RsCabecera.RecordCount
    Screen.MousePointer = vbDefault
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set frmTipoGrupoClienteDetalle = Nothing
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

