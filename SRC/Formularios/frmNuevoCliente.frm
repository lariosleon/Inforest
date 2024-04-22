VERSION 5.00
Begin VB.Form frmNuevoCliente 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Agregar Cliente"
   ClientHeight    =   6525
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5310
   Icon            =   "frmNuevoCliente.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6525
   ScaleWidth      =   5310
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox txtUbigeo 
      Alignment       =   2  'Center
      BackColor       =   &H00E0E0E0&
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
      Height          =   375
      Left            =   2560
      MaxLength       =   200
      TabIndex        =   23
      TabStop         =   0   'False
      Top             =   4800
      Width           =   1320
   End
   Begin VB.CommandButton cmdUbigeo 
      Caption         =   "Ubigeo"
      Height          =   555
      Left            =   3960
      TabIndex        =   22
      Top             =   4680
      Width           =   1275
   End
   Begin VB.CommandButton cmdOpcion 
      Caption         =   "&Urbanizacion"
      Height          =   555
      Index           =   8
      Left            =   3960
      TabIndex        =   21
      Top             =   1750
      Width           =   1275
   End
   Begin VB.TextBox txtUrbanizacion 
      BackColor       =   &H00E0E0E0&
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
      Left            =   45
      MaxLength       =   200
      TabIndex        =   20
      TabStop         =   0   'False
      Top             =   1800
      Width           =   3840
   End
   Begin VB.CommandButton cmdValidar 
      Caption         =   "Validar Ruc SUNAT"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   0
      TabIndex        =   18
      Top             =   5880
      Width           =   1695
   End
   Begin VB.CommandButton cmdTipoCliente 
      Caption         =   "Tipo Cliente"
      Height          =   555
      Left            =   3960
      TabIndex        =   17
      Top             =   4065
      Width           =   1275
   End
   Begin VB.TextBox txtTipoCliente 
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
      Height          =   375
      Left            =   45
      MaxLength       =   200
      TabIndex        =   16
      TabStop         =   0   'False
      Top             =   4185
      Width           =   3840
   End
   Begin VB.TextBox txtEnlace 
      BackColor       =   &H00E0E0E0&
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
      Left            =   40
      MaxLength       =   200
      TabIndex        =   15
      TabStop         =   0   'False
      Top             =   3600
      Width           =   3840
   End
   Begin VB.CommandButton cmdOpcion 
      Caption         =   "Enlace"
      Height          =   555
      Index           =   7
      Left            =   3960
      TabIndex        =   14
      Top             =   3480
      Width           =   1275
   End
   Begin VB.TextBox txtTipoIdentidad 
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
      Height          =   375
      Left            =   45
      MaxLength       =   15
      TabIndex        =   13
      TabStop         =   0   'False
      Top             =   120
      Width           =   3810
   End
   Begin VB.CommandButton cmTipoIdentidad 
      Caption         =   "Tipo Identidad"
      Height          =   555
      Left            =   3960
      TabIndex        =   12
      Top             =   30
      Width           =   1275
   End
   Begin VB.CommandButton cmdOpcion 
      Caption         =   "&Correo"
      Height          =   555
      Index           =   6
      Left            =   3960
      TabIndex        =   11
      Top             =   2895
      Width           =   1275
   End
   Begin VB.TextBox txtCorreo 
      BackColor       =   &H00E0E0E0&
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
      Left            =   40
      MaxLength       =   200
      TabIndex        =   10
      TabStop         =   0   'False
      Top             =   3000
      Width           =   3840
   End
   Begin VB.CommandButton cmdOpcion 
      Caption         =   "Mostrar Visor de Precios"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   555
      Index           =   5
      Left            =   45
      Style           =   1  'Graphical
      TabIndex        =   9
      Top             =   5295
      Visible         =   0   'False
      Width           =   1635
   End
   Begin VB.Frame Frame 
      Height          =   90
      Left            =   30
      TabIndex        =   8
      Top             =   4560
      Width           =   5235
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
      Height          =   555
      Index           =   4
      Left            =   2610
      Picture         =   "frmNuevoCliente.frx":030A
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   5310
      Width           =   1275
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
      Height          =   555
      Index           =   3
      Left            =   3960
      Picture         =   "frmNuevoCliente.frx":040C
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   5310
      Width           =   1275
   End
   Begin VB.CommandButton cmdOpcion 
      Caption         =   "&Dirección"
      Height          =   555
      Index           =   2
      Left            =   3960
      TabIndex        =   5
      Top             =   2340
      Width           =   1275
   End
   Begin VB.CommandButton cmdOpcion 
      Caption         =   "Razón &Social"
      Height          =   555
      Index           =   1
      Left            =   3960
      TabIndex        =   4
      Top             =   1185
      Width           =   1275
   End
   Begin VB.CommandButton cmdOpcion 
      Caption         =   "&Identificador"
      Height          =   555
      Index           =   0
      Left            =   3960
      TabIndex        =   3
      Top             =   600
      Width           =   1275
   End
   Begin VB.TextBox txtRuc 
      BackColor       =   &H00E0E0E0&
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
      Left            =   45
      MaxLength       =   15
      TabIndex        =   2
      TabStop         =   0   'False
      Top             =   720
      Width           =   3840
   End
   Begin VB.TextBox txtRazonSocial 
      BackColor       =   &H00E0E0E0&
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
      Left            =   45
      MaxLength       =   200
      TabIndex        =   1
      TabStop         =   0   'False
      Top             =   1250
      Width           =   3840
   End
   Begin VB.TextBox txtDireccion 
      BackColor       =   &H00E0E0E0&
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
      Left            =   45
      MaxLength       =   200
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   2400
      Width           =   3840
   End
   Begin VB.Label Label1 
      Height          =   375
      Left            =   120
      TabIndex        =   24
      Top             =   4800
      Width           =   2295
   End
   Begin VB.Label lblcondicion 
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
      Left            =   1800
      TabIndex        =   19
      Top             =   6000
      Width           =   3015
   End
End
Attribute VB_Name = "frmNuevoCliente"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim RsTipoIdentidad As ADODB.Recordset
Dim RsTipoCliente As ADODB.Recordset
Dim sCodigoTipoIdentidad As String
Dim sCodigoTipoCliente As String

Dim validaTipoIdentidad As Boolean

'cambio validar DNI
'Dim Isql As String
Dim RsParametroDNI As Recordset


Private Sub cmdUbigeo_Click()
    Dim xCriterio As String
    xCriterio = sCodigo
   Isql = "Select tCodigo as Codigo, tDistrito as Descripcion, tProvincia as Provincia, tDepartamento as Departamento from TUBIGEO order by tCodigo asc"
   
   frmBusca.TipoOperacion = "UBIGEO"
   frmBusca.cboCriterio.Enabled = True
   frmBusca.nPredeterm = 1

   Call ConfGrilla(4, frmBusca.grdGrilla, "Codigo", 2, "Codigo", 1200, 0, 0, "", _
                                          "Distrito", 2, "Descripcion", 1500, 0, 0, "", _
                                          "Provincia", 2, "Provincia", 2500, 0, 0, "", _
                                          "Departamento", 2, "Departamento", 3000, 0, 0, "")
   frmBusca.Show vbModal
   
   If Not wEnter Then
      sCodigo = xCriterio
      Exit Sub
   End If
   
   txtUbigeo.Text = sCodigo
   sCodigo = xCriterio
End Sub

Private Sub cmdValidar_Click()
    Call RucSUNAT(Trim(txtRuc.Text))
End Sub

Private Sub txtRuc_Change()
' validar dni
    Me.txtRuc.Text = Replace(Me.txtRuc, " ", "")
End Sub

Private Sub txtRuc_LostFocus()
    If Trim(txtRuc) = "" Then
     Exit Sub
    End If
End Sub
Private Sub RucSUNAT(Ruc As String)
On Error GoTo fin
    Dim loRUC As vfpsrucperu.vfpsruc
    Set loRUC = New vfpsrucperu.vfpsruc
    
    Dim lcNroRuc As String
 
        lcNroRuc = Ruc
        If loRUC.VFPs_ConsultarRUC(lcNroRuc, False) Then
            'DEVOLVIO LA CONSULTA CORRECTAMENTE
            'PROPIEDADES A CONSULTAR LUEGO DE LA CONSULTA DE RUC
            'loRUC.LCRUC
            txtRazonSocial.Text = loRUC.LCRAZONSOCIAL
            'loRUC.LCTIPOCON
            'loRUC.c
            'loRUC.LCTELEFONO
            'loRUC.LDFECHAINS
            'loRUC.LCESTADO
            lblcondicion.Caption = "Condicion:  " & loRUC.LCCONDICION
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
Private Sub cmdOpcion_Click(Index As Integer)

    Dim xtTipoIdentidad As String
   Dim Numero As Boolean
   Dim te As String

   If sModulo = "ADICION" Then
      frmMozoUsuario.ReseteaTimer
   End If

   Select Case Index
          Case Is = 0 'Ruc
               frmKeyBoard.Tipo = "Cliente"
               frmKeyBoard.txtResultado = txtRuc.Text
               frmKeyBoard.Show vbModal
               
               If Trim(sDescrip) = "" Then
                Exit Sub
               End If
               
                If pais <> "002" Then 'PERU - BOLIVIA
                    If Calcular("Select isnull(nValor,0) As Codigo from vtipoidentidad where Codigo= '" & sCodigoTipoIdentidad & "'", Cn) Then
                    'Consitencia RUC
                    
                            If lLongitud And Len(Trim(sDescrip)) <> nLongitud Then
                               MsgBox "La longitud del Identificador debe ser " & nLongitud, vbCritical, sMensaje
                               wEnter = False
                               Exit Sub
                            ElseIf Not lLongitud And Len(Trim(sDescrip)) < nLongitud Then
                               MsgBox "La longitud del Identificador debe ser mayor igual a " & nLongitud, vbCritical, sMensaje
                               wEnter = False
                               Exit Sub
                            End If

                           validaTipoIdentidad = False
                           validaTipoIdentidad = Calcular("select isnull(nvalor,0) as codigo from vTipoidentidad where Codigo='" & sCodigoTipoIdentidad & "' ", Cn)
                           If validaTipoIdentidad = True Then
                             If Not ValidaRuc(sDescrip) And pais = "000" Then
                               MsgBox "El número ingresado no es válido", vbCritical, sMensaje
                               txtRuc.Text = ""
                               wEnter = False
                               Exit Sub
                             End If
                            End If
                             
                            If xlTipoDocumento = True Then
                                 If Not ValidaRuc(sDescrip) And pais = "000" Then
                                    MsgBox "El número ingresado no es válido", vbCritical, sMensaje
                                    txtRuc.Text = ""
                                    wEnter = False
                                    Exit Sub
                                 End If
                            End If

                    End If
                                        
                Else 'ECUADOR
                
                    'If chkpasaporte.value = 0 Then
                        If Calcular("Select isnull(nValor,0) As Codigo from vtipoidentidad where Codigo= '" & sCodigoTipoIdentidad & "'", Cn) Then
        
                            If Len(Trim(sDescrip)) = 13 Or Len(Trim(sDescrip)) = 10 Then
                                If xlTipoDocumento = True Then
                                    If Not ValidaEcuadorCedulaRuc(sDescrip) Then
                                       MsgBox "Identificador no Válido", vbCritical, sMensaje
                                       wEnter = False
                                       Exit Sub
                                    End If
                                End If
                            Else
                               MsgBox "La longitud del Identificador debe ser 10(Cédula) ó 13(RUC)", vbCritical, sMensaje
                               wEnter = False
                               Exit Sub
                            End If
                            
                        End If
                    'End If
                    
                End If
    
               
                If Val(Calcular("select tIdentidad as Codigo from TCLIENTE where tIdentidad = '" & sDescrip & "'", Cn)) > 0 Then
                   MsgBox "Identificador Repetido", vbCritical, sMensaje
                   wEnter = False
                   Exit Sub
                End If
                xlTipoDocumento = False

                txtRuc.Text = IIf(wEnter, sDescrip, txtRuc.Text)
          
          Case Is = 1 ' Razon social
               frmKeyBoard.txtResultado = txtRazonSocial.Text
               frmKeyBoard.Show vbModal
               txtRazonSocial.Text = IIf(wEnter, sDescrip, txtRazonSocial.Text)
          
          Case Is = 2 ' Direccion
               frmKeyBoard.txtResultado = txtDireccion.Text
               frmKeyBoard.Show vbModal
               txtDireccion.Text = IIf(wEnter, sDescrip, txtDireccion.Text)
               
          Case Is = 3 ' Aceptar
               Dim nCorrela As String
               
               ' cambios para validar DNI
               If RsParametroDNI!lValidaDNI = True Then
               If txtTipoIdentidad.Text = "DNI" Then
               Numero = modProcedimiento.ValidarDNI(LTrim(txtRuc.Text))
                    If Numero = False Then
                    MsgBox "El DNI ingresado no es valido", vbCritical, sMensaje
                    Exit Sub
                    End If
               End If
               End If
               '---------------------------------
               'Chequea Datos
               If txtRuc.Text = "" Then MsgBox "Ingrese el Ruc", vbExclamation, sMensaje: Exit Sub
               If txtRazonSocial = "" Then MsgBox "Ingrese la Razón Social", vbExclamation, sMensaje: Exit Sub
               If txtTipoIdentidad.Text = "" Then MsgBox "Seleccione Tipo de Identidad", vbExclamation, sMensaje: Exit Sub
               If txtTipoIdentidad.Text = "Seleccionar --->" Then MsgBox "Seleccione Tipo de Identidad", vbExclamation, sMensaje: Exit Sub
               
               If sCodigoTipoIdentidad = "" Then MsgBox "Seleccione Tipo de Identidad", vbExclamation, sMensaje: Exit Sub
               
               If lSAP Then
                    If txtTipoCliente.Text = "" Then MsgBox "Seleccione Tipo de Cliente", vbExclamation, sMensaje: Exit Sub
                    If txtTipoCliente.Text = "Seleccionar --->" Then MsgBox "Seleccione Tipo de Cliente", vbExclamation, sMensaje: Exit Sub
                    If sCodigoTipoCliente = "" Then MsgBox "Seleccione Tipo de Cliente", vbExclamation, sMensaje: Exit Sub
               End If
               
               If txtCorreo.Text <> "" Then
                    If Not Validar_Email(txtCorreo.Text) Then
                        MsgBox "El Correo ingresado no es válido", vbCritical, sMensaje
                        wEnter = False
                        Exit Sub
                    End If
               End If
               If lFEpape Then
                    If sCodigoTipoIdentidad = "02" Or sCodigoTipoIdentidad = "01" Then
                        If txtCorreo.Text <> "" Then
                             If Not Validar_Email(txtCorreo.Text) Then
                                 MsgBox "El Correo ingresado no es válido", vbCritical, sMensaje
                                 wEnter = False
                                 Exit Sub
                             End If
                        Else
                            MsgBox "Ingrese un correo Electronico", vbCritical, sMensaje
                        End If
                    End If
               End If

               ' RUC
               If pais <> "002" Then 'PERU - BOLIVIA
                    If sCodigoTipoIdentidad = "02" Then
                         If lLongitud And Len(Trim(txtRuc.Text)) <> nLongitud Then
                            MsgBox "La longitud del Identificador debe ser " & nLongitud, vbCritical, sMensaje
                            wEnter = False
                            Exit Sub
                         ElseIf Not lLongitud And Len(Trim(txtRuc.Text)) < nLongitud Then
                            MsgBox "La longitud del Identificador debe ser mayor igual a " & nLongitud, vbCritical, sMensaje
                            wEnter = False
                            Exit Sub
                         End If
                         
                        validaTipoIdentidad = False
                        validaTipoIdentidad = Calcular("select isnull(nvalor,0) as codigo from vTipoidentidad where Codigo='" & sCodigoTipoIdentidad & "' ", Cn)
                        If validaTipoIdentidad = True Then
                          If Not ValidaRuc(txtRuc.Text) And pais = "000" Then
                            MsgBox "El número ingresado no es válido", vbCritical, sMensaje
                            wEnter = False
                            Exit Sub
                          End If
                         End If
                         
                         If xlTipoDocumento = True Then
                              If Not ValidaRuc(txtRuc.Text) And pais = "000" Then
                                 MsgBox "El número ingresado no es válido", vbCritical, sMensaje
                                 wEnter = False
                                 Exit Sub
                              End If
                         End If
                         
                         xtTipoIdentidad = ""
                    End If
               Else  ' ECUADOR
                    If sCodigoTipoIdentidad = "01" Or sCodigoTipoIdentidad = "02" Then
                        If Len(Trim(txtRuc.Text)) = 13 Or Len(Trim(txtRuc.Text)) = 10 Then
                             If xlTipoDocumento = True Then
                                 If Not ValidaEcuadorCedulaRuc(txtRuc.Text) Then
                                    MsgBox "Identificador no Válido", vbCritical, sMensaje
                                    wEnter = False
                                    Exit Sub
                                 End If
                             End If
                         Else
                            MsgBox "La longitud del Identificador debe ser 10(Cédula) ó 13(RUC)", vbCritical, sMensaje
                            wEnter = False
                            Exit Sub
                         End If
                         
                        'SEGUN SRI
                        If Len(Trim(txtRuc.Text)) = 10 Then
                           xtTipoIdentidad = "01"
                        ElseIf Len(Trim(txtRuc.Text)) = 13 Then
                           xtTipoIdentidad = "02"
                        End If
                        
                    Else
                        xtTipoIdentidad = "03"
                        sCodigoTipoIdentidad = "03"
                    End If
                                          
               End If
               
            
               If frmBusquedaRapida.wAdiciona Then
                  If Val(Calcular("select tIdentidad as Codigo from TCLIENTE where tIdentidad = '" & Apostrofe_v2(txtRuc.Text) & "'", Cn)) > 0 Then
                     MsgBox "Identificador Repetido", vbCritical, sMensaje
                     wEnter = False
                     Exit Sub
                  End If
               
                  'Obtiene el Correlativo
                  nCorrela = Calcular("select Max(tCodigoCliente) as Codigo from TCLIENTE", Cn)
                
                  If IsNull(nCorrela) Or nCorrela = "" Then
                     sCodigo = "00001"
                  Else
                     sCodigo = Lib.Correlativo(nCorrela, 5)
                  End If
                  
                  Isql = "insert into TCLIENTE( " & _
                         "tCodigoCliente, tEmpresa, tIdentidad, tDireccion, tCorreo, tUsuario, tTipoIdentidad, lActivo, tEnlace, tTipoCliente, tubigeo,tUrbanizacion ,fRegistro) " & _
                         "values ('" & sCodigo & "', " & _
                                 " '" & Apostrofe_v2(txtRazonSocial.Text) & "', " & _
                                 " '" & Apostrofe_v2(txtRuc.Text) & "', " & _
                                 " '" & Apostrofe_v2(txtDireccion.Text) & "', " & _
                                 " '" & Apostrofe_v2(txtCorreo.Text) & "', " & _
                                 " '" & sUsuario & "', " & _
                                 " '" & sCodigoTipoIdentidad & "', " & _
                                                          1 & ", " & _
                                 " '" & txtEnlace.Text & "', " & _
                                 " '" & sCodigoTipoCliente & "', " & _
                                 " '" & Me.txtUbigeo.Text & "', " & _
                                 " '" & Me.txtUrbanizacion.Text & "', " & _
                                 " getdate() )"
               Else

                  If Val(Calcular("select tIdentidad as Codigo from TCLIENTE where tIdentidad = '" & txtRuc.Text & "' and tCodigoCliente<>'" & sCodigo & "' ", Cn)) > 0 Then
                     MsgBox "Identificador Repetido", vbCritical, sMensaje
                     wEnter = False
                     Exit Sub
                  End If
                                 
                  Isql = "Update TCLIENTE  SET " & _
                         "tIdentidad='" & Apostrofe_v2(txtRuc.Text) & "', " & _
                         "tEmpresa='" & Apostrofe_v2(txtRazonSocial.Text) & "', " & _
                         "tDireccion='" & Apostrofe_v2(txtDireccion.Text) & "', " & _
                         "tCorreo='" & Apostrofe_v2(txtCorreo.Text) & "', " & _
                         "tTipoIdentidad = '" & sCodigoTipoIdentidad & "', " & _
                         "tTipoCliente = '" & sCodigoTipoCliente & "', " & _
                         "tEnlace = '" & txtEnlace.Text & "', " & _
                         "tUbigeo = '" & Me.txtUbigeo.Text & "', " & _
                         "tUrbanizacion = '" & Me.txtUrbanizacion.Text & "', " & _
                         "fRegistro=getdate() ,lreplica = 1 " & _
                         "where tCodigoCliente='" & sCodigo & "'"
               End If
               Cn.Execute Isql
               wEnter = True
               Unload Me
               
          Case Is = 6 ' Correo
               frmKeyBoard.txtResultado = txtCorreo.Text
               frmKeyBoard.Show vbModal
               txtCorreo.Text = IIf(wEnter, sDescrip, txtCorreo.Text)
               
          Case Is = 7 ' Enlace
               frmKeyBoard.txtResultado = txtEnlace.Text
               frmKeyBoard.Show vbModal
               txtEnlace.Text = IIf(wEnter, sDescrip, txtEnlace.Text)
               
          Case Is = 4 ' Salir
               wEnter = False
               Unload Me
               
          Case Is = 5 ' Mostrar
               Visor txtRazonSocial.Text, txtRuc.Text, nPuerto, "N"
          Case Is = 8 ' Urbanizacion
               frmKeyBoard.txtResultado = txtUrbanizacion.Text
               frmKeyBoard.Show vbModal
               Me.txtUrbanizacion.Text = IIf(wEnter, sDescrip, txtUrbanizacion.Text)
               
   End Select
End Sub

Private Sub cmdTipoCliente_Click()

    RsTipoCliente.MoveNext
    If RsTipoCliente.EOF Then
        RsTipoCliente.MoveFirst
    End If
    sCodigoTipoCliente = RsTipoCliente!codigo
    cmdTipoCliente.Caption = RsTipoCliente!tResumido
    txtTipoCliente.Text = RsTipoCliente!tResumido
    wEnter = True
                              
End Sub

Private Sub cmTipoIdentidad_Click()

    txtRuc.Text = ""
    RsTipoIdentidad.MoveNext
     If RsTipoIdentidad.EOF Then
        RsTipoIdentidad.MoveFirst
     End If
     sCodigoTipoIdentidad = RsTipoIdentidad!codigo
     cmTipoIdentidad.Caption = RsTipoIdentidad!tResumido
     txtTipoIdentidad.Text = RsTipoIdentidad!tResumido
     wEnter = True
                                       
End Sub

Private Sub Form_Initialize()
Set RsTipoIdentidad = New ADODB.Recordset
Set RsTipoCliente = New ADODB.Recordset
End Sub

Private Sub Form_Load()

    ' cambio validar DNI
   Isql = "select lValidaDNI from TPARAMETRO"
   Set RsParametroDNI = Lib.OpenRecordset(Isql, Cn)
    '------------------------------------------

   If sModulo = "ADICION" Then
      frmMozoUsuario.ReseteaTimer
   End If
    If pais = "000" Then
        cmdValidar.Visible = True
    Else
        cmdValidar.Visible = False
    End If

   Set RsTipoIdentidad = Lib.OpenRecordset("select tresumido,codigo from vtipoidentidad where lActivo = 1", Cn)
   Set RsTipoCliente = Lib.OpenRecordset("select tresumido,codigo from vtipogrupocliente where lActivo = 1", Cn)
   
   Select Case pais
    Case "001" 'Bolivia
        cmdOpcion(0).Caption = "&NIT"
    Case Else 'Peru, Ecuador
        cmdOpcion(0).Caption = "&Identificador"
   End Select
   
   Centrar Me
   If frmBusquedaRapida.wAdiciona Then
      cmdOpcion(0).Enabled = True
      Limpiar
      If frmBusquedaRapida.nPredeterm = 2 Then
         txtRazonSocial.Text = frmBusquedaRapida.txtResultado
      Else
         txtRuc.Text = frmBusquedaRapida.txtResultado
      End If
      txtTipoIdentidad.Text = "Seleccionar --->"
        RsTipoIdentidad.MoveNext
        If RsTipoIdentidad.EOF Then
          RsTipoIdentidad.MoveFirst
        End If
        'sCodigoTipoIdentidad = rsTipoIdentidad!codigo
        cmTipoIdentidad.Caption = RsTipoIdentidad!tResumido
        'txtTipoIdentidad.Text = rsTipoIdentidad!tResumido
        
        txtTipoCliente.Text = "Seleccionar --->"
        If RsTipoCliente.RecordCount > 0 Then
            RsTipoCliente.MoveNext
            If RsTipoCliente.EOF Then
              RsTipoCliente.MoveFirst
            End If
            cmdTipoCliente.Caption = RsTipoCliente!tResumido
        End If
        
   Else
      cmdOpcion(0).Enabled = True
      sCodigo = IIf(frmBusquedaRapida.RsGrilla.EOF = True, "", frmBusquedaRapida.RsGrilla!codigo)
      Mostrar
   End If
   
   If lSAP Then
        txtTipoCliente.Visible = True
        cmdTipoCliente.Visible = True
   Else
        txtTipoCliente.Visible = False
        cmdTipoCliente.Visible = False
   End If
   
   
   If nPuerto > 0 Then
      cmdOpcion(5).Visible = True
   End If
   
   If Not lClub Then
    txtEnlace.Visible = False
    cmdOpcion(7).Visible = False
   End If
   
End Sub

Sub Mostrar()
    With frmBusquedaRapida.RsGrilla
        txtRuc.Text = IIf(IsNull(!tIdentidad), "", !tIdentidad)
        txtRazonSocial = IIf(IsNull(!Descripcion), "", !Descripcion)
        txtDireccion = IIf(IsNull(!tDireccion), "", !tDireccion)
        txtCorreo = IIf(IsNull(!tcorreo), "", !tcorreo)
        Me.txtUbigeo = IIf(IsNull(!CodigoUbigeo), "", !CodigoUbigeo)
        Me.txtUrbanizacion = IIf(IsNull(!Urbanizacion), "", !Urbanizacion)
        
        If !TipoIdentidad = "" Then
              txtTipoIdentidad.Text = "Seleccionar --->"
              sCodigoTipoIdentidad = ""
        Else
            txtTipoIdentidad = IIf(IsNull(!TipoIdentidad), "", !TipoIdentidad)
            Me.cmTipoIdentidad.Caption = IIf(IsNull(!TipoIdentidad), "", !TipoIdentidad)
            RsTipoIdentidad.MoveFirst
            RsTipoIdentidad.Find "tresumido='" & Me.cmTipoIdentidad.Caption & "'"
            If Not (RsTipoIdentidad.EOF Or RsTipoIdentidad.BOF) Then
                 sCodigoTipoIdentidad = RsTipoIdentidad!codigo
            End If
        End If
        
        If !TipoCliente = "" Then
              txtTipoCliente.Text = "Seleccionar --->"
              sCodigoTipoCliente = ""
        Else
            txtTipoCliente = IIf(IsNull(!TipoCliente), "", !TipoCliente)
            Me.cmdTipoCliente.Caption = IIf(IsNull(!TipoCliente), "", !TipoCliente)
            RsTipoCliente.MoveFirst
            RsTipoCliente.Find "tresumido='" & Me.cmdTipoCliente.Caption & "'"
            If Not (RsTipoCliente.EOF Or RsTipoCliente.BOF) Then
                 sCodigoTipoCliente = RsTipoCliente!codigo
            End If
        End If
                
        txtEnlace.Text = IIf(IsNull(!tEnlace), "", !tEnlace)
    End With
End Sub

Sub Limpiar()
  Dim Control As Object
  For Each Control In Me.Controls
        If (TypeOf Control Is TextBox) Then
            Control.Text = ""
        End If
    Next Control
End Sub

Private Sub Form_Terminate()
    Set RsTipoIdentidad = Nothing
    Set RsTipoCliente = Nothing
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set frmNuevoCliente = Nothing
End Sub

'cambio de validar DNI
Private Sub txtRuc_GotFocus()
If txtTipoIdentidad.Text = "Seleccionar ---" Then
    MsgBox "Debe colocar un identificador"
    foco
    Exit Sub
End If
End Sub


'cambio de validar DNI
Private Function foco()
Me.txtRazonSocial.SetFocus
End Function


'cambios validar DNI
Private Sub txtRuc_KeyPress(KeyAscii As Integer)
If RsParametroDNI!lValidaDNI = True Then
    If txtTipoIdentidad.Text = "DNI" Then
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
End Sub
