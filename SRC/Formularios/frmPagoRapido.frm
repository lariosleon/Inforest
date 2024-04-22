VERSION 5.00
Begin VB.Form frmPagoRapido 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Pagos"
   ClientHeight    =   8310
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4575
   Icon            =   "frmPagoRapido.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8310
   ScaleWidth      =   4575
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame fraEfectivo 
      Caption         =   "Pago en Efectivo :"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   5430
      Left            =   45
      TabIndex        =   1
      Top             =   0
      Width           =   4500
      Begin VB.CommandButton cmdKey 
         Caption         =   "0"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   855
         Index           =   0
         Left            =   90
         TabIndex        =   28
         Top             =   4425
         Width           =   1725
      End
      Begin VB.CommandButton cmdKey 
         Caption         =   "1"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   855
         Index           =   1
         Left            =   105
         TabIndex        =   27
         Top             =   3570
         Width           =   855
      End
      Begin VB.CommandButton cmdKey 
         Caption         =   "2"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   855
         Index           =   2
         Left            =   960
         TabIndex        =   26
         Top             =   3570
         Width           =   855
      End
      Begin VB.CommandButton cmdKey 
         Caption         =   "3"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   855
         Index           =   3
         Left            =   1815
         TabIndex        =   25
         Top             =   3570
         Width           =   855
      End
      Begin VB.CommandButton cmdKey 
         Caption         =   "4"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   855
         Index           =   4
         Left            =   105
         TabIndex        =   24
         Top             =   2715
         Width           =   855
      End
      Begin VB.CommandButton cmdKey 
         Caption         =   "5"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   855
         Index           =   5
         Left            =   960
         TabIndex        =   23
         Top             =   2715
         Width           =   855
      End
      Begin VB.CommandButton cmdKey 
         Caption         =   "6"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   855
         Index           =   6
         Left            =   1815
         TabIndex        =   22
         Top             =   2715
         Width           =   855
      End
      Begin VB.CommandButton cmdKey 
         Caption         =   "7"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   855
         Index           =   7
         Left            =   105
         TabIndex        =   21
         Top             =   1860
         Width           =   855
      End
      Begin VB.CommandButton cmdKey 
         Caption         =   "8"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   855
         Index           =   8
         Left            =   960
         TabIndex        =   20
         Top             =   1860
         Width           =   855
      End
      Begin VB.CommandButton cmdKey 
         Caption         =   "9"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   855
         Index           =   9
         Left            =   1815
         TabIndex        =   19
         Top             =   1860
         Width           =   855
      End
      Begin VB.CommandButton cmdKey 
         Caption         =   "Esc"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   855
         Index           =   10
         Left            =   2670
         TabIndex        =   18
         Top             =   1860
         Width           =   990
      End
      Begin VB.CommandButton cmdKey 
         Caption         =   "Sup"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   855
         Index           =   11
         Left            =   2670
         TabIndex        =   17
         Top             =   2715
         Width           =   990
      End
      Begin VB.CommandButton cmdKey 
         Caption         =   "Enter"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1715
         Index           =   12
         Left            =   2670
         TabIndex        =   16
         Top             =   3570
         Width           =   990
      End
      Begin VB.CommandButton cmdKey 
         Caption         =   "."
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   855
         Index           =   13
         Left            =   1815
         TabIndex        =   15
         Top             =   4425
         Width           =   855
      End
      Begin VB.TextBox txtTempo 
         Height          =   375
         Left            =   2835
         TabIndex        =   0
         Top             =   4005
         Width           =   690
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "Vuelto"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   3735
         TabIndex        =   34
         Top             =   1395
         Width           =   675
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Abono"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   3735
         TabIndex        =   33
         Top             =   900
         Width           =   690
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Cargo"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   3735
         TabIndex        =   32
         Top             =   405
         Width           =   645
      End
      Begin VB.Label txtCargo 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "0.00"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   420
         Left            =   135
         TabIndex        =   31
         Top             =   315
         Width           =   3510
      End
      Begin VB.Label txtVuelto 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "0.00"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   420
         Left            =   135
         TabIndex        =   30
         Top             =   1305
         Width           =   3510
      End
      Begin VB.Label txtResultado 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "0.00"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00808000&
         Height          =   420
         Left            =   135
         TabIndex        =   29
         Top             =   810
         Width           =   3510
      End
   End
   Begin VB.Frame fraGrilla 
      Caption         =   "Pago con: "
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2775
      Left            =   45
      TabIndex        =   2
      Top             =   5490
      Width           =   4500
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
         Left            =   3015
         Picture         =   "frmPagoRapido.frx":030A
         Style           =   1  'Graphical
         TabIndex        =   35
         Top             =   1935
         Width           =   1335
      End
      Begin VB.CommandButton cmdTarjeta 
         Height          =   630
         Index           =   1
         Left            =   135
         TabIndex        =   14
         Top             =   360
         Width           =   1335
      End
      Begin VB.CommandButton cmdTarjeta 
         Height          =   630
         Index           =   2
         Left            =   1575
         TabIndex        =   13
         Top             =   360
         Width           =   1335
      End
      Begin VB.CommandButton cmdTarjeta 
         Height          =   630
         Index           =   3
         Left            =   3015
         TabIndex        =   12
         Top             =   360
         Width           =   1335
      End
      Begin VB.CommandButton cmdTarjeta 
         Height          =   630
         Index           =   4
         Left            =   135
         TabIndex        =   11
         Top             =   1147
         Width           =   1335
      End
      Begin VB.CommandButton cmdTarjeta 
         Height          =   630
         Index           =   5
         Left            =   1575
         TabIndex        =   10
         Top             =   1147
         Width           =   1335
      End
      Begin VB.CommandButton cmdTarjeta 
         Height          =   630
         Index           =   6
         Left            =   3015
         TabIndex        =   9
         Top             =   1147
         Width           =   1335
      End
      Begin VB.CommandButton cmdTarjeta 
         Height          =   630
         Index           =   7
         Left            =   135
         TabIndex        =   8
         Top             =   1935
         Width           =   1335
      End
      Begin VB.CommandButton cmdTarjeta 
         Height          =   630
         Index           =   8
         Left            =   1575
         TabIndex        =   7
         Top             =   1935
         Width           =   1335
      End
      Begin VB.CommandButton cmdNavegar 
         Height          =   555
         Index           =   0
         Left            =   8940
         Picture         =   "frmPagoRapido.frx":03FC
         Style           =   1  'Graphical
         TabIndex        =   6
         Top             =   240
         Width           =   1215
      End
      Begin VB.CommandButton cmdNavegar 
         Height          =   555
         Index           =   1
         Left            =   8940
         Picture         =   "frmPagoRapido.frx":0CC6
         Style           =   1  'Graphical
         TabIndex        =   5
         Top             =   795
         Width           =   1215
      End
      Begin VB.CommandButton cmdNavegar 
         Height          =   555
         Index           =   2
         Left            =   8940
         Picture         =   "frmPagoRapido.frx":1590
         Style           =   1  'Graphical
         TabIndex        =   4
         Top             =   1350
         Width           =   1215
      End
      Begin VB.CommandButton cmdNavegar 
         Height          =   555
         Index           =   3
         Left            =   8940
         Picture         =   "frmPagoRapido.frx":1E5A
         Style           =   1  'Graphical
         TabIndex        =   3
         Top             =   0
         Width           =   1215
      End
   End
End
Attribute VB_Name = "frmPagoRapido"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim sTarjeta As String
Dim RsTarjeta As Recordset
 
Dim nEN As Double
Dim nEE As Double
Dim nCH As Double
Dim nET As Double
Dim nPT As Double
Dim nDocumento As Double

Dim nAbonoN As Double
  
Dim nSaldo As Double
Dim mTarjeta(8, 3)
Dim sTipoTarjeta As String
Dim sTitulo As String
Dim nIndex As Integer
Dim sCortesia As String
Dim sTipoDocumento As String
Dim sOtroTipoCancelacion As String
Dim sMonDoc As String
Dim nTotalPuntos As Double
Dim sClientePuntos As String
Dim sOtroTipo As String

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
Dim tipoPago As String
Dim tipoTarjeta As String
 
Dim wPunto As Boolean
Dim sTemp As String
 
Private Sub cmdOpcion_Click()
    wEnter = False
    Unload Me
End Sub

Public Sub pagar()
               'Tipo de Cambio
               Dim nCorrelativo As Integer
                  
               If nTC = 0 Then
                  MsgBox "Tipo de Cambio no ingresado", vbCritical, sMensaje
                  Exit Sub
               End If
               
               'JL Correccion
               If lMCPV Then
                    If MultiCajeroOk = False Then
                        Exit Sub
                    End If
               End If
                             
               Screen.MousePointer = vbHourglass
               wEnter = True
               Dim fFechaPago As Date
               Isql = "Update MDOCUMENTO set " & _
                         "tEstadoDocumento = '02', " & _
                         "nAbono = " & nSaldo & "," & _
                         "fPago = getdate()," & _
                         "nVuelto = 0 " & _
                         ",lreplica=1 where tDocumento ='" & sDocumento & "'"

               Cn.Execute Isql
               nCorrelativo = 1
               Dim nEfectivo As Double
               
               Cn.Execute "delete from DPREPAGO where tDocumento='" & sDocumento & "'"
               
                If xTipo = "Modificacion" Then
                    Cn.Execute "delete from DPAGODOCUMENTO where tDocumento='" & sDocumento & "'"
                End If
               
               Select Case tipoPago
                Case "E"
                            'Efectivo Moneda Nacional
                                 Isql = "insert into DPAGODOCUMENTO " & _
                                      "( tDocumento, tCorrelativo, tTurno, tTipoPago, tMoneda, nTipoCambio, nMonto, fRegistro, tUsuario,fDiaContable ) " & _
                                      "Values(   '" & sDocumento & "'," _
                                               & "'" & nCorrelativo & "'," _
                                               & "'" & sTurno & "'," _
                                               & "'01'," _
                                               & "'01'," _
                                               & nTC & ", " _
                                               & nSaldo & ",GETDATE() " _
                                               & ",'" & sUsuario & "','" & Format(obtieneDiaContable, "yyyyMMdd") & "')"
                               Cn.Execute Isql
                 
                Case "T"
                              'Tarjeta
                      Isql = "insert into DPAGODOCUMENTO " & _
                             "( tDocumento, tCorrelativo, tTurno, tTipoPago, tMoneda, tReferencia, nTipoCambio, nMonto, npropina, tTarjeta, tNumero, tFechaVencimiento, fRegistro, tUsuario,fDiaContable ) " & _
                             "Values(    '" & sDocumento & "'," _
                                      & "'" & Trim(str(nCorrelativo)) & "'," _
                                      & "'" & sTurno & "'," _
                                      & "'02'," _
                                      & "'01'," _
                                      & "'', " _
                                      & nTC & ", " _
                                      & nSaldo & ", " _
                                      & 0 & ", " _
                                      & "'" & tipoTarjeta & "', " _
                                      & "'', " _
                                      & "'', " _
                                      & "getdate()," _
                                      & "'" & sUsuario & "','" & Format(obtieneDiaContable, "yyyyMMdd") & "')"
                      Cn.Execute Isql
                End Select
               'Liberacion
               
               If xTipo = "" Then
                  Cn.Execute "Update dbo.DPEDIDO set tFacturado ='P', tCortesia='' where tDocumento ='" & sDocumento & "'"
                  Cn.Execute "UPDATE dbo.MPEDIDO set tEstadoPedido='02', lReplica=1 where tCodigoPedido in (select DISTINCT tCodigoPedido FROM DDOCUMENTO where tDocumento='" & sDocumento & "' )  AND TCODIGOPEDIDO NOT IN (SELECT DISTINCT TCODIGOPEDIDO FROM DPEDIDO WHERE TCODIGOPEDIDO IN(select DISTINCT TCODIGOPEDIDO FROM DDOCUMENTO where tDocumento='" & sDocumento & "') AND ISNULL(TFACTURADO,'') <> 'P')"
                  Cn.Execute "UPDATE dbo.MPEDIDO set fLlegada=getdate(), fEntrega = {fn CURDATE()}  where tCodigoPedido in (select DISTINCT tCodigoPedido FROM DDOCUMENTO where tDocumento='" & sDocumento & "' ) and tTipoPedido='02' and isnull(fLlegada ,0)=0"
                  Cn.Execute "Update dbo.TMESA set tEstadoMesa = '04' where tCodigoMesa in (SELECT DISTINCT TMESA FROM MPEDIDO WHERE TCODIGOPEDIDO IN (SELECT DISTINCT TCODIGOPEDIDO FROM DDOCUMENTO WHERE TDOCUMENTO='" & sDocumento & "') AND TCODIGOPEDIDO NOT IN (SELECT DISTINCT TCODIGOPEDIDO FROM DPEDIDO WHERE TCODIGOPEDIDO IN(select DISTINCT TCODIGOPEDIDO FROM DDOCUMENTO where tDocumento='" & sDocumento & "') AND ISNULL(TFACTURADO,'') <> 'P'))"
                  'Juntar Mesa
                  Cn.Execute "update TMESA set tEstadoMesa='01' where tCodigoMesa in (select tMesa from TPEDIDOMESA where tCodigoPedido='" & sPedido & "')"
               End If
                              
               If CD Then
                    Call ModifcarEstadoDeliveryCabecera(sDocumento)
               End If
               Screen.MousePointer = vbDefault
               wEnter = True
               'Unload Me
               
End Sub

Private Sub cmdTarjeta_Click(Index As Integer)

    If Val(sTemp) < nSaldo Then
       Exit Sub
    End If

    tipoPago = "T"
    tipoTarjeta = ""
    tipoTarjeta = mTarjeta(Index, 1)
                
    If nPuerto > 0 And sFormulario = "CajaRapida" Then
       Dim sss As String
       sss = mTarjeta(Index, 2)
       Visor "Pago con Tarjeta", sss, nPuerto, "N"
    End If

    pagar
    Unload Me
End Sub

Private Sub Form_Load()
   wEnter = False
   tipoPago = ""
   tipoTarjeta = ""
   
   If lCancelacion Then
      cmdOpcion.Enabled = False
   End If
   
   frmPagoRapido.Caption = "Cancelación del Documento " & Format(sDocumento, "@-@@@@@-@@@@@@@@@")
   sMonDoc = "01"
   Limpia
   
    If xTipo = "Modificacion" Then
        txtCargo.Caption = Format(nSaldo, "###,###,###,##0.00")
    Else
        'Obtiene el total del documento para la cancelación!!!!
'        If lModuloPago = "CajaRapida" Then
'            nSaldo = nTotalPR
'        Else
'            nSaldo = nTotalPR
'        End If
        nSaldo = Calcular("select nventa as codigo from MDOCUMENTO where tdocumento= '" & sDocumento & "'", Cn)
        txtCargo.Caption = Format(nSaldo, "###,###,###,##0.00")
    End If
   
   'TARJETAS DE CREDITO
   Isql = "select * from TTARJETACREDITO where nBoton>0 and lActivo = 1 Order by nBoton"
   Set RsTarjeta = Lib.OpenRecordset(Isql, Cn)
   Call AsignaTarjeta(8, RsTarjeta, cmdTarjeta())
   ActivaTarjeta False
   'FIN TARJETAS DE CREDITO

   If sModulo = "ADICION" Then
      frmMozoUsuario.ReseteaTimer
   End If
   
   wPunto = False
   sTemp = "0"
   sDescrip = ""
   txtResultado.Caption = Format(sTemp, "###,###,###,##0.00")

End Sub

Public Sub Limpia()
 nEN = 0
 nEE = 0
 nCH = 0
 nET = 0
 nPT = 0
 nDocumento = 0
 nTotalPuntos = 0
 nAbonoN = 0
 nSaldo = nCargo - nAbonoN
End Sub

Public Sub ActivaTarjeta(Activa As Boolean)
   Dim i As Integer
   For i = 1 To 8
       cmdTarjeta(i).Enabled = Not Activa
   Next i
End Sub

Private Sub Form_Unload(Cancel As Integer)
   Set RsTarjeta = Nothing
   Set frmPagoRapido = Nothing
End Sub

Public Sub AsignaTarjeta(nBoton As Integer, RsAsigna As Recordset, cmdBoton As Object)
   Screen.MousePointer = vbHourglass
   Dim i As Integer
   With RsAsigna
        If .RecordCount > 0 Then
           For i = 1 To nBoton
               .MoveFirst
               .Find "nboton = " & Trim(str(i))
               If Not .EOF Then
                  mTarjeta(i, 1) = !tCodigoTarjeta
                  mTarjeta(i, 2) = IIf(IsNull(!tResumido), "", !tResumido)
                  mTarjeta(i, 3) = IIf(IsNull(!lPinPad), 0, !lPinPad)
                  cmdBoton(i).Visible = True
                  cmdBoton(i).Caption = mTarjeta(i, 2)
                Else
                  cmdBoton(i).Visible = False
                End If
           Next i
        Else
           For i = 1 To nBoton
               cmdBoton(i).Visible = False
           Next i
        End If
   End With
   Screen.MousePointer = vbDefault
End Sub

Private Sub ModifcarEstadoDeliveryCabecera(ByVal qDocumento As String) 'pp
    On Error GoTo ErrorHandler
    'Central Delivery-Motorizado--------------------------------------------------pp
    Isql = "Select Distinct P.tCodigoPedidoCD from DDocumento as D Inner Join MPedido AS P On D.tCodigoPedido = P.tCodigoPedido Where tDocumento = '" + qDocumento + "'"
    Dim RsCodigoPCD As ADODB.Recordset
    Set RsCodigoPCD = Lib.OpenRecordset(Isql, Cn)
    If Not RsCodigoPCD.EOF Then
        If Not IsNull(RsCodigoPCD!tCodigoPedidoCD) Then
'            Call ModifcarEstadoDeliveryCabecera(RsCodigoPCD!tCodigoPedidoCD, "3", txtMotorizado.Caption)
             
            Dim CnCD As Connection
            'Configuración
            Set CnCD = New Connection
            CnCD.Provider = "SQLOLEDB"
            CnCD.CursorLocation = adUseServer
            CnCD.ConnectionString = "User ID=" & sUserName & _
            ";password=" & sUserPassword & _
            ";Data Source=" & sRutaCD & _
            ";Initial Catalog=" & sMDBCD
            CnCD.CommandTimeout = 250
            CnCD.Open
                
            CnCD.Execute "usp_CD_Modificar_EstadoDelivery_Cabecera 4, '" + RsCodigoPCD!tCodigoPedidoCD + "',''"
            CnCD.Close
            
        End If
    End If
    ''''''''''''''''''''''''''''''
    Exit Sub
ErrorHandler:
    MsgBox (err.Description)
End Sub

Private Sub cmdBorra_AfterClick()
   If sModulo = "ADICION" Then
      frmMozoUsuario.ReseteaTimer
   End If

   sTemp = "0"
   wPunto = False
   txtResultado.Caption = Format(sTemp, "###,###,###,##0.00")
End Sub

Private Sub cmdEnter_AfterClick()
   If sModulo = "ADICION" Then
      frmMozoUsuario.ReseteaTimer
   End If

   wEnter = True
   Unload Me
End Sub

Private Sub cmdEsc_AfterClick()
   If sModulo = "ADICION" Then
      frmMozoUsuario.ReseteaTimer
   End If

   wEnter = False
   Unload Me
End Sub

Private Sub cmdkey_Click(Index As Integer)
   If sModulo = "ADICION" Then
      frmMozoUsuario.ReseteaTimer
   End If

    Select Case Index
           Case Is = 10 ' Esc
                wEnter = False
                Unload Me
                
           Case Is = 11 ' Supr
                wPunto = False
                sTemp = "0"
                txtResultado.Caption = Format(sTemp, "###,###,###,##0.00")
                txtTempo.Text = ""
                txtTempo.SetFocus
                
                If nPuerto > 0 And sFormulario = "CajaRapida" Then
                   Visor "Abono :" & sMonN & " " & Right(String(10, " ") & Format(0, "##,##0.00"), 9), "Vuelto:" & sMonN & " " & Right(String(8, " ") & Format(0, "##,##0.00"), 8), nPuerto, "N"
                End If
                txtVuelto.Caption = IIf(Val(sTemp) - nSaldo > 0, Format(Val(sTemp) - nSaldo, "###,###,###,##0.00"), "0.00")
                                               
           Case Is = 12 'Enter
                'If Val(sTemp) < nSaldo And Val(sTemp) <> 0 Then
                If Val(sTemp) < nSaldo Then
                   Exit Sub
                End If
                wEnter = True
                sDescrip = sTemp
                tipoPago = "E" ' EFECTIVO
                pagar
                
                Unload Me
           
           Case Is = 13 'Punto
                If Not wPunto Then
                   sTemp = sTemp & "."
                   wPunto = True
                   txtTempo.SetFocus
                End If
                
                If nPuerto > 0 And sFormulario = "CajaRapida" Then
                   Visor "Abono :" & sMonN & " " & Right(String(10, " ") & Format(Val(sTemp), "##,##0.00"), 9), "Vuelto:" & sMonN & " " & Right(String(8, " ") & Format(IIf(Val(sTemp) - nSaldo > 0, Val(sTemp) - nSaldo, 0), "##,##0.00"), 8), nPuerto, "N"
                End If
                txtVuelto.Caption = IIf(Val(sTemp) - nSaldo > 0, Format(Val(sTemp) - nSaldo, "###,###,###,##0.00"), "0.00")
                
           Case Else
                If (Not wPunto And Len(Trim(sTemp)) >= 16) Or (wPunto And (Len(Right(Trim(sTemp), Trim(InStr(StrReverse(sTemp), "."))))) > 2 And sTipo = "") Or (wPunto And (Len(Right(Trim(sTemp), Trim(InStr(StrReverse(sTemp), "."))))) > 3 And sTipo = "TC") Then
                   Beep
                   txtTempo.SetFocus
                Else
                   sTemp = IIf(sTemp = "0", cmdKey(Index).Caption, sTemp & cmdKey(Index).Caption)
                End If
                txtResultado.Caption = Format(sTemp, "###,###,###,##0.00")
                If nPuerto > 0 And sFormulario = "CajaRapida" Then
                   Visor "Abono :" & sMonN & " " & Right(String(10, " ") & Format(Val(sTemp), "##,##0.00"), 9), "Vuelto:" & sMonN & " " & Right(String(8, " ") & Format(IIf(Val(sTemp) - nSaldo > 0, Val(sTemp) - nSaldo, 0), "##,##0.00"), 8), nPuerto, "N"
                End If
                txtVuelto.Caption = IIf(Val(sTemp) - nSaldo > 0, Format(Val(sTemp) - nSaldo, "###,###,###,##0.00"), "0.00")
                txtVuelto.Caption = IIf(Val(sTemp) - nSaldo > 0, Format(Val(sTemp) - nSaldo, "###,###,###,##0.00"), "0.00")
                txtTempo.SetFocus
    End Select
    

End Sub

Private Sub txtTempo_KeyDown(KeyCode As Integer, Shift As Integer)
   If sModulo = "ADICION" Then
      frmMozoUsuario.ReseteaTimer
   End If

   Select Case KeyCode
   Case 13
        cmdkey_Click (12)
   Case 27
        cmdkey_Click (10)
   Case 46
        cmdkey_Click (11)
   Case 96, 48
        cmdkey_Click (0)
   Case 97, 49
        cmdkey_Click (1)
   Case 98, 50
        cmdkey_Click (2)
   Case 99, 51
        cmdkey_Click (3)
   Case 100, 52
        cmdkey_Click (4)
   Case 101, 53
        cmdkey_Click (5)
   Case 102, 54
        cmdkey_Click (6)
   Case 103, 55
        cmdkey_Click (7)
   Case 104, 56
        cmdkey_Click (8)
   Case 105, 57
        cmdkey_Click (9)
   Case 110, 190
        cmdkey_Click (13)
   End Select
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


