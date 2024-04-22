VERSION 5.00
Begin VB.Form frmKDSConfiguracion 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "KDS Configuracion"
   ClientHeight    =   3195
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3195
   ScaleWidth      =   4680
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton btnCancelar 
      Caption         =   "Cancelar"
      Height          =   345
      Left            =   3435
      TabIndex        =   10
      Top             =   2685
      Width           =   1035
   End
   Begin VB.CommandButton btnAceptar 
      Caption         =   "Aceptar"
      Height          =   345
      Left            =   2175
      TabIndex        =   9
      Top             =   2685
      Width           =   1035
   End
   Begin VB.CommandButton cmdBump 
      Caption         =   "..."
      Height          =   345
      Left            =   3885
      TabIndex        =   8
      Top             =   2000
      Width           =   615
   End
   Begin VB.CommandButton cmdOrderStatus 
      Caption         =   "..."
      Height          =   345
      Left            =   3870
      TabIndex        =   7
      Top             =   1200
      Width           =   615
   End
   Begin VB.CommandButton cmdOrderInfo 
      Caption         =   "..."
      Height          =   345
      Left            =   3870
      TabIndex        =   6
      Top             =   400
      Width           =   615
   End
   Begin VB.TextBox txtBump 
      Height          =   330
      Left            =   135
      TabIndex        =   5
      Top             =   2000
      Width           =   3645
   End
   Begin VB.TextBox txtOrderStatus 
      Height          =   330
      Left            =   135
      TabIndex        =   4
      Top             =   1200
      Width           =   3645
   End
   Begin VB.TextBox txtOrderInfo 
      Height          =   330
      Left            =   135
      TabIndex        =   3
      Top             =   400
      Width           =   3645
   End
   Begin VB.Label lblBump 
      Caption         =   "Fuente de Notificacion del Bump Bar (BumpNotification)"
      Height          =   210
      Left            =   135
      TabIndex        =   2
      Top             =   1800
      Width           =   4170
   End
   Begin VB.Label lblOrderStatus 
      Caption         =   "Destino de archivo de informacion LS (orderstatus)"
      Height          =   255
      Left            =   135
      TabIndex        =   1
      Top             =   1000
      Width           =   3690
   End
   Begin VB.Label lblOrderInfo 
      Caption         =   "Destino de archivos de las ordenes (orderinfo)"
      Height          =   225
      Left            =   135
      TabIndex        =   0
      Top             =   200
      Width           =   3345
   End
End
Attribute VB_Name = "frmKDSConfiguracion"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub btnAceptar_Click()
    Call KDS_GrabarPath(Me.txtOrderInfo.Text, Me.txtOrderStatus.Text, Me.txtBump.Text)
    Unload Me
End Sub

Private Sub btnCancelar_Click()
    Unload Me
End Sub

Private Sub cmdBump_Click()
     Dim ret As String
    ' Le pasa la leyenda del cuadro de iálogo y el path inicial
    ret = Buscar_Carpeta(" ... Seleccione una carpeta ")
    If (ret <> "") Then
        txtBump.Text = ret
    End If
End Sub

Private Sub cmdOrderInfo_Click()
    Dim ret As String
    ' Le pasa la leyenda del cuadro de iálogo y el path inicial
    ret = Buscar_Carpeta(" ... Seleccione una carpeta ")
    If (ret <> "") Then
        txtOrderInfo.Text = ret
    End If
End Sub

Function Buscar_Carpeta(Optional Titulo As String, Optional Path_Inicial As Variant) As String
On Local Error GoTo errFunction
    Dim objShell As Object
    Dim objFolder As Object
    Dim o_Carpeta As Object
    ' Nuevo objeto Shell.Application
    Set objShell = CreateObject("Shell.Application")
       
    On Error Resume Next
    'Abre el cuadro de diálogo para seleccionar
    Set objFolder = objShell.BrowseForFolder(0, Titulo, 0, Path_Inicial)
       
    ' Devuelve solo el nombre de carpeta
    Set o_Carpeta = objFolder.Self
       
    ' Devuelve la ruta completa seleccionada en el diálogo
    Buscar_Carpeta = o_Carpeta.Path
Exit Function
'Error
errFunction:
    MsgBox err.Description, vbCritical
    Buscar_Carpeta = vbNullString
End Function

Private Sub cmdOrderStatus_Click()
    Dim ret As String
    ' Le pasa la leyenda del cuadro de iálogo y el path inicial
    ret = Buscar_Carpeta(" ... Seleccione una carpeta ")
    If (ret <> "") Then
        txtOrderStatus.Text = ret
    End If
End Sub

Sub KDS_GrabarPath(ByVal OrderInfo As String, ByVal OrderStatus As String, ByVal Bump As String)
    Cn.Execute "USP_KDS_GrabarPath '" + OrderInfo + "', '" + OrderStatus + "', '" + Bump + "'"
End Sub

Private Sub Form_Load()
    Dim RsPath As New Recordset
    Set RsPath = Lib.OpenRecordset("USP_KDS_ObtenerPath", Cn)
    Me.txtOrderInfo.Text = RsPath!tOrderInfo
    Me.txtOrderStatus.Text = RsPath!tOrderStatus
    Me.txtBump.Text = RsPath!tBump
End Sub
