VERSION 5.00
Object = "{C4847593-972C-11D0-9567-00A0C9273C2A}#2.1#0"; "CRViewer.dll"
Begin VB.Form frmEmite 
   Caption         =   "Vista Previa"
   ClientHeight    =   8460
   ClientLeft      =   825
   ClientTop       =   1260
   ClientWidth     =   9435
   Icon            =   "frmEmite.frx":0000
   LinkTopic       =   "Form3"
   ScaleHeight     =   8460
   ScaleWidth      =   9435
   WindowState     =   2  'Maximized
   Begin CRVIEWERLibCtl.CRViewer CRViewer 
      Height          =   8310
      Left            =   45
      TabIndex        =   0
      Top             =   45
      Width           =   9330
      DisplayGroupTree=   -1  'True
      DisplayToolbar  =   -1  'True
      EnableGroupTree =   -1  'True
      EnableNavigationControl=   -1  'True
      EnableStopButton=   -1  'True
      EnablePrintButton=   -1  'True
      EnableZoomControl=   -1  'True
      EnableCloseButton=   -1  'True
      EnableProgressControl=   -1  'True
      EnableSearchControl=   -1  'True
      EnableRefreshButton=   -1  'True
      EnableDrillDown =   -1  'True
      EnableAnimationControl=   -1  'True
      EnableSelectExpertControl=   0   'False
      EnableToolbar   =   -1  'True
      DisplayBorder   =   -1  'True
      DisplayTabs     =   -1  'True
      DisplayBackgroundEdge=   -1  'True
      SelectionFormula=   ""
      EnablePopupMenu =   -1  'True
   End
End
Attribute VB_Name = "frmEmite"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub CRViewer_DownloadStarted(ByVal loadingType As CRVIEWERLibCtl.CRLoadingType)
   Screen.MousePointer = vbDefault
End Sub

Private Sub Form_Resize()
    CRViewer.Top = 0
    CRViewer.Left = 0
    CRViewer.Height = ScaleHeight
    CRViewer.Width = ScaleWidth
End Sub

Private Sub Form_Unload(Cancel As Integer)
   Set frmEmite = Nothing
End Sub
