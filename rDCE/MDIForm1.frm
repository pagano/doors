VERSION 5.00
Begin VB.MDIForm MDIForm1 
   BackColor       =   &H8000000C&
   Caption         =   "MDIForm1"
   ClientHeight    =   2400
   ClientLeft      =   132
   ClientTop       =   816
   ClientWidth     =   3744
   Icon            =   "MDIForm1.frx":0000
   LinkTopic       =   "MDIForm1"
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.Menu mnuView 
      Caption         =   "&View"
      Begin VB.Menu mnuViewConnections 
         Caption         =   "&Connections"
      End
   End
   Begin VB.Menu mnuWindow 
      Caption         =   "&Window"
      WindowList      =   -1  'True
      Begin VB.Menu mnuWindowTile 
         Caption         =   "&Tile"
      End
   End
   Begin VB.Menu mnuPopup 
      Caption         =   "Popup"
      Begin VB.Menu mnuPopupEdit 
         Caption         =   "Edit"
      End
      Begin VB.Menu mnuPopupHist 
         Caption         =   "Show History"
      End
   End
   Begin VB.Menu mnuPopupTree 
      Caption         =   "PopupTree"
      Begin VB.Menu mnuPopupTreeRefresh 
         Caption         =   "Refresh"
      End
   End
End
Attribute VB_Name = "MDIForm1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub MDIForm_Load()
    Caption = "Remote DCE v" & App.Major & "." & App.Minor & "." & App.Revision
    
    mnuPopup.Visible = False
    mnuPopupTree.Visible = False
    
    frmConnections.Show
End Sub

Private Sub MDIForm_Unload(Cancel As Integer)
    Dim i As Long
    Dim sAux As String
    Dim ses As Object
    
    If Not Debugging Then
        For Each ses In GcolConnections
            If ses.IsLogged Then ses.Logoff
        Next
    End If
    
    i = 0
    Do
        sAux = GdicURLs.Keys(i)
        If sAux <> "" Then WriteIni "Session", "ServerURL" & i, sAux
        i = i + 1
    Loop Until i = GdicURLs.Count Or i = 10 Or sAux = ""
    Do While i < 10
        WriteIni "Session", "ServerURL" & i, ""
        i = i + 1
    Loop
End Sub

Private Sub mnuViewConnections_Click()
    frmConnections.Show
    frmConnections.SetFocus
End Sub

Private Sub mnuWindowTile_Click()
  MDIForm1.Arrange vbTileHorizontal
End Sub

Private Sub mnuPopupEdit_Click()
    ActiveForm.mnuPopupEditClick
End Sub

Private Sub mnuPopupHist_Click()
    ActiveForm.mnuPopupHistClick
End Sub

Private Sub mnuPopupTreeRefresh_Click()
    ActiveForm.mnuPopupTreeRefreshClick
End Sub

