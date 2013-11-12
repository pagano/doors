VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.Form frmConnections 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Connections"
   ClientHeight    =   4200
   ClientLeft      =   36
   ClientTop       =   420
   ClientWidth     =   8424
   Icon            =   "frmConnections.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   4200
   ScaleWidth      =   8424
   Begin VB.CommandButton cmdModify 
      Caption         =   "Modify"
      Height          =   420
      Left            =   6960
      TabIndex        =   4
      Top             =   2160
      Width           =   1212
   End
   Begin VB.CommandButton cmdClose 
      Cancel          =   -1  'True
      Caption         =   "Close"
      Height          =   420
      Left            =   6960
      TabIndex        =   6
      Top             =   3360
      Width           =   1212
   End
   Begin VB.CommandButton cmdRemove 
      Caption         =   "Remove"
      Height          =   420
      Left            =   6960
      TabIndex        =   5
      Top             =   2760
      Width           =   1212
   End
   Begin VB.CommandButton cmdExplore 
      Caption         =   "Explore"
      Default         =   -1  'True
      Height          =   420
      Left            =   6960
      TabIndex        =   1
      Top             =   360
      Width           =   1212
   End
   Begin VB.CommandButton cmdAddFile 
      Caption         =   "Add File"
      Height          =   420
      Left            =   6960
      TabIndex        =   3
      Top             =   1560
      Width           =   1212
   End
   Begin VB.CommandButton cmdAddInstance 
      Caption         =   "Add Instance"
      Height          =   420
      Left            =   6960
      TabIndex        =   2
      Top             =   960
      Width           =   1212
   End
   Begin MSComctlLib.ListView lstConnections 
      Height          =   3732
      Left            =   240
      TabIndex        =   0
      Top             =   240
      Width           =   6492
      _ExtentX        =   11451
      _ExtentY        =   6583
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      FullRowSelect   =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   0
   End
End
Attribute VB_Name = "frmConnections"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdAddInstance_Click()
    Dim oSession As Object
    Dim frmLogon1 As frmLogon
    Dim frmExplorer1 As frmExplorer
    
    Set oSession = CreateObject("dapihttp.Session")
    
    Set frmLogon1 = New frmLogon
    Set frmLogon1.dSession = oSession
    frmLogon1.Show vbModal
    
    If oSession.IsConnected Then
        GcolConnections.Add oSession
    
        Set frmExplorer1 = New frmExplorer
        Set frmExplorer1.dSession = oSession
        GcolExplorers.Add frmExplorer1
    
        If Not GdicURLs.Exists(oSession.ServerURL) Then GdicURLs.Add oSession.ServerURL, Empty
    End If

    lstConnections.SetFocus
    RefreshList
End Sub

Private Sub cmdClose_Click()
    Me.Hide
End Sub

Private Sub cmdExplore_Click()
    Dim li As MSComctlLib.ListItem
    Dim i As Long
    
    Screen.MousePointer = vbHourglass
    
    Set li = lstConnections.SelectedItem
    If Not li Is Nothing Then
        i = Mid(li.Key, 4)
        With GcolExplorers(i)
            If .dSession.IsLogged Then
                .Show
                If .WindowState = vbMinimized Then .WindowState = vbNormal
                .SetFocus
            Else
                cmdModify_Click
            End If
        End With
    End If
    
    Screen.MousePointer = vbNormal
End Sub

Private Sub cmdModify_Click()
    Dim li As MSComctlLib.ListItem
    Dim oSession As Object
    Dim frmLogon1 As frmLogon
    
    Screen.MousePointer = vbHourglass
    
    Set li = lstConnections.SelectedItem
    If Not li Is Nothing Then
        Set oSession = GcolConnections(CLng(Mid(li.Key, 4)))
        Set frmLogon1 = New frmLogon
        Set frmLogon1.dSession = oSession
        frmLogon1.Show vbModal
        
        RefreshList
    End If
End Sub

Private Sub cmdRemove_Click()
    Dim li As MSComctlLib.ListItem
    Dim i As Long
    
    Set li = lstConnections.SelectedItem
    If Not li Is Nothing Then
        i = Mid(li.Key, 4)
        With GcolExplorers(i)
            If .dSession.IsLogged Then .dSession.Logoff
            Unload GcolExplorers(i)
            GcolExplorers.Remove i
            GcolConnections.Remove i
            lstConnections.ListItems.Remove lstConnections.SelectedItem.Index
        End With
    End If
End Sub

Private Sub Form_GotFocus()
    If Me.Visible Then lstConnections.SetFocus
End Sub

Private Sub Form_Load()
    With lstConnections
        .View = lvwReport
        .MultiSelect = True
    End With
    RefreshList
End Sub

Public Sub RefreshList()
    Dim obj As Object
    Dim i As Long
    Dim li As ListItem
    Dim ch As ColumnHeader
    
    Screen.MousePointer = vbHourglass
    
    With lstConnections
        .ListItems.Clear
        .ColumnHeaders.Clear
        .Sorted = False

        Set ch = .ColumnHeaders.Add(, , "ServerURL")
        ch.Width = 3500
        Set ch = .ColumnHeaders.Add(, , "Instance")
        ch.Width = 1500
        Set ch = .ColumnHeaders.Add(, , "Login")
        ch.Width = 1200

        i = 1
        Do While i <= GcolConnections.Count
            Set obj = GcolConnections.Item(i)
            If obj.ServerURL <> "" Then
                Set li = .ListItems.Add(, "ID=" & i, obj.ServerURL)
                If obj.IsLogged Then
                    li.ListSubItems.Add , , obj.InstanceName
                    li.ListSubItems.Add , , obj.LoggedUser.Login
                Else
                    li.ListSubItems.Add , , "-"
                    li.ListSubItems.Add , , "-"
                End If
                i = i + 1
            Else
                GcolConnections.Remove i
                GcolExplorers.Remove i
            End If
        Loop
    End With

    Screen.MousePointer = vbNormal
End Sub

Private Sub lstConnections_DblClick()
    cmdExplore_Click
End Sub

Private Sub lstConnections_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then lstConnections_DblClick
End Sub


