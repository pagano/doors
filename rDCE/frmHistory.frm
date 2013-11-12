VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmHistory 
   Caption         =   "Form1"
   ClientHeight    =   4704
   ClientLeft      =   60
   ClientTop       =   348
   ClientWidth     =   6492
   Icon            =   "frmHistory.frx":0000
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   4704
   ScaleWidth      =   6492
   Begin MSComctlLib.ListView ListView1 
      Height          =   3135
      Left            =   240
      TabIndex        =   0
      Top             =   240
      Width           =   4335
      _ExtentX        =   7641
      _ExtentY        =   5525
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   0
   End
End
Attribute VB_Name = "frmHistory"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public CodeType As Long '1-Folder, 2-Form, 3-Document, 4-AsyncEvent
Public dForm As Object
Public Folder As Object
Public DocId As Long
Public EventKey As String
Public Field As String
Public blnLoaded As Boolean
Public dSession As Object

Private Sub Form_Activate()
    Dim oRcs As Object
    Dim strSQL As String
    Dim li As MSComctlLib.ListItem
    Dim si As MSComctlLib.ListSubItem
    
    On Error GoTo Error
    
    If Not blnLoaded Then
        strSQL = "select * from DCE_HISTORY where CODETYPE = " & CodeType
        If CodeType = 1 Then
            strSQL = strSQL & " and FLD_ID = " & Folder.id & " and SEV_ID = " & Mid(EventKey, 4)
        ElseIf CodeType = 2 Then
            strSQL = strSQL & " and FRM_ID = " & dForm.id & " and SEV_ID = " & Mid(EventKey, 4)
        ElseIf CodeType = 3 Then
            strSQL = strSQL & " and FLD_ID = " & Folder.id & " and DOC_ID = " & DocId
        End If
        strSQL = strSQL & " order by TIMESTAMP desc"
        
        Set oRcs = dSession.Db.OpenRecordset(strSQL)
        Do While Not oRcs.EOF
            With ListView1
                Set li = .ListItems.Add(, , oRcs("TIMESTAMP").Value)
                li.Tag = oRcs("TIMESTAMP").Value
                li.ToolTipText = li.Text
                ' si 1
                Set si = li.ListSubItems.Add(, , oRcs("ACC_NAME").Value)
                si.ToolTipText = si.Text
                si.Tag = oRcs("ACC_ID").Value
                ' si 2
                Set si = li.ListSubItems.Add(, , Len(oRcs("CODE").Value & ""))
                si.ToolTipText = si.Text
                si.Tag = oRcs("CODE").Value & ""
            End With
            oRcs.MoveNext
        Loop
        
        oRcs.Close
        blnLoaded = True
    End If

    Screen.MousePointer = vbNormal
    If ListView1.ListItems.Count = 0 Then MsgBox "no history", vbInformation
    Exit Sub
Error:
    Screen.MousePointer = vbNormal
    MsgBox Err.Description, vbExclamation
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyEscape And Shift = 0 Or _
           KeyCode = vbKeyF4 And Shift = vbCtrlMask Then
        KeyCode = 0
        Unload Me
    End If
End Sub

Private Sub Form_Load()
    Dim ch As MSComctlLib.ColumnHeader
    
    blnLoaded = False
    KeyPreview = True
    
    With ListView1
        .View = lvwReport
        .LabelEdit = lvwManual
        .LabelWrap = True
        .HideSelection = False
        .FullRowSelect = True
        .BorderStyle = ccFixedSingle
        .Appearance = cc3D
        
        Set ch = .ColumnHeaders.Add(, , "Date")
        ch.Width = 2500
        Set ch = .ColumnHeaders.Add(, , "User")
        ch.Width = 2500
        Set ch = .ColumnHeaders.Add(, , "Size")
        ch.Width = 1000
    End With
End Sub

Private Sub Form_Resize()
    With ListView1
        .Top = 1
        .Left = 1
        .Width = ScaleWidth
        .Height = ScaleHeight
    End With
End Sub

Private Sub ListView1_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
    ListViewColumnClick ListView1, ColumnHeader
End Sub

Private Sub ListView1_DblClick()
    Dim li As MSComctlLib.ListItem
    Dim frmCode As frmEditor
    
    Set li = ListView1.SelectedItem
    If Not li Is Nothing Then
        Set frmCode = New frmEditor
        With frmCode
            .Caption = Caption & " (read only)"
            .CodeMax1.Text = li.ListSubItems(2).Tag
            .CodeMax1.ReadOnly = True
            .Show
        End With
    End If
End Sub

Private Sub ListView1_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then ListView1_DblClick
End Sub
