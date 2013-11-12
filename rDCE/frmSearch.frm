VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.Form frmSearch 
   Caption         =   "Search"
   ClientHeight    =   5508
   ClientLeft      =   60
   ClientTop       =   348
   ClientWidth     =   7236
   Icon            =   "frmSearch.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   5508
   ScaleWidth      =   7236
   WindowState     =   2  'Maximized
   Begin MSComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   372
      Left            =   0
      TabIndex        =   3
      Top             =   5136
      Width           =   7236
      _ExtentX        =   12764
      _ExtentY        =   656
      Style           =   1
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   1
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
         EndProperty
      EndProperty
   End
   Begin VB.CommandButton cmdSearch 
      Caption         =   "Search"
      Default         =   -1  'True
      Height          =   350
      Left            =   4080
      TabIndex        =   1
      Top             =   120
      Width           =   1000
   End
   Begin VB.TextBox txtSearch 
      Height          =   300
      Left            =   360
      TabIndex        =   0
      Top             =   240
      Width           =   3255
   End
   Begin MSComctlLib.ListView ListView1 
      Height          =   3135
      Left            =   360
      TabIndex        =   2
      Top             =   960
      Width           =   5775
      _ExtentX        =   10181
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
Attribute VB_Name = "frmSearch"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public dSession As Object

Dim domFolders As Object
Dim bSearching As Boolean
Dim bCancelSearch As Boolean

Public ParentExplorer As frmExplorer

Private Sub cmdSearch_Click()
    Dim dom As Object
    Dim node As Object
    Dim strForms As String
    Dim strFolders As String
    Dim strRoots As String
    Dim oRcs As Object
    Dim strSQL As String
    Dim li As ListItem
    Dim domForms As Object
    Dim oMaster As Object
    Dim frm As Object
    Dim sCols As String
    Dim sCodeCol As String
    
    On Error GoTo Error
    
    If txtSearch.Text = "" Then
        MsgBox "specify a pattern", vbExclamation
        txtSearch.SetFocus
        Exit Sub
    End If
    
    Caption = "Search - " & dSession.InstanceName & " -> " & txtSearch.Text
    
    cmdSearch.Enabled = False
    bCancelSearch = False
    bSearching = True

    With ListView1
        .ListItems.Clear
        .Sorted = False
    End With
    
    DoEvents
    
    
    ' Eventos de Forms
        
    StatusBar1.SimpleText = "Searching in Form Events..."
    DoEvents
    If bCancelSearch Then GoTo CerrarTodo
    
    strForms = ""
    Set domForms = dSession.FormsList
    For Each node In domForms.documentElement.childNodes
        strForms = strForms & "," & node.getAttribute("id")
    Next
    If strForms <> "" Then strForms = Mid(strForms, 2)

    If strForms <> "" Then
        strSQL = "select SYS_SEV_FRM.SEV_ID, SYS_SEV_FRM.FRM_ID, SYS_SYNCEVENTS.NAME " & _
            "from SYS_SEV_FRM, SYS_SYNCEVENTS " & _
            "where SYS_SEV_FRM.SEV_ID = SYS_SYNCEVENTS.SEV_ID " & _
            "and FRM_ID in (" & strForms & ") and CODE like " & _
            dSession.Db.SqlEncode("%" & txtSearch.Text & "%", 1)
        
        Set oRcs = dSession.Db.OpenRecordset(strSQL)
        Do While Not oRcs.EOF
            Set li = ListView1.ListItems.Add(, , "/Forms/" & FormName(oRcs("FRM_ID").Value, domForms))
            li.ListSubItems.Add , , oRcs("NAME").Value
            li.ListSubItems.Add , , "Form_Event"
            li.ListSubItems.Add , , oRcs("FRM_ID").Value
            li.ListSubItems.Add , , oRcs("SEV_ID").Value
            oRcs.MoveNext
        Loop
        oRcs.Close
    End If
    DoEvents
    If ListView1.ListItems.Count > 0 Then ListView1.SetFocus
    
    
    ' Eventos de Folders
    
    StatusBar1.SimpleText = "Searching in Folder Events..."
    DoEvents
    If bCancelSearch Then GoTo CerrarTodo
        
    strFolders = ""
    Set domFolders = dSession.FoldersTree
    For Each node In domFolders.selectNodes("//d:folder")
        strFolders = strFolders & "," & node.getAttribute("id")
    Next
    If strFolders <> "" Then strFolders = Mid(strFolders, 2)
    
    If strFolders <> "" Then
        strSQL = "select SYS_SEV_FLD.SEV_ID, SYS_SEV_FLD.FLD_ID, SYS_SYNCEVENTS.NAME " & _
            "from SYS_SEV_FLD, SYS_SYNCEVENTS " & _
            "where SYS_SEV_FLD.SEV_ID = SYS_SYNCEVENTS.SEV_ID " & _
            "and FLD_ID in (" & strFolders & ") and CODE like " & _
            dSession.Db.SqlEncode("%" & txtSearch.Text & "%", 1)
    
        Set oRcs = dSession.Db.OpenRecordset(strSQL)
        Do While Not oRcs.EOF
            Set li = ListView1.ListItems.Add(, , FolderPath(oRcs("FLD_ID").Value))
            li.ListSubItems.Add , , oRcs("NAME").Value
            li.ListSubItems.Add , , "Folder_Event"
            li.ListSubItems.Add , , oRcs("FLD_ID").Value
            li.ListSubItems.Add , , oRcs("SEV_ID").Value
            oRcs.MoveNext
        Loop
        oRcs.Close
    End If
    DoEvents
    If ListView1.ListItems.Count > 0 Then ListView1.SetFocus

    
    ' Eventos Asincronos de Folders
    
    StatusBar1.SimpleText = "Searching in Folder AsyncEvents..."
    DoEvents
    If bCancelSearch Then GoTo CerrarTodo
            
    If strFolders <> "" Then
        Set oMaster = dSession.MasterDb
        
        strSQL = "select EVN_ID, FLD_ID " & _
            "from SYS_EVENTS " & _
            "where INS_ID = " & dSession.InstanceId & _
            " and FLD_ID in (" & strFolders & ") and CODE like " & _
            oMaster.SqlEncode("%" & txtSearch.Text & "%", 1)
    
        Set oRcs = oMaster.OpenRecordset(strSQL)
        Do While Not oRcs.EOF
            Set li = ListView1.ListItems.Add(, , FolderPath(oRcs("FLD_ID").Value))
            li.ListSubItems.Add , , "AsyncEvent " & oRcs("EVN_ID")
            li.ListSubItems.Add , , "Folder_AsyncEvent"
            li.ListSubItems.Add , , oRcs("FLD_ID").Value
            li.ListSubItems.Add , , oRcs("EVN_ID").Value
            oRcs.MoveNext
        Loop
        oRcs.Close
    End If
    DoEvents
    If ListView1.ListItems.Count > 0 Then ListView1.SetFocus

    
    ' CodeLib
    
    StatusBar1.SimpleText = "Searching in CodeLibs..."
    DoEvents
    If bCancelSearch Then GoTo CerrarTodo
    
    strRoots = ""
    Set dom = dSession.FoldersList
    For Each node In dom.documentElement.childNodes
        strRoots = strRoots & "," & node.getAttribute("id")
    Next
    If strRoots <> "" Then strRoots = Mid(strRoots, 2)
    
    If strRoots <> "" Then
        Set dom = dSession.Forms("F89ECD42FAFF48FDA229E4D5C5F433ED").Search(strRoots, "DOC_ID,FLD_ID,NAME", "CODE like " & dSession.Db.SqlEncode("%" & txtSearch.Text & "%", 1))
        For Each node In dom.documentElement.childNodes
            Set li = ListView1.ListItems.Add(, , FolderPath(node.getAttribute("fld_id")))
            li.ListSubItems.Add , , node.getAttribute("name")
            li.ListSubItems.Add , , "CodeLib"
            li.ListSubItems.Add , , node.getAttribute("doc_id")
        Next
    End If
    DoEvents
    If ListView1.ListItems.Count > 0 Then ListView1.SetFocus
    
    
    ' Controles
    
    StatusBar1.SimpleText = "Searching in Controls..."
    DoEvents
    If bCancelSearch Then GoTo CerrarTodo
    
    If strRoots <> "" Then
        Set dom = dSession.Forms("EAC99A4211204E1D8EEFEB8273174AC4").Search(strRoots, "DOC_ID,FLD_ID,NAME", "SCRIPTBEFORERENDER like " & dSession.Db.SqlEncode("%" & txtSearch.Text & "%", 1))
        For Each node In dom.documentElement.childNodes
            Set li = ListView1.ListItems.Add(, , FolderPath(node.getAttribute("fld_id")))
            li.ListSubItems.Add , , node.getAttribute("name")
            li.ListSubItems.Add , , "Controls"
            li.ListSubItems.Add , , node.getAttribute("doc_id")
        Next
    End If
    DoEvents
    If ListView1.ListItems.Count > 0 Then ListView1.SetFocus

    
    ' Forms configurados con DCE_HasCode en 1
    
    StatusBar1.SimpleText = "Searching in Custom Codes..."
    DoEvents
    If bCancelSearch Then GoTo CerrarTodo
    
    strSQL = "select OBJ_ID from SYS_PROPERTIES " & _
        "where OBJ_TYPE = 1 and ACC_ID = -1 and NAME = 'DCE_HASCODE' and VALUE like '1'"
    Set oRcs = dSession.Db.OpenRecordset(strSQL)
        
    Do While Not oRcs.EOF
        Set frm = dSession.Forms(oRcs(0).Value)
        sCols = frm.Properties("DCE_ListColumns").Value
        sCodeCol = frm.Properties("DCE_CodeColumn").Value
        Set dom = frm.Search(strRoots, "DOC_ID,FLD_ID," & sCols, sCodeCol & " like " & dSession.Db.SqlEncode("%" & txtSearch.Text & "%", 1))
        For Each node In dom.documentElement.childNodes
            Set li = ListView1.ListItems.Add(, , FolderPath(node.getAttribute("fld_id")))
            li.ListSubItems.Add , , node.Attributes(2).Value
            li.ListSubItems.Add , , frm.Name
            li.ListSubItems.Add , , node.getAttribute("doc_id")
        Next
        DoEvents
        If ListView1.ListItems.Count > 0 Then ListView1.SetFocus
        
        oRcs.MoveNext
    Loop
    oRcs.Close
        
    If ListView1.ListItems.Count > 0 Then
        ListView1.SetFocus
    Else
        MsgBox "No results", vbInformation
        txtSearch.SetFocus
    End If
        
    GoTo CerrarTodo
    
    Exit Sub
Error:
    ErrDisplay Err
    Resume CerrarTodo
CerrarTodo:
    StatusBar1.SimpleText = "Iddle..."
    cmdSearch.Enabled = True
    bSearching = False
    bCancelSearch = False
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyEscape Then
        If bSearching Then
            bCancelSearch = True
        Else
            Unload Me
        End If
    End If
End Sub

Private Sub Form_Load()
    Dim ch As ColumnHeader
    
    Caption = "Search - " & dSession.InstanceName
    
    With ListView1
        .View = lvwReport
        .LabelEdit = lvwManual
        .FullRowSelect = True
        
        Set ch = .ColumnHeaders.Add(, , "Path")
        ch.Width = 6000
        Set ch = .ColumnHeaders.Add(, , "Item")
        ch.Width = 2200
        Set ch = .ColumnHeaders.Add(, , "Type")
        ch.Width = 1700
    End With

    StatusBar1.Height = 300
    StatusBar1.SimpleText = "Iddle..."
    
    bSearching = False
    bCancelSearch = False
End Sub

Private Function FormName(ByRef FormId As Long, ByRef FormsDom As Object)
    Dim node As Object
    
    Set node = FormsDom.selectSingleNode("/d:root/d:item[@id='" & FormId & "']")
    If Not node Is Nothing Then FormName = node.getAttribute("name")
End Function

Private Function FolderPath(ByRef FolderId As Long)
    Dim node As Object
    Dim strPath As String
    Dim strAux As String
    
    strPath = ""
    Set node = domFolders.selectSingleNode("//d:folder[@id='" & FolderId & "']")
    If Not node Is Nothing Then
        Do While node.nodeName <> "root"
            strAux = node.getAttribute("description") & ""
            If strAux = "" Then strAux = node.getAttribute("name") & ""
            strPath = "/" & strAux & strPath
            Set node = node.parentNode
        Loop
    End If
    FolderPath = strPath
End Function

Private Sub Form_Resize()
    If WindowState <> vbMinimized Then
        With txtSearch
            .Top = 250
            .Left = 250
            .Width = ScaleWidth - 1750
        End With
        
        With cmdSearch
            .Top = 225
            .Left = ScaleWidth - 1250
        End With
        
        With ListView1
            .Left = 1
            .Top = 800
            .Width = ScaleWidth
            .Height = ScaleHeight - 800
        End With
    End If
End Sub

Private Sub ListView1_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
    ListViewColumnClick ListView1, ColumnHeader
End Sub

Private Sub ListView1_DblClick()
    Dim li As ListItem
    Dim strType As String
    Dim frm As Object
    Dim fld As Object
    Dim doc As Object
    Dim frmCode As frmEditor
    Dim strKey As String
    Dim sCodeCol As String
    
    Set li = ListView1.SelectedItem
    If Not li Is Nothing Then
        strType = li.ListSubItems(2).Text
        
        If strType = "Form_Event" Then
            Set frm = dSession.Forms(li.ListSubItems(3).Text)
            strKey = "ID=" & li.ListSubItems(4).Text
            Set frmCode = New frmEditor
            With frmCode
                Set .dSession = dSession
                .Caption = "Edit - " & dSession.InstanceName & "://Forms/" & frm.Name & " -> " & li.ListSubItems(1).Text
                Set .ParentExplorer = ParentExplorer
                .CodeMax1.Text = frm.Events(strKey).code
                .CodeType = 2
                Set .dForm = frm
                .EventKey = strKey
                .Show
            End With
            
        ElseIf strType = "Folder_Event" Then
            Set fld = dSession.FoldersGetFromId(li.ListSubItems(3).Text)
            strKey = "ID=" & li.ListSubItems(4).Text
            Set frmCode = New frmEditor
            With frmCode
                Set .dSession = dSession
                .Caption = "Edit - " & dSession.InstanceName & ":/" & li.Text & " -> " & li.ListSubItems(1).Text
                Set .ParentExplorer = ParentExplorer
                .CodeMax1.Text = fld.Events(strKey).code
                .CodeType = 1
                Set .Folder = fld
                .EventKey = strKey
                .Show
            End With

        ElseIf strType = "Folder_AsyncEvent" Then
            Set fld = dSession.FoldersGetFromId(li.ListSubItems(3).Text)
            strKey = "ID=" & li.ListSubItems(4).Text
            Set frmCode = New frmEditor
            With frmCode
                Set .dSession = dSession
                .Caption = "Edit - " & dSession.InstanceName & ":/" & li.Text & " -> " & li.ListSubItems(1).Text
                Set .ParentExplorer = ParentExplorer
                .CodeMax1.Text = fld.AsyncEvents(strKey).code
                .CodeType = 4
                Set .Folder = fld
                .EventKey = strKey
                .Show
            End With

        Else
            Set doc = dSession.DocumentsGetFromId(li.ListSubItems(3).Text)
            Set frmCode = New frmEditor
            With frmCode
                Set .dSession = dSession
                .Caption = "Edit - " & dSession.InstanceName & ":/" & li.Text & " -> " & li.ListSubItems(1).Text
                Set .ParentExplorer = ParentExplorer
                If strType = "CodeLib" Then
                    .CodeMax1.Text = doc("CODE").Value
                    .Field = "code"
                ElseIf strType = "Controls" Then
                    .CodeMax1.Text = doc("SCRIPTBEFORERENDER").Value
                    .Field = "scriptbeforerender"
                Else
                    sCodeCol = doc.Form.Properties("DCE_CodeColumn")
                    .CodeMax1.Text = doc(sCodeCol).Value
                    .Field = sCodeCol
                End If
                .CodeType = 3
                Set .Folder = doc.Parent
                .DocId = doc.id
                .Show
            End With
        End If
    
        If Not frmCode Is Nothing Then
            With frmCode.CodeMax1
                ' Posicionar el editor
                .ExecuteCmd cmCmdSetFindText, txtSearch.Text
                .ExecuteCmd cmCmdFindNext
            End With
        End If
    End If
End Sub

Private Sub ListView1_GotFocus()
    cmdSearch.Default = False
End Sub

Private Sub ListView1_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then ListView1_DblClick
End Sub

Private Sub txtSearch_GotFocus()
    cmdSearch.Default = True
End Sub
