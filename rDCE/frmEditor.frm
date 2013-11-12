VERSION 5.00
Object = "{BCA00000-0F85-414C-A938-5526E9F1E56A}#4.0#0"; "cmax40.dll"
Begin VB.Form frmEditor 
   Caption         =   "Form1"
   ClientHeight    =   2904
   ClientLeft      =   60
   ClientTop       =   348
   ClientWidth     =   3876
   Icon            =   "frmEditor.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   2904
   ScaleWidth      =   3876
   WindowState     =   2  'Maximized
   Begin CodeMax4Ctl.CodeMax CodeMax1 
      Height          =   1695
      Left            =   240
      OleObjectBlob   =   "frmEditor.frx":058A
      TabIndex        =   0
      Top             =   240
      Width           =   2535
   End
End
Attribute VB_Name = "frmEditor"
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
Public CodeChanged As Boolean
Public ParentExplorer As frmExplorer
Public dSession As Object

Private Sub CodeMax1_Change()
    If Not CodeChanged Then
        CodeChanged = True
        Caption = "* " & Caption
    End If
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyS And Shift = vbCtrlMask Then
        KeyCode = 0
        Save
    ElseIf KeyCode = vbKeyEscape And Shift = 0 Or _
           KeyCode = vbKeyF4 And Shift = vbCtrlMask Then
        KeyCode = 0
        Unload Me
    End If
End Sub

Private Sub Form_Load()
    Dim bAdd As Boolean
    Dim lang As CodeMax4Ctl.Language
    Dim i As Long
    Dim sCMaxVersion As String
    Dim lCMaxRevision As Long
    
    With CodeMax1
        .NormalizeCase = False
        .DisplayLeftMargin = False
        .Font.Size = 10
        .LineNumbering = True
        bAdd = True
        For i = 0 To CodeMaxGlobals.Languages.Count - 1
            Set lang = CodeMaxGlobals.Languages(i)
            If lang.Name = "VBScript" Then
                bAdd = False
                Exit For
            End If
        Next
        If bAdd Then
            Set lang = New CodeMax4Ctl.Language
            lang.LoadXmlDefinition App.Path & "\vbscript.lng"
            lang.Register
        End If
        CodeMax1.Language = lang
    End With

    CodeChanged = False
End Sub

Private Sub Form_Resize()
    On Error Resume Next
    If WindowState <> vbMinimized Then
        With CodeMax1
            .Top = 1
            .Left = 1
            .Height = ScaleHeight
            .Width = ScaleWidth
        End With
    End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Dim resp As VbMsgBoxResult
    
    If CodeChanged Then
        resp = MsgBox("wanna save?", vbYesNoCancel + vbQuestion)
        If resp = vbYes Then
            Save
        ElseIf resp = vbCancel Then
            Cancel = 1
        End If
    End If
End Sub

Sub Save()
    Dim sbH As Object
    Dim doc As Object
    Dim fld As Object
    Dim frm As Object
    Dim evn As Object
    Dim sCode As String
    Dim Args As Variant
    
    On Error GoTo Error
    
    If CodeChanged Then
        
        Set sbH = dSession.ConstructNewSqlBuilder
        sbH.Add "TIMESTAMP", Now, 2
        sbH.Add "ACC_ID", dSession.LoggedUser.id, 3
        sbH.Add "ACC_NAME", dSession.LoggedUser.Name, 1
        sbH.Add "CODETYPE", CodeType, 3
        sbH.Add "CODE", "?", 0
        
        Select Case CodeType
            
            Case 1
                'FolderEvents
                Set fld = dSession.FoldersGetFromId(Folder.id)
                Set evn = fld.Events("ID=" & Mid(EventKey, 4))
                evn.code = CodeMax1.Text
                fld.Save
                sbH.Add "FLD_ID", Folder.id, 3
                sbH.Add "SEV_ID", evn.id, 3
            
            Case 2
                'FormEvents
                Set frm = dSession.Forms(dForm.id)
                Set evn = frm.Events("ID=" & Mid(EventKey, 4))
                evn.code = CodeMax1.Text
                frm.Save
                sbH.Add "FRM_ID", dForm.id, 3
                sbH.Add "SEV_ID", Mid(EventKey, 4), 3
            
            Case 3
                'ControlsCode
                Set doc = dSession.DocumentsGetFromId(DocId)
                doc.Fields(Field).Value = CodeMax1.Text
                doc.Save
                dSession.ClearAllCustomCache
                dSession.ClearObjectModelCache "ComCodeLibCache"
                sbH.Add "FLD_ID", Folder.id, 3
                sbH.Add "DOC_ID", DocId, 3
        
            Case 4
                'AsyncEvents
                Set fld = dSession.FoldersGetFromId(Folder.id)
                Set evn = fld.AsyncEvents("ID=" & Mid(EventKey, 4))
                evn.code = CodeMax1.Text
                fld.Save
                sbH.Add "FLD_ID", Folder.id, 3
                sbH.Add "SEV_ID", evn.id, 3
        
        End Select
        
        CodeChanged = False
        Caption = Mid(Caption, 3)
        
        ' Inserta en DCE_HISTORY
        
        sCode = ""
        sCode = sCode & "Set oCmd = dSession.ConstructNewADODBCommand" & vbCrLf
        sCode = sCode & "oCmd.CommandType = 1" & vbCrLf
        sCode = sCode & "Set oPar = oCmd.CreateParameter(""CODE_VALUE"", 201, 1)" & vbCrLf
        sCode = sCode & "oPar.Size = Len(CStr(Arg(2)))" & vbCrLf
        sCode = sCode & "oPar.Value = CStr(Arg(2))" & vbCrLf
        sCode = sCode & "oCmd.Parameters.Append oPar" & vbCrLf
        sCode = sCode & "oCmd.CommandText = CStr(Arg(1))" & vbCrLf
        sCode = sCode & "dSession.Db.ExecuteCommand oCmd"
    
        Args = Array(Empty, Empty)
        Args(0) = "insert into DCE_HISTORY " & sbH.InsertString
        Args(1) = CodeMax1.Text
    
        'No es obligatorio el historial
        On Error Resume Next
        dSession.HttpCallCode sCode, Args
        'MsgBox "DCE_HISTORY: " & IIf(Err.Number <> 0, Err.Description, "ok"), vbInformation
        On Error GoTo Error
        
        ' El Oracle (cuando no) esta dando el error ORA-01036, se arregla con la version 9.2.0.4.0
        ' http://www.oracle.com/technology/software/tech/windows/ole_db/htdocs/readme9204.txt
    End If
    
    Exit Sub
Error:
    MsgBox Err.Description, vbExclamation
End Sub

