Attribute VB_Name = "Module1"
Option Explicit

Public CodeMaxGlobals As CodeMax4Ctl.Globals

Private Declare Function GetPrivateProfileString Lib "kernel32" _
    Alias "GetPrivateProfileStringA" _
    (ByVal lpApplicationName As String, ByVal lpKeyName As Any, _
    ByVal lpDefault As String, ByVal lpReturnedString As String, _
    ByVal nSize As Long, ByVal lpFileName As String) As Long

Private Declare Function WritePrivateProfileString Lib "kernel32" _
    Alias "WritePrivateProfileStringA" _
    (ByVal lpApplicationName As String, ByVal lpKeyName As Any, _
    ByVal lpString As Any, ByVal lpFileName As String) As Long

Public GstrIniFile As String
Public GblnNormalizeCase As Boolean
Public GcolConnections As Collection
Public GcolExplorers As Collection
Public GdicURLs As Scripting.Dictionary

Sub Main()
    Dim lCMaxRevision As Long
    Dim i As Long
    Dim sAux As String
    
    GstrIniFile = App.Path & "\" & App.EXEName & ".ini"
    GblnNormalizeCase = True
    Set GcolConnections = New Collection
    Set GcolExplorers = New Collection
    Set CodeMaxGlobals = New CodeMax4Ctl.Globals
    
    Set GdicURLs = CreateObject("Scripting.Dictionary")
    GdicURLs.CompareMode = 1 ' VBTextCompare
    
    For i = 0 To 9
        sAux = ReadIni("Session", "ServerURL" & i)
        If sAux <> "" Then GdicURLs.Add sAux, Empty
    Next
    
    On Error Resume Next
    lCMaxRevision = CLng(Split(CodeMaxVersion, ".")(3))
    On Error GoTo 0
    
    If lCMaxRevision < 9 Then
        GblnNormalizeCase = False
        MsgBox "This version of CodeMax can change letter case between <##> tags." & vbCrLf & _
            "Because this bug can affect your JavaScript code, case normalization will be disabled." & vbCrLf & _
            "Please upgrade your CodeMax control to the version 4.0.0.9 or greater " & _
            "(http://www.gestar.com/kbview.asp?fld=2374&doc=336504).", vbExclamation
    End If

    MDIForm1.Show
End Sub

Public Function ReadIni(ByRef Section As String, ByRef Key As String) As String
    ReadIni = GetProfileString(GstrIniFile, Section, Key)
End Function

Public Function WriteIni(ByRef Section As String, ByRef Key As String, ByRef Value As String) As Long
    WriteIni = WriteProfileString(GstrIniFile, Section, Key, Value)
End Function

Private Function GetProfileString(ByRef IniFile As String, ByRef Application As String, ByRef Key As String) As String
  Dim strAux As String
  
  strAux = String(1024, vbNullChar)
  If GetPrivateProfileString(Application, Key, "", strAux, 1024, IniFile) = 0 Then
    GetProfileString = ""
  Else
    GetProfileString = Left(strAux, InStr(strAux, vbNullChar) - 1)
  End If
End Function

Private Function WriteProfileString(ByRef IniFile As String, ByRef Application As String, ByRef Key As String, ByRef Value As String) As Long
  WriteProfileString = WritePrivateProfileString(Application, Key, Value, IniFile)
End Function

Private Function CodeMaxVersion() As String
    Dim oShell As Object
    Dim strDll As String
    Dim oFso As Object
    
    Set oShell = CreateObject("WScript.Shell")
    strDll = oShell.RegRead("HKEY_CLASSES_ROOT\CLSID\{BCA00001-18B1-43E0-BB89-FECDDBF0472E}\InprocServer32\")
    Set oFso = CreateObject("Scripting.FileSystemObject")
    CodeMaxVersion = oFso.GetFileVersion(strDll)
End Function

Public Sub ErrDisplay(ByRef Err As Object)
    MsgBox Err.Description & " (" & Err.Number & ")", vbExclamation
End Sub

Public Sub ListViewColumnClick(ByRef pListView As MSComctlLib.ListView, ByVal ColumnHeader As MSComctlLib.ColumnHeader)
    With pListView
        If Not .Sorted Then
            .SortKey = ColumnHeader.Index - 1
            .SortOrder = lvwAscending
            .Sorted = True
        Else
            If .SortKey = ColumnHeader.Index - 1 Then
                ' Invertir
                If .SortOrder = lvwAscending Then
                    .SortOrder = lvwDescending
                Else
                    .SortOrder = lvwAscending
                End If
            Else
                .SortKey = ColumnHeader.Index - 1
                .SortOrder = lvwAscending
            End If
        End If
    End With
End Sub

Public Function Debugging() As Boolean
    On Error Resume Next
    Debug.Assert 1 / 0
    Debugging = (Err <> 0)
End Function
