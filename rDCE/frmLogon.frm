VERSION 5.00
Begin VB.Form frmLogon 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Remote DCE - Log On"
   ClientHeight    =   3576
   ClientLeft      =   48
   ClientTop       =   336
   ClientWidth     =   6048
   Icon            =   "frmLogon.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3576
   ScaleWidth      =   6048
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cmdGo 
      Caption         =   "Go!"
      Height          =   336
      Left            =   5352
      TabIndex        =   1
      Top             =   216
      Width           =   492
   End
   Begin VB.ComboBox cboServerURL 
      Height          =   288
      Left            =   1320
      TabIndex        =   0
      Text            =   "cboServerURL"
      Top             =   240
      Width           =   3972
   End
   Begin VB.CheckBox chkLite 
      Caption         =   "Lite Mode"
      Height          =   192
      Left            =   3096
      TabIndex        =   4
      Top             =   804
      Width           =   1452
   End
   Begin VB.CommandButton cmdLogoff 
      Caption         =   "Logoff"
      Height          =   375
      Left            =   2496
      TabIndex        =   9
      Top             =   3000
      Width           =   912
   End
   Begin VB.OptionButton optLogon 
      Caption         =   "Logon"
      Height          =   195
      Left            =   1296
      TabIndex        =   2
      Top             =   804
      Width           =   975
   End
   Begin VB.OptionButton optWinlogon 
      Caption         =   "Winlogon"
      Height          =   195
      Left            =   1296
      TabIndex        =   3
      Top             =   1164
      Width           =   1455
   End
   Begin VB.CommandButton cmdClose 
      Cancel          =   -1  'True
      Caption         =   "Close"
      Height          =   375
      Left            =   3816
      TabIndex        =   10
      Top             =   3000
      Width           =   912
   End
   Begin VB.CommandButton cmdLogon 
      Caption         =   "Logon"
      Default         =   -1  'True
      Height          =   375
      Left            =   1176
      TabIndex        =   8
      Top             =   3000
      Width           =   912
   End
   Begin VB.ComboBox cboInstance 
      Height          =   288
      Left            =   1440
      Style           =   2  'Dropdown List
      TabIndex        =   7
      Top             =   2412
      Width           =   3735
   End
   Begin VB.TextBox txtPassword 
      Height          =   285
      IMEMode         =   3  'DISABLE
      Left            =   1440
      PasswordChar    =   "*"
      TabIndex        =   6
      Text            =   "txtPassword"
      Top             =   2052
      Width           =   3735
   End
   Begin VB.TextBox txtLogin 
      Height          =   285
      Left            =   1440
      TabIndex        =   5
      Text            =   "txtLogin"
      Top             =   1692
      Width           =   3735
   End
   Begin VB.Label Label7 
      AutoSize        =   -1  'True
      Caption         =   "Instance"
      Height          =   192
      Left            =   540
      TabIndex        =   14
      Top             =   2472
      Width           =   612
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      Caption         =   "Password"
      Height          =   192
      Left            =   540
      TabIndex        =   13
      Top             =   2100
      Width           =   696
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "Login"
      Height          =   192
      Left            =   540
      TabIndex        =   12
      Top             =   1740
      Width           =   396
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Listener URL"
      Height          =   192
      Left            =   240
      TabIndex        =   11
      Top             =   288
      Width           =   924
   End
End
Attribute VB_Name = "frmLogon"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public dSession As Object
Private bLoadError As Boolean

Private Sub cmdClose_Click()
    Unload Me
End Sub

Private Sub cmdGo_Click()
    Dim sStatus As String
    Dim ses As Object
    
    On Error GoTo Error
    
    Screen.MousePointer = vbHourglass
    
    If cboServerURL.Text = "" Then
        MsgBox "must specify a connection", vbExclamation
    Else
        For Each ses In GcolConnections
            If ses.ServerURL = cboServerURL.Text Then
                MsgBox "The connection already exists", vbExclamation
                Exit Sub
            End If
        Next
            
        If cboServerURL.Text <> "" Then
            dSession.ServerURL = cboServerURL.Text
            If Not dSession.IsConnected(sStatus) Then
                MsgBox sStatus, vbExclamation
            End If
            LoadInstances
            EnableControls
        End If
    End If
    
    Screen.MousePointer = vbNormal
    Exit Sub
Error:
    Screen.MousePointer = vbNormal
    ErrDisplay Err
End Sub

Private Sub cmdLogon_Click()
    Dim dom As Object
    Dim oNode As Object
    
    On Error GoTo Error
    
    If optWinlogon Then
        If Not dSession.WinLogon(, dom, chkLite.Value = 1) Then
            With frmInstances.lstInstances
                .Clear
                For Each oNode In dom.documentElement.childNodes
                    .AddItem oNode.getAttribute("description")
                    .ItemData(.NewIndex) = CLng(oNode.getAttribute("id"))
                Next
                .ListIndex = 0
            End With
            frmInstances.Show vbModal
        End If
    Else
        dSession.Logon txtLogin.Text, txtPassword.Text, cboInstance.Text, chkLite = 1
    End If
    
    EnableControls
    
    Exit Sub
Error:
    ErrDisplay Err
End Sub

Private Sub cmdLogoff_Click()
    On Error GoTo Error
    
    dSession.Logoff
    EnableControls
    optLogon.SetFocus
    
    Exit Sub
Error:
    ErrDisplay Err
End Sub

Private Sub Form_Load()
    Dim sIt
    
    On Error GoTo Error
    
    Screen.MousePointer = vbHourglass
    
    Caption = "Logon to an Instance"
        
    For Each sIt In GdicURLs
        cboServerURL.AddItem sIt
    Next
    cboServerURL.Text = dSession.ServerURL
    
    LoadInstances
    
    optLogon.Value = True
    txtLogin.Text = ""
    txtPassword.Text = ""
    
    If dSession.IsConnected Then
        If dSession.IsLogged Then
            txtLogin.Text = dSession.LoggedUser.Login
            cboInstance.Text = dSession.InstanceName
        End If
    End If
    
    EnableControls
    Screen.MousePointer = vbNormal
    bLoadError = False
    Exit Sub
Error:
    Screen.MousePointer = vbNormal
    ErrDisplay Err
    bLoadError = True
End Sub

Private Sub Form_Activate()
    If bLoadError Then Unload Me
End Sub

Private Sub LoadInstances()
    Dim oDom As Object
    Dim oNode As Object
    Dim strInstance As String
    
    cboInstance.Clear
    
    If dSession.IsConnected Then
        Set oDom = dSession.InstanceList
        For Each oNode In oDom.documentElement.childNodes
            strInstance = oNode.getAttribute("name")
            cboInstance.AddItem strInstance
        Next
        If cboInstance.ListCount > 0 Then cboInstance.ListIndex = 0
    End If
End Sub

Private Sub EnableControls()
    If dSession.IsConnected Then
        If dSession.IsLogged Then
            cboServerURL.Enabled = False
            cmdGo.Enabled = False
            cmdGo.Default = False
            optWinlogon.Enabled = False
            optLogon.Enabled = False
            chkLite.Enabled = False
            txtLogin.Enabled = False
            txtPassword.Enabled = False
            cboInstance.Enabled = False
            cmdLogon.Enabled = False
            cmdLogon.Default = False
            cmdLogoff.Enabled = True
        Else
            cboServerURL.Enabled = True
            cmdGo.Enabled = True
            cmdGo.Default = False
            optWinlogon.Enabled = True
            optLogon.Enabled = True
            chkLite.Enabled = True
            txtLogin.Enabled = optLogon.Value
            txtPassword.Enabled = optLogon.Value
            cboInstance.Enabled = optLogon.Value
            cmdLogon.Enabled = True
            cmdLogon.Default = True
            cmdLogoff.Enabled = False
        End If
    Else
        cboServerURL.Enabled = True
        cmdGo.Enabled = True
        cmdGo.Default = True
        optWinlogon.Enabled = False
        optLogon.Enabled = False
        chkLite.Enabled = False
        txtLogin.Enabled = False
        txtPassword.Enabled = False
        cboInstance.Enabled = False
        cmdLogon.Enabled = False
        cmdLogon.Default = False
        cmdLogoff.Enabled = False
    End If
End Sub

Private Sub optLogon_Click()
    txtLogin.Enabled = True
    txtPassword.Enabled = True
    cboInstance.Enabled = True
End Sub

Private Sub optWinlogon_Click()
    txtLogin.Enabled = False
    txtPassword.Enabled = False
    cboInstance.Enabled = False
End Sub


