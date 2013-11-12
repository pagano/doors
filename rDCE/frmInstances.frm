VERSION 5.00
Begin VB.Form frmInstances 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Seleccione la instancia"
   ClientHeight    =   3096
   ClientLeft      =   48
   ClientTop       =   432
   ClientWidth     =   4224
   Icon            =   "frmInstances.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3096
   ScaleWidth      =   4224
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.ListBox lstInstances 
      Height          =   1392
      Left            =   360
      TabIndex        =   0
      Top             =   840
      Width           =   3375
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      Height          =   495
      Left            =   240
      MultiLine       =   -1  'True
      TabIndex        =   3
      TabStop         =   0   'False
      Text            =   "frmInstances.frx":000C
      Top             =   240
      Width           =   3855
   End
   Begin VB.CommandButton cmdOk 
      Caption         =   "Aceptar"
      Default         =   -1  'True
      Height          =   375
      Left            =   360
      TabIndex        =   1
      Top             =   2520
      Width           =   1455
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "Cancelar"
      Height          =   375
      Left            =   2280
      TabIndex        =   2
      Top             =   2520
      Width           =   1455
   End
End
Attribute VB_Name = "frmInstances"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public dSession As Object

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdOk_Click()
    Dim dom As Object
    
    On Error GoTo Error
    
    If lstInstances.ListIndex = -1 Then
        MsgBox "select an instance", vbExclamation
        lstInstances.SetFocus
        Exit Sub
    End If
    
    dSession.WinLogon lstInstances.ItemData(lstInstances.ListIndex), dom, True
    
    Unload Me
    
    Exit Sub
Error:
    ErrDisplay Err
End Sub

Private Sub lstInstances_DblClick()
    cmdOk_Click
End Sub
