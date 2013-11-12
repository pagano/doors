VERSION 5.00
Begin VB.UserControl SplitPanel 
   Alignable       =   -1  'True
   ClientHeight    =   3270
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   3240
   ControlContainer=   -1  'True
   ScaleHeight     =   3270
   ScaleWidth      =   3240
   Begin VB.PictureBox pbSplit 
      BorderStyle     =   0  'None
      Height          =   1380
      Left            =   1320
      ScaleHeight     =   1380
      ScaleWidth      =   165
      TabIndex        =   0
      Top             =   720
      Width           =   165
   End
End
Attribute VB_Name = "SplitPanel"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'*********'*********'*********'*********'*********'*********'*********'*********
'
' SplitPanel
' © 1998-2000 David Crowell
' www.davidcrowell.com
'
' Last updated: November 12, 2000
'
' This source-code is provided 'as-is' with no warranty of any kind.  Use
' at your own risk.
' You may use this source code in any of your own projects.  You may distribute
' this source code as long as you leave this notice intact.
'
' SplitPanel is a VB usercontrol that will compile under VB5 or VB6.  It allows
' a user-adjustable split between two other controls.
'
' Properties:
'
' Control1 As Object (r/w) - Control that will be on left or top
' Control2 As Object (r/w) - Control that will be on right or bottom
' MinPanelSize as Long (r/w) - Minimum size for a panel
' Horizontal as Boolean (r/w) - True makes this a horizontal split
' SnapToEdge as Boolean (r/w) - If true allows the user to move the split most
'       of the way to one side, and the control on that side will disappear.
' SplitterWidth as Long (r/w) - Width of the splitter bar (in twips)
' Position as Long (r/w) - Position of the splitter bar (in twips) from the top
'       or left
' DragColor as OLE_COLOR (r/w) - Color of the splitter bar while dragging
' BackColor as OLE_COLOR (r/w) - Color of the splitter bar
' Snapped as Integer (r/w) - If SnapToEdge is True, this indicates the current
'       Snapped Position:
'           0 - Not Snapped
'           1 - Snapped to top or left
'           2 - Snapped to bottom or right
'
' Events:
'
' Change(Position As Long) - Is fired whenever the split is moved, or the
'       control is resized.  This event may also be fired multiple times
'       during startup.
'
'*********'*********'*********'*********'*********'*********'*********'*********
Option Explicit

Public Event Change(Position As Long)

Private moControl1 As Object
Private moControl2 As Object
Private mlngMinPanelSize As Long
Private mblnHorizontal As Boolean
Private mblnSnapToEdge As Boolean
Private mlngSplitterWidth As Long
Private mlngPosition As Long
Private mlngDragColor As Long
Private mintSnapped As Integer
Private mblnButtonDown As Boolean

'*********'*********'*********'*********'*********'*********'*********'*********
' Public Properties
'*********'*********'*********'*********'*********'*********'*********'*********
Public Property Get Control1() As Object
    Set Control1 = moControl1
End Property
Public Property Set Control1(vData As Object)
    Set moControl1 = vData
    UserControl_Resize
End Property

Public Property Get Control2() As Object
    Set Control2 = moControl2
End Property
Public Property Set Control2(vData As Object)
    Set moControl2 = vData
    UserControl_Resize
End Property

Public Property Get MinPanelSize() As Long
    MinPanelSize = mlngMinPanelSize
End Property
Public Property Let MinPanelSize(vData As Long)
    mlngMinPanelSize = vData
    PropertyChanged "MinPanelSize"
    UserControl_Resize
End Property

Public Property Get Horizontal() As Boolean
    Horizontal = mblnHorizontal
End Property
Public Property Let Horizontal(vData As Boolean)
    mblnHorizontal = vData
    PropertyChanged "Horizontal"
    pbSplit.MousePointer = IIf(vData, vbSizeNS, vbSizeWE)
    UserControl_Resize
End Property

Public Property Get SnapToEdge() As Boolean
    SnapToEdge = mblnSnapToEdge
End Property
Public Property Let SnapToEdge(vData As Boolean)
    mblnSnapToEdge = vData
    PropertyChanged "SnapToEdge"
    UserControl_Resize
End Property

Public Property Get SplitterWidth() As Long
    SplitterWidth = mlngSplitterWidth
End Property
Public Property Let SplitterWidth(vData As Long)
    mlngSplitterWidth = vData
    PropertyChanged "SplitterWidth"
    UserControl_Resize
End Property

Public Property Get Position() As Long
    Position = mlngPosition
End Property
Public Property Let Position(vData As Long)
    mlngPosition = vData
    PropertyChanged "Position"
    mintSnapped = 0
    UserControl_Resize
End Property

Public Property Get BackColor() As OLE_COLOR
    BackColor = UserControl.BackColor
End Property
Public Property Let BackColor(vData As OLE_COLOR)
    pbSplit.BackColor = vData
    UserControl.BackColor = vData
    PropertyChanged "BackColor"
End Property

Public Property Get DragColor() As OLE_COLOR
    DragColor = mlngDragColor
End Property
Public Property Let DragColor(vData As OLE_COLOR)
    mlngDragColor = vData
    PropertyChanged "DragColor"
End Property

Public Property Get Snapped() As Integer
    Snapped = mintSnapped
End Property
Public Property Let Snapped(vData As Integer)
    If Not mblnSnapToEdge Then vData = 0
    mintSnapped = vData
    Select Case vData
    Case 1
        mlngPosition = 0
    Case 2
        mlngPosition = IIf(mblnHorizontal, UserControl.ScaleHeight, _
                    UserControl.ScaleWidth)
    Case Else
        mintSnapped = 0
    End Select
    pbSplit.ZOrder
    UserControl_Resize
End Property

'*********'*********'*********'*********'*********'*********'*********'*********
' Initialization and Termination
'*********'*********'*********'*********'*********'*********'*********'*********
Private Sub UserControl_InitProperties()
    mlngMinPanelSize = 2000
    Horizontal = False
    mblnSnapToEdge = False
    mlngSplitterWidth = 75
    mlngPosition = 3000
    DragColor = vbButtonText
    BackColor = vbButtonFace
    Snapped = 0
    UserControl_Load
End Sub

Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
    mlngMinPanelSize = PropBag.ReadProperty("MinPanelSize", 2000)
    Horizontal = PropBag.ReadProperty("Horizontal", False)
    mblnSnapToEdge = PropBag.ReadProperty("SnapToEdge", False)
    mlngSplitterWidth = PropBag.ReadProperty("SplitterWidth", 75)
    DragColor = PropBag.ReadProperty("DragColor", vbButtonText)
    BackColor = PropBag.ReadProperty("BackColor", vbButtonFace)
    Snapped = PropBag.ReadProperty("Snapped", 0)
    UserControl_Load
End Sub

Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
    PropBag.WriteProperty "MinPanelSize", mlngMinPanelSize, 2000
    PropBag.WriteProperty "Horizontal", mblnHorizontal, False
    PropBag.WriteProperty "SnapToEdge", mblnSnapToEdge, False
    PropBag.WriteProperty "SplitterWidth", mlngSplitterWidth, 75
    PropBag.WriteProperty "DragColor", mlngDragColor, vbButtonText
    PropBag.WriteProperty "BackColor", UserControl.BackColor, vbButtonFace
    PropBag.WriteProperty "Snapped", mintSnapped, 0
End Sub

Private Sub UserControl_Load()
    ' fake event always called during control creation
    If Ambient.UserMode Then
        UserControl.BorderStyle = 0
    Else
        UserControl.BorderStyle = 1
    End If
    
End Sub

Private Sub UserControl_Terminate()
    Set moControl1 = Nothing
    Set moControl2 = Nothing
End Sub

'*********'*********'*********'*********'*********'*********'*********'*********
' Mouse events
'*********'*********'*********'*********'*********'*********'*********'*********
Private Sub pbSplit_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = vbLeftButton Then
        mblnButtonDown = True
        pbSplit.BackColor = mlngDragColor
        pbSplit.ZOrder
    End If
End Sub

Private Sub pbSplit_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    
    If mblnButtonDown Then
        If mblnHorizontal Then
            pbSplit.Top = pbSplit.Top - ((mlngSplitterWidth \ 2) - Y)
        Else
            pbSplit.Left = pbSplit.Left - ((mlngSplitterWidth \ 2) - X)
        End If
    End If
End Sub

Private Sub pbSplit_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    
    Dim lngTotal As Long
    
    If Button = vbLeftButton Then
        mblnButtonDown = False
        pbSplit.BackColor = UserControl.BackColor
        If mblnHorizontal Then
            mlngPosition = pbSplit.Top + (mlngSplitterWidth \ 2)
            lngTotal = UserControl.ScaleHeight
        Else
            mlngPosition = pbSplit.Left + (mlngSplitterWidth \ 2)
            lngTotal = UserControl.ScaleWidth
        End If
        
        ' check if we should snap to the edge
        If mblnSnapToEdge Then
            Select Case mlngPosition
            Case Is < mlngMinPanelSize \ 4
                mlngPosition = 0
                mintSnapped = 1
            Case Is > lngTotal - (mlngMinPanelSize \ 4)
                mlngPosition = lngTotal - mlngSplitterWidth
                mintSnapped = 2
            Case Else
                mintSnapped = 0
            End Select
        End If
            
        UserControl_Resize
    End If
End Sub

'*********'*********'*********'*********'*********'*********'*********'*********
' Most of the code is here in the resize event
'*********'*********'*********'*********'*********'*********'*********'*********
Private Sub UserControl_Resize()
    ' re-position all controls
    
    Dim X1 As Long
    Dim Y1 As Long
    Dim W1 As Long
    Dim H1 As Long
    
    Dim X2 As Long
    Dim Y2 As Long
    Dim W2 As Long
    Dim H2 As Long
    
    Dim XSP As Long
    Dim YSP As Long
    Dim WSP As Long
    Dim HSP As Long
    
    ' check for controls that should not be visible
    On Error Resume Next
    Select Case mintSnapped
    Case 0
        If Not moControl1 Is Nothing Then moControl1.Visible = True
        If Not moControl2 Is Nothing Then moControl2.Visible = True
    Case 1
        If Not moControl1 Is Nothing Then moControl1.Visible = False
        If Not moControl2 Is Nothing Then moControl2.Visible = True
    Case 2
        If Not moControl1 Is Nothing Then moControl1.Visible = True
        If Not moControl2 Is Nothing Then moControl2.Visible = False
    End Select
    On Error GoTo 0
    
    If mblnHorizontal Then
        
        ' check for snapped to bottom edge
        If mintSnapped = 2 Then
            mlngPosition = UserControl.ScaleHeight - (mlngSplitterWidth \ 2)
        End If
        
        ' check minimum height of usercontrol
        If UserControl.ScaleHeight < mlngMinPanelSize * 2 Then
            Select Case mintSnapped
            Case 0
                mlngPosition = UserControl.ScaleHeight \ 2
            Case 1
                mlngPosition = 0
            Case 2
                mlngPosition = UserControl.ScaleHeight - (mlngSplitterWidth \ 2)
            End Select
        Else
        
            ' check first panel size
            If mlngPosition < mlngMinPanelSize Then
                Select Case mintSnapped
                Case 1
                    mlngPosition = 0
                Case Else
                    mlngPosition = mlngMinPanelSize
                End Select
            Else
            
                ' check second panel size
                If mlngPosition > (UserControl.ScaleHeight - mlngMinPanelSize) Then
                    Select Case mintSnapped
                    Case 2
                        mlngPosition = UserControl.ScaleHeight - (mlngSplitterWidth \ 2)
                    Case Else
                        mlngPosition = UserControl.ScaleHeight - mlngMinPanelSize
                    End Select
                End If
            End If
        End If
        
        
        ' figure positions for horizontal split
        X1 = 0
        Y1 = 0
        W1 = UserControl.ScaleWidth
        H1 = mlngPosition - (mlngSplitterWidth \ 2)
        
        X2 = 0
        Y2 = H1 + mlngSplitterWidth
        W2 = W1
        H2 = UserControl.ScaleHeight - Y2
        
        XSP = 0
        YSP = H1
        WSP = W1
        HSP = mlngSplitterWidth
        
        
    Else
    
        ' check for snapped to right edge
        If mintSnapped = 2 Then
            mlngPosition = UserControl.ScaleWidth - mlngSplitterWidth
        End If
        
        ' check minimum width of usercontrol
        If UserControl.ScaleWidth < mlngMinPanelSize * 2 Then
            Select Case mintSnapped
            Case 0
                mlngPosition = UserControl.ScaleWidth \ 2
            Case 1
                mlngPosition = 0
            Case 2
                mlngPosition = UserControl.ScaleWidth - (mlngSplitterWidth \ 2)
            End Select
        Else
        
            ' check first panel size
            If mlngPosition < mlngMinPanelSize Then
                Select Case mintSnapped
                Case 1
                    mlngPosition = 0
                Case Else
                    mlngPosition = mlngMinPanelSize
                End Select
            Else
            
                ' check second panel size
                If mlngPosition > (UserControl.ScaleWidth - mlngMinPanelSize) Then
                    Select Case mintSnapped
                    Case 2
                        mlngPosition = UserControl.ScaleWidth - (mlngSplitterWidth \ 2)
                    Case Else
                        mlngPosition = UserControl.ScaleWidth - mlngMinPanelSize
                    End Select
                End If
            End If
        End If
        
        ' figure position for vertical split
        X1 = 0
        Y1 = 0
        W1 = mlngPosition - (mlngSplitterWidth \ 2)
        H1 = UserControl.ScaleHeight
        
        X2 = W1 + mlngSplitterWidth
        Y2 = 0
        W2 = UserControl.ScaleWidth - X2
        H2 = H1
        
        XSP = W1
        YSP = 0
        WSP = mlngSplitterWidth
        HSP = H1
        
    End If
    
    ' check for illegal values
    If W1 < 1 Then W1 = 1
    If W2 < 1 Then W2 = 1
    If H1 < 1 Then H1 = 1
    If H2 < 1 Then H2 = 1
    If WSP < 1 Then WSP = 1
    If HSP < 1 Then HSP = 1
    
    
    ' move splitter bar
    pbSplit.Move XSP, YSP, WSP, HSP
    
    ' move first control
    If Not moControl1 Is Nothing Then
        moControl1.Move X1, Y1, W1, H1
    End If
    
    ' move second control
    If Not moControl2 Is Nothing Then
        moControl2.Move X2, Y2, W2, H2
    End If
    
    ' let the client know about changes, and return the position
    RaiseEvent Change(mlngPosition)
    
End Sub
