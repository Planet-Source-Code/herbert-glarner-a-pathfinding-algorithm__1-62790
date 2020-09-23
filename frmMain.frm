VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmMain 
   Caption         =   "A+ Pathfinding Algorithm"
   ClientHeight    =   5520
   ClientLeft      =   165
   ClientTop       =   735
   ClientWidth     =   8415
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   368
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   561
   StartUpPosition =   3  'Windows-Standard
   Begin MSComctlLib.StatusBar staMain 
      Align           =   2  'Unten ausrichten
      Height          =   255
      Left            =   0
      TabIndex        =   18
      Top             =   5265
      Width           =   8415
      _ExtentX        =   14843
      _ExtentY        =   450
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   11
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   6615
            Key             =   "Msg"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   2
            Bevel           =   0
            Object.Width           =   344
            MinWidth        =   317
            Text            =   "X"
            TextSave        =   "X"
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            Object.Width           =   952
            MinWidth        =   952
            Key             =   "X"
         EndProperty
         BeginProperty Panel4 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   2
            Bevel           =   0
            Object.Width           =   344
            MinWidth        =   317
            Text            =   "Y"
            TextSave        =   "Y"
         EndProperty
         BeginProperty Panel5 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            Object.Width           =   952
            MinWidth        =   952
            Key             =   "Y"
         EndProperty
         BeginProperty Panel6 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   2
            Bevel           =   0
            Object.Width           =   344
            MinWidth        =   317
            Text            =   "C"
            TextSave        =   "C"
         EndProperty
         BeginProperty Panel7 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            Object.Width           =   1270
            MinWidth        =   1270
            Key             =   "C"
         EndProperty
         BeginProperty Panel8 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   2
            Bevel           =   0
            Object.Width           =   370
            MinWidth        =   317
            Text            =   "H"
            TextSave        =   "H"
         EndProperty
         BeginProperty Panel9 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            Object.Width           =   1270
            MinWidth        =   1270
            Key             =   "H"
         EndProperty
         BeginProperty Panel10 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   2
            Bevel           =   0
            Object.Width           =   344
            MinWidth        =   317
            Text            =   "T"
            TextSave        =   "T"
         EndProperty
         BeginProperty Panel11 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            Object.Width           =   1270
            MinWidth        =   1270
            Key             =   "T"
         EndProperty
      EndProperty
   End
   Begin VB.Frame Frame3 
      Height          =   705
      Left            =   90
      TabIndex        =   17
      Top             =   4470
      Width           =   2025
      Begin VB.CommandButton cmdErasePath 
         Caption         =   "&Erase Path"
         Enabled         =   0   'False
         Height          =   315
         Left            =   180
         TabIndex        =   10
         Top             =   240
         Width           =   1695
      End
   End
   Begin VB.Frame Frame2 
      Height          =   1425
      Left            =   90
      TabIndex        =   16
      Top             =   2940
      Width           =   2025
      Begin VB.CommandButton cmdSingleStep 
         Caption         =   "Single Step (F8)"
         Enabled         =   0   'False
         Height          =   315
         Left            =   150
         TabIndex        =   9
         Top             =   990
         Width           =   1695
      End
      Begin VB.CheckBox chkSingleStep 
         Caption         =   "Si&ngle Step"
         Height          =   225
         Left            =   150
         TabIndex        =   7
         Top             =   240
         Width           =   1695
      End
      Begin VB.CommandButton cmdFindPath 
         Caption         =   "&Find Path"
         Enabled         =   0   'False
         Height          =   315
         Left            =   150
         TabIndex        =   8
         Top             =   600
         Width           =   1695
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Action on Click"
      Height          =   1575
      Left            =   90
      TabIndex        =   15
      Top             =   1260
      Width           =   2025
      Begin VB.OptionButton optSetVoid 
         Caption         =   "&Clear (right button)"
         Height          =   195
         Left            =   150
         TabIndex        =   6
         Top             =   1140
         Width           =   1725
      End
      Begin VB.OptionButton optSetTarget 
         Caption         =   "Set &Target"
         ForeColor       =   &H00C00000&
         Height          =   195
         Left            =   150
         TabIndex        =   4
         Top             =   600
         Width           =   1725
      End
      Begin VB.OptionButton optSetObstacle 
         Caption         =   "Set &Obstacle"
         ForeColor       =   &H000000C0&
         Height          =   195
         Left            =   150
         TabIndex        =   5
         Top             =   870
         Width           =   1725
      End
      Begin VB.OptionButton optSetStart 
         Caption         =   "Set &Start"
         ForeColor       =   &H00008000&
         Height          =   195
         Left            =   150
         TabIndex        =   3
         Top             =   330
         Width           =   1725
      End
   End
   Begin VB.PictureBox picGrid 
      AutoRedraw      =   -1  'True
      BorderStyle     =   0  'Kein
      FillStyle       =   0  'Ausgefüllt
      Height          =   5025
      Left            =   2220
      MousePointer    =   2  'Kreuz
      ScaleHeight     =   335
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   407
      TabIndex        =   12
      TabStop         =   0   'False
      Top             =   150
      Width           =   6105
   End
   Begin VB.Frame fraDims 
      Caption         =   "Grid Dimensions"
      Height          =   1095
      Left            =   90
      TabIndex        =   11
      Top             =   60
      Width           =   2025
      Begin VB.CommandButton cmdMakeGrid 
         Caption         =   "&Redraw"
         Enabled         =   0   'False
         Height          =   315
         Left            =   150
         TabIndex        =   2
         Top             =   660
         Width           =   1695
      End
      Begin VB.TextBox txtH 
         Alignment       =   1  'Rechts
         Height          =   285
         Left            =   1110
         TabIndex        =   1
         Text            =   "10"
         Top             =   270
         Width           =   555
      End
      Begin VB.TextBox txtW 
         Alignment       =   1  'Rechts
         Height          =   285
         Left            =   150
         TabIndex        =   0
         Text            =   "12"
         Top             =   270
         Width           =   555
      End
      Begin VB.Label Label3 
         Caption         =   "H"
         Height          =   195
         Left            =   1710
         TabIndex        =   14
         Top             =   330
         Width           =   195
      End
      Begin VB.Label Label2 
         Caption         =   "W x"
         Height          =   195
         Left            =   750
         TabIndex        =   13
         Top             =   330
         Width           =   315
      End
   End
   Begin VB.Menu mnuFile 
      Caption         =   "&File"
      Begin VB.Menu mncExit 
         Caption         =   "E&xit"
      End
   End
   Begin VB.Menu mnuHelp 
      Caption         =   "&?"
      Begin VB.Menu mncAbout 
         Caption         =   "&About"
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'A+ Pathfinding Algorithm: Implementation by Herbert Glarner (herbert.glarner@bluewin.ch)
'Unlimited use for whatever purpose allowed provided that above credits are given.
'Suggestions and bug reports welcome.

'The grid which contains start, target and obstacles
Private gaeGrid() As eCell
Private Enum eCell
    Void = 0&
    Start = 1&
    Obstacle = 2&
    Target = 3&
End Enum

'Working list to find closest path
Private grList() As tCell
Private Type tCell
    X As Long               'Coordinates of the listed cell
    Y As Long
    Parent As Long          'Parent Index within the list (-1 for start point)
    Cost As Single          'Cost to get til here
    Heuristic As Single     'Estimated cost til target
    Closed As Boolean       'Not considered anymore
End Type

'This enum is needed in APlus.
'Usually this matrix be defined in APlus. It is just here
'because the test form wants to display infos about it's fields.
'The two fields can also be merged into the source matrix.
Private abGridCopy() As tGrid
Private Type tGrid
    ListStat As eListStat   'Status of the list element
    Index As Long           'Index into the open list.
End Type
Private Enum eListStat
    Unprocessed = 0&
    IsOpen = 1&
    IsClosed = 2&
End Enum

'An array of this type is the algorithm's result.
Private Type tPoint
    X As Long
    Y As Long
End Type

'Start and target coordinates held separately to avoid
'to search for them. -1/-1 if not existent.
Private lStartX As Long, lStartY As Long
Private lTargetX As Long, lTargetY As Long

'Dimensions of the Grid
Private glGridW As Long, glGridH As Long

'Edge length of a grid cell
Private glCellLen As Long

'Tells if A+ is currently running
Private bRunning As Boolean

'For single step modus when Animation is checked
'F8 will continue then in single step modus.
Private bSingleStep As Boolean
Private bDoSingleStep As Boolean
Private Declare Function GetAsyncKeyState Lib "user32" (ByVal vKey As Long) As Integer
Private Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)

    

Private Sub cmdErasePath_Click()
    ErasePath
    cmdFindPath.SetFocus
    DoEvents
    staMain.Panels("Msg").Text = "Path erased."
End Sub

Private Sub cmdErasePath_GotFocus()
    If Not bRunning Then staMain.Panels("Msg").Text = "Clean up the grid by erasing the pathfinding indicators."
End Sub

Private Sub cmdFindPath_Click()
    Dim arPath() As tPoint
    Dim sCost As Single
    
    ErasePath
    bRunning = True
    cmdMakeGrid.Enabled = False
    cmdFindPath.Enabled = False
    txtW.Enabled = False
    txtH.Enabled = False
    optSetStart.Enabled = False
    optSetTarget.Enabled = False
    optSetObstacle.Enabled = False
    optSetVoid.Enabled = False
    
    staMain.Panels("Msg").Text = "Finding path..."
    
    If Not APlus(lStartX, lStartY, lTargetX, lTargetY, Void, arPath(), sCost) Then
        staMain.Panels("Msg").Text = "No path exists."
    Else
        'Enable possibility to remove path indicators and play more
        'with the existing data.
        staMain.Panels("Msg").Text = "Path found. Cost is " & Format$(sCost, "0.00") & "."
    End If
    
    optSetStart.Enabled = True
    optSetTarget.Enabled = True
    optSetObstacle.Enabled = True
    optSetObstacle.Value = True
    optSetVoid.Enabled = True
    txtW.Enabled = True
    txtH.Enabled = True
    cmdMakeGrid.Enabled = True
    cmdErasePath.Enabled = True
    cmdErasePath.SetFocus
    DoEvents
    bRunning = False
End Sub

Private Sub cmdFindPath_GotFocus()
    staMain.Panels("Msg").Text = "Start pathfinding."
End Sub

Private Sub cmdMakeGrid_Click()
    MakeGrid False
    optSetStart.Value = True
    optSetStart.SetFocus
End Sub

Private Sub cmdMakeGrid_GotFocus()
    staMain.Panels("Msg").Text = "Create new grid as per above settings."
End Sub

'Called by WaitForKey when cmdSingleStep is pressed in Single Step mode
Private Sub cmdSingleStep_Click()
    'Ignore if NOT in single step mode
    If Not bSingleStep Then Exit Sub

    'Allow single step
    bDoSingleStep = True
End Sub

Private Sub cmdSingleStep_GotFocus()
    'No message here: Single steps provide the activity of the last step;
    'we would overwrite that text.
End Sub

Private Sub Form_Load()
    'Invalidate start and target
    lStartX = -1: lStartY = -1
    lTargetX = -1: lTargetY = -1
    
    Show
    optSetStart.Value = True
    optSetStart.SetFocus
End Sub

Private Sub Form_Resize()
    Dim lW As Long, lH As Long
    
    'No resizing while running
    If bRunning Then Exit Sub
    
    'Resize picGrid
    lW = frmMain.ScaleWidth - 155
    lH = frmMain.ScaleHeight - 36
    If lW > 0 And lH > 0 Then
        picGrid.Width = lW
        picGrid.Height = lH
    
        'Display Grid to allow user to play with cells
        If glGridW Then MakeGrid True Else MakeGrid False
    End If
    
    cmdErasePath.Enabled = False
End Sub

Private Sub Form_Unload(Cancel As Integer)
    End
End Sub

Private Sub mncAbout_Click()
    frmAbout.Show vbModal
End Sub

Private Sub mncExit_Click()
    End
End Sub

Private Sub optSetObstacle_Click()
    If Not bRunning Then staMain.Panels("Msg").Text = "Set obstacles which can not be traversed."
End Sub
Private Sub optSetObstacle_GotFocus()
    If Not bRunning Then staMain.Panels("Msg").Text = "Set obstacles which can not be traversed."
End Sub

Private Sub optSetStart_Click()
    If Not bRunning Then staMain.Panels("Msg").Text = "Set point from where to start the search."
End Sub
Private Sub optSetStart_GotFocus()
    If Not bRunning Then staMain.Panels("Msg").Text = "Set point from where to start the search."
End Sub

Private Sub optSetTarget_Click()
    If Not bRunning Then staMain.Panels("Msg").Text = "Set target point to search a path to."
End Sub
Private Sub optSetTarget_GotFocus()
    If Not bRunning Then staMain.Panels("Msg").Text = "Set target point to search a path to."
End Sub

Private Sub optSetVoid_Click()
    If Not bRunning Then staMain.Panels("Msg").Text = "Clear the cell's content."
End Sub
Private Sub optSetVoid_GotFocus()
    If Not bRunning Then staMain.Panels("Msg").Text = "Clear the cell's content."
End Sub

Private Sub chkSingleStep_GotFocus()
    If Not bRunning Then staMain.Panels("Msg").Text = "Activate to enter single step mode."
End Sub

Private Sub picGrid_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim lX As Long, lY As Long
    
    'Convert clicked pixel into grid coordinates
    lX = X \ glCellLen: lY = Y \ glCellLen
    
    'Not beyond the defined grid
    If lX < 0 Or lX >= glGridW Or lY < 0 Or lY >= glGridH Then Exit Sub
    
    'Ignore if in single step mode
    If bSingleStep Then Exit Sub
    
    cmdFindPath.Enabled = (lStartX <> -1 And lTargetX <> -1)
    
    'If the clicked cell contains the start or the target, clear this first.
    If gaeGrid(lX, lY) = Start Then
        gaeGrid(lX, lY) = Void
        lStartX = -1: lStartY = -1
    ElseIf gaeGrid(lX, lY) = Target Then
        gaeGrid(lX, lY) = Void
        lTargetX = -1: lTargetY = -1
    End If
    
    EraseCell lX, lY
    
    'Right mouse button always is "Erase": Is just more convenient.
    If Button = 2 Then
        gaeGrid(lX, lY) = Void
        Exit Sub
    End If
    
    'Set the new start, obstacle or target, or clear the cell, as per
    'active option button.
    If optSetStart.Value Then
        If lStartX <> -1 Then
            gaeGrid(lStartX, lStartY) = Void
            EraseCell lStartX, lStartY
        End If
        lStartX = lX: lStartY = lY
        gaeGrid(lX, lY) = Start
        DrawDisc lX, lY, &H8000&
    ElseIf optSetTarget.Value Then
        If lTargetX <> -1 Then
            gaeGrid(lTargetX, lTargetY) = Void
            EraseCell lTargetX, lTargetY
        End If
        lTargetX = lX: lTargetY = lY
        gaeGrid(lX, lY) = Target
        DrawDisc lX, lY, &H800000
    ElseIf optSetObstacle.Value Then
        gaeGrid(lX, lY) = Obstacle
        DrawDisc lX, lY, &H80&
    ElseIf optSetVoid.Value Then
        gaeGrid(lX, lY) = Void
    End If
End Sub

Private Sub picGrid_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim lX As Long, lY As Long
    Dim lIndex As Long
    Dim sC As Single, sH As Single
    
    
    'Convert clicked pixel into grid coordinates
    lX = X \ glCellLen: lY = Y \ glCellLen
    
    'Not beyond the defined grid
    If lX < 0 Or lX >= glGridW Or lY < 0 Or lY >= glGridH Then Exit Sub
    
    'Display the current coordinates
    staMain.Panels("X") = CStr(lX): staMain.Panels("Y") = CStr(lY)
    'Also cost and heuristic if any
    If abGridCopy(lX, lY).ListStat <> Unprocessed Then
        lIndex = abGridCopy(lX, lY).Index
        sC = grList(lIndex).Cost: sH = grList(lIndex).Heuristic
        staMain.Panels("C").Text = Format$(sC, "0.00")
        staMain.Panels("H").Text = Format$(sH, "0.00")
        staMain.Panels("T").Text = Format$(sC + sH, "0.00")
    Else
        staMain.Panels("C").Text = ""
        staMain.Panels("H").Text = ""
        staMain.Panels("T").Text = ""
    End If
    
    'Ignore if in single step mode
    If Button = 0 Then Exit Sub
    
    cmdFindPath.Enabled = (lStartX <> -1 And lTargetX <> -1)
    
    'If the clicked cell contains the start or the target, clear this first.
    If gaeGrid(lX, lY) = Start Then
        gaeGrid(lX, lY) = Void
        lStartX = -1: lStartY = -1
    ElseIf gaeGrid(lX, lY) = Target Then
        gaeGrid(lX, lY) = Void
        lTargetX = -1: lTargetY = -1
    End If
    
    EraseCell lX, lY
    
    'Right mouse button always is "Erase": Is just more convenient.
    If Button = 2 Then
        gaeGrid(lX, lY) = Void
        Exit Sub
    End If
    
    'Set the new start, obstacle or target, or clear the cell, as per
    'active option button.
    If optSetStart.Value Then
        If lStartX <> -1 Then
            gaeGrid(lStartX, lStartY) = Void
            EraseCell lStartX, lStartY
        End If
        lStartX = lX: lStartY = lY
        gaeGrid(lX, lY) = Start
        DrawDisc lX, lY, &H8000&
    ElseIf optSetTarget.Value Then
        If lTargetX <> -1 Then
            gaeGrid(lTargetX, lTargetY) = Void
            EraseCell lTargetX, lTargetY
        End If
        lTargetX = lX: lTargetY = lY
        gaeGrid(lX, lY) = Target
        DrawDisc lX, lY, &H800000
    ElseIf optSetObstacle.Value Then
        gaeGrid(lX, lY) = Obstacle
        DrawDisc lX, lY, &H80&
    ElseIf optSetVoid.Value Then
        gaeGrid(lX, lY) = Void
    End If
End Sub

Private Sub picGrid_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    'If this was the start button, switch to Target, as it most
    'likely will be the next point set. If it was the target, switch
    'to obstacles.
    If optSetStart.Value Then
        optSetTarget.Value = True
    ElseIf optSetTarget.Value Then
        optSetObstacle.Value = True
    End If
End Sub

'Constructing a grid in picGrid as per txtW and txtH
Private Sub MakeGrid(ResizeOnly As Boolean)
    Dim lX As Long, lY As Long
    
    
    If Not ResizeOnly Then
        'Get dimensions of grid as per text fields txtW/H and store globally.
        glGridW = CLng(txtW.Text): glGridH = CLng(txtH.Text)
        
        'Construct matrix
        ReDim gaeGrid(0 To glGridW - 1, 0 To glGridH - 1) As eCell
    
        'Usually this would be defined in APlus. It is just here
        'because the test form wants to display infos about it's fields.
        'The two fields can also be merged into the source matrix.
        ReDim abGridCopy(0 To glGridW - 1, 0 To glGridH - 1) As tGrid
    End If
    
    'Define the edge length of a cell within the grid.
    glCellLen = picGrid.Width \ glGridW
    If picGrid.Height \ glGridH < glCellLen Then
        glCellLen = picGrid.Height \ glGridH
    End If
    
    'Clear old grid
    picGrid.Cls
    
    'Vertical lines
    lY = glGridH * glCellLen
    For lX = 0 To glGridW * glCellLen Step glCellLen
        picGrid.Line (lX, 0)-(lX, lY)
    Next lX

    'Horizontal lines
    lX = glGridW * glCellLen
    For lY = 0 To glGridH * glCellLen Step glCellLen
        picGrid.Line (0, lY)-(lX, lY)
    Next lY
        
    'Reconstruct the content
    If ResizeOnly Then
        For lY = 0 To glGridH - 1
            For lX = 0 To glGridW - 1
                'Voids don't need redrawing
                If gaeGrid(lX, lY) = Start Then
                    DrawDisc lX, lY, &H8000&
                ElseIf gaeGrid(lX, lY) = Obstacle Then
                    DrawDisc lX, lY, &H80&
                ElseIf gaeGrid(lX, lY) = Target Then
                    DrawDisc lX, lY, &H800000
                End If
            Next lX
        Next lY
    End If
    
    cmdMakeGrid.Enabled = False
End Sub

Private Sub ErasePath()
    Dim lX As Long, lY As Long
    
    cmdErasePath.Enabled = False
    
    'Reconstruct the content
    For lY = 0 To glGridH - 1
        For lX = 0 To glGridW - 1
            If gaeGrid(lX, lY) = Obstacle Then
                'Obstacles don't need redrawing
            ElseIf gaeGrid(lX, lY) = Start Then
                DrawDisc lX, lY, &H8000&
            ElseIf gaeGrid(lX, lY) = Target Then
                'Target might have parent indicator
                EraseCell lX, lY
                DrawDisc lX, lY, &H800000
            Else
                EraseCell lX, lY
            End If
        Next lX
    Next lY
End Sub

Private Sub EraseCell(X As Long, Y As Long)
    Dim lX As Long, lY As Long

    lX = X * glCellLen
    lY = Y * glCellLen
    picGrid.FillColor = picGrid.BackColor
    picGrid.Line (lX + 1, lY + 1)-(lX + glCellLen - 2, lY + glCellLen - 2), picGrid.BackColor, BF
End Sub

Private Sub DrawDisc(X As Long, Y As Long, Col As Long)
    Dim lX As Long, lY As Long
    
    'Visualize what was clicked
    lX = X * glCellLen + glCellLen \ 2
    lY = Y * glCellLen + glCellLen \ 2
    picGrid.FillColor = Col
    picGrid.Circle (lX, lY), glCellLen \ 2 - 2, Col
    
    'If both a start and a target exist, the algorithm can be run
    'if it is not already running
    If Not bRunning Then
        cmdFindPath.Enabled = (lStartX <> -1 And lTargetX <> -1)
    Else
        cmdFindPath.Enabled = False
    End If
End Sub

Private Sub ParentPointer(X As Long, Y As Long, H As Long, V As Long)
    Dim lX As Long, lY As Long
    Dim lPX As Long, lPY As Long
    
    'Small black disc
    lX = X * glCellLen + glCellLen \ 2
    lY = Y * glCellLen + glCellLen \ 2
    picGrid.FillColor = vbBlack
    picGrid.Circle (lX, lY), glCellLen \ 4, vbBlack
    
    'H and V tell us where the parent is: H=1 to the right, V=1 down, etc.
    'They can have the values -1, 0 and +1.
    'Calculate the coordinates of the line's tip
    lPX = lX + (glCellLen \ 2 - 1) * H
    lPY = lY + (glCellLen \ 2 - 1) * V
    
    'Draw the pointer towards the parent
    picGrid.Line (lX, lY)-(lPX, lPY)
End Sub

Private Sub OpenedCell(X As Long, Y As Long)
    Dim lX As Long, lY As Long
    
    'Tiny yellow disc
    lX = X * glCellLen + glCellLen \ 2
    lY = Y * glCellLen + glCellLen \ 2
    picGrid.FillColor = vbYellow
    picGrid.Circle (lX, lY), glCellLen \ 8, vbYellow
End Sub

Private Sub ClosedCell(X As Long, Y As Long)
    Dim lX As Long, lY As Long
    
    'Tiny red disc
    lX = X * glCellLen + glCellLen \ 2
    lY = Y * glCellLen + glCellLen \ 2
    picGrid.FillColor = vbRed
    picGrid.Circle (lX, lY), glCellLen \ 8, vbRed
End Sub

Private Sub BestPath(X As Long, Y As Long)
    Dim lX As Long, lY As Long
    
    'Tiny green disc
    lX = X * glCellLen + glCellLen \ 2
    lY = Y * glCellLen + glCellLen \ 2
    picGrid.FillColor = vbGreen
    picGrid.Circle (lX, lY), glCellLen \ 8, vbGreen
End Sub

Private Sub txtW_GotFocus()
    txtW.SelStart = 0
    txtW.SelLength = 100
    
    If Not bRunning Then staMain.Panels("Msg").Text = "Horizontal dimension of grid."
End Sub

Private Sub txtH_GotFocus()
    txtH.SelStart = 0
    txtH.SelLength = 100

    If Not bRunning Then staMain.Panels("Msg").Text = "Vertical dimension of grid."
End Sub

Private Sub txtW_Validate(Cancel As Boolean)
    If Not IsNumeric(txtW.Text) Then
        Cancel = True
    ElseIf CStr(CLng(txtW.Text)) <> txtW.Text Then
        Cancel = True
    ElseIf CLng(txtW.Text) <= 0 Then
        Cancel = True
    Else
        cmdMakeGrid.Enabled = txtW.Enabled
    End If
End Sub

Private Sub txtH_Validate(Cancel As Boolean)
    If Not IsNumeric(txtH.Text) Then
        Cancel = True
    ElseIf CStr(CLng(txtH.Text)) <> txtH.Text Then
        Cancel = True
    ElseIf CLng(txtH.Text) <= 0 Then
        Cancel = True
    Else
        cmdMakeGrid.Enabled = txtH.Enabled
    End If
End Sub

'Waiting for F8 or click on cmdSingleStep
Public Sub WaitForSingleStep()
    cmdSingleStep.Enabled = True
    cmdSingleStep.SetFocus
    bSingleStep = True
    WaitForKey vbKeyF8
    bSingleStep = False
    bDoSingleStep = False
    cmdSingleStep.Enabled = False
End Sub

Public Sub WaitForKey(ByVal Key As KeyCodeConstants)
    'Get status of key
    'First wait to be pressed
    Do While Not CBool(GetAsyncKeyState(Key) And &H8000) And Not bDoSingleStep
        Sleep 50
        DoEvents
    Loop
    'Then to be released
    Do While CBool(GetAsyncKeyState(Key) And &H8000) And Not bDoSingleStep
        Sleep 50
        DoEvents
    Loop
End Sub





'The algorithm expects all data to be in the matrix gaeGrid()
'(first dimension being the X axis, second dimension the Y axis).
'SX/SY are the start coordinates (from where we want to reach the target),
'TX/TY are the target coordinates (which we want to reach from the start).
'We may cross only cells which have the value as per the FreeCell argument.
'Returns True if there is a path from SX/SY to TX/TY, else False.
'If there is a path, it can be found in the "argument" Path() which must be
'of type tPoint; also the total Cost is returned.
Private Function APlus(SX As Long, SY As Long, TX As Long, TY As Long, FreeCell As eCell, Path() As tPoint, Cost As Single) As Boolean
    'A+ Pathfinding Algorithm:
    'Implementation by Herbert Glarner (herbert.glarner@bluewin.ch)
    'Unlimited use for whatever purpose allowed provided that above credits are given.
    'Suggestions and bug reports welcome.
    Dim lMaxList As Long
    Dim lActList As Long
    Dim sCheapCost As Single, lCheapIndex As Long
    Dim sTotalCost As Single
    Dim lCheapX As Long, lCheapY As Long
    Dim lOffX As Long, lOffY As Long
    Dim lTestX As Long, lTestY As Long
    Dim lMaxX As Long, lMaxY As Long
    Dim sAdditCost As Single
    Dim lPathPtr As Long
    
    'The test program wants to access this grid. For this reason it is defined
    'and initialized globally. Usually one would define and initialize it only
    'in this procedure.
    'The two fields of tGrid can also be merged into the source matrix.
    '   Dim abGridCopy() As tGrid
    
    Const cSqr2 As Single = 1.4142135623731
    
    'Define the upper boundaries of the grid.
    lMaxX = UBound(gaeGrid, 1): lMaxY = UBound(gaeGrid, 2)
    
    'For each cell of the grid a bit is defined to hold it's "closed" status
    'and the index to the Open-List.
    'The test program wants to access this grid. For this reason it is defined
    'and initialized globally. Usually one would define and initialize it only
    'in this procedure. (Don't omit here: we need an empty matrix.)
    ReDim abGridCopy(0 To lMaxX, 0 To lMaxY) As tGrid
    
    'The starting point is added to the working list. It has no parent (-1).
    'The cost to get here is 0 (we start here). The direct distance enters
    'the Heuristic.
    ReDim grList(0 To 0) As tCell
    With grList(0)
        .X = SX: .Y = SY: .Parent = -1: .Cost = 0
        .Heuristic = Sqr((TX - SX) * (TX - SX) + (TY - SY) * (TY - SY))
    End With
    
    'Start the algorithm
    Do
        'Get the cell with the lowest Cost+Heuristic. Initialize the cheapest cost
        'with an impossible high value (change as needed). The best found index
        'is set to -1 to indicate "none found".
        sCheapCost = 10000000
        lCheapIndex = -1
        'Check all cells of the list. Initially, there is only the start point,
        'but more will be added soon.
        For lActList = 0 To lMaxList
            'Only check if not closed already.
            If Not grList(lActList).Closed Then
                'If this cells total cost (Cost+Heuristic) is lower than the so
                'far lowest cost, then store this total cost and the cell's index
                'as the so far best found.
                sTotalCost = grList(lActList).Cost + grList(lActList).Heuristic
                If sTotalCost < sCheapCost Then
                    'New cheapest cost found.
                    sCheapCost = sTotalCost: lCheapIndex = lActList
                End If
            End If
        Next lActList
        
        'lCheapIndex contains the cell with the lowest total cost now.
        'If no such cell could be found, all cells were already closed and there
        'is no path at all to the target.
        If lCheapIndex = -1 Then
            'There is no path.
            APlus = False: Exit Function
        End If
        
        'Get the cheapest cell's coordinates
        lCheapX = grList(lCheapIndex).X
        lCheapY = grList(lCheapIndex).Y
        
        'If the best field is the target field, we have found our path.
        If lCheapX = TX And lCheapY = TY Then
            'Path found.
            Exit Do
        End If
        
'--- Remove if used standalone ---
'Visualize opening
OpenedCell lCheapX, lCheapY
Refresh
If chkSingleStep.Value Then
    staMain.Panels("Msg").Text = "Opened best cell " & CStr(lCheapX) & "/" & CStr(lCheapY) & " (cheapest total cost " & Format$(sCheapCost, "0.00") & ")"
    WaitForSingleStep
End If
'---------------------------------
        
        'Check all immediate neighbors
        For lOffY = -1 To 1
            For lOffX = -1 To 1
                'Ignore the actual field, process all others (8 neighbors).
                If lOffX <> 0 Or lOffY <> 0 Then
                    'Get the neighbor's coordinates.
                    lTestX = lCheapX + lOffX: lTestY = lCheapY + lOffY
                    'Don't test beyond the grid's boundaries.
                    If lTestX >= 0 And lTestX <= lMaxX And lTestY >= 0 And lTestY <= lMaxY Then
                        'The cell is within the grid's boundaries.
                        'Make sure the field is accessible. To be accessible,
                        'the cell must have the value as per the function
                        'argument FreeCell (change as needed). Of course, the
                        'target is allowed as well.
                        If gaeGrid(lTestX, lTestY) = FreeCell Or gaeGrid(lTestX, lTestY) = Target Then
                            'The cell is accessible.
                            'For this we created the "bitmatrix" abGridCopy().
                            If abGridCopy(lTestX, lTestY).ListStat = Unprocessed Then
                                'Register the new cell in the list.
                                lMaxList = lMaxList + 1
                                ReDim Preserve grList(0 To lMaxList) As tCell
                                With grList(lMaxList)
                                    'The parent is where we come from (the cheapest field);
                                    'it's index is registered.
                                    .X = lTestX: .Y = lTestY: .Parent = lCheapIndex
                                    'Additional cost is 1 for othogonal movement, cSqr2 for
                                    'diagonal movement (change if diagonal steps should have
                                    'a different cost).
                                    If Abs(lOffX) + Abs(lOffY) = 1 Then sAdditCost = 1# Else sAdditCost = cSqr2
                                    'Store cost to get there by summing the actual cell's cost
                                    'and the additional cost.
                                    .Cost = grList(lCheapIndex).Cost + sAdditCost
                                    'Calculate distance to target as the heuristical part
                                    .Heuristic = Sqr((TX - lTestX) * (TX - lTestX) + (TY - lTestY) * (TY - lTestY))
                                End With
                                'Register in the Grid copy as open.
                                abGridCopy(lTestX, lTestY).ListStat = IsOpen
                                'Also register the index to quickly find the element in the
                                '"closed" list.
                                abGridCopy(lTestX, lTestY).Index = lMaxList

'--- Remove if used standalone ---
'Visualize the added field by pointing to the parent
ParentPointer lTestX, lTestY, -lOffX, -lOffY
Refresh
If chkSingleStep.Value Then
    staMain.Panels("Msg").Text = "Added cell " & CStr(lTestX) & "/" & CStr(lTestY)
    WaitForSingleStep
End If
'---------------------------------

                            ElseIf abGridCopy(lTestX, lTestY).ListStat = IsOpen Then
                                'Is the cost to get to this already open field cheaper when using
                                'this path via lTestX/lTestY ?
                                lActList = abGridCopy(lTestX, lTestY).Index
                                sAdditCost = IIf(Abs(lOffX) + Abs(lOffY) = 1, 1#, cSqr2)
                                If grList(lCheapIndex).Cost + sAdditCost < grList(lActList).Cost Then
                                    'The cost to reach the already open field is lower via the
                                    'actual field.
                                    
                                    'Store new cost
                                    grList(lActList).Cost = grList(lCheapIndex).Cost + sAdditCost
                                    'Store new parent
                                    grList(lActList).Parent = lCheapIndex

'--- Remove if used standalone ---
'Visualize the changed parent
EraseCell lTestX, lTestY
ParentPointer lTestX, lTestY, -lOffX, -lOffY
Refresh
If chkSingleStep.Value Then
    staMain.Panels("Msg").Text = "Changed parent for cell " & CStr(lTestX) & "/" & CStr(lTestY)
    WaitForSingleStep
End If
'---------------------------------

                                End If
                            'ElseIf abGridCopy(lTestX, lTestY) = IsClosed Then
                            '   'This cell can be ignored
                            End If
                        End If
                    End If
                End If
            Next lOffX
        Next lOffY
        'Close the just checked cheapest cell.
        grList(lCheapIndex).Closed = True
        abGridCopy(lCheapX, lCheapY).ListStat = IsClosed

'--- Remove if used standalone ---
'Visualize closing
ClosedCell lCheapX, lCheapY
Refresh
If chkSingleStep.Value Then
    staMain.Panels("Msg").Text = "Closed cell " & CStr(lCheapX) & "/" & CStr(lCheapY)
    WaitForSingleStep
End If
'---------------------------------

    Loop
        
    'We arrive here only when a path was found.
    'Provide the total cost in the "argument" Cost.
    Cost = grList(lCheapIndex).Cost
    
'--- Remove if used standalone ---
'Inform about path cost
BestPath grList(lCheapIndex).X, grList(lCheapIndex).Y
Refresh
If chkSingleStep.Value Then
    staMain.Panels("Msg").Text = "Best path found (total cost " & Format(Cost, "0.00") & ")"
    WaitForSingleStep
End If
'---------------------------------
    
    'The path can be found by backtracing from the field TX/TY until SX/SY.
    'The path is traversed in backwards order and stored reversely (!) in
    'the "argument" Path().
    ReDim Path(0 To 0) As tPoint
    lPathPtr = -1
    'lCheapIndex (lCheapX/Y) initially contains the target TX/TY
    Do
        'Store the coordinates of the current cell
        lPathPtr = lPathPtr + 1
        ReDim Preserve Path(0 To lPathPtr) As tPoint
        Path(lPathPtr).X = grList(lCheapIndex).X
        Path(lPathPtr).Y = grList(lCheapIndex).Y

'--- Remove if used standalone ---
'Visualize best path
BestPath Path(lPathPtr).X, Path(lPathPtr).Y
Refresh
If chkSingleStep.Value Then
    staMain.Panels("Msg").Text = "Backtracing best path " & CStr(Path(lPathPtr).X) & "/" & CStr(Path(lPathPtr).Y)
    WaitForSingleStep
End If
'---------------------------------
        
        'Follow the parent
        lCheapIndex = grList(lCheapIndex).Parent
    Loop While lCheapIndex <> -1
    
    APlus = True: Exit Function
End Function

