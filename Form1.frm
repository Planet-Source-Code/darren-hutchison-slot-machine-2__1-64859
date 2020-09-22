VERSION 5.00
Begin VB.Form Form1 
   BackColor       =   &H00808000&
   Caption         =   "Slot Machine 2"
   ClientHeight    =   8115
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   6450
   LinkTopic       =   "Form1"
   ScaleHeight     =   541
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   430
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer tmrRemainder 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   120
      Top             =   4320
   End
   Begin VB.Timer tmrRotateReel 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   120
      Top             =   3720
   End
   Begin VB.Timer tmrHold 
      Enabled         =   0   'False
      Interval        =   500
      Left            =   120
      Top             =   3120
   End
   Begin VB.Timer tmrAddScore 
      Enabled         =   0   'False
      Interval        =   50
      Left            =   120
      Top             =   2520
   End
   Begin VB.Timer tmrLights 
      Enabled         =   0   'False
      Interval        =   100
      Left            =   120
      Top             =   1920
   End
   Begin VB.Timer tmrSpin 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   120
      Top             =   1320
   End
   Begin VB.Timer tmrStart 
      Enabled         =   0   'False
      Interval        =   500
      Left            =   120
      Top             =   720
   End
   Begin VB.Timer tmrRefreshDisplay 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   120
      Top             =   120
   End
   Begin VB.PictureBox picBox 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Height          =   2865
      Index           =   2
      Left            =   4185
      ScaleHeight     =   191
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   63
      TabIndex        =   5
      Top             =   1710
      Width           =   945
   End
   Begin VB.PictureBox picBox 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Height          =   2865
      Index           =   1
      Left            =   3000
      ScaleHeight     =   2865
      ScaleWidth      =   945
      TabIndex        =   4
      Top             =   1710
      Width           =   945
   End
   Begin VB.PictureBox picBox 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Height          =   2865
      Index           =   0
      Left            =   1845
      ScaleHeight     =   191
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   63
      TabIndex        =   3
      Top             =   1710
      Width           =   945
   End
   Begin VB.Label lblCancel 
      Alignment       =   2  'Center
      BackColor       =   &H00808000&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "CANCEL"
      Height          =   255
      Left            =   960
      TabIndex        =   19
      Top             =   6960
      Width           =   735
   End
   Begin VB.Label lblSelector 
      BackStyle       =   0  'Transparent
      Caption         =   "è"
      BeginProperty Font 
         Name            =   "Wingdings"
         Size            =   12
         Charset         =   2
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   255
      Index           =   2
      Left            =   1125
      TabIndex        =   18
      Top             =   4005
      Width           =   255
   End
   Begin VB.Label lblSelector 
      BackStyle       =   0  'Transparent
      Caption         =   "è"
      BeginProperty Font 
         Name            =   "Wingdings"
         Size            =   12
         Charset         =   2
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   255
      Index           =   1
      Left            =   1125
      TabIndex        =   17
      Top             =   3030
      Width           =   255
   End
   Begin VB.Label lblSelector 
      BackStyle       =   0  'Transparent
      Caption         =   "è"
      BeginProperty Font 
         Name            =   "Wingdings"
         Size            =   12
         Charset         =   2
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   255
      Index           =   0
      Left            =   1125
      TabIndex        =   16
      Top             =   2055
      Width           =   255
   End
   Begin VB.Label lblMessage 
      Alignment       =   2  'Center
      BackColor       =   &H80000007&
      Caption         =   "Press Start to Begin"
      ForeColor       =   &H0000FF00&
      Height          =   255
      Left            =   2040
      TabIndex        =   15
      Top             =   7440
      Width           =   2655
   End
   Begin VB.Label lblLight 
      BackColor       =   &H00808080&
      BorderStyle     =   1  'Fixed Single
      Height          =   135
      Index           =   5
      Left            =   1080
      TabIndex        =   14
      Top             =   5160
      Width           =   375
   End
   Begin VB.Label lblLight 
      BackColor       =   &H00808080&
      BorderStyle     =   1  'Fixed Single
      Height          =   135
      Index           =   4
      Left            =   1080
      TabIndex        =   13
      Top             =   5400
      Width           =   375
   End
   Begin VB.Label lblLight 
      BackColor       =   &H00808080&
      BorderStyle     =   1  'Fixed Single
      Height          =   135
      Index           =   3
      Left            =   1080
      TabIndex        =   12
      Top             =   5640
      Width           =   375
   End
   Begin VB.Label lblLight 
      BackColor       =   &H00808080&
      BorderStyle     =   1  'Fixed Single
      Height          =   135
      Index           =   2
      Left            =   1080
      TabIndex        =   11
      Top             =   5880
      Width           =   375
   End
   Begin VB.Label lblLight 
      BackColor       =   &H00808080&
      BorderStyle     =   1  'Fixed Single
      Height          =   135
      Index           =   1
      Left            =   1080
      TabIndex        =   10
      Top             =   6120
      Width           =   375
   End
   Begin VB.Label lblLight 
      BackColor       =   &H00808080&
      BorderStyle     =   1  'Fixed Single
      Height          =   135
      Index           =   0
      Left            =   1080
      TabIndex        =   9
      Top             =   6360
      Width           =   375
   End
   Begin VB.Label lblWin 
      Alignment       =   2  'Center
      BackColor       =   &H80000009&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   255
      Left            =   3240
      TabIndex        =   8
      Top             =   1020
      Width           =   570
   End
   Begin VB.Label lblScore 
      Alignment       =   2  'Center
      BackColor       =   &H80000009&
      Caption         =   "75"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   3240
      TabIndex        =   7
      Top             =   675
      Width           =   570
   End
   Begin VB.Label lblStart 
      Alignment       =   2  'Center
      BackColor       =   &H000000C0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "START"
      Height          =   255
      Left            =   4920
      TabIndex        =   6
      Top             =   7440
      Width           =   615
   End
   Begin VB.Image Image2 
      Height          =   1485
      Left            =   1830
      Picture         =   "Form1.frx":0000
      Top             =   5070
      Width           =   3315
   End
   Begin VB.Label lblHold 
      Alignment       =   2  'Center
      BackColor       =   &H0000C0C0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "HOLD"
      Height          =   255
      Index           =   2
      Left            =   4440
      TabIndex        =   2
      Top             =   6960
      Width           =   615
   End
   Begin VB.Label lblHold 
      Alignment       =   2  'Center
      BackColor       =   &H0000C0C0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "HOLD"
      Height          =   255
      Index           =   1
      Left            =   3240
      TabIndex        =   1
      Top             =   6960
      Width           =   615
   End
   Begin VB.Label lblHold 
      Alignment       =   2  'Center
      BackColor       =   &H0000C0C0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "HOLD"
      Height          =   255
      Index           =   0
      Left            =   2040
      TabIndex        =   0
      Top             =   6960
      Width           =   615
   End
   Begin VB.Image Image1 
      Height          =   8550
      Left            =   600
      Picture         =   "Form1.frx":5AE2
      Top             =   120
      Width           =   5205
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim DX As New DirectX7
Dim DD As DirectDraw7
Dim picBuffer As DirectDrawSurface7

Dim srfPrimary As DirectDrawSurface7
Dim dscPrimary As DDSURFACEDESC2
Dim srfBackBuffer As DirectDrawSurface7
Dim dscBackBuffer As DDSURFACEDESC2

Dim clpClipper(3) As DirectDrawClipper

Dim srfReel(3) As DirectDrawSurface7
Dim dscDescription As DDSURFACEDESC2

Dim rctSource As RECT
Dim rctDest As RECT
Dim intTop(3) As Integer
Dim intCount(3) As Integer
Dim intStartPos(3) As Integer
Dim intMax(3) As Integer
Dim intOffset(3) As Integer
Dim intResult(3) As Integer
Dim booSoundPlayed(3) As Boolean
Dim intLightCount As Integer
Dim booLightForward As Boolean
Dim intScoreCount As Integer
Dim intWinningAmount As Integer
Dim booHoldActive As Boolean
Dim booHold(3) As Boolean
Dim booCancel As Boolean
Dim intFinishedLines As Integer
Dim booBonusLight As Boolean
Dim booRotateReel As Boolean
Dim Winning(0 To 2) As Integer
Dim booNoCheck As Boolean

Option Explicit

Private Sub form_load()

Dim i As Integer

    Randomize
    
    Set DD = DX.DirectDrawCreate("")
    DD.SetCooperativeLevel Me.hWnd, DDSCL_NORMAL

    dscPrimary.lFlags = DDSD_CAPS
    dscPrimary.ddsCaps.lCaps = DDSCAPS_PRIMARYSURFACE
    Set srfPrimary = DD.CreateSurface(dscPrimary)
    
    dscBackBuffer.lFlags = DDSD_CAPS Or DDSD_WIDTH Or DDSD_HEIGHT
    dscBackBuffer.ddsCaps.lCaps = DDSCAPS_OFFSCREENPLAIN Or DDSCAPS_3DDEVICE
    dscBackBuffer.lWidth = 64 * 3
    dscBackBuffer.lHeight = 1472

    Set srfBackBuffer = DD.CreateSurface(dscBackBuffer)
    
    dscDescription.lFlags = DDSD_CAPS
    dscDescription.ddsCaps.lCaps = DDSCAPS_OFFSCREENPLAIN
    
    Set clpClipper(0) = DD.CreateClipper(0)
    clpClipper(0).SetHWnd picBox(0).hWnd
    srfPrimary.SetClipper clpClipper(0)
        
    Set clpClipper(1) = DD.CreateClipper(0)
    clpClipper(1).SetHWnd picBox(1).hWnd
    srfPrimary.SetClipper clpClipper(1)
    
    Set clpClipper(2) = DD.CreateClipper(0)
    clpClipper(2).SetHWnd picBox(2).hWnd
    srfPrimary.SetClipper clpClipper(2)
    
    Set srfReel(0) = DD.CreateSurfaceFromFile(App.Path & "\Wheel1.bmp", dscDescription)
    Set srfReel(1) = DD.CreateSurfaceFromFile(App.Path & "\Wheel2.bmp", dscDescription)
    Set srfReel(2) = DD.CreateSurfaceFromFile(App.Path & "\Wheel3.bmp", dscDescription)
    
    Me.Show
    tmrStart.Enabled = True
    intCount(0) = 0
    intCount(1) = 0
    intCount(2) = 0
    intLightCount = 0
    booLightForward = True
    booNoCheck = False
    
    For i = 0 To 2
        booSoundPlayed(i) = False
        intStartPos(i) = Int(Rnd * 20)
    Next i
    booHoldActive = False
    booCancel = False
    booBonusLight = False
    booRotateReel = False
    
    intTop(0) = intStartPos(0) * 64
    intTop(1) = intStartPos(1) * 64
    intTop(2) = intStartPos(2) * 64
    intFinishedLines = 0
    rctSource.Top = 0: rctSource.Left = 0
    rctSource.Right = 63: rctSource.Bottom = 1472
    For i = 0 To 2
        rctDest.Top = 0: rctDest.Left = i * 64
        rctDest.Right = 63 + rctDest.Left: rctDest.Bottom = 1472
        srfBackBuffer.Blt rctDest, srfReel(i), rctSource, DDBLT_WAIT
    Next i

    Do
        DisplayReel
        DoEvents
    Loop Until False
    
End Sub

Private Sub DisplayReel()

Dim i As Integer

    For i = 0 To 2
        rctSource.Top = intTop(i): rctSource.Left = i * 64
        rctSource.Right = 63 + rctSource.Left: rctSource.Bottom = (64 * 3) + intTop(i)
        DX.GetWindowRect picBox(i).hWnd, rctDest
        srfPrimary.SetClipper clpClipper(i)
        srfPrimary.Blt rctDest, srfBackBuffer, rctSource, DDBLT_WAIT
    Next i

End Sub

Private Sub form_unload(cancel As Integer)

    Set srfBackBuffer = Nothing
    Set srfPrimary = Nothing
    Set DD = Nothing
    Set DX = Nothing
    Unload Me
    End
    
End Sub

Private Sub lblCancel_Click()
Dim i As Integer

    If booCancel = False Then Exit Sub
    For i = 0 To 2
        lblHold(i).BackColor = &HC0C0&
        booHold(i) = False
    Next i
    booCancel = False
    lblCancel.BackColor = &H808000
    
End Sub

Private Sub lblHold_Click(Index As Integer)

    If booHoldActive = False Then Exit Sub
    If booHold(Index) = False Then sndPlaySound App.Path & "\ploff.wav", &H1
    booHold(Index) = True
    lblHold(Index).BackColor = vbYellow
    booCancel = True
    
End Sub

Private Sub lblStart_Click()

Dim i As Integer

    If tmrStart.Enabled = False Then Exit Sub
    
    If booRotateReel = True Then
        tmrStart.Enabled = False
        tmrRotateReel.Enabled = False
        tmrRemainder.Enabled = True
        booNoCheck = True
        Exit Sub
    End If
    
    If booBonusLight = True Then
        If intLightCount = 6 Then
            lblMessage.Caption = "Well done! It gets faster"
            intLightCount = 0
            tmrLights.Enabled = False
            lblScore.Caption = Str(Val(lblScore.Caption) + 2)
            lblWin.Caption = "+ 2"
            sndPlaySound App.Path & "\tik.wav", &H1
            For i = 0 To 5
                lblLight(i).BackColor = &H808080
            Next i
            tmrLights.Interval = tmrLights.Interval - 10
            tmrLights.Enabled = True
            Exit Sub
        Else
            lblMessage.Caption = "You missed!"
            booBonusLight = False
            intLightCount = 0
            tmrLights.Enabled = False
            For i = 0 To 5
                lblLight(i).BackColor = &H808080
            Next i
            tmrLights.Interval = 100
            Exit Sub
        End If
    End If
    
    If Val(lblScore.Caption) <= 4 Then lblMessage.Caption = "No credits left! Game over!": tmrStart.Enabled = False: Exit Sub
    
    lblCancel.BackColor = &H808000
    lblMessage.Caption = "Spinning....."
    If booHoldActive = True Then
        For i = 0 To 2
            If booHold(i) = False Then lblHold(i).BackColor = &HC0C0&
        Next i
        booCancel = False
        tmrHold.Enabled = False
    End If
    
    tmrStart.Enabled = False: lblStart.BackColor = &HC0&
    sndPlaySound App.Path & "\ploff.wav", &H1
    For i = 0 To 2
        booSoundPlayed(i) = False
        intCount(i) = 0
    Next i

    intMax(0) = 20 + (Int(Rnd * 20) * 2)
    intMax(1) = 50 + (Int(Rnd * 20) * 2)
    intMax(2) = 80 + (Int(Rnd * 20) * 2)

    lblWin.Caption = ""
    lblScore.Caption = Str(Val(lblScore.Caption) - 5)
    SpinReel
    
End Sub

Private Sub tmrHold_Timer()

Dim i As Integer

    For i = 0 To 2
        If booHold(i) = False Then If lblHold(i).BackColor = vbYellow Then lblHold(i).BackColor = &HC0C0& Else lblHold(i).BackColor = vbYellow
    Next i
    If booCancel = True Then If lblCancel.BackColor = &H808000 Then lblCancel.BackColor = vbCyan Else lblCancel.BackColor = &H808000
        
End Sub

Private Sub tmrLights_Timer()

    If booLightForward = True Then
        If intLightCount < 6 Then
            lblLight(intLightCount).BackColor = vbRed
            intLightCount = intLightCount + 1
        Else
            booLightForward = False
        End If
    Else
        If intLightCount > 0 Then
            intLightCount = intLightCount - 1
            lblLight(intLightCount).BackColor = &H808080
        Else
            booLightForward = True
        End If
    End If
    
End Sub

Private Sub tmrRemainder_Timer()

    If intTop(2) Mod 64 <> 0 Then
        intTop(2) = intTop(2) - 2
        If intTop(2) < 0 Then intTop(2) = 1280 - 2
        If intTop(2) Mod 32 = 0 Then intCount(2) = intCount(2) + 1
    Else
        tmrRemainder.Enabled = False
        booRotateReel = False
        sndPlaySound App.Path & "\plup.wav", &H1
        CalculateBonusOffset
        Win
    End If
    
End Sub

Private Sub tmrRotateReel_Timer()

    intTop(2) = intTop(2) - 8
    If intTop(2) < 0 Then intTop(2) = 1280 - 8
    If intTop(2) Mod 32 = 0 Then intCount(2) = intCount(2) + 1
    
End Sub

Private Sub tmrSpin_Timer()

Dim i As Integer, j As Integer, intLinesActive As Integer

    intLinesActive = 0
    For i = 0 To 2
        If booHold(i) = False Then intLinesActive = intLinesActive + 1
    Next i

    For i = 0 To 2
        If booHold(i) = False Then
            If intCount(i) < intMax(i) Then
                intTop(i) = intTop(i) - 32
                If intTop(i) < 0 Then intTop(i) = 1280 - 32
                intCount(i) = intCount(i) + 1
            ElseIf intCount(i) = intMax(i) And booSoundPlayed(i) = False Then
                sndPlaySound App.Path & "\plup.wav", &H1
                booSoundPlayed(i) = True
                intFinishedLines = intFinishedLines + 1
            End If
        End If
    Next i

    If intCount(2) >= intMax(2) Or intLinesActive = intFinishedLines Then
        sndPlaySound App.Path & "\plup.wav", &H1
        tmrSpin.Enabled = False
        CalculateOffsets
        Win
        If tmrAddScore.Enabled = False Then
            CheckFeatures
            intCount(0) = 0: intCount(1) = 0: intCount(2) = 0
            For i = 0 To 2
                booHold(i) = False
                lblHold(i).BackColor = &HC0C0&
            Next i
            intFinishedLines = 0
            StartFeatures
        End If
    End If
    
End Sub

Private Sub CalculateOffsets()

Dim i As Integer

    For i = 0 To 2
        intOffset(i) = 0
        intResult(i) = 0
    Next i
    
    For i = 0 To 2
        intOffset(i) = (intCount(i) / 2) Mod 20
        If intOffset(i) < intStartPos(i) Then
            intResult(i) = intStartPos(i) - intOffset(i)
            intStartPos(i) = intResult(i)
        Else
            intOffset(i) = intOffset(i) - intStartPos(i)
            intResult(i) = 20 - intOffset(i)
            If intResult(i) = 20 Then intResult(i) = 0
            intStartPos(i) = intResult(i)
        End If
    Next i
    
End Sub

Private Sub CalculateBonusOffset()

    intOffset(2) = 0
    intResult(2) = 0
    
    intOffset(2) = (intCount(2) / 2) Mod 20
    If intOffset(2) < intStartPos(2) Then
        intResult(2) = intStartPos(2) - intOffset(2)
        intStartPos(2) = intResult(2)
    Else
        intOffset(2) = intOffset(2) - intStartPos(2)
        intResult(2) = 20 - intOffset(2)
        If intResult(2) = 20 Then intResult(2) = 0
        intStartPos(2) = intResult(2)
    End If
    
End Sub

Private Sub tmrStart_Timer()

    If lblStart.BackColor = vbRed Then lblStart.BackColor = &HC0& Else lblStart.BackColor = vbRed
    
End Sub

Private Sub SpinReel()

    tmrSpin.Enabled = True
    
End Sub

Function Wheel1(R As Integer) As String
    
    Select Case R
        Case 4, 7, 13, 19: Wheel1 = "Blueberry"
        Case 0, 2, 5, 8, 12, 18, 20: Wheel1 = "Orange"
        Case 1, 10, 14, 21: Wheel1 = "Bell"
        Case 11, 16: Wheel1 = "Cherry"
        Case 9, 15: Wheel1 = "1Bar"
        Case 3: Wheel1 = "2Bar"
        Case 6: Wheel1 = "3Bar"
        Case 17: Wheel1 = "7"
    End Select
    
End Function

Function Wheel2(R As Integer) As String

    Select Case R
        Case 1, 6, 17, 21: Wheel2 = "Blueberry"
        Case 0, 5, 8, 13, 15, 18, 20: Wheel2 = "Cherry"
        Case 2, 10, 19: Wheel2 = "Orange"
        Case 3, 9, 16: Wheel2 = "Bell"
        Case 4, 11: Wheel2 = "1Bar"
        Case 14: Wheel2 = "2Bar"
        Case 7: Wheel2 = "3Bar"
        Case 12: Wheel2 = "7"
    End Select
    
End Function

Function Wheel3(R As Integer) As String

    Select Case R
        Case 0, 5, 9, 12, 14, 19, 20: Wheel3 = "Cherry"
        Case 1, 7, 11, 13, 21: Wheel3 = "Orange"
        Case 2, 10, 18: Wheel3 = "Bell"
        Case 3, 6, 15: Wheel3 = "Blueberry"
        Case 16: Wheel3 = "1Bar"
        Case 4: Wheel3 = "2Bar"
        Case 8: Wheel3 = "3Bar"
        Case 17: Wheel3 = "7"
    End Select
    
End Function

Sub Win()
    
    Dim B1 As Integer, B2 As Integer, B3 As Integer
    Dim i As Long
    Dim RolA(0 To 2) As String
    Dim RolB(0 To 2) As String
    Dim RolC(0 To 2) As String
    Dim Bar1 As Boolean, Bar2 As Boolean, Bar3 As Boolean, booHasWon As Boolean, intTotalWinnings As Integer
   
    RolA(0) = Wheel1(intResult(0))
    RolA(1) = Wheel1(intResult(0) + 1)
    RolA(2) = Wheel1(intResult(0) + 2)
    
    RolB(0) = Wheel2(intResult(1))
    RolB(1) = Wheel2(intResult(1) + 1)
    RolB(2) = Wheel2(intResult(1) + 2)
    
    RolC(0) = Wheel3(intResult(2))
    RolC(1) = Wheel3(intResult(2) + 1)
    RolC(2) = Wheel3(intResult(2) + 2)
    For i = 0 To 2
        Winning(i) = 0
    Next i

    i = 0
    Do Until i = 3
        If RolA(i) = "Cherry" And RolB(i) <> "Cherry" Then Winning(i) = 2
        If RolA(i) = "Cherry" And RolB(i) = "Cherry" And RolC(i) <> "Cherry" Then Winning(i) = 5
        If RolA(i) = "Cherry" And RolB(i) = "Cherry" And RolC(i) = "Cherry" Then Winning(i) = 10
        If RolA(i) = "Orange" And RolB(i) = "Orange" And RolC(i) = "Orange" Then Winning(i) = 10
        If RolA(i) = "Orange" And RolB(i) = "Orange" And RolC(i) = "3Bar" Then Winning(i) = 10
        If RolA(i) = "Blueberry" And RolB(i) = "Blueberry" And RolC(i) = "Blueberry" Then Winning(i) = 14
        If RolA(i) = "Blueberry" And RolB(i) = "Blueberry" And RolC(i) = "3Bar" Then Winning(i) = 14
        If RolA(i) = "Bell" And RolB(i) = "Bell" And RolC(i) = "Bell" Then Winning(i) = 18
        If RolA(i) = "Bell" And RolB(i) = "Bell" And RolC(i) = "3Bar" Then Winning(i) = 18

        If RolA(i) = "1Bar" And RolB(i) = "1Bar" And RolC(i) = "1Bar" Then
            Winning(i) = 50
            Bar1 = True
        Else
            Bar1 = False
        End If
        If RolA(i) = "2Bar" And RolB(i) = "2Bar" And RolC(i) = "2Bar" Then
            Winning(i) = 100
            Bar2 = True
        Else
            Bar2 = False
        End If
        If RolA(i) = "3Bar" And RolB(i) = "3Bar" And RolC(i) = "3Bar" Then
            Winning(i) = 150
            Bar3 = True
        Else
            Bar3 = False
        End If
        If Bar1 = False And Bar2 = False And Bar3 = False Then
            If RolA(i) = "1Bar" Or RolA(i) = "2Bar" Or RolA(i) = "3Bar" Then
               If RolB(i) = "1Bar" Or RolB(i) = "2Bar" Or RolB(i) = "3Bar" Then
                   If RolC(i) = "1Bar" Or RolC(i) = "2Bar" Or RolC(i) = "3Bar" Then
                        Winning(i) = 25
                   End If
               End If
            End If
        End If
        If RolA(i) = "7" And RolB(i) = "7" And RolC(i) <> "7" Then Winning(i) = 250
        If RolA(i) = "7" And RolB(i) = "7" And RolC(i) = "7" Then Winning(i) = 500
    i = i + 1
    Loop

    intTotalWinnings = 0
    booHasWon = False
    For i = 0 To 2
        If Winning(i) > 0 Then
            lblSelector(i).ForeColor = vbYellow
            intTotalWinnings = intTotalWinnings + Winning(i)
            booHasWon = True
        End If
    Next i
    
    If booHasWon = True Then
        intScoreCount = 0
        Counter intTotalWinnings
        If intTotalWinnings > 25 Then lblMessage.Caption = "Won " & Str(intTotalWinnings) & "p" & " Woo!" Else lblMessage.Caption = "Won " & Str(intTotalWinnings) & "p"
    Else
        lblMessage.Caption = "No win. Spin again.": booNoCheck = False: tmrStart.Enabled = True
    End If
    
End Sub

Private Sub StartFeatures()

    If booRotateReel = True Then
        lblMessage.Caption = "Feature - stop at correct reel."
        sndPlaySound App.Path & "\beep.wav", &H1
        tmrRotateReel.Enabled = True
        Exit Sub
    End If
    If booBonusLight = True Then
        tmrLights.Enabled = True
        lblMessage.Caption = "Feature - stop at top light."
        sndPlaySound App.Path & "\beep.wav", &H1
    End If
    
End Sub

Private Sub CheckFeatures()
    
    Dim RolA(0 To 2) As String
    Dim RolB(0 To 2) As String
    Dim RolC(0 To 2) As String
    
    RolA(0) = Wheel1(intResult(0))
    RolA(1) = Wheel1(intResult(0) + 1)
    RolA(2) = Wheel1(intResult(0) + 2)
    
    RolB(0) = Wheel2(intResult(1))
    RolB(1) = Wheel2(intResult(1) + 1)
    RolB(2) = Wheel2(intResult(1) + 2)
    
    RolC(0) = Wheel3(intResult(2))
    RolC(1) = Wheel3(intResult(2) + 1)
    RolC(2) = Wheel3(intResult(2) + 2)
       
    If booNoCheck = False Then
        If RolA(1) = "Bell" And RolB(1) = "Bell" And RolC(1) <> "Bell" Then booRotateReel = True
        If RolA(1) = "Blueberry" And RolB(1) = "Blueberry" And RolC(1) <> "Blueberry" Then booRotateReel = True
        If RolA(1) = "Orange" And RolB(1) = "Orange" And RolC(1) <> "Orange" Then booRotateReel = True
    End If
    If booRotateReel = True Then
        intStartPos(2) = intResult(2)
        intCount(2) = 0
        Exit Sub
    End If
    If RolA(1) = "7" Or RolB(1) = "7" Or RolC(1) = "7" Then booBonusLight = True: Exit Sub
    If Int(Rnd * 4) = 0 Then
        booHoldActive = True
        tmrHold.Enabled = True
    Else
        booHoldActive = False
        lblCancel.BackColor = &H808000
    End If
    
End Sub

Private Sub Counter(ByVal Amount As Integer)

    lblWin.Caption = "+" & Str(Amount)
    intWinningAmount = Amount
    tmrAddScore.Enabled = True
   
End Sub

Private Sub tmrAddScore_Timer()

Dim i As Integer

    If intScoreCount < intWinningAmount Then
        lblScore.Caption = Str(Val(lblScore.Caption) + 1)
        sndPlaySound App.Path & "\tik.wav", &H1
        DoEvents
        intScoreCount = intScoreCount + 1
    Else
        tmrAddScore.Enabled = False
        CheckFeatures
        booNoCheck = False
        intCount(0) = 0: intCount(1) = 0: intCount(2) = 0
        For i = 0 To 2
            booHold(i) = False
            lblHold(i).BackColor = &HC0C0&
            lblSelector(i).ForeColor = vbBlack
        Next i
        intFinishedLines = 0
        StartFeatures
        If booBonusLight = True Then
            tmrLights.Enabled = True
            lblMessage.Caption = "Feature - stop at top light."
            sndPlaySound App.Path & "\beep.wav", &H1
        End If
        If booRotateReel = True Then
            lblMessage.Caption = "Feature - stop at correct reel."
            sndPlaySound App.Path & "\beep.wav", &H1
            tmrRotateReel.Enabled = True
        End If
        tmrStart.Enabled = True
    End If
    
End Sub
