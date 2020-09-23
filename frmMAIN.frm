VERSION 5.00
Begin VB.Form frmMAIN 
   BackColor       =   &H00C0C0C0&
   Caption         =   "Physic Engine 2D"
   ClientHeight    =   9705
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   12285
   Icon            =   "frmMAIN.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   647
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   819
   StartUpPosition =   2  'CenterScreen
   WindowState     =   2  'Maximized
   Begin VB.Frame oFrame 
      BackColor       =   &H00C0C0C0&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1935
      Left            =   1320
      TabIndex        =   5
      Top             =   7320
      Width           =   5175
      Begin VB.OptionButton oAuto 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Automatic"
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00008000&
         Height          =   255
         Left            =   3000
         TabIndex        =   10
         Top             =   360
         Value           =   -1  'True
         Width           =   1815
      End
      Begin VB.OptionButton oUseKeyboard 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Use Keyboard"
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00008000&
         Height          =   255
         Left            =   3000
         TabIndex        =   9
         Top             =   960
         Width           =   1815
      End
      Begin VB.CheckBox chkTXTR 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Texture"
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00808000&
         Height          =   375
         Left            =   120
         TabIndex        =   6
         ToolTipText     =   "Perform Cool Texture BackGround"
         Top             =   240
         Value           =   1  'Checked
         Width           =   1695
      End
      Begin VB.Label Label1 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Label1"
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00808000&
         Height          =   375
         Left            =   120
         TabIndex        =   7
         Top             =   840
         Width           =   1695
      End
   End
   Begin VB.Timer Timer2 
      Interval        =   1000
      Left            =   4560
      Top             =   120
   End
   Begin VB.PictureBox Picture2 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   6000
      Left            =   10560
      Picture         =   "frmMAIN.frx":030A
      ScaleHeight     =   6000
      ScaleWidth      =   6000
      TabIndex        =   4
      Top             =   4440
      Visible         =   0   'False
      Width           =   6000
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   7680
      Left            =   10560
      Picture         =   "frmMAIN.frx":41D8
      ScaleHeight     =   7680
      ScaleWidth      =   7680
      TabIndex        =   3
      Top             =   2400
      Visible         =   0   'False
      Width           =   7680
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   10
      Left            =   5880
      Top             =   0
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00808080&
      Caption         =   "Start"
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   120
      TabIndex        =   1
      ToolTipText     =   "Start New Scene"
      Top             =   7440
      Width           =   975
   End
   Begin VB.PictureBox PIC 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00E0E0E0&
      ForeColor       =   &H80000008&
      Height          =   6615
      Left            =   120
      ScaleHeight     =   439
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   687
      TabIndex        =   0
      Top             =   600
      Width           =   10335
   End
   Begin VB.PictureBox BCKGRNDPicture 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00C0C000&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   375
      Left            =   5400
      ScaleHeight     =   375
      ScaleWidth      =   3735
      TabIndex        =   2
      Top             =   7200
      Visible         =   0   'False
      Width           =   3735
   End
   Begin VB.TextBox InfoTxt 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      CausesValidation=   0   'False
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   120
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      TabIndex        =   8
      Top             =   240
      Width           =   8415
   End
   Begin VB.Label Label2 
      BackColor       =   &H00C0C0C0&
      Caption         =   $"frmMAIN.frx":92C1
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2055
      Left            =   6600
      TabIndex        =   11
      Top             =   7560
      Width           =   2535
   End
End
Attribute VB_Name = "frmMAIN"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Author :Roberto Mior
'     reexre@gmail.com
'
'If you use source code or part of it please cite the author
'You can use this code however you like providing the above credits remain intact
'
'
'
'
'--------------------------------------------------------------------------------


Dim CNT As Long
Dim oldCNT As Long



Private Sub Command1_Click()
    Dim I As Long
    Dim x As Single
    Dim y As Single
    
    Dim LX As Long
    Dim lY As Long
    
    Dim lX2 As Long
    Dim lY2 As Long
    
    Dim xFrom As Long
    Dim xTo As Long
    Dim yFrom As Long
    Dim yTo As Long
    
    
    Dim color As Long
    
    Command1.Caption = "Re Start"
    
    Timer1.Enabled = False
    
    CNT = 0
    
    
    
    With BOT
        .Clear
        
        '.AddPoint 50, 50, 25, 1.4 '-
        '.AddPoint 90, 60, 15
        '.AddPoint 60, 90, 15
        '.AddPoint 110, 110, 20, 1.4
        
        '.AddLink 1, 2, 0.3
        '.AddLink 2, 3, 0.3
        '.AddLink 3, 4, 0.3  ' 0.5
        '.AddLink 4, 1, 0.5
        '.AddLink 1, 3, 0.3  '0.5
        '.AddLink 2, 4, 0.3
        '-------------------------------------------Test1
        '        .AddPoint 50, 50, 25, 1.4
        '        .AddPoint 100, 50, 16
        '        .AddPoint 110, 110, 20, 1
        '
        '        'AddPoint 250, 50, 20, 2
        '
        '        .AddLink 1, 2, 0.3
        '        .AddLink 2, 3, 0.4
        '        .AddLink 3, 1, 0.3 ' 0.5
        '
        '        '.AddLink 4, 2, 0.01
        '        .Movable = True
        '--------------------------------------------
        
        
        '-------------------------------------------Test2
        '        .AddPoint 100, 200, 18, 1
        '        .AddPoint 150, 200, 12, 1
        '        .AddPoint 200, 200, 12, 1
        '
        '        .AddPoint 100, 150, 12, 1
        '        .AddPoint 150, 150, 18, 1
        '        .AddPoint 200, 150, 18, 1
        '
        '        '--
        '        '.AddPoint 250, 100, 12, 1
        '        '.AddPoint 250, 50, 18, 1
        '
        '        .AddLink 1, 2, 0.15
        '        .AddLink 2, 3, 0.15
        '
        '        .AddLink 4, 5, 0.15
        '        .AddLink 5, 6, 0.15
        '
        '        .AddLink 1, 4, 0.15
        '        .AddLink 2, 5, 0.15
        '        .AddLink 3, 6, 0.15
        '
        '        .AddLink 1, 5, 0.2
        '        .AddLink 2, 6, 0.2
        '        .AddLink 2, 4, 0.2
        '        .AddLink 3, 5, 0.2
        '
        '        '--
        '        '.AddLink 3, 7, 0.15
        '        '.AddLink 7, 8, 0.15
        '        '.AddLink 8, 6, 0.15
        '        '.AddLink 3, 8, 0.2
        '        '.AddLink 6, 7, 0.2
        '
        '        .Movable = True
        ''-------------------------------------------
        
        
        '        '-------------------------------------------Test3
        '        '         100,100
        '        .AddPoint 100, 100 ', 24 , 1
        '        '              60
        '        .AddPoint 130, 60, 8, 1
        '        .AddPoint 220, 60, 8, 1
        '        '         250,100
        '        .AddPoint 250, 100, 24, 1
        '
        '        .AddPoint 125 + 50, 15 ', 8
        '
        '       .AddLink 1, 2, 0.5, 0, 0.5, 1
        '
        '        .AddLink 2, 3, 0.5
        '        .AddLink 3, 4, 0.5
        '
        '        .AddLink 2, 5, 0.5
        '        .AddLink 3, 5, 0.5, 2, 0.5, 1
        '
        '
        '        .ADDMuscle 1, 2, 0.05
        '        .ADDMuscle 3, 2, 0.15 '.15 '.15 '0.1
        '
        '
        '        .Movable = True
        '        '-------------------------------------------
        
        
        '-------------------------------------------Test3
        .AddPoint 100, 100, 24, 1
        .AddPoint 130, 60
        .AddPoint 220, 60
        .AddPoint 280, 70, 20, 1
        
        .AddPoint 125 + 50, 25, 10, 1
        
        .AddLink 1, 2, 0.5, 180, 50, 0.5
        .AddLink 2, 3, 1
        .AddLink 3, 4, 0.5
        
        .AddLink 2, 5, 1
        .AddLink 3, 5, 1
        
        
        .ADDMuscle 1, 2, 0.1
        
        .ADDMuscle 3, 2, 0.1, 180, 45, 0.5
        
        .Movable = True
        
        '        '-------------------------------------------
    End With
    
    With Ground
        .Clear
        
        y = 440
        
        .AddPoint 0, -200
        .AddPoint 2, y - 120
        
        k1 = 200 + Rnd * 100
        k2 = 150 + Rnd * 80
        k3 = 80 + Rnd * 50
        kx1 = Rnd * 10000
        kx2 = Rnd * 10000
        kx3 = Rnd * 10000
        
        For x = 50 To PIC.Width * 5 Step 50
            
            y = -150 + Sin((x + kx1) / k1) * 80 + 440
            y = y + Sin((x + kx2) / k2) * 50
            y = y + Sin((x + kx3) / k3) * 40
            If Rnd < 0.1 Then y = y - 30 + Rnd * 60
            If y > PIC.Height - 10 Then y = PIC.Height - 10
            
            .AddPoint x, y
            
        Next
        .AddPoint x - 50 + 1, -200
        
        For I = 1 To Ground.NP - 1
            .AddLink I, I + 1, -999
        Next I
        
    End With
    
    
    '-------------------------------------------------------------------------
    '-------------------------------------------------------------------------
    
    Me.Caption = "Building Texture... Please wait..."
    Me.Refresh
    
    
    BCKGRNDPicture.Width = PIC.Width * 5
    BCKGRNDPicture.Height = PIC.Height
    BCKGRNDPicture.Cls
    
    If chkTXTR.Value = Checked Then
        '        For x = 0 To BCKGRNDPicture.Width
        '            For y = 0 To BCKGRNDPicture.Height
        '                color = GetPixel(Picture1.hdc, x Mod Picture1.Width, y Mod Picture1.Height)
        '                SetPixel BCKGRNDPicture.hdc, x, y, color
        '            Next
        '        Next
        For x = 0 To BCKGRNDPicture.Width Step Picture1.Width
            For y = 0 To BCKGRNDPicture.Height Step Picture1.Height
                BitBlt BCKGRNDPicture.hdc, x, y, Picture1.Width, Picture1.Height, Picture1.hdc, 0, 0, vbSrcCopy
            Next
        Next
        
    End If
    
    
    For I = 1 To Ground.NP - 1
        
        
        xFrom = Ground.GetPointX(I)
        xTo = Ground.GetPointX(I + 1) + 1
        yFrom = Ground.GetPointY(I)
        yTo = Ground.GetPointY(I + 1)
        
        If chkTXTR.Value = Checked Then
            
            For x = xFrom To xTo
                y = (xTo - x) / (xTo - xFrom)
                y = yFrom * (y) + yTo * (1 - y)
                
                '        FastLine BCKGRNDPicture.hdc, X \ 1, 440, X \ 1, Y \ 1, 1, vbGreen
                
                For y3 = y To 441
                    
                    If y3 < 0 Then y3 = 0
                    
                    Y2 = (y3 - y) / (441 - y)
                    Y2 = Y2 * 256
                    Y2 = 256 - Y2
                    If y3 <= 441 Then
                        'color = RGB(0, Y2, 0)
                        'FastLine BCKGRNDPicture.hdc, X \ 1, Y3 \ 1, X \ 1, Y3 \ 1 + 1, 1, color
                        
                        color = GetPixel(Picture2.hdc, x Mod Picture2.Width, y3 Mod Picture2.Height)
                        color = color - Y2 / 2.7
                        SetPixel BCKGRNDPicture.hdc, x, y3, color
                        
                    End If
                Next y3
            Next x
            
        End If
        
        FastLine BCKGRNDPicture.hdc, xFrom, yFrom, xTo, yTo, 3, vbBlack
        
        Me.Caption = "Building Texture... Please wait..." & I
        Me.Refresh
        
        
    Next I
    
    
    Timer1.Enabled = True
    Me.Caption = "RUN!"
End Sub

Private Sub Form_Activate()
    Command1_Click
End Sub

Private Sub Form_Load()
    Randomize Timer
    
    Gravity = 0.12 ' 0.075 '0.05
    Atmosphere = 0.002 '0.001
    WallBounce = 0.99 '1.2
    WallFriction = 0.35 ' 0.8 '0.95
    ' WIND = 0.02
    
    
    RightWall = PIC.Width
    FloorWall = PIC.Height
    
    
    PIC.Width = 823 ' Screen.Width \ Screen.TwipsPerPixelX - 200
    PIC.Left = (Screen.Width \ Screen.TwipsPerPixelX - PIC.Width) \ 2
    
    
    Command1.Left = PIC.Left + PIC.Width - Command1.Width
    
    oFrame.Left = (Screen.Width \ Screen.TwipsPerPixelX - oFrame.Width) \ 2
    
    InfoTxt.Text = App.Title & " V" & App.Major & "." & App.Minor & " by reexre@gmail.com"
    InfoTxt.Left = (Screen.Width \ Screen.TwipsPerPixelX - InfoTxt.Width) \ 2
    
    Label2.Left = PIC.Left
    Label2.Width = oFrame.Left - PIC.Left - 10
    Label2.Top = oFrame.Top
    Label2.Height = oFrame.Height
    
    
    
End Sub

Private Sub Form_Unload(Cancel As Integer)
End

End Sub

Private Sub oUseKeyboard_Click()
    Command1.SetFocus
    
End Sub

Private Sub Timer1_Timer()
    'Me.Caption = BOT.SpeedX
    
    PanX = PIC.Width * 0.5 - BOT.comX
    
    BOT.DoPhysics
    
    BitBlt PIC.hdc, 0, 0, PIC.Width, 441, BCKGRNDPicture.hdc, -PanX, 0, vbSrcCopy
    BOT.DRAW PIC.hdc, PanX
    ' Ground.DRAW PIC.hdc, PanX
    Part.DRAW_and_Physics PIC.hdc, PanX
    
    PIC.Refresh
    DoEvents
    
    
    chkCircleLine BOT, Ground
    chkLineLine2 BOT, Ground
    BOT.DoMuscles
    
    
    
    
    
    If oUseKeyboard Then KeyBoard
    
    CNT = CNT + 1
    If CNT > 5000 Then CNT = 0:  Command1_Click
End Sub


Sub chkCircleLine(ByRef tBOT As clsBOT, tGround As clsBOT)
    
    Dim TouchPoint As tPoint
    Dim Reflect As tPoint
    
    Dim P1 As tPoint
    Dim P2 As tPoint
    Dim C As tPoint 'Circle center
    Dim R As Single 'Circle Radius
    Dim WallNormalized As tPoint
    
    Dim jFrom As Long
    Dim jTo As Long
    Dim jStep As Long
    
    Dim PB1 As tPoint
    Dim PB2 As tPoint
    
    Dim j As Long
    
    Dim I As Long
    Dim I1 As Long
    Dim I2 As Long
    Dim iNear As Long
    
    
    'Start Scanning Bot Points
    For I = 1 To tBOT.NP
        
        C = MakePoint(tBOT.GetPointX(I), tBOT.GetPointY(I))
        R = tBOT.GetPointR(I)
        'C  = Center of tBOT Point (i)
        'R  = Radius of tBOT Point (i)
        
        'For some reason works better if "For J cicle" dirction
        'follows the MomentumX direction
        If tBOT.GetPointMomX(I) >= 0 Then
            jFrom = 1
            jTo = tGround.NL
            jStep = 1
        Else
            jFrom = tGround.NL
            jTo = 1
            jStep = -1
        End If
        
        
        'Start Scanning Ground Points
        For j = jFrom To jTo Step jStep
            
            P1 = MakePoint(tGround.GetPointX(tGround.GetLinkP1(j)), tGround.GetPointY(tGround.GetLinkP1(j)))
            P2 = MakePoint(tGround.GetPointX(tGround.GetLinkP2(j)), tGround.GetPointY(tGround.GetLinkP2(j)))
            
            'P1 and P2 are the points of (tGround) Link (j)
            'Maybe a better way to Get Them
            
            
            'MaxX is Bot maximum X plus 100
            'MinX is Bot minimun X minus 100
            If P1.x > tBOT.MaxX Or P2.x > tBOT.MaxX Or _
                    P1.x < tBOT.MinX Or P2.x < tBOT.MinX Then
            
            'Point on Ground is too far from BOT
            'to perform Circle Line Collision
            
            Else
            
                TouchPoint = CIRCLE_LINE(P1, P2, C, R)
            
                If TouchPoint.x > 0 And TouchPoint.y > 0 Then
                    
                    '   PIC.Circle (Touchpoint.x + PanX, Touchpoint.y), 5, vbYellow ': Stop
                
                    P1.x = P2.x - P1.x
                    P1.y = P2.y - P1.y
                    WallNormalized = Normalize(P1)
                    '    PIC.Line (Ret.x + PanX, Ret.y - 2)-(Ret.x + PanX + WallNormalized.x * tBOT.GetMotor(i) * 20, Ret.y + WallNormalized.y * tBOT.GetMotor(i) * 20 - 2), vbYellow ': Stop
                
                    Reflect = getREFLECT(MakePoint(tBOT.GetPointMomX(I), tBOT.GetPointMomY(I)), P1)
                
                    Reflect.x = Reflect.x * WallBounce
                    Reflect.y = Reflect.y * WallBounce
                
                    'New Momentum XY
                    tBOT.SetPointMomX(I) = -Reflect.x
                    tBOT.SetPointMomY(I) = -Reflect.y
                
                    'Set Point New Position equal to old Position plus
                    '(Reflect) Bounced Velocity
                    tBOT.SetPointPosEqualOldPos (I)
                    tBOT.SetPointX(I) = tBOT.GetPointX(I) - Reflect.x '* 2
                    tBOT.SetPointY(I) = tBOT.GetPointY(I) - Reflect.y '* 2
                                
                    ' if a Point Touch the First or the Last "tGround" segment
                    ' then Invert motors Forces.
                    If j = 1 Or j = Ground.NL Then
                        For I2 = 1 To tBOT.NP
                            tBOT.SetMotor(I2) = -tBOT.GetMotor(I2)
                        Next
                    End If
                
                    tBOT.ApplyMotorForce I, WallNormalized.x, WallNormalized.y
                
                    'If Motor Add a Particle
                    If Abs(tBOT.GetMotor(I) <> 0) Then
                        Part.AddParticle TouchPoint.x, TouchPoint.y, -WallNormalized.x * tBOT.GetMotor(I), -WallNormalized.y * tBOT.GetMotor(I)
                    End If
                
                    'Exit J For Cicle
                    j = jTo
                
                End If
            
            End If
        
        Next j
    
    Next I

End Sub

Sub chkLineLine2(ByRef tBOT As clsBOT, tGround As clsBOT)
    
    Dim TouchPoint As tPoint
    Dim Reflect As tPoint
    
    Dim P1 As tPoint
    Dim P2 As tPoint

    Dim WallNormalized As tPoint
    
    Dim jFrom As Long
    Dim jTo As Long
    Dim jStep As Long
    
    Dim PB1 As tPoint
    Dim PB2 As tPoint
    
    Dim j As Long
    
    Dim I As Long
    
    Dim ProjWall As New cls2DVector
    Dim ProjPerpWall As New cls2DVector
    
    '-------------------------Line Line
    For I = 1 To tBOT.NP
        
                        
        If BOT.IsWheel(I) = False Then
            
            PB1 = MakePoint(tBOT.GetPointX(I), tBOT.GetPointY(I))
        
            PB2 = MakePoint(tBOT.GetPointOldX(I), tBOT.GetPointOldY(I))
        
            'For some reason works better if "For J cicle" dirction
            'follows the MomentumX direction
            If tBOT.GetPointMomX(I) > 0 Then
                jFrom = 1
                jTo = tGround.NL
                jStep = 1
            Else
                jFrom = tGround.NL
                jTo = 1
                jStep = -1
            End If
            
            
            'Start Scanning Ground Points
            For j = jFrom To jTo Step jStep
                
                P1 = MakePoint(tGround.GetPointX(tGround.GetLinkP1(j)), tGround.GetPointY(tGround.GetLinkP1(j)))
                P2 = MakePoint(tGround.GetPointX(tGround.GetLinkP2(j)), tGround.GetPointY(tGround.GetLinkP2(j)))
                
                'P1 and P2 are the points of (tGround) Link (j)
                'Maybe a better way to Get Them
                
                
                'MaxX is Bot maximum X plus 100
                'MinX is Bot minimun X minus 100
                If P1.x > tBOT.MaxX Or P2.x > tBOT.MaxX Or _
                        P1.x < tBOT.MinX Or P2.x < tBOT.MinX Then
                
                    'Point on Ground is too far from BOT
                    'to perform Circle Line Collision
                
                Else
                
                
                    TouchPoint = LINE_LINE(P1, P2, PB1, PB2)
                
                    If TouchPoint.x <> 0 Or TouchPoint.y <> 0 Then
                    
                        '   PIC.Circle (Touchpoint.x + PanX, Touchpoint.y), 5, vbYellow ': Stop
                    
                        P1.x = P2.x - P1.x
                        P1.y = P2.y - P1.y
                        WallNormalized = Normalize(P1)
                        '    PIC.Line (Ret.x + PanX, Ret.y - 2)-(Ret.x + PanX + WallNormalized.x * tBOT.GetMotor(i) * 20, Ret.y + WallNormalized.y * tBOT.GetMotor(i) * 20 - 2), vbYellow ': Stop
                    
                        Reflect = getREFLECT(MakePoint(tBOT.GetPointMomX(I), tBOT.GetPointMomY(I)), P1)
                    
                    '-------------
                        'Vector ProjWall contains the Projection of Vector "reflect" to the Wall
                        Set ProjWall = MakeVector(Reflect.x, Reflect.y).Projection(MakeVector(P1.x, P1.y))
                        
                        'Vector ProjWall contains the Projection of Vector "reflect" to the Perpendicular of Wall
                        Set ProjPerpWall = MakeVector(Reflect.x, Reflect.y).Projection(MakeVector(P1.x, P1.y).Perpendicular)
                    
                        '-----
                        'The Reaction on Parallel to Wall is affected by Wallfriction
                        'The Reaction on Perppendicular to wall is affected by WallBounce
                        ProjWall.MUL 1 - WallFriction
                        ProjPerpWall.MUL WallBounce
                    
                        Reflect.x = ProjWall.x + ProjPerpWall.x
                        Reflect.y = ProjWall.y + ProjPerpWall.y
                        '------
                        
                    '--------------
                    
                        tBOT.SetPointMomX(I) = -Reflect.x
                        tBOT.SetPointMomY(I) = -Reflect.y
                    
                        tBOT.SetPointPosEqualOldPos (I)
                        tBOT.SetPointX(I) = tBOT.GetPointX(I) - Reflect.x '* 2
                        tBOT.SetPointY(I) = tBOT.GetPointY(I) - Reflect.y '* 2
                    
                        ' tBOT.SetPointX(I) = TouchPoint.x - Reflect.x * 2
                        ' tBOT.SetPointY(I) = TouchPoint.y - Reflect.y * 2
                    
                    
                        j = jTo
                    
                    End If
                
                End If
            
            Next j
        
        End If '(isWheel)
    
    Next I

End Sub

Private Sub Timer2_Timer()
    Label1 = CNT - oldCNT & " FPS"
    oldCNT = CNT
End Sub
