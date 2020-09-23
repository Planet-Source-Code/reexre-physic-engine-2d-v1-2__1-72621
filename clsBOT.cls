VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsBOT"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
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


Public NP As Long 'Number of Points
Private P() As tPoint 'Points Array
Public NL As Long 'Number of Links
Private L() As tLink 'Links Array
Public NMuscles As Long 'Number of Muscles
Private Muscle() As tMuscle 'Muscles Array

Public IntTimer As Long 'Internal Timer

Public Movable As Boolean 'Movable or Stand Still?

Public comX As Single 'Bot Center X
Public comY As Single 'Bot Center Y

Public SpeedX As Single 'Bot V X
Public SpeedY As Single 'Bot V Y

Public MinX As Single 'Here will be put Bot min X -100. Used in during Check Collision
Public MaxX As Single 'Here will be put Bot min X +100

Private Const SpokColor As Long = 13132900 'RGB(100, 100, 200)

Public Function GetPointX(I As Long) As Single
    GetPointX = P(I).x
End Function
Public Function GetPointY(I As Long) As Single
    GetPointY = P(I).y
End Function
Public Function GetPointOldX(I As Long) As Single
    GetPointOldX = P(I).OldX
End Function
Public Function GetPointOldY(I As Long) As Single
    GetPointOldY = P(I).OldY
End Function
Public Function GetMotor(I As Long) As Single
    GetMotor = P(I).Motor
End Function
Public Property Let SetMotor(I As Long, V As Single)
P(I).Motor = V
End Property
Public Property Get IsMotor(I As Long) As Boolean
IsMotor = P(I).IsMotor
End Property
Public Property Get IsWheel(I As Long) As Boolean
IsWheel = P(I).IsWheel
End Property



Public Property Let SetPointX(I As Long, V As Single)
P(I).x = V
End Property
Public Property Let SetPointY(I As Long, V As Single)
P(I).y = V
End Property
Public Function GetPointR(I As Long) As Single
    GetPointR = P(I).Radius
End Function

Public Function GetPointMomX(I As Long) As Single
    GetPointMomX = P(I).momentum_X
End Function
Public Function GetPointMomY(I As Long) As Single
    GetPointMomY = P(I).momentum_Y
End Function
Public Property Let SetPointMomX(I As Long, V As Single)
P(I).momentum_X = V
End Property
Public Property Let SetPointMomY(I As Long, V As Single)
P(I).momentum_Y = V
End Property
Public Function GetLinkP1(I As Long) As Long
    GetLinkP1 = L(I).P1
End Function
Public Function GetLinkP2(I As Long) As Long
    GetLinkP2 = L(I).P2
End Function


Public Function SetPointPosEqualOldPos(I As Long)
    P(I).x = P(I).OldX
    P(I).y = P(I).OldY
End Function



Public Sub Clear()
    Dim I As Long
    For I = 1 To NP
        P(I).momentum_C = 0
        P(I).momentum_X = 0
        P(I).momentum_Y = 0
    Next
    
    IntTimer = 0
    
    NP = 0
    NL = 0
    NMuscles = 0
End Sub

Public Sub AddPoint(x, y, Optional R = 0, Optional Motor = 0)
    NP = NP + 1
    ReDim Preserve P(NP)
    
    With P(NP)
        .x = x
        .OldX = x
        .y = y
        .OldY = y
        .Radius = R
        .Heading = 0
        If R > 0 Then
            .IsWheel = True
            .Motor = Motor
            If Motor <> 0 Then
                .IsMotor = True
            End If
            
        End If
        
    End With
    
End Sub


Public Sub AddLink(P1 As Long, P2 As Long, TENS, _
            Optional dPhase360 As Single = 0, Optional dAmp100 As Single = 50, Optional dSpeed As Single = 0)
    'TENS is the tension between the Points (or Link Stregth). MAXvalue=1
    'dPhase360 is the Dynamic Phase in Degrees
    'dAmp100 is the Dynamic (length) Amplitude: 100 Means then Maximum Length will be 2Times the MainLength and The Minimun Length will be 0
    'dSpeed is the Dynamic Speed
    
    If P1 = P2 Then Exit Sub
    
    NL = NL + 1
    ReDim Preserve L(NL)
    With L(NL)
        .P1 = P1
        .P2 = P2
        .midX = (P(P1).x + P(P2).x) * 0.5
        .midY = (P(P1).y + P(P2).y) * 0.5
        .CurrLen = PointDist(P(P1), P(P2))
        .LastLen = .CurrLen
        .TENS = TENS
        
        If dSpeed <> 0 Then
            .IsDynamic = True
            .DynPhase = dPhase360
            .DynAmp = .CurrLen * dAmp100 / 100
            .DynSpeed = dSpeed
        End If
        
        
    End With
    
End Sub


Public Function ADDMuscle(L1, L2, Strength, _
            Optional dPhase360 As Single = 0, Optional dAmp360 As Single = 20, Optional dSpeed As Single = 0) As Boolean
    'dPhase360 is the Dynamic Phase in Degrees
    'dAmp360 is the Dynamic Angle Amplitude in Degrees (from -Value to +Value)
    'dSpeed is the Dynamic Speed
    
    NMuscles = NMuscles + 1
    ReDim Preserve Muscle(NMuscles)
    
    ADDMuscle = False
    
    With Muscle(NMuscles)
        .L1 = L1
        .L2 = L2
        .f = Strength
        .isNotBroken = True
        
        If L(L1).P1 = L(L2).P1 Then
            .P0 = L(L1).P1
            .P1 = L(L1).P2
            .P2 = L(L2).P2
            ADDMuscle = True
        End If
        
        If L(L1).P1 = L(L2).P2 Then
            .P0 = L(L1).P1
            .P1 = L(L1).P2
            .P2 = L(L2).P1
            ADDMuscle = True
        End If
        
        If L(L1).P2 = L(L2).P1 Then
            .P0 = L(L1).P2
            .P1 = L(L1).P1
            .P2 = L(L2).P2
            ADDMuscle = True
        End If
        
        If L(L1).P2 = L(L2).P2 Then
            .P0 = L(L1).P2
            .P1 = L(L1).P1
            .P2 = L(L2).P1
            ADDMuscle = True
        End If
        'Stop
        
    End With
    
    If ADDMuscle = False Then
        NMuscles = NMuscles - 1
        ReDim Preserve Muscle(NMuscles)
        Exit Function
    End If
    
    
    
    With Muscle(NMuscles)
        .MainA = Atan2(P(.P1).x - P(.P0).x, P(.P1).y - P(.P0).y) - _
                Atan2(P(.P2).x - P(.P0).x, P(.P2).y - P(.P0).y)
        
        
        If IsFixed Then .FixedANG = True
        
        If dSpeed <> 0 Then
            .IsDynamic = True
            .DynPhase = dPhase360
            .DynAmp = dAmp360
            .DynSpeed = dSpeed
        End If
        
    End With
    
End Function


Public Sub DRAW(ByRef PIChdc As Long, PanX As Single)
    
    Dim I As Long
    Dim j As Long
    
    Dim X1 As Long
    Dim Y1 As Long
    Dim X2 As Long
    Dim Y2 As Long
    
    Dim NSpokes As Long
    
    Dim subHeading As Single
    
    Dim pArc0 As tPoint
    Dim pArc1 As tPoint
    Dim pArc2 As tPoint
    
    
    For I = 1 To NL
        
        X1 = P(L(I).P1).x + PanX
        Y1 = P(L(I).P1).y
        X2 = P(L(I).P2).x + PanX
        Y2 = P(L(I).P2).y
        
        If L(I).IsDynamic Then
            
            FastLine PIChdc, X1, Y1, X2, Y2, 2, vbBlue
        Else
            FastLine PIChdc, X1, Y1, X2, Y2, 2, vbBlack
        End If
        
    Next I
    
    For I = 1 To NP
        If P(I).IsWheel Then
            
            X1 = P(I).x + PanX
            Y1 = P(I).y
            
            NSpokes = IIf(P(I).IsMotor, 3, 1)
            
            For j = 1 To NSpokes
                subHeading = P(I).Heading + ((360 / NSpokes) * j)
                X2 = X1 + P(I).Radius * Cos(subHeading * 0.0174) 'PI/180
                Y2 = Y1 + P(I).Radius * Sin(subHeading * 0.0174)
                
                FastLine PIChdc, X1, Y1, X2, Y2, 2, SpokColor
                
            Next
            
            MyCircle PIChdc, X1, Y1, P(I).Radius \ 1, 3, vbBlack
            
        End If
        
        
    Next
    
    'Draw Muscle ARC
    For I = 1 To NMuscles
        
        P0 = Muscle(I).P0
        P1 = Muscle(I).P1
        P2 = Muscle(I).P2
        
        pArc0 = MakePoint(P(P0).x, P(P0).y)
        pArc1 = Normalize(MakePoint(P(P1).x - P(P0).x, P(P1).y - P(P0).y))
        pArc2 = Normalize(MakePoint(P(P2).x - P(P0).x, P(P2).y - P(P0).y))
        
        pArc0.x = pArc0.x + PanX
        pArc1.x = pArc1.x * 15 + pArc0.x
        pArc1.y = pArc1.y * 15 + pArc0.y
        pArc2.x = pArc2.x * 15 + pArc0.x
        pArc2.y = pArc2.y * 15 + pArc0.y
        
        If Muscle(I).IsDynamic Then
            DeleteObject (SelectObject(PIChdc, CreatePen(vbSolid, 1, vbBlue)))
        Else
            DeleteObject (SelectObject(PIChdc, CreatePen(vbSolid, 1, vbBlack)))
        End If
        
        Arc PIChdc, pArc0.x - 15, pArc0.y - 15, pArc0.x + 15, pArc0.y + 15, _
                pArc1.x, pArc1.y, pArc2.x, pArc2.y
        
    Next I
    
    
End Sub

Public Sub DoPhysics()
    
    
    Dim xer  As Single
    Dim yer  As Single
    Dim NewX  As Single
    Dim NewY As Single
    Dim Leng As Double 'Single
    Dim Leng2Go_x As Single
    Dim Leng2Go_y As Single
    Dim LengthTotal As Single
    Dim TempTime As Single
    Dim Fric
    
    Dim I As Long
    Dim P1 As Long
    Dim P2 As Long
    
    
    
    If Not (Movable) Then Exit Sub
    
    IntTimer = IntTimer + 1
    
    
    For I = 1 To NL
        
        With L(I)
            
            
            P1 = .P1
            P2 = .P2
            
            'This will calculate the links 'true' length.  That is the distance
            'between its 2 vertices.   The length stored in .linklength is
            'the length that the link 'should' be.  The link will push or pull
            'to bring the vertices back to that distance.
            LengthTotal = .CurrLen
            xer = (P(P2).x + P(P2).momentum_X) - (P(P1).x + P(P1).momentum_X)
            yer = (P(P2).y + P(P2).momentum_Y) - (P(P1).y + P(P1).momentum_Y)
            Leng = Sqr(xer * xer + yer * yer)
            
            
            If .IsDynamic Then
                'pi/180
                Leng = Leng + Sin((IntTimer * .DynSpeed + .DynPhase) * 0.017453292) * .DynAmp
                
            End If
            
            
            
            
            'Calculate how far in each direction the vertices must go
            'in order to get the link back to its regular length
            Leng2Go_x = ((Leng - LengthTotal) / Leng) * xer
            Leng2Go_y = ((Leng - LengthTotal) / Leng) * yer
            
            
            
            P(P1).momentum_X = P(P1).momentum_X + (Leng2Go_x / 2) * .TENS
            P(P1).momentum_Y = P(P1).momentum_Y + (Leng2Go_y / 2) * .TENS
            P(P2).momentum_X = P(P2).momentum_X + (Leng2Go_x / 2) * -1 * .TENS
            P(P2).momentum_Y = P(P2).momentum_Y + (Leng2Go_y / 2) * -1 * .TENS
            'These lines actually add the neccessary momentum to the Link's
            'two vertices to make them snap into place.
            'It factors in the amount of tension the link has
            'If the link has a .linktension value of 1, it will snap back
            'into place almost instantly.
            'If it has a value of 0, it will not try to re-establish its
            'length.
            
        End With
        
    Next
    
    MinX = 9999999999#
    MaxX = -9999999999#
    
    comX = 0
    comY = 0
    SpeedX = 0
    SpeedY = 0
    
    For I = 1 To NP
        
        With P(I)
            
            If .x > MaxX Then MaxX = .x
            If .x < MinX Then MinX = .x
            
            comX = comX + .x
            comY = comY + .y
            
            SpeedX = SpeedX + .momentum_X
            SpeedY = SpeedY + .momentum_Y
            
            .momentum_Y = .momentum_Y + Gravity
            .momentum_X = .momentum_X + WIND
            
            .momentum_X = .momentum_X * (1 - Atmosphere)
            .momentum_Y = .momentum_Y * (1 - Atmosphere)
            
            .OldX = .x
            .OldY = .y
            
            .x = .x + .momentum_X
            .y = .y + .momentum_Y
            
            Fric = WallFriction
            If .IsWheel = True Then Fric = 0
            
            If .y + .Radius > FloorWall Then 'floor
                '.Y = FloorWall - .Radius
                .y = .OldY - .momentum_Y * 2
                .momentum_X = .momentum_X * (1 - Fric)
                .momentum_Y = -(.momentum_Y * WallBounce)
                If .IsWheel Then
                    If .Motor <> 0 Then .momentum_X = .momentum_X + .Motor * .Radius * 0.05
                    .momentum_C = 1 * .momentum_X
                End If
            End If
            
            
            '            If .X - .Radius < 0.1 Then 'left wall
            '                WIND = -WIND
            '                '.X = 0 + .Radius
            '                .X = .OldX - .momentum_X * 2
            '                .momentum_X = -(.momentum_X * WallBounce)
            '                .momentum_Y = .momentum_Y * (1 - Fric)
            '                'If mnuAutoRev.Checked Then
            '                '    If AutoReverseCycle = 0 Then AutoReverseCycle = 2: ClockSpeed = ClockSpeed * -1: INDspd = "Spd: " & ClockSpeed
            '                '    If AutoReverseCycle = 1 Then AutoReverseCycle = 2: ClockSpeed = ClockSpeed * -1: INDspd = "Spd: " & ClockSpeed
            '                'End If
            '                If .Wheel Then
            '                    .Motor = -.Motor
            '                    If .Motor <> 0 Then .momentum_Y = .momentum_Y + .Motor
            '                    .momentum_C = .momentum_Y
            '
            '                End If
            '            End If
            '            If .X + .Radius > (RightWall - 0.1) Then 'right wall
            '                WIND = -WIND
            '                '.X = RightWall - .Radius
            '                .X = .OldX - .momentum_X * 2
            '                .momentum_X = (.momentum_X * WallBounce) * -1
            '                .momentum_Y = .momentum_Y * (1 - Fric)
            '                'If mnuAutoRev.Checked Then
            '                '    If AutoReverseCycle = 0 Then AutoReverseCycle = 1: ClockSpeed = ClockSpeed * -1: INDspd = "Spd: " & ClockSpeed
            '                '    If AutoReverseCycle = 2 Then AutoReverseCycle = 1: ClockSpeed = ClockSpeed * -1: INDspd = "Spd: " & ClockSpeed
            '                'End If
            '                If .Wheel Then
            '                    .Motor = -.Motor
            '                    If .Motor <> 0 Then .momentum_Y = .momentum_Y - .Motor
            '                    .momentum_C = -.momentum_Y
            '                End If
            '            End If
            
            
            .momentum_C = .momentum_C * (1 - Atmosphere)
            .Heading = .Heading + .momentum_C
            If .Heading > 360 Then .Heading = .Heading - 360
            If .Heading < 0 Then .Heading = .Heading + 360
            
            'momentum_c is clockwise momentum.  If the wheel has momentum
            'it will turn
            
        End With
    Next
    
    comX = comX \ NP
    comY = comY \ NP
    
    SpeedX = SpeedX / NP
    SpeedY = SpeedY / NP
    
    MinX = MinX - 100
    MaxX = MaxX + 100
    
End Sub

Sub DoMuscles()
    Dim X1 As Single
    Dim Y1 As Single
    Dim X2 As Single
    Dim Y2 As Single
    Dim AM As Single
    Dim Dx1 As Single
    Dim Dy1 As Single
    Dim Dx2 As Single
    Dim Dy2 As Single
    
    Dim RetP As tPoint
    
    Dim Perc1 As Single
    Dim Perc2 As Single
    
    'Const Kfixed = 0.1
    Const kMuscleSpeedLimit = 5
    
    For M = 1 To NMuscles
        With Muscle(M)
            P(.P0).NewX = P(.P0).x + P(.P0).momentum_X
            P(.P0).NewY = P(.P0).y + P(.P0).momentum_Y
            P(.P1).NewX = P(.P1).x + P(.P1).momentum_X
            P(.P1).NewY = P(.P1).y + P(.P1).momentum_Y
            P(.P2).NewX = P(.P2).x + P(.P2).momentum_X
            P(.P2).NewY = P(.P2).y + P(.P2).momentum_Y
        End With
    Next
    
    For M = 1 To NMuscles
        
        With Muscle(M)
            
            a1 = Atan2(P(.P1).NewX - P(.P0).NewX, P(.P1).NewY - P(.P0).NewY)
            a2 = Atan2(P(.P2).NewX - P(.P0).NewX, P(.P2).NewY - P(.P0).NewY)
            
            AM = (a1 - a2) - .MainA
            
            If .IsDynamic Then                                  'pi/180
                AM = AM + Sin((IntTimer * .DynSpeed + .DynPhase) * 0.017453292) * .DynAmp * 0.017453292
            End If
            
            
            '--------------------------------------------------
            'This Should Make Longer Link Move Slower
            'and shorter move faster
            X1 = P(.P0).NewX + Cos(a1 - AM) * L(.L1).CurrLen
            Y1 = P(.P0).NewY + Sin(a1 - AM) * L(.L1).CurrLen
            X2 = P(.P0).NewX + Cos(a2 + AM) * L(.L2).CurrLen
            Y2 = P(.P0).NewY + Sin(a2 + AM) * L(.L2).CurrLen
            
            If L(.L1).CurrLen > L(.L2).CurrLen Then
                Perc1 = L(.L2).CurrLen / L(.L1).CurrLen '(L(.L1).CurrLen + L(.L2).CurrLen)
                Perc2 = 1 - Perc1
            Else
                Perc2 = L(.L1).CurrLen / L(.L2).CurrLen '(L(.L1).CurrLen + L(.L2).CurrLen)
                Perc1 = 1 - Perc2
            End If
            '--------------------------------------------------
            
            
            Dx1 = (P(.P1).NewX - X1) * Perc1
            Dy1 = (P(.P1).NewY - Y1) * Perc1
            Dx2 = (P(.P2).NewX - X2) * Perc2
            Dy2 = (P(.P2).NewY - Y2) * Perc2
            
            'Apply OK Speed Limit
            'D = Sqr(Dx1 * Dx1 + Dy1 * Dy1)
            'If D > kMuscleSpeedLimit Then
            '    RetP = Normalize(MakePoint(Dx1, Dy1))
            '    Dx1 = RetP.x * kMuscleSpeedLimit
            '    Dy1 = RetP.y * kMuscleSpeedLimit
            'End If
            'D = Sqr(Dx2 * Dx2 + Dy2 * Dy2)
            'If D > kMuscleSpeedLimit Then
            '    RetP = Normalize(MakePoint(Dx2, Dy2))
            '    Dx2 = RetP.x * kMuscleSpeedLimit
            '    Dy2 = RetP.y * kMuscleSpeedLimit
            'End If
            '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
            
            '            If .FixedANG Then
            '                P(.P1).momentum_X = P(.P1).momentum_X - Kfixed * Dx1
            '                P(.P1).momentum_Y = P(.P1).momentum_Y - Kfixed * Dy1
            '                P(.P2).momentum_X = P(.P2).momentum_X - Kfixed * Dx2
            '                P(.P2).momentum_Y = P(.P2).momentum_Y - Kfixed * Dy2
            '                P(.P0).momentum_X = P(.P0).momentum_X + Kfixed * (Dx1 + Dx2)
            '                P(.P0).momentum_Y = P(.P0).momentum_Y + Kfixed * (Dy1 + Dy2)
            '            Else
            
            P(.P1).momentum_X = P(.P1).momentum_X - .f * Dx1
            P(.P1).momentum_Y = P(.P1).momentum_Y - .f * Dy1
            
            P(.P2).momentum_X = P(.P2).momentum_X - .f * Dx2
            P(.P2).momentum_Y = P(.P2).momentum_Y - .f * Dy2
            P(.P0).momentum_X = P(.P0).momentum_X + .f * (Dx1 + Dx2)
            P(.P0).momentum_Y = P(.P0).momentum_Y + .f * (Dy1 + Dy2)
            
            '    P(.P0).momentum_X = P(.P0).momentum_X + .f * Dx1
            '    P(.P0).momentum_Y = P(.P0).momentum_Y + .f * Dy1
            
            '            End If
            
        End With
        
    Next
    
End Sub
Public Function ApplyMotorForce(wPoint As Long, DirX, DirY)
    
    With P(wPoint)
        
        If .IsWheel Then
            
            If .Motor <> 0 Then
                
                'Apply motor Force
                .momentum_X = .momentum_X + DirX * .Motor * .Radius * 0.05
                .momentum_Y = .momentum_Y + DirY * .Motor * .Radius * 0.05
                
                .momentum_C = 31 * (.momentum_X + Abs(.momentum_Y) * Sgn(.momentum_X)) _
                        / .Radius
                
                .momentum_C = .momentum_C + _
                        31 * (.Motor) / .Radius
                
            Else
                
                .momentum_C = 31 * (.momentum_X + Abs(.momentum_Y) * Sgn(.momentum_X)) _
                        / .Radius
                
            End If
            
        End If
    
    End With
    
End Function




'Sub chkCircleLine(ByRef BOT As clsBOT, Ground As clsBOT)
'
'    Dim I As Long
'    Dim j As Long
'    Dim i2 As Long
'
'    Dim TouchPoint As tPoint
'    Dim Reflect As tPoint
'
'    Dim P1 As tPoint
'    Dim P2 As tPoint
'    Dim C As tPoint 'Circle center
'    Dim R As Single 'Circle Radius
'    Dim WallNormalized As tPoint
'
'    Dim jFrom As Long
'    Dim jTo As Long
'    Dim jStep As Long
'
'    For I = 1 To BOT.NP
'
'        C = MakePoint(BOT.GetPointX(I), BOT.GetPointY(I))
'        R = BOT.GetPointR(I)
'        'C  = Center of Bot Point (i)
'        'R  = Radius of Bot Point (i)
'
'        'For some reason works better if (J) For follow MomentumX direction
'        If BOT.GetPointMomX(I) >= 0 Then
'            jFrom = 1
'            jTo = Ground.NL
'            jStep = 1
'        Else
'            jFrom = Ground.NL
'            jTo = 1
'            jStep = -1
'        End If
'
'        For j = jFrom To jTo Step jStep
'
'
'            P1 = MakePoint(Ground.GetPointX(Ground.GetLinkP1(j)), Ground.GetPointY(Ground.GetLinkP1(j)))
'            P2 = MakePoint(Ground.GetPointX(Ground.GetLinkP2(j)), Ground.GetPointY(Ground.GetLinkP2(j)))
'
'            'P1 and P2 are the points of (Ground) Link (j)
'            'Maybe a better way to Get Them
'
'
'
'            If P1.x > BOT.MaxX Or P2.x > BOT.MaxX Or _
'               P1.x < BOT.MinX Or P2.x < BOT.MinX Then
'            Else
'
'            TouchPoint = CIRCLE_LINE(P1, P2, C, R)
'
'            If TouchPoint.x > 0 And TouchPoint.y > 0 Then
'
'
'                '   PIC.Circle (Touchpoint.x + PanX, Touchpoint.y), 5, vbYellow ': Stop
'
'                P1.x = P2.x - P1.x
'                P1.y = P2.y - P1.y
'                WallNormalized = Normalize(P1)
'                '    PIC.Line (Ret.x + PanX, Ret.y - 2)-(Ret.x + PanX + WallNormalized.x * BOT.GetMotor(i) * 20, Ret.y + WallNormalized.y * BOT.GetMotor(i) * 20 - 2), vbYellow ': Stop
'
'                Reflect = getREFLECT(MakePoint(BOT.GetPointMomX(I), BOT.GetPointMomY(I)), P1)
'
'                'Reflect.X = Reflect.X * WallBounce
'                'Reflect.Y = Reflect.Y * WallBounce
'
'                BOT.SetPointMomX(I) = -Reflect.x
'                BOT.SetPointMomY(I) = -Reflect.y
'
'                BOT.SetPointPosEqualOldPos (I)
'                BOT.SetPointX(I) = BOT.GetPointX(I) - Reflect.x '* 2
'                BOT.SetPointY(I) = BOT.GetPointY(I) - Reflect.y '* 2
'
'
'                ' if a Point Touch the First or the Last "Ground" segment
'                ' then Invert motors Forces.
'                If j = 1 Or j = Ground.NL Then
'                    For i2 = 1 To BOT.NP
'                        BOT.SetMotor(i2) = -BOT.GetMotor(i2)
'                    Next
'                End If
'
'
'                BOT.ApplyMotorForce I, WallNormalized.x, WallNormalized.y
'
'                If Abs(BOT.GetMotor(I) <> 0) Then Part.AddParticle TouchPoint.x, TouchPoint.y, _
'                        -WallNormalized.x * BOT.GetMotor(I), -WallNormalized.y * BOT.GetMotor(I)
'
'
'                'Exit J For
'                j = jTo
'
'                End If
'
'            End If
'
'        Next
'    Next
'
'End Sub

