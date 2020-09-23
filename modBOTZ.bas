Attribute VB_Name = "modBOTZ"
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

Public Type tPoint
    
    x As Single
    y As Single
    OldX As Single
    OldY As Single
    NewX As Single
    NewY As Single
    
    IsWheel As Boolean
    Radius As Single
    Heading As Single
    Motor As Single
    IsMotor As Boolean
    
    momentum_X As Single
    momentum_Y As Single
    momentum_C As Single
    
    
    
    
End Type

Public Type tLink
    
    P1 As Long
    P2 As Long
    CurrLen As Single
    
    TENS As Single
    
    LastLen As Single
    midX As Single
    midY As Single
    
    DynPhase As Single
    DynAmp As Single
    DynSpeed As Single
    
    IsDynamic As Boolean
    
End Type


Public Type tMuscle
    L1 As Integer '     Link1
    L2 As Integer '     Link2
    MainA As Double '   Angle that should be between L1 and L2
    P0 As Integer '     Common point of L1 and L2
    P1 As Integer '     Other point on L1
    P2 As Integer '     Other point on L2
    f As Double '       Muscle Force(strength)
    
    
    
    DynPhase As Single
    DynAmp As Single
    DynSpeed As Single
    
    IsDynamic As Boolean
    
    
    
    
    
    isNotBroken As Boolean
    
    
    FixedANG As Boolean
    
    
End Type


'Global Variables:---------------------------------------------------
Global Gravity As Single
Global Atmosphere As Single
Global WallBounce As Single
Global WallFriction As Single
Global WIND As Single

'Global Tension As Single
'Global ClockSpeed As Single

Global RightWall As Integer
Global FloorWall As Integer


Global PanX As Single

Global Const PI = 3.14159265358979


Public BOT As New clsBOT
Public Ground As New clsBOT

Public Part As New clsParticle


Public Function PointDist(P1 As tPoint, P2 As tPoint) As Single
    
    Dim dX As Single
    Dim dY As Single
    
    dX = P1.x - P2.x
    dY = P1.y - P2.y
    
    PointDist = Sqr(dX * dX + dY * dY)
    
    
End Function


Public Function Atan2(ByVal dX As Double, ByVal dY As Double) As Double
    'This Should return Angle
    
    Dim theta As Double
    
    If (Abs(dX) < 0.0000001) Then
        If (Abs(dY) < 0.0000001) Then
            theta = 0#
        ElseIf (dY > 0#) Then
            theta = 1.5707963267949
            'theta = PI / 2
        Else
            theta = -1.5707963267949
            'theta = -PI / 2
        End If
    Else
        theta = Atn(dY / dX)
        
        If (dX < 0) Then
            If (dY >= 0#) Then
                theta = PI + theta
            Else
                theta = theta - PI
            End If
        End If
    End If
    
    Atan2 = theta
End Function

