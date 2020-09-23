VERSION 1.0 CLASS
BEGIN
MultiUse = -1 'True
Persistable = 0 'NotPersistable
DataBindingBehavior = 0 'vbNone
DataSourceBehavior  = 0 'vbNone
MTSTransactionMode  = 0 'NotAnMTSObject
END
Attribute VB_Name = "clsParticle"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'Instead of creating a new type for this class, it is used "tPoint" Type but
'IsMotor is used as IsAlive
'Healing is used as MaxAge

Public NP As Long 'Number Of Particles
Private P() As tPoint 'Array of Particles

Public IterCNT As Long 'Internal Countr

Public IsNotEmpty As Boolean 'Particle system is not empty?

Private Const PartCol As Long = 12950 'rgb(150,50,0)



Public Sub AddParticle(x, y, vx, vy, Optional Life = 150)
    'IsMotor is used as IsAlive
    'Healing is used as MaxAge
    
    Dim I As Long
    Dim AddTo As Long 'Array Index where to add Particle
    
    
    If NP <> 0 Then
        '------------------------
        'Find first 'Free' particle
        I = 1
        AddTo = 0
        Do
            If P(I).IsMotor = False Then 'p(i) is free
                AddTo = I
                I = NP
            End If
            I = I + 1
        Loop While I <= NP
        
        If AddTo = 0 Then 'Doesnt find free one
            NP = NP + 1
            ReDim Preserve P(NP)
            AddTo = NP
        End If
        '----------------------
    Else
        NP = NP + 1
        ReDim Preserve P(NP)
        AddTo = NP
    End If
    
    With P(AddTo)
        .x = x
        .y = y
        .momentum_X = vx
        .momentum_Y = vy - Gravity * 10 'Should be only "VY" but so it's smarter
        .Heading = Life
        .IsMotor = True
    End With
    
    IsNotEmpty = True
    
End Sub



Public Sub DRAW_and_Physics(ByRef PIChdc As Long, PanX As Single)
    'IsMotor is used as IsAlive
    'Healing is used as MaxAge
    
    Dim I As Long
    
    If IsNotEmpty Then
        
        For I = 1 To NP
            
            With P(I)
                If .y > FloorWall Then .IsMotor = False
                
                If .IsMotor Then
                    
                    FastLine PIChdc, .x + PanX, .y \ 1, .x + PanX + 1, .y + 1, 3, PartCol
                    
                    .x = .x + .momentum_X
                    .y = .y + .momentum_Y
                    
                    .momentum_X = .momentum_X * (1 - Atmosphere)
                    .momentum_Y = .momentum_Y * (1 - Atmosphere)
                    .momentum_Y = .momentum_Y + Gravity * 0.25
                    
                    .Heading = .Heading - 1
                    If .Heading < 0 Then .IsMotor = False
                    
                End If
                
            End With
        Next
        
        IterCNT = IterCNT + 1
        
        'Ever 200 calls to this sub we check if Particles system is empty
        '(Maybe UseLess)
        If IterCNT > 200 Then
            IsNotEmpty = False
            For I = 1 To NP
                If P(I).IsMotor Then IsNotEmpty = True: I = NP
            Next I
            If Not (IsNotEmpty) Then NP = 0
        End If
        
    End If
    
End Sub
