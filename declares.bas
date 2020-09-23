Attribute VB_Name = "Declares"
Option Explicit

Public Declare Function BitBlt Lib "gdi32" (ByVal destdc As Long, ByVal X As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal srcdc As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long) As Long
Public Declare Function GetTickCount& Lib "kernel32" ()
Public Const SRCCOPY = &HCC0020

Type link_type
   used As Boolean
   target1_id As Integer
   target2_id As Integer
   linklength As Single
   linktension As Single
   pushtiming As Integer
   pushspan As Integer
   pushstrength As Single
   Push As Single
   lastlen As Single
   midx As Single
   midy As Single
   phase As Byte
End Type

Type vertex_type
   justreleased As Boolean
   Selected As Boolean
   used As Boolean
   X As Single
   y As Single
   LastX As Single
   Lasty As Single
   momentum_x As Double 'doubles seem to be more accurate
   momentum_y As Double 'than singles *shrug*
   momentum_c As Double
   Radius As Integer
   wheel As Boolean
   heading As Single
   lightmode As Boolean
   phase As Byte
End Type


'Global Variables:---------------------------------------------------
Global Gravity As Single
Global Atmosphere As Single
Global WallBounce As Single
Global WallFriction As Single
Global LeftWind As Single
Global Tension As Single
Global ClockSpeed As Single

'Descriptions:
'
'   Gravity - Force pulling downward, can be negative, thus pushing
'             upward against the ceiling.
'
'   Atmosphere - the air resistance applied to the vertices
'                crazy things should happen if you make this negative
'
'   WallBounce - how much of an objects momentum is relfected when
'                it bounces off the floor, the ceiling or a wall.
'
'   WallFriction - how much of an object's momentum is decreased
'                  as it is dragged along a wall or floor.
'                  0 = perfectly slippery
'
'   LeftWind - The amount of wind blowing from the left.
'              If this value is negative, the wind will blow from the
'              right instead.
'
'   Tension - how hard the links will 'fight' to maintain the distance
'             between the vertices it is attached to.
'
'---------------------------------------------------------------------


Sub experimental_stuff()


'This is incomplete stuff that may one day make
'non-orthogonal walls that bots can interact with


'For A = 1 To MaxLinks
'
'          If Link(A).used = True And Link(A).phase = 0 And Link(A).x1 = Link(A).x2 Then
'            If newx + WallPadding + .Radius > Link(A).x1 And Link(A).y1 <= (newy + 3) And (newy - 3) <= Link(A).y2 And .LastX <= Link(A).x2 - WallPadding Then
'            .X = Link(A).x1 - WallPadding - .Radius
'            .momentum_x = (.momentum_x * WallBounce) * -1
'            .momentum_y = .momentum_y * (1 - Fric)
'            End If
'            If newx - WallPadding + .Radius < Link(A).x1 And Link(A).y1 <= (newy + 3) And (newy - 3) <= Link(A).y2 And .LastX >= Link(A).x2 + WallPadding Then
'            .X = Link(A).x1 + WallPadding + .Radius
'            .momentum_x = (.momentum_x * WallBounce) * -1
'            .momentum_y = .momentum_y * (1 - Fric)
'            End If
'            'GoTo nexterPA
'          End If
'
'          If Link(A).used = True And Link(A).phase = 0 And Link(A).y1 = Link(A).y2 Then
'            If newy + WallPadding + .Radius > Link(A).y1 And Link(A).x1 >= (newx - 3) And (newx + 3) >= Link(A).x2 And .Lasty <= Link(A).y2 - WallPadding Then
'            .y = Link(A).y1 - WallPadding - .Radius
'            .momentum_y = (.momentum_y * WallBounce) * -1
'            .momentum_x = .momentum_x * (1 - Fric)
'            End If
'            If newy - WallPadding - .Radius < Link(A).y1 And Link(A).x1 >= (newx - 3) And (newx + 3) >= Link(A).x2 And .Lasty >= Link(A).y2 + WallPadding Then
'            .y = Link(A).y1 + WallPadding + .Radius
'            .momentum_y = (.momentum_y * WallBounce) * -1
'            .momentum_x = .momentum_x * (1 - Fric)
'            End If
'            'GoTo nexterPA
'          End If
'
'          T1X = vertex(Link(A).target1_id).X: T2X = vertex(Link(A).target2_id).X
'          T1Y = vertex(Link(A).target1_id).y: T2Y = vertex(Link(A).target2_id).y
'
'          If Link(A).used = True And Link(A).phase = 0 And T1Y > T2Y And T1X > T2X Then
'            If .X < newx And .y > newy And CrossLine(newx, newy, .LastX, .Lasty, T1X, T1Y, T2X, T2Y) = True Then
'             'MsgBox "ok"
'             TVA = Abs(T1Y - T2Y): THA = Abs(T1X - T2X)
'
'
'             If newx + WallPadding + .Radius >= rtnX Then
'             .X = rtnX - WallPadding - .Radius
'             .momentum_x = .momentum_x * (Fric - (THA / (THA + TVA)))
'             .momentum_x = .momentum_x - ((.momentum_x * 2) * (WallBounce * THA / (THA + TVA)))
'             End If
'
'             If newy + WallPadding + .Radius <= rtnY Then
'             .y = rtnY + WallPadding - .Radius
'             .momentum_y = .momentum_y * (Fric - (TVA / (TVA + THA)))
'             .momentum_y = .momentum_y - ((.momentum_y * 2) * (WallBounce * TVA / (TVA + THA)))
'             End If
'
'            End If
'          End If
'
'nexterPA:
'
'Next A

End Sub


Function Phase_Color(phase) As Long


If phase = 0 Then Phase_Color = RGB(0, 0, 255): Exit Function
If phase = 1 Then Phase_Color = RGB(0, 0, 0): Exit Function
If phase = 2 Then Phase_Color = RGB(255, 0, 0): Exit Function
If phase = 3 Then Phase_Color = RGB(0, 255, 0): Exit Function
If phase = 4 Then Phase_Color = RGB(255, 0, 255): Exit Function

Phase_Color = RGB(0, 0, 0)


End Function


