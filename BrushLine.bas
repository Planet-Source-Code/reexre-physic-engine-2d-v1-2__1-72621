Attribute VB_Name = "BrushLine"
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

Public Type POINTAPI
    x As Long
    y As Long
End Type

Public poi As POINTAPI


Public Declare Function GetPixel Lib "gdi32" (ByVal hdc As Long, ByVal x As Long, ByVal y As Long) As Long
Public Declare Function SetPixel Lib "gdi32" (ByVal hdc As Long, ByVal x As Long, ByVal y As Long, ByVal crColor As Long) As Long

Public Declare Function LineTo Lib "gdi32" (ByVal hdc As Long, ByVal x As Long, ByVal y As Long) As Long
Public Declare Function MoveToEx Lib "gdi32" (ByVal hdc As Long, ByVal x As Long, ByVal y As Long, lpPoint As POINTAPI) As Long
Public Declare Function CreatePen Lib "gdi32" (ByVal nPenStyle As Long, ByVal nWidth As Long, ByVal crColor As Long) As Long
Public Declare Function SelectObject Lib "gdi32" (ByVal hdc As Long, ByVal hObject As Long) As Long
Public Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long

Public Declare Function Arc Lib "gdi32" (ByVal hdc As Long, _
        ByVal xInizioRettangolo As Long, _
        ByVal yInizioRettangolo As Long, _
        ByVal xFineRettangolo As Long, _
        ByVal yFineRettangolo As Long, _
        ByVal xInizioArco As Long, _
        ByVal yInizioArco As Long, _
        ByVal xFineArco As Long, _
        ByVal yFineArco As Long) As Long

Public Declare Function BitBlt Lib "gdi32" (ByVal hDestDC As Long, ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long) As Long


'Declare Function Arc Lib "gdi32.dll" (ByVal HDC As Long, ByVal X1 As Long, _
ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long, ByVal X3 As Long, _
        ByVal Y3 As Long, ByVal X4 As Long, ByVal Y4 As Long) As Long


Public Sub SetBrush(ByVal hdc As Long, ByVal PenWidth As Long, ByVal PenColor As Long)
    
    
    DeleteObject (SelectObject(hdc, CreatePen(vbSolid, PenWidth, PenColor)))
    'kOBJ = SelectObject(hDC, CreatePen(vbSolid, PenWidth, PenColor))
    'SetBrush = kOBJ
    
    
End Sub



Public Sub FastLine(ByRef hdc As Long, ByRef X1 As Long, ByRef Y1 As Long, _
            ByRef X2 As Long, ByRef Y2 As Long, ByRef W As Long, ByRef color As Long)
    Attribute FastLine.VB_Description = "disegna line veloce"
    
    Dim poi As POINTAPI
    
    'SetBrush hdc, W, color
    DeleteObject (SelectObject(hdc, CreatePen(vbSolid, W, color)))
    
    MoveToEx hdc, X1, Y1, poi
    LineTo hdc, X2, Y2
    
End Sub

Sub MyCircle(ByRef hdc As Long, ByRef x As Long, ByRef y As Long, ByRef R As Long, W, color)
    Dim XpR As Long
    
    'SetBrush hdc, W, color
    DeleteObject (SelectObject(hdc, CreatePen(vbSolid, W, color)))
    
    XpR = x + R
    
    Arc hdc, x - R, y - R, XpR, y + R, XpR, y, XpR, y
    
End Sub


