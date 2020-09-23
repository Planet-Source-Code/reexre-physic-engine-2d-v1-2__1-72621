Attribute VB_Name = "modKeyBoard"
Public Declare Function GetAsyncKeyState Lib "user32" (ByVal vKey As Long) As Integer
Public Declare Function GetKeyState Lib "user32" (ByVal nVirtKey As Long) As Integer


Public Function KeyBoard()
    
    Dim I As Long
    Dim BotNP As Long
    
    BotNP = BOT.NP
    
    For I = 1 To BotNP
        If BOT.IsMotor(I) Then BOT.SetMotor(I) = 0
    Next
    
    If GetAsyncKeyState(vbKeyLeft) <> 0 Then
        For I = 1 To BotNP
            If BOT.IsMotor(I) Then BOT.SetMotor(I) = -1
        Next
        
    End If
    If GetAsyncKeyState(vbKeyRight) <> 0 Then
        For I = 1 To BotNP
            If BOT.IsMotor(I) Then BOT.SetMotor(I) = 1
        Next
    End If
    
End Function
