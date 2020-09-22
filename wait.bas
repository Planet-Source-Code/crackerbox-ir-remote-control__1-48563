Attribute VB_Name = "Pause"
'**************************************
'Windows API/Global Declarations for :A
'     simple Wait Function
'**************************************

Public Declare Function GetTickCount Lib "kernel32.dll" () As Long
'the api for detecting keys pressed or scan coded
Private Declare Function GetAsyncKeyState Lib "user32" (ByVal vKey As Long) As Integer

Public Function Wait(ByVal TimeToWait As Long) 'Time in seconds
    Dim EndTime As Long
    
    EndTime = GetTickCount + TimeToWait / 1 '* 1000 Cause u give seconds and GetTickCount uses Milliseconds

    Do Until GetTickCount > EndTime

        DoEvents
        
    Loop
        
End Function
