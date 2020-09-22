Attribute VB_Name = "XClose"
'**************************************
'Windows API/Global Declarations for :Di
'     sable 'X' on Forms (Including MDI Child


'     Forms)
'**************************************
' Place this in a module:


Public Declare Function GetSystemMenu Lib "user32" (ByVal hwnd As Long, ByVal bRevert As Long) As Long


Public Declare Function RemoveMenu Lib "user32" (ByVal hMenu As Long, ByVal nPosition As Long, ByVal wFlags As Long) As Long
    Public Const SC_CLOSE = &HF060&
    Public Const MF_BYCOMMAND = &H0&
