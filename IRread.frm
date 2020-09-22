VERSION 5.00
Object = "{648A5603-2C6E-101B-82B6-000000000014}#1.1#0"; "MSCOMM32.OCX"
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Begin VB.Form IRread 
   AutoRedraw      =   -1  'True
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Ir Read"
   ClientHeight    =   3585
   ClientLeft      =   45
   ClientTop       =   300
   ClientWidth     =   9870
   DrawMode        =   1  'Blackness
   Icon            =   "IRread.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3585
   ScaleWidth      =   9870
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox Signal_box 
      Height          =   375
      Left            =   120
      MaxLength       =   1024
      TabIndex        =   14
      TabStop         =   0   'False
      Top             =   4200
      Width           =   3375
   End
   Begin VB.TextBox signal_out_box 
      Height          =   375
      Left            =   3600
      MaxLength       =   1024
      TabIndex        =   13
      TabStop         =   0   'False
      Top             =   4200
      Width           =   3015
   End
   Begin VB.Timer Timer2 
      Enabled         =   0   'False
      Interval        =   500
      Left            =   6720
      Top             =   4200
   End
   Begin VB.Frame Frame1 
      Caption         =   "Demo Ir Remote Control record and playback program"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   3375
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   9615
      Begin VB.CommandButton Comm 
         Caption         =   "C&omm"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   7680
         TabIndex        =   19
         Top             =   240
         Width           =   855
      End
      Begin ComctlLib.Slider Slider2 
         Height          =   255
         Left            =   6480
         TabIndex        =   17
         Top             =   3000
         Width           =   3015
         _ExtentX        =   5318
         _ExtentY        =   450
         _Version        =   327682
         LargeChange     =   1
         SelStart        =   6
         Value           =   6
      End
      Begin ComctlLib.Slider Slider1 
         Height          =   255
         Left            =   1560
         TabIndex        =   15
         Top             =   3000
         Width           =   3015
         _ExtentX        =   5318
         _ExtentY        =   450
         _Version        =   327682
         LargeChange     =   1
         Min             =   1
         SelStart        =   2
         Value           =   2
      End
      Begin VB.TextBox Text2 
         Alignment       =   2  'Center
         BackColor       =   &H00000000&
         BeginProperty Font 
            Name            =   "Book Antiqua"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   435
         Left            =   7920
         TabIndex        =   7
         TabStop         =   0   'False
         Top             =   1080
         Width           =   1575
      End
      Begin VB.PictureBox Picture_line 
         AutoRedraw      =   -1  'True
         BackColor       =   &H00000000&
         Height          =   1215
         Left            =   120
         Negotiate       =   -1  'True
         ScaleHeight     =   77
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   621
         TabIndex        =   6
         TabStop         =   0   'False
         Top             =   1680
         Width           =   9375
      End
      Begin VB.TextBox receive_box 
         BackColor       =   &H00000000&
         ForeColor       =   &H000000FF&
         Height          =   375
         Left            =   120
         MaxLength       =   54
         TabIndex        =   5
         TabStop         =   0   'False
         Top             =   240
         Width           =   4935
      End
      Begin VB.TextBox Text1 
         BackColor       =   &H00000000&
         ForeColor       =   &H000000FF&
         Height          =   375
         Left            =   120
         MaxLength       =   32
         TabIndex        =   4
         TabStop         =   0   'False
         Top             =   1200
         Width           =   4935
      End
      Begin VB.TextBox Code_box 
         BackColor       =   &H00000000&
         ForeColor       =   &H000000FF&
         Height          =   375
         Left            =   120
         Locked          =   -1  'True
         MaxLength       =   58
         TabIndex        =   3
         TabStop         =   0   'False
         Top             =   720
         Width           =   4935
      End
      Begin VB.CommandButton cmdExit 
         Caption         =   "E&xit"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   8640
         TabIndex        =   2
         Top             =   240
         Width           =   855
      End
      Begin VB.CommandButton cmdClear 
         Caption         =   "&Clear"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   6720
         TabIndex        =   1
         Top             =   240
         Width           =   855
      End
      Begin VB.Label Label7 
         Caption         =   "Signal Contrast"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   255
         Left            =   4680
         TabIndex        =   18
         Top             =   3000
         Width           =   1695
      End
      Begin VB.Label Label3 
         Caption         =   "Grid Contrast"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   255
         Left            =   120
         TabIndex        =   16
         Top             =   3000
         Width           =   1455
      End
      Begin VB.Label Label6 
         Alignment       =   2  'Center
         Caption         =   "Button Code"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   255
         Left            =   7920
         TabIndex        =   12
         Top             =   720
         Width           =   1455
      End
      Begin VB.Label Label5 
         Caption         =   "Raw Sample"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   255
         Left            =   5160
         TabIndex        =   11
         Top             =   360
         Width           =   1335
      End
      Begin VB.Label Label4 
         Caption         =   "Actual Code"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   375
         Left            =   5160
         TabIndex        =   10
         Top             =   840
         Width           =   1335
      End
      Begin VB.Label Label2 
         Caption         =   "Binary Code"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   375
         Left            =   5160
         TabIndex        =   9
         Top             =   1320
         Width           =   1455
      End
      Begin VB.Label Label1 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   255
         Left            =   4080
         TabIndex        =   8
         Top             =   2280
         Width           =   1695
      End
   End
   Begin MSCommLib.MSComm MSComm1 
      Left            =   7320
      Top             =   4200
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
      CommPort        =   2
      DTREnable       =   0   'False
      Handshaking     =   3
      ParityReplace   =   48
      BaudRate        =   19200
      InputMode       =   1
   End
End
Attribute VB_Name = "IRread"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit
Public linediagram As New Cls_diagram
Public Comm_Port As Integer
Private signal(1 To 1024) As Single
Private Grid As Integer

Private Sub Comm_Click()

frmOptions.Show

End Sub

'
Private Sub Form_Load()

    Dim hSysMenu As Long

    If App.PrevInstance = True Then

        Unload Me
        
    End If
    
    cmdClear_Click
    'Set up the Comm port
    Comm_settings

End Sub
'Close Comport when exiting the program
Private Sub Form_Unload(Cancel As Integer)

    If MSComm1.PortOpen = True Then MSComm1.PortOpen = False

End Sub
'Clear the Send_box and the Receive_box text boxes
Private Sub cmdClear_Click()

Dim l, g As Integer

    Grid = RGB(0, Slider1.value * 25, 0)
    Timer2.Enabled = True
    Code_box.Text = ""
    receive_box.Text = ""
    Text1.Text = ""
    Text2.Text = ""
    signal_out_box.Text = ""
    
    For g = 1 To 1024
    
        signal(g) = 0
        
    Next
    
    'Signal Header
    Signal_box.Text = "00000000011111111000000100"
    
    MSComm1.InBufferCount = 0
    MSComm1.OutBufferCount = 0
    
    With linediagram
        .InitDiagram Picture_line, RGB(Slider2.value * 25, 0, 0), True, Grid
        .Max = 10
        .HorzSplits = 10
        .VertSplits = 75
        .DiagramType = TYPE_LINE
        For g = 1 To Picture_line.ScaleWidth  '512 'Len(receive_box.Text)
            
            linediagram.AddValue signal(g) + 2
        Next
        .RePaint
    End With
    
End Sub
'Print serial data received to Receive_box
Private Sub MSComm1_OnComm()

Dim X, g, bum As Integer
Dim value As String


  Select Case MSComm1.CommEvent
  
    Case comEventRxOver
      MsgBox ("Receive buffer overflow")
      
    Case comEventTxFull
      MsgBox ("Send buffer overflow")
          
    Case comEvReceive
            X = Asc(MSComm1.Input)
        receive_box = receive_box + Hex(X)

    Case comEvCD
    
End Select

End Sub
Private Sub cmdExit_Click()

    Unload Me
    
End Sub
Private Function HexToBinStr(ByVal inHex As String) As String

    Dim mDec As Integer
    Dim s As String
    Dim i
    
    mDec = CInt("&h" & inHex)
    s = Trim(CStr(mDec Mod 2))
    i = mDec \ 2
    
    Do While i <> 0
        s = Trim(CStr(i Mod 2)) & s
        i = i \ 2
    Loop
    
    Do While Len(s) < 4 ' 8
        s = "0" & s
    Loop
    
    HexToBinStr = s
    Exit Function
    
End Function
Private Function bin_to_Hex(ByVal inBin As String) As String

Dim n, size, hex_byte, Filter_byte As Integer
Dim bin, ones, twos, fours, eights, temp, Trans_out As String

        temp = Text1 ' receive_box
        size = Len(temp)
        
        For n = 1 To size Step 4
        
            bin = Mid$(temp, n, 4)
            ones = Val(Mid$(bin, 4, 1))
            twos = Val(Mid$(bin, 3, 1))
            fours = Val(Mid$(bin, 2, 1))
            eights = Val(Mid$(bin, 1, 1))

        
            hex_byte = ones + (2 * twos) + (4 * fours) + _
                       (8 * eights)
            Filter_byte = hex_byte 'And &H5
            'Code_box.Text = Code_box.Text & Hex(Filter_byte)
            Text2.Text = Text2.Text + Hex(Filter_byte)
        Next n
        
        'Trans_out = Code_box.Text
        Trans_out = Text2.Text
        
End Function
'Signal and Button formatting
Private Sub Code_thing()

Dim temp, value, value2 As String
Dim size, g, t, h, ty, bob As Integer

    temp = receive_box
    size = Len(temp)
    
    'filter signal
    For g = 1 To size
    
            value = Mid$(temp, g, 1)
                If value = "F" Or value = "E" Or value = "D" Or value = "C" Then value = "1"
                If value = "8" Then value = "0"
            Code_box = Code_box + value
    Next
    
    'decode signal
    For t = 1 To size
        value = Mid$(Code_box, t, 2)
            If value = "00" Then value = "1"
            If value = "10" Then value = "0"
            If value = "01" Then value = ""
        Text1 = Text1 + value
    Next
        
    'Draw the signal
    For t = 1 To Len(Text1.Text)
        value = Mid$(Text1.Text, t, 1)
        
            If value = "1" Then value2 = "010"
            If value = "0" Then value2 = "10"

        Signal_box.Text = Signal_box.Text + value2
        
    Next
           
    For ty = 1 To Len(Signal_box.Text)

        For h = 1 To 5
 
            signal_out_box.Text = signal_out_box.Text & Mid$(Signal_box.Text, ty, 1)
            
        Next
            
    Next
                     
    For g = 1 To Picture_line.ScaleWidth  '512 'Len(receive_box.Text)
    
            bob = bob + 1
   
            value = Mid$(signal_out_box.Text, bob, 1)
            
                If value = "" Then value = 0
                If value = 1 Then value = value + 5
            
            linediagram.AddValue value + 2
            signal(g) = value

    Next
    bob = 0
    linediagram.RePaint
    
    'Convert Binary code to Hex
    bin_to_Hex (Text1)

End Sub
'Com port setup
Private Sub Comm_settings()

On Error Resume Next

    MSComm1.Settings = ("14400" + "," + "n" + "," + "8" + "," + "1")
    MSComm1.InputLen = 1
    MSComm1.PortOpen = True
    MSComm1.InputMode = comInputModeText
    MSComm1.RThreshold = 1

End Sub
'Grid Contrast control
Private Sub Slider1_Scroll()
Dim g As Integer
Grid = RGB(0, Int(Slider1.value * 25.5), 0)

With linediagram
        .InitDiagram Picture_line, RGB(Int(Slider2.value * 25.5), 0, 0), True, Grid
        
        For g = 1 To Picture_line.ScaleWidth  '512 'Len(receive_box.Text)
            
            linediagram.AddValue signal(g) + 2
        Next
        
        .RePaint
End With

End Sub
'Signal Contrast control
Private Sub Slider2_Scroll()
Dim g As Integer
Grid = RGB(0, Int(Slider1.value * 25.5), 0)

With linediagram

        .InitDiagram Picture_line, RGB(Int(Slider2.value * 25.5), 0, 0), True, Grid
        .Max = 10
        .HorzSplits = 10
        .VertSplits = 75
        .DiagramType = TYPE_LINE
        
        For g = 1 To Picture_line.ScaleWidth
            
            linediagram.AddValue signal(g) + 2
            
        Next
        .RePaint
        
End With

End Sub
Private Sub Timer2_Timer()

If Len(receive_box.Text) = 54 Then

    Code_thing
    Timer2.Enabled = False
    Exit Sub
    
End If

    Code_box.Text = ""
    receive_box.Text = ""
    Text1.Text = ""
    Label1.Caption = ""
    MSComm1.InBufferCount = 0
    MSComm1.OutBufferCount = 0

End Sub
