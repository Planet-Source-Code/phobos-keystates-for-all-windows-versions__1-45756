VERSION 5.00
Begin VB.Form frmKeystates 
   Caption         =   "Keystates Control for all windows versions"
   ClientHeight    =   3030
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   5460
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   3030
   ScaleWidth      =   5460
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command4 
      Caption         =   "Toggle Scroll"
      Height          =   300
      Left            =   3735
      TabIndex        =   9
      Top             =   1680
      Width           =   1290
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Toggle Num"
      Height          =   300
      Left            =   3735
      TabIndex        =   8
      Top             =   480
      Width           =   1290
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Toggle Caps"
      Height          =   300
      Left            =   3735
      TabIndex        =   7
      Top             =   1080
      Width           =   1290
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Refresh"
      Height          =   315
      Left            =   2175
      TabIndex        =   6
      Top             =   2460
      Width           =   1335
   End
   Begin VB.TextBox Text3 
      Alignment       =   2  'Center
      BackColor       =   &H00E0E0E0&
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   480
      Left            =   2625
      TabIndex        =   4
      Text            =   "OFF"
      Top             =   1575
      Width           =   780
   End
   Begin VB.TextBox Text2 
      Alignment       =   2  'Center
      BackColor       =   &H00E0E0E0&
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   480
      Left            =   2625
      TabIndex        =   2
      Text            =   "OFF"
      Top             =   345
      Width           =   780
   End
   Begin VB.TextBox Text1 
      Alignment       =   2  'Center
      BackColor       =   &H00E0E0E0&
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   480
      Left            =   2625
      TabIndex        =   0
      Text            =   "OFF"
      Top             =   975
      Width           =   780
   End
   Begin VB.Label Label3 
      Alignment       =   1  'Right Justify
      Caption         =   "ScrollLock:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   615
      TabIndex        =   5
      Top             =   1620
      Width           =   1875
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      Caption         =   "NumLock:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   615
      TabIndex        =   3
      Top             =   390
      Width           =   1875
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      Caption         =   "CapsLock:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   615
      TabIndex        =   1
      Top             =   1020
      Width           =   1875
   End
End
Attribute VB_Name = "frmKeystates"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Const VK_NUMLOCK = &H90
Private Const VK_SCROLL = &H91
Private Const VK_CAPITAL = &H14
Private Const KEYEVENTF_EXTENDEDKEY = &H1
Private Const KEYEVENTF_KEYUP = &H2
Private Const VER_PLATFORM_WIN32_NT = 2
Private Const VER_PLATFORM_WIN32_WINDOWS = 1

Private Type KeyboardBytes
     kbByte(0 To 255) As Byte
End Type

Private Type OSVERSIONINFO
    dwOSVersionInfoSize As Long
    dwMajorVersion As Long
    dwMinorVersion As Long
    dwBuildNumber As Long
    dwPlatformId As Long
    szCSDVersion As String * 128
End Type
    
Private Declare Function GetVersionEx Lib "kernel32" Alias "GetVersionExA" (lpVersionInformation As OSVERSIONINFO) As Long
Private Declare Function GetKeyState Lib "user32" (ByVal nVirtKey As Long) As Integer
Private Declare Function GetKeyboardState Lib "user32" (kbArray As KeyboardBytes) As Long
Private Declare Function SetKeyboardState Lib "user32" (lppbKeyState As Byte) As Long

Private Declare Sub keybd_event Lib "user32" (ByVal bVK As Byte, ByVal bScan As Byte, ByVal dwFlags As Long, ByVal dwExtraInfo As Long)

Private Sub Command1_Click()
    Call ShowKeyStates
End Sub

Private Sub Command2_Click()
    Call SimulateKeypress(VK_CAPITAL)
    Call ShowKeyStates
End Sub

Private Sub Command3_Click()
    Call SimulateKeypress(VK_NUMLOCK)
    Call ShowKeyStates
End Sub

Private Sub Command4_Click()
    
    Call SimulateKeypress(VK_SCROLL)
    Call ShowKeyStates

End Sub

Private Sub SimulateKeypress(bVK As Byte)
    
    Dim keys(0 To 255) As Byte
    Dim o As OSVERSIONINFO
    
    o.dwOSVersionInfoSize = Len(o)
    GetVersionEx o
    If o.dwPlatformId = VER_PLATFORM_WIN32_WINDOWS Then ' Win95/Win98
        keys(bVK) = IIf(GetKeyState(bVK) = 0, 1, 0)
        Call SetKeyboardState(keys(0))
    Else
        ' Simulate key press then release for win2k.
        Call keybd_event(bVK, &H45, KEYEVENTF_EXTENDEDKEY Or 0, 0)
        Call keybd_event(bVK, &H45, KEYEVENTF_EXTENDEDKEY Or KEYEVENTF_KEYUP, 0)
        DoEvents
    End If
    
End Sub

Private Sub Form_Load()
    Call ShowKeyStates
End Sub

Sub ShowKeyStates()
    Text1.Text = IIf(GetCapslock, "ON", "OFF")
    Text2.Text = IIf(GetNumlock, "ON", "OFF")
    Text3.Text = IIf(GetScrollLock, "ON", "OFF")
End Sub

Function GetCapslock() As Boolean
    ' Return or set the Capslock toggle.
    GetCapslock = CBool(GetKeyState(vbKeyCapital) And 1)
End Function

Function GetNumlock() As Boolean
    ' Return or set the Numlock toggle.
    GetNumlock = CBool(GetKeyState(vbKeyNumlock) And 1)
End Function

Function GetScrollLock() As Boolean
    ' Return or set the ScrollLock toggle.
    GetScrollLock = CBool(GetKeyState(vbKeyScrollLock) And 1)
End Function
