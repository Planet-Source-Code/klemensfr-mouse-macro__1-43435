VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form Form1 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Wood - MouseMacro"
   ClientHeight    =   1950
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8445
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1950
   ScaleWidth      =   8445
   StartUpPosition =   2  'CenterScreen
   Begin VB.CheckBox klickss 
      Alignment       =   1  'Right Justify
      Caption         =   "Mouseclicks"
      Height          =   255
      Left            =   3480
      TabIndex        =   13
      Top             =   840
      Value           =   1  'Checked
      Width           =   1215
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Info"
      Height          =   255
      Left            =   3600
      TabIndex        =   12
      Top             =   1560
      Width           =   1215
   End
   Begin MSComctlLib.ProgressBar ProgressBar1 
      Height          =   255
      Left            =   120
      TabIndex        =   9
      Top             =   1200
      Width           =   8175
      _ExtentX        =   14420
      _ExtentY        =   450
      _Version        =   393216
      BorderStyle     =   1
      Appearance      =   0
      Scrolling       =   1
   End
   Begin MSComctlLib.Slider Slider1 
      Height          =   495
      Left            =   4080
      TabIndex        =   4
      Top             =   360
      Width           =   3375
      _ExtentX        =   5953
      _ExtentY        =   873
      _Version        =   393216
      LargeChange     =   1
      Min             =   1
      Max             =   11
      SelStart        =   3
      Value           =   3
   End
   Begin VB.CommandButton Command5 
      Caption         =   "Stop"
      Height          =   495
      Left            =   1680
      TabIndex        =   3
      Top             =   600
      Width           =   1215
   End
   Begin VB.CommandButton Command4 
      Caption         =   "Stop"
      Height          =   495
      Left            =   120
      TabIndex        =   2
      Top             =   600
      Width           =   1215
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Play"
      Height          =   495
      Left            =   1680
      TabIndex        =   1
      Top             =   120
      Width           =   1215
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Record"
      Height          =   495
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   1215
   End
   Begin VB.Timer Timer2 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   1320
      Top             =   600
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   1320
      Top             =   120
   End
   Begin VB.Label Mousebutton 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   4920
      TabIndex        =   11
      Top             =   1560
      Width           =   3375
   End
   Begin VB.Label lblstatus 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Status"
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   120
      TabIndex        =   10
      Top             =   1560
      Width           =   3375
   End
   Begin VB.Label Label4 
      Caption         =   "Normal"
      Height          =   255
      Left            =   5520
      TabIndex        =   8
      Top             =   120
      Width           =   495
   End
   Begin VB.Label Label3 
      Caption         =   "Quick"
      Height          =   255
      Left            =   3480
      TabIndex        =   7
      Top             =   480
      Width           =   615
   End
   Begin VB.Label Label2 
      Caption         =   "Slow"
      Height          =   255
      Left            =   7440
      TabIndex        =   6
      Top             =   480
      Width           =   855
   End
   Begin VB.Label Label1 
      Caption         =   "Speed:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   3480
      TabIndex        =   5
      Top             =   0
      Width           =   735
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Wood - MouseMacro
'
'Written by: Klemens FRIEDL 02/2003
'Some Codes from: Marcin Bedner & Dodge (do mouseclicks), http://www.activevb.de (show mouseclicks)
'
'My website: www.sharksoft.net.tc
'
'Don't forget to Vote ;)
'

Option Explicit

Public MacroFile As String
Public Länge As Long

'Mouse-Buttons:
Private Const VK_LBUTTON = &H1
Private Const VK_RBUTTON = &H2
Private Const VK_MBUTTON = &H4

'Wait-Funktion:
Private Declare Function GetTickCount Lib "kernel32" () As Long


'API Functions

#If Win32 Then 'Win32 declarations
Private Declare Function GetAsyncKeyState% Lib "user32" (ByVal vKey As Long) 'Gets state of one key
Private Declare Function GetCursorPos& Lib "user32" (lpPoint As POINTAPI) 'Gets current cursor position
Private Declare Function SetCursorPos& Lib "user32" (ByVal X As Long, ByVal Y As Long) 'Sets cursor position
Private Declare Sub mouse_event Lib "user32" (ByVal dwFlags As Long, ByVal dx As Long, ByVal dy As Long, _
ByVal cButtons As Long, ByVal dwExtraInfo As Long) 'Sends a mouse event
#Else ' Win16 declarations
Private Declare Function GetAsyncKeyState% Lib "user" (ByVal vKey As Integer)
Private Declare Sub GetCursorPos Lib "user" (lpPoint As POINTAPI)
Private Declare Sub SetCursorPos Lib "user" (ByVal X As Integer, ByVal Y As Integer)
'Function mouse_event is not available in the WIN16 API.
#End If 'WIN32

'API Types

Private Type POINTAPI
        X As Long
        Y As Long
End Type

'API Constants

Const MOUSEEVENTF_LEFTDOWN = &H2 '  left button down
Const MOUSEEVENTF_LEFTUP = &H4 '  left button up
Const MOUSEEVENTF_MIDDLEDOWN = &H20 '  middle button down
Const MOUSEEVENTF_MIDDLEUP = &H40 '  middle button up
Const MOUSEEVENTF_RIGHTDOWN = &H8 '  right button down
Const MOUSEEVENTF_RIGHTUP = &H10 '  right button up

Dim lasttag As Variant



Sub GetKeys(keys() As Boolean)
Dim ii As Variant
    ReDim keys(256) 'All keys
    For ii = 0 To 255
        If GetAsyncKeyState(ii) <> 0 Then 'If not 0 then the key is pressed
            keys(ii) = True
        Else
            keys(ii) = False
        End If
    Next ii
End Sub





Sub pickout()
Dim Werte() As String, i&
Dim Zw As String
Open MacroFile For Binary As #1
Zw = Space(LOF(1))
Get #1, , Zw
Close 1
Werte = Split(Zw, vbNewLine)
For i = 0 To UBound(Werte) - 1
ProgressBar1.Max = UBound(Werte)
Call interpret(Werte(i))
Next i
End Sub

Sub interpret(ByVal WertStr$)
Dim XX As String
Dim YY As String
Dim v As Variant
Dim L As Long
v = Split(WertStr)
L = UBound(v)
If L <= "1" Then
'Nothing
Else
Dim zerteilt() As String
Dim Wert As String
Dim Variabel As String
zerteilt = Split(WertStr, "#")
XX = zerteilt(1)
YY = zerteilt(3)
End If

Dim tag1 As Variant
Dim tag2 As Variant
tag1 = Split(WertStr, "+")(0)
tag2 = Split(WertStr, "+")(1)

 'Wend
 Wait (Slider1.Value)
Call SetCursorPos(XX, YY)
Dim xxx

If tag2 = "R" Then
    Mousebutton.Caption = "right Mouse-Button"
    If klickss.Value = 1 Then
        Call mouse_event(MOUSEEVENTF_RIGHTDOWN, 0, 0, 0, xxx)
        Call mouse_event(MOUSEEVENTF_RIGHTUP, 0, 0, 0, xxx)
    End If
ElseIf tag2 = "L" Then
    Mousebutton.Caption = "left Mouse-Button"
    If klickss.Value = 1 Then
        Call mouse_event(MOUSEEVENTF_LEFTDOWN, 0, 0, 0, xxx)
        Call mouse_event(MOUSEEVENTF_LEFTUP, 0, 0, 0, xxx)
    End If
ElseIf tag2 = "M" Then
    Mousebutton.Caption = "middle Mouse-Button"
    If klickss.Value = 1 Then
        Call mouse_event(MOUSEEVENTF_MIDDLEDOWN, 0, 0, 0, xxx)
        Call mouse_event(MOUSEEVENTF_MIDDLEUP, 0, 0, 0, xxx)
    End If
Else
    Mousebutton.Caption = ""
End If


Me.Caption = "X: " & XX & "  " & "Y: " & YY
ProgressBar1.Value = ProgressBar1.Value + 1
End Sub

Function Wait(Sekunden As Double)
Dim i As Long
For i = 1 To Sekunden
'Sleep 2
Waitsec (1)
'DoEvents
Next i
End Function


Private Sub Command1_Click()
On Error Resume Next
Kill MacroFile
Länge = "0"
Timer1.Enabled = True
Timer2.Enabled = True
End Sub

Private Sub Command2_Click()
ProgressBar1.Value = "0"
Mousebutton.Caption = ""
lblstatus.Caption = "Play"
DoEvents
Call pickout
End Sub

Private Sub Command3_Click()
MsgBox "Wood - MouseMacro" & vbNewLine & vbNewLine & "Written by: Klemens FRIEDL 02/2003" & vbNewLine & "Some Codes from: Marcin Bedner & Dodge (do mouseclicks), http://www.activevb.de (show mouseclicks)", vbInformation, "Info"
End Sub

Private Sub Command4_Click()
Timer1.Enabled = False
Timer2.Enabled = False
End Sub

Private Sub Form_Activate()
Slider1.Value = 2
End Sub

Private Sub Form_Load()
MacroFile = "macro.mca"
End Sub


Private Sub Timer1_Timer()
Dim tastentemp1 As Variant
'Record

If GetAsyncKeyState(VK_LBUTTON) Then
    Mousebutton.Caption = "left Mouse-Button"
    tastentemp1 = "L"
ElseIf GetAsyncKeyState(VK_RBUTTON) Then
    Mousebutton.Caption = "right Mouse-Button"
    tastentemp1 = "R"
ElseIf GetAsyncKeyState(VK_MBUTTON) Then
    Mousebutton.Caption = "middle Mouse-Button"
    tastentemp1 = "M"
Else
    Mousebutton.Caption = ""
    tastentemp1 = "N"
End If



On Error Resume Next
Dim Result&, P As POINTAPI
    Result = GetCursorPos(P)
    Me.Caption = "X: " & P.X & "     Y: " & P.Y
Open MacroFile For Append As #1
Print #1, "#" & P.X & "#" & "#" & P.Y & "#" & "#     #" & " +" & tastentemp1
'DoEvents
Close #1
End Sub

Private Sub Timer2_Timer()
Länge = Länge + 1
lblstatus.Caption = "Record : " & Länge & " sec"
End Sub


Private Function Waitsec(ByVal TimeToWait As Long)  'Time In seconds
    Dim EndTime As Long
    EndTime = GetTickCount + TimeToWait


    Do Until GetTickCount > EndTime
        DoEvents
        Loop
End Function

Private Sub TimeOut(MS)

  Dim start As Variant

    start = GetTickCount
    While start + MS > GetTickCount
        DoEvents
    Wend

End Sub

