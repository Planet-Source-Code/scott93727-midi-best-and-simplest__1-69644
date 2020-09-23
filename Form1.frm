VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Midi 128 Instrument 100 Note Demo"
   ClientHeight    =   2025
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   5625
   LinkTopic       =   "Form1"
   ScaleHeight     =   2025
   ScaleWidth      =   5625
   StartUpPosition =   1  'CenterOwner
   Begin VB.Timer tmrwelcome 
      Enabled         =   0   'False
      Interval        =   200
      Left            =   5040
      Top             =   1080
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Play Song"
      Height          =   255
      Left            =   3720
      TabIndex        =   12
      Top             =   1080
      Width           =   1215
   End
   Begin VB.Timer tmrrec 
      Left            =   5040
      Top             =   1560
   End
   Begin VB.TextBox Text3 
      Alignment       =   2  'Center
      Height          =   285
      Left            =   2280
      TabIndex        =   9
      Top             =   1320
      Width           =   1095
   End
   Begin VB.HScrollBar sldVol 
      Height          =   255
      Left            =   240
      Max             =   180
      TabIndex        =   7
      Top             =   1320
      Width           =   1695
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Stop"
      Height          =   255
      Left            =   3720
      TabIndex        =   6
      Top             =   720
      Width           =   1215
   End
   Begin VB.HScrollBar note 
      Height          =   255
      Left            =   240
      Max             =   100
      TabIndex        =   5
      Top             =   840
      Width           =   1695
   End
   Begin VB.HScrollBar sldInst 
      Height          =   255
      LargeChange     =   5
      Left            =   240
      Max             =   126
      TabIndex        =   4
      Top             =   360
      Width           =   1695
   End
   Begin VB.CommandButton Command4 
      Caption         =   "Exit"
      Height          =   255
      Left            =   3720
      TabIndex        =   3
      Top             =   1440
      Width           =   1215
   End
   Begin VB.TextBox Text2 
      Alignment       =   2  'Center
      Height          =   285
      Left            =   2280
      TabIndex        =   2
      Top             =   840
      Width           =   1095
   End
   Begin VB.TextBox Text1 
      Alignment       =   2  'Center
      Height          =   285
      Left            =   2280
      TabIndex        =   1
      Top             =   360
      Width           =   1095
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Play"
      Height          =   255
      Left            =   3720
      TabIndex        =   0
      Top             =   360
      Width           =   1215
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      Caption         =   "Instrument"
      Height          =   255
      Left            =   360
      TabIndex        =   11
      Top             =   120
      Width           =   1695
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      Caption         =   "Note"
      Height          =   255
      Left            =   360
      TabIndex        =   10
      Top             =   600
      Width           =   1695
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "Volume"
      Height          =   255
      Left            =   360
      TabIndex        =   8
      Top             =   1080
      Width           =   1695
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'' P I A N O  by Armin Niki
'' Original code:
'' http://www.planet-source-code.com/vb/scripts/ShowCode.asp?txtCodeId=64928&lngWId=1
'' Update - Apr 16 2006 by Paul Bahlawan
'' modified for min code Nov 2007 by Scott Smith
Option Explicit
Private Declare Function midiOutClose Lib "winmm.dll" (ByVal hMidiOut As Long) As Long
Private Declare Function midiOutOpen Lib "winmm.dll" (lphMidiOut As Long, ByVal uDeviceID As Long, ByVal dwCallback As Long, ByVal dwInstance As Long, ByVal dwFlags As Long) As Long
Private Declare Function midiOutShortMsg Lib "winmm.dll" (ByVal hMidiOut As Long, ByVal dwMsg As Long) As Long
Private hmidi As Long
Private baseNote As Long
Private channel As Long
Private volume As Long
Private lNote As Long
Private Playin() As String
Private playinc As Long
Private timers As Long
Private rec As String
Dim midimsg As Long
Dim notep As Long
Private Sub Command1_Click()
 domusic notep
End Sub
Private Sub Command2_Click()
domusicstop notep
End Sub
Private Sub tmrWelcome_Timer() 'plays song
Static pdemo As Long
    domusicstop pdemo - 7
    If pdemo > 64 Then
        pdemo = 0
        tmrwelcome.Enabled = False
        Exit Sub
    End If
    domusic pdemo + 5
    pdemo = pdemo + 12
   End Sub
Private Sub Command3_Click()
tmrwelcome.Enabled = True
End Sub
 Private Sub Command4_Click()
midiOutClose (hmidi)
Unload Me
End Sub
Private Sub domusic(mNote As Long)
Dim midimsg As Long
    'Play note
    midimsg = &H90 + ((baseNote + mNote) * &H100) + (volume * &H10000) + channel
    midiOutShortMsg hmidi, midimsg
    'record the key-down event
    If tmrrec.Enabled Then rec = rec & mNote & "x" & timers & " "
    timers = 0
    lNote = mNote
    'hi-light key being played
    'pKey(mNote - 1).BackColor = &H6060F0
End Sub
'Stop a note
Private Sub domusicstop(mNote As Long)
Dim midimsg As Long
    midimsg = &H80 + ((baseNote + mNote) * &H100) + channel
    midiOutShortMsg hmidi, midimsg
    'record the key-up event
    If tmrrec.Enabled Then rec = rec & -mNote & "x" & timers & " "
    timers = 0
    If mNote = lNote Then lNote = 0 'lNote = 0
    End Sub
Private Sub Form_Load()
Dim rc As Long
Dim curDevice As Long
    midiOutClose (hmidi)
    rc = midiOutOpen(hmidi, curDevice, 0, 0, 0)
    If (rc <> 0) Then
        MsgBox "Couldn't open midi device - Error #" & rc
    End If
    baseNote = 23
    channel = 15
    volume = 127
    sldVol.Value = 127
End Sub
Private Sub Form_Unload(Cancel As Integer)
    midiOutClose (hmidi)
End Sub
Private Sub note_Change()
notep = note.Value
Text2.Text = notep
End Sub
'Change the instrument
Private Sub sldInst_Change()
Dim midimsg As Long
    midimsg = (sldInst.Value * 256) + &HC0 + channel
    Text1.Text = midimsg
    midiOutShortMsg hmidi, midimsg
End Sub
Private Sub sldVol_Change()
    volume = sldVol.Value
    Text3.Text = sldVol.Value
End Sub


