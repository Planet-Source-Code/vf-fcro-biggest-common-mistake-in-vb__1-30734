VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "MS Little Mistake,NO API DECLARES!"
   ClientHeight    =   6855
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   7965
   LinkTopic       =   "Form1"
   ScaleHeight     =   6855
   ScaleWidth      =   7965
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox Text1 
      Height          =   3975
      Left            =   240
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Both
      TabIndex        =   6
      Top             =   2760
      Width           =   7695
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Add 10000 Strings [BUFFERING]"
      Height          =   615
      Index           =   1
      Left            =   3840
      TabIndex        =   1
      Top             =   1080
      Width           =   2895
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Add 10000 Strings [NORMAL]"
      Height          =   615
      Index           =   0
      Left            =   600
      TabIndex        =   0
      Top             =   1080
      Width           =   2895
   End
   Begin VB.Label Label3 
      BackColor       =   &H0000C000&
      Height          =   375
      Index           =   1
      Left            =   2040
      TabIndex        =   5
      Top             =   2280
      Width           =   3255
   End
   Begin VB.Label Label3 
      BackColor       =   &H0000FFFF&
      Height          =   375
      Index           =   0
      Left            =   2040
      TabIndex        =   4
      Top             =   1800
      Width           =   3255
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackColor       =   &H00C0FFFF&
      Caption         =   "Add 10000 times a string to itself,and tell us about time what we spend!              Start with NORMAL"
      Height          =   375
      Left            =   840
      TabIndex        =   3
      Top             =   600
      Width           =   5895
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H000000FF&
      Caption         =   $"Form1.frx":0000
      Height          =   375
      Left            =   840
      TabIndex        =   2
      Top             =   120
      Width           =   5895
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click(Index As Integer)
Dim a As String
Dim b As String
a = "COMM/MIST"
b = ""
Text1 = ""

Select Case Index
Case 0
Label3(0) = "Start:" & Time
DoEvents
For u = 1 To 10000
b = b & a
Next u
Label3(1) = "End:" & Time
Text1 = b
Case 1
Label3(0) = "Start:" & Time
DoEvents
Dim x As Long
x = 1
b = Space(Len(a) * 10000)
For u = 1 To 10000
Mid(b, x, Len(a)) = a
x = x + Len(a)
Next u
Label3(1) = "End:" & Time
Text1 = b
End Select

End Sub


Private Sub Form_Load()
Top = (Screen.Height - Height) / 2
Left = (Screen.Width - Width) / 2
End Sub
