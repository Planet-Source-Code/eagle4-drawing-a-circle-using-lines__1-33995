VERSION 5.00
Begin VB.Form Form1 
   AutoRedraw      =   -1  'True
   Caption         =   "Cool Circle"
   ClientHeight    =   5745
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   6825
   LinkTopic       =   "Form1"
   ScaleHeight     =   5745
   ScaleWidth      =   6825
   StartUpPosition =   3  'Windows Default
   Begin VB.Timer Timer4 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   6060
      Top             =   4470
   End
   Begin VB.Timer Timer3 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   6030
      Top             =   4020
   End
   Begin VB.Timer Timer2 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   6060
      Top             =   3570
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  'Right Justify
      Height          =   285
      Left            =   5805
      TabIndex        =   2
      Text            =   "100"
      Top             =   1140
      Width           =   975
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   5670
      Left            =   30
      ScaleHeight     =   5640
      ScaleWidth      =   5700
      TabIndex        =   1
      Top             =   60
      Width           =   5730
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   6075
      Top             =   3090
   End
   Begin VB.CommandButton Command1 
      Caption         =   "DRAW"
      Height          =   465
      Left            =   5850
      TabIndex        =   0
      Top             =   345
      Width           =   930
   End
   Begin VB.Label Label2 
      Height          =   495
      Left            =   5865
      TabIndex        =   4
      Top             =   1560
      Width           =   1260
   End
   Begin VB.Label Label1 
      Caption         =   "SPACING"
      Height          =   300
      Left            =   5940
      TabIndex        =   3
      Top             =   900
      Width           =   750
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'this code is copyrighted to David Whiddon (me)
'the use of this code, or any part of this code
'is prohibited unless an e-mail is sent to me at
'wedgetaileagle@yahoo.com

'i just decided one day to make a program that draws
'a circle using straight lines. i hope you like it
'Dave

Dim x As Integer 'x and y are variables for the 2 axies
Dim y As Integer
Private Sub Command1_Click()
Picture1.Cls
Timer1.Enabled = True
y = Picture1.Height / 2
x = 0
End Sub

Private Sub Timer1_Timer()
Label2.Caption = "DRAWING..." ' each timer draws the lines
x = x + Text1                 ' in each corner
y = y - Text1

Picture1.Line (0, x)-(y, 0)
If x >= Picture1.Height / 2 Then
    Timer1.Enabled = False
    x = Picture1.Height / 2
    y = 0
    Timer2.Enabled = True
    
End If

End Sub

Private Sub Timer2_Timer()
x = x + Text1
y = y + Text1
Picture1.Line (x, 0)-(Picture1.Width, y)
If x >= Picture1.Width Then
    Timer2.Enabled = False
    Timer3.Enabled = True
    y = Picture1.Height / 2
    x = 0
End If
End Sub

Private Sub Timer3_Timer()
x = x + Text1
y = y + Text1
Picture1.Line (x, Picture1.Height)-(0, y)
If x >= Picture1.Height Then
    Timer3.Enabled = False
    Timer4.Enabled = True
    y = (Picture1.Height / 2) - 100
    x = (Picture1.Width) - 50
End If
End Sub

Private Sub Timer4_Timer()
x = x - Text1
y = y + Text1
Picture1.Line (Picture1.Width, y)-(x, Picture1.Height)
If x <= Picture1.Width / 2 Then
    Timer4.Enabled = False
    Label2.Caption = "DONE"
End If
End Sub


