VERSION 5.00
Object = "*\AClockCtl.vbp"
Begin VB.Form frmClockTest 
   Caption         =   "Form1"
   ClientHeight    =   4020
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   5685
   LinkTopic       =   "Form1"
   ScaleHeight     =   4020
   ScaleWidth      =   5685
   StartUpPosition =   3  'Windows Default
   Begin Project1.Clock Clock1 
      Height          =   1515
      Left            =   750
      TabIndex        =   0
      Top             =   300
      Width           =   1440
      _ExtentX        =   2540
      _ExtentY        =   2672
      HourColor       =   -2147483628
      MinuteColor     =   65535
      Secondcolor     =   4210816
      Hours           =   0
      Minutes         =   0
      Seconds         =   0
      HourTickColor   =   -2147483631
      BorderColor     =   8882055
   End
   Begin VB.Timer tmTimer 
      Interval        =   1000
      Left            =   0
      Top             =   1425
   End
End
Attribute VB_Name = "frmClockTest"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Click()
Static Started As Boolean
    If Started Then Exit Sub
    Started = True
    Do
        Clock1.GetTime
        Clock1.PrintClock
        DoEvents
    Loop
    Started = False
End Sub
