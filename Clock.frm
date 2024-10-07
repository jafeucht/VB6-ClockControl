VERSION 5.00
Begin VB.Form frmClock 
   AutoRedraw      =   -1  'True
   Caption         =   "Form1"
   ClientHeight    =   3165
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4665
   LinkTopic       =   "Form1"
   ScaleHeight     =   3165
   ScaleWidth      =   4665
   StartUpPosition =   3  'Windows Default
   Begin VB.Timer tmTimer 
      Interval        =   1
      Left            =   75
      Top             =   75
   End
End
Attribute VB_Name = "frmClock"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Const PI = 3.14159265358979

Const MIN_COL = &HAFAFAF
Const HOUR_COL = &HCDCDCD

Private Type PointAPI
    X As Double
    Y As Double
End Type

Dim UseWidth As Boolean
Dim MinSize As Integer
Dim CenterPos As PointAPI
Dim HourHand As Double
Dim MinuteHand As Double
Dim SecondHand As Double
Dim StartTick As Double
Dim EndTick As Double

Private Sub Form_Click()
    PrintClock
End Sub

Private Sub Form_Resize()
    UseWidth = (ScaleWidth < ScaleHeight)
    CenterPos.X = ScaleWidth / 2
    CenterPos.Y = ScaleHeight / 2
    If UseWidth Then MinSize = ScaleWidth / 2 Else: MinSize = ScaleHeight / 2
    HourHand = MinSize * 0.6
    MinuteHand = MinSize * 0.8
    SecondHand = MinSize * 0.85
    StartTick = MinSize * 0.95
    EndTick = MinSize * 0.98
End Sub

Sub PrintClock()
Dim Hours As Double
Dim Minutes As Double
Dim Seconds As Double
    Hours = FMod(Timer / &HE10, &HC)
    Minutes = FMod(Timer / &H3C, &H3C)
    Seconds = FMod(Timer, &H3C)
    DrawHand HourHand, Hours / 12 * 360, RGB(175, 175, 175), 4
    DrawHand MinuteHand, Minutes / 60 * 360, RGB(125, 125, 125), 2
    DrawHand SecondHand, Seconds / 60 * 360, 255, 2
    PrintTics
    ForeColor = vbBlue
    Circle (CenterPos.X, CenterPos.Y), MinSize * 0.98
End Sub

Sub PrintTics()
Dim i As Integer
Dim StartPos As PointAPI, Endpos As PointAPI
    DrawWidth = 1
    For i = 1 To 60
        ForeColor = RGB(0, 255, 0): DrawWidth = 1
        If i Mod 5 = 0 Then ForeColor = RGB(255, 125, 0): DrawWidth = 2
        StartPos.X = CenterPos.X + Sine(i * 6) * StartTick
        StartPos.Y = CenterPos.Y - Cosine(i * 6) * StartTick
        Endpos.X = CenterPos.X + Sine(i * 6) * EndTick
        Endpos.Y = CenterPos.Y - Cosine(i * 6) * EndTick
        Line (Endpos.X, Endpos.Y)-(StartPos.X, StartPos.Y)
    Next i
End Sub

Sub DrawHand(Size As Double, Angle As Double, Color As Long, BorderWidth As Integer)
Dim EndPoint As PointAPI
    EndPoint.X = CenterPos.X + Sine(Angle) * Size
    EndPoint.Y = CenterPos.Y - Cosine(Angle) * Size
    DrawWidth = BorderWidth
    ForeColor = Color
    Line (EndPoint.X, EndPoint.Y)-(CenterPos.X, CenterPos.Y)
End Sub

Function FMod(Number As Double, Divisor As Double) As Double
Dim DivideProd As Double, Difference As Long
    DivideProd = Number / Divisor
    Difference = CInt(DivideProd) * Divisor
    FMod = Number - Difference
End Function

Function Sign(Number As Double) As Integer
    Sign = Abs(Number) / Number
End Function

Function Sine(ByVal i As Double) As Double
    Sine = Sin(i * (PI / 180))
End Function

Function Cosine(ByVal i As Double) As Double
    Cosine = Cos(i * (PI / 180))
End Function

Private Sub tmTimer_Timer()
    Cls
    PrintClock
End Sub

