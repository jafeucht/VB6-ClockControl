VERSION 5.00
Begin VB.UserControl Clock 
   AutoRedraw      =   -1  'True
   ClientHeight    =   915
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   1290
   DrawWidth       =   2
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   8.25
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ScaleHeight     =   915
   ScaleWidth      =   1290
End
Attribute VB_Name = "Clock"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Const PI = 3.14159265358979
Const BoldLetters = 1400
Const MinLetters = 900

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
Dim Hour As Double
Dim Minute As Double
Dim Second As Double
Dim HColor As Long
Dim MColor As Long
Dim SColor As Long
Dim T1Color As Long
Dim T2Color As Long
Dim BColor As Long
Dim Label As Long
Dim CTime As Date
Dim ShwLetters As Boolean

Property Get ShowLetters() As Boolean
    ShowLetters = ShwLetters
End Property

Property Let ShowLetters(NewValue As Boolean)
    ShwLetters = NewValue
    UserControl_Resize
    PrintClock
End Property

Property Get HourColor() As OLE_COLOR
    HourColor = HColor
End Property

Property Let HourColor(NewColor As OLE_COLOR)
    HColor = NewColor
    PropertyChanged "HourColor"
    PrintIfAmbient
End Property

Property Get MinuteColor() As OLE_COLOR
    MinuteColor = MColor
End Property

Property Let MinuteColor(NewColor As OLE_COLOR)
    MColor = NewColor
    PropertyChanged "MinuteColor"
    PrintIfAmbient
End Property

Property Get SecondColor() As OLE_COLOR
    SecondColor = SColor
End Property

Property Let SecondColor(NewColor As OLE_COLOR)
    SColor = NewColor
    PropertyChanged "SecondColor"
    PrintIfAmbient
End Property

Property Get Hours() As Double
    Hours = Hour
End Property

Property Let Hours(NewHours As Double)
    Hour = NewHours
    PropertyChanged "Hours"
    PrintIfAmbient
End Property

Property Get Minutes() As Double
    Minutes = Minute
End Property

Property Let Minutes(NewMinutes As Double)
    Minute = NewMinutes
    PropertyChanged "Minutes"
    PrintIfAmbient
End Property

Property Get Seconds() As Double
    Seconds = Second
End Property

Property Let Seconds(NewSeconds As Double)
    Second = NewSeconds
    PropertyChanged "Seconds"
    PrintIfAmbient
End Property

Property Get MinuteTickColor() As OLE_COLOR
    MinuteTickColor = T1Color
End Property

Property Let MinuteTickColor(NewColor As OLE_COLOR)
    T1Color = NewColor
    PropertyChanged "MinuteTickColor"
    PrintIfAmbient
End Property

Property Get HourTickColor() As OLE_COLOR
    HourTickColor = T2Color
End Property

Property Let HourTickColor(NewColor As OLE_COLOR)
    T2Color = NewColor
    PropertyChanged "HourTickColor"
    PrintIfAmbient
End Property

Property Get BorderColor() As OLE_COLOR
    BorderColor = BColor
End Property

Property Let BorderColor(NewColor As OLE_COLOR)
    BColor = NewColor
    PropertyChanged "BorderColor"
    PrintIfAmbient
End Property

Property Get BackColor() As OLE_COLOR
    BackColor = UserControl.BackColor
End Property

Property Let BackColor(NewColor As OLE_COLOR)
    UserControl.BackColor = NewColor
    PropertyChanged "BackColor"
    PrintIfAmbient
End Property

Property Get ClockTime() As Date
    On Error Resume Next
    ClockTime = CDate(Int(Hour) & ":" & Int(Minute) & ":" & Int(Second))
End Property

Property Let ClockTime(NewTime As Date)
    Hour = Int(Format(NewTime, "h"))
    Minute = Int(Format(NewTime, "n"))
    Second = Int(Format(NewTime, "s"))
End Property

Sub PrintIfAmbient()
    On Error Resume Next
    If Not UserControl.Ambient.UserMode Then PrintClock
End Sub

Private Sub UserControl_Initialize()
    HColor = 11513775
    MColor = 12434877
    SColor = 255
    GetTime
    T1Color = 16711680
    T2Color = 32000
    BColor = 8882055
    BackColor = -2147483633
    ShwLetters = True
End Sub

Public Sub Reset()
    ClockTime = #12:00:00 AM#
End Sub

Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
    HColor = PropBag.ReadProperty("HourColor", 11513775)
    MColor = PropBag.ReadProperty("MinuteColor", 12434877)
    SColor = PropBag.ReadProperty("Secondcolor", 255)
    Hours = PropBag.ReadProperty("Hours", CInt(Format(Time, "h")))
    Minutes = PropBag.ReadProperty("Minutes", CInt(Format(Time, "m")))
    Seconds = PropBag.ReadProperty("Seconds", CInt(Format(Time, "s")))
    T1Color = PropBag.ReadProperty("MinuteTickColor", 16711680)
    T2Color = PropBag.ReadProperty("HourTickColor", 32000)
    BColor = PropBag.ReadProperty("BorderColor", 0)
    BackColor = PropBag.ReadProperty("BackColor", -2147483633)
    ShwLetters = PropBag.ReadProperty("ShowLetters", True)
End Sub

Private Sub UserControl_Resize()
    UseWidth = (ScaleWidth < ScaleHeight)
    CenterPos.X = ScaleWidth / 2
    CenterPos.Y = ScaleHeight / 2
    If UseWidth Then MinSize = ScaleWidth / 2 Else: MinSize = ScaleHeight / 2
    Debug.Print MinSize
    If ShwLetters And Not MinSize < MinLetters Then
        HourHand = MinSize * 0.4
        MinuteHand = MinSize * 0.6
        SecondHand = MinSize * 0.65
        StartTick = MinSize * 0.72
        EndTick = MinSize * 0.77
        Label = MinSize * 0.167
    Else
        HourHand = MinSize * 0.6
        MinuteHand = MinSize * 0.8
        SecondHand = MinSize * 0.85
        StartTick = MinSize * 0.92
        EndTick = MinSize * 0.97
    End If
    PrintIfAmbient
End Sub

Public Sub GetTime()
    Hour = FMod(Timer / &HE10, &HC)
    Minute = FMod(Timer / &H3C, &H3C)
    Second = FMod(Timer, &H3C)
End Sub

Public Sub PrintClock()
    Cls
    DrawHand HourHand, Hour / 12 * 360, HourColor, 3
    DrawHand MinuteHand, Minute / 60 * 360, MinuteColor, 2
    DrawHand SecondHand, Second / 60 * 360, SecondColor, 2
    PrintTics
    DrawWidth = 2
    ForeColor = BorderColor
    Circle (CenterPos.X, CenterPos.Y), MinSize * 0.98
End Sub

Sub PrintTics()
Dim i As Integer
Dim StartPos As PointAPI, Endpos As PointAPI
    For i = 1 To 60
        ForeColor = T1Color: DrawWidth = 1
        StartPos.X = CenterPos.X + Sine(i * 6) * StartTick
        StartPos.Y = CenterPos.Y - Cosine(i * 6) * StartTick
        Endpos.X = CenterPos.X + Sine(i * 6) * EndTick
        Endpos.Y = CenterPos.Y - Cosine(i * 6) * EndTick
        If i Mod 5 = 0 Then
            UserControl.ForeColor = T2Color: DrawWidth = 2
            UserControl.Font.Size = 8
            CurrentX = StartPos.X + Sine(i * 6) * Label - TextWidth(i) / 2
            CurrentY = StartPos.Y - Cosine(i * 6) * Label - TextHeight(i) / 2
            Font.Bold = MinSize > BoldLetters
            If MinSize > MinLetters And ShwLetters Then Print i / 5
        End If
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

Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
    PropBag.WriteProperty "HourColor", HColor, 11513775
    PropBag.WriteProperty "MinuteColor", MColor, 12434877
    PropBag.WriteProperty "Secondcolor", SColor, 255
    PropBag.WriteProperty "Hours", Hours, CInt(Format(Time, "h"))
    PropBag.WriteProperty "Minutes", Minutes, CInt(Format(Time, "m"))
    PropBag.WriteProperty "Seconds", Seconds, CInt(Format(Time, "s"))
    PropBag.WriteProperty "MinuteTickColor", T1Color, 16711680
    PropBag.WriteProperty "HourTickColor", T2Color, 32000
    PropBag.WriteProperty "BorderColor", BColor, 0
    PropBag.WriteProperty "BackColor", BackColor, -2147483633
    PropBag.WriteProperty "ShowLetters", ShwLetters, True
End Sub
