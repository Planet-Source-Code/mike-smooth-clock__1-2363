VERSION 5.00
Begin VB.Form Form1 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00C0C0C0&
   Caption         =   "Analog Clock"
   ClientHeight    =   4500
   ClientLeft      =   165
   ClientTop       =   735
   ClientWidth     =   6225
   FillColor       =   &H00FFFFFF&
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   12
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MousePointer    =   2  'Cross
   ScaleHeight     =   300
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   415
   StartUpPosition =   3  'Windows Default
   Begin VB.Menu mnuClock 
      Caption         =   "Clock Type"
      Begin VB.Menu mnuClockType 
         Caption         =   "Lines"
         Index           =   0
      End
      Begin VB.Menu mnuClockType 
         Caption         =   "Polygons"
         Index           =   1
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Type POINTAPI
  X As Long
  Y As Long
End Type

Dim sec_hand() As POINTAPI
Dim min_hand() As POINTAPI
Dim hour_hand() As POINTAPI
Dim tick_mrk() As POINTAPI

Private Declare Function Polygon Lib "gdi32" _
  (ByVal hdc As Long, lpPoint As POINTAPI, ByVal nCount As Long) As Long

Private Declare Function Polyline Lib "gdi32" _
  (ByVal hdc As Long, lpPoint As POINTAPI, ByVal nCount As Long) As Long


Dim hl As Integer, ml As Integer, sl As Integer
Dim cx As Integer, cy As Integer, r As Single, clocktype As Integer
Const vbDkGray = &H808080 'for the shadow
Const vbPI = 3.141592654
Const Deg2Rad = vbPI / 180 'Degrees to Radians



Sub Print_Num(num As Integer)
  strg = CStr(num)
  X = (TextWidth(strg) / 2)
  Y = (TextHeight(strg) / 2)
  offset = TextWidth("00") / 1.3
  CurrentX = Sin((180 - num * 30) * Deg2Rad) * (r - offset) + cx - (TextWidth(strg) / 2)
  CurrentY = Cos((180 - num * 30) * Deg2Rad) * (r - offset) + cy - (TextHeight(strg) / 2)
  Print strg

End Sub

Private Sub Form_Activate()
  Do
    If clocktype = 0 Then
      Time_1
    Else
      Time_2
    End If
    DoEvents
  Loop
End Sub


Private Sub Transform(tmp() As POINTAPI, pnts() As POINTAPI, A As Double, X As Double, Y As Double)
  Dim cos_ As Single, sin_ As Single
  A = A * Deg2Rad 'convert the degrees passed to radians
  p = UBound(pnts)
  ReDim tmp(p)
  
  cos_ = Cos(A) 'No use calculating Cos() and Sin() repeatedly
  sin_ = Sin(A) 'for the same angle 'specially if you have a large array of points
  
  'Now rotate hand to point in the proper direction
  'and translate the hands coordinates to move the
  'it to the center of the form
  For i = 0 To p
    tmp(i).X = pnts(i).X * cos_ - pnts(i).Y * sin_ + X
    tmp(i).Y = pnts(i).X * sin_ + pnts(i).Y * cos_ + Y
  Next
  'I put the transfofrmed coordinates in to a seperate array
  'because I do not want to alter the master copy of the hand coord's
End Sub

Private Sub lblpoints()
  For X = 0 To UBound(sec_hnd)
    PSet (sec_hnd(X).X, sec_hnd(X).Y)
    Print X
  Next
End Sub

Private Sub Form_DblClick()
  Cls
  DoEvents
  Time_2
End Sub

Private Sub Form_Load()
  Caption = "Smooth Clock!........"
  mnuClockType_Click 1
  
  
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
  If Button = vbLeftButton Then
    Caption = X & ", " & Y
  End If
End Sub


Private Sub Form_Resize()
Cls
'do these calculations here as they only
'change when the form is resized

  'find the center of the form
  cx = ScaleWidth / 2
  cy = ScaleHeight / 2
  
  'set the r (radius) to the smaller of the two cx & cy
  r = IIf(cx > cy, cy, cx) - 5
  PSet (cx, cy)
  'sel the lengths for the hands
  hl = r * 0.5  'hour
  ml = r * 0.25 'minute
  sl = r * 0.15 'second
  Load_Hands

End Sub
Private Sub Time_1()

  Dim t  As Double, H As Double, m As Double, s As Double
  
  '** here is what makes it smooth **
  'using the timer value
  t = Timer
  H = Abs(t / 3600)                             'get the hours
  m = Abs((t - (Fix(H) * 3600)) / 60)           'and minutes
  s = Abs(t - (Fix(H) * 3600) - (Fix(m) * 60))  'and seconds
  H = IIf(H >= 13, H - 12, H)                   'fix the hours to 12 hour intervals
  Cls
  Clock_Face
  
'draw the hour hand
  sin_ = Sin((180 - H * 30) * Deg2Rad) * (r - hl) + cx
  cos_ = Cos((180 - H * 30) * Deg2Rad) * (r - hl) + cy
  Draw_Line cx, cy, sin_, cos_, 3

'draw the minute hand
  DrawWidth = 2
  sin_ = Sin((180 - m * 6) * Deg2Rad) * (r - ml) + cx
  cos_ = Cos((180 - m * 6) * Deg2Rad) * (r - ml) + cy
  Draw_Line cx, cy, sin_, cos_, 2

'draw the second hand
  DrawWidth = 1
  sin_ = Sin((180 - s * 6) * Deg2Rad) * (r - sl) + cx
  cos_ = Cos((180 - s * 6) * Deg2Rad) * (r - sl) + cy
  Draw_Line cx, cy, sin_, cos_, 1, vbRed
  

End Sub
Sub Draw_Line(X1, Y1, X2, Y2, dw, Optional color)
  If IsMissing(color) Then color = vbBlack
  oldcolor = ForeColor
  ForeColor = color
  DrawWidth = dw
  Line (X1, Y1)-(X2, Y2)
  ForeColor = oldcolor
End Sub
Private Sub Time_2()
  Dim t  As Double, H As Double, m As Double, s As Double
  Dim tmp() As POINTAPI
  
  '** here is what makes it smooth **
  'using the timer value
  t = Timer
  H = Abs(t / 3600)                             'get the hours
  m = Abs((t - (Fix(H) * 3600)) / 60)           'and minutes
  s = Abs(t - (Fix(H) * 3600) - (Fix(m) * 60))  'and seconds
  H = IIf(H >= 13, H - 12, H)                   'fix the hours to 12 hour intervals
  Cls
  
  Clock_Face

  'draw the hour hand
  Form1.DrawWidth = 1
  FillStyle = vbSolid
  ForeColor = BackColor
  FillColor = vbBlack
  Transform tmp, hour_hand, H * 30, CSng(cx), CSng(cy)
  Polygon Me.hdc, tmp(0), UBound(hour_hand) + 1 'draw the polygon
  
  'draw the minute hand
  ForeColor = BackColor
  FillColor = vbBlack
  Transform tmp, min_hand, m * 6, CSng(cx), CSng(cy)
  Polygon Me.hdc, tmp(0), UBound(min_hand) + 1
  
  'draw the second hand
  ForeColor = BackColor
  FillColor = vbRed
  Transform tmp, sec_hand, s * 6, CSng(cx), CSng(cy)
  Polygon Me.hdc, tmp(0), UBound(sec_hand) + 1
  
End Sub
Sub Clock_Face()
  Dim tmp() As POINTAPI
  
  'draw the face
  Form1.DrawWidth = 1
  FillStyle = vbTransparent
  FillColor = BackColor
  Form1.DrawWidth = 3
  Circle (cx, cy), r, vbBlack   'outer circle
  Form1.DrawWidth = 1
  FillStyle = vbSolid
  ForeColor = vbRed
  FillColor = vbBlack
  Circle (cx, cy), 5, vbBlack   'center of clock
  
  'print the numerals
  ForeColor = vbBlack
  For X = 3 To 12 Step 3  'only 12, 3, 6 and 9
    Print_Num CInt(X)
  Next
  
  'draw the tic marks
  DrawWidth = 2
  ForeColor = vbBlack
  For X = 1 To 8  ' in locations other than 12, 3, 6 and 9
    Y = Choose(X, 5, 10, 20, 25, 35, 40, 50, 55) 'at these minutes
    Transform tmp, tick_mrk, CDbl(Y * 6), CSng(cx), CSng(cy)
    Polyline Me.hdc, tmp(0), UBound(tmp) + 1
  Next
End Sub

Private Sub Form_Unload(Cancel As Integer)
  End
End Sub

Private Sub mnuClockType_Click(Index As Integer)
  Static oldindex
  clocktype = Index
  mnuClockType(oldindex).Checked = False
  mnuClockType(Index).Checked = True
  
  oldindex = Index
End Sub
Sub Load_Hands()

  ReDim sec_hand(4)
  sec_hand(0).X = 0
  sec_hand(0).Y = -25
  sec_hand(1).X = -4
  sec_hand(1).Y = -30
  sec_hand(2).X = -1.5
  sec_hand(2).Y = -(r * 0.85)
  sec_hand(3).X = 1.5
  sec_hand(3).Y = -(r * 0.85)
  sec_hand(4).X = 4
  sec_hand(4).Y = -30

  ReDim min_hand(4)
  min_hand(0).X = 0
  min_hand(0).Y = -25
  min_hand(1).X = -6
  min_hand(1).Y = -30
  min_hand(2).X = -1.5
  min_hand(2).Y = -(r * 0.75)
  min_hand(3).X = 1.5
  min_hand(3).Y = -(r * 0.75)
  min_hand(4).X = 6
  min_hand(4).Y = -30

  ReDim hour_hand(4)
  hour_hand(0).X = 0
  hour_hand(0).Y = -25
  hour_hand(1).X = -10
  hour_hand(1).Y = -30
  hour_hand(2).X = -1.5
  hour_hand(2).Y = -(r * 0.65)
  hour_hand(3).X = 1.5
  hour_hand(3).Y = -(r * 0.65)
  hour_hand(4).X = 10
  hour_hand(4).Y = -30

  ReDim tick_mrk(1)
  tick_mrk(0).X = 0
  tick_mrk(0).Y = -(r - sl)
  tick_mrk(1).X = 0
  tick_mrk(1).Y = -r

End Sub
