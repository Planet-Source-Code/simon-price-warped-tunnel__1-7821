Attribute VB_Name = "WarpedMod"
Public Type POINTAPI
  x As Integer
  y As Integer
End Type

Public Declare Function BitBlt Lib "gdi32" (ByVal hDestDC As Long, ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long) As Long
Public Declare Function StretchBlt Lib "gdi32" (ByVal hdc As Long, ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal nSrcWidth As Long, ByVal nSrcHeight As Long, ByVal dwRop As Long) As Long
Public Declare Function LineTo Lib "gdi32" (ByVal hdc As Long, ByVal x As Long, ByVal y As Long) As Long
Public Declare Function MoveToEx Lib "gdi32" (ByVal hdc As Long, ByVal x As Long, ByVal y As Long, lpPoint As POINTAPI) As Long
Public Declare Function SetPixel Lib "gdi32" (ByVal hdc As Long, ByVal x As Long, ByVal y As Long, ByVal crColor As Long) As Long
Public Declare Function GetPixel Lib "gdi32" (ByVal hdc As Long, ByVal x As Long, ByVal y As Long) As Long

Public Type t3DVector
  x As Single
  y As Single
  z As Single
End Type

Public Type tHole
  x As Single
  y As Single
  z As Single
  Color As Long
End Type

Public Const SPEED = 1

Public Const NUMHOLES = 100
Public Hole(1 To NUMHOLES) As tHole

Public Sine(0 To 359) As Single
Public Cosine(0 To 359) As Single

Public DispWidth As Integer
Public DispHeight As Integer

Public i As Integer
Public i2 As Integer
Public i3 As Integer
Public x As Integer
Public y As Integer
Public z As Integer
Public R As Integer
Public Lastx As Integer
Public Lasty As Integer
Public LastR As Integer

Public Const VIEWDEPTH = 100
Public Const LENS = VIEWDEPTH
Public LensDivDist As Single

Public Const PBWIDTH = 600
Public Const PBHEIGHT = PBWIDTH * 0.75
Public Const CX = PBWIDTH \ 2
Public Const CY = PBHEIGHT \ 2

Public lpPoint As POINTAPI

Public Const PI = 3.14159265358979 'obvious
Public Const PIdiv180 = PI / 180

Public Sub BuildTrigTable()
For i = 0 To 359
  Sine(i) = Sin(i * PIdiv180)
  Cosine(i) = Cos(i * PIdiv180)
Next
End Sub

