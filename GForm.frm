VERSION 5.00
Begin VB.Form GForm 
   BorderStyle     =   0  'None
   Caption         =   "Warped By Simon Price"
   ClientHeight    =   4932
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   6144
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   411
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   512
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox PB 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      ClipControls    =   0   'False
      FillColor       =   &H0000FF00&
      ForeColor       =   &H000000FF&
      Height          =   3372
      Left            =   120
      ScaleHeight     =   281
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   331
      TabIndex        =   0
      Top             =   120
      Visible         =   0   'False
      Width           =   3972
   End
End
Attribute VB_Name = "GForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
Select Case KeyCode
  Case vbKeyC
    PB.Picture = PB.Image
    SavePicture PB.Picture, App.Path & "\WarpedPic.bmp"
    PB.Picture = Nothing
  Case vbKeyEscape
    End
End Select
End Sub

Private Sub Form_Load()
SortLayout
Show
BuildTrigTable
CreateTunnel
Mainloop
End Sub

Private Sub SortLayout()
Move 0, 0, Screen.Width, Screen.Height
PB.Move 0, 0, 600, 450
DispWidth = Screen.Height * 1.3333 / Screen.TwipsPerPixelY
DispHeight = Screen.Height / Screen.TwipsPerPixelY
End Sub

Private Sub CreateTunnel()
For i = 1 To NUMHOLES
   Hole(i).Color = QBColor(Int(Rnd * 14) + 1)
   Hole(i).x = Sine(Int(i * 3.59))
   Hole(i).y = Cosine(Int(i * 3.59))
   Hole(i).z = i
Next
End Sub

Public Sub Mainloop()
Do
DoEvents
PB.Cls
For i = 1 To NUMHOLES
  z = Hole(i).z - SPEED
  If z <= 0 Then
    Hole(i).z = VIEWDEPTH 'place back at redraw distance
  Else
    Hole(i).z = z 'move section of tunnel
  End If
  
  LensDivDist = Hole(i).z
  x = CX + Hole(i).x * LensDivDist
  y = CY - Hole(i).y * LensDivDist
  R = 3 * (VIEWDEPTH - Hole(i).z)
  PB.Circle (x, y), R, Hole(i).Color
  
  Select Case Abs(LastR - R)
    Case Is > 50
    GoTo Next1
  End Select
  PB.ForeColor = Hole(i).Color
  MoveToEx PB.hdc, Lastx + LastR, Lasty + LastR, lpPoint
  LineTo PB.hdc, x + R, y + R
  MoveToEx PB.hdc, Lastx - LastR, Lasty + LastR, lpPoint
  LineTo PB.hdc, x - R, y + R
  MoveToEx PB.hdc, Lastx + LastR, Lasty - LastR, lpPoint
  LineTo PB.hdc, x + R, y - R
  MoveToEx PB.hdc, Lastx - LastR, Lasty - LastR, lpPoint
  LineTo PB.hdc, x - R, y - R
Next1:
  Lastx = x
  Lasty = y
  LastR = R
Next
StretchBlt hdc, 0, 0, DispWidth, DispHeight, PB.hdc, 50, 50, PBWIDTH - 100, PBHEIGHT - 100, vbSrcCopy
Loop
End Sub
