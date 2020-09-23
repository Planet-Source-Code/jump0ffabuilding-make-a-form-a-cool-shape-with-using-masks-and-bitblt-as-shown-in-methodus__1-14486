VERSION 5.00
Begin VB.Form Form1 
   AutoRedraw      =   -1  'True
   BorderStyle     =   0  'None
   Caption         =   "Jump0ffabuilding is cool"
   ClientHeight    =   2640
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   2955
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "Form1.frx":0000
   ScaleHeight     =   2640
   ScaleWidth      =   2955
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command1 
      Caption         =   "Exit"
      Height          =   645
      Left            =   2055
      TabIndex        =   1
      Top             =   1050
      Width           =   690
   End
   Begin VB.PictureBox picBox 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      Height          =   2775
      Left            =   300
      Picture         =   "Form1.frx":9FCE
      ScaleHeight     =   2715
      ScaleWidth      =   3285
      TabIndex        =   0
      Top             =   1965
      Visible         =   0   'False
      Width           =   3345
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Declare Function CombineRgn Lib "gdi32" (ByVal hDestRgn As Long, ByVal hSrcRgn1 As Long, ByVal hSrcRgn2 As Long, ByVal nCombineMode As Long) As Long
Private Declare Function CreateRectRgn Lib "gdi32" (ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long) As Long
Private Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long
Private Declare Function GetPixel Lib "gdi32" (ByVal hdc As Long, ByVal X As Long, ByVal Y As Long) As Long
Private Declare Function ReleaseCapture Lib "user32" () As Long
Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Private Declare Function SetWindowRgn Lib "user32" (ByVal hWnd As Long, ByVal hRgn As Long, ByVal bRedraw As Boolean) As Long

Private Const RGN_OR = 2

Private Const WM_MOVE = &HF012
Private Const WM_SYSCOMMAND = &H112

Private lngRegion As Long
' this is the part that does all the crap to make this cool crap work
Private Function RegionFromBitmap(picSource As PictureBox, Optional lngTransColor As Long) As Long
  Dim lngRetr As Long, lngHeight As Long, lngWidth As Long
  Dim lngRgnFinal As Long, lngRgnTmp As Long
  Dim lngStart As Long, lngRow As Long
  Dim lngCol As Long
  If lngTransColor& < 1 Then 'if that crap is less than 1...
    lngTransColor& = GetPixel(picSource.hdc, 0, 0) 'sets lngTransColor& = to the the 0, 0 pixel
  End If
  lngHeight& = picSource.Height / Screen.TwipsPerPixelY 'get height crap
  lngWidth& = picSource.Width / Screen.TwipsPerPixelX ' get width crap
  lngRgnFinal& = CreateRectRgn(0, 0, 0, 0) 'do that crap
  For lngRow& = 0 To lngHeight& - 1
    lngCol& = 0
    Do While lngCol& < lngWidth&
      Do While lngCol& < lngWidth& And GetPixel(picSource.hdc, lngCol&, lngRow&) = lngTransColor& 'does crap to each pixel
        lngCol& = lngCol& + 1 'goes to next pixel
      Loop
      If lngCol& < lngWidth& Then
        lngStart& = lngCol&
        Do While lngCol& < lngWidth& And GetPixel(picSource.hdc, lngCol&, lngRow&) <> lngTransColor& ' does crap to each pixel
          lngCol& = lngCol& + 1 'goes to next pixel
        Loop
        If lngCol& > lngWidth& Then lngCol& = lngWidth&
        lngRgnTmp& = CreateRectRgn(lngStart&, lngRow&, lngCol&, lngRow& + 1)
        lngRetr& = CombineRgn(lngRgnFinal&, lngRgnFinal&, lngRgnTmp&, RGN_OR) 'combines crap to make region transparent
        DeleteObject (lngRgnTmp&) 'empties stupid buffer
      End If
    Loop
  Next
  RegionFromBitmap& = lngRgnFinal& 'done with crap
End Function

Private Sub Command1_Click()
End
End Sub

Private Sub Form_Load()
  Dim lngRetr As Long
  lngRegion& = RegionFromBitmap(picBox) 'calls that crap
  lngRetr& = SetWindowRgn(Me.hWnd, lngRegion&, True) 'calls that crap
End Sub
Private Sub Form_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Call ReleaseCapture 'this crap is for moving
    Call SendMessage(Me.hWnd, WM_SYSCOMMAND, WM_MOVE, 0) 'more moving crap
End Sub

