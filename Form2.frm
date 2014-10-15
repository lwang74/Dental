VERSION 5.00
Begin VB.Form Form2 
   Caption         =   "Form2"
   ClientHeight    =   6675
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   8400
   LinkTopic       =   "Form2"
   ScaleHeight     =   6675
   ScaleWidth      =   8400
   StartUpPosition =   3  '窗口缺省
   Begin VB.PictureBox Picture1 
      Height          =   4335
      Left            =   840
      ScaleHeight     =   4275
      ScaleWidth      =   5595
      TabIndex        =   2
      Top             =   600
      Width           =   5655
   End
   Begin VB.HScrollBar HScroll1 
      Height          =   495
      Left            =   720
      TabIndex        =   1
      Top             =   5520
      Width           =   6135
   End
   Begin VB.VScrollBar VScroll1 
      Height          =   6015
      Left            =   7680
      TabIndex        =   0
      Top             =   240
      Width           =   375
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Hvalue As Integer
Dim Vvalue As Integer
Dim WidValue As Integer
Dim HeiValue As Integer

Private Sub Form_Load()
WindowState = 2
'Picture1.AutoSize = True
'Picture1.Picture = LoadPicture("D:\Private\外网网站\图片\两条线照片\DSC00220.JPG")
End Sub

Private Sub Form_Resize()
WidValue = ScaleWidth - VScroll1.Width
HeiValue = ScaleHeight - HScroll1.Height
Picture1.Top = 0
Picture1.Left = 0
HScroll1.Top = HeiValue
HScroll1.Left = 0
HScroll1.Width = WidValue
VScroll1.Left = WidValue
VScroll1.Top = 0
VScroll1.Height = HeiValue + HScroll1.Height
HScroll1.Max = Picture1.Width - WidValue
VScroll1.Max = Picture1.Height - HeiValue
HScroll1.Min = 0
VScroll1.Min = 0
VScroll1.LargeChange = 100
VScroll1.SmallChange = 5
HScroll1.LargeChange = 100
HScroll1.SmallChange = 5
End Sub

Private Sub VScroll1_Change()
Vvalue = VScroll1.Value
Picture1.Move -Hvalue, -Vvalue
End Sub
