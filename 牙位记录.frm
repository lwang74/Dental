VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form 牙位记录 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "牙位记录"
   ClientHeight    =   3555
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6510
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3555
   ScaleWidth      =   6510
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  '窗口缺省
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   495
      Left            =   360
      Top             =   480
      Visible         =   0   'False
      Width           =   1815
      _ExtentX        =   3201
      _ExtentY        =   873
      ConnectMode     =   0
      CursorLocation  =   3
      IsolationLevel  =   -1
      ConnectionTimeout=   15
      CommandTimeout  =   30
      CursorType      =   3
      LockType        =   3
      CommandType     =   8
      CursorOptions   =   0
      CacheSize       =   50
      MaxRecords      =   0
      BOFAction       =   0
      EOFAction       =   0
      ConnectStringType=   1
      Appearance      =   1
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      Orientation     =   0
      Enabled         =   -1
      Connect         =   ""
      OLEDBString     =   ""
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   ""
      Caption         =   "Adodc1"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "宋体"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _Version        =   393216
   End
   Begin VB.CommandButton Command2 
      Caption         =   "关 闭"
      BeginProperty Font 
         Name            =   "楷体_GB2312"
         Size            =   14.25
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   3720
      TabIndex        =   67
      Top             =   2760
      Width           =   1455
   End
   Begin VB.CommandButton Command1 
      Caption         =   "保 存"
      BeginProperty Font 
         Name            =   "楷体_GB2312"
         Size            =   14.25
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   1080
      TabIndex        =   66
      Top             =   2760
      Width           =   1455
   End
   Begin VB.TextBox Text1 
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   14.25
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   3480
      TabIndex        =   65
      Top             =   240
      Width           =   1575
   End
   Begin VB.CheckBox Check1 
      Caption         =   "Check1"
      Height          =   255
      Index           =   31
      Left            =   5100
      TabIndex        =   31
      Top             =   2040
      Width           =   255
   End
   Begin VB.CheckBox Check1 
      Caption         =   "Check1"
      Height          =   255
      Index           =   30
      Left            =   4860
      TabIndex        =   30
      Top             =   2040
      Width           =   255
   End
   Begin VB.CheckBox Check1 
      Caption         =   "Check1"
      Height          =   255
      Index           =   29
      Left            =   4620
      TabIndex        =   29
      Top             =   2040
      Width           =   255
   End
   Begin VB.CheckBox Check1 
      Caption         =   "Check1"
      Height          =   255
      Index           =   28
      Left            =   4380
      TabIndex        =   28
      Top             =   2040
      Width           =   255
   End
   Begin VB.CheckBox Check1 
      Caption         =   "Check1"
      Height          =   255
      Index           =   27
      Left            =   4140
      TabIndex        =   27
      Top             =   2040
      Width           =   255
   End
   Begin VB.CheckBox Check1 
      Caption         =   "Check1"
      Height          =   255
      Index           =   26
      Left            =   3900
      TabIndex        =   26
      Top             =   2040
      Width           =   255
   End
   Begin VB.CheckBox Check1 
      Caption         =   "Check1"
      Height          =   255
      Index           =   25
      Left            =   3660
      TabIndex        =   25
      Top             =   2040
      Width           =   255
   End
   Begin VB.CheckBox Check1 
      Caption         =   "Check1"
      Height          =   255
      Index           =   24
      Left            =   3420
      TabIndex        =   24
      Top             =   2040
      Width           =   255
   End
   Begin VB.CheckBox Check1 
      Caption         =   "Check1"
      Height          =   255
      Index           =   23
      Left            =   720
      TabIndex        =   23
      Top             =   2040
      Width           =   255
   End
   Begin VB.CheckBox Check1 
      Caption         =   "Check1"
      Height          =   255
      Index           =   22
      Left            =   1020
      TabIndex        =   22
      Top             =   2040
      Width           =   255
   End
   Begin VB.CheckBox Check1 
      Caption         =   "Check1"
      Height          =   255
      Index           =   21
      Left            =   1260
      TabIndex        =   21
      Top             =   2040
      Width           =   255
   End
   Begin VB.CheckBox Check1 
      Caption         =   "Check1"
      Height          =   255
      Index           =   20
      Left            =   1500
      TabIndex        =   20
      Top             =   2040
      Width           =   255
   End
   Begin VB.CheckBox Check1 
      Caption         =   "Check1"
      Height          =   255
      Index           =   19
      Left            =   1740
      TabIndex        =   19
      Top             =   2040
      Width           =   255
   End
   Begin VB.CheckBox Check1 
      Caption         =   "Check1"
      Height          =   255
      Index           =   18
      Left            =   1980
      TabIndex        =   18
      Top             =   2040
      Width           =   255
   End
   Begin VB.CheckBox Check1 
      Caption         =   "Check1"
      Height          =   255
      Index           =   17
      Left            =   2220
      TabIndex        =   17
      Top             =   2040
      Width           =   255
   End
   Begin VB.CheckBox Check1 
      Caption         =   "Check1"
      Height          =   255
      Index           =   16
      Left            =   2460
      TabIndex        =   16
      Top             =   2040
      Width           =   255
   End
   Begin VB.CheckBox Check1 
      Caption         =   "Check1"
      Height          =   255
      Index           =   15
      Left            =   5100
      TabIndex        =   15
      Top             =   1320
      Width           =   255
   End
   Begin VB.CheckBox Check1 
      Caption         =   "Check1"
      Height          =   255
      Index           =   14
      Left            =   4860
      TabIndex        =   14
      Top             =   1320
      Width           =   255
   End
   Begin VB.CheckBox Check1 
      Caption         =   "Check1"
      Height          =   255
      Index           =   13
      Left            =   4620
      TabIndex        =   13
      Top             =   1320
      Width           =   255
   End
   Begin VB.CheckBox Check1 
      Caption         =   "Check1"
      Height          =   255
      Index           =   12
      Left            =   4380
      TabIndex        =   12
      Top             =   1320
      Width           =   255
   End
   Begin VB.CheckBox Check1 
      Caption         =   "Check1"
      Height          =   255
      Index           =   11
      Left            =   4140
      TabIndex        =   11
      Top             =   1320
      Width           =   255
   End
   Begin VB.CheckBox Check1 
      Caption         =   "Check1"
      Height          =   255
      Index           =   10
      Left            =   3900
      TabIndex        =   10
      Top             =   1320
      Width           =   255
   End
   Begin VB.CheckBox Check1 
      Caption         =   "Check1"
      Height          =   255
      Index           =   9
      Left            =   3660
      TabIndex        =   9
      Top             =   1320
      Width           =   255
   End
   Begin VB.CheckBox Check1 
      Caption         =   "Check1"
      Height          =   255
      Index           =   8
      Left            =   3420
      TabIndex        =   8
      Top             =   1320
      Width           =   255
   End
   Begin VB.CheckBox Check1 
      Caption         =   "Check1"
      Height          =   255
      Index           =   0
      Left            =   2460
      TabIndex        =   7
      Top             =   1320
      Width           =   255
   End
   Begin VB.CheckBox Check1 
      Caption         =   "Check1"
      Height          =   255
      Index           =   1
      Left            =   2220
      TabIndex        =   6
      Top             =   1320
      Width           =   255
   End
   Begin VB.CheckBox Check1 
      Caption         =   "Check1"
      Height          =   255
      Index           =   2
      Left            =   1980
      TabIndex        =   5
      Top             =   1320
      Width           =   255
   End
   Begin VB.CheckBox Check1 
      Caption         =   "Check1"
      Height          =   255
      Index           =   3
      Left            =   1740
      TabIndex        =   4
      Top             =   1320
      Width           =   255
   End
   Begin VB.CheckBox Check1 
      Caption         =   "Check1"
      Height          =   255
      Index           =   4
      Left            =   1530
      TabIndex        =   3
      Top             =   1320
      Width           =   195
   End
   Begin VB.CheckBox Check1 
      Caption         =   "Check1"
      Height          =   255
      Index           =   5
      Left            =   1260
      TabIndex        =   2
      Top             =   1320
      Width           =   255
   End
   Begin VB.CheckBox Check1 
      Caption         =   "Check1"
      Height          =   255
      Index           =   6
      Left            =   1020
      TabIndex        =   1
      Top             =   1320
      Width           =   255
   End
   Begin VB.CheckBox Check1 
      Caption         =   "Check1"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   10.5
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   7
      Left            =   720
      TabIndex        =   0
      Top             =   1320
      Width           =   255
   End
   Begin VB.Label Label2 
      Caption         =   "病历编号："
      BeginProperty Font 
         Name            =   "楷体_GB2312"
         Size            =   14.25
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1200
      TabIndex        =   64
      Top             =   360
      Width           =   1575
   End
   Begin VB.Label Label1 
      Caption         =   "8"
      Height          =   255
      Index           =   31
      Left            =   5160
      TabIndex        =   63
      Top             =   2400
      Width           =   135
   End
   Begin VB.Label Label1 
      Caption         =   "7"
      Height          =   255
      Index           =   30
      Left            =   4920
      TabIndex        =   62
      Top             =   2400
      Width           =   135
   End
   Begin VB.Label Label1 
      Caption         =   "5"
      Height          =   255
      Index           =   29
      Left            =   4440
      TabIndex        =   61
      Top             =   2400
      Width           =   135
   End
   Begin VB.Label Label1 
      Caption         =   "6"
      Height          =   255
      Index           =   28
      Left            =   4680
      TabIndex        =   60
      Top             =   2400
      Width           =   135
   End
   Begin VB.Label Label1 
      Caption         =   "1"
      Height          =   255
      Index           =   27
      Left            =   3480
      TabIndex        =   59
      Top             =   2400
      Width           =   135
   End
   Begin VB.Label Label1 
      Caption         =   "2"
      Height          =   255
      Index           =   26
      Left            =   3720
      TabIndex        =   58
      Top             =   2400
      Width           =   135
   End
   Begin VB.Label Label1 
      Caption         =   "4"
      Height          =   255
      Index           =   25
      Left            =   4200
      TabIndex        =   57
      Top             =   2400
      Width           =   135
   End
   Begin VB.Label Label1 
      Caption         =   "3"
      Height          =   255
      Index           =   24
      Left            =   3960
      TabIndex        =   56
      Top             =   2400
      Width           =   135
   End
   Begin VB.Label Label1 
      Caption         =   "8"
      Height          =   255
      Index           =   23
      Left            =   5160
      TabIndex        =   55
      Top             =   1080
      Width           =   135
   End
   Begin VB.Label Label1 
      Caption         =   "7"
      Height          =   255
      Index           =   22
      Left            =   4920
      TabIndex        =   54
      Top             =   1080
      Width           =   135
   End
   Begin VB.Label Label1 
      Caption         =   "5"
      Height          =   255
      Index           =   21
      Left            =   4440
      TabIndex        =   53
      Top             =   1080
      Width           =   135
   End
   Begin VB.Label Label1 
      Caption         =   "6"
      Height          =   255
      Index           =   20
      Left            =   4680
      TabIndex        =   52
      Top             =   1080
      Width           =   135
   End
   Begin VB.Label Label1 
      Caption         =   "1"
      Height          =   255
      Index           =   19
      Left            =   3480
      TabIndex        =   51
      Top             =   1080
      Width           =   135
   End
   Begin VB.Label Label1 
      Caption         =   "2"
      Height          =   255
      Index           =   18
      Left            =   3720
      TabIndex        =   50
      Top             =   1080
      Width           =   135
   End
   Begin VB.Label Label1 
      Caption         =   "4"
      Height          =   255
      Index           =   17
      Left            =   4200
      TabIndex        =   49
      Top             =   1080
      Width           =   135
   End
   Begin VB.Label Label1 
      Caption         =   "3"
      Height          =   255
      Index           =   16
      Left            =   3960
      TabIndex        =   48
      Top             =   1080
      Width           =   135
   End
   Begin VB.Label Label1 
      Caption         =   "8"
      Height          =   255
      Index           =   15
      Left            =   780
      TabIndex        =   47
      Top             =   2400
      Width           =   135
   End
   Begin VB.Label Label1 
      Caption         =   "7"
      Height          =   255
      Index           =   14
      Left            =   1080
      TabIndex        =   46
      Top             =   2400
      Width           =   135
   End
   Begin VB.Label Label1 
      Caption         =   "5"
      Height          =   255
      Index           =   13
      Left            =   1560
      TabIndex        =   45
      Top             =   2400
      Width           =   135
   End
   Begin VB.Label Label1 
      Caption         =   "6"
      Height          =   255
      Index           =   12
      Left            =   1320
      TabIndex        =   44
      Top             =   2400
      Width           =   135
   End
   Begin VB.Label Label1 
      Caption         =   "1"
      Height          =   255
      Index           =   11
      Left            =   2520
      TabIndex        =   43
      Top             =   2400
      Width           =   135
   End
   Begin VB.Label Label1 
      Caption         =   "2"
      Height          =   255
      Index           =   10
      Left            =   2280
      TabIndex        =   42
      Top             =   2400
      Width           =   135
   End
   Begin VB.Label Label1 
      Caption         =   "4"
      Height          =   255
      Index           =   9
      Left            =   1800
      TabIndex        =   41
      Top             =   2400
      Width           =   135
   End
   Begin VB.Label Label1 
      Caption         =   "3"
      Height          =   255
      Index           =   8
      Left            =   2040
      TabIndex        =   40
      Top             =   2400
      Width           =   135
   End
   Begin VB.Label Label1 
      Caption         =   "8"
      Height          =   255
      Index           =   7
      Left            =   780
      TabIndex        =   39
      Top             =   1080
      Width           =   135
   End
   Begin VB.Label Label1 
      Caption         =   "7"
      Height          =   255
      Index           =   6
      Left            =   1080
      TabIndex        =   38
      Top             =   1080
      Width           =   135
   End
   Begin VB.Label Label1 
      Caption         =   "5"
      Height          =   255
      Index           =   5
      Left            =   1560
      TabIndex        =   37
      Top             =   1080
      Width           =   135
   End
   Begin VB.Label Label1 
      Caption         =   "6"
      Height          =   255
      Index           =   4
      Left            =   1320
      TabIndex        =   36
      Top             =   1080
      Width           =   135
   End
   Begin VB.Label Label1 
      Caption         =   "1"
      Height          =   255
      Index           =   3
      Left            =   2520
      TabIndex        =   35
      Top             =   1080
      Width           =   135
   End
   Begin VB.Label Label1 
      Caption         =   "2"
      Height          =   255
      Index           =   2
      Left            =   2280
      TabIndex        =   34
      Top             =   1080
      Width           =   135
   End
   Begin VB.Label Label1 
      Caption         =   "4"
      Height          =   255
      Index           =   1
      Left            =   1800
      TabIndex        =   33
      Top             =   1080
      Width           =   135
   End
   Begin VB.Label Label1 
      Caption         =   "3"
      Height          =   255
      Index           =   0
      Left            =   2040
      TabIndex        =   32
      Top             =   1080
      Width           =   135
   End
   Begin VB.Line Line1 
      BorderWidth     =   3
      X1              =   120
      X2              =   6240
      Y1              =   1800
      Y2              =   1800
   End
   Begin VB.Line Line2 
      BorderWidth     =   3
      X1              =   3120
      X2              =   3120
      Y1              =   120
      Y2              =   3360
   End
End
Attribute VB_Name = "牙位记录"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Command1_Click()
Dim x As Integer
Dim i As Integer
sql5 = "select * from 牙位记录 where 病历编号 ='" & Trim(Text1.Text) & "'"
Adodc1.RecordSource = sql5
Adodc1.Refresh
If Adodc1.Recordset.RecordCount >= 1 Then
    x = MsgBox("此牙位记录已存在，是否需要覆盖？", 52, "提示")
Else
aaa:
    With Adodc1
        .Recordset.AddNew
        .Recordset("病历编号") = Text1.Text
        .Recordset("th1") = Check1(0).Value
        .Recordset("th2") = Check1(1).Value
        .Recordset("th3") = Check1(2).Value
        .Recordset("th4") = Check1(3).Value
        .Recordset("th5") = Check1(4).Value
        .Recordset("th6") = Check1(5).Value
        .Recordset("th7") = Check1(6).Value
        .Recordset("th8") = Check1(7).Value
        .Recordset("th9") = Check1(8).Value
        .Recordset("th10") = Check1(9).Value
        .Recordset("th11") = Check1(10).Value
        .Recordset("th12") = Check1(11).Value
        .Recordset("th13") = Check1(12).Value
        .Recordset("th14") = Check1(13).Value
        .Recordset("th15") = Check1(14).Value
        .Recordset("th16") = Check1(15).Value
        .Recordset("th17") = Check1(16).Value
        .Recordset("th18") = Check1(17).Value
        .Recordset("th19") = Check1(18).Value
        .Recordset("th20") = Check1(19).Value
        .Recordset("th21") = Check1(20).Value
        .Recordset("th22") = Check1(21).Value
        .Recordset("th23") = Check1(22).Value
        .Recordset("th24") = Check1(23).Value
        .Recordset("th25") = Check1(24).Value
        .Recordset("th26") = Check1(25).Value
        .Recordset("th27") = Check1(26).Value
        .Recordset("th28") = Check1(27).Value
        .Recordset("th29") = Check1(28).Value
        .Recordset("th30") = Check1(29).Value
        .Recordset("th31") = Check1(30).Value
        .Recordset("th32") = Check1(31).Value
    End With
    Adodc1.Recordset.Update
    Adodc1.Refresh
    MsgBox "保存完毕！", 48, "提示"
    Unload Me
    Exit Sub
End If
If x = 6 Then
    Adodc1.Recordset.Delete
    Adodc1.Recordset.Update
    Adodc1.Refresh
    GoTo aaa
    Exit Sub
End If
End Sub

Private Sub Command2_Click()
Unload Me
End Sub

Private Sub Form_Load()
Adodc1.ConnectionString = strconn

End Sub
