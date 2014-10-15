VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "Msadodc.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Begin VB.Form Form3 
   BackColor       =   &H00FFC0C0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "天津康齿口腔医疗门诊病历档案"
   ClientHeight    =   8796
   ClientLeft      =   156
   ClientTop       =   432
   ClientWidth     =   11484
   LinkTopic       =   "Form3"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8796
   ScaleWidth      =   11484
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command6 
      Caption         =   "复  位"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   4560
      TabIndex        =   107
      Top             =   7920
      Width           =   1335
   End
   Begin VB.CommandButton Command4 
      Caption         =   "确  定"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   9960
      TabIndex        =   104
      Top             =   8040
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.TextBox Text1 
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   14.4
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   7800
      TabIndex        =   103
      Top             =   8040
      Visible         =   0   'False
      Width           =   1935
   End
   Begin VB.OptionButton Option4 
      BackColor       =   &H00FFC0C0&
      Caption         =   "姓名"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   9120
      TabIndex        =   101
      Top             =   7680
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.OptionButton Option3 
      BackColor       =   &H00FFC0C0&
      Caption         =   "编号"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   7800
      TabIndex        =   100
      Top             =   7680
      Visible         =   0   'False
      Width           =   1215
   End
   Begin MSAdodcLib.Adodc Adodc4 
      Height          =   495
      Left            =   5520
      Top             =   1320
      Visible         =   0   'False
      Width           =   2055
      _ExtentX        =   3620
      _ExtentY        =   868
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
      Caption         =   "Adodc4"
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
   Begin MSAdodcLib.Adodc Adodc3 
      Height          =   495
      Left            =   3720
      Top             =   1320
      Visible         =   0   'False
      Width           =   1815
      _ExtentX        =   3196
      _ExtentY        =   868
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
      Caption         =   "Adodc3"
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
   Begin MSAdodcLib.Adodc Adodc2 
      Height          =   495
      Left            =   1920
      Top             =   1320
      Visible         =   0   'False
      Width           =   1815
      _ExtentX        =   3196
      _ExtentY        =   868
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
      Caption         =   "Adodc2"
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
   Begin VB.CommandButton Command3 
      Caption         =   "修改病历"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   3120
      TabIndex        =   96
      Top             =   7920
      Width           =   1455
   End
   Begin VB.CommandButton Command2 
      Caption         =   "查询病历"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   1680
      TabIndex        =   94
      Top             =   7920
      Width           =   1455
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H0080FF80&
      Caption         =   "增加病历"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   240
      MaskColor       =   &H00FFFF80&
      TabIndex        =   93
      Top             =   7920
      Width           =   1455
   End
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   495
      Left            =   120
      Top             =   1320
      Visible         =   0   'False
      Width           =   1815
      _ExtentX        =   3196
      _ExtentY        =   868
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
   Begin VB.TextBox id 
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   14.4
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   420
      Left            =   9480
      MaxLength       =   7
      TabIndex        =   75
      Top             =   1410
      Width           =   1455
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H0000C000&
      Height          =   5535
      Left            =   120
      TabIndex        =   0
      Top             =   1920
      Width           =   11175
      Begin VB.PictureBox Picture1 
         Height          =   4095
         Index           =   0
         Left            =   480
         ScaleHeight     =   4044
         ScaleWidth      =   10164
         TabIndex        =   2
         Top             =   840
         Width           =   10215
         Begin MSComCtl2.DTPicker bir 
            Height          =   375
            Left            =   1920
            TabIndex        =   95
            Top             =   1080
            Width           =   3135
            _ExtentX        =   5525
            _ExtentY        =   656
            _Version        =   393216
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "宋体"
               Size            =   12
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Format          =   41091073
            CurrentDate     =   39388
         End
         Begin VB.TextBox txt 
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   12
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Index           =   0
            Left            =   1920
            TabIndex        =   13
            Top             =   360
            Width           =   3135
         End
         Begin VB.OptionButton Option1 
            Caption         =   "M/男"
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   10.8
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Left            =   6840
            TabIndex        =   12
            Top             =   480
            Value           =   -1  'True
            Width           =   735
         End
         Begin VB.OptionButton Option2 
            Caption         =   "F/女"
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   10.8
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Left            =   7800
            TabIndex        =   11
            Top             =   480
            Width           =   735
         End
         Begin VB.TextBox txt 
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   12
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Index           =   1
            Left            =   8040
            TabIndex        =   10
            Top             =   1080
            Width           =   1935
         End
         Begin VB.TextBox txt 
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   12
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Index           =   2
            Left            =   1920
            TabIndex        =   9
            Top             =   1800
            Width           =   3135
         End
         Begin VB.TextBox txt 
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   12
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Index           =   3
            Left            =   8040
            TabIndex        =   8
            Top             =   1800
            Width           =   1935
         End
         Begin VB.TextBox txt 
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   12
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Index           =   4
            Left            =   1920
            TabIndex        =   7
            Top             =   2400
            Width           =   5175
         End
         Begin VB.TextBox txt 
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   12
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Index           =   5
            Left            =   8040
            TabIndex        =   6
            Top             =   2400
            Width           =   1935
         End
         Begin VB.TextBox txt 
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   12
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Index           =   6
            Left            =   1920
            TabIndex        =   5
            Top             =   3000
            Width           =   5175
         End
         Begin VB.TextBox txt 
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   12
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Index           =   7
            Left            =   8040
            TabIndex        =   4
            Top             =   3000
            Width           =   1935
         End
         Begin VB.Label Label4 
            Caption         =   "NAME 姓名"
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   10.8
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   495
            Left            =   360
            TabIndex        =   23
            Top             =   360
            Width           =   495
         End
         Begin VB.Label Label5 
            Caption         =   "SEX   性别"
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   10.8
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   495
            Left            =   6000
            TabIndex        =   22
            Top             =   360
            Width           =   615
         End
         Begin VB.Label Label6 
            Caption         =   "DATE OF BIRTH 出生年月"
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   10.8
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   495
            Left            =   360
            TabIndex        =   21
            Top             =   1080
            Width           =   1575
         End
         Begin VB.Label Label7 
            BackStyle       =   0  'Transparent
            Caption         =   "NATIONALITY 国籍"
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   10.8
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   495
            Left            =   6000
            TabIndex        =   20
            Top             =   1080
            Width           =   1335
         End
         Begin VB.Label Label8 
            Caption         =   "OCCUPATION   职业"
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   10.8
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   495
            Left            =   360
            TabIndex        =   19
            Top             =   1800
            Width           =   1455
         End
         Begin VB.Label Label9 
            BackStyle       =   0  'Transparent
            Caption         =   "ID ORPASSPORT NO. 身份证或护照号"
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   10.8
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   495
            Left            =   6000
            TabIndex        =   18
            Top             =   1800
            Width           =   2055
         End
         Begin VB.Label Label10 
            Caption         =   "HOME ADDRESS 家庭住址"
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   10.8
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   495
            Left            =   360
            TabIndex        =   17
            Top             =   2400
            Width           =   1455
         End
         Begin VB.Label Label11 
            Caption         =   "TEL 电话"
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   10.8
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   495
            Left            =   7320
            TabIndex        =   16
            Top             =   2400
            Width           =   615
         End
         Begin VB.Label Label12 
            BackStyle       =   0  'Transparent
            Caption         =   "OFFICE ADDRESS 单位地址"
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   10.8
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   495
            Left            =   360
            TabIndex        =   15
            Top             =   3000
            Width           =   1695
         End
         Begin VB.Label Label13 
            BackStyle       =   0  'Transparent
            Caption         =   "TEL 电话"
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   10.8
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   495
            Left            =   7320
            TabIndex        =   14
            Top             =   3000
            Width           =   615
         End
      End
      Begin VB.PictureBox Picture1 
         Height          =   3735
         Index           =   7
         Left            =   360
         ScaleHeight     =   3684
         ScaleWidth      =   10164
         TabIndex        =   87
         Top             =   720
         Visible         =   0   'False
         Width           =   10215
         Begin VB.TextBox xg 
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   14.4
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   2655
            Left            =   1320
            MultiLine       =   -1  'True
            TabIndex        =   89
            Top             =   840
            Width           =   7575
         End
         Begin VB.Label lab 
            Caption         =   "X光检查"
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   14.4
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   4200
            TabIndex        =   88
            Top             =   360
            Width           =   1095
         End
      End
      Begin VB.PictureBox Picture1 
         Height          =   3735
         Index           =   6
         Left            =   360
         ScaleHeight     =   3684
         ScaleWidth      =   10164
         TabIndex        =   84
         Top             =   960
         Visible         =   0   'False
         Width           =   10215
         Begin VB.TextBox kqjc 
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   14.4
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   2655
            Left            =   1080
            MultiLine       =   -1  'True
            TabIndex        =   86
            Top             =   960
            Width           =   8175
         End
         Begin VB.Label 口腔检查 
            Caption         =   "口腔检查"
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   15
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   4200
            TabIndex        =   85
            Top             =   360
            Width           =   1455
         End
      End
      Begin VB.PictureBox Picture1 
         Height          =   3735
         Index           =   5
         Left            =   360
         ScaleHeight     =   3684
         ScaleWidth      =   10164
         TabIndex        =   81
         Top             =   960
         Visible         =   0   'False
         Width           =   10215
         Begin VB.TextBox zs 
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   14.4
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   2655
            Left            =   960
            MultiLine       =   -1  'True
            TabIndex        =   83
            Top             =   840
            Width           =   8055
         End
         Begin VB.Label Label44 
            Caption         =   "主  诉"
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   15
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   4320
            TabIndex        =   82
            Top             =   120
            Width           =   975
         End
      End
      Begin VB.PictureBox Picture1 
         Height          =   3735
         Index           =   4
         Left            =   360
         ScaleHeight     =   3684
         ScaleWidth      =   10164
         TabIndex        =   78
         Top             =   960
         Visible         =   0   'False
         Width           =   10215
         Begin VB.TextBox kqjws 
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   14.4
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   2295
            Left            =   1200
            MultiLine       =   -1  'True
            TabIndex        =   80
            Top             =   960
            Width           =   7935
         End
         Begin VB.Label Label43 
            Caption         =   "口腔既往史"
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   15.6
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00C00000&
            Height          =   375
            Left            =   3840
            TabIndex        =   79
            Top             =   240
            Width           =   1935
         End
      End
      Begin VB.PictureBox Picture1 
         Height          =   3735
         Index           =   2
         Left            =   360
         ScaleHeight     =   3684
         ScaleWidth      =   10164
         TabIndex        =   60
         Top             =   960
         Visible         =   0   'False
         Width           =   10215
         Begin VB.TextBox txt 
            Height          =   375
            Index           =   10
            Left            =   360
            TabIndex        =   97
            Top             =   2400
            Width           =   4215
         End
         Begin VB.TextBox txt 
            Height          =   375
            Index           =   11
            Left            =   5160
            TabIndex        =   70
            Top             =   2400
            Width           =   3735
         End
         Begin VB.TextBox txt 
            Height          =   375
            Index           =   8
            Left            =   360
            TabIndex        =   62
            Top             =   1320
            Width           =   4215
         End
         Begin VB.TextBox txt 
            Height          =   375
            Index           =   9
            Left            =   5160
            TabIndex        =   61
            Top             =   1320
            Width           =   3735
         End
         Begin VB.Label Label1 
            Alignment       =   2  'Center
            BackColor       =   &H00808000&
            Caption         =   "(TO BE FILLED IF PATIENT IS BELOW 18 TEARS OLD)  18岁以下病人填写"
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   12
               Charset         =   134
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   615
            Left            =   1920
            TabIndex        =   71
            Top             =   120
            Width           =   6375
         End
         Begin VB.Label Label34 
            Caption         =   "Name of Parent/Guardian: 父母/监护人姓名"
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   10.8
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   360
            TabIndex        =   69
            Top             =   960
            Width           =   4455
         End
         Begin VB.Label Label35 
            Caption         =   "Relationship to Patient: 与病人关系"
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   10.8
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   5160
            TabIndex        =   68
            Top             =   960
            Width           =   4335
         End
         Begin VB.Label Label36 
            Caption         =   "Address(if different from above): 地址（如果与上面的地址不同）"
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   10.8
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   495
            Left            =   360
            TabIndex        =   67
            Top             =   1800
            Width           =   3615
         End
         Begin VB.Label Label37 
            Caption         =   "I wish to have treatment for my children/my self.我希望对我孩子/我本人进行治疗"
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   10.8
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   495
            Left            =   5160
            TabIndex        =   66
            Top             =   1800
            Width           =   4575
         End
         Begin VB.Label Label38 
            BackColor       =   &H000000FF&
            Caption         =   "NOTTOBETAKENAWAY 此病历不要携带走"
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   12
               Charset         =   134
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   615
            Left            =   360
            TabIndex        =   65
            Top             =   3000
            Width           =   2175
         End
         Begin VB.Label Label39 
            Caption         =   "Signature 签字"
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   10.8
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   5160
            TabIndex        =   64
            Top             =   3000
            Width           =   1575
         End
         Begin VB.Label Label40 
            Caption         =   "Date 日期"
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   10.8
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   7920
            TabIndex        =   63
            Top             =   3000
            Width           =   1095
         End
      End
      Begin VB.PictureBox Picture1 
         Height          =   3735
         Index           =   1
         Left            =   360
         ScaleHeight     =   3684
         ScaleWidth      =   10164
         TabIndex        =   3
         Top             =   960
         Visible         =   0   'False
         Width           =   10215
         Begin VB.CheckBox chk 
            Caption         =   "Check1"
            Height          =   255
            Index           =   15
            Left            =   8520
            TabIndex        =   98
            Top             =   2880
            Width           =   255
         End
         Begin VB.CheckBox chk 
            Height          =   255
            Index           =   0
            Left            =   4320
            TabIndex        =   39
            Top             =   1320
            Width           =   255
         End
         Begin VB.CheckBox chk 
            Caption         =   "Check1"
            Height          =   255
            Index           =   1
            Left            =   4320
            TabIndex        =   38
            Top             =   1560
            Width           =   255
         End
         Begin VB.CheckBox chk 
            Caption         =   "Check1"
            Height          =   255
            Index           =   2
            Left            =   4320
            TabIndex        =   37
            Top             =   1800
            Width           =   255
         End
         Begin VB.CheckBox chk 
            Caption         =   "Check1"
            Height          =   255
            Index           =   3
            Left            =   4320
            TabIndex        =   36
            Top             =   2040
            Width           =   255
         End
         Begin VB.CheckBox chk 
            Caption         =   "Check1"
            Height          =   255
            Index           =   4
            Left            =   4320
            TabIndex        =   35
            Top             =   2280
            Width           =   255
         End
         Begin VB.CheckBox chk 
            Caption         =   "Check1"
            Height          =   255
            Index           =   5
            Left            =   4320
            TabIndex        =   34
            Top             =   2520
            Width           =   255
         End
         Begin VB.CheckBox chk 
            Caption         =   "Check1"
            Height          =   255
            Index           =   6
            Left            =   4320
            TabIndex        =   33
            Top             =   2760
            Width           =   255
         End
         Begin VB.CheckBox chk 
            Caption         =   "Check1"
            Height          =   255
            Index           =   7
            Left            =   4320
            TabIndex        =   32
            Top             =   3000
            Width           =   255
         End
         Begin VB.CheckBox chk 
            Caption         =   "Check17"
            Height          =   255
            Index           =   8
            Left            =   4320
            TabIndex        =   31
            Top             =   3240
            Width           =   255
         End
         Begin VB.CheckBox chk 
            Caption         =   "Check1"
            Height          =   255
            Index           =   14
            Left            =   8520
            TabIndex        =   30
            Top             =   2400
            Width           =   255
         End
         Begin VB.CheckBox chk 
            Caption         =   "Check1"
            Height          =   255
            Index           =   13
            Left            =   8520
            TabIndex        =   29
            Top             =   2160
            Width           =   255
         End
         Begin VB.CheckBox chk 
            Caption         =   "Check1"
            Height          =   255
            Index           =   12
            Left            =   8520
            TabIndex        =   28
            Top             =   1920
            Width           =   255
         End
         Begin VB.CheckBox chk 
            Caption         =   "Check1"
            Height          =   255
            Index           =   11
            Left            =   8520
            TabIndex        =   27
            Top             =   1680
            Width           =   255
         End
         Begin VB.CheckBox chk 
            Caption         =   "Check1"
            Height          =   255
            Index           =   10
            Left            =   8520
            TabIndex        =   26
            Top             =   1440
            Width           =   255
         End
         Begin VB.CheckBox chk 
            Height          =   255
            Index           =   9
            Left            =   8520
            TabIndex        =   25
            Top             =   1200
            Width           =   255
         End
         Begin VB.TextBox zhfyw 
            Height          =   375
            Left            =   7080
            TabIndex        =   24
            Top             =   3120
            Width           =   2655
         End
         Begin VB.Label Label41 
            BackColor       =   &H00808000&
            Caption         =   "(MEDICAL HISTORY) 既往史"
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   14.4
               Charset         =   134
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   2640
            TabIndex        =   74
            Top             =   240
            Width           =   4215
         End
         Begin VB.Label Label24 
            Caption         =   "10 Pregnancy      (怀孕)"
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   10.8
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   5280
            TabIndex        =   59
            Top             =   1200
            Width           =   3015
         End
         Begin VB.Label Label25 
            Caption         =   "11 Aids           (艾滋病)"
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   10.8
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   5280
            TabIndex        =   58
            Top             =   1440
            Width           =   3015
         End
         Begin VB.Label Label26 
            Caption         =   "12 Heart Disease  (心脏病)"
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   10.8
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   5280
            TabIndex        =   57
            Top             =   1680
            Width           =   3015
         End
         Begin VB.Label Label27 
            Caption         =   "13 Anemia         (贫血)"
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   10.8
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   5280
            TabIndex        =   56
            Top             =   1920
            Width           =   3015
         End
         Begin VB.Label Label28 
            Caption         =   "14 Epilepsy/Faints(癫痫/昏厥)"
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   10.8
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   5280
            TabIndex        =   55
            Top             =   2160
            Width           =   3135
         End
         Begin VB.Label Label29 
            Caption         =   "15 Others         (其它)"
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   10.8
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   5280
            TabIndex        =   54
            Top             =   2400
            Width           =   3015
         End
         Begin VB.Label Label14 
            Caption         =   "I have/had the following: 我有/会有如下："
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   10.8
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   240
            TabIndex        =   53
            Top             =   840
            Width           =   4455
         End
         Begin VB.Label Label15 
            Caption         =   "1 Drug Allergies     (药物过敏)"
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   10.8
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   840
            TabIndex        =   52
            Top             =   1320
            Width           =   3255
         End
         Begin VB.Label Label16 
            Caption         =   "2 Kidney Disease     (肾病)"
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   10.8
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   840
            TabIndex        =   51
            Top             =   1560
            Width           =   3015
         End
         Begin VB.Label Label17 
            Caption         =   "3 Hepatitis          (肝炎)"
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   10.8
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   840
            TabIndex        =   50
            Top             =   1800
            Width           =   3015
         End
         Begin VB.Label Label18 
            Caption         =   "4 Infectious Disease (传染病)"
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   10.8
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   840
            TabIndex        =   49
            Top             =   2040
            Width           =   3135
         End
         Begin VB.Label Label19 
            Caption         =   "5 Rheumatic Fever    (风湿热)"
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   10.8
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   840
            TabIndex        =   48
            Top             =   2280
            Width           =   3135
         End
         Begin VB.Label Label20 
            Caption         =   "6 High Blood Pressure(高血压)"
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   10.8
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   840
            TabIndex        =   47
            Top             =   2520
            Width           =   3135
         End
         Begin VB.Label Label21 
            Caption         =   "7 Abnormal Bleeding  (出血异常)"
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   10.8
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   840
            TabIndex        =   46
            Top             =   2760
            Width           =   3255
         End
         Begin VB.Label Label22 
            Caption         =   "8 Asthma             (哮喘)"
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   10.8
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   840
            TabIndex        =   45
            Top             =   3000
            Width           =   3015
         End
         Begin VB.Label Label23 
            Caption         =   "9 Diabetes           (糖尿病)"
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   10.8
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   840
            TabIndex        =   44
            Top             =   3240
            Width           =   3255
         End
         Begin VB.Label Label30 
            BackStyle       =   0  'Transparent
            Caption         =   "16 Are you taking any medication? 你正在服用某种药物吗？"
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   10.8
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   495
            Left            =   5280
            TabIndex        =   43
            Top             =   2640
            Width           =   3735
         End
         Begin VB.Label Label31 
            Caption         =   "Yes"
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   10.8
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   4200
            TabIndex        =   42
            Top             =   1080
            Width           =   375
         End
         Begin VB.Label Label32 
            BackStyle       =   0  'Transparent
            Caption         =   "Medication Name 药品名称及用途"
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   10.8
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   495
            Left            =   5280
            TabIndex        =   41
            Top             =   3120
            Width           =   1695
         End
         Begin VB.Label Label33 
            Caption         =   "Yes"
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   10.8
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   8400
            TabIndex        =   40
            Top             =   960
            Width           =   375
         End
      End
      Begin MSComctlLib.TabStrip Tabs1 
         Height          =   5175
         Left            =   240
         TabIndex        =   1
         Top             =   240
         Width           =   10695
         _ExtentX        =   18860
         _ExtentY        =   9123
         MultiRow        =   -1  'True
         TabFixedWidth   =   5292
         _Version        =   393216
         BeginProperty Tabs {1EFB6598-857C-11D1-B16A-00C0F0283628} 
            NumTabs         =   9
            BeginProperty Tab1 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
               Caption         =   "基本情况"
               ImageVarType    =   2
            EndProperty
            BeginProperty Tab2 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
               Caption         =   "既往史"
               ImageVarType    =   2
            EndProperty
            BeginProperty Tab3 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
               Caption         =   "18岁以下病人填写"
               ImageVarType    =   2
            EndProperty
            BeginProperty Tab4 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
               Caption         =   "门诊记录"
               ImageVarType    =   2
            EndProperty
            BeginProperty Tab5 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
               Caption         =   "口腔既往史"
               ImageVarType    =   2
            EndProperty
            BeginProperty Tab6 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
               Caption         =   "主诉"
               ImageVarType    =   2
            EndProperty
            BeginProperty Tab7 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
               Caption         =   "口腔检查"
               ImageVarType    =   2
            EndProperty
            BeginProperty Tab8 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
               Caption         =   "X光检查"
               ImageVarType    =   2
            EndProperty
            BeginProperty Tab9 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
               Caption         =   "治疗计划"
               ImageVarType    =   2
            EndProperty
         EndProperty
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "宋体"
            Size            =   12
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin VB.PictureBox Picture1 
         Height          =   4095
         Index           =   8
         Left            =   360
         ScaleHeight     =   4044
         ScaleWidth      =   10164
         TabIndex        =   90
         Top             =   720
         Visible         =   0   'False
         Width           =   10215
         Begin VB.TextBox zljh 
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   14.4
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   2655
            Left            =   1080
            MultiLine       =   -1  'True
            TabIndex        =   92
            Top             =   720
            Width           =   8055
         End
         Begin VB.Label Label46 
            Caption         =   "治疗计划"
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   14.4
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   4080
            TabIndex        =   91
            Top             =   240
            Width           =   1455
         End
      End
      Begin VB.PictureBox Picture1 
         Height          =   3735
         Index           =   3
         Left            =   360
         ScaleHeight     =   3684
         ScaleWidth      =   10164
         TabIndex        =   77
         Top             =   960
         Visible         =   0   'False
         Width           =   10215
         Begin VB.CommandButton Command5 
            Caption         =   "添加新纪录"
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   12
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   495
            Left            =   4200
            TabIndex        =   106
            Top             =   3000
            Width           =   1695
         End
         Begin MSFlexGridLib.MSFlexGrid msf 
            Height          =   2892
            Left            =   480
            TabIndex        =   105
            Top             =   960
            Width           =   10212
            _ExtentX        =   18013
            _ExtentY        =   5101
            _Version        =   393216
            Cols            =   8
            RowHeightMin    =   10
         End
      End
   End
   Begin VB.Label Label48 
      BackStyle       =   0  'Transparent
      Caption         =   "查询条件："
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   6360
      TabIndex        =   102
      Top             =   8160
      Visible         =   0   'False
      Width           =   1455
   End
   Begin VB.Label Label47 
      BackStyle       =   0  'Transparent
      Caption         =   "查询方式："
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   6360
      TabIndex        =   99
      Top             =   7680
      Visible         =   0   'False
      Width           =   1455
   End
   Begin VB.Label Label42 
      BackStyle       =   0  'Transparent
      Caption         =   "病历编号："
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   14.4
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C000C0&
      Height          =   375
      Left            =   8040
      TabIndex        =   76
      Top             =   1440
      Width           =   1575
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      BackColor       =   &H8000000A&
      BackStyle       =   0  'Transparent
      Caption         =   "TianJin Kang Chi Dental Center Medical Record"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   18
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   375
      Left            =   1560
      TabIndex        =   73
      Top             =   360
      Width           =   8415
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackColor       =   &H8000000A&
      BackStyle       =   0  'Transparent
      Caption         =   "天津康齿口腔医疗门诊病历档案"
      BeginProperty Font 
         Name            =   "华文中宋"
         Size            =   18
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   375
      Left            =   3000
      TabIndex        =   72
      Top             =   840
      Width           =   5535
   End
End
Attribute VB_Name = "Form3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Command1_Click()
Dim sex As String
Dim i As Integer
Dim j As Integer
Dim x As Integer
Dim y As Integer
Dim z As Integer
Dim flag As Boolean
Dim flag1 As Boolean

'**********验证病历编号，必须是7位数字*******************
If Trim(id.Text) = "" Then
    MsgBox "请输入病历编号", 48, "提示"
    Exit Sub
End If
If IsNumeric(id.Text) = False Then
    MsgBox "病历编号输入错误，请输入数字！", 48, "错误"
    Exit Sub
End If
If Len(id.Text) <> 7 Then
    MsgBox "请输入7位病历编号！", 48, "错误"
    Exit Sub
End If

If Option1.Value = True Then
    sex = "男"
Else
    sex = "女"
End If

'********基本信息验证*********************************
For i = 0 To 7
    If Trim(txt(i).Text) = "" Then
        MsgBox "信息输入不完整！", 48, "提示"
        Tabs1.Tabs(1).Selected = True
        'j = Tabs1.SelectedItem.Index
        'Picture1(j - 1).Visible = False
        'Picture1(0).Visible = True
        txt(i).SetFocus
        Exit Sub
    End If
Next i
    
'*******既往史信息验证*******************************
flag = False
For i = 0 To 15
    If chk(i).Value = 1 Then
        flag = True
    End If
Next i
'x = 6
If flag = False Then
    x = MsgBox(" 既往史选择全为空，确定请点击是，重新选择请点击否！", 52, "提示")
End If
If x = 7 Then
    Tabs1.Tabs(2).Selected = True
    Exit Sub
End If

'*******18岁以下信息验证*******************************
For i = 8 To 8
    If Trim(txt(i).Text) = "" Then
       y = MsgBox(" 病人是否是18岁以下？", 52, "提示")
    End If
Next i
If y = 6 Then
    Tabs1.Tabs(3).Selected = True
    Exit Sub
End If

'********保存病人基本信息*******************************
With Adodc1
    .Recordset.AddNew
    .Recordset("病历编号") = id.Text
    .Recordset("姓名") = Trim(txt(0).Text)
    .Recordset("性别") = sex
    .Recordset("出生年月") = Trim(bir.Value)
    .Recordset("国籍") = Trim(txt(1).Text)
    .Recordset("职业") = Trim(txt(2).Text)
    .Recordset("证件号") = Trim(txt(3).Text)
    .Recordset("家庭住址") = Trim(txt(4).Text)
    .Recordset("宅电") = Trim(txt(5).Text)
    .Recordset("单位地址") = Trim(txt(6).Text)
    .Recordset("单位电话") = Trim(txt(7).Text)
    .Recordset("监护人姓名") = Trim(txt(8).Text)
    .Recordset("与病人关系") = Trim(txt(9).Text)
    .Recordset("监护人地址") = Trim(txt(10).Text)
    .Recordset("治疗") = Trim(txt(11).Text)
    .Recordset.Update
    .Refresh
End With
With Adodc2
    .Recordset.AddNew
    .Recordset("病历编号") = id.Text
    .Recordset("药物过敏") = chk(0).Value
    .Recordset("肾病") = chk(1).Value
    .Recordset("肝炎") = chk(2).Value
    .Recordset("传染病") = chk(3).Value
    .Recordset("风湿热") = chk(4).Value
    .Recordset("高血压") = chk(5).Value
    .Recordset("出血异常") = chk(6).Value
    .Recordset("哮喘") = chk(7).Value
    .Recordset("糖尿病") = chk(8).Value
    .Recordset("怀孕") = chk(9).Value
    .Recordset("艾滋病") = chk(10).Value
    .Recordset("心脏病") = chk(11).Value
    .Recordset("贫血") = chk(12).Value
    .Recordset("癫痫/昏厥") = chk(13).Value
    .Recordset("其它") = chk(14).Value
    .Recordset("是否服用药物") = chk(15).Value
    .Recordset("药物名称用途") = Trim(zhfyw.Text)
    .Recordset.Update
    .Refresh
End With
MsgBox "病人基本信息保存成功！", 64, "提示"

id.Text = ""
For i = 0 To 11
    txt(i).Text = ""
Next i
For i = 0 To 15
    chk(i).Value = 0
Next i
zhfyw.Text = ""

'z = MsgBox("是否继续给该病历添加其它内容？", 52, "提示")
'If z = 7 Then
    'id.Text = ""
    'For i = 0 To 11
        'txt(i).Text = ""
    'Next i
    'For i = 0 To 15
        'chk(i).Value = 0
    'Next i
    'zhfyw.Text = ""
'End If
End Sub

'病历记录list
Public Sub set_grid(ByVal c_id As String)
Dim str3 As String

    sql3 = "select * from 门诊记录 where 病历编号= '" & c_id & "'"

    Adodc3.RecordSource = sql3
    Adodc3.Refresh
    
    '*********读取病人门诊记录************************
    
    msf.FixedCols = 1
    msf.FixedRows = 1
    msf.Rows = 2
    
    Do While Not Adodc3.Recordset.EOF
        With Adodc3.Recordset
            
            msf.TextMatrix(msf.Rows - 1, 0) = .Fields(0).Value
            msf.TextMatrix(msf.Rows - 1, 1) = .Fields(1).Value
            msf.TextMatrix(msf.Rows - 1, 2) = .Fields(2).Value
            msf.TextMatrix(msf.Rows - 1, 3) = .Fields(3).Value
            msf.TextMatrix(msf.Rows - 1, 4) = .Fields(4).Value
            msf.TextMatrix(msf.Rows - 1, 5) = .Fields(5).Value
            msf.TextMatrix(msf.Rows - 1, 6) = .Fields(6).Value
            msf.TextMatrix(msf.Rows - 1, 7) = .Fields(7).Value
            'msf.ColAlignment(0) = flexAlignCenterCenter
            .MoveNext
        End With
        msf.Rows = msf.Rows + 1
    Loop
End Sub

Private Sub Init_One(ByVal c_id As String)
Dim i As Integer
Dim m As Integer    '*********************记录重名病历的数量
Dim str1 As String
Dim str2 As String
'Dim str3 As String

str1 = Chr$(13) + Chr$(10)
str2 = ""

aaa:
    sql1 = "select * from 基本表 where 病历编号= '" & c_id & "'"
    sql2 = "select * from 既往史 where 病历编号= '" & c_id & "'"
'    sql3 = "select * from 门诊记录 where 病历编号= '" & c_id & "'"
    sql4 = "select * from 治疗计划 where 病历编号= '" & c_id & "'"
    sql5 = "select * from 牙位记录 where 病历编号 ='" & c_id & "'"

    Adodc1.RecordSource = sql1
    Adodc1.Refresh
    If Adodc1.Recordset.EOF Then
        MsgBox "无此病历，请确认后再查询！", 48, "提示"
        Exit Sub
    End If

    With Adodc1
        id.Text = .Recordset.Fields(0).Value
        Option1.Value = .Recordset.Fields(2).Value = "男"
        Option2.Value = .Recordset.Fields(2).Value = "女"
        
        txt(0).Text = .Recordset.Fields(1).Value
        bir.Value = .Recordset.Fields(3).Value
        For i = 1 To 11
            If IsNull(.Recordset.Fields(i + 3).Value) Then
                txt(i).Text = ""
            Else
                txt(i).Text = .Recordset.Fields(i + 3).Value
            End If
        Next i
    End With
    
    Adodc2.RecordSource = sql2
    Adodc2.Refresh
    With Adodc2
        Do While Not .Recordset.EOF
            For i = 0 To 15
                chk(i).Value = Abs(Int(.Recordset.Fields(i + 1).Value))
            Next i
            zhfyw.Text = .Recordset.Fields(17).Value
            Exit Do
        Loop
    End With

'**********门诊记录表初始化**************************

Call ini


'*********读取病人门诊记录************************
    set_grid c_id

'*********治疗计划等****************

Adodc4.RecordSource = sql4
Adodc4.Refresh
With Adodc4.Recordset
    Do While Not .EOF
        kqjws.Text = .Fields(2).Value
        zs.Text = .Fields(3).Value
        kqjc.Text = .Fields(4).Value
        xg.Text = .Fields(5).Value
        zljh.Text = .Fields(6).Value
        Exit Do
    Loop
End With

'*******牙位记录********************

牙位记录.Adodc1.RecordSource = sql5
牙位记录.Adodc1.Refresh
With 牙位记录.Adodc1.Recordset
    Do While Not .EOF
        For i = 0 To 31
            牙位记录.Check1(i).Value = Abs(Int(.Fields(i + 2).Value))
        Next i
        Exit Do
    Loop
End With
'******************************

Command2.Enabled = True
End Sub

Private Sub Command2_Click()
Dim i As Integer
'Adodc1.Recordset.AddNew
'Command2.Enabled = False
'Label47.Visible = True
'Label48.Visible = True
'Option3.Visible = True
'Option4.Visible = True
'Text1.Visible = True
'Command4.Visible = True
'Text1.Text = ""
Unload 牙位记录
For i = 0 To 31
    牙位记录.Check1(i).Value = False
Next i

    FrmQuery.Show 1, Me
    Init_One g_id
End Sub

Private Sub Command3_Click()
Dim z As Integer
If Trim(id.Text) = "" Then
    MsgBox "请输入病历编号！", 48, "提示"
    id.SetFocus
    Exit Sub
End If

'WangLin 2012-11-1 Comment out below
'If Trim(kqjws.Text) = "" Or Trim(zs.Text) = "" Or Trim(kqjc.Text) = "" Or Trim(xg.Text) = "" Or Trim(zljh.Text) = "" Then
'   z = MsgBox("口腔既往史、主诉、X光检查、治疗计划信息输入不完整！是否重新填写？", 52, "提示")
'   End If
'If z = 6 Then
'    Exit Sub
'End If
'**********修改添加新诊疗记录********************

'WangLin 2012-11-1 Add below
With Adodc1
    .Recordset("病历编号") = g_id
    .Recordset("姓名") = Trim(txt(0).Text)
    If Option1.Value Then
        .Recordset("性别") = "男"
    Else
        .Recordset("性别") = "女"
    End If
    .Recordset("出生年月") = Trim(bir.Value)
    .Recordset("国籍") = Trim(txt(1).Text)
    .Recordset("职业") = Trim(txt(2).Text)
    .Recordset("证件号") = Trim(txt(3).Text)
    .Recordset("家庭住址") = Trim(txt(4).Text)
    .Recordset("宅电") = Trim(txt(5).Text)
    .Recordset("单位地址") = Trim(txt(6).Text)
    .Recordset("单位电话") = Trim(txt(7).Text)
    .Recordset("监护人姓名") = Trim(txt(8).Text)
    .Recordset("与病人关系") = Trim(txt(9).Text)
    .Recordset("监护人地址") = Trim(txt(10).Text)
    .Recordset("治疗") = Trim(txt(11).Text)
    .Recordset.Update
    .Refresh
End With

'**********对病历其它部分进行更新****************

With Adodc4
    .Recordset.AddNew
    .Recordset("病历编号") = id.Text
    .Recordset("口腔既往史") = kqjws.Text
    .Recordset("主诉") = zs.Text
    .Recordset("口腔检查") = kqjc.Text
    .Recordset("X光检查") = xg.Text
    .Recordset("治疗计划") = zljh.Text
    .Recordset.Update
    .Refresh
End With
MsgBox "病历修改成功！", 64, "提示"

id.Text = ""
kqjws.Text = ""
zs.Text = ""
kqjc.Text = ""
xg.Text = ""
zljh.Text = ""
End Sub

Private Sub Command4_Click()
Dim i As Integer
Dim m As Integer    '*********************记录重名病历的数量
Dim str1 As String
Dim str2 As String
Dim str3 As String
str1 = Chr$(13) + Chr$(10)
str2 = ""
If Option3.Value = False And Option4.Value = False Then
    MsgBox "请选择查询方式！", 48, "提示"
    Exit Sub
End If
If Trim(Text1.Text) = "" Then
    MsgBox "请输入查询条件！", 48, "提示"
    Text1.SetFocus
    Exit Sub
End If
aaa:
If Option3.Value = True Then
    sql1 = " select * from 基本表 where 病历编号= '" & Trim(Text1.Text) & "'"
    sql2 = "select * from 既往史 where 病历编号= '" & Trim(Text1.Text) & "'"
    sql3 = "select * from 门诊记录 where 病历编号= '" & Trim(Text1.Text) & "'"
    sql4 = "select * from 治疗计划 where 病历编号= '" & Trim(Text1.Text) & "'"
    sql5 = "select * from 牙位记录 where 病历编号 ='" & Trim(Text1.Text) & "'"
End If
If Option4.Value = True Then
    sql1 = " select * from 基本表 where 姓名= '" & Trim(Text1.Text) & "'"
    'sql2 = "select * from 既往史 where 病历编号= '" & Trim(Text1.Text) & "'"
    'sql3 = "select * from 门诊记录 where 病历编号= '" & Trim(Text1.Text) & "'"
    'sql4 = "select * from 治疗计划 where 病历编号= '" & Trim(Text1.Text) & "'"
End If
Adodc1.RecordSource = sql1
Adodc1.Refresh
If Adodc1.Recordset.EOF Then
    MsgBox "无此病历，请确认后再查询！", 48, "提示"
    Exit Sub
End If
With Adodc1
    If Option4.Value = True And .Recordset.RecordCount > 1 Then
        m = .Recordset.RecordCount
        Do While Not .Recordset.EOF
            str2 = str2 + .Recordset.Fields(0).Value + "    " + .Recordset.Fields(1).Value + str1
            .Recordset.MoveNext
        Loop
        str3 = InputBox("您所查询的病历有重名，请确认后再查询！" + str1 + str1 + str2, "请输入病历编号")
        If str3 = "" Then
            Exit Sub
        End If
        Option3.Value = True
        Text1.Text = Trim(str3)
        GoTo aaa
    End If
    id.Text = .Recordset.Fields(0).Value
    If Option4.Value = True Then
        sql2 = "select * from 既往史 where 病历编号= '" & Trim(id.Text) & "'"
        sql3 = "select * from 门诊记录 where 病历编号= '" & Trim(id.Text) & "'"
        sql4 = "select * from 治疗计划 where 病历编号= '" & Trim(id.Text) & "'"
        sql5 = "select * from 牙位记录 where 病历编号 ='" & Trim(id.Text) & "'"
    End If
    If .Recordset.Fields(2).Value = "男" Then
        Option1.Value = True
    Else
        If .Recordset.Fields(2).Value = "女" Then
            Option2.Value = True
        End If
    End If
    txt(0).Text = .Recordset.Fields(1).Value
    bir.Value = .Recordset.Fields(3).Value
    For i = 1 To 11
        If IsNull(.Recordset.Fields(i + 3).Value) Then
            txt(i).Text = ""
        Else
            txt(i).Text = .Recordset.Fields(i + 3).Value
        End If
    Next i
End With
Adodc2.RecordSource = sql2
Adodc2.Refresh
With Adodc2
    Do While Not .Recordset.EOF
        For i = 0 To 15
            chk(i).Value = Abs(Int(.Recordset.Fields(i + 1).Value))
        Next i
        zhfyw.Text = .Recordset.Fields(17).Value
        Exit Do
    Loop
End With

'**********门诊记录表初始化**************************

Call ini

Adodc3.RecordSource = sql3
Adodc3.Refresh
'MsgBox Adodc3.Recordset.Fields(1).Value

'*********读取病人门诊记录************************

msf.FixedCols = 1
msf.FixedRows = 1
msf.Rows = 2

Do While Not Adodc3.Recordset.EOF
    With Adodc3.Recordset
        
        msf.TextMatrix(msf.Rows - 1, 0) = .Fields(0).Value
        msf.TextMatrix(msf.Rows - 1, 1) = .Fields(1).Value
        msf.TextMatrix(msf.Rows - 1, 2) = .Fields(2).Value
        msf.TextMatrix(msf.Rows - 1, 3) = .Fields(3).Value
        msf.TextMatrix(msf.Rows - 1, 4) = .Fields(4).Value
        msf.TextMatrix(msf.Rows - 1, 5) = .Fields(5).Value
        msf.TextMatrix(msf.Rows - 1, 6) = .Fields(6).Value
        msf.TextMatrix(msf.Rows - 1, 7) = .Fields(7).Value
        'msf.ColAlignment(0) = flexAlignCenterCenter
        .MoveNext
    End With
    msf.Rows = msf.Rows + 1
Loop

'*********治疗计划等****************

Adodc4.RecordSource = sql4
Adodc4.Refresh
With Adodc4.Recordset
    Do While Not .EOF
        kqjws.Text = .Fields(2).Value
        zs.Text = .Fields(3).Value
        kqjc.Text = .Fields(4).Value
        xg.Text = .Fields(5).Value
        zljh.Text = .Fields(6).Value
        Exit Do
    Loop
End With

'*******牙位记录********************

牙位记录.Adodc1.RecordSource = sql5
牙位记录.Adodc1.Refresh
With 牙位记录.Adodc1.Recordset
    Do While Not .EOF
        For i = 0 To 31
            牙位记录.Check1(i).Value = Abs(Int(.Fields(i + 2).Value))
        Next i
        Exit Do
    Loop
End With
'******************************


Command2.Enabled = True
Label47.Visible = False
Label48.Visible = False
Option3.Visible = False
Option4.Visible = False
Text1.Visible = False
Command4.Visible = False


End Sub

Private Sub Command5_Click()
If Trim(id.Text) = "" Then
    MsgBox "请输入病历编号！", 48, "提示"
    id.SetFocus
    Exit Sub
End If
门诊记录.Show (1)
End Sub

Private Sub Command6_Click()
Dim i As Integer
id.Text = ""
For i = 0 To 11
        txt(i).Text = ""
Next i
For i = 0 To 15
        chk(i).Value = 0
Next i
zhfyw.Text = ""
Call ini
kqjws.Text = ""
zs.Text = ""
kqjc.Text = ""
xg.Text = ""
zljh.Text = ""
Command2.Enabled = True
Label47.Visible = False
Label48.Visible = False
Option3.Visible = False
Option4.Visible = False
Text1.Visible = False
Command4.Visible = False
End Sub

Private Sub Form_Load()
'Dim strconn As String
'Dim num As Long
strconn = App.Path & "\db\bl.mdb"
strconn = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source =" & strconn
strconn = strconn & ";Persist Security Info=False"
Adodc1.ConnectionString = strconn
Adodc1.RecordSource = "基本表"
Adodc1.Refresh
Adodc2.ConnectionString = strconn
Adodc2.RecordSource = "既往史"
Adodc2.Refresh
Adodc3.ConnectionString = strconn
Adodc3.RecordSource = "门诊记录"
Adodc3.Refresh
Adodc4.ConnectionString = strconn
Adodc4.RecordSource = "治疗计划"
Adodc4.Refresh

Call ini

    'Set same position.
    For i = 0 To Picture1.Count - 1
        Picture1(i).Left = Picture1(0).Left
        Picture1(i).Top = Picture1(0).Top
    Next i


End Sub

Private Sub msf_Click()
Dim x As Integer
x = msf.Row
If Trim(msf.TextMatrix(x, 2)) = "" Then
    MsgBox "您的选择有误，请重新选择！", 48, "错误"
    Exit Sub
End If
With msf
    门诊记录.Lbl_seq.Caption = .TextMatrix(x, 0)

    门诊记录.DTPicker1.Value = .TextMatrix(x, 2)
    门诊记录.Text1.Text = .TextMatrix(x, 3)
    门诊记录.Text2.Text = .TextMatrix(x, 4)
    门诊记录.Text3.Text = .TextMatrix(x, 5)
    门诊记录.Text4.Text = .TextMatrix(x, 6)
    门诊记录.Text5.Text = .TextMatrix(x, 7)
End With
门诊记录.Command1.Visible = False
门诊记录.Cmd_save.Visible = True
门诊记录.Command2.Visible = True
门诊记录.Show (1)
End Sub

Private Sub Option3_Click()
Text1.Text = ""
End Sub

Private Sub Option4_Click()
Text1.Text = ""
End Sub



Private Sub Tabs1_Click()
Dim i, j As Integer
j = Tabs1.SelectedItem.Index
For i = 0 To 8
    Picture1(i).Visible = False
Next i
j = j - 1
If j = 3 Then
    If Trim(id.Text) = "" Then
        MsgBox "请输入病历编号！", 48, "提示"
        id.SetFocus
        Exit Sub
    End If
    牙位记录.Move 0, 0
    牙位记录.Text1.Text = id.Text
    牙位记录.Show
End If
Picture1(j).Visible = True

End Sub

