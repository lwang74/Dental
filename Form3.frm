VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "Msadodc.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Begin VB.Form Form3 
   BackColor       =   &H00FFC0C0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "��򿵳ݿ�ǻҽ�����ﲡ������"
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
      Caption         =   "��  λ"
      BeginProperty Font 
         Name            =   "����"
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
      Caption         =   "ȷ  ��"
      BeginProperty Font 
         Name            =   "����"
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
         Name            =   "����"
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
      Caption         =   "����"
      BeginProperty Font 
         Name            =   "����"
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
      Caption         =   "���"
      BeginProperty Font 
         Name            =   "����"
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
         Name            =   "����"
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
         Name            =   "����"
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
         Name            =   "����"
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
      Caption         =   "�޸Ĳ���"
      BeginProperty Font 
         Name            =   "����"
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
      Caption         =   "��ѯ����"
      BeginProperty Font 
         Name            =   "����"
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
      Caption         =   "���Ӳ���"
      BeginProperty Font 
         Name            =   "����"
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
         Name            =   "����"
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
         Name            =   "����"
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
               Name            =   "����"
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
               Name            =   "����"
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
            Caption         =   "M/��"
            BeginProperty Font 
               Name            =   "����"
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
            Caption         =   "F/Ů"
            BeginProperty Font 
               Name            =   "����"
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
               Name            =   "����"
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
               Name            =   "����"
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
               Name            =   "����"
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
               Name            =   "����"
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
               Name            =   "����"
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
               Name            =   "����"
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
               Name            =   "����"
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
            Caption         =   "NAME ����"
            BeginProperty Font 
               Name            =   "����"
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
            Caption         =   "SEX   �Ա�"
            BeginProperty Font 
               Name            =   "����"
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
            Caption         =   "DATE OF BIRTH ��������"
            BeginProperty Font 
               Name            =   "����"
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
            Caption         =   "NATIONALITY ����"
            BeginProperty Font 
               Name            =   "����"
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
            Caption         =   "OCCUPATION   ְҵ"
            BeginProperty Font 
               Name            =   "����"
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
            Caption         =   "ID ORPASSPORT NO. ���֤���պ�"
            BeginProperty Font 
               Name            =   "����"
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
            Caption         =   "HOME ADDRESS ��ͥסַ"
            BeginProperty Font 
               Name            =   "����"
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
            Caption         =   "TEL �绰"
            BeginProperty Font 
               Name            =   "����"
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
            Caption         =   "OFFICE ADDRESS ��λ��ַ"
            BeginProperty Font 
               Name            =   "����"
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
            Caption         =   "TEL �绰"
            BeginProperty Font 
               Name            =   "����"
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
               Name            =   "����"
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
            Caption         =   "X����"
            BeginProperty Font 
               Name            =   "����"
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
               Name            =   "����"
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
         Begin VB.Label ��ǻ��� 
            Caption         =   "��ǻ���"
            BeginProperty Font 
               Name            =   "����"
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
               Name            =   "����"
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
            Caption         =   "��  ��"
            BeginProperty Font 
               Name            =   "����"
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
               Name            =   "����"
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
            Caption         =   "��ǻ����ʷ"
            BeginProperty Font 
               Name            =   "����"
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
            Caption         =   "(TO BE FILLED IF PATIENT IS BELOW 18 TEARS OLD)  18�����²�����д"
            BeginProperty Font 
               Name            =   "����"
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
            Caption         =   "Name of Parent/Guardian: ��ĸ/�໤������"
            BeginProperty Font 
               Name            =   "����"
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
            Caption         =   "Relationship to Patient: �벡�˹�ϵ"
            BeginProperty Font 
               Name            =   "����"
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
            Caption         =   "Address(if different from above): ��ַ�����������ĵ�ַ��ͬ��"
            BeginProperty Font 
               Name            =   "����"
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
            Caption         =   "I wish to have treatment for my children/my self.��ϣ�����Һ���/�ұ��˽�������"
            BeginProperty Font 
               Name            =   "����"
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
            Caption         =   "NOTTOBETAKENAWAY �˲�����ҪЯ����"
            BeginProperty Font 
               Name            =   "����"
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
            Caption         =   "Signature ǩ��"
            BeginProperty Font 
               Name            =   "����"
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
            Caption         =   "Date ����"
            BeginProperty Font 
               Name            =   "����"
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
            Caption         =   "(MEDICAL HISTORY) ����ʷ"
            BeginProperty Font 
               Name            =   "����"
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
            Caption         =   "10 Pregnancy      (����)"
            BeginProperty Font 
               Name            =   "����"
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
            Caption         =   "11 Aids           (���̲�)"
            BeginProperty Font 
               Name            =   "����"
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
            Caption         =   "12 Heart Disease  (���ಡ)"
            BeginProperty Font 
               Name            =   "����"
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
            Caption         =   "13 Anemia         (ƶѪ)"
            BeginProperty Font 
               Name            =   "����"
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
            Caption         =   "14 Epilepsy/Faints(���/����)"
            BeginProperty Font 
               Name            =   "����"
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
            Caption         =   "15 Others         (����)"
            BeginProperty Font 
               Name            =   "����"
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
            Caption         =   "I have/had the following: ����/�������£�"
            BeginProperty Font 
               Name            =   "����"
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
            Caption         =   "1 Drug Allergies     (ҩ�����)"
            BeginProperty Font 
               Name            =   "����"
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
            Caption         =   "2 Kidney Disease     (����)"
            BeginProperty Font 
               Name            =   "����"
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
            Caption         =   "3 Hepatitis          (����)"
            BeginProperty Font 
               Name            =   "����"
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
            Caption         =   "4 Infectious Disease (��Ⱦ��)"
            BeginProperty Font 
               Name            =   "����"
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
            Caption         =   "5 Rheumatic Fever    (��ʪ��)"
            BeginProperty Font 
               Name            =   "����"
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
            Caption         =   "6 High Blood Pressure(��Ѫѹ)"
            BeginProperty Font 
               Name            =   "����"
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
            Caption         =   "7 Abnormal Bleeding  (��Ѫ�쳣)"
            BeginProperty Font 
               Name            =   "����"
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
            Caption         =   "8 Asthma             (����)"
            BeginProperty Font 
               Name            =   "����"
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
            Caption         =   "9 Diabetes           (����)"
            BeginProperty Font 
               Name            =   "����"
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
            Caption         =   "16 Are you taking any medication? �����ڷ���ĳ��ҩ����"
            BeginProperty Font 
               Name            =   "����"
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
               Name            =   "����"
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
            Caption         =   "Medication Name ҩƷ���Ƽ���;"
            BeginProperty Font 
               Name            =   "����"
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
               Name            =   "����"
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
               Caption         =   "�������"
               ImageVarType    =   2
            EndProperty
            BeginProperty Tab2 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
               Caption         =   "����ʷ"
               ImageVarType    =   2
            EndProperty
            BeginProperty Tab3 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
               Caption         =   "18�����²�����д"
               ImageVarType    =   2
            EndProperty
            BeginProperty Tab4 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
               Caption         =   "�����¼"
               ImageVarType    =   2
            EndProperty
            BeginProperty Tab5 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
               Caption         =   "��ǻ����ʷ"
               ImageVarType    =   2
            EndProperty
            BeginProperty Tab6 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
               Caption         =   "����"
               ImageVarType    =   2
            EndProperty
            BeginProperty Tab7 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
               Caption         =   "��ǻ���"
               ImageVarType    =   2
            EndProperty
            BeginProperty Tab8 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
               Caption         =   "X����"
               ImageVarType    =   2
            EndProperty
            BeginProperty Tab9 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
               Caption         =   "���Ƽƻ�"
               ImageVarType    =   2
            EndProperty
         EndProperty
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "����"
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
               Name            =   "����"
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
            Caption         =   "���Ƽƻ�"
            BeginProperty Font 
               Name            =   "����"
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
            Caption         =   "����¼�¼"
            BeginProperty Font 
               Name            =   "����"
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
      Caption         =   "��ѯ������"
      BeginProperty Font 
         Name            =   "����"
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
      Caption         =   "��ѯ��ʽ��"
      BeginProperty Font 
         Name            =   "����"
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
      Caption         =   "������ţ�"
      BeginProperty Font 
         Name            =   "����"
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
         Name            =   "����"
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
      Caption         =   "��򿵳ݿ�ǻҽ�����ﲡ������"
      BeginProperty Font 
         Name            =   "��������"
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

'**********��֤������ţ�������7λ����*******************
If Trim(id.Text) = "" Then
    MsgBox "�����벡�����", 48, "��ʾ"
    Exit Sub
End If
If IsNumeric(id.Text) = False Then
    MsgBox "�����������������������֣�", 48, "����"
    Exit Sub
End If
If Len(id.Text) <> 7 Then
    MsgBox "������7λ������ţ�", 48, "����"
    Exit Sub
End If

If Option1.Value = True Then
    sex = "��"
Else
    sex = "Ů"
End If

'********������Ϣ��֤*********************************
For i = 0 To 7
    If Trim(txt(i).Text) = "" Then
        MsgBox "��Ϣ���벻������", 48, "��ʾ"
        Tabs1.Tabs(1).Selected = True
        'j = Tabs1.SelectedItem.Index
        'Picture1(j - 1).Visible = False
        'Picture1(0).Visible = True
        txt(i).SetFocus
        Exit Sub
    End If
Next i
    
'*******����ʷ��Ϣ��֤*******************************
flag = False
For i = 0 To 15
    If chk(i).Value = 1 Then
        flag = True
    End If
Next i
'x = 6
If flag = False Then
    x = MsgBox(" ����ʷѡ��ȫΪ�գ�ȷ�������ǣ�����ѡ��������", 52, "��ʾ")
End If
If x = 7 Then
    Tabs1.Tabs(2).Selected = True
    Exit Sub
End If

'*******18��������Ϣ��֤*******************************
For i = 8 To 8
    If Trim(txt(i).Text) = "" Then
       y = MsgBox(" �����Ƿ���18�����£�", 52, "��ʾ")
    End If
Next i
If y = 6 Then
    Tabs1.Tabs(3).Selected = True
    Exit Sub
End If

'********���没�˻�����Ϣ*******************************
With Adodc1
    .Recordset.AddNew
    .Recordset("�������") = id.Text
    .Recordset("����") = Trim(txt(0).Text)
    .Recordset("�Ա�") = sex
    .Recordset("��������") = Trim(bir.Value)
    .Recordset("����") = Trim(txt(1).Text)
    .Recordset("ְҵ") = Trim(txt(2).Text)
    .Recordset("֤����") = Trim(txt(3).Text)
    .Recordset("��ͥסַ") = Trim(txt(4).Text)
    .Recordset("լ��") = Trim(txt(5).Text)
    .Recordset("��λ��ַ") = Trim(txt(6).Text)
    .Recordset("��λ�绰") = Trim(txt(7).Text)
    .Recordset("�໤������") = Trim(txt(8).Text)
    .Recordset("�벡�˹�ϵ") = Trim(txt(9).Text)
    .Recordset("�໤�˵�ַ") = Trim(txt(10).Text)
    .Recordset("����") = Trim(txt(11).Text)
    .Recordset.Update
    .Refresh
End With
With Adodc2
    .Recordset.AddNew
    .Recordset("�������") = id.Text
    .Recordset("ҩ�����") = chk(0).Value
    .Recordset("����") = chk(1).Value
    .Recordset("����") = chk(2).Value
    .Recordset("��Ⱦ��") = chk(3).Value
    .Recordset("��ʪ��") = chk(4).Value
    .Recordset("��Ѫѹ") = chk(5).Value
    .Recordset("��Ѫ�쳣") = chk(6).Value
    .Recordset("����") = chk(7).Value
    .Recordset("����") = chk(8).Value
    .Recordset("����") = chk(9).Value
    .Recordset("���̲�") = chk(10).Value
    .Recordset("���ಡ") = chk(11).Value
    .Recordset("ƶѪ") = chk(12).Value
    .Recordset("���/����") = chk(13).Value
    .Recordset("����") = chk(14).Value
    .Recordset("�Ƿ����ҩ��") = chk(15).Value
    .Recordset("ҩ��������;") = Trim(zhfyw.Text)
    .Recordset.Update
    .Refresh
End With
MsgBox "���˻�����Ϣ����ɹ���", 64, "��ʾ"

id.Text = ""
For i = 0 To 11
    txt(i).Text = ""
Next i
For i = 0 To 15
    chk(i).Value = 0
Next i
zhfyw.Text = ""

'z = MsgBox("�Ƿ�������ò�������������ݣ�", 52, "��ʾ")
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

'������¼list
Public Sub set_grid(ByVal c_id As String)
Dim str3 As String

    sql3 = "select * from �����¼ where �������= '" & c_id & "'"

    Adodc3.RecordSource = sql3
    Adodc3.Refresh
    
    '*********��ȡ���������¼************************
    
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
Dim m As Integer    '*********************��¼��������������
Dim str1 As String
Dim str2 As String
'Dim str3 As String

str1 = Chr$(13) + Chr$(10)
str2 = ""

aaa:
    sql1 = "select * from ������ where �������= '" & c_id & "'"
    sql2 = "select * from ����ʷ where �������= '" & c_id & "'"
'    sql3 = "select * from �����¼ where �������= '" & c_id & "'"
    sql4 = "select * from ���Ƽƻ� where �������= '" & c_id & "'"
    sql5 = "select * from ��λ��¼ where ������� ='" & c_id & "'"

    Adodc1.RecordSource = sql1
    Adodc1.Refresh
    If Adodc1.Recordset.EOF Then
        MsgBox "�޴˲�������ȷ�Ϻ��ٲ�ѯ��", 48, "��ʾ"
        Exit Sub
    End If

    With Adodc1
        id.Text = .Recordset.Fields(0).Value
        Option1.Value = .Recordset.Fields(2).Value = "��"
        Option2.Value = .Recordset.Fields(2).Value = "Ů"
        
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

'**********�����¼���ʼ��**************************

Call ini


'*********��ȡ���������¼************************
    set_grid c_id

'*********���Ƽƻ���****************

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

'*******��λ��¼********************

��λ��¼.Adodc1.RecordSource = sql5
��λ��¼.Adodc1.Refresh
With ��λ��¼.Adodc1.Recordset
    Do While Not .EOF
        For i = 0 To 31
            ��λ��¼.Check1(i).Value = Abs(Int(.Fields(i + 2).Value))
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
Unload ��λ��¼
For i = 0 To 31
    ��λ��¼.Check1(i).Value = False
Next i

    FrmQuery.Show 1, Me
    Init_One g_id
End Sub

Private Sub Command3_Click()
Dim z As Integer
If Trim(id.Text) = "" Then
    MsgBox "�����벡����ţ�", 48, "��ʾ"
    id.SetFocus
    Exit Sub
End If

'WangLin 2012-11-1 Comment out below
'If Trim(kqjws.Text) = "" Or Trim(zs.Text) = "" Or Trim(kqjc.Text) = "" Or Trim(xg.Text) = "" Or Trim(zljh.Text) = "" Then
'   z = MsgBox("��ǻ����ʷ�����ߡ�X���顢���Ƽƻ���Ϣ���벻�������Ƿ�������д��", 52, "��ʾ")
'   End If
'If z = 6 Then
'    Exit Sub
'End If
'**********�޸���������Ƽ�¼********************

'WangLin 2012-11-1 Add below
With Adodc1
    .Recordset("�������") = g_id
    .Recordset("����") = Trim(txt(0).Text)
    If Option1.Value Then
        .Recordset("�Ա�") = "��"
    Else
        .Recordset("�Ա�") = "Ů"
    End If
    .Recordset("��������") = Trim(bir.Value)
    .Recordset("����") = Trim(txt(1).Text)
    .Recordset("ְҵ") = Trim(txt(2).Text)
    .Recordset("֤����") = Trim(txt(3).Text)
    .Recordset("��ͥסַ") = Trim(txt(4).Text)
    .Recordset("լ��") = Trim(txt(5).Text)
    .Recordset("��λ��ַ") = Trim(txt(6).Text)
    .Recordset("��λ�绰") = Trim(txt(7).Text)
    .Recordset("�໤������") = Trim(txt(8).Text)
    .Recordset("�벡�˹�ϵ") = Trim(txt(9).Text)
    .Recordset("�໤�˵�ַ") = Trim(txt(10).Text)
    .Recordset("����") = Trim(txt(11).Text)
    .Recordset.Update
    .Refresh
End With

'**********�Բ����������ֽ��и���****************

With Adodc4
    .Recordset.AddNew
    .Recordset("�������") = id.Text
    .Recordset("��ǻ����ʷ") = kqjws.Text
    .Recordset("����") = zs.Text
    .Recordset("��ǻ���") = kqjc.Text
    .Recordset("X����") = xg.Text
    .Recordset("���Ƽƻ�") = zljh.Text
    .Recordset.Update
    .Refresh
End With
MsgBox "�����޸ĳɹ���", 64, "��ʾ"

id.Text = ""
kqjws.Text = ""
zs.Text = ""
kqjc.Text = ""
xg.Text = ""
zljh.Text = ""
End Sub

Private Sub Command4_Click()
Dim i As Integer
Dim m As Integer    '*********************��¼��������������
Dim str1 As String
Dim str2 As String
Dim str3 As String
str1 = Chr$(13) + Chr$(10)
str2 = ""
If Option3.Value = False And Option4.Value = False Then
    MsgBox "��ѡ���ѯ��ʽ��", 48, "��ʾ"
    Exit Sub
End If
If Trim(Text1.Text) = "" Then
    MsgBox "�������ѯ������", 48, "��ʾ"
    Text1.SetFocus
    Exit Sub
End If
aaa:
If Option3.Value = True Then
    sql1 = " select * from ������ where �������= '" & Trim(Text1.Text) & "'"
    sql2 = "select * from ����ʷ where �������= '" & Trim(Text1.Text) & "'"
    sql3 = "select * from �����¼ where �������= '" & Trim(Text1.Text) & "'"
    sql4 = "select * from ���Ƽƻ� where �������= '" & Trim(Text1.Text) & "'"
    sql5 = "select * from ��λ��¼ where ������� ='" & Trim(Text1.Text) & "'"
End If
If Option4.Value = True Then
    sql1 = " select * from ������ where ����= '" & Trim(Text1.Text) & "'"
    'sql2 = "select * from ����ʷ where �������= '" & Trim(Text1.Text) & "'"
    'sql3 = "select * from �����¼ where �������= '" & Trim(Text1.Text) & "'"
    'sql4 = "select * from ���Ƽƻ� where �������= '" & Trim(Text1.Text) & "'"
End If
Adodc1.RecordSource = sql1
Adodc1.Refresh
If Adodc1.Recordset.EOF Then
    MsgBox "�޴˲�������ȷ�Ϻ��ٲ�ѯ��", 48, "��ʾ"
    Exit Sub
End If
With Adodc1
    If Option4.Value = True And .Recordset.RecordCount > 1 Then
        m = .Recordset.RecordCount
        Do While Not .Recordset.EOF
            str2 = str2 + .Recordset.Fields(0).Value + "    " + .Recordset.Fields(1).Value + str1
            .Recordset.MoveNext
        Loop
        str3 = InputBox("������ѯ�Ĳ�������������ȷ�Ϻ��ٲ�ѯ��" + str1 + str1 + str2, "�����벡�����")
        If str3 = "" Then
            Exit Sub
        End If
        Option3.Value = True
        Text1.Text = Trim(str3)
        GoTo aaa
    End If
    id.Text = .Recordset.Fields(0).Value
    If Option4.Value = True Then
        sql2 = "select * from ����ʷ where �������= '" & Trim(id.Text) & "'"
        sql3 = "select * from �����¼ where �������= '" & Trim(id.Text) & "'"
        sql4 = "select * from ���Ƽƻ� where �������= '" & Trim(id.Text) & "'"
        sql5 = "select * from ��λ��¼ where ������� ='" & Trim(id.Text) & "'"
    End If
    If .Recordset.Fields(2).Value = "��" Then
        Option1.Value = True
    Else
        If .Recordset.Fields(2).Value = "Ů" Then
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

'**********�����¼���ʼ��**************************

Call ini

Adodc3.RecordSource = sql3
Adodc3.Refresh
'MsgBox Adodc3.Recordset.Fields(1).Value

'*********��ȡ���������¼************************

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

'*********���Ƽƻ���****************

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

'*******��λ��¼********************

��λ��¼.Adodc1.RecordSource = sql5
��λ��¼.Adodc1.Refresh
With ��λ��¼.Adodc1.Recordset
    Do While Not .EOF
        For i = 0 To 31
            ��λ��¼.Check1(i).Value = Abs(Int(.Fields(i + 2).Value))
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
    MsgBox "�����벡����ţ�", 48, "��ʾ"
    id.SetFocus
    Exit Sub
End If
�����¼.Show (1)
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
Adodc1.RecordSource = "������"
Adodc1.Refresh
Adodc2.ConnectionString = strconn
Adodc2.RecordSource = "����ʷ"
Adodc2.Refresh
Adodc3.ConnectionString = strconn
Adodc3.RecordSource = "�����¼"
Adodc3.Refresh
Adodc4.ConnectionString = strconn
Adodc4.RecordSource = "���Ƽƻ�"
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
    MsgBox "����ѡ������������ѡ��", 48, "����"
    Exit Sub
End If
With msf
    �����¼.Lbl_seq.Caption = .TextMatrix(x, 0)

    �����¼.DTPicker1.Value = .TextMatrix(x, 2)
    �����¼.Text1.Text = .TextMatrix(x, 3)
    �����¼.Text2.Text = .TextMatrix(x, 4)
    �����¼.Text3.Text = .TextMatrix(x, 5)
    �����¼.Text4.Text = .TextMatrix(x, 6)
    �����¼.Text5.Text = .TextMatrix(x, 7)
End With
�����¼.Command1.Visible = False
�����¼.Cmd_save.Visible = True
�����¼.Command2.Visible = True
�����¼.Show (1)
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
        MsgBox "�����벡����ţ�", 48, "��ʾ"
        id.SetFocus
        Exit Sub
    End If
    ��λ��¼.Move 0, 0
    ��λ��¼.Text1.Text = id.Text
    ��λ��¼.Show
End If
Picture1(j).Visible = True

End Sub

