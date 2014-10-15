VERSION 5.00
Begin VB.Form FrmQuery 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "查询"
   ClientHeight    =   6660
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   7620
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6660
   ScaleWidth      =   7620
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.ListBox List_names 
      BeginProperty Font 
         Name            =   "楷体_GB2312"
         Size            =   14.25
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4620
      Left            =   240
      TabIndex        =   7
      Top             =   1800
      Width           =   7095
   End
   Begin VB.CommandButton Command_OK 
      Caption         =   "确  定"
      BeginProperty Font 
         Name            =   "楷体_GB2312"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   6240
      TabIndex        =   6
      Top             =   960
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.Frame Frame1 
      BorderStyle     =   0  'None
      Caption         =   "Frame1"
      Height          =   1215
      Left            =   360
      TabIndex        =   0
      Top             =   360
      Width           =   5655
      Begin VB.OptionButton Opt 
         BackColor       =   &H00FFC0C0&
         Caption         =   "编号"
         BeginProperty Font 
            Name            =   "楷体_GB2312"
            Size            =   12
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   1
         Left            =   3240
         TabIndex        =   3
         Top             =   120
         Width           =   1215
      End
      Begin VB.OptionButton Opt 
         BackColor       =   &H00FFC0C0&
         Caption         =   "姓名"
         BeginProperty Font 
            Name            =   "楷体_GB2312"
            Size            =   12
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   0
         Left            =   1800
         TabIndex        =   2
         Top             =   120
         Value           =   -1  'True
         Width           =   1335
      End
      Begin VB.TextBox Tx_Name 
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
         Left            =   1800
         TabIndex        =   1
         Top             =   600
         Width           =   3735
      End
      Begin VB.Label Label47 
         BackStyle       =   0  'Transparent
         Caption         =   "查询方式："
         BeginProperty Font 
            Name            =   "楷体_GB2312"
            Size            =   12
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   5
         Top             =   120
         Width           =   1455
      End
      Begin VB.Label Label48 
         BackStyle       =   0  'Transparent
         Caption         =   "查询条件："
         BeginProperty Font 
            Name            =   "楷体_GB2312"
            Size            =   12
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   4
         Top             =   720
         Width           =   1455
      End
   End
End
Attribute VB_Name = "FrmQuery"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Const DEL As String = "    "

Private Sub Form_Load()
    Me.Top = (Screen.Height - Me.Height) / 2
    Me.Left = (Screen.Width - Me.Width) / 2
End Sub

Private Sub List_names_Click()
Dim arr() As String

    arr = Split(List_names.List(List_names.ListIndex), DEL)
    g_id = arr(0)
    Me.Hide
End Sub

Private Sub Opt_Click(Index As Integer)
    If Index = 0 Then 'name
        Me.Command_OK.Visible = False
    Else
        Me.Command_OK.Visible = True
    End If
End Sub

Private Sub Tx_Name_Change()
Dim sql As String

    If Me.Opt(0) Then 'name
        sql = "select * from 基本表 where 姓名 like  '%" & Trim(Me.Tx_Name.Text) & "%' order by 病历编号"
        set_list sql
    End If
End Sub

Private Sub Command_OK_Click()
Dim sql As String

    If Me.Opt(1) Then 'id
        sql = "select * from 基本表 where 病历编号 like  '%" & Trim(Me.Tx_Name.Text) & "%' order by 病历编号"
        set_list sql
    End If
End Sub

Private Sub set_list(ByVal sql As String)
    Me.List_names.Clear
    With Form3.Adodc1
        If Trim(Me.Tx_Name.Text) <> "" Then
            .RecordSource = sql
            .Refresh
            Do Until .Recordset.EOF
                Me.List_names.AddItem .Recordset.Fields(0).Value & DEL & .Recordset.Fields(1).Value
                .Recordset.MoveNext
                DoEvents
            Loop
        Else
            
        End If
    End With
End Sub


