VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form Form1 
   Caption         =   "欢迎使用本系统，请登录"
   ClientHeight    =   5625
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   8505
   LinkTopic       =   "Form1"
   ScaleHeight     =   5625
   ScaleWidth      =   8505
   StartUpPosition =   2  '屏幕中心
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   615
      Left            =   2520
      Top             =   1080
      Visible         =   0   'False
      Width           =   3375
      _ExtentX        =   5953
      _ExtentY        =   1085
      ConnectMode     =   0
      CursorLocation  =   3
      IsolationLevel  =   -1
      ConnectionTimeout=   15
      CommandTimeout  =   30
      CursorType      =   3
      LockType        =   3
      CommandType     =   2
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
      Connect         =   "Provider=SQLOLEDB.1;Integrated Security=SSPI;Persist Security Info=False;Initial Catalog=shuihu;Data Source=DESKTOP-CKQKJT7"
      OLEDBString     =   "Provider=SQLOLEDB.1;Integrated Security=SSPI;Persist Security Info=False;Initial Catalog=shuihu;Data Source=DESKTOP-CKQKJT7"
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   "login"
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
      Caption         =   "取消"
      Height          =   495
      Left            =   4560
      TabIndex        =   3
      Top             =   4680
      Width           =   1095
   End
   Begin VB.TextBox Text2 
      Appearance      =   0  'Flat
      Height          =   375
      IMEMode         =   3  'DISABLE
      Left            =   3240
      PasswordChar    =   "*"
      TabIndex        =   2
      Top             =   3960
      Width           =   2415
   End
   Begin VB.TextBox Text1 
      Appearance      =   0  'Flat
      Height          =   375
      Left            =   3240
      TabIndex        =   1
      Top             =   3000
      Width           =   2415
   End
   Begin VB.CommandButton Command1 
      Caption         =   "登录"
      Height          =   495
      Left            =   3240
      TabIndex        =   0
      Top             =   4680
      Width           =   1095
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      Caption         =   "IPv6"
      ForeColor       =   &H000000FF&
      Height          =   255
      Left            =   6720
      TabIndex        =   8
      Top             =   1680
      Width           =   975
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "安徽建筑大学"
      BeginProperty Font 
         Name            =   "楷体"
         Size            =   26.25
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   615
      Left            =   2520
      TabIndex        =   7
      Top             =   600
      Width           =   3375
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "建筑用能数据采集系统"
      BeginProperty Font 
         Name            =   "隶书"
         Size            =   24
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   1095
      Left            =   1800
      TabIndex        =   6
      Top             =   1560
      Width           =   4935
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "密码(&P):"
      Height          =   255
      Left            =   2040
      TabIndex        =   5
      Top             =   4080
      Width           =   855
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "用户名称(&U):"
      ForeColor       =   &H00000000&
      Height          =   255
      Left            =   1920
      TabIndex        =   4
      Top             =   3120
      Width           =   1215
   End
   Begin VB.Image Image1 
      Height          =   10800
      Left            =   0
      Picture         =   "Form1.frx":0000
      Top             =   -120
      Width           =   28800
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
    If Text1.Text = "sa" Then
        If Text2.Text = "123456" Then
        MsgBox "登陆成功！欢迎使用"
            smartplus.Show
            Form1.Hide
    Else
        MsgBox "密码错误，请重新输入"
        Text2.Text = ""
        Text2.SetFocus
    End If
Else
    MsgBox "用户名不存在，请重新输入"
    Text1.Text = ""
    Text1.SetFocus
    End If
    
End Sub

Private Sub Picture1_Click()

End Sub

Private Sub Command2_Click()
 Me.Hide
End Sub

Private Sub Image2_Click()

End Sub


