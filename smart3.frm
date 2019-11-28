VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{65E121D4-0C60-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCHRT20.OCX"
Begin VB.Form smartplus 
   Caption         =   "智能插座终端"
   ClientHeight    =   10950
   ClientLeft      =   225
   ClientTop       =   555
   ClientWidth     =   15120
   LinkTopic       =   "Form1"
   ScaleHeight     =   10950
   ScaleWidth      =   15120
   StartUpPosition =   3  '窗口缺省
   Begin VB.Timer Timerplug 
      Enabled         =   0   'False
      Interval        =   5000
      Left            =   14640
      Top             =   9960
   End
   Begin VB.Timer Timercollector 
      Enabled         =   0   'False
      Interval        =   5000
      Left            =   14520
      Top             =   9360
   End
   Begin MSAdodcLib.Adodc collector 
      Height          =   330
      Index           =   0
      Left            =   7200
      Top             =   4440
      Width           =   3015
      _ExtentX        =   5318
      _ExtentY        =   582
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
      Caption         =   "collector"
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
   Begin VB.Timer Timerarea 
      Enabled         =   0   'False
      Interval        =   500
      Left            =   14520
      Top             =   8520
   End
   Begin MSAdodcLib.Adodc area 
      Height          =   495
      Index           =   0
      Left            =   5160
      Top             =   9480
      Visible         =   0   'False
      Width           =   2295
      _ExtentX        =   4048
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
      Caption         =   "area"
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
   Begin VB.Timer Timerbuilding 
      Enabled         =   0   'False
      Interval        =   5000
      Left            =   14400
      Top             =   7680
   End
   Begin MSAdodcLib.Adodc buildingactive 
      Height          =   615
      Index           =   0
      Left            =   4440
      Top             =   7440
      Visible         =   0   'False
      Width           =   4815
      _ExtentX        =   8493
      _ExtentY        =   1085
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
      Caption         =   "buildingactive"
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
   Begin MSDataGridLib.DataGrid DataGrid3 
      Bindings        =   "smart3.frx":0000
      Height          =   1215
      Left            =   840
      TabIndex        =   12
      Top             =   9480
      Width           =   13455
      _ExtentX        =   23733
      _ExtentY        =   2143
      _Version        =   393216
      HeadLines       =   1
      RowHeight       =   15
      BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "宋体"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "宋体"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ColumnCount     =   2
      BeginProperty Column00 
         DataField       =   ""
         Caption         =   ""
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   2052
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column01 
         DataField       =   ""
         Caption         =   ""
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   2052
            SubFormatType   =   0
         EndProperty
      EndProperty
      SplitCount      =   1
      BeginProperty Split0 
         BeginProperty Column00 
         EndProperty
         BeginProperty Column01 
         EndProperty
      EndProperty
   End
   Begin MSAdodcLib.Adodc TotalTemp 
      Height          =   495
      Left            =   720
      Top             =   7800
      Visible         =   0   'False
      Width           =   2055
      _ExtentX        =   3625
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
      Caption         =   "totaltemp"
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
   Begin MSAdodcLib.Adodc test 
      Height          =   975
      Left            =   720
      Top             =   9600
      Visible         =   0   'False
      Width           =   2535
      _ExtentX        =   4471
      _ExtentY        =   1720
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
   Begin MSAdodcLib.Adodc TotalData 
      Height          =   855
      Left            =   600
      Top             =   8520
      Visible         =   0   'False
      Width           =   2175
      _ExtentX        =   3836
      _ExtentY        =   1508
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
      Caption         =   "totaldata"
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
   Begin VB.Timer Timer3 
      Interval        =   5000
      Left            =   240
      Top             =   8400
   End
   Begin MSAdodcLib.Adodc Adodc2 
      Height          =   615
      Left            =   5040
      Top             =   6240
      Visible         =   0   'False
      Width           =   5415
      _ExtentX        =   9551
      _ExtentY        =   1085
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
   Begin VB.Frame Frame4 
      Caption         =   "目录导航"
      Height          =   7575
      Left            =   240
      TabIndex        =   3
      Top             =   120
      Width           =   2895
      Begin MSComctlLib.TreeView TreeView1 
         Height          =   6735
         Left            =   240
         TabIndex        =   4
         Top             =   600
         Width           =   2535
         _ExtentX        =   4471
         _ExtentY        =   11880
         _Version        =   393217
         Style           =   7
         Appearance      =   1
      End
   End
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   495
      Left            =   6240
      Top             =   2520
      Visible         =   0   'False
      Width           =   2895
      _ExtentX        =   5106
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
   Begin VB.Frame Frame3 
      Caption         =   "温度、湿度"
      Height          =   3255
      Left            =   3120
      TabIndex        =   2
      Top             =   7560
      Width           =   11175
      Begin MSDataGridLib.DataGrid DataGrid2 
         Bindings        =   "smart3.frx":0016
         Height          =   1575
         Left            =   480
         TabIndex        =   6
         Top             =   600
         Width           =   9735
         _ExtentX        =   17171
         _ExtentY        =   2778
         _Version        =   393216
         HeadLines       =   1
         RowHeight       =   15
         BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "宋体"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "宋体"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ColumnCount     =   2
         BeginProperty Column00 
            DataField       =   ""
            Caption         =   ""
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   0
               Format          =   ""
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   2052
               SubFormatType   =   0
            EndProperty
         EndProperty
         BeginProperty Column01 
            DataField       =   ""
            Caption         =   ""
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   0
               Format          =   ""
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   2052
               SubFormatType   =   0
            EndProperty
         EndProperty
         SplitCount      =   1
         BeginProperty Split0 
            BeginProperty Column00 
            EndProperty
            BeginProperty Column01 
            EndProperty
         EndProperty
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "插座信息显示区"
      Height          =   4575
      Left            =   3000
      TabIndex        =   1
      Top             =   -240
      Width           =   23895
      Begin MSChart20Lib.MSChart ACTIVEENERGYChart 
         Height          =   2415
         Left            =   -120
         OleObjectBlob   =   "smart3.frx":002E
         TabIndex        =   9
         Top             =   360
         Width           =   4575
      End
      Begin MSChart20Lib.MSChart VOLTAGEchart 
         Height          =   2055
         Left            =   480
         OleObjectBlob   =   "smart3.frx":237B
         TabIndex        =   7
         Top             =   2640
         Visible         =   0   'False
         Width           =   4455
      End
      Begin MSChart20Lib.MSChart CURRENTChart 
         Height          =   2295
         Left            =   4800
         OleObjectBlob   =   "smart3.frx":5E2C
         TabIndex        =   8
         Top             =   2520
         Visible         =   0   'False
         Width           =   4215
      End
      Begin MSChart20Lib.MSChart TEMPChart 
         Height          =   2415
         Left            =   4200
         OleObjectBlob   =   "smart3.frx":98DD
         TabIndex        =   10
         Top             =   360
         Width           =   4575
      End
      Begin MSChart20Lib.MSChart HUMIDITYChart 
         Height          =   2415
         Left            =   8520
         OleObjectBlob   =   "smart3.frx":BC2A
         TabIndex        =   11
         Top             =   360
         Width           =   4575
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "电量、功率、电流、电压"
      Height          =   2535
      Left            =   3240
      TabIndex        =   0
      Top             =   4800
      Width           =   19455
      Begin MSAdodcLib.Adodc plug 
         Height          =   495
         Index           =   0
         Left            =   7440
         Top             =   600
         Width           =   1935
         _ExtentX        =   3413
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
         Caption         =   "plug"
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
      Begin MSDataGridLib.DataGrid DataGrid1 
         Bindings        =   "smart3.frx":DF77
         Height          =   2895
         Left            =   120
         TabIndex        =   5
         Top             =   240
         Width           =   15255
         _ExtentX        =   26908
         _ExtentY        =   5106
         _Version        =   393216
         HeadLines       =   1
         RowHeight       =   15
         BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "宋体"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "宋体"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ColumnCount     =   2
         BeginProperty Column00 
            DataField       =   ""
            Caption         =   ""
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   0
               Format          =   ""
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   2052
               SubFormatType   =   0
            EndProperty
         EndProperty
         BeginProperty Column01 
            DataField       =   ""
            Caption         =   ""
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   0
               Format          =   ""
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   2052
               SubFormatType   =   0
            EndProperty
         EndProperty
         SplitCount      =   1
         BeginProperty Split0 
            BeginProperty Column00 
            EndProperty
            BeginProperty Column01 
            EndProperty
         EndProperty
      End
   End
   Begin MSAdodcLib.Adodc buildingactive 
      Height          =   615
      Index           =   1
      Left            =   120
      Top             =   8040
      Visible         =   0   'False
      Width           =   2415
      _ExtentX        =   4260
      _ExtentY        =   1085
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
      Caption         =   "buildingactive"
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
   Begin MSAdodcLib.Adodc area 
      Height          =   495
      Index           =   1
      Left            =   120
      Top             =   10320
      Visible         =   0   'False
      Width           =   2295
      _ExtentX        =   4048
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
      Caption         =   "area"
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
   Begin MSAdodcLib.Adodc collector 
      Height          =   330
      Index           =   1
      Left            =   4080
      Top             =   4440
      Width           =   3015
      _ExtentX        =   5318
      _ExtentY        =   582
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
      Caption         =   "collector"
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
   Begin MSAdodcLib.Adodc plug 
      Height          =   495
      Index           =   1
      Left            =   11280
      Top             =   4440
      Width           =   1935
      _ExtentX        =   3413
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
      Caption         =   "plug"
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
End
Attribute VB_Name = "smartplus"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Dim i As Single







Private Sub Form_Initialize()
Adodc1.ConnectionString = "Provider=SQLOLEDB.1;Persist Security Info=False;User ID=sa;Password=ahjzu;Initial Catalog=master;Data Source=219.231.15.100"

Adodc1.CommandType = adCmdText
TotalData.ConnectionString = "Provider=SQLOLEDB.1;Persist Security Info=False;User ID=sa;Password=ahjzu;Initial Catalog=master;Data Source=219.231.15.100" '全部插座信息表的ado控件
TotalTemp.ConnectionString = "Provider=SQLOLEDB.1;Persist Security Info=False;User ID=sa;Password=ahjzu;Initial Catalog=master;Data Source=219.231.15.100" '全部插座信息表的ado控件
TotalData.CommandType = adCmdText
TotalTemp.CommandType = adCmdText
buildingactive(0).ConnectionString = "Provider=SQLOLEDB.1;Persist Security Info=False;User ID=sa;Password=ahjzu;Initial Catalog=master;Data Source=219.231.15.100"
buildingactive(0).CommandType = adCmdText
buildingactive(1).ConnectionString = "Provider=SQLOLEDB.1;Persist Security Info=False;User ID=sa;Password=ahjzu;Initial Catalog=master;Data Source=219.231.15.100"
buildingactive(1).CommandType = adCmdText
area(0).ConnectionString = "Provider=SQLOLEDB.1;Persist Security Info=False;User ID=sa;Password=ahjzu;Initial Catalog=master;Data Source=219.231.15.100"
area(0).CommandType = adCmdText
area(1).ConnectionString = "Provider=SQLOLEDB.1;Persist Security Info=False;User ID=sa;Password=ahjzu;Initial Catalog=master;Data Source=219.231.15.100"
area(1).CommandType = adCmdText
collector(0).ConnectionString = "Provider=SQLOLEDB.1;Persist Security Info=False;User ID=sa;Password=ahjzu;Initial Catalog=master;Data Source=219.231.15.100"
collector(0).CommandType = adCmdText
collector(1).ConnectionString = "Provider=SQLOLEDB.1;Persist Security Info=False;User ID=sa;Password=ahjzu;Initial Catalog=master;Data Source=219.231.15.100"
collector(1).CommandType = adCmdText

plug(0).ConnectionString = "Provider=SQLOLEDB.1;Persist Security Info=False;User ID=sa;Password=ahjzu;Initial Catalog=master;Data Source=219.231.15.100"
plug(0).CommandType = adCmdText
plug(1).ConnectionString = "Provider=SQLOLEDB.1;Persist Security Info=False;User ID=sa;Password=ahjzu;Initial Catalog=master;Data Source=219.231.15.100"
plug(1).CommandType = adCmdText
'buildingactive.Refresh
'test.ConnectionString = "Provider=SQLOLEDB.1;Persist Security Info=False;User ID=sa;Password=ahjzu;Initial Catalog=master;Data Source=219.231.15.100" '全部插座信息表的ado控件
'
'test.CommandType = adCmdText
Adodc2.ConnectionString = "Provider=SQLOLEDB.1;Persist Security Info=False;User ID=sa;Password=ahjzu;Initial Catalog=master;Data Source=219.231.15.100"

Adodc2.CommandType = adCmdText
TotalData.RecordSource = "Select top 100 ACTIVEPOWER,REACTIVEPOWER,VOLTAGE,[CURRENT],ACTIVEENERGY,processTime,DBTIME,MAC from analyze_electrictable order by DBTIME desc"
TotalTemp.RecordSource = "select top 100 TEMP,HUMIDITY,MAC,COLLECTIP,DBTIME from analyze_humituretable order by DBTIME desc"
  
  TotalData.Refresh
  TotalTemp.Refresh

'Adodc2.RecordSource = "select top 10 TEMP,HUMIDITY,MAC,COLLECTIP,DBTIME from analyze_humituretable order by DBTIME desc"


'Adodc2.Refresh

VOLTAGEchart.chartType = VtChChartType2dXY '电压二维散点图，显示一条曲线
CURRENTChart.chartType = VtChChartType2dXY '电流二维散点图，显示一条曲线
ACTIVEENERGYChart.chartType = VtChChartType2dXY '电量二维散点图，显示一条曲线
TEMPChart.chartType = VtChChartType2dXY '温度二维散点图，显示一条曲线
With ACTIVEENERGYChart '初始化电量曲线样式
 .chartType = VtChChartType2dXY '二维散点图，只能显示一条曲线注意本语句的位置，如果放在最后X轴坐标将显示成小数而非时间格式
        '图上只画条曲线
        .ColumnCount = 2
        .Plot.Axis(VtChAxisIdX).ValueScale.Auto = False
.Plot.Axis(VtChAxisIdY).ValueScale.Auto = False
 '// 设置最大值
     .Plot.Axis(VtChAxisIdY).ValueScale.Maximum = 1
       '// 设置最小值
      .Plot.Axis(VtChAxisIdY).ValueScale.Minimum = 0
  .Plot.Axis(VtChAxisIdX).ValueScale.Minimum = 0

     '//设置图表标题
      .Title.Text = "电量汇总"


 .Plot.Axis(VtChAxisIdX).ValueScale.MajorDivision = 3 'X轴主要网格数量
    .Plot.Axis(VtChAxisIdY).ValueScale.MajorDivision = 6 'Y轴主要网格数量
    .Plot.Axis(VtChAxisIdX).ValueScale.MinorDivision = 0 'X轴次要网格数量
    .Plot.Axis(VtChAxisIdY).ValueScale.MinorDivision = 0 'Y轴次要网格数量
 .Plot.Axis(VtChAxisIdX).Labels(1).Format = "hh:mm:ss"
    .Plot.Axis(VtChAxisIdX).AxisGrid.MajorPen.Style = VtPenStyleNull '//X轴网格不显示
   .Plot.Axis(VtChAxisIdY).AxisGrid.MajorPen.Style = VtPenStyleDotted '//Y轴网格为实线
   .Plot.AutoLayout = False   '//改为手动设置大小
   .Plot.UniformAxis = False '//指定图表的所有值坐标轴的单位刻度不一致(X,Y轴坐标不需要一致).
End With
With CURRENTChart '初始化电流曲线样式
 .chartType = VtChChartType2dXY '二维散点图，只能显示一条曲线注意本语句的位置，如果放在最后X轴坐标将显示成小数而非时间格式
        '图上只画条曲线
        .ColumnCount = 2
        .Plot.Axis(VtChAxisIdX).ValueScale.Auto = False
.Plot.Axis(VtChAxisIdY).ValueScale.Auto = False
 '// 设置最大值
     .Plot.Axis(VtChAxisIdY).ValueScale.Maximum = 1
       '// 设置最小值
      .Plot.Axis(VtChAxisIdY).ValueScale.Minimum = 0
  .Plot.Axis(VtChAxisIdX).ValueScale.Minimum = 0

     '//设置图表标题
      .Title.Text = "电流"


 .Plot.Axis(VtChAxisIdX).ValueScale.MajorDivision = 3 'X轴主要网格数量
    .Plot.Axis(VtChAxisIdY).ValueScale.MajorDivision = 6 'Y轴主要网格数量
    .Plot.Axis(VtChAxisIdX).ValueScale.MinorDivision = 0 'X轴次要网格数量
    .Plot.Axis(VtChAxisIdY).ValueScale.MinorDivision = 0 'Y轴次要网格数量
 .Plot.Axis(VtChAxisIdX).Labels(1).Format = "hh:mm:ss"
    .Plot.Axis(VtChAxisIdX).AxisGrid.MajorPen.Style = VtPenStyleNull '//X轴网格不显示
   .Plot.Axis(VtChAxisIdY).AxisGrid.MajorPen.Style = VtPenStyleDotted '//Y轴网格为实线
   .Plot.AutoLayout = False   '//改为手动设置大小
   .Plot.UniformAxis = False '//指定图表的所有值坐标轴的单位刻度不一致(X,Y轴坐标不需要一致).
End With
With VOLTAGEchart '初始化电压曲线样式
 .chartType = VtChChartType2dXY '二维散点图，只能显示一条曲线注意本语句的位置，如果放在最后X轴坐标将显示成小数而非时间格式
        '图上只画条曲线
        .ColumnCount = 2
        .Plot.Axis(VtChAxisIdX).ValueScale.Auto = False
.Plot.Axis(VtChAxisIdY).ValueScale.Auto = False
 '// 设置最大值
     .Plot.Axis(VtChAxisIdY).ValueScale.Maximum = 1
       '// 设置最小值
      .Plot.Axis(VtChAxisIdY).ValueScale.Minimum = 0
  .Plot.Axis(VtChAxisIdX).ValueScale.Minimum = 0

     '//设置图表标题
      .Title.Text = "电压"


 .Plot.Axis(VtChAxisIdX).ValueScale.MajorDivision = 3 'X轴主要网格数量
    .Plot.Axis(VtChAxisIdY).ValueScale.MajorDivision = 6 'Y轴主要网格数量
    .Plot.Axis(VtChAxisIdX).ValueScale.MinorDivision = 0 'X轴次要网格数量
    .Plot.Axis(VtChAxisIdY).ValueScale.MinorDivision = 0 'Y轴次要网格数量
 .Plot.Axis(VtChAxisIdX).Labels(1).Format = "hh:mm:ss"
    .Plot.Axis(VtChAxisIdX).AxisGrid.MajorPen.Style = VtPenStyleNull '//X轴网格不显示
   .Plot.Axis(VtChAxisIdY).AxisGrid.MajorPen.Style = VtPenStyleDotted '//Y轴网格为实线
   .Plot.AutoLayout = False   '//改为手动设置大小
   .Plot.UniformAxis = False '//指定图表的所有值坐标轴的单位刻度不一致(X,Y轴坐标不需要一致).
End With


With TEMPChart '初始化温度曲线样式
 .chartType = VtChChartType2dXY '二维散点图，只能显示一条曲线注意本语句的位置，如果放在最后X轴坐标将显示成小数而非时间格式
        '图上只画条曲线
        .ColumnCount = 2
        .Plot.Axis(VtChAxisIdX).ValueScale.Auto = False
.Plot.Axis(VtChAxisIdY).ValueScale.Auto = False
 '// 设置最大值
     .Plot.Axis(VtChAxisIdY).ValueScale.Maximum = 1
       '// 设置最小值
      .Plot.Axis(VtChAxisIdY).ValueScale.Minimum = 0
  .Plot.Axis(VtChAxisIdX).ValueScale.Minimum = 0

     '//设置图表标题
      .Title.Text = "温度平均值"


 .Plot.Axis(VtChAxisIdX).ValueScale.MajorDivision = 3 'X轴主要网格数量
    .Plot.Axis(VtChAxisIdY).ValueScale.MajorDivision = 6 'Y轴主要网格数量
    .Plot.Axis(VtChAxisIdX).ValueScale.MinorDivision = 0 'X轴次要网格数量
    .Plot.Axis(VtChAxisIdY).ValueScale.MinorDivision = 0 'Y轴次要网格数量
 .Plot.Axis(VtChAxisIdX).Labels(1).Format = "hh:mm:ss"
    .Plot.Axis(VtChAxisIdX).AxisGrid.MajorPen.Style = VtPenStyleNull '//X轴网格不显示
   .Plot.Axis(VtChAxisIdY).AxisGrid.MajorPen.Style = VtPenStyleDotted '//Y轴网格为实线
   .Plot.AutoLayout = False   '//改为手动设置大小
   .Plot.UniformAxis = False '//指定图表的所有值坐标轴的单位刻度不一致(X,Y轴坐标不需要一致).
End With
With HUMIDITYChart '初始化温度曲线样式
 .chartType = VtChChartType2dXY '二维散点图，只能显示一条曲线注意本语句的位置，如果放在最后X轴坐标将显示成小数而非时间格式
        '图上只画条曲线
        .ColumnCount = 2
        .Plot.Axis(VtChAxisIdX).ValueScale.Auto = False
.Plot.Axis(VtChAxisIdY).ValueScale.Auto = False
 '// 设置最大值
     .Plot.Axis(VtChAxisIdY).ValueScale.Maximum = 1
       '// 设置最小值
      .Plot.Axis(VtChAxisIdY).ValueScale.Minimum = 0
  .Plot.Axis(VtChAxisIdX).ValueScale.Minimum = 0

     '//设置图表标题
      .Title.Text = "湿度平均值"


 .Plot.Axis(VtChAxisIdX).ValueScale.MajorDivision = 3 'X轴主要网格数量
    .Plot.Axis(VtChAxisIdY).ValueScale.MajorDivision = 6 'Y轴主要网格数量
    .Plot.Axis(VtChAxisIdX).ValueScale.MinorDivision = 0 'X轴次要网格数量
    .Plot.Axis(VtChAxisIdY).ValueScale.MinorDivision = 0 'Y轴次要网格数量
 .Plot.Axis(VtChAxisIdX).Labels(1).Format = " hh:mm:ss"
    .Plot.Axis(VtChAxisIdX).AxisGrid.MajorPen.Style = VtPenStyleNull '//X轴网格不显示
   .Plot.Axis(VtChAxisIdY).AxisGrid.MajorPen.Style = VtPenStyleDotted '//Y轴网格为实线
   .Plot.AutoLayout = False   '//改为手动设置大小
   .Plot.UniformAxis = False '//指定图表的所有值坐标轴的单位刻度不一致(X,Y轴坐标不需要一致).
End With
t = buildtree()

End Sub

Private Sub Form_Load() '数据库连接
  
End Sub






Private Sub Timer3_Timer()
'Adodc2.Refresh
'Adodc1.Refresh

TotalData.Refresh
TotalTemp.Refresh
End Sub

Private Sub Timerarea_Timer()
area(0).Refresh
area(1).Refresh
End Sub

Private Sub Timerbuilding_Timer()
buildingactive(0).Refresh
buildingactive(1).Refresh

End Sub

Private Sub Timercollector_Timer()
collector(0).Refresh
collector(1).Refresh
End Sub

Private Sub Timerplug_Timer()
plug(0).Refresh
plug(1).Refresh
End Sub

'Dim MyData() As Double '电量数组
'Dim VOLTAGEData() As Double '电流数组
'Dim CURRENTData() As Double '电流数组
Private Sub TreeView1_NodeClick(ByVal Node As MSComctlLib.Node)   '读取数据中的对映表并显示电量等

Dim Mydata() As Double '电量数组
Dim VOLTAGEData() As Double '电流数组
Dim CURRENTData() As Double '电流数组
Dim TEMPData() As Double '温度数组
Dim HUMIDITYData() As Double '湿度数据
DblSec = 1.1574074074074E-05 '时间类型的数据1秒赋值给double类型变量时的值，即：1秒（时间类型）=1.1574074074074E-05（double类型）

Dim buildingR As New ADODB.Recordset '根目录节点的记录集对象（建筑物）
Dim plugR As New ADODB.Recordset '三级目录节点的记录集对象（插座）
Dim collectorR As New ADODB.Recordset '二级目录节点的记录集对象（采集器）
Dim AreaR As New ADODB.Recordset '一级目录节点的记录集对象（区域）


'*************************************************************************************************************************显示建筑物的汇总数据的相关曲线
buildingR.Open "select * from buildingInformation ", buildingactive(0).ConnectionString

        Do Until buildingR.EOF  '遍历二级采集器目录节点并显示
        strCodebuildingtagName = Trim(buildingR.Fields("tagName"))
         
If Node.Text = strCodebuildingtagName Then '选中建筑物节点
Timerbuilding.Enabled = True

Timerplug.Enabled = False
Timercollector = False
Timerarea = False
 VOLTAGEchart.Visible = False '不显示电压曲线表
 CURRENTChart.Visible = False '不显示电流曲线表
            'buildingactive(0).RecordSource = "Select top 5 sum(ACTIVEENERGY) as 电量汇总,DBTIME,areaInformation.BID from analyze_electrictable join collectorInformation join areaInformation on collectorInformation.AID=areaInformation.AID on COLLECTIP=collectorInformation.macAddress  group by DBTIME,areaInformation.BID Order by DBTIME desc"

           buildingactive(0).RecordSource = "Select sum(ACTIVEENERGY) as 电量汇总,DBTIME,areaInformation.BID from analyze_electrictable join collectorInformation join areaInformation on collectorInformation.AID=areaInformation.AID on COLLECTIP=collectorInformation.macAddress where DBTIME>dateadd(MINUTE,-10,GETDATE()) group by DBTIME,areaInformation.BID Order by DBTIME"
             buildingactive(1).RecordSource = "Select avg(cast(TEMP AS int)) as 温度平均值,avg(cast(HUMIDITY as int)) as 湿度平均值,DBTIME,areaInformation.BID from analyze_humituretable join collectorInformation join areaInformation on collectorInformation.AID=areaInformation.AID on COLLECTIP=collectorInformation.macAddress where DBTIME>dateadd(MINUTE,-10,GETDATE()) group by DBTIME,areaInformation.BID Order by DBTIME"
buildingactive(1).Refresh

          buildingactive(0).Refresh
'

If buildingactive(0).Recordset.RecordCount > 0 Then  '如果电量记录集有数据才显示电量之和曲线
ReDim Mydata(buildingactive(0).Recordset.RecordCount - 1, 1) '重定义采集器电量汇总之和的数组
ReDim HUMIDITYData(buildingactive(1).Recordset.RecordCount - 1, 1) As Double '重定义湿度平均值的数组

       
  For i = 0 To buildingactive(0).Recordset.RecordCount - 1
 TempSec = DateDiff("s", "0:0:0", TimeValue(buildingactive(0).Recordset(1).Value))
 Maxdate = buildingactive(0).Recordset(1).Value
    If i = 0 Then
       MaxSec = TempSec: MinSec = TempSec
       Maxdate0 = buildingactive(0).Recordset(1).Value
    End If
    If TempSec > MaxSec Then MaxSec = TempSec '求x轴最大值
    If TempSec < MinSec Then MinSec = TempSec '求x轴最小值
If Maxdate > Maxdata0 Then Maxdata0 = Maxdate
    TempACTIVEENERGYY = buildingactive(0).Recordset(0).Value '将电量第一个数设为y轴最小值

    If i = 0 Then

       MaxACTIVEENERGYY = TempACTIVEENERGYY: MinACTIVEENERGYY = TempACTIVEENERGYY
    End If

    If TempACTIVEENERGYY > MaxACTIVEENERGYY Then MaxACTIVEENERGYY = TempACTIVEENERGYY '求采集器电量之和Y轴最大值
    If TempACTIVEENERGYY < MinACTIVEENERGYY Then MinACTIVEENERGYY = TempACTIVEENERGYY '求采集器电量之和Y轴最小值

  Mydata(i, 0) = TimeValue(buildingactive(0).Recordset(1).Value) '第2列时间值存放电量X轴数据
    Mydata(i, 1) = buildingactive(0).Recordset(0).Value '第4列存入电量之和Y轴数据

       buildingactive(0).Recordset.MoveNext

    Next i
   
   With ACTIVEENERGYChart
  .Plot.Axis(VtChAxisIdY).ValueScale.Maximum = MaxACTIVEENERGYY + 5

    .Plot.Axis(VtChAxisIdY).ValueScale.Minimum = MinACTIVEENERGYY

           .Plot.SeriesCollection(1).Pen.VtColor.Set 0, 0, 255 '//第一条为理想曲线,设置为蓝色
        .Plot.SeriesCollection(1).Pen.Width = 20 '//设置为较细

          .Plot.Axis(VtChAxisIdX).ValueScale.Maximum = MaxSec * DblSec
           .Plot.Axis(VtChAxisIdX).ValueScale.Minimum = MinSec * DblSec
 .ChartData = Mydata

End With

'||||||||||||||||||||||||||||||||

ReDim TEMPData(buildingactive(1).Recordset.RecordCount - 1, 1) '重定义建筑物的温度数组
ReDim HUMIDITYData(buildingactive(1).Recordset.RecordCount - 1, 1) '重定义建筑物的温度数组


  For i = 0 To buildingactive(1).Recordset.RecordCount - 1
 TempSec = DateDiff("s", "0:0:0", TimeValue(buildingactive(1).Recordset(2).Value))
 
    If i = 0 Then
       MaxSec = TempSec: MinSec = TempSec
    End If
    If TempSec > MaxSec Then MaxSec = TempSec '求x轴最大值
    If TempSec < MinSec Then MinSec = TempSec '求x轴最小值
    
    Temptemp = buildingactive(1).Recordset(0).Value '将温度第一个数设为y轴最小值
    TempHUMIDITYY = buildingactive(1).Recordset(1).Value '将湿度第一个数设为y轴最小值
    If i = 0 Then
       MaxtempY = Temptemp: MintempY = Temptemp
      MaxHUMIDITYY = TempHUMIDITYY: MinHUMIDITYYY = TempHUMIDITYY

    End If
    
    If Temptemp > MaxtempY Then MaxtempY = Temptemp '求建筑物温度Y轴最大值
    If Temptemp < MintempY Then MintempY = Temptemp '求建筑物温度Y轴最小值
      If TempHUMIDITYY > MaxtempY Then MaxtempY = TempHUMIDITYY '求建筑物温度Y轴最大值
    If TempHUMIDITYY < MinHUMIDITYYY Then MinHUMIDITYYY = TempHUMIDITYY '求建筑物湿度Y轴最小值
    
  TEMPData(i, 0) = TimeValue(buildingactive(1).Recordset(2).Value) '第3列时间值存放建筑物X轴数据
    TEMPData(i, 1) = buildingactive(1).Recordset(0).Value '第1列存入温度Y轴数据

    HUMIDITYData(i, 0) = TimeValue(buildingactive(1).Recordset(2).Value) '第3列时间值存放X轴数据
    HUMIDITYData(i, 1) = buildingactive(1).Recordset(1).Value '第1列存入湿度平均值Y轴数据
   
       buildingactive(1).Recordset.MoveNext
     
    Next i
  
   With TEMPChart
  .Plot.Axis(VtChAxisIdY).ValueScale.Maximum = MaxtempY + 5
  
    .Plot.Axis(VtChAxisIdY).ValueScale.Minimum = MintempY
 
           .Plot.SeriesCollection(1).Pen.VtColor.Set 0, 0, 255 '//第一条为理想曲线,设置为蓝色
        .Plot.SeriesCollection(1).Pen.Width = 40 '//设置为较细
        
          .Plot.Axis(VtChAxisIdX).ValueScale.Maximum = MaxSec * DblSec
           .Plot.Axis(VtChAxisIdX).ValueScale.Minimum = MinSec * DblSec
 .ChartData = TEMPData

End With
      
       With HUMIDITYChart
  .Plot.Axis(VtChAxisIdY).ValueScale.Maximum = MaxHUMIDITYY + 5
  
    .Plot.Axis(VtChAxisIdY).ValueScale.Minimum = MinHUMIDITYY
 
           .Plot.SeriesCollection(1).Pen.VtColor.Set 0, 0, 255 '//第一条为理想曲线,设置为蓝色
        .Plot.SeriesCollection(1).Pen.Width = 40 '//设置为较细
        
          .Plot.Axis(VtChAxisIdX).ValueScale.Maximum = MaxSec * DblSec
           .Plot.Axis(VtChAxisIdX).ValueScale.Minimum = MinSec * DblSec
 .ChartData = HUMIDITYData

End With

'|||||||||||||||||||||||||||||||||||||||||||||||



     End If '结束一个有数据的建筑物曲线显示

End If  '结束一个建筑物的曲线显示

     buildingR.MoveNext
        Loop
        buildingR.Close

'*************************************************************************************************************************end显示建筑物的汇总数据的相关曲线
'*************************************************************************************************************************显示区域的汇总数据的相关曲线
AreaR.Open "select * from areaInformation ", area(0).ConnectionString

        Do Until AreaR.EOF  '遍历二级采集器目录节点并显示
        strCodeAreatagName = Trim(AreaR.Fields("tagName"))
          'strCodeCOLLECTIP = AreaR.Fields("macAddress")
If Node.Text = strCodeAreatagName Then '选中区域节点
Timerarea.Enabled = True
Timerbuilding.Enabled = False

Timerplug.Enabled = False
Timercollector = False



 VOLTAGEchart.Visible = False '不显示电压曲线表
 CURRENTChart.Visible = False '不显示电流曲线表
       
        area(0).RecordSource = "Select sum(ACTIVEENERGY) as 电量汇总,DBTIME,collectorInformation.AID from analyze_electrictable join collectorInformation on COLLECTIP=collectorInformation.macAddress where DBTIME>dateadd(MINUTE,-10,GETDATE()) group by DBTIME,collectorInformation.AID Order by DBTIME"
        area(1).RecordSource = "Select avg(cast(TEMP AS int)) as 温度平均值,avg(cast(HUMIDITY as int)) as 湿度平均值,DBTIME,collectorInformation.AID from analyze_humituretable join collectorInformation on COLLECTIP=collectorInformation.macAddress  where  DBTIME>dateadd(MINUTE,-10,GETDATE()) group by DBTIME,collectorInformation.AID Order by DBTIME"
      
           'test.RecordSource = "Select sum(ACTIVEENERGY) as 电量汇总,DBTIME,collectorInformation.AID from analyze_electrictable join collectorInformation on COLLECTIP=collectorInformation.macAddress where DBTIME>dateadd(MINUTE,-10,GETDATE())group by DBTIME,COLLECTIP"
'test.Refresh
          area(1).Refresh

          area(0).Refresh
'

If area(0).Recordset.RecordCount > 0 Then  '如果记录集有数据才显示曲线
ReDim Mydata(area(0).Recordset.RecordCount - 1, 1) '重定义采集器电量汇总之和的数组
       
  For i = 0 To area(0).Recordset.RecordCount - 1
 TempSec = DateDiff("s", "0:0:0", TimeValue(area(0).Recordset(1).Value))

    If i = 0 Then
       MaxSec = TempSec: MinSec = TempSec
    End If
    If TempSec > MaxSec Then MaxSec = TempSec '求x轴最大值
    If TempSec < MinSec Then MinSec = TempSec '求x轴最小值

    TempACTIVEENERGYY = area(0).Recordset(0).Value '将电量第一个数设为y轴最小值

    If i = 0 Then

       MaxACTIVEENERGYY = TempACTIVEENERGYY: MinACTIVEENERGYY = TempACTIVEENERGYY
    End If

    If TempACTIVEENERGYY > MaxACTIVEENERGYY Then MaxACTIVEENERGYY = TempACTIVEENERGYY '求采集器电量之和Y轴最大值
    If TempACTIVEENERGYY < MinACTIVEENERGYY Then MinACTIVEENERGYY = TempACTIVEENERGYY '求采集器电量之和Y轴最小值

  Mydata(i, 0) = TimeValue(area(0).Recordset(1).Value) '第2列时间值存放电量X轴数据
    Mydata(i, 1) = area(0).Recordset(0).Value '第4列存入电量之和Y轴数据

       area(0).Recordset.MoveNext

    Next i
   
   With ACTIVEENERGYChart
  .Plot.Axis(VtChAxisIdY).ValueScale.Maximum = MaxACTIVEENERGYY + 5

    .Plot.Axis(VtChAxisIdY).ValueScale.Minimum = MinACTIVEENERGYY

           .Plot.SeriesCollection(1).Pen.VtColor.Set 0, 0, 255 '//第一条为理想曲线,设置为蓝色
        .Plot.SeriesCollection(1).Pen.Width = 20 '//设置为较细

          .Plot.Axis(VtChAxisIdX).ValueScale.Maximum = MaxSec * DblSec
           .Plot.Axis(VtChAxisIdX).ValueScale.Minimum = MinSec * DblSec
 .ChartData = Mydata

End With


ReDim TEMPData(area(1).Recordset.RecordCount - 1, 1) '重定义建筑物的温度数组
ReDim HUMIDITYData(area(1).Recordset.RecordCount - 1, 1) '重定义建筑物的温度数组


  For i = 0 To area(1).Recordset.RecordCount - 1
 TempSec = DateDiff("s", "0:0:0", TimeValue(area(1).Recordset(2).Value))
 
    If i = 0 Then
       MaxSec = TempSec: MinSec = TempSec
    End If
    If TempSec > MaxSec Then MaxSec = TempSec '求x轴最大值
    If TempSec < MinSec Then MinSec = TempSec '求x轴最小值
    
    Temptemp = area(1).Recordset(0).Value '将温度第一个数设为y轴最小值
    TempHUMIDITYY = area(1).Recordset(1).Value '将湿度第一个数设为y轴最小值
    If i = 0 Then
       MaxtempY = Temptemp: MintempY = Temptemp
      MaxHUMIDITYY = TempHUMIDITYY: MinHUMIDITYYY = TempHUMIDITYY

    End If
    
    If Temptemp > MaxtempY Then MaxtempY = Temptemp '求采集器温度Y轴最大值
    If Temptemp < MintempY Then MintempY = Temptemp '求采集器温度Y轴最小值
      If TempHUMIDITYY > MaxtempY Then MaxtempY = TempHUMIDITYY '求采集器温度Y轴最大值
    If TempHUMIDITYY < MinHUMIDITYYY Then MinHUMIDITYYY = TempHUMIDITYY '求采集器湿度Y轴最小值
    
  TEMPData(i, 0) = TimeValue(area(1).Recordset(2).Value) '第3列时间值存放采集器X轴数据
    TEMPData(i, 1) = area(1).Recordset(0).Value '第1列存入温度Y轴数据

    HUMIDITYData(i, 0) = TimeValue(area(1).Recordset(2).Value) '第3列时间值存放X轴数据
    HUMIDITYData(i, 1) = area(1).Recordset(1).Value '第1列存入湿度平均值Y轴数据
   
       area(1).Recordset.MoveNext
     
    Next i
  
   With TEMPChart
  .Plot.Axis(VtChAxisIdY).ValueScale.Maximum = MaxtempY + 5
  
    .Plot.Axis(VtChAxisIdY).ValueScale.Minimum = MintempY
 
           .Plot.SeriesCollection(1).Pen.VtColor.Set 0, 0, 255 '//第一条为理想曲线,设置为蓝色
        .Plot.SeriesCollection(1).Pen.Width = 40 '//设置为较细
        
          .Plot.Axis(VtChAxisIdX).ValueScale.Maximum = MaxSec * DblSec
           .Plot.Axis(VtChAxisIdX).ValueScale.Minimum = MinSec * DblSec
 .ChartData = TEMPData

End With
      
       With HUMIDITYChart
  .Plot.Axis(VtChAxisIdY).ValueScale.Maximum = MaxHUMIDITYY + 5
  
    .Plot.Axis(VtChAxisIdY).ValueScale.Minimum = MinHUMIDITYY
 
           .Plot.SeriesCollection(1).Pen.VtColor.Set 0, 0, 255 '//第一条为理想曲线,设置为蓝色
        .Plot.SeriesCollection(1).Pen.Width = 40 '//设置为较细
        
          .Plot.Axis(VtChAxisIdX).ValueScale.Maximum = MaxSec * DblSec
           .Plot.Axis(VtChAxisIdX).ValueScale.Minimum = MinSec * DblSec
 .ChartData = HUMIDITYData

End With







     End If '结束一个有数据的区域曲线显示

End If  '结束一个区域的曲线显示

     AreaR.MoveNext
        Loop
        AreaR.Close

'*************************************************************************************************************************end显示区域的汇总数据的相关曲线


'$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$显示采集器的汇总数据的相关曲线
collectorR.Open "select * from collectorInformation  ", collector(0).ConnectionString

        Do Until collectorR.EOF  '遍历二级采集器目录节点并显示
        strCodecollectortagName = Trim(collectorR.Fields("tagName"))
       ' strCodeCOLLECTIP = Trim(collectorR.Fields("macAddress"))
          strCodeCOLLECTIP = collectorR.Fields("macAddress")
If Node.Text = strCodecollectortagName Then '选中采集器节点
Timercollector.Enabled = True
Timerbuilding.Enabled = False
Timerarea.Enabled = False


Timerplug.Enabled = False

 VOLTAGEchart.Visible = False '不显示电压曲线表
 CURRENTChart.Visible = False '不显示电流曲线表
       
             collector(0).RecordSource = "Select sum([CURRENT]) as 电流汇总,DBTIME,COLLECTIP,sum(VOLTAGE) as 电压汇总,sum(ACTIVEENERGY) as 电量汇总 from analyze_electrictable where COLLECTIP='" & strCodeCOLLECTIP & "' and DBTIME>dateadd(MINUTE,-10,GETDATE()) group by DBTIME,COLLECTIP Order by DBTIME"
             collector(1).RecordSource = "Select avg(cast(TEMP AS int)) as 温度平均值,avg(cast(HUMIDITY as int)) as 湿度平均值,DBTIME,COLLECTIP from analyze_humituretable where COLLECTIP='" & strCodeCOLLECTIP & "' and DBTIME>dateadd(MINUTE,-10,GETDATE())group by DBTIME,COLLECTIP"

          collector(0).Refresh
          collector(1).Refresh
'           test.RecordSource = "Select sum([CURRENT]) as 电流汇总,DBTIME,COLLECTIP,sum(VOLTAGE)as 电压汇总,sum(ACTIVEENERGY)as 电量汇总 from analyze_electrictable where COLLECTIP='" & strCodeCOLLECTIP & "' group by DBTIME,COLLECTIP Order by DBTIME"
'           test.Refresh
        
        
           
        
If collector(0).Recordset.RecordCount > 0 Then  '如果记录集有数据才显示曲线
ReDim Mydata(collector(0).Recordset.RecordCount - 1, 1) '重定义采集器电量汇总之和的数组
         ReDim CURRENTData(collector(0).Recordset.RecordCount - 1, 1) '重定义采集器电流汇总之和的数组
 ReDim VOLTAGEData(collector(0).Recordset.RecordCount - 1, 1) '重定义采集器电压之和数组
  For i = 0 To collector(0).Recordset.RecordCount - 1
 TempSec = DateDiff("s", "0:0:0", TimeValue(collector(0).Recordset(1).Value))
 
    If i = 0 Then
       MaxSec = TempSec: MinSec = TempSec
    End If
    If TempSec > MaxSec Then MaxSec = TempSec '求x轴最大值
    If TempSec < MinSec Then MinSec = TempSec '求x轴最小值
    
    TempACTIVEENERGYY = collector(0).Recordset(4).Value '将电量第一个数设为y轴最小值
     TempCurrentY = collector(0).Recordset(0).Value '将电流第一个数设为y轴最小值
    TempVoltageY = collector(0).Recordset(3).Value '将电压第一个数设为y轴最小值
    If i = 0 Then
       MaxY = TempCurrentY: MinY = TempCurrentY
       MaxVoltageY = TempVoltageY: MinVoltageY = TempVoltageY
       MaxACTIVEENERGYY = TempACTIVEENERGYY: MinACTIVEENERGYY = TempACTIVEENERGYY
    End If
    
    If TempACTIVEENERGYY > MaxACTIVEENERGYY Then MaxACTIVEENERGYY = TempACTIVEENERGYY '求采集器电量之和Y轴最大值
    If TempACTIVEENERGYY < MinACTIVEENERGYY Then MinACTIVEENERGYY = TempACTIVEENERGYY '求采集器电量之和Y轴最小值
    If TempCurrentY > MaxY Then MaxY = TempCurrentY '求采集器电流之和Y轴最大值
    If TempCurrentY < MinY Then MinY = TempCurrentY '求采集器电流之和Y轴最小值
    
    If TempVoltageY > MaxVoltageY Then MaxVoltageY = TempVoltageY '求采集器电压之和Y轴最大值
    If TempVoltageY < MinVoltageY Then MinVoltageY = TempVoltageY '求采集器电压之和Y轴最小值
    
    
 Mydata(i, 0) = TimeValue(collector(0).Recordset(1).Value) '第2列时间值存放电量X轴数据
 'Mydata(i, 0) = collector(0).Recordset(1).Value '第2列时间值存放电量X轴数据


    Mydata(i, 1) = collector(0).Recordset(4).Value '第4列存入电量之和Y轴数据
 CURRENTData(i, 0) = TimeValue(collector(0).Recordset(1).Value) '第2列时间值存放电流X轴数据
    CURRENTData(i, 1) = collector(0).Recordset(0).Value '第1列存入电流之和Y轴数据
      VOLTAGEData(i, 0) = TimeValue(collector(0).Recordset(1).Value) '第2列时间值存放电压X轴数据
    VOLTAGEData(i, 1) = collector(0).Recordset(3).Value '第4列存入电压之和Y轴数据
       
       collector(0).Recordset.MoveNext
     
    Next i
   
      'collector(0).Recordset.Close
   With ACTIVEENERGYChart
  .Plot.Axis(VtChAxisIdY).ValueScale.Maximum = MaxACTIVEENERGYY + 5
  
    .Plot.Axis(VtChAxisIdY).ValueScale.Minimum = MinACTIVEENERGYY
 
           .Plot.SeriesCollection(1).Pen.VtColor.Set 0, 0, 255 '//第一条为理想曲线,设置为蓝色
        .Plot.SeriesCollection(1).Pen.Width = 10 '//设置为较细
        
          .Plot.Axis(VtChAxisIdX).ValueScale.Maximum = MaxSec * DblSec
           .Plot.Axis(VtChAxisIdX).ValueScale.Minimum = MinSec * DblSec
 .ChartData = Mydata

End With
      
      
      
      
With CURRENTChart
  .Plot.Axis(VtChAxisIdY).ValueScale.Maximum = MaxX
  
    .Plot.Axis(VtChAxisIdY).ValueScale.Minimum = MinY
 ' .Plot.Axis(VtChAxisIdX).ValueScale.Minimum = 0
 
           .Plot.SeriesCollection(1).Pen.VtColor.Set 0, 0, 255 '//第一条为理想曲线,设置为蓝色
        .Plot.SeriesCollection(1).Pen.Width = 10 '//设置为较细
        
          .Plot.Axis(VtChAxisIdX).ValueScale.Maximum = MaxSec * DblSec
           .Plot.Axis(VtChAxisIdY).ValueScale.Maximum = MaxY + 5
 .Plot.Axis(VtChAxisIdX).ValueScale.Minimum = MinSec * DblSec
 .ChartData = CURRENTData

End With

With VOLTAGEchart
  .Plot.Axis(VtChAxisIdY).ValueScale.Maximum = MaxVoltageY + 5
  
    .Plot.Axis(VtChAxisIdY).ValueScale.Minimum = MinVoltageY
  '.Plot.Axis(VtChAxisIdX).ValueScale.Minimum = 0
 
           .Plot.SeriesCollection(1).Pen.VtColor.Set 0, 0, 255 '//第一条为理想曲线,设置为蓝色
        .Plot.SeriesCollection(1).Pen.Width = 10 '//设置为较细
        
          .Plot.Axis(VtChAxisIdX).ValueScale.Maximum = MaxSec * DblSec
           .Plot.Axis(VtChAxisIdX).ValueScale.Minimum = MinSec * DblSec
 .ChartData = VOLTAGEData

End With

'||||||||||||||||||||||||||||||||

ReDim TEMPData(collector(1).Recordset.RecordCount - 1, 1) '重定义采集器的温度数组
ReDim HUMIDITYData(collector(1).Recordset.RecordCount - 1, 1) '重定义采集器温度数组


  For i = 0 To collector(1).Recordset.RecordCount - 1
 TempSec = DateDiff("s", "0:0:0", TimeValue(collector(1).Recordset(2).Value))
 
    If i = 0 Then
       MaxSec = TempSec: MinSec = TempSec
    End If
    If TempSec > MaxSec Then MaxSec = TempSec '求x轴最大值
    If TempSec < MinSec Then MinSec = TempSec '求x轴最小值
    
    Temptemp = collector(1).Recordset(0).Value '将温度第一个数设为y轴最小值
    TempHUMIDITYY = collector(1).Recordset(1).Value '将湿度第一个数设为y轴最小值
    If i = 0 Then
       MaxtempY = Temptemp: MintempY = Temptemp
      MaxHUMIDITYY = TempHUMIDITYY: MinHUMIDITYYY = TempHUMIDITYY

    End If
    
    If Temptemp > MaxtempY Then MaxtempY = Temptemp '求区域温度Y轴最大值
    If Temptemp < MintempY Then MintempY = Temptemp '求区域温度Y轴最小值
      If TempHUMIDITYY > MaxtempY Then MaxtempY = TempHUMIDITYY '求区域温度Y轴最大值
    If TempHUMIDITYY < MinHUMIDITYYY Then MinHUMIDITYYY = TempHUMIDITYY '求区域湿度Y轴最小值
    
  TEMPData(i, 0) = TimeValue(collector(1).Recordset(2).Value) '第3列时间值存放区域X轴数据
    TEMPData(i, 1) = collector(1).Recordset(0).Value '第1列存入温度Y轴数据

    HUMIDITYData(i, 0) = TimeValue(collector(1).Recordset(2).Value) '第3列时间值存放X轴数据
    HUMIDITYData(i, 1) = collector(1).Recordset(1).Value   '第1列存入湿度平均值Y轴数据
   
       collector(1).Recordset.MoveNext
     
    Next i
  
   With TEMPChart
  .Plot.Axis(VtChAxisIdY).ValueScale.Maximum = MaxtempY + 5
  
    .Plot.Axis(VtChAxisIdY).ValueScale.Minimum = MintempY
 
           .Plot.SeriesCollection(1).Pen.VtColor.Set 0, 0, 255 '//第一条为理想曲线,设置为蓝色
        .Plot.SeriesCollection(1).Pen.Width = 40 '//设置为较细
        
          .Plot.Axis(VtChAxisIdX).ValueScale.Maximum = MaxSec * DblSec
           .Plot.Axis(VtChAxisIdX).ValueScale.Minimum = MinSec * DblSec
 .ChartData = TEMPData

End With
      
       With HUMIDITYChart
  .Plot.Axis(VtChAxisIdY).ValueScale.Maximum = MaxHUMIDITYY + 5
  
    .Plot.Axis(VtChAxisIdY).ValueScale.Minimum = MinHUMIDITYY
 
           .Plot.SeriesCollection(1).Pen.VtColor.Set 0, 0, 255 '//第一条为理想曲线,设置为蓝色
        .Plot.SeriesCollection(1).Pen.Width = 40 '//设置为较细
        
          .Plot.Axis(VtChAxisIdX).ValueScale.Maximum = MaxSec * DblSec
           .Plot.Axis(VtChAxisIdX).ValueScale.Minimum = MinSec * DblSec
 .ChartData = HUMIDITYData

End With




































'|||||||||||||||||||||||||||||||||||||||||||||||
     End If '结束一个有数据的采集器曲线显示
           
End If  '结束一个采集器的曲线显示
      
     collectorR.MoveNext
        Loop
        collectorR.Close
        
'$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$end显示采集器的汇总数据的相关曲线


Dim analyze_electrictableR As New ADODB.Recordset '插座电量等数据表的记录集对象
Dim LastTimeR As New ADODB.Recordset '用于取最后一条记录时间的记录对象
plugR.Open "select * from pluginInformation", plug(0).ConnectionString
        Do Until plugR.EOF  '遍历三级插座目录节点并显示
         
        strCodeplugtagName = Trim(plugR.Fields("tagName"))
        strCodeplugmacAddress = Trim(plugR.Fields("macAddress"))
If Node.Text = strCodeplugtagName Then   '如果选中插座级的节点
Timerplug.Enabled = True
Timercollector = False
Timerbuilding = False
Timerarea = False

 CURRENTChart.Visible = True
 VOLTAGEchart.Visible = True
 
 '*+++++++++++++++++++++++++++++++++++++++++++++获取插座最后一条记录的时间
 LastTimeR.Open "Select ACTIVEPOWER,REACTIVEPOWER,VOLTAGE,[CURRENT],ACTIVEENERGY,DBTIME,MAC from analyze_electrictable where MAC='" & strCodeplugmacAddress & "'Order by DBTIME", plug(0).ConnectionString
    Do Until LastTimeR.EOF
    pluglastTime = LastTimeR(5)
  LastTimeR.MoveNext
  Loop
 Print pluglastTime
 LastTimeR.Close
  '*+++++++++++++++++++++++++++++++++++++++++++++获取插座最后一条记录的时间

                plug(0).RecordSource = "Select ACTIVEPOWER,REACTIVEPOWER,VOLTAGE,[CURRENT],ACTIVEENERGY,DBTIME,MAC from analyze_electrictable where MAC='" & strCodeplugmacAddress & "'and DBTIME>dateadd(MINUTE,-10,GETDATE()) Order by DBTIME"


'Adodc2.RecordSource = "select TEMP,HUMIDITY,MAC,COLLECTIP,DBTIME from analyze_humituretable  where DBTIME>dateadd(MINUTE,-10,GETDATE())"  '显示近10分钟的插座温度，演示系统时要把系统时间设定为10-21 9：01：50，因为这是数据库记录的最后一条记录的插入时间
plug(1).RecordSource = "select TEMP,HUMIDITY,MAC,COLLECTIP,DBTIME from analyze_humituretable  where MAC='" & strCodeplugmacAddress & "' and DBTIME>dateadd(MINUTE,-10,GETDATE()) Order by DBTIME"  '显示近10分钟的插座温度，演示系统时要把系统时间设定为10-21 9：01：50，因为这是数据库记录的最后一条记录的插入时间

        plug(0).Refresh
        plug(1).Refresh
          


With ACTIVEENERGYChart
     .chartType = VtChChartType2dXY '二维散点图，只能显示一条曲线注意本语句的位置，如果放在最后X轴坐标将显示成小数而非时间格式
        '图上只画条曲线
        .ColumnCount = 2
       VOLTAGEchart.ColumnCount = 2
      
'        'X轴显示10个单位
'          .RowCount = plug(0).Recordset.RecordCount
        
          '设置XY轴?
.Plot.Axis(VtChAxisIdX).ValueScale.Auto = False
.Plot.Axis(VtChAxisIdY).ValueScale.Auto = False
VOLTAGEchart.Plot.Axis(VtChAxisIdX).ValueScale.Auto = False
VOLTAGEchart.Plot.Axis(VtChAxisIdY).ValueScale.Auto = False

 
  
   VOLTAGEchart.Plot.Axis(VtChAxisIdY).ValueScale.Minimum = 0
 VOLTAGEchart.Plot.Axis(VtChAxisIdX).ValueScale.Minimum = 0

      '//设置图表标题
      .Title.Text = "电量"
        VOLTAGEchart.Title.Text = "电压"

ReDim Mydata(plug(0).Recordset.RecordCount - 1, 1) '重定义电量数组
ReDim CURRENTData(plug(0).Recordset.RecordCount - 1, 1) '重定义电流数组
ReDim VOLTAGEData(plug(0).Recordset.RecordCount - 1, 1) '电压数组
For i = 0 To plug(0).Recordset.RecordCount - 1
  TempSec = DateDiff("s", "0:0:0", TimeValue(plug(0).Recordset(5).Value))
  'Print TempSec
If i = 0 Then
MaxSec = TempSec: MinSec = TempSec
End If
If TempSec > MaxSec Then MaxSec = TempSec '求最大值
If TempSec < MinSec Then MinSec = TempSec '求最小值

TempACTIVEENERGYY = plug(0).Recordset(4).Value '将电量第一个数设为y轴最小值

    If i = 0 Then

       MaxACTIVEENERGYY = TempACTIVEENERGYY: MinACTIVEENERGYY = TempACTIVEENERGYY
    End If

    If TempACTIVEENERGYY > MaxACTIVEENERGYY Then MaxACTIVEENERGYY = TempACTIVEENERGYY '求采集器电量之和Y轴最大值
    If TempACTIVEENERGYY < MinACTIVEENERGYY Then MinACTIVEENERGYY = TempACTIVEENERGYY '求采集器电量之和Y轴最小值


VOLTAGEData(i, 0) = TimeValue(plug(0).Recordset(5).Value) '第6列时间值存放电压X轴数据

CURRENTData(i, 0) = TimeValue(plug(0).Recordset(5).Value) '第6列时间值存放电流X轴数据
Mydata(i, 0) = TimeValue(plug(0).Recordset(5).Value) '第6列时间值存放X轴数据
Mydata(i, 1) = plug(0).Recordset(4).Value '第5列电量存放Y轴数据
'MyData(i, 2) = plug(0).Recordset(1).Value '第2列存放Y轴数据

VOLTAGEData(i, 1) = plug(0).Recordset(2).Value '第3列存入电压Y轴数据
CURRENTData(i, 1) = plug(0).Recordset(3).Value '第4列存入电流Y轴数据
         plug(0).Recordset.MoveNext

Next i
          
          
           .Plot.SeriesCollection(1).Pen.VtColor.Set 0, 0, 255 '//第一条为理想曲线,设置为蓝色
        .Plot.SeriesCollection(1).Pen.Width = 50 '//设置为较细
       
        
          .Plot.Axis(VtChAxisIdX).ValueScale.Maximum = MaxSec * DblSec
 .Plot.Axis(VtChAxisIdX).ValueScale.Minimum = MinSec * DblSec
  '.Plot.Axis(VtChAxisIdY).ValueScale.Maximum = 1
  
    .Plot.Axis(VtChAxisIdY).ValueScale.Maximum = MaxACTIVEENERGYY + 0.1
  .Plot.Axis(VtChAxisIdY).ValueScale.Minimum = MinACTIVEENERGYY
 .ChartData = Mydata '电量数据?
 
 '&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&
With VOLTAGEchart
  .Plot.Axis(VtChAxisIdY).ValueScale.Maximum = 300
  
    .Plot.Axis(VtChAxisIdY).ValueScale.Minimum = 0
  .Plot.Axis(VtChAxisIdX).ValueScale.Minimum = 0
 
           .Plot.SeriesCollection(1).Pen.VtColor.Set 0, 0, 255 '//第一条为理想曲线,设置为蓝色
        .Plot.SeriesCollection(1).Pen.Width = 10 '//设置为较细
        
       
        
        
          .Plot.Axis(VtChAxisIdX).ValueScale.Maximum = MaxSec * DblSec
 .Plot.Axis(VtChAxisIdX).ValueScale.Minimum = MinSec * DblSec

End With
With CURRENTChart
  .Plot.Axis(VtChAxisIdY).ValueScale.Maximum = 1
  
    .Plot.Axis(VtChAxisIdY).ValueScale.Minimum = 0
  .Plot.Axis(VtChAxisIdX).ValueScale.Minimum = 0
 
           .Plot.SeriesCollection(1).Pen.VtColor.Set 0, 0, 255 '//第一条为理想曲线,设置为蓝色
        .Plot.SeriesCollection(1).Pen.Width = 40 '//设置为较细
        
        
        
        
          .Plot.Axis(VtChAxisIdX).ValueScale.Maximum = MaxSec * DblSec
 .Plot.Axis(VtChAxisIdX).ValueScale.Minimum = MinSec * DblSec
 .ChartData = CURRENTData

End With

VOLTAGEchart.ChartData = VOLTAGEData '数据?
      End With



  MSChart1.Plot.UniformAxis = False

'#############################################################################插座级温度显示
If plug(1).Recordset.RecordCount > 0 Then  '如果记录集有数据才显示曲线
ReDim TEMPData(plug(1).Recordset.RecordCount - 1, 1) '重定义插座的温度数组
ReDim HUMIDITYData(plug(1).Recordset.RecordCount - 1, 1) '重定义插座的温度数组


  For i = 0 To plug(1).Recordset.RecordCount - 1
 TempSec = DateDiff("s", "0:0:0", TimeValue(plug(1).Recordset(4).Value))
 
    If i = 0 Then
       MaxSec = TempSec: MinSec = TempSec
    End If
    If TempSec > MaxSec Then MaxSec = TempSec '求x轴最大值
    If TempSec < MinSec Then MinSec = TempSec '求x轴最小值
    
    Temptemp = plug(1).Recordset(0).Value '将温度第一个数设为y轴最小值
    TempHUMIDITYY = plug(1).Recordset(1).Value '将湿度第一个数设为y轴最小值
    If i = 0 Then
       MaxtempY = Temptemp: MintempY = Temptemp
      MaxHUMIDITYY = TempHUMIDITYY: MinHUMIDITYYY = TempHUMIDITYY

    End If
    
    If Temptemp > MaxtempY Then MaxtempY = Temptemp '求插座温度Y轴最大值
    If Temptemp < MintempY Then MintempY = Temptemp '求插座温度Y轴最小值
      If TempHUMIDITYY > MaxtempY Then MaxtempY = TempHUMIDITYY '求插座温度Y轴最大值
    If TempHUMIDITYY < MinHUMIDITYYY Then MinHUMIDITYYY = TempHUMIDITYY '求插座湿度Y轴最小值
    
  TEMPData(i, 0) = TimeValue(plug(1).Recordset(4).Value) '第5列时间值存放电量X轴数据
    TEMPData(i, 1) = plug(1).Recordset(0).Value '第1列存入温度Y轴数据

    HUMIDITYData(i, 0) = TimeValue(plug(1).Recordset(4).Value) '第5列时间值存放电量X轴数据
    HUMIDITYData(i, 1) = plug(1).Recordset(1).Value '第1列存入温度Y轴数据
   
       plug(1).Recordset.MoveNext
     
    Next i
  
   With TEMPChart
       .Title.Text = "温度"
  .Plot.Axis(VtChAxisIdY).ValueScale.Maximum = MaxtempY + 5
  
    .Plot.Axis(VtChAxisIdY).ValueScale.Minimum = MintempY
 
           .Plot.SeriesCollection(1).Pen.VtColor.Set 0, 0, 255 '//第一条为理想曲线,设置为蓝色
        .Plot.SeriesCollection(1).Pen.Width = 40 '//设置为较细
        
          .Plot.Axis(VtChAxisIdX).ValueScale.Maximum = MaxSec * DblSec
           .Plot.Axis(VtChAxisIdX).ValueScale.Minimum = MinSec * DblSec
 .ChartData = TEMPData

End With
      
       With HUMIDITYChart
           .Title.Text = "湿度"
  .Plot.Axis(VtChAxisIdY).ValueScale.Maximum = MaxHUMIDITYY + 5
  
    .Plot.Axis(VtChAxisIdY).ValueScale.Minimum = MinHUMIDITYY
 
           .Plot.SeriesCollection(1).Pen.VtColor.Set 0, 0, 255 '//第一条为理想曲线,设置为蓝色
        .Plot.SeriesCollection(1).Pen.Width = 40 '//设置为较细
        
          .Plot.Axis(VtChAxisIdX).ValueScale.Maximum = MaxSec * DblSec
           .Plot.Axis(VtChAxisIdX).ValueScale.Minimum = MinSec * DblSec
 .ChartData = HUMIDITYData

End With
      
     End If '结束一个有数据的插座温度曲线显示
    
'################################################################################end插座级温度曲线显示
    
    
    
    
        
        
        
        
        
        
        
          End If
           
  
            
            
            
            
            
        plugR.MoveNext
        
        Loop
        


End Sub
Public Function buildtree() '建立导航目录树

'打开记录集，根据实际情况修改
'Set daoRs = daoDB.OpenRecordset("Select BID,tagname from buildingInformation Order by BID", dbOpenForwardOnly, dbReadOnly)
'Set areaRs = daoDB.OpenRecordset("Select AID,BID,tagname from areaInformation Order by BID", dbOpenForwardOnly, dbReadOnly)
Dim R As New ADODB.Recordset '根目录节点的记录集对象（建筑物）
Dim AreaR As New ADODB.Recordset '一级目录节点的记录集对象（区域）
Dim collectorR As New ADODB.Recordset '二级目录节点的记录集对象（采集器）
Dim plugR As New ADODB.Recordset '三级目录节点的记录集对象（插座）

R.Open "select * from buildingInformation", Adodc1.ConnectionString
With TreeView1.Nodes
    '清空列表
    .Clear
    Do Until R.EOF '遍历根目录节点并显示
        strCode = Trim(R.Fields("BID"))
       ' Print strCode
'        Select Case Len(strCode)
'        Case 1
    TreeView1.Nodes.Add , , "A" & strCode, R.Fields("tagName")  'Treeview 的 Key 用于直接访问节点。不能是数字，前面要加一个字符
    
    
AreaR.Open "select * from areaInformation ", Adodc1.ConnectionString
      Do Until AreaR.EOF '遍历一级区域目录节点并显示
        strCodeAreaBID = Trim(AreaR.Fields("BID"))
        strCodeAreaAID = Trim(AreaR.Fields("AID"))
       ' Print strCodeAreaBID
        If strCodeAreaBID = strCode Then
        Set nodeTemp = .Item("A" & strCode)
          'Print nodeTemp
           If nodeTemp Is Nothing Then Exit Do 'error
            TreeView1.Nodes.Add nodeTemp, tvwChild, "A" & strCodeAreaBID & strCodeAreaAID, AreaR.Fields("tagName")
        End If
        
         
 
collectorR.Open "select * from collectorInformation  ", Adodc1.ConnectionString
        Do Until collectorR.EOF  '遍历二级采集器目录节点并显示
        strCodecollectorAID = Trim(collectorR.Fields("AID"))
        strCodecollectorCID = Trim(collectorR.Fields("CID"))
          If strCodecollectorAID = strCodeAreaAID Then
        Set nodeTemp = .Item("A" & strCode & strCodeAreaAID)
          
           If nodeTemp Is Nothing Then Exit Do 'error
            TreeView1.Nodes.Add nodeTemp, tvwChild, "A" & strCodeAreaBID & strCodeAreaAID & strCodecollectorCID, collectorR.Fields("tagName")
        
        
        plugR.Open "select * from pluginInformation", Adodc1.ConnectionString
        Do Until plugR.EOF  '遍历三级插座目录节点并显示
        strCodeplugPID = Trim(plugR.Fields("PID"))
        strCodeplugCID = Trim(plugR.Fields("CID"))
          If strCodeplugCID = strCodecollectorCID Then
        Set nodeTemp = .Item("A" & strCodeAreaBID & strCodeAreaAID & strCodecollectorCID)
       
           If nodeTemp Is Nothing Then Exit Do 'error
            TreeView1.Nodes.Add nodeTemp, tvwChild, "A" & strCodeAreaBID & strCodeAreaAID & strCodecollectorCID & strCodeplugPID, plugR.Fields("tagName")
            
        
            
            
            
        End If
         plugR.MoveNext
        
        Loop
        plugR.Close
        
        End If
         collectorR.MoveNext
        Loop
        collectorR.Close
        
        AreaR.MoveNext
     Loop
     AreaR.Close
     
'
        R.MoveNext
    Loop
      R.Close
   
End With
Set R = Nothing

End Function

