VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{65E121D4-0C60-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCHRT20.OCX"
Begin VB.Form smartplus 
   Caption         =   "���ܲ����ն�"
   ClientHeight    =   10950
   ClientLeft      =   225
   ClientTop       =   555
   ClientWidth     =   15120
   LinkTopic       =   "Form1"
   ScaleHeight     =   10950
   ScaleWidth      =   15120
   StartUpPosition =   3  '����ȱʡ
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
         Name            =   "����"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "����"
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
   Begin VB.Frame Frame4 
      Caption         =   "Ŀ¼����"
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
   Begin VB.Frame Frame3 
      Caption         =   "�¶ȡ�ʪ��"
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
            Name            =   "����"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "����"
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
      Caption         =   "������Ϣ��ʾ��"
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
      Caption         =   "���������ʡ���������ѹ"
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
            Name            =   "����"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "����"
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
TotalData.ConnectionString = "Provider=SQLOLEDB.1;Persist Security Info=False;User ID=sa;Password=ahjzu;Initial Catalog=master;Data Source=219.231.15.100" 'ȫ��������Ϣ���ado�ؼ�
TotalTemp.ConnectionString = "Provider=SQLOLEDB.1;Persist Security Info=False;User ID=sa;Password=ahjzu;Initial Catalog=master;Data Source=219.231.15.100" 'ȫ��������Ϣ���ado�ؼ�
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
'test.ConnectionString = "Provider=SQLOLEDB.1;Persist Security Info=False;User ID=sa;Password=ahjzu;Initial Catalog=master;Data Source=219.231.15.100" 'ȫ��������Ϣ���ado�ؼ�
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

VOLTAGEchart.chartType = VtChChartType2dXY '��ѹ��άɢ��ͼ����ʾһ������
CURRENTChart.chartType = VtChChartType2dXY '������άɢ��ͼ����ʾһ������
ACTIVEENERGYChart.chartType = VtChChartType2dXY '������άɢ��ͼ����ʾһ������
TEMPChart.chartType = VtChChartType2dXY '�¶ȶ�άɢ��ͼ����ʾһ������
With ACTIVEENERGYChart '��ʼ������������ʽ
 .chartType = VtChChartType2dXY '��άɢ��ͼ��ֻ����ʾһ������ע�Ȿ����λ�ã�����������X�����꽫��ʾ��С������ʱ���ʽ
        'ͼ��ֻ��������
        .ColumnCount = 2
        .Plot.Axis(VtChAxisIdX).ValueScale.Auto = False
.Plot.Axis(VtChAxisIdY).ValueScale.Auto = False
 '// �������ֵ
     .Plot.Axis(VtChAxisIdY).ValueScale.Maximum = 1
       '// ������Сֵ
      .Plot.Axis(VtChAxisIdY).ValueScale.Minimum = 0
  .Plot.Axis(VtChAxisIdX).ValueScale.Minimum = 0

     '//����ͼ�����
      .Title.Text = "��������"


 .Plot.Axis(VtChAxisIdX).ValueScale.MajorDivision = 3 'X����Ҫ��������
    .Plot.Axis(VtChAxisIdY).ValueScale.MajorDivision = 6 'Y����Ҫ��������
    .Plot.Axis(VtChAxisIdX).ValueScale.MinorDivision = 0 'X���Ҫ��������
    .Plot.Axis(VtChAxisIdY).ValueScale.MinorDivision = 0 'Y���Ҫ��������
 .Plot.Axis(VtChAxisIdX).Labels(1).Format = "hh:mm:ss"
    .Plot.Axis(VtChAxisIdX).AxisGrid.MajorPen.Style = VtPenStyleNull '//X��������ʾ
   .Plot.Axis(VtChAxisIdY).AxisGrid.MajorPen.Style = VtPenStyleDotted '//Y������Ϊʵ��
   .Plot.AutoLayout = False   '//��Ϊ�ֶ����ô�С
   .Plot.UniformAxis = False '//ָ��ͼ�������ֵ������ĵ�λ�̶Ȳ�һ��(X,Y�����겻��Ҫһ��).
End With
With CURRENTChart '��ʼ������������ʽ
 .chartType = VtChChartType2dXY '��άɢ��ͼ��ֻ����ʾһ������ע�Ȿ����λ�ã�����������X�����꽫��ʾ��С������ʱ���ʽ
        'ͼ��ֻ��������
        .ColumnCount = 2
        .Plot.Axis(VtChAxisIdX).ValueScale.Auto = False
.Plot.Axis(VtChAxisIdY).ValueScale.Auto = False
 '// �������ֵ
     .Plot.Axis(VtChAxisIdY).ValueScale.Maximum = 1
       '// ������Сֵ
      .Plot.Axis(VtChAxisIdY).ValueScale.Minimum = 0
  .Plot.Axis(VtChAxisIdX).ValueScale.Minimum = 0

     '//����ͼ�����
      .Title.Text = "����"


 .Plot.Axis(VtChAxisIdX).ValueScale.MajorDivision = 3 'X����Ҫ��������
    .Plot.Axis(VtChAxisIdY).ValueScale.MajorDivision = 6 'Y����Ҫ��������
    .Plot.Axis(VtChAxisIdX).ValueScale.MinorDivision = 0 'X���Ҫ��������
    .Plot.Axis(VtChAxisIdY).ValueScale.MinorDivision = 0 'Y���Ҫ��������
 .Plot.Axis(VtChAxisIdX).Labels(1).Format = "hh:mm:ss"
    .Plot.Axis(VtChAxisIdX).AxisGrid.MajorPen.Style = VtPenStyleNull '//X��������ʾ
   .Plot.Axis(VtChAxisIdY).AxisGrid.MajorPen.Style = VtPenStyleDotted '//Y������Ϊʵ��
   .Plot.AutoLayout = False   '//��Ϊ�ֶ����ô�С
   .Plot.UniformAxis = False '//ָ��ͼ�������ֵ������ĵ�λ�̶Ȳ�һ��(X,Y�����겻��Ҫһ��).
End With
With VOLTAGEchart '��ʼ����ѹ������ʽ
 .chartType = VtChChartType2dXY '��άɢ��ͼ��ֻ����ʾһ������ע�Ȿ����λ�ã�����������X�����꽫��ʾ��С������ʱ���ʽ
        'ͼ��ֻ��������
        .ColumnCount = 2
        .Plot.Axis(VtChAxisIdX).ValueScale.Auto = False
.Plot.Axis(VtChAxisIdY).ValueScale.Auto = False
 '// �������ֵ
     .Plot.Axis(VtChAxisIdY).ValueScale.Maximum = 1
       '// ������Сֵ
      .Plot.Axis(VtChAxisIdY).ValueScale.Minimum = 0
  .Plot.Axis(VtChAxisIdX).ValueScale.Minimum = 0

     '//����ͼ�����
      .Title.Text = "��ѹ"


 .Plot.Axis(VtChAxisIdX).ValueScale.MajorDivision = 3 'X����Ҫ��������
    .Plot.Axis(VtChAxisIdY).ValueScale.MajorDivision = 6 'Y����Ҫ��������
    .Plot.Axis(VtChAxisIdX).ValueScale.MinorDivision = 0 'X���Ҫ��������
    .Plot.Axis(VtChAxisIdY).ValueScale.MinorDivision = 0 'Y���Ҫ��������
 .Plot.Axis(VtChAxisIdX).Labels(1).Format = "hh:mm:ss"
    .Plot.Axis(VtChAxisIdX).AxisGrid.MajorPen.Style = VtPenStyleNull '//X��������ʾ
   .Plot.Axis(VtChAxisIdY).AxisGrid.MajorPen.Style = VtPenStyleDotted '//Y������Ϊʵ��
   .Plot.AutoLayout = False   '//��Ϊ�ֶ����ô�С
   .Plot.UniformAxis = False '//ָ��ͼ�������ֵ������ĵ�λ�̶Ȳ�һ��(X,Y�����겻��Ҫһ��).
End With


With TEMPChart '��ʼ���¶�������ʽ
 .chartType = VtChChartType2dXY '��άɢ��ͼ��ֻ����ʾһ������ע�Ȿ����λ�ã�����������X�����꽫��ʾ��С������ʱ���ʽ
        'ͼ��ֻ��������
        .ColumnCount = 2
        .Plot.Axis(VtChAxisIdX).ValueScale.Auto = False
.Plot.Axis(VtChAxisIdY).ValueScale.Auto = False
 '// �������ֵ
     .Plot.Axis(VtChAxisIdY).ValueScale.Maximum = 1
       '// ������Сֵ
      .Plot.Axis(VtChAxisIdY).ValueScale.Minimum = 0
  .Plot.Axis(VtChAxisIdX).ValueScale.Minimum = 0

     '//����ͼ�����
      .Title.Text = "�¶�ƽ��ֵ"


 .Plot.Axis(VtChAxisIdX).ValueScale.MajorDivision = 3 'X����Ҫ��������
    .Plot.Axis(VtChAxisIdY).ValueScale.MajorDivision = 6 'Y����Ҫ��������
    .Plot.Axis(VtChAxisIdX).ValueScale.MinorDivision = 0 'X���Ҫ��������
    .Plot.Axis(VtChAxisIdY).ValueScale.MinorDivision = 0 'Y���Ҫ��������
 .Plot.Axis(VtChAxisIdX).Labels(1).Format = "hh:mm:ss"
    .Plot.Axis(VtChAxisIdX).AxisGrid.MajorPen.Style = VtPenStyleNull '//X��������ʾ
   .Plot.Axis(VtChAxisIdY).AxisGrid.MajorPen.Style = VtPenStyleDotted '//Y������Ϊʵ��
   .Plot.AutoLayout = False   '//��Ϊ�ֶ����ô�С
   .Plot.UniformAxis = False '//ָ��ͼ�������ֵ������ĵ�λ�̶Ȳ�һ��(X,Y�����겻��Ҫһ��).
End With
With HUMIDITYChart '��ʼ���¶�������ʽ
 .chartType = VtChChartType2dXY '��άɢ��ͼ��ֻ����ʾһ������ע�Ȿ����λ�ã�����������X�����꽫��ʾ��С������ʱ���ʽ
        'ͼ��ֻ��������
        .ColumnCount = 2
        .Plot.Axis(VtChAxisIdX).ValueScale.Auto = False
.Plot.Axis(VtChAxisIdY).ValueScale.Auto = False
 '// �������ֵ
     .Plot.Axis(VtChAxisIdY).ValueScale.Maximum = 1
       '// ������Сֵ
      .Plot.Axis(VtChAxisIdY).ValueScale.Minimum = 0
  .Plot.Axis(VtChAxisIdX).ValueScale.Minimum = 0

     '//����ͼ�����
      .Title.Text = "ʪ��ƽ��ֵ"


 .Plot.Axis(VtChAxisIdX).ValueScale.MajorDivision = 3 'X����Ҫ��������
    .Plot.Axis(VtChAxisIdY).ValueScale.MajorDivision = 6 'Y����Ҫ��������
    .Plot.Axis(VtChAxisIdX).ValueScale.MinorDivision = 0 'X���Ҫ��������
    .Plot.Axis(VtChAxisIdY).ValueScale.MinorDivision = 0 'Y���Ҫ��������
 .Plot.Axis(VtChAxisIdX).Labels(1).Format = " hh:mm:ss"
    .Plot.Axis(VtChAxisIdX).AxisGrid.MajorPen.Style = VtPenStyleNull '//X��������ʾ
   .Plot.Axis(VtChAxisIdY).AxisGrid.MajorPen.Style = VtPenStyleDotted '//Y������Ϊʵ��
   .Plot.AutoLayout = False   '//��Ϊ�ֶ����ô�С
   .Plot.UniformAxis = False '//ָ��ͼ�������ֵ������ĵ�λ�̶Ȳ�һ��(X,Y�����겻��Ҫһ��).
End With
t = buildtree()

End Sub

Private Sub Form_Load() '���ݿ�����
  
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

'Dim MyData() As Double '��������
'Dim VOLTAGEData() As Double '��������
'Dim CURRENTData() As Double '��������
Private Sub TreeView1_NodeClick(ByVal Node As MSComctlLib.Node)   '��ȡ�����еĶ�ӳ����ʾ������

Dim Mydata() As Double '��������
Dim VOLTAGEData() As Double '��������
Dim CURRENTData() As Double '��������
Dim TEMPData() As Double '�¶�����
Dim HUMIDITYData() As Double 'ʪ������
DblSec = 1.1574074074074E-05 'ʱ�����͵�����1�븳ֵ��double���ͱ���ʱ��ֵ������1�루ʱ�����ͣ�=1.1574074074074E-05��double���ͣ�

Dim buildingR As New ADODB.Recordset '��Ŀ¼�ڵ�ļ�¼�����󣨽����
Dim plugR As New ADODB.Recordset '����Ŀ¼�ڵ�ļ�¼�����󣨲�����
Dim collectorR As New ADODB.Recordset '����Ŀ¼�ڵ�ļ�¼�����󣨲ɼ�����
Dim AreaR As New ADODB.Recordset 'һ��Ŀ¼�ڵ�ļ�¼����������


'*************************************************************************************************************************��ʾ������Ļ������ݵ��������
buildingR.Open "select * from buildingInformation ", buildingactive(0).ConnectionString

        Do Until buildingR.EOF  '���������ɼ���Ŀ¼�ڵ㲢��ʾ
        strCodebuildingtagName = Trim(buildingR.Fields("tagName"))
         
If Node.Text = strCodebuildingtagName Then 'ѡ�н�����ڵ�
Timerbuilding.Enabled = True

Timerplug.Enabled = False
Timercollector = False
Timerarea = False
 VOLTAGEchart.Visible = False '����ʾ��ѹ���߱�
 CURRENTChart.Visible = False '����ʾ�������߱�
            'buildingactive(0).RecordSource = "Select top 5 sum(ACTIVEENERGY) as ��������,DBTIME,areaInformation.BID from analyze_electrictable join collectorInformation join areaInformation on collectorInformation.AID=areaInformation.AID on COLLECTIP=collectorInformation.macAddress  group by DBTIME,areaInformation.BID Order by DBTIME desc"

           buildingactive(0).RecordSource = "Select sum(ACTIVEENERGY) as ��������,DBTIME,areaInformation.BID from analyze_electrictable join collectorInformation join areaInformation on collectorInformation.AID=areaInformation.AID on COLLECTIP=collectorInformation.macAddress where DBTIME>dateadd(MINUTE,-10,GETDATE()) group by DBTIME,areaInformation.BID Order by DBTIME"
             buildingactive(1).RecordSource = "Select avg(cast(TEMP AS int)) as �¶�ƽ��ֵ,avg(cast(HUMIDITY as int)) as ʪ��ƽ��ֵ,DBTIME,areaInformation.BID from analyze_humituretable join collectorInformation join areaInformation on collectorInformation.AID=areaInformation.AID on COLLECTIP=collectorInformation.macAddress where DBTIME>dateadd(MINUTE,-10,GETDATE()) group by DBTIME,areaInformation.BID Order by DBTIME"
buildingactive(1).Refresh

          buildingactive(0).Refresh
'

If buildingactive(0).Recordset.RecordCount > 0 Then  '���������¼�������ݲ���ʾ����֮������
ReDim Mydata(buildingactive(0).Recordset.RecordCount - 1, 1) '�ض���ɼ�����������֮�͵�����
ReDim HUMIDITYData(buildingactive(1).Recordset.RecordCount - 1, 1) As Double '�ض���ʪ��ƽ��ֵ������

       
  For i = 0 To buildingactive(0).Recordset.RecordCount - 1
 TempSec = DateDiff("s", "0:0:0", TimeValue(buildingactive(0).Recordset(1).Value))
 Maxdate = buildingactive(0).Recordset(1).Value
    If i = 0 Then
       MaxSec = TempSec: MinSec = TempSec
       Maxdate0 = buildingactive(0).Recordset(1).Value
    End If
    If TempSec > MaxSec Then MaxSec = TempSec '��x�����ֵ
    If TempSec < MinSec Then MinSec = TempSec '��x����Сֵ
If Maxdate > Maxdata0 Then Maxdata0 = Maxdate
    TempACTIVEENERGYY = buildingactive(0).Recordset(0).Value '��������һ������Ϊy����Сֵ

    If i = 0 Then

       MaxACTIVEENERGYY = TempACTIVEENERGYY: MinACTIVEENERGYY = TempACTIVEENERGYY
    End If

    If TempACTIVEENERGYY > MaxACTIVEENERGYY Then MaxACTIVEENERGYY = TempACTIVEENERGYY '��ɼ�������֮��Y�����ֵ
    If TempACTIVEENERGYY < MinACTIVEENERGYY Then MinACTIVEENERGYY = TempACTIVEENERGYY '��ɼ�������֮��Y����Сֵ

  Mydata(i, 0) = TimeValue(buildingactive(0).Recordset(1).Value) '��2��ʱ��ֵ��ŵ���X������
    Mydata(i, 1) = buildingactive(0).Recordset(0).Value '��4�д������֮��Y������

       buildingactive(0).Recordset.MoveNext

    Next i
   
   With ACTIVEENERGYChart
  .Plot.Axis(VtChAxisIdY).ValueScale.Maximum = MaxACTIVEENERGYY + 5

    .Plot.Axis(VtChAxisIdY).ValueScale.Minimum = MinACTIVEENERGYY

           .Plot.SeriesCollection(1).Pen.VtColor.Set 0, 0, 255 '//��һ��Ϊ��������,����Ϊ��ɫ
        .Plot.SeriesCollection(1).Pen.Width = 20 '//����Ϊ��ϸ

          .Plot.Axis(VtChAxisIdX).ValueScale.Maximum = MaxSec * DblSec
           .Plot.Axis(VtChAxisIdX).ValueScale.Minimum = MinSec * DblSec
 .ChartData = Mydata

End With

'||||||||||||||||||||||||||||||||

ReDim TEMPData(buildingactive(1).Recordset.RecordCount - 1, 1) '�ض��彨������¶�����
ReDim HUMIDITYData(buildingactive(1).Recordset.RecordCount - 1, 1) '�ض��彨������¶�����


  For i = 0 To buildingactive(1).Recordset.RecordCount - 1
 TempSec = DateDiff("s", "0:0:0", TimeValue(buildingactive(1).Recordset(2).Value))
 
    If i = 0 Then
       MaxSec = TempSec: MinSec = TempSec
    End If
    If TempSec > MaxSec Then MaxSec = TempSec '��x�����ֵ
    If TempSec < MinSec Then MinSec = TempSec '��x����Сֵ
    
    Temptemp = buildingactive(1).Recordset(0).Value '���¶ȵ�һ������Ϊy����Сֵ
    TempHUMIDITYY = buildingactive(1).Recordset(1).Value '��ʪ�ȵ�һ������Ϊy����Сֵ
    If i = 0 Then
       MaxtempY = Temptemp: MintempY = Temptemp
      MaxHUMIDITYY = TempHUMIDITYY: MinHUMIDITYYY = TempHUMIDITYY

    End If
    
    If Temptemp > MaxtempY Then MaxtempY = Temptemp '�������¶�Y�����ֵ
    If Temptemp < MintempY Then MintempY = Temptemp '�������¶�Y����Сֵ
      If TempHUMIDITYY > MaxtempY Then MaxtempY = TempHUMIDITYY '�������¶�Y�����ֵ
    If TempHUMIDITYY < MinHUMIDITYYY Then MinHUMIDITYYY = TempHUMIDITYY '������ʪ��Y����Сֵ
    
  TEMPData(i, 0) = TimeValue(buildingactive(1).Recordset(2).Value) '��3��ʱ��ֵ��Ž�����X������
    TEMPData(i, 1) = buildingactive(1).Recordset(0).Value '��1�д����¶�Y������

    HUMIDITYData(i, 0) = TimeValue(buildingactive(1).Recordset(2).Value) '��3��ʱ��ֵ���X������
    HUMIDITYData(i, 1) = buildingactive(1).Recordset(1).Value '��1�д���ʪ��ƽ��ֵY������
   
       buildingactive(1).Recordset.MoveNext
     
    Next i
  
   With TEMPChart
  .Plot.Axis(VtChAxisIdY).ValueScale.Maximum = MaxtempY + 5
  
    .Plot.Axis(VtChAxisIdY).ValueScale.Minimum = MintempY
 
           .Plot.SeriesCollection(1).Pen.VtColor.Set 0, 0, 255 '//��һ��Ϊ��������,����Ϊ��ɫ
        .Plot.SeriesCollection(1).Pen.Width = 40 '//����Ϊ��ϸ
        
          .Plot.Axis(VtChAxisIdX).ValueScale.Maximum = MaxSec * DblSec
           .Plot.Axis(VtChAxisIdX).ValueScale.Minimum = MinSec * DblSec
 .ChartData = TEMPData

End With
      
       With HUMIDITYChart
  .Plot.Axis(VtChAxisIdY).ValueScale.Maximum = MaxHUMIDITYY + 5
  
    .Plot.Axis(VtChAxisIdY).ValueScale.Minimum = MinHUMIDITYY
 
           .Plot.SeriesCollection(1).Pen.VtColor.Set 0, 0, 255 '//��һ��Ϊ��������,����Ϊ��ɫ
        .Plot.SeriesCollection(1).Pen.Width = 40 '//����Ϊ��ϸ
        
          .Plot.Axis(VtChAxisIdX).ValueScale.Maximum = MaxSec * DblSec
           .Plot.Axis(VtChAxisIdX).ValueScale.Minimum = MinSec * DblSec
 .ChartData = HUMIDITYData

End With

'|||||||||||||||||||||||||||||||||||||||||||||||



     End If '����һ�������ݵĽ�����������ʾ

End If  '����һ���������������ʾ

     buildingR.MoveNext
        Loop
        buildingR.Close

'*************************************************************************************************************************end��ʾ������Ļ������ݵ��������
'*************************************************************************************************************************��ʾ����Ļ������ݵ��������
AreaR.Open "select * from areaInformation ", area(0).ConnectionString

        Do Until AreaR.EOF  '���������ɼ���Ŀ¼�ڵ㲢��ʾ
        strCodeAreatagName = Trim(AreaR.Fields("tagName"))
          'strCodeCOLLECTIP = AreaR.Fields("macAddress")
If Node.Text = strCodeAreatagName Then 'ѡ������ڵ�
Timerarea.Enabled = True
Timerbuilding.Enabled = False

Timerplug.Enabled = False
Timercollector = False



 VOLTAGEchart.Visible = False '����ʾ��ѹ���߱�
 CURRENTChart.Visible = False '����ʾ�������߱�
       
        area(0).RecordSource = "Select sum(ACTIVEENERGY) as ��������,DBTIME,collectorInformation.AID from analyze_electrictable join collectorInformation on COLLECTIP=collectorInformation.macAddress where DBTIME>dateadd(MINUTE,-10,GETDATE()) group by DBTIME,collectorInformation.AID Order by DBTIME"
        area(1).RecordSource = "Select avg(cast(TEMP AS int)) as �¶�ƽ��ֵ,avg(cast(HUMIDITY as int)) as ʪ��ƽ��ֵ,DBTIME,collectorInformation.AID from analyze_humituretable join collectorInformation on COLLECTIP=collectorInformation.macAddress  where  DBTIME>dateadd(MINUTE,-10,GETDATE()) group by DBTIME,collectorInformation.AID Order by DBTIME"
      
           'test.RecordSource = "Select sum(ACTIVEENERGY) as ��������,DBTIME,collectorInformation.AID from analyze_electrictable join collectorInformation on COLLECTIP=collectorInformation.macAddress where DBTIME>dateadd(MINUTE,-10,GETDATE())group by DBTIME,COLLECTIP"
'test.Refresh
          area(1).Refresh

          area(0).Refresh
'

If area(0).Recordset.RecordCount > 0 Then  '�����¼�������ݲ���ʾ����
ReDim Mydata(area(0).Recordset.RecordCount - 1, 1) '�ض���ɼ�����������֮�͵�����
       
  For i = 0 To area(0).Recordset.RecordCount - 1
 TempSec = DateDiff("s", "0:0:0", TimeValue(area(0).Recordset(1).Value))

    If i = 0 Then
       MaxSec = TempSec: MinSec = TempSec
    End If
    If TempSec > MaxSec Then MaxSec = TempSec '��x�����ֵ
    If TempSec < MinSec Then MinSec = TempSec '��x����Сֵ

    TempACTIVEENERGYY = area(0).Recordset(0).Value '��������һ������Ϊy����Сֵ

    If i = 0 Then

       MaxACTIVEENERGYY = TempACTIVEENERGYY: MinACTIVEENERGYY = TempACTIVEENERGYY
    End If

    If TempACTIVEENERGYY > MaxACTIVEENERGYY Then MaxACTIVEENERGYY = TempACTIVEENERGYY '��ɼ�������֮��Y�����ֵ
    If TempACTIVEENERGYY < MinACTIVEENERGYY Then MinACTIVEENERGYY = TempACTIVEENERGYY '��ɼ�������֮��Y����Сֵ

  Mydata(i, 0) = TimeValue(area(0).Recordset(1).Value) '��2��ʱ��ֵ��ŵ���X������
    Mydata(i, 1) = area(0).Recordset(0).Value '��4�д������֮��Y������

       area(0).Recordset.MoveNext

    Next i
   
   With ACTIVEENERGYChart
  .Plot.Axis(VtChAxisIdY).ValueScale.Maximum = MaxACTIVEENERGYY + 5

    .Plot.Axis(VtChAxisIdY).ValueScale.Minimum = MinACTIVEENERGYY

           .Plot.SeriesCollection(1).Pen.VtColor.Set 0, 0, 255 '//��һ��Ϊ��������,����Ϊ��ɫ
        .Plot.SeriesCollection(1).Pen.Width = 20 '//����Ϊ��ϸ

          .Plot.Axis(VtChAxisIdX).ValueScale.Maximum = MaxSec * DblSec
           .Plot.Axis(VtChAxisIdX).ValueScale.Minimum = MinSec * DblSec
 .ChartData = Mydata

End With


ReDim TEMPData(area(1).Recordset.RecordCount - 1, 1) '�ض��彨������¶�����
ReDim HUMIDITYData(area(1).Recordset.RecordCount - 1, 1) '�ض��彨������¶�����


  For i = 0 To area(1).Recordset.RecordCount - 1
 TempSec = DateDiff("s", "0:0:0", TimeValue(area(1).Recordset(2).Value))
 
    If i = 0 Then
       MaxSec = TempSec: MinSec = TempSec
    End If
    If TempSec > MaxSec Then MaxSec = TempSec '��x�����ֵ
    If TempSec < MinSec Then MinSec = TempSec '��x����Сֵ
    
    Temptemp = area(1).Recordset(0).Value '���¶ȵ�һ������Ϊy����Сֵ
    TempHUMIDITYY = area(1).Recordset(1).Value '��ʪ�ȵ�һ������Ϊy����Сֵ
    If i = 0 Then
       MaxtempY = Temptemp: MintempY = Temptemp
      MaxHUMIDITYY = TempHUMIDITYY: MinHUMIDITYYY = TempHUMIDITYY

    End If
    
    If Temptemp > MaxtempY Then MaxtempY = Temptemp '��ɼ����¶�Y�����ֵ
    If Temptemp < MintempY Then MintempY = Temptemp '��ɼ����¶�Y����Сֵ
      If TempHUMIDITYY > MaxtempY Then MaxtempY = TempHUMIDITYY '��ɼ����¶�Y�����ֵ
    If TempHUMIDITYY < MinHUMIDITYYY Then MinHUMIDITYYY = TempHUMIDITYY '��ɼ���ʪ��Y����Сֵ
    
  TEMPData(i, 0) = TimeValue(area(1).Recordset(2).Value) '��3��ʱ��ֵ��Ųɼ���X������
    TEMPData(i, 1) = area(1).Recordset(0).Value '��1�д����¶�Y������

    HUMIDITYData(i, 0) = TimeValue(area(1).Recordset(2).Value) '��3��ʱ��ֵ���X������
    HUMIDITYData(i, 1) = area(1).Recordset(1).Value '��1�д���ʪ��ƽ��ֵY������
   
       area(1).Recordset.MoveNext
     
    Next i
  
   With TEMPChart
  .Plot.Axis(VtChAxisIdY).ValueScale.Maximum = MaxtempY + 5
  
    .Plot.Axis(VtChAxisIdY).ValueScale.Minimum = MintempY
 
           .Plot.SeriesCollection(1).Pen.VtColor.Set 0, 0, 255 '//��һ��Ϊ��������,����Ϊ��ɫ
        .Plot.SeriesCollection(1).Pen.Width = 40 '//����Ϊ��ϸ
        
          .Plot.Axis(VtChAxisIdX).ValueScale.Maximum = MaxSec * DblSec
           .Plot.Axis(VtChAxisIdX).ValueScale.Minimum = MinSec * DblSec
 .ChartData = TEMPData

End With
      
       With HUMIDITYChart
  .Plot.Axis(VtChAxisIdY).ValueScale.Maximum = MaxHUMIDITYY + 5
  
    .Plot.Axis(VtChAxisIdY).ValueScale.Minimum = MinHUMIDITYY
 
           .Plot.SeriesCollection(1).Pen.VtColor.Set 0, 0, 255 '//��һ��Ϊ��������,����Ϊ��ɫ
        .Plot.SeriesCollection(1).Pen.Width = 40 '//����Ϊ��ϸ
        
          .Plot.Axis(VtChAxisIdX).ValueScale.Maximum = MaxSec * DblSec
           .Plot.Axis(VtChAxisIdX).ValueScale.Minimum = MinSec * DblSec
 .ChartData = HUMIDITYData

End With







     End If '����һ�������ݵ�����������ʾ

End If  '����һ�������������ʾ

     AreaR.MoveNext
        Loop
        AreaR.Close

'*************************************************************************************************************************end��ʾ����Ļ������ݵ��������


'$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$��ʾ�ɼ����Ļ������ݵ��������
collectorR.Open "select * from collectorInformation  ", collector(0).ConnectionString

        Do Until collectorR.EOF  '���������ɼ���Ŀ¼�ڵ㲢��ʾ
        strCodecollectortagName = Trim(collectorR.Fields("tagName"))
       ' strCodeCOLLECTIP = Trim(collectorR.Fields("macAddress"))
          strCodeCOLLECTIP = collectorR.Fields("macAddress")
If Node.Text = strCodecollectortagName Then 'ѡ�вɼ����ڵ�
Timercollector.Enabled = True
Timerbuilding.Enabled = False
Timerarea.Enabled = False


Timerplug.Enabled = False

 VOLTAGEchart.Visible = False '����ʾ��ѹ���߱�
 CURRENTChart.Visible = False '����ʾ�������߱�
       
             collector(0).RecordSource = "Select sum([CURRENT]) as ��������,DBTIME,COLLECTIP,sum(VOLTAGE) as ��ѹ����,sum(ACTIVEENERGY) as �������� from analyze_electrictable where COLLECTIP='" & strCodeCOLLECTIP & "' and DBTIME>dateadd(MINUTE,-10,GETDATE()) group by DBTIME,COLLECTIP Order by DBTIME"
             collector(1).RecordSource = "Select avg(cast(TEMP AS int)) as �¶�ƽ��ֵ,avg(cast(HUMIDITY as int)) as ʪ��ƽ��ֵ,DBTIME,COLLECTIP from analyze_humituretable where COLLECTIP='" & strCodeCOLLECTIP & "' and DBTIME>dateadd(MINUTE,-10,GETDATE())group by DBTIME,COLLECTIP"

          collector(0).Refresh
          collector(1).Refresh
'           test.RecordSource = "Select sum([CURRENT]) as ��������,DBTIME,COLLECTIP,sum(VOLTAGE)as ��ѹ����,sum(ACTIVEENERGY)as �������� from analyze_electrictable where COLLECTIP='" & strCodeCOLLECTIP & "' group by DBTIME,COLLECTIP Order by DBTIME"
'           test.Refresh
        
        
           
        
If collector(0).Recordset.RecordCount > 0 Then  '�����¼�������ݲ���ʾ����
ReDim Mydata(collector(0).Recordset.RecordCount - 1, 1) '�ض���ɼ�����������֮�͵�����
         ReDim CURRENTData(collector(0).Recordset.RecordCount - 1, 1) '�ض���ɼ�����������֮�͵�����
 ReDim VOLTAGEData(collector(0).Recordset.RecordCount - 1, 1) '�ض���ɼ�����ѹ֮������
  For i = 0 To collector(0).Recordset.RecordCount - 1
 TempSec = DateDiff("s", "0:0:0", TimeValue(collector(0).Recordset(1).Value))
 
    If i = 0 Then
       MaxSec = TempSec: MinSec = TempSec
    End If
    If TempSec > MaxSec Then MaxSec = TempSec '��x�����ֵ
    If TempSec < MinSec Then MinSec = TempSec '��x����Сֵ
    
    TempACTIVEENERGYY = collector(0).Recordset(4).Value '��������һ������Ϊy����Сֵ
     TempCurrentY = collector(0).Recordset(0).Value '��������һ������Ϊy����Сֵ
    TempVoltageY = collector(0).Recordset(3).Value '����ѹ��һ������Ϊy����Сֵ
    If i = 0 Then
       MaxY = TempCurrentY: MinY = TempCurrentY
       MaxVoltageY = TempVoltageY: MinVoltageY = TempVoltageY
       MaxACTIVEENERGYY = TempACTIVEENERGYY: MinACTIVEENERGYY = TempACTIVEENERGYY
    End If
    
    If TempACTIVEENERGYY > MaxACTIVEENERGYY Then MaxACTIVEENERGYY = TempACTIVEENERGYY '��ɼ�������֮��Y�����ֵ
    If TempACTIVEENERGYY < MinACTIVEENERGYY Then MinACTIVEENERGYY = TempACTIVEENERGYY '��ɼ�������֮��Y����Сֵ
    If TempCurrentY > MaxY Then MaxY = TempCurrentY '��ɼ�������֮��Y�����ֵ
    If TempCurrentY < MinY Then MinY = TempCurrentY '��ɼ�������֮��Y����Сֵ
    
    If TempVoltageY > MaxVoltageY Then MaxVoltageY = TempVoltageY '��ɼ�����ѹ֮��Y�����ֵ
    If TempVoltageY < MinVoltageY Then MinVoltageY = TempVoltageY '��ɼ�����ѹ֮��Y����Сֵ
    
    
 Mydata(i, 0) = TimeValue(collector(0).Recordset(1).Value) '��2��ʱ��ֵ��ŵ���X������
 'Mydata(i, 0) = collector(0).Recordset(1).Value '��2��ʱ��ֵ��ŵ���X������


    Mydata(i, 1) = collector(0).Recordset(4).Value '��4�д������֮��Y������
 CURRENTData(i, 0) = TimeValue(collector(0).Recordset(1).Value) '��2��ʱ��ֵ��ŵ���X������
    CURRENTData(i, 1) = collector(0).Recordset(0).Value '��1�д������֮��Y������
      VOLTAGEData(i, 0) = TimeValue(collector(0).Recordset(1).Value) '��2��ʱ��ֵ��ŵ�ѹX������
    VOLTAGEData(i, 1) = collector(0).Recordset(3).Value '��4�д����ѹ֮��Y������
       
       collector(0).Recordset.MoveNext
     
    Next i
   
      'collector(0).Recordset.Close
   With ACTIVEENERGYChart
  .Plot.Axis(VtChAxisIdY).ValueScale.Maximum = MaxACTIVEENERGYY + 5
  
    .Plot.Axis(VtChAxisIdY).ValueScale.Minimum = MinACTIVEENERGYY
 
           .Plot.SeriesCollection(1).Pen.VtColor.Set 0, 0, 255 '//��һ��Ϊ��������,����Ϊ��ɫ
        .Plot.SeriesCollection(1).Pen.Width = 10 '//����Ϊ��ϸ
        
          .Plot.Axis(VtChAxisIdX).ValueScale.Maximum = MaxSec * DblSec
           .Plot.Axis(VtChAxisIdX).ValueScale.Minimum = MinSec * DblSec
 .ChartData = Mydata

End With
      
      
      
      
With CURRENTChart
  .Plot.Axis(VtChAxisIdY).ValueScale.Maximum = MaxX
  
    .Plot.Axis(VtChAxisIdY).ValueScale.Minimum = MinY
 ' .Plot.Axis(VtChAxisIdX).ValueScale.Minimum = 0
 
           .Plot.SeriesCollection(1).Pen.VtColor.Set 0, 0, 255 '//��һ��Ϊ��������,����Ϊ��ɫ
        .Plot.SeriesCollection(1).Pen.Width = 10 '//����Ϊ��ϸ
        
          .Plot.Axis(VtChAxisIdX).ValueScale.Maximum = MaxSec * DblSec
           .Plot.Axis(VtChAxisIdY).ValueScale.Maximum = MaxY + 5
 .Plot.Axis(VtChAxisIdX).ValueScale.Minimum = MinSec * DblSec
 .ChartData = CURRENTData

End With

With VOLTAGEchart
  .Plot.Axis(VtChAxisIdY).ValueScale.Maximum = MaxVoltageY + 5
  
    .Plot.Axis(VtChAxisIdY).ValueScale.Minimum = MinVoltageY
  '.Plot.Axis(VtChAxisIdX).ValueScale.Minimum = 0
 
           .Plot.SeriesCollection(1).Pen.VtColor.Set 0, 0, 255 '//��һ��Ϊ��������,����Ϊ��ɫ
        .Plot.SeriesCollection(1).Pen.Width = 10 '//����Ϊ��ϸ
        
          .Plot.Axis(VtChAxisIdX).ValueScale.Maximum = MaxSec * DblSec
           .Plot.Axis(VtChAxisIdX).ValueScale.Minimum = MinSec * DblSec
 .ChartData = VOLTAGEData

End With

'||||||||||||||||||||||||||||||||

ReDim TEMPData(collector(1).Recordset.RecordCount - 1, 1) '�ض���ɼ������¶�����
ReDim HUMIDITYData(collector(1).Recordset.RecordCount - 1, 1) '�ض���ɼ����¶�����


  For i = 0 To collector(1).Recordset.RecordCount - 1
 TempSec = DateDiff("s", "0:0:0", TimeValue(collector(1).Recordset(2).Value))
 
    If i = 0 Then
       MaxSec = TempSec: MinSec = TempSec
    End If
    If TempSec > MaxSec Then MaxSec = TempSec '��x�����ֵ
    If TempSec < MinSec Then MinSec = TempSec '��x����Сֵ
    
    Temptemp = collector(1).Recordset(0).Value '���¶ȵ�һ������Ϊy����Сֵ
    TempHUMIDITYY = collector(1).Recordset(1).Value '��ʪ�ȵ�һ������Ϊy����Сֵ
    If i = 0 Then
       MaxtempY = Temptemp: MintempY = Temptemp
      MaxHUMIDITYY = TempHUMIDITYY: MinHUMIDITYYY = TempHUMIDITYY

    End If
    
    If Temptemp > MaxtempY Then MaxtempY = Temptemp '�������¶�Y�����ֵ
    If Temptemp < MintempY Then MintempY = Temptemp '�������¶�Y����Сֵ
      If TempHUMIDITYY > MaxtempY Then MaxtempY = TempHUMIDITYY '�������¶�Y�����ֵ
    If TempHUMIDITYY < MinHUMIDITYYY Then MinHUMIDITYYY = TempHUMIDITYY '������ʪ��Y����Сֵ
    
  TEMPData(i, 0) = TimeValue(collector(1).Recordset(2).Value) '��3��ʱ��ֵ�������X������
    TEMPData(i, 1) = collector(1).Recordset(0).Value '��1�д����¶�Y������

    HUMIDITYData(i, 0) = TimeValue(collector(1).Recordset(2).Value) '��3��ʱ��ֵ���X������
    HUMIDITYData(i, 1) = collector(1).Recordset(1).Value   '��1�д���ʪ��ƽ��ֵY������
   
       collector(1).Recordset.MoveNext
     
    Next i
  
   With TEMPChart
  .Plot.Axis(VtChAxisIdY).ValueScale.Maximum = MaxtempY + 5
  
    .Plot.Axis(VtChAxisIdY).ValueScale.Minimum = MintempY
 
           .Plot.SeriesCollection(1).Pen.VtColor.Set 0, 0, 255 '//��һ��Ϊ��������,����Ϊ��ɫ
        .Plot.SeriesCollection(1).Pen.Width = 40 '//����Ϊ��ϸ
        
          .Plot.Axis(VtChAxisIdX).ValueScale.Maximum = MaxSec * DblSec
           .Plot.Axis(VtChAxisIdX).ValueScale.Minimum = MinSec * DblSec
 .ChartData = TEMPData

End With
      
       With HUMIDITYChart
  .Plot.Axis(VtChAxisIdY).ValueScale.Maximum = MaxHUMIDITYY + 5
  
    .Plot.Axis(VtChAxisIdY).ValueScale.Minimum = MinHUMIDITYY
 
           .Plot.SeriesCollection(1).Pen.VtColor.Set 0, 0, 255 '//��һ��Ϊ��������,����Ϊ��ɫ
        .Plot.SeriesCollection(1).Pen.Width = 40 '//����Ϊ��ϸ
        
          .Plot.Axis(VtChAxisIdX).ValueScale.Maximum = MaxSec * DblSec
           .Plot.Axis(VtChAxisIdX).ValueScale.Minimum = MinSec * DblSec
 .ChartData = HUMIDITYData

End With




































'|||||||||||||||||||||||||||||||||||||||||||||||
     End If '����һ�������ݵĲɼ���������ʾ
           
End If  '����һ���ɼ�����������ʾ
      
     collectorR.MoveNext
        Loop
        collectorR.Close
        
'$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$end��ʾ�ɼ����Ļ������ݵ��������


Dim analyze_electrictableR As New ADODB.Recordset '�������������ݱ�ļ�¼������
Dim LastTimeR As New ADODB.Recordset '����ȡ���һ����¼ʱ��ļ�¼����
plugR.Open "select * from pluginInformation", plug(0).ConnectionString
        Do Until plugR.EOF  '������������Ŀ¼�ڵ㲢��ʾ
         
        strCodeplugtagName = Trim(plugR.Fields("tagName"))
        strCodeplugmacAddress = Trim(plugR.Fields("macAddress"))
If Node.Text = strCodeplugtagName Then   '���ѡ�в������Ľڵ�
Timerplug.Enabled = True
Timercollector = False
Timerbuilding = False
Timerarea = False

 CURRENTChart.Visible = True
 VOLTAGEchart.Visible = True
 
 '*+++++++++++++++++++++++++++++++++++++++++++++��ȡ�������һ����¼��ʱ��
 LastTimeR.Open "Select ACTIVEPOWER,REACTIVEPOWER,VOLTAGE,[CURRENT],ACTIVEENERGY,DBTIME,MAC from analyze_electrictable where MAC='" & strCodeplugmacAddress & "'Order by DBTIME", plug(0).ConnectionString
    Do Until LastTimeR.EOF
    pluglastTime = LastTimeR(5)
  LastTimeR.MoveNext
  Loop
 Print pluglastTime
 LastTimeR.Close
  '*+++++++++++++++++++++++++++++++++++++++++++++��ȡ�������һ����¼��ʱ��

                plug(0).RecordSource = "Select ACTIVEPOWER,REACTIVEPOWER,VOLTAGE,[CURRENT],ACTIVEENERGY,DBTIME,MAC from analyze_electrictable where MAC='" & strCodeplugmacAddress & "'and DBTIME>dateadd(MINUTE,-10,GETDATE()) Order by DBTIME"


'Adodc2.RecordSource = "select TEMP,HUMIDITY,MAC,COLLECTIP,DBTIME from analyze_humituretable  where DBTIME>dateadd(MINUTE,-10,GETDATE())"  '��ʾ��10���ӵĲ����¶ȣ���ʾϵͳʱҪ��ϵͳʱ���趨Ϊ10-21 9��01��50����Ϊ�������ݿ��¼�����һ����¼�Ĳ���ʱ��
plug(1).RecordSource = "select TEMP,HUMIDITY,MAC,COLLECTIP,DBTIME from analyze_humituretable  where MAC='" & strCodeplugmacAddress & "' and DBTIME>dateadd(MINUTE,-10,GETDATE()) Order by DBTIME"  '��ʾ��10���ӵĲ����¶ȣ���ʾϵͳʱҪ��ϵͳʱ���趨Ϊ10-21 9��01��50����Ϊ�������ݿ��¼�����һ����¼�Ĳ���ʱ��

        plug(0).Refresh
        plug(1).Refresh
          


With ACTIVEENERGYChart
     .chartType = VtChChartType2dXY '��άɢ��ͼ��ֻ����ʾһ������ע�Ȿ����λ�ã�����������X�����꽫��ʾ��С������ʱ���ʽ
        'ͼ��ֻ��������
        .ColumnCount = 2
       VOLTAGEchart.ColumnCount = 2
      
'        'X����ʾ10����λ
'          .RowCount = plug(0).Recordset.RecordCount
        
          '����XY��?
.Plot.Axis(VtChAxisIdX).ValueScale.Auto = False
.Plot.Axis(VtChAxisIdY).ValueScale.Auto = False
VOLTAGEchart.Plot.Axis(VtChAxisIdX).ValueScale.Auto = False
VOLTAGEchart.Plot.Axis(VtChAxisIdY).ValueScale.Auto = False

 
  
   VOLTAGEchart.Plot.Axis(VtChAxisIdY).ValueScale.Minimum = 0
 VOLTAGEchart.Plot.Axis(VtChAxisIdX).ValueScale.Minimum = 0

      '//����ͼ�����
      .Title.Text = "����"
        VOLTAGEchart.Title.Text = "��ѹ"

ReDim Mydata(plug(0).Recordset.RecordCount - 1, 1) '�ض����������
ReDim CURRENTData(plug(0).Recordset.RecordCount - 1, 1) '�ض����������
ReDim VOLTAGEData(plug(0).Recordset.RecordCount - 1, 1) '��ѹ����
For i = 0 To plug(0).Recordset.RecordCount - 1
  TempSec = DateDiff("s", "0:0:0", TimeValue(plug(0).Recordset(5).Value))
  'Print TempSec
If i = 0 Then
MaxSec = TempSec: MinSec = TempSec
End If
If TempSec > MaxSec Then MaxSec = TempSec '�����ֵ
If TempSec < MinSec Then MinSec = TempSec '����Сֵ

TempACTIVEENERGYY = plug(0).Recordset(4).Value '��������һ������Ϊy����Сֵ

    If i = 0 Then

       MaxACTIVEENERGYY = TempACTIVEENERGYY: MinACTIVEENERGYY = TempACTIVEENERGYY
    End If

    If TempACTIVEENERGYY > MaxACTIVEENERGYY Then MaxACTIVEENERGYY = TempACTIVEENERGYY '��ɼ�������֮��Y�����ֵ
    If TempACTIVEENERGYY < MinACTIVEENERGYY Then MinACTIVEENERGYY = TempACTIVEENERGYY '��ɼ�������֮��Y����Сֵ


VOLTAGEData(i, 0) = TimeValue(plug(0).Recordset(5).Value) '��6��ʱ��ֵ��ŵ�ѹX������

CURRENTData(i, 0) = TimeValue(plug(0).Recordset(5).Value) '��6��ʱ��ֵ��ŵ���X������
Mydata(i, 0) = TimeValue(plug(0).Recordset(5).Value) '��6��ʱ��ֵ���X������
Mydata(i, 1) = plug(0).Recordset(4).Value '��5�е������Y������
'MyData(i, 2) = plug(0).Recordset(1).Value '��2�д��Y������

VOLTAGEData(i, 1) = plug(0).Recordset(2).Value '��3�д����ѹY������
CURRENTData(i, 1) = plug(0).Recordset(3).Value '��4�д������Y������
         plug(0).Recordset.MoveNext

Next i
          
          
           .Plot.SeriesCollection(1).Pen.VtColor.Set 0, 0, 255 '//��һ��Ϊ��������,����Ϊ��ɫ
        .Plot.SeriesCollection(1).Pen.Width = 50 '//����Ϊ��ϸ
       
        
          .Plot.Axis(VtChAxisIdX).ValueScale.Maximum = MaxSec * DblSec
 .Plot.Axis(VtChAxisIdX).ValueScale.Minimum = MinSec * DblSec
  '.Plot.Axis(VtChAxisIdY).ValueScale.Maximum = 1
  
    .Plot.Axis(VtChAxisIdY).ValueScale.Maximum = MaxACTIVEENERGYY + 0.1
  .Plot.Axis(VtChAxisIdY).ValueScale.Minimum = MinACTIVEENERGYY
 .ChartData = Mydata '��������?
 
 '&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&
With VOLTAGEchart
  .Plot.Axis(VtChAxisIdY).ValueScale.Maximum = 300
  
    .Plot.Axis(VtChAxisIdY).ValueScale.Minimum = 0
  .Plot.Axis(VtChAxisIdX).ValueScale.Minimum = 0
 
           .Plot.SeriesCollection(1).Pen.VtColor.Set 0, 0, 255 '//��һ��Ϊ��������,����Ϊ��ɫ
        .Plot.SeriesCollection(1).Pen.Width = 10 '//����Ϊ��ϸ
        
       
        
        
          .Plot.Axis(VtChAxisIdX).ValueScale.Maximum = MaxSec * DblSec
 .Plot.Axis(VtChAxisIdX).ValueScale.Minimum = MinSec * DblSec

End With
With CURRENTChart
  .Plot.Axis(VtChAxisIdY).ValueScale.Maximum = 1
  
    .Plot.Axis(VtChAxisIdY).ValueScale.Minimum = 0
  .Plot.Axis(VtChAxisIdX).ValueScale.Minimum = 0
 
           .Plot.SeriesCollection(1).Pen.VtColor.Set 0, 0, 255 '//��һ��Ϊ��������,����Ϊ��ɫ
        .Plot.SeriesCollection(1).Pen.Width = 40 '//����Ϊ��ϸ
        
        
        
        
          .Plot.Axis(VtChAxisIdX).ValueScale.Maximum = MaxSec * DblSec
 .Plot.Axis(VtChAxisIdX).ValueScale.Minimum = MinSec * DblSec
 .ChartData = CURRENTData

End With

VOLTAGEchart.ChartData = VOLTAGEData '����?
      End With



  MSChart1.Plot.UniformAxis = False

'#############################################################################�������¶���ʾ
If plug(1).Recordset.RecordCount > 0 Then  '�����¼�������ݲ���ʾ����
ReDim TEMPData(plug(1).Recordset.RecordCount - 1, 1) '�ض���������¶�����
ReDim HUMIDITYData(plug(1).Recordset.RecordCount - 1, 1) '�ض���������¶�����


  For i = 0 To plug(1).Recordset.RecordCount - 1
 TempSec = DateDiff("s", "0:0:0", TimeValue(plug(1).Recordset(4).Value))
 
    If i = 0 Then
       MaxSec = TempSec: MinSec = TempSec
    End If
    If TempSec > MaxSec Then MaxSec = TempSec '��x�����ֵ
    If TempSec < MinSec Then MinSec = TempSec '��x����Сֵ
    
    Temptemp = plug(1).Recordset(0).Value '���¶ȵ�һ������Ϊy����Сֵ
    TempHUMIDITYY = plug(1).Recordset(1).Value '��ʪ�ȵ�һ������Ϊy����Сֵ
    If i = 0 Then
       MaxtempY = Temptemp: MintempY = Temptemp
      MaxHUMIDITYY = TempHUMIDITYY: MinHUMIDITYYY = TempHUMIDITYY

    End If
    
    If Temptemp > MaxtempY Then MaxtempY = Temptemp '������¶�Y�����ֵ
    If Temptemp < MintempY Then MintempY = Temptemp '������¶�Y����Сֵ
      If TempHUMIDITYY > MaxtempY Then MaxtempY = TempHUMIDITYY '������¶�Y�����ֵ
    If TempHUMIDITYY < MinHUMIDITYYY Then MinHUMIDITYYY = TempHUMIDITYY '�����ʪ��Y����Сֵ
    
  TEMPData(i, 0) = TimeValue(plug(1).Recordset(4).Value) '��5��ʱ��ֵ��ŵ���X������
    TEMPData(i, 1) = plug(1).Recordset(0).Value '��1�д����¶�Y������

    HUMIDITYData(i, 0) = TimeValue(plug(1).Recordset(4).Value) '��5��ʱ��ֵ��ŵ���X������
    HUMIDITYData(i, 1) = plug(1).Recordset(1).Value '��1�д����¶�Y������
   
       plug(1).Recordset.MoveNext
     
    Next i
  
   With TEMPChart
       .Title.Text = "�¶�"
  .Plot.Axis(VtChAxisIdY).ValueScale.Maximum = MaxtempY + 5
  
    .Plot.Axis(VtChAxisIdY).ValueScale.Minimum = MintempY
 
           .Plot.SeriesCollection(1).Pen.VtColor.Set 0, 0, 255 '//��һ��Ϊ��������,����Ϊ��ɫ
        .Plot.SeriesCollection(1).Pen.Width = 40 '//����Ϊ��ϸ
        
          .Plot.Axis(VtChAxisIdX).ValueScale.Maximum = MaxSec * DblSec
           .Plot.Axis(VtChAxisIdX).ValueScale.Minimum = MinSec * DblSec
 .ChartData = TEMPData

End With
      
       With HUMIDITYChart
           .Title.Text = "ʪ��"
  .Plot.Axis(VtChAxisIdY).ValueScale.Maximum = MaxHUMIDITYY + 5
  
    .Plot.Axis(VtChAxisIdY).ValueScale.Minimum = MinHUMIDITYY
 
           .Plot.SeriesCollection(1).Pen.VtColor.Set 0, 0, 255 '//��һ��Ϊ��������,����Ϊ��ɫ
        .Plot.SeriesCollection(1).Pen.Width = 40 '//����Ϊ��ϸ
        
          .Plot.Axis(VtChAxisIdX).ValueScale.Maximum = MaxSec * DblSec
           .Plot.Axis(VtChAxisIdX).ValueScale.Minimum = MinSec * DblSec
 .ChartData = HUMIDITYData

End With
      
     End If '����һ�������ݵĲ����¶�������ʾ
    
'################################################################################end�������¶�������ʾ
    
    
    
    
        
        
        
        
        
        
        
          End If
           
  
            
            
            
            
            
        plugR.MoveNext
        
        Loop
        


End Sub
Public Function buildtree() '��������Ŀ¼��

'�򿪼�¼��������ʵ������޸�
'Set daoRs = daoDB.OpenRecordset("Select BID,tagname from buildingInformation Order by BID", dbOpenForwardOnly, dbReadOnly)
'Set areaRs = daoDB.OpenRecordset("Select AID,BID,tagname from areaInformation Order by BID", dbOpenForwardOnly, dbReadOnly)
Dim R As New ADODB.Recordset '��Ŀ¼�ڵ�ļ�¼�����󣨽����
Dim AreaR As New ADODB.Recordset 'һ��Ŀ¼�ڵ�ļ�¼����������
Dim collectorR As New ADODB.Recordset '����Ŀ¼�ڵ�ļ�¼�����󣨲ɼ�����
Dim plugR As New ADODB.Recordset '����Ŀ¼�ڵ�ļ�¼�����󣨲�����

R.Open "select * from buildingInformation", Adodc1.ConnectionString
With TreeView1.Nodes
    '����б�
    .Clear
    Do Until R.EOF '������Ŀ¼�ڵ㲢��ʾ
        strCode = Trim(R.Fields("BID"))
       ' Print strCode
'        Select Case Len(strCode)
'        Case 1
    TreeView1.Nodes.Add , , "A" & strCode, R.Fields("tagName")  'Treeview �� Key ����ֱ�ӷ��ʽڵ㡣���������֣�ǰ��Ҫ��һ���ַ�
    
    
AreaR.Open "select * from areaInformation ", Adodc1.ConnectionString
      Do Until AreaR.EOF '����һ������Ŀ¼�ڵ㲢��ʾ
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
        Do Until collectorR.EOF  '���������ɼ���Ŀ¼�ڵ㲢��ʾ
        strCodecollectorAID = Trim(collectorR.Fields("AID"))
        strCodecollectorCID = Trim(collectorR.Fields("CID"))
          If strCodecollectorAID = strCodeAreaAID Then
        Set nodeTemp = .Item("A" & strCode & strCodeAreaAID)
          
           If nodeTemp Is Nothing Then Exit Do 'error
            TreeView1.Nodes.Add nodeTemp, tvwChild, "A" & strCodeAreaBID & strCodeAreaAID & strCodecollectorCID, collectorR.Fields("tagName")
        
        
        plugR.Open "select * from pluginInformation", Adodc1.ConnectionString
        Do Until plugR.EOF  '������������Ŀ¼�ڵ㲢��ʾ
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

