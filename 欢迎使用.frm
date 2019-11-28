VERSION 5.00
Begin VB.Form smartplus 
   BackColor       =   &H00FFFFFF&
   Caption         =   "建筑用能采集系统"
   ClientHeight    =   10935
   ClientLeft      =   2865
   ClientTop       =   255
   ClientWidth     =   15120
   LinkTopic       =   "Form1"
   ScaleHeight     =   10935
   ScaleWidth      =   15120
   StartUpPosition =   2  '屏幕中心
   Begin VB.CommandButton Command1 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000D&
      Caption         =   "注销登录"
      Height          =   375
      Left            =   1800
      MaskColor       =   &H8000000D&
      TabIndex        =   13
      Top             =   2160
      UseMaskColor    =   -1  'True
      Width           =   975
   End
   Begin VB.Timer Timerplug 
      Enabled         =   0   'False
      Interval        =   5000
      Left            =   14520
      Top             =   9960
   End
   Begin VB.Timer Timercollector 
      Enabled         =   0   'False
      Left            =   14520
      Top             =   9360
   End
   Begin VB.PictureBox collector 
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   330
      Index           =   0
      Left            =   7440
      ScaleHeight     =   270
      ScaleWidth      =   2955
      TabIndex        =   22
      Top             =   4440
      Visible         =   0   'False
      Width           =   3015
   End
   Begin VB.Timer Timerarea 
      Enabled         =   0   'False
      Interval        =   500
      Left            =   14520
      Top             =   8520
   End
   Begin VB.PictureBox area 
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   495
      Index           =   0
      Left            =   5160
      ScaleHeight     =   435
      ScaleWidth      =   2235
      TabIndex        =   23
      Top             =   9480
      Visible         =   0   'False
      Width           =   2295
   End
   Begin VB.Timer Timerbuilding 
      Enabled         =   0   'False
      Interval        =   5000
      Left            =   14400
      Top             =   7680
   End
   Begin VB.PictureBox buildingactive 
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   615
      Index           =   0
      Left            =   4440
      ScaleHeight     =   555
      ScaleWidth      =   4755
      TabIndex        =   24
      Top             =   7440
      Visible         =   0   'False
      Width           =   4815
   End
   Begin VB.PictureBox DataGrid3 
      Height          =   2175
      Left            =   11760
      ScaleHeight     =   2115
      ScaleWidth      =   8235
      TabIndex        =   12
      Top             =   8640
      Width           =   8295
   End
   Begin VB.PictureBox TotalTemp 
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   495
      Left            =   720
      ScaleHeight     =   435
      ScaleWidth      =   1995
      TabIndex        =   25
      Top             =   7800
      Visible         =   0   'False
      Width           =   2055
   End
   Begin VB.PictureBox test 
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   855
      Left            =   360
      ScaleHeight     =   795
      ScaleWidth      =   2475
      TabIndex        =   26
      Top             =   10080
      Visible         =   0   'False
      Width           =   2535
   End
   Begin VB.PictureBox TotalData 
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   855
      Left            =   600
      ScaleHeight     =   795
      ScaleWidth      =   2115
      TabIndex        =   27
      Top             =   8520
      Visible         =   0   'False
      Width           =   2175
   End
   Begin VB.Timer Timer3 
      Interval        =   5000
      Left            =   240
      Top             =   8400
   End
   Begin VB.PictureBox Adodc2 
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   615
      Left            =   5040
      ScaleHeight     =   555
      ScaleWidth      =   5355
      TabIndex        =   28
      Top             =   6240
      Visible         =   0   'False
      Width           =   5415
   End
   Begin VB.Frame Frame4 
      BackColor       =   &H00404000&
      Caption         =   "目录导航"
      BeginProperty Font 
         Name            =   "华文中宋"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   8295
      Left            =   120
      TabIndex        =   3
      Top             =   2640
      Width           =   3015
      Begin VB.PictureBox TreeView1 
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   12
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   6855
         Left            =   240
         ScaleHeight     =   6795
         ScaleWidth      =   2475
         TabIndex        =   4
         Top             =   600
         Width           =   2535
      End
   End
   Begin VB.PictureBox Adodc1 
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   495
      Left            =   3840
      ScaleHeight     =   435
      ScaleWidth      =   2835
      TabIndex        =   29
      Top             =   3360
      Visible         =   0   'False
      Width           =   2895
   End
   Begin VB.Frame Frame3 
      BackColor       =   &H00404000&
      Caption         =   "温度、湿度"
      ForeColor       =   &H8000000B&
      Height          =   2535
      Left            =   3120
      TabIndex        =   2
      Top             =   8400
      Width           =   17025
      Begin VB.PictureBox DataGrid2 
         Height          =   2175
         Left            =   120
         ScaleHeight     =   2115
         ScaleWidth      =   8235
         TabIndex        =   6
         Top             =   240
         Width           =   8295
      End
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00FFFFFF&
      Caption         =   "能耗数据显示区"
      ForeColor       =   &H00FFFFFF&
      Height          =   4575
      Left            =   3120
      TabIndex        =   1
      Top             =   1320
      Width           =   17025
      Begin VB.PictureBox CURRENTChart 
         Height          =   2055
         Left            =   9360
         ScaleHeight     =   1995
         ScaleWidth      =   4395
         TabIndex        =   8
         Top             =   2520
         Visible         =   0   'False
         Width           =   4455
      End
      Begin VB.PictureBox VOLTAGEchart 
         Height          =   2055
         Left            =   3480
         ScaleHeight     =   1995
         ScaleWidth      =   4395
         TabIndex        =   7
         Top             =   2520
         Visible         =   0   'False
         Width           =   4455
      End
      Begin VB.PictureBox ACTIVEENERGYChart 
         Height          =   2415
         Left            =   120
         ScaleHeight     =   2355
         ScaleWidth      =   4515
         TabIndex        =   9
         Top             =   240
         Width           =   4575
      End
      Begin VB.PictureBox TEMPChart 
         Height          =   2415
         Left            =   6360
         ScaleHeight     =   2355
         ScaleWidth      =   4515
         TabIndex        =   10
         Top             =   240
         Width           =   4575
      End
      Begin VB.PictureBox HUMIDITYChart 
         Height          =   2415
         Left            =   12480
         ScaleHeight     =   2355
         ScaleWidth      =   4515
         TabIndex        =   11
         Top             =   240
         Width           =   4575
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00404000&
      Caption         =   "电量、功率、电流、电压"
      ForeColor       =   &H8000000B&
      Height          =   2535
      Left            =   3120
      TabIndex        =   0
      Top             =   5880
      Width           =   17025
      Begin VB.PictureBox plug 
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   495
         Index           =   0
         Left            =   7440
         ScaleHeight     =   435
         ScaleWidth      =   1875
         TabIndex        =   30
         Top             =   600
         Visible         =   0   'False
         Width           =   1935
      End
      Begin VB.PictureBox DataGrid1 
         Height          =   2895
         Left            =   120
         ScaleHeight     =   2835
         ScaleWidth      =   16755
         TabIndex        =   5
         Top             =   240
         Width           =   16815
      End
   End
   Begin VB.PictureBox buildingactive 
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   615
      Index           =   1
      Left            =   120
      ScaleHeight     =   555
      ScaleWidth      =   2355
      TabIndex        =   31
      Top             =   8040
      Visible         =   0   'False
      Width           =   2415
   End
   Begin VB.PictureBox area 
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   495
      Index           =   1
      Left            =   120
      ScaleHeight     =   435
      ScaleWidth      =   2235
      TabIndex        =   32
      Top             =   10320
      Visible         =   0   'False
      Width           =   2295
   End
   Begin VB.PictureBox collector 
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   330
      Index           =   1
      Left            =   4080
      ScaleHeight     =   270
      ScaleWidth      =   2955
      TabIndex        =   33
      Top             =   4440
      Visible         =   0   'False
      Width           =   3015
   End
   Begin VB.PictureBox plug 
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   495
      Index           =   1
      Left            =   11280
      ScaleHeight     =   435
      ScaleWidth      =   1875
      TabIndex        =   34
      Top             =   4440
      Visible         =   0   'False
      Width           =   1935
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "◎基于移动IPv6的智慧校园建筑用能数据采集终端Ver2.0"
      ForeColor       =   &H000000FF&
      Height          =   255
      Index           =   1
      Left            =   15360
      TabIndex        =   21
      Top             =   960
      Width           =   4815
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "欢迎使用本系统"
      ForeColor       =   &H00FFFFFF&
      Height          =   180
      Index           =   1
      Left            =   480
      TabIndex        =   18
      Top             =   2280
      Width           =   1260
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "您好！管理员"
      ForeColor       =   &H00FFFFFF&
      Height          =   180
      Index           =   0
      Left            =   960
      TabIndex        =   17
      Top             =   1800
      Width           =   1080
   End
   Begin VB.Image Image2 
      Height          =   495
      Left            =   360
      Picture         =   "欢迎使用.frx":0000
      Stretch         =   -1  'True
      Top             =   1680
      Width           =   495
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "An Hui Jian Zhu University"
      BeginProperty Font 
         Name            =   "隶书"
         Size            =   10.5
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   2040
      TabIndex        =   16
      Top             =   960
      Width           =   3495
   End
   Begin VB.Line Line2 
      BorderColor     =   &H000000FF&
      X1              =   5640
      X2              =   5640
      Y1              =   0
      Y2              =   1200
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "建筑用能数据采集系统"
      BeginProperty Font 
         Name            =   "华文楷体"
         Size            =   21.75
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   615
      Index           =   1
      Left            =   9120
      TabIndex        =   15
      Top             =   360
      Width           =   4575
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "安徽建筑大学"
      BeginProperty Font 
         Name            =   "华文行楷"
         Size            =   21.75
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   495
      Index           =   0
      Left            =   2160
      TabIndex        =   14
      Top             =   360
      Width           =   3135
   End
   Begin VB.Line Line1 
      BorderColor     =   &H000000FF&
      X1              =   1680
      X2              =   5640
      Y1              =   840
      Y2              =   840
   End
   Begin VB.Image Image1 
      Height          =   1245
      Left            =   120
      Picture         =   "欢迎使用.frx":348A
      Stretch         =   -1  'True
      Top             =   0
      Width           =   1245
   End
   Begin VB.Label Label5 
      BackColor       =   &H00404000&
      Height          =   1455
      Left            =   120
      TabIndex        =   19
      Top             =   1320
      Width           =   3015
   End
   Begin VB.Label Label6 
      BackColor       =   &H00FF8080&
      Height          =   10935
      Left            =   -480
      TabIndex        =   20
      Top             =   0
      Width           =   20730
   End
End
Attribute VB_Name = "smartplus"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim pluglastTime
Dim pluglastTime2
Dim i As Single
Dim LastTimeR As New ADODB.Recordset '用于取最后一条电量等表的记录时间的记录对象
Dim LastTimeR2 As New ADODB.Recordset '用于取最后一条温度等表的记录时间的记录对象
Dim strCodeplugmacAddress









Private Sub Command1_Click()
 MsgBox "用户注销成功"
    smartplus.Hide
    Form1.Show
End Sub

Private Sub Form_Initialize()
Adodc1.ConnectionString = "Provider=SQLOLEDB.1;Persist Security Info=False;User ID=sa;Password=123456;Initial Catalog=shuihu;Data Source=127.0.0.1"

Adodc1.CommandType = adCmdText
TotalData.ConnectionString = "Provider=SQLOLEDB.1;Persist Security Info=False;User ID=sa;Password=123456;Initial Catalog=shuihu;Data Source=127.0.0.1" '全部插座信息表的ado控件
TotalTemp.ConnectionString = "Provider=SQLOLEDB.1;Persist Security Info=False;User ID=sa;Password=123456;Initial Catalog=shuihu;Data Source=127.0.0.1" '全部插座信息表的ado控件
TotalData.CommandType = adCmdText
TotalTemp.CommandType = adCmdText
buildingactive(0).ConnectionString = "Provider=SQLOLEDB.1;Persist Security Info=False;User ID=sa;Password=123456;Initial Catalog=shuihu;Data Source=127.0.0.1"
buildingactive(0).CommandType = adCmdText
buildingactive(1).ConnectionString = "Provider=SQLOLEDB.1;Persist Security Info=False;User ID=sa;Password=123456;Initial Catalog=shuihu;Data Source=127.0.0.1"
buildingactive(1).CommandType = adCmdText
area(0).ConnectionString = "Provider=SQLOLEDB.1;Persist Security Info=False;User ID=sa;Password=123456;Initial Catalog=shuihu;Data Source=127.0.0.1"
area(0).CommandType = adCmdText
area(1).ConnectionString = "Provider=SQLOLEDB.1;Persist Security Info=False;User ID=sa;Password=123456;Initial Catalog=shuihu;Data Source=127.0.0.1"
area(1).CommandType = adCmdText
collector(0).ConnectionString = "Provider=SQLOLEDB.1;Persist Security Info=False;User ID=sa;Password=123456;Initial Catalog=shuihu;Data Source=127.0.0.1"
collector(0).CommandType = adCmdText
collector(1).ConnectionString = "Provider=SQLOLEDB.1;Persist Security Info=False;User ID=sa;Password=123456;Initial Catalog=shuihu;Data Source=127.0.0.1"
collector(1).CommandType = adCmdText

plug(0).ConnectionString = "Provider=SQLOLEDB.1;Persist Security Info=False;User ID=sa;Password=123456;Initial Catalog=shuihu;Data Source=127.0.0.1"
plug(0).CommandType = adCmdText
plug(1).ConnectionString = "Provider=SQLOLEDB.1;Persist Security Info=False;User ID=sa;Password=123456;Initial Catalog=shuihu;Data Source=127.0.0.1"
plug(1).CommandType = adCmdText

Adodc2.ConnectionString = "Provider=SQLOLEDB.1;Persist Security Info=False;User ID=sa;Password=123456;Initial Catalog=shuihu;Data Source=127.0.0.1"

Adodc2.CommandType = adCmdText
TotalData.RecordSource = "Select top 100 ACTIVEPOWER,REACTIVEPOWER,VOLTAGE,[CURRENT],ACTIVEENERGY,processTime,DBTIME,MAC from analyze_electrictable order by DBTIME desc"
TotalTemp.RecordSource = "select top 100 TEMP,HUMIDITY,MAC,COLLECTIP,DBTIME from analyze_humituretable order by DBTIME desc"
  
  TotalData.Refresh
  TotalTemp.Refresh



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
 
  '*+++++++++++++++++++++++++++++++++++++++++++++获取建筑物电量等表最后一条记录的时间

 LastTimeR.Open "Select max(DBTIME) from analyze_electrictable join collectorInformation join areaInformation on collectorInformation.AID=areaInformation.AID on COLLECTIP=collectorInformation.macAddress  group by DBTIME,areaInformation.BID Order by DBTIME", buildingactive(0).ConnectionString
    Do Until LastTimeR.EOF
    buildinglastTime = CStr(LastTimeR(0))
  LastTimeR.MoveNext
  Loop

 LastTimeR.Close
  '*+++++++++++++++++++++++++++++++++++++++++++++获取建筑物最后一条记录的时间
   '*+++++++++++++++++++++++++++++++++++++++++++++获取采建筑物温度等表最后一条记录的时间
 '
 LastTimeR2.Open "select max(DBTIME) from analyze_humituretable join collectorInformation join areaInformation on collectorInformation.AID=areaInformation.AID on COLLECTIP=collectorInformation.macAddress  group by DBTIME,areaInformation.BID Order by DBTIME", buildingactive(1).ConnectionString
    Do Until LastTimeR2.EOF
    buildinglastTime2 = CStr(LastTimeR2(0))
  LastTimeR2.MoveNext
  Loop

 LastTimeR2.Close
  '*+++++++++++++++++++++++++++++++++++++++++++++获取建筑物温度等表最后一条记录的时间
 
 
 

 buildingactive(0).RecordSource = "Select sum(ACTIVEENERGY) as 电量汇总,DBTIME,areaInformation.BID from analyze_electrictable join collectorInformation join areaInformation on collectorInformation.AID=areaInformation.AID on COLLECTIP=collectorInformation.macAddress where DBTIME>dateadd(MINUTE,-10,cast('" & buildinglastTime & "'as datetime)) group by DBTIME,areaInformation.BID Order by DBTIME"
             buildingactive(1).RecordSource = "Select avg(cast(TEMP AS int)) as 温度平均值,avg(cast(HUMIDITY as int)) as 湿度平均值,CONVERT(varchar,DBTIME,120),areaInformation.BID from analyze_humituretable join collectorInformation join areaInformation on collectorInformation.AID=areaInformation.AID on COLLECTIP=collectorInformation.macAddress where DBTIME>dateadd(MINUTE,-10,cast('" & buildinglastTime2 & "'as datetime)) group by CONVERT(varchar,DBTIME,120),areaInformation.BID Order by CONVERT(varchar,DBTIME,120)"
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
 
 
  '*+++++++++++++++++++++++++++++++++++++++++++++获取区域电量等表最后一条记录的时间

 LastTimeR.Open "Select max(DBTIME) from analyze_electrictable join collectorInformation on COLLECTIP=collectorInformation.macAddress group by DBTIME,collectorInformation.AID ", area(0).ConnectionString
    Do Until LastTimeR.EOF
    arealastTime = CStr(LastTimeR(0))
  LastTimeR.MoveNext
  Loop

 LastTimeR.Close
  '*+++++++++++++++++++++++++++++++++++++++++++++获取区域最后一条记录的时间
   '*+++++++++++++++++++++++++++++++++++++++++++++获取区域温度等表最后一条记录的时间
 '
 LastTimeR2.Open "select max(DBTIME) from analyze_humituretable join collectorInformation on COLLECTIP=collectorInformation.macAddress  group by DBTIME,collectorInformation.AID ", area(1).ConnectionString
    Do Until LastTimeR2.EOF
    arealastTime2 = CStr(LastTimeR2(0))
  LastTimeR2.MoveNext
  Loop

 LastTimeR2.Close
  '*+++++++++++++++++++++++++++++++++++++++++++++获取区域温度等表最后一条记录的时间
 
       
        area(0).RecordSource = "Select sum(ACTIVEENERGY) as 电量汇总,DBTIME,collectorInformation.AID from analyze_electrictable join collectorInformation on COLLECTIP=collectorInformation.macAddress where DBTIME>dateadd(MINUTE,-10,cast('" & arealastTime & "'as datetime)) group by DBTIME,collectorInformation.AID Order by DBTIME"
        area(1).RecordSource = "Select avg(cast(TEMP AS int)) as 温度平均值,avg(cast(HUMIDITY as int)) as 湿度平均值,CONVERT(varchar,DBTIME,120),collectorInformation.AID from analyze_humituretable join collectorInformation on COLLECTIP=collectorInformation.macAddress  where  DBTIME>dateadd(MINUTE,-10,cast('" & arealastTime2 & "'as datetime)) group by CONVERT(varchar,DBTIME,120),collectorInformation.AID Order by CONVERT(varchar,DBTIME,120)"
       '   area(1).RecordSource = "Select avg(cast(TEMP AS int)) as 温度平均值,avg(cast(HUMIDITY as int)) as 湿度平均值,DBTIME,collectorInformation.AID from analyze_humituretable join collectorInformation on COLLECTIP=collectorInformation.macAddress  where  DBTIME>dateadd(MINUTE,-10,cast('" & arealastTime2 & "'as datetime)) group by DBTIME,collectorInformation.AID Order by DBTIME"
    
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
 
  
 '*+++++++++++++++++++++++++++++++++++++++++++++获取采集器电量等表最后一条记录的时间

 LastTimeR.Open "Select max(DBTIME) from analyze_electrictable where  COLLECTIP='" & strCodeCOLLECTIP & "' group by COLLECTIP ", collector(0).ConnectionString
    Do Until LastTimeR.EOF
    collectorlastTime = CStr(LastTimeR(0))
  LastTimeR.MoveNext
  Loop
 Print collectorlastTime
 LastTimeR.Close
  '*+++++++++++++++++++++++++++++++++++++++++++++获取采集器最后一条记录的时间
   '*+++++++++++++++++++++++++++++++++++++++++++++获取采集器温度等表最后一条记录的时间
 '
 LastTimeR2.Open "select max(DBTIME) from analyze_humituretable  where COLLECTIP='" & strCodeCOLLECTIP & "'group by COLLECTIP", collector(1).ConnectionString
    Do Until LastTimeR2.EOF
    collectorlastTime2 = CStr(LastTimeR2(0))
  LastTimeR2.MoveNext
  Loop
 Print collectorlastTime2
 LastTimeR2.Close
  '*+++++++++++++++++++++++++++++++++++++++++++++获取采集器最后一条记录的时间
 collector(0).RecordSource = "Select sum([CURRENT]) as 电流汇总,DBTIME,COLLECTIP,sum(VOLTAGE) as 电压汇总,sum(ACTIVEENERGY) as 电量汇总 from analyze_electrictable where COLLECTIP='" & strCodeCOLLECTIP & "' and DBTIME>dateadd(MINUTE,-30,cast('" & collectorlastTime & "'as datetime)) group by DBTIME,COLLECTIP Order by DBTIME"

       
collector(1).RecordSource = "Select avg(cast(TEMP AS int)) as 温度平均值,avg(cast(HUMIDITY as int)) as 湿度平均值,CONVERT(varchar,DBTIME,120),COLLECTIP from analyze_humituretable where COLLECTIP='" & strCodeCOLLECTIP & "' and DBTIME>dateadd(MINUTE,-10,cast('" & collectorlastTime2 & "'as datetime))group by CONVERT(varchar,DBTIME,120),COLLECTIP"

          collector(0).Refresh
          collector(1).Refresh
        
        
           
        
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
 
 
  LastTimeR.Open "Select max(DBTIME) from analyze_electrictable where MAC='" & strCodeplugmacAddress & "'", plug(0).ConnectionString

    Do Until LastTimeR.EOF
    pluglastTime = CStr(LastTimeR(0))
  LastTimeR.MoveNext
  Loop
 Print pluglastTime
 LastTimeR.Close
  '*+++++++++++++++++++++++++++++++++++++++++++++获取插座最后一条记录的时间
   '*+++++++++++++++++++++++++++++++++++++++++++++获取插座电量等表最后一条记录的时间
 '
 LastTimeR2.Open "select max(DBTIME) from analyze_humituretable  where MAC='" & strCodeplugmacAddress & "'", plug(1).ConnectionString
    Do Until LastTimeR2.EOF
    pluglastTime2 = CStr(LastTimeR2(0))
  LastTimeR2.MoveNext
  Loop
 Print pluglastTime2
 LastTimeR2.Close
  '*+++++++++++++++++++++++++++++++++++++++++++++获取插座最后一条记录的时间
  
  

plug(0).RecordSource = "Select ACTIVEPOWER,REACTIVEPOWER,VOLTAGE,[CURRENT],ACTIVEENERGY,DBTIME,MAC from analyze_electrictable where MAC='" & strCodeplugmacAddress & "'and DBTIME>dateadd(MINUTE,-10,cast('" & pluglastTime & "'as datetime)) Order by DBTIME"


'Adodc2.RecordSource = "select TEMP,HUMIDITY,MAC,COLLECTIP,DBTIME from analyze_humituretable  where DBTIME>dateadd(MINUTE,-10,GETDATE())"  '显示近10分钟的插座温度，演示系统时要把系统时间设定为10-21 9：01：50，因为这是数据库记录的最后一条记录的插入时间
plug(1).RecordSource = "select TEMP,HUMIDITY,MAC,COLLECTIP,DBTIME from analyze_humituretable  where MAC='" & strCodeplugmacAddress & "' and DBTIME>dateadd(MINUTE,-10,cast('" & pluglastTime2 & "'as datetime)) Order by DBTIME"  '显示近10分钟的插座温度，演示系统时要把系统时间设定为10-21 9：01：50，因为这是数据库记录的最后一条记录的插入时间

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



  'MSChart1.Plot.UniformAxis = False

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

