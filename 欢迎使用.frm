VERSION 5.00
Begin VB.Form smartplus 
   BackColor       =   &H00FFFFFF&
   Caption         =   "�������ܲɼ�ϵͳ"
   ClientHeight    =   10935
   ClientLeft      =   2865
   ClientTop       =   255
   ClientWidth     =   15120
   LinkTopic       =   "Form1"
   ScaleHeight     =   10935
   ScaleWidth      =   15120
   StartUpPosition =   2  '��Ļ����
   Begin VB.CommandButton Command1 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000D&
      Caption         =   "ע����¼"
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
      Caption         =   "Ŀ¼����"
      BeginProperty Font 
         Name            =   "��������"
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
            Name            =   "����"
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
      Caption         =   "�¶ȡ�ʪ��"
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
      Caption         =   "�ܺ�������ʾ��"
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
      Caption         =   "���������ʡ���������ѹ"
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
      Caption         =   "������ƶ�IPv6���ǻ�У԰�����������ݲɼ��ն�Ver2.0"
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
      Caption         =   "��ӭʹ�ñ�ϵͳ"
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
      Caption         =   "���ã�����Ա"
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
      Picture         =   "��ӭʹ��.frx":0000
      Stretch         =   -1  'True
      Top             =   1680
      Width           =   495
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "An Hui Jian Zhu University"
      BeginProperty Font 
         Name            =   "����"
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
      Caption         =   "�����������ݲɼ�ϵͳ"
      BeginProperty Font 
         Name            =   "���Ŀ���"
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
      Caption         =   "���ս�����ѧ"
      BeginProperty Font 
         Name            =   "�����п�"
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
      Picture         =   "��ӭʹ��.frx":348A
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
Dim LastTimeR As New ADODB.Recordset '����ȡ���һ�������ȱ�ļ�¼ʱ��ļ�¼����
Dim LastTimeR2 As New ADODB.Recordset '����ȡ���һ���¶ȵȱ�ļ�¼ʱ��ļ�¼����
Dim strCodeplugmacAddress









Private Sub Command1_Click()
 MsgBox "�û�ע���ɹ�"
    smartplus.Hide
    Form1.Show
End Sub

Private Sub Form_Initialize()
Adodc1.ConnectionString = "Provider=SQLOLEDB.1;Persist Security Info=False;User ID=sa;Password=123456;Initial Catalog=shuihu;Data Source=127.0.0.1"

Adodc1.CommandType = adCmdText
TotalData.ConnectionString = "Provider=SQLOLEDB.1;Persist Security Info=False;User ID=sa;Password=123456;Initial Catalog=shuihu;Data Source=127.0.0.1" 'ȫ��������Ϣ���ado�ؼ�
TotalTemp.ConnectionString = "Provider=SQLOLEDB.1;Persist Security Info=False;User ID=sa;Password=123456;Initial Catalog=shuihu;Data Source=127.0.0.1" 'ȫ��������Ϣ���ado�ؼ�
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
 
  '*+++++++++++++++++++++++++++++++++++++++++++++��ȡ����������ȱ����һ����¼��ʱ��

 LastTimeR.Open "Select max(DBTIME) from analyze_electrictable join collectorInformation join areaInformation on collectorInformation.AID=areaInformation.AID on COLLECTIP=collectorInformation.macAddress  group by DBTIME,areaInformation.BID Order by DBTIME", buildingactive(0).ConnectionString
    Do Until LastTimeR.EOF
    buildinglastTime = CStr(LastTimeR(0))
  LastTimeR.MoveNext
  Loop

 LastTimeR.Close
  '*+++++++++++++++++++++++++++++++++++++++++++++��ȡ���������һ����¼��ʱ��
   '*+++++++++++++++++++++++++++++++++++++++++++++��ȡ�ɽ������¶ȵȱ����һ����¼��ʱ��
 '
 LastTimeR2.Open "select max(DBTIME) from analyze_humituretable join collectorInformation join areaInformation on collectorInformation.AID=areaInformation.AID on COLLECTIP=collectorInformation.macAddress  group by DBTIME,areaInformation.BID Order by DBTIME", buildingactive(1).ConnectionString
    Do Until LastTimeR2.EOF
    buildinglastTime2 = CStr(LastTimeR2(0))
  LastTimeR2.MoveNext
  Loop

 LastTimeR2.Close
  '*+++++++++++++++++++++++++++++++++++++++++++++��ȡ�������¶ȵȱ����һ����¼��ʱ��
 
 
 

 buildingactive(0).RecordSource = "Select sum(ACTIVEENERGY) as ��������,DBTIME,areaInformation.BID from analyze_electrictable join collectorInformation join areaInformation on collectorInformation.AID=areaInformation.AID on COLLECTIP=collectorInformation.macAddress where DBTIME>dateadd(MINUTE,-10,cast('" & buildinglastTime & "'as datetime)) group by DBTIME,areaInformation.BID Order by DBTIME"
             buildingactive(1).RecordSource = "Select avg(cast(TEMP AS int)) as �¶�ƽ��ֵ,avg(cast(HUMIDITY as int)) as ʪ��ƽ��ֵ,CONVERT(varchar,DBTIME,120),areaInformation.BID from analyze_humituretable join collectorInformation join areaInformation on collectorInformation.AID=areaInformation.AID on COLLECTIP=collectorInformation.macAddress where DBTIME>dateadd(MINUTE,-10,cast('" & buildinglastTime2 & "'as datetime)) group by CONVERT(varchar,DBTIME,120),areaInformation.BID Order by CONVERT(varchar,DBTIME,120)"
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
 
 
  '*+++++++++++++++++++++++++++++++++++++++++++++��ȡ��������ȱ����һ����¼��ʱ��

 LastTimeR.Open "Select max(DBTIME) from analyze_electrictable join collectorInformation on COLLECTIP=collectorInformation.macAddress group by DBTIME,collectorInformation.AID ", area(0).ConnectionString
    Do Until LastTimeR.EOF
    arealastTime = CStr(LastTimeR(0))
  LastTimeR.MoveNext
  Loop

 LastTimeR.Close
  '*+++++++++++++++++++++++++++++++++++++++++++++��ȡ�������һ����¼��ʱ��
   '*+++++++++++++++++++++++++++++++++++++++++++++��ȡ�����¶ȵȱ����һ����¼��ʱ��
 '
 LastTimeR2.Open "select max(DBTIME) from analyze_humituretable join collectorInformation on COLLECTIP=collectorInformation.macAddress  group by DBTIME,collectorInformation.AID ", area(1).ConnectionString
    Do Until LastTimeR2.EOF
    arealastTime2 = CStr(LastTimeR2(0))
  LastTimeR2.MoveNext
  Loop

 LastTimeR2.Close
  '*+++++++++++++++++++++++++++++++++++++++++++++��ȡ�����¶ȵȱ����һ����¼��ʱ��
 
       
        area(0).RecordSource = "Select sum(ACTIVEENERGY) as ��������,DBTIME,collectorInformation.AID from analyze_electrictable join collectorInformation on COLLECTIP=collectorInformation.macAddress where DBTIME>dateadd(MINUTE,-10,cast('" & arealastTime & "'as datetime)) group by DBTIME,collectorInformation.AID Order by DBTIME"
        area(1).RecordSource = "Select avg(cast(TEMP AS int)) as �¶�ƽ��ֵ,avg(cast(HUMIDITY as int)) as ʪ��ƽ��ֵ,CONVERT(varchar,DBTIME,120),collectorInformation.AID from analyze_humituretable join collectorInformation on COLLECTIP=collectorInformation.macAddress  where  DBTIME>dateadd(MINUTE,-10,cast('" & arealastTime2 & "'as datetime)) group by CONVERT(varchar,DBTIME,120),collectorInformation.AID Order by CONVERT(varchar,DBTIME,120)"
       '   area(1).RecordSource = "Select avg(cast(TEMP AS int)) as �¶�ƽ��ֵ,avg(cast(HUMIDITY as int)) as ʪ��ƽ��ֵ,DBTIME,collectorInformation.AID from analyze_humituretable join collectorInformation on COLLECTIP=collectorInformation.macAddress  where  DBTIME>dateadd(MINUTE,-10,cast('" & arealastTime2 & "'as datetime)) group by DBTIME,collectorInformation.AID Order by DBTIME"
    
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
 
  
 '*+++++++++++++++++++++++++++++++++++++++++++++��ȡ�ɼ��������ȱ����һ����¼��ʱ��

 LastTimeR.Open "Select max(DBTIME) from analyze_electrictable where  COLLECTIP='" & strCodeCOLLECTIP & "' group by COLLECTIP ", collector(0).ConnectionString
    Do Until LastTimeR.EOF
    collectorlastTime = CStr(LastTimeR(0))
  LastTimeR.MoveNext
  Loop
 Print collectorlastTime
 LastTimeR.Close
  '*+++++++++++++++++++++++++++++++++++++++++++++��ȡ�ɼ������һ����¼��ʱ��
   '*+++++++++++++++++++++++++++++++++++++++++++++��ȡ�ɼ����¶ȵȱ����һ����¼��ʱ��
 '
 LastTimeR2.Open "select max(DBTIME) from analyze_humituretable  where COLLECTIP='" & strCodeCOLLECTIP & "'group by COLLECTIP", collector(1).ConnectionString
    Do Until LastTimeR2.EOF
    collectorlastTime2 = CStr(LastTimeR2(0))
  LastTimeR2.MoveNext
  Loop
 Print collectorlastTime2
 LastTimeR2.Close
  '*+++++++++++++++++++++++++++++++++++++++++++++��ȡ�ɼ������һ����¼��ʱ��
 collector(0).RecordSource = "Select sum([CURRENT]) as ��������,DBTIME,COLLECTIP,sum(VOLTAGE) as ��ѹ����,sum(ACTIVEENERGY) as �������� from analyze_electrictable where COLLECTIP='" & strCodeCOLLECTIP & "' and DBTIME>dateadd(MINUTE,-30,cast('" & collectorlastTime & "'as datetime)) group by DBTIME,COLLECTIP Order by DBTIME"

       
collector(1).RecordSource = "Select avg(cast(TEMP AS int)) as �¶�ƽ��ֵ,avg(cast(HUMIDITY as int)) as ʪ��ƽ��ֵ,CONVERT(varchar,DBTIME,120),COLLECTIP from analyze_humituretable where COLLECTIP='" & strCodeCOLLECTIP & "' and DBTIME>dateadd(MINUTE,-10,cast('" & collectorlastTime2 & "'as datetime))group by CONVERT(varchar,DBTIME,120),COLLECTIP"

          collector(0).Refresh
          collector(1).Refresh
        
        
           
        
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
 
 
  LastTimeR.Open "Select max(DBTIME) from analyze_electrictable where MAC='" & strCodeplugmacAddress & "'", plug(0).ConnectionString

    Do Until LastTimeR.EOF
    pluglastTime = CStr(LastTimeR(0))
  LastTimeR.MoveNext
  Loop
 Print pluglastTime
 LastTimeR.Close
  '*+++++++++++++++++++++++++++++++++++++++++++++��ȡ�������һ����¼��ʱ��
   '*+++++++++++++++++++++++++++++++++++++++++++++��ȡ���������ȱ����һ����¼��ʱ��
 '
 LastTimeR2.Open "select max(DBTIME) from analyze_humituretable  where MAC='" & strCodeplugmacAddress & "'", plug(1).ConnectionString
    Do Until LastTimeR2.EOF
    pluglastTime2 = CStr(LastTimeR2(0))
  LastTimeR2.MoveNext
  Loop
 Print pluglastTime2
 LastTimeR2.Close
  '*+++++++++++++++++++++++++++++++++++++++++++++��ȡ�������һ����¼��ʱ��
  
  

plug(0).RecordSource = "Select ACTIVEPOWER,REACTIVEPOWER,VOLTAGE,[CURRENT],ACTIVEENERGY,DBTIME,MAC from analyze_electrictable where MAC='" & strCodeplugmacAddress & "'and DBTIME>dateadd(MINUTE,-10,cast('" & pluglastTime & "'as datetime)) Order by DBTIME"


'Adodc2.RecordSource = "select TEMP,HUMIDITY,MAC,COLLECTIP,DBTIME from analyze_humituretable  where DBTIME>dateadd(MINUTE,-10,GETDATE())"  '��ʾ��10���ӵĲ����¶ȣ���ʾϵͳʱҪ��ϵͳʱ���趨Ϊ10-21 9��01��50����Ϊ�������ݿ��¼�����һ����¼�Ĳ���ʱ��
plug(1).RecordSource = "select TEMP,HUMIDITY,MAC,COLLECTIP,DBTIME from analyze_humituretable  where MAC='" & strCodeplugmacAddress & "' and DBTIME>dateadd(MINUTE,-10,cast('" & pluglastTime2 & "'as datetime)) Order by DBTIME"  '��ʾ��10���ӵĲ����¶ȣ���ʾϵͳʱҪ��ϵͳʱ���趨Ϊ10-21 9��01��50����Ϊ�������ݿ��¼�����һ����¼�Ĳ���ʱ��

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



  'MSChart1.Plot.UniformAxis = False

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

