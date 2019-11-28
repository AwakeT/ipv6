Attribute VB_Name = "Data_sql"
Public conn As ADODB.Connection
Public rs As ADODB.Recordset
Public addFlag As Boolean      '声明部分
'Public Function OpenCn(ByVal Cip As String, ByVal users As String, ByVal pw As String) As Boolean '连接模块 填写数据库等信息
Dim mag As String

On Error GoTo strerrmag
Set conn = New ADODB.Connection
conn.ConnectionTimeout = 25
conn.Provider = "sqloledb"
conn.Properties("data source").Value = Cip    '服务器的名字
conn.Properties("initial catalog").Value = "pubs"          '库名
conn.Properties("integrated security").Value = "SSPI"      '登陆类型
conn.Properties("user id").Value = users 'SQL库名"
conn.Properties("password").Value = pw '密码
'SQL = "Provider=SQLOLEDB.1;Persist Security Info=False;User ID=sa;password=;Initial Catalog=master;Data Source=219.231.15.100"    '如果不用这个模块也行可以，这一句便是常用的引擎
'SQL = "Provider=SQLOLEDB.1;Integrated Security=SSPI;Persist Security Info=False;Initial Catalog=master;Data Source=DENJIN"
'conn.ConnectionString = SQL
conn.Open OpenCn = True
If conn.State = 1 Then addFlag = True
Exit Function
strerrmag:
     mag = "Data can't connect"
     Call MsgBox(mag, vbOKOnly, "Error:Data connect")
     addFlag = False
     Exit Function      '连接错误消息 End Function
'关闭数据库，释放连接 Public Sub cloCn() On Error Resume Next
If conn.State <> adStateClosed Then conn.Close
Set conn = Nothing
End Sub
Public Function openRs(ByVal strsql As String) As Boolean      '连接数据库记录集
Dim mag As String
Dim rpy As Boolean
On Error GoTo strerrmag
     Set rs = New ADODB.Recordset
     If addFlag = False Then rpy = True

 With rs
     .ActiveConnection = conn
     .CursorLocation = adUseClient
     .CursorType = adOpenKeyset
     .LockType = adLockOptimistic
     .Open strsql
End With
     addFlag = True
     openRs = True
 'End        '将记录集给rs
 Exit Function
strerrmag:
     mag = "data not connect"
     Call MsgBox(mag, vbOKOnly, "error:connect")
     openRs = False
     End
     'Exit Function '连接错误消息
     End Function
Public Sub cloRs()

End Sub
On Error Resume Next
If rs.State <> adStateClosed Then rs.Clone
Set rs = Nothing '释放记录集
End Sub
