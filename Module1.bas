Attribute VB_Name = "Data_sql"
Public conn As ADODB.Connection
Public rs As ADODB.Recordset
Public addFlag As Boolean      '��������
'Public Function OpenCn(ByVal Cip As String, ByVal users As String, ByVal pw As String) As Boolean '����ģ�� ��д���ݿ����Ϣ
Dim mag As String

On Error GoTo strerrmag
Set conn = New ADODB.Connection
conn.ConnectionTimeout = 25
conn.Provider = "sqloledb"
conn.Properties("data source").Value = Cip    '������������
conn.Properties("initial catalog").Value = "pubs"          '����
conn.Properties("integrated security").Value = "SSPI"      '��½����
conn.Properties("user id").Value = users 'SQL����"
conn.Properties("password").Value = pw '����
'SQL = "Provider=SQLOLEDB.1;Persist Security Info=False;User ID=sa;password=;Initial Catalog=master;Data Source=219.231.15.100"    '����������ģ��Ҳ�п��ԣ���һ����ǳ��õ�����
'SQL = "Provider=SQLOLEDB.1;Integrated Security=SSPI;Persist Security Info=False;Initial Catalog=master;Data Source=DENJIN"
'conn.ConnectionString = SQL
conn.Open OpenCn = True
If conn.State = 1 Then addFlag = True
Exit Function
strerrmag:
     mag = "Data can't connect"
     Call MsgBox(mag, vbOKOnly, "Error:Data connect")
     addFlag = False
     Exit Function      '���Ӵ�����Ϣ End Function
'�ر����ݿ⣬�ͷ����� Public Sub cloCn() On Error Resume Next
If conn.State <> adStateClosed Then conn.Close
Set conn = Nothing
End Sub
Public Function openRs(ByVal strsql As String) As Boolean      '�������ݿ��¼��
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
 'End        '����¼����rs
 Exit Function
strerrmag:
     mag = "data not connect"
     Call MsgBox(mag, vbOKOnly, "error:connect")
     openRs = False
     End
     'Exit Function '���Ӵ�����Ϣ
     End Function
Public Sub cloRs()

End Sub
On Error Resume Next
If rs.State <> adStateClosed Then rs.Clone
Set rs = Nothing '�ͷż�¼��
End Sub
