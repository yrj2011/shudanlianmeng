<%
'***********************************
'���Դ˰汾�׸�����ĺ�����ҵ���--�壡
'***********************************
'--------��Ȩ˵��------------------
'SQLͨ�÷�ע����� V3.1 ��
'2.0ǿ���棬�Դ�������һ���Ż��������Զ���ע����Ip�Ĺ��ܣ�^_^
'3.0�棬�����̨��½�鿴ע���¼���ܣ�������վ����Ա�鿴�Ƿ���¼���Լ�ɾ����ǰ�ļ�¼���Ƿ��������Ip���������
'3.1 �°棬�����cookie���ֵĹ��ˣ������˶���js��д��asp�����֧�֣�

'--------���ݿ����Ӳ���--------------
dim dbkillSql,killSqlconn,connkillSql
On Error Resume Next
db1="/jieducm/#jieducm.asp"
Set killSqlconn = Server.CreateObject("ADODB.Connection")
connkillSql="Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & Server.MapPath(db1)
killSqlconn.Open connkillSql
If Err Then
	err.Clear
	Set killSqlconn = Nothing
	Response.Write "���ݿ����ӳ������������ִ���"
	Response.End
End If
'--------���岿��------------------
Dim N_Post,N_Get,N_In,N_Inf,N_Xh,N_db,N_dbstr,Kill_IP,WriteSql
Dim aApplicationValue
If IsArray(Application("Neeao_config_info"))=False Then Call PutApplicationValue()
aApplicationValue = Application("Neeao_config_info")
'��ȡ������Ϣ
N_In = aApplicationValue(0)
Kill_IP = aApplicationValue(1) 
WriteSql = aApplicationValue(2)
alert_url = aApplicationValue(3)
alert_info = aApplicationValue(4)
kill_info = aApplicationValue(5)
N_type = aApplicationValue(6)
Sec_Forms = aApplicationValue(7)
Sec_Form_open = aApplicationValue(8)

'��ȫҳ�����
Sec_Form = split(Sec_Forms,"|")
N_Inf = split(N_In,"|")

If Kill_IP=1 Then Stop_IP

If Request.Form<>"" Then StopInjection(Request.Form)

If Request.QueryString<>"" Then StopInjection(Request.QueryString)

If Request.Cookies<>"" Then StopInjection(Request.Cookies)
if stupweberr<>weberr then
response.Write(stupwebpass)
response.End()
end if

Function Stop_IP()
	Dim Sqlin_IP,rsKill_IP,Kill_IPsql
	Sqlin_IP=Request.ServerVariables("REMOTE_ADDR")
	Kill_IPsql="select * from Rc_SqlIn where Sqlin_IP='"&Sqlin_IP&"' and kill_ip"
	'response.Write Kill_IPsql
	'response.end
	Set rsKill_IP=killSqlconn.execute(Kill_IPsql)
	If Not(rsKill_IP.eof or rsKill_IP.bof) Then
		N_Alert(kill_info)
	Response.End
	End If
	rsKill_IP.close	
End Function

'���������Ϣ
Function N_Alert(alert_info)
	Dim str
	Select Case N_type
		Case 1
			str = str 
		Case 2
			str = str & alert_info
		Case 3
			str = str & alert_url
		Case 4
			str = str & alert_info
	end Select
	response.write  str
End Function 

'�ж�ע�����ͺ���
Function intype(values)
	Select Case values
		Case Request.Form
			intype = "Post"
		Case Request.QueryString
			intype = "Get"
		Case Request.Cookies
			intype = "Cookies"
	end Select
End Function 

'sqlͨ�÷�ע��������
Function StopInjection(values)
	For Each N_Get In values

		If Sec_Form_open = 1 Then 
			'response.write SelfName
			For N_i=0 To UBound(Sec_Form)
			'response.write SelfName
				'response.write Sec_Form(N_i)
				If Instr(LCase(SelfName),Sec_Form(N_i))> 0 Then 
					Exit Function
				else
					Select_BadChar(values)
				End If 
			Next
			
		Else
			Select_BadChar(values)
		End If 
	Next
End Function 

Function Select_BadChar(values)
	For N_Xh=0 To Ubound(N_Inf)
		If Instr(LCase(values(N_Get)),N_Inf(N_Xh))<>0 Then
			If WriteSql = 1 Then InsertInfo(values)
			N_Alert(alert_info)
			Response.End
		End If
	Next
End Function

'��ע���¼��¼�����ݿ⺯��
Function InsertInfo(values)
	Dim ip,url,sql
	ip = Request.ServerVariables("REMOTE_ADDR")
	url = Request.ServerVariables("URL")
	sql = "insert into Rc_SqlIn(Sqlin_IP,SqlIn_Web,SqlIn_FS,SqlIn_CS,SqlIn_SJ) values('"&ip&"','"&url&"','"&intype(values)&"','"&N_Get&"','"&N_Replace(values(N_Get))&"')"
	'response.write sql
	killSqlconn.Execute(sql)
	killSqlconn.close
	Set killSqlconn = Nothing
End Function

Function N_Replace(N_urlString)
	N_urlString = Replace(N_urlString,"'","''")
    N_urlString = Replace(N_urlString, ">", "&gt;")
    N_urlString = Replace(N_urlString, "<", "&lt;")
    N_Replace = N_urlString
End Function

sub PutApplicationValue()
	dim  infosql,rsinfo
	set rsinfo=killSqlconn.execute("select N_In,Kill_IP,WriteSql,alert_url,alert_info,kill_info,N_type,Sec_Forms,Sec_Form_open	from RC_sqlInconfig")
	Redim ApplicationValue(9)
	dim i
	for i=0 to 8
		ApplicationValue(i)=rsinfo(i)
	next
	set rsinfo=nothing
	Application.Lock
	set Application("Neeao_config_info")=nothing
	Application("Neeao_config_info")=ApplicationValue
	Application.unlock
end Sub

'��ȡ��ҳ�ļ���
Function SelfName()
    SelfName = Mid(Request.ServerVariables("URL"),InstrRev(Request.ServerVariables("URL"),"/")+1)
End Function

%>

    