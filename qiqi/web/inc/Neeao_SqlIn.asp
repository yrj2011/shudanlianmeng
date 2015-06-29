<%
'***********************************
'谨以此版本献给我深爱的和深爱着我的人--彦！
'***********************************
'--------版权说明------------------
'SQL通用防注入程序 V3.1 β
'2.0强化版，对代码做了一点优化，加入自动封注入者Ip的功能！^_^
'3.0版，加入后台登陆查看注入记录功能，方便网站管理员查看非法记录，以及删除以前的记录，是否对入侵者Ip解除封锁！
'3.1 β版，加入对cookie部分的过滤，加入了对用js书写的asp程序的支持！

'--------数据库连接部分--------------
dim dbkillSql,killSqlconn,connkillSql
On Error Resume Next
db1="/jieducm/#jieducm.asp"
Set killSqlconn = Server.CreateObject("ADODB.Connection")
connkillSql="Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & Server.MapPath(db1)
killSqlconn.Open connkillSql
If Err Then
	err.Clear
	Set killSqlconn = Nothing
	Response.Write "数据库连接出错，请检查连接字串。"
	Response.End
End If
'--------定义部份------------------
Dim N_Post,N_Get,N_In,N_Inf,N_Xh,N_db,N_dbstr,Kill_IP,WriteSql
Dim aApplicationValue
If IsArray(Application("Neeao_config_info"))=False Then Call PutApplicationValue()
aApplicationValue = Application("Neeao_config_info")
'获取配置信息
N_In = aApplicationValue(0)
Kill_IP = aApplicationValue(1) 
WriteSql = aApplicationValue(2)
alert_url = aApplicationValue(3)
alert_info = aApplicationValue(4)
kill_info = aApplicationValue(5)
N_type = aApplicationValue(6)
Sec_Forms = aApplicationValue(7)
Sec_Form_open = aApplicationValue(8)

'安全页面参数
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

'输出警告信息
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

'判断注入类型函数
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

'sql通用防注入主函数
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

'将注入记录记录到数据库函数
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

'获取本页文件名
Function SelfName()
    SelfName = Mid(Request.ServerVariables("URL"),InstrRev(Request.ServerVariables("URL"),"/")+1)
End Function

%>

    