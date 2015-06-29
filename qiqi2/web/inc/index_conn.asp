<%
'*************************************
'* 捷度传媒-智慧拓展未来！
'* Web : http://www.jieducm.cn/
'* By  : Piao 
'* QQ  : 616591415 
'* Date: 2009-8-18
'本文件非常重要请不要自行修改。否则将导致正个平台无法使用
'*************************************
'自定义需要过滤的字串,用 "|" 分隔
On Error Resume Next
Dim Fy_Get, Fy_In, Fy_Inf, Fy_Xh
If LCase(Request.ServerVariables("HTTPS")) = "off" Then
 strTemp = "http://"
Else
 strTemp = "https://"
End If
strTemp = strTemp & Request.ServerVariables("SERVER_NAME")
If Request.ServerVariables("SERVER_PORT") <> 80 Then strTemp = strTemp & ":" & Request.ServerVariables("SERVER_PORT")
strTemp = strTemp & Request.ServerVariables("URL")
If Trim(Request.QueryString) <> "" Then strTemp = strTemp & "?" & Trim(Request.QueryString)
strtempurl = LCase(strTemp)
strTemp=Trim(Request.QueryString)
If Instr(strTemp,"'") or Instr(strTemp,",") or Instr(strTemp,"%20and%20") or Instr(strTemp,"%20or%20") or Instr(strTemp,"select%20") or Instr(strTemp,"insert%20") or Instr(strTemp,"delete%20from") or Instr(strTemp,"count(") or Instr(strTemp,"drop%20table") or Instr(strTemp,"update%20") or Instr(strTemp,"truncate%20") or Instr(strTemp,"asc(") or Instr(strTemp,"mid(") or Instr(strTemp,"char(") or Instr(strTemp,"0x61006E006400") or Instr(strTemp,"0x63006800610072002800") or Instr(strTemp,"0x6F007200") or Instr(strTemp,"0x2700") or Instr(strTemp,"min(") or Instr(strTemp,"sum(") or Instr(strTemp,"max(") or Instr(strTemp,"xp_cmdshell") or Instr(strTemp,"0x31003D003100") or Instr(strTemp,"exec%20master") or Instr(strTemp,"net%20localgroup%20administrators")  or Instr(strTemp,"cmd") or Instr(strTemp,"net%20user")  then
response.write("<script language=javascript>;history.back();</script>")
   response.end
End If

Fy_In = "'|and|exec|insert|select|delete|update|count|mid|master|truncate|char|declare|fabudian|cunkuan|jifei|xp_cmdshell"
Fy_In2 = "'|UNION| |%|or|+|(|)|where|'|table|drop|and|exec|insert|select|delete|update|count|mid|master|truncate|char|declare|fabudian|cunkuan|jifei|xp_cmdshell"
Fy_Inf = Split(Fy_In, "|")
Fy_Inf2 = Split(Fy_In2, "|")
If Request.QueryString <> "" Then
    For Each Fy_Get In Request.QueryString
        For Fy_Xh = 0 To UBound(Fy_Inf2)
            If InStr(LCase(Request.QueryString(Fy_Get)), Fy_Inf2(Fy_Xh)) <> 0 Then
                Response.Write "<Script Language=JavaScript>alert('请不要在参数中包含非法字符尝试注入！');window.location.href='../index.asp';</Script>"
                Fy_Get = ""
                Fy_In = ""
                Fy_Inf = ""
                Fy_Xh = ""
                Response.End
            End If
        Next
    Next
End If

If Request.form <> "" Then
    For Each Fy_Get In Request.form
        For Fy_Xh = 0 To UBound(Fy_Inf)
            If InStr(LCase(Request.form(Fy_Get)), Fy_Inf(Fy_Xh)) <> 0 Then
                Response.Write "<Script Language=JavaScript>alert('请不要在参数中包含非法字符尝试注入！');window.location.href='../index.asp';</Script>"
                Fy_Get = ""
                Fy_In = ""
                Fy_Inf = ""
                Fy_Xh = ""
                Response.End
            End If
        Next
    Next
End If

Dim conn,WebSqlType,MySQLUser,MySqlPass,MySqlDb
'******************************************************************************************************
MyAccEssFile="Access/jieducm.mdb" '数据库的路径,AccEss专用
'---------------------------------
MySQLUser="liuzhong1177" '连接数据库的用户名
'---------------------------------
MySqlPass="liuzhong1177a" '连接数据库的密码
'---------------------------------
MySqlDb="liuzhong1177" '数据库名称
'---------------------------------
SQLAccEssType=1 '1为SQL数据库存 , 0为ACCESS数据库
'---------------------------------
jieducode="fe0b9e59fa853198" '网站授权码

jieducm="jieducm_gong"

smsname="778892" '手机短信用户名
smspwdphone="778892" '手机短信密码


'请尊重程序员的劳动成果，谢谢！
'以下内容请不要自行修改。否则将导致整个平台无法使用，数据丢失后果自负！
WebPassXS="试用版用户请及时升级为正式版或授权版,同时也感谢你使用!捷度传媒独立开发的互刷平台<br><a href=http://www.jieducm.cn target=_blank>www.jieducm.cn</a>"
weberr="网站正在维护中......感谢您使用捷度传媒独立开发的互刷平台<a href=http://www.jieducm.cn target=_blank>www.jieducm.cn</a>"
'******************************************************************************************************
If SQLAccEssType=1 then
On Error Resume Next
Set Conn=server.createobject("adodb.connection")  
Conn.ConnectionTimeout=20 
Session.Timeout=30                    
' Conn.open "driver={SQL Server};server=(local);uid="&MySQLUser&";pwd="&MySqlPass&";database="&MySqlDb&""
 strconn="Provider=SQLOLEDB;uid=sa;pwd=123456;Server=ZHAO-JISHU013\SQLEXPRESS;DATABASE=yitop90800"
    Conn.Open strconn
' 如上面连接字符不能用，可用下面这个。把里面的信息改为你自己的即可
'Conn.Open "provider=sqloledb;data source=211.144.65.79,1433;User ID=yitop90800;pwd=yitop90800;Initial Catalog=yitop90800"

If Err.Number <> 0 Then
  Response.write(weberr)
   Response.End()
end if
Else
conndata="Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & Server.MapPath(""&MyAccEssFile&"")
On Error Resume Next
set conn = Server.CreateObject("ADODB.Connection")
conn.Open conndata
If Err.Number <> 0 Then
  Response.write(weberr)
  Response.End()
end if
End If

'******************************************************************************************************
set Rsstup=server.createobject("adodb.recordset")
sql="select * from "&jieducm&"_stup where password='"&jieducode&"' and webpass='"&WebPassXS&"'"
Rsstup.open sql,conn,1,1
if (Rsstup.bof) and (Rsstup.eof) then
response.Write WebPassXS
response.End()
else
stupname= Rsstup("name")
 stuptitle=Rsstup("title")
 stupkey=  Rsstup("key")
 stupdesc= Rsstup("desc")
 stupgonggao=Rsstup("gonggao")
 stupkouhao=Rsstup("kouhao")
 stupurl=  Rsstup("url")
 stuplogo= Rsstup("logo")
 stupaddress=Rsstup("address")
 stupqq=   Rsstup("qq")
 stupqq2=   Rsstup("qq2")
 stupqq3=   Rsstup("qq3")
 stupqq4=   Rsstup("qq4")
  stupqq5=   Rsstup("qq5")
 stupqq6=   Rsstup("qq6")
 stupzfb=  Rsstup("zfb")
 stupcft=  Rsstup("cft")
 stupemail=Rsstup("email")
 stupphone=Rsstup("phone")
 stuppass=Rsstup("password")
 stuptel=Rsstup("tel")
 stupicp=  Rsstup("icp")
 stupcopy= Rsstup("copy")
 stupjifei   = Rsstup("jifei")
 stupfa = Rsstup("fa")
 stupkuan= Rsstup("kuan")
 stupkuanvip=Rsstup("kuanvip")
 stupkou=Rsstup("kou")
 stupgeshu=Rsstup("geshu")
 stupzhang=Rsstup("zhang")
 stupping=Rsstup("ping")
 stupzhu=Rsstup("zhu")
 stuptjrjifei=Rsstup("tjrjifei")
 stupvipjifei=Rsstup("vipjifei")
 stupfajifei=Rsstup("fajifei")
 stupshouc=Rsstup("shouc")
 stupshouf=Rsstup("shouf")
 stupshouf2=Rsstup("shouf2")
 stupvipjin=Rsstup("vipjin")
 stupshouxu=Rsstup("shouxu")
 stupwu=Rsstup("wu")
 stupshi=Rsstup("shi")
 stupwushi=Rsstup("wushi")
 stupyibai=Rsstup("yibai")
 stupwubai=Rsstup("wubai")
 stupka=Rsstup("ka")
 stupjifeidi=Rsstup("jifeidi")
 stupjifeigao=Rsstup("jifeigao")
 stupalllogin=Rsstup("alllogin")
 stupfhc=Rsstup("fhc")
 stupfhcshu=Rsstup("fhcshu")
 stupquxin=rsstup("quxin")
 stupqushou=rsstup("qushou")
 stupquvip=rsstup("quvip")
 stupbuyhao=rsstup("buyhao")
 stupxinshoufa=rsstup("xinshoufa")
 stupshouchangfa=rsstup("shouchangfa")
 stupvipfa=rsstup("vipfa")
 stupjifeibig=rsstup("jifeibig")
 stupsongfa=rsstup("songfa")
 stupenchfa=rsstup("enchfa")
 stupsongshou=rsstup("songshou")
 stupsongji=rsstup("songji")
 stupcar=rsstup("car")
 stupwangyin=rsstup("wangyin")
 stupalipay=rsstup("alipay")
 stuplogin=rsstup("login")
 stupCheckCode=rsstup("CheckCode")
 stupisgive=rsstup("isgive")
 stuppopup=rsstup("popup")
 stupwebpass=rsstup("webpass")
 stupweberr=rsstup("weberr")
 stupregister_taobao=rsstup("register_taobao")
 stupregister_pai=rsstup("register_pai")
 stupregister_you=rsstup("register_you")
 stuptjrzhu=rsstup("tjrzhu")
 stuptjrfabudian=rsstup("tjrfabudian")
 stupMailServer=rsstup("MailServer")
 stupMailServerUserName=rsstup("MailServerUserName")
 stupMailServerPassWord=rsstup("MailServerPassWord")
 stupMailDomain=rsstup("MailDomain")
 stupjieducmwebad=rsstup("jieducm_webad")
 stupjieducmadweb=rsstup("jieducm_adweb")
 stupisemail=rsstup("isemail")
 stupismessage=rsstup("ismessage")
 stupemailcontent=rsstup("emailcontent")
 stupmessagecontent=rsstup("messagecontent")
 stupemailactive=rsstup("emailactive")
  stuppuuser=rsstup("puuser")
 stupvipuser=rsstup("vipuser")
 stuptribestart=rsstup("tribestart")
 stupphonestart=rsstup("phonestart")
 stupjieweiok=rsstup("jieweiok")
 stupfaweiok=rsstup("faweiok")
 stupinvite=rsstup("invite")
 stupexitauto=rsstup("exitauto")
 stupvip_car_jieducm=rsstup("vip_car_jieducm")
  stupcang=rsstup("cang")
 stupkefu_pic=rsstup("kefu_pic")
 stupjieducm_gonggao=rsstup("jieducm_gonggao")
 stupjieducm_MerId=rsstup("jieducm_MerId")
 stupjieducm_MerKey=rsstup("jieducm_MerKey")
 stupyibao=rsstup("yibao")
 stuppayis=rsstup("payis")
 stuppayjifei=rsstup("payjifei")
 stuppaynum=rsstup("paynum")
 stupvippaynum=rsstup("vippaynum")
 stuppupaynum=rsstup("pupaynum")
 stup_RndQueryNum=rsstup("RndQueryNum")
 stup_DefQueryNum=rsstup("DefQueryNum")
 Rsstup.close
 set Rsstup=Nothing
end if
if stupinvite=2 then
response.Redirect("../invite.html")
response.End()
end if
shij=dateadd("n",10,now())
Conn.execute("Delete  from "&jieducm&"_online where ActiveTime<='"&now()&"'")

If session("useridname")<>"" or session("useridname")<>empty then
'删除活动时间在10分种以前的用户
Set rsline=server.createobject("adodb.recordset")
'在线信息更新开始
yousql="select * from "&jieducm&"_online where username='"&session("useridname")&"'"
set OnlineIp=conn.execute(yousql)
rsline.open yousql,conn,1,3
'搜索用户名，如果在线中没有你的用户名则把你加为在线者，如果有都把你新的活动时间加进去
if onlineIp.eof then
'增加
rsline.addnew
rsline("username")=session("useridname")
rsline("ActiveTime")=shij
rsline.Update
rsline.Close
else
'更新
rsline("ActiveTime")=shij
rsline.update
end if
end if
Set OnlineIp=nothing
Set rsline=nothing
Function HtmlEncode(Content)
  Content = Replace(Content, ">", "&gt;") 
  Content = Replace(Content, "<", "&lt;")
  Content = Replace(Content, "'", "") 
  HtmlEncode = content 
End Function

sub CloseConn()
On Error Resume Next
	Conn.close
	set Conn=nothing
end sub
%>
