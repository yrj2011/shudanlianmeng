<%
restr="index.asp"'文件必须存在当前文件夹时
pathc= replace(replace(server.mappath(restr),server.mappath("/"),""),restr,"")
pathch=replace(pathc,"\","")

Url=split(request.servervariables("script_name"),"/")
FileName=Url(ubound(Url))
Files=Left(FileName,InstrRev(FileName,".")-1)

if stupquxin=1 or stupqushou=1 then
if stupquxin=1 and pathch="taobao" then
  Response.write "<div align='center'><font size=6 color=#666666>淘宝区数据正在维护中，暂停关闭......</font><br>"
   response.Write"<object classid='clsid:D27CDB6E-AE6D-11cf-96B8-444553540000' codebase='http://download.macromedia.com/pub/shockwave/cabs/flash/swflash.cab#version=7,0,19,0' width='550' height='400'>"
  response.Write"<param name='movie' value='../images/4759.swf' />"
  response.Write"<param name='quality' value='high' />"
  response.Write"<embed src='../images/4759.swf' quality='high' pluginspage='http://www.macromedia.com/go/getflashplayer' type='application/x-shockwave-flash' width='550' height='400'></embed>"
 response.Write"</object><div>"
  response.Write("<br><br><a href='javascript:history.back()'>返回</a>")
   Response.End()
elseif stupqushou=1 and pathch="pai" then
  Response.write "<div align='center'><font size=6 color=#666666>拍拍区数据正在维护中，暂停关闭......</font><br>"
   response.Write"<object classid='clsid:D27CDB6E-AE6D-11cf-96B8-444553540000' codebase='http://download.macromedia.com/pub/shockwave/cabs/flash/swflash.cab#version=7,0,19,0' width='550' height='400'>"
  response.Write"<param name='movie' value='../images/4759.swf' />"
  response.Write"<param name='quality' value='high' />"
  response.Write"<embed src='../images/4759.swf' quality='high' pluginspage='http://www.macromedia.com/go/getflashplayer' type='application/x-shockwave-flash' width='550' height='400'></embed>"
 response.Write"</object><div>"
  response.Write("<br><br><a href='javascript:history.back()'>返回</a>")
   Response.End()
end if
 
end if
     
			  	Sql = "select * from "&jieducm&"_register where username='"&session("useridname")&"'"
				Set rs = Server.CreateObject("Adodb.RecordSet")
				rs.Open Sql,conn,1,1
				IF rs.Eof Then
                 Response.Write("<script>alert('您没有登录，或登录超时，请登录后操作!');window.location.href='../login.asp';</script>")
			     response.End()
				else		
				fabudian=rs("fabudian")
				cunkuan=rs("cunkuan")
				jifei=rs("jifei")
				isgives=rs("isgives")
				czm=rs("czm")
				isdun=rs("isdun")
				tjr=Replace(Trim(rs("tjr")),"'","''")	
				level1=Replace(Trim(rs("level1")),"'","''")
				taobao=Replace(Trim(rs("taobao")),"'","''")
				paipai=Replace(Trim(rs("paipai")),"'","''")
				vip=Replace(Trim(rs("vip")),"'","''")	
				vipdel=rs("vipdel")
				vipdate=rs("vipdate")	
				isfa=Replace(Trim(rs("isfa")),"'","''")
				isjie=Replace(Trim(rs("isjie")),"'","''")
				dst=Replace(Trim(rs("dst")),"'","''")
				tribeid=rs("tribeid")
				login_ip=rs("login_ip")
			  end if
			  
               if level1="0" then
			   		  session("useridname")=""
                     session("czm")=""
				    Response.Write("<script>alert('对不起此账号还没有通过管理员审核!');window.location.href='../index.asp';</script>")
					response.End()
		        elseif taobao=0 then
				   	Response.Write("<script>alert('对不起!您还没有开淘宝区会员！');window.location.href='../index.asp';</script>")
					response.End()
				elseif paipai=0 and pathch="pai" then
				   	Response.Write("<script>alert('对不起!您还没有开拍拍区会员！');window.location.href='../index.asp';</script>")
					response.End()
				elseif tribeid=0 and stuptribestart=1 then
				     Response.Write("<script>alert('请先加入一个部落才能进行其它操作!');window.location.href='../tribe/';</script>")
					response.End()
				elseif stupphonestart=1 and dst =0 then
				     Response.Write("<script>alert('请先绑定手机才能进行其它操作!');window.location.href='../user/Center_Userlock.asp';</script>")
					 response.End()
				elseif login_ip<>Request.ServerVariables("REMOTE_ADDR") and login_ip<>"0" then
				     session("useridname")=""
                     session("czm")=""
			         Response.Write("<script>alert('你设置的ip和登陆的ip不同!');window.location.href='../index.asp';</script>")
			         response.End()
				end if
				
if vip=1  then			
sss= DateDiff( "d", vipdel, Now())
if int(sss)>int(vipdate) then
sqlr="update "&jieducm&"_register set  vip=0 where username='"&session("useridname")&"'"
conn.execute sqlr
Response.Write("<script>alert('您vip会员身份已到期!');window.location.href='index.asp';</script>")	
response.end()
end if
end if
								  
sql="select count(*) as weidu1 from "&jieducm&"_china_message where uid='"&session("useridname")&"' and hits=0"
Set Rs = Server.CreateObject("Adodb.RecordSet")
Rs.Open Sql,conn,1,1
if not rs.eof then
weidu1=rs("weidu1")
end if
rs.close
set rs=nothing
 %>