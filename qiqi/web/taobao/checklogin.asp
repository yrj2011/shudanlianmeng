<%
restr="index.asp"'�ļ�������ڵ�ǰ�ļ���ʱ
pathc= replace(replace(server.mappath(restr),server.mappath("/"),""),restr,"")
pathch=replace(pathc,"\","")

Url=split(request.servervariables("script_name"),"/")
FileName=Url(ubound(Url))
Files=Left(FileName,InstrRev(FileName,".")-1)

if stupquxin=1 or stupqushou=1 then
if stupquxin=1 and pathch="taobao" then
  Response.write "<div align='center'><font size=6 color=#666666>�Ա�����������ά���У���ͣ�ر�......</font><br>"
   response.Write"<object classid='clsid:D27CDB6E-AE6D-11cf-96B8-444553540000' codebase='http://download.macromedia.com/pub/shockwave/cabs/flash/swflash.cab#version=7,0,19,0' width='550' height='400'>"
  response.Write"<param name='movie' value='../images/4759.swf' />"
  response.Write"<param name='quality' value='high' />"
  response.Write"<embed src='../images/4759.swf' quality='high' pluginspage='http://www.macromedia.com/go/getflashplayer' type='application/x-shockwave-flash' width='550' height='400'></embed>"
 response.Write"</object><div>"
  response.Write("<br><br><a href='javascript:history.back()'>����</a>")
   Response.End()
elseif stupqushou=1 and pathch="pai" then
  Response.write "<div align='center'><font size=6 color=#666666>��������������ά���У���ͣ�ر�......</font><br>"
   response.Write"<object classid='clsid:D27CDB6E-AE6D-11cf-96B8-444553540000' codebase='http://download.macromedia.com/pub/shockwave/cabs/flash/swflash.cab#version=7,0,19,0' width='550' height='400'>"
  response.Write"<param name='movie' value='../images/4759.swf' />"
  response.Write"<param name='quality' value='high' />"
  response.Write"<embed src='../images/4759.swf' quality='high' pluginspage='http://www.macromedia.com/go/getflashplayer' type='application/x-shockwave-flash' width='550' height='400'></embed>"
 response.Write"</object><div>"
  response.Write("<br><br><a href='javascript:history.back()'>����</a>")
   Response.End()
end if
 
end if
     
			  	Sql = "select * from "&jieducm&"_register where username='"&session("useridname")&"'"
				Set rs = Server.CreateObject("Adodb.RecordSet")
				rs.Open Sql,conn,1,1
				IF rs.Eof Then
                 Response.Write("<script>alert('��û�е�¼�����¼��ʱ�����¼�����!');window.location.href='../login.asp';</script>")
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
				    Response.Write("<script>alert('�Բ�����˺Ż�û��ͨ������Ա���!');window.location.href='../index.asp';</script>")
					response.End()
		        elseif taobao=0 then
				   	Response.Write("<script>alert('�Բ���!����û�п��Ա�����Ա��');window.location.href='../index.asp';</script>")
					response.End()
				elseif paipai=0 and pathch="pai" then
				   	Response.Write("<script>alert('�Բ���!����û�п���������Ա��');window.location.href='../index.asp';</script>")
					response.End()
				elseif tribeid=0 and stuptribestart=1 then
				     Response.Write("<script>alert('���ȼ���һ��������ܽ�����������!');window.location.href='../tribe/';</script>")
					response.End()
				elseif stupphonestart=1 and dst =0 then
				     Response.Write("<script>alert('���Ȱ��ֻ����ܽ�����������!');window.location.href='../user/Center_Userlock.asp';</script>")
					 response.End()
				elseif login_ip<>Request.ServerVariables("REMOTE_ADDR") and login_ip<>"0" then
				     session("useridname")=""
                     session("czm")=""
			         Response.Write("<script>alert('�����õ�ip�͵�½��ip��ͬ!');window.location.href='../index.asp';</script>")
			         response.End()
				end if
				
if vip=1  then			
sss= DateDiff( "d", vipdel, Now())
if int(sss)>int(vipdate) then
sqlr="update "&jieducm&"_register set  vip=0 where username='"&session("useridname")&"'"
conn.execute sqlr
Response.Write("<script>alert('��vip��Ա����ѵ���!');window.location.href='index.asp';</script>")	
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