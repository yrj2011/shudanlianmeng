<!--#include file="../inc/index_conn.asp"-->
<!--#include file="../inc/function.asp"-->
<!--#include file="checklogin.asp"-->
<%
Set rs=server.createobject("ADODB.RECORDSET")
rs.open "select * from "&jieducm&"_jieshou where classid='1' and start='4' and username='"&session("useridname")&"'",conn,1,1
if  (rs.eof) or  (rs.bof) then
  Response.Write("<script>alert('对不起,暂时没有超过12小时需确认好评的任务!');window.history.back();</script>")
  response.end()
else
  DO WHILE NOT rs.EOF	
  fusername=rs("username2")
  pid=rs("pid")
  nowj=rs("nowok")
  numbersstart=0
sss= DateDiff( "h", nowj, Now())
if int(sss)>12 then

numbersstart=numbersstart+1
Set rsjieducm=server.createobject("ADODB.RECORDSET")
rsjieducm.open "select * from "&jieducm&"_zhongxin where pid='"&pid&"' and start='4' ",conn,3,3
if  not(rsjieducm.eof) then
  bei=rsjieducm("bei")
  addfabu=rsjieducm("addfabu")
  price=rsjieducm("price")
  jieducm_fb=rs("jieducm_fb")
  rsjieducm("start")="5"
  rsjieducm.update
  rsjieducm.close  
end if

Set rsjieducm=server.createobject("ADODB.RECORDSET")
rsjieducm.open "select * from "&jieducm&"_jieshou where pid='"&pid&"' and start='4'",conn,3,3
if  not(rsjieducm.eof) then
  rsjieducm("start")="5"
  rsjieducm.update
  rsjieducm.close  
end if


fabu=jieducm_fb+addfabu

if vip="1" then
fabu1=stupvipjifei*int(fabu)
else		
if jifei>stupjifeibig then
fabu1=stupkou*int(fabu)
else
fabu1=fabu
end if
end if

Set rsjieducm=server.createobject("ADODB.RECORDSET")
rsjieducm.open "Select * From "&jieducm&"_register where username='"&session("useridname")&"'" ,Conn,3,3 
rsjieducm("jifei")=rsjieducm("jifei")+ stupfajifei
rsjieducm("fabudian")=rsjieducm("fabudian")+fabu1
rsjieducm("cunkuan")=rsjieducm("cunkuan")+price
rsjieducm.update
rsjieducm.close

num=filepath&year(now)&month(now)&day(now)&hour(now)&minute(now)&second(now)
Set rsjieducm=server.createobject("ADODB.RECORDSET")
rsjieducm.open "Select * From "&jieducm&"_record " ,Conn,3,3  
rsjieducm.addnew
rsjieducm("username")=session("useridname")
rsjieducm("num")=num
rsjieducm("class")="淘宝区任务"
rsjieducm("content")="任务ID:"&id&"超过12小时系统自动确认好评你得到"&price&"元担保金和"&fabu1&"个发布点"&stupfajifei&"积分"
rsjieducm("price")="+"&price
rsjieducm("jiegou")="确认已好评任务完成!"
rsjieducm.update
rsjieducm.close	

sqlr="update "&jieducm&"_register set  jifei=jifei+"&stupfajifei&" where username='"&fusername&"'"
conn.execute sqlr

Set rsjieducm=server.createobject("ADODB.RECORDSET")
rsjieducm.open "Select * From "&jieducm&"_record " ,Conn,3,3  
rsjieducm.addnew
rsjieducm("username")=fusername
rsjieducm("num")=num
rsjieducm("class")="淘宝区任务"
rsjieducm("content")="任务ID:"&id&"超过12小时系统自动确认好评,对方得到"&price&"元担保金和"&fabu1&"个发布点"&stupfajifei&"积分"
rsjieducm("price")=0
rsjieducm("jiegou")="确认已好评任务完成!"
rsjieducm.update
rsjieducm.close
  
end if

  rs.MoveNext
  Loop
end if
rs.close
set rs=nothing
set rsjieducm=nothing
conn.close
set conn=nothing

if numbersstart=0 then
  Response.Write("<script>alert('对不起,暂时没有超过12小时需确认好评的任务!');window.history.back();</script>")
  response.end()	
else
  Response.Write("<script>alert('本次共自动确认好评超过12小时的任务"&numbersstart&"条!');window.location.href='ReMission.asp';</script>")
  response.end()	
end if


%>