<!--#include file="../inc/index_conn.asp"-->
<!--#include file="../inc/function.asp"-->
<!--#include file="checklogin.asp"-->
<%
Set rs=server.createobject("ADODB.RECORDSET")
rs.open "select * from "&jieducm&"_jieshou where classid='1' and start='4' and username='"&session("useridname")&"'",conn,1,1
if  (rs.eof) or  (rs.bof) then
  Response.Write("<script>alert('�Բ���,��ʱû�г���12Сʱ��ȷ�Ϻ���������!');window.history.back();</script>")
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
rsjieducm("class")="�Ա�������"
rsjieducm("content")="����ID:"&id&"����12Сʱϵͳ�Զ�ȷ�Ϻ�����õ�"&price&"Ԫ�������"&fabu1&"��������"&stupfajifei&"����"
rsjieducm("price")="+"&price
rsjieducm("jiegou")="ȷ���Ѻ����������!"
rsjieducm.update
rsjieducm.close	

sqlr="update "&jieducm&"_register set  jifei=jifei+"&stupfajifei&" where username='"&fusername&"'"
conn.execute sqlr

Set rsjieducm=server.createobject("ADODB.RECORDSET")
rsjieducm.open "Select * From "&jieducm&"_record " ,Conn,3,3  
rsjieducm.addnew
rsjieducm("username")=fusername
rsjieducm("num")=num
rsjieducm("class")="�Ա�������"
rsjieducm("content")="����ID:"&id&"����12Сʱϵͳ�Զ�ȷ�Ϻ���,�Է��õ�"&price&"Ԫ�������"&fabu1&"��������"&stupfajifei&"����"
rsjieducm("price")=0
rsjieducm("jiegou")="ȷ���Ѻ����������!"
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
  Response.Write("<script>alert('�Բ���,��ʱû�г���12Сʱ��ȷ�Ϻ���������!');window.history.back();</script>")
  response.end()	
else
  Response.Write("<script>alert('���ι��Զ�ȷ�Ϻ�������12Сʱ������"&numbersstart&"��!');window.location.href='ReMission.asp';</script>")
  response.end()	
end if


%>