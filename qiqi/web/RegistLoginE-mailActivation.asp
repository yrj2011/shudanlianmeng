<!--#include file="inc/index_conn.asp"-->
<!--#include file="inc/function.asp"-->
<!--#include file="config.asp"-->
<!--#include file="md5.asp"-->
<%
nameactive=request.QueryString("name")
cid=request.QueryString("cid")
sql="select *  from "&jieducm&"_register where id="&cid&" and activestart=0"
Set Rs = Server.CreateObject("Adodb.RecordSet")
Rs.Open Sql,conn,3,3
if NOT rs.EOF  then
 username=rs("username")
 rnow=rs("now")
 sss= DateDiff( "h", rnow, Now())
 if int(sss)>24 then
    response.write "<script>alert('�Բ�����ļ���ʱ���ѹ���������ע��ɣ�');window.location.href='index.asp';</script>"
    response.End()
 elseif md5(username)<>nameactive then
    response.write "<script>alert('��ڳ���');window.location.href='index.asp';</script>"
    response.End()
 end if
 rs("activestart")=1
 rs.update
 rs.close
 session("useridname")=username
 response.write "<script>alert('�˻�����ɹ���');window.location.href='index.asp';</script>"
 response.End()
else
   response.write "<script>alert('��ڳ������˻��Ѽ��');window.location.href='index.asp';</script>"
   response.End()
End IF
%>