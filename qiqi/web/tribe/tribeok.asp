<!--#include file="../inc/index_conn.asp"-->
<!--#include file="../inc/function.asp"-->
<!--#INCLUDE FILE="../config.asp"-->
<!--#include file="../user/checklogin.asp"-->
<%
tribecode=Replace(Trim(request("tribecode")),"'","''")
bulouid=request("bulouid")
action=request("action")
if action="add"  then

      Set rs=server.createobject("ADODB.RECORDSET")
	  rs.open "Select * From "&jieducm&"_tribe where triberad='"&tribecode&"' and id="&bulouid&"" ,Conn,3,3  
	  if  rs.eof then
            Response.Write("<script>alert('�����벻��ȷ������������');history.back();</script>")
			response.End()
      end if
	  
      Set rs=server.createobject("ADODB.RECORDSET")
	  rs.open "Select tribeid From "&jieducm&"_register where username='"&session("useridname")&"' " ,Conn,3,3  
	  if not rs.eof then
    	rs("tribeid")=bulouid
    	rs.update
		rs.close	
		Response.Write("<script>alert('�����ɹ�!���Ѽ����˴˲���');window.location.href='index.asp';</script>")
		response.End()
      end if
	  
elseif action="del" then

      Set rs=server.createobject("ADODB.RECORDSET")
	  rs.open "Select * From "&jieducm&"_tribe where username='"&session("useridname")&"' and id="&request("tribeidid")&"" ,Conn,3,3  
	  if  not rs.eof then
            Response.Write("<script>alert('���Ǳ��������������˳��˲��䣡');window.location.href='index.asp';</script>")
			response.End()
      end if
	  
      Set rs=server.createobject("ADODB.RECORDSET")
	  rs.open "Select tribeid From "&jieducm&"_register where username='"&session("useridname")&"' " ,Conn,3,3  
	  if rs.eof then
	  		Response.Write("<script>alert('�������������ԣ�');window.location.href='index.asp';</script>")
			response.End()
	  else	
    	rs("tribeid")=0
    	rs.update
		rs.close	
		Response.Write("<script>alert('�����ɹ�!�����˳��˲��䣡');window.location.href='index.asp';</script>")
		response.End()
      end if
end if
%>