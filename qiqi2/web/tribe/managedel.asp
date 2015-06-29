<!--#include file="../inc/index_conn.asp"-->
<!--#include file="../inc/function.asp"-->
<!--#INCLUDE FILE="../config.asp"-->
<!--#include file="../user/checklogin.asp"-->
<%
id=request("id")
action=request("action")
if action="del"  then
      Set rs=server.createobject("ADODB.RECORDSET")
	  rs.open "Select * From "&jieducm&"_tribe where username='"&session("useridname")&"'" ,Conn,3,3  
	  if  rs.eof then
            Response.Write("<script>alert('你没有权限踢出此会员！');window.location.href='index.asp';</script>")
			response.End()
      end if
	  
      Set rs=server.createobject("ADODB.RECORDSET")
	  rs.open "Select tribeid ,username From "&jieducm&"_register where id="&id&"" ,Conn,3,3  
	  if not rs.eof then
	    if rs("username")=session("useridname")  then
	        Response.Write("<script>alert('酋长不能踢出！');history.back();</script>")
			response.End()
	    end if
		username=rs("username")
    	rs("tribeid")=0
    	rs.update
		rs.close	
		
      end if
	   sql="delete from "&jieducm&"_tribebook where username='"&username&"'"
      conn.execute(sql)
      response.Redirect("manage.asp")
end if
%>